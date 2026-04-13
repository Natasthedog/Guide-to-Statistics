from __future__ import annotations

import logging

import pandas as pd
from pptx import Presentation

from deck.engine.pptx.text import _replace_placeholders_in_slide_runs

from .media_kpis_summary_service_layer import (
    _NET_REVENUE_ROI_TITLE,
    _compute_brand_roi_values,
    _ensure_total_row,
    _find_chart_shape_by_title,
    _find_workbook_header_column,
    _iter_shapes_recursive,
    _load_chart_workbook,
    _ordered_unique_brands,
    _resolve_scope_year_range,
    _save_chart_workbook,
    _update_nr_roi_yoy_arrow,
    _update_series_and_category_caches,
)
from .waterfall_service_layer import WaterfallSlideMapper

logger = logging.getLogger(__name__)

_MARKER_PHRASE = "Net Revenue ROI Breakdown"
_BRAND_PLACEHOLDER = "<BRAND>"
_MAT_1_PLACEHOLDER = "<MAT-1>"
_MAT_PLACEHOLDER = "<MAT>"
_TOTAL_MAT_1_PLACEHOLDER = "Total <MAT-1>"
_TOTAL_MAT_PLACEHOLDER = "Total <MAT>"


def _normalize_token(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return "".join(ch for ch in str(value).strip().lower() if ch.isalnum())


def _find_net_revenue_roi_breakdown_template_slide(prs: Presentation):
    marker = _MARKER_PHRASE
    marker_norm = _normalize_token(marker)

    for slide in prs.slides:
        for shape in _iter_shapes_recursive(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            text_value = shape.text_frame.text or ""
            if marker in text_value:
                return slide
            if marker_norm and marker_norm in _normalize_token(text_value):
                return slide
    return None


def _apply_net_revenue_roi_breakdown_placeholders(slide, *, brand: str, start_year: int, end_year: int) -> bool:
    replacements = {
        _BRAND_PLACEHOLDER: str(brand),
        _MAT_1_PLACEHOLDER: str(start_year),
        _MAT_PLACEHOLDER: str(end_year),
    }
    return _replace_placeholders_in_slide_runs(slide, replacements)


def _update_net_revenue_roi_breakdown_chart(
    slide,
    *,
    brand: str,
    year_range,
    media_template_brand_df: pd.DataFrame,
):
    chart_shape = _find_chart_shape_by_title(slide, _NET_REVENUE_ROI_TITLE)
    if chart_shape is None:
        logger.info(
            "No '%s' chart found on Net Revenue ROI Breakdown slide for brand %r.",
            _NET_REVENUE_ROI_TITLE,
            brand,
        )
        return _compute_brand_roi_values(
            media_template_brand_df,
            brand=brand,
            year_range=year_range,
        )

    roi_values = _compute_brand_roi_values(
        media_template_brand_df,
        brand=brand,
        year_range=year_range,
    )

    chart = chart_shape.chart
    workbook = _load_chart_workbook(chart)
    ws = workbook.active

    start_col = _find_workbook_header_column(
        ws,
        (_TOTAL_MAT_1_PLACEHOLDER, _MAT_1_PLACEHOLDER, f"Total {year_range.start_year}", str(year_range.start_year)),
    )
    end_col = _find_workbook_header_column(
        ws,
        (_TOTAL_MAT_PLACEHOLDER, _MAT_PLACEHOLDER, f"Total {year_range.end_year}", str(year_range.end_year)),
    )
    total_row = _ensure_total_row(ws)

    if start_col is None or end_col is None:
        raise ValueError(
            "Could not find Net Revenue ROI Breakdown chart year columns in embedded workbook. "
            f"Expected headers like {_TOTAL_MAT_1_PLACEHOLDER!r}/{_TOTAL_MAT_PLACEHOLDER!r} "
            f"or {_MAT_1_PLACEHOLDER!r}/{_MAT_PLACEHOLDER!r}."
        )

    # Leave Column1 untouched. Keep the literal 'Total' prefix in the year headers.
    ws.cell(row=1, column=start_col, value=f"Total {year_range.start_year}")
    ws.cell(row=1, column=end_col, value=f"Total {year_range.end_year}")
    ws.cell(row=total_row, column=start_col, value=roi_values.mat_1_roi)
    ws.cell(row=total_row, column=end_col, value=roi_values.mat_roi)

    _save_chart_workbook(chart, workbook)
    _update_series_and_category_caches(chart, workbook)
    return roi_values


def _populate_net_revenue_roi_breakdown_slide(
    slide,
    *,
    brand: str,
    year_range,
    media_template_brand_df: pd.DataFrame,
) -> None:
    _apply_net_revenue_roi_breakdown_placeholders(
        slide,
        brand=brand,
        start_year=year_range.start_year,
        end_year=year_range.end_year,
    )
    roi_values = _update_net_revenue_roi_breakdown_chart(
        slide,
        brand=brand,
        year_range=year_range,
        media_template_brand_df=media_template_brand_df,
    )
    _update_nr_roi_yoy_arrow(
        slide,
        mat_1_roi=float(roi_values.mat_1_roi or 0.0),
        mat_roi=float(roi_values.mat_roi or 0.0),
    )


def populate_net_revenue_roi_breakdown_slides(
    prs: Presentation,
    *,
    media_template_brand_df: pd.DataFrame,
    scope_df: pd.DataFrame,
) -> None:
    template_slide = _find_net_revenue_roi_breakdown_template_slide(prs)
    if template_slide is None:
        return

    brands = _ordered_unique_brands(media_template_brand_df)
    if not brands:
        return

    year_range = _resolve_scope_year_range(scope_df)

    if len(brands) == 1:
        _populate_net_revenue_roi_breakdown_slide(
            template_slide,
            brand=brands[0],
            year_range=year_range,
            media_template_brand_df=media_template_brand_df,
        )
        return

    slide_mapper = WaterfallSlideMapper()
    template_idx = list(prs.slides).index(template_slide)

    for idx, brand in enumerate(brands):
        slide = slide_mapper.duplicate_slide(prs, template_slide, insert_idx=template_idx + idx + 1)
        _populate_net_revenue_roi_breakdown_slide(
            slide,
            brand=brand,
            year_range=year_range,
            media_template_brand_df=media_template_brand_df,
        )

    slide_mapper.delete_slide(prs, template_slide)
