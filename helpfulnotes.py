from __future__ import annotations

import logging
from copy import deepcopy
from dataclasses import dataclass

import pandas as pd
from pptx import Presentation

from deck.engine.pptx.text import _replace_placeholders_in_slide_runs

from .media_kpis_summary_service_layer import (
    _ALLOWED_EFFECT_TYPE_TOKENS,
    _BRAND_COLUMN_CANDIDATES,
    _EFFECT_TYPE_COLUMN_CANDIDATES,
    _MAT_COLUMN_CANDIDATES,
    _NET_REVENUE_ROI_TITLE,
    _PROFIT_COLUMN_CANDIDATES,
    _SPEND_COLUMN_CANDIDATES,
    _coerce_numeric,
    _compute_brand_roi_values,
    _ensure_total_row,
    _find_chart_shape_by_title,
    _find_workbook_header_column,
    _iter_shapes_recursive,
    _load_chart_workbook,
    _mat_bucket_key,
    _normalize_media_template_brand_df,
    _ordered_unique_brands,
    _resolve_column,
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

_CHANNEL_COLUMN_CANDIDATES = (
    "Channel",
    "Media Type",
    "MediaType",
    "Media Channel",
    "Channel Name",
)


def _normalize_token(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return "".join(ch for ch in str(value).strip().lower() if ch.isalnum())


@dataclass
class ChannelBreakdownRow:
    channel: str
    mat_1_roi: float
    mat_1_spend: float
    mat_1_budget_pct: float
    mat_roi: float
    mat_spend: float
    mat_budget_pct: float


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

    ws.cell(row=1, column=start_col, value=f"Total {year_range.start_year}")
    ws.cell(row=1, column=end_col, value=f"Total {year_range.end_year}")
    ws.cell(row=total_row, column=start_col, value=roi_values.mat_1_roi)
    ws.cell(row=total_row, column=end_col, value=roi_values.mat_roi)

    _save_chart_workbook(chart, workbook)
    _update_series_and_category_caches(chart, workbook)
    return roi_values


def _find_breakdown_table_on_slide(slide):
    candidates = []
    for shape in _iter_shapes_recursive(slide.shapes):
        if not getattr(shape, "has_table", False):
            continue
        table = shape.table
        candidates.append(table)

    if not candidates:
        return None

    for table in candidates:
        try:
            if len(table.columns) == 7 and len(table.rows) >= 3:
                return table
        except Exception:
            continue

    for table in candidates:
        try:
            if len(table.columns) >= 7:
                return table
        except Exception:
            continue

    return candidates[0]


def _set_cell_text_preserve_formatting(cell, text: str) -> None:
    text_frame = cell.text_frame
    if not text_frame.paragraphs:
        paragraph = text_frame.add_paragraph()
    else:
        paragraph = text_frame.paragraphs[0]

    if not paragraph.runs:
        run = paragraph.add_run()
    else:
        run = paragraph.runs[0]

    run.text = str(text)

    for extra_run in paragraph.runs[1:]:
        extra_run.text = ""

    for extra_paragraph in text_frame.paragraphs[1:]:
        for extra_run in extra_paragraph.runs:
            extra_run.text = ""


def _ensure_breakdown_table_data_row_count(table, *, needed_data_rows: int, header_rows: int = 2) -> None:
    tbl = table._tbl
    current_rows = len(table.rows)

    if current_rows < header_rows + 1:
        raise ValueError(
            "Net Revenue ROI Breakdown table template must contain at least "
            f"{header_rows} header rows and 1 dummy data row."
        )

    template_tr = deepcopy(tbl.tr_lst[header_rows])

    while len(table.rows) < header_rows + needed_data_rows:
        tbl.append(deepcopy(template_tr))

    while len(table.rows) > header_rows + needed_data_rows:
        tbl.remove(tbl.tr_lst[-1])


def _format_roi(value: float) -> str:
    return f"{float(value or 0.0):.2f}"


def _format_spend_thousands(value: float) -> str:
    return f"{int(round(float(value or 0.0) / 1000.0)):,}"


def _format_budget_pct(value: float) -> str:
    return f"{float(value or 0.0) * 100:.1f}%"


def _ordered_unique_channel_labels(series: pd.Series) -> list[str]:
    ordered: list[str] = []
    seen: set[str] = set()
    for value in series.tolist():
        label = "" if value is None or pd.isna(value) else str(value).strip()
        key = _normalize_token(label)
        if not label or not key or key in seen:
            continue
        seen.add(key)
        ordered.append(label)
    return ordered


def _compute_brand_channel_breakdown_rows(
    media_template_brand_df: pd.DataFrame,
    *,
    brand: str,
    year_range,
) -> list[ChannelBreakdownRow]:
    df = _normalize_media_template_brand_df(media_template_brand_df)

    brand_column = _resolve_column(df, _BRAND_COLUMN_CANDIDATES)
    effect_type_column = _resolve_column(df, _EFFECT_TYPE_COLUMN_CANDIDATES)
    mat_column = _resolve_column(df, _MAT_COLUMN_CANDIDATES)
    channel_column = _resolve_column(df, _CHANNEL_COLUMN_CANDIDATES)
    profit_column = _resolve_column(df, _PROFIT_COLUMN_CANDIDATES)
    spend_column = _resolve_column(df, _SPEND_COLUMN_CANDIDATES)

    brand_key = _normalize_token(brand)
    brand_series = df[brand_column].map(_normalize_token)
    effect_series = df[effect_type_column].map(_normalize_token)
    mat_series = df[mat_column].map(lambda value: _mat_bucket_key(value, year_range=year_range))
    channel_display_series = df[channel_column].map(lambda value: "" if value is None or pd.isna(value) else str(value).strip())
    channel_key_series = channel_display_series.map(_normalize_token)

    filtered_mask = (
        brand_series.eq(brand_key)
        & effect_series.isin(_ALLOWED_EFFECT_TYPE_TOKENS)
        & mat_series.isin({"year1", "year2"})
        & channel_key_series.ne("")
    )
    filtered_df = df.loc[filtered_mask].copy()
    if filtered_df.empty:
        return []

    filtered_df["_bucket"] = mat_series.loc[filtered_mask].values
    filtered_df["_channel_display"] = channel_display_series.loc[filtered_mask].values
    filtered_df["_channel_key"] = channel_key_series.loc[filtered_mask].values
    filtered_df["_profit"] = pd.to_numeric(filtered_df[profit_column], errors="coerce").fillna(0.0)
    filtered_df["_spend"] = pd.to_numeric(filtered_df[spend_column], errors="coerce").fillna(0.0)

    total_spend_by_bucket = {
        "year1": float(filtered_df.loc[filtered_df["_bucket"] == "year1", "_spend"].sum()),
        "year2": float(filtered_df.loc[filtered_df["_bucket"] == "year2", "_spend"].sum()),
    }

    ordered_channels = _ordered_unique_channel_labels(filtered_df["_channel_display"])

    rows: list[ChannelBreakdownRow] = []
    for channel in ordered_channels:
        channel_key = _normalize_token(channel)
        channel_df = filtered_df.loc[filtered_df["_channel_key"] == channel_key]

        def _metrics(bucket_key: str) -> tuple[float, float, float]:
            bucket_df = channel_df.loc[channel_df["_bucket"] == bucket_key]
            spend_value = float(bucket_df["_spend"].sum())
            profit_value = float(bucket_df["_profit"].sum())
            roi_value = 0.0 if spend_value == 0 else round(profit_value / spend_value, 2)
            total_spend = float(total_spend_by_bucket.get(bucket_key, 0.0) or 0.0)
            budget_pct = 0.0 if total_spend == 0 else float(spend_value) / total_spend
            return roi_value, spend_value, budget_pct

        mat_1_roi, mat_1_spend, mat_1_budget_pct = _metrics("year1")
        mat_roi, mat_spend, mat_budget_pct = _metrics("year2")

        rows.append(
            ChannelBreakdownRow(
                channel=channel,
                mat_1_roi=mat_1_roi,
                mat_1_spend=mat_1_spend,
                mat_1_budget_pct=mat_1_budget_pct,
                mat_roi=mat_roi,
                mat_spend=mat_spend,
                mat_budget_pct=mat_budget_pct,
            )
        )

    return rows


def _update_net_revenue_roi_breakdown_table(
    slide,
    *,
    brand: str,
    year_range,
    media_template_brand_df: pd.DataFrame,
) -> None:
    table = _find_breakdown_table_on_slide(slide)
    if table is None:
        logger.info("No Net Revenue ROI Breakdown table found on slide for brand %r.", brand)
        return

    rows = _compute_brand_channel_breakdown_rows(
        media_template_brand_df,
        brand=brand,
        year_range=year_range,
    )

    _ensure_breakdown_table_data_row_count(table, needed_data_rows=len(rows), header_rows=2)

    if len(table.rows) >= 1 and len(table.columns) >= 5:
        _set_cell_text_preserve_formatting(table.cell(0, 1), str(year_range.start_year))
        _set_cell_text_preserve_formatting(table.cell(0, 4), str(year_range.end_year))

    for idx, row in enumerate(rows, start=2):
        _set_cell_text_preserve_formatting(table.cell(idx, 0), row.channel)
        _set_cell_text_preserve_formatting(table.cell(idx, 1), _format_roi(row.mat_1_roi))
        _set_cell_text_preserve_formatting(table.cell(idx, 2), _format_spend_thousands(row.mat_1_spend))
        _set_cell_text_preserve_formatting(table.cell(idx, 3), _format_budget_pct(row.mat_1_budget_pct))
        _set_cell_text_preserve_formatting(table.cell(idx, 4), _format_roi(row.mat_roi))
        _set_cell_text_preserve_formatting(table.cell(idx, 5), _format_spend_thousands(row.mat_spend))
        _set_cell_text_preserve_formatting(table.cell(idx, 6), _format_budget_pct(row.mat_budget_pct))


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
    _update_net_revenue_roi_breakdown_table(
        slide,
        brand=brand,
        year_range=year_range,
        media_template_brand_df=media_template_brand_df,
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

    normalized_media_template_brand_df = _normalize_media_template_brand_df(media_template_brand_df)
    brands = _ordered_unique_brands(normalized_media_template_brand_df)
    if not brands:
        return

    year_range = _resolve_scope_year_range(scope_df)

    if len(brands) == 1:
        _populate_net_revenue_roi_breakdown_slide(
            template_slide,
            brand=brands[0],
            year_range=year_range,
            media_template_brand_df=normalized_media_template_brand_df,
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
            media_template_brand_df=normalized_media_template_brand_df,
        )

    slide_mapper.delete_slide(prs, template_slide)
