from __future__ import annotations

import pandas as pd
from pptx import Presentation

from deck.engine.pptx.text import _replace_placeholders_in_slide_runs

from .media_kpis_summary_service_layer import _iter_shapes_recursive, _ordered_unique_brands, _resolve_scope_year_range
from .waterfall_service_layer import WaterfallSlideMapper

_MARKER_PHRASE = "Net Revenue ROI Breakdown"
_BRAND_PLACEHOLDER = "<BRAND>"
_MAT_1_PLACEHOLDER = "<MAT-1>"
_MAT_PLACEHOLDER = "<MAT>"


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
        _apply_net_revenue_roi_breakdown_placeholders(
            template_slide,
            brand=brands[0],
            start_year=year_range.start_year,
            end_year=year_range.end_year,
        )
        return

    slide_mapper = WaterfallSlideMapper()
    template_idx = list(prs.slides).index(template_slide)

    for idx, brand in enumerate(brands):
        slide = slide_mapper.duplicate_slide(prs, template_slide, insert_idx=template_idx + idx + 1)
        _apply_net_revenue_roi_breakdown_placeholders(
            slide,
            brand=brand,
            start_year=year_range.start_year,
            end_year=year_range.end_year,
        )

    slide_mapper.delete_slide(prs, template_slide)
