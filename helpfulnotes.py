from __future__ import annotations

import logging
from copy import deepcopy
from dataclasses import dataclass

import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

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
_ARROW_UTILITIES_MARKER = "<Arrow Utilities>"
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

_GREEN_INCREASE_SHAPE_NAME = "ROIStatus_GreenIncreaseArrow"
_PINK_DECREASE_SHAPE_NAME = "ROIStatus_PinkDecreaseArrow"
_NEW_SHAPE_NAME = "ROIStatus_NewSymbol"
_EQUAL_SHAPE_NAME = "ROIStatus_EqualSymbol"

_ROI_STATUS_EQUALITY_TOLERANCE = 0.01
_HEADER_ROW_COUNT = 2
_TABLE_COLUMN_COUNT = 8
# Channel | MAT-1 ROI | Spend | Budget % | ROI Status | MAT ROI | Spend | Budget %
_ROI_STATUS_COL_IDX = 4


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


def _find_arrow_utilities_slide(prs: Presentation):
    marker_norm = _normalize_token(_ARROW_UTILITIES_MARKER)
    for slide in prs.slides:
        for shape in _iter_shapes_recursive(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            text_value = shape.text_frame.text or ""
            if _ARROW_UTILITIES_MARKER in text_value:
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


def _find_breakdown_table_shape_and_table(slide):
    candidates = []
    for shape in _iter_shapes_recursive(slide.shapes):
        if not getattr(shape, "has_table", False):
            continue
        table = shape.table
        candidates.append((shape, table))

    if not candidates:
        return None, None

    for shape, table in candidates:
        try:
            if len(table.columns) == _TABLE_COLUMN_COUNT and len(table.rows) >= _HEADER_ROW_COUNT + 1:
                return shape, table
        except Exception:
            continue

    for shape, table in candidates:
        try:
            if len(table.columns) >= _TABLE_COLUMN_COUNT:
                return shape, table
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


def _clear_cell_text(cell) -> None:
    _set_cell_text_preserve_formatting(cell, "")


def _ensure_breakdown_table_data_row_count(table, *, needed_data_rows: int, header_rows: int = _HEADER_ROW_COUNT) -> None:
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
    return f"€{float(value or 0.0):.2f}"


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


def _find_named_utility_shape(slide, shape_name: str):
    target = _normalize_token(shape_name)
    for shape in _iter_shapes_recursive(slide.shapes):
        if _normalize_token(getattr(shape, 'name', '')) == target:
            return shape
    return None


def _copy_shape_to_slide(slide, source_shape):
    newel = deepcopy(source_shape._element)
    slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
    return slide.shapes[-1]


def _table_cell_box(table_shape, table, row_idx: int, col_idx: int):
    left = int(table_shape.left)
    top = int(table_shape.top)
    for idx in range(col_idx):
        left += int(table.columns[idx].width)
    for idx in range(row_idx):
        top += int(table.rows[idx].height)
    width = int(table.columns[col_idx].width)
    height = int(table.rows[row_idx].height)
    return left, top, width, height


def _fit_shape_in_box(
    shape,
    left: int,
    top: int,
    width: int,
    height: int,
    *,
    target_width_ratio: float = 0.52,
    target_height_ratio: float = 0.58,
) -> None:
    target_width = max(1, int(width * target_width_ratio))
    target_height = max(1, int(height * target_height_ratio))

    orig_width = max(1, int(shape.width))
    orig_height = max(1, int(shape.height))
    scale = min(target_width / orig_width, target_height / orig_height)
    new_width = max(1, int(orig_width * scale))
    new_height = max(1, int(orig_height * scale))

    shape.width = new_width
    shape.height = new_height
    shape.left = int(left + ((width - new_width) / 2))
    shape.top = int(top + ((height - new_height) / 2))


def _rounded_roi_cents(value: float) -> int:
    return int(round(float(value or 0.0) * 100))


def _roi_status_symbol_name(row: ChannelBreakdownRow) -> str:
    mat_1_roi_cents = _rounded_roi_cents(row.mat_1_roi)
    mat_roi_cents = _rounded_roi_cents(row.mat_roi)

    if mat_1_roi_cents == 0 and mat_roi_cents > 0:
        return _NEW_SHAPE_NAME

    tolerance_cents = int(round(_ROI_STATUS_EQUALITY_TOLERANCE * 100))
    delta_cents = mat_roi_cents - mat_1_roi_cents
    if abs(delta_cents) <= tolerance_cents:
        return _EQUAL_SHAPE_NAME
    if delta_cents > tolerance_cents:
        return _GREEN_INCREASE_SHAPE_NAME
    return _PINK_DECREASE_SHAPE_NAME


def _apply_roi_status_symbols(
    slide,
    *,
    utility_slide,
    table_shape,
    table,
    rows: list[ChannelBreakdownRow],
) -> None:
    if utility_slide is None:
        logger.info('No <Arrow Utilities> slide found; skipping ROI Status symbols.')
        return

    prototypes = {
        _GREEN_INCREASE_SHAPE_NAME: _find_named_utility_shape(utility_slide, _GREEN_INCREASE_SHAPE_NAME),
        _PINK_DECREASE_SHAPE_NAME: _find_named_utility_shape(utility_slide, _PINK_DECREASE_SHAPE_NAME),
        _NEW_SHAPE_NAME: _find_named_utility_shape(utility_slide, _NEW_SHAPE_NAME),
        _EQUAL_SHAPE_NAME: _find_named_utility_shape(utility_slide, _EQUAL_SHAPE_NAME),
    }

    missing = [name for name, shape in prototypes.items() if shape is None]
    if missing:
        raise ValueError(
            'Could not find ROI Status utility shape(s) on <Arrow Utilities> slide: ' + ', '.join(missing)
        )

    for idx, row in enumerate(rows, start=_HEADER_ROW_COUNT):
        symbol_name = _roi_status_symbol_name(row)
        source_shape = prototypes.get(symbol_name)
        if source_shape is None:
            continue
        _clear_cell_text(table.cell(idx, _ROI_STATUS_COL_IDX))
        clone = _copy_shape_to_slide(slide, source_shape)
        left, top, width, height = _table_cell_box(table_shape, table, idx, _ROI_STATUS_COL_IDX)
        _fit_shape_in_box(clone, left, top, width, height)


def _update_net_revenue_roi_breakdown_table(
    slide,
    *,
    utility_slide,
    brand: str,
    year_range,
    media_template_brand_df: pd.DataFrame,
) -> None:
    table_shape, table = _find_breakdown_table_shape_and_table(slide)
    if table is None or table_shape is None:
        logger.info("No Net Revenue ROI Breakdown table found on slide for brand %r.", brand)
        return

    rows = _compute_brand_channel_breakdown_rows(
        media_template_brand_df,
        brand=brand,
        year_range=year_range,
    )

    _ensure_breakdown_table_data_row_count(table, needed_data_rows=len(rows), header_rows=_HEADER_ROW_COUNT)

    if len(table.rows) >= 1 and len(table.columns) >= 6:
        _set_cell_text_preserve_formatting(table.cell(0, 1), str(year_range.start_year))
        _set_cell_text_preserve_formatting(table.cell(0, 5), str(year_range.end_year))

    for idx, row in enumerate(rows, start=_HEADER_ROW_COUNT):
        _set_cell_text_preserve_formatting(table.cell(idx, 0), row.channel)
        _set_cell_text_preserve_formatting(table.cell(idx, 1), _format_roi(row.mat_1_roi))
        _set_cell_text_preserve_formatting(table.cell(idx, 2), _format_spend_thousands(row.mat_1_spend))
        _set_cell_text_preserve_formatting(table.cell(idx, 3), _format_budget_pct(row.mat_1_budget_pct))
        _clear_cell_text(table.cell(idx, _ROI_STATUS_COL_IDX))
        _set_cell_text_preserve_formatting(table.cell(idx, 5), _format_roi(row.mat_roi))
        _set_cell_text_preserve_formatting(table.cell(idx, 6), _format_spend_thousands(row.mat_spend))
        _set_cell_text_preserve_formatting(table.cell(idx, 7), _format_budget_pct(row.mat_budget_pct))

    _apply_roi_status_symbols(
        slide,
        utility_slide=utility_slide,
        table_shape=table_shape,
        table=table,
        rows=rows,
    )


def _populate_net_revenue_roi_breakdown_slide(
    slide,
    *,
    utility_slide,
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
        utility_slide=utility_slide,
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

    utility_slide = _find_arrow_utilities_slide(prs)

    normalized_media_template_brand_df = _normalize_media_template_brand_df(media_template_brand_df)
    brands = _ordered_unique_brands(normalized_media_template_brand_df)
    if not brands:
        return

    year_range = _resolve_scope_year_range(scope_df)

    slide_mapper = WaterfallSlideMapper()

    if len(brands) == 1:
        _populate_net_revenue_roi_breakdown_slide(
            template_slide,
            utility_slide=utility_slide,
            brand=brands[0],
            year_range=year_range,
            media_template_brand_df=normalized_media_template_brand_df,
        )
        if utility_slide is not None and utility_slide != template_slide:
            slide_mapper.delete_slide(prs, utility_slide)
        return

    template_idx = list(prs.slides).index(template_slide)

    for idx, brand in enumerate(brands):
        slide = slide_mapper.duplicate_slide(prs, template_slide, insert_idx=template_idx + idx + 1)
        _populate_net_revenue_roi_breakdown_slide(
            slide,
            utility_slide=utility_slide,
            brand=brand,
            year_range=year_range,
            media_template_brand_df=normalized_media_template_brand_df,
        )

    slide_mapper.delete_slide(prs, template_slide)
    if utility_slide is not None:
        slide_mapper.delete_slide(prs, utility_slide)
