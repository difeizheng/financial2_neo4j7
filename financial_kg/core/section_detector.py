"""Business item detector — identifies financial concepts spanning rows.

Detects patterns like:
  序号 | 名称 | 合计公式 | 单位 | 2023-01 | 2023-02 | ...
  (B)    (C)    (D)       (E)    (F)      (G)

Strategy:
1. Profile each sheet's column usage (text vs number vs formula per column)
2. Identify header rows
3. Group rows into items
4. LLM fallback for ambiguous cases
"""

from __future__ import annotations

import re
from collections import defaultdict

from models.business_item import BusinessItem, ColumnRoles
from models.cell_node import col_to_index, index_to_col
from models.graph import FinancialGraph

# Patterns that indicate a value is NOT a real business item name
_DATE_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}")
_YEAR_ONLY_RE = re.compile(r"^\d{4}$")
_COL_LETTER_RE = re.compile(r"^[A-Z]{1,3}$")  # "A", "AB", etc.
_NUMBER_RE = re.compile(r"^[+-]?[\d,]+\.?\d*$")
# Time-point auto-reference pattern: "2023年3月，自动取时点"
_TIME_POINT_RE = re.compile(r"\d{4}年\d{1,2}月.*自动取时点")
# Sub-entry pattern: "XX第N年", "XX第N期"
_SUB_ENTRY_RE = re.compile(r".+第\d+[年期间次月]+")


def detect_business_items(fg: FinancialGraph) -> list[BusinessItem]:
    """Detect business items in all sheets."""
    items = []

    for sheet_name in sorted({c.sheet for c in fg.cells.values()}):
        sheet_cells = [c for c in fg.cells.values() if c.sheet == sheet_name]

        # Group by row
        rows: dict[int, list] = {}
        for cell in sheet_cells:
            rows.setdefault(cell.row, []).append(cell)

        # Profile columns across the sheet
        col_profile = _profile_columns(rows)

        # Detect item rows
        sheet_items = _detect_item_rows(fg, sheet_name, rows, col_profile)
        items.extend(sheet_items)

    return items


def _profile_columns(rows: dict[int, list]) -> dict[int, dict]:
    """Profile each column: what % text, number, formula, empty."""
    profile: dict[int, dict] = {}

    for row_cells in rows.values():
        for cell in row_cells:
            ci = cell.col_index
            if ci not in profile:
                profile[ci] = {"text": 0, "number": 0, "formula": 0, "date": 0, "empty": 0, "total": 0}

            p = profile[ci]
            p["total"] += 1
            if cell.formula_raw:
                p["formula"] += 1
            elif cell.data_type == "number":
                p["number"] += 1
            elif cell.data_type == "text":
                p["text"] += 1
            elif cell.data_type == "date":
                p["date"] += 1

    return profile


def _detect_item_rows(
    fg: FinancialGraph,
    sheet_name: str,
    rows: dict[int, list],
    col_profile: dict[int, dict],
) -> list[BusinessItem]:
    """Detect rows that form business items.

    Pattern heuristics:
    - 参数输入表: B=类别, C=序号, D=参数, I=数值, J=单位
    - 表1-10: B=序号, C=名称, D=合计, E=单位, F+=时间序列
    - 时间序列: B=名称, C+=时间序列值
    """
    items = []

    # Determine sheet type by patterns
    sheet_type = _classify_sheet(sheet_name, rows)

    if sheet_type == "parameter_input":
        items = _detect_parameter_items(fg, sheet_name, rows)
    elif sheet_type == "financial_table":
        items = _detect_table_items(fg, sheet_name, rows)
    elif sheet_type == "time_series":
        items = _detect_time_series_items(fg, sheet_name, rows)
    else:
        # Generic detection
        items = _detect_generic_items(fg, sheet_name, rows)

    return items


def _classify_sheet(sheet_name: str, rows: dict[int, list]) -> str:
    """Classify sheet type by name and content patterns."""
    if "参数输入" in sheet_name or "param" in sheet_name.lower():
        return "parameter_input"
    if "时间序列" in sheet_name or "time" in sheet_name.lower():
        return "time_series"
    if "表" in sheet_name or sheet_name.startswith("表"):
        return "financial_table"

    # Check row 3 for header pattern
    row3 = rows.get(3, [])
    has_param_header = any(c.value == "参数" for c in row3)
    has_category_header = any(c.value == "类别" for c in row3)
    if has_param_header or has_category_header:
        return "parameter_input"

    return "financial_table"


def _is_noise_name(name: str) -> bool:
    """Return True if name looks like noise (date, column label, number, sub-entry, etc.)."""
    if len(name) < 3:
        return True
    if _DATE_RE.match(name):
        return True
    if _YEAR_ONLY_RE.match(name):
        return True
    if _COL_LETTER_RE.match(name):
        return True
    if _NUMBER_RE.match(name):
        return True
    # Time-point auto-references are data refs, not items
    if _TIME_POINT_RE.search(name):
        return True
    # Sub-entries like "资本金第1年" are breakdown rows, not independent items
    if _SUB_ENTRY_RE.match(name):
        return True
    return False


def _detect_parameter_items(
    fg: FinancialGraph, sheet_name: str, rows: dict[int, list]
) -> list[BusinessItem]:
    """Detect items in 参数输入表 style:
    Col B=类别, C=序号, D=参数, I=数值, J=单位
    """
    items = []

    # Find header row
    for row_num, row_cells in sorted(rows.items()):
        col_vals = {c.col: c.value for c in row_cells}
        if col_vals.get("D") == "参数":
            # This is the header row — data starts next row
            data_start = row_num + 1
            break
    else:
        data_start = 2  # Default

    SKIP_NAMES = {
        "项目", "单位", "合计", "总计", "小计", "参数", "操作类型",
        "类别", "序号", "取值说明", "备注", "数据结构化整理",
        "收入/成本归类", "参数释义/计算公式",
        "数据来源", "填写说明", "参数名称", "参数值",
    }

    # Process data rows
    current_section = None
    for row_num in sorted(rows.keys()):
        if row_num < data_start:
            continue

        row_cells = rows[row_num]
        col_vals = {c.col: c for c in row_cells}

        # Section marker in col B — detect merged-cell style headers
        b_cell = col_vals.get("B")
        if b_cell and b_cell.data_type == "text" and not b_cell.formula_raw:
            val = str(b_cell.value).strip()
            # Check if this is a section title (text in B, no value in I)
            i_cell_b = col_vals.get("I")
            has_other_data = any(
                c.data_type in ("number", "formula") or c.formula_raw
                for c in row_cells if c.col not in ("B", "D", "J")
            )
            if not i_cell_b or (i_cell_b.data_type == "text" and not i_cell_b.formula_raw):
                if not has_other_data:
                    current_section = val
                    continue

        # Item has name in col D and value in col I
        d_cell = col_vals.get("D")
        i_cell = col_vals.get("I")

        if d_cell and d_cell.data_type == "text" and not d_cell.formula_raw:
            name = str(d_cell.value).strip()
            if name in SKIP_NAMES:
                continue
            if _is_noise_name(name):
                continue

            # Require meaningful value in col I (number or formula referencing data)
            if not i_cell:
                continue
            if i_cell.data_type == "text" and not i_cell.formula_raw:
                continue  # Skip text-only rows

            item_id = f"BI_{sheet_name}_{row_num}"
            cell_ids = [c.id for c in row_cells]

            cols = ColumnRoles(
                name_col="D",
                value_col="I",
                unit_col="J",
            )

            item = BusinessItem(
                id=item_id,
                name=name,
                sheet=sheet_name,
                source_rows=[row_num],
                columns=cols,
                cell_ids=cell_ids,
                value_cell=i_cell.id if i_cell else None,
                unit=str(col_vals.get("J", {}).value) if col_vals.get("J") else None,
                section=current_section,
            )
            items.append(item)
            fg.add_business_item(item)
            fg.add_belongs_to(d_cell.id, item_id, role="name")
            if i_cell:
                fg.add_belongs_to(i_cell.id, item_id, role="value")

    return items


def _detect_table_items(
    fg: FinancialGraph, sheet_name: str, rows: dict[int, list]
) -> list[BusinessItem]:
    """Detect items in 表1-10 style:
    Col C=名称, D=合计, E=单位, F-BE=时间序列, B=序号
    """
    items = []

    # Find name column by scanning for text in col C
    name_col = None
    for row_cells in rows.values():
        c_cell = None
        for cell in row_cells:
            if cell.col == "C":
                c_cell = cell
                break
        if c_cell and c_cell.data_type == "text" and not c_cell.formula_raw:
            name_col = "C"
            break

    if not name_col:
        # Fallback: try col B
        name_col = "B"

    for row_num, row_cells in sorted(rows.items()):
        col_vals = {c.col: c for c in row_cells}

        name_cell = col_vals.get(name_col)
        if not name_cell or name_cell.data_type != "text" or name_cell.formula_raw:
            continue

        name = str(name_cell.value).strip()
        if _is_noise_name(name):
            continue

        # Skip rows that look like section headers or subtotals
        lower = name.lower()
        if any(k in name for k in ("合计", "总计", "小计", "一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、")):
            continue

        # Find value/formula columns
        value_col = None
        unit_col = None
        ts_start = None
        ts_end = None
        formula_cols = []
        value_cols = []

        sorted_cols = sorted(row_cells, key=lambda c: c.col_index)
        for cell in sorted_cols:
            if cell.col == name_col:
                continue
            if cell.formula_raw:
                formula_cols.append(cell)
                if value_col is None:
                    value_col = cell.col
            elif cell.data_type == "text" and not cell.formula_raw:
                if len(str(cell.value)) <= 4:
                    unit_col = cell.col
            elif cell.data_type == "number":
                value_cols.append(cell)
                if value_col is None:
                    value_col = cell.col

        # Determine time series range
        if formula_cols:
            # First formula col = total/合计, rest = time series
            if len(formula_cols) > 1:
                ts_start = formula_cols[1].col
                ts_end = formula_cols[-1].col
            elif len(value_cols) > 2:
                ts_start = value_cols[0].col
                ts_end = value_cols[-1].col
        elif value_cols and len(value_cols) > 3:
            ts_start = value_cols[0].col
            ts_end = value_cols[-1].col

        item_id = f"BI_{sheet_name}_{row_num}"
        cell_ids = [c.id for c in row_cells]

        cols = ColumnRoles(
            name_col=name_col,
            value_col=value_col,
            unit_col=unit_col,
            time_series_start=ts_start,
            time_series_end=ts_end,
        )

        value_cell = None
        if value_col:
            vc = col_vals.get(value_col)
            if vc:
                value_cell = vc.id

        item = BusinessItem(
            id=item_id,
            name=name,
            sheet=sheet_name,
            source_rows=[row_num],
            columns=cols,
            cell_ids=cell_ids,
            value_cell=value_cell,
            unit=str(col_vals[unit_col].value) if unit_col and col_vals.get(unit_col) else None,
            has_time_series=ts_start is not None,
            formula_source=formula_cols[0].id if formula_cols else value_cell,
        )
        items.append(item)
        fg.add_business_item(item)
        fg.add_belongs_to(name_cell.id, item_id, role="name")
        if value_cell:
            fg.add_belongs_to(value_cell, item_id, role="value")

    return items


def _detect_time_series_items(
    fg: FinancialGraph, sheet_name: str, rows: dict[int, list]
) -> list[BusinessItem]:
    """Detect items in time series sheets."""
    items = []

    for row_num, row_cells in sorted(rows.items()):
        col_vals = {c.col: c for c in row_cells}

        # Name in col B or C
        name_cell = col_vals.get("B") or col_vals.get("C")
        if not name_cell or name_cell.data_type != "text" or name_cell.formula_raw:
            continue

        name = str(name_cell.value).strip()
        if _is_noise_name(name):
            continue

        # All remaining cells are time series values
        value_cells = [c for c in row_cells if c.col not in ("B", "C") and c.data_type in ("number", "formula")]
        if not value_cells:
            continue

        value_cells.sort(key=lambda c: c.col_index)

        item_id = f"BI_{sheet_name}_{row_num}"
        cell_ids = [c.id for c in row_cells]

        cols = ColumnRoles(
            name_col=name_cell.col,
            value_col=value_cells[0].col if value_cells else None,
            time_series_start=value_cells[0].col if value_cells else None,
            time_series_end=value_cells[-1].col if value_cells else None,
        )

        item = BusinessItem(
            id=item_id,
            name=name,
            sheet=sheet_name,
            source_rows=[row_num],
            columns=cols,
            cell_ids=cell_ids,
            value_cell=value_cells[0].id if value_cells else None,
            has_time_series=True,
        )
        items.append(item)
        fg.add_business_item(item)

    return items


def _detect_generic_items(
    fg: FinancialGraph, sheet_name: str, rows: dict[int, list]
) -> list[BusinessItem]:
    """Fallback: detect rows with text + value pattern."""
    items = []

    for row_num, row_cells in sorted(rows.items()):
        text_cells = [c for c in row_cells if c.data_type == "text" and not c.formula_raw]
        val_cells = [c for c in row_cells if c.data_type == "number" or c.formula_raw]

        if not text_cells or not val_cells:
            continue

        name_cell = text_cells[0]
        name = str(name_cell.value).strip()
        if _is_noise_name(name):
            continue

        item_id = f"BI_{sheet_name}_{row_num}"
        cell_ids = [c.id for c in row_cells]

        val_cells.sort(key=lambda c: c.col_index)

        cols = ColumnRoles(
            name_col=name_cell.col,
            value_col=val_cells[0].col if val_cells else None,
        )

        if len(val_cells) > 3:
            cols.time_series_start = val_cells[0].col
            cols.time_series_end = val_cells[-1].col

        item = BusinessItem(
            id=item_id,
            name=name,
            sheet=sheet_name,
            source_rows=[row_num],
            columns=cols,
            cell_ids=cell_ids,
            value_cell=val_cells[0].id if val_cells else None,
            has_time_series=len(val_cells) > 3,
        )
        items.append(item)
        fg.add_business_item(item)

    return items
