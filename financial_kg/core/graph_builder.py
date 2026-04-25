"""Graph builder — assembles cells + formula dependencies into FinancialGraph."""

from __future__ import annotations

import re

from models.cell_node import CellNode, CellRef, col_to_index, index_to_col
from models.graph import FinancialGraph
from core.formula_parser import FormulaParser
from core.parser import parse_workbook_with_values


def _resolve_references(
    cell: CellNode,
    sheet_names: list[str],
    all_cell_ids: set[str],
) -> list[tuple[str, str]]:
    """Resolve formula references to actual cell IDs. Returns list of (target_id, fragment)."""
    if not cell.formula_raw:
        return []

    # Try AST parsing first
    try:
        ast = FormulaParser(cell.formula_raw).parse()
        cell.formula_ast = ast
    except Exception:
        return _regex_fallback(cell)

    results = []
    for ref in ast.references:
        sheet = ref.sheet or cell.sheet

        if ref.is_range and ref.range_end:
            start_col_idx = col_to_index(ref.col or "A")
            end_col_idx = col_to_index(ref.range_end.col or "A")
            start_row = ref.row or 1
            end_row = ref.range_end.row or start_row

            for r in range(start_row, end_row + 1):
                for c_idx in range(start_col_idx, end_col_idx + 1):
                    col_letter = index_to_col(c_idx)
                    target_id = f"{sheet}_{r}_{col_letter}"
                    fragment = f"{ref.col or ''}{ref.row or ''}:{ref.range_end.col or ''}{ref.range_end.row or ''}"
                    if target_id in all_cell_ids:
                        results.append((target_id, fragment))
        else:
            if ref.row is not None and ref.col is not None:
                target_id = f"{sheet}_{ref.row}_{ref.col}"
                fragment = str(ref)
                if target_id in all_cell_ids:
                    results.append((target_id, fragment))

    return results


def _regex_fallback(cell: CellNode) -> list[tuple[str, str]]:
    """Regex-based reference extraction as fallback."""
    if not cell.formula_raw:
        return []

    pattern = r"(?:([^\s!]+)!)?(\$?)([A-Za-z]+)(\$?)(\d+)"
    results = []
    for m in re.finditer(pattern, cell.formula_raw):
        sheet = m.group(1) or cell.sheet
        col = m.group(3).upper()
        row = m.group(5)
        target_id = f"{sheet}_{row}_{col}"
        results.append((target_id, m.group(0)))
    return results


def _mark_circular_cells(fg: FinancialGraph) -> set[str]:
    """Detect and mark cells involved in circular references."""
    import networkx as nx

    cycles = set()
    for component in nx.strongly_connected_components(fg.graph):
        if len(component) > 1:
            cycles.update(component)

    for cid in cycles:
        cell = fg.cells.get(cid)
        if cell:
            cell.data_type = "circular"

    return cycles


def build_graph(excel_path: str) -> FinancialGraph:
    """Build complete FinancialGraph from Excel file."""
    cells = parse_workbook_with_values(excel_path)
    fg = FinancialGraph()

    sheet_names = list({c.sheet for c in cells})
    all_cell_ids = {c.id for c in cells}

    # Add all cells
    for cell in cells:
        fg.add_cell(cell)

    # Build dependency edges
    for cell in cells:
        if cell.formula_raw:
            deps = _resolve_references(cell, sheet_names, all_cell_ids)
            for target_id, fragment in deps:
                fg.add_dependency(cell.id, target_id, fragment)

    # Detect and mark circular references
    _mark_circular_cells(fg)

    return fg
