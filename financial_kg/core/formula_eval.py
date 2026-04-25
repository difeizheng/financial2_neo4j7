"""Hybrid formula evaluator — custom AST for simple, original values for complex.

Priority:
1. Custom AST evaluator (transparent, dependency-tracked)
   - Binary ops: +, -, *, /, ^, &, comparisons
   - Functions: ROUND, ROUNDUP, ROUNDDOWN, SUM, ABS, MAX, MIN, AVERAGE, IF, AND, OR, NOT, POWER, SQRT, MOD, INT, LEN, CONCATENATE, SIGN, YEAR, MONTH, DAY, ISBLANK, COUNT, DATEDIF, EDATE, PMT, COUNTIF, DATE
2. Original value preservation for complex formulas
   - SUMIF, IRR, XIRR
   - For full recalculation, fall back to `formulas` library's ExcelModel
"""

from __future__ import annotations

import re
from typing import Any

from models.graph import FinancialGraph


# Functions handled by custom AST evaluator
SIMPLE_FUNCTIONS = {
    "ROUND", "ROUNDUP", "ROUNDDOWN", "SUM", "ABS", "MAX", "MIN",
    "AVERAGE", "IF", "AND", "OR", "NOT", "POWER", "SQRT", "MOD",
    "INT", "LEN", "CONCATENATE", "SIGN", "YEAR", "MONTH", "DAY",
    "ISBLANK", "COUNT", "DATEDIF", "EDATE", "PMT", "COUNTIF", "DATE",
    "SUMIF", "IRR", "XIRR",
    "MATCH", "INDEX", "CHOOSE", "VLOOKUP",
}


def is_simple_formula(formula_raw: str) -> bool:
    """Check if a formula can be handled by the custom AST evaluator."""
    if not formula_raw:
        return True

    text = formula_raw.lstrip("=").strip()

    # Find all function names
    func_names = re.findall(r'([A-Z_]\w*)\s*\(', text, re.IGNORECASE)
    for name in func_names:
        if name.upper() not in SIMPLE_FUNCTIONS:
            return False

    return True


def needs_fallback(formula_raw: str) -> bool:
    """Check if formula needs fallback."""
    if not formula_raw:
        return False
    return not is_simple_formula(formula_raw)


def compare_evaluation(
    excel_path: str,
    fg: FinancialGraph,
) -> dict:
    """Analyze formula coverage.

    Returns coverage statistics showing which functions need fallback.
    """
    custom_count = 0
    fallback_count = 0
    fallback_funcs = set()
    fallback_cells = []

    for cell in fg.cells.values():
        if not cell.formula_raw:
            continue

        if is_simple_formula(cell.formula_raw):
            custom_count += 1
        else:
            fallback_count += 1
            # Extract function names that need fallback
            func_names = re.findall(r'([A-Z_]\w*)\s*\(', cell.formula_raw, re.IGNORECASE)
            for name in func_names:
                if name.upper() not in SIMPLE_FUNCTIONS:
                    fallback_funcs.add(name.upper())
            fallback_cells.append(cell.id)

    total = custom_count + fallback_count
    return {
        "total_formulas": total,
        "custom_handled": custom_count,
        "custom_pct": custom_count / total * 100 if total > 0 else 0,
        "needs_fallback": fallback_count,
        "fallback_pct": fallback_count / total * 100 if total > 0 else 0,
        "fallback_functions": sorted(fallback_funcs),
        "fallback_cells": fallback_cells[:50],
    }


def load_excel_model(excel_path: str):
    """Load Excel with formulas library for full recalculation.

    Use this when you need to recalculate complex formulas (SUMIF, IRR, etc.)
    after changing input values.
    """
    from formulas import ExcelModel
    return ExcelModel().load(excel_path)


def recalculate_with_formulas(
    excel_path: str,
    overrides: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Recalculate entire Excel with formulas library.

    Args:
        excel_path: Path to Excel file
        overrides: Dict of {cell_key: value} to override before recalculation.
                   cell_key format: ('SheetName', 'A1')

    Returns:
        Dict mapping cell keys to computed values.
    """
    model = load_excel_model(excel_path)

    # Apply overrides
    if overrides:
        for key, value in overrides.items():
            if isinstance(key, str):
                # Parse "SheetName!A1" format
                if "!" in key:
                    sheet, ref = key.split("!", 1)
                    cell_key = (sheet, ref)
                else:
                    continue
            else:
                cell_key = key
            if cell_key in model.cells:
                model.cells[cell_key].value = value

    # Calculate
    model.calculate()

    # Collect results
    results = {}
    for cell_key, cell in model.cells.items():
        results[cell_key] = cell.value

    return results
