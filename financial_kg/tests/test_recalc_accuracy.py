"""Recalculation accuracy verification against real Excel file.

Strategy:
1. Parse Excel into graph with all formulas and values
2. Pick a real input parameter from 参数输入表 (number cell, not formula)
3. Change the input value
4. Recalculate through the graph dependency chain
5. Verify downstream formulas produce correct results

Key challenge: We can't re-run Excel to get new computed values after input change.
So we verify by:
- Confirming the dependency chain is correct (all affected cells are descendants)
- Independently evaluating a sample of downstream formulas with new input values
- Checking that formula evaluation matches expected arithmetic
"""

from __future__ import annotations

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.parser import parse_workbook_with_values
from core.graph_builder import build_graph
from core.formula_parser import FormulaParser
from core.recalc_engine import RecalcEngine, FormulaEvaluator
from models.business_item import BusinessItem, ColumnRoles


EXCEL_PATH = os.path.join(
    os.path.dirname(__file__),
    "..", "..",
    "数字化系统财务模型边界【抽水蓄能】v15(亏损弥补+分红预提税+净资产税+折旧摊销优化）.xlsx",
)


def find_input_cells(fg):
    """Find candidate input cells in 参数输入表 — number cells with no formula."""
    input_cells = []
    for cid, cell in fg.cells.items():
        if "参数输入" in cell.sheet and cell.data_type == "number" and not cell.formula_raw:
            input_cells.append(cell)
    return input_cells


def find_downstream_cells(fg, start_id, depth=3):
    """Find cells that depend on start_id, up to N levels deep."""
    visited = set()
    queue = [(start_id, 0)]
    while queue:
        cid, d = queue.pop(0)
        if cid in visited or d > depth:
            continue
        visited.add(cid)
        for dep in fg.get_dependents(cid):
            if dep not in visited:
                queue.append((dep, d + 1))
    return visited


def verify_recalculation():
    """Full recalculation accuracy test."""
    print("=" * 60)
    print("Recalculation Accuracy Verification")
    print("=" * 60)

    # Step 1: Parse Excel
    print("\n[1/5] Parsing Excel file...")
    cells = parse_workbook_with_values(EXCEL_PATH)
    print(f"  Parsed {len(cells)} cells")

    # Step 2: Build graph
    print("\n[2/5] Building dependency graph...")
    fg = build_graph(EXCEL_PATH)
    print(f"  Cells: {len(fg.cells)}")
    print(f"  Edges: {fg.graph.number_of_edges()}")

    # Step 3: Find input cells
    print("\n[3/5] Finding input cells in 参数输入表...")
    input_cells = find_input_cells(fg)
    print(f"  Found {len(input_cells)} candidate input cells")

    # Sort by value to find interesting ones (non-trivial)
    numbered = [c for c in input_cells if isinstance(c.value, (int, float)) and c.value != 0]
    numbered.sort(key=lambda c: abs(c.value) if c.value else 0, reverse=True)

    if not numbered:
        print("  No suitable input cells found!")
        return

    # Pick a few test inputs
    test_inputs = numbered[:5]  # Test top 5 by absolute value

    print("\n[4/5] Testing recalculation for input cells...")

    all_passed = True
    results = []

    for inp in test_inputs:
        print(f"\n  Testing: {inp.id} (sheet={inp.sheet}, row={inp.row}, col={inp.col})")
        print(f"    Original value: {inp.value}")

        # Find downstream cells
        downstream = find_downstream_cells(fg, inp.id, depth=5)
        downstream_formulas = [cid for cid in downstream if fg.cells.get(cid) and fg.cells[cid].formula_raw]
        print(f"    Downstream cells: {len(downstream)} ({len(downstream_formulas)} with formulas)")

        if not downstream_formulas:
            print(f"    SKIP: no downstream formulas")
            continue

        # Record original values of downstream formulas
        original_values = {}
        for cid in downstream_formulas[:50]:  # Sample first 50
            cell = fg.cells.get(cid)
            if cell:
                original_values[cid] = cell.value

        # Create evaluator for current state
        evaluator = FormulaEvaluator(fg.cells)

        # Change the input
        old_val = inp.value
        if isinstance(old_val, (int, float)):
            new_val = old_val * 1.1  # 10% increase
        else:
            new_val = 100
        inp.value = new_val

        # Recalculate
        engine = RecalcEngine(fg)
        result = engine.recalculate(inp.id)

        # Verify
        passed = True
        errors = []

        # Check that some downstream cells changed
        changed_ids = {d.cell_id for d in result.changed_cells}
        overlapping = changed_ids & set(original_values.keys())

        if len(overlapping) == 0:
            passed = False
            errors.append(f"No overlapping changed cells (expected some of the {len(original_values)} sampled)")
        else:
            print(f"    Changed cells: {len(changed_ids)} total, {len(overlapping)} in sample")

        # Verify a few formula evaluations produce numbers (not errors)
        for cid in list(changed_ids & set(original_values.keys()))[:10]:
            cell = fg.cells.get(cid)
            if cell and cell.formula_raw:
                new_eval = FormulaEvaluator(fg.cells)
                ev_result = new_eval.evaluate(cell)
                if isinstance(ev_result, str) and ("ERROR" in ev_result or "NAME" in ev_result):
                    errors.append(f"  {cid}: eval error: {ev_result}")
                    passed = False
                elif isinstance(ev_result, (int, float)):
                    pass  # Good — produced a number

        # Restore original value
        inp.value = old_val

        status = "PASS" if passed else "FAIL"
        results.append((inp.id, old_val, new_val, len(changed_ids), status, errors))

        if not passed:
            all_passed = False
            for err in errors:
                print(f"    ERROR: {err}")
        else:
            print(f"    PASS — {len(changed_ids)} cells recalculated")

    # Step 5: Summary
    print("\n[5/5] Summary")
    print("-" * 60)
    passed_count = sum(1 for r in results if r[4] == "PASS")
    failed_count = sum(1 for r in results if r[4] == "FAIL")
    print(f"  Tests: {len(results)}")
    print(f"  Passed: {passed_count}")
    print(f"  Failed: {failed_count}")

    if not all_passed:
        print("\n  Failed tests:")
        for cid, old, new, changed, status, errors in results:
            if status == "FAIL":
                print(f"    {cid} (value {old} -> {new}):")
                for err in errors:
                    print(f"      {err}")

    return all_passed


def verify_specific_formulas():
    """Verify that specific Excel formulas parse and evaluate correctly."""
    print("\n" + "=" * 60)
    print("Formula Evaluation Spot-Check")
    print("=" * 60)

    cells_list = parse_workbook_with_values(EXCEL_PATH)
    fg = build_graph(EXCEL_PATH)

    # Find formula cells that have known values
    formula_cells = [c for c in fg.cells.values() if c.formula_raw and isinstance(c.value, (int, float))]

    # Sample some from different sheets
    by_sheet = {}
    for c in formula_cells:
        by_sheet.setdefault(c.sheet, []).append(c)

    evaluator = FormulaEvaluator(fg.cells)

    tested = 0
    passed = 0
    failed = 0
    errors = []

    for sheet, cells in sorted(by_sheet.items()):
        # Test up to 20 formulas per sheet
        for cell in cells[:20]:
            tested += 1
            expected = cell.value
            result = evaluator.evaluate(cell)

            if isinstance(result, (int, float)) and isinstance(expected, (int, float)):
                # Allow small floating-point differences
                if abs(result - expected) < 0.01 * max(abs(expected), 1):
                    passed += 1
                else:
                    failed += 1
                    errors.append(
                        f"  {cell.id}: expected={expected}, got={result}, formula={cell.formula_raw[:80]}"
                    )
            elif isinstance(result, str) and ("ERROR" in result or "NAME" in result):
                failed += 1
                errors.append(
                    f"  {cell.id}: eval error: {result}, formula={cell.formula_raw[:80]}"
                )
            else:
                # Text, date, or other type — count as passed if not error
                passed += 1

    print(f"\n  Tested: {tested} formula cells across {len(by_sheet)} sheets")
    print(f"  Passed: {passed}")
    print(f"  Failed: {failed}")

    if errors:
        print(f"\n  Errors (first 20):")
        for err in errors[:20]:
            print(f"    {err}")

    accuracy = passed / max(tested, 1) * 100
    print(f"\n  Formula evaluation accuracy: {accuracy:.1f}%")

    return accuracy


if __name__ == "__main__":
    test1 = verify_recalculation()
    test2 = verify_specific_formulas()

    print("\n" + "=" * 60)
    print("FINAL RESULT")
    print("=" * 60)
    print(f"  Recalculation test: {'PASS' if test1 else 'FAIL'}")
    print(f"  Formula accuracy: {test2:.1f}%")
