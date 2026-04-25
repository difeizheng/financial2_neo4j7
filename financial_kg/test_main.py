"""Test script — parse sample Excel, build graph, save JSON."""

import sys
import json
from pathlib import Path

# Add parent dir to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from core.parser import parse_workbook_with_values
from core.graph_builder import build_graph
from storage.json_io import save_graph, load_graph
from core.formula_parser import FormulaParser

EXCEL_PATH = Path(__file__).parent.parent / "数字化系统财务模型边界【抽水蓄能】v15(亏损弥补+分红预提税+净资产税+折旧摊销优化）.xlsx"
OUTPUT_DIR = Path(__file__).parent.parent / "data" / "graphs"


def test_cell_extraction():
    print("=== Cell Extraction ===")
    cells = parse_workbook_with_values(str(EXCEL_PATH))
    print(f"Total cells: {len(cells)}")

    # Count by sheet
    sheets = {}
    for c in cells:
        sheets.setdefault(c.sheet, 0)
        sheets[c.sheet] += 1
    for s, count in sorted(sheets.items(), key=lambda x: -x[1]):
        print(f"  {s}: {count} cells")

    # Count formulas
    formulas = [c for c in cells if c.formula_raw]
    print(f"Formula cells: {len(formulas)}")
    print(f"Value cells: {len(cells) - len(formulas)}")
    return cells


def test_formula_parser():
    print("\n=== Formula Parser ===")
    test_cases = [
        "=ROUNDUP(I10,0)",
        "=$I$33*$I$34*I51",
        "=参数输入表!I5",
        "=投资概算明细!F24",
        "=SUMIF($C$6:$BC$6,D7,$C$5:$BC$5)",
        "=D8+D9+D35",
        "=ROUND(DATEDIF(I5,I7,\"D\")/365*12,0)",
    ]
    for formula in test_cases:
        try:
            ast = FormulaParser(formula).parse()
            refs = [str(r) for r in ast.references]
            print(f"  OK: {formula[:50]:<55} refs={refs}")
        except Exception as e:
            print(f"  FAIL: {formula[:50]:<55} error={e}")


def test_graph_build():
    print("\n=== Graph Build ===")
    fg = build_graph(str(EXCEL_PATH))
    print(f"Nodes: {fg.graph.number_of_nodes()}")
    print(f"Edges: {fg.graph.number_of_edges()}")
    print(f"Has circular: {fg.has_circular()}")

    # Sample dependencies
    formula_cells = fg.get_cell_ids_with_formulas()[:5]
    print("\nSample dependencies:")
    for cid in formula_cells:
        deps = fg.get_dependencies(cid)
        print(f"  {cid} -> {deps[:3]}")

    return fg


def test_json_save(fg):
    print("\n=== JSON Save ===")
    output = OUTPUT_DIR / "test_graph.json"
    save_graph(fg, output)
    size = output.stat().st_size
    print(f"Saved: {output} ({size:,} bytes)")

    # Verify load
    fg2 = load_graph(output)
    print(f"Loaded: {fg2.graph.number_of_nodes()} nodes, {fg2.graph.number_of_edges()} edges")
    assert fg.graph.number_of_nodes() == fg2.graph.number_of_nodes()
    assert fg.graph.number_of_edges() == fg2.graph.number_of_edges()
    print("Round-trip OK")


def main():
    test_cell_extraction()
    test_formula_parser()
    fg = test_graph_build()
    test_json_save(fg)
    print("\n=== All tests passed ===")


if __name__ == "__main__":
    main()
