"""Integration tests — end-to-end pipeline on real Excel."""

import json
from pathlib import Path

import pytest

EXCEL_PATH = str(Path(__file__).parent.parent.parent /
                 "数字化系统财务模型边界【抽水蓄能】v15(亏损弥补+分红预提税+净资产税+折旧摊销优化）.xlsx")


class TestExcelParsing:
    def test_parse_all_sheets(self):
        from core.parser import parse_workbook_with_values
        cells = parse_workbook_with_values(EXCEL_PATH)

        sheets = {c.sheet for c in cells}
        assert len(sheets) == 14
        assert "参数输入表" in sheets
        assert "表1-资金筹措及还本付息表" in sheets
        assert "表10-资产负债表" in sheets

    def test_cell_count(self):
        from core.parser import parse_workbook_with_values
        cells = parse_workbook_with_values(EXCEL_PATH)
        assert len(cells) > 50000

    def test_formulas_preserved(self):
        from core.parser import parse_workbook_with_values
        cells = parse_workbook_with_values(EXCEL_PATH)
        formulas = [c for c in cells if c.formula_raw]
        assert len(formulas) > 40000


class TestGraphBuild:
    def test_graph_structure(self):
        from core.graph_builder import build_graph
        fg = build_graph(EXCEL_PATH)

        assert fg.graph.number_of_nodes() > 50000
        assert fg.graph.number_of_edges() > 100000
        # Excel may have real circular references (iterative calculation)
        # System should detect and report them, not crash
        if fg.has_circular():
            cycles = fg.get_circular_refs(max_cycles=5)
            assert len(cycles) > 0  # Circular refs detected

    def test_cross_sheet_dependencies(self):
        from core.graph_builder import build_graph
        fg = build_graph(EXCEL_PATH)

        # 时间序列!C4 = 参数输入表!I5
        ts_cell = fg.cells.get("时间序列_4_C")
        if ts_cell and ts_cell.formula_raw:
            deps = fg.get_dependencies(ts_cell.id)
            assert len(deps) > 0


class TestBusinessItems:
    def test_items_detected(self):
        from core.graph_builder import build_graph
        from core.section_detector import detect_business_items

        fg = build_graph(EXCEL_PATH)
        items = detect_business_items(fg)
        assert len(items) > 100

    def test_key_items_found(self):
        from core.graph_builder import build_graph
        from core.section_detector import detect_business_items

        fg = build_graph(EXCEL_PATH)
        items = detect_business_items(fg)
        names = [i.name for i in items]

        # Key financial items should be detected
        found_items = [n for n in names if "总投资" in n or "营业收入" in n]
        assert len(found_items) > 0


class TestQueryResolver:
    def test_resolve_entity(self):
        from core.graph_builder import build_graph
        from core.section_detector import detect_business_items
        from llm.query_resolver import QueryResolver

        fg = build_graph(EXCEL_PATH)
        detect_business_items(fg)
        resolver = QueryResolver(fg)

        result = resolver.resolve("建设期是多少？")
        assert result.entity is not None

    def test_resolve_not_found(self):
        from core.graph_builder import build_graph
        from core.section_detector import detect_business_items
        from llm.query_resolver import QueryResolver

        fg = build_graph(EXCEL_PATH)
        detect_business_items(fg)
        resolver = QueryResolver(fg)

        # Use a query with no recognizable entity
        result = resolver.resolve("abcdefg")
        assert result.entity is None


class TestFormulaEval:
    def test_simple_formula_coverage(self):
        from core.graph_builder import build_graph
        from core.formula_eval import compare_evaluation

        fg = build_graph(EXCEL_PATH)
        stats = compare_evaluation(EXCEL_PATH, fg)

        # At least 80% of formulas should be handled by custom evaluator
        assert stats["custom_pct"] > 80
        assert stats["total_formulas"] > 40000

    def test_fallback_functions(self):
        from core.graph_builder import build_graph
        from core.formula_eval import compare_evaluation

        fg = build_graph(EXCEL_PATH)
        stats = compare_evaluation(EXCEL_PATH, fg)

        # SUMIF, DATEDIF, EDATE, PMT, COUNTIF, IRR, XIRR now handled by custom evaluator
        # Only SUMIF used to need fallback — verify it's now covered
        assert "SUMIF" not in stats["fallback_functions"]


class TestVersionDiff:
    def test_same_file_no_diff(self):
        from core.graph_builder import build_graph
        from storage.version_diff import compare_graphs

        fg = build_graph(EXCEL_PATH)
        result = compare_graphs(fg, fg, "A", "A")

        assert result.added_count == 0
        assert result.removed_count == 0
        assert result.modified_count == 0
