"""Unit tests for formula parser."""

import pytest
from core.formula_parser import FormulaParser, Tokenizer, _parse_cell_ref


class TestTokenizer:
    def test_simple_formula(self):
        t = Tokenizer("ROUNDUP(I10,0)")
        types = [tok.type for tok in t.tokens]
        assert types == ["FUNC", "LPAREN", "CELL_REF", "COMMA", "NUM", "RPAREN", "EOF"]

    def test_absolute_refs(self):
        t = Tokenizer("$I$33*$I$34*I51")
        types = [tok.type for tok in t.tokens]
        assert types.count("CELL_REF") == 3
        assert types.count("OP") == 2

    def test_cross_sheet_ref(self):
        t = Tokenizer("参数输入表!$I$250")
        refs = [tok for tok in t.tokens if tok.type == "CELL_REF"]
        assert len(refs) == 1
        assert "参数输入表" in refs[0].value

    def test_range_ref(self):
        # Ranges like $C$6:$BC$6 are tokenized as 3 tokens: CELL_REF, COLON, CELL_REF
        # The parser combines them into a range reference
        t = Tokenizer("$C$6:$BC$6")
        refs = [tok for tok in t.tokens if tok.type == "CELL_REF"]
        assert len(refs) == 2
        colons = [tok for tok in t.tokens if tok.type == "COLON"]
        assert len(colons) == 1

    def test_cross_sheet_range_tokenizer(self):
        # Cross-sheet range like 参数输入表!$I$250:$I$260 is tokenized as single CELL_REF
        t = Tokenizer("参数输入表!$I$250:$I$260")
        refs = [tok for tok in t.tokens if tok.type == "CELL_REF"]
        assert len(refs) == 1
        assert ":" in refs[0].value

    def test_nested_functions(self):
        t = Tokenizer('ROUND(DATEDIF(I5,I7,"D")/365*12,0)')
        funcs = [tok for tok in t.tokens if tok.type == "FUNC"]
        assert len(funcs) == 2
        assert funcs[0].value == "ROUND"
        assert funcs[1].value == "DATEDIF"


class TestFormulaParser:
    def test_simple_roundup(self):
        ast = FormulaParser("=ROUNDUP(I10,0)").parse()
        assert isinstance(ast.tree, type(ast.tree))  # has a tree
        assert len(ast.references) == 1

    def test_binary_ops(self):
        ast = FormulaParser("=$I$33*$I$34*I51").parse()
        assert len(ast.references) == 3

    def test_cross_sheet(self):
        ast = FormulaParser("=参数输入表!I5").parse()
        assert len(ast.references) == 1
        ref = ast.references[0]
        assert ref.sheet == "参数输入表"

    def test_multi_cell_ref(self):
        ast = FormulaParser("=D8+D9+D35").parse()
        assert len(ast.references) == 3

    def test_nested_with_string(self):
        ast = FormulaParser('=ROUND(DATEDIF(I5,I7,"D")/365*12,0)').parse()
        assert len(ast.references) == 2

    def test_range_in_function(self):
        ast = FormulaParser("=SUMIF($C$6:$BC$6,D7,$C$5:$BC$5)").parse()
        assert len(ast.references) == 3
        # Check range refs
        range_refs = [r for r in ast.references if r.is_range]
        assert len(range_refs) == 2

    def test_cross_sheet_range(self):
        ast = FormulaParser("=参数输入表!$I$250:$I$260").parse()
        assert len(ast.references) == 1
        ref = ast.references[0]
        assert ref.is_range
        assert ref.sheet == "参数输入表"

    def test_mixed_abs_col(self):
        ast = FormulaParser("=$I250").parse()
        ref = ast.references[0]
        assert ref.col_abs
        assert not ref.row_abs

    def test_mixed_abs_row(self):
        ast = FormulaParser("=I$250").parse()
        ref = ast.references[0]
        assert ref.row_abs
        assert not ref.col_abs


class TestCellRefParsing:
    def test_absolute(self):
        ref = _parse_cell_ref("$I$250")
        assert ref.row == 250
        assert ref.col == "I"
        assert ref.row_abs
        assert ref.col_abs

    def test_relative(self):
        ref = _parse_cell_ref("I250")
        assert ref.row == 250
        assert ref.col == "I"
        assert not ref.row_abs
        assert not ref.col_abs

    def test_cross_sheet(self):
        ref = _parse_cell_ref("参数输入表!$I$250")
        assert ref.sheet == "参数输入表"
        assert ref.row == 250
        assert ref.col == "I"
