"""Excel formula parser — tokenizes and builds AST with full reference resolution.

Handles:
  - Absolute refs: $I$250
  - Relative refs: I250
  - Mixed: $I250, I$250
  - Cross-sheet: 参数输入表!$I$250
  - Ranges: $I$250:$I$260, 参数输入表!$A$1:$C$10
  - Nested functions: ROUND(SUM(A1:A10)/10, 2)
  - Operators: +, -, *, /, ^, %, &, =, <>, <, >, <=, >=
  - Literals: numbers, strings ("..."), booleans (TRUE/FALSE)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from models.cell_node import (
    CellRef, FormulaAST, FunctionCall, BinaryOp, UnaryOp, Literal
)

# --- Token types ---
TK_CELL_REF = "CELL_REF"
TK_RANGE = "RANGE"
TK_FUNC = "FUNC"
TK_LPAREN = "LPAREN"
TK_RPAREN = "RPAREN"
TK_COLON = "COLON"
TK_COMMA = "COMMA"
TK_OP = "OP"
TK_NUM = "NUM"
TK_STR = "STR"
TK_BOOL = "BOOL"
TK_EOF = "EOF"


@dataclass
class Token:
    type: str
    value: Any
    pos: int = 0


# Regex for cell reference (with optional sheet prefix, optional $)
_CELL_REF_RE = re.compile(
    r"""
    (?:(?P<sheet>[^!]+)!)??   # optional sheet name followed by !
    \$?(?P<col>[A-Z]+)         # column (optional $)
    \$?(?P<row>\d+)            # row (optional $)
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Pattern to match sheet!cell_ref at current position (for tokenizer lookahead)
# Sheet names may be quoted with single quotes for special chars: 'Sheet-Name'!A1
_SHEET_CELL_RE = re.compile(
    r"""
    (?P<quote>')?                # optional opening quote
    (?P<sheet>(?(quote)[^']+|[^!]+))  # quoted: everything until closing quote; unquoted: until !
    (?(quote)'|)!                # closing quote or just !
    \$?[A-Z]+                    # column (optional $)
    \$?\d+                       # row (optional $)
    (?::\$?[A-Z]+\$\d+)?        # optional range end
    """,
    re.IGNORECASE | re.VERBOSE,
)

_FUNC_RE = re.compile(r"[A-Z_]\w*\(", re.IGNORECASE)
_NUM_RE = re.compile(r"-?\d+(\.\d+)?([eE][+-]?\d+)?")
_STR_RE = re.compile(r'"[^"]*"')
_OP_CHARS = {"+", "-", "*", "/", "^", "=", "<", ">", "&", "%"}


def _is_cell_ref(s: str) -> bool:
    return bool(_CELL_REF_RE.match(s))


def _parse_cell_ref(s: str) -> CellRef:
    """Parse a cell reference string like '参数输入表!$I$250' into CellRef.
    Also handles range suffixes like '参数输入表!$I$250:$I$260'."""
    # Check for range suffix
    is_range = False
    range_end = None
    range_match = re.search(r":(\$?)([A-Za-z]+)(\$?)(\d+)$", s)
    if range_match:
        is_range = True
        range_end = CellRef(
            sheet=None,  # same sheet as start
            row=int(range_match.group(4)),
            col=range_match.group(2).upper(),
            row_abs=range_match.group(3) == "$",
            col_abs=range_match.group(1) == "$",
        )
        s = s[:range_match.start()]  # trim to just the start ref

    m = _CELL_REF_RE.match(s)
    if not m:
        raise ValueError(f"Not a cell ref: {s}")

    sheet = m.group("sheet")
    if sheet:
        sheet = sheet.strip("'")  # Excel quotes sheet names with special chars
    col = m.group("col").upper()
    row_str = m.group("row")

    # Determine absoluteness by checking $ before col/row positions
    col_start = m.start("col")
    row_start = m.start("row")
    col_abs = col_start > 0 and s[col_start - 1] == "$"
    row_abs = row_start > 0 and s[row_start - 1] == "$"

    ref = CellRef(
        sheet=sheet,
        row=int(row_str),
        col=col,
        row_abs=row_abs,
        col_abs=col_abs,
        is_range=is_range,
        range_end=range_end,
    )
    return ref


class Tokenizer:
    """Breaks formula string (without leading '=') into tokens."""

    def __init__(self, text: str):
        self.text = text.strip()
        self.pos = 0
        self.tokens: list[Token] = []
        self._tokenize()

    def _tokenize(self) -> None:
        t = self.text
        i = 0
        while i < len(t):
            # Skip whitespace
            if t[i] == " ":
                i += 1
                continue

            # String literal
            if t[i] == '"':
                j = i + 1
                while j < len(t) and t[j] != '"':
                    j += 1
                self.tokens.append(Token(TK_STR, t[i:j + 1], i))
                i = j + 1
                continue

            # Comma
            if t[i] == ",":
                self.tokens.append(Token(TK_COMMA, ",", i))
                i += 1
                continue

            # Colon (range)
            if t[i] == ":":
                self.tokens.append(Token(TK_COLON, ":", i))
                i += 1
                continue

            # Parentheses
            if t[i] == "(":
                self.tokens.append(Token(TK_LPAREN, "(", i))
                i += 1
                continue
            if t[i] == ")":
                self.tokens.append(Token(TK_RPAREN, ")", i))
                i += 1
                continue

            # Comparison operators: <>, <=, >=
            if t[i] in ("<", ">") and i + 1 < len(t) and t[i + 1] == ">":
                self.tokens.append(Token(TK_OP, t[i:i + 2], i))
                i += 2
                continue
            if t[i] in ("<", ">") and i + 1 < len(t) and t[i + 1] == "=":
                self.tokens.append(Token(TK_OP, t[i:i + 2], i))
                i += 2
                continue

            # Arithmetic operators
            if t[i] in _OP_CHARS and t[i] not in ("<", ">"):
                self.tokens.append(Token(TK_OP, t[i], i))
                i += 1
                continue

            # Boolean
            if t[i:i + 4].upper() == "TRUE":
                self.tokens.append(Token(TK_BOOL, True, i))
                i += 4
                continue
            if t[i:i + 5].upper() == "FALSE":
                self.tokens.append(Token(TK_BOOL, False, i))
                i += 5
                continue

            # Number
            m = _NUM_RE.match(t, i)
            if m:
                val = m.group()
                self.tokens.append(Token(TK_NUM, float(val) if "." in val or "e" in val.lower() else int(val), i))
                i = m.end()
                continue

            # Sheet-prefixed cell ref (e.g., 参数输入表!$I$250 or 参数输入表!$A$1:$C$10)
            # Must match from position i exactly — sheet name can contain Chinese but NOT operators
            sheet_m = _SHEET_CELL_RE.match(t, i)
            if sheet_m:
                segment = sheet_m.group()
                # Verify it doesn't include operator chars (prevents over-matching)
                if not any(op in segment for op in ('+', '-', '*', '/', '^', '&', '=', '<', '>', '%')):
                    # Check if it's a function (followed by '(')
                    j = i + len(segment)
                    if j < len(t) and t[j] == "(":
                        self.tokens.append(Token(TK_FUNC, segment.rstrip("("), i))
                        i = j
                        continue
                    self.tokens.append(Token(TK_CELL_REF, segment, i))
                    i = j
                    continue

            # Function or cell ref — scan forward to find the end
            if t[i].isalpha() or t[i] in "$_":
                j = i
                while j < len(t) and (t[j].isalnum() or t[j] in "$_."):
                    j += 1
                segment = t[i:j]

                # Check for sheet prefix in scanned segment (e.g., 表1!A1)
                if "!" in segment:
                    # Try to match as sheet!cell
                    parts = segment.split("!", 1)
                    ref_part = parts[1]
                    if _is_cell_ref(ref_part) or ":" in ref_part:
                        # Check if followed by (  => function
                        if j < len(t) and t[j] == "(":
                            self.tokens.append(Token(TK_FUNC, parts[0].rstrip("("), i))
                            i = j
                            continue
                        self.tokens.append(Token(TK_CELL_REF, segment, i))
                        i = j
                        continue

                # Check if followed by (  => function
                if j < len(t) and t[j] == "(":
                    self.tokens.append(Token(TK_FUNC, segment.rstrip("("), i))
                    i = j
                    continue

                # Otherwise it's a cell reference
                # Try to parse as cell ref
                if _is_cell_ref(segment):
                    self.tokens.append(Token(TK_CELL_REF, segment, i))
                    i = j
                    continue

                # Could be a named range or sheet reference — treat as cell ref anyway
                self.tokens.append(Token(TK_CELL_REF, segment, i))
                i = j
                continue

            # Unknown char — skip
            i += 1

        self.tokens.append(Token(TK_EOF, None, i))


class FormulaParser:
    """Recursive descent parser for Excel formulas."""

    def __init__(self, formula_raw: str):
        self.raw = formula_raw
        text = formula_raw.lstrip("=").strip()
        self.tokenizer = Tokenizer(text)
        self.tokens = self.tokenizer.tokens
        self.pos = 0
        self.references: list[CellRef] = []

    def parse(self) -> FormulaAST:
        tree = self._expression()
        return FormulaAST(tree=tree, references=self.references, raw=self.raw)

    def _current(self) -> Token:
        if self.pos < len(self.tokens):
            return self.tokens[self.pos]
        return self.tokens[-1]  # EOF

    def _advance(self) -> Token:
        tok = self._current()
        self.pos += 1
        return tok

    def _expect(self, ttype: str) -> Token:
        tok = self._current()
        if tok.type != ttype:
            raise ValueError(f"Expected {ttype}, got {tok.type} ('{tok.value}') at pos {tok.pos}")
        return self._advance()

    # --- Grammar rules (precedence: low to high) ---

    def _expression(self):
        """expr = concat_expr (('&' | '=' | '<>' | '<' | '>' | '<=' | '>=') concat_expr)*"""
        left = self._concat()
        while self._current().type == TK_OP and self._current().value in ("&", "=", "<>", "<", ">", "<=", ">="):
            op = self._advance().value
            right = self._concat()
            left = BinaryOp(op=op, left=left, right=right)
        return left

    def _concat(self):
        """concat = add_expr ('&' add_expr)*"""
        left = self._additive()
        while self._current().type == TK_OP and self._current().value == "&":
            self._advance()
            right = self._additive()
            left = BinaryOp(op="&", left=left, right=right)
        return left

    def _additive(self):
        """additive = multiplicative (('+' | '-') multiplicative)*"""
        left = self._multiplicative()
        while self._current().type == TK_OP and self._current().value in ("+", "-"):
            op = self._advance().value
            right = self._multiplicative()
            left = BinaryOp(op=op, left=left, right=right)
        return left

    def _multiplicative(self):
        """multiplicative = power (('*' | '/') power)*"""
        left = self._power()
        while self._current().type == TK_OP and self._current().value in ("*", "/"):
            op = self._advance().value
            right = self._power()
            left = BinaryOp(op=op, left=left, right=right)
        return left

    def _power(self):
        """power = unary ('^' unary)*"""
        base = self._unary()
        if self._current().type == TK_OP and self._current().value == "^":
            self._advance()
            exp = self._power()  # right-associative
            return BinaryOp(op="^", left=base, right=exp)
        return base

    def _unary(self):
        """unary = ('-' | '+') unary | primary"""
        if self._current().type == TK_OP and self._current().value in ("-", "+"):
            op = self._advance().value
            operand = self._unary()
            return UnaryOp(op=op, operand=operand)
        return self._primary()

    def _primary(self):
        """primary = NUMBER | STRING | BOOL | CELL_REF | RANGE | FUNC | '(' expr ')'"""
        tok = self._current()

        if tok.type == TK_NUM:
            self._advance()
            return Literal(value=tok.value)

        if tok.type == TK_STR:
            self._advance()
            return Literal(value=tok.value.strip('"'))

        if tok.type == TK_BOOL:
            self._advance()
            return Literal(value=tok.value)

        if tok.type == TK_CELL_REF:
            self._advance()
            ref = _parse_cell_ref(tok.value)

            # Check for range: CELL_REF : CELL_REF
            if self._current().type == TK_COLON:
                self._advance()  # skip ':'
                next_tok = self._current()
                if next_tok.type == TK_CELL_REF:
                    self._advance()
                    end_ref = _parse_cell_ref(next_tok.value)
                    ref.is_range = True
                    ref.range_end = end_ref
                else:
                    raise ValueError(f"Expected cell ref after ':', got {next_tok.type}")

            self.references.append(ref)
            return ref

        if tok.type == TK_FUNC:
            return self._function_call()

        if tok.type == TK_LPAREN:
            self._advance()
            expr = self._expression()
            self._expect(TK_RPAREN)
            return expr

        # Handle percentage (postfix)
        if tok.type == TK_OP and tok.value == "%":
            self._advance()
            val = Literal(value=0.01)
            return BinaryOp(op="*", left=Literal(value=1), right=val)  # placeholder

        raise ValueError(f"Unexpected token: {tok.type} '{tok.value}' at pos {tok.pos}")

    def _function_call(self) -> FunctionCall:
        name = self._advance().value  # function name
        self._expect(TK_LPAREN)
        args = []
        if self._current().type != TK_RPAREN:
            args.append(self._expression())
            while self._current().type == TK_COMMA:
                self._advance()
                args.append(self._expression())
        self._expect(TK_RPAREN)
        return FunctionCall(name=name.upper(), args=args)
