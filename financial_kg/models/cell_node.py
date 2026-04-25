"""Cell Node data model — represents one non-empty Excel cell."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


def col_to_index(col: str) -> int:
    """Convert column letter to 1-based index. A=1, Z=26, AA=27."""
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def index_to_col(index: int) -> str:
    """Convert 1-based column index to letter. 1=A, 26=Z, 27=AA."""
    chars = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        chars.append(chr(65 + remainder))
    return "".join(reversed(chars))


@dataclass
class CellRef:
    """Reference to one or more cells (supports ranges)."""
    sheet: str | None = None  # None = same sheet as formula
    row: int | None = None
    col: str | None = None
    row_abs: bool = False
    col_abs: bool = False
    is_range: bool = False
    range_end: "CellRef | None" = None

    def resolve(self, formula_sheet: str, formula_row: int, formula_col: str) -> "CellRef":
        """Resolve relative references to absolute coordinates."""
        sheet = self.sheet or formula_sheet
        row = self.row
        col = self.col

        if not self.row_abs and self.row is not None and formula_row is not None:
            row = self.row  # relative offset already computed by parser

        if not self.col_abs and self.col is not None and formula_col is not None:
            col = self.col  # relative offset already computed by parser

        return CellRef(
            sheet=sheet, row=row, col=col,
            row_abs=True, col_abs=True,
            is_range=self.is_range,
            range_end=self.range_end.resolve(formula_sheet, formula_row, formula_col) if self.range_end else None,
        )

    @property
    def cell_id(self) -> str:
        if self.is_range and self.range_end:
            return f"{self.sheet or ''}_{self.row}_{self.col}:{self.range_end.row}_{self.range_end.col}"
        return f"{self.sheet}_{self.row}_{self.col}" if self.sheet else f"{self.row}_{self.col}"

    def __str__(self) -> str:
        prefix = f"{self.sheet}!" if self.sheet else ""
        row_str = str(self.row) if self.row is not None else "?"
        col_str = self.col or "?"
        ref = f"${col_str}${row_str}" if self.col_abs and self.row_abs else f"{col_str}{row_str}"
        if self.is_range and self.range_end:
            return f"{prefix}{ref}:{self.range_end}"
        return f"{prefix}{ref}"


@dataclass
class FunctionCall:
    """Parsed Excel function call."""
    name: str
    args: list[Any]  # ExprNode | Literal | CellRef


@dataclass
class BinaryOp:
    """Binary operation in formula."""
    op: str  # +, -, *, /, ^, &
    left: Any  # ExprNode
    right: Any  # ExprNode


@dataclass
class UnaryOp:
    """Unary operation in formula."""
    op: str  # +, -
    operand: Any  # ExprNode


@dataclass
class Literal:
    """Literal value in formula."""
    value: int | float | str | bool


@dataclass
class FormulaAST:
    """Parsed formula with references."""
    tree: FunctionCall | BinaryOp | UnaryOp | Literal | CellRef | None = None
    references: list[CellRef] = field(default_factory=list)
    raw: str = ""


@dataclass
class CellNode:
    """One non-empty cell in the workbook."""
    id: str
    sheet: str
    row: int
    col: str
    value: str | int | float | None = None
    formula_raw: str | None = None
    formula_ast: FormulaAST | None = None
    data_type: str = "text"  # number|text|date|formula|error|boolean
    is_header: bool = False
    is_merged: bool = False
    merge_range: str | None = None
    section: str | None = None
    format_code: str | None = None

    @property
    def col_index(self) -> int:
        return col_to_index(self.col)

    @classmethod
    def from_cell_id(cls, cell_id: str) -> "CellNode":
        """Parse cell_id like '参数输入表_5_I' into CellNode."""
        parts = cell_id.rsplit("_", 2)
        if len(parts) != 3:
            # Fallback: last part might be multi-char column
            match = re.match(r"(.+?)_(\d+)_([A-Z]+)$", cell_id, re.IGNORECASE)
            if match:
                sheet, row, col = match.groups()
            else:
                raise ValueError(f"Invalid cell_id: {cell_id}")
        else:
            sheet, row, col = parts
        return cls(
            id=cell_id,
            sheet=sheet,
            row=int(row),
            col=col.upper(),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "sheet": self.sheet,
            "row": self.row,
            "col": self.col,
            "col_index": self.col_index,
            "value": self.value,
            "formula_raw": self.formula_raw,
            "formula_ast": _ast_to_dict(self.formula_ast) if self.formula_ast else None,
            "data_type": self.data_type,
            "is_header": self.is_header,
            "is_merged": self.is_merged,
            "merge_range": self.merge_range,
            "section": self.section,
            "format_code": self.format_code,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "CellNode":
        ast = _dict_to_ast(d.get("formula_ast")) if d.get("formula_ast") else None
        return cls(
            id=d["id"],
            sheet=d["sheet"],
            row=d["row"],
            col=d["col"],
            value=d.get("value"),
            formula_raw=d.get("formula_raw"),
            formula_ast=ast,
            data_type=d.get("data_type", "text"),
            is_header=d.get("is_header", False),
            is_merged=d.get("is_merged", False),
            merge_range=d.get("merge_range"),
            section=d.get("section"),
            format_code=d.get("format_code"),
        )


def _ast_to_dict(ast: FormulaAST | None) -> dict | None:
    if ast is None:
        return None
    return {
        "tree": _node_to_dict(ast.tree),
        "references": [str(r) for r in ast.references],
        "raw": ast.raw,
    }


def _node_to_dict(node) -> dict | None:
    if node is None:
        return None
    if isinstance(node, CellRef):
        return {"type": "CellRef", "sheet": node.sheet, "row": node.row,
                "col": node.col, "row_abs": node.row_abs, "col_abs": node.col_abs,
                "is_range": node.is_range,
                "range_end": _node_to_dict(node.range_end) if node.range_end else None}
    if isinstance(node, FunctionCall):
        return {"type": "FunctionCall", "name": node.name,
                "args": [_node_to_dict(a) for a in node.args]}
    if isinstance(node, BinaryOp):
        return {"type": "BinaryOp", "op": node.op,
                "left": _node_to_dict(node.left), "right": _node_to_dict(node.right)}
    if isinstance(node, UnaryOp):
        return {"type": "UnaryOp", "op": node.op,
                "operand": _node_to_dict(node.operand)}
    if isinstance(node, Literal):
        return {"type": "Literal", "value": node.value}
    return {"type": "unknown", "value": str(node)}


def _dict_to_ast(d: dict) -> FormulaAST:
    ast = FormulaAST(tree=None, references=[], raw=d.get("raw", ""))
    tree = d.get("tree")
    if tree:
        ast.tree = _dict_to_node(tree)
    return ast


def _dict_to_node(d: dict):
    t = d.get("type")
    if t == "CellRef":
        ref = CellRef(sheet=d.get("sheet"), row=d.get("row"), col=d.get("col"),
                      row_abs=d.get("row_abs", False), col_abs=d.get("col_abs", False),
                      is_range=d.get("is_range", False))
        if d.get("range_end"):
            ref.range_end = _dict_to_node(d["range_end"])
        return ref
    if t == "FunctionCall":
        return FunctionCall(name=d["name"], args=[_dict_to_node(a) for a in d.get("args", [])])
    if t == "BinaryOp":
        return BinaryOp(op=d["op"], left=_dict_to_node(d["left"]), right=_dict_to_node(d["right"]))
    if t == "UnaryOp":
        return UnaryOp(op=d["op"], operand=_dict_to_node(d["operand"]))
    if t == "Literal":
        return Literal(value=d["value"])
    return None
