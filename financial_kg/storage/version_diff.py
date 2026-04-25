"""Version comparison — diff two graph versions.

Modes:
1. Same file different versions (upload1 vs upload2)
2. Cross-project comparison (project A vs project B)
3. Change impact analysis (before vs after input change)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from models.graph import FinancialGraph


@dataclass
class NodeDiff:
    """Diff for a single node."""
    cell_id: str
    status: str  # added | removed | modified
    old_value: Any = None
    new_value: Any = None
    old_formula: str | None = None
    new_formula: str | None = None
    sheet: str = ""


@dataclass
class ItemDiff:
    """Diff for a business item."""
    item_name: str
    status: str  # added | removed | modified
    old_value: Any = None
    new_value: Any = None
    unit: str = ""


@dataclass
class ComparisonResult:
    """Full comparison result."""
    version_a: str
    version_b: str
    node_diffs: list[NodeDiff] = field(default_factory=list)
    item_diffs: list[ItemDiff] = field(default_factory=list)
    added_count: int = 0
    removed_count: int = 0
    modified_count: int = 0
    unchanged_count: int = 0
    summary: str = ""

    @property
    def total_diffs(self) -> int:
        return self.added_count + self.removed_count + self.modified_count


def compare_graphs(
    fg_a: FinancialGraph, fg_b: FinancialGraph,
    name_a: str = "Version A", name_b: str = "Version B",
) -> ComparisonResult:
    """Compare two FinancialGraphs.

    Returns diff report with added/removed/modified cells and business items.
    """
    result = ComparisonResult(version_a=name_a, version_b=name_b)

    cells_a = set(fg_a.cells.keys())
    cells_b = set(fg_b.cells.keys())

    # Added cells (in B but not in A)
    for cell_id in cells_b - cells_a:
        cell = fg_b.cells[cell_id]
        result.node_diffs.append(NodeDiff(
            cell_id=cell_id, status="added",
            new_value=cell.value, new_formula=cell.formula_raw,
            sheet=cell.sheet,
        ))
        result.added_count += 1

    # Removed cells (in A but not in B)
    for cell_id in cells_a - cells_b:
        cell = fg_a.cells[cell_id]
        result.node_diffs.append(NodeDiff(
            cell_id=cell_id, status="removed",
            old_value=cell.value, old_formula=cell.formula_raw,
            sheet=cell.sheet,
        ))
        result.removed_count += 1

    # Modified cells (in both but different)
    for cell_id in cells_a & cells_b:
        cell_a = fg_a.cells[cell_id]
        cell_b = fg_b.cells[cell_id]

        if cell_a.value != cell_b.value or cell_a.formula_raw != cell_b.formula_raw:
            result.node_diffs.append(NodeDiff(
                cell_id=cell_id, status="modified",
                old_value=cell_a.value, new_value=cell_b.value,
                old_formula=cell_a.formula_raw, new_formula=cell_b.formula_raw,
                sheet=cell_a.sheet,
            ))
            result.modified_count += 1
        else:
            result.unchanged_count += 1

    # Compare business items
    items_a = {i.name: i for i in fg_a.business_items.values()}
    items_b = {i.name: i for i in fg_b.business_items.values()}

    for name in set(items_a.keys()) | set(items_b.keys()):
        item_a = items_a.get(name)
        item_b = items_b.get(name)

        if item_a and not item_b:
            result.item_diffs.append(ItemDiff(item_name=name, status="removed", unit=item_a.unit or ""))
        elif item_b and not item_a:
            result.item_diffs.append(ItemDiff(item_name=name, status="added", unit=item_b.unit or ""))
        elif item_a and item_b:
            val_a = fg_a.cells.get(item_a.value_cell).value if item_a.value_cell else None
            val_b = fg_b.cells.get(item_b.value_cell).value if item_b.value_cell else None
            if val_a != val_b or item_a.has_time_series != item_b.has_time_series:
                result.item_diffs.append(ItemDiff(
                    item_name=name, status="modified",
                    old_value=val_a, new_value=val_b,
                    unit=item_a.unit or "",
                ))

    result.summary = (
        f"{name_a} vs {name_b}: "
        f"+{result.added_count} / -{result.removed_count} / ~{result.modified_count} "
        f"({result.unchanged_count} unchanged)"
    )

    return result


def compare_versions_by_excel(
    excel_path_a: str, excel_path_b: str,
    name_a: str = "Version A", name_b: str = "Version B",
) -> ComparisonResult:
    """Compare two Excel files directly."""
    from core.graph_builder import build_graph

    fg_a = build_graph(excel_path_a)
    fg_b = build_graph(excel_path_b)

    return compare_graphs(fg_a, fg_b, name_a, name_b)
