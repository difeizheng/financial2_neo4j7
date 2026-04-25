"""Business Item data model — logical grouping of cells forming a financial concept."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ColumnRoles:
    """Column role mapping for a business item row."""
    name_col: str | None = None
    value_col: str | None = None
    unit_col: str | None = None
    time_series_start: str | None = None
    time_series_end: str | None = None
    extra: dict[str, str] = field(default_factory=dict)


@dataclass
class BusinessItem:
    """Logical financial concept spanning one or more cells."""
    id: str
    name: str
    sheet: str
    source_rows: list[int] = field(default_factory=list)
    columns: ColumnRoles = field(default_factory=ColumnRoles)
    cell_ids: list[str] = field(default_factory=list)
    value_cell: str | None = None
    unit: str | None = None
    has_time_series: bool = False
    formula_source: str | None = None  # cell_id where the main formula lives
    section: str | None = None
    semantic_type: str | None = None  # investment_total|cost|revenue|tax|...

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "sheet": self.sheet,
            "source_rows": self.source_rows,
            "columns": {
                "name_col": self.columns.name_col,
                "value_col": self.columns.value_col,
                "unit_col": self.columns.unit_col,
                "time_series_start": self.columns.time_series_start,
                "time_series_end": self.columns.time_series_end,
                **self.columns.extra,
            },
            "cell_ids": self.cell_ids,
            "value_cell": self.value_cell,
            "unit": self.unit,
            "has_time_series": self.has_time_series,
            "formula_source": self.formula_source,
            "section": self.section,
            "semantic_type": self.semantic_type,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "BusinessItem":
        cols_data = d.get("columns", {})
        cols = ColumnRoles(
            name_col=cols_data.get("name_col"),
            value_col=cols_data.get("value_col"),
            unit_col=cols_data.get("unit_col"),
            time_series_start=cols_data.get("time_series_start"),
            time_series_end=cols_data.get("time_series_end"),
        )
        return cls(
            id=d["id"],
            name=d["name"],
            sheet=d["sheet"],
            source_rows=d.get("source_rows", []),
            columns=cols,
            cell_ids=d.get("cell_ids", []),
            value_cell=d.get("value_cell"),
            unit=d.get("unit"),
            has_time_series=d.get("has_time_series", False),
            formula_source=d.get("formula_source"),
            section=d.get("section"),
            semantic_type=d.get("semantic_type"),
        )
