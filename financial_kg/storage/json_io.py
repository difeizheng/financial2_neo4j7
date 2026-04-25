"""JSON I/O for graph serialization."""

from __future__ import annotations

import json
from pathlib import Path

from models.graph import FinancialGraph


def save_graph(fg: FinancialGraph, path: str | Path) -> None:
    """Save FinancialGraph to JSON."""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    fg.save_json(path)


def load_graph(path: str | Path) -> FinancialGraph:
    """Load FinancialGraph from JSON."""
    return FinancialGraph.load_json(path)
