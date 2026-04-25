"""Graph wrapper — NetworkX operations + JSON serialization."""

from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

import networkx as nx

from models.cell_node import CellNode
from models.business_item import BusinessItem


class _JSONEncoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, (datetime, date)):
            return o.isoformat()
        return super().default(o)


class FinancialGraph:
    """NetworkX graph holding cells, business items, and their relationships."""

    def __init__(self):
        self.graph = nx.DiGraph()
        self.cells: dict[str, CellNode] = {}
        self.business_items: dict[str, BusinessItem] = {}

    # --- Node operations ---

    def add_cell(self, cell: CellNode) -> None:
        self.cells[cell.id] = cell
        self.graph.add_node(cell.id, type="cell", data=cell.to_dict())

    def add_business_item(self, item: BusinessItem) -> None:
        self.business_items[item.id] = item
        self.graph.add_node(item.id, type="business_item", data=item.to_dict())

    # --- Edge operations ---

    def add_dependency(self, source_id: str, target_id: str, fragment: str = "") -> None:
        """DEPENDS_ON: source cell formula references target cell."""
        self.graph.add_edge(source_id, target_id, relation="DEPENDS_ON", fragment=fragment)

    def add_belongs_to(self, cell_id: str, item_id: str, role: str = "") -> None:
        """BELONGS_TO: cell belongs to business item."""
        self.graph.add_edge(cell_id, item_id, relation="BELONGS_TO", role=role)

    def add_part_of_section(self, cell_id: str, section_name: str) -> None:
        """PART_OF_SECTION: cell belongs to a section."""
        section_id = f"SECTION_{section_name}"
        if not self.graph.has_node(section_id):
            self.graph.add_node(section_id, type="section", name=section_name)
        self.graph.add_edge(cell_id, section_id, relation="PART_OF_SECTION")

    def add_time_series_of(self, cell_id: str, item_id: str, period: str = "") -> None:
        """TIME_SERIES_OF: time-point cell belongs to business item."""
        self.graph.add_edge(cell_id, item_id, relation="TIME_SERIES_OF", period=period)

    # --- Query ---

    def get_dependencies(self, cell_id: str) -> list[str]:
        """Get cells that this cell's formula depends on."""
        return list(self.graph.successors(cell_id))

    def get_dependents(self, cell_id: str) -> list[str]:
        """Get cells that depend on this cell (reverse lookup)."""
        return list(self.graph.predecessors(cell_id))

    def get_topo_order(self) -> list[str]:
        """Topological sort for recalculation order."""
        return list(nx.topological_sort(self.graph))

    def has_circular(self) -> bool:
        """Check for circular references."""
        try:
            list(nx.topological_sort(self.graph))
            return False
        except nx.NetworkXUnfeasible:
            return True

    def get_circular_refs(self, max_cycles: int = 10) -> list[list[str]]:
        """Find circular reference chains. Excel may use iterative calculation for these."""
        try:
            list(nx.topological_sort(self.graph))
            return []
        except nx.NetworkXUnfeasible:
            cycles = []
            seen = set()
            for cycle in nx.simple_cycles(self.graph):
                key = tuple(sorted(cycle))
                if key not in seen and len(cycle) <= 8:
                    seen.add(key)
                    cycles.append(cycle)
                    if len(cycles) >= max_cycles:
                        break
            return cycles

    def get_cell_ids_with_formulas(self) -> list[str]:
        return [cid for cid, cell in self.cells.items() if cell.formula_raw]

    # --- Serialization ---

    def to_dict(self) -> dict:
        nodes = []
        for nid, data in self.graph.nodes(data=True):
            nodes.append({"id": nid, **data})
        edges = []
        for src, tgt, data in self.graph.edges(data=True):
            edges.append({"source": src, "target": tgt, **data})
        return {"nodes": nodes, "edges": edges}

    def save_json(self, path: str | Path) -> None:
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2, cls=_JSONEncoder)

    @classmethod
    def load_json(cls, path: str | Path) -> "FinancialGraph":
        fg = cls()
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for node in data.get("nodes", []):
            nid = node["id"]
            ntype = node.get("type")
            if ntype == "cell":
                cell = CellNode.from_dict(node.get("data", {}))
                fg.add_cell(cell)
            elif ntype == "business_item":
                item = BusinessItem.from_dict(node.get("data", {}))
                fg.add_business_item(item)
            elif ntype == "section":
                fg.graph.add_node(nid, type="section", name=node.get("name", nid))
        for edge in data.get("edges", []):
            rel = edge.get("relation")
            if rel == "DEPENDS_ON":
                fg.add_dependency(edge["source"], edge["target"], edge.get("fragment", ""))
            elif rel == "BELONGS_TO":
                fg.add_belongs_to(edge["source"], edge["target"], edge.get("role", ""))
            elif rel == "PART_OF_SECTION":
                fg.add_part_of_section(edge["source"], edge.get("target", "").replace("SECTION_", ""))
            elif rel == "TIME_SERIES_OF":
                fg.add_time_series_of(edge["source"], edge["target"], edge.get("period", ""))
        return fg
