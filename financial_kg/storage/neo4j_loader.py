"""Neo4j integration — load FinancialGraph into Neo4j database.

Usage:
    from storage.neo4j_loader import Neo4jLoader
    loader = Neo4jLoader("bolt://localhost:7687", "neo4j", "password")
    loader.load_graph(fg)
    # Query: loader.run_cypher("MATCH (n:Cell {sheet: '参数输入表'}) RETURN count(n)")
"""

from __future__ import annotations

from typing import Any

from config import Config


class Neo4jLoader:
    """Loads FinancialGraph into Neo4j."""

    def __init__(self, uri: str | None = None,
                 user: str | None = None, password: str | None = None):
        self.uri = uri or Config.NEO4J_URI
        self.user = user or Config.NEO4J_USER
        self.password = password or Config.NEO4J_PASSWORD
        self.driver = None

    def connect(self):
        """Connect to Neo4j."""
        try:
            from neo4j import GraphDatabase
            self.driver = GraphDatabase.driver(
                self.uri, auth=(self.user, self.password)
            )
            self.driver.verify_connectivity()
        except ImportError:
            raise ImportError("Install neo4j: pip install neo4j")
        except Exception as e:
            raise ConnectionError(f"Cannot connect to Neo4j at {self.uri}: {e}")

    def close(self):
        if self.driver:
            self.driver.close()

    def load_graph(self, fg, batch_size: int = 100) -> int:
        """Load FinancialGraph into Neo4j.

        Returns number of operations executed.
        """
        if not self.driver:
            self.connect()

        ops = 0
        with self.driver.session() as session:
            # Clear existing data
            session.run("MATCH (n) DETACH DELETE n")

            # Create constraints
            session.run("CREATE CONSTRAINT cell_id IF NOT EXISTS FOR (c:Cell) REQUIRE c.id IS UNIQUE")
            session.run("CREATE CONSTRAINT item_id IF NOT EXISTS FOR (i:BusinessItem) REQUIRE i.id IS UNIQUE")

            # Batch insert cells
            cells = list(fg.cells.values())
            for i in range(0, len(cells), batch_size):
                batch = cells[i:i + batch_size]
                data = [c.to_dict() for c in batch]
                session.run(
                    """
                    UNWIND $data AS d
                    MERGE (c:Cell {id: d.id})
                    SET c += {
                        sheet: d.sheet, row: d.row, col: d.col,
                        value: d.value, formula_raw: d.formula_raw,
                        data_type: d.data_type, is_header: d.is_header,
                        section: d.section
                    }
                    """,
                    data=data,
                )
                ops += 1

            # Batch insert business items
            items = list(fg.business_items.values())
            if items:
                for i in range(0, len(items), batch_size):
                    batch = items[i:i + batch_size]
                    data = [item.to_dict() for item in batch]
                    session.run(
                        """
                        UNWIND $data AS d
                        MERGE (i:BusinessItem {id: d.id})
                        SET i += {
                            name: d.name, sheet: d.sheet,
                            unit: d.unit, has_time_series: d.has_time_series,
                            section: d.section, semantic_type: d.semantic_type
                        }
                        """,
                        data=data,
                    )
                    ops += 1

            # Create relationships in batches
            # Separate by type for efficient batch insertion
            # Filter out dangling edges (refs to empty cells not in graph)
            cell_ids = set(fg.cells.keys())
            item_ids = set(fg.business_items.keys())
            def _valid_dep(d):
                return d["src"] in cell_ids and d["tgt"] in cell_ids
            def _valid_bi(d):
                return d["src"] in cell_ids and d["tgt"] in item_ids

            dep_edges = []
            belongs_edges = []
            ts_edges = []
            section_edges = []

            for src, tgt, data in fg.graph.edges(data=True):
                rel = data.get("relation", "DEPENDS_ON")
                props = {k: v for k, v in data.items() if k != "relation"}
                entry = {"src": src, "tgt": tgt, "props": props}

                if rel == "DEPENDS_ON" and _valid_dep(entry):
                    dep_edges.append(entry)
                elif rel == "BELONGS_TO" and _valid_bi(entry):
                    belongs_edges.append(entry)
                elif rel == "TIME_SERIES_OF" and _valid_bi(entry):
                    ts_edges.append(entry)
                elif rel == "PART_OF_SECTION":
                    section_edges.append({"src": src, "tgt": tgt.replace("SECTION_", "")})

            # DEPENDS_ON edges (Cell -> Cell) — batch via UNWIND
            for i in range(0, len(dep_edges), batch_size):
                batch = dep_edges[i:i + batch_size]
                # Flatten props into top-level keys for UNWIND compatibility
                flat = []
                for e in batch:
                    d = {"src": e["src"], "tgt": e["tgt"]}
                    for k, v in e["props"].items():
                        d[f"p_{k}"] = v
                    flat.append(d)
                props_keys = set()
                for e in batch:
                    props_keys.update(e["props"].keys())
                set_clause = ", ".join(f"r.{k} = d.p_{k}" for k in props_keys) if props_keys else ""
                session.run(
                    f"""
                    UNWIND $data AS d
                    MATCH (a:Cell {{id: d.src}}), (b:Cell {{id: d.tgt}})
                    CREATE (a)-[r:DEPENDS_ON]->(b)
                    {f"SET {set_clause}" if set_clause else ""}
                    """,
                    data=flat,
                )
                ops += 1

            # BELONGS_TO edges (Cell -> BusinessItem)
            for i in range(0, len(belongs_edges), batch_size):
                batch = belongs_edges[i:i + batch_size]
                flat = []
                for e in batch:
                    d = {"src": e["src"], "tgt": e["tgt"]}
                    for k, v in e["props"].items():
                        d[f"p_{k}"] = v
                    flat.append(d)
                props_keys = set()
                for e in batch:
                    props_keys.update(e["props"].keys())
                set_clause = ", ".join(f"r.{k} = d.p_{k}" for k in props_keys) if props_keys else ""
                session.run(
                    f"""
                    UNWIND $data AS d
                    MATCH (a:Cell {{id: d.src}}), (b:BusinessItem {{id: d.tgt}})
                    CREATE (a)-[r:BELONGS_TO]->(b)
                    {f"SET {set_clause}" if set_clause else ""}
                    """,
                    data=flat,
                )
                ops += 1

            # TIME_SERIES_OF edges
            for i in range(0, len(ts_edges), batch_size):
                batch = ts_edges[i:i + batch_size]
                flat = []
                for e in batch:
                    d = {"src": e["src"], "tgt": e["tgt"]}
                    for k, v in e["props"].items():
                        d[f"p_{k}"] = v
                    flat.append(d)
                props_keys = set()
                for e in batch:
                    props_keys.update(e["props"].keys())
                set_clause = ", ".join(f"r.{k} = d.p_{k}" for k in props_keys) if props_keys else ""
                session.run(
                    f"""
                    UNWIND $data AS d
                    MATCH (a:Cell {{id: d.src}}), (b:BusinessItem {{id: d.tgt}})
                    CREATE (a)-[r:TIME_SERIES_OF]->(b)
                    {f"SET {set_clause}" if set_clause else ""}
                    """,
                    data=flat,
                )
                ops += 1

            # PART_OF_SECTION edges (Cell -> Section)
            for i in range(0, len(section_edges), batch_size):
                batch = section_edges[i:i + batch_size]
                session.run(
                    """
                    UNWIND $data AS d
                    MATCH (a:Cell {id: d.src})
                    MERGE (s:Section {name: d.tgt})
                    CREATE (a)-[:PART_OF_SECTION]->(s)
                    """,
                    data=batch,
                )
                ops += 1

        return ops

    def run_cypher(self, query: str, params: dict | None = None) -> list[dict]:
        """Run a Cypher query and return results."""
        if not self.driver:
            self.connect()

        with self.driver.session() as session:
            result = session.run(query, params or {})
            return [r.data() for r in result]

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, *args):
        self.close()
