"""Recalculation engine — propagates cell changes through dependency graph.

Workflow:
1. Build reverse dependency map (who depends on me)
2. Mark changed input cell as dirty
3. BFS through reverse deps to find all affected cells
4. Evaluate dirty cells in topological order
5. Track (old_value, new_value) for each change
"""

from __future__ import annotations

import math
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from typing import Any

from models.cell_node import CellNode, CellRef, col_to_index, index_to_col
from models.graph import FinancialGraph


# --- Helper functions (must be before FormulaEvaluator class) ---

def _to_num(val: Any) -> float:
    """Convert value to number for arithmetic."""
    if isinstance(val, bool):
        return 1 if val else 0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        try:
            return float(val)
        except ValueError:
            return 0
    return 0


def _flatten(items) -> list:
    """Flatten nested lists (for range results)."""
    result = []
    for item in items:
        if isinstance(item, list):
            result.extend(_flatten(item))
        else:
            result.append(item)
    return result


def _if_func(condition: Any, true_val: Any, false_val: Any = 0) -> Any:
    return true_val if condition else false_val


def _excel_serial_to_datetime(serial: float) -> datetime:
    """Convert Excel serial date number to datetime.
    Excel epoch: 1899-12-30 (serial 1 = 1899-12-31).
    """
    return datetime(1899, 12, 30) + timedelta(days=serial)


def _to_date(val: Any) -> datetime:
    """Convert any value to datetime."""
    if isinstance(val, datetime):
        return val
    if isinstance(val, (int, float)) and val > 1000:
        # Excel serial date number
        try:
            return _excel_serial_to_datetime(val)
        except (OverflowError, ValueError, OSError):
            return datetime(1900, 1, 1)
    if isinstance(val, str):
        # Try parsing common date formats
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(val, fmt)
            except ValueError:
                continue
    return datetime(1900, 1, 1)


def _datedif(start: Any, end: Any, unit: str) -> float:
    """Excel DATEDIF function."""
    d1 = _to_date(start)
    d2 = _to_date(end)
    unit = unit.upper().strip().strip('"')

    if unit == "D":
        return (d2 - d1).days
    elif unit == "M":
        return (d2.year - d1.year) * 12 + (d2.month - d1.month) - (1 if d2.day < d1.day else 0)
    elif unit == "Y":
        years = d2.year - d1.year
        if (d2.month, d2.day) < (d1.month, d1.day):
            years -= 1
        return years
    elif unit in ("MD", "YM", "YD"):
        # Less common variants
        if unit == "MD":
            # Day difference ignoring months/years
            d1_adj = d1.replace(year=d2.year, month=d2.month)
            if d1_adj > d2:
                d1_adj = d1.replace(year=d2.year - 1, month=d2.month)
            return (d2 - d1_adj).days
        elif unit == "YM":
            return (d2.month - d1.month) % 12 - (1 if d2.day < d1.day else 0)
        elif unit == "YD":
            d1_adj = d1.replace(year=d2.year)
            if d1_adj > d2:
                d1_adj = d1.replace(year=d2.year - 1)
            return (d2 - d1_adj).days
    return 0


def _edate(start: Any, months: int) -> float:
    """Excel EDATE function — returns Excel serial date number."""
    d = _to_date(start)
    m = d.month - 1 + int(months)
    year = d.year + m // 12
    month = m % 12 + 1
    day = min(d.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
                       31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
    new_date = datetime(year, month, day)
    # Convert back to Excel serial
    return (new_date - datetime(1899, 12, 30)).days


def _pmt(rate: float, nper: float, pv: float, fv: float = 0, when_beg: int = 0) -> float:
    """Excel PMT function — payment for a loan."""
    if rate == 0:
        return -(pv + fv) / nper
    temp = (1 + rate) ** nper
    factor = (1 + rate * when_beg)
    return -(rate * (pv * temp + fv)) / ((temp - 1) * factor)


def _countif(values, criteria) -> int:
    """Excel COUNTIF function — simplified."""
    count = 0
    for v in _flatten(values) if isinstance(values, list) else [values]:
        if _countif_match(v, criteria):
            count += 1
    return count


def _countif_match(value, criteria) -> bool:
    """Match a value against a COUNTIF criteria."""
    if isinstance(criteria, str):
        if criteria.startswith(">="):
            return _to_num(value) >= _to_num(criteria[2:])
        elif criteria.startswith("<="):
            return _to_num(value) <= _to_num(criteria[2:])
        elif criteria.startswith("<>"):
            return _to_num(value) != _to_num(criteria[2:])
        elif criteria.startswith(">"):
            return _to_num(value) > _to_num(criteria[1:])
        elif criteria.startswith("<"):
            return _to_num(value) < _to_num(criteria[1:])
        elif criteria.startswith("="):
            return _to_num(value) == _to_num(criteria[1:])
        elif "*" in criteria or "?" in criteria:
            # Wildcard matching — treat as text
            import re
            pattern = criteria.replace("*", ".*").replace("?", ".")
            return bool(re.match(f"^{pattern}$", str(value)))
        else:
            return _to_num(value) == _to_num(criteria)
    return _to_num(value) == _to_num(criteria)


def _sumif(rng, criteria, sum_rng=None):
    """Excel SUMIF function.

    SUMIF(range, criteria, [sum_range])
    """
    values = _flatten(rng) if isinstance(rng, list) else [rng]
    if sum_rng is None:
        sum_values = values
    else:
        sum_values = _flatten(sum_rng) if isinstance(sum_rng, list) else [sum_rng]

    total = 0
    for i, v in enumerate(values):
        if _countif_match(v, criteria):
            sv = sum_values[i] if i < len(sum_values) else v
            total += _to_num(sv)
    return total


def _irr(values, guess=0.1):
    """Excel IRR function — Newton-Raphson iteration.

    values: list of cash flows (first is initial investment, usually negative)
    guess: initial estimate of IRR
    """
    rate = guess
    for _ in range(100):
        npv = 0
        d_npv = 0
        for t, cf in enumerate(values):
            cf = _to_num(cf)
            npv += cf / (1 + rate) ** t
            d_npv -= t * cf / (1 + rate) ** (t + 1)
        if abs(npv) < 1e-10:
            return rate
        if abs(d_npv) < 1e-15:
            break
        rate = rate - npv / d_npv
        # Clamp rate to avoid divergence
        rate = max(-0.9999, min(rate, 10))
    return f"#NUM!"


def _xirr(values, dates, guess=0.1):
    """Excel XIRR function — irregular cash flows with dates."""
    from datetime import datetime as dt

    dt_dates = [_to_date(d) for d in dates]
    values_list = [_to_num(v) for v in values]

    paired = sorted(zip(dt_dates, values_list), key=lambda x: x[0])
    dt_dates = [p[0] for p in paired]
    values_list = [p[1] for p in paired]

    t0 = dt_dates[0]

    def npv(rate):
        total = 0
        for d, v in zip(dt_dates, values_list):
            days = (d - t0).days
            total += v / (1 + rate) ** (days / 365)
        return total

    rate = guess
    for _ in range(100):
        f = npv(rate)
        if abs(f) < 1e-10:
            return rate
        h = 1e-8
        f_prime = (npv(rate + h) - f) / h
        if abs(f_prime) < 1e-15:
            break
        rate = rate - f / f_prime
        rate = max(-0.9999, min(rate, 10))
    return f"#NUM!"


# --- Lookup functions ---

def _match(lookup_val, lookup_range, match_type=0):
    """Excel MATCH function. match_type: 0=exact, 1=less than, -1=greater than."""
    items = _flatten(lookup_range) if isinstance(lookup_range, list) else [lookup_range]
    lookup_val = _to_num(lookup_val) if not isinstance(lookup_val, str) else str(lookup_val).strip()

    if match_type == 0:  # Exact match
        for i, item in enumerate(items):
            if _is_equal(lookup_val, item):
                return i + 1  # 1-based index
        return f"#N/A"

    if match_type == 1:  # Less than (ascending)
        best = -1
        for i, item in enumerate(items):
            if _to_num(item) <= lookup_val:
                best = i + 1
        return best if best >= 0 else f"#N/A"

    if match_type == -1:  # Greater than (descending)
        best = -1
        for i, item in enumerate(items):
            if _to_num(item) >= lookup_val:
                best = i + 1
        return best if best >= 0 else f"#N/A"

    return f"#VALUE!"


def _index(arr, row_num, col_num=1):
    """Excel INDEX function."""
    if isinstance(arr, list) and len(arr) > 0:
        if isinstance(arr[0], list):
            # 2D array
            r = int(row_num) - 1
            c = int(col_num) - 1
            if 0 <= r < len(arr) and 0 <= c < len(arr[r]):
                return arr[r][c]
            return f"#REF!"
        else:
            # 1D array
            idx = int(row_num) - 1
            if 0 <= idx < len(arr):
                return arr[idx]
            return f"#REF!"
    return f"#VALUE!"


def _choose(index_val, *values):
    """Excel CHOOSE function."""
    idx = int(_to_num(index_val))
    if 1 <= idx <= len(values):
        return values[idx - 1]
    return f"#VALUE!"


def _vlookup(lookup_val, table, col_idx, approximate=True):
    """Excel VLOOKUP function."""
    if not isinstance(table, list):
        return f"#VALUE!"

    # Flatten to rows if 2D
    rows = table if isinstance(table[0], list) else [table]

    lookup_val = _to_num(lookup_val) if not isinstance(lookup_val, str) else str(lookup_val).strip()

    if not approximate:  # Exact match
        for row in rows:
            if row and _is_equal(lookup_val, row[0]):
                ci = int(_to_num(col_idx)) - 1
                if 0 <= ci < len(row):
                    return row[ci]
                return f"#REF!"
        return f"#N/A"

    # Approximate match (ascending)
    best_row = None
    for row in rows:
        if row and _to_num(row[0]) <= lookup_val:
            best_row = row
        elif _to_num(row[0]) > lookup_val:
            break
    if best_row:
        ci = int(_to_num(col_idx)) - 1
        if 0 <= ci < len(best_row):
            return best_row[ci]
    return f"#N/A"


def _is_equal(a, b) -> bool:
    """Excel-style equality comparison."""
    if isinstance(a, str) and isinstance(b, str):
        return a.strip().lower() == b.strip().lower()
    return _to_num(a) == _to_num(b)


@dataclass
class ChangeDelta:
    """Record of a single cell change."""
    cell_id: str
    old_value: Any
    new_value: Any


@dataclass
class RecalcResult:
    """Full recalculation result."""
    changed_cells: list[ChangeDelta] = field(default_factory=list)
    unchanged_count: int = 0
    error_cells: list[tuple[str, str]] = field(default_factory=list)  # (cell_id, error_msg)

    @property
    def total_changed(self) -> int:
        return len(self.changed_cells)


class FormulaEvaluator:
    """Evaluates parsed formula AST against a cell lookup."""

    # Supported Excel functions
    FUNCTIONS = {
        "ROUND": lambda x, n: round(x, n),
        "ROUNDUP": lambda x, n: math.ceil(x * 10**n) / 10**n if n >= 0 else round(x, n),
        "ROUNDDOWN": lambda x, n: math.floor(x * 10**n) / 10**n if n >= 0 else round(x, n),
        "SUM": lambda *args: sum(_flatten(args)),
        "ABS": abs,
        "MAX": lambda *args: max(_flatten(args)),
        "MIN": lambda *args: min(_flatten(args)),
        "AVERAGE": lambda *args: sum(_flatten(args)) / len(_flatten(args)),
        "IF": _if_func,
        "AND": lambda *args: all(_flatten(args)),
        "OR": lambda *args: any(_flatten(args)),
        "NOT": lambda x: not x,
        "POWER": pow,
        "SQRT": math.sqrt,
        "MOD": lambda x, y: x % y,
        "INT": lambda x: int(x),
        "LEN": len,
        "CONCATENATE": lambda *args: "".join(str(a) for a in _flatten(args)),
        "SIGN": lambda x: (1 if x > 0 else -1 if x < 0 else 0),
        # Date functions
        "YEAR": lambda d: _to_date(d).year,
        "MONTH": lambda d: _to_date(d).month,
        "DAY": lambda d: _to_date(d).day,
        "DATEDIF": _datedif,
        "EDATE": _edate,
        "DATE": lambda y, m, d: (_excel_serial_to_datetime(
            (datetime(int(y), int(m), int(d)) - datetime(1899, 12, 30)).days
        ) if isinstance(datetime(int(y), int(m), int(d)), datetime) else
         (datetime(int(y), int(m), int(d)) - datetime(1899, 12, 30)).days),
        "ISBLANK": lambda x: x is None or x == "",
        "COUNT": lambda *args: sum(1 for a in _flatten(args) if isinstance(a, (int, float))),
        "COUNTIF": _countif,
        "SUMIF": _sumif,
        "PMT": _pmt,
        "IRR": _irr,
        "XIRR": _xirr,
        # Lookup functions
        "MATCH": _match,
        "INDEX": _index,
        "CHOOSE": _choose,
        "VLOOKUP": _vlookup,
    }

    def __init__(self, cells: dict[str, CellNode]):
        self.cells = cells
        self._cache: dict[str, Any] = {}

    def evaluate(self, cell: CellNode) -> Any:
        """Evaluate a cell's value or formula."""
        if cell.id in self._cache:
            return self._cache[cell.id]

        if not cell.formula_raw or not cell.formula_ast:
            self._cache[cell.id] = cell.value
            return cell.value

        try:
            result = self._eval_node(cell.formula_ast.tree, cell)
            self._cache[cell.id] = result
            return result
        except Exception as e:
            return f"#ERROR: {e}"

    def _eval_node(self, node, formula_cell: CellNode) -> Any:
        from core.formula_parser import BinaryOp, UnaryOp, FunctionCall, Literal

        if node is None:
            return 0
        if isinstance(node, Literal):
            return node.value
        if isinstance(node, CellRef):
            return self._resolve_ref(node, formula_cell)
        if isinstance(node, BinaryOp):
            left = self._eval_node(node.left, formula_cell)
            right = self._eval_node(node.right, formula_cell)
            return self._apply_op(node.op, left, right)
        if isinstance(node, UnaryOp):
            val = self._eval_node(node.operand, formula_cell)
            if node.op == "-":
                return -val
            return val
        if isinstance(node, FunctionCall):
            return self._eval_function(node, formula_cell)
        return 0

    def _resolve_ref(self, ref: CellRef, formula_cell: CellNode) -> Any:
        """Resolve a cell reference to its value."""
        sheet = ref.sheet or formula_cell.sheet
        row = ref.row
        col = ref.col

        if ref.is_range and ref.range_end:
            # Return list of values for range
            values = []
            start_ci = col_to_index(ref.col or "A")
            end_ci = col_to_index(ref.range_end.col or "A")
            for r in range(row, (ref.range_end.row or row) + 1):
                for ci in range(start_ci, end_ci + 1):
                    cid = f"{sheet}_{r}_{index_to_col(ci)}"
                    val = self._get_cell_value(cid)
                    values.append(val)
            return values

        cell_id = f"{sheet}_{row}_{col}"
        return self._get_cell_value(cell_id)

    def _get_cell_value(self, cell_id: str) -> Any:
        """Get value of a cell, evaluating if needed."""
        cell = self.cells.get(cell_id)
        if cell is None:
            return 0
        return self.evaluate(cell)

    def _apply_op(self, op: str, left: Any, right: Any) -> Any:
        # Handle datetime arithmetic: datetime + number = new datetime
        if isinstance(left, datetime) and isinstance(right, (int, float)):
            if op == "+":
                return left + timedelta(days=right)
            if op == "-":
                return left - timedelta(days=right)
        if isinstance(right, datetime) and isinstance(left, (int, float)):
            if op == "+":
                return right + timedelta(days=left)
            if op == "-":
                return (datetime(1899, 12, 30) + timedelta(days=left)) - right
        if op == "+":
            return _to_num(left) + _to_num(right)
        if op == "-":
            return _to_num(left) - _to_num(right)
        if op == "*":
            return _to_num(left) * _to_num(right)
        if op == "/":
            r = _to_num(right)
            if r == 0:
                return "#DIV/0!"
            return _to_num(left) / r
        if op == "^":
            return _to_num(left) ** _to_num(right)
        if op == "&":
            return f"{left}{right}"
        if op == "=":
            return _to_num(left) == _to_num(right)
        if op == "<>":
            return _to_num(left) != _to_num(right)
        if op == "<":
            return _to_num(left) < _to_num(right)
        if op == ">":
            return _to_num(left) > _to_num(right)
        if op == "<=":
            return _to_num(left) <= _to_num(right)
        if op == ">=":
            return _to_num(left) >= _to_num(right)
        return 0

    def _eval_function(self, func: Any, formula_cell: CellNode) -> Any:
        from core.formula_parser import FunctionCall

        args = []
        for arg in func.args:
            val = self._eval_node(arg, formula_cell)
            args.append(val)

        fn = self.FUNCTIONS.get(func.name.upper())
        if fn is None:
            # Unknown function — return error
            return f"#NAME?: {func.name}"

        try:
            return fn(*args)
        except Exception as e:
            return f"#ERROR: {func.name}({e})"


class RecalcEngine:
    """Dependency-based recalculation with change tracking."""

    def __init__(self, fg: FinancialGraph, max_iterations: int = 100, tolerance: float = 1e-6):
        self.graph = fg
        self.cells = fg.cells
        self.evaluator = FormulaEvaluator(self.cells)
        self.max_iterations = max_iterations
        self.tolerance = tolerance

        # Build reverse dependency map: "who depends on me?"
        self._reverse_deps: dict[str, list[str]] = {}
        for src in fg.graph.nodes():
            for tgt in fg.graph.successors(src):
                self._reverse_deps.setdefault(tgt, []).append(src)

        # Detect circular references via strongly connected components
        self._circular_cells: set[str] = set()
        self._scc_groups: list[set[str]] = []
        self._find_circular_groups()

        # Get topological order (ignoring cycles)
        self._topo_order = list(fg.graph.nodes())
        try:
            self._topo_order = list(fg.get_topo_order())
        except Exception:
            # Fall back: process DAG nodes first, then circular groups
            self._topo_order = self._build_fallback_order()

    def recalculate(self, changed_cell_id: str, apply: bool = False) -> RecalcResult:
        """Recalculate starting from a changed cell.

        Args:
            changed_cell_id: The input cell that was modified (must be set to new value before calling).
            apply: If True, write new values back to cells. If False (default), only compute.
        """
        result = RecalcResult()

        # Find all affected cells via reverse deps (BFS)
        dirty = self._find_dirty(changed_cell_id)
        dirty_set = set(dirty)
        self.evaluator._cache = {}

        # Identify which circular groups are affected
        affected_sccs = []
        for group in self._scc_groups:
            if group & dirty_set:
                affected_sccs.append(group)

        # Phase 1: Evaluate DAG nodes (non-circular) in topological order
        for cell_id in self._topo_order:
            if cell_id not in dirty_set:
                continue
            if cell_id in self._circular_cells:
                continue  # Skip circular cells — handled in Phase 2

            cell = self.cells.get(cell_id)
            if cell is None:
                continue

            old_val = cell.value
            new_val = self.evaluator.evaluate(cell)

            if apply:
                cell.value = new_val

            if old_val != new_val:
                result.changed_cells.append(ChangeDelta(
                    cell_id=cell_id, old_value=old_val, new_value=new_val,
                ))

        # Phase 2: Iterate circular groups until convergence
        for scc in affected_sccs:
            scc_dirty = scc & dirty_set
            if not scc_dirty:
                continue
            iteration_result = self._evaluate_scc(scc, apply)
            result.changed_cells.extend(iteration_result.changed_cells)

        result.unchanged_count = len(self.cells) - len(dirty)
        return result

    def _find_circular_groups(self) -> None:
        """Find strongly connected components with >1 node."""
        import networkx as nx
        for component in nx.strongly_connected_components(self.graph.graph):
            if len(component) > 1:
                self._scc_groups.append(set(component))
                self._circular_cells.update(component)

    def _build_fallback_order(self) -> list[str]:
        """Build evaluation order: DAG nodes first (topo), then circular groups."""
        import networkx as nx
        try:
            dag_nodes = list(nx.topological_sort(self.graph.graph))
            return dag_nodes
        except nx.NetworkXUnfeasible:
            # Remove cycle edges temporarily to get partial order
            dag = self.graph.graph.copy()
            for component in nx.strongly_connected_components(dag):
                if len(component) > 1:
                    subgraph = dag.subgraph(component)
                    edges_to_remove = list(subgraph.edges())
                    dag.remove_edges_from(edges_to_remove)
            try:
                return list(nx.topological_sort(dag))
            except Exception:
                return list(self.graph.graph.nodes())

    def _evaluate_scc(self, scc: set[str], apply: bool) -> RecalcResult:
        """Iteratively evaluate a strongly connected component until convergence."""
        result = RecalcResult()

        for iteration in range(self.max_iterations):
            max_delta = 0
            iteration_changes = []

            for cell_id in scc:
                cell = self.cells.get(cell_id)
                if cell is None:
                    continue

                old_val = cell.value
                try:
                    # Clear cache for this cell only
                    self.evaluator._cache.pop(cell_id, None)
                    new_val = self.evaluator.evaluate(cell)
                except Exception:
                    new_val = "#CIRC!"

                if isinstance(old_val, (int, float)) and isinstance(new_val, (int, float)):
                    delta = abs(new_val - old_val)
                    max_delta = max(max_delta, delta)

                if old_val != new_val:
                    iteration_changes.append(ChangeDelta(
                        cell_id=cell_id, old_value=old_val, new_value=new_val,
                    ))
                    if apply:
                        cell.value = new_val

            if max_delta < self.tolerance:
                # Converged — record final changes
                result.changed_cells.extend(iteration_changes)
                return result

        # Did not converge — mark as circular error
        for cell_id in scc:
            cell = self.cells.get(cell_id)
            if cell:
                if apply:
                    cell.value = "#CIRC!"
                result.changed_cells.append(ChangeDelta(
                    cell_id=cell_id, old_value=cell.value, new_value="#CIRC!",
                ))

        return result

    def _find_dirty(self, start_id: str) -> list[str]:
        """BFS through reverse deps to find all affected cells."""
        dirty = set()
        queue = [start_id]
        while queue:
            current = queue.pop(0)
            if current in dirty:
                continue
            dirty.add(current)
            for dependent in self._reverse_deps.get(current, []):
                if dependent not in dirty:
                    queue.append(dependent)
        return list(dirty)
