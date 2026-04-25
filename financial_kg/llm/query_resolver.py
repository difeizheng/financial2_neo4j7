"""LLM query resolver — natural language → graph query.

Uses Bailian (DashScope) API for entity resolution when rule-based match fails.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Any

from config import Config
from models.graph import FinancialGraph
from models.business_item import BusinessItem


@dataclass
class QueryIntent:
    entity_name: str | None = None
    entity_id: str | None = None
    time_period: str | None = None
    metric: str | None = None
    compare_with: str | None = None
    change_input: dict[str, Any] | None = None


@dataclass
class QueryResult:
    entity: BusinessItem | None = None
    value: Any = None
    time_series: dict[str, Any] | None = None
    explanation: str = ""
    related_items: list[str] | None = None
    # Comparison support
    compare_entity: BusinessItem | None = None
    compare_value: Any = None
    compare_time_series: dict[str, Any] | None = None


# Common Chinese financial term aliases
FINANCIAL_ALIASES = {
    "总投资": ["总投资", "投资"],
    "动态投资": ["动态总投资", "动态投"],
    "静态投资": ["静态总投资", "静态投"],
    "营业收入": ["营业收入", "营收入", "收入"],
    "营业成本": ["营业成本", "营成本", "成本"],
    "净利润": ["净利润", "净利", "利润"],
    "所得税": ["所得税", "税费", "税"],
    "建设期": ["建设期"],
    "运营期": ["运营期", "经营期"],
    "资本金": ["资本金", "资本"],
    "贷款": ["贷款", "借款", "融资"],
    "现金流": ["现金流", "现金"],
    "资产负债": ["资产负债", "资产", "负债"],
    "还本付息": ["还本付息", "还本", "付息"],
}


class QueryResolver:
    def __init__(self, fg: FinancialGraph):
        self.fg = fg
        self._item_index: dict[str, BusinessItem] = {}
        self._item_names: list[str] = []
        self._build_index()

    def _build_index(self) -> None:
        for item in self.fg.business_items.values():
            self._item_index[item.name] = item
            self._item_names.append(item.name)
            for keyword in self._extract_keywords(item.name):
                if keyword not in self._item_index:
                    self._item_index[keyword] = item

    def _extract_keywords(self, name: str) -> list[str]:
        clean = re.sub(r"[（）()（）]", "", name)
        parts = re.split(r"[、，,·]", clean)
        return [p.strip() for p in parts if len(p.strip()) > 1]

    def resolve(self, query: str) -> QueryResult:
        intent = self._parse_query(query)

        if intent.change_input:
            return self._handle_what_if(intent)

        if intent.compare_with:
            return self._compare_entities(intent)

        if intent.entity_name:
            return self._lookup_entity(intent)

        return QueryResult(explanation="无法理解查询，请重试。例如: 建设期是多少？2030年营业收入是多少？")

    def _parse_query(self, query: str) -> QueryIntent:
        intent = QueryIntent()

        # Pattern 1: "X是多少/的值/的数值/情况"
        name_patterns = [
            r"(.+?)是多少", r"(.+?)的数值", r"(.+?)的值",
            r"(.+?)的数据", r"(.+?)情况", r"(.+?)为多少",
        ]
        for pattern in name_patterns:
            m = re.search(pattern, query)
            if m:
                entity_text = m.group(1).strip()
                self._extract_time_and_entity(entity_text, intent)
                return intent

        # Pattern 2: what-if
        if "如果" in query and ("会" in query or "影响" in query):
            return self._parse_what_if(query)

        # Pattern 3: comparison
        if ("和" in query or "与" in query) and ("对比" in query or "区别" in query or "比较" in query):
            parts = re.split(r"和|与", query)
            if len(parts) >= 2:
                intent.entity_name = parts[0].strip()
                intent.compare_with = parts[1].strip()
                return intent

        # Pattern 4: Direct entity + year extraction
        self._extract_time_and_entity(query, intent)
        if intent.entity_name:
            return intent

        # Pattern 5: Exact/prefix match
        for name, item in self._item_index.items():
            if name in query:
                intent.entity_name = name
                return intent

        # Pattern 6: Fuzzy match
        best = self._fuzzy_match_best(query)
        if best:
            intent.entity_name = best

        # Pattern 7: LLM fallback
        if not intent.entity_name and Config.has_llm():
            llm_result = self._llm_resolve(query)
            if llm_result:
                intent.entity_name = llm_result

        return intent

    def _extract_time_and_entity(self, text: str, intent: QueryIntent) -> None:
        """Extract year and entity from text like '2030年动态总投资'."""
        time_match = re.search(r"(\d{4})[年度]", text)
        if time_match:
            intent.time_period = time_match.group(1)
            text = re.sub(r"\d{4}[年度]", "", text).strip()

        q_match = re.search(r"(\d{4})[年]([一二三四1-4])季度?", text)
        if q_match:
            intent.time_period = f"{q_match.group(1)}Q{q_match.group(2)}"
            text = re.sub(r"\d{4}[年][一二三四1-4]季度?", "", text).strip()

        if text:
            intent.entity_name = text

    def _fuzzy_match_best(self, query: str) -> str | None:
        """Best fuzzy match using SequenceMatcher."""
        clean_query = re.sub(r"是多少|的值|的数值|的数据|情况|为多少|请|问|查询", "", query).strip()
        if not clean_query:
            return None

        best_score = 0
        best_name = None

        for name in self._item_names:
            score = SequenceMatcher(None, clean_query, name).ratio()
            if clean_query in name:
                score = max(score, len(clean_query) / len(name))
            for alias, keywords in FINANCIAL_ALIASES.items():
                for kw in keywords:
                    if kw in clean_query and kw in name:
                        score = max(score, 0.7)
            if score > best_score:
                best_score = score
                best_name = name

        if best_score > 0.4 and best_name:
            return best_name
        return None

    def _llm_resolve(self, query: str) -> str | None:
        """Use Bailian API to resolve entity with structured output."""
        sheet_items: dict[str, list[str]] = {}
        for item in self.fg.business_items.values():
            sheet_items.setdefault(item.sheet, []).append(item.name)

        sample = []
        for sheet, names in sorted(sheet_items.items()):
            sample.extend(names[:10])
            if len(sample) >= 150:
                break

        names_str = "\n".join(f"- {n}" for n in sample[:150])

        prompt = f"""你是一个财务模型查询助手。根据用户问题，从以下财务指标列表中找出最匹配的一个。

可用指标：
{names_str}

用户问题：{query}

只返回最匹配的指标名称，不要其他内容。如果找不到，返回"无匹配"。"""

        try:
            import dashscope
            from dashscope import Generation

            dashscope.api_key = Config.DASHSCOPE_API_KEY
            response = Generation.call(
                model=Config.DASHSCOPE_MODEL,
                prompt=prompt,
                max_tokens=50,
            )

            if response.status_code == 200:
                result = response.output.text.strip().strip('"').strip("'")
                if result != "无匹配":
                    if result in self._item_index:
                        return result
                    best = self._fuzzy_match_best(result)
                    if best:
                        return best
        except Exception:
            pass

        return None

    def _lookup_entity(self, intent: QueryIntent) -> QueryResult:
        item = self._item_index.get(intent.entity_name)

        if not item:
            item = self._fuzzy_match_entity(intent.entity_name)

        if not item:
            return QueryResult(explanation=f"未找到 '{intent.entity_name}'")

        result = QueryResult(
            entity=item,
            explanation=f"找到: {item.name}",
        )

        value_cell = self.fg.cells.get(item.value_cell) if item.value_cell else None
        if value_cell:
            result.value = value_cell.value

        if intent.time_period:
            ts_data = self._get_time_series_for_year(item, intent.time_period)
            if ts_data:
                result.time_series = ts_data
                result.explanation += f"\n{intent.time_period} 相关数据: {len(ts_data)} 个值"

        return result

    def _fuzzy_match_entity(self, name: str) -> BusinessItem | None:
        """Find best matching BusinessItem using fuzzy matching."""
        clean = re.sub(r"[是|多少|的|值|数据|情况|问|请|查询]", "", name).strip()
        if not clean:
            return None

        best_score = 0
        best_item = None

        for item_name, item in self._item_index.items():
            score = SequenceMatcher(None, clean, item_name).ratio()
            if clean in item_name or item_name in clean:
                score = max(score, len(clean) / max(len(item_name), len(clean)) * 0.8)
            for alias, keywords in FINANCIAL_ALIASES.items():
                for kw in keywords:
                    if kw in clean and kw in item_name:
                        score = max(score, 0.7)
            if score > best_score:
                best_score = score
                best_item = item

        if best_score > 0.35:
            return best_item
        return None

    def _get_time_series_for_year(self, item: BusinessItem, year: str) -> dict[str, Any]:
        if not item.columns.time_series_start:
            return {}

        from models.cell_node import col_to_index
        start_ci = col_to_index(item.columns.time_series_start)
        end_ci = col_to_index(item.columns.time_series_end) if item.columns.time_series_end else start_ci

        result = {}
        for cid in item.cell_ids:
            cell = self.fg.cells.get(cid)
            if cell and start_ci <= cell.col_index <= end_ci:
                label = f"{cell.col}{cell.row}"
                result[label] = cell.value

        return result

    def _compare_entities(self, intent: QueryIntent) -> QueryResult:
        """Handle comparison queries like 'A和B的区别/对比'."""
        item_a = self._item_index.get(intent.entity_name) or self._fuzzy_match_entity(intent.entity_name)
        item_b = self._item_index.get(intent.compare_with) or self._fuzzy_match_entity(intent.compare_with)

        if not item_a and not item_b:
            return QueryResult(explanation=f"未找到 '{intent.entity_name}' 和 '{intent.compare_with}'")
        if not item_a:
            return QueryResult(explanation=f"未找到 '{intent.entity_name}'")
        if not item_b:
            return QueryResult(explanation=f"未找到 '{intent.compare_with}'")

        result = QueryResult(
            entity=item_a,
            compare_entity=item_b,
            explanation=f"对比: {item_a.name} vs {item_b.name}",
        )

        val_a = self.fg.cells.get(item_a.value_cell) if item_a.value_cell else None
        val_b = self.fg.cells.get(item_b.value_cell) if item_b.value_cell else None
        result.value = val_a.value if val_a else None
        result.compare_value = val_b.value if val_b else None

        return result

    def _parse_what_if(self, query: str) -> QueryIntent:
        intent = QueryIntent()
        m = re.search(r"如果(.+?)(变成|变为|改为|增加|减少)(.+?)[，,]", query)
        if m:
            intent.change_input = {
                "entity": m.group(1).strip(),
                "action": m.group(2).strip(),
                "value": m.group(3).strip(),
            }
        return intent

    def _handle_what_if(self, intent: QueryIntent) -> QueryResult:
        from core.recalc_engine import RecalcEngine

        if not intent.change_input:
            return QueryResult(explanation="无法解析假设条件")

        change = intent.change_input
        item = self._fuzzy_match_entity(change["entity"])
        if not item or not item.value_cell:
            return QueryResult(explanation=f"未找到 '{change['entity']}'")

        try:
            new_value = float(change["value"])
        except ValueError:
            return QueryResult(explanation=f"无法解析值: {change['value']}")

        old_cell = self.fg.cells.get(item.value_cell)
        old_value = old_cell.value if old_cell else None

        if old_cell:
            old_cell.value = new_value

        engine = RecalcEngine(self.fg)
        result = engine.recalculate(item.value_cell)

        if old_cell:
            old_cell.value = old_value

        changed_count = result.total_changed
        top_changes = result.changed_cells[:5]

        explanation = f"修改 {item.name}: {old_value} -> {new_value}\n"
        explanation += f"影响 {changed_count} 个单元格:\n"
        for delta in top_changes:
            explanation += f"  {delta.cell_id}: {delta.old_value} -> {delta.new_value}\n"

        return QueryResult(explanation=explanation)
