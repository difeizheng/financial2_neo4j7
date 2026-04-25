"""LLM-based business item validator — confirms ambiguous detections.

Used when rule-based detection is uncertain about:
1. Whether a row is a data row or header/section marker
2. Column role assignment (which column is name/value/unit)
3. Item name extraction from merged cells
"""

from __future__ import annotations

from typing import Any

PROMPT_TEMPLATE = """You are validating financial model structure detection.

## Context
An Excel financial model has been parsed. The system detected a potential business item row.
Your job: confirm if this is a real financial concept and assign column roles.

## Sheet: {sheet}
## Row {row} content:
{row_content}

## Questions:
1. Is this row a real financial data row (not a header, section title, or empty row)?
2. What is the item name?
3. Which column contains the main value?
4. Which column contains the unit?
5. Is this a time series row (values across multiple time periods)?

## Row context:
{formatted_cells}

Answer in JSON format:
{{"is_data_row": true/false, "name": "string", "value_col": "letter", "unit_col": "letter", "has_time_series": true/false}}
"""


class ItemValidator:
    """Validates business items using LLM."""

    def __init__(self, llm_client: Any = None):
        self.llm = llm_client

    def validate(self, sheet_name: str, row_num: int, cells: list[dict]) -> dict:
        """Validate a potential business item row.

        Args:
            sheet_name: Sheet name
            row_num: Row number
            cells: List of {col, value, formula, data_type} dicts

        Returns:
            {is_data_row, name, value_col, unit_col, has_time_series}
        """
        if not self.llm:
            # Fallback: rule-based decision
            return self._rule_fallback(cells)

        prompt = PROMPT_TEMPLATE.format(
            sheet=sheet_name,
            row=row_num,
            row_content=cells,
            formatted_cells=self._format_cells(cells),
        )

        try:
            response = self.llm.generate(prompt)
            import json
            result = json.loads(response)
            return result
        except Exception:
            return self._rule_fallback(cells)

    def _rule_fallback(self, cells: list[dict]) -> dict:
        """Rule-based fallback when LLM unavailable."""
        text_cells = [c for c in cells if c.get("data_type") == "text" and not c.get("formula")]
        val_cells = [c for c in cells if c.get("data_type") == "number" or c.get("formula")]

        is_data_row = len(text_cells) >= 1 and len(val_cells) >= 1
        name = text_cells[0].get("value", "") if text_cells else ""

        # Check if name looks like a header/section
        is_header = len(name) <= 3 and name in ("项目", "合计", "总计", "小计", "单位")

        return {
            "is_data_row": is_data_row and not is_header,
            "name": name,
            "value_col": val_cells[0].get("col") if val_cells else None,
            "unit_col": None,
            "has_time_series": len(val_cells) > 5,
        }

    def _format_cells(self, cells: list[dict]) -> str:
        """Format cells for LLM context."""
        lines = []
        for c in sorted(cells, key=lambda x: x.get("col", "")):
            col = c.get("col", "?")
            val = c.get("value", "")
            formula = c.get("formula", "")
            dtype = c.get("data_type", "")
            line = f"  {col}: value={val}, type={dtype}"
            if formula:
                line += f", formula={formula}"
            lines.append(line)
        return "\n".join(lines)
