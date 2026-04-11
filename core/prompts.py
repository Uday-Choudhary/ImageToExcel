"""
Vision model prompts for the ImageToExcel pipeline.

Single source of truth — both the Streamlit UI and CLI pipeline
import from here to avoid prompt drift between interfaces.
"""

from __future__ import annotations

VISION_EXTRACTION_PROMPT: str = """### INPUT ANALYSIS PHASE
1. **Identify Script Style**: First, detect if the text is printed, handwritten, or a mix. Adjust your internal recognition weights for handwritten characters.
2. **Spatial Anchoring**: Use visual lines and margins to define table boundaries.

### EXTRACTION REQUIREMENTS
1. **Hierarchical Extraction**:
   - **Header Data**: Extract titles, dates, and names.
   - **Tabular Structures**: Reconstruct grid-like data.
   - **Currency Detection**: Capture currency symbols exactly.
2. **Handwritten Edge Cases**:
   - Ignore crossed-out text.
   - Use [illegible] for unreadable words.
3. **The "Logical Audit"**:
   - Manually recalculate sums.
   - Report discrepancies as `handwritten_total` and `computed_total`.

### OUTPUT STRUCTURE (JSON)
Constraints:
- Return ONLY valid JSON.
- Do not include comments or markdown formatting (```json).
- "normalized_value" must be a single number or string, NOT a mathematical expression (e.g., use 100.50, not 50+50.50).
- If a calculation is needed, perform it internally and output the result.

{
  "document_summary": { "style": "handwritten/printed", "domain": "auto-detect" },
  "entities": { "label": "value" },
  "tables": [
    {
      "table_description": "string",
      "headers": [],
      "rows": [
        { "column_name": "raw_text", "normalized_value": "numeric_or_string", "currency": "ISO_CODE" }
      ],
      "validation": {
        "math_check": "passed/failed",
        "notes": "string"
      }
    }
  ]
}"""
