"""LLM-facing functions: formula generation, task classification, clarification.

All prompts live as module-level constants so they can be reviewed and
versioned independently. Every external call is wrapped in retries and
falls back to deterministic heuristics when the API is unavailable.

The OpenAI SDK is used as a thin client against Mistral's
OpenAI-compatible endpoint (https://api.mistral.ai/v1). All four call sites
exercise the standard ``chat.completions.create`` shape that Mistral's
compatibility layer implements, including JSON-mode responses.
"""

from __future__ import annotations

import json
import re
from typing import Any

from openai import OpenAI, OpenAIError

from config import CONFIG
from utils.logger import get_logger
from utils.validators import validate_excel_formula

logger = get_logger(__name__)

# ──────────────────────────────────────────────────────────────────────
# Prompt constants
# ──────────────────────────────────────────────────────────────────────

FORMULA_SYSTEM_PROMPT = """You are an expert Excel formula generator for finance professionals.

Output rules:
- Return ONLY the formula text, starting with "=" — nothing before, nothing after
- The output must be a single Excel expression. Do NOT add a name/label prefix
  inside the formula (e.g. "=EBITDA_Margin:C5/B5" is INVALID — return "=C5/B5")
- No explanations, no markdown, no code fences, no trailing punctuation
- Optimize for financial accuracy
- Use these cell references if provided: {context}

Finance formula reference:
- Growth Rate: =(new-old)/old
- CAGR: =(end/start)^(1/years)-1
- NPV: =NPV(rate, cashflow_range)
- IRR: =IRR(cashflow_range)
- WDV Depreciation: =ROUND(asset_value*(1-rate)^year, 2)
- PMT (loan): =PMT(rate/12, tenure_months, -principal)
- EBITDA Margin: =EBITDA_cell/Revenue_cell

Examples (input → correct output):
- "EBITDA margin if EBITDA in C5 and revenue in B5" → =C5/B5
- "growth rate between B2 and B3" → =(B3-B2)/B2
- "IRR for cash flows in B2:B7" → =IRR(B2:B7)
"""

TASK_DETECT_SYSTEM_PROMPT = """You classify finance Excel instructions.

Return ONLY a JSON object with these exact keys:
{
  "type": "formula" | "table" | "chart",
  "complexity": "simple" | "complex",
  "finance_domain": string,
  "requires_clarification": boolean,
  "clarification_question": string or null
}

Guidance:
- "formula" = a single calculation that fits in one cell (growth rate, IRR, NPV, EMI, margin)
- "table"   = a multi-row layout (depreciation schedule, amortization, projection, forecast)
- "chart"   = explicit request for a visual chart
- requires_clarification=true ONLY when essential numeric inputs are missing
- finance_domain examples: "valuation", "lending", "depreciation", "ratios", "forecasting"
"""

CLARIFY_SYSTEM_PROMPT = """You ask ONE concise clarifying question to fill in
missing finance inputs (asset value, rate, tenure, cell ranges, etc.).
Return only the question text, no preamble. Keep it under 25 words."""

TABLE_CONFIG_SYSTEM_PROMPT = """Extract a table_config JSON object from a finance instruction.

Return ONLY JSON. Schema by table type:

depreciation:
  {"type":"depreciation","asset_value":number,"rate":number,"years":int,
   "method":"wdv"|"straight_line","salvage_value":number,
   "start_row":int,"start_col":int}

amortization:
  {"type":"amortization","principal":number,"annual_rate":number,
   "tenure_months":int,"start_row":int,"start_col":int}

projection:
  {"type":"projection","base_revenue":number,"growth_rate":number,
   "cost_ratio":number,"years":int,"growth_method":"compound"|"linear",
   "start_row":int,"start_col":int}

Defaults: start_row=1, start_col=1, years=5, cost_ratio=0.6, method="wdv",
salvage_value=0, growth_method="compound" if absent.
Set method="straight_line" only if the instruction explicitly says
"straight line", "straight-line", "SL", or "linear depreciation".
Set growth_method="linear" only if the instruction explicitly says
"linear growth", "simple growth", or "flat growth".
Convert percentages to decimals (25% -> 0.25, "9 percent" -> 0.09).
Convert "20 lakh" -> 2000000, "50 lakh" -> 5000000, "1 crore" -> 10000000.
"""

# ──────────────────────────────────────────────────────────────────────
# Client construction
# ──────────────────────────────────────────────────────────────────────

_client: OpenAI | None = None


def _get_client() -> OpenAI:
    """Lazily build a singleton client pointed at Mistral's compatible endpoint."""
    global _client
    if _client is None:
        if not CONFIG.mistral_api_key:
            raise RuntimeError(
                "MISTRAL_API_KEY is not set. Add it to your .env file or "
                "export it in the shell before running the agent."
            )
        kwargs: dict[str, Any] = {
            "api_key": CONFIG.mistral_api_key,
            "base_url": CONFIG.mistral_base_url,
        }
        _client = OpenAI(**kwargs)
    return _client


def reset_llm_client() -> None:
    """Drop the cached client so the next call rebuilds with current CONFIG values.

    Called by :func:`config.update_config` whenever the API key, model, or
    base URL changes — typically from the Streamlit UI when the user pastes
    a new key or switches models mid-session.
    """
    global _client
    _client = None


# ──────────────────────────────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────────────────────────────


def text_to_formula(instruction: str, context: dict | None = None) -> str:
    """Convert a natural-language instruction into an Excel formula via LLM.

    Args:
        instruction: User instruction in plain English.
        context:     Optional dict of cell references / domain hints, e.g.
            ``{"revenue_cell": "B2", "cost_cell": "C2"}``.

    Returns:
        A validated formula string starting with ``=``.

    Raises:
        ValueError: If the LLM returns malformed output after all retries.
        RuntimeError: If the LLM client cannot be constructed.
    """
    client = _get_client()
    ctx_str = json.dumps(context or {}, ensure_ascii=False)
    system_prompt = FORMULA_SYSTEM_PROMPT.format(context=ctx_str)

    last_error = ""
    for attempt in range(CONFIG.max_retries + 1):
        try:
            response = client.chat.completions.create(
                model=CONFIG.mistral_model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": instruction},
                ],
                temperature=0,
            )
            raw = (response.choices[0].message.content or "").strip()
            formula = _strip_to_formula(raw)
            ok, err = validate_excel_formula(formula)
            if ok:
                logger.info("text_to_formula success on attempt %d", attempt + 1)
                return formula
            last_error = err
            logger.warning("text_to_formula attempt %d returned invalid formula: %s", attempt + 1, err)
        except OpenAIError as exc:
            last_error = type(exc).__name__
            logger.warning("text_to_formula attempt %d API error: %s", attempt + 1, type(exc).__name__)

    raise ValueError(f"LLM did not return a valid formula after retries: {last_error}")


def detect_task_type(instruction: str) -> dict:
    """Classify *instruction* into formula / table / chart.

    Falls back to keyword matching if the LLM call fails so the agent
    remains usable in offline or rate-limited scenarios.

    Returns:
        Dict with keys ``type``, ``complexity``, ``finance_domain``,
        ``requires_clarification``, ``clarification_question``.
    """
    try:
        client = _get_client()
        response = client.chat.completions.create(
            model=CONFIG.mistral_model,
            messages=[
                {"role": "system", "content": TASK_DETECT_SYSTEM_PROMPT},
                {"role": "user", "content": instruction},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        raw = (response.choices[0].message.content or "").strip()
        parsed = json.loads(raw)
        return _normalise_task_dict(parsed)
    except (OpenAIError, json.JSONDecodeError, RuntimeError) as exc:
        logger.warning("detect_task_type LLM failed (%s); using keyword fallback", exc)
        return _keyword_classify(instruction)


def clarify_input(instruction: str) -> str:
    """Generate a single clarifying question when inputs are ambiguous.

    Args:
        instruction: The original instruction that lacked enough information.

    Returns:
        A short, focused question string. Falls back to a generic prompt if
        the LLM is unavailable.
    """
    try:
        client = _get_client()
        response = client.chat.completions.create(
            model=CONFIG.mistral_model,
            messages=[
                {"role": "system", "content": CLARIFY_SYSTEM_PROMPT},
                {"role": "user", "content": instruction},
            ],
            temperature=0.2,
        )
        question = (response.choices[0].message.content or "").strip()
        return question or "Could you specify the exact values, rate, and time period?"
    except (OpenAIError, RuntimeError) as exc:
        logger.warning("clarify_input failed (%s); using default", exc)
        return "Could you specify the exact values, rate, and time period?"


def extract_table_config(instruction: str) -> dict:
    """Parse a table-creation instruction into a structured config dict via LLM.

    Args:
        instruction: User instruction describing a multi-row table.

    Returns:
        A dict matching one of the schemas documented in
        :data:`TABLE_CONFIG_SYSTEM_PROMPT`.

    Raises:
        ValueError: If the LLM is unreachable or returns malformed JSON.
            Callers should catch and decide whether to fall back to
            :func:`_heuristic_table_config` (for source tracking).
        RuntimeError: If the LLM client cannot be constructed.
    """
    client = _get_client()
    try:
        response = client.chat.completions.create(
            model=CONFIG.mistral_model,
            messages=[
                {"role": "system", "content": TABLE_CONFIG_SYSTEM_PROMPT},
                {"role": "user", "content": instruction},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        raw = (response.choices[0].message.content or "").strip()
        return json.loads(raw)
    except (OpenAIError, json.JSONDecodeError) as exc:
        logger.warning("extract_table_config LLM error: %s", type(exc).__name__)
        raise ValueError(f"LLM table_config extraction failed: {type(exc).__name__}") from exc


# ──────────────────────────────────────────────────────────────────────
# Internal helpers
# ──────────────────────────────────────────────────────────────────────


def _strip_to_formula(raw: str) -> str:
    """Remove markdown fences/labels and isolate the first formula token."""
    text = raw.strip()
    # Drop ```...``` fences if present.
    text = re.sub(r"^```(?:excel|formula|text)?\s*|\s*```$", "", text, flags=re.IGNORECASE | re.MULTILINE)
    text = text.strip().strip("`").strip()
    # If the model prefixed something like "Formula: =..." pull out the =-clause.
    match = re.search(r"=.+", text)
    if match:
        return match.group(0).split("\n")[0].strip()
    return text


def _normalise_task_dict(parsed: dict) -> dict:
    """Coerce arbitrary LLM JSON into the canonical task-type schema."""
    return {
        "type": str(parsed.get("type", "formula")).lower(),
        "complexity": str(parsed.get("complexity", "simple")).lower(),
        "finance_domain": str(parsed.get("finance_domain", "general")),
        "requires_clarification": bool(parsed.get("requires_clarification", False)),
        "clarification_question": parsed.get("clarification_question"),
    }


_TABLE_KEYWORDS = (
    "schedule", "table", "amortization", "amortisation", "projection",
    "forecast", "plan", "5-year", "5 year", "annual",
)


def _keyword_classify(instruction: str) -> dict:
    """Pure-Python fallback classifier for when the LLM is unavailable."""
    lower = instruction.lower()
    if any(k in lower for k in _TABLE_KEYWORDS):
        task_type = "table"
    elif "chart" in lower or "graph" in lower or "plot" in lower:
        task_type = "chart"
    else:
        task_type = "formula"
    return {
        "type": task_type,
        "complexity": "complex" if task_type == "table" else "simple",
        "finance_domain": "general",
        "requires_clarification": False,
        "clarification_question": None,
    }


def _heuristic_table_config(instruction: str, *, strict: bool = False) -> dict | None:
    """Best-effort regex-driven table_config for offline mode.

    Looks for *labelled* numbers ("5-year", "20 lakh", "25%") rather than
    naively grabbing every digit in the string, so phrases like "5-year
    schedule for 20 lakh at 25%" parse correctly.

    Args:
        instruction: Free-text instruction describing the table.
        strict: When True, returns ``None`` if the instruction does not name
            a recognised table type (depreciation / amortization / projection
            keywords). When False (default), falls through to a depreciation
            config with sensible defaults — preserving the legacy behaviour
            used by the LLM fallback path.
    """
    lower = instruction.lower()

    def find_money(text: str) -> float | None:
        # "20 lakh", "50 lakhs", "1.5 crore", "5,00,000", plain "5000000"
        m = re.search(r"(\d+(?:[.,]\d+)?)\s*(crore|cr|lakh|lakhs|lac|lacs|k|m)?", text)
        if not m:
            return None
        value = float(m.group(1).replace(",", ""))
        unit = (m.group(2) or "").lower()
        if unit in {"crore", "cr"}:
            return value * 10_000_000
        if unit in {"lakh", "lakhs", "lac", "lacs"}:
            return value * 100_000
        if unit == "k":
            return value * 1_000
        if unit == "m":
            return value * 1_000_000
        return value

    def find_percent(text: str) -> float | None:
        m = re.search(r"(\d+(?:\.\d+)?)\s*%", text)
        if m:
            return float(m.group(1)) / 100
        m = re.search(r"(\d+(?:\.\d+)?)\s*percent", text)
        if m:
            return float(m.group(1)) / 100
        return None

    def find_years(text: str) -> int | None:
        m = re.search(r"(\d+)[\s-]*year", text)
        return int(m.group(1)) if m else None

    if any(k in lower for k in ("amortization", "amortisation", "loan", "emi")):
        # "50 lakh loan at 8.5% for 20 years"
        money_match = re.search(r"(\d+(?:[.,]\d+)?\s*(?:lakh|lakhs|lac|lacs|crore|cr|k|m)?)\s+(?:loan|principal)", lower)
        principal = find_money(money_match.group(1)) if money_match else find_money(lower)
        if principal is None:
            principal = 5_000_000
        rate = find_percent(lower) or 0.09
        years = find_years(lower) or 10
        return {
            "type": "amortization",
            "principal": principal,
            "annual_rate": rate,
            "tenure_months": years * 12,
            "start_row": 1,
            "start_col": 1,
        }

    if any(k in lower for k in ("projection", "forecast", "revenue")):
        # Pick the LARGEST money-like value as the base. Same pattern as the
        # depreciation default branch — robust against phrasings like "at 50
        # lakh" where the legacy prefix-only regex would mistake "5-year" for 5.
        money_candidates = re.findall(
            r"(\d+(?:[.,]\d+)?)\s*(crore|cr|lakh|lakhs|lac|lacs|k|m)?",
            lower,
        )
        base = 5_000_000
        best = 0.0
        for num, unit in money_candidates:
            candidate = find_money(f"{num} {unit}".strip())
            if candidate is not None and candidate > best:
                best = candidate
                base = candidate
        growth = find_percent(lower) or 0.15
        years = find_years(lower) or 5
        # Linear growth requested explicitly?
        growth_method = "linear" if any(
            k in lower for k in ("linear growth", "simple growth", "flat growth")
        ) else "compound"
        return {
            "type": "projection",
            "base_revenue": base,
            "growth_rate": growth,
            "cost_ratio": 0.6,
            "years": years,
            "growth_method": growth_method,
            "start_row": 1,
            "start_col": 1,
        }

    # Depreciation branch: only fire when the instruction names it explicitly,
    # OR (in non-strict mode) as the catch-all default.
    has_dep_keyword = any(
        k in lower for k in ("depreciation", "depreciate", "wdv", "asset value", "asset")
    )
    if strict and not has_dep_keyword:
        return None

    # Pick the LARGEST money-like value in the string as the asset value to
    # avoid mistaking "5-year" for 5.
    money_candidates = re.findall(
        r"(\d+(?:[.,]\d+)?)\s*(crore|cr|lakh|lakhs|lac|lacs|k|m)?",
        lower,
    )
    asset = 2_000_000
    best = 0.0
    for num, unit in money_candidates:
        candidate = find_money(f"{num} {unit}".strip())
        if candidate is not None and candidate > best:
            best = candidate
            asset = candidate
    rate = find_percent(lower) or 0.25
    years = find_years(lower) or 5
    # Straight-line depreciation requested?
    method = "straight_line" if any(
        k in lower for k in ("straight line", "straight-line", "linear depreciation")
    ) or re.search(r"\bsl\b", lower) else "wdv"
    cfg: dict = {
        "type": "depreciation",
        "asset_value": asset,
        "rate": rate,
        "years": years,
        "method": method,
        "start_row": 1,
        "start_col": 1,
    }
    # Salvage value extraction is optional; only set if the instruction names it.
    salvage_match = re.search(
        r"salvage(?:\s+value)?\s+(?:of\s+)?(\d+(?:[.,]\d+)?\s*(?:lakh|lakhs|lac|lacs|crore|cr|k|m)?)",
        lower,
    )
    if salvage_match:
        salvage = find_money(salvage_match.group(1))
        if salvage is not None:
            cfg["salvage_value"] = salvage
    return cfg
