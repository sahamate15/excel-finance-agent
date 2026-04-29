"""Validation helpers for instructions, formulas, and cell references.

These functions are deliberately conservative: they reject anything that
looks suspicious so the LLM cannot inject control characters, infinitely
long strings, or malformed Excel syntax into a workbook.
"""

from __future__ import annotations

import re

_CELL_REF_RE = re.compile(r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")
_FORMULA_FUNC_OR_OP_RE = re.compile(r"[A-Z]{2,}\s*\(|[\+\-\*/\^]")
_FORBIDDEN_CHARS = {"\n", "\r", "\t", "\x00"}
_MAX_FORMULA_LEN = 1000
_MAX_INSTRUCTION_LEN = 500


def validate_excel_formula(formula: str) -> tuple[bool, str]:
    """Validate that *formula* is a syntactically plausible Excel formula.

    The check is intentionally lightweight — it cannot guarantee that the
    formula evaluates correctly, only that it will not corrupt the workbook
    or trigger obvious Excel parsing errors.

    Args:
        formula: The candidate formula string returned by the LLM or fallback.

    Returns:
        A ``(is_valid, error_message)`` pair. ``error_message`` is empty when
        the formula passes all checks.
    """
    if not isinstance(formula, str):
        return False, "formula must be a string"

    if not formula:
        return False, "formula is empty"

    if not formula.startswith("="):
        return False, "formula must start with ="

    if len(formula) > _MAX_FORMULA_LEN:
        return False, f"formula exceeds {_MAX_FORMULA_LEN} characters"

    if any(ch in formula for ch in _FORBIDDEN_CHARS):
        return False, "formula contains forbidden control characters"

    body = formula[1:].strip()
    if not body:
        return False, "formula has no body after ="

    # Balanced parentheses
    depth = 0
    for ch in body:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                return False, "unbalanced parentheses (extra closing)"
    if depth != 0:
        return False, "unbalanced parentheses"

    # Must contain at least one function call or arithmetic operator, OR be a
    # pure cell-reference passthrough like =A1.
    if not _FORMULA_FUNC_OR_OP_RE.search(body) and not _CELL_REF_RE.match(body):
        return False, "formula has no function or operator"

    return True, ""


def validate_cell_reference(cell: str) -> bool:
    """Return True if *cell* is a valid A1-style Excel reference.

    Accepts references from ``A1`` up to the Excel-2007 maximum
    ``XFD1048576``.
    """
    if not isinstance(cell, str):
        return False
    return bool(_CELL_REF_RE.match(cell.strip().upper()))


def sanitize_instruction(text: str) -> str:
    """Normalise a free-text instruction for downstream processing.

    Strips whitespace, collapses internal whitespace runs, removes characters
    outside the printable ASCII + common Latin-1 punctuation range, and
    truncates to ``MAX_INSTRUCTION_LEN`` characters.

    Args:
        text: Raw user instruction.

    Returns:
        A cleaned instruction safe to embed in an LLM prompt.
    """
    if not isinstance(text, str):
        return ""
    cleaned = re.sub(r"[^\w\s\.,:;!\?\-\(\)\[\]\%\$\+\*/=^&'\"]+", " ", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:_MAX_INSTRUCTION_LEN]


# ──────────────────────────────────────────────────────────────────────
# Finance-domain table_config validation
# ──────────────────────────────────────────────────────────────────────
#
# Bounds chosen for IB use cases:
# - years <= 50: covers long infrastructure / PE horizons; blocks runaway tables
# - tenure_months <= 600 (50 years): covers any real loan
# - growth_rate <= 5.0 (500%): allows aggressive early-stage projections, blocks
#   common fat-finger errors like entering "50" instead of "0.5"
# - cost_ratio <= 2.0: allows unprofitable businesses (costs 2x revenue) while
#   blocking nonsense values
# - rate (depreciation) capped at 1.0: >100% is meaningless for WDV
# - annual_rate (loan) allowed at 0: zero-interest intercompany loans exist


def _check_number(
    value: object,
    *,
    field: str,
    lower: float,
    upper: float,
    lower_inclusive: bool,
    upper_inclusive: bool,
) -> tuple[bool, str]:
    """Return (ok, error) for a numeric field bounded by ``lower`` / ``upper``."""
    try:
        n = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return False, f"{field} must be a number, got {type(value).__name__}"
    if lower_inclusive and n < lower:
        return False, f"{field} must be >= {lower}, got {n}"
    if not lower_inclusive and n <= lower:
        return False, f"{field} must be > {lower}, got {n}"
    if upper_inclusive and n > upper:
        return False, f"{field} must be <= {upper}, got {n}"
    if not upper_inclusive and n >= upper:
        return False, f"{field} must be < {upper}, got {n}"
    return True, ""


def _check_int(value: object, *, field: str, lower: int, upper: int) -> tuple[bool, str]:
    """Return (ok, error) for an integer field bounded by ``[lower, upper]``."""
    try:
        n = int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return False, f"{field} must be an integer, got {type(value).__name__}"
    if n < lower:
        return False, f"{field} must be >= {lower}, got {n}"
    if n > upper:
        return False, f"{field} must be <= {upper}, got {n}"
    return True, ""


def validate_table_config(cfg: dict) -> tuple[bool, str]:
    """Validate a table-builder config against IB-domain bounds.

    Returns:
        ``(True, "")`` if all fields are present and within sane ranges, or
        ``(False, error_message)`` describing which field failed.
    """
    if not isinstance(cfg, dict):
        return False, "table_config must be a dict"

    table_type = cfg.get("type")
    if table_type not in {"depreciation", "amortization", "projection"}:
        return False, f"unsupported table type: {table_type!r}"

    if table_type == "depreciation":
        if "asset_value" not in cfg:
            return False, "depreciation requires 'asset_value'"
        if "rate" not in cfg:
            return False, "depreciation requires 'rate'"
        if "years" not in cfg:
            return False, "depreciation requires 'years'"

        ok, err = _check_number(
            cfg["asset_value"], field="asset_value",
            lower=0, upper=1e15, lower_inclusive=False, upper_inclusive=True,
        )
        if not ok:
            return ok, err

        method = cfg.get("method", "wdv")
        if method not in {"wdv", "straight_line"}:
            return False, f"depreciation method must be 'wdv' or 'straight_line', got {method!r}"

        # Rate semantics depend on the method. For WDV: 0 < rate <= 1.
        # For straight-line: rate is unused; salvage_value drives the schedule.
        if method == "wdv":
            ok, err = _check_number(
                cfg["rate"], field="rate",
                lower=0, upper=1, lower_inclusive=False, upper_inclusive=True,
            )
            if not ok:
                return ok, err

        ok, err = _check_int(cfg["years"], field="years", lower=1, upper=50)
        if not ok:
            return ok, err

        if "salvage_value" in cfg:
            ok, err = _check_number(
                cfg["salvage_value"], field="salvage_value",
                lower=0, upper=1e15, lower_inclusive=True, upper_inclusive=True,
            )
            if not ok:
                return ok, err
            if float(cfg["salvage_value"]) >= float(cfg["asset_value"]):
                return False, "salvage_value must be less than asset_value"

        return True, ""

    if table_type == "amortization":
        if "principal" not in cfg:
            return False, "amortization requires 'principal'"
        if "annual_rate" not in cfg:
            return False, "amortization requires 'annual_rate'"
        if "tenure_months" not in cfg:
            return False, "amortization requires 'tenure_months'"

        ok, err = _check_number(
            cfg["principal"], field="principal",
            lower=0, upper=1e15, lower_inclusive=False, upper_inclusive=True,
        )
        if not ok:
            return ok, err
        ok, err = _check_number(
            cfg["annual_rate"], field="annual_rate",
            lower=0, upper=1, lower_inclusive=True, upper_inclusive=True,
        )
        if not ok:
            return ok, err
        ok, err = _check_int(cfg["tenure_months"], field="tenure_months", lower=1, upper=600)
        if not ok:
            return ok, err
        return True, ""

    # projection
    if "base_revenue" not in cfg:
        return False, "projection requires 'base_revenue'"
    if "growth_rate" not in cfg:
        return False, "projection requires 'growth_rate'"
    if "years" not in cfg:
        return False, "projection requires 'years'"

    ok, err = _check_number(
        cfg["base_revenue"], field="base_revenue",
        lower=0, upper=1e15, lower_inclusive=False, upper_inclusive=True,
    )
    if not ok:
        return ok, err
    ok, err = _check_number(
        cfg["growth_rate"], field="growth_rate",
        lower=-1, upper=5.0, lower_inclusive=False, upper_inclusive=True,
    )
    if not ok:
        return ok, err
    ok, err = _check_int(cfg["years"], field="years", lower=1, upper=50)
    if not ok:
        return ok, err

    cost_ratio = cfg.get("cost_ratio", 0.6)
    ok, err = _check_number(
        cost_ratio, field="cost_ratio",
        lower=0, upper=2.0, lower_inclusive=False, upper_inclusive=True,
    )
    if not ok:
        return ok, err

    growth_method = cfg.get("growth_method", "compound")
    if growth_method not in {"compound", "linear"}:
        return False, f"projection growth_method must be 'compound' or 'linear', got {growth_method!r}"

    return True, ""
