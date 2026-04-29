"""Offline, deterministic fallback formulas for common finance tasks.

Used by the orchestrator before contacting the LLM (when ``FALLBACK_FIRST``
is true) and as a graceful degradation path if the LLM is unreachable.
"""

from __future__ import annotations

import re

FORMULA_FALLBACKS: dict[str, str] = {
    "growth rate": "=(B2-A2)/A2",
    "cagr": "=(B2/A2)^(1/5)-1",
    "irr": "=IRR(B2:B10)",
    "npv": "=NPV(0.1,B2:B10)",
    "wdv depreciation": "=ROUND(B2*(1-0.25)^A2,2)",
    "net profit margin": "=B2/A2",
    "gross margin": "=(A2-B2)/A2",
    "debt to equity": "=A2/B2",
    "roe": "=A2/B2",
    "pmt": "=PMT(B2/12,C2,-A2)",
    "average": "=AVERAGE(A2:A10)",
    "median": "=MEDIAN(A2:A10)",
    "sum": "=SUM(A2:A10)",
    "compound interest": "=A2*(1+B2)^C2",
    "simple interest": "=A2*B2*C2",
}

# Stop words that don't help us match a finance concept.
_STOP_WORDS = {
    "calculate", "compute", "find", "what", "is", "the", "a", "an", "for",
    "of", "in", "on", "at", "to", "and", "or", "with", "between", "from",
    "give", "show", "me", "please", "want", "need", "value",
}


def _tokenise(text: str) -> set[str]:
    """Lowercase a string and return the set of non-stopword tokens."""
    tokens = re.findall(r"[a-zA-Z]+", text.lower())
    return {t for t in tokens if t not in _STOP_WORDS and len(t) > 1}


def get_fallback_formula(instruction: str) -> str | None:
    """Return the best-matching hardcoded formula for *instruction*.

    Matching is a token-overlap score — the dictionary key whose tokens have
    the highest Jaccard-like overlap with the instruction wins, provided the
    overlap is non-zero and at least one key-token is present.

    Args:
        instruction: Free-text user instruction.

    Returns:
        The matching formula string, or ``None`` if no key shares any token
        with the instruction.
    """
    if not isinstance(instruction, str) or not instruction.strip():
        return None

    instr_tokens = _tokenise(instruction)
    if not instr_tokens:
        return None

    best_key: str | None = None
    best_score: float = 0.0

    for key in FORMULA_FALLBACKS:
        key_tokens = _tokenise(key)
        if not key_tokens:
            continue
        overlap = len(instr_tokens & key_tokens)
        if overlap == 0:
            continue
        # Reward keys whose tokens are *all* present in the instruction.
        score = overlap / len(key_tokens)
        # Tie-breaker: prefer longer (more specific) keys.
        if score > best_score or (score == best_score and best_key is not None
                                  and len(key) > len(best_key)):
            best_score = score
            best_key = key

    if best_key is None or best_score < 0.5:
        return None
    return FORMULA_FALLBACKS[best_key]
