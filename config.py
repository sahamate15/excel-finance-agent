"""Centralised application configuration.

Runtime settings are loaded once from environment variables (a ``.env`` file
when present) into a mutable :class:`AppConfig` singleton. The Streamlit UI
calls :func:`update_config` to apply per-session overrides without persisting
them. Never hard-code secrets — always go through ``os.getenv``.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

# Load .env from the project root, regardless of where Python is launched from.
PROJECT_ROOT: Path = Path(__file__).resolve().parent
load_dotenv(PROJECT_ROOT / ".env")


def _str_to_bool(value: str | None, default: bool = False) -> bool:
    """Convert a string env var to bool. Accepts true/false/1/0/yes/no."""
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


@dataclass
class AppConfig:
    """Application configuration container.

    Mutable so UI surfaces (Streamlit) can override individual fields per
    session via :func:`update_config`. CLI / test environments still get
    immutable-by-default behaviour because nothing mutates the defaults
    unless explicitly asked.

    Attributes:
        mistral_api_key: Secret key for Mistral's chat completions endpoint.
        mistral_model:   Chat completion model name (e.g. mistral-small-latest).
        mistral_base_url: Base URL for the OpenAI-compatible endpoint.
        default_excel_file: Path to the workbook used when none is given.
        default_sheet:  Sheet name used when none is given.
        log_level:      Logging level string (DEBUG/INFO/WARN/ERROR).
        max_retries:    LLM retry budget for malformed responses.
        fallback_first: If True, try hardcoded formulas before calling the LLM.
        strict_mode:    If True, the LLM is never contacted; only the offline
            formula map and heuristic table parser are used. Instructions the
            offline path cannot resolve fail loudly with a clear error. For
            compliance contexts where no instruction-derived data may leave
            the machine.
        log_dir:        Directory where rotating logs are written.
        data_dir:       Directory containing sample workbooks and user data.
    """

    mistral_api_key: str
    mistral_model: str
    mistral_base_url: str
    default_excel_file: Path
    default_sheet: str
    log_level: str
    max_retries: int
    fallback_first: bool
    strict_mode: bool
    log_dir: Path
    data_dir: Path


def load_config() -> AppConfig:
    """Build an :class:`AppConfig` from the current process environment.

    Returns:
        A fully populated :class:`AppConfig`. Missing optional values fall back
        to sensible defaults; missing required values yield empty strings so
        the caller can decide how to react (the AI engine raises if the key
        is needed but absent).
    """
    default_file = os.getenv("DEFAULT_EXCEL_FILE", "data/sample_workbook.xlsx")
    excel_path = (PROJECT_ROOT / default_file).resolve() if not Path(default_file).is_absolute() else Path(default_file)

    return AppConfig(
        mistral_api_key=os.getenv("MISTRAL_API_KEY", ""),
        mistral_model=os.getenv("MISTRAL_MODEL", "mistral-small-latest"),
        mistral_base_url=os.getenv("MISTRAL_BASE_URL", "https://api.mistral.ai/v1"),
        default_excel_file=excel_path,
        default_sheet=os.getenv("DEFAULT_SHEET", "Sheet1"),
        log_level=os.getenv("LOG_LEVEL", "INFO").upper(),
        max_retries=int(os.getenv("MAX_RETRIES", "2")),
        fallback_first=_str_to_bool(os.getenv("FALLBACK_FIRST"), default=True),
        strict_mode=_str_to_bool(os.getenv("STRICT_MODE"), default=False),
        log_dir=PROJECT_ROOT / "logs",
        data_dir=PROJECT_ROOT / "data",
    )


# Module-level singleton — import this everywhere instead of re-loading.
CONFIG: AppConfig = load_config()


def update_config(**overrides: Any) -> AppConfig:
    """Mutate fields on the singleton :data:`CONFIG`.

    Used by the Streamlit UI to apply per-session overrides without writing
    them back to ``.env``. If the LLM key, model, or base URL changed, the
    cached LLM client is reset so subsequent calls pick up the new values.

    Args:
        **overrides: Field names and new values. Unknown fields raise
            :class:`AttributeError`.

    Returns:
        The (mutated) :data:`CONFIG` singleton.
    """
    llm_changed = False
    for field_name, value in overrides.items():
        if not hasattr(CONFIG, field_name):
            raise AttributeError(f"AppConfig has no field {field_name!r}")
        if field_name in {"mistral_api_key", "mistral_model", "mistral_base_url"}:
            if getattr(CONFIG, field_name) != value:
                llm_changed = True
        setattr(CONFIG, field_name, value)

    if llm_changed:
        # Local import to avoid a circular import at module load time.
        from agents.ai_engine import reset_llm_client  # noqa: PLC0415
        reset_llm_client()

    return CONFIG


# Re-export Any for the type annotation used above.
__all__ = ["AppConfig", "CONFIG", "PROJECT_ROOT", "load_config", "update_config"]
