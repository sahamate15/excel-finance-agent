"""Logging setup for the Excel agent.

Provides a single :func:`get_logger` helper that returns a configured logger
writing to ``logs/agent.log`` with rotation. All modules in the project
should obtain their logger via this helper rather than calling
``logging.getLogger`` directly, to ensure the rotating handler is attached
exactly once.
"""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from config import CONFIG

_LOG_FORMAT = "[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
_MAX_BYTES = 5 * 1024 * 1024  # 5 MB
_BACKUP_COUNT = 3

_initialised = False


def _initialise_root() -> None:
    """Attach the rotating file handler and console handler to the root logger.

    Called lazily on first use of :func:`get_logger`. Idempotent — safe to
    call multiple times.
    """
    global _initialised
    if _initialised:
        return

    log_dir: Path = CONFIG.log_dir
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "agent.log"

    formatter = logging.Formatter(fmt=_LOG_FORMAT, datefmt=_DATE_FORMAT)

    file_handler = RotatingFileHandler(
        log_path, maxBytes=_MAX_BYTES, backupCount=_BACKUP_COUNT, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    # Console only shows warnings+ to keep the CLI clean; file gets everything.
    console_handler.setLevel(logging.WARNING)

    root = logging.getLogger("excel_agent")
    root.setLevel(getattr(logging, CONFIG.log_level, logging.INFO))
    root.addHandler(file_handler)
    root.addHandler(console_handler)
    root.propagate = False

    _initialised = True


def get_logger(name: str) -> logging.Logger:
    """Return a namespaced child logger.

    Args:
        name: Module name (typically ``__name__``).

    Returns:
        A logger that writes to the rotating file at ``logs/agent.log``.
    """
    _initialise_root()
    # Always nest under the project root so the rotating handler captures it.
    if not name.startswith("excel_agent"):
        name = f"excel_agent.{name}"
    return logging.getLogger(name)
