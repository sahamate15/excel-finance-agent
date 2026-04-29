"""Append-only, tamper-evident audit log for the Excel Finance Agent.

Schema (one JSON object per line, daily files):

    logs/audit/YYYY-MM-DD.jsonl

Each line:

    {
      "ts":         "2026-04-29T10:15:23.451Z",
      "session_id": "<uuid>",
      "seq":        42,                         # increments per session
      "event_type": "formula_written",
      "payload":    {...event-specific fields...},
      "prev_hash":  "sha256:..." | null         # SHA-256 of the previous
                                                # event's full JSON line within
                                                # the same session and file.
    }

Hash chain rules:
- ``prev_hash`` is the SHA-256 of the *previous event's complete JSON line*
  (UTF-8 encoded, no trailing newline) within the same session, within the
  same daily file.
- The first event of a session in a given daily file has ``prev_hash: null``.
  When a session spans midnight UTC the chain restarts in the new file.
  Verification therefore operates per-file, per-session.

Compliance constraint:
- Cell *values* from the workbook MUST NEVER appear in any audit payload.
  Only formulas (what was written), target locations (file basename, sheet,
  cell), and the user's instruction text are recorded. File contents are
  identified by SHA-256 hash, never by raw content.
"""

from __future__ import annotations

import hashlib
import json
import threading
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable, Literal

from config import CONFIG

Source = Literal["llm", "offline_fallback"]

_AUDIT_DIR = "audit"
_RETENTION_DAYS = 90  # files older than this get flagged in the archive manifest

# ──────────────────────────────────────────────────────────────────────
# Module-level session state. One Python process == one session by default;
# Streamlit overrides via init_session(uuid). All access goes through _LOCK.
# ──────────────────────────────────────────────────────────────────────

_LOCK = threading.Lock()
_SESSION_ID: str = str(uuid.uuid4())
_SEQ: int = 0
# Per UTC date, the SHA-256 of the last event line we wrote for this session.
# Resets when the date changes so chain verification works per-file.
_CHAIN: dict[str, str | None] = {}


# ──────────────────────────────────────────────────────────────────────
# Session management
# ──────────────────────────────────────────────────────────────────────


def init_session(session_id: str | None = None) -> str:
    """Start a fresh audit session.

    Resets the seq counter and per-date hash-chain state. Streamlit calls
    this once per ``st.session_state`` lifetime; the CLI calls it from
    ``main()``.

    Args:
        session_id: If provided, use this UUID; otherwise generate a fresh one.

    Returns:
        The active session ID.
    """
    global _SESSION_ID, _SEQ, _CHAIN
    with _LOCK:
        _SESSION_ID = session_id or str(uuid.uuid4())
        _SEQ = 0
        _CHAIN = {}
    return _SESSION_ID


def get_session_id() -> str:
    """Return the active session ID."""
    return _SESSION_ID


# ──────────────────────────────────────────────────────────────────────
# Internals — paths, time, hashing, emission
# ──────────────────────────────────────────────────────────────────────


def _audit_dir() -> Path:
    """Return ``logs/audit/`` (creating the directory if needed)."""
    path = CONFIG.log_dir / _AUDIT_DIR
    path.mkdir(parents=True, exist_ok=True)
    return path


def _daily_path(now: datetime | None = None) -> Path:
    """Return ``logs/audit/YYYY-MM-DD.jsonl`` for the given UTC datetime."""
    when = now or datetime.now(timezone.utc)
    if when.tzinfo is None:
        when = when.replace(tzinfo=timezone.utc)
    when = when.astimezone(timezone.utc)
    return _audit_dir() / f"{when.strftime('%Y-%m-%d')}.jsonl"


def _now_iso(now: datetime | None = None) -> str:
    """ISO-8601 UTC timestamp with millisecond precision and trailing Z."""
    n = now or datetime.now(timezone.utc)
    if n.tzinfo is None:
        n = n.replace(tzinfo=timezone.utc)
    n = n.astimezone(timezone.utc)
    return n.strftime("%Y-%m-%dT%H:%M:%S.") + f"{n.microsecond // 1000:03d}Z"


def _date_key(now: datetime | None = None) -> str:
    """Return the UTC date string used as both filename stem and chain key."""
    n = now or datetime.now(timezone.utc)
    if n.tzinfo is None:
        n = n.replace(tzinfo=timezone.utc)
    return n.astimezone(timezone.utc).strftime("%Y-%m-%d")


def _sha256(line: str) -> str:
    """SHA-256 of *line* encoded as UTF-8, returned as ``sha256:<hex>``."""
    return "sha256:" + hashlib.sha256(line.encode("utf-8")).hexdigest()


def file_sha256(path: str | Path) -> str:
    """Return the SHA-256 of a file's contents as ``sha256:<hex>``.

    Used to identify uploaded workbooks in the audit log without recording
    any of their contents. Reads in 64 KB chunks so large files don't
    blow memory.
    """
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return "sha256:" + h.hexdigest()


def _emit(event_type: str, payload: dict[str, Any]) -> dict[str, Any]:
    """Serialize and append one event. Returns the event dict (for testing)."""
    global _SEQ
    with _LOCK:
        now = datetime.now(timezone.utc)
        date_key = _date_key(now)
        prev = _CHAIN.get(date_key)  # None = no previous event today this session
        _SEQ += 1
        event: dict[str, Any] = {
            "ts": _now_iso(now),
            "session_id": _SESSION_ID,
            "seq": _SEQ,
            "event_type": event_type,
            "payload": payload,
            "prev_hash": prev,
        }
        line = json.dumps(event, ensure_ascii=False, separators=(",", ":"))

        path = _daily_path(now)
        with path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")

        # Chain forward: this event's hash becomes the next event's prev_hash.
        _CHAIN[date_key] = _sha256(line)
        return event


# ──────────────────────────────────────────────────────────────────────
# Public event helpers — one per category from the spec.
# ──────────────────────────────────────────────────────────────────────


def record_event(event_type: str, **payload: Any) -> dict[str, Any]:
    """Generic event emitter. Prefer the typed wrappers below where they fit."""
    return _emit(event_type, payload)


# --- Session lifecycle ------------------------------------------------


def record_session_started(*, mode: str, surface: str = "cli") -> dict[str, Any]:
    """First event of a session. Records the surface (cli, streamlit) and mode."""
    return _emit("session_started", {"mode": mode, "surface": surface})


def record_session_ended(*, mode: str) -> dict[str, Any]:
    return _emit("session_ended", {"mode": mode})


def record_mode_changed(*, from_mode: str, to_mode: str) -> dict[str, Any]:
    return _emit("mode_changed", {"from": from_mode, "to": to_mode})


def record_api_key_loaded() -> dict[str, Any]:
    """Records that a key was provided to the session. The key is NEVER logged."""
    return _emit("api_key_loaded", {})


def record_api_key_cleared() -> dict[str, Any]:
    return _emit("api_key_cleared", {})


def record_model_selected(*, model: str) -> dict[str, Any]:
    return _emit("model_selected", {"model": model})


# --- File events ------------------------------------------------------


def record_file_uploaded(*, file: str | Path, size_bytes: int, file_hash: str) -> dict[str, Any]:
    return _emit("file_uploaded", {
        "file": Path(file).name,
        "size_bytes": size_bytes,
        "file_hash": file_hash,
    })


def record_sheet_selected(*, file: str | Path, sheet: str) -> dict[str, Any]:
    return _emit("sheet_selected", {
        "file": Path(file).name,
        "sheet": sheet,
    })


def record_file_downloaded(*, file: str | Path) -> dict[str, Any]:
    return _emit("file_downloaded", {"file": Path(file).name})


# --- Instruction events ----------------------------------------------


def record_instruction_submitted(*, instruction: str) -> dict[str, Any]:
    """The raw (sanitized) instruction. Compliance reviews this; cell values are not."""
    return _emit("instruction_submitted", {"instruction": instruction})


def record_task_type_detected(
    *, task_type: str, source: Literal["llm", "offline_classifier"],
    requires_clarification: bool,
) -> dict[str, Any]:
    return _emit("task_type_detected", {
        "task_type": task_type,
        "source": source,
        "requires_clarification": requires_clarification,
    })


def record_clarification_requested(*, question: str) -> dict[str, Any]:
    return _emit("clarification_requested", {"question": question})


# --- Formula / table events ------------------------------------------


def record_formula_generated(
    *, formula: str, target_cell: str, sheet: str, source: Source,
) -> dict[str, Any]:
    return _emit("formula_generated", {
        "formula": formula,
        "target_cell": target_cell,
        "sheet": sheet,
        "source": source,
    })


def record_dry_run_preview(
    *, sheet: str, cells: int, target_cell: str | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {"sheet": sheet, "cells": cells}
    if target_cell is not None:
        payload["target_cell"] = target_cell
    return _emit("dry_run_preview_shown", payload)


def record_dry_run_confirmed(*, sheet: str) -> dict[str, Any]:
    return _emit("dry_run_confirmed", {"sheet": sheet})


def record_dry_run_cancelled(*, sheet: str) -> dict[str, Any]:
    return _emit("dry_run_cancelled", {"sheet": sheet})


def record_formula_write(
    *, file: str | Path, sheet: str, cell: str, formula: str, source: Source,
) -> dict[str, Any]:
    return _emit("formula_written", {
        "file": Path(file).name,
        "sheet": sheet,
        "cell": cell,
        "formula": formula,
        "source": source,
    })


def record_formula_write_failed(
    *, file: str | Path, sheet: str, cell: str, error: str,
) -> dict[str, Any]:
    return _emit("formula_write_failed", {
        "file": Path(file).name,
        "sheet": sheet,
        "cell": cell,
        "error": error,
    })


def record_table_config_extracted(
    *, table_type: str, source: Source,
) -> dict[str, Any]:
    """Records WHICH table type and from where; numeric inputs are NOT logged."""
    return _emit("table_config_extracted", {
        "table_type": table_type,
        "source": source,
    })


def record_table_built(
    *, file: str | Path, sheet: str, table_type: str, source: Source,
) -> dict[str, Any]:
    return _emit("table_built", {
        "file": Path(file).name,
        "sheet": sheet,
        "table_type": table_type,
        "source": source,
    })


def record_table_build_failed(
    *, file: str | Path, sheet: str, table_type: str, error: str,
) -> dict[str, Any]:
    return _emit("table_build_failed", {
        "file": Path(file).name,
        "sheet": sheet,
        "table_type": table_type,
        "error": error,
    })


# --- Validation / error events ---------------------------------------


def record_input_rejected(*, reason: str, field: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {"reason": reason}
    if field is not None:
        payload["field"] = field
    return _emit("input_rejected", payload)


def record_formula_rejected(*, formula: str, reason: str) -> dict[str, Any]:
    return _emit("formula_rejected_by_validator", {
        "formula": formula,
        "reason": reason,
    })


def record_error(*, action: str, error: str, traceback: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {"action": action, "error": error}
    if traceback is not None:
        payload["traceback"] = traceback
    return _emit("unhandled_exception", payload)


# ──────────────────────────────────────────────────────────────────────
# Reading / verifying
# ──────────────────────────────────────────────────────────────────────


def read_audit_file(date: str | datetime | Path) -> list[dict[str, Any]]:
    """Read all events from a daily audit file in original write order.

    Args:
        date: Either ``"YYYY-MM-DD"``, a datetime, or a direct Path to a
            ``.jsonl`` file (useful for tests).
    """
    if isinstance(date, Path):
        path = date
    elif isinstance(date, datetime):
        path = _daily_path(date)
    else:
        # Parse "YYYY-MM-DD" into a UTC midnight datetime.
        dt = datetime.strptime(date, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        path = _daily_path(dt)

    if not path.exists():
        return []

    events: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            try:
                events.append(json.loads(line))
            except json.JSONDecodeError as exc:
                events.append({"_parse_error": str(exc), "_raw": line})
    return events


def verify_audit_log(date: str | datetime | Path) -> list[dict[str, Any]]:
    """Verify the per-session hash chains in a daily audit file.

    Returns a list of "break" records. An empty list means every chain is
    intact. Each break is a dict like::

        {
          "session_id": "...",
          "seq": 12,
          "reason": "prev_hash_mismatch" | "parse_error" | "missing_seq" | "duplicate_seq",
          "expected": "sha256:...",
          "got": "sha256:..."
        }

    Args:
        date: ``"YYYY-MM-DD"`` string, datetime, or direct Path.
    """
    breaks: list[dict[str, Any]] = []

    # Resolve to a Path so we can re-read raw lines for hashing.
    if isinstance(date, Path):
        path = date
    elif isinstance(date, datetime):
        path = _daily_path(date)
    else:
        dt = datetime.strptime(date, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        path = _daily_path(dt)

    if not path.exists():
        return breaks

    # We need both the parsed event and the original line bytes for hashing.
    raw_lines: list[str] = []
    parsed: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as f:
        for i, raw in enumerate(f, start=1):
            raw = raw.rstrip("\n")
            if not raw.strip():
                continue
            raw_lines.append(raw)
            try:
                parsed.append(json.loads(raw))
            except json.JSONDecodeError as exc:
                breaks.append({
                    "session_id": None,
                    "seq": None,
                    "line_number": i,
                    "reason": "parse_error",
                    "error": str(exc),
                })
                parsed.append(None)  # type: ignore[arg-type]

    # Group by session, preserving each session's order of appearance.
    sessions: dict[str, list[tuple[int, str, dict[str, Any]]]] = {}
    for idx, ev in enumerate(parsed):
        if ev is None:
            continue
        sid = ev.get("session_id") or "<missing>"
        sessions.setdefault(sid, []).append((idx, raw_lines[idx], ev))

    for sid, items in sessions.items():
        prev_hash: str | None = None
        seen_seqs: set[int] = set()
        last_seq = 0
        for line_idx, raw, ev in items:
            seq = ev.get("seq")
            if not isinstance(seq, int):
                breaks.append({
                    "session_id": sid, "seq": seq, "line_number": line_idx + 1,
                    "reason": "missing_seq",
                })
                continue
            if seq in seen_seqs:
                breaks.append({
                    "session_id": sid, "seq": seq, "line_number": line_idx + 1,
                    "reason": "duplicate_seq",
                })
                continue
            seen_seqs.add(seq)
            # seq monotonicity within a session in this file
            if seq <= last_seq:
                breaks.append({
                    "session_id": sid, "seq": seq, "line_number": line_idx + 1,
                    "reason": "non_monotonic_seq",
                    "expected_min": last_seq + 1, "got": seq,
                })
            last_seq = max(last_seq, seq)

            stored_prev = ev.get("prev_hash")
            if stored_prev != prev_hash:
                breaks.append({
                    "session_id": sid, "seq": seq, "line_number": line_idx + 1,
                    "reason": "prev_hash_mismatch",
                    "expected": prev_hash,
                    "got": stored_prev,
                })
            # Advance the chain regardless so we can flag every divergence.
            prev_hash = _sha256(raw)

    return breaks


# ──────────────────────────────────────────────────────────────────────
# Querying
# ──────────────────────────────────────────────────────────────────────


def iter_events_in_range(
    start_date: str,
    end_date: str,
) -> Iterable[dict[str, Any]]:
    """Yield events from every daily file in ``[start_date, end_date]`` inclusive.

    Dates are ``YYYY-MM-DD`` UTC strings. Files that don't exist are skipped
    silently — the query CLI should report missing days separately if needed.
    """
    start = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    end = datetime.strptime(end_date, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    if end < start:
        return
    cur = start
    while cur <= end:
        for ev in read_audit_file(cur):
            yield ev
        cur += timedelta(days=1)


# ──────────────────────────────────────────────────────────────────────
# Retention manifest (no actual deletion — flag-only)
# ──────────────────────────────────────────────────────────────────────


def write_archive_manifest(retention_days: int = _RETENTION_DAYS) -> dict[str, Any]:
    """Scan ``logs/audit/`` and list daily files older than ``retention_days``.

    Writes ``logs/audit/archive_manifest.json``. Does NOT move or delete any
    files — that is left to the firm's archival pipeline.

    Returns the manifest dict.
    """
    cutoff = (datetime.now(timezone.utc) - timedelta(days=retention_days)).date()
    candidates: list[dict[str, Any]] = []
    audit_root = _audit_dir()
    for p in sorted(audit_root.glob("*.jsonl")):
        # Filename stem is YYYY-MM-DD by construction.
        try:
            file_date = datetime.strptime(p.stem, "%Y-%m-%d").date()
        except ValueError:
            continue
        if file_date < cutoff:
            candidates.append({
                "file": p.name,
                "date": p.stem,
                "size_bytes": p.stat().st_size,
                "sha256": file_sha256(p),
            })

    manifest = {
        "generated_at": _now_iso(),
        "retention_days": retention_days,
        "cutoff_date": cutoff.strftime("%Y-%m-%d"),
        "ready_to_archive": candidates,
    }
    out_path = audit_root / "archive_manifest.json"
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    return manifest


__all__ = [
    "Source",
    # Session
    "init_session", "get_session_id",
    # Generic
    "record_event", "file_sha256",
    # Session lifecycle
    "record_session_started", "record_session_ended",
    "record_mode_changed", "record_api_key_loaded", "record_api_key_cleared",
    "record_model_selected",
    # File
    "record_file_uploaded", "record_sheet_selected", "record_file_downloaded",
    # Instruction
    "record_instruction_submitted", "record_task_type_detected",
    "record_clarification_requested",
    # Formula / table
    "record_formula_generated", "record_dry_run_preview",
    "record_dry_run_confirmed", "record_dry_run_cancelled",
    "record_formula_write", "record_formula_write_failed",
    "record_table_config_extracted", "record_table_built",
    "record_table_build_failed",
    # Validation / error
    "record_input_rejected", "record_formula_rejected", "record_error",
    # Reading / verifying / querying
    "read_audit_file", "verify_audit_log", "iter_events_in_range",
    # Retention
    "write_archive_manifest",
]
