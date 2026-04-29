"""Query and verify the audit log from the command line.

Usage:
    python audit_query.py --from 2026-04-01 --to 2026-04-29 --event-type formula_written --file deal_model.xlsx
    python audit_query.py --session-id <uuid> --format csv > session_export.csv
    python audit_query.py --verify 2026-04-29

Filters (combine freely):
    --from YYYY-MM-DD       inclusive lower date bound (UTC)
    --to YYYY-MM-DD         inclusive upper date bound (UTC)
    --session-id UUID       events from one session only
    --event-type STR        e.g. formula_written, table_built (repeatable)
    --file STR              substring match on file basename
    --file-hash STR         exact match on the recorded SHA-256
    --source STR            llm | offline_fallback | offline_classifier

Output:
    --format table          aligned text columns (default)
    --format json           one JSON array of events
    --format csv            CSV with all distinct payload keys as columns

Verification:
    --verify YYYY-MM-DD     check the per-session hash chains in that day's
                            file. Exits 0 if intact, 1 if any breaks were
                            detected (and prints them).
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import datetime, timezone
from typing import Any, Iterable

from utils import audit


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="audit_query",
        description="Query and verify the Excel Finance Agent audit log.",
    )
    p.add_argument("--from", dest="date_from", help="Inclusive start date (UTC), YYYY-MM-DD")
    p.add_argument("--to", dest="date_to", help="Inclusive end date (UTC), YYYY-MM-DD")
    p.add_argument("--session-id", help="Filter to a single session UUID")
    p.add_argument(
        "--event-type", action="append", default=[],
        help="Event type filter (repeatable)",
    )
    p.add_argument("--file", help="Substring match on file basename")
    p.add_argument("--file-hash", help="Exact match on file SHA-256 hash")
    p.add_argument(
        "--source",
        choices=["llm", "offline_fallback", "offline_classifier"],
        help="Filter by resolution source",
    )
    p.add_argument(
        "--format", choices=["table", "json", "csv"], default="table",
        help="Output format",
    )
    p.add_argument(
        "--verify", metavar="YYYY-MM-DD",
        help="Run hash-chain verification on a daily file and exit",
    )
    return p.parse_args(argv)


def _events_matching(args: argparse.Namespace) -> Iterable[dict[str, Any]]:
    """Yield events satisfying every CLI filter."""
    today = datetime.now(timezone.utc).date().strftime("%Y-%m-%d")
    date_from = args.date_from or today
    date_to = args.date_to or today

    for ev in audit.iter_events_in_range(date_from, date_to):
        if args.session_id and ev.get("session_id") != args.session_id:
            continue
        if args.event_type and ev.get("event_type") not in args.event_type:
            continue
        payload = ev.get("payload") or {}
        if args.file:
            file_val = str(payload.get("file") or "")
            if args.file.lower() not in file_val.lower():
                continue
        if args.file_hash:
            if payload.get("file_hash") != args.file_hash:
                continue
        if args.source:
            if payload.get("source") != args.source:
                continue
        yield ev


# ──────────────────────────────────────────────────────────────────────
# Output formatters
# ──────────────────────────────────────────────────────────────────────


def _flatten(ev: dict[str, Any]) -> dict[str, Any]:
    flat = {
        "ts": ev.get("ts"),
        "session_id": ev.get("session_id"),
        "seq": ev.get("seq"),
        "event_type": ev.get("event_type"),
    }
    payload = ev.get("payload") or {}
    flat.update(payload)
    return flat


def _format_table(events: list[dict[str, Any]]) -> str:
    """Aligned text columns. Default human-friendly format."""
    if not events:
        return "(no events match)\n"

    rows = [_flatten(ev) for ev in events]
    columns: list[str] = []
    for r in rows:
        for k in r:
            if k not in columns:
                columns.append(k)

    str_rows = [[("" if r.get(c) is None else str(r.get(c))) for c in columns] for r in rows]
    widths = [max(len(c), max((len(r[i]) for r in str_rows), default=0)) for i, c in enumerate(columns)]
    # Cap any single column width to keep the table readable on narrow terminals.
    widths = [min(w, 60) for w in widths]

    def _format_cell(text: str, width: int) -> str:
        return (text[: width - 1] + "…") if len(text) > width else text.ljust(width)

    header = "  ".join(_format_cell(c, w) for c, w in zip(columns, widths))
    sep = "  ".join("-" * w for w in widths)
    body = "\n".join("  ".join(_format_cell(c, w) for c, w in zip(r, widths)) for r in str_rows)
    return f"{header}\n{sep}\n{body}\n"


def _format_json(events: list[dict[str, Any]]) -> str:
    return json.dumps(events, indent=2, ensure_ascii=False)


def _format_csv(events: list[dict[str, Any]]) -> str:
    if not events:
        return ""
    rows = [_flatten(ev) for ev in events]
    columns: list[str] = []
    for r in rows:
        for k in r:
            if k not in columns:
                columns.append(k)
    out = []
    import io
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=columns, extrasaction="ignore")
    writer.writeheader()
    for r in rows:
        writer.writerow(r)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    if args.verify:
        breaks = audit.verify_audit_log(args.verify)
        if not breaks:
            print(f"✓ {args.verify}: hash chain intact across all sessions")
            return 0
        print(f"✗ {args.verify}: {len(breaks)} chain break(s) detected:")
        for b in breaks:
            print(json.dumps(b, indent=2))
        return 1

    events = list(_events_matching(args))

    if args.format == "json":
        sys.stdout.write(_format_json(events))
        sys.stdout.write("\n")
    elif args.format == "csv":
        sys.stdout.write(_format_csv(events))
    else:
        sys.stdout.write(_format_table(events))
    return 0


if __name__ == "__main__":
    sys.exit(main())
