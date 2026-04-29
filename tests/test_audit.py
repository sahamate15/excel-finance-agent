"""Tests for the Step 6 audit logging system.

Covers:
- Event-write integrity (schema, daily file, append-only)
- Hash chain construction and verification
- Tamper detection (modified payload, swapped lines, deleted line)
- Query filtering (date range, session, event type, file, source)
- CSV export shape
- Integration: full upload → instruct → preview → confirm → write → download
  workflow with an intact chain across the expected event types.
"""

from __future__ import annotations

import importlib
import json
import sys
import uuid
from pathlib import Path

import pytest


# Make project root importable.
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from utils import audit  # noqa: E402
from config import CONFIG  # noqa: E402


@pytest.fixture
def isolated_audit(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Redirect audit writes to a tmp dir and start a fresh session."""
    monkeypatch.setattr(audit.CONFIG, "log_dir", tmp_path)
    audit.init_session()
    return tmp_path / "audit"


# ──────────────────────────────────────────────────────────────────────
# Event-write integrity
# ──────────────────────────────────────────────────────────────────────


class TestEventWriteIntegrity:
    """Schema, daily-file location, append-only behaviour."""

    def test_event_schema(self, isolated_audit: Path) -> None:
        ev = audit.record_formula_write(
            file="deal_model.xlsx", sheet="Revenue", cell="B2",
            formula="=A1+1", source="offline_fallback",
        )
        assert ev["event_type"] == "formula_written"
        assert ev["session_id"] == audit.get_session_id()
        assert ev["seq"] == 1
        assert "ts" in ev
        assert ev["payload"]["file"] == "deal_model.xlsx"
        assert ev["payload"]["formula"] == "=A1+1"
        assert ev["prev_hash"] is None  # first event of the session

    def test_writes_to_daily_file(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="x.xlsx", sheet="S", cell="A1",
            formula="=1", source="offline_fallback",
        )
        files = sorted(isolated_audit.glob("*.jsonl"))
        assert len(files) == 1
        # Filename is YYYY-MM-DD.jsonl (UTC).
        assert len(files[0].stem) == 10
        assert files[0].stem.count("-") == 2

    def test_seq_increments(self, isolated_audit: Path) -> None:
        e1 = audit.record_formula_write(
            file="f", sheet="s", cell="A1", formula="=1", source="offline_fallback",
        )
        e2 = audit.record_formula_write(
            file="f", sheet="s", cell="A2", formula="=2", source="offline_fallback",
        )
        e3 = audit.record_table_built(
            file="f", sheet="s", table_type="depreciation", source="offline_fallback",
        )
        assert (e1["seq"], e2["seq"], e3["seq"]) == (1, 2, 3)

    def test_init_session_resets_seq(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="f", sheet="s", cell="A1", formula="=1", source="offline_fallback",
        )
        audit.init_session()
        ev = audit.record_formula_write(
            file="f", sheet="s", cell="A2", formula="=2", source="offline_fallback",
        )
        assert ev["seq"] == 1
        assert ev["prev_hash"] is None  # fresh session = fresh chain

    def test_payload_strips_path_to_basename(self, isolated_audit: Path) -> None:
        ev = audit.record_formula_write(
            file="/clients/Acme/q4/deal_model.xlsx",
            sheet="S", cell="A1", formula="=1", source="llm",
        )
        assert ev["payload"]["file"] == "deal_model.xlsx"

    def test_no_cell_values_in_payload(self, isolated_audit: Path) -> None:
        """Compliance invariant: payload must never carry cell values."""
        ev = audit.record_formula_write(
            file="f", sheet="s", cell="B2",
            formula="=AVERAGE(B2:B100)", source="offline_fallback",
        )
        # Sanity: the only "value-like" data in the payload is the formula
        # string itself, which IS allowed (it describes what was written).
        # No "previous_value", "old_value", "cell_contents", etc. keys exist.
        forbidden = {"value", "previous_value", "old_value", "cell_contents", "raw_value"}
        assert not (set(ev["payload"].keys()) & forbidden)


# ──────────────────────────────────────────────────────────────────────
# Hash chain construction
# ──────────────────────────────────────────────────────────────────────


class TestHashChain:
    """Each event's prev_hash links to the previous event's full JSON line."""

    def test_chain_links_consecutive_events(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="f", sheet="s", cell="A1", formula="=1", source="offline_fallback",
        )
        audit.record_formula_write(
            file="f", sheet="s", cell="A2", formula="=2", source="offline_fallback",
        )
        audit.record_formula_write(
            file="f", sheet="s", cell="A3", formula="=3", source="offline_fallback",
        )

        breaks = audit.verify_audit_log(_today_str())
        assert breaks == [], f"chain should be intact, got: {breaks}"

    def test_first_event_has_null_prev_hash(self, isolated_audit: Path) -> None:
        ev = audit.record_formula_write(
            file="f", sheet="s", cell="A1", formula="=1", source="offline_fallback",
        )
        assert ev["prev_hash"] is None

    def test_subsequent_events_have_non_null_prev_hash(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="f", sheet="s", cell="A1", formula="=1", source="offline_fallback",
        )
        ev = audit.record_formula_write(
            file="f", sheet="s", cell="A2", formula="=2", source="offline_fallback",
        )
        assert ev["prev_hash"] is not None
        assert ev["prev_hash"].startswith("sha256:")


# ──────────────────────────────────────────────────────────────────────
# Tamper detection
# ──────────────────────────────────────────────────────────────────────


class TestTamperDetection:
    """verify_audit_log must catch silent edits to any persisted line."""

    def test_modified_payload_detected(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="x", sheet="s", cell="A1", formula="=1", source="offline_fallback",
        )
        audit.record_formula_write(
            file="x", sheet="s", cell="A2", formula="=2", source="offline_fallback",
        )
        audit.record_formula_write(
            file="x", sheet="s", cell="A3", formula="=3", source="offline_fallback",
        )

        # Verify clean first.
        assert audit.verify_audit_log(_today_str()) == []

        # Now corrupt the second event's formula.
        path = isolated_audit / f"{_today_str()}.jsonl"
        lines = path.read_text(encoding="utf-8").splitlines()
        ev = json.loads(lines[1])
        ev["payload"]["formula"] = "=999"  # silent edit
        lines[1] = json.dumps(ev, ensure_ascii=False, separators=(",", ":"))
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")

        # Verification should now report a chain break on event 3 (whose
        # prev_hash no longer matches the modified line 2).
        breaks = audit.verify_audit_log(_today_str())
        assert len(breaks) >= 1
        assert any(b["reason"] == "prev_hash_mismatch" for b in breaks)
        # The break should land on seq 3 (the event whose prev_hash no longer matches).
        assert any(b.get("seq") == 3 for b in breaks)

    def test_swapped_lines_detected(self, isolated_audit: Path) -> None:
        for i in range(1, 5):
            audit.record_formula_write(
                file="x", sheet="s", cell=f"A{i}", formula=f"={i}", source="offline_fallback",
            )

        path = isolated_audit / f"{_today_str()}.jsonl"
        lines = path.read_text(encoding="utf-8").splitlines()
        # Swap lines 1 and 2 (0-indexed: lines[1] and lines[2]).
        lines[1], lines[2] = lines[2], lines[1]
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")

        breaks = audit.verify_audit_log(_today_str())
        assert len(breaks) >= 1
        # Either a prev_hash_mismatch or a non_monotonic_seq will fire here.
        assert any(b["reason"] in {"prev_hash_mismatch", "non_monotonic_seq"} for b in breaks)

    def test_deleted_line_detected(self, isolated_audit: Path) -> None:
        for i in range(1, 5):
            audit.record_formula_write(
                file="x", sheet="s", cell=f"A{i}", formula=f"={i}", source="offline_fallback",
            )

        path = isolated_audit / f"{_today_str()}.jsonl"
        lines = path.read_text(encoding="utf-8").splitlines()
        # Drop line 2 (0-indexed lines[1]).
        del lines[1]
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")

        breaks = audit.verify_audit_log(_today_str())
        assert len(breaks) >= 1


# ──────────────────────────────────────────────────────────────────────
# Query filtering
# ──────────────────────────────────────────────────────────────────────


class TestQueryFiltering:
    """iter_events_in_range honours date bounds; CLI _events_matching honours all filters."""

    def test_iter_events_today(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="a", sheet="s", cell="A1", formula="=1", source="llm",
        )
        audit.record_table_built(
            file="b", sheet="s", table_type="projection", source="offline_fallback",
        )

        events = list(audit.iter_events_in_range(_today_str(), _today_str()))
        assert len(events) == 2

    def test_session_filter_via_cli(
        self, isolated_audit: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        # Two sessions writing to the same daily file.
        sid_a = audit.init_session(str(uuid.uuid4()))
        audit.record_formula_write(
            file="a", sheet="s", cell="A1", formula="=1", source="llm",
        )
        sid_b = audit.init_session(str(uuid.uuid4()))
        audit.record_formula_write(
            file="b", sheet="s", cell="B1", formula="=2", source="offline_fallback",
        )

        # Re-import audit_query AFTER the audit module's CONFIG.log_dir has been
        # monkeypatched, since the CLI's `from utils import audit` was bound
        # at import time on the same module instance.
        if "audit_query" in sys.modules:
            del sys.modules["audit_query"]
        import audit_query  # noqa: PLC0415

        events = list(audit_query._events_matching(_args(
            date_from=_today_str(), date_to=_today_str(), session_id=sid_a,
        )))
        assert len(events) == 1
        assert events[0]["session_id"] == sid_a

        events_b = list(audit_query._events_matching(_args(
            date_from=_today_str(), date_to=_today_str(), session_id=sid_b,
        )))
        assert len(events_b) == 1
        assert events_b[0]["session_id"] == sid_b

    def test_event_type_filter(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="x", sheet="s", cell="A1", formula="=1", source="llm",
        )
        audit.record_table_built(
            file="y", sheet="s", table_type="depreciation", source="offline_fallback",
        )
        audit.record_formula_write(
            file="x", sheet="s", cell="A2", formula="=2", source="offline_fallback",
        )

        if "audit_query" in sys.modules:
            del sys.modules["audit_query"]
        import audit_query  # noqa: PLC0415

        events = list(audit_query._events_matching(_args(
            date_from=_today_str(), date_to=_today_str(),
            event_type=["formula_written"],
        )))
        assert len(events) == 2
        assert all(e["event_type"] == "formula_written" for e in events)

    def test_file_substring_filter(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="acme_q4_deal.xlsx", sheet="s", cell="A1", formula="=1", source="llm",
        )
        audit.record_formula_write(
            file="other.xlsx", sheet="s", cell="A2", formula="=2", source="llm",
        )

        if "audit_query" in sys.modules:
            del sys.modules["audit_query"]
        import audit_query  # noqa: PLC0415

        events = list(audit_query._events_matching(_args(
            date_from=_today_str(), date_to=_today_str(), file="acme",
        )))
        assert len(events) == 1
        assert events[0]["payload"]["file"] == "acme_q4_deal.xlsx"

    def test_source_filter(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="x", sheet="s", cell="A1", formula="=1", source="llm",
        )
        audit.record_formula_write(
            file="x", sheet="s", cell="A2", formula="=2", source="offline_fallback",
        )

        if "audit_query" in sys.modules:
            del sys.modules["audit_query"]
        import audit_query  # noqa: PLC0415

        events = list(audit_query._events_matching(_args(
            date_from=_today_str(), date_to=_today_str(), source="llm",
        )))
        assert len(events) == 1
        assert events[0]["payload"]["source"] == "llm"


# ──────────────────────────────────────────────────────────────────────
# Export format correctness
# ──────────────────────────────────────────────────────────────────────


class TestExportFormat:
    """CSV export carries every distinct payload key; JSON is parseable."""

    def test_csv_has_payload_columns(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="x", sheet="s", cell="A1", formula="=1", source="llm",
        )
        audit.record_table_built(
            file="x", sheet="s", table_type="projection", source="llm",
        )

        if "audit_query" in sys.modules:
            del sys.modules["audit_query"]
        import audit_query  # noqa: PLC0415

        events = list(audit.iter_events_in_range(_today_str(), _today_str()))
        csv_text = audit_query._format_csv(events)
        header = csv_text.splitlines()[0].split(",")
        # Top-level keys + every distinct payload key must be present.
        for col in ["ts", "session_id", "seq", "event_type", "file", "sheet", "formula", "table_type", "source"]:
            assert col in header

    def test_json_round_trip(self, isolated_audit: Path) -> None:
        audit.record_formula_write(
            file="x", sheet="s", cell="A1", formula="=AVG(B2:B10)", source="llm",
        )
        if "audit_query" in sys.modules:
            del sys.modules["audit_query"]
        import audit_query  # noqa: PLC0415

        events = list(audit.iter_events_in_range(_today_str(), _today_str()))
        text = audit_query._format_json(events)
        parsed = json.loads(text)
        assert isinstance(parsed, list)
        assert parsed[0]["payload"]["formula"] == "=AVG(B2:B10)"


# ──────────────────────────────────────────────────────────────────────
# Integration: full workflow end-to-end
# ──────────────────────────────────────────────────────────────────────


class TestWorkflowIntegration:
    """A complete UI-equivalent workflow records the expected events in order."""

    def test_full_dry_run_then_write_workflow(
        self, isolated_audit: Path, tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """upload → select sheet → instruct → preview → confirm → write → download."""
        from agents import task_executor  # noqa: PLC0415

        # Stub the LLM-backed pieces so the workflow is deterministic.
        monkeypatch.setattr(task_executor, "detect_task_type", task_executor._keyword_classify)
        monkeypatch.setattr(
            task_executor, "extract_table_config",
            lambda instr: task_executor._heuristic_table_config(instr, strict=False),
        )

        # 1. Simulate file upload (creates the file + records event).
        wb_path = tmp_path / "deal_model.xlsx"
        wb_path.write_bytes(b"PK\x03\x04 not really an xlsx")  # dummy bytes for hashing
        file_hash = audit.file_sha256(wb_path)
        audit.record_file_uploaded(
            file=wb_path, size_bytes=wb_path.stat().st_size, file_hash=file_hash,
        )

        # 2. Simulate sheet selection.
        audit.record_sheet_selected(file=wb_path, sheet="Revenue")

        # 3. Replace the dummy bytes with a real xlsx so the orchestrator can write.
        from openpyxl import Workbook  # noqa: PLC0415
        wb = Workbook()
        wb["Sheet"].title = "Revenue"
        wb.save(wb_path)

        # 4. Submit the instruction in dry-run mode (orchestrator records
        #    instruction + task_type + formula_generated + dry_run_preview).
        r1 = task_executor.execute_excel_task(
            "calculate IRR for B2:B10",
            filepath=wb_path, sheet_name="Revenue", dry_run=True,
        )
        assert r1["success"] and r1["dry_run"]

        # 5. Simulate the user clicking "Confirm and write".
        audit.record_dry_run_confirmed(sheet="Revenue")

        # 6. Re-run with dry_run=False to actually write.
        r2 = task_executor.execute_excel_task(
            "calculate IRR for B2:B10",
            filepath=wb_path, sheet_name="Revenue", dry_run=False,
        )
        assert r2["success"] and not r2["dry_run"]

        # 7. Simulate download.
        audit.record_file_downloaded(file=wb_path)

        # Now verify the chain on today's audit file.
        breaks = audit.verify_audit_log(_today_str())
        assert breaks == [], f"chain breaks in integration test: {breaks}"

        # Confirm the expected event types appear, in order.
        events = audit.read_audit_file(_today_str())
        types_in_order = [e["event_type"] for e in events]

        # Each of these must appear, and in this relative order.
        expected_in_order = [
            "file_uploaded",
            "sheet_selected",
            "instruction_submitted",
            "task_type_detected",
            "formula_generated",
            "dry_run_preview_shown",
            "dry_run_confirmed",
            "instruction_submitted",      # second pass, dry_run=False
            "task_type_detected",
            "formula_generated",
            "formula_written",
            "file_downloaded",
        ]
        last_index = -1
        for et in expected_in_order:
            try:
                last_index = types_in_order.index(et, last_index + 1)
            except ValueError as exc:
                pytest.fail(f"Expected event {et!r} not found after index {last_index}: {types_in_order}")

        # And no cell-value leakage anywhere.
        for ev in events:
            payload_str = json.dumps(ev.get("payload") or {})
            assert "old_value" not in payload_str
            assert "previous_value" not in payload_str


# ──────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────


def _today_str() -> str:
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).date().strftime("%Y-%m-%d")


def _args(**kwargs):
    """Build a Namespace mirroring audit_query's argparse output."""
    import argparse
    defaults = dict(
        date_from=None, date_to=None, session_id=None,
        event_type=[], file=None, file_hash=None, source=None,
        format="table", verify=None,
    )
    defaults.update(kwargs)
    return argparse.Namespace(**defaults)
