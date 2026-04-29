"""CLI entry point for the Excel Finance Automation Agent.

Run with: ``python main.py``

Type ``help`` inside the prompt for sample instructions, ``file <path>`` to
switch active workbook, ``sheet <name>`` to switch the sheet, ``preview`` to
peek at the last few rows, and ``exit`` / ``quit`` to leave.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pandas as pd
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt
from rich.table import Table
from rich.text import Text

from agents.task_executor import execute_excel_task
from config import CONFIG
from utils import audit
from utils.logger import get_logger

logger = get_logger(__name__)
console = Console()

EXAMPLES: dict[str, list[str]] = {
    "Formulas": [
        "calculate growth rate between B2 and B3",
        "what is the CAGR from B2 to B6 over 4 years",
        "calculate IRR for cash flows in B2:B7",
        "NPV at 12% discount rate for B3:B7",
        "EMI for 50 lakh loan at 8.5% for 20 years",
        "net profit margin if revenue is in B2 and profit in C2",
        "WDV depreciation for year 3 on asset value in B2 at 25%",
    ],
    "Tables": [
        "create a 5-year WDV depreciation schedule for 20 lakh at 25%",
        "build loan amortization table for 50 lakh at 9% for 10 years",
        "make a 5-year revenue projection starting at 50 lakh with 15% growth",
        "create EBITDA table for years 2020–2024 using Sheet Revenue data",
    ],
}


def _print_banner() -> None:
    """Display the welcome banner."""
    banner = Text()
    banner.append("💹 Excel Finance Automation Agent\n", style="bold cyan")
    banner.append("Type 'help' for examples · 'file <path>' to switch file\n", style="dim")
    banner.append("Type 'sheet <name>' to switch sheet · 'preview' to peek\n", style="dim")
    banner.append("Type 'exit' to quit", style="dim")
    console.print(Panel(banner, border_style="cyan"))


def _print_help() -> None:
    """Show example instructions grouped by category."""
    for category, items in EXAMPLES.items():
        table = Table(title=category, show_header=False, border_style="blue")
        table.add_column("Example", style="white")
        for ex in items:
            table.add_row(f"• {ex}")
        console.print(table)


def _print_result(result: dict[str, Any], filepath: Path) -> None:
    """Render a result dict from the orchestrator."""
    table = Table(show_header=False, border_style="green" if result["success"] else "red")
    table.add_column("k", style="bold")
    table.add_column("v")
    icon = "✅" if result["success"] else "❌"
    table.add_row(f"{icon} Task Type", str(result.get("task_type") or "—"))
    if result.get("formula"):
        table.add_row("📐 Formula", str(result["formula"]))
    if result.get("cell_written"):
        table.add_row("📍 Written to", str(result["cell_written"]))
    if result.get("table_created"):
        table.add_row("📊 Table", "created")
    table.add_row("📁 File", str(filepath.name))
    table.add_row("💬 Message", str(result.get("message") or ""))
    if result.get("error"):
        table.add_row("⚠️  Error", str(result["error"]))
    console.print(Panel(table, title="Result", border_style="green" if result["success"] else "red"))


def _preview(filepath: Path, sheet_name: str, n: int = 5) -> None:
    """Print the last *n* rows of *sheet_name* as a pandas table."""
    if not filepath.exists():
        console.print(f"[yellow]File does not exist yet:[/yellow] {filepath}")
        return
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name)
    except (ValueError, FileNotFoundError) as exc:
        console.print(f"[red]Cannot preview {sheet_name}: {exc}[/red]")
        return
    console.print(Panel(str(df.tail(n)), title=f"{filepath.name} :: {sheet_name} (tail {n})"))


def main() -> int:
    """Run the interactive REPL.

    Returns:
        Process exit code: 0 on graceful exit, 1 on fatal startup error.
    """
    audit.init_session()
    audit.record_session_started(
        mode="strict" if CONFIG.strict_mode else "ai_assisted",
        surface="cli",
    )
    _print_banner()

    active_file: Path = CONFIG.default_excel_file
    active_sheet: str = CONFIG.default_sheet

    if not active_file.exists():
        console.print(
            f"[yellow]Note:[/yellow] default workbook {active_file} does not exist. "
            "Run [bold]python generate_sample_workbook.py[/bold] to create it."
        )

    while True:
        try:
            user_input = Prompt.ask("[bold cyan]agent[/bold cyan]").strip()
        except (EOFError, KeyboardInterrupt):
            console.print("\n[dim]Goodbye.[/dim]")
            return 0

        if not user_input:
            continue

        lower = user_input.lower()

        if lower in {"exit", "quit"}:
            audit.record_session_ended(mode="strict" if CONFIG.strict_mode else "ai_assisted")
            console.print("[dim]Goodbye.[/dim]")
            return 0

        if lower == "help":
            _print_help()
            continue

        if lower.startswith("file "):
            new_path = Path(user_input.split(" ", 1)[1]).expanduser().resolve()
            active_file = new_path
            console.print(f"[green]Active file →[/green] {active_file}")
            continue

        if lower.startswith("sheet "):
            active_sheet = user_input.split(" ", 1)[1].strip()
            console.print(f"[green]Active sheet →[/green] {active_sheet}")
            continue

        if lower == "preview":
            _preview(active_file, active_sheet)
            continue

        # Treat anything else as an instruction.
        try:
            result = execute_excel_task(
                instruction=user_input,
                filepath=active_file,
                sheet_name=active_sheet,
                cell=None,
            )
        except Exception as exc:  # noqa: BLE001 — REPL must keep running
            logger.exception("Unhandled error in REPL")
            console.print(f"[red]Unhandled error:[/red] {exc}")
            continue

        _print_result(result, active_file)


if __name__ == "__main__":
    sys.exit(main())
