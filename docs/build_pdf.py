#!/usr/bin/env python3
"""Build the Excel Finance Agent user guide as a professional PDF.

Reads ../GETTING_STARTED.md and produces ./Excel_Finance_Agent_User_Guide.pdf
via a two-pass build (pass 1 collects page numbers for the TOC; pass 2 emits
the final document).

Adapted from the project's standard ReportLab builder. Uses ReportLab's
built-in Helvetica family so it runs on macOS, Linux, and Windows without
any font installation. Run from this directory:

    pip install reportlab pypdf
    python docs/build_pdf.py
"""

import os
import re
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    HRFlowable,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


# ──────────────────────────────────────────────────────────────────────
# Fonts: ReportLab ships Helvetica + Courier as Standard 14 fonts that
# need no registration and render identically on every platform.
# ──────────────────────────────────────────────────────────────────────

F = "Helvetica"
FB = "Helvetica-Bold"
FI = "Helvetica-Oblique"
FBI = "Helvetica-BoldOblique"
FM = "Courier"


# ──────────────────────────────────────────────────────────────────────
# Colour palette (matches the project's house style)
# ──────────────────────────────────────────────────────────────────────

DARK_NAVY = HexColor("#1a2332")
TEAL = HexColor("#0d7377")
DARK_GRAY = HexColor("#2d3436")
MED_GRAY = HexColor("#636e72")
BORDER_GRAY = HexColor("#dfe6e9")
TBL_HDR_BG = HexColor("#1a2332")
TBL_ALT_BG = HexColor("#f8f9fa")
CALLOUT_BG = HexColor("#f0f8f8")


PAGE_W, PAGE_H = A4
LM = 22 * mm
RM = 22 * mm
TM = 25 * mm
BM = 25 * mm
CW = PAGE_W - LM - RM


# ──────────────────────────────────────────────────────────────────────
# Document plumbing
# ──────────────────────────────────────────────────────────────────────


class TOCAnchor(Flowable):
    """Invisible flowable that records its page number for the TOC."""

    width = 0
    height = 0

    def __init__(self, key: str, level: int, text: str) -> None:
        Flowable.__init__(self)
        self.key = key
        self.level = level
        self.text = text

    def draw(self) -> None:  # noqa: D401
        pass


class DocTemplate(BaseDocTemplate):
    def __init__(self, filename: str, **kw) -> None:
        super().__init__(filename, **kw)
        frame_cover = Frame(LM, BM, CW, PAGE_H - TM - BM, id="cover")
        frame_body = Frame(LM, BM + 8 * mm, CW, PAGE_H - TM - BM - 8 * mm, id="body")
        self.addPageTemplates([
            PageTemplate(id="cover", frames=[frame_cover], onPage=self._cover_page),
            PageTemplate(id="body", frames=[frame_body], onPage=self._body_page),
        ])
        self.toc_entries: list[tuple[int, str, int]] = []

    def _cover_page(self, c, doc) -> None:
        c.saveState()
        c.setFillColor(TEAL)
        c.rect(0, PAGE_H - 8 * mm, PAGE_W, 8 * mm, fill=1, stroke=0)
        c.setFillColor(DARK_NAVY)
        c.rect(0, 0, PAGE_W, 18 * mm, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont(F, 8)
        c.drawString(LM, 8 * mm, "Excel Finance Agent  |  User Guide  |  April 2026")
        c.restoreState()

    def _body_page(self, c, doc) -> None:
        c.saveState()
        pn = doc.page
        c.setStrokeColor(TEAL)
        c.setLineWidth(0.5)
        c.line(LM, PAGE_H - TM + 6 * mm, PAGE_W - RM, PAGE_H - TM + 6 * mm)
        c.setFont(F, 7)
        c.setFillColor(MED_GRAY)
        c.drawString(LM, PAGE_H - TM + 9 * mm, "Excel Finance Agent  |  A Simple User Guide")
        c.drawRightString(PAGE_W - RM, PAGE_H - TM + 9 * mm, f"Page {pn}")
        c.line(LM, BM + 4 * mm, PAGE_W - RM, BM + 4 * mm)
        c.drawString(LM, BM, "github.com/sahamate15/excel-finance-agent")
        c.restoreState()

    def afterFlowable(self, flowable) -> None:
        if isinstance(flowable, TOCAnchor):
            self.toc_entries.append((flowable.level, flowable.text, self.page))


# ──────────────────────────────────────────────────────────────────────
# Paragraph styles
# ──────────────────────────────────────────────────────────────────────


def make_styles() -> dict:
    s: dict = {}
    s["cover_title"] = ParagraphStyle("CT", fontName=FB, fontSize=28, leading=36, textColor=DARK_NAVY)
    s["cover_sub"] = ParagraphStyle("CS", fontName=F, fontSize=14, leading=20, textColor=TEAL)
    s["cover_author"] = ParagraphStyle("CA", fontName=F, fontSize=11, leading=16, textColor=DARK_GRAY)
    s["cover_meta"] = ParagraphStyle("CM", fontName=F, fontSize=10, leading=15, textColor=MED_GRAY)

    s["part"] = ParagraphStyle(
        "Part", fontName=FB, fontSize=18, leading=24,
        textColor=DARK_NAVY, spaceBefore=14 * mm, spaceAfter=5 * mm,
    )
    s["section"] = ParagraphStyle(
        "Sec", fontName=FB, fontSize=13, leading=18,
        textColor=TEAL, spaceBefore=8 * mm, spaceAfter=3 * mm,
    )
    s["subsection"] = ParagraphStyle(
        "Sub", fontName=FB, fontSize=11, leading=15.5,
        textColor=DARK_NAVY, spaceBefore=6 * mm, spaceAfter=2.5 * mm,
    )
    s["h2_other"] = ParagraphStyle(
        "H2O", fontName=FB, fontSize=16, leading=22,
        textColor=DARK_NAVY, spaceBefore=12 * mm, spaceAfter=5 * mm,
    )

    s["body"] = ParagraphStyle(
        "Body", fontName=F, fontSize=9.5, leading=14.5,
        textColor=DARK_GRAY, spaceAfter=3 * mm, alignment=TA_JUSTIFY,
    )
    s["bullet"] = ParagraphStyle(
        "Bul", fontName=F, fontSize=9.5, leading=14.5,
        textColor=DARK_GRAY, spaceAfter=1.5 * mm, leftIndent=14, bulletIndent=0,
    )
    s["code"] = ParagraphStyle(
        "Code", fontName=FM, fontSize=7.5, leading=10.5,
        textColor=DARK_GRAY, backColor=HexColor("#f8f9fa"),
        borderColor=BORDER_GRAY, borderWidth=0.5, borderPadding=6,
        spaceAfter=3 * mm, leftIndent=4, rightIndent=4,
    )
    s["blockquote"] = ParagraphStyle(
        "BQ", fontName=FI, fontSize=9.5, leading=14,
        textColor=MED_GRAY, leftIndent=14, rightIndent=14,
        borderColor=TEAL, borderWidth=0, borderPadding=4,
        spaceAfter=3 * mm,
    )

    s["tbl_hdr"] = ParagraphStyle("TH", fontName=FB, fontSize=8, leading=11, textColor=white)
    s["tbl_cell"] = ParagraphStyle("TC", fontName=F, fontSize=8, leading=11, textColor=DARK_GRAY)

    s["toc_title"] = ParagraphStyle(
        "TT", fontName=FB, fontSize=18, leading=24,
        textColor=DARK_NAVY, spaceAfter=8 * mm,
    )
    s["toc_part"] = ParagraphStyle(
        "TP", fontName=FB, fontSize=10, leading=18, textColor=DARK_NAVY, leftIndent=0,
    )
    s["toc_sec"] = ParagraphStyle(
        "TS", fontName=F, fontSize=9, leading=15, textColor=DARK_GRAY, leftIndent=10,
    )
    s["toc_sub"] = ParagraphStyle(
        "TSS", fontName=F, fontSize=8.5, leading=13, textColor=MED_GRAY, leftIndent=20,
    )
    return s


# ──────────────────────────────────────────────────────────────────────
# Markdown helpers
# ──────────────────────────────────────────────────────────────────────


def esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def md_inline(t: str) -> str:
    t = esc(t)
    t = re.sub(r"\*\*\*(.*?)\*\*\*", r"<b><i>\1</i></b>", t)
    t = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<i>\1</i>", t)
    # Inline code: monospace, dark red, slightly smaller
    t = re.sub(
        r"`([^`]+)`",
        lambda m: f'<font face="{FM}" size="8" color="#c0392b">{m.group(1)}</font>',
        t,
    )
    # Plain links: <https://...> or markdown [text](url)
    t = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        lambda m: f'<font color="#0d7377"><u>{m.group(1)}</u></font>',
        t,
    )
    t = re.sub(
        r"&lt;(https?://[^&]+)&gt;",
        lambda m: f'<font color="#0d7377"><u>{m.group(1)}</u></font>',
        t,
    )
    return t


def parse_table(lines: list[str]) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in lines:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        rows.append(cells)
    return [r for r in rows if not all(re.match(r"^[-:]+$", c) or c == "" for c in r)]


def build_table(rows: list[list[str]], styles: dict):
    if not rows or len(rows) < 2:
        return None
    hdr = rows[0]
    data = rows[1:]
    nc = len(hdr)
    td = [[Paragraph(md_inline(c), styles["tbl_hdr"]) for c in hdr]]
    for r in data:
        while len(r) < nc:
            r.append("")
        td.append([Paragraph(md_inline(c), styles["tbl_cell"]) for c in r[:nc]])
    cw = (CW - 4 * mm) / nc
    t = Table(td, colWidths=[cw] * nc, repeatRows=1)
    sc = [
        ("BACKGROUND", (0, 0), (-1, 0), TBL_HDR_BG),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.5, BORDER_GRAY),
        ("LINEBELOW", (0, 0), (-1, 0), 1, TEAL),
    ]
    for i in range(1, len(td)):
        if i % 2 == 0:
            sc.append(("BACKGROUND", (0, i), (-1, i), TBL_ALT_BG))
    t.setStyle(TableStyle(sc))
    return t


# ──────────────────────────────────────────────────────────────────────
# Markdown to ReportLab flowables
# ──────────────────────────────────────────────────────────────────────


def md_to_flowables(md_text: str, styles: dict) -> list:
    lines = md_text.split("\n")
    fl: list = []
    i = 0
    toc_n = [0]

    def heading(level: int, style_key: str, raw_text: str) -> None:
        clean = raw_text.strip()
        display = md_inline(clean)
        toc_text = re.sub(r"<[^>]+>", "", esc(clean))
        key = f"toc_{toc_n[0]}"
        toc_n[0] += 1
        fl.append(TOCAnchor(key, level, toc_text))
        fl.append(Paragraph(f'<a name="{key}"/>{display}', styles[style_key]))

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue

        # Horizontal rule
        if stripped in ("---", "***", "___"):
            fl.append(HRFlowable(
                width="100%", thickness=0.5, color=TEAL,
                spaceBefore=4 * mm, spaceAfter=4 * mm,
            ))
            i += 1
            continue

        # Table
        if "|" in stripped and stripped.startswith("|"):
            tl = [stripped]
            i += 1
            while i < len(lines) and "|" in lines[i].strip() and lines[i].strip().startswith("|"):
                tl.append(lines[i].strip())
                i += 1
            rows = parse_table(tl)
            tbl = build_table(rows, styles)
            if tbl:
                fl.append(Spacer(1, 2 * mm))
                fl.append(tbl)
                fl.append(Spacer(1, 3 * mm))
            continue

        # Code block
        if stripped.startswith("```"):
            code_lines: list[str] = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            if i < len(lines):
                i += 1  # skip closing ```
            code_text = "\n".join(code_lines)
            code_escaped = esc(code_text).replace("\n", "<br/>")
            fl.append(Paragraph(code_escaped, styles["code"]))
            continue

        # Block quote
        if stripped.startswith(">"):
            quote_lines = [stripped.lstrip("> ").rstrip()]
            i += 1
            while i < len(lines) and lines[i].strip().startswith(">"):
                quote_lines.append(lines[i].strip().lstrip("> ").rstrip())
                i += 1
            fl.append(Paragraph(md_inline(" ".join(quote_lines)), styles["blockquote"]))
            continue

        # ####
        if stripped.startswith("####") and not stripped.startswith("#####"):
            heading(2, "subsection", stripped.lstrip("#").strip())
            i += 1
            continue

        # ###
        if stripped.startswith("###") and not stripped.startswith("####"):
            heading(1, "section", stripped.lstrip("#").strip())
            i += 1
            continue

        # ##
        if stripped.startswith("##") and not stripped.startswith("###"):
            heading(0, "h2_other", stripped.lstrip("#").strip())
            i += 1
            continue

        # # — title line, skip in body
        if stripped.startswith("#") and not stripped.startswith("##"):
            i += 1
            continue

        # Bullet
        if stripped.startswith("- ") or (stripped.startswith("* ") and not stripped.startswith("**")):
            bt = stripped[2:].strip()
            i += 1
            while i < len(lines) and lines[i].strip() and \
                    not lines[i].strip().startswith(("-", "*", "#", "|", ">")) and \
                    not re.match(r"^\d+\.", lines[i].strip()):
                if lines[i].startswith("  ") or lines[i].startswith("\t"):
                    bt += " " + lines[i].strip()
                    i += 1
                else:
                    break
            fl.append(Paragraph("•  " + md_inline(bt), styles["bullet"]))
            continue

        # Numbered list
        m = re.match(r"^(\d+)\.\s+(.*)", stripped)
        if m:
            num, txt = m.group(1), m.group(2)
            i += 1
            while i < len(lines) and lines[i].strip() and \
                    not re.match(r"^\d+\.", lines[i].strip()) and \
                    not lines[i].strip().startswith(("#", "-", "|", ">")):
                if lines[i].startswith("  ") or lines[i].startswith("\t"):
                    txt += " " + lines[i].strip()
                    i += 1
                else:
                    break
            fl.append(Paragraph(f"{num}.  {md_inline(txt)}", styles["bullet"]))
            continue

        # Body paragraph
        para = stripped
        i += 1
        while i < len(lines):
            nl = lines[i].strip()
            if not nl:
                break
            if nl.startswith(("#", "-", "|", "---", "```", ">")):
                break
            if re.match(r"^\d+\.\s", nl):
                break
            if nl.startswith("* ") and not nl.startswith("**"):
                break
            para += " " + nl
            i += 1
        fl.append(Paragraph(md_inline(para), styles["body"]))

    return fl


# ──────────────────────────────────────────────────────────────────────
# Cover and TOC
# ──────────────────────────────────────────────────────────────────────


def build_cover(styles: dict) -> list:
    el: list = []
    el.append(Spacer(1, 35 * mm))
    el.append(Paragraph("Excel Finance Agent", styles["cover_title"]))
    el.append(Spacer(1, 6 * mm))
    el.append(Paragraph("A Simple User Guide", styles["cover_sub"]))
    el.append(Spacer(1, 10 * mm))
    el.append(HRFlowable(width="40%", thickness=2, color=TEAL, spaceAfter=8 * mm, hAlign="LEFT"))

    el.append(Paragraph(
        "Built by <b>Shubham Sahamate</b><br/>"
        "with the help of Claude Opus 4.7 (1M context)",
        styles["cover_author"],
    ))
    el.append(Spacer(1, 8 * mm))

    for m in [
        "<b>For:</b> First-time users, both technical and non-technical",
        "<b>Scope:</b> End-to-end setup and daily use of the agent",
        "<b>Repo:</b> github.com/sahamate15/excel-finance-agent",
        "<b>As of:</b> April 2026",
    ]:
        el.append(Paragraph(m, styles["cover_meta"]))
        el.append(Spacer(1, 1 * mm))

    el.append(Spacer(1, 14 * mm))

    callout = Table(
        [[Paragraph(
            "Plain-English finance instructions become Excel formulas and "
            "multi-row models. Built for investment-banking workflows with "
            "audit-grade compliance logging. This guide takes you from zero "
            "to your first depreciation schedule in about 20 minutes.",
            ParagraphStyle("blurb", fontName=F, fontSize=9, leading=13, textColor=DARK_GRAY),
        )]],
        colWidths=[CW],
    )
    callout.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CALLOUT_BG),
        ("LINEABOVE", (0, 0), (-1, 0), 1, TEAL),
        ("LINEBELOW", (0, -1), (-1, -1), 1, TEAL),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    el.append(callout)

    el.append(NextPageTemplate("body"))
    el.append(PageBreak())
    return el


def build_toc(entries: list, styles: dict) -> list:
    fl = [Paragraph("Table of Contents", styles["toc_title"])]
    for level, text, page in entries:
        sk = ["toc_part", "toc_sec", "toc_sub"][min(level, 2)]
        dots = '<font color="#b2bec3"> ... </font>'
        pg = f'<font color="#636e72">{page}</font>'
        fl.append(Paragraph(f"{esc(text)}{dots}{pg}", styles[sk]))
    fl.append(PageBreak())
    return fl


# ──────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────


def main() -> None:
    here = Path(__file__).resolve().parent
    src = here.parent / "GETTING_STARTED.md"
    out = here / "Excel_Finance_Agent_User_Guide.pdf"

    md = src.read_text(encoding="utf-8")
    styles = make_styles()

    # Pass 1: collect TOC entries with placeholder TOC pages.
    tmp = str(here / "_pass1.pdf")
    doc1 = DocTemplate(
        tmp, pagesize=A4,
        leftMargin=LM, rightMargin=RM, topMargin=TM, bottomMargin=BM,
    )
    s1 = build_cover(styles)
    s1.append(Spacer(1, 200 * mm))
    s1.append(PageBreak())
    s1.extend(md_to_flowables(md, styles))
    doc1.build(s1)
    toc = doc1.toc_entries
    print(f"Pass 1: {len(toc)} TOC entries collected")

    # Pass 2: real TOC with the page numbers from pass 1.
    doc2 = DocTemplate(
        str(out), pagesize=A4,
        leftMargin=LM, rightMargin=RM, topMargin=TM, bottomMargin=BM,
        title="Excel Finance Agent: A Simple User Guide",
        author="Shubham Sahamate",
    )
    s2 = build_cover(styles)
    s2.extend(build_toc(toc, styles))
    s2.extend(md_to_flowables(md, styles))
    doc2.build(s2)

    os.remove(tmp)

    try:
        from pypdf import PdfReader
        pages = len(PdfReader(str(out)).pages)
        print(f"Done: {out}  ({pages} pages)")
    except Exception as exc:
        print(f"Done: {out}  (page count unavailable: {exc})")


if __name__ == "__main__":
    main()
