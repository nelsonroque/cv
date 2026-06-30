#!/usr/bin/env python3
"""Convert a DOCX CV into a first-pass Quarto Markdown source file.

This is intentionally conservative:
- preserve paragraph order
- map Word Title/Heading3 styles to Markdown headings
- convert tables into bullet lists so the content stays editable

The goal is a clean starting point for later manual cleanup and
phase-2 bibliography extraction.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"


def para_text(p: ET.Element) -> str:
    parts: list[str] = []
    for node in p.iter():
        if node.tag == f"{W}t":
            parts.append(node.text or "")
        elif node.tag == f"{W}tab":
            parts.append("\t")
        elif node.tag == f"{W}br":
            parts.append("\n")
    text = "".join(parts).strip()
    text = text.replace("Email:", "Email: ")
    text = text.replace("University  of Central Florida", "University of Central Florida")
    text = text.replace("2024- Present", "2024-present")
    text = text.replace("2020-2023", "2020-2023")
    return text


def table_rows(tbl: ET.Element) -> list[list[str]]:
    rows: list[list[str]] = []
    for tr in tbl.findall("./w:tr", NS):
        cells: list[str] = []
        for tc in tr.findall("./w:tc", NS):
            text = para_text(tc)
            if text:
                cells.append(text)
        if cells:
            rows.append(cells)
    return rows


def is_year(text: str) -> bool:
    return text.isdigit() and len(text) == 4 and 1900 <= int(text) <= 2100


def is_section_heading(text: str) -> bool:
    compact = " ".join(text.split())
    if not compact or len(compact) > 60:
        return False
    if is_year(compact[:4]) or re.match(r"^(19|20)\d{2}[-,]", compact):
        return False
    if compact.startswith(("NIA-", "NIA ", "Member,", "Ad hoc Reviewer", "Roque", "Dzombak", "King,", "Kelly", "Hawkins", "Lomte", "Bennett", "Facher", "Gopal", "Bioengineer", "News, Mirage", "SCIENMAG", "State, Penn")):
        return False
    if compact.count(".") > 0:
        return False
    words = compact.split()
    return 1 <= len(words) <= 8


PLAIN_HEADINGS = {
    "Peer-Reviewed Articles (Supervised students & postdocs underlined)",
    "Other Publications (graduate students & postdocs when written underlined)",
    "Book Chapters",
    "Technical Reports",
    "Conference Proceedings",
    "Current Funding Applications In Review",
    "External Research Funding",
    "Teaching Preparation and Training",
    "Service to Penn State",
    "Service to Funding Agencies",
    "Service to Professional Organizations",
    "Service to Profession, Peer Review",
    "Community Outreach",
    "Invited Talks",
    "Conferences: Organized Symposia",
    "Conferences: Oral Presentations",
    "Conferences: Posters",
    "University / Department / Center: Oral Presentations",
    "Presentations (approximately last ten years)",
}


def is_entry_like(text: str) -> bool:
    compact = " ".join(text.split())
    if re.match(r"^\d{4}([–-]\d{4}|[–-]\s*present|[–-]present|[–-]\d{2}|\b|,)", compact, re.I):
        return True
    if compact.startswith(
        (
            "NIA-",
            "NIA ",
            "Member,",
            "Ad hoc Reviewer",
            "Roque",
            "Dzombak,",
            "King,",
            "Kelly,",
            "Hawkins,",
            "Douglas,",
            "Lomte,",
            "Bennett,",
            "Facher,",
            "Gopal,",
            "Bioengineer",
            "News, Mirage",
            "MSN.",
            "DNYUZ.",
            "Laboratory Equipment.",
            "SCIENMAG.",
            "State, Penn",
            "Wagner,",
            "Charness,",
            "Boot,",
            "Harrington,",
            "Sliwinski,",
            "Oravecz,",
            "Hyun,",
            "Hoogendoorn,",
            "De La Garza,",
            "Sergeyev,",
            "Andrews,",
            "Garza,",
            "Prieto,",
            "Flores-Cruz,",
            "Thai,",
            "Roque N.",
            "Roque, N.",
            "Roque, N.A.",
        )
    ):
        return True
    if compact.startswith(("202", "201", "May ", "Fall ", "Spring ", "Summer ")):
        return True
    return False


def heading_level(text: str) -> int:
    major = {
        "Appointments & Affiliations",
        "Education",
        "Research and Teaching Interests",
        "Awards, Honors, & Distinctions",
        "Teaching Experience",
        "Professional Service",
        "Professional Memberships",
        "Scientific Outreach",
    }
    if text in major:
        return 2
    return 3


def format_table_row(cells: list[str]) -> str:
    clean = [c.strip() for c in cells if c.strip()]
    if not clean:
        return ""
    if len(clean) == 1:
        return clean[0]

    first = clean[0]
    rest = " ".join(clean[1:]).strip()

    if is_year(first):
        return f"{first} — {rest}" if rest else first

    if first.isdigit() and len(first) <= 2 and rest:
        return rest

    if len(clean) == 2 and rest:
        return f"{first} — {rest}"

    return " — ".join(clean)


def convert_docx(docx_path: Path) -> str:
    with ZipFile(docx_path) as z:
        document = ET.fromstring(z.read("word/document.xml"))

    body = document.find("./w:body", NS)
    if body is None:
        raise RuntimeError("DOCX body not found")

    blocks: list[str] = []

    title_text = "Nelson Roque Curriculum Vitae"

    for child in list(body):
        if child.tag == f"{W}p":
            text = para_text(child)
            if not text:
                continue
            style_el = child.find("./w:pPr/w:pStyle", NS)
            style = style_el.attrib.get(f"{W}val") if style_el is not None else ""

            if style == "Title":
                title_text = text
                continue
            if style == "Heading3":
                if is_section_heading(text):
                    blocks.append(f"{'#' * heading_level(text)} {text}")
                elif is_entry_like(text):
                    blocks.append(f"- {text}")
                else:
                    blocks.append(f"### {text}")
                continue
            if text in PLAIN_HEADINGS:
                blocks.append(f"### {text}")
                continue
            if text in {
                "The Pennsylvania State University",
                "Department of Human Development and Family Studies",
                "422 Biobehavioral Health Building, University Park, 16802",
                "Email: nur375@psu.edu",
            }:
                continue
            blocks.append(text)

        elif child.tag == f"{W}tbl":
            rows = table_rows(child)
            if not rows:
                continue
            blocks.append("")
            for row in rows:
                formatted = format_table_row(row)
                if formatted:
                    blocks.append(f"- {formatted}")
            blocks.append("")

    while blocks and not blocks[0].strip():
        blocks.pop(0)
    while blocks and not blocks[-1].strip():
        blocks.pop()

    lines = [
        "---",
        f"title: {title_text!r}",
        "format:",
        "  html:",
        "    toc: false",
        "    page-layout: full",
        "---",
        "",
        "<style>",
        "main.content { max-width: 960px; margin: 0 auto; }",
        "#title-block-header, header.quarto-title-block { display: none !important; }",
        "h2 { margin-top: 1.4rem; border-bottom: 1px solid #ddd; padding-bottom: 0.2rem; }",
        "h3 { margin-top: 1.1rem; }",
        "p, li { line-height: 1.35; }",
        "ul { margin-top: 0.35rem; margin-bottom: 0.7rem; }",
        "</style>",
        "",
        "# Nelson Roque",
        "",
        "The Pennsylvania State University  \nDepartment of Human Development and Family Studies  \n422 Biobehavioral Health Building, University Park, 16802  \nEmail: nur375@psu.edu",
        "",
    ]

    pending_blank = False
    for block in blocks:
        if not block.strip():
            pending_blank = True
            continue
        if pending_blank and lines and lines[-1] != "":
            lines.append("")
        pending_blank = False
        lines.append(block)
        if block.startswith("#"):
            lines.append("")
        elif not block.startswith("- "):
            lines.append("")

    while lines and lines[-1] == "":
        lines.pop()
    lines.append("")
    return "\n".join(lines)


def main() -> int:
    if len(sys.argv) != 3:
        print(
            "usage: convert_docx_to_qmd.py INPUT.docx OUTPUT.qmd",
            file=sys.stderr,
        )
        return 2

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    output_path.write_text(convert_docx(input_path), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
