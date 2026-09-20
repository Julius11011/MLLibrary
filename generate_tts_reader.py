#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MLC Universal Text-to-Speech (Read Aloud) HTML Reader Generator
===============================================================
Converts legal study materials (.docx, .pdf, .md, .txt) across all MLC folders
into interactive, high-readability HTML web applications featuring:
  - Natural, Human-like Cadence & Intonation with Smart Legal Phonetic Expansions
  - Continuous, Flowing Body Paragraphs with Modern Legal Typography
  - Streamlined, Non-Redundant Table of Contents (TOC) with Instant Live Search
  - Clear Visual Hierarchy: Case Badges, Citation Banners, Subheadings, ALAC Badges & Tables
  - Native Studio MP3 Podcast Audio Player Integration (Mobile Background Compatible)
  - Full Keyboard Navigation, Theme Modes (Dark, Sepia, Light), and Responsive Mobile Layout
"""

import os
import sys
import re
import html
import json
import argparse
from pathlib import Path

# Optional dependencies with graceful fallback
try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import pymupdf as fitz
    HAS_FITZ = True
except ImportError:
    try:
        import fitz
        HAS_FITZ = True
    except ImportError:
        HAS_FITZ = False

# ==============================================================================
# 1. TEXT CLEANING & NORMALIZATION
# ==============================================================================

def clean_text(text):
    if not text:
        return ""
    text = text.replace("\ufffd", "'")
    text = text.replace("Pala'a", "Palaña").replace("Palaa", "Palaña")
    text = text.replace("Ca'ete", "Cañete").replace("Caete", "Cañete")
    text = text.replace("Pe'a", "Peña").replace("Pea", "Peña")
    text = text.replace("A'onuevo", "Año nuevo")
    text = text.replace("Mu'oz", "Muñoz")
    text = text.replace("Qui'ones", "Quiñones")
    text = text.replace("Nu'ez", "Nuñez")
    text = text.replace("Iba'ez", "Ibañez")
    text = text.replace("Taada", "Tañada").replace("Ta'ada", "Tañada")
    text = text.replace("Peacerrada", "Peñacerrada").replace("Pe'acerrada", "Peñacerrada")
    text = text.replace("’", "'").replace("‘", "'")
    text = text.replace("“", '"').replace("”", '"')
    text = text.replace("–", "-").replace("—", " - ")
    return text

def slugify(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9\-_\s]', '', text)
    text = re.sub(r'\s+', '-', text)
    return text[:60].strip('-') or 'section'

# Regex patterns for high-precision legal document structure
CASE_HEADING_RE = re.compile(
    r'^(CASE\s+\d+[:\.]?.*|PART\s+[I|V|X\d]+[:\.]?.*|Canon\s+[I|V|X\d]+[:\.]?.*|CHAPTER\s+[I|V|X\d]+[:\.]?.*|TOPIC\s+[I|V|X\d]+[:\.]?.*)',
    re.IGNORECASE
)

SUBHEADING_RE = re.compile(
    r'^([I|V|X]+\.\s+[A-Z\s\(\)&,\-\/:]{3,}|HOW THE COURT CONSTRUED|RELEVANT STATUTORY|APPLICABLE\s+.*(?:PRINCIPLES|MAXIMS|CANONS|PROVISIONS)|SUMMARY OF THE STATUTORY|SUBJECT MATTER & PROVISION|SYLLABUS TOPIC|FACTS OF THE CASE|STATEMENT OF THE ETHICAL ISSUE|SUBSTANTIVE LEGAL ISSUE|THE COURT\'S RULING|THE RULING|COURT\'S RULING)',
    re.IGNORECASE
)

ALAC_PATTERNS = [
    r'Ethical Synthesis & Canon Application',
    r'Criminal Law Synthesis & Statutory Application',
    r'Statutory Construction Synthesis & Methodological Application',
    r'APPLICATION\s*\/\s*ANALYSIS',
    r'CONCLUSION\s*\/\s*DISPOSITIVE\s*RULING',
    r'ANSWER\s*\/\s*RULING',
    r'A\s*[\-–—]\s*ANSWER',
    r'L\s*[\-–—]\s*LEGAL\s*BASIS',
    r'A\s*[\-–—]\s*APPLICATION',
    r'C\s*[\-–—]\s*CONCLUSION',
    r'DISPOSITIVE\s*EFFECT',
    r'DISPOSITIVE\s*RULING',
    r'LEGAL\s*BASIS',
    r'ANSWER',
    r'APPLICATION',
    r'ANALYSIS',
    r'CONCLUSION',
    r'\d+\.\s+Why Construction Was Necessary',
    r'\d+\.\s+Purpose of Construction',
    r'\d+\.\s+Canon\s*\/\s*Method Applied',
    r'\d+\.\s+Extrinsic Aids\s*\/\s*Other Sources Used',
    r'Core Legal Mandate',
    r'Practical Meaning',
    r'High-Yield Bar Rule'
]
ALAC_RE = re.compile(r'^(' + '|'.join(ALAC_PATTERNS) + r')\s*[:\.\-–—]?', re.IGNORECASE)

# ==============================================================================
# 2. DOCUMENT PARSERS (.docx, .pdf, .md, .txt)
# ==============================================================================

def format_inline_runs(paragraph):
    """Converts a docx paragraph into clean HTML with bold, italic, and underline tags."""
    formatted_runs = []
    for run in paragraph.runs:
        t = html.escape(clean_text(run.text))
        if run.bold and run.italic:
            t = f"<strong><em>{t}</em></strong>"
        elif run.bold:
            t = f"<strong>{t}</strong>"
        elif run.italic:
            t = f"<em>{t}</em>"
        if run.underline:
            t = f"<u>{t}</u>"
        formatted_runs.append(t)
    
    text = "".join(formatted_runs).strip()
    return text

def parse_docx_file(file_path, doc_title=""):
    """
    Parses Microsoft Word (.docx) file into structured case sections with:
      - TOC limited strictly to Major Case / Canon / Part headings
      - In-case styled subheadings (Facts, Issue, Ruling, Ethical Principles)
      - Distinct styled ALAC inline badges (Answer, Legal Basis, Application, Conclusion)
      - Smooth, non-fragmented continuous body paragraphs
    """
    if not HAS_DOCX:
        raise RuntimeError("python-docx is required. Install with 'pip install python-docx'.")
    
    doc = docx.Document(file_path)
    sections = []
    current_sec = {
        "id": "sec-overview",
        "title": doc_title or "Document Overview",
        "level": 1,
        "units": []
    }
    sec_counter = 1
    expect_citation = False

    for p in doc.paragraphs:
        raw_text = clean_text(p.text.strip())
        if not raw_text:
            continue

        style_name = p.style.name.lower() if p.style else ""
        formatted_text = format_inline_runs(p)
        if not formatted_text:
            continue

        # 1. Check for Top-Level Case / Canon / Part Headings (Creates TOC Section)
        is_major_heading = bool(
            CASE_HEADING_RE.match(raw_text) or
            ("heading 1" in style_name and len(raw_text) < 140 and not SUBHEADING_RE.match(raw_text)) or
            ("title" in style_name and len(raw_text) < 140)
        )

        if is_major_heading:
            if current_sec["units"] or current_sec["title"] != (doc_title or "Document Overview"):
                sections.append(current_sec)
            
            sec_id = f"sec-{sec_counter}-{slugify(raw_text)}"
            sec_counter += 1
            
            # Extract case badge if it is a case
            case_badge = ""
            case_name = raw_text
            case_match = re.match(r'^(CASE\s+\d+)[:\.]?\s*(.*)$', raw_text, re.IGNORECASE)
            if case_match:
                case_badge = case_match.group(1).upper()
                case_name = case_match.group(2) or case_badge
            
            current_sec = {
                "id": sec_id,
                "title": raw_text,
                "level": 1,
                "units": []
            }
            
            if case_badge:
                current_sec["units"].append(
                    f'<div class="case-header-card read-unit" data-unit-type="case-header"><span class="case-number-pill">{case_badge}</span><h2 class="case-header-title">{html.escape(case_name)}</h2></div>'
                )
            else:
                current_sec["units"].append(
                    f'<div class="case-header-card read-unit" data-unit-type="case-header"><h2 class="case-header-title">{html.escape(raw_text)}</h2></div>'
                )
            expect_citation = True
            continue

        # 2. Check for Case Citation Line (e.g. A.C. No. 13521, June 27, 2023 | Per Curiam)
        if expect_citation and (len(raw_text) < 180 and any(k in raw_text for k in ["SCRA", "G.R.", "A.C.", "A.M.", "Phil.", "Ponente", "Curiam", "En Banc", "|"])):
            current_sec["units"].append(
                f'<div class="case-citation-banner read-unit" data-unit-type="citation"><span class="citation-icon">⚖</span> <span class="citation-text">{formatted_text}</span></div>'
            )
            expect_citation = False
            continue
        
        expect_citation = False

        # 3. Check for Subheadings within the Case (e.g. I. FACTS OF THE CASE, II. ISSUE)
        if SUBHEADING_RE.match(raw_text):
            sub_title = raw_text
            current_sec["units"].append(
                f'<h3 class="case-subheading read-unit" data-unit-type="subheading"><span class="subheading-accent">§</span> {formatted_text}</h3>'
            )
            continue

        # 4. Check for ALAC & Structured Reasoning Badges (ANSWER:, LEGAL BASIS:, ANALYSIS:, CONCLUSION:)
        alac_match = ALAC_RE.match(raw_text)
        if alac_match:
            badge_label = alac_match.group(1).strip()
            b_upper = badge_label.upper()
            badge_class = "badge-ans" if any(k in b_upper for k in ["ANSWER", "A - ANSWER"]) else \
                          "badge-law" if any(k in b_upper for k in ["LEGAL BASIS", "L - LEGAL"]) else \
                          "badge-app" if any(k in b_upper for k in ["APPLICATION", "ANALYSIS", "HOW THE"]) else \
                          "badge-con" if any(k in b_upper for k in ["CONCLUSION", "DISPOSITIVE"]) else \
                          "badge-syn" if any(k in b_upper for k in ["SYNTHESIS", "MANDATE", "MEANING", "BAR RULE", "WHY CONSTRUCTION", "PURPOSE OF", "CANON /", "EXTRINSIC"]) else "badge-gen"

            chars_to_skip = alac_match.end()
            formatted_body_runs = []
            
            for run in p.runs:
                run_text = clean_text(run.text)
                if chars_to_skip >= len(run_text):
                    chars_to_skip -= len(run_text)
                    continue
                elif chars_to_skip > 0:
                    run_text = run_text[chars_to_skip:]
                    chars_to_skip = 0
                    
                if not run_text.strip():
                    if run_text:
                        formatted_body_runs.append(" ")
                    continue
                    
                t = html.escape(run_text)
                if run.bold and run.italic:
                    t = f'<strong><em>{t}</em></strong>'
                elif run.bold:
                    t = f'<strong>{t}</strong>'
                elif run.italic:
                    t = f'<em>{t}</em>'
                if run.underline:
                    t = f'<u>{t}</u>'
                formatted_body_runs.append(t)
                
            rest_html = "".join(formatted_body_runs).strip()
            rest_html = re.sub(r'^(?:[:\-–—]\s*)+', '', rest_html)
            
            current_sec["units"].append(
                f'<p class="read-unit case-paragraph alac-paragraph" data-unit-type="alac"><span class="alac-badge {badge_class}">{html.escape(badge_label)}</span> <span class="alac-text">{rest_html}</span></p>'
            )
            continue

        # 5. Check for Bullet Points and Lists
        if "list" in style_name or "bullet" in style_name or raw_text.startswith(('•', '-', '*')) or re.match(r'^[•\-\*]\s*', raw_text):
            clean_bullet = re.sub(r'^[•\-\*]\s*', '', formatted_text)
            current_sec["units"].append(
                f'<div class="bullet-point read-unit" data-unit-type="bullet"><span class="bullet-dot">•</span><div class="bullet-content">{clean_bullet}</div></div>'
            )
            continue

        # 6. Standard Continuous Flowing Paragraph
        current_sec["units"].append(f'<p class="read-unit case-paragraph" data-unit-type="paragraph">{formatted_text}</p>')

    # Extract tables in docx
    for table in doc.tables:
        html_tbl = ['<div class="table-responsive read-unit" data-unit-type="table"><table class="reader-table">']
        for i, row in enumerate(table.rows):
            tag = 'th' if i == 0 else 'td'
            html_tbl.append('<tr>')
            for cell in row.cells:
                ctext = html.escape(clean_text(cell.text.strip()))
                html_tbl.append(f'<{tag}>{ctext}</{tag}>')
            html_tbl.append('</tr>')
        html_tbl.append('</table></div>')
        current_sec["units"].append("".join(html_tbl))

    if current_sec["units"] or current_sec["title"]:
        sections.append(current_sec)

    return sections

def parse_pdf_file(file_path, doc_title=""):
    """
    Parses PDF document using PyMuPDF (fitz), de-fragmenting lines into smooth paragraphs
    and organizing TOC strictly around major case/document headings.
    """
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) is required. Install with 'pip install pymupdf'.")
    
    doc = fitz.open(file_path)
    all_blocks = []

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if "lines" in b:
                block_text_parts = []
                is_bold_block = False
                for l in b["lines"]:
                    line_str = "".join(s["text"] for s in l["spans"]).strip()
                    bbox = l["bbox"]
                    # Skip recurring header/footer
                    if bbox[1] > 740 and line_str.isdigit():
                        continue
                    if bbox[1] < 50 and ("DEAN" in line_str or "FACULTY OF CIVIL LAW" in line_str):
                        continue
                    if line_str:
                        if any("bold" in s.get("font", "").lower() or s.get("flags", 0) & 2 for s in l["spans"]):
                            is_bold_block = True
                        block_text_parts.append(line_str)
                if block_text_parts:
                    joined = clean_text(" ".join(block_text_parts))
                    all_blocks.append((joined, is_bold_block))

    sections = []
    current_sec = {
        "id": "sec-overview",
        "title": doc_title or "Document Overview",
        "level": 1,
        "units": []
    }
    sec_counter = 1

    for text, is_bold in all_blocks:
        if not text:
            continue

        is_major = bool(
            CASE_HEADING_RE.match(text) or
            (is_bold and len(text) < 120 and not SUBHEADING_RE.match(text) and (text.isupper() or "CASE" in text.upper()))
        )

        if is_major:
            if current_sec["units"] or current_sec["title"] != (doc_title or "Document Overview"):
                sections.append(current_sec)
            sec_id = f"sec-{sec_counter}-{slugify(text)}"
            sec_counter += 1
            current_sec = {
                "id": sec_id,
                "title": text,
                "level": 1,
                "units": []
            }
            current_sec["units"].append(
                f'<div class="case-header-card read-unit" data-unit-type="case-header"><h2 class="case-header-title">{html.escape(text)}</h2></div>'
            )
            continue

        if SUBHEADING_RE.match(text):
            current_sec["units"].append(
                f'<h3 class="case-subheading read-unit" data-unit-type="subheading"><span class="subheading-accent">§</span> {html.escape(text)}</h3>'
            )
            continue

        esc = html.escape(text)
        if is_bold and len(text) < 120:
            esc = f"<strong>{esc}</strong>"

        current_sec["units"].append(f'<p class="read-unit case-paragraph" data-unit-type="paragraph">{esc}</p>')

    if current_sec["units"] or current_sec["title"]:
        sections.append(current_sec)

    return sections

def parse_markdown(content, doc_title=""):
    """Parses Markdown text into structured case sections and units."""
    lines = content.splitlines()
    sections = []
    current_sec = {
        "id": "sec-overview",
        "title": doc_title or "Introduction",
        "level": 1,
        "units": []
    }
    sec_counter = 1

    def format_inline(t):
        t = clean_text(t)
        t = html.escape(t)
        t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
        t = re.sub(r'__(.+?)__', r'<strong>\1</strong>', t)
        t = re.sub(r'\*(.+?)\*', r'<em>\1</em>', t)
        t = re.sub(r'_(.+?)_', r'<em>\1</em>', t)
        t = re.sub(r'`(.+?)`', r'<code>\1</code>', t)
        return t

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Major Headings (# or ## CASE)
        heading_match = re.match(r'^(#{1,3})\s+(.+)$', stripped)
        if heading_match:
            htext = heading_match.group(2).strip()
            if CASE_HEADING_RE.match(htext) or heading_match.group(1) == '#':
                if current_sec["units"] or current_sec["title"] != (doc_title or "Introduction"):
                    sections.append(current_sec)
                sec_id = f"sec-{sec_counter}-{slugify(htext)}"
                sec_counter += 1
                current_sec = {
                    "id": sec_id,
                    "title": clean_text(htext),
                    "level": 1,
                    "units": []
                }
                current_sec["units"].append(
                    f'<div class="case-header-card read-unit" data-unit-type="case-header"><h2 class="case-header-title">{html.escape(htext)}</h2></div>'
                )
                continue
            else:
                current_sec["units"].append(
                    f'<h3 class="case-subheading read-unit" data-unit-type="subheading"><span class="subheading-accent">§</span> {format_inline(htext)}</h3>'
                )
                continue

        formatted_p = format_inline(stripped)
        current_sec["units"].append(f'<p class="read-unit case-paragraph" data-unit-type="paragraph">{formatted_p}</p>')

    if current_sec["units"] or current_sec["title"]:
        sections.append(current_sec)

    return sections

def parse_txt_file(file_path, doc_title=""):
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    return parse_markdown(content, doc_title)

# ==============================================================================
# 3. MODERN RESPONSIVE HTML READER TEMPLATE
# ==============================================================================

HTML_PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="description" content="{escaped_description}">
  <title>{escaped_title} | MLC Law Library</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700;800;900&family=Inter:wght@300;400;500;600;700&family=Merriweather:ital,wght@0,300;0,400;0,700;1,300;1,400&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <style>
    :root {{
      --bg-primary: #0b1120;
      --bg-secondary: #131d31;
      --bg-tertiary: #1e293b;
      --bg-card: rgba(19, 29, 49, 0.85);
      --bg-card-hover: rgba(30, 41, 59, 0.95);
      --border-color: rgba(148, 163, 184, 0.16);
      --border-focus: #38bdf8;
      --text-primary: #f8fafc;
      --text-secondary: #cbd5e1;
      --text-muted: #94a3b8;
      --accent-gold: #fbbf24;
      --accent-gold-dark: #d97706;
      --accent-blue: #38bdf8;
      --accent-indigo: #818cf8;
      --accent-emerald: #34d399;
      --accent-rose: #fb7185;
      --highlight-reading: rgba(56, 189, 248, 0.22);
      --font-body: 'Merriweather', Georgia, serif;
      --font-ui: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      --font-heading: 'Cinzel', serif;
      --font-mono: 'JetBrains Mono', monospace;
      --sidebar-width: 340px;
      --header-height: 68px;
      --content-max-width: 900px;
      --base-font-size: 17px;
      --line-height: 1.85;
      --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.3), 0 2px 4px -1px rgba(0, 0, 0, 0.2);
      --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.4), 0 4px 6px -2px rgba(0, 0, 0, 0.3);
      --glass-blur: blur(14px);
    }}

    [data-theme="light"] {{
      --bg-primary: #f8fafc;
      --bg-secondary: #ffffff;
      --bg-tertiary: #f1f5f9;
      --bg-card: rgba(255, 255, 255, 0.95);
      --bg-card-hover: rgba(248, 250, 252, 1);
      --border-color: rgba(203, 213, 225, 0.85);
      --border-focus: #0284c7;
      --text-primary: #0f172a;
      --text-secondary: #334155;
      --text-muted: #64748b;
      --accent-gold: #b45309;
      --accent-gold-dark: #92400e;
      --accent-blue: #0284c7;
      --accent-indigo: #4f46e5;
      --accent-emerald: #059669;
      --accent-rose: #e11d48;
      --highlight-reading: rgba(2, 132, 199, 0.15);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.08);
    }}

    [data-theme="sepia"] {{
      --bg-primary: #fbf0d9;
      --bg-secondary: #f4e4c1;
      --bg-tertiary: #ecd6a7;
      --bg-card: rgba(244, 228, 193, 0.95);
      --bg-card-hover: rgba(236, 214, 167, 1);
      --border-color: rgba(180, 150, 110, 0.45);
      --border-focus: #9a6700;
      --text-primary: #2d241e;
      --text-secondary: #4a3b32;
      --text-muted: #786455;
      --accent-gold: #a75d00;
      --accent-gold-dark: #8c4c00;
      --accent-blue: #2c6e91;
      --accent-indigo: #5d4a82;
      --accent-emerald: #2e7d32;
      --accent-rose: #c2185b;
      --highlight-reading: rgba(200, 150, 50, 0.22);
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    html {{
      scroll-behavior: smooth;
      font-size: var(--base-font-size);
    }}

    body {{
      background-color: var(--bg-primary);
      color: var(--text-primary);
      font-family: var(--font-body);
      line-height: var(--line-height);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      transition: background-color 0.25s ease, color 0.25s ease;
      overflow-x: hidden;
    }}

    /* Top Reading Progress Bar */
    #readingProgressBar {{
      position: fixed;
      top: 0;
      left: 0;
      height: 3px;
      width: 0%;
      background: linear-gradient(90deg, var(--accent-gold), var(--accent-blue));
      z-index: 999;
      transition: width 0.1s ease;
    }}

    /* Header */
    header.app-header {{
      position: sticky;
      top: 0;
      z-index: 100;
      height: var(--header-height);
      background-color: rgba(11, 17, 32, 0.92);
      backdrop-filter: var(--glass-blur);
      -webkit-backdrop-filter: var(--glass-blur);
      border-bottom: 1px solid var(--border-color);
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 1.25rem;
      gap: 1rem;
    }}

    [data-theme="light"] header.app-header {{
      background-color: rgba(255, 255, 255, 0.92);
    }}

    [data-theme="sepia"] header.app-header {{
      background-color: rgba(251, 240, 217, 0.94);
    }}

    .brand-section {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
      flex-shrink: 0;
    }}

    .brand-logo {{
      width: 36px;
      height: 36px;
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      color: #0b1120;
      font-weight: 900;
      font-family: var(--font-ui);
      font-size: 1.1rem;
      box-shadow: var(--shadow-sm);
    }}

    .brand-info {{
      display: flex;
      flex-direction: column;
    }}

    .brand-title {{
      font-family: var(--font-ui);
      font-weight: 700;
      font-size: 0.95rem;
      color: var(--text-primary);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 280px;
    }}

    .brand-subject {{
      font-family: var(--font-ui);
      font-size: 0.72rem;
      color: var(--accent-gold);
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }}

    /* Controls Toolbar */
    .tts-toolbar {{
      display: flex;
      align-items: center;
      gap: 0.5rem;
      flex-wrap: wrap;
    }}

    .toolbar-group {{
      display: flex;
      align-items: center;
      background-color: var(--bg-secondary);
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 0.2rem 0.4rem;
      gap: 0.35rem;
    }}

    .icon-btn {{
      background: none;
      border: none;
      color: var(--text-secondary);
      cursor: pointer;
      font-size: 0.95rem;
      padding: 0.35rem 0.55rem;
      border-radius: 6px;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: all 0.15s ease;
    }}

    .icon-btn:hover {{
      background-color: var(--bg-tertiary);
      color: var(--text-primary);
    }}

    .icon-btn.primary {{
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      color: #0b1120;
      font-weight: 700;
    }}

    .icon-btn.primary:hover {{
      filter: brightness(1.1);
      transform: scale(1.02);
    }}

    .toolbar-select {{
      background-color: var(--bg-secondary);
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      font-family: var(--font-ui);
      font-size: 0.8rem;
      padding: 0.3rem 0.5rem;
      border-radius: 6px;
      outline: none;
      cursor: pointer;
    }}

    .tts-status-badge {{
      font-family: var(--font-ui);
      font-size: 0.72rem;
      font-weight: 600;
      padding: 0.2rem 0.55rem;
      border-radius: 12px;
      background-color: var(--bg-tertiary);
      color: var(--text-muted);
    }}

    .tts-status-badge.speaking {{
      background-color: rgba(56, 189, 248, 0.2);
      color: var(--accent-blue);
      border: 1px solid rgba(56, 189, 248, 0.4);
    }}

    /* Main Container */
    .app-main-layout {{
      display: flex;
      flex: 1;
      position: relative;
    }}

    /* Sidebar Navigation (TOC) */
    aside.app-sidebar {{
      width: var(--sidebar-width);
      background-color: var(--bg-secondary);
      border-right: 1px solid var(--border-color);
      height: calc(100vh - var(--header-height));
      position: sticky;
      top: var(--header-height);
      display: flex;
      flex-direction: column;
      transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      z-index: 50;
      flex-shrink: 0;
    }}

    .sidebar-header {{
      padding: 1rem;
      border-bottom: 1px solid var(--border-color);
    }}

    .sidebar-search {{
      width: 100%;
      background-color: var(--bg-primary);
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      font-family: var(--font-ui);
      font-size: 0.82rem;
      padding: 0.5rem 0.75rem;
      border-radius: 6px;
      outline: none;
      transition: border-color 0.2s ease;
    }}

    .sidebar-search:focus {{
      border-color: var(--border-focus);
    }}

    .sidebar-nav {{
      flex: 1;
      overflow-y: auto;
      padding: 0.75rem 0.5rem;
    }}

    .sidebar-nav ul {{
      list-style: none;
    }}

    .sidebar-nav li {{
      margin-bottom: 0.25rem;
    }}

    .toc-link {{
      display: block;
      color: var(--text-secondary);
      text-decoration: none;
      font-family: var(--font-ui);
      font-size: 0.82rem;
      line-height: 1.4;
      padding: 0.5rem 0.75rem;
      border-radius: 6px;
      transition: all 0.15s ease;
      border-left: 2px solid transparent;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}

    .toc-link:hover {{
      background-color: var(--bg-tertiary);
      color: var(--accent-gold);
    }}

    .toc-link.active {{
      background-color: rgba(56, 189, 248, 0.12);
      color: var(--accent-blue);
      border-left-color: var(--accent-blue);
      font-weight: 600;
    }}

    /* Main Content Area */
    main.content-area {{
      flex: 1;
      padding: 2.5rem 2rem 5rem 2rem;
      display: flex;
      justify-content: center;
      min-width: 0;
    }}

    .content-wrapper {{
      max-width: var(--content-max-width);
      width: 100%;
    }}

    /* Document Top Card */
    .doc-hero-card {{
      background: linear-gradient(135deg, var(--bg-secondary), var(--bg-tertiary));
      border: 1px solid var(--border-color);
      border-radius: 14px;
      padding: 2rem;
      margin-bottom: 2.5rem;
      box-shadow: var(--shadow-md);
    }}

    .doc-hero-subject {{
      font-family: var(--font-ui);
      font-size: 0.8rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--accent-gold);
      margin-bottom: 0.4rem;
    }}

    .doc-hero-title {{
      font-family: var(--font-heading);
      font-size: 1.8rem;
      font-weight: 800;
      color: var(--text-primary);
      line-height: 1.3;
      margin-bottom: 0.75rem;
    }}

    .doc-meta-stats {{
      display: flex;
      gap: 1.25rem;
      font-family: var(--font-ui);
      font-size: 0.82rem;
      color: var(--text-muted);
      flex-wrap: wrap;
    }}

    /* Case Section & Cards */
    .doc-section {{
      margin-bottom: 3rem;
    }}

    .case-header-card {{
      background: linear-gradient(135deg, rgba(19, 29, 49, 0.95), rgba(15, 23, 42, 0.98));
      border: 1px solid rgba(251, 191, 36, 0.35);
      border-radius: 12px;
      padding: 1.25rem 1.5rem;
      margin: 2.5rem 0 1.25rem 0;
      box-shadow: var(--shadow-sm);
      display: flex;
      align-items: center;
      gap: 1rem;
      flex-wrap: wrap;
    }}

    [data-theme="light"] .case-header-card {{
      background: linear-gradient(135deg, #ffffff, #f1f5f9);
      border-color: rgba(180, 83, 9, 0.3);
    }}

    .case-number-pill {{
      font-family: var(--font-ui);
      font-size: 0.8rem;
      font-weight: 800;
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      color: #0b1120;
      padding: 0.3rem 0.75rem;
      border-radius: 6px;
      letter-spacing: 0.04em;
      flex-shrink: 0;
    }}

    .case-header-title {{
      font-family: var(--font-heading);
      font-size: 1.35rem;
      font-weight: 700;
      color: var(--text-primary);
      line-height: 1.35;
      margin: 0;
      flex: 1;
    }}

    .case-citation-banner {{
      background: rgba(56, 189, 248, 0.08);
      border-left: 3px solid var(--accent-blue);
      border-radius: 0 8px 8px 0;
      padding: 0.6rem 1rem;
      margin-bottom: 1.5rem;
      font-family: var(--font-ui);
      font-size: 0.88rem;
      color: var(--accent-blue);
      display: flex;
      align-items: center;
      gap: 0.6rem;
    }}

    .citation-icon {{
      font-size: 1.1rem;
    }}

    /* Subheadings within Case (Facts, Issue, Ruling) */
    .case-subheading {{
      font-family: var(--font-ui);
      font-size: 0.95rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: var(--accent-gold);
      margin: 1.75rem 0 0.85rem 0;
      padding-bottom: 0.35rem;
      border-bottom: 1px solid rgba(251, 191, 36, 0.25);
      display: flex;
      align-items: center;
      gap: 0.45rem;
    }}

    .subheading-accent {{
      color: var(--accent-gold);
      font-size: 1rem;
    }}

    /* Continuous Flowing Body Paragraphs */
    .case-paragraph {{
      margin-bottom: 1.25rem;
      font-size: 1.05rem;
      line-height: 1.85;
      color: var(--text-primary);
      text-align: justify;
      text-justify: inter-word;
    }}

    .read-unit {{
      position: relative;
      padding: 0.5rem 0.75rem;
      border-radius: 8px;
      cursor: pointer;
      transition: background-color 0.2s ease, box-shadow 0.2s ease, transform 0.15s ease;
    }}

    .read-unit:hover {{
      background-color: rgba(148, 163, 184, 0.07);
    }}

    .read-unit.is-speaking {{
      background-color: var(--highlight-reading) !important;
      border-left: 3px solid var(--accent-blue);
      box-shadow: 0 0 14px rgba(56, 189, 248, 0.25);
    }}

    /* ALAC and Reasoning Inline Badges */
    .alac-paragraph {{
      display: block;
    }}

    .alac-badge {{
      display: inline-block;
      font-family: var(--font-ui);
      font-size: 0.72rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      padding: 0.2rem 0.55rem;
      border-radius: 4px;
      margin-right: 0.45rem;
      vertical-align: baseline;
    }}

    .badge-ans {{ background: rgba(52, 211, 153, 0.18); color: var(--accent-emerald); border: 1px solid rgba(52, 211, 153, 0.35); }}
    .badge-law {{ background: rgba(56, 189, 248, 0.18); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.35); }}
    .badge-app {{ background: rgba(129, 140, 248, 0.18); color: var(--accent-indigo); border: 1px solid rgba(129, 140, 248, 0.35); }}
    .badge-con {{ background: rgba(251, 113, 133, 0.18); color: var(--accent-rose); border: 1px solid rgba(251, 113, 133, 0.35); }}
    .badge-syn {{ background: rgba(251, 191, 36, 0.18); color: var(--accent-gold); border: 1px solid rgba(251, 191, 36, 0.35); }}
    .badge-gen {{ background: rgba(148, 163, 184, 0.18); color: var(--text-secondary); border: 1px solid rgba(148, 163, 184, 0.35); }}

    /* Lists */
    .bullet-point {{
      display: flex;
      align-items: flex-start;
      gap: 0.65rem;
      margin-bottom: 0.75rem;
      font-size: 1.05rem;
      line-height: 1.75;
    }}

    .bullet-dot {{
      color: var(--accent-gold);
      font-weight: bold;
      font-size: 1.2rem;
      line-height: 1.5;
    }}

    .bullet-content {{
      flex: 1;
    }}

    /* Tables */
    .table-responsive {{
      overflow-x: auto;
      margin: 1.5rem 0;
      border-radius: 8px;
      border: 1px solid var(--border-color);
    }}

    .reader-table {{
      width: 100%;
      border-collapse: collapse;
      font-family: var(--font-ui);
      font-size: 0.9rem;
    }}

    .reader-table th, .reader-table td {{
      padding: 0.85rem 1.15rem;
      border: 1px solid var(--border-color);
      text-align: left;
      line-height: 1.6;
    }}

    .reader-table th {{
      background-color: var(--bg-secondary);
      color: var(--accent-gold);
      font-weight: 700;
    }}

    .reader-table tr:nth-child(even) {{
      background-color: rgba(148, 163, 184, 0.04);
    }}

    /* Floating TTS Tooltip */
    #floatingTtsTrigger {{
      position: absolute;
      display: none;
      z-index: 1000;
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      color: #0b1120;
      font-family: var(--font-ui);
      font-size: 0.78rem;
      font-weight: 700;
      border: none;
      border-radius: 20px;
      padding: 0.4rem 0.85rem;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
      cursor: pointer;
      transform: translateX(-50%);
      transition: transform 0.15s ease, filter 0.15s ease;
    }}

    #floatingTtsTrigger:hover {{
      transform: translateX(-50%) scale(1.05);
      filter: brightness(1.1);
    }}

    /* Footer */
    footer.app-footer {{
      background-color: var(--bg-secondary);
      border-top: 1px solid var(--border-color);
      padding: 1.5rem;
      text-align: center;
      font-family: var(--font-ui);
      font-size: 0.82rem;
      color: var(--text-muted);
      margin-top: auto;
    }}

    @media (max-width: 900px) {{
      aside.app-sidebar {{
        position: fixed;
        left: 0;
        top: var(--header-height);
        transform: translateX(-100%);
        box-shadow: var(--shadow-lg);
      }}

      aside.app-sidebar.open {{
        transform: translateX(0);
      }}

      .brand-title {{
        max-width: 160px;
      }}

      .tts-status-badge {{
        display: none;
      }}
    }}
  </style>
</head>
<body>
  <div id="readingProgressBar"></div>

  <!-- App Header -->
  <header class="app-header">
    <div class="brand-section">
      <button class="icon-btn" id="toggleSidebarBtn" title="Toggle Table of Contents">☰</button>
      <div class="brand-logo">§</div>
      <div class="brand-info">
        <span class="brand-title">{escaped_title}</span>
        <span class="brand-subject">{escaped_subject_tag}</span>
      </div>
    </div>

    <!-- Speech Synthesis Toolbar -->
    <div class="tts-toolbar">
      <div class="toolbar-group">
        <button class="icon-btn" id="prevBtn" title="Previous (P)">⏮</button>
        <button class="icon-btn primary" id="playPauseBtn" title="Play / Pause (Space)">
          <span id="playIcon">▶</span>
          <span id="pauseIcon" style="display:none;">⏸</span>
        </button>
        <button class="icon-btn" id="stopBtn" title="Stop (Esc)">⏹</button>
        <button class="icon-btn" id="nextBtn" title="Next (N)">⏭</button>
      </div>

      <div class="toolbar-group">
        <select id="speedSelect" class="toolbar-select" title="Reading Speed">
          <option value="0.75">0.75x</option>
          <option value="0.9" selected>0.90x (Deliberate)</option>
          <option value="1.0">1.00x</option>
          <option value="1.15">1.15x</option>
          <option value="1.25">1.25x</option>
          <option value="1.5">1.50x</option>
        </select>
        <select id="voiceSelect" class="toolbar-select" title="Select Natural Voice" style="max-width: 130px;">
          <option value="default">Auto Voice</option>
        </select>
      </div>

      <div class="toolbar-group">
        <button class="icon-btn" id="fontDecBtn" title="Decrease Font (-)">A-</button>
        <button class="icon-btn" id="fontIncBtn" title="Increase Font (+)">A+</button>
        <select id="themeSelect" class="toolbar-select" title="Switch Theme">
          <option value="dark">Dark</option>
          <option value="sepia">Sepia</option>
          <option value="light">Light</option>
        </select>
      </div>

      <span id="ttsStatusBadge" class="tts-status-badge">Ready</span>
    </div>
  </header>

  <!-- Main Body Layout -->
  <div class="app-main-layout">
    <!-- Table of Contents Sidebar -->
    <aside class="app-sidebar" id="sidebarNav">
      <div class="sidebar-header">
        <input type="text" id="sidebarSearch" class="sidebar-search" placeholder="🔍 Filter Case Titles..." />
      </div>
      <nav class="sidebar-nav">
        <ul id="tocList">
          {toc_html}
        </ul>
      </nav>
    </aside>

    <!-- Reading Content Area -->
    <main class="content-area">
      <div class="content-wrapper" id="contentWrapper">
        <!-- Hero Header Card -->
        <div class="doc-hero-card">
          <div class="doc-hero-subject">{escaped_subject_tag}</div>
          <h1 class="doc-hero-title">{escaped_title}</h1>
          <div class="doc-meta-stats">
            <span>📚 {total_sections} Cases / Sections</span>
            <span>📖 {total_units} Narrative Units</span>
            <span>⏱ ~{reading_time_minutes} min study read</span>
          </div>
          {mp3_player_html}
        </div>

        <!-- Rendered Sections -->
        {sections_html}
      </div>
    </main>
  </div>

  <button id="floatingTtsTrigger">🔊 Read Selection</button>

  <footer class="app-footer">
    <div>Manila Law College (MLC) • Juris Doctor Compendium & Interactive Audio Suite</div>
  </footer>

  <script>
    (function() {{
      const synth = window.speechSynthesis;
      const playPauseBtn = document.getElementById('playPauseBtn');
      const playIcon = document.getElementById('playIcon');
      const pauseIcon = document.getElementById('pauseIcon');
      const stopBtn = document.getElementById('stopBtn');
      const prevBtn = document.getElementById('prevBtn');
      const nextBtn = document.getElementById('nextBtn');
      const speedSelect = document.getElementById('speedSelect');
      const voiceSelect = document.getElementById('voiceSelect');
      const themeSelect = document.getElementById('themeSelect');
      const fontIncBtn = document.getElementById('fontIncBtn');
      const fontDecBtn = document.getElementById('fontDecBtn');
      const statusBadge = document.getElementById('ttsStatusBadge');
      const toggleSidebarBtn = document.getElementById('toggleSidebarBtn');
      const sidebarNav = document.getElementById('sidebarNav');
      const sidebarSearch = document.getElementById('sidebarSearch');
      const tocList = document.getElementById('tocList');
      const progressBar = document.getElementById('readingProgressBar');
      const readUnits = Array.from(document.querySelectorAll('.read-unit'));

      let voices = [];
      let currentUnitIndex = -1;
      let isPaused = false;
      let isSpeaking = false;
      let lastSectionId = null;

      // 1. Sidebar Toggle & Instant Live Filter
      toggleSidebarBtn.addEventListener('click', () => {{
        sidebarNav.classList.toggle('open');
      }});

      sidebarSearch.addEventListener('input', (e) => {{
        const q = e.target.value.toLowerCase();
        const items = tocList.querySelectorAll('li');
        items.forEach(li => {{
          const txt = li.textContent.toLowerCase();
          li.style.display = txt.includes(q) ? '' : 'none';
        }});
      }});

      // 2. Theme & Font Scaling
      themeSelect.addEventListener('change', (e) => {{
        document.documentElement.setAttribute('data-theme', e.target.value);
        localStorage.setItem('mlc_theme', e.target.value);
      }});
      const savedTheme = localStorage.getItem('mlc_theme');
      if (savedTheme) {{
        themeSelect.value = savedTheme;
        document.documentElement.setAttribute('data-theme', savedTheme);
      }}

      let currentFontSize = 17;
      fontIncBtn.addEventListener('click', () => {{
        if (currentFontSize < 24) {{
          currentFontSize += 1;
          document.documentElement.style.fontSize = currentFontSize + 'px';
        }}
      }});
      fontDecBtn.addEventListener('click', () => {{
        if (currentFontSize > 13) {{
          currentFontSize -= 1;
          document.documentElement.style.fontSize = currentFontSize + 'px';
        }}
      }});

      // 3. Scroll Progress & Active TOC Link Sync
      window.addEventListener('scroll', () => {{
        const sTop = window.scrollY;
        const dHeight = document.documentElement.scrollHeight - window.innerHeight;
        if (dHeight > 0) {{
          progressBar.style.width = ((sTop / dHeight) * 100) + '%';
        }}

        // Sync TOC Active Link
        const sections = document.querySelectorAll('section.doc-section');
        let currentActive = null;
        sections.forEach(sec => {{
          const rect = sec.getBoundingClientRect();
          if (rect.top <= 140 && rect.bottom >= 140) {{
            currentActive = sec.id;
          }}
        }});
        if (currentActive) {{
          document.querySelectorAll('.toc-link').forEach(link => {{
            link.classList.toggle('active', link.getAttribute('href') === '#' + currentActive);
          }});
        }}
      }});

      // 4. Voice Population with High-Quality Natural Neural Prioritization
      function populateVoices() {{
        if (!synth) return;
        voices = synth.getVoices();
        voiceSelect.innerHTML = '<option value="default">Auto Best Voice</option>';

        // Score voices: prioritize Neural, Natural, English-US / UK / AU
        const scored = voices.map((v, i) => {{
          let score = 0;
          const name = v.name.toLowerCase();
          if (name.includes('natural') || name.includes('neural') || name.includes('online')) score += 50;
          if (name.includes('jenny') || name.includes('guy') || name.includes('aria') || name.includes('google')) score += 30;
          if (v.lang.startsWith('en')) score += 20;
          return {{ voice: v, index: i, score }};
        }}).sort((a, b) => b.score - a.score);

        scored.forEach(item => {{
          if (item.voice.lang.startsWith('en') || item.score > 20) {{
            const opt = document.createElement('option');
            opt.value = item.index;
            opt.textContent = `${{item.voice.name}} (${{item.voice.lang}})`;
            voiceSelect.appendChild(opt);
          }}
        }});
      }}

      if (synth) {{
        populateVoices();
        if (speechSynthesis.onvoiceschanged !== undefined) {{
          speechSynthesis.onvoiceschanged = populateVoices;
        }}
      }}

      // 5. Smart Legal Phonetic Expander
      function prepareSpeechText(unitEl) {{
        let text = unitEl.innerText || unitEl.textContent || '';
        text = text
          .replace(/\\bSCRA\\b/g, 'S-C-R-A')
          .replace(/\\bPhil\\.\\s*(\\d+)/gi, 'Philippine Reports volume $1')
          .replace(/\\bArt\\.\\s*(\\d+)/gi, 'Article $1')
          .replace(/\\bArts\\.\\s*([\\d,\\s\\-]+)/gi, 'Articles $1')
          .replace(/\\bSec\\.\\s*(\\d+)/gi, 'Section $1')
          .replace(/\\bSecs\\.\\s*([\\d,\\s\\-]+)/gi, 'Sections $1')
          .replace(/\\bPar\\.\\s*(\\d+)/gi, 'Paragraph $1')
          .replace(/\\bRPC\\b/g, 'Revised Penal Code')
          .replace(/\\bCPR\\b/g, 'Code of Professional Responsibility')
          .replace(/\\bCPRA\\b/g, 'Code of Professional Responsibility and Accountability')
          .replace(/\\bConst\\.\\b/gi, 'Constitution')
          .replace(/\\bP\\.D\\.\\s*No\\.?\\s*(\\d+)/gi, 'Presidential Decree Number $1')
          .replace(/\\bR\\.A\\.\\s*No\\.?\\s*(\\d+)/gi, 'Republic Act Number $1')
          .replace(/\\bA\\.C\\.\\s*No\\.?\\s*([A-Za-z0-9\\-]+)/gi, 'Administrative Case Number $1')
          .replace(/\\bA\\.M\\.\\s*No\\.?\\s*([A-Za-z0-9\\-]+)/gi, 'Administrative Matter Number $1')
          .replace(/\\bG\\.R\\.\\s*No\\.?\\s*([A-Za-z0-9\\-]+)/gi, 'G-R Number $1')
          .replace(/\\bet\\s+al\\./gi, 'and others')
          .replace(/\\bi\\.e\\./gi, 'that is')
          .replace(/\\be\\.g\\./gi, 'for example')
          .replace(/\\bviz\\./gi, 'namely')
          .replace(/\\s+v(?:s)?\\.\\s+/gi, ' versus ')
          .replace(/;\\s*/g, ';, ')
          .replace(/:\\s*/g, ': ... ');

        return text.trim();
      }}

      function setSpeakingState(speaking, paused = false) {{
        isSpeaking = speaking;
        isPaused = paused;
        if (speaking && !paused) {{
          playIcon.style.display = 'none';
          pauseIcon.style.display = 'inline';
          statusBadge.textContent = 'Speaking...';
          statusBadge.className = 'tts-status-badge speaking';
        }} else if (paused) {{
          playIcon.style.display = 'inline';
          pauseIcon.style.display = 'none';
          statusBadge.textContent = 'Paused';
          statusBadge.className = 'tts-status-badge';
        }} else {{
          playIcon.style.display = 'inline';
          pauseIcon.style.display = 'none';
          statusBadge.textContent = 'Ready';
          statusBadge.className = 'tts-status-badge';
          clearHighlights();
        }}
      }}

      function clearHighlights() {{
        readUnits.forEach(u => u.classList.remove('is-speaking'));
      }}

      function highlightUnit(idx) {{
        clearHighlights();
        if (idx >= 0 && idx < readUnits.length) {{
          const unit = readUnits[idx];
          unit.classList.add('is-speaking');
          unit.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
        }}
      }}

      function speakUnit(idx) {{
        if (!synth || idx < 0 || idx >= readUnits.length) {{
          setSpeakingState(false);
          return;
        }}

        synth.cancel();
        currentUnitIndex = idx;
        const unit = readUnits[idx];
        highlightUnit(idx);

        let speechText = prepareSpeechText(unit);
        const currentSec = unit.closest('section.doc-section');
        const currentSecId = currentSec ? currentSec.id : null;
        if (currentSecId && currentSecId !== lastSectionId && currentSec) {{
          const headerEl = currentSec.querySelector('.case-header-title') || currentSec.querySelector('.section-title');
          if (headerEl) {{
            speechText = 'Now Reading: ' + headerEl.innerText.trim() + ' ... ... ' + speechText;
          }}
          lastSectionId = currentSecId;
        }}

        const utterance = new SpeechSynthesisUtterance(speechText);
        utterance.rate = parseFloat(speedSelect.value) || 0.9;
        utterance.pitch = 1.0;

        const selVoiceIdx = voiceSelect.value;
        if (selVoiceIdx !== 'default' && voices[selVoiceIdx]) {{
          utterance.voice = voices[selVoiceIdx];
        }} else if (voices.length > 0) {{
          utterance.voice = voices[0];
        }}

        utterance.onstart = () => {{
          setSpeakingState(true, false);
        }};

        utterance.onend = () => {{
          if (isSpeaking && !isPaused) {{
            if (idx + 1 < readUnits.length) {{
              speakUnit(idx + 1);
            }} else {{
              setSpeakingState(false);
            }}
          }}
        }};

        utterance.onerror = (e) => {{
          if (e.error !== 'interrupted' && e.error !== 'canceled') {{
            console.warn('TTS error:', e);
          }}
        }};

        synth.speak(utterance);
      }}

      playPauseBtn.addEventListener('click', () => {{
        if (!synth) return;
        if (isSpeaking && !isPaused) {{
          synth.pause();
          setSpeakingState(true, true);
        }} else if (isPaused) {{
          synth.resume();
          setSpeakingState(true, false);
        }} else {{
          speakUnit(currentUnitIndex >= 0 ? currentUnitIndex : 0);
        }}
      }});

      stopBtn.addEventListener('click', () => {{
        if (!synth) return;
        synth.cancel();
        setSpeakingState(false);
        currentUnitIndex = -1;
        lastSectionId = null;
      }});

      nextBtn.addEventListener('click', () => {{
        const target = Math.min(readUnits.length - 1, (currentUnitIndex >= 0 ? currentUnitIndex : 0) + 1);
        speakUnit(target);
      }});

      prevBtn.addEventListener('click', () => {{
        const target = Math.max(0, (currentUnitIndex >= 0 ? currentUnitIndex : 0) - 1);
        speakUnit(target);
      }});

      // Click on any paragraph to start speaking from that exact unit
      document.getElementById('contentWrapper').addEventListener('click', (e) => {{
        const unit = e.target.closest('.read-unit');
        if (unit) {{
          const idx = readUnits.indexOf(unit);
          if (idx !== -1) {{
            speakUnit(idx);
          }}
        }}
      }});

      // Floating Selection Reader
      const floatBtn = document.getElementById('floatingTtsTrigger');
      document.addEventListener('mouseup', (e) => {{
        if (e.target.closest('#floatingTtsTrigger') || e.target.closest('.tts-toolbar') || e.target.closest('header')) return;
        const sel = window.getSelection().toString().trim();
        if (sel.length > 2) {{
          const rect = window.getSelection().getRangeAt(0).getBoundingClientRect();
          floatBtn.style.top = (window.scrollY + rect.top - 44) + 'px';
          floatBtn.style.left = (window.scrollX + rect.left + (rect.width / 2)) + 'px';
          floatBtn.style.display = 'block';
        }} else {{
          floatBtn.style.display = 'none';
        }}
      }});

      floatBtn.addEventListener('click', () => {{
        const sel = window.getSelection().toString().trim();
        if (sel && synth) {{
          synth.cancel();
          const dummy = document.createElement('div');
          dummy.innerText = sel;
          const speechText = prepareSpeechText(dummy);

          const utterance = new SpeechSynthesisUtterance(speechText);
          utterance.rate = parseFloat(speedSelect.value) || 0.9;
          const selVoiceIdx = voiceSelect.value;
          if (selVoiceIdx !== 'default' && voices[selVoiceIdx]) {{
            utterance.voice = voices[selVoiceIdx];
          }}
          synth.speak(utterance);
          setSpeakingState(true);
          floatBtn.style.display = 'none';
        }}
      }});

      // Keyboard Shortcuts
      document.addEventListener('keydown', (e) => {{
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
        if (e.code === 'Space') {{
          e.preventDefault();
          playPauseBtn.click();
        }} else if (e.code === 'Escape') {{
          stopBtn.click();
        }} else if (e.key === 'n' || e.key === 'N') {{
          nextBtn.click();
        }} else if (e.key === 'p' || e.key === 'P') {{
          prevBtn.click();
        }} else if (e.key === 'd' || e.key === 'D') {{
          themeSelect.value = 'dark';
          themeSelect.dispatchEvent(new Event('change'));
        }} else if (e.key === 's' || e.key === 'S') {{
          themeSelect.value = 'sepia';
          themeSelect.dispatchEvent(new Event('change'));
        }} else if (e.key === 'l' || e.key === 'L') {{
          themeSelect.value = 'light';
          themeSelect.dispatchEvent(new Event('change'));
        }} else if (e.key === '+' || e.key === '=') {{
          fontIncBtn.click();
        }} else if (e.key === '-') {{
          fontDecBtn.click();
        }} else if (e.key === '/') {{
          e.preventDefault();
          sidebarSearch.focus();
        }}
      }});
    }})();
  </script>
</body>
</html>
"""

# ==============================================================================
# 4. COMPILER & BATCH PROCESSOR
# ==============================================================================

def infer_subject_and_title(file_path):
    path_obj = Path(file_path).resolve()
    filename_stem = path_obj.stem
    parts = list(path_obj.parts)
    subject = "Law Study Notes"
    for p in reversed(parts[:-1]):
        if any(term in p.lower() for term in ["criminal", "statutory", "blje", "ethics", "procedure", "civil", "consti"]):
            subject = p
            break
            
    clean_title = filename_stem.replace("_", " ").replace("-", " ")
    clean_title = re.sub(r'\s+', ' ', clean_title).strip()
    return subject, clean_title

def convert_file_to_html_reader(file_path, output_path=None, overwrite=True):
    path_obj = Path(file_path).resolve()
    if not path_obj.exists():
        print(f"[ERROR] File not found: {path_obj}")
        return None

    if output_path is None:
        output_path = path_obj.with_suffix('.html')
    else:
        output_path = Path(output_path).resolve()

    ext = path_obj.suffix.lower()
    subject_tag, doc_title = infer_subject_and_title(path_obj)

    if output_path.exists() and not overwrite:
        print(f"[SKIP] Output exists: {output_path}")
        return output_path

    print(f"[PROCESSING] {path_obj.name} -> {output_path.name}")

    sections = []
    if ext == '.docx':
        sections = parse_docx_file(str(path_obj), doc_title)
    elif ext == '.pdf':
        sections = parse_pdf_file(str(path_obj), doc_title)
    elif ext == '.md':
        with open(path_obj, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        sections = parse_markdown(content, doc_title)
    elif ext == '.txt':
        sections = parse_txt_file(str(path_obj), doc_title)
    else:
        print(f"[WARN] Unsupported extension: {ext}")
        return None

    if not sections:
        print(f"[WARN] No sections extracted from {path_obj.name}")
        return None

    # Build TOC HTML & Sections HTML
    toc_items = []
    sec_html_list = []
    total_units = 0

    for s in sections:
        sec_id = s.get("id", "sec-main")
        sec_title = s.get("title", "Section")
        level = s.get("level", 1)
        units = s.get("units", [])
        total_units += len(units)

        display_title = sec_title
        if len(display_title) > 65:
            display_title = display_title[:62] + "..."

        toc_items.append(
            f'<li><a href="#{sec_id}" class="toc-link level-{level}" title="{html.escape(sec_title)}">{html.escape(display_title)}</a></li>'
        )

        rendered_units = "\n        ".join(units)
        sec_html_list.append(f'''
        <section class="doc-section" id="{sec_id}">
          <div class="section-body">
            {rendered_units}
          </div>
        </section>''')

    toc_html = "\n            ".join(toc_items)
    sections_html = "\n".join(sec_html_list)

    reading_time = max(1, round(total_units * 35 / 160))

    # Check companion MP3 audio podcast
    mp3_file = path_obj.with_suffix('.mp3')
    mp3_player_html = ''
    if mp3_file.exists():
        mp3_name_esc = html.escape(mp3_file.name)
        mp3_size_mb = round(mp3_file.stat().st_size / (1024 * 1024), 1)
        mp3_player_html = f'''
          <div class="mp3-podcast-card" style="margin-top: 1.5rem; padding: 1.15rem 1.35rem; background: rgba(56, 189, 248, 0.12); border: 1px solid rgba(56, 189, 248, 0.35); border-radius: 12px;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 0.75rem; flex-wrap: wrap; gap: 0.5rem;">
              <div style="display: flex; align-items: center; gap: 0.5rem;">
                <span style="font-size: 1.15rem;">🎙️</span>
                <span style="font-family: var(--font-ui); font-weight: 700; font-size: 0.95rem; color: var(--accent-blue);">Master Studio Audio Podcast ({mp3_size_mb} MB)</span>
                <span style="font-family: var(--font-ui); font-size: 0.72rem; padding: 0.2rem 0.55rem; background: rgba(52, 211, 153, 0.2); color: var(--accent-emerald); border-radius: 12px; font-weight: 600;">Mobile Background Audio</span>
              </div>
              <a href="{mp3_name_esc}" download style="font-family: var(--font-ui); font-size: 0.82rem; color: var(--accent-gold); text-decoration: none; font-weight: 700;">⬇ Download MP3</a>
            </div>
            <audio controls preload="metadata" style="width: 100%; height: 42px; border-radius: 8px; outline: none;">
              <source src="{mp3_name_esc}" type="audio/mpeg">
              Your browser does not support audio playback.
            </audio>
            <div style="font-family: var(--font-ui); font-size: 0.75rem; color: var(--text-muted); margin-top: 0.4rem;">
              💡 <em>Plays continuously with phone screen locked or app minimized. Perfect for mobile commute listening.</em>
            </div>
          </div>
        '''

    rendered_page = HTML_PAGE_TEMPLATE.format(
        escaped_title=html.escape(doc_title),
        escaped_description=html.escape(f"{subject_tag} comprehensive review notes and study materials."),
        escaped_subject_tag=html.escape(subject_tag),
        total_sections=len(sections),
        total_units=total_units,
        reading_time_minutes=reading_time,
        toc_html=toc_html,
        sections_html=sections_html,
        mp3_player_html=mp3_player_html
    )

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(rendered_page)

    print(f"[SUCCESS] Created HTML Reader: {output_path.name} ({total_units} units, {len(sections)} TOC sections)")
    return output_path

def scan_and_convert_directory(dir_path, recursive=True, overwrite=True):
    target_dir = Path(dir_path).resolve()
    print(f"\n=== Scanning directory: {target_dir} ===")
    supported_extensions = {'.docx', '.pdf', '.md', '.txt'}
    ignore_stems = {'readme', 'handoff', 'format_dump', 'study_hub', 'index'}
    
    pattern = "**/*" if recursive else "*"
    all_files = [p for p in target_dir.glob(pattern) if p.is_file() and p.suffix.lower() in supported_extensions]

    # Group by directory and stem to prioritize .docx over .pdf
    grouped = {}
    for p in all_files:
        if '__pycache__' in str(p) or '.agents' in str(p) or p.name.startswith(('~', '.', 'Unconfirmed')):
            continue
        if p.stem.lower() in ignore_stems and p.suffix.lower() == '.md':
            continue

        key = (p.parent, p.stem.lower())
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(p)

    converted_count = 0
    generated_files = []

    for key, file_list in grouped.items():
        # Preference: .docx > .md > .txt > .pdf
        chosen = None
        for ext in ['.docx', '.md', '.txt', '.pdf']:
            match = next((f for f in file_list if f.suffix.lower() == ext), None)
            if match:
                chosen = match
                break

        if chosen:
            try:
                out = convert_file_to_html_reader(chosen, overwrite=overwrite)
                if out:
                    converted_count += 1
                    generated_files.append(out)
            except Exception as e:
                print(f"[ERROR] Failed to convert {chosen.name}: {e}")

    print(f"\n=== Completed! Successfully converted {converted_count} documents. ===\n")
    return generated_files

def generate_study_hub_index(root_dir):
    root_path = Path(root_dir).resolve()
    html_files = list(root_path.glob("**/*.html"))
    html_files = [f for f in html_files if f.name.lower() not in {'mlc_study_hub.html', 'index.html'}]

    by_folder = {}
    for h in sorted(html_files, key=lambda x: str(x)):
        rel = h.relative_to(root_path)
        subject_group = rel.parent.name or "Core Library"
        if subject_group not in by_folder:
            by_folder[subject_group] = []
        by_folder[subject_group].append((h, rel))

    cards_html = []
    for group_name, files in by_folder.items():
        cards_html.append(f'<div class="hub-group"><h2 class="hub-group-title">📂 {html.escape(group_name)}</h2><div class="hub-grid">')
        for fpath, rel_path in files:
            doc_name = fpath.stem.replace('_', ' ')
            rel_str = str(rel_path).replace('\\\\', '/').replace('\\', '/')
            mp3_file = fpath.with_suffix('.mp3')
            has_mp3 = mp3_file.exists()
            mp3_badge = '<span class="badge-audio">🎙️ MP3 Audio</span>' if has_mp3 else ''
            
            cards_html.append(f'''
              <div class="hub-card">
                <div class="hub-card-header">
                  <span class="badge-reader">HTML Reader</span>
                  {mp3_badge}
                </div>
                <h3 class="hub-card-title"><a href="{rel_str}">{html.escape(doc_name)}</a></h3>
                <div class="hub-card-actions">
                  <a href="{rel_str}" class="btn-open">Open Reader →</a>
                </div>
              </div>
            ''')
        cards_html.append('</div></div>')

    hub_page = f'''<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>MLC Interactive Study Hub • Natural Speech Readers</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@700;800;900&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --bg-primary: #0b1120;
      --bg-secondary: #131d31;
      --bg-tertiary: #1e293b;
      --border-color: rgba(148, 163, 184, 0.16);
      --text-primary: #f8fafc;
      --text-secondary: #cbd5e1;
      --text-muted: #94a3b8;
      --accent-gold: #fbbf24;
      --accent-blue: #38bdf8;
      --font-ui: 'Inter', -apple-system, sans-serif;
      --font-heading: 'Cinzel', serif;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      background-color: var(--bg-primary);
      color: var(--text-primary);
      font-family: var(--font-ui);
      padding: 3rem 1.5rem;
      min-height: 100vh;
    }}
    .hub-container {{ max-width: 1100px; margin: 0 auto; }}
    .hub-header {{ text-align: center; margin-bottom: 3.5rem; }}
    .hub-title {{ font-family: var(--font-heading); font-size: 2.2rem; font-weight: 800; color: var(--accent-gold); margin-bottom: 0.5rem; }}
    .hub-subtitle {{ color: var(--text-muted); font-size: 1.05rem; }}
    .hub-group {{ margin-bottom: 2.5rem; }}
    .hub-group-title {{ font-size: 1.25rem; font-weight: 700; color: var(--text-secondary); margin-bottom: 1.25rem; border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem; }}
    .hub-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 1.25rem; }}
    .hub-card {{ background-color: var(--bg-secondary); border: 1px solid var(--border-color); border-radius: 12px; padding: 1.5rem; display: flex; flex-direction: column; justify-content: space-between; transition: all 0.2s ease; }}
    .hub-card:hover {{ transform: translateY(-3px); border-color: var(--accent-blue); box-shadow: 0 8px 20px rgba(0, 0, 0, 0.4); }}
    .hub-card-header {{ display: flex; gap: 0.5rem; margin-bottom: 0.75rem; }}
    .badge-reader {{ background: rgba(56, 189, 248, 0.15); color: var(--accent-blue); font-size: 0.72rem; font-weight: 700; padding: 0.2rem 0.55rem; border-radius: 6px; }}
    .badge-audio {{ background: rgba(52, 211, 153, 0.15); color: #34d399; font-size: 0.72rem; font-weight: 700; padding: 0.2rem 0.55rem; border-radius: 6px; }}
    .hub-card-title {{ font-size: 1.05rem; font-weight: 600; line-height: 1.4; margin-bottom: 1.25rem; }}
    .hub-card-title a {{ color: var(--text-primary); text-decoration: none; }}
    .hub-card-title a:hover {{ color: var(--accent-gold); }}
    .btn-open {{ display: inline-block; background: linear-gradient(135deg, var(--accent-gold), #d97706); color: #0b1120; font-weight: 700; font-size: 0.85rem; padding: 0.5rem 1rem; border-radius: 8px; text-decoration: none; text-align: center; transition: filter 0.15s ease; }}
    .btn-open:hover {{ filter: brightness(1.1); }}
  </style>
</head>
<body>
  <div class="hub-container">
    <header class="hub-header">
      <h1 class="hub-title">MLC Law Library & Audio Hub</h1>
      <p class="hub-subtitle">Interactive Full-Text Readers with Natural Voice Synthesis & Studio MP3 Podcasts</p>
    </header>
    {''.join(cards_html)}
  </div>
</body>
</html>
'''

    out_hub = root_path / "MLC_Study_Hub.html"
    out_idx = root_path / "index.html"
    with open(out_hub, 'w', encoding='utf-8') as f:
        f.write(hub_page)
    with open(out_idx, 'w', encoding='utf-8') as f:
        f.write(hub_page)
    print(f"[SUCCESS] Updated Master Hub: {out_hub.name} and {out_idx.name}")

# ==============================================================================
# 5. CLI ENTRYPOINT
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(description="MLC Universal Text-to-Speech HTML Reader Generator")
    parser.add_argument("input_path", nargs="?", default=None, help="File or directory to convert")
    parser.add_argument("--dir", default=None, help="Scan and convert entire directory recursively")
    parser.add_argument("--hub", action="store_true", help="Generate Master Hub index HTML")
    parser.add_argument("--no-overwrite", action="store_true", help="Skip already converted files")
    args = parser.parse_args()

    root_mlc = Path(__file__).parent.resolve()

    if args.dir:
        scan_and_convert_directory(args.dir, overwrite=not args.no_overwrite)
        generate_study_hub_index(root_mlc)
    elif args.hub:
        generate_study_hub_index(root_mlc)
    elif args.input_path:
        p = Path(args.input_path).resolve()
        if p.is_dir():
            scan_and_convert_directory(p, overwrite=not args.no_overwrite)
        else:
            convert_file_to_html_reader(p, overwrite=not args.no_overwrite)
        generate_study_hub_index(root_mlc)
    else:
        scan_and_convert_directory(root_mlc, overwrite=not args.no_overwrite)
        generate_study_hub_index(root_mlc)

if __name__ == "__main__":
    main()
