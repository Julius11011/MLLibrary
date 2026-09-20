#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MLC Universal Text-to-Speech (Read Aloud) HTML Reader Generator
===============================================================
Converts legal study materials (.md, .docx, .txt, .pdf) across all MLC subfolders
into interactive, high-readability HTML web applications featuring:
  - Natural, Human-like Cadence & Intonation
  - Smart Punctuation & Clause Pauses (commas, semicolons, dashes, periods)
  - Clear Context Transition Pauses between major topics, subtopics, and case briefs
  - Audio Labeling of Topics, Subtopics, Facts, Issues, Held, and Doctrines
  - Smart Phonetic Legal Abbreviation Expander (SCRA, G.R. No., Art., Sec., RPC, etc.)
  - Intelligent Neural/Natural Voice Auto-Selector (Jenny, Guy, Aria, Google Natural, etc.)
  - Optimized deliberate study pacing (default 0.9x rate)
  - Dark, Sepia, and Light Themes + Font Scaling
  - Table of Contents Sidebar with Instant Search Filter
  - Keyboard Navigation (Space, Esc, N, P, D, S, L, +, -)
  - Full compatibility with browser extensions (Brave Read Aloud, Speechify, Edge Read Aloud)
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
    text = text.replace("“", '"').replace("”", '"')
    text = text.replace("’", "'").replace("‘", "'")
    text = text.replace("—", " — ").replace("–", " – ")
    return text

def slugify(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9\-_\s]', '', text)
    text = re.sub(r'\s+', '-', text)
    return text[:60].strip('-') or 'section'

# ==============================================================================
# 2. DOCUMENT PARSERS (.md, .docx, .txt, .pdf)
# ==============================================================================

def parse_markdown(content, doc_title=""):
    """Parses Markdown text into structured sections and units."""
    lines = content.splitlines()
    sections = []
    current_sec = {
        "id": "sec-intro",
        "title": doc_title or "Introduction",
        "level": 1,
        "units": []
    }
    
    in_code_block = False
    code_lines = []
    code_lang = ""
    
    in_table = False
    table_rows = []

    def flush_table():
        nonlocal in_table, table_rows, current_sec
        if not table_rows:
            in_table = False
            return
        
        # Build HTML table
        header_row = table_rows[0]
        data_rows = table_rows[1:] if len(table_rows) > 1 else []
        
        # Check if second row is separator
        if data_rows and all(re.match(r'^[\s\-:\|]+$', cell) for cell in data_rows[0]):
            data_rows = data_rows[1:]
            
        html_tbl = ['<div class="table-responsive read-unit" data-unit-type="table"><table class="reader-table">']
        if header_row:
            html_tbl.append('<thead><tr>')
            for c in header_row:
                html_tbl.append(f'<th>{html.escape(c.strip())}</th>')
            html_tbl.append('</tr></thead>')
        if data_rows:
            html_tbl.append('<tbody>')
            for r in data_rows:
                html_tbl.append('<tr>')
                for c in r:
                    html_tbl.append(f'<td>{format_inline(c.strip())}</td>')
                html_tbl.append('</tr>')
            html_tbl.append('</tbody>')
        html_tbl.append('</table></div>')
        current_sec["units"].append("".join(html_tbl))
        table_rows = []
        in_table = False

    def format_inline(t):
        t = clean_text(t)
        t = html.escape(t)
        # Bold
        t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
        t = re.sub(r'__(.+?)__', r'<strong>\1</strong>', t)
        # Italic
        t = re.sub(r'\*(.+?)\*', r'<em>\1</em>', t)
        t = re.sub(r'_(.+?)_', r'<em>\1</em>', t)
        # Inline code
        t = re.sub(r'`(.+?)`', r'<code>\1</code>', t)
        # Case citations highlight
        t = re.sub(r'\b([A-Z][a-zA-Z\s\.,]+?\s+v(?:s)?\.?\s+[A-Z][a-zA-Z\s\.,]+?(?:\b|\d+ SCRA \d+|G\.R\. No\. [\d-]+))', r'<span class="legal-citation">\1</span>', t)
        return t

    sec_counter = 1
    for line in lines:
        raw = line
        stripped = line.strip()

        # Code fences
        if stripped.startswith('```'):
            if in_code_block:
                in_code_block = False
                joined = html.escape("\n".join(code_lines))
                current_sec["units"].append(
                    f'<pre class="reader-code read-unit" data-unit-type="code"><code class="lang-{code_lang}">{joined}</code></pre>'
                )
                code_lines = []
                code_lang = ""
            else:
                if in_table:
                    flush_table()
                in_code_block = True
                code_lang = stripped[3:].strip()
                code_lines = []
            continue

        if in_code_block:
            code_lines.append(raw)
            continue

        # Tables
        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped.strip('|').split('|')]
            table_rows.append(cells)
            in_table = True
            continue
        elif in_table:
            flush_table()

        if not stripped:
            continue

        # Headings
        heading_match = re.match(r'^(#{1,6})\s+(.+)$', stripped)
        if heading_match:
            level = len(heading_match.group(1))
            htext = heading_match.group(2).strip()
            
            # Save previous section if it has content
            if current_sec["units"] or current_sec["title"] != (doc_title or "Introduction"):
                sections.append(current_sec)
                
            sec_id = f"sec-{sec_counter}-{slugify(htext)}"
            sec_counter += 1
            current_sec = {
                "id": sec_id,
                "title": clean_text(htext),
                "level": level,
                "units": []
            }
            continue

        # Horizontal Rule
        if re.match(r'^(\-{3,}|\*{3,}|_{3,})$', stripped):
            current_sec["units"].append('<hr class="section-divider" />')
            continue

        # Blockquote
        if stripped.startswith('>'):
            bq_content = stripped.lstrip('>').strip()
            formatted = format_inline(bq_content)
            current_sec["units"].append(
                f'<blockquote class="reader-quote read-unit" data-unit-type="quote">{formatted}</blockquote>'
            )
            continue

        # Bullet or list item
        bullet_match = re.match(r'^[\*\-\+]\s+(.+)$', stripped)
        num_match = re.match(r'^(\d+\.)\s+(.+)$', stripped)
        if bullet_match:
            formatted = format_inline(bullet_match.group(1))
            current_sec["units"].append(
                f'<div class="bullet-point read-unit" data-unit-type="bullet"><span class="bullet-dot">•</span><div class="bullet-content">{formatted}</div></div>'
            )
            continue
        elif num_match:
            num_prefix = num_match.group(1)
            formatted = format_inline(num_match.group(2))
            current_sec["units"].append(
                f'<div class="bullet-point read-unit" data-unit-type="numbered"><span class="bullet-num">{num_prefix}</span><div class="bullet-content">{formatted}</div></div>'
            )
            continue

        # Check for Case Digest Callouts (FACTS / ISSUE / HELD / DOCTRINE / RULING / PRINCIPLE)
        if re.match(r'^(FACTS|ISSUE|HELD|DOCTRINE|RULING|PRINCIPLE|TOPIC|CANON|ARTICLE|RULE):?', stripped, re.IGNORECASE):
            parts = stripped.split(':', 1)
            tag = parts[0].upper().strip()
            rest = parts[1].strip() if len(parts) > 1 else ''
            badge_class = 'badge-doctrine' if 'DOCTRINE' in tag else 'badge-held' if 'HELD' in tag or 'RULING' in tag else 'badge-issue' if 'ISSUE' in tag else 'badge-facts'
            formatted_rest = format_inline(rest)
            current_sec["units"].append(
                f'<div class="case-element read-unit" data-unit-type="case-badge" data-badge-label="{tag}"><span class="case-badge {badge_class}">{tag}</span> <div class="case-element-text">{formatted_rest}</div></div>'
            )
            continue

        # Standard Paragraph
        formatted_p = format_inline(stripped)
        current_sec["units"].append(f'<p class="read-unit" data-unit-type="paragraph">{formatted_p}</p>')

    if in_table:
        flush_table()
    if current_sec["units"] or current_sec["title"]:
        sections.append(current_sec)

    return sections

def parse_docx_file(file_path, doc_title=""):
    """Parses Microsoft Word (.docx) file into structured sections and units."""
    if not HAS_DOCX:
        raise RuntimeError("python-docx is required to parse .docx files. Install with 'pip install python-docx'.")
    
    doc = docx.Document(file_path)
    sections = []
    current_sec = {
        "id": "sec-intro",
        "title": doc_title or "Document Overview",
        "level": 1,
        "units": []
    }
    sec_counter = 1

    for p in doc.paragraphs:
        raw_text = clean_text(p.text.strip())
        if not raw_text:
            continue

        style_name = p.style.name.lower() if p.style else ""
        
        # Detect headings based on style or typography
        is_heading = (
            "heading 1" in style_name or
            "heading 2" in style_name or
            "heading 3" in style_name or
            "title" in style_name or
            (len(p.runs) > 0 and p.runs[0].bold and len(raw_text) < 120 and (
                raw_text.isupper() or
                re.match(r'^(CANON|ARTICLE|CHAPTER|TOPIC|PART|SECTION|RULE|CASE|I|II|III|IV|V|VI|VII|VIII|IX|X)\b', raw_text, re.IGNORECASE)
            ))
        )

        if is_heading:
            if current_sec["units"] or current_sec["title"] != (doc_title or "Document Overview"):
                sections.append(current_sec)
            level = 1 if ("1" in style_name or "title" in style_name) else 2 if "2" in style_name else 3
            sec_id = f"sec-{sec_counter}-{slugify(raw_text)}"
            sec_counter += 1
            current_sec = {
                "id": sec_id,
                "title": raw_text,
                "level": level,
                "units": []
            }
            continue

        # Render formatted runs
        formatted_runs = []
        for run in p.runs:
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
        
        formatted_text = "".join(formatted_runs).strip()
        if not formatted_text:
            continue

        # Check list item style
        if "list" in style_name or "bullet" in style_name or raw_text.startswith(('•', '-', '*')):
            clean_item = re.sub(r'^[•\-\*]\s*', '', formatted_text)
            current_sec["units"].append(
                f'<div class="bullet-point read-unit" data-unit-type="bullet"><span class="bullet-dot">•</span><div class="bullet-content">{clean_item}</div></div>'
            )
        else:
            current_sec["units"].append(f'<p class="read-unit" data-unit-type="paragraph">{formatted_text}</p>')

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

def parse_txt_file(file_path, doc_title=""):
    """Parses plain text file with intelligent header & paragraph heuristics."""
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    
    lines = [clean_text(l.strip()) for l in content.splitlines()]
    sections = []
    current_sec = {
        "id": "sec-intro",
        "title": doc_title or "Document Text",
        "level": 1,
        "units": []
    }
    sec_counter = 1

    case_title_re = re.compile(r'^(\d+\.)?\s*([A-Z][a-zA-Z\s\.,]+?\s+v(?:s)?\.?\s+[A-Z][a-zA-Z\s\.,]+.*)$')
    major_heading_re = re.compile(r'^([I|V|X|L|C|D|M]+\.\s+[A-Z\s,]{3,}|CHAPTER\s+[\w\d]+|CANON\s+\d+|ARTICLE\s+\d+|PART\s+[\w\d]+|RULE\s+\d+)', re.IGNORECASE)

    for line in lines:
        if not line:
            continue

        is_heading = False
        level = 2

        if major_heading_re.match(line) and len(line) < 140:
            is_heading = True
            level = 2
        elif case_title_re.match(line) and len(line) < 140:
            is_heading = True
            level = 3
        elif line.isupper() and 4 < len(line) < 100:
            is_heading = True
            level = 2

        if is_heading:
            if current_sec["units"] or current_sec["title"] != (doc_title or "Document Text"):
                sections.append(current_sec)
            sec_id = f"sec-{sec_counter}-{slugify(line)}"
            sec_counter += 1
            current_sec = {
                "id": sec_id,
                "title": line,
                "level": level,
                "units": []
            }
            continue

        esc = html.escape(line)
        # Bullet detection
        if line.startswith(('•', '-', '*')) or re.match(r'^\d+\.\s+', line):
            current_sec["units"].append(
                f'<div class="bullet-point read-unit" data-unit-type="bullet"><span class="bullet-dot">•</span><div class="bullet-content">{esc}</div></div>'
            )
        else:
            current_sec["units"].append(f'<p class="read-unit" data-unit-type="paragraph">{esc}</p>')

    if current_sec["units"] or current_sec["title"]:
        sections.append(current_sec)

    return sections

def parse_pdf_file(file_path, doc_title=""):
    """Parses PDF document using PyMuPDF (fitz), redacting repeated running headers."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) is required to parse .pdf files. Install with 'pip install pymupdf'.")
    
    doc = fitz.open(file_path)
    pages_data = []

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        blocks = page.get_text("dict")["blocks"]
        page_lines = []
        for b in blocks:
            if "lines" in b:
                for l in b["lines"]:
                    line_text = "".join(s["text"] for s in l["spans"]).strip()
                    bbox = l["bbox"]
                    # Filter recurring page footers and headers
                    if bbox[1] > 720 and line_text.isdigit():
                        continue
                    if bbox[1] < 60 and ("DEAN" in line_text or "FACULTY OF CIVIL LAW" in line_text):
                        continue
                    if line_text:
                        is_bold = any("bold" in s.get("font", "").lower() or s.get("flags", 0) & 2 for s in l["spans"])
                        page_lines.append((line_text, is_bold))
        pages_data.append((page_idx + 1, page_lines))

    sections = []
    current_sec = {
        "id": "sec-intro",
        "title": doc_title or "Document Content",
        "level": 1,
        "units": []
    }
    sec_counter = 1

    case_re = re.compile(r'^(\d+\.)?\s*([A-Z][a-zA-Z\s\.,]+?\s+v(?:s)?\.?\s+[A-Z].*)')
    topic_re = re.compile(r'^([I|V|X|L|C|D|M]+\.\s+[A-Z\s,]{3,}|CHAPTER\s+[\w\d]+|CANON\s+\d+|ARTICLE\s+\d+)', re.IGNORECASE)

    for pno, plines in pages_data:
        for line, is_bold in plines:
            cleaned = clean_text(line)
            if not cleaned:
                continue

            is_heading = False
            level = 2
            if (case_re.match(cleaned) and len(cleaned) < 150) or (is_bold and len(cleaned) < 100 and (cleaned.isupper() or topic_re.match(cleaned))):
                is_heading = True
                level = 3 if case_re.match(cleaned) else 2

            if is_heading:
                if current_sec["units"] or current_sec["title"] != (doc_title or "Document Content"):
                    sections.append(current_sec)
                sec_id = f"sec-{sec_counter}-{slugify(cleaned)}"
                sec_counter += 1
                current_sec = {
                    "id": sec_id,
                    "title": cleaned,
                    "level": level,
                    "units": []
                }
                continue

            esc = html.escape(cleaned)
            if is_bold:
                esc = f"<strong>{esc}</strong>"
            current_sec["units"].append(f'<p class="read-unit" data-unit-type="paragraph">{esc}</p>')

    if current_sec["units"] or current_sec["title"]:
        sections.append(current_sec)

    return sections

# ==============================================================================
# 3. HTML TEMPLATE & NATURAL SPEECH SYNTHESIS ENGINE
# ==============================================================================

HTML_PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="description" content="{escaped_description}">
  <title>{escaped_title}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700;800&family=Inter:wght@300;400;500;600;700&family=Merriweather:ital,wght@0,300;0,400;0,700;1,300;1,400&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <style>
    :root {{
      --bg-primary: #0f172a;
      --bg-secondary: #1e293b;
      --bg-tertiary: #334155;
      --bg-card: rgba(30, 41, 59, 0.75);
      --bg-card-hover: rgba(51, 65, 85, 0.85);
      --border-color: rgba(148, 163, 184, 0.18);
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
      --highlight-reading: rgba(56, 189, 248, 0.28);
      --font-body: 'Merriweather', Georgia, serif;
      --font-ui: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      --font-heading: 'Cinzel', serif;
      --font-mono: 'JetBrains Mono', monospace;
      --sidebar-width: 320px;
      --header-height: 72px;
      --content-max-width: 880px;
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
      --bg-tertiary: #e2e8f0;
      --bg-card: rgba(255, 255, 255, 0.9);
      --bg-card-hover: rgba(241, 245, 249, 0.98);
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
      --highlight-reading: rgba(2, 132, 199, 0.18);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.08), 0 2px 4px -1px rgba(0, 0, 0, 0.04);
    }}

    [data-theme="sepia"] {{
      --bg-primary: #fbf0d9;
      --bg-secondary: #f4e4c1;
      --bg-tertiary: #ecd6a7;
      --bg-card: rgba(244, 228, 193, 0.9);
      --bg-card-hover: rgba(236, 214, 167, 0.98);
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
      --highlight-reading: rgba(200, 150, 50, 0.28);
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

    /* Top Sticky Navigation Bar */
    header.app-header {{
      position: sticky;
      top: 0;
      z-index: 100;
      height: var(--header-height);
      background-color: rgba(15, 23, 42, 0.88);
      backdrop-filter: var(--glass-blur);
      -webkit-backdrop-filter: var(--glass-blur);
      border-bottom: 1px solid var(--border-color);
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 1.5rem;
      gap: 1rem;
    }}

    [data-theme="light"] header.app-header {{
      background-color: rgba(255, 255, 255, 0.9);
    }}

    [data-theme="sepia"] header.app-header {{
      background-color: rgba(251, 240, 217, 0.92);
    }}

    .brand-section {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
      flex-shrink: 0;
    }}

    .brand-logo {{
      width: 38px;
      height: 38px;
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      color: #0f172a;
      font-weight: 800;
      font-family: var(--font-heading);
      font-size: 1.1rem;
      box-shadow: 0 2px 8px rgba(251, 191, 36, 0.35);
      text-decoration: none;
    }}

    .brand-titles {{
      display: flex;
      flex-direction: column;
    }}

    .brand-title {{
      font-family: var(--font-ui);
      font-size: 0.95rem;
      font-weight: 700;
      color: var(--text-primary);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 300px;
    }}

    .brand-subtitle {{
      font-family: var(--font-ui);
      font-size: 0.72rem;
      color: var(--accent-gold);
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }}

    /* Global TTS Controls Toolbar */
    .tts-toolbar {{
      display: flex;
      align-items: center;
      background: var(--bg-secondary);
      border: 1px solid var(--border-color);
      padding: 0.35rem 0.65rem;
      border-radius: 40px;
      gap: 0.45rem;
      box-shadow: var(--shadow-md);
    }}

    .tts-btn {{
      background: transparent;
      border: none;
      color: var(--text-primary);
      cursor: pointer;
      width: 34px;
      height: 34px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: all 0.2s ease;
      font-size: 0.95rem;
    }}

    .tts-btn:hover {{
      background-color: var(--bg-tertiary);
      color: var(--accent-gold);
    }}

    .tts-btn.primary {{
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      color: #0f172a;
      font-weight: 700;
      box-shadow: 0 2px 6px rgba(251, 191, 36, 0.4);
    }}

    .tts-btn.primary:hover {{
      transform: scale(1.05);
      filter: brightness(1.1);
    }}

    .tts-status-badge {{
      font-family: var(--font-ui);
      font-size: 0.75rem;
      font-weight: 600;
      padding: 0.2rem 0.55rem;
      border-radius: 20px;
      background-color: var(--bg-tertiary);
      color: var(--text-muted);
      min-width: 90px;
      text-align: center;
    }}

    .tts-status-badge.speaking {{
      background-color: rgba(56, 189, 248, 0.2);
      color: var(--accent-blue);
      border: 1px solid rgba(56, 189, 248, 0.4);
      animation: pulseGlow 1.5s infinite alternate;
    }}

    @keyframes pulseGlow {{
      from {{ box-shadow: 0 0 4px rgba(56, 189, 248, 0.2); }}
      to {{ box-shadow: 0 0 10px rgba(56, 189, 248, 0.5); }}
    }}

    .header-actions {{
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }}

    .control-select {{
      background-color: var(--bg-secondary);
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      padding: 0.4rem 0.65rem;
      border-radius: 8px;
      font-family: var(--font-ui);
      font-size: 0.8rem;
      font-weight: 500;
      cursor: pointer;
      outline: none;
    }}

    .control-select:focus {{
      border-color: var(--border-focus);
    }}

    .icon-btn {{
      background-color: var(--bg-secondary);
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      width: 36px;
      height: 36px;
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      font-size: 0.9rem;
      transition: all 0.2s ease;
    }}

    .icon-btn:hover {{
      background-color: var(--bg-tertiary);
      border-color: var(--border-focus);
    }}

    /* Main Application Layout */
    .app-layout {{
      display: flex;
      flex: 1;
      position: relative;
    }}

    /* Sticky Sidebar Navigation */
    aside.app-sidebar {{
      width: var(--sidebar-width);
      height: calc(100vh - var(--header-height));
      position: sticky;
      top: var(--header-height);
      background-color: var(--bg-secondary);
      border-right: 1px solid var(--border-color);
      display: flex;
      flex-direction: column;
      flex-shrink: 0;
      overflow-y: auto;
      z-index: 50;
      transition: transform 0.3s ease;
    }}

    .sidebar-header {{
      padding: 1.25rem 1rem 0.75rem 1rem;
      border-bottom: 1px solid var(--border-color);
    }}

    .sidebar-search {{
      position: relative;
      width: 100%;
    }}

    .sidebar-search input {{
      width: 100%;
      background-color: var(--bg-primary);
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 0.5rem 0.75rem 0.5rem 2rem;
      color: var(--text-primary);
      font-family: var(--font-ui);
      font-size: 0.85rem;
      outline: none;
      transition: border-color 0.2s ease;
    }}

    .sidebar-search input:focus {{
      border-color: var(--border-focus);
      box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.2);
    }}

    .sidebar-search .search-icon {{
      position: absolute;
      left: 0.65rem;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-muted);
      font-size: 0.8rem;
    }}

    .toc-container {{
      padding: 0.75rem 0.5rem;
      flex: 1;
    }}

    .toc-nav ul {{
      list-style: none;
    }}

    .toc-nav li {{
      margin-bottom: 0.2rem;
    }}

    .toc-link {{
      display: block;
      padding: 0.45rem 0.75rem;
      border-radius: 6px;
      color: var(--text-secondary);
      text-decoration: none;
      font-family: var(--font-ui);
      font-size: 0.82rem;
      font-weight: 500;
      line-height: 1.4;
      transition: all 0.2s ease;
      border-left: 2px solid transparent;
    }}

    .toc-link:hover {{
      background-color: var(--bg-tertiary);
      color: var(--text-primary);
    }}

    .toc-link.active {{
      background-color: rgba(56, 189, 248, 0.12);
      color: var(--accent-blue);
      border-left-color: var(--accent-blue);
      font-weight: 600;
    }}

    .toc-link.level-2 {{
      padding-left: 1.25rem;
      font-size: 0.78rem;
    }}

    .toc-link.level-3 {{
      padding-left: 1.75rem;
      font-size: 0.75rem;
      color: var(--text-muted);
    }}

    /* Main Content Container */
    main.app-content {{
      flex: 1;
      display: flex;
      justify-content: center;
      padding: 2.5rem 2rem;
      overflow-y: visible;
    }}

    .content-wrapper {{
      max-width: var(--content-max-width);
      width: 100%;
    }}

    /* Document Header Banner */
    .document-hero {{
      background: linear-gradient(145deg, var(--bg-card), var(--bg-secondary));
      border: 1px solid var(--border-color);
      border-radius: 16px;
      padding: 2.25rem 2rem;
      margin-bottom: 2.5rem;
      box-shadow: var(--shadow-lg);
      position: relative;
      overflow: hidden;
    }}

    .document-hero::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 4px;
      background: linear-gradient(90deg, var(--accent-gold), var(--accent-blue), var(--accent-indigo));
    }}

    .doc-badge {{
      display: inline-block;
      font-family: var(--font-ui);
      font-size: 0.75rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      padding: 0.25rem 0.75rem;
      border-radius: 20px;
      background: rgba(251, 191, 36, 0.15);
      color: var(--accent-gold);
      border: 1px solid rgba(251, 191, 36, 0.3);
      margin-bottom: 0.85rem;
    }}

    .doc-main-title {{
      font-family: var(--font-heading);
      font-size: 2rem;
      font-weight: 700;
      line-height: 1.3;
      color: var(--text-primary);
      margin-bottom: 0.75rem;
    }}

    .doc-subtitle {{
      font-family: var(--font-ui);
      font-size: 0.95rem;
      color: var(--text-secondary);
      line-height: 1.6;
    }}

    .doc-meta-bar {{
      display: flex;
      flex-wrap: wrap;
      gap: 1.25rem;
      margin-top: 1.25rem;
      padding-top: 1rem;
      border-top: 1px solid var(--border-color);
      font-family: var(--font-ui);
      font-size: 0.8rem;
      color: var(--text-muted);
    }}

    .meta-item {{
      display: flex;
      align-items: center;
      gap: 0.35rem;
    }}

    /* Section & Content Styling */
    section.doc-section {{
      background-color: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 2rem;
      margin-bottom: 2.25rem;
      box-shadow: var(--shadow-md);
      transition: border-color 0.2s ease;
    }}

    section.doc-section:hover {{
      border-color: rgba(148, 163, 184, 0.35);
    }}

    .section-header {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      border-bottom: 1px solid var(--border-color);
      padding-bottom: 0.85rem;
      margin-bottom: 1.5rem;
    }}

    .section-title {{
      font-family: var(--font-heading);
      font-size: 1.4rem;
      font-weight: 700;
      color: var(--accent-gold);
      line-height: 1.4;
    }}

    .section-title.level-2 {{
      font-size: 1.2rem;
      color: var(--text-primary);
    }}

    .section-title.level-3 {{
      font-size: 1.05rem;
      color: var(--accent-blue);
    }}

    .section-listen-btn {{
      font-family: var(--font-ui);
      font-size: 0.75rem;
      font-weight: 600;
      padding: 0.3rem 0.7rem;
      border-radius: 20px;
      background-color: var(--bg-tertiary);
      border: 1px solid var(--border-color);
      color: var(--text-secondary);
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.3rem;
      transition: all 0.2s ease;
    }}

    .section-listen-btn:hover {{
      background-color: var(--accent-gold);
      color: #0f172a;
      border-color: var(--accent-gold);
    }}

    /* Read Units & Paragraphs */
    .read-unit {{
      position: relative;
      margin-bottom: 1.15rem;
      padding: 0.4rem 0.7rem;
      border-radius: 6px;
      cursor: pointer;
      transition: background-color 0.2s ease, box-shadow 0.2s ease;
    }}

    .read-unit:hover {{
      background-color: rgba(148, 163, 184, 0.08);
    }}

    .read-unit.is-speaking {{
      background-color: var(--highlight-reading) !important;
      border-left: 3px solid var(--accent-blue);
      box-shadow: 0 0 12px rgba(56, 189, 248, 0.25);
    }}

    .bullet-point {{
      display: flex;
      align-items: flex-start;
      gap: 0.65rem;
    }}

    .bullet-dot {{
      color: var(--accent-gold);
      font-weight: bold;
      font-size: 1.1rem;
      line-height: 1.5;
    }}

    .bullet-num {{
      color: var(--accent-blue);
      font-family: var(--font-ui);
      font-weight: 700;
      font-size: 0.9rem;
      line-height: 1.6;
    }}

    .bullet-content {{
      flex: 1;
    }}

    .reader-quote {{
      border-left: 4px solid var(--accent-gold);
      padding: 0.75rem 1.25rem;
      background-color: rgba(251, 191, 36, 0.06);
      border-radius: 0 8px 8px 0;
      font-style: italic;
      color: var(--text-secondary);
      margin: 1.25rem 0;
    }}

    .reader-code {{
      font-family: var(--font-mono);
      font-size: 0.85rem;
      background-color: var(--bg-secondary);
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 1rem;
      overflow-x: auto;
      line-height: 1.5;
    }}

    .table-responsive {{
      overflow-x: auto;
      margin: 1.25rem 0;
    }}

    .reader-table {{
      width: 100%;
      border-collapse: collapse;
      font-family: var(--font-ui);
      font-size: 0.88rem;
    }}

    .reader-table th, .reader-table td {{
      padding: 0.75rem 1rem;
      border: 1px solid var(--border-color);
      text-align: left;
    }}

    .reader-table th {{
      background-color: var(--bg-secondary);
      color: var(--accent-gold);
      font-weight: 600;
    }}

    .reader-table tr:nth-child(even) {{
      background-color: rgba(148, 163, 184, 0.04);
    }}

    /* Case Elements Badges */
    .case-element {{
      display: flex;
      align-items: flex-start;
      gap: 0.75rem;
      margin-bottom: 0.85rem;
    }}

    .case-badge {{
      font-family: var(--font-ui);
      font-size: 0.7rem;
      font-weight: 700;
      text-transform: uppercase;
      padding: 0.2rem 0.55rem;
      border-radius: 4px;
      letter-spacing: 0.05em;
      flex-shrink: 0;
      margin-top: 0.25rem;
    }}

    .badge-facts {{ background: rgba(56, 189, 248, 0.18); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.35); }}
    .badge-issue {{ background: rgba(251, 191, 36, 0.18); color: var(--accent-gold); border: 1px solid rgba(251, 191, 36, 0.35); }}
    .badge-held {{ background: rgba(52, 211, 153, 0.18); color: var(--accent-emerald); border: 1px solid rgba(52, 211, 153, 0.35); }}
    .badge-doctrine {{ background: rgba(251, 113, 133, 0.18); color: var(--accent-rose); border: 1px solid rgba(251, 113, 133, 0.35); }}

    .legal-citation {{
      font-weight: 600;
      color: var(--accent-gold);
      border-bottom: 1px dotted rgba(251, 191, 36, 0.5);
    }}

    .section-divider {{
      border: 0;
      height: 1px;
      background: var(--border-color);
      margin: 1.5rem 0;
    }}

    /* Floating TTS Trigger on Selection */
    #floatingTtsTrigger {{
      position: absolute;
      display: none;
      z-index: 1000;
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      color: #0f172a;
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
      font-size: 0.8rem;
      color: var(--text-muted);
      margin-top: auto;
    }}

    /* Mobile Responsiveness */
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
        max-width: 180px;
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
      <a href="#" class="brand-logo" title="MLC Law Reader">⚖</a>
      <div class="brand-titles">
        <span class="brand-title">{escaped_title}</span>
        <span class="brand-subtitle">{escaped_subject_tag} • Interactive Reader</span>
      </div>
    </div>

    <!-- TTS Toolbar -->
    <div class="tts-toolbar">
      <button class="tts-btn" id="ttsPrevBtn" title="Previous Paragraph (P)">⏮</button>
      <button class="tts-btn primary" id="ttsPlayPauseBtn" title="Play / Pause (Space)">
        <span id="ttsPlayIcon">▶</span>
        <span id="ttsPauseIcon" style="display:none;">⏸</span>
      </button>
      <button class="tts-btn" id="ttsStopBtn" title="Stop (Esc)">⏹</button>
      <button class="tts-btn" id="ttsNextBtn" title="Next Paragraph (N)">⏭</button>
      <span class="tts-status-badge" id="ttsStatus">Ready</span>
    </div>

    <!-- Right Controls -->
    <div class="header-actions">
      <select class="control-select" id="ttsSpeedSelect" title="Reading Speed (Deliberate Study Pace)">
        <option value="0.75">0.75x (Slow)</option>
        <option value="0.85">0.85x (Very Clear)</option>
        <option value="0.9" selected>0.9x (Deliberate)</option>
        <option value="1.0">1.0x (Normal)</option>
        <option value="1.15">1.15x</option>
        <option value="1.25">1.25x</option>
        <option value="1.5">1.5x</option>
      </select>
      <select class="control-select" id="ttsVoiceSelect" title="Speech Voice (Neural Natural Voices First)" style="max-width: 155px;">
        <option value="default">Natural Voice (Auto)</option>
      </select>
      <button class="icon-btn" id="fontDecBtn" title="Decrease Font Size ( - )">A-</button>
      <button class="icon-btn" id="fontIncBtn" title="Increase Font Size ( + )">A+</button>
      <select class="control-select" id="themeSelect" title="Visual Theme">
        <option value="dark">🌙 Dark</option>
        <option value="sepia">📜 Sepia</option>
        <option value="light">☀️ Light</option>
      </select>
    </div>
  </header>

  <!-- Layout Container -->
  <div class="app-layout">
    <!-- Table of Contents Sidebar -->
    <aside class="app-sidebar" id="appSidebar">
      <div class="sidebar-header">
        <div class="sidebar-search">
          <span class="search-icon">🔍</span>
          <input type="text" id="sidebarSearch" placeholder="Search topics, cases... ( / )" autocomplete="off">
        </div>
      </div>
      <div class="toc-container">
        <nav class="toc-nav">
          <ul id="tocList">
            {toc_html}
          </ul>
        </nav>
      </div>
    </aside>

    <!-- Main Content Area -->
    <main class="app-content" id="mainContentArea">
      <div class="content-wrapper" id="contentWrapper">
        <!-- Hero Header -->
        <div class="document-hero">
          <span class="doc-badge">{escaped_subject_tag}</span>
          <h1 class="doc-main-title">{escaped_title}</h1>
          <p class="doc-subtitle">{escaped_description}</p>
          <div class="doc-meta-bar">
            <div class="meta-item"><span>📑</span> <span>{total_sections} Sections</span></div>
            <div class="meta-item"><span>🔊</span> <span id="totalUnitsCounter">{total_units} Read Units</span></div>
            <div class="meta-item"><span>⏱</span> <span>~{reading_time_minutes} min read</span></div>
            <div class="meta-item"><span>🎙</span> <span>Intelligent Natural Speech</span></div>
          </div>
        </div>

        <!-- Document Sections -->
        {sections_html}
      </div>
    </main>
  </div>

  <footer class="app-footer">
    <p>MLC Interactive Legal Study Hub • Natural Cadence Web Speech Synthesis Engine</p>
    <p style="margin-top: 0.25rem; opacity: 0.7;">Click any paragraph to start reading aloud • Fully compatible with Brave Read Aloud & extensions</p>
  </footer>

  <!-- Floating Selection Audio Trigger -->
  <button id="floatingTtsTrigger" aria-label="Read Selection Aloud">🔊 Read Selection</button>

  <!-- Natural Voice & Speech Synthesis Application Logic -->
  <script>
    (function() {{
      // 1. Theme Management
      const themeSelect = document.getElementById('themeSelect');
      const savedTheme = localStorage.getItem('mlc_reader_theme') || 'dark';
      document.documentElement.setAttribute('data-theme', savedTheme);
      themeSelect.value = savedTheme;

      themeSelect.addEventListener('change', (e) => {{
        document.documentElement.setAttribute('data-theme', e.target.value);
        localStorage.setItem('mlc_reader_theme', e.target.value);
      }});

      // 2. Font Size Controls
      let currentFontSize = parseInt(localStorage.getItem('mlc_font_size') || '17', 10);
      document.documentElement.style.fontSize = currentFontSize + 'px';

      document.getElementById('fontIncBtn').addEventListener('click', () => {{
        if (currentFontSize < 26) {{
          currentFontSize += 1;
          document.documentElement.style.fontSize = currentFontSize + 'px';
          localStorage.setItem('mlc_font_size', currentFontSize);
        }}
      }});

      document.getElementById('fontDecBtn').addEventListener('click', () => {{
        if (currentFontSize > 12) {{
          currentFontSize -= 1;
          document.documentElement.style.fontSize = currentFontSize + 'px';
          localStorage.setItem('mlc_font_size', currentFontSize);
        }}
      }});

      // 3. Sidebar Mobile Toggle & Search
      const sidebar = document.getElementById('appSidebar');
      const toggleSidebarBtn = document.getElementById('toggleSidebarBtn');
      toggleSidebarBtn.addEventListener('click', () => {{
        sidebar.classList.toggle('open');
      }});

      const searchInput = document.getElementById('sidebarSearch');
      const tocItems = document.querySelectorAll('.toc-nav li');
      searchInput.addEventListener('input', (e) => {{
        const q = e.target.value.toLowerCase().trim();
        tocItems.forEach(li => {{
          const txt = li.textContent.toLowerCase();
          li.style.display = txt.includes(q) ? '' : 'none';
        }});
      }});

      // 4. Reading Progress Bar
      window.addEventListener('scroll', () => {{
        const winScroll = document.body.scrollTop || document.documentElement.scrollTop;
        const height = document.documentElement.scrollHeight - document.documentElement.clientHeight;
        const scrolled = (winScroll / (height || 1)) * 100;
        document.getElementById('readingProgressBar').style.width = scrolled + '%';
      }});

      // =========================================================================
      // 5. ENHANCED NATURAL SPEECH SYNTHESIS ENGINE
      // =========================================================================
      const synth = window.speechSynthesis;
      let readUnits = [];
      let currentUnitIndex = -1;
      let isPaused = false;
      let voices = [];
      let nextUtteranceTimeout = null;
      let lastSectionId = null;

      const playPauseBtn = document.getElementById('ttsPlayPauseBtn');
      const stopBtn = document.getElementById('ttsStopBtn');
      const nextBtn = document.getElementById('ttsNextBtn');
      const prevBtn = document.getElementById('ttsPrevBtn');
      const playIcon = document.getElementById('ttsPlayIcon');
      const pauseIcon = document.getElementById('ttsPauseIcon');
      const statusBadge = document.getElementById('ttsStatus');
      const speedSelect = document.getElementById('ttsSpeedSelect');
      const voiceSelect = document.getElementById('ttsVoiceSelect');

      function updateReadUnits() {{
        readUnits = Array.from(document.querySelectorAll('.read-unit'));
      }}
      updateReadUnits();

      // Intelligent Voice Ranking (Prefers Natural / Neural / High-Tier Human Voices)
      function scoreVoice(v) {{
        const name = (v.name || '').toLowerCase();
        let score = 0;
        if (v.lang.startsWith('en')) score += 50;
        if (v.lang === 'en-US' || v.lang === 'en-GB') score += 20;
        if (name.includes('natural') || name.includes('neural') || name.includes('online')) score += 100;
        if (name.includes('jenny') || name.includes('guy') || name.includes('aria') || name.includes('ryan') || name.includes('sonia')) score += 80;
        if (name.includes('google') && !name.includes('legacy')) score += 70;
        if (name.includes('premium') || name.includes('enhanced') || name.includes('siri')) score += 60;
        if (name.includes('david') || name.includes('zira') || name.includes('mark')) score += 30;
        return score;
      }}

      function populateVoiceList() {{
        if (!synth) return;
        const allVoices = synth.getVoices();
        const enVoices = allVoices.filter(v => v.lang.startsWith('en'));
        enVoices.sort((a, b) => scoreVoice(b) - scoreVoice(a));
        voices = enVoices;

        voiceSelect.innerHTML = '';
        voices.forEach((v, idx) => {{
          const opt = document.createElement('option');
          opt.value = idx;
          const isNatural = (v.name.toLowerCase().includes('natural') || v.name.toLowerCase().includes('neural') || v.name.toLowerCase().includes('online'));
          opt.textContent = (isNatural ? '✨ ' : '') + v.name + ' (' + v.lang + ')';
          if (idx === 0) opt.selected = true;
          voiceSelect.appendChild(opt);
        }});
      }}

      if (synth) {{
        populateVoiceList();
        if (speechSynthesis.onvoiceschanged !== undefined) {{
          speechSynthesis.onvoiceschanged = populateVoiceList;
        }}
      }}

      // Smart Legal Phonetic Pre-Processor (Expands abbreviations & introduces natural prosody)
      function prepareSpeechText(unit) {{
        let text = (unit.innerText || unit.textContent || '').trim();
        const parentSec = unit.closest('section.doc-section');
        const unitType = unit.getAttribute('data-unit-type') || '';
        const badgeLabel = unit.getAttribute('data-badge-label') || '';

        // Label announcing for case components
        if (badgeLabel) {{
          if (badgeLabel.includes('DOCTRINE') || badgeLabel.includes('PRINCIPLE')) {{
            text = 'Key Doctrine: ' + text;
          }} else if (badgeLabel.includes('FACTS')) {{
            text = 'Case Facts: ' + text.replace(/^FACTS:?\\\\s*/i, '');
          }} else if (badgeLabel.includes('ISSUE')) {{
            text = 'Legal Issue: ' + text.replace(/^ISSUE:?\\\\s*/i, '');
          }} else if (badgeLabel.includes('HELD') || badgeLabel.includes('RULING')) {{
            text = 'Court Ruling and Held: ' + text.replace(/^(HELD|RULING):?\\\\s*/i, '');
          }}
        }} else if (unitType === 'quote') {{
          text = 'Quote: ' + text;
        }} else if (unitType === 'numbered') {{
          text = text.replace(/^(\\d+)\\.\\\\s*/, 'Item $1: ');
        }}

        // Clean & expand common legal abbreviations for natural pronunciation
        text = text
          .replace(/\\bG\\.R\\.\\s*No\\.?\\s*([\\d\\-]+)/gi, 'G R Number $1')
          .replace(/\\bSCRA\\b/g, 'S-C-R-A')
          .replace(/\\bPhil\\.\\s*(\\d+)/gi, 'Philippine Reports volume $1')
          .replace(/\\bArt\\.\\s*(\\d+)/gi, 'Article $1')
          .replace(/\\bArts\\.\\s*([\\d,\\s\\-]+)/gi, 'Articles $1')
          .replace(/\\bSec\\.\\s*(\\d+)/gi, 'Section $1')
          .replace(/\\bSecs\\.\\s*([\\d,\\s\\-]+)/gi, 'Sections $1')
          .replace(/\\bPar\\.\\s*(\\d+)/gi, 'Paragraph $1')
          .replace(/\\bPars\\.\\s*([\\d,\\s\\-]+)/gi, 'Paragraphs $1')
          .replace(/\\bRPC\\b/g, 'Revised Penal Code')
          .replace(/\\bCPR\\b/g, 'Code of Professional Responsibility')
          .replace(/\\bCPRA\\b/g, 'Code of Professional Responsibility and Accountability')
          .replace(/\\bConst\\.\\b/gi, 'Constitution')
          .replace(/\\bP\\.D\\.\\s*No\\.?\\s*(\\d+)/gi, 'Presidential Decree Number $1')
          .replace(/\\bR\\.A\\.\\s*No\\.?\\s*(\\d+)/gi, 'Republic Act Number $1')
          .replace(/\\bA\\.C\\.\\s*No\\.?\\s*([A-Za-z0-9\\-]+)/gi, 'Administrative Case Number $1')
          .replace(/\\bA\\.M\\.\\s*No\\.?\\s*([A-Za-z0-9\\-]+)/gi, 'Administrative Matter Number $1')
          .replace(/\\bet\\s+al\\./gi, 'and others')
          .replace(/\\bi\\.e\\./gi, 'that is')
          .replace(/\\be\\.g\\./gi, 'for example')
          .replace(/\\bviz\\./gi, 'namely')
          .replace(/\\s+v(?:s)?\\.\\s+/gi, ' versus ')
          .replace(/—/g, ', — ')
          .replace(/–/g, ', ')
          .replace(/;\\s*/g, ';, ') // Introduce micro-pause at semicolons
          .replace(/:\\s*/g, ': ... ') // Introduce clear pause at colons
          .replace(/\\(([a-z\\d])\\)/gi, 'sub-item $1, ');

        return text.trim();
      }}

      // 5.1 Mobile Background Audio Carrier & WakeLock
      let wakeLock = null;
      let audioCtx = null;
      let silentGain = null;
      let oscillator = null;

      function startKeepAliveAudio() {{
        try {{
          if (!audioCtx) {{
            const AudioContextClass = window.AudioContext || window.webkitAudioContext;
            if (AudioContextClass) {{
              audioCtx = new AudioContextClass();
              oscillator = audioCtx.createOscillator();
              silentGain = audioCtx.createGain();
              silentGain.gain.setValueAtTime(0.0001, audioCtx.currentTime); // Inaudible keep-alive signal
              oscillator.frequency.setValueAtTime(440, audioCtx.currentTime);
              oscillator.connect(silentGain);
              silentGain.connect(audioCtx.destination);
              oscillator.start();
            }}
          }}
          if (audioCtx && audioCtx.state === 'suspended') {{
            audioCtx.resume();
          }}
        }} catch (e) {{}}
      }}

      function stopKeepAliveAudio() {{
        try {{
          if (audioCtx && audioCtx.state === 'running') {{
            audioCtx.suspend();
          }}
        }} catch (e) {{}}
      }}

      async function requestWakeLock() {{
        try {{
          if ('wakeLock' in navigator) {{
            wakeLock = await navigator.wakeLock.request('screen');
          }}
        }} catch (err) {{}}
      }}

      function releaseWakeLock() {{
        if (wakeLock) {{
          wakeLock.release().catch(() => {{}});
          wakeLock = null;
        }}
      }}

      function updateMediaSession(trackTitle) {{
        if ('mediaSession' in navigator) {{
          navigator.mediaSession.metadata = new MediaMetadata({{
            title: trackTitle || document.title,
            artist: 'MLC Law Library',
            album: 'Legal Audio Reader'
          }});
          navigator.mediaSession.setActionHandler('play', () => playPauseBtn.click());
          navigator.mediaSession.setActionHandler('pause', () => playPauseBtn.click());
          navigator.mediaSession.setActionHandler('nexttrack', () => nextBtn.click());
          navigator.mediaSession.setActionHandler('previoustrack', () => prevBtn.click());
        }}
      }}

      function setSpeakingState(speaking, paused = false) {{
        if (speaking && !paused) {{
          playIcon.style.display = 'none';
          pauseIcon.style.display = 'inline';
          statusBadge.textContent = 'Speaking...';
          statusBadge.className = 'tts-status-badge speaking';
          startKeepAliveAudio();
          requestWakeLock();
        }} else if (paused) {{
          playIcon.style.display = 'inline';
          pauseIcon.style.display = 'none';
          statusBadge.textContent = 'Paused';
          statusBadge.className = 'tts-status-badge';
          stopKeepAliveAudio();
          releaseWakeLock();
        }} else {{
          playIcon.style.display = 'inline';
          pauseIcon.style.display = 'none';
          statusBadge.textContent = 'Ready';
          statusBadge.className = 'tts-status-badge';
          stopKeepAliveAudio();
          releaseWakeLock();
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

      // =========================================================================
      // Native Multi-Utterance Queueing Engine (Persists through Android Minimization)
      // =========================================================================
      const QUEUE_AHEAD_COUNT = 8;
      let isPlaying = false;
      let nextQueueIdx = 0;

      function createUtteranceForIndex(idx) {{
        if (idx < 0 || idx >= readUnits.length) return null;
        const unit = readUnits[idx];
        const currentSection = unit.closest('section.doc-section');
        const currentSecId = currentSection ? currentSection.id : null;
        
        let processedText = prepareSpeechText(unit);
        if (currentSecId && currentSecId !== lastSectionId && currentSection) {{
          const headerEl = currentSection.querySelector('.section-title');
          if (headerEl) {{
            processedText = 'New Topic: ' + headerEl.innerText.trim() + ' ... ... ' + processedText;
          }}
          lastSectionId = currentSecId;
        }}

        if (!processedText) return null;

        const utterance = new SpeechSynthesisUtterance(processedText);
        utterance.rate = parseFloat(speedSelect.value) || 0.9;
        utterance.pitch = 1.0;

        const selectedVoiceIdx = voiceSelect.value;
        if (selectedVoiceIdx !== 'default' && voices[selectedVoiceIdx]) {{
          utterance.voice = voices[selectedVoiceIdx];
        }} else if (voices.length > 0) {{
          utterance.voice = voices[0];
        }}

        utterance.onstart = () => {{
          currentUnitIndex = idx;
          highlightUnit(idx);
          updateMediaSession(processedText);
          setSpeakingState(true, false);
        }};

        utterance.onend = () => {{
          if (isPlaying && !isPaused) {{
            fillSpeechQueue();
            if (idx >= readUnits.length - 1) {{
              setSpeakingState(false);
              isPlaying = false;
            }}
          }}
        }};

        utterance.onerror = (e) => {{
          if (e.error !== 'interrupted' && e.error !== 'canceled') {{
            console.warn('TTS error:', e);
          }}
        }};

        return utterance;
      }}

      function fillSpeechQueue() {{
        if (!isPlaying || isPaused || !synth) return;
        while (nextQueueIdx < readUnits.length && nextQueueIdx < currentUnitIndex + QUEUE_AHEAD_COUNT) {{
          const u = createUtteranceForIndex(nextQueueIdx);
          nextQueueIdx++;
          if (u) {{
            synth.speak(u);
          }}
        }}
      }}

      function startPlaybackFrom(idx) {{
        if (!synth) return;
        synth.cancel();
        isPlaying = true;
        isPaused = false;
        currentUnitIndex = idx;
        lastSectionId = null;
        nextQueueIdx = idx;
        setSpeakingState(true, false);
        fillSpeechQueue();
      }}

      playPauseBtn.addEventListener('click', () => {{
        if (!synth) return;
        if (isPlaying && !isPaused) {{
          synth.pause();
          isPaused = true;
          setSpeakingState(true, true);
        }} else if (isPaused) {{
          synth.resume();
          isPaused = false;
          setSpeakingState(true, false);
          fillSpeechQueue();
        }} else {{
          startPlaybackFrom(currentUnitIndex >= 0 ? currentUnitIndex : 0);
        }}
      }});

      stopBtn.addEventListener('click', () => {{
        if (!synth) return;
        synth.cancel();
        isPlaying = false;
        isPaused = false;
        setSpeakingState(false);
        currentUnitIndex = -1;
        lastSectionId = null;
      }});

      nextBtn.addEventListener('click', () => {{
        if (!synth) return;
        const target = Math.min(readUnits.length - 1, (currentUnitIndex >= 0 ? currentUnitIndex : 0) + 1);
        startPlaybackFrom(target);
      }});

      prevBtn.addEventListener('click', () => {{
        if (!synth) return;
        const target = Math.max(0, (currentUnitIndex >= 0 ? currentUnitIndex : 0) - 1);
        startPlaybackFrom(target);
      }});

      speedSelect.addEventListener('change', () => {{
        if (isPlaying) {{
          startPlaybackFrom(currentUnitIndex >= 0 ? currentUnitIndex : 0);
        }}
      }});

      voiceSelect.addEventListener('change', () => {{
        if (isPlaying) {{
          startPlaybackFrom(currentUnitIndex >= 0 ? currentUnitIndex : 0);
        }}
      }});

      // Click on any paragraph to read from that exact spot with natural cadence
      document.getElementById('contentWrapper').addEventListener('click', (e) => {{
        const unit = e.target.closest('.read-unit');
        if (unit) {{
          const idx = readUnits.indexOf(unit);
          if (idx !== -1) {{
            startPlaybackFrom(idx);
          }}
        }}
      }});

      // Section Quick-Listen Handler
      window.readSection = function(sectionId) {{
        const el = document.getElementById(sectionId);
        if (el) {{
          const firstUnit = el.querySelector('.read-unit');
          if (firstUnit) {{
            const idx = readUnits.indexOf(firstUnit);
            if (idx !== -1) {{
              startPlaybackFrom(idx);
            }}
          }}
        }}
      }};
      window.readSection = function(sectionId) {{
        const el = document.getElementById(sectionId);
        if (el) {{
          const firstUnit = el.querySelector('.read-unit');
          if (firstUnit) {{
            const idx = readUnits.indexOf(firstUnit);
            if (idx !== -1) {{
              isPaused = false;
              lastSectionId = null; // force announcement of section
              speakUnit(idx);
            }}
          }}
        }}
      }};

      // Floating Selection Audio Button (with natural legal abbreviation expander)
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
          const dummyUnit = document.createElement('div');
          dummyUnit.innerText = sel;
          const naturalText = prepareSpeechText(dummyUnit);

          const utterance = new SpeechSynthesisUtterance(naturalText);
          utterance.rate = parseFloat(speedSelect.value) || 0.9;
          utterance.pitch = 1.0;

          const selectedVoiceIdx = voiceSelect.value;
          if (selectedVoiceIdx !== 'default' && voices[selectedVoiceIdx]) {{
            utterance.voice = voices[selectedVoiceIdx];
          }} else if (voices.length > 0) {{
            utterance.voice = voices[0];
          }}

          synth.speak(utterance);
          setSpeakingState(true);
          floatBtn.style.display = 'none';
        }}
      }});

      // Keyboard Navigation
      document.addEventListener('keydown', (e) => {{
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
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
          document.getElementById('fontIncBtn').click();
        }} else if (e.key === '-') {{
          document.getElementById('fontDecBtn').click();
        }} else if (e.key === '/') {{
          e.preventDefault();
          searchInput.focus();
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
    
    # Infer subject from parent folder names
    parts = list(path_obj.parts)
    subject = "Law Study Notes"
    for p in reversed(parts[:-1]):
        if any(term in p.lower() for term in ["criminal", "csl", "statutory", "blje", "ethics", "procedure", "civil", "consti"]):
            subject = p
            break
            
    # Format readable title
    clean_title = filename_stem.replace("_", " ").replace("-", " ")
    clean_title = re.sub(r'\s+', ' ', clean_title).strip()
    return subject, clean_title

def convert_file_to_html_reader(file_path, output_path=None, overwrite=True):
    path_obj = Path(file_path).resolve()
    if not path_obj.exists():
        print(f"[ERROR] File not found: {file_path}")
        return None

    ext = path_obj.suffix.lower()
    subject_tag, doc_title = infer_subject_and_title(path_obj)
    
    if output_path is None:
        output_path = path_obj.with_suffix('.html')
    else:
        output_path = Path(output_path).resolve()

    if output_path.exists() and not overwrite:
        print(f"[SKIP] Output exists: {output_path}")
        return output_path

    print(f"[PROCESSING] {path_obj.name} -> {output_path.name}")

    # 1. Parse content by extension
    sections = []
    if ext == '.md':
        with open(path_obj, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        sections = parse_markdown(content, doc_title)
    elif ext == '.docx':
        sections = parse_docx_file(str(path_obj), doc_title)
    elif ext == '.txt':
        sections = parse_txt_file(str(path_obj), doc_title)
    elif ext == '.pdf':
        sections = parse_pdf_file(str(path_obj), doc_title)
    else:
        print(f"[WARN] Unsupported extension: {ext}")
        return None

    if not sections:
        print(f"[WARN] No sections extracted from {path_obj.name}")
        return None

    # 2. Build TOC HTML & Sections HTML
    toc_items = []
    sec_html_list = []
    total_units = 0

    for s in sections:
        sec_id = s.get("id", "sec-main")
        sec_title = s.get("title", "Section")
        level = s.get("level", 1)
        units = s.get("units", [])
        total_units += len(units)

        toc_items.append(
            f'<li><a href="#{sec_id}" class="toc-link level-{level}">{html.escape(sec_title)}</a></li>'
        )

        rendered_units = "\n        ".join(units)
        sec_html_list.append(f'''
        <section class="doc-section" id="{sec_id}">
          <header class="section-header">
            <h{min(level + 1, 4)} class="section-title level-{level}">{html.escape(sec_title)}</h{min(level + 1, 4)}>
            <button class="section-listen-btn" onclick="readSection('{sec_id}')" title="Read this topic aloud with natural pacing">🔊 Listen Topic</button>
          </header>
          <div class="section-body">
            {rendered_units}
          </div>
        </section>''')

    toc_html = "\n            ".join(toc_items)
    sections_html = "\n".join(sec_html_list)

    # Calculate estimated reading time (~160 words/min for deliberate study)
    reading_time = max(1, round(total_units * 35 / 160))

    rendered_page = HTML_PAGE_TEMPLATE.format(
        escaped_title=html.escape(doc_title),
        escaped_description=html.escape(f"{subject_tag} comprehensive review notes and study materials."),
        escaped_subject_tag=html.escape(subject_tag),
        total_sections=len(sections),
        total_units=total_units,
        reading_time_minutes=reading_time,
        toc_html=toc_html,
        sections_html=sections_html
    )

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(rendered_page)

    print(f"[SUCCESS] Created HTML Reader: {output_path} ({total_units} read units, {len(sections)} sections)")
    return output_path

def scan_and_convert_directory(dir_path, recursive=True, overwrite=True):
    target_dir = Path(dir_path).resolve()
    print(f"\n=== Scanning directory: {target_dir} ===")
    supported_extensions = {'.md', '.docx', '.txt', '.pdf'}
    
    # Ignore patterns
    ignore_stems = {'readme', 'handoff', 'format_dump', 'study_hub', 'index'}
    
    converted_count = 0
    generated_files = []

    pattern = "**/*" if recursive else "*"
    for p in target_dir.glob(pattern):
        if not p.is_file():
            continue
        if p.suffix.lower() not in supported_extensions:
            continue
        
        # Skip temporary, cache, or existing html
        if '__pycache__' in str(p) or '.agents' in str(p) or p.name.startswith(('~', '.', 'Unconfirmed')):
            continue
        if p.stem.lower() in ignore_stems and p.suffix.lower() == '.md':
            continue

        try:
            out = convert_file_to_html_reader(p, overwrite=overwrite)
            if out:
                converted_count += 1
                generated_files.append(out)
        except Exception as e:
            print(f"[ERROR] Failed to convert {p.name}: {e}")

    print(f"\n=== Completed! Successfully converted {converted_count} documents. ===\n")
    return generated_files

def generate_study_hub_index(root_dir):
    """Generates a master index dashboard (MLC_Study_Hub.html) listing all generated readers."""
    root_path = Path(root_dir).resolve()
    html_files = list(root_path.glob("**/*.html"))
    
    # Exclude the hub itself
    html_files = [f for f in html_files if f.name.lower() not in {'mlc_study_hub.html', 'index.html'}]
    
    # Group by subject / parent folder
    by_folder = {}
    for h in sorted(html_files, key=lambda x: str(x)):
        rel = h.relative_to(root_path)
        subject_group = rel.parent.name or "Root"
        if subject_group not in by_folder:
            by_folder[subject_group] = []
        by_folder[subject_group].append((h, rel))

    hub_html = ['''<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>MLC Interactive Study Hub • Natural Speech Readers</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg-primary: #0f172a;
      --bg-secondary: #1e293b;
      --bg-tertiary: #334155;
      --bg-card: rgba(30, 41, 59, 0.7);
      --border-color: rgba(148, 163, 184, 0.15);
      --text-primary: #f8fafc;
      --text-secondary: #cbd5e1;
      --text-muted: #94a3b8;
      --accent-gold: #fbbf24;
      --accent-blue: #38bdf8;
      --font-ui: 'Inter', -apple-system, sans-serif;
      --font-heading: 'Cinzel', serif;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background-color: var(--bg-primary);
      color: var(--text-primary);
      font-family: var(--font-ui);
      padding: 3rem 1.5rem;
      min-height: 100vh;
    }
    .hub-container {
      max-width: 1100px;
      margin: 0 auto;
    }
    .hub-header {
      text-align: center;
      margin-bottom: 3.5rem;
    }
    .hub-badge {
      display: inline-block;
      font-size: 0.8rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      padding: 0.35rem 1rem;
      border-radius: 20px;
      background: rgba(251, 191, 36, 0.15);
      color: var(--accent-gold);
      border: 1px solid rgba(251, 191, 36, 0.3);
      margin-bottom: 1rem;
    }
    .hub-title {
      font-family: var(--font-heading);
      font-size: 2.5rem;
      font-weight: 800;
      color: var(--text-primary);
      margin-bottom: 0.75rem;
    }
    .hub-desc {
      color: var(--text-secondary);
      font-size: 1.05rem;
      max-width: 650px;
      margin: 0 auto;
      line-height: 1.6;
    }
    .search-box {
      margin: 2rem auto 0 auto;
      max-width: 480px;
      position: relative;
    }
    .search-box input {
      width: 100%;
      background: var(--bg-secondary);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 0.85rem 1.25rem 0.85rem 2.75rem;
      color: var(--text-primary);
      font-size: 0.95rem;
      outline: none;
      transition: border-color 0.2s;
    }
    .search-box input:focus {
      border-color: var(--accent-blue);
    }
    .search-icon {
      position: absolute;
      left: 1rem;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-muted);
    }
    .group-section {
      margin-bottom: 3rem;
    }
    .group-title {
      font-family: var(--font-heading);
      font-size: 1.4rem;
      color: var(--accent-gold);
      margin-bottom: 1.25rem;
      padding-bottom: 0.5rem;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }
    .cards-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
      gap: 1.25rem;
    }
    .reader-card {
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.5rem;
      text-decoration: none;
      color: inherit;
      display: flex;
      flex-direction: column;
      transition: all 0.25s ease;
      position: relative;
    }
    .reader-card:hover {
      transform: translateY(-4px);
      border-color: var(--accent-blue);
      background: rgba(51, 65, 85, 0.8);
      box-shadow: 0 10px 20px -5px rgba(0,0,0,0.4);
    }
    .card-top {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 0.75rem;
    }
    .card-icon {
      font-size: 1.4rem;
    }
    .card-tts-badge {
      font-size: 0.72rem;
      font-weight: 700;
      color: var(--accent-blue);
      background: rgba(56, 189, 248, 0.15);
      padding: 0.2rem 0.5rem;
      border-radius: 12px;
      border: 1px solid rgba(56, 189, 248, 0.3);
    }
    .card-title {
      font-size: 1.1rem;
      font-weight: 700;
      color: var(--text-primary);
      margin-bottom: 0.5rem;
      line-height: 1.4;
    }
    .card-relpath {
      font-size: 0.78rem;
      color: var(--text-muted);
      margin-top: auto;
      padding-top: 0.75rem;
      border-top: 1px solid var(--border-color);
      word-break: break-all;
    }
  </style>
</head>
<body>
  <div class="hub-container">
    <header class="hub-header">
      <span class="hub-badge">MLC Academic Law Library</span>
      <h1 class="hub-title">Interactive Study Hub & TTS Readers</h1>
      <p class="hub-desc">Access all your course modules, digests, and outlines with full built-in Natural Speech Synthesis and visual reading guides.</p>
      <div class="search-box">
        <span class="search-icon">🔍</span>
        <input type="text" id="hubSearch" placeholder="Search subject, digest, or topic...">
      </div>
    </header>
''']

    for group_name, items in sorted(by_folder.items()):
        hub_html.append(f'''
    <section class="group-section">
      <h2 class="group-title">📁 {html.escape(group_name)}</h2>
      <div class="cards-grid">''')
        for full_p, rel_p in items:
            stem = full_p.stem.replace('_', ' ')
            url_path = rel_p.as_posix()
            hub_html.append(f'''
        <a href="{url_path}" class="reader-card">
          <div class="card-top">
            <span class="card-icon">📖</span>
            <span class="card-tts-badge">🎙 Natural Voice</span>
          </div>
          <h3 class="card-title">{html.escape(stem)}</h3>
          <div class="card-relpath">{html.escape(url_path)}</div>
        </a>''')
        hub_html.append('''
      </div>
    </section>''')

    hub_html.append('''
  </div>
  <script>
    const searchInput = document.getElementById('hubSearch');
    const cards = document.querySelectorAll('.reader-card');
    searchInput.addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase().trim();
      cards.forEach(card => {
        const text = card.textContent.toLowerCase();
        card.style.display = text.includes(q) ? '' : 'none';
      });
    });
  </script>
</body>
</html>''')

    hub_content = "\n".join(hub_html)
    hub_output_path = root_path / 'MLC_Study_Hub.html'
    index_output_path = root_path / 'index.html'
    with open(hub_output_path, 'w', encoding='utf-8') as f:
        f.write(hub_content)
    with open(index_output_path, 'w', encoding='utf-8') as f:
        f.write(hub_content)
    print(f"[SUCCESS] Generated Master Study Hub Index: {hub_output_path} & {index_output_path}")
    return hub_output_path

# ==============================================================================
# 5. CLI ENTRYPOINT
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Generate Interactive Natural Speech HTML Web Readers for MLC Legal Notes."
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="Path to a single study document (.md, .docx, .txt, .pdf) to convert."
    )
    parser.add_argument(
        "--dir",
        dest="directory",
        help="Path to a directory to convert all supported documents within."
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Scan and convert ALL supported documents across all subfolders of MLC."
    )
    parser.add_argument(
        "--hub",
        action="store_true",
        help="Generate/refresh the MLC Master Study Hub HTML dashboard."
    )
    parser.add_argument(
        "--no-overwrite",
        action="store_true",
        help="Do not overwrite existing HTML reader files."
    )

    args = parser.parse_args()
    root_mlc = Path(__file__).resolve().parent

    overwrite = not args.no_overwrite

    if args.all:
        scan_and_convert_directory(root_mlc, recursive=True, overwrite=overwrite)
        generate_study_hub_index(root_mlc)
    elif args.directory:
        scan_and_convert_directory(args.directory, recursive=True, overwrite=overwrite)
        generate_study_hub_index(root_mlc)
    elif args.file:
        convert_file_to_html_reader(args.file, overwrite=overwrite)
        generate_study_hub_index(root_mlc)
    elif args.hub:
        generate_study_hub_index(root_mlc)
    else:
        print("No specific file or flag provided. Running full scan on MLC folder...")
        scan_and_convert_directory(root_mlc, recursive=True, overwrite=overwrite)
        generate_study_hub_index(root_mlc)

if __name__ == '__main__':
    main()
