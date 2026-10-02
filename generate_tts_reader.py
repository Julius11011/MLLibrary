#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MLC Universal Text-to-Speech (Read Aloud) HTML Reader Generator
===============================================================
Converts legal study materials (.docx) across all MLC folders
into interactive, high-readability HTML web applications featuring:
  - Smart Reading Normalization & Deduplication:
      * Word-Number pairs: 'one (1)' -> 'one', 'thirty (30)' -> 'thirty'
      * Bracketed duplicate tags: '[Applicable] Applicable' -> 'Applicable', '[Answer] Answer:' -> 'Answer:'
      * Elimination of all emojis/icons from spoken streams and visual UI for luxury typography aesthetic
  - Master Unified Table (Citations + Cyber & Digital Laws: SKRA, JEE-AR, SIP-ruh, DEE-PEE-AY, etc.)
  - Roman Numeral Cardinal & Ordinal Spoken Pronunciation
  - Topic Change Audio Notifications ("Now Reading: CASE X: ...")
  - Smooth Intonation Pauses at Colons, Semicolons, Commas, and Periods
  - Continuous, Flowing Body Paragraphs with Modern Legal Typography
  - Streamlined, Clean Table of Contents (TOC) focused on Main Topics & Sub Topics
  - Clear Visual Hierarchy: Case Badges, Citation Banners, Subheadings, ALAC Inline Badges & Tables
  - Native Studio MP3 Podcast Audio Player Integration with GitHub CDN Cloud Streaming
  - Full Keyboard Navigation, Theme Modes (Dark, Sepia, Light), and Responsive Mobile Layout
"""

import os
import sys
import re
import html
import json
import urllib.parse
import argparse
from pathlib import Path

# Optional dependencies
try:
    import docx
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# GitHub Release Audio CDN base URL
GITHUB_AUDIO_BASE_URL = "https://github.com/Julius11011/MLLibrary/releases/download/audio-v1"

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
    r'^(CASE\s+\d+[:\.]?.*|^\d+\.\s+[A-Z0-9\s\.,\(\)\'\-&]+?\s+V[\.S]?\s+.*)',
    re.IGNORECASE
)

MAIN_TOPIC_RE = re.compile(
    r'^(PART\s+[I|V|X\d]+[:\.]?.*|Canon\s+[I|V|X\d]+[:\.]?.*|CHAPTER\s+[I|V|X\d]+[:\.]?.*|\b[I|V|X]+\.\s+[A-Z\s\(\)&,\-\/:]{3,}|STEP\s+\d+[:\.]?.*|\d+\.\s+[A-Z\s\(\)&,\-\/:]{4,}|EXECUTIVE CASE DISTRIBUTION MATRIX)',
    re.IGNORECASE
)

SUB_TOPIC_RE = re.compile(
    r'^(Topic:\s+.*|^[A-Z]\.\s+[A-Za-z0-9\s\(\)&,\-\/:]{3,}|Section\s+\d+[:\.]?.*|Rule\s+\d+[:\.]?.*|STEP\s+\d+[:\.]?.*|STEP\s+0[:\.]?.*)',
    re.IGNORECASE
)

INTERNAL_SUBHEADING_RE = re.compile(
    r'^([I|V|X]+\.\s+(?:FACTS|ISSUE|THE COURT|COURT|APPLICABLE|SUBJECT MATTER|SYLLABUS|SUMMARY|HOW THE COURT|RELEVANT).*|FACTS OF THE CASE|STATEMENT OF THE ETHICAL ISSUE|STATEMENT OF THE ISSUE|SUBSTANTIVE LEGAL ISSUE|THE COURT\'S RULING|THE RULING|COURT\'S RULING|HOW THE COURT CONSTRUED|RELEVANT STATUTORY|APPLICABLE\s+.*(?:PRINCIPLES|MAXIMS|CANONS|PROVISIONS)|SUMMARY OF THE STATUTORY|SUBJECT MATTER & PROVISION|SYLLABUS TOPIC)',
    re.IGNORECASE
)

ALAC_PATTERNS = [
    r'Ethical Synthesis & Canon Application',
    r'Criminal Law Synthesis & Statutory Application',
    r'Statutory Construction Synthesis & Maxim Application',
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
# 2. DOCUMENT PARSER (.docx)
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
      - TOC limited strictly to Major Case Titles / Main Topics / Sub Topics
      - In-case styled subheadings (Facts, Issue, Ruling, Ethical Principles)
      - Smooth, non-fragmented continuous body paragraphs
      - Embedded tables at their exact contextual positions
      - Distinct styled ALAC inline badges (Answer, Legal Basis, Application, Conclusion)
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

    # Strictly identify Case Digest documents
    filename_lower = file_path.name.lower()
    is_case_digest_doc = "digest" in filename_lower

    for element in doc.element.body:
        # 1. Handle Tables
        if element.tag.endswith('tbl'):
            tbl = Table(element, doc)
            html_tbl = ['<div class="table-responsive read-unit" data-unit-type="table"><table class="reader-table">']
            for i, row in enumerate(tbl.rows):
                tag = 'th' if i == 0 else 'td'
                html_tbl.append('<tr>')
                for cell in row.cells:
                    ctext = html.escape(clean_text(cell.text.strip()))
                    html_tbl.append(f'<{tag}>{ctext}</{tag}>')
                html_tbl.append('</tr>')
            html_tbl.append('</table></div>')
            current_sec["units"].append("".join(html_tbl))
            continue

        # 2. Handle Paragraphs
        if not element.tag.endswith('p'):
            continue

        p = Paragraph(element, doc)
        raw_text = clean_text(p.text.strip())
        if not raw_text:
            continue

        formatted_text = format_inline_runs(p)
        style_name = (p.style.name or "").lower()

        # A. Check for Case Title Heading: "CASE 1: ...", "CASE 13: ..."
        case_match = CASE_HEADING_RE.match(raw_text)
        if case_match:
            if current_sec["units"] or current_sec["title"] != (doc_title or "Document Overview"):
                sections.append(current_sec)
            
            sec_id = f"sec-{sec_counter}-{slugify(raw_text)}"
            sec_counter += 1
            
            case_badge = ""
            case_name = raw_text
            c_split = re.match(r'^(CASE\s+\d+)[:\.]?\s*(.*)$', raw_text, re.IGNORECASE)
            if c_split:
                case_badge = c_split.group(1).upper()
                case_name = c_split.group(2) or case_badge

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

        # B. Check for Main Topics in Non-Digest Guides (e.g. Canons, StatCon Methodology, Outline Parts)
        if not is_case_digest_doc and MAIN_TOPIC_RE.match(raw_text):
            if current_sec["units"] or current_sec["title"] != (doc_title or "Document Overview"):
                sections.append(current_sec)
            
            sec_id = f"sec-{sec_counter}-{slugify(raw_text)}"
            sec_counter += 1
            
            current_sec = {
                "id": sec_id,
                "title": raw_text,
                "level": 1,
                "units": []
            }
            current_sec["units"].append(
                f'<div class="topic-header-card read-unit" data-unit-type="topic-header"><h2 class="topic-header-title">{formatted_text}</h2></div>'
            )
            continue

        # C. Check for Sub-Topics in Non-Digest Guides (e.g. "Topic: Theories of Criminal Law", "A. Definitions & Nature")
        if not is_case_digest_doc and SUB_TOPIC_RE.match(raw_text):
            if current_sec["units"] or current_sec["title"] != (doc_title or "Document Overview"):
                sections.append(current_sec)
            
            sec_id = f"sec-{sec_counter}-{slugify(raw_text)}"
            sec_counter += 1
            
            current_sec = {
                "id": sec_id,
                "title": raw_text,
                "level": 2,
                "units": []
            }
            current_sec["units"].append(
                f'<div class="subtopic-header-card read-unit" data-unit-type="subtopic-header"><h3 class="subtopic-header-title">{formatted_text}</h3></div>'
            )
            continue

        # D. Check for Case Citation Line (e.g. A.C. No. 13521, June 27, 2023 | Per Curiam)
        if expect_citation and (len(raw_text) < 180 and any(k in raw_text for k in ["SCRA", "G.R.", "A.C.", "A.M.", "Phil.", "Ponente", "Curiam", "En Banc", "|"])):
            current_sec["units"].append(
                f'<div class="case-citation-banner read-unit" data-unit-type="citation"><span class="citation-label">CITATION</span> <span class="citation-text">{formatted_text}</span></div>'
            )
            expect_citation = False
            continue
        
        expect_citation = False

        # E. Check for Internal In-Case Subheadings (e.g. I. FACTS OF THE CASE, II. ISSUE, III. RULING)
        if INTERNAL_SUBHEADING_RE.match(raw_text):
            current_sec["units"].append(
                f'<h3 class="case-subheading read-unit" data-unit-type="subheading"><span class="subheading-accent">§</span> {formatted_text}</h3>'
            )
            continue

        # F. Check for ALAC & Structured Reasoning Badges (ANSWER:, LEGAL BASIS:, ANALYSIS:, CONCLUSION:)
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

        # G. Check for Bullet Points and Lists
        if "list" in style_name or "bullet" in style_name or raw_text.startswith(('•', '-', '*')) or re.match(r'^[•\-\*]\s*', raw_text):
            clean_bullet = re.sub(r'^[•\-\*]\s*', '', formatted_text)
            current_sec["units"].append(
                f'<div class="bullet-point read-unit" data-unit-type="bullet"><span class="bullet-dot">•</span><div class="bullet-content">{clean_bullet}</div></div>'
            )
            continue

        # H. Standard Continuous Flowing Paragraph
        current_sec["units"].append(f'<p class="read-unit case-paragraph" data-unit-type="paragraph">{formatted_text}</p>')

    if current_sec["units"] or current_sec["title"]:
        sections.append(current_sec)

    return sections

# ==============================================================================
# 3. HTML GENERATOR & TEMPLATE
# ==============================================================================

READER_HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>__ESCAPED_TITLE__ | MLC Law Library</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700;800;900&family=Inter:wght@400;500;600;700&family=Merriweather:ital,wght@0,300;0,400;0,700;1,300;1,400&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg-primary: #0b1120;
      --bg-secondary: #131d31;
      --bg-tertiary: #1e293b;
      --border-color: rgba(148, 163, 184, 0.14);
      --border-focus: #fbbf24;
      --text-primary: #f8fafc;
      --text-secondary: #cbd5e1;
      --text-muted: #94a3b8;
      --accent-gold: #fbbf24;
      --accent-gold-dark: #d97706;
      --accent-blue: #38bdf8;
      --accent-emerald: #34d399;
      --accent-crimson: #f87171;
      --accent-purple: #c084fc;
      --highlight-bg: rgba(251, 191, 36, 0.22);
      --highlight-border: #fbbf24;
      --card-bg: rgba(19, 29, 49, 0.7);
      --font-ui: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      --font-body: 'Merriweather', Georgia, serif;
      --font-heading: 'Cinzel', serif;
      --font-mono: 'JetBrains Mono', monospace;
      --sidebar-width: 320px;
      --header-height: 60px;
      --toolbar-height: 52px;
      --sticky-top-total: calc(var(--header-height) + var(--toolbar-height));
      --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
      --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
    }

    [data-theme="sepia"] {
      --bg-primary: #fbf7ee;
      --bg-secondary: #f4ecd8;
      --bg-tertiary: #e9dfc4;
      --border-color: rgba(120, 97, 65, 0.18);
      --border-focus: #b45309;
      --text-primary: #2d241e;
      --text-secondary: #4a3e35;
      --text-muted: #786c60;
      --accent-gold: #b45309;
      --accent-gold-dark: #92400e;
      --accent-blue: #0284c7;
      --accent-emerald: #059669;
      --accent-crimson: #dc2626;
      --accent-purple: #7c3aed;
      --highlight-bg: rgba(217, 119, 6, 0.2);
      --highlight-border: #b45309;
      --card-bg: rgba(244, 236, 216, 0.85);
    }

    [data-theme="light"] {
      --bg-primary: #ffffff;
      --bg-secondary: #f8fafc;
      --bg-tertiary: #f1f5f9;
      --border-color: #e2e8f0;
      --border-focus: #d97706;
      --text-primary: #0f172a;
      --text-secondary: #334155;
      --text-muted: #64748b;
      --accent-gold: #d97706;
      --accent-gold-dark: #b45309;
      --accent-blue: #0284c7;
      --accent-emerald: #059669;
      --accent-crimson: #dc2626;
      --accent-purple: #7c3aed;
      --highlight-bg: rgba(254, 240, 138, 0.5);
      --highlight-border: #eab308;
      --card-bg: rgba(248, 250, 252, 0.9);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      background-color: var(--bg-primary);
      color: var(--text-primary);
      font-family: var(--font-body);
      font-size: 17px;
      line-height: 1.8;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      transition: background-color 0.2s ease, color 0.2s ease;
    }

    /* Scroll Progress Bar */
    #readingProgressBar {
      position: fixed;
      top: 0;
      left: 0;
      height: 3px;
      background: linear-gradient(90deg, var(--accent-gold), var(--accent-blue));
      width: 0%;
      z-index: 1000;
      transition: width 0.1s ease;
    }

    /* Floating Selection Reader Trigger */
    #floatingTtsTrigger {
      position: absolute;
      display: none;
      z-index: 999;
      background: var(--accent-gold);
      color: #0b1120;
      font-family: var(--font-ui);
      font-size: 0.8rem;
      font-weight: 700;
      padding: 6px 14px;
      border-radius: 20px;
      box-shadow: var(--shadow-lg);
      cursor: pointer;
      border: none;
      transform: translate(-50%, -100%);
      transition: transform 0.15s ease, background-color 0.15s ease;
    }
    #floatingTtsTrigger:hover {
      background: var(--accent-gold-dark);
      color: #ffffff;
      transform: translate(-50%, -105%) scale(1.05);
    }

    /* Header & Navigation Bar - STICKY TOP */
    header {
      position: sticky;
      top: 0;
      z-index: 100;
      background-color: var(--bg-secondary);
      border-bottom: 1px solid var(--border-color);
      height: var(--header-height);
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 1.5rem;
      backdrop-filter: blur(12px);
    }

    .header-left {
      display: flex;
      align-items: center;
      gap: 1rem;
    }

    .btn-icon {
      background: transparent;
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      width: 38px;
      height: 38px;
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      font-size: 1.1rem;
      transition: all 0.15s ease;
    }
    .btn-icon:hover {
      background-color: var(--bg-tertiary);
      border-color: var(--border-focus);
      color: var(--accent-gold);
    }

    .brand-title {
      font-family: var(--font-heading);
      font-size: 1.1rem;
      font-weight: 700;
      letter-spacing: 0.05em;
      color: var(--accent-gold);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 340px;
    }

    .subject-pill {
      display: inline-block;
      font-family: var(--font-ui);
      font-size: 0.72rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      padding: 2px 8px;
      border-radius: 4px;
      background: rgba(56, 189, 248, 0.12);
      color: var(--accent-blue);
      border: 1px solid rgba(56, 189, 248, 0.25);
    }

    .header-actions {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .select-control {
      background-color: var(--bg-tertiary);
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      font-family: var(--font-ui);
      font-size: 0.85rem;
      padding: 6px 10px;
      border-radius: 6px;
      outline: none;
      cursor: pointer;
    }
    .select-control:focus {
      border-color: var(--border-focus);
    }

    /* TTS Audio Control Bar - STICKY UNDER HEADER */
    .tts-toolbar {
      position: sticky;
      top: var(--header-height);
      z-index: 95;
      background-color: var(--bg-secondary);
      border-bottom: 1px solid var(--border-color);
      padding: 0.55rem 1.5rem;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 1rem;
      font-family: var(--font-ui);
      font-size: 0.88rem;
      backdrop-filter: blur(12px);
      box-shadow: 0 4px 14px rgba(0, 0, 0, 0.18);
      transition: background-color 0.2s ease, border-color 0.2s ease;
    }

    .tts-controls-group {
      display: flex;
      align-items: center;
      gap: 0.5rem;
      flex-wrap: wrap;
    }

    .btn-tts {
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      border: none;
      color: #0b1120;
      font-weight: 700;
      font-size: 0.85rem;
      padding: 7px 16px;
      border-radius: 6px;
      display: flex;
      align-items: center;
      gap: 0.4rem;
      cursor: pointer;
      transition: transform 0.1s ease, filter 0.15s ease;
    }
    .btn-tts:hover {
      filter: brightness(1.1);
      transform: translateY(-1px);
    }
    .btn-tts:active {
      transform: translateY(1px);
    }

    .btn-tts-secondary {
      background-color: var(--bg-tertiary);
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      font-weight: 600;
      font-size: 0.85rem;
      padding: 7px 12px;
      border-radius: 6px;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.3rem;
      transition: all 0.15s ease;
    }
    .btn-tts-secondary:hover {
      border-color: var(--border-focus);
      color: var(--accent-gold);
    }

    .btn-tts-stop {
      border-color: rgba(248, 113, 113, 0.35);
      color: var(--accent-crimson);
    }
    .btn-tts-stop:hover {
      background-color: rgba(248, 113, 113, 0.15);
      border-color: var(--accent-crimson);
      color: #ffffff;
    }

    .tts-status-badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 0.8rem;
      color: var(--text-muted);
      background: var(--bg-tertiary);
      padding: 4px 10px;
      border-radius: 12px;
      border: 1px solid var(--border-color);
      transition: all 0.2s ease;
    }
    .tts-status-badge.speaking {
      color: var(--accent-emerald);
      border-color: var(--accent-emerald);
      background: rgba(52, 211, 153, 0.1);
    }
    .tts-status-badge.paused {
      color: var(--accent-gold);
      border-color: var(--accent-gold);
      background: rgba(251, 191, 36, 0.1);
    }
    .pulse-dot {
      width: 8px;
      height: 8px;
      background-color: currentColor;
      border-radius: 50%;
      display: inline-block;
    }
    .speaking .pulse-dot {
      animation: pulse 1.2s infinite;
    }
    @keyframes pulse {
      0% { transform: scale(0.9); opacity: 0.6; }
      50% { transform: scale(1.3); opacity: 1; }
      100% { transform: scale(0.9); opacity: 0.6; }
    }

    /* Main Content Layout */
    .app-layout {
      display: flex;
      flex: 1;
      position: relative;
      width: 100%;
    }

    /* Table of Contents Sidebar */
    .sidebar-toc {
      width: var(--sidebar-width);
      background-color: var(--bg-secondary);
      border-right: 1px solid var(--border-color);
      position: sticky;
      top: var(--sticky-top-total);
      height: calc(100vh - var(--sticky-top-total));
      overflow-y: auto;
      padding: 1.25rem 1rem;
      flex-shrink: 0;
      transition: width 0.3s cubic-bezier(0.4, 0, 0.2, 1),
                  transform 0.3s cubic-bezier(0.4, 0, 0.2, 1),
                  padding 0.3s cubic-bezier(0.4, 0, 0.2, 1),
                  opacity 0.2s ease;
      z-index: 80;
    }

    /* Sidebar Minimized / Collapsed state */
    body.sidebar-collapsed .sidebar-toc,
    .sidebar-toc.minimized {
      width: 0 !important;
      min-width: 0 !important;
      padding: 0 !important;
      margin: 0 !important;
      border-right: none !important;
      opacity: 0 !important;
      overflow: hidden !important;
      pointer-events: none !important;
      transform: translateX(-15px);
    }

    .toc-header-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 0.85rem;
      padding-bottom: 0.4rem;
      border-bottom: 1px solid var(--border-color);
    }

    .toc-heading {
      font-family: var(--font-ui);
      font-size: 0.76rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      color: var(--accent-gold);
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .btn-toc-minimize {
      background: var(--bg-tertiary);
      border: 1px solid var(--border-color);
      color: var(--text-muted);
      cursor: pointer;
      width: 28px;
      height: 28px;
      border-radius: 6px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 1.05rem;
      line-height: 1;
      transition: all 0.15s ease;
    }
    .btn-toc-minimize:hover {
      background: var(--accent-gold);
      color: #0b1120;
      border-color: var(--accent-gold);
    }

    .toc-search-box {
      width: 100%;
      background-color: var(--bg-primary);
      border: 1px solid var(--border-color);
      color: var(--text-primary);
      font-family: var(--font-ui);
      font-size: 0.82rem;
      padding: 8px 12px;
      border-radius: 6px;
      margin-bottom: 1rem;
      outline: none;
    }
    .toc-search-box:focus {
      border-color: var(--border-focus);
    }

    .toc-list {
      list-style: none;
    }

    .toc-link {
      display: block;
      color: var(--text-secondary);
      text-decoration: none;
      font-family: var(--font-ui);
      font-size: 0.85rem;
      padding: 6px 10px 6px 12px;
      border-radius: 4px;
      margin-bottom: 2px;
      border-left: 2px solid transparent;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      transition: all 0.15s ease;
    }
    .toc-link:hover {
      background-color: var(--bg-tertiary);
      color: var(--accent-gold);
      border-left-color: var(--accent-gold);
    }
    .toc-link.active {
      background-color: var(--bg-tertiary);
      color: var(--accent-gold);
      border-left-color: var(--accent-gold);
      font-weight: 600;
    }
    .toc-link.level-2 {
      padding-left: 22px;
      font-size: 0.8rem;
      color: var(--text-muted);
    }

    /* Floating / Fixed Restore Sidebar Tab when Minimized */
    #restoreSidebarBtn {
      position: fixed;
      left: 1rem;
      top: calc(var(--sticky-top-total) + 14px);
      z-index: 90;
      background: var(--bg-secondary);
      border: 1px solid var(--accent-gold);
      color: var(--accent-gold);
      font-family: var(--font-ui);
      font-size: 0.82rem;
      font-weight: 700;
      padding: 7px 14px;
      border-radius: 20px;
      box-shadow: var(--shadow-lg);
      cursor: pointer;
      display: none;
      align-items: center;
      gap: 8px;
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    }
    body.sidebar-collapsed #restoreSidebarBtn {
      display: flex;
    }
    #restoreSidebarBtn:hover {
      background: var(--accent-gold);
      color: #0b1120;
      transform: translateY(-2px) scale(1.03);
      box-shadow: 0 8px 24px rgba(251, 191, 36, 0.35);
    }

    /* Header toggle button active status */
    body.sidebar-collapsed #toggleSidebarBtn {
      background-color: var(--bg-tertiary);
      border-color: var(--accent-gold);
      color: var(--accent-gold);
    }

    /* Main Reading Article Container - Adapts width smoothly */
    .reader-main {
      flex: 1;
      width: 100%;
      max-width: 920px;
      margin: 0 auto;
      padding: 2.5rem 2rem 6rem;
      transition: max-width 0.3s cubic-bezier(0.4, 0, 0.2, 1),
                  padding 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    /* Responsive Expanded Width when Sidebar is Minimized */
    body.sidebar-collapsed .reader-main {
      max-width: 1200px;
      padding-left: 3rem;
      padding-right: 3rem;
    }

    @media (min-width: 1600px) {
      .reader-main {
        max-width: 1020px;
      }
      body.sidebar-collapsed .reader-main {
        max-width: 1400px;
      }
    }

    .doc-meta-banner {
      margin-bottom: 2.5rem;
      padding-bottom: 1.5rem;
      border-bottom: 1px solid var(--border-color);
    }
    .doc-headline {
      font-family: var(--font-heading);
      font-size: 2.2rem;
      font-weight: 800;
      color: var(--accent-gold);
      line-height: 1.25;
      margin-bottom: 0.75rem;
      letter-spacing: 0.02em;
    }
    .doc-stats {
      font-family: var(--font-ui);
      font-size: 0.85rem;
      color: var(--text-muted);
      display: flex;
      align-items: center;
      gap: 1.25rem;
    }

    /* Studio Native MP3 Audio Player */
    .studio-audio-player {
      background: var(--bg-secondary);
      border: 1px solid var(--border-color);
      border-left: 4px solid var(--accent-emerald);
      border-radius: 10px;
      padding: 1.25rem 1.5rem;
      margin-bottom: 2.5rem;
      box-shadow: var(--shadow-md);
    }
    .audio-player-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 0.85rem;
      font-family: var(--font-ui);
    }
    .audio-badge {
      font-size: 0.8rem;
      font-weight: 700;
      color: var(--accent-emerald);
      text-transform: uppercase;
      letter-spacing: 0.06em;
    }
    .audio-filename {
      font-size: 0.78rem;
      color: var(--text-muted);
      font-family: var(--font-mono);
    }
    .native-audio-element {
      width: 100%;
      outline: none;
      border-radius: 30px;
    }

    /* Document Sections & Continuous Body Elements */
    .doc-section {
      margin-bottom: 3rem;
      scroll-margin-top: calc(var(--sticky-top-total) + 20px);
    }
    .section-divider {
      border: 0;
      height: 1px;
      background: linear-gradient(90deg, transparent, var(--border-color), transparent);
      margin: 3rem 0;
    }

    /* Case Header Card */
    .case-header-card {
      background: var(--card-bg);
      border: 1px solid var(--border-color);
      border-left: 5px solid var(--accent-gold);
      border-radius: 10px;
      padding: 1.25rem 1.5rem;
      margin-bottom: 1.5rem;
      box-shadow: var(--shadow-sm);
    }
    .case-number-pill {
      display: inline-block;
      background: var(--accent-gold);
      color: #0b1120;
      font-family: var(--font-ui);
      font-size: 0.72rem;
      font-weight: 800;
      padding: 2px 8px;
      border-radius: 4px;
      margin-bottom: 0.5rem;
      letter-spacing: 0.05em;
    }
    .case-header-title {
      font-family: var(--font-heading);
      font-size: 1.35rem;
      font-weight: 700;
      color: var(--text-primary);
      line-height: 1.35;
    }

    .topic-header-card {
      background: var(--bg-secondary);
      border-left: 4px solid var(--accent-blue);
      border-radius: 8px;
      padding: 1rem 1.25rem;
      margin-bottom: 1.5rem;
    }
    .topic-header-title {
      font-family: var(--font-heading);
      font-size: 1.2rem;
      font-weight: 700;
      color: var(--accent-blue);
    }

    .subtopic-header-card {
      margin: 1.5rem 0 1rem;
    }
    .subtopic-header-title {
      font-family: var(--font-ui);
      font-size: 1.05rem;
      font-weight: 700;
      color: var(--accent-gold);
    }

    /* Case Citation Banner */
    .case-citation-banner {
      background-color: var(--bg-tertiary);
      border: 1px dashed var(--border-color);
      border-radius: 8px;
      padding: 0.75rem 1.25rem;
      margin-bottom: 1.5rem;
      font-family: var(--font-mono);
      font-size: 0.85rem;
      color: var(--text-secondary);
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }
    .citation-label {
      display: inline-block;
      font-family: var(--font-ui);
      font-size: 0.72rem;
      font-weight: 800;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--accent-gold);
      background: rgba(251, 191, 36, 0.12);
      border: 1px solid rgba(251, 191, 36, 0.25);
      padding: 2px 7px;
      border-radius: 4px;
    }

    /* Subheadings within Case (Facts, Issue, Ruling) */
    .case-subheading {
      font-family: var(--font-ui);
      font-size: 1.05rem;
      font-weight: 700;
      color: var(--text-primary);
      margin: 1.75rem 0 0.75rem;
      display: flex;
      align-items: center;
      gap: 0.5rem;
      border-bottom: 1px solid var(--border-color);
      padding-bottom: 0.4rem;
    }
    .subheading-accent {
      color: var(--accent-gold);
      font-size: 1.2rem;
    }

    /* Continuous Flowing Paragraphs */
    .case-paragraph {
      margin-bottom: 1.25rem;
      text-align: justify;
      text-justify: inter-word;
      line-height: 1.85;
      padding: 4px 6px;
      border-radius: 6px;
      transition: background-color 0.15s ease;
    }

    /* ALAC & Reasoning Paragraph Badges */
    .alac-paragraph {
      background: rgba(255, 255, 255, 0.02);
      border-left: 3px solid var(--border-color);
      padding: 0.6rem 0.85rem;
      margin-bottom: 1.25rem;
    }
    .alac-badge {
      display: inline-block;
      font-family: var(--font-ui);
      font-size: 0.75rem;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      padding: 2px 7px;
      border-radius: 4px;
      margin-right: 0.5rem;
      vertical-align: baseline;
    }
    .badge-ans { background: rgba(56, 189, 248, 0.2); color: var(--accent-blue); border: 1px solid var(--accent-blue); }
    .badge-law { background: rgba(192, 132, 252, 0.2); color: var(--accent-purple); border: 1px solid var(--accent-purple); }
    .badge-app { background: rgba(251, 191, 36, 0.2); color: var(--accent-gold); border: 1px solid var(--accent-gold); }
    .badge-con { background: rgba(52, 211, 153, 0.2); color: var(--accent-emerald); border: 1px solid var(--accent-emerald); }
    .badge-syn { background: rgba(248, 113, 113, 0.2); color: var(--accent-crimson); border: 1px solid var(--accent-crimson); }
    .badge-gen { background: rgba(148, 163, 184, 0.2); color: var(--text-secondary); border: 1px solid var(--border-color); }

    /* Bullet Points */
    .bullet-point {
      display: flex;
      align-items: baseline;
      gap: 0.75rem;
      margin-bottom: 0.75rem;
      padding: 2px 6px;
      border-radius: 4px;
    }
    .bullet-dot {
      color: var(--accent-gold);
      font-size: 1.2rem;
      line-height: 1;
    }
    .bullet-content {
      flex: 1;
      text-align: justify;
    }

    /* Tables */
    .table-responsive {
      overflow-x: auto;
      margin: 1.5rem 0;
      border: 1px solid var(--border-color);
      border-radius: 8px;
    }
    .reader-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.92rem;
      font-family: var(--font-ui);
    }
    .reader-table th, .reader-table td {
      padding: 10px 14px;
      border: 1px solid var(--border-color);
      text-align: left;
    }
    .reader-table th {
      background-color: var(--bg-secondary);
      font-weight: 700;
      color: var(--accent-gold);
    }
    .reader-table tr:nth-child(even) {
      background-color: rgba(255, 255, 255, 0.02);
    }

    /* Active TTS Highlight */
    .read-unit.is-speaking {
      background-color: var(--highlight-bg) !important;
      border-left: 3px solid var(--highlight-border) !important;
      border-radius: 4px;
      box-shadow: 0 0 15px rgba(251, 191, 36, 0.15);
    }

    /* Responsive adjustments for Tablets and Mobile */
    @media (max-width: 900px) {
      :root {
        --header-height: 56px;
        --toolbar-height: auto;
      }
      .tts-toolbar {
        top: var(--header-height);
        padding: 0.5rem 1rem;
        gap: 0.6rem;
      }
      .tts-controls-group {
        flex-wrap: wrap;
      }
      .sidebar-toc {
        position: fixed;
        left: 0;
        top: 0;
        height: 100vh;
        width: 300px;
        max-width: 85vw;
        z-index: 1000;
        transform: translateX(-100%);
        box-shadow: 0 0 30px rgba(0, 0, 0, 0.5);
        background-color: var(--bg-secondary);
        display: block !important;
        opacity: 1 !important;
        visibility: visible !important;
        pointer-events: auto !important;
        padding: 1.25rem 1rem !important;
      }
      .sidebar-toc.open {
        transform: translateX(0) !important;
      }
      body.sidebar-collapsed .sidebar-toc {
        transform: translateX(-100%) !important;
      }
      #restoreSidebarBtn {
        display: none !important;
      }
      .reader-main {
        padding: 1.5rem 1rem 5rem !important;
        max-width: 100% !important;
      }
      body.sidebar-collapsed .reader-main {
        padding: 1.5rem 1rem 5rem !important;
        max-width: 100% !important;
      }
      #sidebarOverlay {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0, 0, 0, 0.65);
        backdrop-filter: blur(4px);
        z-index: 999;
        display: none;
      }
      #sidebarOverlay.active {
        display: block;
      }
    }
  </style>
</head>
<body>
  <div id="readingProgressBar"></div>
  <button id="floatingTtsTrigger" title="Read Selected Text">Read Selection</button>
  <button id="restoreSidebarBtn" title="Show Table of Contents (Ctrl+B)">☰ Table of Contents</button>
  <div id="sidebarOverlay"></div>

  <header>
    <div class="header-left">
      <button class="btn-icon" id="toggleSidebarBtn" title="Toggle Table of Contents (Ctrl+B)">☰</button>
      <div>
        <h1 class="brand-title">__ESCAPED_TITLE__</h1>
        <span class="subject-pill">__ESCAPED_SUBJECT_TAG__</span>
      </div>
    </div>
    <div class="header-actions">
      <button class="btn-icon" id="fontDecBtn" title="Decrease Font Size">A-</button>
      <button class="btn-icon" id="fontIncBtn" title="Increase Font Size">A+</button>
      <select class="select-control" id="themeSelect" title="Select Theme">
        <option value="dark">Dark</option>
        <option value="sepia">Sepia</option>
        <option value="light">Light</option>
      </select>
    </div>
  </header>

  <div class="tts-toolbar">
    <div class="tts-controls-group">
      <button class="btn-tts" id="playPauseBtn">
        <span id="playIcon">▶</span>
        <span id="pauseIcon" style="display:none;">❚❚</span>
        <span id="playBtnText">Read Aloud</span>
      </button>
      <button class="btn-tts-secondary" id="prevBtn" title="Previous Paragraph (Key: P)">Previous</button>
      <button class="btn-tts-secondary" id="nextBtn" title="Next Paragraph (Key: N)">Next</button>
      <button class="btn-tts-secondary btn-tts-stop" id="stopBtn" title="Stop Reading Aloud (Key: Esc)">⏹ Stop</button>
    </div>

    <div class="tts-controls-group">
      <label for="voiceSelect" style="font-size:0.8rem; color:var(--text-muted);">Voice:</label>
      <select class="select-control" id="voiceSelect" style="max-width: 220px;">
        <option value="default">Auto Best Voice</option>
      </select>

      <label for="speedSelect" style="font-size:0.8rem; color:var(--text-muted);">Speed:</label>
      <select class="select-control" id="speedSelect">
        <option value="0.8">0.8x</option>
        <option value="0.9" selected>0.9x (Natural)</option>
        <option value="1.0">1.0x (Standard)</option>
        <option value="1.1">1.1x</option>
        <option value="1.2">1.2x</option>
      </select>

      <div class="tts-status-badge" id="ttsStatusBadge">
        <span class="pulse-dot"></span>
        <span id="statusText">Ready</span>
      </div>
    </div>
  </div>

  <div class="app-layout">
    <nav class="sidebar-toc" id="sidebarNav">
      <div class="toc-header-bar">
        <div class="toc-heading">
          <span>Table of Contents</span>
        </div>
        <button class="btn-toc-minimize" id="minimizeSidebarBtn" title="Minimize / Toggle Sidebar (Ctrl+B)">☰</button>
      </div>
      <input type="text" class="toc-search-box" id="sidebarSearch" placeholder="Filter topics & cases..." />
      <ul class="toc-list" id="tocList">
__TOC_HTML__
      </ul>
    </nav>

    <main class="reader-main" id="mainArticle">
      <div class="doc-meta-banner">
        <h1 class="doc-headline">__ESCAPED_TITLE__</h1>
        <div class="doc-stats">
          <span>__TOTAL_SECTIONS__ Sections</span>
          <span>~__READING_TIME__ min read</span>
          <span>Manila Law College</span>
        </div>
      </div>

      __MP3_PLAYER_HTML__

      <div id="contentWrapper">
__SECTIONS_HTML__
      </div>
    </main>
  </div>

  <script>
    (function() {
      const synth = window.speechSynthesis;
      let voices = [];
      const readUnits = Array.from(document.querySelectorAll('.read-unit'));
      const playPauseBtn = document.getElementById('playPauseBtn');
      const playIcon = document.getElementById('playIcon');
      const pauseIcon = document.getElementById('pauseIcon');
      const playBtnText = document.getElementById('playBtnText');
      const stopBtn = document.getElementById('stopBtn');
      const prevBtn = document.getElementById('prevBtn');
      const nextBtn = document.getElementById('nextBtn');
      const voiceSelect = document.getElementById('voiceSelect');
      const speedSelect = document.getElementById('speedSelect');
      const statusBadge = document.getElementById('ttsStatusBadge');
      const progressBar = document.getElementById('readingProgressBar');
      const toggleSidebarBtn = document.getElementById('toggleSidebarBtn');
      const minimizeSidebarBtn = document.getElementById('minimizeSidebarBtn');
      const restoreSidebarBtn = document.getElementById('restoreSidebarBtn');
      const sidebarOverlay = document.getElementById('sidebarOverlay');
      const sidebarNav = document.getElementById('sidebarNav');
      const sidebarSearch = document.getElementById('sidebarSearch');
      const tocList = document.getElementById('tocList');
      const themeSelect = document.getElementById('themeSelect');
      const fontIncBtn = document.getElementById('fontIncBtn');
      const fontDecBtn = document.getElementById('fontDecBtn');

      let currentUnitIndex = -1;
      let isPaused = false;
      let isSpeaking = false;

      // 1. Sidebar Minimize / Maximize & Responsive Adaptations
      function isMobile() {
        return window.innerWidth <= 900;
      }

      function toggleSidebar(forceState) {
        if (isMobile()) {
          const isOpen = (forceState !== undefined) ? forceState : !sidebarNav.classList.contains('open');
          sidebarNav.classList.toggle('open', isOpen);
          if (sidebarOverlay) {
            sidebarOverlay.classList.toggle('active', isOpen);
          }
        } else {
          // Desktop toggle minimize / maximize
          const isCurrentlyCollapsed = document.body.classList.contains('sidebar-collapsed');
          const shouldCollapse = (forceState !== undefined) ? forceState : !isCurrentlyCollapsed;
          document.body.classList.toggle('sidebar-collapsed', shouldCollapse);
          sidebarNav.classList.toggle('minimized', shouldCollapse);
          try {
            localStorage.setItem('mlc_sidebar_collapsed', shouldCollapse ? 'true' : 'false');
          } catch (e) {}
        }
      }

      // Restore saved sidebar state
      try {
        const savedCollapsed = localStorage.getItem('mlc_sidebar_collapsed');
        if (!isMobile() && savedCollapsed === 'true') {
          toggleSidebar(true);
        }
      } catch (e) {}

      if (toggleSidebarBtn) {
        toggleSidebarBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          toggleSidebar();
        });
      }
      if (minimizeSidebarBtn) {
        minimizeSidebarBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          toggleSidebar();
        });
      }
      if (restoreSidebarBtn) {
        restoreSidebarBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          toggleSidebar(false);
        });
      }
      if (sidebarOverlay) {
        sidebarOverlay.addEventListener('click', () => {
          sidebarNav.classList.remove('open');
          sidebarOverlay.classList.remove('active');
        });
      }

      // Auto close mobile drawer when a TOC link is clicked
      if (tocList) {
        tocList.querySelectorAll('.toc-link').forEach(link => {
          link.addEventListener('click', () => {
            if (isMobile()) {
              sidebarNav.classList.remove('open');
              if (sidebarOverlay) sidebarOverlay.classList.remove('active');
            }
          });
        });
      }

      // Instant live filter
      if (sidebarSearch && tocList) {
        sidebarSearch.addEventListener('input', (e) => {
          const q = e.target.value.toLowerCase();
          const items = tocList.querySelectorAll('li');
          items.forEach(li => {
            const txt = li.textContent.toLowerCase();
            li.style.display = txt.includes(q) ? '' : 'none';
          });
        });
      }

      // 2. Theme & Font Scaling
      if (themeSelect) {
        themeSelect.addEventListener('change', (e) => {
          document.documentElement.setAttribute('data-theme', e.target.value);
          try { localStorage.setItem('mlc_theme', e.target.value); } catch(err) {}
        });
        try {
          const savedTheme = localStorage.getItem('mlc_theme');
          if (savedTheme) {
            themeSelect.value = savedTheme;
            document.documentElement.setAttribute('data-theme', savedTheme);
          }
        } catch(err) {}
      }

      let currentFontSize = 17;
      if (fontIncBtn) {
        fontIncBtn.addEventListener('click', () => {
          if (currentFontSize < 24) {
            currentFontSize += 1;
            document.documentElement.style.fontSize = currentFontSize + 'px';
          }
        });
      }
      if (fontDecBtn) {
        fontDecBtn.addEventListener('click', () => {
          if (currentFontSize > 13) {
            currentFontSize -= 1;
            document.documentElement.style.fontSize = currentFontSize + 'px';
          }
        });
      }

      // 3. Scroll Progress & Active TOC Link Sync
      window.addEventListener('scroll', () => {
        const sTop = window.scrollY;
        const dHeight = document.documentElement.scrollHeight - window.innerHeight;
        if (dHeight > 0 && progressBar) {
          progressBar.style.width = ((sTop / dHeight) * 100) + '%';
        }

        // Sync TOC Active Link
        const sections = document.querySelectorAll('section.doc-section');
        let currentActive = null;
        sections.forEach(sec => {
          const rect = sec.getBoundingClientRect();
          if (rect.top <= 160 && rect.bottom >= 160) {
            currentActive = sec.id;
          }
        });
        if (currentActive) {
          document.querySelectorAll('.toc-link').forEach(link => {
            link.classList.toggle('active', link.getAttribute('href') === '#' + currentActive);
          });
        }
      });

      // 4. Voice Population with High-Quality Natural Neural Prioritization
      function populateVoices() {
        if (!synth || !voiceSelect) return;
        voices = synth.getVoices();
        voiceSelect.innerHTML = '<option value="default">Auto Best Voice</option>';

        const scored = voices.map((v, i) => {
          let score = 0;
          const name = v.name.toLowerCase();
          const lang = v.lang.toLowerCase();

          if (lang.startsWith('en')) score += 50;
          if (lang.includes('ph') || lang.includes('fil')) score += 30;
          if (lang.includes('us')) score += 20;
          if (lang.includes('gb') || lang.includes('uk')) score += 15;
          if (name.includes('natural')) score += 40;
          if (name.includes('neural')) score += 40;
          if (name.includes('online')) score += 25;
          if (name.includes('multilingual')) score += 20;
          if (name.includes('guy') || name.includes('aria') || name.includes('jenny') || name.includes('andrew')) score += 15;
          if (v.default) score += 5;

          return { voice: v, index: i, score: score };
        });

        scored.sort((a, b) => b.score - a.score);

        scored.forEach(item => {
          const opt = document.createElement('option');
          opt.value = item.index;
          opt.textContent = item.voice.name + ' (' + item.voice.lang + ')';
          voiceSelect.appendChild(opt);
        });
      }

      populateVoices();
      if (speechSynthesis && speechSynthesis.onvoiceschanged !== undefined) {
        speechSynthesis.onvoiceschanged = populateVoices;
      }

      // 5. Roman Numeral & Phonetics Normalization
      function romanToInt(s) {
        if (!s) return null;
        s = s.toUpperCase().trim();
        if (!/^[IVXLCDM]+$/.test(s)) return null;
        const map = { 'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000 };
        let total = 0;
        let i = 0;
        while (i < s.length) {
          if (i + 1 < s.length && map[s[i]] < map[s[i+1]]) {
            total += map[s[i+1]] - map[s[i]];
            i += 2;
          } else if (map[s[i]]) {
            total += map[s[i]];
            i += 1;
          } else {
            return null;
          }
        }
        return total;
      }

      const ORDINALS = {
        'I': 'the first', 'II': 'the second', 'III': 'the third', 'IV': 'the fourth', 'V': 'the fifth',
        'VI': 'the sixth', 'VII': 'the seventh', 'VIII': 'the eighth', 'IX': 'the ninth', 'X': 'the tenth'
      };

      function cleanSmartReading(txt) {
        if (!txt) return '';

        // A. Strip all emojis and decorative symbols completely
        txt = Array.from(txt).filter(c => {
          const cp = c.codePointAt(0);
          return cp < 0x2000 || (cp > 0x2BFF && cp < 0x1F000);
        }).join('');
        txt = txt.replace(/[•§]/g, ' ');

        // B. Smart Number-Word Deduplication
        const numWordPattern = /\b(zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|twenty-five|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|thousand|million)\s*\(\s*\d+\s*\)/gi;
        txt = txt.replace(numWordPattern, '$1');
        txt = txt.replace(/\b\d+\s*\(\s*([a-zA-Z\-]+)\s*\)/g, '$1');

        // C. Smart Tag / Bracket Deduplication
        txt = txt.replace(/\[\s*[A-Z]\s*[\-–—]?\s*(?:ANSWER|LEGAL\s+BASIS|APPLICATION|ANALYSIS|CONCLUSION)?\s*\]\s*/gi, '');
        txt = txt.replace(/\[\s*([^\]]+)\s*\]\s*[:\-–—]?\s*([\s\S]*)/g, function(match, tag, rest) {
          tag = tag.trim();
          if (rest.toLowerCase().startsWith(tag.toLowerCase())) {
            return rest;
          }
          return tag + ': ' + rest;
        });

        txt = txt.replace(/\b([A-Za-z]{3,})\b\s+\1\b/gi, '$1');
        txt = txt.replace(/:\s*:/g, ':');
        txt = txt.replace(/\s+/g, ' ').trim();
        return txt;
      }

      function expandRomanAndPhonetics(txt) {
        if (!txt) return '';
        txt = cleanSmartReading(txt);

        const replacements = [
          [/\bA\.C\.\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)/gi, 'Administrative Case Number $1'],
          [/\bA\.M\.\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)/gi, 'Administrative Matter Number $1'],
          [/\bG\.R\.\s*Nos\.?\s*([A-Za-z0-9\-,\s]+)/gi, 'JEE-AR Numbers $1'],
          [/\bG\.R\.\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)/gi, 'JEE-AR Number $1'],
          [/\bG\.R\.\b/gi, 'JEE-AR'],
          [/\bP\.D\.\s*(?:No\.?\s*)?(\d+)/gi, 'Presidential Decree Number $1'],
          [/\bPD\s*(\d+)/gi, 'Presidential Decree $1'],
          [/\bP\.D\.\b/gi, 'Presidential Decree'],
          [/\bR\.A\.\s*(?:No\.?\s*)?(\d+)/gi, 'Republic Act Number $1'],
          [/\bRA\s*(\d+)/gi, 'Republic Act $1'],
          [/\bR\.A\.\b/gi, 'Republic Act'],
          [/\bPhil\.\s*(\d+)/gi, 'Philippine Reports volume $1'],
          [/\bPhil\.\b/gi, 'Phil'],
          [/\bSCRA\b/g, 'SKRA'],
          [/\bCPRA\b/g, 'SIP-ruh'],
          [/\bCJCA\b/g, 'SEE-JAY-SEE-AY'],
          [/\bCPR\b/g, 'Code of Professional Responsibility'],
          [/\bCCCP\b/g, 'Code of Conduct for Court Personnel'],
          [/\bJIO\b/g, 'Judicial Integrity Office'],
          [/\bOCA\b/g, 'Office of the Court Administrator'],
          [/\bDPA\b/g, 'DEE-PEE-AY'],
          [/\bITA\b/g, 'EYE-tuh'],
          [/\bOSAEC\b/g, 'OH-sak'],
          [/\bAFASA\b/g, 'ah-FAH-suh'],
          [/\bCPA\b/g, 'Cybercrime Prevention Act'],
          [/\bREED\b/g, 'REED'],
          [/\bDICT\b/g, 'DIK-tee'],
          [/\bCICC\b/g, 'SIK-see'],
          [/\bRPC\b/g, 'Revised Penal Code'],
          [/\bIn\s+re\b/gi, 'in Ree'],
          [/\bet\s+al\./gi, 'et AHL,'],
          [/\bet\s+al\b/gi, 'et AHL,'],
          [/\bi\.e\./gi, 'that is,'],
          [/\be\.g\./gi, 'for example,'],
          [/\bArt\.\s*(\d+)/gi, 'Article $1'],
          [/\bArts\.\s*([\d,\s\-]+)/gi, 'Articles $1'],
          [/\bSec\.\s*(\d+)/gi, 'Section $1'],
          [/\bSecs\.\s*([\d,\s\-]+)/gi, 'Sections $1'],
          [/\bPar\.\s*(\d+)/gi, 'Paragraph $1'],
          [/\s+v(?:s)?\.\s+/gi, ' versus ']
        ];

        replacements.forEach(([pattern, rep]) => {
          txt = txt.replace(pattern, rep);
        });

        txt = txt.replace(/(^|\n|\.\s+|;\s+)([IVXLCDM]+)\.\s+/gi, function(match, prefix, roman) {
          const val = romanToInt(roman);
          return val ? prefix + 'Topic ' + val + ': ' : match;
        });

        txt = txt.replace(/\b([A-Z][a-z]+)\s+([IVXLCDM]{1,4})\b/g, function(match, name, roman) {
          if (/^(?:Canon|Part|Chapter|Article|Title|Book|Section|Sec|Art|Vol|Volume|Rule|Topic|Case|No|Nos)$/i.test(name)) {
            return match;
          }
          const upperR = roman.toUpperCase();
          return ORDINALS[upperR] ? name + ' ' + ORDINALS[upperR] : match;
        });

        txt = txt.replace(/\b(Canon|Part|Chapter|Article|Title|Book|Section|Sec|Art|Vol|Volume|Rule)\s+([IVXLCDM]+)\b/gi, function(match, prefix, roman) {
          const val = romanToInt(roman);
          return val ? prefix + ' ' + val : match;
        });

        txt = txt.replace(/\(([ivxlcdm]+)\)/gi, function(match, roman) {
          const val = romanToInt(roman);
          return val ? 'sub-item ' + val + ', ' : match;
        });

        txt = txt.replace(/:\s*/g, ': ... ');
        txt = txt.replace(/;\s*/g, ';, ');
        txt = txt.replace(/\s*—\s*/g, ' — ... ');

        return txt.trim();
      }

      function prepareSpeechText(unitEl) {
        let text = unitEl.innerText || unitEl.textContent || '';
        return expandRomanAndPhonetics(text);
      }

      function setSpeakingState(speaking, paused = false) {
        isSpeaking = speaking;
        isPaused = paused;
        if (speaking && !paused) {
          if (playIcon) playIcon.style.display = 'none';
          if (pauseIcon) pauseIcon.style.display = 'inline';
          if (playBtnText) playBtnText.textContent = 'Pause';
          if (statusBadge) {
            statusBadge.innerHTML = '<span class="pulse-dot"></span><span>Speaking...</span>';
            statusBadge.className = 'tts-status-badge speaking';
          }
        } else if (paused) {
          if (playIcon) playIcon.style.display = 'inline';
          if (pauseIcon) pauseIcon.style.display = 'none';
          if (playBtnText) playBtnText.textContent = 'Resume';
          if (statusBadge) {
            statusBadge.innerHTML = '<span class="pulse-dot"></span><span>Paused</span>';
            statusBadge.className = 'tts-status-badge paused';
          }
        } else {
          if (playIcon) playIcon.style.display = 'inline';
          if (pauseIcon) pauseIcon.style.display = 'none';
          if (playBtnText) playBtnText.textContent = 'Read Aloud';
          if (statusBadge) {
            statusBadge.innerHTML = '<span class="pulse-dot"></span><span>Ready</span>';
            statusBadge.className = 'tts-status-badge';
          }
          clearHighlights();
        }
      }

      function clearHighlights() {
        readUnits.forEach(u => u.classList.remove('is-speaking'));
      }

      function highlightUnit(idx) {
        clearHighlights();
        if (idx >= 0 && idx < readUnits.length) {
          const unit = readUnits[idx];
          unit.classList.add('is-speaking');
          unit.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }

      function speakUnit(idx) {
        if (!synth || idx < 0 || idx >= readUnits.length) {
          setSpeakingState(false);
          return;
        }

        synth.cancel();
        currentUnitIndex = idx;
        const unit = readUnits[idx];
        highlightUnit(idx);

        const speechText = prepareSpeechText(unit);
        const utterance = new SpeechSynthesisUtterance(speechText);
        utterance.rate = (speedSelect ? parseFloat(speedSelect.value) : 0.9) || 0.9;
        utterance.pitch = 1.0;

        const selVoiceIdx = voiceSelect ? voiceSelect.value : 'default';
        if (selVoiceIdx !== 'default' && voices[selVoiceIdx]) {
          utterance.voice = voices[selVoiceIdx];
        } else if (voices.length > 0) {
          utterance.voice = voices[0];
        }

        utterance.onstart = () => {
          setSpeakingState(true, false);
        };

        utterance.onend = () => {
          if (isSpeaking && !isPaused) {
            if (idx + 1 < readUnits.length) {
              speakUnit(idx + 1);
            } else {
              setSpeakingState(false);
            }
          }
        };

        utterance.onerror = (e) => {
          if (e.error !== 'interrupted' && e.error !== 'canceled') {
            console.warn('TTS error:', e);
          }
        };

        synth.speak(utterance);
      }

      if (playPauseBtn) {
        playPauseBtn.addEventListener('click', () => {
          if (!synth) return;
          if (isSpeaking && !isPaused) {
            synth.pause();
            setSpeakingState(true, true);
          } else if (isPaused) {
            synth.resume();
            setSpeakingState(true, false);
          } else {
            speakUnit(currentUnitIndex >= 0 ? currentUnitIndex : 0);
          }
        });
      }

      if (stopBtn) {
        stopBtn.addEventListener('click', () => {
          if (!synth) return;
          synth.cancel();
          setSpeakingState(false);
          currentUnitIndex = -1;
          clearHighlights();
        });
      }

      if (nextBtn) {
        nextBtn.addEventListener('click', () => {
          const target = Math.min(readUnits.length - 1, (currentUnitIndex >= 0 ? currentUnitIndex : 0) + 1);
          speakUnit(target);
        });
      }

      if (prevBtn) {
        prevBtn.addEventListener('click', () => {
          const target = Math.max(0, (currentUnitIndex >= 0 ? currentUnitIndex : 0) - 1);
          speakUnit(target);
        });
      }

      // Click on any paragraph to start speaking from that exact unit
      const contentWrap = document.getElementById('contentWrapper');
      if (contentWrap) {
        contentWrap.addEventListener('click', (e) => {
          const unit = e.target.closest('.read-unit');
          if (unit) {
            const idx = readUnits.indexOf(unit);
            if (idx !== -1) {
              speakUnit(idx);
            }
          }
        });
      }

      // Floating Selection Reader
      const floatBtn = document.getElementById('floatingTtsTrigger');
      if (floatBtn) {
        document.addEventListener('mouseup', (e) => {
          if (e.target.closest('#floatingTtsTrigger') || e.target.closest('.tts-toolbar') || e.target.closest('header')) return;
          const sel = window.getSelection().toString().trim();
          if (sel.length > 2) {
            const rect = window.getSelection().getRangeAt(0).getBoundingClientRect();
            floatBtn.style.top = (window.scrollY + rect.top - 44) + 'px';
            floatBtn.style.left = (window.scrollX + rect.left + (rect.width / 2)) + 'px';
            floatBtn.style.display = 'block';
          } else {
            floatBtn.style.display = 'none';
          }
        });

        floatBtn.addEventListener('click', () => {
          const sel = window.getSelection().toString().trim();
          if (sel && synth) {
            synth.cancel();
            const utterance = new SpeechSynthesisUtterance(expandRomanAndPhonetics(sel));
            utterance.rate = (speedSelect ? parseFloat(speedSelect.value) : 0.9) || 0.9;
            const selVoiceIdx = voiceSelect ? voiceSelect.value : 'default';
            if (selVoiceIdx !== 'default' && voices[selVoiceIdx]) {
              utterance.voice = voices[selVoiceIdx];
            }
            synth.speak(utterance);
            floatBtn.style.display = 'none';
          }
        });
      }

      // Keyboard Shortcuts
      document.addEventListener('keydown', (e) => {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
        if ((e.ctrlKey && (e.key === 'b' || e.key === 'B')) || (e.altKey && (e.key === 't' || e.key === 'T'))) {
          e.preventDefault();
          toggleSidebar();
        } else if (e.code === 'Space') {
          e.preventDefault();
          if (playPauseBtn) playPauseBtn.click();
        } else if (e.key === 'Escape') {
          if (stopBtn) stopBtn.click();
        } else if (e.key === 'n' || e.key === 'N') {
          if (nextBtn) nextBtn.click();
        } else if (e.key === 'p' || e.key === 'P') {
          if (prevBtn) prevBtn.click();
        }
      });
    })();
  </script>
</body>
</html>'''

def generate_reader_html(doc_title, subject_tag, sections, mp3_filename=None):
    """
    Renders the complete self-contained interactive reader HTML app using token replacement.
    """
    escaped_title = html.escape(doc_title)
    escaped_subject_tag = html.escape(subject_tag)
    
    total_sections = len(sections)
    total_units = sum(len(s.get("units", [])) for s in sections)
    reading_time_minutes = max(1, round(total_units * 0.4))

    # Build TOC HTML
    toc_items = []
    for s in sections:
        s_title = html.escape(s.get("title", "Section"))
        s_id = s.get("id", "sec")
        s_lvl = s.get("level", 1)
        lvl_class = f"level-{s_lvl}"
        toc_items.append(f'<li><a href="#{s_id}" class="toc-link {lvl_class}" title="{s_title}">{s_title}</a></li>')
    toc_html = "\n".join(toc_items)

    # Build Content HTML
    sec_html_list = []
    for s in sections:
        s_id = s.get("id", "sec")
        units_html = "\n".join(s.get("units", []))
        sec_html_list.append(f'<section id="{s_id}" class="doc-section">\n{units_html}\n</section>')
    sections_html = "\n<hr class=\"section-divider\" />\n".join(sec_html_list)

    # Audio player snippet with Multi-Release Cloud Stream URL & Local Fallback
    mp3_player_html = ""
    if mp3_filename:
        encoded_mp3 = urllib.parse.quote(mp3_filename)
        stem = mp3_filename[:-4] if mp3_filename.endswith(".mp3") else mp3_filename
        dot_stem = re.sub(r"[ &()]", ".", stem)
        dot_name = re.sub(r"\.+", ".", dot_stem).strip(".") + ".mp3"
        
        candidate_names = [
            dot_name,
            mp3_filename,
            encoded_mp3,
        ]
        if "canons & its section" in mp3_filename.lower() or "new" in mp3_filename.lower():
            candidate_names.extend([
                "New.Canons.its.Section.mp3",
                "%5BNew%5D%20Canons%20%26%20its%20Section.mp3",
                "[New] Canons & its Section.mp3",
                "Canons.its.Section.mp3",
                "Canons.and.its.Section.mp3"
            ])
            
        tags = ["audio-v0", "audio-v1", "audio-v2"]
        source_elements = []
        seen_urls = set()
        
        for tag in tags:
            for cname in candidate_names:
                url_cname = urllib.parse.quote(cname) if not cname.startswith("%") and any(c in cname for c in " &()[]") else cname
                url = f"https://github.com/Julius11011/MLLibrary/releases/download/{tag}/{url_cname}"
                if url not in seen_urls:
                    seen_urls.add(url)
                    source_elements.append(f'            <source src="{url}" type="audio/mpeg">')
                    
        source_elements.append(f'            <source src="{html.escape(mp3_filename)}" type="audio/mpeg">')
        sources_str = "\n".join(source_elements)
        
        mp3_player_html = f"""
        <div class="studio-audio-player">
          <div class="audio-player-header">
            <span class="audio-badge">Studio Voice Podcast (Cloud Stream)</span>
            <span class="audio-filename">{html.escape(mp3_filename)}</span>
          </div>
          <audio controls preload="metadata" class="native-audio-element">
{sources_str}
            Your browser does not support the audio element.
          </audio>
        </div>
        """

    # Template token injection
    rendered = READER_HTML_TEMPLATE
    rendered = rendered.replace('__ESCAPED_TITLE__', escaped_title)
    rendered = rendered.replace('__ESCAPED_SUBJECT_TAG__', escaped_subject_tag)
    rendered = rendered.replace('__TOTAL_SECTIONS__', str(total_sections))
    rendered = rendered.replace('__READING_TIME__', str(reading_time_minutes))
    rendered = rendered.replace('__MP3_PLAYER_HTML__', mp3_player_html)
    rendered = rendered.replace('__TOC_HTML__', toc_html)
    rendered = rendered.replace('__SECTIONS_HTML__', sections_html)
    return rendered

def convert_file_to_html_reader(input_file_path, output_html_path=None, overwrite=True):
    path = Path(input_file_path).resolve()
    if not path.exists():
        print(f"[ERROR] File not found: {path}")
        return None

    if output_html_path is None:
        out_path = path.with_suffix('.html')
    else:
        out_path = Path(output_html_path).resolve()

    if out_path.exists() and not overwrite:
        print(f"[SKIP] Output file already exists: {out_path.name}")
        return out_path

    doc_title = path.stem.replace('_', ' ')
    subject_tag = path.parent.name
    if subject_tag.lower() == "case digest":
        subject_tag = f"{path.parent.parent.name} • Case Digest"

    # Audio file pairing
    mp3_file = path.with_suffix('.mp3')
    mp3_filename = mp3_file.name if mp3_file.exists() else None

    print(f"[CONVERTING] {path.name} -> {out_path.name} (Subject: {subject_tag})")

    suffix = path.suffix.lower()
    if suffix == '.docx':
        sections = parse_docx_file(path, doc_title=doc_title)
    else:
        print(f"[WARN] Unsupported or non-docx file: {path.name}")
        return None

    html_content = generate_reader_html(
        doc_title=doc_title,
        subject_tag=subject_tag,
        sections=sections,
        mp3_filename=mp3_filename
    )

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"[SUCCESS] Created HTML Reader: {out_path.name} ({len(sections)} sections)")
    return out_path

def scan_and_convert_directory(dir_path, overwrite=True):
    base_dir = Path(dir_path).resolve()
    print(f"\n=== Scanning Subjects in {base_dir} ===")

    # Group files by stem to prioritize docx
    grouped = {}
    for p in base_dir.rglob('*'):
        if not p.is_file() or p.name.startswith(('~$', '.~')):
            continue
        if p.suffix.lower() != '.docx':
            continue
        key = (p.parent, p.stem.lower())
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(p)

    converted_count = 0
    generated_files = []

    for key, file_list in grouped.items():
        chosen = file_list[0]
        try:
            out = convert_file_to_html_reader(chosen, overwrite=overwrite)
            if out:
                converted_count += 1
                generated_files.append(out)
        except Exception as e:
            print(f"[ERROR] Failed to convert {chosen.name}: {e}")

    print(f"\n=== Successfully converted {converted_count} documents. ===\n")
    return generated_files

def generate_study_hub_index(root_dir):
    root_path = Path(root_dir).resolve()
    
    # Strictly scan First Sem 1st Year/Subjects
    subjects_dir = root_path / "First Sem 1st Year" / "Subjects"
    if not subjects_dir.exists():
        subjects_dir = root_path / "Subjects"
    
    html_files = list(subjects_dir.rglob("*.html")) if subjects_dir.exists() else []

    # Map each subject group cleanly with explicit support for Constitutional Law & Criminal Law
    by_subject = {
        "Basic Legal and Judiciary Ethics": [],
        "Constitutional Law": [],
        "Criminal Law": [],
        "Statutory Construction": []
    }

    subject_info = {
        "Basic Legal and Judiciary Ethics": {
            "icon": "⚖️",
            "tag": "BLJE",
            "desc": "Code of Professional Responsibility and Accountability (CPRA, A.M. No. 22-09-01-SC), Canons of Ethics, & Landmark Precedents",
            "accent": "#c084fc",
            "badge_class": "badge-ethics"
        },
        "Constitutional Law": {
            "icon": "🏛️",
            "tag": "CONSTI-1",
            "desc": "The 1987 Philippine Constitution, State Immunity Doctrine, Separation of Powers, Judicial Review, & Landmark ALAC Digests",
            "accent": "#38bdf8",
            "badge_class": "badge-consti"
        },
        "Criminal Law": {
            "icon": "🛡️",
            "tag": "CRIM-1",
            "desc": "Revised Penal Code (Act No. 3815) Book I (Articles 1–113), Felonies, Criminal Liability, Modifying Circumstances, & Supreme Court Doctrines",
            "accent": "#f87171",
            "badge_class": "badge-crim"
        },
        "Statutory Construction": {
            "icon": "📜",
            "tag": "STATCON",
            "desc": "Canons of Statutory Interpretation, Latin Maxims, Legislative Intent, Extrinsic/Intrinsic Aids, & Case Law Analysis",
            "accent": "#34d399",
            "badge_class": "badge-statcon"
        },
        "Other Subjects": {
            "icon": "📚",
            "tag": "GENERAL",
            "desc": "Supplementary Legal Compilations and Juris Doctor Course Modules",
            "accent": "#fbbf24",
            "badge_class": "badge-other"
        }
    }

    total_modules = 0
    total_audio_count = 0
    total_digest_count = 0

    for h in sorted(html_files, key=lambda x: str(x)):
        rel = h.relative_to(root_path)
        rel_str = str(rel).replace('\\\\', '/').replace('\\', '/')
        
        # Categorize
        cat = "Other Subjects"
        for sname in by_subject.keys():
            if sname.lower() in str(rel).lower():
                cat = sname
                break
        
        if cat not in by_subject:
            by_subject[cat] = []
            
        by_subject[cat].append((h, rel_str))
        total_modules += 1

    groups_html = []

    for group_name, files in by_subject.items():
        if not files:
            continue
        info = subject_info.get(group_name, subject_info["Other Subjects"])
        cards_html = []
        for fpath, rel_str in files:
            doc_name = fpath.stem.replace('_', ' ')
            
            # Companion format paths
            mp3_file = fpath.with_suffix('.mp3')
            pdf_file = fpath.with_suffix('.pdf')
            docx_file = fpath.with_suffix('.docx')
            
            has_mp3 = mp3_file.exists()
            has_pdf = pdf_file.exists()
            has_docx = docx_file.exists()

            mp3_rel = str(mp3_file.relative_to(root_path)).replace('\\\\', '/').replace('\\', '/') if has_mp3 else ''
            pdf_rel = str(pdf_file.relative_to(root_path)).replace('\\\\', '/').replace('\\', '/') if has_pdf else ''
            docx_rel = str(docx_file.relative_to(root_path)).replace('\\\\', '/').replace('\\', '/') if has_docx else ''

            if has_mp3:
                total_audio_count += 1
            
            # Type categorization
            d_lower = doc_name.lower()
            if 'digest' in d_lower:
                type_badge = '<span class="badge-type badge-digest">⚖️ Case Digest</span>'
                total_digest_count += 1
            elif 'outline' in d_lower:
                type_badge = '<span class="badge-type badge-outline">📋 Course Outline</span>'
            elif 'lecture' in d_lower:
                type_badge = '<span class="badge-type badge-lecture">📖 Comprehensive Lecture</span>'
            elif 'canon' in d_lower:
                type_badge = '<span class="badge-type badge-canon">📜 CPRA Canons</span>'
            elif 'case' in d_lower or 'landmark' in d_lower:
                type_badge = '<span class="badge-type badge-cases">🏛️ Landmark Cases</span>'
                total_digest_count += 1
            else:
                type_badge = '<span class="badge-type badge-module">📚 Study Module</span>'
            
            # Format availability pills
            format_pills = []
            format_pills.append('<span class="fmt-pill fmt-html" title="Interactive Full-Text Reader with Web Speech Synthesis">📖 HTML Reader</span>')
            if has_mp3:
                size_mb = mp3_file.stat().st_size / (1024 * 1024)
                format_pills.append(f'<span class="fmt-pill fmt-mp3" title="Studio Neural Voice MP3 Podcast ({size_mb:.1f} MB)">🎙️ MP3 Podcast</span>')
            if has_pdf:
                format_pills.append('<span class="fmt-pill fmt-pdf" title="Adobe PDF Format">📄 PDF</span>')
            if has_docx:
                format_pills.append('<span class="fmt-pill fmt-docx" title="Microsoft Word DOCX">📝 DOCX</span>')

            # Quick resource links
            resource_actions = []
            if has_mp3:
                resource_actions.append(f'<a href="{mp3_rel}" class="btn-sub btn-sub-audio" download title="Download Studio MP3 Podcast">🎧 Podcast</a>')
            if has_pdf:
                resource_actions.append(f'<a href="{pdf_rel}" class="btn-sub btn-sub-pdf" target="_blank" title="View PDF Document">📄 PDF</a>')
            if has_docx:
                resource_actions.append(f'<a href="{docx_rel}" class="btn-sub btn-sub-docx" download title="Download Word DOCX Document">📝 DOCX</a>')

            cards_html.append(f'''
              <div class="hub-card" data-subject="{html.escape(group_name)}" data-title="{html.escape(doc_name.lower())}" data-has-audio="{str(has_mp3).lower()}" data-is-digest="{str('digest' in d_lower or 'case' in d_lower).lower()}">
                <div class="hub-card-header">
                  {type_badge}
                  <div class="fmt-pills-row">
                    {''.join(format_pills)}
                  </div>
                </div>
                <h3 class="hub-card-title"><a href="{rel_str}">{html.escape(doc_name)}</a></h3>
                <div class="hub-card-footer">
                  <a href="{rel_str}" class="btn-open">📖 Open Interactive Reader →</a>
                  <div class="hub-sub-actions">
                    {''.join(resource_actions)}
                  </div>
                </div>
              </div>
            ''')

        groups_html.append(f'''
          <section class="hub-group" data-subject-group="{html.escape(group_name)}">
            <div class="hub-group-header">
              <div class="hub-group-title-wrap">
                <span class="hub-group-icon">{info['icon']}</span>
                <div>
                  <h2 class="hub-group-title">{html.escape(group_name)}</h2>
                  <p class="hub-group-desc">{html.escape(info['desc'])}</p>
                </div>
              </div>
              <span class="hub-group-count">{len(files)} Modules</span>
            </div>
            <div class="hub-grid">
              {''.join(cards_html)}
            </div>
          </section>
        ''')

    hub_page = f'''<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>MLC Law Library & Interactive Audio Suite | Manila Law College</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@700;800;900&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --bg-primary: #080d1a;
      --bg-secondary: #0f172a;
      --bg-card: rgba(19, 29, 49, 0.75);
      --bg-card-hover: rgba(28, 41, 68, 0.85);
      --border-color: rgba(148, 163, 184, 0.14);
      --border-hover: rgba(56, 189, 248, 0.4);
      --text-primary: #f8fafc;
      --text-secondary: #cbd5e1;
      --text-muted: #94a3b8;
      --accent-gold: #fbbf24;
      --accent-gold-dark: #d97706;
      --accent-gold-light: #fef3c7;
      --accent-blue: #38bdf8;
      --accent-emerald: #34d399;
      --accent-purple: #c084fc;
      --accent-rose: #fb7185;
      --font-ui: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      --font-heading: 'Cinzel', serif;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      background: radial-gradient(circle at 50% 0%, #172554 0%, var(--bg-primary) 70%);
      color: var(--text-primary);
      font-family: var(--font-ui);
      padding: 3rem 1.5rem 5rem;
      min-height: 100vh;
      -webkit-font-smoothing: antialiased;
    }}
    .hub-container {{ max-width: 1200px; margin: 0 auto; }}
    
    /* Header Section */
    .hub-header {{ text-align: center; margin-bottom: 3rem; }}
    .hub-badge {{
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      background: rgba(251, 191, 36, 0.12);
      color: var(--accent-gold);
      font-size: 0.8rem;
      font-weight: 700;
      padding: 0.4rem 1.1rem;
      border-radius: 9999px;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      margin-bottom: 1.25rem;
      border: 1px solid rgba(251, 191, 36, 0.3);
      backdrop-filter: blur(8px);
    }}
    .hub-title {{
      font-family: var(--font-heading);
      font-size: 2.75rem;
      font-weight: 900;
      background: linear-gradient(135deg, #ffffff 30%, var(--accent-gold) 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      margin-bottom: 0.85rem;
      letter-spacing: 0.02em;
    }}
    .hub-subtitle {{
      color: var(--text-secondary);
      font-size: 1.1rem;
      max-width: 780px;
      margin: 0 auto 2rem;
      line-height: 1.6;
    }}

    /* Stat Counters Grid */
    .hub-stats-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 1rem;
      margin-bottom: 3rem;
    }}
    .stat-card {{
      background: rgba(15, 23, 42, 0.6);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 1.25rem 1rem;
      text-align: center;
      backdrop-filter: blur(6px);
      transition: border-color 0.2s;
    }}
    .stat-card:hover {{ border-color: rgba(251, 191, 36, 0.35); }}
    .stat-val {{
      font-size: 1.85rem;
      font-weight: 800;
      color: var(--accent-gold);
      font-family: var(--font-heading);
      line-height: 1.2;
    }}
    .stat-label {{
      font-size: 0.8rem;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.05em;
      margin-top: 0.35rem;
    }}

    /* Search & Filter Bar */
    .hub-controls {{
      background: rgba(15, 23, 42, 0.8);
      border: 1px solid var(--border-color);
      border-radius: 16px;
      padding: 1.5rem;
      margin-bottom: 3rem;
      backdrop-filter: blur(12px);
      box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
    }}
    .search-input-wrap {{
      position: relative;
      margin-bottom: 1.25rem;
    }}
    .search-icon {{
      position: absolute;
      left: 1rem;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-muted);
      font-size: 1.1rem;
      pointer-events: none;
    }}
    .hub-search-input {{
      width: 100%;
      padding: 0.9rem 1rem 0.9rem 2.85rem;
      background: rgba(8, 13, 26, 0.7);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      color: var(--text-primary);
      font-size: 0.95rem;
      font-family: var(--font-ui);
      transition: all 0.2s ease;
    }}
    .hub-search-input:focus {{
      outline: none;
      border-color: var(--accent-blue);
      box-shadow: 0 0 0 3px rgba(56, 189, 248, 0.2);
    }}
    .hub-filters {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.6rem;
      align-items: center;
    }}
    .filter-chip {{
      background: rgba(30, 41, 59, 0.7);
      border: 1px solid var(--border-color);
      color: var(--text-secondary);
      font-size: 0.8rem;
      font-weight: 600;
      padding: 0.45rem 0.9rem;
      border-radius: 8px;
      cursor: pointer;
      transition: all 0.15s ease;
      user-select: none;
    }}
    .filter-chip:hover {{
      background: rgba(56, 189, 248, 0.15);
      border-color: var(--accent-blue);
      color: var(--text-primary);
    }}
    .filter-chip.active {{
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      border-color: var(--accent-gold);
      color: #080d1a;
      font-weight: 700;
    }}
    .results-count {{
      margin-left: auto;
      font-size: 0.82rem;
      color: var(--text-muted);
      font-weight: 500;
    }}

    /* Subject Groups */
    .hub-group {{
      margin-bottom: 3.5rem;
      transition: opacity 0.2s ease;
    }}
    .hub-group-header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 1.5rem;
      padding-bottom: 0.85rem;
      border-bottom: 1px solid var(--border-color);
    }}
    .hub-group-title-wrap {{
      display: flex;
      align-items: flex-start;
      gap: 0.85rem;
    }}
    .hub-group-icon {{
      font-size: 1.75rem;
      line-height: 1;
      padding: 0.4rem;
      background: rgba(30, 41, 59, 0.6);
      border-radius: 10px;
      border: 1px solid var(--border-color);
    }}
    .hub-group-title {{
      font-family: var(--font-heading);
      font-size: 1.45rem;
      font-weight: 800;
      color: var(--text-primary);
      letter-spacing: 0.02em;
      line-height: 1.2;
    }}
    .hub-group-desc {{
      color: var(--text-muted);
      font-size: 0.85rem;
      margin-top: 0.3rem;
      max-width: 750px;
      line-height: 1.45;
    }}
    .hub-group-count {{
      background: rgba(56, 189, 248, 0.12);
      color: var(--accent-blue);
      font-size: 0.75rem;
      font-weight: 700;
      padding: 0.3rem 0.75rem;
      border-radius: 9999px;
      border: 1px solid rgba(56, 189, 248, 0.25);
      white-space: nowrap;
    }}

    /* Cards Grid */
    .hub-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
      gap: 1.5rem;
    }}
    .hub-card {{
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 14px;
      padding: 1.5rem;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      backdrop-filter: blur(8px);
      transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    }}
    .hub-card:hover {{
      transform: translateY(-4px);
      background: var(--bg-card-hover);
      border-color: var(--border-hover);
      box-shadow: 0 16px 32px -8px rgba(0, 0, 0, 0.5), 0 0 20px rgba(56, 189, 248, 0.15);
    }}
    .hub-card.hidden {{
      display: none !important;
    }}
    .hub-card-header {{
      display: flex;
      flex-direction: column;
      gap: 0.6rem;
      margin-bottom: 1rem;
    }}
    .fmt-pills-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.35rem;
    }}
    .badge-type {{
      align-self: flex-start;
      font-size: 0.72rem;
      font-weight: 700;
      padding: 0.25rem 0.65rem;
      border-radius: 6px;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }}
    .badge-digest {{ background: rgba(56, 189, 248, 0.15); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.3); }}
    .badge-outline {{ background: rgba(192, 132, 252, 0.15); color: var(--accent-purple); border: 1px solid rgba(192, 132, 252, 0.3); }}
    .badge-lecture {{ background: rgba(251, 191, 36, 0.15); color: var(--accent-gold); border: 1px solid rgba(251, 191, 36, 0.3); }}
    .badge-canon {{ background: rgba(52, 211, 153, 0.15); color: var(--accent-emerald); border: 1px solid rgba(52, 211, 153, 0.3); }}
    .badge-cases {{ background: rgba(251, 113, 133, 0.15); color: var(--accent-rose); border: 1px solid rgba(251, 113, 133, 0.3); }}
    .badge-module {{ background: rgba(148, 163, 184, 0.15); color: var(--text-secondary); border: 1px solid rgba(148, 163, 184, 0.3); }}

    .fmt-pill {{
      font-size: 0.68rem;
      font-weight: 600;
      padding: 0.18rem 0.5rem;
      border-radius: 4px;
      border: 1px solid transparent;
    }}
    .fmt-html {{ background: rgba(56, 189, 248, 0.1); color: #7dd3fc; border-color: rgba(56, 189, 248, 0.2); }}
    .fmt-mp3 {{ background: rgba(52, 211, 153, 0.12); color: #6ee7b7; border-color: rgba(52, 211, 153, 0.25); }}
    .fmt-pdf {{ background: rgba(251, 146, 60, 0.1); color: #fdba74; border-color: rgba(251, 146, 60, 0.2); }}
    .fmt-docx {{ background: rgba(168, 85, 247, 0.1); color: #d8b4fe; border-color: rgba(168, 85, 247, 0.2); }}

    .hub-card-title {{
      font-size: 1.12rem;
      font-weight: 700;
      line-height: 1.45;
      margin-bottom: 1.75rem;
    }}
    .hub-card-title a {{
      color: var(--text-primary);
      text-decoration: none;
      transition: color 0.15s ease;
    }}
    .hub-card-title a:hover {{
      color: var(--accent-gold);
    }}

    .hub-card-footer {{
      margin-top: auto;
      display: flex;
      flex-direction: column;
      gap: 0.75rem;
    }}
    .btn-open {{
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 0.5rem;
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      color: #080d1a;
      font-weight: 700;
      font-size: 0.9rem;
      padding: 0.7rem 1.25rem;
      border-radius: 10px;
      text-decoration: none;
      text-align: center;
      transition: filter 0.15s ease, transform 0.15s ease, box-shadow 0.15s ease;
      box-shadow: 0 4px 12px rgba(217, 119, 6, 0.3);
    }}
    .btn-open:hover {{
      filter: brightness(1.1);
      transform: translateY(-2px);
      box-shadow: 0 6px 18px rgba(217, 119, 6, 0.45);
    }}
    .hub-sub-actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.4rem;
    }}
    .btn-sub {{
      flex: 1;
      min-width: 75px;
      text-align: center;
      font-size: 0.74rem;
      font-weight: 600;
      padding: 0.35rem 0.5rem;
      border-radius: 6px;
      text-decoration: none;
      transition: all 0.15s ease;
      border: 1px solid var(--border-color);
      background: rgba(30, 41, 59, 0.5);
      color: var(--text-secondary);
    }}
    .btn-sub:hover {{
      background: rgba(56, 189, 248, 0.2);
      border-color: var(--accent-blue);
      color: var(--text-primary);
    }}
    .btn-sub-audio:hover {{
      background: rgba(52, 211, 153, 0.2);
      border-color: var(--accent-emerald);
      color: #a7f3d0;
    }}
    .btn-sub-pdf:hover {{
      background: rgba(251, 146, 60, 0.2);
      border-color: #fb923c;
      color: #fed7aa;
    }}
    .btn-sub-docx:hover {{
      background: rgba(168, 85, 247, 0.2);
      border-color: #c084fc;
      color: #f3e8ff;
    }}

    footer {{
      text-align: center;
      margin-top: 5rem;
      color: var(--text-muted);
      font-size: 0.88rem;
      border-top: 1px solid var(--border-color);
      padding-top: 2.5rem;
      line-height: 1.6;
    }}
    footer a {{ color: var(--accent-gold); text-decoration: none; }}

    @media (max-width: 768px) {{
      .hub-title {{ font-size: 2rem; }}
      .hub-grid {{ grid-template-columns: 1fr; }}
      .hub-filters {{ justify-content: center; }}
      .results-count {{ width: 100%; text-align: center; margin-top: 0.5rem; }}
    }}
  </style>
</head>
<body>
  <div class="hub-container">
    <header class="hub-header">
      <div class="hub-badge">🏛️ Manila Law College • Juris Doctor Program</div>
      <h1 class="hub-title">MLC Law Library &amp; Audio Hub</h1>
      <p class="hub-subtitle">Interactive Full-Text Legal Readers with Natural Voice Synthesis, ALAC Reasoning Precedents, and High-Fidelity Studio Podcasts</p>
    </header>

    <!-- Stats Overview -->
    <div class="hub-stats-grid">
      <div class="stat-card">
        <div class="stat-val">{len([k for k, v in by_subject.items() if v])}</div>
        <div class="stat-label">Core JD Subjects</div>
      </div>
      <div class="stat-card">
        <div class="stat-val">{total_modules}</div>
        <div class="stat-label">Digital Modules</div>
      </div>
      <div class="stat-card">
        <div class="stat-val">{total_audio_count}</div>
        <div class="stat-label">Studio Podcasts</div>
      </div>
      <div class="stat-card">
        <div class="stat-val">{total_digest_count}</div>
        <div class="stat-label">ALAC Case Digests</div>
      </div>
    </div>

    <!-- Search and Filter Suite -->
    <div class="hub-controls">
      <div class="search-input-wrap">
        <span class="search-icon">🔍</span>
        <input type="text" id="hubSearch" class="hub-search-input" placeholder="Search case digests, topics, lecture titles, articles, or doctrines...">
      </div>
      <div class="hub-filters">
        <span class="filter-chip active" data-filter="all">All Subjects</span>
        <span class="filter-chip" data-filter="Basic Legal and Judiciary Ethics">⚖️ Legal Ethics</span>
        <span class="filter-chip" data-filter="Constitutional Law">🏛️ Constitutional Law</span>
        <span class="filter-chip" data-filter="Criminal Law">🛡️ Criminal Law</span>
        <span class="filter-chip" data-filter="Statutory Construction">📜 Statutory Construction</span>
        <span class="filter-chip" data-filter="audio-only">🎙️ MP3 Podcasts</span>
        <span class="filter-chip" data-filter="digest-only">⚖️ Case Digests</span>
        <span class="results-count" id="resultsCount">Showing all {total_modules} modules</span>
      </div>
    </div>

    <!-- Subject Modules -->
    <main id="hubMain">
      {''.join(groups_html)}
    </main>

    <footer>
      <div>Manila Law College (MLC) • Academic Year 2026–2027 • Juris Doctor First Semester</div>
      <div style="margin-top: 0.4rem; font-size: 0.8rem; color: var(--text-muted);">
        Interactive Speech Synthesis &amp; Studio Audio Engine • Pure Paragraph Editions
      </div>
    </footer>
  </div>

  <script>
    document.addEventListener('DOMContentLoaded', () => {{
      const searchInput = document.getElementById('hubSearch');
      const filterChips = document.querySelectorAll('.filter-chip');
      const cards = document.querySelectorAll('.hub-card');
      const groups = document.querySelectorAll('.hub-group');
      const resultsCount = document.getElementById('resultsCount');

      let currentFilter = 'all';
      let currentQuery = '';

      function filterCards() {{
        let visibleCount = 0;

        cards.forEach(card => {{
          const cardSubject = card.getAttribute('data-subject') || '';
          const cardTitle = card.getAttribute('data-title') || '';
          const hasAudio = card.getAttribute('data-has-audio') === 'true';
          const isDigest = card.getAttribute('data-is-digest') === 'true';

          let matchesFilter = true;
          if (currentFilter === 'all') {{
            matchesFilter = true;
          }} else if (currentFilter === 'audio-only') {{
            matchesFilter = hasAudio;
          }} else if (currentFilter === 'digest-only') {{
            matchesFilter = isDigest;
          }} else {{
            matchesFilter = (cardSubject === currentFilter);
          }}

          let matchesQuery = true;
          if (currentQuery) {{
            matchesQuery = cardTitle.includes(currentQuery) || cardSubject.toLowerCase().includes(currentQuery);
          }}

          if (matchesFilter && matchesQuery) {{
            card.classList.remove('hidden');
            visibleCount++;
          }} else {{
            card.classList.add('hidden');
          }}
        }});

        // Hide empty groups
        groups.forEach(group => {{
          const groupCards = group.querySelectorAll('.hub-card:not(.hidden)');
          if (groupCards.length === 0) {{
            group.style.display = 'none';
          }} else {{
            group.style.display = 'block';
          }}
        }});

        resultsCount.textContent = `Showing ${{visibleCount}} of ${{cards.length}} modules`;
      }}

      searchInput.addEventListener('input', (e) => {{
        currentQuery = e.target.value.toLowerCase().trim();
        filterCards();
      }});

      filterChips.forEach(chip => {{
        chip.addEventListener('click', () => {{
          filterChips.forEach(c => c.classList.remove('active'));
          chip.classList.add('active');
          currentFilter = chip.getAttribute('data-filter');
          filterCards();
        }});
      }});
    }});
  </script>
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
        # Default: scan First Sem 1st Year/Subjects
        subjects_dir = root_mlc / "First Sem 1st Year" / "Subjects"
        if subjects_dir.exists():
            scan_and_convert_directory(subjects_dir, overwrite=not args.no_overwrite)
        else:
            scan_and_convert_directory(root_mlc, overwrite=not args.no_overwrite)
        generate_study_hub_index(root_mlc)

if __name__ == "__main__":
    main()
