#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MLC Universal Text-to-Speech (Read Aloud) HTML Reader Generator
===============================================================
Converts legal study materials (.docx) across all MLC folders
into interactive, high-readability HTML web applications featuring:
  - Natural, Human-like Cadence & Intonation with Smart Legal Phonetic Expansions
  - Continuous, Flowing Body Paragraphs with Modern Legal Typography (No Chat-Like Fragmentation)
  - Streamlined, Clean Table of Contents (TOC) focused on Main Topics & Sub Topics (e.g. Case Titles Only)
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
    r'^(CASE\s+\d+[:\.]?.*)',
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
                f'<div class="case-citation-banner read-unit" data-unit-type="citation"><span class="citation-icon">⚖️</span> <span class="citation-text">{formatted_text}</span></div>'
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

def generate_reader_html(doc_title, subject_tag, sections, mp3_filename=None):
    """
    Renders the complete self-contained interactive reader HTML app.
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

    # Audio player snippet with Cloud Stream URL & Local Fallback
    mp3_player_html = ""
    if mp3_filename:
        encoded_mp3 = urllib.parse.quote(mp3_filename)
        cloud_stream_url = f"{GITHUB_AUDIO_BASE_URL}/{encoded_mp3}"
        mp3_player_html = f'''
        <div class="studio-audio-player">
          <div class="audio-player-header">
            <span class="audio-badge">🎙️ Studio Voice Podcast (Cloud Stream)</span>
            <span class="audio-filename">{html.escape(mp3_filename)}</span>
          </div>
          <audio controls preload="metadata" class="native-audio-element">
            <source src="{cloud_stream_url}" type="audio/mpeg">
            <source src="{html.escape(mp3_filename)}" type="audio/mpeg">
            Your browser does not support the audio element.
          </audio>
        </div>
        '''

    html_template = f'''<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{escaped_title} | MLC Law Library</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700;800;900&family=Inter:wght@400;500;600;700&family=Merriweather:ital,wght@0,300;0,400;0,700;1,300;1,400&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <style>
    :root {{
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
      --header-height: 64px;
      --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
      --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
    }}

    [data-theme="sepia"] {{
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
    }}

    [data-theme="light"] {{
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
    }}

    * {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      background-color: var(--bg-primary);
      color: var(--text-primary);
      font-family: var(--font-body);
      font-size: 17px;
      line-height: 1.8;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      transition: background-color 0.2s ease, color 0.2s ease;
    }}

    /* Scroll Progress Bar */
    #readingProgressBar {{
      position: fixed;
      top: 0;
      left: 0;
      height: 3px;
      background: linear-gradient(90deg, var(--accent-gold), var(--accent-blue));
      width: 0%;
      z-index: 9999;
      transition: width 0.1s ease;
    }}

    /* App Header */
    header.app-header {{
      height: var(--header-height);
      background-color: var(--bg-secondary);
      border-bottom: 1px solid var(--border-color);
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 1.25rem;
      position: sticky;
      top: 0;
      z-index: 100;
      backdrop-filter: blur(8px);
    }}

    .brand-section {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }}

    .brand-logo {{
      font-size: 1.4rem;
    }}

    .brand-info {{
      display: flex;
      flex-direction: column;
    }}

    .brand-title {{
      font-family: var(--font-heading);
      font-weight: 700;
      font-size: 0.95rem;
      color: var(--accent-gold);
      letter-spacing: 0.04em;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 320px;
    }}

    .brand-subject {{
      font-family: var(--font-ui);
      font-size: 0.72rem;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-weight: 600;
    }}

    /* TTS Toolbar */
    .tts-toolbar {{
      display: flex;
      align-items: center;
      gap: 0.6rem;
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

    .toc-link.level-2 {{
      padding-left: 1.5rem;
      font-size: 0.78rem;
      color: var(--text-muted);
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
      max-width: 900px;
      margin: 0 auto;
      padding: 2.5rem 2rem 5rem 2rem;
    }}

    /* Hero Card */
    .doc-hero-card {{
      background-color: var(--bg-secondary);
      border: 1px solid var(--border-color);
      border-radius: 12px;
      padding: 2rem;
      margin-bottom: 2.5rem;
      box-shadow: var(--shadow-md);
      position: relative;
      overflow: hidden;
    }}

    .doc-hero-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      width: 4px;
      height: 100%;
      background: linear-gradient(180deg, var(--accent-gold), var(--accent-blue));
    }}

    .doc-hero-subject {{
      font-family: var(--font-ui);
      font-size: 0.75rem;
      font-weight: 700;
      color: var(--accent-gold);
      text-transform: uppercase;
      letter-spacing: 0.08em;
      margin-bottom: 0.5rem;
    }}

    .doc-hero-title {{
      font-family: var(--font-heading);
      font-size: 1.85rem;
      font-weight: 800;
      color: var(--text-primary);
      line-height: 1.3;
      margin-bottom: 1rem;
    }}

    .doc-meta-stats {{
      display: flex;
      flex-wrap: wrap;
      gap: 1.25rem;
      font-family: var(--font-ui);
      font-size: 0.8rem;
      color: var(--text-muted);
    }}

    /* Studio Audio Player */
    .studio-audio-player {{
      margin-top: 1.5rem;
      background-color: var(--bg-primary);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      padding: 1rem;
    }}

    .audio-player-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 0.75rem;
      font-family: var(--font-ui);
      font-size: 0.8rem;
    }}

    .audio-badge {{
      font-weight: 700;
      color: var(--accent-emerald);
    }}

    .audio-filename {{
      color: var(--text-muted);
      font-family: var(--font-mono);
      font-size: 0.75rem;
    }}

    audio.native-audio-element {{
      width: 100%;
      height: 40px;
      outline: none;
    }}

    /* Case Section & Flowing Prose */
    section.doc-section {{
      margin-bottom: 3.5rem;
    }}

    .section-divider {{
      border: 0;
      height: 1px;
      background: var(--border-color);
      margin: 3.5rem 0;
    }}

    .case-header-card {{
      background: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      padding: 1.25rem 1.5rem;
      margin-bottom: 1.25rem;
      display: flex;
      align-items: center;
      gap: 1rem;
      box-shadow: var(--shadow-sm);
    }}

    .case-number-pill {{
      font-family: var(--font-ui);
      background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark));
      color: #0b1120;
      font-weight: 800;
      font-size: 0.78rem;
      padding: 0.35rem 0.75rem;
      border-radius: 20px;
      letter-spacing: 0.05em;
      white-space: nowrap;
    }}

    .case-header-title {{
      font-family: var(--font-heading);
      font-size: 1.35rem;
      font-weight: 700;
      color: var(--text-primary);
      line-height: 1.3;
    }}

    .topic-header-card {{
      background: var(--card-bg);
      border-left: 4px solid var(--accent-gold);
      border-radius: 8px;
      padding: 1rem 1.25rem;
      margin-bottom: 1.5rem;
    }}

    .topic-header-title {{
      font-family: var(--font-heading);
      font-size: 1.3rem;
      font-weight: 700;
      color: var(--accent-gold);
    }}

    .subtopic-header-card {{
      margin: 1.5rem 0 1rem 0;
      border-bottom: 1px solid var(--border-color);
      padding-bottom: 0.4rem;
    }}

    .subtopic-header-title {{
      font-family: var(--font-ui);
      font-size: 1.1rem;
      font-weight: 700;
      color: var(--accent-blue);
    }}

    .case-citation-banner {{
      background-color: var(--bg-tertiary);
      border-left: 3px solid var(--accent-blue);
      border-radius: 6px;
      padding: 0.6rem 1rem;
      font-family: var(--font-ui);
      font-size: 0.85rem;
      color: var(--text-secondary);
      margin-bottom: 1.5rem;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }}

    .case-subheading {{
      font-family: var(--font-ui);
      font-size: 1.05rem;
      font-weight: 700;
      color: var(--accent-gold);
      margin: 2rem 0 0.85rem 0;
      display: flex;
      align-items: center;
      gap: 0.4rem;
      letter-spacing: 0.02em;
    }}

    .subheading-accent {{
      color: var(--accent-gold);
      font-weight: 800;
    }}

    /* Continuous Flowing Body Paragraphs */
    .case-paragraph {{
      margin-bottom: 1.25rem;
      line-height: 1.85;
      color: var(--text-primary);
      text-align: justify;
      transition: background-color 0.2s ease, transform 0.15s ease;
      padding: 0.35rem 0.5rem;
      border-radius: 6px;
      cursor: pointer;
    }}

    .case-paragraph:hover {{
      background-color: rgba(255, 255, 255, 0.03);
    }}

    /* ALAC Inline Badges & Paragraphs */
    .alac-paragraph {{
      margin-bottom: 1.25rem;
      line-height: 1.85;
    }}

    .alac-badge {{
      display: inline-block;
      font-family: var(--font-ui);
      font-size: 0.72rem;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      padding: 0.2rem 0.55rem;
      border-radius: 5px;
      margin-right: 0.5rem;
      vertical-align: middle;
    }}

    .badge-ans {{ background-color: rgba(56, 189, 248, 0.18); color: var(--accent-blue); border: 1px solid rgba(56, 189, 248, 0.4); }}
    .badge-law {{ background-color: rgba(251, 191, 36, 0.18); color: var(--accent-gold); border: 1px solid rgba(251, 191, 36, 0.4); }}
    .badge-app {{ background-color: rgba(192, 132, 252, 0.18); color: var(--accent-purple); border: 1px solid rgba(192, 132, 252, 0.4); }}
    .badge-con {{ background-color: rgba(248, 113, 113, 0.18); color: var(--accent-crimson); border: 1px solid rgba(248, 113, 113, 0.4); }}
    .badge-syn {{ background-color: rgba(52, 211, 153, 0.18); color: var(--accent-emerald); border: 1px solid rgba(52, 211, 153, 0.4); }}
    .badge-gen {{ background-color: rgba(148, 163, 184, 0.18); color: var(--text-secondary); border: 1px solid rgba(148, 163, 184, 0.4); }}

    /* Bullet Points */
    .bullet-point {{
      display: flex;
      align-items: flex-start;
      gap: 0.6rem;
      margin-bottom: 0.85rem;
      padding-left: 0.5rem;
      line-height: 1.75;
    }}

    .bullet-dot {{
      color: var(--accent-gold);
      font-weight: 700;
    }}

    .bullet-content {{
      flex: 1;
    }}

    /* Tables */
    .table-responsive {{
      overflow-x: auto;
      margin: 1.5rem 0;
      border: 1px solid var(--border-color);
      border-radius: 8px;
    }}

    .reader-table {{
      width: 100%;
      border-collapse: collapse;
      font-family: var(--font-ui);
      font-size: 0.85rem;
      text-align: left;
    }}

    .reader-table th {{
      background-color: var(--bg-tertiary);
      color: var(--accent-gold);
      font-weight: 700;
      padding: 0.75rem 1rem;
      border-bottom: 1px solid var(--border-color);
    }}

    .reader-table td {{
      padding: 0.75rem 1rem;
      border-bottom: 1px solid var(--border-color);
      color: var(--text-secondary);
    }}

    .reader-table tr:last-child td {{
      border-bottom: none;
    }}

    /* TTS Active Unit Highlighting */
    .read-unit.is-speaking {{
      background-color: var(--highlight-bg) !important;
      outline: 2px solid var(--highlight-border);
      border-radius: 6px;
      box-shadow: 0 0 16px rgba(251, 191, 36, 0.3);
      transition: all 0.15s ease;
    }}

    /* Floating Selection Trigger */
    #floatingTtsTrigger {{
      display: none;
      position: absolute;
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
      <div class="brand-logo">⚖️</div>
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
        <input type="text" id="sidebarSearch" class="sidebar-search" placeholder="🔍 Filter Topics / Cases..." />
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
            <span>📑 {total_sections} Topics / Cases</span>
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

      // 5. Smart Legal Phonetic & Roman Numeral Expander
      function romanToInt(roman) {{
        const map = {{
          'M': 1000, 'CM': 900, 'D': 500, 'CD': 400,
          'C': 100, 'XC': 90, 'L': 50, 'XL': 40,
          'X': 10, 'IX': 9, 'V': 5, 'IV': 4, 'I': 1
        }};
        const s = roman.toUpperCase();
        let total = 0, i = 0;
        while (i < s.length) {{
          if (i + 1 < s.length && map[s.substr(i, 2)]) {{
            total += map[s.substr(i, 2)];
            i += 2;
          }} else if (map[s[i]]) {{
            total += map[s[i]];
            i += 1;
          }} else {{
            return null;
          }}
        }}
        return total;
      }}

      function expandRomanNumerals(txt) {{
        // 1. Topic Roman Numeral Headings: "I. FACTS", "XVI. CONCLUSION"
        txt = txt.replace(/\\b([IVXLCDM]+)\\.\\s+/g, function(match, roman) {{
          const val = romanToInt(roman);
          return val ? 'Topic ' + val + ': ' : match;
        }});

        // 2. Legal prefix + Roman: Canon II, Part I, Chapter XVI, Article XIV, Rule IV, Book I
        txt = txt.replace(/\\b(Canon|Part|Chapter|Article|Title|Book|Section|Sec|Art|Vol|Volume|Rule)\\s+([IVXLCDM]+)\\b/gi, function(match, prefix, roman) {{
          const val = romanToInt(roman);
          return val ? prefix + ' ' + val : match;
        }});

        // 3. Parenthesized Roman numerals: (i), (ii), (iv), (xvi)
        txt = txt.replace(/\\(([ivxlcdm]+)\\)/gi, function(match, roman) {{
          const val = romanToInt(roman);
          return val ? 'sub-item ' + val + ', ' : match;
        }});

        return txt;
      }}

      function prepareSpeechText(unitEl) {{
        let text = unitEl.innerText || unitEl.textContent || '';
        
        // Expand Roman Numerals first so "Canon II" becomes "Canon 2", "XVI" becomes "16"
        text = expandRomanNumerals(text);

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
          const headerEl = currentSec.querySelector('.case-header-title') || currentSec.querySelector('.topic-header-title') || currentSec.querySelector('.subtopic-header-title');
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
          const utterance = new SpeechSynthesisUtterance(expandRomanNumerals(sel));
          utterance.rate = parseFloat(speedSelect.value) || 0.9;
          const selVoiceIdx = voiceSelect.value;
          if (selVoiceIdx !== 'default' && voices[selVoiceIdx]) {{
            utterance.voice = voices[selVoiceIdx];
          }}
          synth.speak(utterance);
          floatBtn.style.display = 'none';
        }}
      }});

      // Keyboard Shortcuts
      document.addEventListener('keydown', (e) => {{
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
        if (e.code === 'Space') {{
          e.preventDefault();
          playPauseBtn.click();
        }} else if (e.key === 'Escape') {{
          stopBtn.click();
        }} else if (e.key === 'n' || e.key === 'N') {{
          nextBtn.click();
        }} else if (e.key === 'p' || e.key === 'P') {{
          prevBtn.click();
        }}
      }});
    }})();
  </script>
</body>
</html>
'''
    return html_template

# ==============================================================================
# 4. CONVERSION & HUB SCANNING
# ==============================================================================

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
    
    # We strictly search inside First Sem 1st Year/Subjects
    subjects_dir = root_path / "First Sem 1st Year" / "Subjects"
    if not subjects_dir.exists():
        subjects_dir = root_path / "Subjects"
    
    html_files = list(subjects_dir.rglob("*.html")) if subjects_dir.exists() else []

    # Map each subject group cleanly
    by_subject = {
        "Basic Legal and Judiciary Ethics": [],
        "Criminal Law": [],
        "Statutory Construction": []
    }

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

    groups_html = []
    subject_icons = {
        "Basic Legal and Judiciary Ethics": "⚖️",
        "Criminal Law": "🏛️",
        "Statutory Construction": "📜"
    }

    for group_name, files in by_subject.items():
        if not files:
            continue
        icon = subject_icons.get(group_name, "📚")
        cards_html = []
        for fpath, rel_str in files:
            doc_name = fpath.stem.replace('_', ' ')
            mp3_file = fpath.with_suffix('.mp3')
            has_mp3 = mp3_file.exists()
            mp3_badge = '<span class="badge-audio">🎙️ MP3 Audio</span>' if has_mp3 else ''
            is_digest = 'digest' in doc_name.lower()
            type_badge = '<span class="badge-reader">Case Digest</span>' if is_digest else '<span class="badge-outline">Course Outline</span>'
            
            cards_html.append(f'''
              <div class="hub-card">
                <div class="hub-card-header">
                  {type_badge}
                  {mp3_badge}
                </div>
                <h3 class="hub-card-title"><a href="{rel_str}">{html.escape(doc_name)}</a></h3>
                <div class="hub-card-actions">
                  <a href="{rel_str}" class="btn-open">Open Reader →</a>
                </div>
              </div>
            ''')

        groups_html.append(f'''
          <div class="hub-group">
            <h2 class="hub-group-title">{icon} {html.escape(group_name)}</h2>
            <div class="hub-grid">
              {''.join(cards_html)}
            </div>
          </div>
        ''')

    hub_page = f'''<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>MLC Law Library & Interactive Audio Suite</title>
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
      --accent-gold-dark: #d97706;
      --accent-blue: #38bdf8;
      --accent-emerald: #34d399;
      --font-ui: 'Inter', -apple-system, sans-serif;
      --font-heading: 'Cinzel', serif;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      background-color: var(--bg-primary);
      color: var(--text-primary);
      font-family: var(--font-ui);
      padding: 3.5rem 1.5rem;
      min-height: 100vh;
    }}
    .hub-container {{ max-width: 1100px; margin: 0 auto; }}
    .hub-header {{ text-align: center; margin-bottom: 3.5rem; }}
    .hub-badge {{ display: inline-block; background: rgba(251, 191, 36, 0.15); color: var(--accent-gold); font-size: 0.78rem; font-weight: 700; padding: 0.3rem 0.8rem; border-radius: 20px; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 1rem; border: 1px solid rgba(251, 191, 36, 0.3); }}
    .hub-title {{ font-family: var(--font-heading); font-size: 2.3rem; font-weight: 800; color: var(--accent-gold); margin-bottom: 0.75rem; letter-spacing: 0.03em; }}
    .hub-subtitle {{ color: var(--text-muted); font-size: 1.05rem; max-width: 650px; margin: 0 auto; line-height: 1.6; }}
    .hub-group {{ margin-bottom: 3rem; }}
    .hub-group-title {{ font-size: 1.3rem; font-weight: 700; color: var(--text-primary); margin-bottom: 1.25rem; border-bottom: 1px solid var(--border-color); padding-bottom: 0.6rem; display: flex; align-items: center; gap: 0.5rem; }}
    .hub-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 1.25rem; }}
    .hub-card {{ background-color: var(--bg-secondary); border: 1px solid var(--border-color); border-radius: 12px; padding: 1.5rem; display: flex; flex-direction: column; justify-content: space-between; transition: all 0.2s ease; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2); }}
    .hub-card:hover {{ transform: translateY(-3px); border-color: var(--accent-blue); box-shadow: 0 10px 25px rgba(0, 0, 0, 0.4); }}
    .hub-card-header {{ display: flex; gap: 0.5rem; margin-bottom: 0.85rem; }}
    .badge-reader {{ background: rgba(56, 189, 248, 0.15); color: var(--accent-blue); font-size: 0.72rem; font-weight: 700; padding: 0.25rem 0.6rem; border-radius: 6px; border: 1px solid rgba(56, 189, 248, 0.3); }}
    .badge-outline {{ background: rgba(192, 132, 252, 0.15); color: #c084fc; font-size: 0.72rem; font-weight: 700; padding: 0.25rem 0.6rem; border-radius: 6px; border: 1px solid rgba(192, 132, 252, 0.3); }}
    .badge-audio {{ background: rgba(52, 211, 153, 0.15); color: #34d399; font-size: 0.72rem; font-weight: 700; padding: 0.25rem 0.6rem; border-radius: 6px; border: 1px solid rgba(52, 211, 153, 0.3); }}
    .hub-card-title {{ font-size: 1.05rem; font-weight: 600; line-height: 1.45; margin-bottom: 1.5rem; }}
    .hub-card-title a {{ color: var(--text-primary); text-decoration: none; transition: color 0.15s ease; }}
    .hub-card-title a:hover {{ color: var(--accent-gold); }}
    .hub-card-actions {{ margin-top: auto; }}
    .btn-open {{ display: inline-block; background: linear-gradient(135deg, var(--accent-gold), var(--accent-gold-dark)); color: #0b1120; font-weight: 700; font-size: 0.85rem; padding: 0.55rem 1.1rem; border-radius: 8px; text-decoration: none; text-align: center; transition: filter 0.15s ease, transform 0.15s ease; }}
    .btn-open:hover {{ filter: brightness(1.1); transform: scale(1.02); }}
    footer {{ text-align: center; margin-top: 4rem; color: var(--text-muted); font-size: 0.85rem; border-top: 1px solid var(--border-color); padding-top: 2rem; }}
  </style>
</head>
<body>
  <div class="hub-container">
    <header class="hub-header">
      <div class="hub-badge">Manila Law College • Juris Doctor Program</div>
      <h1 class="hub-title">MLC Law Library & Audio Hub</h1>
      <p class="hub-subtitle">Interactive Full-Text Readers with Natural Voice Synthesis, ALAC Reasoning, and Studio Podcasts</p>
    </header>
    {''.join(groups_html)}
    <footer>
      <div>Manila Law College (MLC) • Academic Year 2026–2027 • First Semester Subjects</div>
    </footer>
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
        # Default: scan First Sem 1st Year/Subjects
        subjects_dir = root_mlc / "First Sem 1st Year" / "Subjects"
        if subjects_dir.exists():
            scan_and_convert_directory(subjects_dir, overwrite=not args.no_overwrite)
        else:
            scan_and_convert_directory(root_mlc, overwrite=not args.no_overwrite)
        generate_study_hub_index(root_mlc)

if __name__ == "__main__":
    main()
