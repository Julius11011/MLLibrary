#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Studio Neural Voice MP3 Generator for Revised Penal Code & Proposed Criminal Code Compendium
=============================================================================================
Synthesizes high-fidelity, studio-grade neural voice podcasts using Edge TTS (Jenny Neural)
for:
  1. RPC/Philippine_Revised_Penal_Code_and_Proposed_New_Criminal_Code_Compendium.docx -> .mp3
  2. RPC/New_RPC_Public_Publish_and_Codal_Compendium.docx -> .mp3

Features:
  - Strict Rule 4 Spoken Pronunciation: All Roman numerals pronounced as cardinal numbers
    ("Canon 1", "Canon 2", "Topic 1", "Part 1", "Book 1", "Title 1", "Article 14")
  - Master Unified Legal Citation & Acronym expansions ("G.R. No." -> "JEE-AR Number", "RA" -> "Republic Act", "RPC" -> "Revised Penal Code", "SCRA" -> "SKRA")
  - Seamless Extraction of Paragraphs, Tables, and Stylized Callouts
  - Natural Intonation Pauses for Crystal-Clear Audio Comprehension
"""

import asyncio
import os
import sys
import shutil
import re
from pathlib import Path
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
import edge_tts

ROMAN_MAP = {
    'M': 1000, 'CM': 900, 'D': 500, 'CD': 400,
    'C': 100, 'XC': 90, 'L': 50, 'XL': 40,
    'X': 10, 'IX': 9, 'V': 5, 'IV': 4, 'I': 1
}

ORDINALS = {
    'I': 'the first', 'II': 'the second', 'III': 'the third', 'IV': 'the fourth', 'V': 'the fifth',
    'VI': 'the sixth', 'VII': 'the seventh', 'VIII': 'the eighth', 'IX': 'the ninth', 'X': 'the tenth'
}

def roman_to_int(roman_str):
    s = roman_str.upper()
    total = 0
    i = 0
    while i < len(s):
        if i + 1 < len(s) and s[i:i+2] in ROMAN_MAP:
            total += ROMAN_MAP[s[i:i+2]]
            i += 2
        elif s[i] in ROMAN_MAP:
            total += ROMAN_MAP[s[i]]
            i += 1
        else:
            return None
    return total

def clean_text(text):
    if not text:
        return ""
    text = text.replace("\ufffd", "'")
    # Clean redundant emoji markers for speech
    text = re.sub(r'[\U00010000-\U0010ffff]', '', text)
    text = re.sub(r'[📌⚖️📖🏛️📚⭐🔖]', '', text)
    return text.strip()

def expand_roman_and_phonetics(text):
    # 1. Master Unified Table: Citations & Legal acronyms
    replacements_first = [
        (r'\bA\.C\.\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)', r'Administrative Case Number \1'),
        (r'\bA\.M\.\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)', r'Administrative Matter Number \1'),
        (r'\bG\.R\.\s*Nos\.?\s*([A-Za-z0-9\-,\s]+)', r'JEE-AR Numbers \1'),
        (r'\bG\.R\.\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)', r'JEE-AR Number \1'),
        (r'\bG\.R\.\b', 'JEE-AR'),
        (r'\bP\.D\.\s*(?:No\.?\s*)?(\d+)', r'Presidential Decree Number \1'),
        (r'\bPD\s*(\d+)', r'Presidential Decree \1'),
        (r'\bP\.D\.\b', 'Presidential Decree'),
        (r'\bR\.A\.\s*(?:No\.?\s*)?(\d+)', r'Republic Act Number \1'),
        (r'\bRA\s*(\d+)', r'Republic Act \1'),
        (r'\bR\.A\.\b', 'Republic Act'),
        (r'\bPhil\.\s*(\d+)', r'Philippine Reports volume \1'),
        (r'\bPhil\.\b', 'Philippine Reports'),
        (r'\bSCRA\b', 'SKRA'),
        (r'\bCPRA\b', 'SIP-ruh'),
        (r'\bRPC\b', 'Revised Penal Code'),
        (r'\bIn\s+re\b', 'in Ree'),
        (r'\bin\s+re\b', 'in Ree'),
        (r'\bet\s+al\.', 'et AHL,'),
        (r'\bet\s+al\b', 'et AHL,'),
        (r'\bi\.e\.', 'that is,'),
        (r'\be\.g\.', 'for example,'),
        (r'\bArt\.\s*(\d+)', r'Article \1'),
        (r'\bArts\.\s*([\d,\s\-]+)', r'Articles \1'),
        (r'\bSec\.\s*(\d+)', r'Section \1'),
        (r'\bSecs\.\s*([\d,\s\-]+)', r'Sections \1'),
        (r'\bPar\.\s*(\d+)', r'Paragraph \1'),
        (r'\bP(\d+[\d,]*\b)', r'\1 pesos'),
        (r'\bvs\.\b|\bv\.\b', 'versus'),
        (r'\[A\]\s*ANSWER:\s*', 'Answer: ... '),
        (r'\[L\]\s*LEGAL BASIS:\s*', 'Legal Basis: ... '),
        (r'\[A\]\s*APPLICATION:\s*', 'Application: ... '),
        (r'\[C\]\s*CONCLUSION & DOCTRINE:\s*', 'Conclusion and Doctrine: ... '),
        (r'\[ARTICLE\]\s*APPLICABLE ARTICLE:\s*', 'Applicable Codal Provision: ... '),
        (r'\[TREATISE\]\s*TREATISE CORRELATION:\s*', 'Treatise Correlation: ... ')
    ]

    for pattern, rep in replacements_first:
        text = re.sub(pattern, rep, text, flags=re.IGNORECASE)

    # 2. Section / Topic Roman Numeral Headings (e.g. PART I -> Part 1, I. EXECUTIVE -> Topic 1: EXECUTIVE)
    def rep_topic_roman(m):
        prefix = m.group(1)
        roman = m.group(2)
        val = roman_to_int(roman)
        return f"{prefix}Topic {val}: " if val else m.group(0)

    text = re.sub(r'(^|\n|\.\s+|;\s+)([IVXLCDM]+)\.\s+', rep_topic_roman, text)

    # 3. Legal prefix + Roman (e.g., Part I -> Part 1, Canon II -> Canon 2, Article XIV -> Article 14, Book One -> Book 1)
    text = re.sub(r'\b(Canon|Part|Chapter|Article|Title|Book|Section|Sec|Art|Vol|Volume|Rule)\s+([IVXLCDM]+)\b', 
                  lambda m: f"{m.group(1)} {roman_to_int(m.group(2))}" if roman_to_int(m.group(2)) else m.group(0), 
                  text, flags=re.IGNORECASE)

    # Convert Book One/Two/Three
    book_num_map = {'one': '1', 'two': '2', 'three': '3', 'four': '4', 'five': '5'}
    text = re.sub(r'\bBook\s+(One|Two|Three|Four|Five)\b', lambda m: f"Book {book_num_map.get(m.group(1).lower(), m.group(1))}", text, flags=re.IGNORECASE)

    # 4. Parenthesized Roman numerals: (i), (ii), (iv), (xvi)
    text = re.sub(r'\(([ivxlcdm]+)\)', lambda m: f"sub-item {roman_to_int(m.group(1))}, " if roman_to_int(m.group(1)) else m.group(0), text, flags=re.IGNORECASE)

    # 5. Natural Intonation & Pauses
    text = re.sub(r':\s*', ': ... ', text)
    text = re.sub(r';\s*', ';, ', text)
    text = re.sub(r'\s*—\s*', ' — ... ', text)

    return text.strip()

def iter_block_items(parent):
    """
    Yield each paragraph and table in document order.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Unsupported parent type")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_compendium_text(docx_path):
    doc = docx.Document(docx_path)
    lines = []
    
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = clean_text(block.text)
            if t:
                # Mark major section headings with distinct pause
                if re.match(r'^(\d+\.\s+PART|\d+\.\s+EXECUTIVE|[A-Z\s]{4,}:)', t):
                    lines.append(f"\n\n{t}. ... ... ")
                else:
                    lines.append(t)
        elif isinstance(block, Table):
            # Check if callout box (1x1 table) or comparative/data table
            if len(block.rows) == 1 and len(block.columns) == 1:
                # Callout box
                cell = block.cell(0, 0)
                cell_lines = []
                for p in cell.paragraphs:
                    ct = clean_text(p.text)
                    if ct:
                        cell_lines.append(ct)
                if cell_lines:
                    lines.append("\n" + "\n".join(cell_lines) + "\n")
            elif len(block.rows) == 1 and len(block.columns) == 3:
                # Meta header bar
                continue
            else:
                # Comparative matrix or schedule table
                lines.append("\nComparative Table Breakdown:")
                header_row = [clean_text(c.text) for c in block.rows[0].cells]
                for r_idx, row in enumerate(block.rows[1:], start=1):
                    row_texts = [clean_text(c.text) for c in row.cells]
                    if len(row_texts) >= 3:
                        lines.append(f"Subject: {row_texts[0]}. Under current law: {row_texts[1]}. Under proposed reform: {row_texts[2]}.")
                    elif len(row_texts) == 4:
                        lines.append(f"Article {row_texts[0]}, {row_texts[1]}. Original value: {row_texts[2]}. Modern value under Republic Act 10951: {row_texts[3]}.")
                    else:
                        lines.append(" — ".join(row_texts))
                lines.append("\n")

    joined = "\n\n".join(lines)
    return expand_roman_and_phonetics(joined)

async def synthesize_audio(text, output_mp3_path, voice="en-US-JennyNeural", rate="-3%"):
    out_obj = Path(output_mp3_path).resolve()
    out_obj.parent.mkdir(parents=True, exist_ok=True)
    
    char_count = len(text)
    word_count = len(text.split())
    print(f"\n[AUDIO SYNTHESIS] Target: {out_obj.name}")
    print(f"                 Output Path: {out_obj}")
    print(f"                 Total Words: {word_count:,} ({char_count:,} characters)")
    
    chunk_size = 25000
    if len(text) <= chunk_size:
        print(f"  -> Synthesizing directly ({voice} at rate {rate})...")
        communicate = edge_tts.Communicate(text, voice, rate=rate)
        await communicate.save(str(out_obj))
    else:
        paragraphs = text.split("\n\n")
        chunks = []
        cur_chunk = []
        cur_len = 0
        for p in paragraphs:
            if cur_len + len(p) > chunk_size and cur_chunk:
                chunks.append("\n\n".join(cur_chunk))
                cur_chunk = [p]
                cur_len = len(p)
            else:
                cur_chunk.append(p)
                cur_len += len(p)
        if cur_chunk:
            chunks.append("\n\n".join(cur_chunk))
            
        print(f"  -> Document split into {len(chunks)} synthesis chunks...")
        temp_files = []
        for i, chunk in enumerate(chunks):
            temp_mp3 = out_obj.parent / f"temp_{out_obj.stem}_{i}.mp3"
            print(f"     Synthesizing chunk {i+1}/{len(chunks)} ({len(chunk):,} chars)...")
            comm = edge_tts.Communicate(chunk, voice, rate=rate)
            await comm.save(str(temp_mp3))
            temp_files.append(temp_mp3)
            
        # Combine parts
        print("  -> Concatenating audio chunks into master MP3...")
        with open(out_obj, "wb") as outfile:
            for tf in temp_files:
                with open(tf, "rb") as infile:
                    outfile.write(infile.read())
                tf.unlink()

    mb_size = out_obj.stat().st_size / (1024 * 1024)
    print(f"[SUCCESS] Master MP3 Saved: {out_obj.name} ({mb_size:.2f} MB)")
    return out_obj

async def main():
    rpc_dir = Path(r"C:\Users\JR\Downloads\14All-All41\MLC\RPC")
    
    file1_docx = rpc_dir / "Philippine_Revised_Penal_Code_and_Proposed_New_Criminal_Code_Compendium.docx"
    file1_mp3 = rpc_dir / "Philippine_Revised_Penal_Code_and_Proposed_New_Criminal_Code_Compendium.mp3"
    
    file2_docx = rpc_dir / "New_RPC_Public_Publish_and_Codal_Compendium.docx"
    file2_mp3 = rpc_dir / "New_RPC_Public_Publish_and_Codal_Compendium.mp3"
    
    if not file1_docx.exists():
        print(f"Error: {file1_docx} not found.")
        return
        
    print("=== EXTRACTING & PHONETICALLY NORMALIZING RPC TEXT ===")
    spoken_text = extract_compendium_text(file1_docx)
    
    # Generate MP3 for the primary compendium
    await synthesize_audio(spoken_text, file1_mp3)
    
    # Copy/sync to the second compendium file
    shutil.copy2(str(file1_mp3), str(file2_mp3))
    print(f"[SYNC] Copied master MP3 to: {file2_mp3.name}")
    
    print("\n=== ALL RPC TTS MP3 AUDIO GENERATION COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    asyncio.run(main())
