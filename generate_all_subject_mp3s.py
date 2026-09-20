#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MLC Studio Neural MP3 Generator
===============================
Generates high-fidelity studio voice MP3 podcasts for all MLC legal documents
using Edge TTS Neural voice synthesis with:
  - Master Unified Table (Citations + Cyber & Digital Laws: SKRA, JEE-AR, SIP-ruh, DEE-PEE-AY, etc.)
  - Roman Numeral Cardinal & Ordinal Pronunciation
  - Topic / Case Change Voice Announcements ("Now Reading: CASE X: ...")
  - Smooth Intonation Pauses at Colons (: ...), Semicolons (;,), and Em-Dashes (— ...)
"""

import asyncio
import os
import sys
from pathlib import Path
import docx
import edge_tts
import re

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
    return text.strip()

def expand_roman_and_phonetics(text):
    # 1. Master Unified Table: Citations + Cyber & Digital Laws
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
        (r'\bPhil\.\b', 'Phil'),
        (r'\bSCRA\b', 'SKRA'),
        (r'\bCPRA\b', 'SIP-ruh'),
        (r'\bCPR\b', 'Code of Professional Responsibility'),
        (r'\bDPA\b', 'DEE-PEE-AY'),
        (r'\bITA\b', 'EYE-tuh'),
        (r'\bOSAEC\b', 'OH-sak'),
        (r'\bAFASA\b', 'ah-FAH-suh'),
        (r'\bCPA\b', 'Cybercrime Prevention Act'),
        (r'\bREED\b', 'REED'),
        (r'\bDICT\b', 'DIK-tee'),
        (r'\bCICC\b', 'SIK-see'),
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
        (r'\s+v(?:s)?\.\s+', ' versus '),
    ]

    for pattern, rep in replacements_first:
        text = re.sub(pattern, rep, text, flags=re.IGNORECASE)

    # 2. Topic Roman Numeral Headings: "(?:^|\n|;\s*|\.\s+)([IVXLCDM]+)\.\s+"
    def rep_topic_roman(m):
        prefix = m.group(1)
        roman = m.group(2)
        val = roman_to_int(roman)
        return f"{prefix}Topic {val}: " if val else m.group(0)

    text = re.sub(r'(^|\n|\.\s+|;\s+)([IVXLCDM]+)\.\s+', rep_topic_roman, text)

    # 3. Person names with ordinals: Macario Ramos II, King Henry VIII
    def rep_name_ordinal(m):
        first_word = m.group(1)
        roman = m.group(2).upper()
        if re.match(r'^(?:Canon|Part|Chapter|Article|Title|Book|Section|Sec|Art|Vol|Volume|Rule|Topic|Case|No|Nos)$', first_word, re.IGNORECASE):
            return m.group(0)
        return f"{first_word} {ORDINALS.get(roman, roman)}"

    text = re.sub(r'\b([A-Z][a-z]+)\s+([IVXLCDM]{1,4})\b', rep_name_ordinal, text)

    # 4. Legal prefix + Roman: Canon II, Part I, Chapter XVI, Article XIV, Rule IV, Book I
    text = re.sub(r'\b(Canon|Part|Chapter|Article|Title|Book|Section|Sec|Art|Vol|Volume|Rule)\s+([IVXLCDM]+)\b', lambda m: f"{m.group(1)} {roman_to_int(m.group(2))}" if roman_to_int(m.group(2)) else m.group(0), text, flags=re.IGNORECASE)

    # 5. Parenthesized Roman numerals: (i), (ii), (iv), (xvi)
    text = re.sub(r'\(([ivxlcdm]+)\)', lambda m: f"sub-item {roman_to_int(m.group(1))}, " if roman_to_int(m.group(1)) else m.group(0), text, flags=re.IGNORECASE)

    # 6. Natural Intonation & Pauses (Commas, Colons, Semicolons)
    text = re.sub(r':\s*', ': ... ', text)
    text = re.sub(r';\s*', ';, ', text)
    text = re.sub(r'\s*—\s*', ' — ... ', text)

    return text.strip()

def extract_clean_text(docx_path):
    doc = docx.Document(docx_path)
    lines = []
    for p in doc.paragraphs:
        t = clean_text(p.text)
        if t:
            # Announce case titles and main topics with transition cue
            if re.match(r'^(CASE\s+\d+|PART\s+[IVX\d]+|Canon\s+[IVX\d]+)', t, re.IGNORECASE):
                lines.append(f"Now Reading: {t}. ... ... ")
            else:
                lines.append(t)
    joined = "\n\n".join(lines)
    return expand_roman_and_phonetics(joined)

async def generate_mp3_for_file(docx_path, voice="en-US-JennyNeural", rate="-3%"):
    path_obj = Path(docx_path).resolve()
    output_mp3 = path_obj.with_suffix('.mp3')
    
    print(f"\n[AUDIO ENGINE] Reading: {path_obj.name}...")
    full_text = extract_clean_text(path_obj)
    
    char_count = len(full_text)
    word_count = len(full_text.split())
    print(f"  -> Extracted {word_count:,} words ({char_count:,} characters)")
    print(f"  -> Synthesizing Neural Audio ({voice}) -> {output_mp3.name}...")
    
    chunk_size = 25000
    if len(full_text) <= chunk_size:
        communicate = edge_tts.Communicate(full_text, voice, rate=rate)
        await communicate.save(str(output_mp3))
    else:
        paragraphs = full_text.split("\n\n")
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
            
        print(f"  -> Large document split into {len(chunks)} synthesis batches...")
        temp_files = []
        for i, chunk in enumerate(chunks):
            temp_mp3 = path_obj.parent / f"temp_{path_obj.stem}_{i}.mp3"
            print(f"     Synthesizing part {i+1}/{len(chunks)} ({len(chunk):,} chars)...")
            comm = edge_tts.Communicate(chunk, voice, rate=rate)
            await comm.save(str(temp_mp3))
            temp_files.append(temp_mp3)
            
        # Combine parts
        with open(output_mp3, "wb") as outfile:
            for tf in temp_files:
                with open(tf, "rb") as infile:
                    outfile.write(infile.read())
                tf.unlink()

    mb_size = output_mp3.stat().st_size / (1024 * 1024)
    print(f"[SUCCESS] Audio Saved: {output_mp3.name} ({mb_size:.2f} MB)")
    return output_mp3

async def main():
    root_dir = Path(r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Subjects")
    if not root_dir.exists():
        root_dir = Path(r"C:\Users\JR\Downloads\14All-All41\Subjects")
        
    docx_files = [p for p in root_dir.rglob("*.docx") if not p.name.startswith(("~$", ".~"))]
    print(f"=== Found {len(docx_files)} Word Documents for Neural Speech Synthesis ===")
    for df in docx_files:
        await generate_mp3_for_file(df)

if __name__ == "__main__":
    asyncio.run(main())
