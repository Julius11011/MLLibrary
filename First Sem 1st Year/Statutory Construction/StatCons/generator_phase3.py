# generator_phase3.py
"""
Phase 3 Generator: Chapter IV (Cases 50 to 69)
Dual Layer: Full ALAC & Recitation Guides + 12-Section Case Digest Format.
"""

import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core_builder import (
    create_docx_builder, add_header_box, add_chapter_banner
)
from case_renderer import render_case_to_docx, render_case_to_txt

from phase3_cases_part1 import cases_p3_part1
from phase3_cases_part2 import cases_p3_part2

all_p3_cases = cases_p3_part1 + cases_p3_part2
all_p3_cases.sort(key=lambda x: x['num'])
print(f"Total Unique Phase 3 Cases: {len(all_p3_cases)}")

def build_phase3_documents():
    txt_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase3_Chapter_IV.txt"
    docx_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase3_Chapter_IV.docx"
    
    # 1. Build Text File
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("MANILA LAW COLLEGE\n")
        f.write("Subject: Statutory Construction (2 units)\n")
        f.write("Instructor: Atty. Jason P. Raposas\n")
        f.write("Reference Syllabus: StatCon_Syllabus_JPR.pdf\n\n")
        f.write("=" * 95 + "\n")
        f.write("PHASE 3: CHAPTER IV CASE DIGESTS & COMPREHENSIVE STATCON ANALYSIS\n")
        f.write("CHAPTER IV: AIDS TO CONSTRUCTION (Cases 50 to 69)\n")
        f.write("(Dual Layer: Full ALAC + Recitation Scripts + 12-Section Case Digest Format)\n")
        f.write("=" * 95 + "\n\n")
        
        f.write("\n" + "#" * 95 + "\n")
        f.write("CHAPTER IV: AIDS TO CONSTRUCTION\n")
        f.write("Syllabus Coverage: Intrinsic Aids (Preambles, Title, Punctuations, Capitalizations, Headnotes, Lingual Text); Extrinsic Aids (Policy, Purpose, Legislative History, Legal Environment, Contemporaneous Construction, Dictionaries); Presumptions (Against Unconstitutionality, Injustice, Implied Repeals, Ineffectiveness, Absurdity)\n")
        f.write("#" * 95 + "\n\n")
        
        for c in all_p3_cases:
            f.write(render_case_to_txt(c))
            
    print(f"Successfully generated Phase 3 TXT file: {txt_path}")

    # 2. Build DOCX File
    doc = create_docx_builder()
    add_header_box(
        doc,
        title="STATUTORY CONSTRUCTION COMPENDIUM: PHASE 3 (CHAPTER IV)",
        subtitle="Chapter IV: Aids to Construction (Cases 50 to 69) | Dual Layer ALAC + 12-Section Case Digest Format",
        course_info="Manila Law College | Statutory Construction (2 Units) | Instructor: Atty. Jason P. Raposas"
    )
    
    for idx, c in enumerate(all_p3_cases):
        doc.add_page_break()
        
        if idx == 0:
            add_chapter_banner(
                doc,
                chapter_num="IV",
                chapter_title="Aids to Construction",
                chapter_subtitle="Cases 50 to 69 | Intrinsic Aids, Extrinsic Aids & Statutory Presumptions",
                syllabus_topics=[
                    "A. Intrinsic Aids (Cases 50–59): Preambles, Title, Punctuation Marks, Capitalization, Headnotes/Epigraphs, Lingual Text and Language Reconciliations",
                    "B. Extrinsic Aids (Cases 60–65): Policy and Purpose of Law, Legislative History (Debates, Reports, Committee Journals), Legal Environment, Contemporaneous Construction (Executive/Administrative Interpretation), Dictionaries",
                    "C. Presumptions in Aid of Construction (Cases 66–69): Presumption Against Unconstitutionality, Presumption Against Injustice, Presumption Against Implied Repeals, Presumption Against Ineffectiveness, Presumption Against Absurdity"
                ]
            )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(8)
            
        render_case_to_docx(doc, c)

    doc.save(docx_path)
    print(f"Successfully generated Phase 3 DOCX file: {docx_path}")

if __name__ == "__main__":
    build_phase3_documents()
