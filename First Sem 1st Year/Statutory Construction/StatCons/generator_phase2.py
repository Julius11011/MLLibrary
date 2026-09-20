# generator_phase2.py
"""
Phase 2 Generator: Chapter III (Cases 22 to 49)
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

from phase2_cases_part1 import cases_p2_part1
from phase2_cases_part2 import cases_p2_part2
from phase2_cases_part3 import cases_p2_part3

all_p2_cases = cases_p2_part1 + cases_p2_part2 + cases_p2_part3
all_p2_cases.sort(key=lambda x: x['num'])
print(f"Total Unique Phase 2 Cases: {len(all_p2_cases)}")

def build_phase2_documents():
    txt_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase2_Chapter_III.txt"
    docx_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase2_Chapter_III.docx"
    
    # 1. Build Text File
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("MANILA LAW COLLEGE\n")
        f.write("Subject: Statutory Construction (2 units)\n")
        f.write("Instructor: Atty. Jason P. Raposas\n")
        f.write("Reference Syllabus: StatCon_Syllabus_JPR.pdf\n\n")
        f.write("=" * 95 + "\n")
        f.write("PHASE 2: CHAPTER III CASE DIGESTS & COMPREHENSIVE STATCON ANALYSIS\n")
        f.write("CHAPTER III: BASIC GUIDELINES IN THE CONSTRUCTION AND INTERPRETATION OF LAW (Cases 22 to 49)\n")
        f.write("(Dual Layer: Full ALAC + Recitation Scripts + 12-Section Case Digest Format)\n")
        f.write("=" * 95 + "\n\n")
        
        f.write("\n" + "#" * 95 + "\n")
        f.write("CHAPTER III: BASIC GUIDELINES IN THE CONSTRUCTION AND INTERPRETATION OF LAW\n")
        f.write("Syllabus Coverage: Legislative Intent, Verba Intentioni, Nemo Tenetur Ad Impossibile, Avoid Absurdity, Whole Statute, Ubi Lex Non Distinguit, Ejusdem Generis, Noscitur A Sociis, Expressio Unius, Casus Omissus, Generalia Specialibus, Lex Posterior, Lex Specialis, Borrowed Statute, In Pari Materia\n")
        f.write("#" * 95 + "\n\n")
        
        for c in all_p2_cases:
            f.write(render_case_to_txt(c))
            
    print(f"Successfully generated Phase 2 TXT file: {txt_path}")

    # 2. Build DOCX File
    doc = create_docx_builder()
    add_header_box(
        doc,
        title="STATUTORY CONSTRUCTION COMPENDIUM: PHASE 2 (CHAPTER III)",
        subtitle="Chapter III: Basic Guidelines in Construction (Cases 22 to 49) | Dual Layer ALAC + 12-Section Case Digest Format",
        course_info="Manila Law College | Statutory Construction (2 Units) | Instructor: Atty. Jason P. Raposas"
    )
    
    for idx, c in enumerate(all_p2_cases):
        doc.add_page_break()
        
        if idx == 0:
            add_chapter_banner(
                doc,
                chapter_num="III",
                chapter_title="Basic Guidelines in the Construction and Interpretation of Law",
                chapter_subtitle="Cases 22 to 49 | General Principles, Considerations Within & Outside the Statute",
                syllabus_topics=[
                    "A. General Guidelines: Legislative Intent Must Prevail; Verba Intentioni, Non E Contra, Debent Inservire; Purposeful Construction; Nemo Tenetur Ad Impossibile; Avoid Absurdity and Injustice; Whole Statute; Every Part Given Effect",
                    "B. Considerations Within the Statute: Ubi Lex Non Distinguit Nec Nos Distinguere Debemos; Ejusdem Generis; Noscitur A Sociis; Doctrine of Necessary Implication; Expressio Unius Est Exclusio Alterius; Reddendo Singula Singulis; Casus Omissus Pro Omisso Habendus Est",
                    "C. Considerations Outside the Statute: Generalia Specialibus Non Derogat; Lex Posterior Derogat Priori; Lex Specialis Derogat Legi Generalis; Borrowed Statute Rule; Interpretare Et Concordare Leges Legibus Est Optimus Interpretandi"
                ]
            )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(8)
            
        render_case_to_docx(doc, c)

    doc.save(docx_path)
    print(f"Successfully generated Phase 2 DOCX file: {docx_path}")

if __name__ == "__main__":
    build_phase2_documents()
