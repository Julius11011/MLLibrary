# generator_phase1.py
"""
Phase 1 Generator: Chapters I & II (Cases 1 to 21)
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
from phase1_cases_complete import all_phase1_cases

all_p1_cases = all_phase1_cases
all_p1_cases.sort(key=lambda x: x['num'])
print(f"Total Unique Phase 1 Cases: {len(all_p1_cases)}")

def build_phase1_documents():
    txt_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase1_Chapters_I_II.txt"
    docx_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase1_Chapters_I_II.docx"
    
    # 1. Build Text File
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("MANILA LAW COLLEGE\n")
        f.write("Subject: Statutory Construction (2 units)\n")
        f.write("Instructor: Atty. Jason P. Raposas\n")
        f.write("Reference Syllabus: StatCon_Syllabus_JPR.pdf\n\n")
        f.write("=" * 95 + "\n")
        f.write("PHASE 1: CHAPTERS I & II CASE DIGESTS & COMPREHENSIVE STATCON ANALYSIS\n")
        f.write("CHAPTER I: BACKGROUND OF STATUTORY CONSTRUCTION (Cases 1 to 7)\n")
        f.write("CHAPTER II: PRELUDE TO THE EXERCISE OF THE POWER (Cases 8 to 21)\n")
        f.write("(Dual Layer: Full ALAC + Recitation Scripts + 12-Section Case Digest Format)\n")
        f.write("=" * 95 + "\n\n")
        
        current_ch = 0
        for c in all_p1_cases:
            if c['num'] == 1 and current_ch < 1:
                current_ch = 1
                f.write("\n" + "#" * 95 + "\n")
                f.write("CHAPTER I: BACKGROUND OF STATUTORY CONSTRUCTION\n")
                f.write("Syllabus Coverage: Statutes, Situs, Purpose, When to Construe, Ambiguity, Judicial Legislation, Legis Interpretatio, Strict/Liberal, Prospective/Retrospective, Limits of Power\n")
                f.write("#" * 95 + "\n\n")
            elif c['num'] == 8 and current_ch < 2:
                current_ch = 2
                f.write("\n" + "#" * 95 + "\n")
                f.write("CHAPTER II: PRELUDE TO THE EXERCISE OF THE POWER\n")
                f.write("Syllabus Coverage: Verba Legis Non Est Recedendum, Absoluta Sententia Expositore Non Indigent, Dura Lex Sed Lex\n")
                f.write("#" * 95 + "\n\n")
                
            f.write(render_case_to_txt(c))
            
    print(f"Successfully generated Phase 1 TXT file: {txt_path}")

    # 2. Build DOCX File
    doc = create_docx_builder()
    add_header_box(
        doc,
        title="STATUTORY CONSTRUCTION COMPENDIUM: PHASE 1 (CHAPTERS I & II)",
        subtitle="Chapters I & II (Cases 1 to 21) | Dual Layer ALAC + 12-Section Case Digest Format",
        course_info="Manila Law College | Statutory Construction (2 Units) | Instructor: Atty. Jason P. Raposas"
    )
    
    current_ch = 0
    for c in all_p1_cases:
        doc.add_page_break()
        
        if c['num'] == 1 and current_ch < 1:
            current_ch = 1
            add_chapter_banner(
                doc,
                chapter_num="I",
                chapter_title="Background of Statutory Construction",
                chapter_subtitle="Cases 1 to 7 | Foundations of Statutory Interpretation & Judicial Power",
                syllabus_topics=[
                    "A. Statutes: Definition, Nature, Essential Requisites, Classifications",
                    "B. Statutory Construction: Definition, Nature, Function and Scope",
                    "C. Situs of Construction: Judicial Power and Constitutional Prerogative (Art. VIII)",
                    "D. Purpose of Construction: Ascertaining and Effectuating Legislative Intent",
                    "E. When to Construe: Threshold Cardinal Rule of Ambiguity vs. Clarity",
                    "F. Ambiguity / Vagueness: Void-for-Vagueness Doctrine and Types of Ambiguity",
                    "G. Statutory Construction vs. Judicial Legislation (Jus Dicere vs. Jus Dare)",
                    "H. Legis Interpretatio Legis Vim Obtinet (Art. 8 Civil Code & Stare Decisis)",
                    "I. Kinds of Construction: Strict vs. Liberal, Prospective vs. Retrospective",
                    "J. Extent and Limits of Judicial Power to Construe"
                ]
            )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(8)
            
        elif c['num'] == 8 and current_ch < 2:
            current_ch = 2
            add_chapter_banner(
                doc,
                chapter_num="II",
                chapter_title="Prelude to the Exercise of the Power",
                chapter_subtitle="Cases 8 to 21 | Cardinal Threshold Maxims of Textual Fidelity",
                syllabus_topics=[
                    "A. Verba Legis Non Est Recedendum: From the words of the statute there must be no departure",
                    "B. Absoluta Sententia Expositore Non Indigent: When the language is clear, it needs no expositor",
                    "C. Dura Lex Sed Lex: The law may be harsh, but it is the law; equity cannot override positive statutory command",
                    "D. Plain Meaning Rule: Clear statutory text is the best exponent of legislative will"
                ]
            )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(8)
            
        render_case_to_docx(doc, c)

    doc.save(docx_path)
    print(f"Successfully generated Phase 1 DOCX file: {docx_path}")

if __name__ == "__main__":
    build_phase1_documents()
