# generator_phase4.py
"""
Phase 4 Generator: Chapters V & VI (Cases 70 to 103)
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

from phase4_cases_part1 import cases_p4_part1
from phase4_cases_part2 import cases_p4_part2
from phase4_cases_part3 import cases_p4_part3

all_p4_cases = cases_p4_part1 + cases_p4_part2 + cases_p4_part3
all_p4_cases.sort(key=lambda x: x['num'])
print(f"Total Unique Phase 4 Cases: {len(all_p4_cases)}")

def build_phase4_documents():
    txt_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase4_Chapters_V_VI.txt"
    docx_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Phase4_Chapters_V_VI.docx"
    
    # 1. Build Text File
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("MANILA LAW COLLEGE\n")
        f.write("Subject: Statutory Construction (2 units)\n")
        f.write("Instructor: Atty. Jason P. Raposas\n")
        f.write("Reference Syllabus: StatCon_Syllabus_JPR.pdf\n\n")
        f.write("=" * 95 + "\n")
        f.write("PHASE 4: CHAPTERS V & VI CASE DIGESTS & COMPREHENSIVE STATCON ANALYSIS\n")
        f.write("CHAPTER V: INTERPRETATION OF WORDS AND PHRASES (Cases 70 to 93)\n")
        f.write("CHAPTER VI: CONSTRUCTION OF CONSTITUTION AND SPECIFIC KINDS OF STATUTES (Cases 94 to 103)\n")
        f.write("(Dual Layer: Full ALAC + Recitation Scripts + 12-Section Case Digest Format)\n")
        f.write("=" * 95 + "\n\n")
        
        current_ch = 0
        for c in all_p4_cases:
            if c['num'] == 70 and current_ch < 5:
                current_ch = 5
                f.write("\n" + "#" * 95 + "\n")
                f.write("CHAPTER V: INTERPRETATION OF WORDS AND PHRASES\n")
                f.write("Syllabus Coverage: Generic Words, Commercial Meaning, Technical/Legal Terms, Disjunctive ('or') / Conjunctive ('and'), Mandatory ('shall') / Directory ('may'), Progressive Interpretation, Purpose-Qualified Meaning, Provisos, Exceptions, Saving Clauses, Whole Statute Harmonization\n")
                f.write("#" * 95 + "\n\n")
            elif c['num'] == 94 and current_ch < 6:
                current_ch = 6
                f.write("\n" + "#" * 95 + "\n")
                f.write("CHAPTER VI: CONSTRUCTION OF CONSTITUTION AND SPECIFIC KINDS OF STATUTES\n")
                f.write("Syllabus Coverage: Constitutional Construction (Verba Legis, Ratio Legis, Ut Magis Valeat), Statutes Strictly Construed (Penal, Tax, Derogation of Rights, Naturalization), Statutes Liberally Construed (Remedial, Social Justice, Labor, Pension, Election Laws)\n")
                f.write("#" * 95 + "\n\n")
                
            f.write(render_case_to_txt(c))
            
    print(f"Successfully generated Phase 4 TXT file: {txt_path}")

    # 2. Build DOCX File
    doc = create_docx_builder()
    add_header_box(
        doc,
        title="STATUTORY CONSTRUCTION COMPENDIUM: PHASE 4 (CHAPTERS V & VI)",
        subtitle="Chapters V & VI (Cases 70 to 103) | Dual Layer ALAC + 12-Section Case Digest Format",
        course_info="Manila Law College | Statutory Construction (2 Units) | Instructor: Atty. Jason P. Raposas"
    )
    
    current_ch = 0
    for c in all_p4_cases:
        doc.add_page_break()
        
        if c['num'] == 70 and current_ch < 5:
            current_ch = 5
            add_chapter_banner(
                doc,
                chapter_num="V",
                chapter_title="Interpretation of Words and Phrases",
                chapter_subtitle="Cases 70 to 93 | Semantic Rules, Specific Connectives & Textual Modifiers",
                syllabus_topics=[
                    "A. General Words and Phrases: Generic Words, Commercial Meaning, Technical and Legal Terms",
                    "B. Specific Words: Disjunctive ('or') vs. Conjunctive ('and'); Mandatory ('shall') vs. Directory ('may')",
                    "C. General Words Construed Generally; Progressive Interpretation Doctrine",
                    "D. Meaning of Words Qualified by Purpose of Statute",
                    "E. Provisos, Exceptions, and Saving Clauses",
                    "F. Statute Construed as a Whole and in Relation to Related Statutes"
                ]
            )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(8)
            
        elif c['num'] == 94 and current_ch < 6:
            current_ch = 6
            add_chapter_banner(
                doc,
                chapter_num="VI",
                chapter_title="Construction of Constitution and Specific Kinds of Statutes",
                chapter_subtitle="Cases 94 to 103 | Constitutional Interpretation & Strict vs. Liberal Construction",
                syllabus_topics=[
                    "A. Constitutional Construction (Cases 94–97): Verba Legis (Plain Meaning), Ratio Legis Est Anima (Intent/Spirit), Ut Res Magis Valeat Quam Pereat (Harmonious Whole)",
                    "B. Statutes Strictly Construed (Cases 98–101): Penal Statutes, Tax Statutes, Statutes in Derogation of Rights, Naturalization Laws",
                    "C. Statutes Liberally Construed (Cases 102–103): Remedial Statutes, Social Justice Legislation, Labor Laws, Pension and Retirement Laws, Election Laws"
                ]
            )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(8)
            
        render_case_to_docx(doc, c)

    doc.save(docx_path)
    print(f"Successfully generated Phase 4 DOCX file: {docx_path}")

if __name__ == "__main__":
    build_phase4_documents()
