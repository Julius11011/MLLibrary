# generator_master.py
"""
Master Compendium Generator for Statutory Construction (All 103 Cases)
Complete 6 Chapters | Dual Layer: Full ALAC & Recitations + 12-Section Case Digest Format.
"""

import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core_builder import (
    create_docx_builder, add_header_box, add_chapter_banner,
    add_heading_1, add_heading_2, add_body_p
)
from case_renderer import render_case_to_docx, render_case_to_txt

# Import cases across all 4 phases
from phase1_cases_complete import all_phase1_cases
from phase2_cases_part1 import cases_p2_part1
from phase2_cases_part2 import cases_p2_part2
from phase2_cases_part3 import cases_p2_part3
from phase3_cases_part1 import cases_p3_part1
from phase3_cases_part2 import cases_p3_part2
from phase4_cases_part1 import cases_p4_part1
from phase4_cases_part2 import cases_p4_part2
from phase4_cases_part3 import cases_p4_part3

p1_cases = all_phase1_cases
all_p2_cases = cases_p2_part1 + cases_p2_part2 + cases_p2_part3
all_p3_cases = cases_p3_part1 + cases_p3_part2
all_p4_cases = cases_p4_part1 + cases_p4_part2 + cases_p4_part3

all_master_cases = p1_cases + all_p2_cases + all_p3_cases + all_p4_cases
all_master_cases.sort(key=lambda x: x['num'])
print(f"Total Unique Master Cases: {len(all_master_cases)}")

CHAPTER_DEFINITIONS = [
    {
        "num": "I",
        "start": 1,
        "end": 7,
        "title": "Background of Statutory Construction",
        "subtitle": "Cases 1 to 7 | Foundations of Statutory Interpretation & Judicial Power",
        "topics": [
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
    },
    {
        "num": "II",
        "start": 8,
        "end": 21,
        "title": "Prelude to the Exercise of the Power",
        "subtitle": "Cases 8 to 21 | Cardinal Threshold Maxims of Textual Fidelity",
        "topics": [
            "A. Verba Legis Non Est Recedendum: From the words of the statute there must be no departure",
            "B. Absoluta Sententia Expositore Non Indigent: When the language is clear, it needs no expositor",
            "C. Dura Lex Sed Lex: The law may be harsh, but it is the law; equity cannot override positive statutory command",
            "D. Plain Meaning Rule: Clear statutory text is the best exponent of legislative will"
        ]
    },
    {
        "num": "III",
        "start": 22,
        "end": 49,
        "title": "Basic Guidelines in the Construction and Interpretation of Law",
        "subtitle": "Cases 22 to 49 | General Principles, Considerations Within & Outside the Statute",
        "topics": [
            "A. General Guidelines: Legislative Intent Must Prevail; Verba Intentioni, Non E Contra, Debent Inservire; Purposeful Construction; Nemo Tenetur Ad Impossibile; Avoid Absurdity and Injustice; Whole Statute; Every Part Given Effect",
            "B. Considerations Within the Statute: Ubi Lex Non Distinguit Nec Nos Distinguere Debemos; Ejusdem Generis; Noscitur A Sociis; Doctrine of Necessary Implication; Expressio Unius Est Exclusio Alterius; Reddendo Singula Singulis; Casus Omissus Pro Omisso Habendus Est",
            "C. Considerations Outside the Statute: Generalia Specialibus Non Derogat; Lex Posterior Derogat Priori; Lex Specialis Derogat Legi Generalis; Borrowed Statute Rule; Interpretare Et Concordare Leges Legibus Est Optimus Interpretandi"
        ]
    },
    {
        "num": "IV",
        "start": 50,
        "end": 69,
        "title": "Aids to Construction",
        "subtitle": "Cases 50 to 69 | Intrinsic Aids, Extrinsic Aids & Statutory Presumptions",
        "topics": [
            "A. Intrinsic Aids (Cases 50–59): Preambles, Title, Punctuation Marks, Capitalization, Headnotes/Epigraphs, Lingual Text and Language Reconciliations",
            "B. Extrinsic Aids (Cases 60–65): Policy and Purpose of Law, Legislative History (Debates, Reports, Committee Journals), Legal Environment, Contemporaneous Construction (Executive/Administrative Interpretation), Dictionaries",
            "C. Presumptions in Aid of Construction (Cases 66–69): Presumption Against Unconstitutionality, Presumption Against Injustice, Presumption Against Implied Repeals, Presumption Against Ineffectiveness, Presumption Against Absurdity"
        ]
    },
    {
        "num": "V",
        "start": 70,
        "end": 93,
        "title": "Interpretation of Words and Phrases",
        "subtitle": "Cases 70 to 93 | Semantic Rules, Specific Connectives & Textual Modifiers",
        "topics": [
            "A. General Words and Phrases: Generic Words, Commercial Meaning, Technical and Legal Terms",
            "B. Specific Words: Disjunctive ('or') vs. Conjunctive ('and'); Mandatory ('shall') vs. Directory ('may')",
            "C. General Words Construed Generally; Progressive Interpretation Doctrine",
            "D. Meaning of Words Qualified by Purpose of Statute",
            "E. Provisos, Exceptions, and Saving Clauses",
            "F. Statute Construed as a Whole and in Relation to Related Statutes"
        ]
    },
    {
        "num": "VI",
        "start": 94,
        "end": 103,
        "title": "Construction of Constitution and Specific Kinds of Statutes",
        "subtitle": "Cases 94 to 103 | Constitutional Interpretation & Strict vs. Liberal Construction",
        "topics": [
            "A. Constitutional Construction (Cases 94–97): Verba Legis (Plain Meaning), Ratio Legis Est Anima (Intent/Spirit), Ut Res Magis Valeat Quam Pereat (Harmonious Whole)",
            "B. Statutes Strictly Construed (Cases 98–101): Penal Statutes, Tax Statutes, Statutes in Derogation of Rights, Naturalization Laws",
            "C. Statutes Liberally Construed (Cases 102–103): Remedial Statutes, Social Justice Legislation, Labor Laws, Pension and Retirement Laws, Election Laws"
        ]
    }
]

def build_master_documents():
    txt_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Master_Complete_All_Cases.txt"
    docx_path = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\StatCons\StatCons_Master_Complete_All_Cases.docx"
    
    # 1. Build Master TXT File
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("MANILA LAW COLLEGE\n")
        f.write("Subject: Statutory Construction (2 units)\n")
        f.write("Instructor: Atty. Jason P. Raposas\n")
        f.write("Reference Syllabus: StatCon_Syllabus_JPR.pdf\n\n")
        f.write("=" * 95 + "\n")
        f.write("MASTER STATUTORY CONSTRUCTION DIGEST COMPENDIUM & TREATISE (ALL 103 CASES)\n")
        f.write("COMPLETE 6 MAJOR CHAPTERS | DUAL LAYER ALAC & 12-SECTION CASE DIGEST FORMAT\n")
        f.write("=" * 95 + "\n\n")
        
        current_ch_idx = 0
        for c in all_master_cases:
            # Check if entering a new chapter
            if current_ch_idx < len(CHAPTER_DEFINITIONS) and c['num'] == CHAPTER_DEFINITIONS[current_ch_idx]['start']:
                ch = CHAPTER_DEFINITIONS[current_ch_idx]
                f.write("\n" + "=" * 95 + "\n")
                f.write(f"CHAPTER {ch['num']}: {ch['title'].upper()}\n")
                f.write(f"{ch['subtitle']}\n")
                f.write("Topics Covered:\n")
                for t in ch['topics']:
                    f.write(f"  * {t}\n")
                f.write("=" * 95 + "\n\n")
                current_ch_idx += 1
                
            f.write(render_case_to_txt(c))
            
    print(f"Successfully generated Master TXT file: {txt_path}")

    # 2. Build Master DOCX File
    doc = create_docx_builder()
    add_header_box(
        doc,
        title="MASTER STATUTORY CONSTRUCTION COMPENDIUM & TREATISE (ALL 103 CASES)",
        subtitle="Complete 6 Major Chapters | Dual Layer ALAC + 12-Section Case Digest Format",
        course_info="Manila Law College | Statutory Construction (2 Units) | Instructor: Atty. Jason P. Raposas"
    )
    
    current_ch_idx = 0
    for c in all_master_cases:
        doc.add_page_break()
        
        if current_ch_idx < len(CHAPTER_DEFINITIONS) and c['num'] == CHAPTER_DEFINITIONS[current_ch_idx]['start']:
            ch = CHAPTER_DEFINITIONS[current_ch_idx]
            add_chapter_banner(
                doc,
                chapter_num=ch['num'],
                chapter_title=ch['title'],
                chapter_subtitle=ch['subtitle'],
                syllabus_topics=ch['topics']
            )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(8)
            current_ch_idx += 1
            
        render_case_to_docx(doc, c)

    doc.save(docx_path)
    print(f"Successfully generated Master DOCX file: {docx_path}")

if __name__ == "__main__":
    build_master_documents()
