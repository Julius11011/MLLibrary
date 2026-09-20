# generator_week6.py
"""
Generator for Statutory Construction Week 6 Case Compendium
Coverage:
- Chapter III.B: Considerations Within the Statute (Cases 1-2)
- Chapter III.C: Considerations Outside the Statute (Cases 3-8)
- Chapter IV.A: Intrinsic Aids (Cases 9-18)
- Chapter IV.B: Extrinsic Aids (Cases 19-24)
- Chapter IV.C: Presumptions in Statutory Construction (Cases 25-33)
"""

import os
import sys
import copy
import docx
from docx.shared import Pt, Inches, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core_builder import (
    create_docx_builder, add_header_box, add_chapter_banner
)
from case_renderer import render_case_to_docx, render_case_to_txt

import phase1_cases_complete as p1
import phase2_cases_part1 as p21
import phase2_cases_part2 as p22
import phase2_cases_part3 as p23
import phase3_cases_part1 as p31
import phase3_cases_part2 as p32

def get_case_by_title_keyword(keyword, *modules):
    for mod in modules:
        for attr in dir(mod):
            if 'case' in attr.lower():
                val = getattr(mod, attr)
                if isinstance(val, list):
                    for c in val:
                        if isinstance(c, dict) and keyword.lower() in c.get('title', '').lower():
                            return copy.deepcopy(c)
    raise ValueError(f"Case with keyword '{keyword}' not found in provided modules.")

def build_week6_cases():
    cases_week6 = []

    # -------------------------------------------------------------------------
    # TOPIC III.B: CONSIDERATIONS WITHIN THE STATUTE
    # -------------------------------------------------------------------------
    # Case 1: Escribano v. Avila
    c1 = get_case_by_title_keyword('escribano', p23)
    c1['num'] = 1
    c1['syllabus'] = "III. Basic Guidelines in Construction -> B. Considerations within the Statute -> Associated Words / Ejusdem Generis / Expressio Unius"
    c1['chapter_section'] = "III.B"
    cases_week6.append(c1)

    # Case 2: People v. Manantan
    c2 = get_case_by_title_keyword('manantan', p23)
    c2['num'] = 2
    c2['syllabus'] = "III. Basic Guidelines in Construction -> B. Considerations within the Statute -> Expressio Unius Est Exclusio Alterius / Casus Omissus"
    c2['chapter_section'] = "III.B"
    cases_week6.append(c2)

    # -------------------------------------------------------------------------
    # TOPIC III.C: CONSIDERATIONS OUTSIDE THE STATUTE
    # -------------------------------------------------------------------------
    # Case 3: PEZA v. Green Asia
    c3 = get_case_by_title_keyword('green asia', p23)
    c3['num'] = 3
    c3['syllabus'] = "III. Basic Guidelines in Construction -> C. Considerations outside the Statute -> a. Generalia Specialibus Non Derogat & b. Lex Posterior Derogat Priori"
    c3['chapter_section'] = "III.C"
    cases_week6.append(c3)

    # Case 4: Tuna Processing, Inc. v. Philippine Kingford, Inc.
    c4 = get_case_by_title_keyword('tuna processing', p23)
    c4['num'] = 4
    c4['syllabus'] = "III. Basic Guidelines in Construction -> C. Considerations outside the Statute -> c. Lex Specialis Derogat Legi Generalis (ADR Act vs Corporation Code)"
    c4['chapter_section'] = "III.C"
    cases_week6.append(c4)

    # Case 5: Magno v. COMELEC
    c5 = get_case_by_title_keyword('magno', p23)
    c5['num'] = 5
    c5['syllabus'] = "III. Basic Guidelines in Construction -> C. Considerations outside the Statute -> d. Lex Posterior Generalis Non Derogat Legi Priori Specialis"
    c5['chapter_section'] = "III.C"
    cases_week6.append(c5)

    # Case 6: CIR v. PAL
    c6 = get_case_by_title_keyword('philippine airlines', p23)
    c6['num'] = 6
    c6['syllabus'] = "III. Basic Guidelines in Construction -> C. Considerations outside the Statute -> d. Lex Posterior Generalis Non Derogat Legi Priori Specialis (Special Tax Franchise vs NIRC)"
    c6['chapter_section'] = "III.C"
    cases_week6.append(c6)

    # Case 7: Republic v. Del Monte Motors, Inc.
    c7 = get_case_by_title_keyword('del monte', p23)
    c7['num'] = 7
    c7['syllabus'] = "III. Basic Guidelines in Construction -> C. Considerations outside the Statute -> e. Borrowed Statute Rule (UCC / Chattel Mortgage Law)"
    c7['chapter_section'] = "III.C"
    cases_week6.append(c7)

    # Case 8: Remo v. Secretary of Foreign Affairs
    c8 = get_case_by_title_keyword('remo', p23)
    c8['num'] = 8
    c8['syllabus'] = "III. Basic Guidelines in Construction -> C. Considerations outside the Statute -> f. Interpretare et Concordare Leges Legibus Est Optimus Interpretandi"
    c8['chapter_section'] = "III.C"
    cases_week6.append(c8)

    # -------------------------------------------------------------------------
    # TOPIC IV.A: INTRINSIC AIDS TO CONSTRUCTION
    # -------------------------------------------------------------------------
    # Case 9: People v. Francisco
    c9 = get_case_by_title_keyword('francisco', p31)
    c9['num'] = 9
    c9['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> b. Title of the Statute (RA 7659 / Complex Crime of Illegal Recruitment & Rape)"
    c9['chapter_section'] = "IV.A"
    cases_week6.append(c9)

    # Case 10: PNB v. Office of the President
    c10 = get_case_by_title_keyword('tapio', p31)
    c10['num'] = 10
    c10['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> a. Preambles (PD 957 / Buyer Protection against Foreclosing Bank)"
    c10['chapter_section'] = "IV.A"
    cases_week6.append(c10)

    # Case 11: Ebarle v. Sucaldito
    c11 = get_case_by_title_keyword('sucaldito', p31)
    c11['num'] = 11
    c11['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> a. Preambles & b. Title (EO 264 / Complaints against Local Officials)"
    c11['chapter_section'] = "IV.A"
    cases_week6.append(c11)

    # Case 12: Commissioner of Customs v. Relunia
    c12 = get_case_by_title_keyword('relunia', p31)
    c12['num'] = 12
    c12['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> e. Headnotes or Epigraphs (Tariff and Customs Code Sec. 3514)"
    c12['chapter_section'] = "IV.A"
    cases_week6.append(c12)

    # Case 13: In Re: Estate of Emil H. Johnson
    c13 = get_case_by_title_keyword('johnson', p31)
    c13['num'] = 13
    c13['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> e. Headnotes / Section Headings (Code of Civil Procedure Sec. 636)"
    c13['chapter_section'] = "IV.A"
    cases_week6.append(c13)

    # Case 14: People v. Subido
    c14 = get_case_by_title_keyword('subido', p31)
    c14['num'] = 14
    c14['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> c. Punctuations (Semicolon Separating Penal Sanctions / Subsidiary Imprisonment)"
    c14['chapter_section'] = "IV.A"
    cases_week6.append(c14)

    # Case 15: Florentino v. PNB
    c15 = get_case_by_title_keyword('florentino and florentino', p31)
    c15['num'] = 15
    c15['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> c. Punctuations (Semicolon Separating Qualifying Clauses / RA 897 Backpay Certificates)"
    c15['chapter_section'] = "IV.A"
    cases_week6.append(c15)

    # Case 16: Mapa v. Arroyo
    c16 = get_case_by_title_keyword('mapa and candido mapa', p31)
    c16['num'] = 16
    c16['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> a. Preambles, b. Title & c. Punctuations (PD 957 / Scope of Subdivision Protection)"
    c16['chapter_section'] = "IV.A"
    cases_week6.append(c16)

    # Case 17: People v. Yabut
    c17 = get_case_by_title_keyword('yabut', p31)
    c17['num'] = 17
    c17['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> f. Lingual Text (RPC Art. 160 Quasi-Recidivism / Spanish Text Prevails over English)"
    c17['chapter_section'] = "IV.A"
    cases_week6.append(c17)

    # Case 18: Baking v. Director of Prisons
    c18 = get_case_by_title_keyword('baking', p31)
    c18['num'] = 18
    c18['syllabus'] = "IV. Aids to Construction -> A. Intrinsic Aids -> f. Lingual Text (Spanish Enacted Text vs English Translation / Complex Crimes)"
    c18['chapter_section'] = "IV.A"
    cases_week6.append(c18)

    # -------------------------------------------------------------------------
    # TOPIC IV.B: EXTRINSIC AIDS TO CONSTRUCTION
    # -------------------------------------------------------------------------
    # Case 19: MMDA v. Garin
    c19 = get_case_by_title_keyword('garin', p32)
    c19['num'] = 19
    c19['syllabus'] = "IV. Aids to Construction -> B. Extrinsic Aids -> a. Policy & b. Purpose of the Law (RA 7924 / Absence of Police Power to Confiscate Driver's Licenses)"
    c19['chapter_section'] = "IV.B"
    cases_week6.append(c19)

    # Case 20: CIR v. SM Prime Holdings, Inc.
    c20 = get_case_by_title_keyword('sm prime', p32)
    c20['num'] = 20
    c20['syllabus'] = "IV. Aids to Construction -> B. Extrinsic Aids -> c. Legislative History & e. Contemporaneous Construction (Movie Theaters Not Subject to VAT)"
    c20['chapter_section'] = "IV.B"
    cases_week6.append(c20)

    # Case 21: Laxamana v. Baltazar
    c21 = get_case_by_title_keyword('laxamana', p32)
    c21['num'] = 21
    c21['syllabus'] = "IV. Aids to Construction -> B. Extrinsic Aids -> c. Legislative History & Legislative Evolution (RAC Sec. 2195 / Vice-Mayor Succession)"
    c21['chapter_section'] = "IV.B"
    cases_week6.append(c21)

    # Case 22: Alvarez v. Guingona
    c22 = get_case_by_title_keyword('alvarez', p32)
    c22['num'] = 22
    c22['syllabus'] = "IV. Aids to Construction -> B. Extrinsic Aids -> c. Legislative History & Congressional Debates (RA 7720 / Internal Revenue Allotment for City Conversion)"
    c22['chapter_section'] = "IV.B"
    cases_week6.append(c22)

    # Case 23: Boie-Takeda Chemicals v. De La Serna
    c23 = get_case_by_title_keyword('boie-takeda', p32)
    c23['num'] = 23
    c23['syllabus'] = "IV. Aids to Construction -> B. Extrinsic Aids -> e. Contemporaneous Construction (PD 851 13th Month Pay / Administrative Rules vs Basic Salary)"
    c23['chapter_section'] = "IV.B"
    cases_week6.append(c23)

    # Case 24: Ang v. Court of Appeals
    c24 = get_case_by_title_keyword('sagud', p32)
    c24['num'] = 24
    c24['syllabus'] = "IV. Aids to Construction -> B. Extrinsic Aids -> f. Dictionaries & Ordinary Meaning (Rules on Electronic Evidence / Ephemeral Communications)"
    c24['chapter_section'] = "IV.B"
    cases_week6.append(c24)

    # -------------------------------------------------------------------------
    # TOPIC IV.C: PRESUMPTIONS IN AID OF CONSTRUCTION
    # -------------------------------------------------------------------------
    # Case 25: Alvarez v. Guingona (Presumption against unconstitutionality)
    c25 = copy.deepcopy(c22)
    c25['num'] = 25
    c25['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> a. Presumption Against Unconstitutionality (Burden of Proof to Overcome Presumption of Validity of RA 7720)"
    c25['chapter_section'] = "IV.C"
    cases_week6.append(c25)

    # Case 26: LAMP v. DBM Secretary
    c26 = get_case_by_title_keyword('lamp', p32)
    c26['num'] = 26
    c26['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> a. Presumption Against Unconstitutionality (Presumption Favoring Constitutionality of GAA CDF/PDAF Provisions)"
    c26['chapter_section'] = "IV.C"
    cases_week6.append(c26)

    # Case 27: NHA v. Reyes
    c27 = get_case_by_title_keyword('posadas', p32)
    c27['num'] = 27
    c27['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> b. Presumption Against Injustice & Absurdity (PD 464 Expropriation Just Compensation Rules)"
    c27['chapter_section'] = "IV.C"
    cases_week6.append(c27)

    # Case 28: Aparri v. Court of Appeals
    c28 = get_case_by_title_keyword('aparri', p32)
    c28['num'] = 28
    c28['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> d. Presumption Against Ineffectiveness & e. Presumption Against Absurdity (Expiration of Term vs Removal)"
    c28['chapter_section'] = "IV.C"
    cases_week6.append(c28)

    # Case 29: Alonzo v. Intermediate Appellate Court
    c29 = get_case_by_title_keyword('alonzo', p21)
    c29['num'] = 29
    c29['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> b. Presumption Against Injustice & e. Presumption Against Absurdity (Civil Code Art. 1088 / Spirit over Letter)"
    c29['chapter_section'] = "IV.C"
    cases_week6.append(c29)

    # Case 30: Melchor v. Commission on Audit
    c30 = get_case_by_title_keyword('melchor', p21)
    c30['num'] = 30
    c30['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> b. Presumption Against Injustice (Quantum Meruit Compensation on Executed Government Contracts)"
    c30['chapter_section'] = "IV.C"
    cases_week6.append(c30)

    # Case 31: Republic v. Court of Appeals (Molina)
    c31 = get_case_by_title_keyword('molina', p32)
    c31['num'] = 31
    c31['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> c. Presumption Against Implied Repeals & Presumption in Favor of Validity of Marriage (FC Art. 36)"
    c31['chapter_section'] = "IV.C"
    cases_week6.append(c31)

    # Case 32: COA v. Province of Cebu
    c32 = get_case_by_title_keyword('cebu', p22)
    c32['num'] = 32
    c32['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> c. Presumption Against Implied Repeals & Harmonization (RA 5447 vs Local Government Code Special Education Fund)"
    c32['chapter_section'] = "IV.C"
    cases_week6.append(c32)

    # Case 33: Remo v. Secretary of Foreign Affairs
    c33 = copy.deepcopy(c8)
    c33['num'] = 33
    c33['syllabus'] = "IV. Aids to Construction -> C. Presumptions -> c. Presumption Against Implied Repeals & Harmonization of Statutes (Civil Code Art. 370 vs Passport Act RA 8239)"
    c33['chapter_section'] = "IV.C"
    cases_week6.append(c33)

    return cases_week6

def generate_week6_files():
    out_dir = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\week 6"
    os.makedirs(out_dir, exist_ok=True)
    
    txt_path = os.path.join(out_dir, "Statutory_Construction_Week_6_Cases.txt")
    docx_path = os.path.join(out_dir, "Statutory_Construction_Week_6_Cases.docx")
    
    cases = build_week6_cases()
    print(f"Total Week 6 Cases to render: {len(cases)}")
    
    # 1. Plain Text Generation
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("MANILA LAW COLLEGE\n")
        f.write("Subject: Statutory Construction (2 units)\n")
        f.write("Instructor: Atty. Jason P. Raposas\n")
        f.write("WEEK 6 CASE DIGESTS & STATUTORY CONSTRUCTION COMPENDIUM\n")
        f.write("Standardized Format: Statutory Construction Case Digest Format.docx (12-Section Architecture)\n")
        f.write("=" * 95 + "\n\n")
        
        current_section = None
        for c in cases:
            sec = c.get('chapter_section', '')
            if sec != current_section:
                current_section = sec
                f.write("\n" + "=" * 95 + "\n")
                if sec == "III.B":
                    f.write("CHAPTER III: BASIC GUIDELINES IN THE CONSTRUCTION AND INTERPRETATION OF LAWS\n")
                    f.write("TOPIC III.B: CONSIDERATIONS WITHIN THE STATUTE (Associated Words, Ejusdem Generis, Expressio Unius)\n")
                elif sec == "III.C":
                    f.write("CHAPTER III: BASIC GUIDELINES IN THE CONSTRUCTION AND INTERPRETATION OF LAWS\n")
                    f.write("TOPIC III.C: CONSIDERATIONS OUTSIDE THE STATUTE (Generalia Specialibus, Lex Posterior, Lex Specialis, Borrowed Statute, Harmonization)\n")
                elif sec == "IV.A":
                    f.write("CHAPTER IV: AIDS TO CONSTRUCTION\n")
                    f.write("TOPIC IV.A: INTRINSIC AIDS (Preambles, Title, Punctuations, Capitalizations, Headnotes/Epigraphs, Lingual Text)\n")
                elif sec == "IV.B":
                    f.write("CHAPTER IV: AIDS TO CONSTRUCTION\n")
                    f.write("TOPIC IV.B: EXTRINSIC AIDS (Policy, Purpose, Legislative History, Contemporaneous Construction, Dictionaries)\n")
                elif sec == "IV.C":
                    f.write("CHAPTER IV: AIDS TO CONSTRUCTION\n")
                    f.write("TOPIC IV.C: PRESUMPTIONS IN STATUTORY CONSTRUCTION (Against Unconstitutionality, Injustice, Implied Repeals, Ineffectiveness, Absurdity)\n")
                f.write("=" * 95 + "\n\n")
            
            f.write(render_case_to_txt(c))
            
    print(f"[OK] Generated TXT: {txt_path}")
    
    # 2. DOCX Generation
    doc = create_docx_builder()
    add_header_box(
        doc,
        title="STATUTORY CONSTRUCTION COMPENDIUM: WEEK 6 CASE STUDIES",
        subtitle="Chapters III.B, III.C & IV (Intrinsic Aids, Extrinsic Aids & Presumptions) | 12-Section Case Digest Architecture",
        course_info="Manila Law College | Statutory Construction (2 Units) | Instructor: Atty. Jason P. Raposas"
    )
    
    current_sec = None
    for idx, c in enumerate(cases):
        doc.add_page_break()
        sec = c.get('chapter_section', '')
        
        if sec != current_sec:
            current_sec = sec
            if sec == "III.B":
                add_chapter_banner(
                    doc,
                    chapter_num="III (Part 2)",
                    chapter_title="Basic Guidelines in Construction: Considerations Within the Statute",
                    chapter_subtitle="Cases 1 to 2 | Associated Words, Ejusdem Generis, Expressio Unius Est Exclusio Alterius",
                    syllabus_topics=[
                        "• Associated Words & Noscitur a Sociis: Meaning of words ascertained by associated terms",
                        "• Expressio Unius Est Exclusio Alterius: Express mention of one thing implies exclusion of others",
                        "• Casus Omissus: Words omitted cannot be supplied under the guise of judicial construction"
                    ]
                )
            elif sec == "III.C":
                add_chapter_banner(
                    doc,
                    chapter_num="III (Part 3)",
                    chapter_title="Considerations Outside the Statute",
                    chapter_subtitle="Cases 3 to 8 | Inter-Statutory Canons, Borrowed Statutes & Harmonization",
                    syllabus_topics=[
                        "• Generalia Specialibus Non Derogat: General law does not derogate from a special law",
                        "• Lex Posterior Derogat Priori: Later statute repeals an earlier inconsistent statute",
                        "• Lex Specialis Derogat Legi Generalis: Special statute governs over general law",
                        "• Lex Posterior Generalis Non Derogat Legi Priori Specialis: Later general law does not repeal prior special law",
                        "• Borrowed Statute Rule: Judicial construction of parent jurisdiction has persuasive authority",
                        "• Interpretare et Concordare Leges Legibus Est Optimus Interpretandi: Harmonization of laws is the best method of interpretation"
                    ]
                )
            elif sec == "IV.A":
                add_chapter_banner(
                    doc,
                    chapter_num="IV (Part 1)",
                    chapter_title="Aids to Construction: Intrinsic Aids",
                    chapter_subtitle="Cases 9 to 18 | Textual & Structural Aids Found Within the Statute Itself",
                    syllabus_topics=[
                        "• Preambles & Title: Explaining legislative intent, object, and scope",
                        "• Headnotes & Epigraphs: Non-controlling topical guides to statutory sections",
                        "• Punctuations & Capitalizations: Semicolons and commas determining clause independence",
                        "• Lingual Text & Translation: Controlling enacted language (Spanish vs. English text)"
                    ]
                )
            elif sec == "IV.B":
                add_chapter_banner(
                    doc,
                    chapter_num="IV (Part 2)",
                    chapter_title="Aids to Construction: Extrinsic Aids",
                    chapter_subtitle="Cases 19 to 24 | External Sources, Legislative History & Contemporary Context",
                    syllabus_topics=[
                        "• Policy & Purpose of the Law: Ascertaining legislative mischief and remedial objective",
                        "• Legislative History: Committee reports, legislative debates, and statutory evolution",
                        "• Contemporaneous Construction: Executive and administrative interpretations given great weight",
                        "• Dictionaries & Ordinary Meaning: Natural and customary definitions of words"
                    ]
                )
            elif sec == "IV.C":
                add_chapter_banner(
                    doc,
                    chapter_num="IV (Part 3)",
                    chapter_title="Aids to Construction: Presumptions in Statutory Construction",
                    chapter_subtitle="Cases 25 to 33 | Judicial Presumptions Guiding Interpretation",
                    syllabus_topics=[
                        "• Presumption Against Unconstitutionality: Statutes are presumed valid; burden on challenger",
                        "• Presumption Against Injustice & Absurdity: Spirit over letter; avoiding unjust interpretations",
                        "• Presumption Against Implied Repeals: All laws presumed enacted with knowledge of prior laws and intended to coexist",
                        "• Presumption Against Ineffectiveness: Statutes presumed intended to operate effectively"
                    ]
                )
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(6)
            
        render_case_to_docx(doc, c)
        
    doc.save(docx_path)
    print(f"[OK] Generated DOCX: {docx_path}")
    return docx_path, txt_path

if __name__ == "__main__":
    generate_week6_files()
