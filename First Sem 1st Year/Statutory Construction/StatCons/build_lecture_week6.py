# build_lecture_week6.py
"""
Generates the Statutory Construction Week 6 Master Lecture & Study Guide
in DOCX, TXT, and converts to PDF in:
C:\\Users\\JR\\Downloads\\14All-All41\\MLC\\First Sem 1st Year\\Statutory Construction\\week 6\\
"""

import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

COLOR_PRIMARY = RGBColor(15, 32, 67)      # Deep Navy
COLOR_SECONDARY = RGBColor(120, 30, 30)  # Deep Crimson
COLOR_STUDENT = RGBColor(13, 110, 75)    # Emerald Green
COLOR_LATIN = RGBColor(140, 50, 0)       # Amber
COLOR_TEXT = RGBColor(33, 37, 41)        # Charcoal

HEX_PRIMARY = "0F2043"
HEX_SECONDARY = "781E1E"
HEX_LIGHT_BG = "F4F6F9"
HEX_STUDENT_BG = "EBF7F0"
HEX_LATIN_BG = "FDF6EE"
HEX_CALLOUT_BG = "F0F4FA"
HEX_WARNING_BG = "FDF2E9"

def add_header_banner(doc, title, subtitle, course_info):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, HEX_PRIMARY)
    set_cell_margins(cell, top=200, bottom=200, left=240, right=240)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    run1 = p.add_run("MANILA LAW COLLEGE\n")
    run1.font.name = "Arial"
    run1.font.size = Pt(14)
    run1.font.bold = True
    run1.font.color.rgb = RGBColor(255, 255, 255)
    
    run2 = p.add_run(f"{course_info}\n")
    run2.font.name = "Arial"
    run2.font.size = Pt(10)
    run2.font.color.rgb = RGBColor(210, 225, 245)
    
    run3 = p.add_run(f"{title}\n")
    run3.font.name = "Arial"
    run3.font.size = Pt(12.5)
    run3.font.bold = True
    run3.font.color.rgb = RGBColor(255, 215, 0)
    
    run4 = p.add_run(subtitle)
    run4.font.name = "Arial"
    run4.font.size = Pt(9.5)
    run4.font.italic = True
    run4.font.color.rgb = RGBColor(230, 235, 245)

def add_module_banner(doc, module_num, module_title, module_subtitle):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, HEX_PRIMARY)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    
    r_ch = p.add_run(f"MODULE {module_num}: ")
    r_ch.font.name = "Arial"
    r_ch.font.size = Pt(12)
    r_ch.font.bold = True
    r_ch.font.color.rgb = RGBColor(255, 215, 0)
    
    r_title = p.add_run(f"{module_title.upper()}\n")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(12)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(255, 255, 255)
    
    r_sub = p.add_run(module_subtitle)
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(215, 230, 250)

def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY
    return p

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_SECONDARY
    return p

def add_p(doc, text, bold_prefix="", italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = "Calibri"
        r_b.font.size = Pt(10.5)
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_TEXT
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.italic = italic
    r.font.color.rgb = COLOR_TEXT
    return p

def add_callout(doc, title, items, hex_bg=HEX_LIGHT_BG, title_color=COLOR_PRIMARY):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, hex_bg)
    set_cell_margins(cell, top=100, bottom=100, left=160, right=160)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    r_head = p.add_run(f"📌 {title.upper()}\n")
    r_head.font.name = "Arial"
    r_head.font.size = Pt(10)
    r_head.font.bold = True
    r_head.font.color.rgb = title_color
    
    for it in items:
        p_it = cell.add_paragraph()
        p_it.paragraph_format.space_before = Pt(1)
        p_it.paragraph_format.space_after = Pt(2)
        p_it.paragraph_format.line_spacing = 1.15
        if isinstance(it, tuple):
            r1 = p_it.add_run(it[0] + " ")
            r1.font.name = "Calibri"
            r1.font.size = Pt(9.5)
            r1.font.bold = True
            r1.font.color.rgb = COLOR_TEXT
            r2 = p_it.add_run(it[1])
            r2.font.name = "Calibri"
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = COLOR_TEXT
        else:
            r = p_it.add_run(it)
            r.font.name = "Calibri"
            r.font.size = Pt(9.5)
            r.font.color.rgb = COLOR_TEXT

def add_recitation_tip(doc, concept_name, plain_meaning, recitation_script):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, HEX_STUDENT_BG)
    set_cell_margins(cell, top=100, bottom=100, left=160, right=160)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    r_h = p.add_run(f"💡 HOW LAW STUDENTS CAN GRASP & RECITE: {concept_name.upper()}\n")
    r_h.font.name = "Arial"
    r_h.font.size = Pt(9.5)
    r_h.font.bold = True
    r_h.font.color.rgb = COLOR_STUDENT
    
    p1 = cell.add_paragraph()
    p1.paragraph_format.space_before = Pt(1)
    p1.paragraph_format.space_after = Pt(2)
    p1.paragraph_format.line_spacing = 1.15
    r1_l = p1.add_run("• Plain English Concept: ")
    r1_l.font.name = "Calibri"
    r1_l.font.size = Pt(9.5)
    r1_l.font.bold = True
    r1_l.font.color.rgb = COLOR_STUDENT
    r1_v = p1.add_run(plain_meaning)
    r1_v.font.name = "Calibri"
    r1_v.font.size = Pt(9.5)
    r1_v.font.color.rgb = COLOR_TEXT
    
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(1)
    p2.paragraph_format.space_after = Pt(2)
    p2.paragraph_format.line_spacing = 1.15
    r2_l = p2.add_run("• Golden Recitation Script: ")
    r2_l.font.name = "Calibri"
    r2_l.font.size = Pt(9.5)
    r2_l.font.bold = True
    r2_l.font.color.rgb = COLOR_STUDENT
    r2_v = p2.add_run(f'"{recitation_script}"')
    r2_v.font.name = "Calibri"
    r2_v.font.size = Pt(9.5)
    r2_v.font.italic = True
    r2_v.font.color.rgb = COLOR_TEXT

def generate_lecture_guide():
    out_dir = r"C:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Statutory Construction\week 6"
    os.makedirs(out_dir, exist_ok=True)
    
    docx_path = os.path.join(out_dir, "Statutory_Construction_Week_6_Lecture_Guide.docx")
    txt_path = os.path.join(out_dir, "Statutory_Construction_Week_6_Lecture_Guide.txt")
    
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.85)
        s.bottom_margin = Inches(0.85)
        s.left_margin = Inches(0.85)
        s.right_margin = Inches(0.85)
        
    add_header_banner(
        doc,
        title="STATUTORY CONSTRUCTION COMPREHENSIVE LECTURE & STUDY GUIDE",
        subtitle="Coverage: Chapter II (Prelude to Exercise of Power) to Chapter IV.C (Presumptions in StatCon)",
        course_info="Manila Law College | Statutory Construction (2 Units) | Instructor: Atty. Jason P. Raposas"
    )
    
    # -------------------------------------------------------------------------
    # MODULE 1
    # -------------------------------------------------------------------------
    doc.add_paragraph().paragraph_format.space_before = Pt(8)
    add_module_banner(doc, "1", "Prelude to the Exercise of Judicial Power", "The Threshold Test: Can the Court Even Interpret the Law?")
    
    add_h1(doc, "A. Verba Legis Non Est Recedendum (The Plain Meaning Rule)")
    add_p(doc, "From the words of a statute there must be no departure. When the language of the law is clear, plain, and free from ambiguity, the court has no authority to construe—it must simply apply the law as written.", bold_prefix="• Doctrinal Core: ")
    add_p(doc, "The legislature (Congress) is the sole constitutional authority vested with lawmaking power. When Congress expresses its will in plain terms, judges cannot insert exceptions or alter the scope under the guise of interpretation without committing unconstitutional judicial legislation.", bold_prefix="• Why the Rule Exists: ")
    
    add_callout(doc, "Classroom & Practical Example", [
        ("Hypothetical:", "A city ordinance penalizes: 'Any person who brings any motorized vehicle inside the botanical gardens.' An individual rides a modern high-end electric scooter into the garden and is cited for a violation. He argues that electric scooters are eco-friendly and did not exist when the ordinance was enacted."),
        ("Application:", "Under Verba Legis, an electric scooter is unequivocally a motorized vehicle. Because the text makes no distinction between gasoline and electric motors, the plain words govern without exception.")
    ], hex_bg=HEX_LIGHT_BG)
    
    add_callout(doc, "Landmark Case: People v. Mapa (G.R. No. L-22301, Aug 30, 1967)", [
        ("Facts:", "Mario Mapa was prosecuted for illegal possession of firearms. He invoked his appointment as a secret agent of the Provincial Governor of Batangas."),
        ("Ruling:", "Section 879 of the Revised Administrative Code explicitly enumerated the public officials exempt from firearm licensing (e.g. mayors, municipal treasurers, governors). Secret agents were omitted from the text. The Supreme Court applied Verba Legis: where the statute is clear, courts cannot supply an exemption that Congress did not write.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    add_recitation_tip(
        doc,
        "Verba Legis in Recitation",
        "If the statutory wording is clear and unambiguous, interpretation is forbidden; only literal application is permitted.",
        "Your Honor, under the doctrine of Verba Legis Non Est Recedendum, where the law speaks in clear and categorical language, there is no room for interpretation. There is only room for application. The duty of the court is jus dicere, non jus dare."
    )
    
    add_h1(doc, "B. Absoluta Sententia Expositore Non Indiget & Dura Lex Sed Lex")
    add_p(doc, "An absolute or clear statement needs no expositor. When the statute is plain and unequivocal, bringing in external historical opinions or secondary aids only clouds what Congress already made clear.", bold_prefix="• Absoluta Sententia: ")
    add_p(doc, "The law may be harsh, but it is the law. Courts cannot refuse to apply a valid statute merely because its enforcement works hardship, inconvenience, or perceived unfairness to a particular litigant.", bold_prefix="• Dura Lex Sed Lex: ")
    
    add_callout(doc, "Landmark Case: Pascual v. Pascual-Bautista (G.R. No. 84240, Mar 25, 1992)", [
        ("Facts:", "An illegitimate child sought to inherit from her father's legitimate brother via intestate succession."),
        ("Ruling:", "Article 992 of the Civil Code explicitly bars illegitimate children from inheriting intestate from the legitimate children and relatives of their father or mother (the Iron Curtain Rule). The Court ruled that even if the rule feels harsh, courts must apply the absolute text: Dura lex sed lex.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    # -------------------------------------------------------------------------
    # MODULE 2
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_module_banner(doc, "2", "Basic Guidelines in Construction and Interpretation", "General Guidelines, Intent, Harmony, and Preventing Absurdity")
    
    add_h1(doc, "A. Intention of the Statute Must Prevail (Ratio Legis Est Anima Legis)")
    add_p(doc, "The reason of the law is the soul of the law. The primary objective of all statutory interpretation is to discover and effectuate the true legislative intent. Words ought to be subservient to the intent, not the intent to the words (Verba intentioni, non e contra, debent inservire).", bold_prefix="• Core Principle: ")
    
    add_h1(doc, "B. Construction to Avoid Absurdity and Injustice")
    add_p(doc, "The legislature is presumed never to intend an unjust, absurd, or ridiculous consequence. If a strictly literal reading of a statute defeats its manifest spirit or produces a monstrously absurd result, the spirit of the law controls over the letter.", bold_prefix="• Spirit Over Letter: ")
    
    add_callout(doc, "Landmark Case: Alonzo v. IAC (G.R. No. 72873, May 28, 1987)", [
        ("Facts:", "Co-heirs sold their inherited property to a buyer who lived openly on the land and built a house for 13 years. A co-heir later sought to exercise legal redemption under Civil Code Article 1088, which allows redemption within 30 days from written notice of the sale, arguing no written notice had ever been served."),
        ("Ruling:", "The Supreme Court (per Justice Cruz) held that actual open knowledge for over a decade satisfied the purpose of the law. To allow redemption after 13 years under a hyper-technical reading of 'written notice' would perpetuate grave injustice. The spirit of the law transcends its literal letters.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    add_h1(doc, "C. Ut Res Magis Valeat Quam Pereat & Nemo Tenetur Ad Impossibile")
    add_p(doc, "That the thing may rather have effect than perish. An interpretation that gives life, validity, and operational effect to the statute is favored over one that destroys it or renders it inoperative.", bold_prefix="• Ut Res Magis Valeat: ")
    add_p(doc, "No one is bound to perform the impossible. Statutory requirements are never interpreted to compel acts that are physically or legally impossible without fault of the party.", bold_prefix="• Nemo Tenetur: ")
    
    # -------------------------------------------------------------------------
    # MODULE 3
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_module_banner(doc, "3", "Considerations Within the Statute (Intrinsic Canons)", "Noscitur a Sociis, Ejusdem Generis, Expressio Unius, Casus Omissus, Reddendo Singula")
    
    add_h1(doc, "A. Ubi Lex Non Distinguit Nec Nos Distinguere Debemus")
    add_p(doc, "Where the law makes no distinction, the court must make no distinction. When a statute uses general terms (e.g., 'all public officers', 'any action'), courts cannot carve out classifications or restrict the scope.", bold_prefix="• Rule: ")
    
    add_h1(doc, "B. Noscitur A Sociis (Associated Words)")
    add_p(doc, "A word is known by the company it keeps. The meaning of a doubtful or ambiguous statutory term is ascertained and colored by reference to the words with which it is associated in the sentence.", bold_prefix="• Rule: ")
    
    add_h1(doc, "C. Ejusdem Generis (Of the Same Kind)")
    add_p(doc, "Where general words follow an enumeration of specific persons or things of a distinct class, the general words are construed to apply only to persons or things of the same general character as those specifically listed.", bold_prefix="• The Formula: ")
    add_p(doc, "(1) Enumeration of specific words belonging to a homogenous class; (2) General catch-all term follows; (3) Specific terms do not exhaust the entire class.", bold_prefix="• Strict Requisites: ")
    
    add_callout(doc, "Landmark Case: People v. Manantan (G.R. No. L-14129, July 31, 1962)", [
        ("Facts:", "A Justice of the Peace (Judge) was charged with partisan political campaigning under Section 54 of the Revised Election Code, which prohibited 'any justice, judge, fiscal, treasurer, or assessor' from aiding candidates. He argued that the omission of 'justice of the peace' (which appeared in the predecessor law) meant he was excluded under Expressio Unius."),
        ("Ruling:", "Convicted. Expressio Unius is merely an auxiliary aid and cannot defeat clear legislative intent. The general word 'judge' comprehends all judges, including Justices of the Peace. The purpose of insulating the judiciary from political partisanship applied equally to all judicial officers.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    add_h1(doc, "D. Expressio Unius Est Exclusio Alterius & Casus Omissus")
    add_p(doc, "The express mention of one person, thing, or consequence implies the exclusion of all others. If Congress writes a specific list, what is excluded is presumed intentionally omitted.", bold_prefix="• Expressio Unius: ")
    add_p(doc, "A case omitted is held to have been omitted intentionally. Even if a court believes Congress made an accidental omission, the court cannot supply the missing term under the guise of construction.", bold_prefix="• Casus Omissus: ")
    
    add_h1(doc, "E. Reddendo Singula Singulis & Doctrine of Necessary Implication")
    add_p(doc, "Referring each to each. Distributing respective words and phrases to their appropriate subjects or antecedents in the statute.", bold_prefix="• Reddendo Singula: ")
    add_p(doc, "What is implied in a statute is as much a part of it as what is expressed (Ex necessitate legis). Every statutory grant of power carries with it all incidental powers necessary to execute the main grant.", bold_prefix="• Necessary Implication: ")
    
    # -------------------------------------------------------------------------
    # MODULE 4
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_module_banner(doc, "4", "Considerations Outside the Statute (Inter-Statutory Canons)", "Generalia Specialibus, Lex Posterior, Lex Specialis, Borrowed Statutes & Harmonization")
    
    add_h1(doc, "A. Generalia Specialibus Non Derogat & Lex Specialis Derogat Generali")
    add_p(doc, "A general law does not repeal or derogate from a special law. A special law represents the concentrated legislative will on a specific subject and governs over a general law, even if the general law was enacted later.", bold_prefix="• General vs Special: ")
    
    add_callout(doc, "Landmark Case: Tuna Processing, Inc. v. Philippine Kingford, Inc. (G.R. No. 185582, Feb 29, 2012)", [
        ("The Clash:", "Corporation Code Sec. 133 (General Law: unlicensed foreign corporation cannot sue) vs. Alternative Dispute Resolution Act / RA 9285 (Special Law: any party to arbitration may enforce an award)."),
        ("Ruling:", "The ADR Act as Lex Specialis governs the specific enforcement of international arbitral awards over the general capacity-to-sue provisions of the Corporation Code. Foreign corporations may enforce arbitral awards without a local license.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    add_h1(doc, "B. Lex Posterior Generalis Non Derogat Legi Priori Specialis")
    add_p(doc, "A subsequent general statute does not repeal a prior special statute unless there is a clear, manifest, and irreconcilable intention to abrogate the special law.", bold_prefix="• Rule: ")
    
    add_callout(doc, "Landmark Case: CIR v. Philippine Airlines, Inc. (G.R. No. 180066, July 8, 2009)", [
        ("Facts:", "PAL held a special legislative charter (PD 1590) allowing it to pay a franchise tax 'in lieu of all other taxes.' Congress later enacted broad VAT amendments to the NIRC."),
        ("Ruling:", "The later general NIRC amendments did not repeal PAL's prior special charter exemption. A general tax law will not impliedly repeal a specialized franchise charter.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    add_h1(doc, "C. Interpretare Et Concordare Leges Legibus (Harmonization)")
    add_p(doc, "To interpret and harmonize laws with laws is the best method of interpretation. Every statute is presumed enacted in harmony with the existing body of laws. Courts must strive to reconcile conflicting statutes so that both may stand.", bold_prefix="• Golden Rule: ")
    
    add_callout(doc, "Landmark Case: Remo v. Secretary of Foreign Affairs (G.R. No. 169202, Mar 5, 2010)", [
        ("Harmonization:", "Civil Code Art. 370 gives a married woman the option to use her husband's surname, while the Philippine Passport Act (RA 8239) regulates passport issuance. Harmonized: A woman has the right to choose her surname, but once chosen and entered into official passport records, she cannot revert to her maiden name while married except upon death, annulment, or divorce.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    add_h1(doc, "D. The Borrowed Statute Rule")
    add_p(doc, "When a Philippine statute is copied or patterned after a foreign law (e.g. U.S. Uniform Commercial Code, Negotiable Instruments Law, Rules of Court), the authoritative decisions of the highest court of the originating jurisdiction are entitled to great persuasive weight in Philippine jurisprudence.", bold_prefix="• Rule: ")
    
    # -------------------------------------------------------------------------
    # MODULE 5
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_module_banner(doc, "5", "Aids to Construction (Intrinsic & Extrinsic Aids)", "Title, Preambles, Punctuation, Lingual Text, Legislative History, Contemporaneous Construction")
    
    add_h1(doc, "A. Intrinsic Aids (Found Within the Four Corners)")
    add_p(doc, "• Title: Defines the object and scope of the statute; persuasive when the body is ambiguous, but cannot override clear operational text.\n• Preamble: The 'Whereas' clauses stating the mischief and purpose; the key to unlock the minds of the lawmakers.\n• Punctuations: Semicolons separate distinct independent thoughts; commas set off qualifying clauses.\n• Headnotes & Epigraphs: Editorial section headings; convenient for reference, but non-controlling.\n• Lingual Text: Where an enacted statute was originally drafted in Spanish (e.g., Revised Penal Code), the Spanish text controls over defective English translations.")
    
    add_callout(doc, "Landmark Case: People of the Philippines v. Subido (G.R. No. L-21734, Sept 5, 1975)", [
        ("Facts:", "Accused claimed subsidiary imprisonment could not be imposed because the penalty did not explicitly attach to the fine."),
        ("Ruling:", "The Court looked at the semicolon separating the penalty clauses. The semicolon indicated a distinct, independent clause, establishing that subsidiary imprisonment attached to the fine.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    add_h1(doc, "B. Extrinsic Aids (Found Outside the Four Corners)")
    add_p(doc, "• Legislative History: Committee reports, explanatory notes, and floor sponsorship speeches showing the evil sought to be remedied.\n• Contemporaneous Construction: Executive and administrative interpretations given by departments enforcing the law carry great respect, but cannot override clear text.\n• Dictionaries: Defining ordinary words according to their natural, customary, and accepted usage.")
    
    add_callout(doc, "Landmark Case: CIR v. SM Prime Holdings, Inc. (G.R. No. 183505, Feb 26, 2010)", [
        ("Facts:", "The CIR attempted to impose VAT on cinema ticket sales by administrative circular."),
        ("Ruling:", "Tracing the legislative history of VAT enactments, the Supreme Court found that Congress intentionally excluded movie theaters from VAT because they were already subject to local amusement taxes. Administrative circulars cannot expand a tax statute beyond legislative intent.")
    ], hex_bg=HEX_CALLOUT_BG)
    
    # -------------------------------------------------------------------------
    # MODULE 6
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_module_banner(doc, "6", "Presumptions in Statutory Construction", "The Five Cardinal Default Truths Assumed by the Judiciary")
    
    add_h1(doc, "A. The Five Cardinal Presumptions")
    add_p(doc, "1. Presumption Against Unconstitutionality: Statutes are presumed valid; the challenger bears the heavy burden of showing clear constitutional violation (LAMP v. DBM, G.R. 164987).\n2. Presumption Against Injustice: In case of doubt, it is presumed Congress intended right and justice to prevail (Melchor v. COA, G.R. 95398 / Alonzo v. IAC).\n3. Presumption Against Implied Repeals: All statutes are presumed enacted with knowledge of prior laws and intended to coexist (COA v. Cebu, G.R. 141386).\n4. Presumption Against Ineffectiveness: Congress is presumed to intend that every word and section shall have operational effect.\n5. Presumption Against Absurdity: Statutes must receive a sensible construction that avoids irrational or unworkable results.")
    
    # -------------------------------------------------------------------------
    # MODULE 7
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_module_banner(doc, "7", "The Law Student's 4-Step Attack Algorithm & Recitation Cheat Sheet", "Step-by-Step Decision Flowchart for Recitations and Bar Exams")
    
    add_callout(doc, "4-Step Bar Exam Attack Flowchart", [
        ("STEP 1: IS THE TEXT CLEAR?", "If YES -> Apply VERBA LEGIS / ABSOLUTA SENTENTIA / DURA LEX without interpretation."),
        ("STEP 2: INTERPRETING WORDS IN ONE LAW?", "Check for EJUSDEM GENERIS (class + general term), NOSCITUR A SOCIIS (associated words), EXPRESSIO UNIUS (express list), or UBI LEX NON DISTINGUIT (broad terms)."),
        ("STEP 3: CONFLICT BETWEEN TWO LAWS?", "Apply GENERALIA SPECIALIBUS (special law prevails), LEX POSTERIOR GENERALIS (special law survives general law), or HARMONIZATION."),
        ("STEP 4: CLAIMS OF INJUSTICE OR INVALIDITY?", "Apply PRESUMPTION OF CONSTITUTIONALITY, PRESUMPTION AGAINST INJUSTICE (Ratio Legis), and PRESUMPTION AGAINST IMPLIED REPEAL.")
    ], hex_bg=HEX_STUDENT_BG, title_color=COLOR_STUDENT)
    
    add_callout(doc, "Top 10 Latin Maxims to Memorize for Recitations", [
        ("1. Verba legis non est recedendum:", "From the words of a statute there must be no departure."),
        ("2. Ratio legis est anima legis:", "The reason of the law is the soul of the law."),
        ("3. Absoluta sententia expositore non indiget:", "A clear statement needs no interpreter."),
        ("4. Ubi lex non distinguit nec nos distinguere debemus:", "Where the law does not distinguish, we ought not to distinguish."),
        ("5. Ejusdem generis:", "Of the same kind or nature."),
        ("6. Noscitur a sociis:", "A word is known by its associates."),
        ("7. Expressio unius est exclusio alterius:", "The express mention of one is the exclusion of others."),
        ("8. Generalia specialibus non derogat:", "General words do not derogate from special provisions."),
        ("9. Interpretare et concordare leges legibus:", "To harmonize laws with laws is the best method of interpretation."),
        ("10. Legis interpretatio legis vim obtinet:", "The judicial interpretation of the law has the force of law.")
    ], hex_bg=HEX_LATIN_BG, title_color=COLOR_LATIN)
    
    # Save DOCX
    doc.save(docx_path)
    print(f"[OK] Generated DOCX: {docx_path}")
    
    # Save TXT
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("MANILA LAW COLLEGE\n")
        f.write("STATUTORY CONSTRUCTION COMPREHENSIVE LECTURE & STUDY GUIDE\n")
        f.write("Coverage: Chapter II (Prelude to Exercise of Power) to Chapter IV.C (Presumptions in StatCon)\n")
        f.write("=" * 95 + "\n\n")
        
        for p in doc.paragraphs:
            if p.text.strip():
                f.write(p.text + "\n")
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    f.write(c.text + "\n")
            f.write("-" * 80 + "\n")
            
    print(f"[OK] Generated TXT: {txt_path}")
    return docx_path, txt_path

if __name__ == "__main__":
    generate_lecture_guide()
