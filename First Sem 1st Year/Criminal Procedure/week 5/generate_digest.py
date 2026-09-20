import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def generate_comprehensive_digest_with_canons():
    week5_dir = r"c:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Criminal Procedure\week 5"
    downloads_dir = r"c:\Users\JR\Downloads"
    os.makedirs(week5_dir, exist_ok=True)
    
    txt_path = os.path.join(week5_dir, "Pimentel_v_Legal_Education_Board_Case_Digest.txt")
    docx_path = os.path.join(week5_dir, "Pimentel_v_Legal_Education_Board_Case_Digest.docx")
    txt_copy = os.path.join(week5_dir, "Criminal_Procedure_Week_5_Case_Digests.txt")
    docx_copy = os.path.join(week5_dir, "Criminal_Procedure_Week_5_Case_Digests.docx")

    # =========================================================================
    # 1. TEXT CONTENT GENERATION
    # =========================================================================
    txt_content = """================================================================================
COMPREHENSIVE CASE DIGEST, RECITATION GUIDE & ETHICAL CANON ANALYSIS
CASE: PIMENTEL, ET AL. v. LEGAL EDUCATION BOARD (LEB)
CONSOLIDATED WITH: AGUILAR, ET AL. v. MEDIALDEA, ET AL.
CITATIONS: G.R. Nos. 230642 & 242954
DECISIONS: Main Decision: Sept. 10, 2019 (J. J.C. Reyes, Jr.) | Resolution on MR: Nov. 9, 2021 (J. R.V. Zalameda)
FRAMEWORK: CLASSROOM RECITATION GUIDE + BLJE + ALAC METHOD + ETHICAL CANONS
================================================================================

================================================================================
PART A: CLASSROOM DISCUSSION & RECITATION GUIDE (HOW TO RECITE IN CLASS)
================================================================================

When called upon by your professor to recite Pimentel v. Legal Education Board, follow this systematic, 4-step delivery:

--------------------------------------------------------------------------------
STEP 1: HOW TO STATE THE WHOLE FACTS (ORAL RECITATION DELIVERY)
--------------------------------------------------------------------------------
"May it please the Court / Good day, Attorney/Professor:

This case involves consolidated petitions challenging the constitutionality of Republic Act No. 7662, or the Legal Education Reform Act of 1993, and various issuances of the Legal Education Board (LEB), most notably LEBMO No. 7-2016 which mandated the PhiLSAT.

In 1993, Congress enacted R.A. No. 7662 to reform and uplift legal education standards in the Philippines, creating the Legal Education Board (LEB). Under Section 7 of the law, the LEB was granted regulatory powers over law schools, including setting accreditation standards, faculty qualifications, admission requirements, establishing a mandatory pre-Bar law internship, and adopting a system of Mandatory Continuing Legal Education (MCLE).

Pursuant to this law, the LEB issued several controversial orders:
First, under LEBMO Nos. 1-2011 and 2-2013, the LEB required all law school deans and faculty members to obtain a Master of Laws (LL.M.) degree under pain of disqualification, and prohibited law schools from admitting non-law graduates into LL.M. programs.

Second, and most notably, the LEB issued LEBMO No. 7-2016, which established the Philippine Law School Admission Test (PhiLSAT). Under this order, passing the PhiLSAT with a 55th percentile score was made an absolute, mandatory prerequisite for admission to basic law degrees (LL.B. and J.D.). Law schools nationwide were strictly prohibited from admitting any student who did not pass the PhiLSAT, under threat of fines, cancellation of accreditation, and program closure.

Consequently, two sets of petitions were filed before the Supreme Court:
1. G.R. No. 230642 was filed by Atty. Oscar Pimentel, legal educators, and law students, arguing that R.A. 7662 and LEB issuances usurped the Supreme Court's exclusive constitutional authority over the Bar under Article VIII, Section 5(5), and infringed upon the institutional academic freedom of law schools under Article XIV, Section 5(2).
2. G.R. No. 242954 was filed by aspiring law students who were disqualified from enrolling for failing or not taking the PhiLSAT, alleging violations of their right to education, due process, and equal protection."

--------------------------------------------------------------------------------
STEP 2: HOW TO STATE THE ISSUES (ORAL RECITATION DELIVERY)
--------------------------------------------------------------------------------
"The main legal issues in this case are:

1. JURISDICTION & SEPARATION OF POWERS:
   Does the State's regulation of legal education through the LEB under R.A. No. 7662 encroach upon the Supreme Court's exclusive constitutional authority over admission to the practice of law under Article VIII, Section 5(5)?

2. INSTITUTIONAL ACADEMIC FREEDOM (ADMISSION TO STUDY):
   Does the mandatory, exclusionary PhiLSAT requirement under LEBMO No. 7-2016 violate the institutional academic freedom of law schools under Article XIV, Section 5(2) to determine who may be admitted to study?

3. FACULTY & GRADUATE PROGRAM RESTRICTIONS:
   Are the LEB regulations mandating a Master of Laws (LL.M.) degree for law deans/professors and restricting LL.M. admissions exclusively to basic law degree holders constitutional?

4. PRE-BAR INTERNSHIP & MCLE REGULATION:
   Are the provisions of R.A. No. 7662 authorizing the LEB to establish a mandatory pre-Bar law practice internship (Sec. 7[g]) and to administer Mandatory Continuing Legal Education (Sec. 7[h]) constitutional?"

--------------------------------------------------------------------------------
STEP 3: HOW TO STATE THE DOCTRINES APPLIED (ORAL RECITATION DELIVERY)
--------------------------------------------------------------------------------
"The Supreme Court established three fundamental constitutional doctrines:

FIRST DOCTRINE — 'STUDY OF LAW' VS. 'PRACTICE OF LAW':
The Supreme Court established a clear constitutional boundary:
- The STUDY OF LAW is an academic endeavor falling under the State's sovereign police power over higher education (Art. XIV, Sec. 4[1]). The State may create an administrative body like the LEB to supervise law schools and set reasonable minimum quality standards.
- In contrast, the PRACTICE OF LAW and BAR ADMISSION fall under the exclusive constitutional prerogative of the Supreme Court pursuant to Article VIII, Section 5(5). Administrative bodies cannot impose prerequisites for the Bar Examinations or regulate licensed attorneys.

SECOND DOCTRINE — INSTITUTIONAL ACADEMIC FREEDOM & THE 4 ESSENTIAL FREEDOMS:
Under Article XIV, Section 5(2), institutions of higher learning enjoy academic freedom, encompassing the four essential freedoms:
  (1) Who may teach;
  (2) What may be taught;
  (3) How it shall be taught; and
  (4) Who may be admitted to study.
The State's regulatory power is strictly confined to setting 'reasonable minimum standards.' It cannot impose an absolute, exclusionary pass/fail cutoff (like PhiLSAT) that substitutes the LEB's judgment for the sound admission discretion of individual law schools.

THIRD DOCTRINE — STANDARDIZED APTITUDE TESTS AS ADVISORY TOOLS:
Standardized aptitude tests are valid in principle if used as an ADVISORY, NON-EXCLUSIONARY baseline or diagnostic guide. However, making it an exclusionary requirement that bans law schools from enrolling non-passers is unconstitutional."

--------------------------------------------------------------------------------
STEP 4: HOW TO CONCLUDE (DISPOSITIVE RULING & PRACTICAL TAKEAWAY)
--------------------------------------------------------------------------------
"In conclusion, the Supreme Court PARTIALLY GRANTED the petitions:

1. DECLARED CONSTITUTIONAL & UPHELD:
   - R.A. No. 7662 as a whole and the LEB's general authority to supervise legal education and prescribe reasonable minimum standards.
   - Standardized aptitude testing, but purely as an ADVISORY, NON-MANDATORY guideline.

2. DECLARED UNCONSTITUTIONAL & STRUCK DOWN:
   - LEBMO No. 7-2016 (PhiLSAT) in its ENTIRETY, for violating institutional academic freedom.
   - Section 7(g) of R.A. 7662 (Pre-Bar Internship), for usurping the Supreme Court's exclusive authority over Bar admission under Rule 138.
   - Section 7(h) of R.A. 7662 (MCLE regulation), for usurping the Supreme Court's exclusive jurisdiction over the Integrated Bar under BM No. 850.
   - The mandatory Master of Laws (LL.M.) requirement for faculty and deans, and the restriction on non-law graduates in LL.M. programs.

ETHICAL TAKEAWAY (BLJE):
Pimentel v. LEB preserves the independence of the legal profession under judicial custody, balances state regulatory power with university autonomy, and ensures that access to legal education is protected against arbitrary administrative exclusion."

--------------------------------------------------------------------------------
HOT-SEAT RECITATION CHEAT SHEET (RAPID PROFESSOR Q&A):
--------------------------------------------------------------------------------
Q1: Did the Supreme Court declare the LEB itself unconstitutional?
A1: NO. The Court upheld R.A. 7662 and the existence of the LEB under the State's police power over higher education (Art. XIV, Sec. 4[1]).

Q2: Did the Supreme Court say entrance exams are completely illegal?
A2: NO. The Court held that standardized aptitude tests are permissible as an ADVISORY tool, but making passing it a mandatory, exclusionary prerequisite violates academic freedom.

Q3: Why was the pre-Bar internship in Section 7(g) struck down?
A3: Because prescribing prerequisites for taking the Bar Examination is an exclusive constitutional power of the Supreme Court under Article VIII, Section 5(5) and Rule 138 of the Rules of Court.

Q4: Why can't the LEB administer MCLE under Section 7(h)?
A4: Because the LEB's jurisdiction ends when a student graduates. Once admitted to the Bar, practicing lawyers are governed exclusively by the Supreme Court pursuant to Bar Matter No. 850.

Q5: Why did the Court strike down the LL.M. requirement for law professors?
A5: Because it infringes on the law school's academic freedom to choose 'who may teach.' In law, practical wisdom from seasoned practitioners and judges is invaluable, even without a formal master's degree.

================================================================================
PART B: SYSTEMATIC ISSUE-BY-ISSUE ALAC BREAKDOWNS
================================================================================

[TOPIC 1: State Regulation of Legal Education vs. Judicial Authority]
• Issue: Does State regulation of legal education via LEB under R.A. 7662 encroach on Supreme Court authority under Art. VIII, Sec. 5(5)?
• Doctrine: State Police Power over Higher Education (Art. XIV, Sec. 4[1]) vs. Judicial Rule-Making Power (Art. VIII, Sec. 5[5]).
• ALAC:
  A — NO. The LEB's general supervision of legal education does not encroach on judicial authority.
  L — Art. XIV, Sec. 4(1) grants State supervisory power over education; Art. VIII, Sec. 5(5) grants SC exclusive power over Bar admission and legal practice.
  A — The study of law is higher education governed by State police power. Congress legitimately created the LEB to set minimum academic standards. So long as LEB does not regulate Bar exams or practicing lawyers, it is valid.
  C — R.A. No. 7662 and the LEB's jurisdiction over legal education are UPHELD as CONSTITUTIONAL.

[TOPIC 2: Mandatory PhiLSAT Exam (LEBMO No. 7-2016)]
• Issue: Is the mandatory, exclusionary PhiLSAT requirement under LEBMO No. 7-2016 constitutional?
• Doctrine: Institutional Academic Freedom — Freedom to Determine 'Who May Be Admitted to Study' (Art. XIV, Sec. 5[2]).
• ALAC:
  A — YES, it is UNCONSTITUTIONAL.
  L — Art. XIV, Sec. 5(2) guarantees academic freedom to universities. State power is limited to reasonable minimum standards, not absolute dictation.
  A — Imposing a centralized 55th percentile cutoff and prohibiting law schools from admitting non-passers strips schools of their discretion to admit students based on holistic factors (GPA, interviews, character, diversity).
  C — LEBMO No. 7-2016 is STRUCK DOWN as UNCONSTITUTIONAL in its entirety.

[TOPIC 3: Mandatory Master of Laws (LL.M.) for Faculty & Deans]
• Issue: Does mandating an LL.M. for law professors and deans violate institutional academic freedom?
• Doctrine: Institutional Academic Freedom — Freedom to Determine 'Who May Teach' (Art. XIV, Sec. 5[2]).
• ALAC:
  A — YES, it is UNCONSTITUTIONAL.
  L — Academic freedom protects the university's autonomy to select faculty based on its own assessment of expertise.
  A — Practical wisdom from veteran litigators and judges is essential in legal education. Disqualifying them for lacking an LL.M. is arbitrary and infringes on academic freedom.
  C — Mandatory LL.M. requirements in LEB issuances are STRUCK DOWN as UNCONSTITUTIONAL.

[TOPIC 4: Restrictions on Graduate Law Program Admissions]
• Issue: Does barring non-law graduates from Master of Laws programs violate academic freedom?
• Doctrine: Institutional Academic Freedom — Autonomy Over Curricula and Graduate Admissions (Art. XIV, Sec. 5[2]).
• ALAC:
  A — YES, it is UNCONSTITUTIONAL.
  L — Academic freedom protects the right to design interdisciplinary curricula and set graduate admission standards.
  A — Universities must have the latitude to offer specialized Master's degrees to professionals from diverse fields (doctors, economists, scientists) without rigid LEB restrictions.
  C — Restrictions on graduate law admissions are NULL AND VOID.

[TOPIC 5: Mandatory Pre-Bar Legal Internship (Sec. 7[g], R.A. 7662)]
• Issue: Is Section 7(g) empowering the LEB to establish a pre-Bar internship constitutional?
• Doctrine: Exclusive Judicial Authority over Admission to the Bar (Art. VIII, Sec. 5[5]; Rule 138).
• ALAC:
  A — NO, it is UNCONSTITUTIONAL.
  L — Article VIII, Section 5(5) gives the Supreme Court exclusive power to prescribe requirements for taking the Bar.
  A — Determining Bar eligibility is an exclusive judicial function. The LEB cannot add prerequisites to the Bar Examination.
  C — Section 7(g) of R.A. 7662 and Section 11(g) of LEBMO 1-2011 are STRUCK DOWN as UNCONSTITUTIONAL.

[TOPIC 6: Mandatory Continuing Legal Education (Sec. 7[h], R.A. 7662)]
• Issue: Is Section 7(h) empowering the LEB to regulate MCLE for lawyers constitutional?
• Doctrine: Exclusive Judicial Authority over the Integrated Bar & Legal Practice (Art. VIII, Sec. 5[5]; BM No. 850).
• ALAC:
  A — NO, it is UNCONSTITUTIONAL.
  L — Article VIII, Section 5(5) establishes the Supreme Court's exclusive jurisdiction over the Integrated Bar.
  A — The LEB's authority ends at law school graduation. Once admitted to the Bar, practicing lawyers are governed exclusively by the Supreme Court under BM No. 850.
  C — Section 7(h) of R.A. 7662 and Section 11(h) of LEBMO 1-2011 are STRUCK DOWN as UNCONSTITUTIONAL.

================================================================================
PART C: QUICK REFERENCE MASTER MATRIX
================================================================================
------------------------------------------------------------------------------------------------------------------------
PROVISION / REGULATION          | RULING              | CORE CONSTITUTIONAL DOCTRINE & BLJE RULE
------------------------------------------------------------------------------------------------------------------------
Creation of LEB (R.A. 7662)     | CONSTITUTIONAL      | Valid exercise of State police power over higher education (Art. XIV, Sec. 4[1]).
Mandatory PhiLSAT (LEBMO 7-2016)| UNCONSTITUTIONAL    | Infringes on institutional academic freedom regarding 'who may be admitted to study'.
Pre-Bar Internship (Sec. 7[g])  | UNCONSTITUTIONAL    | Usurps Supreme Court's exclusive authority over Bar admission (Art. VIII, Sec. 5[5]).
MCLE Regulation (Sec. 7[h])     | UNCONSTITUTIONAL    | Usurps Supreme Court's exclusive authority over the Integrated Bar (BM No. 850).
Mandatory LL.M. for Faculty     | UNCONSTITUTIONAL    | Infringes on institutional academic freedom regarding 'who may teach' (Art. XIV, Sec. 5[2]).
Graduate Program Restrictions   | UNCONSTITUTIONAL    | Violates institutional academic freedom in curricular design and graduate admissions.
------------------------------------------------------------------------------------------------------------------------

================================================================================
PART D: WHAT CANONS ARE INVOLVED, VIOLATED & APPLIED (ETHICAL CANON ANALYSIS)
================================================================================

In Basic Legal and Judicial Ethics (BLJE), Pimentel v. Legal Education Board intersects directly with the ethical duties of lawyers, legal educators, administrative authorities, and the Judiciary under the Code of Professional Responsibility and Accountability (CPRA, A.M. No. 22-09-01-SC), the former Code of Professional Responsibility (CPR), and the New Code of Judicial Conduct:

1. CPRA CANON I: INDEPENDENCE (Sections 1, 2, & 3) / CANON 1 JUDICIAL CONDUCT
   • Canon Violated/Threatened: The independence of the Judiciary and the legal profession from administrative encroachment.
   • Ethical Principle: The regulation of the legal profession, admission to the Bar, and discipline of lawyers are constitutional monopolies of the Supreme Court.
   • How Applied in Pimentel: When the LEB attempted to prescribe prerequisites for taking the Bar Examination (Sec. 7[g]) and supervise Mandatory Continuing Legal Education (Sec. 7[h]), it committed an ultra vires act that directly encroached upon Judicial Independence under Art. VIII, Sec. 5(5). The Supreme Court reaffirmed that judicial independence demands that no executive or administrative agency may dictate conditions for entering or remaining in the practice of law.

2. CPRA CANON II: PROPRIETY & RESPECT FOR COURTS (Sections 1, 2, & 12) / CPR CANONS 10 & 11
   • Canon Violated: Duty of legal practitioners and government officers to observe respect for the courts, judicial hierarchy, and constitutional authority.
   • Ethical Principle: Lawyers and administrative bodies must not usurp judicial functions or enact regulations that undermine established Supreme Court rules (such as Rule 138 on Bar Admission and Bar Matter No. 850 on MCLE).
   • How Applied in Pimentel: Administrative overreach by an agency created by Congress into areas reserved exclusively for the Judiciary violates the fundamental ethical obligation to respect and maintain the constitutional separation of powers.

3. CPRA CANON III: FIDELITY & UPHOLDING THE CONSTITUTION / CPR CANON 1 (Rules 1.01 & 1.02)
   • Canon Violated: Duty of lawyers, educators, and public officials to uphold the Constitution, obey the laws, and promote respect for legal processes.
   • Ethical Principle: An administrative issuance that violates constitutional rights is void ab initio and offends the ethical duty to uphold the Constitution.
   • How Applied in Pimentel: By issuing LEBMO No. 7-2016 which mandated an absolute exclusionary cutoff score, the LEB violated Article XIV, Section 5(2) (Institutional Academic Freedom) and Article XIV, Section 1 (Right to Accessible Education). Public officers and lawyers heading administrative boards have an affirmative ethical duty to ensure their issuances conform strictly to the Constitution.

4. CPRA CANON IV: COMPETENCE AND DILIGENCE (Sections 1, 2, & 3) / CPR CANON 5
   • Canon Applied & Clarified: Duty of lawyers and educators to keep abreast of legal developments and participate in improving the legal education system.
   • Ethical Principle: Legal competence is multifaceted; it encompasses rigorous academic learning, extensive courtroom experience, ethical grounding, and judicial service.
   • How Applied in Pimentel: While Canon IV encourages continuous improvement and high educational standards, the LEB's rigid mandate requiring a formal Master of Laws (LL.M.) for all faculty arbitrarily disregarded the profound practical competence and pedagogical value of seasoned judges and veteran trial advocates. The Court emphasized that true competence in legal education cannot be measured by a single postgraduate degree alone.

================================================================================
"""

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(txt_content)
    with open(txt_copy, "w", encoding="utf-8") as f:
        f.write(txt_content)
    
    # Also write text file to Downloads
    try:
        with open(os.path.join(downloads_dir, "Pimentel_v_Legal_Education_Board_Case_Digest.txt"), "w", encoding="utf-8") as f:
            f.write(txt_content)
    except Exception as e:
        print(f"Note on Downloads txt: {e}")

    # =========================================================================
    # 2. DOCX FILE GENERATION
    # =========================================================================
    doc = docx.Document()

    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    NAVY_HEX = "1E3A8A"
    ACCENT_BLUE = "2563EB"
    DARK_TEXT = "0F172A"
    BORDER_GREY = "CBD5E1"
    LIGHT_BG = "F8FAFC"
    ALAC_BG = "F1F5F9"
    GREEN_HEX = "15803D"
    RED_HEX = "B91C1C"
    PURPLE_HEX = "6B21A8"

    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(10.0)
    style_normal.font.color.rgb = RGBColor(15, 23, 42)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)

    # Title Block
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    title_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = title_p.add_run("PIMENTEL, ET AL. v. LEGAL EDUCATION BOARD")
    r_title.bold = True
    r_title.font.size = Pt(17)
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(12)
    sub_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = sub_p.add_run("COMPLETE CASE DIGEST, CLASSROOM RECITATION GUIDE & ETHICAL CANONS (CPRA/CPR)\nApplied Framework: Whole Facts | Issues | Doctrines | Conclusions | ALAC | Ethical Canons")
    r_sub.font.size = Pt(10.5)
    r_sub.font.color.rgb = RGBColor(71, 85, 105)
    r_sub.italic = True

    # Metadata Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    col_widths = [Inches(2.2), Inches(4.3)]
    meta_data = [
        ("Case Titles & Citations", "Oscar B. Pimentel, et al. v. Legal Education Board (LEB), G.R. No. 230642\nConsolidated with Abigail Valerie H. Aguilar, et al. v. Salvador Medialdea, et al., G.R. No. 242954"),
        ("Dates of Decisions", "Main Decision: Sept. 10, 2019 (Penned by J. J.C. Reyes, Jr., En Banc)\nResolution on MR: Nov. 9, 2021 (Penned by J. R.V. Zalameda, En Banc)"),
        ("Subject Matter", "Basic Legal and Judicial Ethics (BLJE) | Constitutional Law | CPRA Canons | Institutional Academic Freedom | Judicial Rule-Making Power"),
        ("Key Statutes & Orders", "Republic Act No. 7662 (Legal Education Reform Act of 1993)\nLEBMO No. 7-2016 (PhiLSAT), LEBMO No. 1-2011, LEBMO No. 2-2013"),
        ("Ethical Canons & Consti", "CPRA Canons I, II, III, IV | CPR Canons 1, 5, 10, 11 | 1987 Consti: Art. VIII, Sec. 5(5); Art. XIV, Sec. 1, 4(1), 5(2)")
    ]
    for row_idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        c_lbl, c_val = row.cells[0], row.cells[1]
        c_lbl.width, c_val.width = col_widths[0], col_widths[1]
        
        p_l = c_lbl.paragraphs[0]
        p_l.paragraph_format.space_after = Pt(2)
        r_l = p_l.add_run(label)
        r_l.bold = True
        r_l.font.size = Pt(9.0)
        r_l.font.color.rgb = RGBColor(30, 58, 138)
        
        p_v = c_val.paragraphs[0]
        p_v.paragraph_format.space_after = Pt(2)
        r_v = p_v.add_run(val)
        r_v.font.size = Pt(9.0)
        
        shd_l = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{LIGHT_BG}"/>')
        c_lbl._tc.get_or_add_tcPr().append(shd_l)
        
        for cell in [c_lbl, c_val]:
            tcPr = cell._tc.get_or_add_tcPr()
            borders = parse_xml(f'''
                <w:tcBorders {nsdecls("w")}>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                </w:tcBorders>
            ''')
            tcPr.append(borders)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_sec_hdr(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(30, 58, 138)
        pBdr = parse_xml(f'''
            <w:pBdr {nsdecls("w")}>
                <w:bottom w:val="single" w:sz="12" w:space="4" w:color="{NAVY_HEX}"/>
            </w:pBdr>
        ''')
        p._p.get_or_add_pPr().append(pBdr)

    def add_subsec_hdr(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(11.0)
        r.font.color.rgb = RGBColor(37, 99, 235)

    def add_bullet(bold_txt, regular_txt):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        r_b = p.add_run(bold_txt)
        r_b.bold = True
        p.add_run(regular_txt)

    def add_callout_box(title, text, border_color_hex=ACCENT_BLUE, fill_color_hex=LIGHT_BG):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.rows[0].cells[0]
        cell.width = Inches(6.5)
        
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color_hex}"/>')
        tcPr.append(shd)
        
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:bottom w:val="none"/>
                <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color_hex}"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(2)
        r_t = p.add_run(title + "\n")
        r_t.bold = True
        r_t.font.size = Pt(9.5)
        r_t.font.color.rgb = RGBColor(30, 58, 138)
        
        r_txt = p.add_run(text)
        r_txt.font.size = Pt(9.5)
        
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # -------------------------------------------------------------
    # PART A: RECITATION SCRIPT
    # -------------------------------------------------------------
    add_sec_hdr("PART A: CLASSROOM DISCUSSION & RECITATION SCRIPT")
    p = doc.add_paragraph()
    p.add_run("Use this 4-step oral delivery script when called for recitation in Basic Legal and Judicial Ethics (BLJE):")

    add_subsec_hdr("Step 1: Oral Delivery for the Whole Facts")
    add_callout_box(
        "🎙️ RECITATION SCRIPT — THE WHOLE FACTS:",
        "\"May it please the Professor/Attorney:\n\n"
        "This case involves consolidated petitions challenging the constitutionality of Republic Act No. 7662 (Legal Education Reform Act of 1993) and various orders of the Legal Education Board (LEB), most notably LEBMO No. 7-2016 which mandated the PhiLSAT.\n\n"
        "In 1993, Congress enacted R.A. No. 7662 to reform and uplift the standards of legal education, creating the Legal Education Board (LEB). Section 7 granted the LEB regulatory powers over law schools, including accreditation, faculty standards, admission requirements, establishing a mandatory pre-Bar internship, and adopting mandatory continuing legal education (MCLE).\n\n"
        "Pursuant to this law, the LEB issued several restrictive orders:\n"
        "1. LEBMO Nos. 1-2011 and 2-2013 mandated that all law school deans and professors must obtain a Master of Laws (LL.M.) degree under pain of disqualification, and barred non-law graduates from entering LL.M. programs.\n"
        "2. LEBMO No. 7-2016 established the nationwide Philippine Law School Admission Test (PhiLSAT). Passing the PhiLSAT (at least 55th percentile) was made an absolute, mandatory prerequisite for admission to law school nationwide. Law schools were strictly barred from enrolling non-passers under threat of closure, fines, and revocation of accreditation.\n\n"
        "This prompted two consolidated petitions:\n"
        "• G.R. No. 230642 filed by Atty. Oscar Pimentel, legal educators, and law students, arguing that the LEB encroached upon the Supreme Court's exclusive authority over the Bar (Art. VIII, Sec. 5[5]) and violated institutional academic freedom (Art. XIV, Sec. 5[2]).\n"
        "• G.R. No. 242954 filed by aspiring law students challenging the exclusionary PhiLSAT for violating their right to education, due process, and equal protection.\"",
        border_color_hex=NAVY_HEX, fill_color_hex="F8FAFC"
    )

    add_subsec_hdr("Step 2: Oral Delivery for the Core Issues")
    add_callout_box(
        "🎙️ RECITATION SCRIPT — THE CORE ISSUES:",
        "\"The principal issues raised before the Supreme Court are:\n\n"
        "1. JURISDICTION: Does the State's regulation of legal education through the LEB under R.A. 7662 encroach upon the Supreme Court's exclusive constitutional authority over admission to the practice of law under Article VIII, Section 5(5)?\n"
        "2. ACADEMIC FREEDOM (ADMISSIONS): Does the mandatory, exclusionary PhiLSAT requirement under LEBMO No. 7-2016 violate the institutional academic freedom of law schools under Article XIV, Section 5(2) to determine who may be admitted to study?\n"
        "3. FACULTY & CURRICULUM RESTRICTIONS: Are the LEB regulations mandating a Master of Laws (LL.M.) for faculty and barring non-law graduates from LL.M. programs constitutional?\n"
        "4. BAR REQUIREMENTS & MCLE: Are Section 7(g) (mandatory pre-Bar internship) and Section 7(h) (MCLE power) of R.A. 7662 constitutional?\"",
        border_color_hex=ACCENT_BLUE, fill_color_hex="EFF6FF"
    )

    add_subsec_hdr("Step 3: Oral Delivery for the Legal Doctrines")
    add_callout_box(
        "🎙️ RECITATION SCRIPT — THE LEGAL DOCTRINES:",
        "\"The Supreme Court laid down three cardinal constitutional doctrines:\n\n"
        "1. 'STUDY OF LAW' VS. 'PRACTICE OF LAW':\n"
        "• The STUDY OF LAW is higher education governed by the State's sovereign police power (Art. XIV, Sec. 4[1]). Congress may create an administrative agency (LEB) to set minimum educational standards.\n"
        "• The PRACTICE OF LAW and BAR ADMISSION are exclusive constitutional prerogatives of the Supreme Court (Art. VIII, Sec. 5[5]). Administrative bodies cannot impose prerequisites for the Bar Exams or regulate practicing lawyers.\n\n"
        "2. INSTITUTIONAL ACADEMIC FREEDOM (THE FOUR ESSENTIAL FREEDOMS):\n"
        "Under Article XIV, Section 5(2), universities enjoy the autonomy to determine: (1) who may teach; (2) what may be taught; (3) how it shall be taught; and (4) who may be admitted to study. State power is strictly limited to prescribing reasonable 'minimum standards' and cannot impose total dictation or exclusionary barriers.\n\n"
        "3. STANDARDIZED APTITUDE TESTS AS ADVISORY TOOLS:\n"
        "A standardized test is valid only as an ADVISORY or diagnostic guide. Transforming it into an exclusionary pass/fail requirement that prevents law schools from admitting non-passers violates institutional academic freedom.\"",
        border_color_hex=NAVY_HEX, fill_color_hex="F8FAFC"
    )

    add_subsec_hdr("Step 4: Oral Delivery for the Conclusion & Disposition")
    add_callout_box(
        "🎙️ RECITATION SCRIPT — THE CONCLUSION & DISPOSITION:",
        "\"In conclusion, the Supreme Court PARTIALLY GRANTED the petitions:\n\n"
        "1. UPHELD AS CONSTITUTIONAL:\n"
        "• R.A. No. 7662 as a whole and the LEB's general authority to supervise legal education and prescribe reasonable minimum standards.\n"
        "• Standardized aptitude tests, but purely as an ADVISORY, NON-MANDATORY guideline.\n\n"
        "2. STRUCK DOWN AS UNCONSTITUTIONAL:\n"
        "• LEBMO No. 7-2016 (PhiLSAT) in its ENTIRETY, for violating institutional academic freedom.\n"
        "• Section 7(g) of R.A. 7662 (Pre-Bar Internship), for usurping the Supreme Court's exclusive authority over Bar admission under Rule 138.\n"
        "• Section 7(h) of R.A. 7662 (MCLE regulation), for usurping the Supreme Court's exclusive jurisdiction over the Integrated Bar under BM No. 850.\n"
        "• The mandatory LL.M. degree for faculty/deans and restrictions on graduate law admissions.\n\n"
        "ETHICAL RELEVANCE (BLJE): The ruling safeguards the independence of the Judiciary and the Bar, respects university autonomy, and protects citizens' right to accessible legal education.\"",
        border_color_hex=GREEN_HEX, fill_color_hex="F0FDF4"
    )

    add_subsec_hdr("Hot-Seat Recitation Cheat Sheet (Rapid Professor Q&A)")
    qas = [
        ("Q1: Did the Supreme Court declare the LEB itself unconstitutional?", "NO. The Court upheld R.A. 7662 and the LEB's regulatory authority under State police power over higher education (Art. XIV, Sec. 4[1])."),
        ("Q2: Did the Supreme Court say entrance aptitude exams are completely illegal?", "NO. Standardized aptitude tests are valid as an ADVISORY baseline tool, but making it an absolute, exclusionary pass/fail cutoff violates academic freedom."),
        ("Q3: Why was the pre-Bar internship in Section 7(g) struck down?", "Because setting prerequisites to sit for the Bar Examination is an exclusive judicial function of the Supreme Court under Article VIII, Section 5(5) and Rule 138 of the Rules of Court."),
        ("Q4: Why can't the LEB regulate MCLE under Section 7(h)?", "Because the LEB's jurisdiction ends when a student graduates. Once admitted to the Bar, practicing lawyers are governed exclusively by the Supreme Court under Bar Matter No. 850."),
        ("Q5: Why was the mandatory LL.M. for professors struck down?", "Because it violates the school's freedom to choose 'who may teach.' In law, seasoned practitioners and judges possess immense practical wisdom even without a master's degree.")
    ]
    for q, a in qas:
        add_bullet(q + " -> ", a)

    # -------------------------------------------------------------
    # PART B: ALAC ISSUE BREAKDOWNS
    # -------------------------------------------------------------
    add_sec_hdr("PART B: SYSTEMATIC ISSUE-BY-ISSUE ALAC BREAKDOWNS")

    def add_alac_block(topic_title, issue_text, doctrine_text, ans_text, is_ans_yes, law_text, app_text, concl_text):
        add_subsec_hdr(topic_title)
        add_callout_box("⚖️ LEGAL ISSUE:", issue_text, border_color_hex=NAVY_HEX, fill_color_hex="F8FAFC")
        add_callout_box("📖 DOCTRINE APPLIED:", doctrine_text, border_color_hex=ACCENT_BLUE, fill_color_hex="EFF6FF")
        
        tbl = doc.add_table(rows=4, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        alac_widths = [Inches(1.4), Inches(5.1)]
        alac_elements = [
            ("A — ANSWER", ans_text, GREEN_HEX if is_ans_yes else RED_HEX),
            ("L — LEGAL BASIS", law_text, ACCENT_BLUE),
            ("A — ANALYSIS", app_text, DARK_TEXT),
            ("C — CONCLUSION", concl_text, NAVY_HEX)
        ]
        for idx, (lbl, content, col_hex) in enumerate(alac_elements):
            row = tbl.rows[idx]
            c_l, c_c = row.cells[0], row.cells[1]
            c_l.width, c_c.width = alac_widths[0], alac_widths[1]
            
            p_l = c_l.paragraphs[0]
            p_l.paragraph_format.space_before = Pt(3)
            p_l.paragraph_format.space_after = Pt(3)
            r_l = p_l.add_run(lbl)
            r_l.bold = True
            r_l.font.size = Pt(9.0)
            r_l.font.color.rgb = RGBColor(30, 58, 138)
            
            p_c = c_c.paragraphs[0]
            p_c.paragraph_format.space_before = Pt(3)
            p_c.paragraph_format.space_after = Pt(3)
            r_c = p_c.add_run(content)
            r_c.font.size = Pt(9.0)
            if idx == 0:
                r_c.bold = True
            
            shd_l = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{ALAC_BG}"/>')
            c_l._tc.get_or_add_tcPr().append(shd_l)
            
            for cell in [c_l, c_c]:
                tcPr = cell._tc.get_or_add_tcPr()
                borders = parse_xml(f'''
                    <w:tcBorders {nsdecls("w")}>
                        <w:top w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                        <w:left w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                        <w:right w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                    </w:tcBorders>
                ''')
                tcPr.append(borders)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_alac_block(
        topic_title="Topic 1: State Regulation of Legal Education vs. Judicial Authority",
        issue_text="Does State regulation of legal education via LEB under R.A. 7662 encroach on Supreme Court authority under Art. VIII, Sec. 5(5)?",
        doctrine_text="State Police Power over Higher Education (Art. XIV, Sec. 4[1]) vs. Judicial Rule-Making Power (Art. VIII, Sec. 5[5]).",
        ans_text="NO. The LEB's general supervision of legal education does not encroach on judicial authority.",
        is_ans_yes=True,
        law_text="Art. XIV, Sec. 4(1) grants State supervisory power over education; Art. VIII, Sec. 5(5) grants SC exclusive power over Bar admission and legal practice.",
        app_text="The study of law is higher education governed by State police power. Congress legitimately created the LEB to set minimum academic standards. So long as LEB does not regulate Bar exams or practicing lawyers, it is valid.",
        concl_text="R.A. No. 7662 and the LEB's jurisdiction over legal education are UPHELD as CONSTITUTIONAL."
    )

    add_alac_block(
        topic_title="Topic 2: Mandatory PhiLSAT Exam (LEBMO No. 7-2016)",
        issue_text="Is the mandatory, exclusionary PhiLSAT requirement under LEBMO No. 7-2016 constitutional?",
        doctrine_text="Institutional Academic Freedom — Freedom to Determine 'Who May Be Admitted to Study' (Art. XIV, Sec. 5[2]).",
        ans_text="YES, it is UNCONSTITUTIONAL.",
        is_ans_yes=False,
        law_text="Art. XIV, Sec. 5(2) guarantees academic freedom to universities. State power is limited to reasonable minimum standards, not absolute dictation.",
        app_text="Imposing a centralized 55th percentile cutoff and prohibiting law schools from admitting non-passers strips schools of their discretion to admit students based on holistic factors (GPA, interviews, character, diversity).",
        concl_text="LEBMO No. 7-2016 is STRUCK DOWN as UNCONSTITUTIONAL in its entirety."
    )

    add_alac_block(
        topic_title="Topic 3: Mandatory Master of Laws (LL.M.) for Faculty & Deans",
        issue_text="Does mandating an LL.M. for law professors and deans violate institutional academic freedom?",
        doctrine_text="Institutional Academic Freedom — Freedom to Determine 'Who May Teach' (Art. XIV, Sec. 5[2]).",
        ans_text="YES, it is UNCONSTITUTIONAL.",
        is_ans_yes=False,
        law_text="Academic freedom protects the university's autonomy to select faculty based on its own assessment of expertise.",
        app_text="Practical wisdom from veteran litigators and judges is essential in legal education. Disqualifying them for lacking an LL.M. is arbitrary and infringes on academic freedom.",
        concl_text="Mandatory LL.M. requirements in LEB issuances are STRUCK DOWN as UNCONSTITUTIONAL."
    )

    add_alac_block(
        topic_title="Topic 4: Restrictions on Graduate Law Program Admissions",
        issue_text="Does barring non-law graduates from Master of Laws programs violate academic freedom?",
        doctrine_text="Institutional Academic Freedom — Autonomy Over Curricula and Graduate Admissions (Art. XIV, Sec. 5[2]).",
        ans_text="YES, it is UNCONSTITUTIONAL.",
        is_ans_yes=False,
        law_text="Academic freedom protects the right to design interdisciplinary curricula and set graduate admission standards.",
        app_text="Universities must have the latitude to offer specialized Master's degrees to professionals from diverse fields (doctors, economists, scientists) without rigid LEB restrictions.",
        concl_text="Restrictions on graduate law admissions are NULL AND VOID."
    )

    add_alac_block(
        topic_title="Topic 5: Mandatory Pre-Bar Legal Internship (Sec. 7[g], R.A. 7662)",
        issue_text="Is Section 7(g) empowering the LEB to establish a pre-Bar internship constitutional?",
        doctrine_text="Exclusive Judicial Authority over Admission to the Bar (Art. VIII, Sec. 5[5]; Rule 138).",
        ans_text="NO, it is UNCONSTITUTIONAL.",
        is_ans_yes=False,
        law_text="Article VIII, Section 5(5) gives the Supreme Court exclusive power to prescribe requirements for taking the Bar.",
        app_text="Determining Bar eligibility is an exclusive judicial function. The LEB cannot add prerequisites to the Bar Examination.",
        concl_text="Section 7(g) of R.A. 7662 and Section 11(g) of LEBMO 1-2011 are STRUCK DOWN as UNCONSTITUTIONAL."
    )

    add_alac_block(
        topic_title="Topic 6: Mandatory Continuing Legal Education (Sec. 7[h], R.A. 7662)",
        issue_text="Is Section 7(h) empowering the LEB to regulate MCLE for lawyers constitutional?",
        doctrine_text="Exclusive Judicial Authority over the Integrated Bar & Legal Practice (Art. VIII, Sec. 5[5]; BM No. 850).",
        ans_text="NO, it is UNCONSTITUTIONAL.",
        is_ans_yes=False,
        law_text="Article VIII, Section 5(5) establishes the Supreme Court's exclusive jurisdiction over the Integrated Bar.",
        app_text="The LEB's authority ends at law school graduation. Once admitted to the Bar, practicing lawyers are governed exclusively by the Supreme Court under BM No. 850.",
        concl_text="Section 7(h) of R.A. 7662 and Section 11(h) of LEBMO 1-2011 are STRUCK DOWN as UNCONSTITUTIONAL."
    )

    # -------------------------------------------------------------
    # PART C: REFERENCE MATRIX
    # -------------------------------------------------------------
    add_sec_hdr("PART C: QUICK REFERENCE MASTER MATRIX")
    matrix_table = doc.add_table(rows=7, cols=3)
    matrix_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    matrix_table.autofit = False
    t_widths = [Inches(2.0), Inches(1.8), Inches(2.7)]
    headers = ["Provision / Regulation", "Supreme Court Ruling", "Core Constitutional Doctrine & BLJE Rule"]

    hdr_row = matrix_table.rows[0]
    for idx, text in enumerate(headers):
        cell = hdr_row.cells[idx]
        cell.width = t_widths[idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(255, 255, 255)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{NAVY_HEX}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    matrix_rows = [
        ("Creation of LEB (R.A. No. 7662)", "UPHELD (Constitutional)", "Valid exercise of State police power over higher education (Art. XIV, Sec. 4[1])."),
        ("Mandatory PhiLSAT (LEBMO 7-2016)", "STRUCK DOWN (Unconstitutional)", "Infringes on institutional academic freedom regarding 'who may be admitted to study' (Art. XIV, Sec. 5[2])."),
        ("Pre-Bar Internship (Sec. 7[g])", "STRUCK DOWN (Unconstitutional)", "Usurps Supreme Court's exclusive authority over Bar admission (Art. VIII, Sec. 5[5])."),
        ("MCLE Supervision (Sec. 7[h])", "STRUCK DOWN (Unconstitutional)", "Usurps Supreme Court's exclusive authority over the Integrated Bar and legal practice (BM No. 850)."),
        ("Mandatory LL.M. for Faculty", "STRUCK DOWN (Unconstitutional)", "Infringes on institutional academic freedom regarding 'who may teach' (Art. XIV, Sec. 5[2])."),
        ("Graduate Program Restrictions", "STRUCK DOWN (Unconstitutional)", "Violates academic freedom in designing curricula and admitting graduate students.")
    ]
    for row_idx, (prov, ruling, doctrine) in enumerate(matrix_rows, start=1):
        row = matrix_table.rows[row_idx]
        for col_idx, text in enumerate([prov, ruling, doctrine]):
            cell = row.cells[col_idx]
            cell.width = t_widths[col_idx]
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(text)
            r.font.size = Pt(9.0)
            if col_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if "UPHELD" in text:
                    r.bold = True
                    r.font.color.rgb = RGBColor(21, 128, 61)
                else:
                    r.bold = True
                    r.font.color.rgb = RGBColor(185, 28, 28)
            bg_color = LIGHT_BG if row_idx % 2 == 1 else "FFFFFF"
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
            cell._tc.get_or_add_tcPr().append(shd)
            tcPr = cell._tc.get_or_add_tcPr()
            borders = parse_xml(f'''
                <w:tcBorders {nsdecls("w")}>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>
                </w:tcBorders>
            ''')
            tcPr.append(borders)

    # -------------------------------------------------------------
    # PART D: CANONS INVOLVED, VIOLATED & APPLIED
    # -------------------------------------------------------------
    add_sec_hdr("PART D: WHAT CANONS ARE INVOLVED, VIOLATED & APPLIED (ETHICAL CANON ANALYSIS)")
    p = doc.add_paragraph()
    p.add_run("In Basic Legal and Judicial Ethics (BLJE), the ruling in ").font.size = Pt(10)
    p.add_run("Pimentel v. Legal Education Board").bold = True
    p.add_run(" directly applies and interprets the following ethical Canons under the ")
    p.add_run("Code of Professional Responsibility and Accountability (CPRA / A.M. No. 22-09-01-SC)").bold = True
    p.add_run(", the former CPR, and the New Code of Judicial Conduct:")

    canons_data = [
        ("1. CPRA CANON I: INDEPENDENCE (Sections 1, 2, & 3) / CANON 1 JUDICIAL CONDUCT",
         "• Canon Violated/Threatened: The independence of the Judiciary and the legal profession from administrative encroachment.\n"
         "• Ethical Principle: The regulation of the legal profession, admission to the Bar, and discipline of lawyers are constitutional monopolies of the Supreme Court.\n"
         "• Application in Pimentel: When the LEB attempted to prescribe prerequisites for taking the Bar Examination (Sec. 7[g]) and supervise Mandatory Continuing Legal Education (Sec. 7[h]), it committed an ultra vires act that directly encroached upon Judicial Independence under Art. VIII, Sec. 5(5). The Supreme Court reaffirmed that judicial independence demands that no executive or administrative agency may dictate conditions for entering or remaining in the practice of law.",
         NAVY_HEX, "F8FAFC"),
        
        ("2. CPRA CANON II: PROPRIETY & RESPECT FOR COURTS (Sections 1, 2, & 12) / CPR CANONS 10 & 11",
         "• Canon Violated: Duty of legal practitioners and government officers to observe respect for the courts, judicial hierarchy, and constitutional authority.\n"
         "• Ethical Principle: Lawyers and administrative bodies must not usurp judicial functions or enact regulations that undermine established Supreme Court rules (such as Rule 138 on Bar Admission and Bar Matter No. 850 on MCLE).\n"
         "• Application in Pimentel: Administrative overreach by an agency created by Congress into areas reserved exclusively for the Judiciary violates the fundamental ethical obligation to respect and maintain the constitutional separation of powers.",
         ACCENT_BLUE, "EFF6FF"),
        
        ("3. CPRA CANON III: FIDELITY & UPHOLDING THE CONSTITUTION / CPR CANON 1 (Rules 1.01 & 1.02)",
         "• Canon Violated: Duty of lawyers, educators, and public officials to uphold the Constitution, obey the laws, and promote respect for legal processes.\n"
         "• Ethical Principle: An administrative issuance that violates constitutional rights is void ab initio and offends the ethical duty to uphold the Constitution.\n"
         "• Application in Pimentel: By issuing LEBMO No. 7-2016 which mandated an absolute exclusionary cutoff score, the LEB violated Article XIV, Section 5(2) (Institutional Academic Freedom) and Article XIV, Section 1 (Right to Accessible Education). Public officers and lawyers heading administrative boards have an affirmative ethical duty to ensure their issuances conform strictly to the Constitution.",
         PURPLE_HEX, "FAF5FF"),
        
        ("4. CPRA CANON IV: COMPETENCE AND DILIGENCE (Sections 1, 2, & 3) / CPR CANON 5",
         "• Canon Applied & Clarified: Duty of lawyers and educators to keep abreast of legal developments and participate in improving the legal education system.\n"
         "• Ethical Principle: Legal competence is multifaceted; it encompasses rigorous academic learning, extensive courtroom experience, ethical grounding, and judicial service.\n"
         "• Application in Pimentel: While Canon IV encourages continuous improvement and high educational standards, the LEB's rigid mandate requiring a formal Master of Laws (LL.M.) for all faculty arbitrarily disregarded the profound practical competence and pedagogical value of seasoned judges and veteran trial advocates. The Court emphasized that true competence in legal education cannot be measured by a single postgraduate degree alone.",
         GREEN_HEX, "F0FDF4")
    ]

    for title_c, text_c, border_c, fill_c in canons_data:
        add_callout_box(title_c, text_c, border_color_hex=border_c, fill_color_hex=fill_c)

    doc.save(docx_path)
    doc.save(docx_copy)
    print(f"Written Word documents: {docx_path} and {docx_copy}")
    
    # Attempt to copy/save to Downloads
    try:
        doc.save(os.path.join(downloads_dir, "Pimentel_v_Legal_Education_Board_Case_Digest.docx"))
        print("Updated Downloads docx file.")
    except Exception as e:
        # If open in Word, save as updated copy
        alt_docx = os.path.join(downloads_dir, "Pimentel_v_Legal_Education_Board_Case_Digest_Updated.docx")
        doc.save(alt_docx)
        print(f"File in Downloads was locked by Word. Saved updated file to: {alt_docx}")

if __name__ == "__main__":
    generate_comprehensive_digest_with_canons()
