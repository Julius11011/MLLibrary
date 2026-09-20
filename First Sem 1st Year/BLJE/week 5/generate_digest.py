import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_files():
    week5_dir = r"c:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Criminal Procedure\week 5"
    os.makedirs(week5_dir, exist_ok=True)
    
    txt_path = os.path.join(week5_dir, "Pimentel_v_Legal_Education_Board_Case_Digest.txt")
    docx_path = os.path.join(week5_dir, "Pimentel_v_Legal_Education_Board_Case_Digest.docx")
    
    # -------------------------------------------------------------
    # 1. WRITE PLAIN TEXT FILE
    # -------------------------------------------------------------
    txt_content = """================================================================================
CASE DIGEST: PIMENTEL, ET AL. v. LEGAL EDUCATION BOARD (LEB)
APPLIED FRAMEWORK: BLJE (Background, Legal Issues, Judgment, Explanation)
================================================================================

CITATION & METADATA:
- Case Title: Oscar B. Pimentel, Errol B. Comafay, Jr., Rene B. Gorospe, Edwin R. Sandoval, Victoria B. Florido, Maricel P. Seno, Dennis R. Gorecho v. Legal Education Board (LEB), represented by its Chairperson Emerson B. Aquende, and LEB Member Zenaida N. Elepaño
- Consolidated With: Abigail Valerie H. Aguilar, et al. v. Hon. Salvador Medialdea, in his capacity as Executive Secretary, et al. (G.R. No. 242954)
- Docket Numbers: G.R. Nos. 230642 & 242954
- Decisions:
    * Main Decision: September 10, 2019 (En Banc, penned by Associate Justice Jose C. Reyes, Jr.)
    * Resolution on Motion for Reconsideration: November 9, 2021 (En Banc, penned by Associate Justice Rodil V. Zalameda)
- Key Laws & Issuances:
    * Republic Act No. 7662 (Legal Education Reform Act of 1993)
    * LEB Memorandum Order No. 7, Series of 2016 (LEBMO No. 7-2016 / PhiLSAT)
    * LEB Memorandum Order No. 1, Series of 2011 (LEBMO No. 1-2011)
    * LEB Memorandum Order No. 2, Series of 2013 (LEBMO No. 2-2013)
    * 1987 Philippine Constitution: Art. VIII, Sec. 5(5); Art. XIV, Sec. 1, 4(1), 5(2); Art. III, Sec. 1

================================================================================
B — BACKGROUND & WHOLE FACTS
================================================================================

1. LEGISLATIVE ENACTMENT (REPUBLIC ACT NO. 7662):
On December 23, 1993, Congress passed Republic Act No. 7662, known as the "Legal Education Reform Act of 1993." The law declared the policy of the State to uplift the standards of legal education, prepare law students for advocacy and judicial service, and meet the needs of a developing society.

To implement these goals, R.A. No. 7662 created the Legal Education Board (LEB) as an administrative body to administer and supervise the legal education system in the country. Section 7 of R.A. No. 7662 granted the LEB various powers, including:
  - Setting standards for accreditation of law schools (Sec. 7[c]);
  - Prescribing minimum requirements for admission to legal education and minimum qualifications of faculty members (Sec. 7[e]);
  - Establishing a law practice internship as a requirement for taking the Bar Examinations (Sec. 7[g]);
  - Adopting a system of mandatory continuing legal education (MCLE) for practicing lawyers (Sec. 7[h]).

2. LEB IMPLEMENTING ISSUANCES & THE PHILSAT MANDATE:
Pursuant to R.A. No. 7662, the LEB issued several administrative orders:
  a. LEBMO No. 1-2011 & LEBMO No. 2 (2013):
     - Established policies, standards, and curriculum guidelines.
     - Imposed a mandatory requirement that all law school faculty and deans must obtain a Master of Laws (LL.M.) degree within a prescribed period under pain of disqualification.
     - Restricted admission to graduate law degree programs (Master of Laws) exclusively to holders of basic law degrees (Bachelor of Laws / Juris Doctor).
  b. LEBMO No. 7-2016 (Issued on December 29, 2016):
     - Established the nationwide Philippine Law School Admission Test (PhiLSAT).
     - Standardized the PhiLSAT as a prerequisite for admission to all basic law programs (LL.B. and J.D.) in all Philippine law schools.
     - Mandated a passing score (55th percentile rank) determined by the LEB.
     - Imposed an absolute exclusionary rule: law schools were strictly prohibited from admitting any applicant who did not take or pass the PhiLSAT, or who lacked an LEB Certificate of Exemption.
     - Prescribed severe administrative sanctions on non-compliant law schools, including fines, cancellation of accreditation, and program closure/phase-out.
  c. LEB Memorandum Circular No. 6 (Series of 2017) & Related Circulars:
     - Imposed strict conditions on "conditional enrollment," requiring students to pass subsequent PhiLSAT test dates or face mandatory disqualification from continuing their law studies.

3. THE CONSOLIDATED PETITIONS BEFORE THE SUPREME COURT:
Aggrieved by the LEB's regulations, two sets of petitioners filed petitions before the Supreme Court:
  - G.R. No. 230642 (Oscar B. Pimentel, et al.): Filed by legal educators, lawyers, and law students challenging the constitutionality of R.A. No. 7662 and LEBMO No. 7-2016 on grounds of separation of powers and academic freedom.
  - G.R. No. 242954 (Abigail Valerie H. Aguilar, et al.): Filed by aspiring law students who were unable to take or pass the PhiLSAT, challenging the exclusionary nature of the exam under the equal protection and due process clauses.

4. CORE ARGUMENTS:
  - Petitioners argued:
    1. Encroachment on Judicial Power: R.A. No. 7662 encroaches upon the exclusive constitutional jurisdiction of the Supreme Court under Art. VIII, Sec. 5(5) of the 1987 Constitution to promulgate rules concerning admission to the practice of law, the Integrated Bar, and legal assistance.
    2. Infringement on Academic Freedom: Mandating the PhiLSAT as an absolute exclusionary entrance filter infringes on the constitutional academic freedom of law schools (Art. XIV, Sec. 5[2]) to determine for themselves "who may be admitted to study" and "who may teach."
    3. Violation of Due Process and Right to Education: Arbitrary disqualification denies aspiring students their fundamental right to accessible education (Art. XIV, Sec. 1).
    4. Undue Delegation of Legislative Power: R.A. No. 7662 lacks adequate standards.
  - Respondents (LEB & Executive Secretary) argued:
    1. Legal education is higher education falling under the State's broad police power to regulate educational institutions under Art. XIV, Sec. 4(1).
    2. The Supreme Court's constitutional authority applies only to the "practice of law" and admission to the Bar, not to the prior "study of law."
    3. The PhiLSAT is a reasonable regulatory measure designed to elevate the standard of legal education.

================================================================================
L — LEGAL ISSUES
================================================================================

1. JURISDICTION & SEPARATION OF POWERS:
   Does the State's regulation and supervision of legal education through an administrative body (LEB) under R.A. No. 7662 unconstitutionally encroach upon the Supreme Court's exclusive power to promulgate rules concerning admission to the practice of law under Article VIII, Section 5(5) of the 1987 Constitution?

2. INSTITUTIONAL ACADEMIC FREEDOM (ADMISSION TO STUDY):
   Does the mandatory, exclusionary nature of the PhiLSAT under LEBMO No. 7-2016 violate the constitutional academic freedom of higher education institutions (law schools) under Article XIV, Section 5(2) to determine who may be admitted to study?

3. FACULTY QUALIFICATIONS & GRADUATE PROGRAMS (WHO MAY TEACH & WHAT TO TEACH):
   Are the LEB regulations mandating a Master of Laws (LL.M.) degree for law deans/professors and restricting admission to LL.M. programs exclusively to LL.B./J.D. graduates constitutional?

4. ENCROACHMENT ON BAR ADMISSION & PRACTICE REGULATION:
   Are the provisions of R.A. No. 7662 authorizing the LEB to establish a mandatory law practice internship as a prerequisite for the Bar Examination (Sec. 7[g]) and to administer Mandatory Continuing Legal Education (Sec. 7[h]) constitutional?

================================================================================
J — JUDGMENT (DISPOSITIVE RULING)
================================================================================

The Supreme Court PARTIALLY GRANTED the petitions.

1. DECLARED CONSTITUTIONAL:
   - R.A. No. 7662 as a Whole: The creation and general regulatory jurisdiction of the Legal Education Board (LEB) over legal education is UPHELD as a valid exercise of State police power over higher education (Art. XIV, Sec. 4[1]).
   - Sections 7(c) and 7(e) of R.A. No. 7662: Valid insofar as they authorize the LEB to prescribe reasonable minimum standards for law school accreditation and admission, without infringing on institutional academic freedom.
   - Standardized Aptitude Testing in Principle: The LEB has the authority to conduct a nationwide standardized aptitude exam, BUT solely as an ADVISORY, NON-EXCLUSIONARY guide/baseline for law schools.

2. DECLARED UNCONSTITUTIONAL & NULLIFIED:
   - LEBMO No. 7-2016 in its ENTIRETY (and all related circulars): Struck down for being unconstitutional. The mandatory, exclusionary PhiLSAT requirement unlawfully deprives law schools of their academic freedom to admit applicants who fail or do not take the test.
   - Section 7(g) of R.A. No. 7662 & Section 11(g) of LEBMO No. 1-2011: The LEB's power to establish a mandatory law practice internship as a requirement for taking the Bar Examinations is NULL AND VOID for encroaching on the Supreme Court's exclusive jurisdiction over Bar admission.
   - Section 7(h) of R.A. No. 7662 & Section 11(h) of LEBMO No. 1-2011: The LEB's power to adopt and supervise Mandatory Continuing Legal Education (MCLE) for practicing lawyers is NULL AND VOID for encroaching on the Supreme Court's exclusive authority over the Integrated Bar (BM No. 850).
   - Mandatory Master of Laws (LL.M.) Requirement: Provisions in LEB issuances disqualifying law faculty or deans solely for lacking an LL.M. degree were struck down for infringing on the school's autonomy to determine "who may teach."
   - Restriction on Graduate Law Programs: Provisions prohibiting law schools from admitting non-law graduates into Master of Laws programs were struck down for violating academic freedom.

================================================================================
E — EXPLANATION & RATIO DECIDENDI
================================================================================

1. "STUDY OF LAW" VS. "PRACTICE OF LAW":
The Supreme Court clarified the constitutional boundary between executive/legislative regulation of education and judicial control of the legal profession:
  - PRACTICE OF LAW (Art. VIII, Sec. 5[5]): The Supreme Court possesses exclusive constitutional jurisdiction over admission to the Bar, legal ethics, disciplines, the Integrated Bar, and the actual practice of law. Congress cannot authorize an administrative agency to impose prerequisites for taking the Bar Exam (e.g., pre-bar internship) or to regulate practicing attorneys (e.g., MCLE).
  - STUDY OF LAW (Art. XIV, Sec. 4[1]): Legal education is part of the higher educational system. The State, through valid police power, has the sovereign authority to regulate and supervise all educational institutions to ensure quality education. Thus, Congress legitimately created the LEB to oversee law schools.

2. INSTITUTIONAL ACADEMIC FREEDOM & THE FATAL DEFECT OF PHILSAT:
Under Article XIV, Section 5(2) of the 1987 Constitution, institutions of higher learning enjoy academic freedom, which embodies four essential freedoms:
  (1) Who may teach;
  (2) What may be taught;
  (3) How it shall be taught; and
  (4) Who may be admitted to study.

The LEB's authority is confined to setting MINIMUM STANDARDS, not imposing an absolute exclusionary regime:
  - By making the PhiLSAT mandatory and dictating that no law school may admit any applicant who fails to achieve the LEB's cut-off score, the LEB completely usurped the law schools' constitutional freedom to decide for themselves "who may be admitted to study."
  - Aptitude tests may be recommended as an advisory tool or diagnostic guide, but the final, ultimate decision on student admission must rest in the sound academic discretion of the individual law schools, who evaluate candidates based on holistic factors (e.g., interviews, grades, character, background).

3. ACADEMIC FREEDOM OVER FACULTY & CURRICULUM:
  - Who May Teach: While the LEB may set basic baseline qualifications, mandating a formal Master of Laws (LL.M.) degree for all professors and deans arbitrarily disqualifies experienced judges, seasoned practitioners, and subject-matter experts who have extensive practical knowledge. Law schools have the academic freedom to choose qualified educators.
  - Graduate Programs: Higher education institutions have the freedom to craft interdisciplinary graduate curricula (e.g., LL.M. in energy, tax, or human rights) and admit professionals from other fields without being restrained by rigid LEB mandates.

4. EXCLUSIVE JUDICIAL PREROGATIVES (BAR ADMISSION & MCLE):
  - Prerequisites for the Bar Exam are strictly governed by Rule 138 of the Rules of Court under the exclusive domain of the Supreme Court. The LEB cannot add extra conditions, such as mandatory apprenticeship.
  - Continuing legal education for licensed lawyers is exclusively governed by the Supreme Court pursuant to Bar Matter No. 850 (MCLE Governing Board). The LEB's jurisdiction is strictly limited to formal law school education.

================================================================================
QUICK SUMMARY MATRIX:
--------------------------------------------------------------------------------
PROVISION / REGULATION          | RULING              | DOCTRINE / REASON
--------------------------------------------------------------------------------
Creation of LEB (R.A. 7662)     | CONSTITUTIONAL      | Valid exercise of State police power over higher education (Art. XIV, Sec. 4[1]).
Mandatory PhiLSAT (LEBMO 7-2016)| UNCONSTITUTIONAL    | Infringes on institutional academic freedom (who may be admitted to study).
Pre-Bar Internship (Sec. 7[g])  | UNCONSTITUTIONAL    | Encroaches on Supreme Court's exclusive authority over Bar admission (Art. VIII, Sec. 5[5]).
MCLE Regulation (Sec. 7[h])     | UNCONSTITUTIONAL    | Usurps Supreme Court's exclusive power over the Integrated Bar (BM No. 850).
Mandatory LL.M. for Faculty     | UNCONSTITUTIONAL    | Violates institutional academic freedom regarding who may teach.
Graduate Program Restrictions   | UNCONSTITUTIONAL    | Violates institutional academic freedom in curriculum and admissions.
================================================================================
"""

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(txt_content)
    print(f"Written text file: {txt_path}")

    # -------------------------------------------------------------
    # 2. WRITE BEAUTIFULLY FORMATTED DOCX FILE
    # -------------------------------------------------------------
    doc = docx.Document()

    # Set page margins to 1 inch (72 pt)
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Styling helper colors
    NAVY_HEX = "1E3A8A"      # Primary header / accent
    SLATE_HEX = "334155"     # Subheaders
    DARK_TEXT = "0F172A"     # Body text
    BORDER_GREY = "CBD5E1"   # Borders
    LIGHT_BG = "F1F5F9"      # Background fills
    ACCENT_BLUE = "2563EB"

    # Base Normal style setup
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = RGBColor(15, 23, 42)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)

    # Document Header Title Block
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    title_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("PIMENTEL v. LEGAL EDUCATION BOARD")
    run_title.bold = True
    run_title.font.size = Pt(18)
    run_title.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(12)
    sub_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run("CASE DIGEST & COMPREHENSIVE LEGAL ANALYSIS\nApplied Framework: BLJE (Background, Legal Issues, Judgment, Explanation)")
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(71, 85, 105)
    run_sub.italic = True

    # Metadata Callout Box / Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False

    col_widths = [Inches(2.2), Inches(4.3)]
    meta_data = [
        ("Case Title & Citations", "Oscar B. Pimentel, et al. v. Legal Education Board (LEB), G.R. No. 230642\nConsolidated with Abigail Valerie H. Aguilar, et al. v. Hon. Salvador Medialdea, et al., G.R. No. 242954"),
        ("Dates of Promulgation", "Main Decision: September 10, 2019 (Penned by J. J.C. Reyes, Jr., En Banc)\nResolution on MR: November 9, 2021 (Penned by J. R.V. Zalameda, En Banc)"),
        ("Subject Matter", "Basic Legal and Judicial Ethics (BLJE) | Constitutional Law | Separation of Powers | Institutional Academic Freedom | State Police Power"),
        ("Key Statutes & Orders", "Republic Act No. 7662 (Legal Education Reform Act of 1993)\nLEBMO No. 7-2016 (PhiLSAT), LEBMO No. 1-2011, LEBMO No. 2-2013"),
        ("Constitutional Bases", "1987 Constitution: Art. VIII, Sec. 5(5); Art. XIV, Sec. 1, 4(1), 5(2); Art. III, Sec. 1")
    ]

    for row_idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        
        cell_lbl.width = col_widths[0]
        cell_val.width = col_widths[1]
        
        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.paragraph_format.space_after = Pt(2)
        r_lbl = p_lbl.add_run(label)
        r_lbl.bold = True
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.color.rgb = RGBColor(30, 58, 138)
        
        p_val = cell_val.paragraphs[0]
        p_val.paragraph_format.space_after = Pt(2)
        r_val = p_val.add_run(val)
        r_val.font.size = Pt(9.5)
        
        # Style table cells with light gray border & fill
        shd_lbl = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{LIGHT_BG}"/>')
        cell_lbl._tc.get_or_add_tcPr().append(shd_lbl)
        
        for c in [cell_lbl, cell_val]:
            tcPr = c._tc.get_or_add_tcPr()
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

    def add_section_header(title_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title_text)
        r.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(30, 58, 138)
        
        # Add bottom border under heading paragraph
        pBdr = parse_xml(f'''
            <w:pBdr {nsdecls("w")}>
                <w:bottom w:val="single" w:sz="12" w:space="4" w:color="{NAVY_HEX}"/>
            </w:pBdr>
        ''')
        p._p.get_or_add_pPr().append(pBdr)

    def add_subsection_header(title_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title_text)
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(37, 99, 235)

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        r_pre = p.add_run(bold_prefix)
        r_pre.bold = True
        p.add_run(text)

    # -------------------------------------------------------------
    # SECTION B: BACKGROUND & WHOLE FACTS
    # -------------------------------------------------------------
    add_section_header("B — BACKGROUND & WHOLE FACTS")

    add_subsection_header("1. Legislative Enactment (Republic Act No. 7662)")
    p = doc.add_paragraph()
    p.add_run("On December 23, 1993, Congress passed Republic Act No. 7662, known as the \"Legal Education Reform Act of 1993.\" The law declared the national policy to uplift the standards of legal education, prepare law students for advocacy and judicial service, and meet the needs of a developing society.")
    
    p = doc.add_paragraph()
    p.add_run("To accomplish these objectives, R.A. No. 7662 established the Legal Education Board (LEB) as an administrative agency attached to the Department of Education, Culture and Sports (now CHED) to administer and supervise the legal education system in the Philippines. Section 7 of R.A. No. 7662 conferred upon the LEB broad regulatory powers, including:")
    
    add_bullet("Accreditation Standards (Sec. 7[c]): ", "Authority to set accreditation standards for law schools taking into account faculty qualifications and facilities.")
    add_bullet("Admission & Faculty Qualifications (Sec. 7[e]): ", "Authority to prescribe minimum requirements for admission to legal education and minimum qualifications of faculty members.")
    add_bullet("Pre-Bar Internship (Sec. 7[g]): ", "Authority to establish a law practice internship as a requirement for taking the Bar Examinations.")
    add_bullet("Mandatory Continuing Legal Education (Sec. 7[h]): ", "Authority to adopt and mandate a system of continuing legal education for practicing lawyers.")

    add_subsection_header("2. LEB Implementing Orders & The PhiLSAT Mandate")
    p = doc.add_paragraph()
    p.add_run("Pursuant to its statutory mandate, the LEB promulgated various orders and circulars:")
    
    add_bullet("LEBMO No. 1-2011 & LEBMO No. 2 (2013): ", "Prescribed comprehensive policies, standards, and curricula for legal education. It mandated that law faculty members and deans must obtain a Master of Laws (LL.M.) degree under pain of administrative disqualification, and barred law schools from admitting non-law graduates into graduate law degree programs.")
    add_bullet("LEBMO No. 7, Series of 2016 (PhiLSAT Mandate): ", "Issued on December 29, 2016, establishing the nationwide Philippine Law School Admission Test (PhiLSAT). Under Sections 1, 2, 7, 8, and 9 of LEBMO No. 7-2016:")
    
    # Nested bullet points
    p_b1 = doc.add_paragraph(style='List Bullet 2' if 'List Bullet 2' in doc.styles else 'List Bullet')
    p_b1.paragraph_format.space_after = Pt(2)
    p_b1.add_run("The PhiLSAT was made an ").font.size = Pt(10)
    p_b1.add_run("absolute, mandatory prerequisite").bold = True
    p_b1.add_run(" for admission to any basic law degree program (LL.B. or J.D.) in all Philippine law schools.")
    
    p_b2 = doc.add_paragraph(style='List Bullet 2' if 'List Bullet 2' in doc.styles else 'List Bullet')
    p_b2.paragraph_format.space_after = Pt(2)
    p_b2.add_run("Examinees were required to achieve a cut-off score (55th percentile) set unilaterally by the LEB.")
    
    p_b3 = doc.add_paragraph(style='List Bullet 2' if 'List Bullet 2' in doc.styles else 'List Bullet')
    p_b3.paragraph_format.space_after = Pt(2)
    p_b3.add_run("Law schools were ").font.size = Pt(10)
    p_b3.add_run("strictly prohibited").bold = True
    p_b3.add_run(" from enrolling students who failed or did not take the PhiLSAT, under threat of severe administrative sanctions, program closure, and accreditation revocation.")

    add_subsection_header("3. The Consolidated Petitions Before the Supreme Court")
    p = doc.add_paragraph()
    p.add_run("The constitutionality of R.A. No. 7662 and the LEB issuances was challenged through two consolidated petitions:")
    add_bullet("G.R. No. 230642 (Pimentel Petition): ", "Filed by Atty. Oscar B. Pimentel, law educators, and law students (Petition for Prohibition), asserting that R.A. No. 7662 and LEBMO No. 7-2016 encroached on judicial authority and violated institutional academic freedom.")
    add_bullet("G.R. No. 242954 (Aguilar Petition): ", "Filed by aspiring law students (Petition for Certiorari and Prohibition), arguing that the exclusionary PhiLSAT violated equal protection, due process, and the constitutional right to education.")

    add_subsection_header("4. Core Contentions of the Parties")
    p = doc.add_paragraph()
    p.add_run("Petitioners Contended: ")
    p.runs[0].bold = True
    p.add_run("(1) R.A. 7662 encroaches upon the Supreme Court's exclusive constitutional authority under Article VIII, Section 5(5) to govern admission to the practice of law; (2) The mandatory PhiLSAT violates the institutional academic freedom of law schools under Article XIV, Section 5(2) to choose who may be admitted to study; (3) The exclusionary cut-off violates the constitutional right to education and due process; and (4) The law suffers from undue delegation of legislative power.")

    p = doc.add_paragraph()
    p.add_run("Respondents Contended: ")
    p.runs[0].bold = True
    p.add_run("(1) Legal education is part of higher education and falls squarely under the State's police power to regulate educational institutions under Article XIV, Section 4(1); (2) The Supreme Court's constitutional authority applies only to the practice of law and admission to the Bar, not to the prior preparatory study of law; and (3) The PhiLSAT is a reasonable regulatory measure designed to uplift legal education standards.")

    # -------------------------------------------------------------
    # SECTION L: LEGAL ISSUES
    # -------------------------------------------------------------
    add_section_header("L — LEGAL ISSUES")

    add_bullet("Issue 1 (Separation of Powers / Judicial Authority): ", "Does the State's regulation and supervision of legal education through an administrative body (LEB) under R.A. No. 7662 unconstitutionally encroach upon the Supreme Court's exclusive authority over admission to the practice of law under Article VIII, Section 5(5) of the 1987 Constitution?")
    add_bullet("Issue 2 (Institutional Academic Freedom / Admission to Study): ", "Does the mandatory and exclusionary nature of the PhiLSAT under LEBMO No. 7-2016 violate the constitutional academic freedom of law schools under Article XIV, Section 5(2) to determine who may be admitted to study?")
    add_bullet("Issue 3 (Faculty Qualifications & Graduate Studies): ", "Are LEB issuances mandating a Master of Laws (LL.M.) degree for law professors/deans and prohibiting non-law graduates from entering LL.M. graduate programs constitutional?")
    add_bullet("Issue 4 (Bar Admission Prerequisites & MCLE): ", "Are Section 7(g) (mandatory pre-Bar internship) and Section 7(h) (mandatory continuing legal education for lawyers) of R.A. No. 7662 constitutional?")

    # -------------------------------------------------------------
    # SECTION J: JUDGMENT (DISPOSITIVE RULING)
    # -------------------------------------------------------------
    add_section_header("J — JUDGMENT (DISPOSITIVE RULING)")

    p = doc.add_paragraph()
    p.add_run("The Supreme Court ").font.size = Pt(10.5)
    r_pg = p.add_run("PARTIALLY GRANTED")
    r_pg.bold = True
    p.add_run(" the consolidated petitions, ruling as follows:")

    add_subsection_header("1. Declared CONSTITUTIONAL & UPHELD:")
    add_bullet("Creation and Jurisdiction of the LEB: ", "R.A. No. 7662 as a whole is UPHELD as a valid exercise of State police power over higher education (Art. XIV, Sec. 4[1]).")
    add_bullet("Sections 7(c) and 7(e) of R.A. No. 7662: ", "UPHELD insofar as they empower the LEB to set reasonable minimum standards for law school accreditation and admission requirements, provided institutional academic freedom is respected.")
    add_bullet("Standardized Aptitude Testing in Principle: ", "The LEB has the power to administer a nationwide aptitude examination, but solely as an ADVISORY, NON-EXCLUSIONARY baseline guide.")

    add_subsection_header("2. Declared UNCONSTITUTIONAL & STRUCK DOWN:")
    add_bullet("LEBMO No. 7-2016 in its ENTIRETY (PhiLSAT): ", "STRUCK DOWN as unconstitutional for violating institutional academic freedom. The LEB cannot make passing the PhiLSAT a mandatory, exclusionary prerequisite for law school admission.")
    add_bullet("Section 7(g) of R.A. 7662 & Sec. 11(g) of LEBMO 1-2011: ", "NULL AND VOID for usurping the Supreme Court's exclusive constitutional authority to prescribe requirements for taking the Bar Examinations.")
    add_bullet("Section 7(h) of R.A. 7662 & Sec. 11(h) of LEBMO 1-2011: ", "NULL AND VOID for encroaching upon the Supreme Court's exclusive authority over the Integrated Bar and Mandatory Continuing Legal Education (BM No. 850).")
    add_bullet("Mandatory Master of Laws (LL.M.) Requirement: ", "STRUCK DOWN for infringing on institutional academic freedom regarding who may teach.")
    add_bullet("Restriction on Graduate Legal Admissions: ", "STRUCK DOWN for infringing on law schools' autonomy to determine admission standards for specialized master's programs.")

    # -------------------------------------------------------------
    # SECTION E: EXPLANATION & RATIO DECIDENDI
    # -------------------------------------------------------------
    add_section_header("E — EXPLANATION & RATIO DECIDENDI")

    add_subsection_header("1. Distinction: 'Study of Law' vs. 'Practice of Law'")
    p = doc.add_paragraph()
    p.add_run("The Supreme Court established a fundamental constitutional demarcation between the regulation of the legal profession and the regulation of legal education:")
    add_bullet("Practice of Law (Article VIII, Section 5[5]): ", "The Supreme Court possesses exclusive constitutional jurisdiction over admission to the Bar, legal ethics, disciplines, the Integrated Bar, and the actual practice of law. Congress cannot authorize an administrative agency to prescribe prerequisites for taking the Bar Examinations (such as mandatory apprenticeships) or to supervise practicing lawyers (such as MCLE).")
    add_bullet("Study of Law (Article XIV, Section 4[1]): ", "Legal education is higher education. The State, in the exercise of its sovereign police power, has the constitutional duty to regulate and supervise all educational institutions to ensure educational quality. Hence, Congress acted within its legislative power when it created the LEB to supervise law schools.")

    add_subsection_header("2. Institutional Academic Freedom & The Fatal Defect of PhiLSAT")
    p = doc.add_paragraph()
    p.add_run("Under Article XIV, Section 5(2) of the 1987 Constitution, institutions of higher learning enjoy academic freedom. As established in landmark jurisprudence (Sweezy v. New Hampshire / Garcia v. Faculty Admission Committee), academic freedom guarantees the ")
    p.add_run("four essential freedoms:").bold = True
    
    add_bullet("(1) Who may teach; ", "The autonomy to select qualified educators.")
    add_bullet("(2) What may be taught; ", "The autonomy to design specialized and interdisciplinary curricula.")
    add_bullet("(3) How it shall be taught; ", "The autonomy to determine pedagogical methods.")
    add_bullet("(4) Who may be admitted to study; ", "The autonomy to determine criteria for student admission.")

    p = doc.add_paragraph()
    p.add_run("The Court held that the LEB's statutory mandate is strictly limited to prescribing ")
    p.add_run("minimum standards").bold = True
    p.add_run(", not imposing absolute control or complete exclusion. By making the PhiLSAT mandatory and dictating that no student may enroll without passing the test, the LEB usurped the academic freedom of law schools to admit students who demonstrate potential through other holistic metrics (e.g., undergraduate GPA, personal interviews, entrance exams, or socioeconomic background).")

    add_subsection_header("3. Faculty Qualifications & Graduate Programs")
    p = doc.add_paragraph()
    p.add_run("While the LEB may set reasonable baseline faculty standards, arbitrarily requiring a formal ")
    p.add_run("Master of Laws (LL.M.) degree").bold = True
    p.add_run(" ignores practical realities. Seasoned trial lawyers, retired judges, and renowned legal experts possess immense practical wisdom and teaching capability despite lacking an LL.M. Law schools must retain the academic autonomy to hire such experts. Furthermore, law schools have the freedom to accept non-law graduates into interdisciplinary Master's programs without rigid LEB restrictions.")

    add_subsection_header("4. Exclusive Prerogatives of the Judiciary")
    add_bullet("Bar Admission Prerequisites (Sec. 7[g]): ", "Requirements for taking the Bar Examinations are governed exclusively by Rule 138 of the Rules of Court under the sole prerogative of the Supreme Court. The LEB cannot add extra hurdles such as mandatory internships.")
    add_bullet("MCLE for Lawyers (Sec. 7[h]): ", "Mandatory Continuing Legal Education for licensed attorneys is governed exclusively by the Supreme Court pursuant to Bar Matter No. 850. The LEB's authority terminates upon the student's graduation from law school.")

    # -------------------------------------------------------------
    # QUICK REFERENCE MATRIX TABLE
    # -------------------------------------------------------------
    add_section_header("QUICK REFERENCE MATRIX TABLE")

    matrix_table = doc.add_table(rows=7, cols=3)
    matrix_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    matrix_table.autofit = False

    t_widths = [Inches(2.0), Inches(1.8), Inches(2.7)]
    headers = ["Provision / Regulation", "Supreme Court Ruling", "Core Constitutional Doctrine"]

    # Table Header Row
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
        ("Mandatory PhiLSAT (LEBMO No. 7-2016)", "STRUCK DOWN (Unconstitutional)", "Infringes on institutional academic freedom regarding 'who may be admitted to study' (Art. XIV, Sec. 5[2])."),
        ("Pre-Bar Internship (Sec. 7[g], R.A. 7662)", "STRUCK DOWN (Unconstitutional)", "Usurps Supreme Court's exclusive authority over admission to the Bar (Art. VIII, Sec. 5[5])."),
        ("MCLE Supervision (Sec. 7[h], R.A. 7662)", "STRUCK DOWN (Unconstitutional)", "Usurps Supreme Court's exclusive authority over the Integrated Bar and legal practice (BM No. 850)."),
        ("Mandatory LL.M. for Faculty", "STRUCK DOWN (Unconstitutional)", "Infringes on institutional academic freedom regarding 'who may teach' (Art. XIV, Sec. 5[2])."),
        ("Graduate Program Restrictions", "STRUCK DOWN (Unconstitutional)", "Violates academic freedom in designing curricula and admitting graduate students.")
    ]

    for row_idx, (prov, ruling, doctrine) in enumerate(matrix_rows, start=1):
        row = matrix_table.rows[row_idx]
        
        # Fill cells
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
                    r.font.color.rgb = RGBColor(22, 101, 52)
                else:
                    r.bold = True
                    r.font.color.rgb = RGBColor(185, 28, 28)
            
            # Shading for alternating rows
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

    doc.save(docx_path)
    print(f"Written Word document: {docx_path}")

if __name__ == "__main__":
    create_files()
