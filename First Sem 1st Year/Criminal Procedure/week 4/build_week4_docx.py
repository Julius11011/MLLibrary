import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def add_callout_box(doc, title, text, bg_hex="F1F5F9", border_color="3B82F6"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    # Set left border thick, others none
    tcPr = cell._element.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    run_t = p.add_run(f"💡 {title}\n")
    run_t.bold = True
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(10.5)
    run_t.font.color.rgb = RGBColor(15, 23, 42)
    
    run_b = p.add_run(text)
    run_b.font.name = "Calibri"
    run_b.font.size = Pt(10)
    run_b.font.italic = True
    run_b.font.color.rgb = RGBColor(51, 65, 85)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def create_case_digest_docx(filename):
    doc = docx.Document()
    
    # Page Setup - Margins 1 inch
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Add page numbering in footer
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("Remedial Law - Criminal Procedure (Week 4) | Page ")
        f_run.font.name = "Calibri"
        f_run.font.size = Pt(9)
        f_run.font.color.rgb = RGBColor(148, 163, 184)
        
        # XML page number
        f_pPr = f_p._element.get_or_add_pPr()
        fldSimple = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
        f_p._element.append(fldSimple)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    
    # Header Title
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title_p.add_run("REMEDIAL LAW — CRIMINAL PROCEDURE")
    t_run.bold = True
    t_run.font.name = "Calibri"
    t_run.font.size = Pt(20)
    t_run.font.color.rgb = RGBColor(15, 41, 66)
    
    subtitle_p = doc.add_paragraph()
    subtitle_p.paragraph_format.space_before = Pt(0)
    subtitle_p.paragraph_format.space_after = Pt(4)
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    st_run = subtitle_p.add_run("WEEK 4 CASE DIGESTS & COMPREHENSIVE SYLLABUS CORRELATION")
    st_run.bold = True
    st_run.font.name = "Calibri"
    st_run.font.size = Pt(13)
    st_run.font.color.rgb = RGBColor(30, 58, 138)
    
    meta_p = doc.add_paragraph()
    meta_p.paragraph_format.space_before = Pt(0)
    meta_p.paragraph_format.space_after = Pt(16)
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    m_run = meta_p.add_run("Focus: Rule 110 (Prosecution of Offenses), Control of Public Prosecutor, Territorial Authority of Prosecutors, Sufficiency of Information, and Distinctions Between Complaint & Information\nCorrelated with DOJ Circulars 15 & 28 (s. 2024), B.P. 129 / R.A. 7691 / R.A. 11576, UST Golden Notes & Dean Tan Reviewer")
    m_run.font.name = "Calibri"
    m_run.font.size = Pt(9.5)
    m_run.font.italic = True
    m_run.font.color.rgb = RGBColor(100, 116, 139)
    
    # Divider line
    div_table = doc.add_table(rows=1, cols=1)
    div_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    div_cell = div_table.cell(0, 0)
    div_cell.width = Inches(6.5)
    set_cell_background(div_cell, "1E3A8A")
    set_cell_margins(div_cell, top=20, bottom=20, left=0, right=0)
    div_cell.paragraphs[0].paragraph_format.space_before = Pt(0)
    div_cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # Cases Data
    cases = [
        {
            "num": 1,
            "title": "Wilson Chua, Renita Chua, The Secretary of Justice, and The City Prosecutor of Lucena City v. Rodrigo Padillo and Marietta Padillo",
            "citation": "G.R. No. 163797 | April 24, 2007 | 522 SCRA 128 | First Division",
            "ponente": "Justice Angelina Sandoval-Gutierrez",
            "topic": "Rule 110 (Prosecution of Offenses) — Who Must Prosecute Criminal Actions / Control of the Public Prosecutor / Executive Determination of Probable Cause / Exception to Non-Interference: Grave Abuse of Discretion by the Secretary of Justice",
            "info": [
                ("1) Who is the complainant?", 
                 "• Formal / Public Complainant: People of the Philippines (representing the State).\n• Private Complainants / Offended Parties: Spouses Rodrigo Padillo and Marietta Padillo (owners/proprietors of Padillo Lending Investor in Lucena City)."),
                ("2) What is the ground of the case filed / accusation?", 
                 "Systematic embezzlement and fraudulent diversion of approximately P7,000,000.00 from Padillo Lending Investor through the falsification of commercial checks and loan documents. Marissa Padillo-Chua (manager/cashier) fabricated loan applications for simulated borrowers, altered checks by inserting alternative payees ('OR Wilson Chua' / 'OR Renita Chua'), and diverted the funds directly into the personal bank accounts of her husband (Wilson Chua) and sister-in-law (Renita Chua)."),
                ("3) Committed crime/violation and if there is a probable cause?", 
                 "Multiple counts of Estafa through Falsification of Commercial Documents (Article 315, par. 1(b) and par. 2(a) in relation to Article 172 and Article 48 of the Revised Penal Code).\n• Probable Cause: Yes. Probable cause was established by the City Prosecutor and confirmed by the Court of Appeals and Supreme Court against Marissa Padillo-Chua, Wilson Chua, and Renita Chua. The Secretary of Justice's exclusion of Wilson and Renita was ruled a grave abuse of discretion."),
                ("4) If there is a law/s to punish / or is the case file has ground?", 
                 "Yes. Article 315 (Estafa), Article 172 (Falsification of Commercial Documents), and Article 48 (Complex Crimes) of the Revised Penal Code."),
                ("5) When the case was filed?", 
                 "The criminal complaint was filed with the Office of the City Prosecutor of Lucena City on November 19, 1999. The City Prosecutor issued the resolution finding probable cause on April 14, 2000. Informations were filed with the RTC of Lucena City in 2000. The Supreme Court promulgated its decision on April 24, 2007."),
                ("6) Place/location the case is filed?", 
                 "Office of the City Prosecutor of Lucena City / Regional Trial Court (RTC) of Lucena City, Quezon Province."),
                ("7) What court has jurisdiction in the case and why?", 
                 "Regional Trial Court (RTC) of Lucena City.\n• Subject-Matter Jurisdiction: Under B.P. Blg. 129, as amended by R.A. No. 7691 and R.A. No. 11576, first-level courts (MTC/MTCC) have jurisdiction over offenses with imprisonment not exceeding 6 years. Because the complex crime of Estafa through Falsification of Commercial Documents involving millions of pesos carries the penalty for the graver offense in its maximum period (prision mayor in its maximum period to reclusion temporal, exceeding 6 years), exclusive original jurisdiction belongs to the Regional Trial Court.\n• Territorial Jurisdiction: The loan applications were fabricated, the checks drawn and altered, and the proceeds deposited in Lucena City (locus delicti)."),
                ("8) Where will the trial be held?", 
                 "Regional Trial Court (RTC) of Lucena City, Quezon Province.")
            ],
            "facts": "Spouses Rodrigo and Marietta Padillo owned and operated a money-lending enterprise in Lucena City named 'Padillo Lending Investor.' They hired their niece, Marissa Padillo-Chua, as the firm's manager and cashier. Taking advantage of her position and the trust reposed in her by the Spouses Padillo, Marissa devised a fraudulent scheme to misappropriate business funds amounting to approximately P7,000,000.00.\n\nMarissa fabricated fictitious loan applications and promissory notes using the names of simulated borrowers. She then prepared commercial checks drawn against the bank accounts of Padillo Lending Investor (Urban Bank and Metrobank Lucena branches) payable to these fictitious borrowers. After obtaining the signature of Marietta Padillo on the checks, Marissa altered them by inserting alternative payees (adding 'OR Wilson Chua' or 'OR Renita Chua') or having them encashed. The proceeds were directly deposited into the personal bank accounts of her husband, Wilson Chua, and his sister, Renita Chua. A collector, Ernesto Alcantara, confirmed via affidavit that Wilson Chua was actively involved, monitored the check encashments, and received the stolen money.\n\nOn November 19, 1999, the Padillos filed a criminal complaint for multiple counts of Estafa through Falsification of Commercial Documents before the Office of the City Prosecutor of Lucena City. On April 14, 2000, 1st Assistant City Prosecutor Rome S. Melendres found probable cause against Marissa, Wilson, and Renita Chua. However, upon petition for review, DOJ Undersecretary Manuel A.J. Teehankee (acting for the Secretary of Justice) reversed the resolution in part by directing the prosecutor to drop and exclude Wilson Chua and Renita Chua from the Informations, claiming there was no direct evidence of conspiracy.\n\nThe Padillos filed a Rule 65 petition for certiorari with the Court of Appeals (CA). The CA granted the petition, nullified the DOJ Secretary's directives for grave abuse of discretion, and ordered the inclusion of Wilson and Renita Chua as co-accused. Wilson and Renita Chua appealed to the Supreme Court.",
            "a1_facts": "A married couple owned a money-lending shop in Lucena City and trusted their niece to manage it. The niece secretly stole about P7 million by making fake loan papers for people who didn't exist, writing company checks for those fake loans, and altering the checks so the money went straight into the personal bank accounts of her husband and sister-in-law. When the owners caught them and sued everyone for fraud and falsification, the Department of Justice tried to drop the charges against the husband and sister-in-law, claiming there was no direct proof they helped write the fake papers. The owners appealed, arguing that since the stolen money landed in their personal bank accounts, they were clearly part of the crime.",
            "issues": "Whether the Secretary of Justice committed grave abuse of discretion in directing the exclusion of Wilson Chua and Renita Chua from the criminal Informations despite the presence of strong circumstantial evidence establishing probable cause.",
            "a2_issues": "Can the Secretary of Justice refuse to file criminal charges against family members who received millions in stolen money in their bank accounts, or did the Secretary abuse his power by ignoring clear evidence of their involvement?",
            "ruling": "Yes. The Supreme Court affirmed the Court of Appeals and ruled that the Secretary of Justice committed grave abuse of discretion.\n\n1. Limits of Prosecutorial Discretion: While the determination of probable cause is an executive function and the public prosecutor and Secretary of Justice have broad discretion in the direction and control of criminal prosecutions under Section 5, Rule 110 of the Rules of Court, this discretion is not unbridled. When the Secretary of Justice exercises this power arbitrarily or in disregard of facts establishing probable cause, courts may intervene via certiorari under Rule 65.\n\n2. Standard of Probable Cause: Probable cause does not require absolute certainty or evidence establishing guilt beyond reasonable doubt; it merely requires facts and circumstances sufficient to engender a well-founded belief that a crime has been committed and that the respondent is probably guilty thereof. Conspiracy need not be proved by direct evidence but may be inferred from the collective acts and circumstances showing a community of criminal design.\n\n3. Overwhelming Circumstantial Evidence: The following circumstances clearly established probable cause against Wilson Chua and Renita Chua:\n(a) The altered checks containing the embezzled funds were directly deposited into their personal bank accounts;\n(b) Wilson Chua lived with Marissa as husband and wife and shared their financial household;\n(c) The affidavit of collector Ernesto Alcantara showed Wilson's active participation and knowledge; and\n(d) Neither Wilson nor Renita offered any legitimate commercial explanation for receiving millions of pesos from Padillo Lending Investor's accounts.\n\nTherefore, the Secretary of Justice acted with grave abuse of discretion, and both Wilson Chua and Renita Chua must be included as co-accused in the Informations filed before the RTC of Lucena City.",
            "a3_ruling": "Yes. The Supreme Court ruled that the Secretary of Justice made a serious legal error (grave abuse of discretion). To charge someone in court, the prosecutor does not need 100% proof of guilt—only 'probable cause,' meaning a reasonable belief that they took part in the crime. Here, the stolen money went directly into the bank accounts of the husband and sister-in-law, and a witness confirmed the husband knew about the scheme. People who receive stolen money without any legal explanation cannot claim to be innocent bystanders. Therefore, the court ordered that the husband and sister-in-law must be charged and tried in court alongside the niece.",
            "doctrine": "• Rule 110, Section 5 & Rule 112, Section 4 (Control of Prosecution & Executive Determination of Probable Cause): Public prosecutors and the Secretary of Justice possess executive authority to determine whether to file a criminal case. However, this power is subject to judicial review under Rule 65 when exercised with grave abuse of discretion.\n• DOJ Department Circular No. 015 (s. 2024) and DOJ Circular No. 028 (s. 2024) Integration: Under modern DOJ rules, prosecutors must evaluate complaints under the 'reasonable certainty of conviction' standard based on admissible evidence. In evaluating conspiracy and financial crimes, prosecutors must account for circumstantial evidence and money trails (e.g., bank deposits and unexplained receipt of funds). If an executive officer arbitrarily ignores established financial evidence, the ruling is vulnerable to judicial nullification.\n• Proof Required at Inquest/Preliminary Investigation: Preliminary investigation is not a trial. Probable cause requires only reasonable grounds of suspicion supported by circumstances sufficiently strong in themselves to warrant a cautious man in the belief that the person accused is guilty of the offense.",
            "a4_doctrine": "While government prosecutors have the power to decide who to charge, they cannot arbitrarily protect or drop people when clear evidence (like bank records) connects them to the crime. If the prosecutor or Secretary of Justice ignores obvious evidence, the courts will step in and order the charges filed.",
            "relevance": "Under the syllabus topic 'Rule 110 — Who Must Prosecute Criminal Actions / Control of the Public Prosecutor / Review of Probable Cause,' this case establishes the crucial boundary between executive prosecutorial discretion and judicial review. It illustrates that while the executive branch controls the institution of criminal actions, that power cannot be used as an arbitrary shield to drop co-conspirators when the evidence demonstrates probable cause.",
            "a5_relevance": "Shows the limits of a prosecutor's power. It proves that prosecutorial discretion is not absolute, and victims of crimes can ask the courts to force prosecutors to file charges against all guilty parties if the prosecutor acted with grave abuse of discretion."
        },
        {
            "num": 2,
            "title": "Bureau of Customs v. Peter Sherman, Michael Whelan, Teodoro B. Lingan, Atty. Ofelia B. Cajigal, and the Court of Tax Appeals",
            "citation": "G.R. No. 190487 | April 13, 2011 | 648 SCRA 615 | Third Division",
            "ponente": "Justice Conchita Carpio Morales",
            "topic": "Rule 110 (Prosecution of Offenses) — Who Must Prosecute Criminal Actions / Control of the Public Prosecutor / Representation of the Government by the Office of the Solicitor General (OSG) / Nominal Complainant vs. Public Prosecutor / Withdrawal of Information",
            "info": [
                ("1) Who is the complainant?", 
                 "• Formal / Public Complainant: People of the Philippines (represented by the Public Prosecutor and Office of the Solicitor General).\n• Nominal / Initiating Complainant: Bureau of Customs (BOC), represented by the Commissioner of Customs (acting through its 'Run After The Smugglers' / RATS legal officers)."),
                ("2) What is the ground of the case filed / accusation?", 
                 "Alleged unlawful importation and fraudulent entry of finished printed bet slips and thermal rolls from Australia into the Philippines (transferred from Clark Special Economic Zone to the Philippine Charity Sweepstakes Office) without payment of customs duties and value-added taxes totaling P45,862,693.00."),
                ("3) Committed crime/violation and if there is a probable cause?", 
                 "Unlawful Importation / Fraudulent Practices Against Customs Revenue under Section 3601 in relation to Section 3602 of the Tariff and Customs Code of the Philippines (TCCP).\n• Probable Cause: Probable cause was initially found by the State Prosecutor, but upon review, Secretary of Justice Raul M. Gonzalez found NO probable cause (holding that MSPI was an exempt CSEZ enterprise under R.A. 7227). The Court of Tax Appeals (CTA) independently affirmed the absence of probable cause and granted the withdrawal of the Information."),
                ("4) If there is a law/s to punish / or is the case file has ground?", 
                 "Sections 3601 and 3602 of the Tariff and Customs Code of the Philippines (TCCP), as amended."),
                ("5) When the case was filed?", 
                 "The BOC filed its complaint-affidavit in 2007. The State Prosecutor issued the resolution finding probable cause on January 3, 2008. The Information was filed before the CTA in 2008 (CTA Crim. Case No. O-116). The CTA granted the motion to withdraw the Information on July 28, 2009. The Supreme Court decided the petition on April 13, 2011."),
                ("6) Place/location the case is filed?", 
                 "Department of Justice / Court of Tax Appeals (CTA) First Division, Diliman, Quezon City."),
                ("7) What court has jurisdiction in the case and why?", 
                 "Court of Tax Appeals (CTA).\n• Subject-Matter Jurisdiction: Under R.A. No. 1125, as amended by R.A. No. 9282, the Court of Tax Appeals exercises exclusive original jurisdiction in all criminal offenses arising from violations of the Tariff and Customs Code and National Internal Revenue Code where the principal amount of taxes and fees, exclusive of charges and penalties, is P1,000,000.00 or more. Here, the disputed customs duties and taxes amounted to P45,862,693.00, well above the statutory threshold."),
                ("8) Where will the trial be held?", 
                 "Court of Tax Appeals (CTA), Quezon City.")
            ],
            "facts": "Mark Sensing Philippines, Inc. (MSPI) was an enterprise registered with the Clark Development Corporation (CDC) inside the Clark Special Economic Zone (CSEZ). MSPI was contracted by the Philippine Charity Sweepstakes Office (PCSO) to print on-line lottery bet slips and thermal paper rolls. From June 2005 to January 2007, MSPI imported printed bet slips and thermal rolls from its parent company in Australia. The shipments entered the country through the Port of Manila and Subic, and were transferred to CSEZ under official transshipment permits before final delivery to PCSO.\n\nThe Bureau of Customs (BOC), through its 'Run After The Smugglers' (RATS) Program, alleged that MSPI failed to pay customs duties and value-added taxes totaling P45,862,693.00, claiming that the bet slips were finished consumer goods that did not qualify for tax incentives. The BOC filed a criminal complaint with the DOJ against MSPI officers Peter Sherman, Michael Whelan, Teodoro B. Lingan, Atty. Ofelia B. Cajigal, and customs brokers for violation of Sections 3601 and 3602 of the TCCP.\n\nState Prosecutor Rohairah A. Lao initially found probable cause and filed an Information before the Court of Tax Appeals (CTA Crim. Case No. O-116). However, the respondents filed a Petition for Review with the Secretary of Justice. On June 19, 2008, Secretary of Justice Raul M. Gonzalez reversed the State Prosecutor, holding that MSPI, as a CSEZ enterprise, was legally exempt from customs duties under R.A. 7227, and that all importations were accompanied by valid transshipment permits approved by the BOC itself. The Secretary of Justice ordered the State Prosecutor to immediately withdraw the Information.\n\nPursuant to this directive, the State Prosecutor filed a Motion to Withdraw Information before the CTA First Division. The BOC, represented solely by its RATS legal officers (without the OSG), opposed the motion. On July 28, 2009, the CTA granted the motion to withdraw and dismissed the criminal case, finding upon its own independent evaluation of the evidence that there was no probable cause to hold the accused for trial.\n\nThe BOC, acting through its in-house RATS lawyers without the participation or conformity of the Office of the Solicitor General (OSG), filed a Rule 65 petition for certiorari directly with the Supreme Court.",
            "a1_facts": "A company in the Clark Special Economic Zone imported lottery paper and bet slips for the PCSO. The Bureau of Customs claimed the company smuggled the goods without paying P45.8 million in customs taxes, and sued the company directors. A prosecutor initially filed charges in the Court of Tax Appeals. However, the Secretary of Justice reviewed the case, found that the company had valid tax-free permits under the Clark ecozone law, and ordered the prosecutor to drop the case. The prosecutor asked the court to withdraw the charges, and the court agreed and dismissed the case. The Bureau of Customs got upset and filed an appeal directly in the Supreme Court using only its own in-house lawyers, without getting permission or help from the government's official top lawyer (the Solicitor General).",
            "issues": "1. Whether the Bureau of Customs (as a government agency) has the legal personality to file a petition for certiorari in a criminal case before the Supreme Court without the representation or authorization of the Office of the Solicitor General (OSG).\n2. Whether the Court of Tax Appeals committed grave abuse of discretion in granting the public prosecutor's Motion to Withdraw Information.",
            "a2_issues": "1. Can a government bureau (like the Bureau of Customs) file a criminal appeal in the Supreme Court by itself, or does the law require the Solicitor General to represent it?\n2. Did the tax court make an illegal mistake when it agreed with the prosecutor to withdraw and dismiss the criminal charges?",
            "ruling": "No on both counts. The Supreme Court dismissed the BOC's petition.\n\n1. Lack of Legal Personality and Mandatory OSG Representation: Under Section 35(1), Chapter 12, Title III, Book IV of the 1987 Administrative Code (E.O. 292), the Office of the Solicitor General (OSG) is the sole legal representative of the Government of the Philippines, its agencies, and officials in all appellate proceedings before the Court of Appeals and the Supreme Court. In criminal proceedings, the State is the real party in interest, represented by the OSG. The Bureau of Customs is merely a government bureau acting as the nominal or private complainant. A private complainant or government agency has no standing to assail the dismissal of a criminal case or file a petition for certiorari without the authorization, conformity, and active representation of the OSG. The petition filed solely by the BOC RATS lawyers suffered from a fatal procedural defect.\n\n2. Prosecutorial Control and Independent Judicial Determination: Under Section 5, Rule 110 of the Rules of Court, all criminal actions are under the direction and control of the public prosecutor. When the Secretary of Justice directs the withdrawal of an Information, the prosecutor is bound to follow. Once a motion to withdraw is filed, the court must make its own independent assessment of the merits of the case (Crespo v. Mogul).\n\nHere, the CTA First Division did not merely rubber-stamp the prosecutor's motion; it conducted its own independent evaluation of the records and correctly concluded that MSPI operated under valid tax-exempt CSEZ permits under R.A. 7227, completely negating any fraudulent intent or probable cause for smuggling. The CTA acted well within its jurisdiction.",
            "a3_ruling": "No. The Supreme Court threw out the Bureau of Customs' appeal for two major reasons:\n1. The Bureau of Customs had no legal right to file the appeal on its own. By law, only the Office of the Solicitor General (the chief lawyer of the Philippine government) can represent government agencies in criminal cases before the Supreme Court.\n2. In criminal cases, the public prosecutor controls the prosecution. When the prosecutor moved to withdraw the case and the tax court independently checked the evidence and agreed that no crime was committed, the court acted properly.",
            "doctrine": "• Rule 110, Section 5 (Control of Prosecution): Criminal actions are instituted on behalf of the People of the Philippines and are under the exclusive control of the public prosecutor. Private complainants and government bureaus cannot usurp the prosecutorial functions of the State.\n• Administrative Code of 1987 (Mandatory OSG Representation): In all appellate proceedings (Rule 45, Rule 65) before the Court of Appeals and Supreme Court concerning criminal cases, the State must be represented exclusively by the Solicitor General. Any petition filed by private counsel or in-house agency lawyers without OSG deputation is dismissible outright.\n• Independent Assessment Rule (Crespo v. Mogul): Once an Information is filed in court, any disposition thereof (dismissal, withdrawal, or amendment) rests in the sound discretion of the court. The court cannot blindly rely on the executive findings of the DOJ Secretary, but must perform an independent review.\n• DOJ Circulars No. 015 & 028 (s. 2024) Alignment: In current practice, motions to withdraw informations must be supported by certified findings of absence of reasonable certainty of conviction, preserving the court's prerogative to evaluate the evidence independently.",
            "a4_doctrine": "Only the public prosecutor controls a criminal lawsuit, and only the Solicitor General can represent the government in the Supreme Court. A government agency cannot act like a private rogue litigant; if the prosecutor and the judge agree to drop a case because there is no crime, the agency cannot overturn that decision on its own.",
            "relevance": "Under the syllabus topic 'Rule 110 — Who Must Prosecute Criminal Actions / Control of Prosecution / Role of the OSG,' this case demonstrates that:\n(1) Government agencies acting as complainants in criminal cases are subject to the same procedural hierarchy as private offended parties;\n(2) Control of prosecution belongs strictly to the public prosecutor at trial and the OSG on appeal; and\n(3) The trial court possesses the final judicial authority to grant or deny the withdrawal of an Information upon independent assessment.",
            "a5_relevance": "Clarifies who has the legal power to handle criminal cases in court: the trial prosecutor controls the case at the trial level, the Solicitor General controls it at the Supreme Court level, and a government department cannot bypass these official lawyers."
        },
        {
            "num": 3,
            "title": "People of the Philippines v. Gualberto Cinco y Solloza (Soyosa)",
            "citation": "G.R. No. 186460 | December 4, 2009 | 607 SCRA 739 | Second Division",
            "ponente": "Justice Antonio T. Carpio",
            "topic": "Rule 110 (Prosecution of Offenses) — Sufficiency of Complaint or Information / Date and Time of Commission of Offense / Material Ingredients vs. Gravamen of Offense / Right to be Informed of Nature and Cause of Accusation",
            "info": [
                ("1) Who is the complainant?", 
                 "• Formal / Public Complainant: People of the Philippines (representing the State).\n• Private Complainant / Offended Party: 'AAA', a 14-year-old minor victim, assisted by her mother."),
                ("2) What is the ground of the case filed / accusation?", 
                 "The accused, Gualberto Cinco y Solloza (live-in partner of the victim's maternal aunt), committed two counts of rape against 14-year-old AAA by taking advantage of being alone with her in their residence and having carnal knowledge of her through force, threat of violence, and intimidation."),
                ("3) Committed crime/violation and if there is a probable cause?", 
                 "Two (2) counts of Simple Rape under Article 266-A, paragraph 1(a) of the Revised Penal Code.\n• Probable Cause: Yes. Probable cause was found by the investigating prosecutor and affirmed by the trial court, the Court of Appeals, and the Supreme Court."),
                ("4) If there is a law/s to punish / or is the case file has ground?", 
                 "Article 266-A, paragraph 1(a) and Article 266-B of the Revised Penal Code (as amended by R.A. No. 8353, The Anti-Rape Law of 1997)."),
                ("5) When the case was filed?", 
                 "The two Informations were filed before the Regional Trial Court of Quezon City on October 5, 2001 (Criminal Case Nos. Q-01-100233 and Q-01-100234). The Supreme Court decided the appeal on December 4, 2009."),
                ("6) Place/location the case is filed?", 
                 "Regional Trial Court (RTC) of Quezon City, Branch 104 (Family Court)."),
                ("7) What court has jurisdiction in the case and why?", 
                 "Regional Trial Court (RTC) of Quezon City (Family Court).\n• Subject-Matter Jurisdiction: Under B.P. Blg. 129, as amended, and Republic Act No. 8369 (The Family Courts Act of 1997), the Regional Trial Court designated as a Family Court has exclusive original jurisdiction over criminal cases punishable by reclusion perpetua (such as rape) and all criminal offenses where one of the victims is a child/minor.\n• Territorial Jurisdiction: The criminal acts were committed within the territorial boundaries of Quezon City."),
                ("8) Where will the trial be held?", 
                 "Regional Trial Court (RTC) of Quezon City, Branch 104.")
            ],
            "facts": "In October 2000, 14-year-old 'AAA' was living in Quezon City with her maternal aunt, Maritess, and Maritess' live-in partner, accused-appellant Gualberto Cinco y Solloza. On a morning in October 2000, while Maritess was away at work, Cinco approached AAA, dragged her into the bedroom, forcibly undressed her, pinned her down, and had carnal knowledge of her against her will while threatening to kill her if she told anyone.\n\nOn November 19, 2000, Cinco repeated the sexual assault under similar circumstances when Maritess again left the house. In September 2001, AAA finally disclosed the ordeal to her mother. A physical examination by a PNP medico-legal officer revealed healed deep hymenal lacerations consistent with sexual intercourse.\n\nTwo Informations were filed charging Cinco with two counts of Simple Rape before the RTC of Quezon City (Criminal Case Nos. Q-01-100233 and Q-01-100234).\nThe first Information alleged: 'That on or about the month of October 2000, in Quezon City, Philippines, the above-named accused, by means of force, violence and intimidation, did then and there willfully, unlawfully and feloniously have carnal knowledge of the undersigned AAA, a minor, 14 years of age, against her will...'\nThe second Information alleged: 'That on or about November 19, 2000, in Quezon City, Philippines, the said accused, by means of force, violence and intimidation, did then and there willfully, unlawfully and feloniously have carnal knowledge of the undersigned AAA, a minor, 14 years of age, against her will...'\n\nAt the trial, Cinco raised the defense of denial and alibi, testifying that he worked as a driver and was never alone with the victim. Crucially, Cinco argued that the Informations were fatally defective because the first Information stated 'on or about the month of October 2000' without giving the exact day or time, which allegedly violated his constitutional right to be informed of the nature and cause of the accusation and prevented him from proving an alibi.\n\nThe RTC convicted Cinco of two counts of Simple Rape and sentenced him to reclusion perpetua for each count. The Court of Appeals affirmed the conviction in toto. Cinco appealed to the Supreme Court.",
            "a1_facts": "A 14-year-old girl was living with her aunt and her aunt's live-in partner in Quezon City. While the aunt was away at work, the partner raped the young girl twice (once in October 2000 and once on November 19, 2000) while threatening to kill her. When she finally told her mother, medical exams confirmed the sexual abuse, and the man was charged with two counts of rape. In his defense, the man claimed he didn't do it and argued that the charge sheet (Information) was invalid because it only said the crime happened 'in the month of October 2000' without listing the exact day and time, claiming this made it impossible for him to provide an alibi.",
            "issues": "Whether the Information charging the accused with rape was fatally defective for failing to allege the precise date and time of the commission of the offense.",
            "a2_issues": "Is a criminal charge for rape invalid if the prosecutor only writes the month and year (like 'October 2000') instead of the exact date and time?",
            "ruling": "No. The Supreme Court affirmed the conviction of Gualberto Cinco y Solloza for two counts of Simple Rape.\n\n1. Date and Time Are Not Material Ingredients in Rape: Under Section 11, Rule 110 of the Revised Rules of Criminal Procedure:\n'It is not necessary to state in the complaint or information the precise date the offense was committed except when it is a material ingredient of the offense. The offense may be alleged to have been committed on a date as near to the actual date at which the offense was committed as the complaint or information will permit.'\nThe gravamen of the crime of rape is carnal knowledge of a woman through force, threat, or intimidation (or where the victim is under 12 or deprived of reason). The exact date and time are not essential elements of the crime. An Information that states the offense occurred 'on or about the month of October 2000' or 'on or about November 19, 2000' is legally sufficient.\n\n2. Right to Be Informed of the Accusation: An Information satisfies the constitutional requirement under Section 6, Rule 110 if it states the name of the accused, the designation of the offense, the acts or omissions complained of as constituting the offense, the name of the offended party, the approximate date of commission, and the place of commission. The Information here enabled a person of common understanding to know what offense was intended to be charged.\n\n3. Waiver and Bill of Particulars: If the accused felt that the allegation of time was too indefinite to prepare a defense, his proper procedural remedy was to file a Motion for a Bill of Particulars under Section 9, Rule 116, or a Motion to Quash under Rule 117 before entering his plea. By entering a plea of not guilty and participating in the trial without objection, the accused waived any formal defect in the Information.",
            "a3_ruling": "No. The Supreme Court held that the charge sheet was completely valid. In rape cases, the exact day and time are not essential parts of the crime; what matters is whether the accused forced himself upon the victim. Under the Rules of Court (Rule 110, Section 11), saying the crime happened 'around October 2000' is specific enough. Furthermore, if the accused wanted more specific details, he should have asked the judge for a 'Bill of Particulars' before his arraignment. Because he pleaded 'not guilty' and went through the whole trial without objecting to the wording, he legally gave up any right to complain about it later.",
            "doctrine": "• Rule 110, Section 6 (Sufficiency of Complaint or Information): An Information is sufficient if it states the acts or omissions in ordinary and concise language sufficient to enable a person of common understanding to know the offense charged.\n• Rule 110, Section 11 (Date of Commission of Offense): The precise date is NOT required in the Information UNLESS the date itself is a material element of the crime (e.g., violation of Sunday election bans, infanticide committed within 72 hours of birth). For general crimes such as rape, homicide, or theft, using the phrase 'on or about' a month and year satisfies the law.\n• Rule 116, Section 9 (Bill of Particulars) & Rule 117 (Waiver): Any ambiguity regarding the time, place, or specific circumstances must be attacked through a motion for bill of particulars or motion to quash prior to arraignment. Pleading to the charge cures and waives any non-jurisdictional ambiguity.\n• DOJ Circulars 15 & 28 (s. 2024) Drafting Rules: Directs prosecutors to formulate Informations with clarity, setting forth the timeframe with as much proximity as the evidence allows, while preserving prosecutorial validity against hyper-technical objections.",
            "a4_doctrine": "Charge sheets do not need to list the exact hour or day unless the date itself makes the act illegal. For crimes like rape, giving the approximate month and year is enough. If an accused person believes the charge is too vague, they must object before entering a plea; otherwise, they waive their objection.",
            "relevance": "Under the syllabus topic 'Rule 110 — Sufficiency of Complaint or Information / Date of Commission of Offense,' this case provides the standard interpretation of Section 6 and Section 11 of Rule 110. It establishes that technical imprecision in the date does not invalidate an Information where the essential elements (gravamen) of the felony are clearly described.",
            "a5_relevance": "Teaches law students how to determine if a criminal Information is legally sufficient under Rule 110, proving that approximate dates are valid and that procedural objections must be raised before entering a plea."
        },
        {
            "num": 4,
            "title": "Visitacion L. Estodillo and Jovelyn Estudillo v. Judge Teofilo D. Baluma",
            "citation": "A.M. No. RTJ-04-1837 | March 23, 2004 | 426 SCRA 1 | Second Division",
            "ponente": "Justice Romeo J. Callejo, Sr.",
            "topic": "Rule 110 (Prosecution of Offenses) — Distinctions Between Complaint and Information / Formal Requisites of an Information (Subscription under Official Oath vs. Sworn Statement) / Rule 110 Sections 3 and 4 / Gross Ignorance of Basic Law",
            "info": [
                ("1) Who is the complainant?", 
                 "• Complainants in the Administrative Case: Visitacion L. Estodillo and her minor daughter, Jovelyn Estudillo.\n• Formal / Public Complainant in Underlying Criminal Case: People of the Philippines, upon complaint of Jovelyn Estudillo."),
                ("2) What is the ground of the case filed / accusation?", 
                 "• Administrative Charge: Gross and Inexcusable Ignorance of the Law and Grave Abuse of Authority against Judge Teofilo D. Baluma (Presiding Judge, RTC Bohol, Branch 1) for dismissively throwing out a criminal Information for child abuse on the legally erroneous ground that the Information was 'not subscribed and sworn to' by the prosecutor.\n• Underlying Criminal Charge: Accused Fredie Cirilo Nocos y Urot committed acts of child abuse against 13-year-old Jovelyn Estudillo."),
                ("3) Committed crime/violation and if there is a probable cause?", 
                 "• Administrative Violation: Gross Ignorance of the Law and Procedure (Rule 140 of the Rules of Court / Code of Judicial Conduct).\n• Underlying Criminal Violation: Violation of Section 10(a) of Republic Act No. 7610 (Special Protection of Children Against Child Abuse, Exploitation and Discrimination Act). Probable cause was duly established by the City Prosecutor."),
                ("4) If there is a law/s to punish / or is the case file has ground?", 
                 "Rule 140 of the Rules of Court (Discipline of Judges); Section 3 & Section 4 of Rule 110 of the Revised Rules of Criminal Procedure; Section 10(a) of R.A. 7610."),
                ("5) When the case was filed?", 
                 "The criminal Information was filed on August 21, 2002. Judge Baluma dismissed it on August 27, 2002. The administrative complaint was filed in October 2002. The Supreme Court promulgated its decision on March 23, 2004."),
                ("6) Place/location the case is filed?", 
                 "Regional Trial Court (RTC) of Bohol, Branch 1, Tagbilaran City (Family Court) / Supreme Court of the Philippines (Administrative Matter)."),
                ("7) What court has jurisdiction in the case and why?", 
                 "Supreme Court of the Philippines (Administrative Jurisdiction).\n• Under Article VIII, Section 6 of the 1987 Constitution, the Supreme Court has exclusive administrative supervision and disciplinary power over all courts and judges.\n• Underlying Criminal Jurisdiction: RTC of Bohol, Branch 1 (Family Court under R.A. 8369)."),
                ("8) Where will the trial be held?", 
                 "Supreme Court of the Philippines (Administrative Matter) / RTC of Bohol, Branch 1, Tagbilaran City (Criminal Case No. 11627).")
            ],
            "facts": "On August 21, 2002, 2nd Assistant City Prosecutor Sisinio C. Virtudazo of Tagbilaran City filed a criminal Information for 'Other Acts of Child Abuse' under Section 10(a) of R.A. 7610 against Fredie Cirilo Nocos y Urot before the Regional Trial Court of Bohol, Branch 1 (designated Family Court), presided over by respondent Judge Teofilo D. Baluma (Criminal Case No. 11627). The criminal case stemmed from a sworn complaint filed by 13-year-old Jovelyn Estudillo and her mother Visitacion L. Estodillo.\n\nThe Information was signed ('subscribed') by Assistant City Prosecutor Virtudazo with the prior written approval of City Prosecutor Rodolfo R. Ligason. At the bottom of the Information, Prosecutor Virtudazo executed the standard certification under oath that a preliminary investigation had been conducted in accordance with law and that there was reasonable ground to believe that the crime was committed and the accused was probably guilty.\n\nOn August 27, 2002, without any motion from the accused and before arraignment, Judge Teofilo D. Baluma motu proprio issued an Order dismissing Criminal Case No. 11627 on the ground that the Information was 'not subscribed and sworn to by the Assistant City Prosecutor,' declaring the Information fatally defective and null and void.\n\nProsecutor Virtudazo filed a Motion for Reconsideration, citing Section 4, Rule 110 of the Rules of Court which explicitly provides that an Information only needs to be subscribed by the prosecutor and is not required to be sworn to. Judge Baluma subsequently granted the reconsideration and reinstated the case.\n\nHowever, the private complainants, Visitacion L. Estodillo and Jovelyn Estudillo, filed an administrative complaint before the Office of the Court Administrator (OCA) charging Judge Baluma with Gross and Inexcusable Ignorance of the Law and Grave Abuse of Authority.",
            "a1_facts": "A mother and her 13-year-old daughter filed a child abuse case against a man in Tagbilaran City. The prosecutor investigated the case, signed the formal charge sheet (Information), certified under oath that there was probable cause, and filed it in the Family Court. However, the trial judge immediately threw out and dismissed the criminal case on his own, claiming that the Information was completely void because the prosecutor did not have the main text of the charge sheet 'sworn to under oath.' The prosecutor had to point out the basic rule book to the judge, showing that only a private 'complaint' must be sworn to, while a prosecutor's 'Information' only needs to be signed. The mother and daughter then filed an administrative complaint against the judge for not knowing elementary law.",
            "issues": "Whether an Information filed in court by a public prosecutor is required to be subscribed and sworn to under oath, and whether Judge Baluma's dismissal of the criminal case on that ground constituted Gross Ignorance of the Law.",
            "a2_issues": "Does a criminal charge sheet (Information) signed by a government prosecutor have to be sworn under oath like a private complaint, and is a judge guilty of gross ignorance of the law for dismissing a case because it wasn't sworn to?",
            "ruling": "No (an Information does not need to be sworn to), and Yes (Judge Baluma was guilty of Gross Ignorance of the Law). The Supreme Court reprimanded Judge Baluma with a stern warning.\n\n1. Distinction Between Complaint and Information (Rule 110, Sections 3 & 4):\nThe Supreme Court underscored the elementary distinction between a Complaint and an Information:\n• Complaint (Section 3, Rule 110): 'A complaint is a sworn written statement charging a person with an offense, subscribed by the offended party, any peace officer, or other public officer charged with the enforcement of the law violated.'\n• Information (Section 4, Rule 110): 'An information is an accusation in writing charging a person with an offense, subscribed by the prosecutor and filed with the court.'\n\n2. Why an Information Does NOT Need an Oath: An Information is filed by a public prosecutor who is an officer of the law acting under the solemn responsibility of his official oath of office. The prosecutor's official oath accompanies every official act he performs. Therefore, there is NO statutory or procedural requirement that the Information itself be sworn to. The certification executed by the prosecutor regarding the conduct of a preliminary investigation is sufficient.\n\n3. Gross Ignorance of the Law: The distinction between a complaint and an information is an elementary rule of criminal procedure taught in basic law studies. When the law is so elementary, not to know it or to act as if one does not know it constitutes gross ignorance of the law. A judge is expected to exhibit more than just a cursory acquaintance with the basic Rules of Court. Judge Baluma's erroneous dismissal caused undue delay in a child abuse prosecution and undermined public confidence in the judiciary.",
            "a3_ruling": "No, an Information does NOT need to be sworn to under oath, and Yes, the judge was guilty of gross ignorance of the law.\nThe Supreme Court explained that:\n1. A private 'Complaint' (Rule 110, Section 3) must be sworn under oath by the victim or police officer.\n2. A prosecutor's 'Information' (Rule 110, Section 4) only needs to be signed ('subscribed'). The prosecutor does not need to swear to it every time because they already took a solemn oath of office when they became a government prosecutor.\nConfusing an Information with a Complaint is a first-year law school error. The Supreme Court reprimanded the judge and warned him that repeating such basic mistakes would lead to severe penalties.",
            "doctrine": "• Rule 110, Section 3 vs. Section 4 (The Definitive Distinction):\n  - Form: Complaint is a sworn written statement; Information is an accusation in writing (need not be sworn).\n  - Subscriber: Complaint is subscribed by offended party, peace officer, or enforcing officer; Information is subscribed by the Public Prosecutor.\n  - Target: Complaint is filed in Prosecutor's Office or MTC; Information is filed directly in the Trial Court having jurisdiction.\n  - Oath: Complaint requires separate jurat/oath; Information is protected by the prosecutor's official oath of office.\n• Rule 112, Section 4 (Prosecutor's Certification): The certification at the foot of an Information that preliminary investigation was conducted and probable cause exists does not transform the Information into a sworn complaint.\n• DOJ Circular No. 015 & Circular No. 028 (s. 2024) Standard Forms: All standard Information templates issued by the National Prosecution Service require only the subscription and certification of the investigating prosecutor and the written approval of the Head of the Prosecution Office.\n• Judicial Competence (Canon 3, Code of Judicial Conduct): Judges are duty-bound to master elementary procedural rules to avoid aborting valid criminal prosecutions.",
            "a4_doctrine": "An Information filed by a prosecutor only requires a signature, not a notarized oath, because the prosecutor is an official acting under their government oath. A judge cannot throw out a case because of a non-existent oath requirement.",
            "relevance": "Under the syllabus topic 'Rule 110 — Distinction between Complaint and Information / Requisites of Information,' this case serves as the leading administrative authority illustrating the statutory and practical differences between a complaint under Section 3 and an information under Section 4 of Rule 110.",
            "a5_relevance": "Provides the foundational rule on the difference between a Complaint and an Information, emphasizing that a prosecutor's charge sheet needs only a signature, not a sworn oath."
        },
        {
            "num": 5,
            "title": "Renato Cudia v. The Court of Appeals, The Hon. Carlos D. Rustia, Presiding Judge of RTC, Branch 56, Angeles City, and People of the Philippines",
            "citation": "G.R. No. 110315 | January 16, 1998 | 284 SCRA 173 | Third Division",
            "ponente": "Justice Jose C. Vitug (Romero, J., Acting Chairman)",
            "topic": "Rule 110 (Prosecution of Offenses) — Who Must Prosecute Criminal Actions / Authority of Prosecutor to File Information / Territorial Authority of Prosecutors / Void Information / Requisites of Double Jeopardy (Rule 117, Section 7)",
            "info": [
                ("1) Who is the complainant?", 
                 "• Formal / Public Complainant: People of the Philippines (representing the sovereign State).\n• Initiating Officers / Complainants: Arresting police officers of the Philippine National Police (PNP) / 173rd PC Company."),
                ("2) What is the ground of the case filed / accusation?", 
                 "The accused, Renato Cudia, was apprehended in possession of an unlicensed .38 caliber revolver with serial number and live ammunition without any permit or legal authority to possess the same."),
                ("3) Committed crime/violation and if there is a probable cause?", 
                 "Illegal Possession of Firearms and Ammunition under Presidential Decree No. 1866.\n• Probable Cause: Yes, probable cause was established. However, the first Information filed was legally void because the investigating prosecutor lacked territorial authority to file an Information for a crime committed in a different municipality outside his city."),
                ("4) If there is a law/s to punish / or is the case file has ground?", 
                 "Presidential Decree No. 1866 (Codifying the Laws on Illegal/Unlawful Possession, Manufacture, Dealing In, Acquisition or Disposition of Firearms, Ammunition or Explosives)."),
                ("5) When the case was filed?", 
                 "The first Information was filed by the City Prosecutor of Angeles City on June 28, 1989 (Criminal Case No. 11542, RTC Angeles City). The second Information was filed by the Provincial Prosecutor of Pampanga on April 22, 1991 (Criminal Case No. 11987, RTC San Fernando, Pampanga). The Supreme Court promulgated its decision on January 16, 1998."),
                ("6) Place/location the case is filed?", 
                 "Originally filed with the Regional Trial Court (RTC) of Angeles City, Branch 56 (incorrect venue / lack of territorial jurisdiction); subsequently re-filed with the Regional Trial Court (RTC) of San Fernando, Pampanga, Branch 46 (proper venue)."),
                ("7) What court has jurisdiction in the case and why?", 
                 "Regional Trial Court (RTC) of San Fernando, Pampanga.\n• Subject-Matter Jurisdiction: Under B.P. Blg. 129, as amended, and P.D. 1866, illegal possession of firearms carries a penalty of reclusion temporal in its maximum period to reclusion perpetua, falling squarely under the exclusive original jurisdiction of the Regional Trial Court.\n• Territorial Jurisdiction: Venue in criminal cases is jurisdictional (locus delicti). The firearm was discovered and seized in Mabalacat, Pampanga. The City Prosecutor and RTC of Angeles City had NO territorial jurisdiction over crimes committed in Mabalacat. Only the Provincial Prosecutor and RTC of Pampanga had jurisdiction."),
                ("8) Where will the trial be held?", 
                 "Regional Trial Court (RTC) of San Fernando, Pampanga, Branch 46.")
            ],
            "facts": "On June 28, 1989, petitioner Renato Cudia was arrested by law enforcement operatives in the municipality of Mabalacat, Pampanga, for carrying an unlicensed .38 caliber revolver with ammunition. Following the arrest, the City Prosecutor of Angeles City conducted an inquest and filed an Information for Illegal Possession of Firearms and Ammunition under Presidential Decree No. 1866 before the Regional Trial Court of Angeles City, Branch 56 (Criminal Case No. 11542). Cudia was arraigned and entered a plea of not guilty.\n\nDuring the trial, it was established that the arrest and the apprehension of the firearm took place in Barangay Malabanias / Mabalacat, Pampanga, an area outside the territorial limits of Angeles City. Realizing that the crime occurred in the municipality of Mabalacat (which was under the exclusive territorial jurisdiction of the Provincial Prosecutor of Pampanga and the RTC of San Fernando, Pampanga), the City Prosecutor of Angeles City filed a motion to dismiss Criminal Case No. 11542 on the ground of lack of territorial jurisdiction. Over Cudia's objection, Judge Carlos D. Rustia of RTC Angeles City dismissed the case.\n\nSubsequently, the Provincial Prosecutor of Pampanga conducted a preliminary investigation and filed a new Information for the exact same offense against Cudia before the Regional Trial Court of San Fernando, Pampanga (Criminal Case No. 11987). Cudia filed a Motion to Quash the new Information on the ground of Double Jeopardy (Rule 117, Section 7), asserting that he had already been arraigned and the first case was dismissed without his express consent.\n\nThe RTC of San Fernando denied Cudia's motion to quash, and the Court of Appeals affirmed the denial. Cudia filed a Petition for Review on Certiorari before the Supreme Court.",
            "a1_facts": "A man was caught carrying an unlicensed .38 caliber gun in the town of Mabalacat, Pampanga. However, he was mistakenly brought before the City Prosecutor of Angeles City, who charged him in the Angeles City court. After the man pleaded 'not guilty,' it was discovered that the crime happened outside Angeles City (in Mabalacat), so the Angeles City prosecutor had no legal power over the case, and the Angeles judge dismissed it. Later, the correct Provincial Prosecutor of Pampanga filed the gun charges in the correct Pampanga court. The man objected, arguing 'Double Jeopardy'—claiming he was being tried twice for the same crime after the first case was dismissed against his will.",
            "issues": "1. Whether an Information filed by a prosecutor who has no territorial authority over the place of commission of the crime is valid and confers jurisdiction upon the court.\n2. Whether the dismissal of the first Information barred the filing of the second Information on the ground of double jeopardy.",
            "a2_issues": "1. Can a city prosecutor file criminal charges for a crime that happened in a completely different town outside his city?\n2. Does dismissing a charge sheet that was filed in the wrong city by the wrong prosecutor protect the accused from being charged again in the correct court under 'Double Jeopardy'?",
            "ruling": "No on both counts. The Supreme Court affirmed the Court of Appeals and ruled that double jeopardy did not attach.\n\n1. Authority to File Information is Jurisdictional: Under Section 4 and Section 5, Rule 110 of the Rules of Court, an Information must be subscribed and filed by an officer authorized by law. The authority of a prosecutor to investigate and file criminal charges is strictly confined within his designated territorial jurisdiction. The City Prosecutor of Angeles City had no legal authority to file an Information for an offense committed in Mabalacat, Pampanga. An Information filed by a prosecutor who lacks authority is void ab initio. It confers NO jurisdiction upon the trial court over the offense.\n\n2. Requisites of Double Jeopardy Not Satisfied: Under Section 7, Rule 117 of the Rules of Court, for double jeopardy to attach, four indispensable requisites must concur:\n(a) A valid complaint or information;\n(b) Before a court of competent jurisdiction;\n(c) The accused has been arraigned and entered a plea; and\n(d) The accused was convicted, acquitted, or the case was dismissed without his express consent.\n\nHere, the first two requisites were fatally absent:\n- The first Information was NOT valid because it was filed by an officer lacking statutory and territorial authority; and\n- The RTC of Angeles City was NOT a court of competent jurisdiction because territorial venue in criminal cases is jurisdictional (Rule 110, Section 15).\n\nBecause the first trial court never acquired jurisdiction, legal jeopardy never attached. Therefore, the dismissal of Criminal Case No. 11542 did not constitute a bar to the subsequent filing of a valid Information by the Provincial Prosecutor of Pampanga before the RTC of San Fernando.",
            "a3_ruling": "No to both questions. The Supreme Court ruled that the man could be tried in the second court and that 'Double Jeopardy' did not apply:\n1. A prosecutor only has power inside their assigned city or province. Because the gun was found in Mabalacat, the Angeles City prosecutor had zero authority to file the case, making the first charge sheet completely null and void.\n2. For Double Jeopardy to protect you, the first trial must have started with a VALID charge sheet in a court that actually had LEGAL POWER (jurisdiction) over the case. Because the first case was filed by the wrong prosecutor in the wrong city, the first court never had legal power, and the accused was never in real legal danger (jeopardy). Therefore, the correct prosecutor in Pampanga has every right to file the case and put him on trial.",
            "doctrine": "• Rule 110, Section 4 & 5 (Authority of Filing Prosecutor): The filing of a criminal Information by an officer without statutory authority is a jurisdictional defect that renders the proceedings void ab initio. It is non-curable by amendment or waiver.\n• Rule 110, Section 15 (Territorial Jurisdiction & Venue): In Philippine criminal procedure, venue is jurisdictional. A court cannot try an offense committed outside its territorial boundaries.\n• Rule 117, Section 7 (Elements of Double Jeopardy): Double jeopardy requires a valid information before a competent court. If the first information is void or the court lacks jurisdiction, the accused is not placed in legal jeopardy, and dismissal will not bar subsequent prosecution before the proper forum.\n• DOJ Circulars No. 015 & 028 (s. 2024) Territorial Referral Mandate: If an investigating prosecutor receives a complaint for an offense committed outside their territorial jurisdiction, the prosecutor must immediately issue an Order of Referral transmitting the records to the proper City or Provincial Prosecution Office having territorial jurisdiction, preventing void filings.",
            "a4_doctrine": "A prosecutor cannot file cases for crimes that happened in another municipality. If a case is filed by the wrong prosecutor in the wrong city, the proceedings are legally invisible and void. Dismissing that void case does not trigger Double Jeopardy, and the correct prosecutor can file the charges again in the proper court.",
            "relevance": "Under the syllabus topic 'Rule 110 — Who Must Prosecute Criminal Actions / Territorial Authority of Prosecutors / Requisites of Valid Information / Double Jeopardy,' this case is the landmark precedent establishing that:\n(1) An unauthorized prosecutor produces a void Information;\n(2) Venue is strictly jurisdictional in criminal cases; and\n(3) Legal jeopardy never attaches to a proceeding founded on a void Information before a court without territorial jurisdiction.",
            "a5_relevance": "Serves as the foundational case proving that a charge sheet filed by an unauthorized prosecutor is void from the beginning and cannot trigger Double Jeopardy to block a valid trial in the proper court."
        }
    ]

    for c in cases:
        # Case Header Table
        h_table = doc.add_table(rows=1, cols=1)
        h_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        h_table.autofit = False
        h_cell = h_table.cell(0, 0)
        h_cell.width = Inches(6.5)
        set_cell_background(h_cell, "0F2942")
        set_cell_margins(h_cell, top=140, bottom=140, left=180, right=180)
        
        hp = h_cell.paragraphs[0]
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(2)
        hrun = hp.add_run(f"CASE NO. {c['num']}: {c['title']}")
        hrun.bold = True
        hrun.font.name = "Calibri"
        hrun.font.size = Pt(13)
        hrun.font.color.rgb = RGBColor(255, 255, 255)
        
        sub_hp = h_cell.add_paragraph()
        sub_hp.paragraph_format.space_before = Pt(2)
        sub_hp.paragraph_format.space_after = Pt(0)
        sub_hrun = sub_hp.add_run(f"Citation: {c['citation']}\nPonente: {c['ponente']}\nTopic: {c['topic']}")
        sub_hrun.font.name = "Calibri"
        sub_hrun.font.size = Pt(9.5)
        sub_hrun.font.color.rgb = RGBColor(203, 213, 225)
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
        # 8-Point Information Table
        info_header = doc.add_paragraph()
        info_header.paragraph_format.space_before = Pt(6)
        info_header.paragraph_format.space_after = Pt(4)
        ih_run = info_header.add_run("📋 8-POINT CRIMINAL INFORMATION & JURISDICTIONAL ANALYSIS")
        ih_run.bold = True
        ih_run.font.name = "Calibri"
        ih_run.font.size = Pt(11.5)
        ih_run.font.color.rgb = RGBColor(30, 58, 138)
        
        i_table = doc.add_table(rows=len(c['info']), cols=2)
        i_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        i_table.autofit = False
        set_table_borders(i_table, color="CBD5E1", sz="4", val="single")
        
        for idx, (q, a) in enumerate(c['info']):
            row = i_table.rows[idx]
            c0, c1 = row.cells[0], row.cells[1]
            c0.width = Inches(2.2)
            c1.width = Inches(4.3)
            set_cell_background(c0, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
            set_cell_background(c1, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
            set_cell_margins(c0, top=80, bottom=80, left=100, right=100)
            set_cell_margins(c1, top=80, bottom=80, left=100, right=100)
            
            p0 = c0.paragraphs[0]
            p0.paragraph_format.space_before = Pt(0)
            p0.paragraph_format.space_after = Pt(0)
            p0.paragraph_format.line_spacing = 1.1
            r0 = p0.add_run(q)
            r0.bold = True
            r0.font.name = "Calibri"
            r0.font.size = Pt(9.5)
            r0.font.color.rgb = RGBColor(15, 23, 42)
            
            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.space_after = Pt(0)
            p1.paragraph_format.line_spacing = 1.15
            r1 = p1.add_run(a)
            r1.font.name = "Calibri"
            r1.font.size = Pt(9.5)
            r1.font.color.rgb = RGBColor(51, 65, 85)
            
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
        
        # 1. Facts
        sec1 = doc.add_paragraph()
        sec1.paragraph_format.space_before = Pt(6)
        sec1.paragraph_format.space_after = Pt(2)
        r_sec1 = sec1.add_run("1. FACTS:")
        r_sec1.bold = True
        r_sec1.font.name = "Calibri"
        r_sec1.font.size = Pt(11)
        r_sec1.font.color.rgb = RGBColor(15, 41, 66)
        
        p_facts = doc.add_paragraph(c['facts'])
        p_facts.paragraph_format.space_before = Pt(0)
        p_facts.paragraph_format.space_after = Pt(4)
        p_facts.paragraph_format.line_spacing = 1.15
        
        add_callout_box(doc, "a1. FACTS (Plain English / Layman's Summary)", c['a1_facts'], bg_hex="F8FAFC", border_color="0284C7")
        
        # 2. Issues
        sec2 = doc.add_paragraph()
        sec2.paragraph_format.space_before = Pt(6)
        sec2.paragraph_format.space_after = Pt(2)
        r_sec2 = sec2.add_run("2. ISSUE/S:")
        r_sec2.bold = True
        r_sec2.font.name = "Calibri"
        r_sec2.font.size = Pt(11)
        r_sec2.font.color.rgb = RGBColor(15, 41, 66)
        
        p_issues = doc.add_paragraph(c['issues'])
        p_issues.paragraph_format.space_before = Pt(0)
        p_issues.paragraph_format.space_after = Pt(4)
        p_issues.paragraph_format.line_spacing = 1.15
        
        add_callout_box(doc, "a2. ISSUE/S (Plain English Question)", c['a2_issues'], bg_hex="F8FAFC", border_color="EAB308")
        
        # 3. Ruling
        sec3 = doc.add_paragraph()
        sec3.paragraph_format.space_before = Pt(6)
        sec3.paragraph_format.space_after = Pt(2)
        r_sec3 = sec3.add_run("3. COURT'S RULING:")
        r_sec3.bold = True
        r_sec3.font.name = "Calibri"
        r_sec3.font.size = Pt(11)
        r_sec3.font.color.rgb = RGBColor(15, 41, 66)
        
        p_ruling = doc.add_paragraph(c['ruling'])
        p_ruling.paragraph_format.space_before = Pt(0)
        p_ruling.paragraph_format.space_after = Pt(4)
        p_ruling.paragraph_format.line_spacing = 1.15
        
        add_callout_box(doc, "a3. COURT'S RULING (Plain English Summary)", c['a3_ruling'], bg_hex="F8FAFC", border_color="10B981")
        
        # 4. Doctrine & Procedure
        sec4 = doc.add_paragraph()
        sec4.paragraph_format.space_before = Pt(6)
        sec4.paragraph_format.space_after = Pt(2)
        r_sec4 = sec4.add_run("4. APPLICABLE DOCTRINE & CRIMINAL PROCEDURE APPLICATION:")
        r_sec4.bold = True
        r_sec4.font.name = "Calibri"
        r_sec4.font.size = Pt(11)
        r_sec4.font.color.rgb = RGBColor(15, 41, 66)
        
        p_doc = doc.add_paragraph(c['doctrine'])
        p_doc.paragraph_format.space_before = Pt(0)
        p_doc.paragraph_format.space_after = Pt(4)
        p_doc.paragraph_format.line_spacing = 1.15
        
        add_callout_box(doc, "a4. APPLICABLE DOCTRINE (Plain English Takeaway)", c['a4_doctrine'], bg_hex="F8FAFC", border_color="8B5CF6")
        
        # 5. Relevance
        sec5 = doc.add_paragraph()
        sec5.paragraph_format.space_before = Pt(6)
        sec5.paragraph_format.space_after = Pt(2)
        r_sec5 = sec5.add_run("5. RELEVANCE TO THE TOPIC & EXPLANATION:")
        r_sec5.bold = True
        r_sec5.font.name = "Calibri"
        r_sec5.font.size = Pt(11)
        r_sec5.font.color.rgb = RGBColor(15, 41, 66)
        
        p_rel = doc.add_paragraph(c['relevance'])
        p_rel.paragraph_format.space_before = Pt(0)
        p_rel.paragraph_format.space_after = Pt(4)
        p_rel.paragraph_format.line_spacing = 1.15
        
        add_callout_box(doc, "a5. RELEVANCE TO THE TOPIC (Plain English Syllabus Note)", c['a5_relevance'], bg_hex="F8FAFC", border_color="EC4899")
        
        doc.add_page_break()

    # Summary Table Section
    st_p = doc.add_paragraph()
    st_p.paragraph_format.space_before = Pt(0)
    st_p.paragraph_format.space_after = Pt(4)
    st_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    st_run = st_p.add_run("WEEK 4 SUMMARY COMPARISON TABLE")
    st_run.bold = True
    st_run.font.name = "Calibri"
    st_run.font.size = Pt(14)
    st_run.font.color.rgb = RGBColor(15, 41, 66)
    
    st_sub = doc.add_paragraph()
    st_sub.paragraph_format.space_before = Pt(0)
    st_sub.paragraph_format.space_after = Pt(12)
    st_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    st_srun = st_sub.add_run("Rule 110 (Prosecution of Offenses) Key Doctrines & Procedural Distinctions")
    st_srun.font.name = "Calibri"
    st_srun.font.size = Pt(10)
    st_srun.font.italic = True
    st_srun.font.color.rgb = RGBColor(100, 116, 139)
    
    summary_data = [
        ("Chua v. Padillo\n(G.R. No. 163797, 2007)", 
         "Sec. 5 (Control of Prosecution) & Sec. 4, Rule 112 (Probable Cause)", 
         "Executive prosecutorial discretion is subject to judicial review via Rule 65 when the DOJ Secretary acts with grave abuse of discretion by ignoring financial/circumstantial proof of conspiracy.", 
         "Prevents arbitrary dismissal/exclusion of co-conspirators when probable cause is supported by bank records."),
        ("Bureau of Customs v. Sherman\n(G.R. No. 190487, 2011)", 
         "Sec. 5 (Control of Prosecution), Sec. 16 (Intervention) & OSG Mandate", 
         "Government agencies/bureaus acting as nominal complainants cannot appeal criminal dismissals before the SC without the authorization and representation of the OSG.", 
         "Establishes the sole authority of the OSG in appellate criminal proceedings and affirms prosecutorial control over motions to withdraw."),
        ("People v. Cinco\n(G.R. No. 186460, 2009)", 
         "Sec. 6 (Sufficiency of Info) & Sec. 11 (Date of Offense)", 
         "In offenses where date is not a material element (e.g. rape), alleging 'on or about [month/year]' is legally sufficient. Objections to vagueness must be raised via Bill of Particulars prior to plea.", 
         "Prevents technical dismissals based on approximate dates and enforces waiver of formal defects upon entering a plea."),
        ("Estodillo v. Judge Baluma\n(A.M. No. RTJ-04-1837, 2004)", 
         "Sec. 3 (Complaint) vs. Sec. 4 (Information) & Rule 112", 
         "An Information is NOT required to be sworn to; it is only subscribed by the prosecutor who acts under official oath of office. Dismissing an Information for lack of oath is gross ignorance.", 
         "Clarifies the structural difference between private complaints and public informations."),
        ("Cudia v. Court of Appeals\n(G.R. No. 110315, 1998)", 
         "Sec. 4 & 5 (Prosecutor's Authority), Sec. 15 (Venue) & Rule 117 (Double Jeopardy)", 
         "An Information filed by a prosecutor without territorial authority over the place of the crime is void ab initio. The court acquires no jurisdiction, and double jeopardy does not attach to bar re-filing.", 
         "Enforces strict territorial limits on prosecutors and establishes that void informations do not trigger double jeopardy.")
    ]
    
    sum_table = doc.add_table(rows=len(summary_data) + 1, cols=4)
    sum_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_table.autofit = False
    set_table_borders(sum_table, color="94A3B8", sz="6", val="single")
    
    headers = ["Case Name & Citation", "Rule 110 / Procedural Topic", "Key Ratio Decidendi", "Impact on Criminal Procedure"]
    col_widths = [Inches(1.5), Inches(1.5), Inches(1.9), Inches(1.6)]
    
    hdr_row = sum_table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.width = col_widths[i]
        set_cell_background(cell, "0F2942")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    for r_idx, (case_n, top_n, ratio_n, imp_n) in enumerate(summary_data):
        row = sum_table.rows[r_idx + 1]
        bg = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate([case_n, top_n, ratio_n, imp_n]):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.1
            r = p.add_run(val)
            if c_idx == 0:
                r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(30, 41, 59)
            
    doc.save(filename)
    print(f"Successfully created: {filename}")

if __name__ == "__main__":
    out_path = r"c:\Users\JR\Downloads\14All-All41\MLC\First Sem 1st Year\Criminal Procedure\week 4\Criminal_Procedure_Week_4_Case_Digests.docx"
    create_case_digest_docx(out_path)
