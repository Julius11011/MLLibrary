# generate_phase1_full.py
import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from core_builder import (
    create_docx_builder, add_header_box, add_heading_1, add_heading_2,
    add_body_p, add_student_explanation_box, add_alac_box,
    add_latin_maxims_box, add_statcon_table
)

from statcon_database_phase1 import cases_p1

# Cases 11 to 21 for Phase 1
cases_p1_part2 = [
    # 11. People v. Avecilla / Eastern Med v. Surio
    {
        "num": 11,
        "title": "People of the Philippines v. Macario Avecilla / Eastern Mediterranean Maritime Ltd. v. Surio",
        "citation": "People v. Avecilla (G.R. No. 117033 | February 15, 2001 | 351 SCRA 635) / Eastern Med v. Surio (G.R. No. 154213 | August 21, 2013 | 704 SCRA 232)",
        "plaintiff": "People of the Philippines / Eastern Mediterranean Maritime Ltd. (Petitioners)",
        "defendant": "Macario Avecilla / Estanislao Surio, et al. (Respondents)",
        "court": "RTC Manila / Court of Appeals & NLRC",
        "nature": "Consolidated study on Retroactivity of Favorable Penal Laws (RA 8294 in Avecilla) and Prospective Construction of POEA Standard Employment Contracts (in Eastern Med).",
        "facts": "In People v. Avecilla, accused was convicted of illegal possession of firearms used in homicide under PD 1866. While on appeal, RA 8294 was enacted, which provided that if homicide is committed with an unlicensed firearm, such possession is merely an aggravating circumstance, not a separate offense. Accused invoked retroactive application of RA 8294. In Eastern Med v. Surio, seafarers claimed disability benefits under a 2000 POEA Standard Employment Contract amendment for an illness contracted during a 1999 voyage governed by the 1996 POEA contract.",
        "plain_facts": "In Avecilla, a new law made illegal gun possession less penal when homicide was committed. The court applied it retroactively because it was favorable to the accused. In Eastern Med, seafarers tried to apply a new 2000 employment contract rule to a 1999 case, but the court ruled that contracts and laws apply prospectively unless retroactivity is expressly provided.",
        "issue": "1. In Avecilla: Whether or not a new penal law that decriminalizes separate gun possession when homicide is committed applies retroactively to pending cases if favorable to the accused.\n2. In Eastern Med: Whether or not amendments to administrative contracts (POEA SEC) apply retroactively to preexisting employment contracts.",
        "issue_plain": "Do favorable criminal laws apply retroactively, and do new administrative contracts apply prospectively?",
        "issue_recite": "\"Sir/Ma'am, these cases contrast the retroactivity of favorable penal statutes under Article 22 RPC with the general rule of prospective construction of laws and contracts under Article 4 of the Civil Code.\"",
        "answer": "1. In Avecilla: YES. Penal laws that are favorable to the accused must be given retroactive effect pursuant to Article 22 of the Revised Penal Code.\n2. In Eastern Med: NO. Laws and administrative regulations operate prospectively (Lex prospicit, non respicit) unless the contrary is clearly expressed.",
        "answer_plain": "1. Favorable penal laws are retroactive (Favorabilia sunt amplianda).\n2. Civil laws, regulations, and contracts are prospective (Lex prospicit, non respicit).",
        "answer_recite": "\"Sir/Ma'am: Under Article 22 RPC, penal laws operate retroactively if favorable to the accused (In dubio pro reo). In contrast, under Article 4 of the Civil Code, laws and contracts apply prospectively (Lex prospicit, non respicit).\"",
        "legal_basis": [
            ("1. Article 22, Revised Penal Code:", "Retroactive effect of penal laws favorable to the person guilty of a felony."),
            ("2. Article 4, Civil Code of the Philippines:", "'Laws shall have no retroactive effect, unless the contrary is provided.'"),
            ("3. Republic Act No. 8294 vs. P.D. No. 1866:", "Statutory amendment on illegal firearms.")
        ],
        "legal_basis_plain": "Article 22 RPC (favorable penal retroactivity) and Article 4 Civil Code (prospective application of laws).",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Article 22 RPC (retroactivity of favorable penal laws) and Article 4 Civil Code (prospective operation of statutes).\"",
        "analysis": [
            ("1. Avecilla (Penal Retroactivity):", "Since RA 8294 treats unlicensed gun possession as a mere aggravating circumstance rather than a distinct crime when homicide is committed, it favors the accused and must be applied retroactively to dismiss the separate gun charge."),
            ("2. Eastern Med (Civil/Contractual Prospectivity):", "Subsequent amendments to the POEA standard contract create new substantive rights and obligations; thus, they cannot be applied retroactively to impair rights under existing contracts without violating due process."),
            ("3. Harmonization:", "Statutory construction distinguishes between penal leniency (which looks back to favor liberty) and civil/contractual rules (which look forward to protect vested rights).")
        ],
        "analysis_plain": "If a new criminal law helps the accused, courts must apply it retroactively. But if a new civil law or contract rule creates new rights, it only applies forward to protect vested rights.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Favorabilia sunt amplianda in criminal law to give retroactive relief, while strictly enforcing Lex prospicit, non respicit in labor contracts to protect vested obligations.\"",
        "conclusion": "In Avecilla, the separate conviction for illegal possession of firearms was ACQUITTED/DISMISSED. In Eastern Med, the application of the 2000 amended contract was REVERSED and the 1996 contract terms were upheld.",
        "conclusion_plain": "Avecilla was acquitted of the separate gun charge; Eastern Med won because contract amendments are prospective.",
        "conclusion_recite": "\"The Supreme Court applied retroactive construction in Avecilla to favor the accused, and prospective construction in Eastern Med to protect contractual reliance.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - I. Kinds of construction -> b. Prospective v. retrospective (Article 22 RPC vs. Article 4 Civil Code)",
        "latin_maxims": [
            ("Lex prospicit, non respicit (the law looks forward, not backward)", "The law looks forward, not backward.", "Statutes, regulations, and contracts are presumed prospective under Article 4 Civil Code."),
            ("Favorabilia sunt amplianda, odiosa sunt restringenda (favorable things are to be expanded, hateful things restricted)", "Things favorable are to be enlarged; things odious are to be restrained.", "Penal laws favorable to the accused are given liberal retroactive application."),
            ("In dubio pro reo (when in doubt, rule in favor of the accused)", "In case of doubt, rule for the accused.", "Ambiguities and new statutory benefits are resolved in favor of the defendant.")
        ],
        "primary_maxim": "Lex prospicit, non respicit\n(the law looks forward, not backward / general rule of prospectivity)",
        "secondary_maxims": "• Favorabilia sunt amplianda, odiosa sunt restringenda (favorable things expanded, odious restricted)\n• In dubio pro reo (when in doubt, favor the accused)",
        "violation_text": "In Eastern Med, the NLRC violated Article 4 of the Civil Code by applying a new administrative contract retroactively to impair existing rights.",
        "execution_text": "Core Principles on Statutory Prospectivity vs. Retroactivity: General rule is prospectivity (Art. 4 CC), but favorable penal laws are an express exception (Art. 22 RPC).",
        "statcon_plain": "Proves that laws look forward, except favorable criminal laws which apply backward.",
        "statcon_recite": "\"Sir/Ma'am, under Chapter I (I-b. Prospective v. Retrospective): Avecilla and Eastern Med establish that statutes operate prospectively (Lex prospicit, non respicit) unless they are penal laws favorable to the accused (Art. 22 RPC).\""
    },

    # 12. Tawang Multi-Purpose Coop. v. La Trinidad Water District
    {
        "num": 12,
        "title": "Tawang Multi-Purpose Cooperative v. La Trinidad Water District",
        "citation": "G.R. No. 166471 | March 22, 2011 | 646 SCRA 21 | En Banc | Ponente: J. Carpio",
        "plaintiff": "Tawang Multi-Purpose Cooperative (Petitioner)",
        "defendant": "La Trinidad Water District (Respondent)",
        "court": "National Water Resources Board (NWRB) / Court of Appeals",
        "nature": "Petition for Review on Certiorari challenging the constitutionality of Section 47 of Presidential Decree No. 198 (The Provincial Water Utilities Act) for creating an exclusive franchise.",
        "facts": "Tawang Multi-Purpose Cooperative applied with the NWRB for a Certificate of Public Convenience (CPC) to operate a communal water supply system in Barangay Tawang, La Trinidad, Benguet. La Trinidad Water District (LTWD) opposed, invoking Section 47 of P.D. No. 198, which granted local water districts an exclusive franchise, prohibiting any other water service provider without the district's consent. Tawang Coop challenged Section 47 as unconstitutional under Section 11, Article XII of the 1987 Constitution, which bans exclusive franchises.",
        "plain_facts": "A local cooperative wanted to provide water to its village. The government water district tried to block them, citing an old Marcos decree that gave water districts an 'exclusive monopoly'. The cooperative challenged the law as unconstitutional because the Constitution forbids exclusive franchises.",
        "issue": "Whether Section 47 of P.D. No. 198, which grants exclusive franchise privileges to local water districts, violates Section 11, Article XII of the Constitution prohibiting exclusive franchises.",
        "issue_plain": "Can a presidential decree grant an exclusive water monopoly when the Constitution expressly says no franchise shall be exclusive?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether Section 47 of PD 198 is unconstitutional for granting an exclusive water franchise in direct violation of Section 11, Article XII of the 1987 Constitution, testing Verba Legis in Constitutional Construction.\"",
        "answer": "YES. Section 47 of P.D. No. 198 is UNCONSTITUTIONAL and VOID. The constitutional command is clear and unambiguous: no franchise, certificate, or authorization for the operation of a public utility shall be exclusive in character.",
        "answer_plain": "Yes, Section 47 is unconstitutional. The Constitution literally says no franchise can be exclusive. The decree's monopoly is void.",
        "answer_recite": "\"Yes, Sir/Ma'am. Section 47 of PD 198 is unconstitutional. Under Verba Legis (the words of the law), Section 11, Article XII of the Constitution explicitly forbids exclusive public utility franchises.\"",
        "legal_basis": [
            ("1. Section 11, Article XII, 1987 Constitution:", "'No franchise, certificate, or any other form of authorization for the operation of a public utility shall be exclusive in character...'"),
            ("2. Section 47, Presidential Decree No. 198:", "Exclusive franchise clause of local water districts."),
            ("3. Constitutional Supremacy & Verba Legis:", "The words of the Constitution are paramount; any statute conflicting with its clear text is void.")
        ],
        "legal_basis_plain": "Section 11, Article XII of the Constitution and the principle of Constitutional Supremacy.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Section 11, Article XII of the 1987 Constitution and the doctrine of Constitutional Supremacy under Verba legis non est recedendum (from the words of the statute there should be no departure).\"",
        "analysis": [
            ("1. Plain Meaning of the Constitution:", "The phrasing of Section 11, Article XII is plain, clear, and absolute: no franchise shall be exclusive. When the Constitution speaks in unequivocal terms, there is no room for interpretation."),
            ("2. Nullity of Section 47 PD 198:", "Section 47 granted water districts an absolute monopoly by requiring their consent before any other entity could operate. This statutory exclusivity directly collided with the constitutional prohibition."),
            ("3. Free Competition in Utilities:", "The constitutional policy is to prevent monopolies and protect consumers by allowing qualified cooperatives to operate.")
        ],
        "analysis_plain": "The Constitution says in plain English: 'No franchise shall be exclusive.' A decree cannot give a water district an exclusive monopoly. The constitutional text controls.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Verba Legis in constitutional construction. The constitutional ban against exclusive franchises is plain and unequivocal. Section 47 of PD 198 directly contradicted the Constitution and was void ab initio.\"",
        "conclusion": "The Supreme Court GRANTED the petition, declared Section 47 of P.D. No. 198 UNCONSTITUTIONAL, and directed the NWRB to process Tawang Cooperative's CPC application.",
        "conclusion_plain": "The cooperative won. Section 47 of PD 198 was struck down as unconstitutional.",
        "conclusion_recite": "\"The Supreme Court struck down Section 47 of PD 198 for violating the constitutional ban on exclusive franchises.\"",
        "syllabus": "• Chapter II: Prelude to the Exercise of the Power\n  - A. Verba legis non est recedendum (from the words of the statute there should be no departure)\n  - B. Absoluta sententia expositore non indigent (when language is plain, it needs no interpreter)\n• Chapter VI: Constitutional Construction",
        "latin_maxims": [
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "Applied to the Constitution: clear constitutional provisions must be enforced literally."),
            ("Absoluta sententia expositore non indigent (when language is plain, it needs no interpreter)", "When language is plain, it needs no interpreter.", "The constitutional prohibition on exclusive franchises needs no exposition."),
            ("Lex superior derogat legi inferiori (a higher law overrides a lower law)", "A higher law overrides a lower law.", "The Constitution reigns supreme over conflicting presidential decrees.")
        ],
        "primary_maxim": "Verba legis non est recedendum\n(from the words of the statute/constitution there should be no departure)",
        "secondary_maxims": "• Absoluta sententia expositore non indigent (plain language needs no interpreter)\n• Lex superior derogat legi inferiori (constitutional supremacy)",
        "violation_text": "The respondent water district attempted to enforce an unconstitutional statutory monopoly in defiance of the clear text of the Constitution.",
        "execution_text": "Constitutional Verba Legis: Where the Constitution speaks in plain, mandatory language banning exclusivity, any conflicting statutory provision is void ab initio.",
        "statcon_plain": "When the Constitution says 'No franchise shall be exclusive', it means exactly what it says.",
        "statcon_recite": "\"Sir/Ma'am, Tawang v. La Trinidad Water District demonstrates Verba Legis in constitutional construction: the clear constitutional text banning exclusive franchises overrides statutory monopolies.\""
    },

    # 13. Nippon Express v. CIR
    {
        "num": 13,
        "title": "Nippon Express (Philippines) Corporation v. Commissioner of Internal Revenue",
        "citation": "G.R. No. 196907 | March 13, 2013 | 693 SCRA 356 | Second Division | Ponente: J. Del Castillo",
        "plaintiff": "Nippon Express (Philippines) Corporation (Petitioner)",
        "defendant": "Commissioner of Internal Revenue (Respondent)",
        "court": "Court of Tax Appeals (CTA En Banc)",
        "nature": "Petition for Review on Certiorari regarding a claim for refund/tax credit of unutilized input VAT under Section 112 of the National Internal Revenue Code (NIRC).",
        "facts": "Nippon Express filed an administrative claim for refund of excess input VAT with the BIR on April 24, 2003. Without waiting for the 120-day period granted to the CIR to decide the claim under Section 112(D) (now 112(C)) of the NIRC, Nippon Express filed a judicial petition with the CTA on April 25, 2003 (just one day later). The CIR argued the CTA had no jurisdiction because the 120-day waiting period is mandatory and jurisdictional under the Tax Code.",
        "plain_facts": "A company filed a tax refund claim with the BIR and immediately ran to the Tax Court the next day without waiting for the 120 days given by law to the BIR. The BIR argued the lawsuit was premature and void.",
        "issue": "Whether or not the 120-day period provided in Section 112 of the NIRC is mandatory and jurisdictional before a judicial claim for VAT refund can be filed with the CTA.",
        "issue_plain": "Does a taxpayer have to wait the full 120 days for the BIR to act before going to court?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether the 120-day period under Section 112 of the Tax Code is mandatory and jurisdictional, testing the literal interpretation of tax refund statutes under Verba Legis.\"",
        "answer": "YES. The 120-day waiting period is MANDATORY and JURISDICTIONAL. A judicial claim filed before the expiration of the 120-day period is premature, depriving the CTA of jurisdiction.",
        "answer_plain": "Yes. The 120-day period is strict and mandatory. Going to court after 1 day is premature and the court has zero jurisdiction.",
        "answer_recite": "\"Yes, Sir/Ma'am. Under Section 112 of the NIRC and the doctrine in CIR v. Aichi Forging, the 120-day period is mandatory and jurisdictional. Premature filing renders the judicial claim dismissible for lack of jurisdiction.\"",
        "legal_basis": [
            ("1. Section 112(D) (now 112(C)), NIRC of 1997:", "The Commissioner has 120 days from submission of documents to grant or deny refund; taxpayer has 30 days from denial or inaction to appeal to CTA."),
            ("2. Strictissimi Juris in Tax Refunds:", "Tax refunds are in the nature of tax exemptions and are strictly construed against the claimant."),
            ("3. Verba Legis (the words of the law):", "Mandatory statutory time frames cannot be relaxed by courts.")
        ],
        "legal_basis_plain": "Section 112 of the Tax Code and the rule that tax refund rules are strictly construed against the taxpayer.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Section 112 of the Tax Code, strictly construed under Strictissimi Juris (strict construction of tax refunds) and Verba Legis (the words of the law).\"",
        "analysis": [
            ("1. Mandatory Word 'Shall':", "The statute uses mandatory phrasing granting the CIR 120 days to examine documents. The taxpayer cannot preempt administrative action by filing in court on day 1."),
            ("2. Jurisdictional Defect:", "Compliance with the 120-day period is a condition precedent to invoking the CTA's jurisdiction. Non-compliance is fatal."),
            ("3. Strict Construction of Tax Refunds:", "Tax refunds represent a derogation of state sovereignty and must be strictly pursued in exact accordance with statutory procedures.")
        ],
        "analysis_plain": "The Tax Code explicitly gives the BIR 120 days to decide. If you sue on Day 1, you violated the law's procedure. Tax refunds are strictly interpreted against taxpayers.",
        "analysis_recite": "\"Sir/Ma'am, the Court strictly applied Verba Legis. The 120-day period gives the BIR administrative jurisdiction to decide. Filing on Day 1 was premature, rendering the CTA without jurisdiction under Strictissimi Juris.\"",
        "conclusion": "The Supreme Court DENIED the petition and AFFIRMED the dismissal of Nippon Express's refund claim for lack of jurisdiction.",
        "conclusion_plain": "Nippon Express lost its refund because it jumped the gun and filed in court prematurely.",
        "conclusion_recite": "\"The Supreme Court dismissed the tax refund petition for premature filing and lack of jurisdiction.\"",
        "syllabus": "• Chapter II: Prelude to the Exercise of the Power\n  - A. Verba legis non est recedendum (from the words of the statute there should be no departure)\n• Chapter VI: Construction of Specific Statutes -> B. Statutes strictly construed (Tax refund provisions)",
        "latin_maxims": [
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "Tax statutes setting mandatory prescriptive periods must be strictly applied."),
            ("Strictissimi juris (of the strictest interpretation)", "Of the strictest right/law.", "Tax refunds and exemptions are strictly construed against the taxpayer."),
            ("Dura lex, sed lex (the law is harsh, but it is the law)", "The law is harsh, but it is the law.", "Failure to follow statutory refund timeframes forfeits the claim.")
        ],
        "primary_maxim": "Strictissimi juris\n(of the strictest interpretation / tax refund statutes strictly construed)",
        "secondary_maxims": "• Verba legis non est recedendum (from the words of the statute there should be no departure)\n• Dura lex, sed lex (the law is harsh, but it is the law)",
        "violation_text": "The taxpayer violated the statutory procedural period by filing a premature judicial claim without waiting for the mandatory 120-day administrative period.",
        "execution_text": "Strict Application of Tax Prescriptive Periods: The words of Section 112 NIRC are mandatory and jurisdictional; courts cannot excuse premature filings under Verba Legis.",
        "statcon_plain": "Tax refund time periods are mandatory. If you file before the 120 days are up, your case is dead.",
        "statcon_recite": "\"Sir/Ma'am, Nippon Express v. CIR establishes that tax refund periods in Section 112 NIRC are mandatory and jurisdictional under Verba Legis and Strictissimi Juris.\""
    },

    # 14. Bolos v. Bolos
    {
        "num": 14,
        "title": "Cynthia S. Bolos v. Danilo T. Bolos",
        "citation": "G.R. No. 186400 | October 20, 2010 | 634 SCRA 429 | Second Division | Ponente: J. Mendoza",
        "plaintiff": "Cynthia S. Bolos (Petitioner)",
        "defendant": "Danilo T. Bolos (Respondent)",
        "court": "RTC Pasig City (Branch 69) / Court of Appeals",
        "nature": "Petition for Review on Certiorari regarding whether the Rule on Declaration of Absolute Nullity of Void Marriages (A.M. No. 02-11-10-SC) requiring counseling applies to marriages celebrated before its effectivity.",
        "facts": "Danilo Bolos filed a petition for declaration of absolute nullity of his marriage to Cynthia on the ground of lack of marriage license under Article 4 of the Family Code (a void ab initio marriage under Article 35). The RTC declared the marriage void. Cynthia appealed, arguing that Danilo failed to comply with A.M. No. 02-11-10-SC, which took effect in 2003 and provided procedural rules for nullity petitions, claiming it applied to all void marriages.",
        "plain_facts": "A husband sued to declare his marriage void for lack of a marriage license. The wife argued he didn't follow the Supreme Court's 2003 procedural rule on nullity. The Court had to clarify what marriages the rule applies to.",
        "issue": "Whether or not the phrase 'under the Family Code' in Section 1 of A.M. No. 02-11-10-SC applies to marriages celebrated before the Family Code took effect, and whether courts can distinguish where the rule does not.",
        "issue_plain": "Does the Supreme Court's procedural rule on void marriages apply strictly according to its clear text?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether A.M. No. 02-11-10-SC applies to marriages celebrated prior to the Family Code, testing the maxim Ubi lex non distinguit nec nos distinguere debemus (where the law does not distinguish, we should not distinguish).\"",
        "answer": "NO. Section 1 of A.M. No. 02-11-10-SC is explicit: it applies solely to petitions for declaration of absolute nullity of void marriages celebrated under the Family Code (from August 3, 1988 onward). Where the rule is clear, there is no room for interpretation.",
        "answer_plain": "No. The rule clearly states it only applies to marriages celebrated under the Family Code. Where the law makes a clear distinction, courts must apply it literally.",
        "answer_recite": "\"No, Sir/Ma'am. Under Verba Legis (the words of the law), Section 1 of AM No. 02-11-10-SC clearly and explicitly limits its scope to marriages celebrated under the Family Code.\"",
        "legal_basis": [
            ("1. Section 1, A.M. No. 02-11-10-SC:", "'This Rule shall govern petitions for declaration of absolute nullity of void marriages and annulment of voidable marriages under the Family Code of the Philippines.'"),
            ("2. Article 35 & 4, Family Code:", "Marriages void ab initio for lack of license."),
            ("3. Verba legis non est recedendum (from the words of the statute there should be no departure):", "Clear procedural rules must be applied literally.")
        ],
        "legal_basis_plain": "A.M. No. 02-11-10-SC and Verba Legis (the words of the law).",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Section 1 of AM No. 02-11-10-SC and the cardinal canon Verba legis non est recedendum (from the words of the statute there should be no departure).\"",
        "analysis": [
            ("1. Literal Phrasing:", "The rule explicitly specifies marriages 'under the Family Code'. A marriage celebrated in 1979 under the Civil Code is not covered by the 2003 procedural rule requiring mandatory pre-trial counseling."),
            ("2. Duty of Court:", "The Court must apply the literal phrasing. It is a cardinal rule that when a statute is clear and explicit, there is no room for interpretation."),
            ("3. Void Marriage for Lack of License:", "The marriage was void ab initio, and the trial court correctly nullified it under the applicable Civil Code provisions.")
        ],
        "analysis_plain": "The rule says 'under the Family Code'. The marriage happened in 1979 under the Civil Code. Judges must follow the plain words of the rule.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Verba Legis. The procedural rule explicitly applies only to marriages celebrated under the Family Code. Since the marriage occurred under the Civil Code, the procedural requirement did not apply.\"",
        "conclusion": "The Supreme Court DENIED the petition and AFFIRMED the nullity of the marriage.",
        "conclusion_plain": "Danilo won. The marriage was declared null and void.",
        "conclusion_recite": "\"The Supreme Court affirmed the nullity of the marriage under the plain language of the procedural rule.\"",
        "syllabus": "• Chapter II: Prelude to the Exercise of the Power\n  - A. Verba legis non est recedendum (from the words of the statute there should be no departure)\n  - B. Absoluta sententia expositore non indigent (when language is plain, it needs no interpreter)",
        "latin_maxims": [
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "The words 'under the Family Code' must be applied literally."),
            ("Absoluta sententia expositore non indigent (when language is plain, it needs no interpreter)", "When language is plain, it needs no interpreter.", "Clear procedural scopes need no judicial expansion."),
            ("Ubi lex non distinguit nec nos distinguere debemus (where the law does not distinguish, we should not distinguish)", "Where the law does not distinguish, we ought not to distinguish.", "Courts cannot create exceptions or expansions not written in the rule.")
        ],
        "primary_maxim": "Verba legis non est recedendum\n(from the words of the statute there should be no departure)",
        "secondary_maxims": "• Absoluta sententia expositore non indigent (plain language needs no interpreter)\n• Ubi lex non distinguit nec nos distinguere debemus (where the law does not distinguish, neither should we)",
        "violation_text": "Petitioner sought to expand a procedural rule to cover pre-Family Code marriages despite explicit statutory phrasing to the contrary.",
        "execution_text": "Literal Construction of Court Rules: Where a procedural rule clearly states its coverage, courts must apply it literally without judicial expansion.",
        "statcon_plain": "Procedural rules that state their coverage explicitly must be applied literally.",
        "statcon_recite": "\"Sir/Ma'am, Bolos v. Bolos illustrates Verba legis non est recedendum (from the words of the statute there should be no departure) applied to procedural rules.\""
    },

    # 15. Gan v. Reyes
    {
        "num": 15,
        "title": "Bernadette S. Gan v. Hon. Antonio C. Reyes and Frisco F. Reyes",
        "citation": "G.R. No. 145527 | May 28, 2002 | 382 SCRA 352 | Third Division | Ponente: J. Bellosillo",
        "plaintiff": "Bernadette S. Gan, representing her minor child Francheska Gan (Petitioner)",
        "defendant": "Hon. Antonio C. Reyes (Presiding Judge, RTC Pasig, Branch 61) and Frisco F. Reyes (Respondents)",
        "court": "Regional Trial Court of Pasig City (Branch 61) / Court of Appeals",
        "nature": "Petition for Review on Certiorari regarding whether a judgment for support is immediately executory under Section 4, Rule 39 of the Rules of Court.",
        "facts": "Bernadette Gan filed an action for recognition and support on behalf of her illegitimate minor daughter against Frisco Reyes. The RTC rendered judgment recognizing the child and ordering Frisco to pay monthly support. Frisco appealed. Bernadette moved for execution pending appeal under Section 4, Rule 39 of the Rules of Court, which explicitly states that judgments in actions for support are immediately executory and shall not be stayed by appeal unless ordered by the court. The CA stayed execution, holding that paternity was not yet final.",
        "plain_facts": "A mother won child support in trial court. The father appealed, and the Court of Appeals froze the child support. The mother sued, arguing that under the Rules of Court, judgments for support are immediately executory because a child cannot wait to eat.",
        "issue": "Whether or not a trial court judgment for support is immediately executory pending appeal pursuant to the literal mandate of Section 4, Rule 39 of the Rules of Court.",
        "issue_plain": "Does child support have to be paid immediately even if the father appeals the decision?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether a judgment for child support is immediately executory under Section 4, Rule 39 of the Rules of Court, testing Verba Legis and construction to prevent injustice and absurdity.\"",
        "answer": "YES. Judgments for support are IMMEDIATELY EXECUTORY. Section 4, Rule 39 is explicit, clear, and mandatory: judgments in actions for support shall not be stayed by appeal unless the appellate court directs otherwise.",
        "answer_plain": "Yes. Child support must be paid immediately. A child cannot put eating, schooling, and living on hold during years of appeal.",
        "answer_recite": "\"Yes, Sir/Ma'am. Section 4, Rule 39 explicitly provides that judgments for support are immediately executory. A child cannot starve while paternity is being appealed.\"",
        "legal_basis": [
            ("1. Section 4, Rule 39, Rules of Court:", "Judgments in actions for support, injunction, receivership, and accounting are immediately executory."),
            ("2. Article 194-208, Family Code:", "Duty to provide support."),
            ("3. Construction to Avoid Injustice and Absurdity:", "Statutes must be construed to accomplish their humanitarian purpose.")
        ],
        "legal_basis_plain": "Rule 39, Section 4 of the Rules of Court and the principle that laws must be interpreted to prevent injustice.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Section 4, Rule 39 of the Rules of Court and the rule of Construction to Avoid Injustice (Nemo tenetur ad impossibile - no one is bound to the impossible).\"",
        "analysis": [
            ("1. Mandatory Text of Rule 39:", "The rule explicitly classifies support among judgments that are immediately executory. The words are mandatory."),
            ("2. Purpose of Support:", "Support is for survival—food, clothing, shelter, education. Staying execution defeats the entire purpose of the law, as a child cannot wait years for an appeal to conclude before receiving food."),
            ("3. Harmonization with Equity:", "Literal interpretation of Rule 39 perfectly harmonizes with substantive justice and child protection.")
        ],
        "analysis_plain": "Rule 39 says support is immediately executory. Freezing child support during appeal creates extreme injustice and absurdity.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Verba Legis coupled with Construction to Avoid Injustice. The text of Rule 39 is explicit: support is immediately executory because subsistence cannot be postponed.\"",
        "conclusion": "The Supreme Court GRANTED the petition, REVERSED the CA, and ordered the IMMEDIATE EXECUTION of the support judgment pending appeal.",
        "conclusion_plain": "The mother won. The father was ordered to pay child support immediately.",
        "conclusion_recite": "\"The Supreme Court ordered the immediate execution of child support under Section 4, Rule 39.\"",
        "syllabus": "• Chapter II: Prelude to the Exercise of the Power\n  - A. Verba legis non est recedendum (from the words of the statute there should be no departure)\n  - C. Dura lex sed lex (the law is harsh, but it is the law)\n• Chapter III: Basic Guidelines -> A-g. Construction to avoid absurdity and injustice",
        "latin_maxims": [
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "Section 4 of Rule 39 explicitly commands immediate execution of support."),
            ("Lex non cogit ad impossibilia (the law does not compel the impossible)", "The law does not compel the impossible.", "A child cannot be expected to survive without sustenance pending appeal."),
            ("Interpretatio fienda est ut res magis valeat quam pereat (interpretation must be made so that the purpose may be effective)", "Interpretation should be made so the thing may stand rather than fall.", "Support laws must be construed to achieve their life-sustaining purpose.")
        ],
        "primary_maxim": "Verba legis non est recedendum\n(from the words of the statute there should be no departure / immediate execution of support)",
        "secondary_maxims": "• Lex non cogit ad impossibilia (the law does not compel the impossible)\n• Interpretatio fienda est ut res magis valeat quam pereat (interpreted to be effective)",
        "violation_text": "The Court of Appeals violated the explicit text of Section 4, Rule 39 by staying the execution of a child support judgment.",
        "execution_text": "Harmonizing Plain Meaning with Substantive Purpose: The literal command of procedural rules on support prevents absurdity and fulfills the humanitarian policy of the Family Code.",
        "statcon_plain": "Support judgments are immediately executory under the literal words of the rule to prevent the absurdity of a child starving during appeal.",
        "statcon_recite": "\"Sir/Ma'am, Gan v. Reyes shows how Verba Legis aligns with Construction to Avoid Absurdity: judgments for support are immediately executory under Section 4, Rule 39.\""
    }
]

cases_p1.extend(cases_p1_part2)
print(f"Loaded {len(cases_p1)} cases for Phase 1.")
