# generate_statcon_phase1.py
import os
import sys
from core_builder import (
    create_docx_builder, add_header_box, add_heading_1, add_heading_2,
    add_body_p, add_student_explanation_box, add_alac_box,
    add_latin_maxims_box, add_statcon_table
)

cases_phase1 = [
    # 1. Caltex v. Palomar
    {
        "num": 1,
        "title": "Caltex (Philippines), Inc. v. Enrico Palomar (Postmaster General)",
        "citation": "G.R. No. L-19650 | September 29, 1966 | 18 SCRA 247 | En Banc | Ponente: J. Castro",
        "plaintiff": "Caltex (Philippines), Inc. (Petitioner-Appellee)",
        "defendant": "Enrico Palomar, in his capacity as The Postmaster General (Respondent-Appellant)",
        "court": "Court of First Instance (CFI) of Manila (Branch XIII)",
        "nature": "Petition for Declaratory Relief under Rule 64 (now Rule 63) to determine the legality of a promotional contest under postal laws.",
        "facts": "In 1960, Caltex conceived a promotional campaign called the 'Caltex Hooded Pump Contest'. The contest called on motor vehicle owners and licensed drivers to estimate the number of liters a hooded pump would dispense at a given Caltex service station. Entry forms were available for free at all Caltex stations. No purchase of oil, gas, or other Caltex products was required, nor was any fee collected to participate. Prizes (including motor vehicles) were to be determined by a drawing of lots. Caltex requested the Postmaster General to clear the mailing of contest rules and entry forms. Postmaster General Enrico Palomar denied the request, ruling that the contest fell under the prohibitory provisions of Section 1954-A of the Revised Administrative Code (The Postal Law), which bars lottery or gift enterprise materials from the mails. Palomar threatened to ban all Caltex mail. Caltex filed a petition for declaratory relief.",
        "plain_facts": "Caltex held a free contest where people guessed gas numbers to win cars without buying anything. The Postmaster General banned their mail, claiming the contest was an illegal lottery or gift enterprise. Caltex sued to have the court declare the contest legal under postal laws.",
        "issue": "Whether or not the 'Caltex Hooded Pump Contest' constitutes a 'lottery' or 'gift enterprise' prohibited from the postal service under Section 1954-A of the Revised Administrative Code.",
        "issue_plain": "Does a free promotional contest that requires no purchase or payment count as an illegal 'lottery' or 'gift enterprise' under the Postal Law?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether the Caltex Hooded Pump Contest falls within the statutory prohibition against lotteries and gift enterprises under Section 1954-A of the Revised Administrative Code, testing the legal definitions of lottery elements and the boundaries between statutory construction and administrative overreach.\"",
        "answer": "No. The Caltex Hooded Pump Contest does not constitute a lottery or a gift enterprise. It is a completely lawful promotional scheme that is entitled to use the postal facilities.",
        "answer_plain": "No, it is not an illegal lottery or gift enterprise because it is 100% free—participants do not pay any money or buy any product.",
        "answer_recite": "\"No, Sir/Ma'am. The contest is not a lottery or gift enterprise. The essential element of consideration is completely absent since participants are not required to pay any entrance fee or purchase any Caltex products.\"",
        "legal_basis": [
            ("1. Section 1954-A, Revised Administrative Code:", "Prohibits the mailing of materials concerning any lottery, gift enterprise, or similar scheme offering prizes dependent upon lot or chance."),
            ("2. Legal Definition of Lottery:", "Under established jurisprudence, a lottery consists of three concurring essential elements: (a) Consideration (payment/cost); (b) Chance; and (c) Prize."),
            ("3. Construction of Prohibitory Statutes:", "Administrative officials cannot expand the meaning of statutory terms beyond their established legal definitions.")
        ],
        "legal_basis_plain": "Section 1954-A of the Postal Law, the 3 elements of a lottery (Consideration, Chance, Prize), and the rule that administrative officers cannot invent new statutory definitions.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis rests on Section 1954-A of the Revised Administrative Code and the three elements of a lottery: Consideration, Chance, and Prize. Under Noscitur a sociis (words are known by their associates), 'gift enterprise' must be construed in conjunction with 'lottery', both requiring consideration.\"",
        "analysis": [
            ("1. Absence of Consideration:", "The three elements of a lottery are consideration, chance, and prize. While chance and prize were present, consideration was completely absent. Participants did not pay any money or buy any merchandise."),
            ("2. Definition of Gift Enterprise:", "Under the rule of Noscitur a sociis (words are known by their associates), 'gift enterprise' is grouped with 'lottery' and must partake of the same nature—meaning the distribution of prizes must be coupled with the purchase of goods. A purely gratuitous contest is not a gift enterprise."),
            ("3. Judicial Construction vs. Administrative Expansion:", "The Postmaster General improperly attempted to expand the statute by punishing an act that lacked the essential element of consideration, which constitutes unauthorized administrative legislation.")
        ],
        "analysis_plain": "To be an illegal lottery, there must be (1) Prize, (2) Chance, and (3) Consideration (paying money). Here, there was prize and chance, but NO consideration because it was completely free. Under Noscitur a sociis (words are known by their associates), 'gift enterprise' also requires consideration. The Postmaster General had no right to expand the law.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Noscitur a sociis (words are known by their associates). Since 'gift enterprise' is associated with 'lottery', it must involve consideration (payment). Because Caltex required no purchase or fee, the element of consideration was absent. The Postmaster General engaged in administrative overreach by attempting to expand a prohibitory statute.\"",
        "conclusion": "The Supreme Court AFFIRMED the CFI judgment declaring that the Caltex contest was not prohibited by the Postal Law, and the Postmaster General had no legal authority to ban Caltex promotional mail.",
        "conclusion_plain": "Caltex won. The contest was declared completely legal and the Postmaster General's postal ban was struck down.",
        "conclusion_recite": "\"The Supreme Court affirmed the declaratory relief in favor of Caltex, holding that the contest was legal and not subject to postal mail prohibition.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - B. Statutory construction (definition & purpose)\n  - F. Ambiguity / Vagueness\n  - G. Statutory construction v. judicial legislation",
        "latin_maxims": [
            ("Noscitur a sociis (words are known by their companions)", "A word is known by the company it keeps.", "When a statutory term like 'gift enterprise' is ambiguous, its meaning should be understood in connection with associated words ('lottery'), requiring consideration."),
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "Administrative officers cannot expand clear statutory terms to penalize lawful, free promotional campaigns."),
            ("Optima statuti interpretatrix est ipsum statutum (the best interpreter of a statute is the statute itself)", "The best interpreter of a statute is the statute itself.", "Statutes must be construed according to their internal context and established legal definitions.")
        ],
        "primary_maxim": "Noscitur a sociis\n(a word is known by its company / companions)",
        "secondary_maxims": "• Verba legis non est recedendum (from the words of the statute there should be no departure)\n• Optima statuti interpretatrix est ipsum statutum (the best interpreter of a statute is the statute itself)",
        "violation_text": "The Postmaster General committed an administrative violation of statutory construction by attempting to expand the definition of 'lottery' and 'gift enterprise' beyond their settled legal components to ban a free contest.",
        "execution_text": "Defines the essence of Statutory Construction: Courts interpret statutory terms in accordance with established legal definitions, preventing administrative officers from rewriting laws through administrative fiat.",
        "statcon_plain": "This case is the classic introduction to StatCon. It shows that words must be interpreted according to their legal elements and associations (Noscitur a sociis - words are known by their associates). Administrative agencies cannot rewrite the law to create new bans.",
        "statcon_recite": "\"Sir/Ma'am, in Chapter I of our syllabus: (1) Primary Doctrine: Noscitur a sociis (words are known by their associates)—'gift enterprise' takes its character from 'lottery'; (2) StatCon vs. Administrative Legislation: Administrative agencies cannot expand statutory prohibitions beyond the explicit legal elements enacted by Congress.\""
    },

    # 2. Tañada v. Tuvera
    {
        "num": 2,
        "title": "Lorenzo M. Tañada, et al. v. Hon. Juan C. Tuvera, etc., et al.",
        "citation": "G.R. No. L-63915 | December 29, 1986 | 146 SCRA 446 | En Banc | Ponente: J. Cruz",
        "plaintiff": "Lorenzo M. Tañada, Abraham F. Sarmiento, and MABINI (Petitioners)",
        "defendant": "Hon. Juan C. Tuvera, Executive Secretary, et al. (Respondents)",
        "court": "Supreme Court of the Philippines (Original Action for Mandamus)",
        "nature": "Petition for Mandamus under Rule 65 to compel the publication in the Official Gazette of presidential decrees, executive orders, and administrative issuances.",
        "facts": "During martial law under President Ferdinand Marcos, numerous presidential decrees, general orders, and executive orders were enforced without publication in the Official Gazette. Lorenzo Tañada and public-interest lawyers petitioned for mandamus. The Supreme Court in 1985 ordered publication, and in this 1986 Resolution on Reconsideration, the Court firmly settled the publication mandate under Article 2 of the Civil Code.",
        "plain_facts": "The Marcos government enforced secret presidential decrees without publishing them. Lawyers sued to compel publication, asserting that people cannot follow secret laws.",
        "issue": "Whether or not the clause 'unless it is otherwise provided' in Article 2 of the Civil Code dispenses with the mandatory requirement of publication of laws in the Official Gazette.",
        "issue_plain": "If a law says it takes effect 'immediately', does that allow the government to skip publishing it in the Official Gazette?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether the phrase 'unless it is otherwise provided' in Article 2 of the Civil Code dispenses with publication altogether, or merely modifies the 15-day effectivity period.\"",
        "answer": "No. The clause 'unless it is otherwise provided' refers solely to the date of effectivity, but never dispenses with the absolute requirement of full publication.",
        "answer_plain": "No. Publication is 100% mandatory for all laws. The clause only lets Congress change the 15-day waiting period.",
        "answer_recite": "\"No, Sir/Ma'am. Publication in the Official Gazette or a newspaper of general circulation is an indispensable requirement of due process. The clause 'unless it is otherwise provided' modifies only the transition period.\"",
        "legal_basis": [
            ("1. Article 2, Civil Code:", "Mandates publication as a condition for statutory effectivity."),
            ("2. Article III, Section 1, 1987 Constitution:", "Procedural Due Process Clause."),
            ("3. Article 3, Civil Code:", "Ignorantia legis non excusat (ignorance of the law excuses no one).")
        ],
        "legal_basis_plain": "Article 2 of the Civil Code, Due Process Clause, and Article 3 (ignorance of the law is no excuse).",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis rests on Article 2 and 3 of the Civil Code and the Due Process Clause of the Constitution.\"",
        "analysis": [
            ("1. Due Process Requirement:", "It violates due process to penalize citizens under secret laws while applying the rule that ignorance of the law excuses no one."),
            ("2. Scope of Clause:", "'Unless it is otherwise provided' modifies the 15-day period, not the publication itself."),
            ("3. All Laws Included:", "All statutes, presidential decrees, and administrative circulars with penal sanctions must be published in full.")
        ],
        "analysis_plain": "Secret laws violate due process. The law must be published so the public has notice before anyone can be held liable.",
        "analysis_recite": "\"Sir/Ma'am, the Court harmonized Article 2 with due process. Publication is absolute; without publication, laws are void.\"",
        "conclusion": "The Supreme Court declared that all unpublished presidential decrees and executive issuances are of NO FORCE AND EFFECT.",
        "conclusion_plain": "Unpublished laws are void and legally non-existent.",
        "conclusion_recite": "\"The Supreme Court granted mandamus, holding that unpublished laws are completely unenforceable.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - A. Statutes (validity and effectivity)\n• Chapter II: Prelude to the Exercise of the Power\n  - A. Verba legis non est recedendum (from the words of the statute there should be no departure)\n  - B. Absoluta sententia expositore non indigent (when language is plain, it needs no interpreter)",
        "latin_maxims": [
            ("Ignorantia legis non excusat (ignorance of the law excuses no one)", "Ignorance of the law excuses no one.", "Operates only when laws are published and made accessible to the public."),
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "Article 2 literally demands publication."),
            ("Lex non cogit ad impossibilia (the law does not compel the impossible)", "The law does not compel the impossible.", "Citizens cannot be compelled to obey secret laws.")
        ],
        "primary_maxim": "Ignorantia legis non excusat\n(ignorance of the law excuses no one / demands publication as a prerequisite)",
        "secondary_maxims": "• Verba legis non est recedendum (from the words of the statute there should be no departure)\n• Lex non cogit ad impossibilia (the law does not compel the impossible)",
        "violation_text": "Executive branch treated 'unless it is otherwise provided' as an excuse to enforce secret laws without publication.",
        "execution_text": "Establishes publication as the absolute prerequisite for the validity and binding effect of all statutes.",
        "statcon_plain": "Proves that without publication, no statute exists to be construed.",
        "statcon_recite": "\"Sir/Ma'am, Tañada v. Tuvera defines the threshold requirement of publication under Article 2 of the Civil Code.\""
    },

    # 3. Romualdez v. Sandiganbayan
    {
        "num": 3,
        "title": "Alfredo Romualdez v. Sandiganbayan (Fifth Division)",
        "citation": "G.R. No. 152259 | July 29, 2004 | 435 SCRA 371 | En Banc | Ponente: J. Panganiban",
        "plaintiff": "Alfredo T. Romualdez (Petitioner)",
        "defendant": "Honorable Sandiganbayan and People of the Philippines (Respondents)",
        "court": "Sandiganbayan (Fifth Division)",
        "nature": "Petition for Certiorari challenging Section 5 of R.A. No. 3019 for alleged unconstitutional vagueness.",
        "facts": "Alfredo Romualdez was charged with violating Section 5 of R.A. No. 3019 for intervening in a contract with the National Power Corporation. He claimed the law was facially void for vagueness because 'intervene' was not defined.",
        "plain_facts": "Romualdez was charged with graft for meddling in a state contract and claimed the law was too vague to be constitutional.",
        "issue": "Whether Section 5 of R.A. No. 3019 is void for vagueness on its face.",
        "issue_plain": "Can a penal statute be struck down on facial vagueness?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether Section 5 of RA 3019 is void for vagueness, testing the limits of facial challenges on penal laws.\"",
        "answer": "No. Facial challenges on vagueness are not allowed against penal statutes (only as-applied challenges, except in free speech cases). The word 'intervene' has a clear common meaning.",
        "answer_plain": "No. The law is valid. Facial challenges do not apply to criminal statutes, and 'intervene' is well-understood.",
        "answer_recite": "\"No, Sir/Ma'am. Facial challenges for vagueness do not apply to penal statutes. 'Intervene' has a clear ordinary meaning.\"",
        "legal_basis": [
            ("1. Section 5, R.A. No. 3019:", "Anti-Graft prohibition against intervention by presidential relatives."),
            ("2. Presumptio pro iustitia legis (presumption in favor of the constitutionality of a law):", "Statutes are presumed valid."),
            ("3. Verba generalia generaliter sunt intelligenda (general words are understood generally):", "Words are understood in common usage.")
        ],
        "legal_basis_plain": "Section 5 of RA 3019, Presumption of Constitutionality, and the Plain Meaning Rule.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Presumptio pro iustitia legis (presumption of constitutionality) and Verba generalia generaliter sunt intelligenda (general words are understood generally).\"",
        "analysis": [
            ("1. Inapplicability of Facial Challenge:", "Facial challenge is limited to free speech cases."),
            ("2. Ordinary Meaning:", "'Intervene' means to consciously get involved or meddle."),
            ("3. Presumption of Validity:", "Courts must interpret statutes to sustain their constitutionality.")
        ],
        "analysis_plain": "Criminal laws cannot be invalidated facially. Common words like 'intervene' give fair notice to the public.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied ordinary meaning to 'intervene' and rejected the facial challenge on a penal law.\"",
        "conclusion": "The Supreme Court DISMISSED the petition and ordered the graft trial to proceed.",
        "conclusion_plain": "Romualdez lost and his trial proceeded.",
        "conclusion_recite": "\"The Supreme Court sustained the validity of Section 5 of RA 3019.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - F. Ambiguity / Vagueness (limits of vagueness doctrine)",
        "latin_maxims": [
            ("Presumptio pro iustitia legis (there is a presumption in favor of the justice and constitutionality of a law)", "There is a presumption in favor of the justice and constitutionality of a law.", "Statutes are presumed constitutional."),
            ("Verba generalia generaliter sunt intelligenda (general words are to be understood generally)", "General words are to be understood in a general sense.", "Statutory words are given their natural meaning."),
            ("Ut res magis valeat quam pereat (that the thing may rather have effect than perish)", "That the thing may rather have effect than perish.", "Construing a statute to maintain its validity.")
        ],
        "primary_maxim": "Presumptio pro iustitia legis\n(presumption in favor of the justice and constitutionality of a law)",
        "secondary_maxims": "• Verba generalia generaliter sunt intelligenda (general words are understood generally)\n• Ut res magis valeat quam pereat (that the thing may rather have effect than perish)",
        "violation_text": "Petitioner tried to strike down a penal statute through an improper facial challenge on vagueness.",
        "execution_text": "Establishes that broad statutory words in penal laws do not render the statute void for vagueness.",
        "statcon_plain": "A law is not unconstitutionally vague if ordinary words can be understood by their plain meaning.",
        "statcon_recite": "\"Sir/Ma'am, Romualdez limits the void-for-vagueness doctrine to as-applied challenges in penal laws.\""
    },

    # 4. Pesca v. Pesca
    {
        "num": 4,
        "title": "Lorna G. Pesca v. Zosimo A. Pesca",
        "citation": "G.R. No. 136921 | April 17, 2001 | 356 SCRA 588 | Second Division | Ponente: J. Gonzaga-Reyes",
        "plaintiff": "Lorna G. Pesca (Petitioner)",
        "defendant": "Zosimo A. Pesca (Respondent)",
        "court": "RTC Pasig / Court of Appeals",
        "nature": "Petition for Review seeking declaration of nullity of marriage under Article 36 of the Family Code.",
        "facts": "Lorna Pesca filed for nullity of marriage due to husband's physical abuse and emotional cruelty. The CA applied the Molina guidelines. Lorna argued Molina should not apply retroactively to her pending case.",
        "plain_facts": "Wife sued to annul marriage for cruelty and argued strict new Supreme Court guidelines (Molina) should not apply retroactively.",
        "issue": "Whether Supreme Court decisions interpreting statutes apply retroactively to pending cases.",
        "issue_plain": "Do judicial interpretations apply retroactively to existing cases?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether judicial interpretations of statutes apply retroactively under Legis interpretado legis vim obtinet (the interpretation of the law has the force of law).\"",
        "answer": "Yes. Judicial decisions interpreting a statute form part of that law from the date of its enactment and apply retroactively to all pending cases.",
        "answer_plain": "Yes. Judicial interpretations are retroactive because courts only declare what the statute always meant.",
        "answer_recite": "\"Yes, Sir/Ma'am. Under Article 8 of the Civil Code, judicial interpretations apply retroactively to the enactment date of the law.\"",
        "legal_basis": [
            ("1. Article 8, Civil Code:", "Judicial decisions form part of the legal system."),
            ("2. Article 36, Family Code:", "Psychological incapacity grounds."),
            ("3. Legis interpretado legis vim obtinet (the interpretation of the law has the force of law):", "Judicial interpretation has the force of law.")
        ],
        "legal_basis_plain": "Article 8 of the Civil Code and the doctrine of retroactivity of judicial interpretations.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Article 8 of the Civil Code and Legis interpretado legis vim obtinet (the interpretation of the law has the force of law).\"",
        "analysis": [
            ("1. Interpretative Function:", "Courts do not make new laws; they reveal original legislative intent."),
            ("2. Retroactive Effect:", "The Molina guidelines apply retroactively to the date of effectivity of the Family Code in 1988."),
            ("3. Application:", "Physical abuse alone is a ground for legal separation, not psychological incapacity under Article 36.")
        ],
        "analysis_plain": "Judicial rulings are retroactive because they reveal what the statute always meant from the beginning.",
        "analysis_recite": "\"Sir/Ma'am, the Court held that judicial interpretations are retroactive because they declare what the statute always meant.\"",
        "conclusion": "The Supreme Court AFFIRMED the dismissal of the nullity petition.",
        "conclusion_plain": "Nullity denied; petitioner was advised to file for legal separation.",
        "conclusion_recite": "\"The Supreme Court affirmed the CA, holding that judicial interpretations apply retroactively.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - H. Legis interpretado legis vim obtinet (interpretation has the force of law)\n  - I. Kinds of construction -> b. Prospective v. retrospective",
        "latin_maxims": [
            ("Legis interpretado legis vim obtinet (the interpretation of the law has the force of law)", "The interpretation of the law acquires the force of the law itself.", "Judicial decisions interpreting statutes form part of the law itself under Article 8."),
            ("Lex prospicit, non respicit (the law looks forward, not backward)", "The law looks forward, not backward.", "Statutes look forward, but judicial interpretations look backward to enactment."),
            ("Stare decisis et non quieta movere (to stand by decisions and not disturb what is settled)", "To stand by decisions and not disturb settled matters.", "Precedents ensure stability in statutory meaning.")
        ],
        "primary_maxim": "Legis interpretado legis vim obtinet\n(the interpretation of the law has the force of law itself)",
        "secondary_maxims": "• Lex prospicit, non respicit (statutes look forward, interpretations apply retroactively)\n• Stare decisis et non quieta movere (stand by settled decisions)",
        "violation_text": "Petitioner erroneously treated judicial statutory construction as a prospective legislative enactment.",
        "execution_text": "Defines the retroactivity of judicial construction: Supreme Court interpretations apply retroactively to all pending cases.",
        "statcon_plain": "Supreme Court rulings interpreting statutes apply retroactively to all pending cases.",
        "statcon_recite": "\"Sir/Ma'am, Pesca v. Pesca establishes the retroactivity of judicial interpretation under Article 8 of the Civil Code.\""
    },

    # 5. People v. Mapa
    {
        "num": 5,
        "title": "People of the Philippines v. Mario Mapa y Mapulong",
        "citation": "G.R. No. L-22301 | August 30, 1967 | 20 SCRA 1164 | En Banc | Ponente: J. Fernando",
        "plaintiff": "People of the Philippines (Plaintiff-Appellee)",
        "defendant": "Mario Mapa y Mapulong (Accused-Appellant)",
        "court": "Court of First Instance of Manila",
        "nature": "Criminal prosecution for Illegal Possession of Firearms under Section 878 RAC.",
        "facts": "Mario Mapa was caught with an unlicensed gun and argued he was exempt as a governor's secret agent.",
        "plain_facts": "Mapa claimed his governor's secret agent appointment exempted him from firearm licensing laws.",
        "issue": "Whether a secret agent is exempt from firearm licensing under Section 879 RAC.",
        "issue_plain": "Does a secret agent badge excuse illegal gun possession?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether secret agents are exempt under Section 879 of the Revised Administrative Code.\"",
        "answer": "No. Secret agents are not included in the statutory list of exempt officers under Section 879 RAC.",
        "answer_plain": "No. Secret agents are not exempt; only listed officers are exempt.",
        "answer_recite": "\"No, Sir/Ma'am. Section 879 RAC strictly enumerates exempt officers and does not include secret agents.\"",
        "legal_basis": [
            ("1. Section 878 & 879 RAC:", "Penalizes possession and lists exempt officials."),
            ("2. Verba Legis (the words of the law):", "Clear words require no interpretation."),
            ("3. Expressio unius est exclusio alterius (the express mention of one thing is the exclusion of another):", "Enumeration excludes omitted items.")
        ],
        "legal_basis_plain": "Section 878 & 879 RAC and the Verba Legis rule.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Section 879 RAC and Verba legis non est recedendum (from the words of the statute there should be no departure).\"",
        "analysis": [
            ("1. Clear Text:", "Secret agents are omitted from the statute."),
            ("2. Overruled Precedent:", "Macarandang and Lucero were overruled to prevent judicial legislation."),
            ("3. Strict Application:", "Courts must apply the law as written.")
        ],
        "analysis_plain": "The statute is clear; judges cannot add exceptions that Congress omitted.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Verba Legis and overruled prior cases to prevent judicial legislation.\"",
        "conclusion": "The Supreme Court AFFIRMED the conviction of Mario Mapa.",
        "conclusion_plain": "Conviction affirmed.",
        "conclusion_recite": "\"The Supreme Court affirmed the conviction under Verba Legis.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction (E, F, G, J)\n• Chapter II: Prelude to the Exercise of the Power (A, B, C)",
        "latin_maxims": [
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "The Plain Meaning Rule."),
            ("Absoluta sententia expositore non indigent (when language is plain, it needs no interpreter)", "Plain language needs no interpreter.", "Clear statutes apply automatically."),
            ("Dura lex, sed lex (the law is harsh, but it is the law)", "The law is harsh, but it is the law.", "Hardship cannot justify altering clear statutes.")
        ],
        "primary_maxim": "Verba legis non est recedendum\n(from the words of the statute there should be no departure)",
        "secondary_maxims": "• Absoluta sententia expositore non indigent (plain language needs no interpreter)\n• Dura lex, sed lex (the law is harsh, but it is the law)",
        "violation_text": "Appellant urged the court to insert an unwritten exemption into a penal law.",
        "execution_text": "Leading authority on Verba Legis: When the law is clear, there is zero room for construction.",
        "statcon_plain": "Never interpret when the law is clear.",
        "statcon_recite": "\"Sir/Ma'am, People v. Mapa is the leading authority on Verba legis non est recedendum (from the words of the statute there should be no departure).\""
    }
]

print("Phase 1 cases script part 1 ready.")
