# build_phase1_cases.py
import os
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from core_builder import (
    create_docx_builder, add_header_box, add_heading_1, add_heading_2,
    add_body_p, add_student_explanation_box, add_alac_box,
    add_latin_maxims_box, add_statcon_table
)

phase1_cases = [
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
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis rests on Section 1954-A of the Revised Administrative Code and the three elements of a lottery: Consideration, Chance, and Prize. Under Noscitur a sociis (known by its associates), 'gift enterprise' must be construed in conjunction with 'lottery', both requiring consideration.\"",
        "analysis": [
            ("1. Absence of Consideration:", "The three elements of a lottery are consideration, chance, and prize. While chance and prize were present, consideration was completely absent. Participants did not pay any money or buy any merchandise."),
            ("2. Definition of Gift Enterprise:", "Under the rule of Noscitur a sociis (words are known by their associates), 'gift enterprise' is grouped with 'lottery' and must partake of the same nature—meaning the distribution of prizes must be coupled with the purchase of goods. A purely gratuitous contest is not a gift enterprise."),
            ("3. Judicial Construction vs. Administrative Expansion:", "The Postmaster General improperly attempted to expand the statute by punishing an act that lacked the essential element of consideration, which constitutes unauthorized administrative legislation.")
        ],
        "analysis_plain": "To be an illegal lottery, there must be (1) Prize, (2) Chance, and (3) Consideration (paying money). Here, there was prize and chance, but NO consideration because it was completely free. Under Noscitur a sociis, 'gift enterprise' also requires consideration. The Postmaster General had no right to expand the law.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Noscitur a sociis (known by its associates). Since 'gift enterprise' is associated with 'lottery', it must involve consideration (payment). Because Caltex required no purchase or fee, the element of consideration was absent. The Postmaster General engaged in administrative overreach by attempting to expand a prohibitory statute.\"",
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
        "statcon_plain": "This case is the classic introduction to StatCon. It shows that words must be interpreted according to their legal elements and associations (Noscitur a sociis). Administrative agencies cannot rewrite the law to create new bans.",
        "statcon_recite": "\"Sir/Ma'am, in Chapter I of our syllabus: (1) Primary Doctrine: Noscitur a sociis (words are known by their associates)—'gift enterprise' takes its character from 'lottery'; (2) StatCon vs. Administrative Legislation: Administrative agencies cannot expand statutory prohibitions beyond the explicit legal elements enacted by Congress.\""
    },
    {
        "num": 2,
        "title": "Lorenzo M. Tañada, et al. v. Hon. Juan C. Tuvera, etc., et al.",
        "citation": "G.R. No. L-63915 | December 29, 1986 | 146 SCRA 446 | En Banc | Ponente: J. Cruz",
        "plaintiff": "Lorenzo M. Tañada, Abraham F. Sarmiento, and Movement of Attorneys for Brotherhood, Integrity and Nationalism, Inc. (MABINI) (Petitioners)",
        "defendant": "Hon. Juan C. Tuvera, Executive Secretary to the President, et al. (Respondents)",
        "court": "Supreme Court of the Philippines (Original Petition for Mandamus)",
        "nature": "Petition for Mandamus under Rule 65 to compel the publication in the Official Gazette of presidential decrees, executive orders, and administrative issuances.",
        "facts": "During the martial law regime of President Ferdinand Marcos, numerous presidential decrees, letters of instructions, general orders, and executive orders of general applicability were enforced without being published in the Official Gazette. Petitioners Lorenzo Tañada and public-interest lawyers filed a petition for mandamus to compel publication. In the initial 1985 decision, the Court ordered publication. In the 1986 Resolution on the Motion for Reconsideration, the Supreme Court clarified the scope of the publication requirement under Article 2 of the Civil Code, especially concerning laws that state they shall 'take effect immediately upon approval'.",
        "plain_facts": "The Marcos government passed hundreds of secret presidential decrees and enforced them against citizens without printing them in the Official Gazette. Lawyers sued to force the government to publish every law, arguing that secret laws violate due process.",
        "issue": "Whether or not the clause 'unless it is otherwise provided' in Article 2 of the Civil Code dispenses with the mandatory requirement of publication of laws in the Official Gazette.",
        "issue_plain": "If a law states that it takes effect 'immediately upon approval', does that mean the government can skip publishing it in the Official Gazette?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether the phrase 'unless it is otherwise provided' in Article 2 of the Civil Code allows the legislature to dispense with publication altogether, or whether it only allows variation of the 15-day effectivity period.\"",
        "answer": "No. The clause 'unless it is otherwise provided' does not dispense with the requirement of publication. Publication in full in the Official Gazette (or newspaper of general circulation) is an indispensable, mandatory condition for the effectivity of all laws.",
        "answer_plain": "No. Publication is NEVER optional. All laws must be published in full. The phrase 'unless it is otherwise provided' only refers to the date of effectivity, not the requirement of publication itself.",
        "answer_recite": "\"No, Sir/Ma'am. Publication is an indispensable requirement of due process. The phrase 'unless it is otherwise provided' refers strictly to the date of effectivity (e.g., shortening or lengthening the 15-day period), but never dispenses with publication itself.\"",
        "legal_basis": [
            ("1. Article 2, Civil Code of the Philippines:", "'Laws shall take effect after fifteen days following the completion of their publication in the Official Gazette, unless it is otherwise provided.'"),
            ("2. Article III, Section 1, 1987 Constitution:", "Due process of law clause. No person shall be deprived of life, liberty, or property without due process of law."),
            ("3. Article 3, Civil Code:", "'Ignorance of the law excuses no one from compliance therewith' (Ignorantia legis non excusat).")
        ],
        "legal_basis_plain": "Article 2 of the Civil Code (publication rule), the Due Process Clause of the Constitution, and Article 3 of the Civil Code (ignorance of the law excuses no one).",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis rests on Article 2 and Article 3 of the Civil Code, harmonized with the Due Process Clause of the Constitution. For 'Ignorantia legis non excusat' to apply, prior publication is mandatory.\"",
        "analysis": [
            ("1. Harmonization with Due Process:", "The Court interpreted Article 2 in light of constitutional due process. It would be the height of injustice to punish citizens for violating secret, unpublished laws while presuming that ignorance of the law excuses no one."),
            ("2. Meaning of 'Unless it is otherwise provided':", "The clause 'unless it is otherwise provided' solely modifies the 15-day transition period (e.g., a law may take effect on a specified date or immediately), but it can NEVER be construed as authorizing non-publication."),
            ("3. Scope of Laws Covered:", "All presidential decrees, executive orders, administrative circulars enforcing penal sanctions, and charter laws of general application must be published in full.")
        ],
        "analysis_plain": "People cannot follow laws they cannot read. The Civil Code says ignorance of the law is no excuse, so the government MUST publish every law before enforcing it. The words 'unless it is otherwise provided' only mean Congress can change the 15-day waiting period, but they can NEVER skip printing the law.",
        "analysis_recite": "\"Sir/Ma'am, the Court applied Constitutional and Harmonious Construction. The phrase 'unless it is otherwise provided' refers exclusively to the 15-day period. Publication is absolute; without publication, a law has no binding legal effect whatsoever because secret laws violate procedural due process.\"",
        "conclusion": "The Supreme Court declared that all presidential decrees and executive issuances of general application are of NO FORCE AND EFFECT until published in full in the Official Gazette.",
        "conclusion_plain": "The lawyers won. The Supreme Court ruled that unpublished laws are completely void and unenforceable.",
        "conclusion_recite": "\"The Supreme Court granted mandamus, ruling that publication in full in the Official Gazette is an absolute condition precedent for the effectivity of all statutes and decrees.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - A. Statutes (enactment, validity, effectivity)\n• Chapter II: Prelude to the Exercise of the Power\n  - A. Verba legis non est recedendum\n  - B. Absoluta sententia expositore non indigent",
        "latin_maxims": [
            ("Ignorantia legis non excusat (ignorance of the law excuses no one)", "Ignorance of the law excuses no one.", "A foundational legal maxim that can only morally and legally operate if laws are first made publicly available through full publication."),
            ("Verba legis non est recedendum (from the words of the statute there should be no departure)", "From the words of a statute there should be no departure.", "Article 2 explicitly requires publication following its literal command."),
            ("Lex non cogit ad impossibilia (the law does not compel the impossible)", "The law does not compel a man to do that which he cannot possibly do.", "Citizens cannot be compelled to obey secret laws they have no means of knowing.")
        ],
        "primary_maxim": "Ignorantia legis non excusat\n(ignorance of the law excuses no one / demands publication as a prerequisite)",
        "secondary_maxims": "• Verba legis non est recedendum (from the words of the statute there should be no departure)\n• Lex non cogit ad impossibilia (the law does not compel the impossible)",
        "violation_text": "The executive branch violated constitutional due process and statutory construction by treating the clause 'unless it is otherwise provided' as a license to bypass the publication requirement entirely.",
        "execution_text": "Cornerstone of Statutory Validity: Defines the mandatory prerequisite for any statute to become law, ensuring that statutory interpretation always upholds procedural due process and public notice.",
        "statcon_plain": "Teaches that Article 2 of the Civil Code makes publication 100% mandatory. Without publication, there is no valid statute to interpret or apply.",
        "statcon_recite": "\"Sir/Ma'am, under Chapter I (A. Statutes): Publication is an indispensable element of statutory validity. Tañada v. Tuvera harmonizes statutory interpretation with constitutional due process, establishing that secret laws are absolute nullities.\""
    },
    {
        "num": 3,
        "title": "Alfredo Romualdez v. Sandiganbayan (Fifth Division)",
        "citation": "G.R. No. 152259 | July 29, 2004 | 435 SCRA 371 | En Banc | Ponente: J. Panganiban",
        "plaintiff": "Alfredo T. Romualdez (Petitioner)",
        "defendant": "Honorable Sandiganbayan (Fifth Division) and People of the Philippines (Respondents)",
        "court": "Sandiganbayan (Fifth Division)",
        "nature": "Petition for Certiorari and Prohibition under Rule 65 challenging the constitutionality of Section 5 of Republic Act No. 3019 (The Anti-Graft and Corrupt Practices Act) for alleged vagueness.",
        "facts": "Alfredo Romualdez, brother-in-law of former President Ferdinand Marcos, was charged with violating Section 5 of R.A. No. 3019 for unlawfully intervening in a contract between the National Power Corporation (NPC) and a private shipping firm. Romualdez filed a motion to quash, mounting a 'facial challenge' on the ground that Section 5 was unconstitutionally vague because the terms 'intervene' and 'family relation' were not precisely defined in the statute. The Sandiganbayan denied his motion. Romualdez petitioned the Supreme Court.",
        "plain_facts": "Marcos's brother-in-law was prosecuted for graft for meddling in a government contract. He tried to have the charge dismissed by claiming the graft law was unconstitutionally vague because words like 'intervene' were unclear.",
        "issue": "Whether or not Section 5 of R.A. No. 3019 is unconstitutionally vague on its face, and whether a facial challenge on vagueness grounds can be used to invalidate a penal statute.",
        "issue_plain": "Can a criminal defendant strike down an entire anti-graft law by claiming its wording is vague on its face?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether Section 5 of the Anti-Graft Law is void for vagueness, and whether the 'facial challenge' doctrine can be applied to strike down penal statutes outside the realm of free speech.\"",
        "answer": "No. Section 5 of R.A. No. 3019 is constitutional and NOT unconstitutionally vague. A facial challenge on the ground of vagueness cannot be used to invalidate penal statutes; it can only be raised in 'as-applied' challenges (except in free speech cases).",
        "answer_plain": "No. The law is constitutional. Criminal laws cannot be struck down using facial challenges unless they involve free speech. The word 'intervene' is easily understood in context.",
        "answer_recite": "\"No, Sir/Ma'am. A facial challenge on vagueness grounds is not applicable to penal statutes, as they must be evaluated 'as applied' to the defendant's specific conduct. Moreover, terms like 'intervene' have clear, ordinary, and statutory meanings.\"",
        "legal_basis": [
            ("1. Section 5, Republic Act No. 3019:", "Penalizes relatives of the President, Vice-President, and cabinet members who intervene in government contracts."),
            ("2. Presumption of Constitutionality:", "Every statute is presumed constitutional; courts must favor a construction that validates the statute."),
            ("3. Plain Meaning / Words in Common Usage:", "Words in a statute that are not technically defined are understood in their ordinary, dictionary, and contextual sense.")
        ],
        "legal_basis_plain": "Section 5 of RA 3019, the Presumption of Constitutionality, and the rule that everyday statutory words are construed according to their common, ordinary meaning.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis rests on the Presumption of Constitutionality and the rule on Plain and Ordinary Meaning of Words. Facial challenges on vagueness are restricted to First Amendment/free speech cases.\"",
        "analysis": [
            ("1. Inapplicability of Facial Challenge to Penal Laws:", "The 'facial challenge' doctrine is an exceptional tool reserved for free speech cases to prevent chilling effects. In criminal law, a statute must be examined 'as applied' to the actual facts of the defendant's case."),
            ("2. Plain Meaning of 'Intervene':", "The word 'intervene' is not vague; it has a clear dictionary and statutory meaning: to consciously enter into or influence a transaction."),
            ("3. Statutory Purpose:", "Construing Section 5 broadly to uphold its validity fulfills the legislative purpose of deterring corruption and nepotism in government transactions.")
        ],
        "analysis_plain": "You cannot throw out a whole criminal law just by claiming words are vague in abstract theory; you have to test it on the actual facts of the case. 'Intervene' is a normal English word that means getting involved or meddling. The law is clear enough to give fair warning.",
        "analysis_recite": "\"Sir/Ma'am, the Court held that penal statutes cannot be challenged facially for vagueness. Furthermore, statutory words are construed in their common and commercial meaning. 'Intervene' provides sufficient fair notice to citizens of prohibited corrupt conduct.\"",
        "conclusion": "The Supreme Court DISMISSED the petition and AFFIRMED the Sandiganbayan's denial of the motion to quash, ordering the criminal trial of Romualdez to proceed.",
        "conclusion_plain": "Romualdez lost. The Supreme Court upheld the anti-graft law and ordered his corruption trial to continue.",
        "conclusion_recite": "\"The Supreme Court sustained the constitutionality of Section 5 of RA 3019 and affirmed the order for the trial of Romualdez to proceed.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - F. Ambiguity / Vagueness (Vagueness Doctrine limits)\n• Chapter VI: Construction of Specific Statutes\n  - B. Statutes strictly construed (Penal statutes vs. facial challenge)",
        "latin_maxims": [
            ("Presumptio pro iustitia legis (there is a presumption in favor of the justice and constitutionality of a law)", "There is a presumption in favor of the justice and validity of the law.", "Courts must presume statutes are constitutional and refrain from striking them down on speculative facial claims."),
            ("Verba generalia generaliter sunt intelligenda (general words are to be understood generally)", "General words are to be understood in a general sense.", "Words like 'intervene' must be understood in their ordinary, plain English meaning."),
            ("Ut res magis valeat quam pereat (that the thing may rather have effect than perish)", "That the thing may rather have effect than perish.", "Statutes should be construed to preserve their validity and effectiveness rather than destroyed.")
        ],
        "primary_maxim": "Presumptio pro iustitia legis\n(presumption in favor of the justice and constitutionality of a statute)",
        "secondary_maxims": "• Verba generalia generaliter sunt intelligenda (general words understood generally)\n• Ut res magis valeat quam pereat (that the thing may rather have effect than perish)",
        "violation_text": "The petitioner attempted to misapply the vagueness doctrine via a facial challenge to invalidate an anti-graft penal statute, violating the presumption of constitutionality.",
        "execution_text": "Establishes strict limits on the Vagueness Doctrine: Penal statutes cannot be invalidated through facial challenges on vagueness if statutory terms can be reasonably understood through ordinary dictionary and contextual meaning.",
        "statcon_plain": "Clarifies that a law is not void for vagueness just because it uses broad words. Judges must look at common meaning and presume the law is valid.",
        "statcon_recite": "\"Sir/Ma'am, under Chapter I (F. Ambiguity / Vagueness): Romualdez establishes that the void-for-vagueness doctrine cannot be used facially to strike down penal laws; words in common usage are presumed clear under Verba generalia generaliter sunt intelligenda.\""
    },
    {
        "num": 4,
        "title": "Lorna G. Pesca v. Zosimo A. Pesca",
        "citation": "G.R. No. 136921 | April 17, 2001 | 356 SCRA 588 | Second Division | Ponente: J. Gonzaga-Reyes",
        "plaintiff": "Lorna G. Pesca (Petitioner)",
        "defendant": "Zosimo A. Pesca (Respondent)",
        "court": "Regional Trial Court (RTC) of Pasig City (Branch 160) / Court of Appeals",
        "nature": "Petition for Review on Certiorari under Rule 45 seeking the declaration of nullity of marriage under Article 36 of the Family Code.",
        "facts": "Lorna Pesca filed a petition for declaration of nullity of marriage against her husband Zosimo on the ground of psychological incapacity under Article 36 of the Family Code. She alleged that Zosimo was emotionally immature, irresponsible, habitually drank, and physically battered her. The RTC declared the marriage void. The Court of Appeals reversed, citing the strict guidelines established in the landmark Supreme Court cases of Santos v. CA (1995) and Republic v. Molina (1997). On appeal to the Supreme Court, Lorna argued that the Molina guidelines should not apply retroactively to her case since she filed her lawsuit before the Molina decision was promulgated.",
        "plain_facts": "A wife sued to annul her marriage because of domestic violence and cruelty, claiming it proved psychological incapacity. While her case was ongoing, the Supreme Court issued strict guidelines in the Molina case. She argued the Molina rules should not apply retroactively to her case.",
        "issue": "Whether or not the guidelines interpreting 'psychological incapacity' in Republic v. Molina apply retroactively to pending cases filed prior to the promulgation of the Molina decision.",
        "issue_plain": "Do Supreme Court decisions interpreting a statute apply retroactively to all pending cases, or only prospectively like new laws?",
        "issue_recite": "\"Sir/Ma'am, the issue is whether judicial interpretations of statutes apply retroactively to pending actions, testing the fundamental doctrine of Legis interpretado legis vim obtinet under Article 8 of the Civil Code.\"",
        "answer": "Yes. Judicial decisions interpreting the law form part of the legal system and apply retroactively from the date the statute was enacted. The Molina guidelines applied to Lorna's case.",
        "answer_plain": "Yes. When the Supreme Court interprets a law, it does not make a new law; it merely explains what the law has meant from day one. Therefore, judicial decisions apply retroactively to all pending cases.",
        "answer_recite": "\"Yes, Sir/Ma'am. Under Article 8 of the Civil Code, judicial interpretations of a statute form part of that statute from the date of its original enactment. Consequently, judicial interpretations apply retroactively to all pending actions.\"",
        "legal_basis": [
            ("1. Article 8, Civil Code of the Philippines:", "'Judicial decisions applying or interpreting the laws or the Constitution shall form a part of the legal system of the Philippines.'"),
            ("2. Article 36, Family Code (E.O. No. 209):", "Psychological incapacity as a ground for nullity of marriage."),
            ("3. Doctrine of Legis Interpretado Legis Vim Obtinet:", "The interpretation of a law by the Supreme Court acquires the force of the law itself.")
        ],
        "legal_basis_plain": "Article 8 of the Civil Code (judicial decisions form part of the legal system), Article 36 of the Family Code, and the doctrine that judicial interpretation dates back to the enactment of the statute.",
        "legal_basis_recite": "\"Sir/Ma'am, the legal basis is Article 8 of the Civil Code and the maxim Legis interpretado legis vim obtinet (the interpretation of the law has the force of law). Unlike statutes which are prospective, judicial interpretations apply retroactively.\"",
        "analysis": [
            ("1. Nature of Judicial Interpretation:", "The Supreme Court does not pass new legislation when interpreting a statute; it merely declares the contemporaneous meaning the legislature intended from the date of enactment."),
            ("2. Retroactivity of Jurisprudence:", "Because the Molina guidelines merely clarified what Article 36 of the Family Code has always meant since its enactment in 1988, the guidelines necessarily govern all pending cases involving marriages under the Family Code."),
            ("3. Application to Facts:", "While physical violence and emotional immaturity are grounds for legal separation under Article 55, they do not satisfy the strict requirements of gravity, juridical antecedence, and incurability for psychological incapacity under Article 36.")
        ],
        "analysis_plain": "Congress makes new laws (which look forward/prospective), but the Supreme Court only interprets what the existing law always meant (which looks backward/retroactive). Therefore, the Molina guidelines apply to all pending cases. Physical abuse is a ground for legal separation, not automatic psychological incapacity.",
        "analysis_recite": "\"Sir/Ma'am, the Court explained that judicial construction is retroactive because the Court does not create law, but merely discovers and declares what the law has always meant. Thus, the Molina guidelines governed petitioner's case, and domestic cruelty, without medical proof of grave incapacity, was insufficient under Article 36.\"",
        "conclusion": "The Supreme Court AFFIRMED the Court of Appeals decision, denying the petition for declaration of nullity of marriage without prejudice to the filing of an action for legal separation.",
        "conclusion_plain": "Lorna lost her nullity petition because judicial interpretations apply retroactively to pending cases, but she was permitted to file for legal separation.",
        "conclusion_recite": "\"The Supreme Court denied the petition, holding that judicial interpretations apply retroactively to all pending cases from the date of the statute's effectivity.\"",
        "syllabus": "• Chapter I: Background of the Study of Statutory Construction\n  - H. Legis interpretado legis vim obtinet (interpretation has the force of law)\n  - I. Kinds of construction -> b. Prospective v. retrospective (retroactivity of judicial construction)",
        "latin_maxims": [
            ("Legis interpretado legis vim obtinet (the interpretation of the law has the force of law)", "The interpretation of the law acquires the force of the law itself.", "Supreme Court decisions interpreting statutes become part of the law itself under Article 8 of the Civil Code."),
            ("Lex prospicit, non respicit (the law looks forward, not backward)", "The law looks forward, not backward.", "Statutes enacted by Congress are prospective, contrasting with judicial interpretations which apply retroactively to the date of statutory enactment."),
            ("Stare decisis et non quieta movere (to stand by decisions and not disturb what is settled)", "To stand by decisions and not disturb settled matters.", "Adherence to judicial precedents ensures stability and consistency in statutory interpretation.")
        ],
        "primary_maxim": "Legis interpretado legis vim obtinet\n(the interpretation of the law has the force of law itself)",
        "secondary_maxims": "• Lex prospicit, non respicit (statutes are prospective, but interpretations are retroactive)\n• Stare decisis et non quieta movere (stand by settled decisions)",
        "violation_text": "The petitioner erroneously treated Supreme Court statutory interpretations as prospective legislative acts, failing to recognize that judicial interpretation applies retroactively from the statute's enactment.",
        "execution_text": "Core Authority on Retroactivity of Judicial Construction: Judicial decisions interpreting statutes are incorporated into the law as of the date of the law's passage, applying to all pending and future disputes.",
        "statcon_plain": "The #1 case proving that Supreme Court interpretations of statutes apply retroactively to all pending cases because the Court merely reveals what the statute always meant.",
        "statcon_recite": "\"Sir/Ma'am, under Chapter I (H. Legis interpretado legis vim obtinet): Pesca v. Pesca establishes that judicial interpretations under Article 8 of the Civil Code apply retroactively to the date of the statute's enactment, distinguishing judicial interpretation from prospective legislative enactments.\""
    }
]

print("Phase 1 case definitions loaded successfully.")
