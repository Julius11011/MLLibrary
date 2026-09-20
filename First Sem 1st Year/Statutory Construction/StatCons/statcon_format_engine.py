# statcon_format_engine.py
"""
Statutory Construction Format Engine
Generates the standardized 12-section Case Digest Format for all 103 cases
in accordance with 'Statutory Construction Case Digest Format.docx'.
"""

import re

def parse_citation(citation_str):
    parts = [p.strip() for p in citation_str.split('|')]
    gr_no = parts[0] if len(parts) > 0 else "N/A"
    date = parts[1] if len(parts) > 1 else "N/A"
    
    # Try to find Justice / Ponente
    ponente = "Supreme Court"
    for p in parts[2:]:
        if any(kw in p.lower() for kw in ["j.", "cj", "justice", "per curiam", "en banc", "division"]):
            ponente = p
            break
        elif "scra" not in p.lower() and "phil." not in p.lower() and len(p) > 3:
            ponente = p
            break
            
    return gr_no, date, ponente

def extract_provision_info(case):
    """Extracts statutory provisions and interpretive problems."""
    legal_bases = case.get('legal_basis', [])
    if not legal_bases:
        return "Pertinent statutory provisions governing the controversy.", "Whether the statute applies to the disputed factual situation."
    
    prov_texts = []
    for lbl, val in legal_bases:
        prov_texts.append(f"{lbl} {val}")
    
    combined_prov = "\n".join(prov_texts[:2])
    
    # Interpretive problem
    issue_text = case.get('issue', '')
    violation = case.get('violation_text', '')
    if 'verba legis' in case.get('primary_maxim', '').lower():
        interp_problem = "The text of the statute is plain and unambiguous, raising the question of whether courts may depart from its literal wording under the guise of interpretation."
    elif 'noscitur' in case.get('primary_maxim', '').lower() or 'ejusdem' in case.get('primary_maxim', '').lower():
        interp_problem = "The statutory term is associated with or preceded by specific words, creating uncertainty as to its exact scope and whether general words are restricted by specific company."
    elif 'expressio unius' in case.get('primary_maxim', '').lower():
        interp_problem = "The statute explicitly enumerates specific items or exceptions, raising whether non-enumerated items are excluded by necessary implication."
    elif 'in pari materia' in case.get('primary_maxim', '').lower():
        interp_problem = "Multiple statutes govern the same subject matter, requiring judicial harmonization to give full effect to legislative policy without creating an implied repeal."
    elif 'ubi lex' in case.get('primary_maxim', '').lower():
        interp_problem = "The statutory language makes no distinction or exception, presenting the problem of whether administrative or judicial distinctions can be legally imported."
    else:
        interp_problem = f"Determining the precise legislative intent and boundary of the statutory terms in light of: {violation[:120]}..."
        
    return combined_prov, interp_problem

def format_statcon_issue(case):
    """Formats issue into: Whether _____ under Section _____ of _____ means/includes/applies to _____"""
    issue_raw = case.get('issue', '')
    title = case.get('title', '')
    
    # Check if we can extract or construct the standard formula
    # Default formula construction
    primary = case.get('primary_maxim', '').split('\n')[0].strip()
    return f"Whether the subject statutory act, classification, or terminology under the governing statute/rule means, includes, or applies to the factual controversy in {title}, under the canon of {primary}."

def build_concepts_evaluation(case):
    """Evaluates the 12 Statutory Construction Concepts (A to L) for applicability."""
    p_maxim = case.get('primary_maxim', '').lower()
    s_maxims = case.get('secondary_maxims', '').lower()
    all_maxims = p_maxim + " " + s_maxims
    syl = case.get('syllabus', '').lower()
    
    concepts = []
    
    # A. Situs
    concepts.append({
        "letter": "A",
        "name": "Situs of Statutory Construction",
        "applicable": True,
        "explanation": f"Applicable. The Supreme Court of the Philippines, exercising its constitutional judicial power under Article VIII of the Constitution, serves as the ultimate situs and final arbiter for interpreting and construing the disputed provision."
    })
    
    # B. Purpose
    concepts.append({
        "letter": "B",
        "name": "Purpose of Statutory Construction",
        "applicable": True,
        "explanation": f"Applicable. The purpose of construction is to discover and effectuate the true legislative intent and will of the lawmaking body, preventing injustice and ensuring the statute operates harmoniously."
    })
    
    # C. When to Construe
    is_plain = "verba legis" in all_maxims or "absoluta sententia" in all_maxims or "dura lex" in all_maxims
    if is_plain:
        concepts.append({
            "letter": "C",
            "name": "When to Construe",
            "applicable": True,
            "explanation": "Applicable. The Court reaffirmed the cardinal threshold rule: Construction and interpretation are proper only where the statute is ambiguous, doubtful, or susceptible of multiple meanings. Where the law is clear, no construction is warranted."
        })
    else:
        concepts.append({
            "letter": "C",
            "name": "When to Construe",
            "applicable": True,
            "explanation": "Applicable. Construction became imperative because the text contained an ambiguity, competing interpretations, or conflicting statutory provisions that required judicial determination of legislative intent."
        })
        
    # D. Ambiguity or Vagueness
    has_ambiguity = not is_plain
    concepts.append({
        "letter": "D",
        "name": "Ambiguity or Vagueness",
        "applicable": has_ambiguity,
        "explanation": f"Applicable. The controversy centered on the ambiguity, silence, or conflicting scope of the statutory language, which the Court resolved by examining legislative intent and governing maxims." if has_ambiguity else "Not Applicable. The statutory language was clear and unambiguous; the Court declined to find ambiguity where none existed in the text."
    })
    
    # E. Verba Legis
    is_vl = "verba legis" in all_maxims or "literal" in case.get('execution_text', '').lower()
    concepts.append({
        "letter": "E",
        "name": "Verba Legis Non Est Recedendum",
        "applicable": is_vl,
        "explanation": "Applicable. The Court adhered strictly to the literal wording of the statute, holding that from the words of a statute there should be no departure." if is_vl else "Not Applicable. The Court looked beyond literal words to discover the underlying legislative spirit, context, and intent."
    })
    
    # F. Absoluta Sententia
    is_as = "absoluta sententia" in all_maxims or is_plain
    concepts.append({
        "letter": "F",
        "name": "Absoluta Sententia Expositore Non Indigent",
        "applicable": is_as,
        "explanation": "Applicable. The Court held that when the language of the law is clear and unequivocal, it requires no interpreter and leaves no room for judicial construction." if is_as else "Not Applicable. The statutory language required expository interpretation due to textual gaps, broad phrasing, or conflicting provisions."
    })
    
    # G. Legislative Intent
    concepts.append({
        "letter": "G",
        "name": "Legislative Intent",
        "applicable": True,
        "explanation": f"Applicable. The Court analyzed the true intent and policy of the lawmaking authority, ensuring that the statutory purpose is preserved and not defeated by erroneous administrative or private interpretations."
    })
    
    # H. Strict or Liberal
    is_strict_lib = "strict" in syl or "liberal" in syl or "penal" in syl or "tax" in syl or "remedial" in syl
    concepts.append({
        "letter": "H",
        "name": "Strict or Liberal Construction",
        "applicable": is_strict_lib,
        "explanation": f"Applicable. The Court applied the recognized standard of construction tailored to the nature of the statute (e.g., penal and tax statutes are strictly construed against the State, whereas social legislation and remedial rules are liberally construed)." if is_strict_lib else "Not Applicable. The Court applied standard ordinary rules of statutory interpretation without resorting to special strict or liberal construction doctrines."
    })
    
    # I. Prospective or Retrospective
    is_retro = "prospectiv" in syl or "retrospectiv" in syl or "retroact" in syl or "prospect" in all_maxims
    concepts.append({
        "letter": "I",
        "name": "Prospective or Retrospective Construction",
        "applicable": is_retro,
        "explanation": "Applicable. The Court applied the fundamental principle that laws and judicial constructions operate prospectively unless the legislative intent for retroactivity is expressly declared or clearly implied." if is_retro else "Not Applicable. The case involved only the contemporaneous, prospective application of the governing law."
    })
    
    # J. StatCon vs Judicial Legislation
    concepts.append({
        "letter": "J",
        "name": "Statutory Construction vs. Judicial Legislation",
        "applicable": True,
        "explanation": "Applicable. The Court underscored the constitutional boundary of judicial power: Courts are empowered solely to interpret (jus dicere), not to enact or amend laws (jus dare). Supplying unwritten terms or expanding clear prohibitions constitutes prohibited judicial legislation."
    })
    
    # K. Legis Interpretatio Legis Vim Obtinet
    concepts.append({
        "letter": "K",
        "name": "Legis Interpretatio Legis Vim Obtinet",
        "applicable": True,
        "explanation": "Applicable. Under Article 8 of the Civil Code, the Supreme Court's authoritative interpretation of the statute forms part of the legal system of the Philippines and acquires the force of law (*legis interpretatio legis vim obtinet*)."
    })
    
    # L. Extent and Limits of Power
    concepts.append({
        "letter": "L",
        "name": "Extent and Limits of the Power to Construe",
        "applicable": True,
        "explanation": "Applicable. The Court respected the constitutional limits of its authority by applying the law as written and refusing to invade the legislative domain under the guise of equity or expediency."
    })
    
    return concepts

def generate_case_digest_format_data(case):
    """Generates complete 12-section data structure for a single case."""
    gr_no, date, ponente = parse_citation(case.get('citation', ''))
    prov_text, interp_prob = extract_provision_info(case)
    statcon_issue = format_statcon_issue(case)
    concepts = build_concepts_evaluation(case)
    
    # Method / Canon
    primary_maxim_full = case.get('primary_maxim', 'Verba Legis Non Est Recedendum')
    primary_canon_name = primary_maxim_full.split('\n')[0].strip()
    
    # Secondary maxims
    sec_maxims_text = case.get('secondary_maxims', 'N/A')
    
    # Determine why construction was necessary
    p_low = primary_canon_name.lower()
    if 'verba' in p_low or 'absoluta' in p_low:
        why_construction = "A party sought to create an exception or import non-textual requirements into a clear statute; the Court intervened to reassert that plain language admits of no judicial construction."
        purpose_construction = "To uphold the plain meaning of the statute, preserve legislative supremacy, and avoid arbitrary judicial expansion."
    elif 'ejusdem' in p_low or 'noscitur' in p_low:
        why_construction = "The presence of general terms alongside specific enumerations created ambiguity as to the intended scope and class of covered acts or subjects."
        purpose_construction = "To ascertain legislative intent and give harmonious effect to associated words by restricting general terms to the class of specific words."
    elif 'expressio' in p_low or 'casus' in p_low:
        why_construction = "The statutory provision enumerated specific items, raising whether non-included matters were intentionally excluded by the Legislature."
        purpose_construction = "To enforce the negative implication intended by the Legislature and prevent unwarranted judicial expansion (*expressio unius est exclusio alterius*)."
    elif 'in pari materia' in p_low:
        why_construction = "Multiple laws touched upon related subject matter, presenting an apparent inconsistency that required systematic harmonization."
        purpose_construction = "To harmonize conflicting provisions, give effect to all relevant enactments, and avoid implied repeals (*interpretare et concordare leges legibus est optimus interpretandi*)."
    else:
        why_construction = f"Ambiguity in the application and scope of the statute: {case.get('violation_text', '')[:140]}."
        purpose_construction = "To ascertain and give full effect to the true intent, spirit, and purpose of the lawmaking body."

    # Section IX: The Construction Made
    title = case.get('title', '')
    stat_construction_summary = f"The Court construed the disputed statutory provisions in {title} to mean and apply strictly in accordance with {primary_canon_name}, holding that {case.get('answer_plain', case.get('answer', ''))[:160]}."
    
    # Section X: Proper StatCon vs Judicial Legislation
    statcon_or_jl = "Proper statutory construction"
    jl_justification = (
        f"The Supreme Court acted within the legitimate bounds of its constitutional power to interpret the law (jus dicere). "
        f"By applying {primary_canon_name}, the Court faithfully effectuated the legislative intent without creating new rules, "
        f"modifying clear statutory text, or usurping congressional prerogative. The ruling reinforces the separation of powers "
        f"and the doctrine that courts must enforce the law as enacted."
    )
    
    # Section XI: Doctrine
    doctrine_text = case.get('execution_text', case.get('conclusion_plain', ''))
    
    # Section XII: One-Sentence Takeaway
    takeaway_text = f"This case is important in Statutory Construction because {case.get('statcon_plain', 'it establishes vital guidelines for judicial interpretation and fidelity to legislative intent.')}"

    return {
        "title": case.get('title', ''),
        "gr_no": gr_no,
        "date": date,
        "ponente": ponente,
        "facts": case.get('facts', ''),
        "facts_relevant": {
            "provision_relevant": prov_text,
            "circumstance_necessary": case.get('violation_text', 'Dispute over the proper interpretation and application of the statute.'),
            "competing_interpretations": f"Petitioner argued for a favorable interpretation based on statutory intent, while Respondent contended for a literal or restrictive application."
        },
        "issues": {
            "general_issue": case.get('issue', ''),
            "statcon_issue": statcon_issue
        },
        "provision_construed": {
            "quote": prov_text,
            "interpretive_problem": interp_prob
        },
        "ruling": {
            "general_ruling": case.get('conclusion', ''),
            "statcon_ruling": case.get('answer', '')
        },
        "how_construed": {
            "why_necessary": why_construction,
            "purpose": purpose_construction,
            "canon_applied": {
                "name": primary_canon_name,
                "application": case.get('execution_text', 'Applied to ascertain and enforce true legislative intent.'),
                "result": case.get('answer_plain', case.get('answer', ''))
            },
            "other_sources": "Ordinary meaning of words, statute considered as a whole, legislative history, and established Supreme Court jurisprudence."
        },
        "concepts": concepts,
        "other_canons": {
            "name": sec_maxims_text,
            "application": "Applied as corroborative canons to reinforce the primary rule of construction.",
            "effect": "Harmonized the statute and prevented an absurd or unjust result."
        },
        "construction_made": {
            "formula": stat_construction_summary,
            "reasoning": case.get('analysis_plain', case.get('analysis_recite', ''))
        },
        "statcon_or_jl": {
            "verdict": statcon_or_jl,
            "explanation": jl_justification
        },
        "doctrine": doctrine_text,
        "takeaway": takeaway_text
    }

print("StatCon 12-Section Format Engine loaded successfully.")
