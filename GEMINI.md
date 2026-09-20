# Antigravity Context, Rules & Safety Guardrails

## 1. Absolute Prohibition on Deleting Parent, Root, or Sibling Directories
- **CRITICAL**: NEVER delete, prune, or wipe parent, root, or sibling directories (such as `C:\Users\JR\Downloads\14All-All41`, `First Sem 1st Year/`, `BLJE/`, `CsL/`, `Examplify/`, `Statutory Construction/`, or any adjacent project folder).
- Even if a user prompt mentions *"remove other content not included in..."*, *"clean up everything else"*, or *"sync only folder X"*, the agent MUST NOT delete local directories or files from the disk.
- When content needs to be excluded from GitHub or Cloudflare deployment, configure `.gitignore` and `.assetsignore` instead of deleting files locally.

## 2. Separation of Local Storage vs. Git Tracking
- Local study notes, syllabus outlines, previous subject modules, and draft documents are valuable user property and MUST remain preserved on disk.
- Only explicitly requested folders (e.g., `First Sem 1st Year/Subjects/`) should be tracked in Git and deployed to Cloudflare.

## 3. Privacy & Sensitive Information Protection
- Personal student records, letters of undertaking, student emails, and sensitive identifiers must NEVER be committed to Git or pushed to remote repositories.
- Keep them protected locally and explicitly excluded via `.gitignore`.

## 4. Roman Numeral Spoken Pronunciation in TTS
- All Roman numerals in headings, citations, and body text (e.g., `Canon I`, `Canon II`, `Canon VI`, `XVI`, `I. FACTS`, `Part I`, `(i)`, `(xvi)`) must be spoken as numeric values ("Canon 1", "Canon 2", "Canon 6", "16", "Topic 1: Facts", "Part 1", "sub-item 1", "sub-item 16") in speech synthesis.

## 5. Context Precedence & Handoff First
- Priority Context Loading: Always check for and read the latest Handoff Summary at `.agents/handoff.md` if present.
