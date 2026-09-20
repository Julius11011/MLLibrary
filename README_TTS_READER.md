# MLC Universal Text-to-Speech (Read Aloud) HTML Reader

A centralized, automated Text-to-Speech and high-readability HTML reader suite for all law course materials across the **MLC** workspace.

---

## 🚀 Quick Start

### Option 1: One-Click Interactive Launcher
Double-click **`Generate-TTS-Readers.bat`** located in this folder:
- **[1] Convert ALL study notes** across all subfolders (`Criminal Law`, `CsL`, `Statutory Construction`, `Criminal Procedure`, `BLJE`, etc.).
- **[2-6] Convert specific subject folders**.
- **[7] Refresh Master Study Hub Index (`MLC_Study_Hub.html`)**.
- **[8] Start Local Web Server** at `http://localhost:8080`.

### Option 2: Drag and Drop
Drag and drop any `.md`, `.docx`, `.txt`, or `.pdf` file directly onto **`Generate-TTS-Readers.bat`** to generate an instant HTML reader beside the original document.

### Option 3: Command Line (PowerShell / Terminal)
```powershell
# Convert a single document:
python generate_tts_reader.py "First Sem 1st Year\Criminal Law\Felonies.md"

# Convert an entire subject folder:
python generate_tts_reader.py --dir "First Sem 1st Year\CsL"

# Convert all study materials across all subfolders:
python generate_tts_reader.py --all

# Refresh the master study hub dashboard:
python generate_tts_reader.py --hub
```

---

## 🎧 Built-in Reader Features

- **Embedded Web Speech Synthesis API**:
  - **Play / Pause**: Click `▶` / `⏸` or press `Space`.
  - **Stop**: Click `⏹` or press `Esc`.
  - **Skip Paragraphs**: Next (`N`) and Previous (`P`).
  - **Reading Speed**: Adjustable from `0.75x` up to `2.0x`.
  - **Voice Selector**: Switch between installed browser voices (Natural US/UK/PH accents).
  - **Active Paragraph Highlight**: The current speaking paragraph glows with a subtle pulse and automatically scrolls to remain centered in view.
  - **Click to Read**: Click directly on any paragraph, heading, or case brief to instantly start reading aloud from that exact point.
  - **Floating Selection Tooltip**: Select any snippet of text with your mouse and click the instant **🔊 Read Selection** button.
- **Reading Comfort & Themes**:
  - **3 Color Modes**: 🌙 Dark (Slate), 📜 Sepia (Warm Parchment), ☀️ Light (Clean Paper). Preferences are saved across sessions.
  - **Font Resizing**: `A+` and `A-` buttons (`+` / `-` keys).
  - **Table of Contents Sidebar**: Dynamic navigation with live search filtering (`/` key).
  - **Top Reading Progress Bar**: Real-time visual indicator of document completion percentage.
- **Browser Extension Compatibility**:
  - Semantic HTML structure (`<article>`, `<section>`, `<div class="read-unit">`) designed for compatibility with Brave Text-to-Read, Read Aloud, Speechify, NaturalReader, and Edge Read Aloud.

---

## ⌨ Keyboard Shortcuts

| Key | Action |
|---|---|
| `Space` | Play / Pause Text-to-Speech |
| `Esc` | Stop Speaking & Reset |
| `N` / `n` | Skip to Next Paragraph / Unit |
| `P` / `p` | Skip to Previous Paragraph / Unit |
| `D` / `d` | Switch to Dark Theme |
| `S` / `s` | Switch to Sepia Theme |
| `L` / `l` | Switch to Light Theme |
| `+` / `=` | Increase Font Size |
| `-` | Decrease Font Size |
| `/` | Focus Table of Contents Search Bar |

---

## 🌐 Local Web Server

To view all files with full extension support without browser `file:///` restrictions, double-click **`start_local_reader_server.bat`** or run:
```powershell
python -m http.server 8080
```
Then open: **`http://localhost:8080/MLC_Study_Hub.html`** in your browser.
