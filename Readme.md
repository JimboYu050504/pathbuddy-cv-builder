# CV & Cover Letter Generator (MVP)

A minimal Flask app (1 Python file + several HTML file) that generates a CV and a Cover Letter from a multi‑step form, exports DOCX and tries to export PDF for in‑browser preview.

---

## 1) Setup Instructions

### Prerequisites
- Python 3.9+ (3.10/3.11 recommended)
- (Optional for PDF) One of:
  - **Microsoft Word** (Windows/macOS) for `docx2pdf`, or
  - **LibreOffice** (`soffice` in PATH) for headless DOCX→PDF

### Create and activate a virtual environment
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
```

### Install dependencies
```bash
pip install -r requirements.txt
```
Preferred: docx2pdf (needs Microsoft Word installed).
Fallback: LibreOffice headless. Ensure soffice is on PATH:
macOS (Homebrew): brew install --cask libreoffice
Linux: install from your distro (e.g., sudo apt install libreoffice)
Windows: install LibreOffice and add its program folder to PATH.

### Run the App Locally
```bash
python app.py
```
Visit through web browser: http://127.0.0.1:5000

#Description of Features

Multi‑step form (Personal → Education → Experience → Skills → Cover Letter → Review).

Inline preview (Review step) and New‑tab PDF preview and download.


