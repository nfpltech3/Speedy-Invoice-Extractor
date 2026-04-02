# Speedy Invoice Extractor

Automates the extraction of line-item data from Speedy Multimodes Limited PDFs and converts them into the Logisys Purchase Upload CSV format (41-columns).

## Tech Stack
- Python 3.11
- Tkinter (GUI)
- pdfplumber (PDF Data Extraction)
- pandas / openpyxl (Data formatting/export)

---

## Installation

### Clone
```bash
git clone https://github.com/yourusername/speedy-invoice-extractor.git
cd speedy-invoice-extractor
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
```bash
python -m venv venv
```

2. Activate (REQUIRED)

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. Install dependencies
```bash
pip install -r requirements.txt
```

4. Run application
```bash
python gui_app.py
```

---

### Build Executable (For Desktop Apps)

1. Install PyInstaller (Inside venv):
```bash
pip install pyinstaller
```

2. Build using the included Spec file (Ensure you do not run gui_app.py directly):
```bash
pyinstaller Speedy_Invoice_Extractor.spec
```

3. Locate Executable:
The application will be generated in the `dist/` folder.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv.
- Run and test before pushing.
