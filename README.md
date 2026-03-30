# Skoda License Extractor

A Python-based GUI tool for extracting license details from Bill of Entry (BOE) PDFs. This tool identifies **BE Number**, **BE Date**, and iterates through the **F. LICENCE DETAILS** section to extract **Licence No** and **Debit Duty**.

## Tech Stack
- **Python**: 3.11+
- **GUI**: Tkinter (Nagarkot Brand Theme)
- **PDF Extraction**: PyMuPDF (fitz)
- **Data Export**: Pandas / Openpyxl

---

## Installation

### Clone & Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/Nagarkot-Forwarders/skoda-lic-extractor.git
   cd skoda-lic-extractor
   ```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. **Create virtual environment**
   ```bash
   python -m venv venv
   ```

2. **Activate (REQUIRED)**
   - **Windows:** `venv\Scripts\activate`
   - **Mac/Linux:** `source venv/bin/activate`

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run application**
   ```bash
   python boe_extractor_gui.py
   ```

---

## Build Executable (For Desktop Use)

We use PyInstaller with a specific `.spec` file to bundle assets like the Nagarkot logo.

1. **Install PyInstaller (Inside venv)**
   ```bash
   pip install pyinstaller
   ```

2. **Build using the Spec file**
   ```bash
   pyinstaller boe_extractor_gui.spec
   ```

3. **Locate Executable**
   The standalone app will be generated in the `dist/` folder.

---

## Notes
- **ALWAYS use virtual environment for development.**
- The logo is bundled inside the EXE using `sys._MEIPASS`.
- Ensure all input PDFs are standard India Customs BE formats (Home Consumption).
- **Nagarkot Internal Project** - Developed for Skoda 1702 Documentation.
