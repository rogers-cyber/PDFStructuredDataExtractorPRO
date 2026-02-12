# PDF Structured Data Extractor PRO — Python Tkinter App (Full Source Code)

**PDF Structured Data Extractor PRO** is an enterprise-grade Python desktop application built with **Tkinter** and **ttkbootstrap** for scanning folders, extracting structured fields from PDFs, storing results in SQLite, and exporting formatted Excel reports.

This repository contains the full source code, allowing you to customize:

- Structured field detection logic  
- OCR confidence filtering  
- AI-based handwriting suppression  
- Multi-threaded processing  
- SQLite database storage  
- Excel export formatting  
- UI styling and branding  

Designed for professional document automation workflows.

------------------------------------------------------------
🌟 SCREENSHOT
------------------------------------------------------------

<img alt="PDF Structured Data Extractor PRO" src="https://github.com/rogers-cyber/PDFStructuredDataExtractorPRO/blob/main/PDF-Structured-Data-Extractor-PRO.jpg" />

------------------------------------------------------------
🌟 FEATURES
------------------------------------------------------------

- 📂 Recursive Folder Scanning — Scan any directory for PDF files  
- 📄 Digital PDF Text Extraction — Extract embedded text using `pdfplumber`  
- 🔎 Smart OCR Fallback — Automatically applies OCR when digital text is unavailable  
- 🧠 AI-Based Handwriting Suppression — Filters low-confidence and irregular OCR results  
- 🏷 Structured Field Extraction — Extracts fields such as:
  - Name  
  - Date  
  - Document ID  
  - (Custom fields easily configurable via regex)  
- 🗄 SQLite Database Storage — Automatically stores extracted records  
- 📊 Live Results Table — Displays extracted structured fields in real time  
- 📤 Formatted Excel Export (.xlsx) — Styled headers, auto column width, frozen rows  
- ⚡ Multi-threaded Processing — Faster concurrent PDF handling  
- ⏸ Pause / Resume / Stop Controls — Real-time scanning control  
- 🖥 Modern UI — Clean ttkbootstrap-powered interface  

------------------------------------------------------------
🚀 INSTALLATION
------------------------------------------------------------

1. Clone or download this repository:

```bash
git clone https://github.com/rogers-cyber/PDFStructuredDataExtractorPRO.git
cd PDFStructuredDataExtractorPRO
```

2. Install required Python packages:

```bash
pip install ttkbootstrap pdfplumber pytesseract pdf2image pillow openpyxl
```

3. Install external dependencies:

Tesseract OCR (Required for OCR fallback)  
Download: https://github.com/tesseract-ocr/tesseract  
Ensure Tesseract is added to your system PATH.

Poppler (Required for pdf2image)

Windows:
- Install Poppler and add to PATH  

macOS:
```bash
brew install poppler
```

Linux:
```bash
sudo apt install poppler-utils
```

4. Run the application:

```bash
python main.py
```

------------------------------------------------------------
💡 USAGE
------------------------------------------------------------

1. Select Folder  
   - Click "Select Folder"  
   - Choose the directory containing PDF documents  

2. Automatic Processing  
   - The app scans recursively for `.pdf` files  
   - If digital text exists → extracts directly  
   - If not → applies OCR  
   - AI-based filtering removes low-confidence and irregular (handwritten) text  

3. Structured Field Detection  

The engine automatically extracts:

- Name (example: `Name: John Smith`)  
- Date (example: `12/31/2025`)  
- Document ID (example: `ID: ABC-12345`)  

You can modify the regex logic inside:

- `extract_name()`  
- `extract_date()`  
- `extract_id()`  

4. Pause / Resume / Stop Scan  

- Pause anytime  
- Resume processing  
- Stop the scan completely  

5. Database Storage  

Each processed document is stored in:

`extracted_data.db`

Table Structure:

| Column       | Description              |
|-------------|--------------------------|
| file_name   | PDF file name            |
| name        | Extracted name field     |
| date        | Extracted date field     |
| document_id | Extracted ID field       |

6. Export to Excel  

- Click "Export Excel"  
- Generates formatted `.xlsx` file  
- Bold headers  
- Auto-sized columns  
- Frozen header row  

------------------------------------------------------------
⚙ CONFIGURATION OPTIONS
------------------------------------------------------------

Option | Description
------ | --------------------------------------------------
OCR Confidence Threshold | Default > 75 (adjust in code)
Minimum Word Length | Filters short noisy text
Bounding Box Filter | Removes irregular OCR shapes
Thread Handling | Adjustable worker threads
Database Engine | SQLite (Upgradeable to PostgreSQL)
Field Patterns | Regex-based, fully customizable
Excel Styling | OpenPyXL formatting logic

------------------------------------------------------------
📦 OUTPUT
------------------------------------------------------------

Application Displays:

- File Name  
- Extracted Name  
- Extracted Date  
- Extracted Document ID  

Database Output:

- `extracted_data.db` (SQLite)

Excel Output:

- Formatted `.xlsx` report  
- Clean structured rows  
- Ready for analysis or reporting  

------------------------------------------------------------
📦 DEPENDENCIES
------------------------------------------------------------

- Python 3.10+  
- Tkinter (Built-in)  
- ttkbootstrap — Modern UI styling  
- pdfplumber — Digital text extraction  
- pytesseract — OCR engine interface  
- pdf2image — PDF to image conversion  
- Pillow — Image processing  
- openpyxl — Excel export formatting  
- sqlite3 — Embedded database  
- concurrent.futures — Multi-threading  
- threading / os / re / time — System operations  

------------------------------------------------------------
🧠 AI HANDWRITING SUPPRESSION LOGIC
------------------------------------------------------------

This application reduces handwritten noise using:

- OCR confidence filtering (>75 default)  
- Minimum word length validation  
- Bounding box size filtering  
- Structured pattern validation  
- Regex-based field validation  

Only validated text is stored in the database.

------------------------------------------------------------
🌐 SaaS EXPANSION READY
------------------------------------------------------------

This desktop version can be upgraded into:

- FastAPI backend service  
- Cloud deployment (AWS / DigitalOcean / Render)  
- PostgreSQL database  
- Stripe subscription billing  
- Web dashboard (React / Next.js)  
- REST API document processing  

Designed with modular architecture for scaling.

------------------------------------------------------------
📝 NOTES
------------------------------------------------------------

- Fully offline desktop application  
- Designed for consistently formatted structured PDFs  
- OCR is only triggered when digital text is unavailable  
- Handwritten annotations are filtered out  
- Performance depends on CPU and PDF size  
- Cross-platform compatible (Windows, macOS, Linux)  

------------------------------------------------------------
👤 ABOUT
------------------------------------------------------------

PDF Structured Data Extractor PRO is maintained by Mate Technologies, delivering intelligent document automation tools for structured PDF processing, business workflows, and scalable SaaS deployment.

Website / Contact:  
https://matetools.gumroad.com  

------------------------------------------------------------
📜 LICENSE
------------------------------------------------------------

Distributed as source code.

You may use it for personal or educational projects.  
Redistribution, resale, or commercial use requires explicit written permission from Mate Technologies.

