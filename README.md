# Timao DOCX to Excel Hierarchy Extractor

This project extracts numbered hierarchical outlines from DOCX documents and exports them as structured Excel files. It also supports capturing associated blocks of descriptive text (CCTP) following each leaf node in the hierarchy.

---

## Features

- **Extract numbered outlines** (e.g., 1., 1.1, 1.1.1) from DOCX files.
- **Build full hierarchical paths** for numbered paragraphs.
- **Identify and filter leaf nodes** (those without sub-items).
- **Extract associated descriptive text blocks (CCTP)** immediately following leaf nodes.
- **Export hierarchical data and CCTP content to Excel**, with each hierarchy level in separate columns plus a dedicated CCTP column.
- Ignores unnecessary system and temporary files via `.gitignore`.

---

## Project Structure

- `Macaron - Plan detaille.docx` — Sample DOCX input file with numbered outline.
- Python scripts (e.g., `test_docx_to_excel.py`) — Main code for extracting hierarchy and exporting Excel.
- Excel output files (`*.xlsx`) — Various output files showing hierarchical exports, including:
  - `structured_output_with_cctp.xlsx` — Hierarchy plus CCTP text.
  - Other intermediate or filtered output Excel files.
- `.gitignore` — Configured to exclude system files (`.DS_Store`), temporary Office files (`~$*`), Python cache, and other common ignored files.
- `requirements.txt` — Python dependencies (e.g., `python-docx`, `openpyxl`).

---

## Usage

1. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
