# Timao AI Automation Project
This repo is part of an automation pipeline using n8n to orchestrate AI-powered workflows, focusing on document processing and data extraction.
## Project Overview
- **n8n workflows** automate AI tasks and data pipelines.
- **DOCX to Excel extraction** parses Word files to get numbered outlines and related descriptive text.
- Outputs structured Excel files for further AI processing.
- Additional AI and automation modules are planned.
## Repo Contents
- `test_docx_to_excel.py`: Extracts hierarchical numbered outlines and CCTP text from DOCX files.
- Sample DOCX and Excel output files.
- `.gitignore` to exclude temp/system files.
- `requirements.txt` for Python dependencies.
- `My workflow.json`: n8n workflow (in development).
## Usage
1. Install dependencies by running: `pip install -r requirements.txt`
2. Run extraction script: `python test_docx_to_excel.py`
3. Use generated Excel files in the n8n pipeline.
## Next Steps
- Finish n8n workflow integration.
- Add AI modules for analysis and automation.
- Improve error handling and pipeline stability.
## Contact
Maintained by Scarlett Francini.
