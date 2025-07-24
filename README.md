# Timao N8N Document Conversion Pipeline

Ce dépôt contient des scripts et instructions pour convertir des fichiers Word structurés (.docx) en Excel (.xlsx), les enrichir avec n8n et GPT, puis les reconvertir en document .docx correctement formaté.

---

## ✅ Step-by-Step Setup / Étapes de configuration

### 1. 🔁 Clone the repository / Cloner le dépôt
```
git clone https://github.com/shfrancini/Timao.git
cd Timao
```

### 2. 📂 Organize your files / Organisez vos fichiers
Place the following in the root folder:
- `source_input.docx` : your input Word document
- `docx_to_xlsx.py` : conversion script
- `xlsx_to_docx.py` : final output script

### 3. ⚙️ Run the scripts / Lancez les scripts
#### Convert Word to Excel
```
python docx_to_xlsx.py
```
This will create `output.xlsx`

#### Run your enrichment pipeline (n8n + GPT)
- Use `output.xlsx` as input to your n8n workflow
- The workflow updates the Excel file with GPT outputs
- Save the new file as `enriched_output.xlsx`

#### Convert enriched Excel to Word
```
python xlsx_to_docx.py
```
This generates a fully formatted `final_output.docx`

---

## 🚀 Requirements / Prérequis
- Python 3.9+
- `pandas`, `openpyxl`, `python-docx`, etc.
- n8n with OpenAI & Google Drive credentials

---

## 🔐 Authentication Setup / Connexions API

### Pinecone
1. Create a Pinecone account and index
2. Copy your API key and environment to `.env` or n8n secret node
3. Use `Pinecone` node in your workflow to vectorize enriched content

### Google Drive API
1. Create a Google Cloud project
2. Enable the Drive API
3. Create OAuth credentials
4. Store the credentials in n8n

### OpenAI
Already shared across the team; no additional config needed unless switching accounts

---

## 📊 Folder Structure / Organisation des fichiers
```
Timao/
├── source_input.docx
├── output.xlsx
├── enriched_output.xlsx
├── final_output.docx
├── docx_to_xlsx.py
├── xlsx_to_docx.py
├── workflow.json
```

---

## ✍️ Notes
- The document must respect a strict heading structure for conversion to work
- GPT enrichment requires column matching (e.g. section, context, prompt)
- The workflow is modular and can be adapted per use-case
