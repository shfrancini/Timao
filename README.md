# Timao N8N Document Conversion Pipeline

Ce dépôt contient des scripts et instructions pour convertir des fichiers Word structurés (.docx) en Excel (.xlsx), les enrichir avec n8n et GPT, puis les reconvertir en document .docx correctement formaté.

---

## ✅ Full Setup from Scratch / Configuration complète depuis zéro

### 1. 🔁 Clone the repository / Cloner le dépôt
```
git clone https://github.com/shfrancini/Timao.git
cd Timao
```

### 2. 🐳 Install Docker / Installer Docker
Download Docker Desktop: https://www.docker.com/products/docker-desktop  
Téléchargez et installez Docker Desktop puis lancez-le.

### 3. 🚀 Run n8n locally / Lancer n8n en local
```
docker run -it --rm \
  -p 5678:5678 \
  -v ~/.n8n:/home/node/.n8n \
  n8nio/n8n
```
Open your browser at: http://localhost:5678  
Ouvrez votre navigateur à l’adresse : http://localhost:5678

### 4. 📥 Import the n8n workflow / Importer le workflow n8n
Go to the n8n UI → "Workflows" → "Import from file"  
Aller dans l’interface n8n → "Workflows" → "Import from file"  
Import: `tests/My workflow.json`

Set credentials (OpenAI, Google Drive) in the Credentials section.  
Configurer vos identifiants (OpenAI, Google Drive) dans la section "Credentials".

### 5. 🐍 Install Python dependencies / Installer les dépendances Python
```
pip install -r requirements.txt
```
(Or use Docker as described below / ou utilisez directement Docker comme ci-dessous)

---

## 📂 Folder Hierarchy & Required Files / Structure des dossiers et fichiers requis

To ensure the pipeline runs smoothly, maintain the following structure in your project root:  
Pour que le pipeline fonctionne correctement, conservez la structure suivante à la racine de votre projet :

```
Timao/
├── scripts/
│   ├── docx_to_xlsx.py                  ← script de conversion DOCX ➝ XLSX
│   └── xlsx_to_docx.py                  ← script de reconversion XLSX ➝ DOCX
├── files/
│   ├── source_input.docx                ← fichier Word source (entrée)
│   ├── output.xlsx                      ← fichier Excel brut généré par le script
│   ├── enriched_output.xlsx             ← fichier enrichi sauvegardé depuis n8n
│   ├── final_output_test.docx           ← document final Word généré
├── n8n_workflow.json                    ← fichier de workflow n8n à importer
├── requirements.txt                     ← dépendances Python
├── README.md                            ← ce fichier
└── Dockerfile (facultatif)              ← pour exécution via conteneur Docker
```

**Important notes / Notes importantes :**  
- The input file must be named `source_input.docx`  
  Le fichier source doit s'appeler `source_input.docx`
- The workflow produces: `output.xlsx` ➝ `File.xlsx` ➝ `output.docx`  
  Le workflow produit successivement : `output.xlsx` ➝ `File.xlsx` ➝ `output.docx`
- The n8n workflow is saved in `n8n_workflow.json` at the root  
  Le workflow n8n est enregistré dans `n8n_workflow.json` à la racine du projet
- Respect this exact structure for the scripts and volume mounting to work  
  Cette structure est nécessaire pour que les scripts et le montage Docker fonctionnent

## 🗐 Usage Process / Processus d'utilisation

### 1. Place your DOCX source file
Placez votre fichier Word source dans le dossier `files/`

- The document must use hierarchical headings: `1`, `1.1`, `1.1.1`
- Le fichier doit contenir des titres hiérarchiques numérotés
- Filename: `source_input.docx`

### 2. Convert DOCX → XLSX (via Python)
```
docker run --rm \
  -v $(pwd)/scripts:/data/scripts \
  -v $(pwd)/files:/data/files \
  python:3.10 \
  /bin/bash -c "pip install -r /data/scripts/requirements.txt && \
                python3 /data/scripts/docx_to_xlsx.py /data/files/source_input.docx /data/files/output.xlsx"
```

### 3. Enrich the Excel file via n8n
- The file `output.xlsx` is used as input in the workflow  
- GPT enriches each row using `context` + `prompt`  
- Columns like `gpt1_response`, `gpt2_response`, etc. are added  
- Save the enriched file as `File.xlsx` manually inside `files/`

### 4. Convert enriched XLSX → DOCX
```
docker run --rm \
  -v $(pwd)/scripts:/data/scripts \
  -v $(pwd)/files:/data/files \
  python:3.10 \
  /bin/bash -c "pip install -r /data/scripts/requirements.txt && \
                python3 /data/scripts/xlsx_to_docx.py /data/files/File.xlsx /data/files/output.docx"
```

---

## 🔧 Python Dependencies / Dépendances Python

Required packages are listed in `requirements.txt`, but minimum includes:
- `openpyxl`
- `python-docx`
- `pandas`

Install locally with:
```
pip install -r requirements.txt
```

---

## 🧠 GPT Prompt & Output Customization

### A. Modify GPT prompts inside n8n
Example prompt (in `prompt` field):  
```
Context: {{context}}

CCTP Content:
{{cctp}}

Please analyze this section and provide your output.
```

### B. Customize DOCX output formatting
Edit `xlsx_to_docx.py` to control formatting:
```python
run.font.bold = True
run.font.size = Pt(11)
```

---

## 📌 File Summary / Récapitulatif des fichiers

| File / Fichier              | Role / Rôle                                  |
|----------------------------|-----------------------------------------------|
| source_input.docx           | Original input / Entrée initiale              |
| output.xlsx                 | Raw extracted content                         |
| enriched_output.xlsx        | Intermediate with GPT responses               |
| File.xlsx                   | Final enriched file for reconversion          |
| output.docx                 | Final DOCX document (auto-formatted)          |

---

## ⚠️ Notes & Requirements

- Only rows with CCTP content are processed in the GPT loop  
- Numbering is rebuilt from structured headings  
- The pipeline will fail if filenames or folder structure differ  
- You must manually save `File.xlsx` from n8n between steps
