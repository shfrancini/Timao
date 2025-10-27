# Timao N8N Document Conversion Pipeline

This repository contains scripts and instructions to convert a structured Word document (.docx) into Excel (.xlsx), enrich it using GPT via n8n, then regenerate a formatted Word document.

Ce dépôt contient des scripts et instructions pour convertir un document Word structuré (.docx) en Excel (.xlsx), l’enrichir avec GPT via n8n, puis le retransformer en document Word mis en forme.

---

## Full Setup from Scratch / Configuration complète depuis zéro

### 1. Clone the repository / Cloner le dépôt

#### Option A – With Git / Avec Git
If you have Git installed, run:
```bash
git clone https://github.com/shfrancini/Timao.git
cd Timao
```

#### Option B – Without Git / Sans Git
If you don’t have Git:
1. Go to [https://github.com/shfrancini/Timao](https://github.com/shfrancini/Timao)
2. Click the green **"Code"** button
3. Choose **"Download ZIP"**
4. Extract the ZIP and open the `Timao` folder

### 2. Install Docker & Create Account / Installer Docker et créer un compte

#### Step 1 – Install Docker Desktop / Installer Docker Desktop
Download and install Docker Desktop:  
Téléchargez et installez Docker Desktop :  
https://www.docker.com/products/docker-desktop

#### Step 2 – Create a Docker account / Créer un compte Docker
Before running any Docker commands, you must **create a free Docker account**:  
Avant d'exécuter des commandes Docker, vous devez **créer un compte Docker gratuit** :  
https://hub.docker.com/signup

Once your account is created and Docker is running, continue with the next steps.  
Une fois votre compte créé et Docker lancé, poursuivez les étapes suivantes.

### 3. 🚀 Launch n8n locally / Lancer n8n en local
Open your terminal or command prompt and run:

```bash
docker run -it --rm \
  -p 5678:5678 \
  -v ~/.n8n:/home/node/.n8n \
  n8nio/n8n
```

Open your browser and go to http://localhost:5678  
Ouvrez votre navigateur et allez à l’adresse http://localhost:5678

Sign in / create an n8n account.
Connectez-vous / créez un compte n8n.

### 4. 📥 Import the n8n workflow / Importer le workflow n8n
- In the n8n interface → click “Workflows” → “Import from file”
- File to import: `n8n_workflow.json`

Dans l’interface n8n → cliquez sur “Workflows” → “Import from file”  
Fichier à importer : `n8n_workflow.json`

Configure your credentials (OpenAI, Google Drive, Pinecone, etc.) in the “Credentials” section.  
Configurez vos identifiants (OpenAI, Google Drive, Pinecone, etc.) dans la section “Credentials”.

---

## 📂 Folder Structure & Required Files / Structure des dossiers et fichiers requis

Maintain this structure in the root of your project:  
Gardez cette structure à la racine du projet :

```
Timao/
├── scripts/
│   ├── docx_to_xlsx.py           ← script: Word ➝ Excel
│   └── xlsx_to_docx.py           ← script: Excel ➝ Word
├── files/
│   ├── input.docx                ← your original Word file
│   ├── int_output.xlsx           ← auto-generated Excel file (initial)
│   ├── gpt_output.xlsx           ← enriched Excel file from n8n
│   ├── output.docx               ← final Word file
├── n8n_workflow.json             ← n8n pipeline to import
├── requirements.txt              ← Python packages
├── README.md                     ← this file
└── Dockerfile (optional)
```

---

## 🗐 Step-by-Step Usage / Étapes d’utilisation

### 1. Prepare your Word document / Préparez votre document Word

- Place it in the `files/` folder  
- Name it `input.docx`  
- Use structured headings like `1`, `1.1`, `1.1.1`  

Déposez le document dans le dossier `files/`, nommé `input.docx`, en utilisant une structure de titres hiérarchiques.

---

### 2. Run the full process via n8n / Lancez tout le processus via n8n

Once the workflow is imported in n8n:

- Click **"Execute Workflow"**
- It will:
  1. Convert `input.docx` → `int_output.xlsx`
  2. Enrich each row with GPT → `gpt_output.xlsx`
  3. Convert that into a Word doc → `output.docx`

Une fois le workflow importé dans n8n :
- Cliquez sur **"Execute Workflow"**
- Il va :
  1. Convertir `input.docx` → `int_output.xlsx`
  2. Enrichir chaque ligne avec GPT → `gpt_output.xlsx`
  3. Générer un document Word final → `output.docx`

---

## Requirements / Prérequis

### Python (if running manually) / Python (si vous exécutez les scripts à la main)
Make sure you have Python 3.9+ installed.  
Assurez-vous d’avoir Python 3.9+ installé.

Install dependencies:
```bash
pip install -r requirements.txt
```

---


## Run the Conversion Script via Docker Only / Exécuter le script via Docker uniquement

All required Python libraries are automatically installed inside the Docker container.  
You do **not** need to install Python or dependencies manually on your machine.

### Run this from the root of the repo:

```bash
docker run --rm \
  -v $(pwd)/scripts:/data/scripts \
  -v $(pwd)/files:/data/files \
  python:3.10 \
  /bin/bash -c "pip install openpyxl python-docx && \
                python3 /data/scripts/xlsx_to_docx.py /data/files/gpt_output.xlsx /data/files/output.docx"
```

This command:
- Mounts your local `scripts/` and `files/` folders inside the Docker container
- Installs required libraries (`openpyxl`, `python-docx`)
- Runs the script that converts `gpt_output.xlsx` into `output.docx`

---

## GPT Prompt Customization / Personnalisation du prompt GPT

Modify the prompt directly in the n8n node where GPT is called.

Example:
```
Context: {{context}}

CCTP Content:
{{cctp}}

Please analyze this section and provide your output.
```

Vous pouvez modifier le prompt directement dans le nœud OpenAI de n8n.

---

## File Summary / Récapitulatif des fichiers

| File / Fichier        | Role / Rôle                                               |
|-----------------------|-----------------------------------------------------------|
| `input.docx`          | Original Word file / Fichier Word d’origine              |
| `int_output.xlsx`     | Extracted content / Contenu extrait                      |
| `gpt_output.xlsx`     | GPT-enriched Excel / Fichier enrichi par GPT             |
| `output.docx`         | Final formatted Word doc / Document Word final formaté   |

---

## Notes

- File names **must match exactly**  
- GPT enrichment works only if `context` and `prompt` columns are present  
- Heading structure in Word is required for section mapping  
- Use consistent folder structure for Docker volume access

Les noms de fichiers doivent correspondre exactement.  
La structure des titres est essentielle pour que la conversion fonctionne correctement.
