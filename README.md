# Timao N8N Document Conversion Pipeline

This repository contains everything needed to convert a Word document (`.docx`) into an enriched and formatted version using n8n and GPT. The entire process — from DOCX to XLSX, GPT enrichment, and back to DOCX — is managed via one automated workflow in n8n.

Ce dépôt contient tout le nécessaire pour transformer un fichier Word (`.docx`) en une version enrichie et formatée à l’aide de n8n et GPT. Tout le processus — de DOCX à XLSX, enrichissement GPT, puis retour à DOCX — est automatisé dans un seul workflow n8n.

---

## ✅ Quick Start Guide / Guide de démarrage rapide

### 1. 🔁 Download the repository / Télécharger le dépôt

You can:
- Download the ZIP here: https://github.com/shfrancini/Timao/archive/refs/heads/main.zip  
- Or, if you have Git:
  ```
  git clone https://github.com/shfrancini/Timao.git
  cd Timao
  ```

Téléchargez le dossier ou utilisez Git pour le cloner.

---

### 2. 🐳 Install Docker & Run n8n / Installer Docker & lancer n8n

- Download Docker Desktop: https://www.docker.com/products/docker-desktop
- Launch Docker
- Then run this command in your terminal (Command Prompt, Terminal, etc.):

```
docker run -it --rm \
  -p 5678:5678 \
  -v ~/.n8n:/home/node/.n8n \
  -v $(pwd):/data \
  n8nio/n8n
```

Open your browser at http://localhost:5678  
Ouvrez votre navigateur à l’adresse : http://localhost:5678

---

### 3. 📥 Import the n8n workflow

In the n8n interface:
- Click "Workflows" → "Import from file"
- Select `n8n_workflow.json` from the repo folder

Dans l’interface n8n :
- Cliquez sur “Workflows” → “Import from file”
- Sélectionnez `n8n_workflow.json` depuis le dossier téléchargé

---

### 4. 📂 Prepare your input document

- Place your source Word document in the `/files` folder
- It **must be named**: `source_input.docx`
- Use clear hierarchical headings: `1`, `1.1`, `1.1.1`, etc.

Placez votre document Word source dans le dossier `/files` avec le nom `source_input.docx`.

---

### 5. ⚙️ Execute the workflow

Run the workflow inside n8n.  
It will:
- Convert the Word file into a structured Excel
- Enrich each row using GPT (based on context + CCTP)
- Reconvert it into a clean `.docx` file

Lancez le workflow dans n8n.  
Il va :
- Convertir le fichier Word en Excel
- L’enrichir ligne par ligne avec GPT
- Le reconvertir automatiquement en document Word final

---

## 📁 Folder Structure / Structure du projet

```
Timao/
├── scripts/                     ← Python scripts (used inside the workflow)
│   ├── docx_to_xlsx.py
│   └── xlsx_to_docx.py
├── files/                       ← Your working files go here
│   ├── source_input.docx
│   ├── output.xlsx              (auto-generated)
│   ├── File.xlsx                (enriched file)
│   ├── output.docx              (final formatted output)
├── n8n_workflow.json            ← Main workflow file
├── requirements.txt             ← Python dependencies (already used inside the workflow)
└── README.md
```

---

## 🔐 Credentials Setup

The workflow uses OpenAI and optionally Google Drive. In the n8n UI:
- Go to **Credentials**
- Add your OpenAI API key
- (Optional) Add Google Drive credentials if needed for storage

Dans l’interface n8n, configurez vos identifiants dans la section **Credentials**.

---

## 🧠 GPT Prompt Logic (Customizable)

You can edit the GPT prompts inside the workflow. Example prompt:

```
Context: {{context}}

CCTP:
{{cctp}}

Please rewrite the response to this technical requirement.
```

Vous pouvez modifier les prompts directement dans le champ "prompt" du workflow.

---

## ⚠️ Tips & Requirements

- The filenames must remain the same (`source_input.docx`, `output.xlsx`, etc.)
- The workflow runs in sequence — don’t skip steps
- You must manually export `File.xlsx` inside n8n after enrichment
- Only sections with CCTP content will be enriched

Le pipeline repose sur une structure de noms de fichiers stricte et une exécution complète dans n8n. Veillez à ne pas interrompre le processus.
