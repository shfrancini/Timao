Timao N8N Document Conversion Pipeline
Pipeline de conversion de documents avec N8N pour Timao

This repository contains scripts and instructions to convert structured Word files (.docx) into Excel (.xlsx), enrich them using n8n and GPT, then reconvert them into a properly formatted .docx document.

Ce dépôt contient des scripts et instructions pour convertir des fichiers Word structurés (.docx) en Excel (.xlsx), les enrichir avec n8n et GPT, puis les reconvertir en document .docx correctement formaté.

============================
✅ Full Setup from Scratch / Configuration complète depuis zéro
============================

1. Clone the repository / Cloner le dépôt

git clone https://github.com/shfrancini/Timao.git
cd Timao

2. Install Docker / Installer Docker

Download Docker Desktop: https://www.docker.com/products/docker-desktop
Téléchargez et installez Docker Desktop puis lancez-le.

3. Run n8n locally / Lancer n8n en local

docker run -it --rm \
  -p 5678:5678 \
  -v ~/.n8n:/home/node/.n8n \
  n8nio/n8n

Open your browser at: http://localhost:5678
Ouvrez votre navigateur à l’adresse : http://localhost:5678

4. Import the n8n workflow / Importer le workflow n8n

Go to the n8n UI → "Workflows" → "Import from file"
Aller dans l’interface n8n → "Workflows" → "Import from file"

Import: tests/My workflow.json  
Importer : tests/My workflow.json

Set credentials (OpenAI, Pinecone, Google Drive) in the Credentials section.  
Configurer vos identifiants (OpenAI, Pinecone, Google Drive) dans la section "Credentials".

5. Install Python dependencies / Installer les dépendances Python

pip install -r requirements.txt

(Or use Docker as described below / ou utilisez directement Docker comme ci-dessous)

=======================
🧱 Folder Structure / Structure des dossiers
=======================

Timao/
├── Dockerfile
├── README.md
├── requirements.txt
├── recherches.py
├── scripts/
│   ├── docx_to_xlsx.py
│   └── xlsx_to_docx.py
├── files/
│   ├── source_input.docx
│   ├── output.xlsx
│   ├── enriched_output.xlsx
│   ├── final_output_test.docx
├── tests/
│   └── My workflow.json

==============================
🗐 Usage Process / Processus d'utilisation
==============================

1. Place your DOCX source file in files/  
   Placez votre fichier Word source dans le dossier files/

- The document must use numbered heading levels: 1, 1.1, 1.1.1  
- Le fichier doit contenir des titres hiérarchiques numérotés : 1, 1.1, 1.1.1

- Expected filename: source_input.docx  
- Nom attendu : source_input.docx

2. Convert DOCX to XLSX / Conversion DOCX → XLSX

docker run --rm \
  -v $(pwd)/scripts:/data/scripts \
  -v $(pwd)/files:/data/files \
  python:3.10 \
  /bin/bash -c "pip install -r /data/scripts/requirements.txt && \
                python3 /data/scripts/docx_to_xlsx.py /data/files/source_input.docx /data/files/output.xlsx"

3. Enrich via n8n / Enrichissement via n8n

- The file output.xlsx is used as input in n8n  
- Le fichier output.xlsx est lu dans le workflow n8n

- GPT responses are added to columns like gpt_summary, gpt_keywords, etc.  
- Les réponses GPT sont ajoutées dans les colonnes gpt_summary, gpt_keywords, etc.

- Save the output as enriched_output.xlsx  
- Sauvegardez sous enriched_output.xlsx

4. Convert enriched XLSX back to DOCX / Conversion XLSX enrichi → DOCX

docker run --rm \
  -v $(pwd)/scripts:/data/scripts \
  -v $(pwd)/files:/data/files \
  python:3.10 \
  /bin/bash -c "pip install -r /data/scripts/requirements.txt && \
                python3 /data/scripts/xlsx_to_docx.py /data/files/enriched_output.xlsx /data/files/final_output_test.docx"

========================
🔧 Python Dependencies / Dépendances Python
========================

openpyxl  
python-docx

pip install -r requirements.txt

===================================
🧠 GPT Prompt & Output Customization
Personnalisation des réponses GPT
===================================

A. Modify GPT prompt in n8n / Modifier le prompt GPT dans n8n

Example / Exemple :
Format your response with:
- A bold heading
- Bulleted list
- Formal tone

B. Customize output formatting in xlsx_to_docx.py  
   Personnalisez le style de sortie dans le script xlsx_to_docx.py

run.font.bold = True  
run.font.size = Pt(11)

====================================
📌 File Summary / Récapitulatif des fichiers
====================================

File / Fichier              | Role / Rôle
--------------------------- | ------------------------------------------
source_input.docx           | Original Word input / Fichier Word source
output.xlsx                 | Raw extracted content / Version brute Excel
enriched_output.xlsx        | With GPT responses / Version enrichie avec GPT
final_output_test.docx      | Final output / Résultat final Word

=====================================
🔹 Notes
=====================================

- Only rows with CCTP content are processed  
- Seuls les blocs contenant du contenu CCTP sont traités

- Numbering is reconstructed from the original structure  
- La numérotation est reconstruite automatiquement

- If the file names and folder structure are preserved, the whole pipeline is replicable  
- Le processus est réplicable si le nom des fichiers et la structure sont respectés

For questions or contributions, contact the Timao team.  
Pour toute contribution ou question, contactez l’équipe Timao.
