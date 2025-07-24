# Timao N8N Document Conversion Pipeline

Ce dépôt contient des scripts et instructions pour convertir des fichiers Word structurés (.docx) en Excel (.xlsx), les enrichir avec n8n et GPT, puis les reconvertir en document .docx correctement formaté.

---

## 🧱 Folder Structure / Structure des dossiers

```
Timao/
├── Dockerfile
├── README.md
├── requirements.txt
├── recherches.py
├── scripts/
│   ├── docx_to_xlsx.py         # Converts DOCX to structured XLSX
│   └── xlsx_to_docx.py         # Converts enriched XLSX back to DOCX
├── files/
│   ├── source_input.docx       # Original input file, with numbered hierarchy
│   ├── output.xlsx             # XLSX generated from DOCX (not enriched)
│   ├── final_output_test.docx  # Final DOCX output (with GPT content)
│   └── ...
├── tests/                      # Sample files for manual or automated testing
└── ...
```

---

## 🧭 Workflow Overview / Vue d'ensemble du processus

1. **Start with** a `.docx` file that uses numbered headings (e.g., 1, 1.1, 1.1.1, etc.).
   - Place this in `files/` and name it: `source_input.docx`

   **Commencez avec** un fichier `.docx` structuré avec des titres numérotés.
   - Placez-le dans `files/` sous le nom `source_input.docx`

2. **Convert to Excel** using the script `scripts/docx_to_xlsx.py`:

   ```bash
   docker run --rm \
     -v ~/Projects/Timao/scripts:/data/scripts \
     -v ~/Projects/Timao/files:/data/files \
     python:3.10 \
     /bin/bash -c "pip install -r /data/scripts/requirements.txt && \
                   python3 /data/scripts/docx_to_xlsx.py /data/files/source_input.docx /data/files/output.xlsx"
   ```

   **Convertissez en Excel** avec le script `scripts/docx_to_xlsx.py`

3. **Enrich `output.xlsx` using n8n**.
   - This process adds GPT outputs (columns like `gpt_summary`, `gpt_keywords`, etc.).
   - Save the enriched file as `output.xlsx` (overwrite) or as a new file if preferred.

   **Enrichissez `output.xlsx` avec n8n**.
   - Ce processus ajoute les réponses GPT.
   - Sauvegardez le fichier enrichi (écrasement ou nouveau nom).

4. **Convert back to DOCX** using the script `scripts/xlsx_to_docx.py`:

   ```bash
   docker run --rm \
     -v ~/Projects/Timao/scripts:/data/scripts \
     -v ~/Projects/Timao/files:/data/files \
     python:3.10 \
     /bin/bash -c "pip install -r /data/scripts/requirements.txt && \
                   python3 /data/scripts/xlsx_to_docx.py /data/files/output.xlsx /data/files/final_output_test.docx"
   ```

   **Convertissez à nouveau en DOCX** avec le script `scripts/xlsx_to_docx.py`

---

## 🔧 Required Python Packages / Dépendances Python

```
openpyxl
python-docx
```

To manually install:
```bash
pip install -r requirements.txt
```

Pour une installation manuelle :
```bash
pip install -r requirements.txt
```

---

## 🧠 GPT Agent Customization / Personnalisation des réponses GPT

To change how GPT responses appear:

### 🔧 Modify the GPT prompt in your assistant (n8n):

```text
Format your response with:
- A bold heading
- Bulleted list
- Formal tone
```

Change to use numbered list, markdown, etc.

### 🎨 Modify output formatting in `xlsx_to_docx.py`:

```python
run.font.bold = True
run.font.size = Pt(11)
```

---

## 📌 Notes

- Only rows with `CCTP` content are processed.
- Numbering is reconstructed from the logical heading levels in the original DOCX.
- GPT responses are inserted below each `CCTP` item.

---

## 💡 Example File Flow / Exemple de chemin de fichiers

| File                     | Purpose                                   | Utilité                              |
|--------------------------|--------------------------------------------|---------------------------------------|
| `source_input.docx`      | Original input with numbered structure     | Fichier source                        |
| `output.xlsx`            | Raw DOCX → XLSX conversion (not enriched) | Conversion Excel (non enrichi)       |
| `output.xlsx` (enriched) | Enriched version with GPT via n8n         | Version enrichie avec GPT            |
| `final_output_test.docx` | Final output after reintegration          | Résultat final                        |