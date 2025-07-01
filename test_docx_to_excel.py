import openpyxl
from docx import Document

def extract_numbered_hierarchy_from_docx(docx_path):
    doc = Document(docx_path)
    rows = []
    hierarchy = []

    for para in doc.paragraphs:
        p = para._p  # low-level lxml element
        numPr = p.find('.//w:numPr', namespaces=p.nsmap)
        if numPr is not None:
            ilvl = numPr.find('.//w:ilvl', namespaces=p.nsmap)
            if ilvl is not None:
                level = int(ilvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
            else:
                level = 0
            text = para.text.strip()
            if len(hierarchy) <= level:
                hierarchy += [''] * (level - len(hierarchy) + 1)
            hierarchy[level] = text
            hierarchy = hierarchy[:level + 1]
            rows.append(hierarchy.copy())
        else:
            # Paragraph without numbering — skip it
            continue

    return rows

def filter_leaf_paths(rows):
    leaf_rows = []
    for i, row in enumerate(rows):
        row_trimmed = [x for x in row if x]

        is_leaf = True
        for j, other in enumerate(rows):
            if i == j:
                continue
            other_trimmed = [x for x in other if x]
            if len(other_trimmed) > len(row_trimmed) and other_trimmed[:len(row_trimmed)] == row_trimmed:
                is_leaf = False
                break
        if is_leaf:
            leaf_rows.append(row)
    return leaf_rows

def write_rows_to_excel(rows, output_file):
    if not rows:
        raise ValueError("⚠️ No valid hierarchy items found.")

    max_depth = max(len(row) for row in rows)
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append([f"H{i+1}" for i in range(max_depth)])

    for row in rows:
        row += [''] * (max_depth - len(row))
        ws.append(row)

    wb.save(output_file)
    print(f"✅ Excel file saved as: {output_file}")

# === CONFIGURATION ===
docx_file = "docx_ex.docx"                # Replace with your actual DOCX filename
excel_file = "structured_output.xlsx"     # Final Excel output

# === RUN PIPELINE ===
rows = extract_numbered_hierarchy_from_docx(docx_file)
leaf_rows = filter_leaf_paths(rows)
write_rows_to_excel(leaf_rows, excel_file)
