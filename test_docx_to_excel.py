from docx import Document
import openpyxl

def get_indent_level(paragraph):
    if paragraph.paragraph_format.left_indent:
        return int(paragraph.paragraph_format.left_indent.pt // 10)  # adjust as needed
    return 0

def parse_hierarchy(docx_path):
    doc = Document(docx_path)
    hierarchy = []
    rows = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        level = get_indent_level(para)

        # Expand hierarchy if needed
        if len(hierarchy) <= level:
            hierarchy += [None] * (level - len(hierarchy) + 1)
        hierarchy[level] = text
        hierarchy = hierarchy[:level + 1]

        # If it's a leaf (likely contains ': CTTP' or just end info), log full path
        if ':' in text or level >= 1:
            rows.append(hierarchy.copy())

    return rows

def write_to_excel(data, excel_path):
    max_depth = max(len(row) for row in data)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([f"H{i+1}" for i in range(max_depth)])  # header row
    for row in data:
        row += [''] * (max_depth - len(row))  # pad empty cols
        ws.append(row)
    wb.save(excel_path)

# Usage
docx_file = "input.docx"    # replace with your file
output_excel = "output.xlsx"

rows = parse_hierarchy(docx_file)
write_to_excel(rows, output_excel)