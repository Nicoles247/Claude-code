"""Convert markdown CV/Cover Letter files to .docx format."""
import os
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TEAL = RGBColor(0x00, 0xAC, 0xAC)
DARK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x55, 0x55, 0x55)

def set_font(run, name="Calibri", size=11, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color

def add_horizontal_rule(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '00ACAC')
    pBdr.append(bottom)
    pPr.append(pBdr)

def md_to_docx(md_path, out_path, is_cover_letter=False):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    with open(md_path, encoding='utf-8') as f:
        lines = f.readlines()

    i = 0
    while i < len(lines):
        line = lines[i].rstrip('\n')

        # Skip empty lines at start
        if not line.strip() and i == 0:
            i += 1
            continue

        # H1 — Name (big teal header)
        if line.startswith('# ') and not line.startswith('## '):
            name = line[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(name)
            set_font(run, "Calibri", 26, bold=True, color=TEAL)
            i += 1
            continue

        # H2 — Section headers
        if line.startswith('## '):
            heading = line[3:].strip()
            add_horizontal_rule(doc)
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(3)
            run = p.add_run(heading.upper())
            set_font(run, "Calibri", 11, bold=True, color=TEAL)
            i += 1
            continue

        # H3 — Job titles / subsections
        if line.startswith('### '):
            title = line[4:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(5)
            p.paragraph_format.space_after = Pt(1)
            # Handle bold + rest
            run = p.add_run(title)
            set_font(run, "Calibri", 11, bold=True, color=DARK)
            i += 1
            continue

        # Italic date lines (*...*) after H3
        if line.startswith('*') and line.endswith('*') and not line.startswith('**'):
            content = line.strip('*')
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(content)
            set_font(run, "Calibri", 10, color=GRAY)
            run.font.italic = True
            i += 1
            continue

        # Horizontal rule ---
        if line.strip() in ('---', '***', '___'):
            add_horizontal_rule(doc)
            i += 1
            continue

        # Bullet points
        if line.startswith('- ') or line.startswith('* '):
            content = line[2:].strip()
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.left_indent = Inches(0.2)
            # Handle bold inline (**text**)
            _add_inline_formatted_run(p, content, size=10.5)
            i += 1
            continue

        # Table rows |...|
        if line.startswith('|') and '|' in line[1:]:
            # Collect table rows
            table_lines = []
            while i < len(lines) and lines[i].startswith('|'):
                table_lines.append(lines[i].rstrip('\n'))
                i += 1
            # Filter out separator rows
            data_rows = [r for r in table_lines if not re.match(r'^\|[-| :]+\|$', r.strip())]
            if data_rows:
                cols = len(data_rows[0].split('|')) - 2
                t = doc.add_table(rows=len(data_rows), cols=max(cols, 1))
                t.style = 'Table Grid'
                for ri, row_line in enumerate(data_rows):
                    cells = [c.strip() for c in row_line.split('|')[1:-1]]
                    for ci, cell_text in enumerate(cells):
                        if ci < len(t.rows[ri].cells):
                            cell = t.rows[ri].cells[ci]
                            cell.paragraphs[0].clear()
                            run = cell.paragraphs[0].add_run(cell_text)
                            set_font(run, "Calibri", 10)
            continue

        # Bold line standalone (**...**)
        if line.startswith('**') and line.endswith('**') and line.count('**') == 2:
            content = line.strip('*')
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(1)
            run = p.add_run(content)
            set_font(run, "Calibri", 11, bold=True, color=DARK)
            i += 1
            continue

        # Regular paragraph
        if line.strip():
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(3)
            _add_inline_formatted_run(p, line.strip(), size=10.5)
        else:
            # blank line = small space
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

        i += 1

    doc.save(out_path)
    print(f"  ✓ {os.path.basename(out_path)}")


def _add_inline_formatted_run(paragraph, text, size=10.5):
    """Parse inline **bold** and regular text."""
    parts = re.split(r'(\*\*[^*]+\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            content = part[2:-2]
            run = paragraph.add_run(content)
            set_font(run, "Calibri", size, bold=True, color=DARK)
        else:
            run = paragraph.add_run(part)
            set_font(run, "Calibri", size, color=DARK)


def main():
    apps_dir = "/home/user/Claude-code/applications"
    out_dir = "/home/user/Claude-code/applications/word"
    os.makedirs(out_dir, exist_ok=True)

    md_files = sorted([f for f in os.listdir(apps_dir) if f.endswith('.md')])

    print(f"Convirtiendo {len(md_files)} archivos a Word...\n")
    for fname in md_files:
        md_path = os.path.join(apps_dir, fname)
        docx_name = fname.replace('.md', '.docx')
        out_path = os.path.join(out_dir, docx_name)
        is_cover = 'CoverLetter' in fname
        md_to_docx(md_path, out_path, is_cover_letter=is_cover)

    print(f"\n✅ {len(md_files)} archivos generados en {out_dir}")


if __name__ == "__main__":
    main()
