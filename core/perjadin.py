"""Laporan Perjadin generator: renders template + adds a documentation page with photo grid."""
import io

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Emu
from docxtpl import DocxTemplate

PERJADIN_TEMPLATE = "template/template_laporan_perjadin.docx"

# A4 usable area (with 2.54cm margins on 21cm x 29.7cm)
A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
MARGIN_CM = 2.54
USABLE_W = A4_WIDTH_CM - 2 * MARGIN_CM
USABLE_H = A4_HEIGHT_CM - 2 * MARGIN_CM


def add_page_break(doc):
    """Append a page break paragraph."""
    from docx.oxml import OxmlElement
    para = doc.add_paragraph()
    run = para.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._element.append(br)


def build_photo_grid(doc, photo_bytes_list, cols=2):
    """Add a 'DOKUMENTASI' page with an A4 photo grid.

    Args:
        doc: python-docx Document to append to
        photo_bytes_list: list of raw image bytes
        cols: number of columns in the grid (2 or 3)
    """
    # Page break
    add_page_break(doc)

    # Title
    title = doc.add_paragraph()
    title.alignment = 1  # center
    run = title.add_run("DOKUMENTASI")
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = "Arial"

    doc.add_paragraph()  # spacing

    if not photo_bytes_list:
        return

    photos = photo_bytes_list
    rows = (len(photos) + cols - 1) // cols

    # Reserve space for title (~1.5cm) + spacing paragraph (~1cm)
    TITLE_RESERVE_CM = 2.5
    grid_max_h = USABLE_H - TITLE_RESERVE_CM

    # Cell dimensions
    cell_w_cm = USABLE_W / cols
    cell_h_cm = grid_max_h / rows

    # Inner padding inside each cell
    pad_cm = 0.15
    img_max_w = cell_w_cm - 2 * pad_cm
    img_max_h = cell_h_cm - 2 * pad_cm

    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"

    # Set row height exactly (AT_LEAST would grow beyond)
    from docx.oxml import OxmlElement
    for row in table.rows:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement("w:trHeight")
        trHeight.set(qn("w:val"), str(int(Cm(cell_h_cm))))
        trHeight.set(qn("w:hRule"), "exact")
        trPr.append(trHeight)

    for idx, photo_data in enumerate(photos):
        row_idx = idx // cols
        col_idx = idx % cols
        cell = table.cell(row_idx, col_idx)
        cell.width = Cm(cell_w_cm)
        cell.paragraphs[0].clear()

        try:
            from PIL import Image
            img = Image.open(io.BytesIO(photo_data))
            w, h = img.size
            # Scale to fit within cell
            ratio_w = img_max_w / (w / 914400 * 2.54)
            ratio_h = img_max_h / (h / 914400 * 2.54)
            ratio = min(ratio_w, ratio_h)
            new_w_emu = Emu(int(w * ratio))
            new_h_emu = Emu(int(h * ratio))

            para = cell.paragraphs[0]
            para.alignment = 1
            # Reduce spacing in cell paragraph
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            run = para.add_run()
            run.add_picture(io.BytesIO(photo_data), width=new_w_emu, height=new_h_emu)
        except Exception:
            para = cell.paragraphs[0]
            para.alignment = 1
            run = para.add_run(f"[Foto {idx + 1}]")
            run.font.size = Pt(10)

    # Remove extra paragraph Word adds after table
    doc.add_paragraph()


def render_perjadin(context, photo_bytes_list, tpl_path=None, tpl_file=None):
    """Render the perjadin report with a documentation page.

    Args:
        context: dict of template variables
        photo_bytes_list: list of raw image bytes for the documentation page
        tpl_path: path to template on disk (defaults to PERJADIN_TEMPLATE)
        tpl_file: uploaded file object (overrides tpl_path)

    Returns:
        bytes of the rendered .docx
    """
    if tpl_file is not None:
        tpl_file.seek(0)
        doc = DocxTemplate(tpl_file)
    else:
        doc = DocxTemplate(tpl_path or PERJADIN_TEMPLATE)

    doc.render(context)

    # Save rendered doc, then reopen to add documentation page
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    rendered = Document(buf)

    if photo_bytes_list:
        build_photo_grid(rendered, photo_bytes_list, cols=2)

    out = io.BytesIO()
    rendered.save(out)
    out.seek(0)
    return out.getvalue()
