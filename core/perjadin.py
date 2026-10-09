"""Laporan Perjadin generator: renders template + adds a documentation page with photo grid."""
import io

from docx import Document
from docx.oxml import OxmlElement
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
    """Add a 'DOKUMENTASI' page with an A4 photo grid."""
    add_page_break(doc)

    # Compact title
    title = doc.add_paragraph()
    title.alignment = 1
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(6)
    run = title.add_run("DOKUMENTASI")
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = "Arial"

    if not photo_bytes_list:
        return

    photos = photo_bytes_list
    rows = (len(photos) + cols - 1) // cols

    # Available height: A4 usable minus title (~1cm) minus a safety margin (~0.5cm)
    TITLE_AND_MARGIN_CM = 1.5
    grid_max_h = USABLE_H - TITLE_AND_MARGIN_CM

    cell_w_cm = USABLE_W / cols
    cell_h_cm = grid_max_h / rows

    pad_cm = 0.1
    img_max_w = cell_w_cm - 2 * pad_cm
    img_max_h = cell_h_cm - 2 * pad_cm

    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"

    # Lock row heights in TWIPS (1 cm = 567 twips), rule = exact
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trHeight = OxmlElement("w:trHeight")
        trHeight.set(qn("w:val"), str(int(cell_h_cm * 567)))
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
            # Old working logic: fit image to cell
            ratio = min(img_max_w / (w / 914400 * 2.54), img_max_h / (h / 914400 * 2.54))
            new_w_emu = Emu(int(w * ratio))
            new_h_emu = Emu(int(h * ratio))

            para = cell.paragraphs[0]
            para.alignment = 1
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
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
