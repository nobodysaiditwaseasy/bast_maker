"""Format 1: N halaman muka per PPL + 1 lampiran rekap di akhir.

Atomic modular approach: each page is rendered as its own standalone docx
from a split sub-template, so no XML surgery is needed on a merged document.
This guarantees exactly 1 page per PPL and prevents blank-page spillover.
"""
import io
import tempfile

from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docxtpl import DocxTemplate

from core.template_split import split_master_template


# ---------------------------------------------------------------------------
# Appendix table helpers
# ---------------------------------------------------------------------------

def set_cell_margin(cell, top=80, start=100, bottom=80, end=100):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def find_appendix_table(doc):
    target = {"no", "nama", "volume", "satuan", "status pekerjaan"}
    for table in doc.tables:
        if not table.rows:
            continue
        header = {cell.text.strip().lower() for cell in table.rows[0].cells}
        if target.issubset(header):
            return table
    return None


def delete_row(table, row):
    table._tbl.remove(row._tr)


def fill_appendix_table(table, rows):
    while len(table.rows) > 2:
        delete_row(table, table.rows[-1])

    widths = [Cm(1.1), Cm(6.2), Cm(2.2), Cm(2.6), Cm(3.4)]

    for nomor, item in enumerate(rows, start=1):
        cells = table.add_row().cells
        values = [
            str(nomor),
            item.get("nama_ppl", ""),
            item.get("vol_kegiatan", ""),
            item.get("satuan_vol", "Dokumen"),
            "Selesai",
        ]
        for idx, (cell, value) in enumerate(zip(cells, values)):
            cell.text = value
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.width = widths[idx]
            set_cell_margin(cell)
            for paragraph in cell.paragraphs:
                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
                )
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(10)


# ---------------------------------------------------------------------------
# Main render
# ---------------------------------------------------------------------------

def render_collective(tpl_file, tpl_path, df, context_fn):
    """Render Format 1 as atomic per-page docx parts.

    Returns:
        list of (filename, bytes) — one part per PPL, then the appendix part.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        muka_path, lampiran_path = split_master_template(tpl_file, tpl_path, tmpdir)

        parts = []

        # --- Phase A: one standalone page per PPL ---
        for idx, (_, row) in enumerate(df.iterrows()):
            doc = DocxTemplate(muka_path)
            doc.render(context_fn(row))

            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)

            name = str(row.get("nama_ppl", idx + 1)).strip().replace(" ", "_")
            parts.append((f"part_{idx + 1:02d}_{name}.docx", buf.getvalue()))

        # --- Phase B: one appendix page with the recap table ---
        lampiran_doc = DocxTemplate(lampiran_path)
        lampiran_doc.render(context_fn(df.iloc[0]))

        table = find_appendix_table(lampiran_doc)
        if table is not None:
            fill_appendix_table(table, df.to_dict("records"))

        buf = io.BytesIO()
        lampiran_doc.save(buf)
        buf.seek(0)
        parts.append(("part_final_lampiran.docx", buf.getvalue()))

        return parts
