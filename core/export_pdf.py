import io
import os
import tempfile
from PyPDF2 import PdfMerger
from core.converter import docx_to_pdf


def export_merged_pdf(docs):
    """Convert each docx to PDF individually, then merge into one PDF.

    Args:
        docs: list of (name, bytes) tuples from render_docs()

    Returns:
        io.BytesIO containing the merged PDF
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        merger = PdfMerger()
        for idx, (name, data) in enumerate(docs):
            docx_path = os.path.join(tmpdir, f"bast_{idx}.docx")
            pdf_path = os.path.join(tmpdir, f"bast_{idx}.pdf")
            with open(docx_path, "wb") as f:
                f.write(data)

            docx_to_pdf(docx_path, pdf_path)

            # LibreOffice outputs PDF with original filename in outdir
            # Verify it exists before merging
            if os.path.exists(pdf_path):
                merger.append(pdf_path)

        buf = io.BytesIO()
        merger.write(buf)
        merger.close()
        buf.seek(0)
    return buf
