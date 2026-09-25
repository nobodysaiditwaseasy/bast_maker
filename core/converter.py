import subprocess
import os


def docx_to_pdf(docx_path, pdf_path):
    """Convert a .docx to .pdf using Microsoft Word via PowerShell COM."""
    docx_abs = os.path.abspath(docx_path).replace("'", "''")
    pdf_abs = os.path.abspath(pdf_path).replace("'", "''")
    ps_script = (
        f"$word = New-Object -ComObject Word.Application; "
        f"$word.Visible = $false; "
        f"$doc = $word.Documents.Open('{docx_abs}'); "
        f"$doc.SaveAs2('{pdf_abs}', 17); "
        f"$doc.Close(); "
        f"$word.Quit()"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps_script],
        capture_output=True, timeout=60
    )
