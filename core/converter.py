import subprocess
import os
import platform


def docx_to_pdf(docx_path, pdf_path):
    """Convert a .docx to .pdf using the best available tool for the platform.

    Windows: Microsoft Word via PowerShell COM
    Linux: LibreOffice headless mode
    """
    docx_abs = os.path.abspath(docx_path)
    pdf_dir = os.path.abspath(os.path.dirname(pdf_path))

    if platform.system() == "Windows":
        _convert_windows(docx_abs, pdf_path)
    else:
        _convert_linux(docx_abs, pdf_dir)


def _convert_windows(docx_path, pdf_path):
    docx_safe = docx_path.replace("'", "''")
    pdf_safe = pdf_path.replace("'", "''")
    ps_script = (
        f"$word = New-Object -ComObject Word.Application; "
        f"$word.Visible = $false; "
        f"$doc = $word.Documents.Open('{docx_safe}'); "
        f"$doc.SaveAs2('{pdf_safe}', 17); "
        f"$doc.Close(); "
        f"$word.Quit()"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps_script],
        capture_output=True, timeout=60
    )


def _convert_linux(docx_path, output_dir):
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf",
         "--outdir", output_dir, docx_path],
        capture_output=True, timeout=120
    )
