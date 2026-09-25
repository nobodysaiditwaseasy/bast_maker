import io
from docxtpl import DocxTemplate
from core.format_collective import render_collective


def render_docs(tpl_file, tpl_path, df, context_fn, format_mode="individual"):
    """Render DataFrame rows through the template.

    Args:
        tpl_file: uploaded file object (or None for default template)
        tpl_path: path to default template on disk
        df: DataFrame of rows to render
        context_fn: callable(row) -> dict for building template context
        format_mode: "individual" or "collective"

    Returns:
        list of (filename, bytes) tuples
    """
    if format_mode == "collective":
        # Returns atomic per-page parts: one per PPL + one appendix
        return render_collective(tpl_file, tpl_path, df, context_fn)

    # Individual: one doc per row
    docs = []
    for idx, (_, row) in enumerate(df.iterrows()):
        if tpl_file is not None:
            tpl_file.seek(0)
            doc = DocxTemplate(tpl_file)
        else:
            doc = DocxTemplate(tpl_path)

        doc.render(context_fn(row))

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        name = str(row.get("nama_ppl", idx + 1)).strip().replace(" ", "_")
        docs.append((f"BAST_{name}.docx", buf.getvalue()))
    return docs
