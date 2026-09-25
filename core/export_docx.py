import io
import zipfile


def export_zip(docs):
    """Pack (name, bytes) pairs into a ZIP archive.

    Returns:
        io.BytesIO containing the zip
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in docs:
            z.writestr(name, data)
    buf.seek(0)
    return buf


def _sectpr(body):
    """Find the final w:sectPr (must stay the last child of w:body)."""
    from docx.oxml.ns import qn
    return body.find(qn("w:sectPr"))


def _insert_before_sectpr(body, elem):
    from docx.oxml.ns import qn
    sect = _sectpr(body)
    if sect is not None:
        sect.addprevious(elem)
    else:
        body.append(elem)


def _insert_page_break(body):
    """Add a hard page break into the last paragraph (before final sectPr)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    sect = _sectpr(body)
    content = [c for c in body if c is not sect]

    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run.append(br)

    if content and content[-1].tag == qn("w:p"):
        content[-1].append(run)
    else:
        para = OxmlElement("w:p")
        para.append(run)
        _insert_before_sectpr(body, para)


def export_merged_docx(docs):
    """Merge multiple rendered docx byte-arrays into one document.

    Each part is copied in sequence with a hard page break between them.
    The final w:sectPr always remains the last child of w:body.

    Returns:
        io.BytesIO containing the merged .docx
    """
    from copy import deepcopy
    from docx import Document as DocxReader
    from docx.oxml.ns import qn

    merged_doc = None
    body = None

    for _name, data in docs:
        rendered = DocxReader(io.BytesIO(data))
        src_body = rendered.element.body

        if merged_doc is None:
            merged_doc = rendered
            body = merged_doc.element.body
            continue

        # Separate this part from the previous one
        _insert_page_break(body)

        # Copy body elements, excluding the source's own sectPr
        for child in list(src_body):
            if child.tag != qn("w:sectPr"):
                _insert_before_sectpr(body, deepcopy(child))

    out = io.BytesIO()
    merged_doc.save(out)
    out.seek(0)
    return out
