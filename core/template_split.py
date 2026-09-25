"""Split a master BAST template into standalone page-1 and page-2 templates.

The master template separates page 1 (BAST body + signature) and page 2
(appendix + table) with a hard page break. We split it into two independent
docx files so each can be rendered as exactly one page.
"""
import os
import shutil

from docx import Document
from docx.oxml.ns import qn


def _find_hard_break_index(elements):
    for i, elem in enumerate(elements):
        for br in elem.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return i
    return -1


def _strip_page_break_runs(elem):
    """Remove hard page break runs from an element, keeping its text."""
    for br in list(elem.iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page":
            run = br.getparent()
            run.remove(br)
            remaining = [c for c in run if c.tag != qn("w:rPr")]
            if not remaining:
                run.getparent().remove(run)


def _is_empty_para(elem):
    return not any((t.text or "").strip() for t in elem.iter(qn("w:t")))


def _content(body):
    return [e for e in body if e.tag != qn("w:sectPr")]


def _strip_trailing_empties(body):
    """Remove empty paragraphs at the very end of the body."""
    for elem in reversed(_content(body)):
        if elem.tag == qn("w:p") and _is_empty_para(elem):
            body.remove(elem)
        else:
            break


def _strip_leading_empties(body):
    """Remove empty paragraphs at the very start of the body."""
    for elem in _content(body):
        if elem.tag == qn("w:p") and _is_empty_para(elem):
            body.remove(elem)
        else:
            break


def split_master_template(tpl_file, tpl_path, tmpdir):
    """Split the master template into two standalone docx files.

    Returns:
        (muka_path, lampiran_path) — file paths inside tmpdir.
    """
    master_path = os.path.join(tmpdir, "_master.docx")
    if tpl_file is not None:
        tpl_file.seek(0)
        with open(master_path, "wb") as f:
            f.write(tpl_file.read())
    else:
        shutil.copyfile(tpl_path, master_path)

    muka_path = os.path.join(tmpdir, "tpl_muka.docx")
    lampiran_path = os.path.join(tmpdir, "tpl_lampiran.docx")

    # --- Page 1: keep everything BEFORE the break, strip trailing empties ---
    shutil.copyfile(master_path, muka_path)
    doc = Document(muka_path)
    body = doc.element.body
    elements = list(body)
    idx = _find_hard_break_index(elements)
    if idx >= 0:
        for elem in elements[idx:]:
            if elem.tag != qn("w:sectPr"):
                body.remove(elem)
    _strip_trailing_empties(body)
    doc.save(muka_path)

    # --- Page 2: keep everything FROM the break, strip break + leading empties ---
    shutil.copyfile(master_path, lampiran_path)
    doc = Document(lampiran_path)
    body = doc.element.body
    elements = list(body)
    idx = _find_hard_break_index(elements)
    if idx >= 0:
        for elem in elements[:idx]:
            if elem.tag != qn("w:sectPr"):
                body.remove(elem)
        # Strip the break from the first (Lampiran title) paragraph
        for elem in body:
            if elem.tag == qn("w:p"):
                _strip_page_break_runs(elem)
                break
    _strip_leading_empties(body)
    doc.save(lampiran_path)

    return muka_path, lampiran_path
