# -*- coding: utf-8 -*-
"""把缺件清单写进《资料交接单》Word，保留原表格格式与自动编号。"""
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def find_anchor_cell(doc, anchor: str):
    """在全部表格单元格里找首段以 anchor 开头的单元格。"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.paragraphs and cell.paragraphs[0].text.strip().startswith(anchor):
                    return cell
    return None


def set_after_label(doc, label: str, value: str) -> bool:
    """把「工程名称：xxx」里的值换成 value，保留原有字体字号。"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text.strip()
                    if not text.startswith(label):
                        continue

                    prefix = label + "："
                    if "：" in text:
                        prefix = text.split("：", 1)[0] + "："
                    elif ":" in text:
                        prefix = text.split(":", 1)[0] + "："

                    runs = p.runs
                    if not runs:
                        return False
                    runs[0].text = prefix
                    for r in runs[1:]:
                        r._r.getparent().remove(r._r)

                    new_run = deepcopy(runs[0]._r)
                    t = new_run.find(qn("w:t"))
                    if t is None:
                        t = new_run.makeelement(qn("w:t"), {})
                        new_run.append(t)
                    t.text = value
                    t.set(qn("xml:space"), "preserve")
                    runs[0]._r.addnext(new_run)
                    return True
    return False


def fill_items(docx_path, items, out_path, anchor="工作事项",
               prefix="未见", suffix="；", project_name=None):
    doc = Document(str(docx_path))

    if project_name:
        set_after_label(doc, "工程名称", project_name.strip())

    cell = find_anchor_cell(doc, anchor)
    if cell is None:
        raise RuntimeError(f"没有找到以「{anchor}」开头的单元格，请检查模板")

    paras = cell.paragraphs
    head = paras[0]

    if len(paras) > 1:
        body_template = deepcopy(paras[1]._p)
        had_trailing_blank = not paras[-1].text.strip()
        tail_template = deepcopy(paras[-1]._p) if had_trailing_blank else None
    else:
        body_template = deepcopy(head._p)
        for r in body_template.findall(qn("w:r")):
            body_template.remove(r)
        pPr = body_template.find(qn("w:pPr"))
        if pPr is not None:
            numPr = pPr.find(qn("w:numPr"))
            if numPr is not None:
                pPr.remove(numPr)
            ind = pPr.find(qn("w:ind"))
            if ind is not None:
                ind.set(qn("w:left"), "560")
        had_trailing_blank, tail_template = False, None

    for p in paras[1:]:
        p._p.getparent().remove(p._p)

    anchor_p = head._p
    for text in items:
        new_p = deepcopy(body_template)
        runs = new_p.findall(qn("w:r"))
        for r in runs[1:]:
            new_p.remove(r)
        if runs:
            t = runs[0].find(qn("w:t"))
            if t is None:
                t = runs[0].makeelement(qn("w:t"), {})
                runs[0].append(t)
            t.text = f"{prefix}{text}{suffix}"
            t.set(qn("xml:space"), "preserve")
        else:
            r = new_p.makeelement(qn("w:r"), {})
            t = new_p.makeelement(qn("w:t"), {})
            t.text = f"{prefix}{text}{suffix}"
            t.set(qn("xml:space"), "preserve")
            r.append(t)
            new_p.append(r)
        anchor_p.addnext(new_p)
        anchor_p = new_p

    if had_trailing_blank and tail_template is not None:
        blank = deepcopy(tail_template)
        for r in blank.findall(qn("w:r")):
            blank.remove(r)
        anchor_p.addnext(blank)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def read_back(path, anchor="工作事项"):
    doc = Document(str(path))
    cell = find_anchor_cell(doc, anchor)
    return [p.text for p in cell.paragraphs] if cell else []
