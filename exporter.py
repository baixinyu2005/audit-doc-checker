# -*- coding: utf-8 -*-
"""导出比对明细 Excel（比对结果 + 文件内容清单 + 未归类文件）。"""
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

EDGE = Side(style="thin", color="C9D2E0")
BOX = Border(left=EDGE, right=EDGE, top=EDGE, bottom=EDGE)
FILL = {"已提供": ("E2EFDA", "375623"), "缺失": ("FCE4E4", "C00000"),
        "不适用": ("EDEDED", "595959"), "待确认": ("FFF2CC", "9C6500")}
HEAD_FILL = "1F3864"
SUB_FILL = "EEF2FA"

# Excel 不接受控制字符；OCR/二进制文件里可能混进来
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ud800-\udfff\ufffe\uffff]")


def _clean(value):
    if isinstance(value, str):
        value = _ILLEGAL.sub("", value)
        return value[:32000]
    return value


def short_name(path: str, limit: int = 40) -> str:
    name = (path or "").split("/")[-1]
    if len(name) <= limit:
        return name
    return name[: limit // 2 - 1] + "…" + name[-(limit // 2):]


def _fmt_size(n) -> str:
    n = n or 0
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def _header(ws, headers, widths, row=3):
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
    ws.row_dimensions[row].height = 28


def _title(ws, ncols, project):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value="报审资料比对明细")
    c.font = Font(name="微软雅黑", size=15, bold=True, color=HEAD_FILL)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1,
                value=_clean(f"项目：{project.get('name', '')}    核查时间：{project.get('created', '')}"
                             f"    文件 {len(project.get('files') or [])} 个"))
    c.font = Font(name="微软雅黑", size=9, color="7F7F7F")
    ws.row_dimensions[2].height = 20


def export(project: dict, out_path: Path) -> Path:
    wb = Workbook()

    # ---------------- 比对结果 ----------------
    ws = wb.active
    ws.title = "比对结果"
    ws.sheet_view.showGridLines = False
    headers = ["序号", "模板目录", "资料名称", "状态", "识别依据", "命中文件（页码）", "内容摘要", "说明", "写进交接单"]
    _title(ws, len(headers), project)
    _header(ws, headers, [10, 30, 26, 9, 22, 40, 44, 26, 24])

    rows = project.get("rows") or []
    for r, row in enumerate(rows, start=4):
        hits = row.get("hits") or []
        hit_text = "；".join(
            f"{short_name(h.get('name', ''))}"
            + (" P" + ",".join(f"{a}-{b}" if b != a else str(a) for a, b in h.get("pages", [])) if h.get("pages") else "")
            for h in hits[:4])
        if len(hits) > 4:
            hit_text += f" 等 {len(hits)} 份"
        how = "；".join(sorted({h.get("how", "") for h in hits if h.get("how")}))
        snip = next((h.get("snippet") for h in hits if h.get("snippet")), "")
        vals = [_clean(x) for x in [
            row.get("key", ""), row.get("dir", ""), row.get("label", ""),
            row.get("status", ""), how, hit_text, snip,
            row.get("note", ""), "✓" if row.get("write") else ""]]
        for i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = Font(name="微软雅黑", size=9)
            cell.alignment = Alignment(vertical="top", wrap_text=True,
                                       horizontal="center" if i in (1, 4, 9) else "left")
            cell.border = BOX
        cell = ws.cell(row=r, column=4)
        bg, fg = FILL.get(row.get("status"), ("FFFFFF", "000000"))
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(name="微软雅黑", size=9, bold=True, color=fg)
        ws.row_dimensions[r].height = 46

    ws.freeze_panes = "A4"
    if rows:
        ws.auto_filter.ref = f"A3:I{3 + len(rows)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # ---------------- 文件内容清单 ----------------
    ws2 = wb.create_sheet("文件内容清单")
    ws2.sheet_view.showGridLines = False
    headers2 = ["文件", "大小", "页数", "字数", "识别方式", "归类到", "内容摘要", "来源"]
    _title(ws2, len(headers2), project)
    _header(ws2, headers2, [46, 10, 7, 8, 22, 30, 60, 30])

    files = project.get("files") or []
    for r, f in enumerate(files, start=4):
        labels = "、".join(a.get("label", "") for a in (f.get("assigned") or []))
        vals = [_clean(x) for x in [
            f.get("name", ""), _fmt_size(f.get("size")), f.get("pages", 0), f.get("chars", 0),
            f.get("method", ""), labels or "未归类", f.get("snippet", ""),
            f.get("archive", "")]]
        for i, v in enumerate(vals, start=1):
            cell = ws2.cell(row=r, column=i, value=v)
            cell.font = Font(name="微软雅黑", size=9)
            cell.alignment = Alignment(vertical="top", wrap_text=True,
                                       horizontal="center" if i in (2, 3, 4) else "left")
            cell.border = BOX
        if not labels:
            ws2.cell(row=r, column=6).fill = PatternFill("solid", fgColor="FFF2CC")
        ws2.row_dimensions[r].height = 34

    ws2.freeze_panes = "A4"
    if files:
        ws2.auto_filter.ref = f"A3:H{3 + len(files)}"
    ws2.page_setup.orientation = "landscape"
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0

    # ---------------- 未归类文件 ----------------
    ws3 = wb.create_sheet("未归类文件")
    ws3.sheet_view.showGridLines = False
    headers3 = ["文件", "大小", "识别方式", "内容摘要"]
    _title(ws3, len(headers3), project)
    _header(ws3, headers3, [50, 10, 22, 80])
    for r, f in enumerate(project.get("unclassified") or [], start=4):
        vals = [_clean(x) for x in [
            f.get("name", ""), _fmt_size(f.get("size")), f.get("method", ""), f.get("snippet", "")]]
        for i, v in enumerate(vals, start=1):
            cell = ws3.cell(row=r, column=i, value=v)
            cell.font = Font(name="微软雅黑", size=9)
            cell.alignment = Alignment(vertical="top", wrap_text=True,
                                       horizontal="center" if i == 2 else "left")
            cell.border = BOX
        ws3.row_dimensions[r].height = 34
    ws3.freeze_panes = "A4"
    ws3.page_setup.orientation = "landscape"
    ws3.page_setup.fitToWidth = 1
    ws3.page_setup.fitToHeight = 0

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path
