# -*- coding: utf-8 -*-
"""读取送审文件的全部内容。

支持：PDF（文字层 + 扫描页 OCR）、Word（docx/doc/wps）、Excel（xlsx/xls）、
OFD、图片 OCR、文本文件，以及认不出的二进制文件（尽力提取关键词）。

OCR 用 Windows 自带的离线识别（Windows.Media.Ocr），不联网、不上传。
"""
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import unicodedata
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

from paths import BASE, RESOURCE

CACHE_DIR = BASE / "data" / "cache"
TEXT_CACHE = CACHE_DIR / "text"
OCR_CACHE = CACHE_DIR / "ocr"
MEDIA_DIR = CACHE_DIR / "media"
TMP_DIR = CACHE_DIR / "tmp"
for _d in (TEXT_CACHE, OCR_CACHE, MEDIA_DIR, TMP_DIR):
    _d.mkdir(parents=True, exist_ok=True)

OCR_PS1 = RESOURCE / "tools" / "win_ocr.ps1"
OCR_VERSION = 3
TEXT_VERSION = 5
OCR_WIDTH = 1700
OCR_JOB_PAGES = 40
OCR_WORKERS = min(4, max(1, int(os.environ.get("AUDIT_OCR_WORKERS", "0")) or ((os.cpu_count() or 4) - 1)))

PDF_EXT = {".pdf"}
DOCX_EXT = {".docx", ".docm"}
DOC_EXT = {".doc", ".wps", ".wpt"}
XLSX_EXT = {".xlsx", ".xlsm"}
XLS_EXT = {".xls", ".et"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp", ".gif"}
TEXT_EXT = {".txt", ".csv", ".md", ".json", ".xml", ".html", ".htm", ".log", ".ini"}
OFD_EXT = {".ofd"}
CAD_EXT = {".dwg", ".dxf"}
PPT_EXT = {".ppt", ".pptx", ".dps"}

MAX_TEXT = 300_000
MAX_PAGES = 800

_CJK = r"\u4e00-\u9fa5\u3400-\u4dbf"


def is_image(ext: str) -> bool:
    return ext.lower() in IMAGE_EXT


def is_pdf(ext: str) -> bool:
    return ext.lower() in PDF_EXT


def normalize_for_match(text: str) -> str:
    """归一化：全半角、大小写、去空白和标点，便于关键词匹配。"""
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text).casefold()
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", text)


def tidy_text(text: str) -> str:
    """整理 OCR 结果：Windows OCR 会在汉字之间加空格，去掉后便于阅读与匹配。"""
    if not text:
        return ""
    text = text.replace("\u00a0", " ").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(rf"(?<=[{_CJK}])\s+(?=[{_CJK}])", "", text)
    text = re.sub(rf"(?<=[{_CJK}])\s+(?=\d)", "", text)
    text = re.sub(rf"(?<=\d)\s+(?=[{_CJK}])", "", text)
    text = re.sub(r"(?<=\d)\s+(?=\d)", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def snippet(text: str, limit: int = 150) -> str:
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text or "")
    s = re.sub(r"\s+", " ", s).strip()
    return s[:limit] + ("…" if len(s) > limit else "")


def file_signature(path: Path) -> str:
    try:
        st = path.stat()
        raw = f"{path.resolve()}|{st.st_size}|{st.st_mtime_ns}"
    except OSError:
        raw = str(path)
    return raw


def text_key(path: Path, use_ocr: bool) -> str:
    raw = f"{file_signature(path)}|v{TEXT_VERSION}|ocr{int(use_ocr)}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def ocr_key(path: Path, page: int) -> str:
    raw = f"{file_signature(path)}|p{page}|w{OCR_WIDTH}|v{OCR_VERSION}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:20]


def _read_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _write_json(path: Path, obj) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


# --------------------------------------------------------------------------
# 各格式读取（不含 OCR）
# --------------------------------------------------------------------------

def _empty_result(method="", warn=""):
    return {"pages": [{"no": 1, "text": "", "src": ""}], "method": method,
            "warn": warn, "pending": [], "images": []}


def _pdf_pages(path: Path):
    from pypdf import PdfReader
    pages, pending = [], []
    try:
        reader = PdfReader(str(path), strict=False)
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                pass
        total = len(reader.pages)
        for i, page in enumerate(reader.pages[:MAX_PAGES]):
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""
            text = tidy_text(text)
            pages.append({"no": i + 1, "text": text, "src": "PDF文字层" if text else ""})
            if len(normalize_for_match(text)) < 12:
                pending.append(i + 1)
        if total > MAX_PAGES:
            warn = f"PDF 共 {total} 页，只读取了前 {MAX_PAGES} 页"
        else:
            warn = ""
        method = "PDF 文字层" if any(p["text"] for p in pages) else "PDF（需 OCR）"
        return {"pages": pages, "method": method, "warn": warn,
                "pending": pending, "images": []}
    except Exception as exc:
        return _empty_result("PDF 读取失败", f"{exc}")


def _media_images(path: Path, prefix: str, limit: int = 20):
    """把 Office 文件里内嵌的图片取出来，交给 OCR。"""
    images = []
    try:
        with zipfile.ZipFile(path) as z:
            media = [n for n in z.namelist() if n.startswith(prefix)]
            for name in media[:limit]:
                ext = Path(name).suffix.lower()
                if ext not in IMAGE_EXT:
                    continue
                data = z.read(name)
                if len(data) < 2000:
                    continue
                out = MEDIA_DIR / f"{hashlib.sha1(data).hexdigest()[:16]}{ext}"
                if not out.exists():
                    out.write_bytes(data)
                images.append({"no": 1, "path": str(out), "src": "文档内图片"})
    except Exception:
        pass
    return images


def _docx_pages(path: Path):
    from docx import Document
    from docx.oxml.ns import qn
    try:
        doc = Document(str(path))
    except Exception as exc:
        return _empty_result("Word 读取失败", str(exc))
    parts = []

    def add(text):
        if text and text.strip():
            parts.append(text.strip())

    for p in doc.paragraphs:
        add(p.text)
    for table in doc.tables:
        for row in table.rows:
            cells = []
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                cells.append(cell.text.replace("\n", " ").strip())
            add("\t".join(cells))
    for section in doc.sections:
        for hf in (section.header, section.footer):
            try:
                for p in hf.paragraphs:
                    add(p.text)
            except Exception:
                pass
    text = "\n".join(parts)

    images = []
    if len(normalize_for_match(text)) < 30:
        images = _media_images(path, "word/media/", limit=20)
    warn = ""
    if images:
        warn = "Word 正文几乎没有文字，已尝试识别文档内图片"
    return {"pages": [{"no": 1, "text": tidy_text(text), "src": "Word正文"}],
            "method": "Word 文档", "warn": warn, "pending": [], "images": images}


_WORD_LOCAL = threading.local()


def _get_word():
    """复用同一个 Word 进程，避免每份 .doc 都重启 Word（每份能省 2-4 秒）。"""
    import pythoncom
    import win32com.client as win32
    app = getattr(_WORD_LOCAL, "app", None)
    if app is None:
        pythoncom.CoInitialize()
        _WORD_LOCAL.inited = True
        last = None
        for prog in ("Word.Application", "KWPS.Application", "WPS.Application"):
            try:
                app = win32.DispatchEx(prog)
                break
            except Exception as exc:
                last = exc
                app = None
        if app is None:
            raise RuntimeError(f"没有可用的 Word/WPS：{last}")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            app.AutomationSecurity = 3
        except Exception:
            pass
        _WORD_LOCAL.app = app
    return app


def close_word():
    app = getattr(_WORD_LOCAL, "app", None)
    _WORD_LOCAL.app = None
    if app is not None:
        for _ in range(3):
            try:
                app.Quit(0)
                break
            except Exception:
                time.sleep(0.2)
    if getattr(_WORD_LOCAL, "inited", False):
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _WORD_LOCAL.inited = False


def _doc_pages(path: Path):
    """.doc / .wps：调用本机 Word 转文本；没有 Word 时退回二进制关键词。"""
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return _binary_pages(path, "Word 组件不可用")
    text, err, doc = "", "", None
    try:
        app = _get_word()
        doc = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False,
                                 Visible=False)
        text = doc.Content.Text or ""
        try:
            for sec in doc.Sections:
                for hf in list(sec.Headers) + list(sec.Footers):
                    text += "\n" + (hf.Range.Text or "")
        except Exception:
            pass
    except Exception as exc:
        err = str(exc)
        close_word()
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                close_word()
    text = text.replace("\x07", "\t").replace("\r", "\n")
    if text.strip():
        return {"pages": [{"no": 1, "text": tidy_text(text), "src": "Word正文"}],
                "method": "Word 97-2003（本机 Word 转换）", "warn": "",
                "pending": [], "images": []}
    res = _binary_pages(path, err or "Word 未能读出文字")
    return res


def _xlsx_pages(path: Path):
    from openpyxl import load_workbook
    try:
        wb = load_workbook(str(path), data_only=True, read_only=True)
    except Exception as exc:
        return _empty_result("Excel 读取失败", str(exc))
    pages = []
    total_chars = 0
    try:
        for idx, ws in enumerate(wb.worksheets, start=1):
            lines = [f"[工作表] {ws.title}"]
            count = 0
            for row in ws.iter_rows(values_only=True):
                vals = []
                for v in row:
                    if v is None or v == "":
                        continue
                    if isinstance(v, datetime):
                        vals.append(v.strftime("%Y-%m-%d"))
                    elif isinstance(v, float) and v.is_integer():
                        vals.append(str(int(v)))
                    else:
                        vals.append(str(v).strip())
                if vals:
                    lines.append("\t".join(vals))
                    count += 1
                if count >= 4000 or total_chars > MAX_TEXT:
                    break
            text = "\n".join(lines)
            total_chars += len(text)
            pages.append({"no": idx, "text": tidy_text(text), "src": f"工作表 {ws.title}"})
            if total_chars > MAX_TEXT:
                break
    finally:
        try:
            wb.close()
        except Exception:
            pass
    warn = ""
    images = []
    if total_chars <= 20:
        warn = "表格里没有读到内容，可能是图片表格，已尝试识别表格里的图片"
        images = _media_images(path, "xl/media/", limit=20)
    return {"pages": pages or [{"no": 1, "text": "", "src": ""}],
            "method": "Excel 表格", "warn": warn, "pending": [], "images": images}


def _xls_pages(path: Path):
    try:
        import xlrd
        book = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        return _empty_result("Excel 读取失败", str(exc))
    pages = []
    total = 0
    try:
        for idx, sheet in enumerate(book.sheets(), start=1):
            lines = [f"[工作表] {sheet.name}"]
            for r in range(min(sheet.nrows, 4000)):
                vals = []
                for c in range(sheet.ncols):
                    cell = sheet.cell(r, c)
                    if cell.value in ("", None):
                        continue
                    v = cell.value
                    if cell.ctype == 3:
                        try:
                            v = datetime(*xlrd.xldate_as_tuple(v, book.datemode)).strftime("%Y-%m-%d")
                        except Exception:
                            v = str(v)
                    elif isinstance(v, float) and v.is_integer():
                        v = str(int(v))
                    else:
                        v = str(v).strip()
                    if v:
                        vals.append(v)
                if vals:
                    lines.append("\t".join(vals))
                if total > MAX_TEXT:
                    break
            text = "\n".join(lines)
            total += len(text)
            pages.append({"no": idx, "text": tidy_text(text), "src": f"工作表 {sheet.name}"})
            if total > MAX_TEXT:
                break
    finally:
        try:
            book.release_resources()
        except Exception:
            pass
    return {"pages": pages or [{"no": 1, "text": "", "src": ""}],
            "method": "Excel 97-2003", "warn": "", "pending": [], "images": []}


def _ofd_pages(path: Path):
    """OFD 本质是 zip，文字在 Content.xml 的 TextCode 里。"""
    texts = []
    pages = []
    try:
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.endswith("Content.xml")]
            names.sort(key=lambda n: [int(x) if x.isdigit() else x
                                      for x in re.split(r"[^0-9]+", n)])
            for i, name in enumerate(names, start=1):
                try:
                    xml = z.read(name).decode("utf-8", errors="replace")
                except Exception:
                    continue
                found = re.findall(r"<(?:[A-Za-z0-9]+:)?TextCode\b[^>]*>(.*?)</(?:[A-Za-z0-9]+:)?TextCode>",
                                   xml, flags=re.S)
                page_text = tidy_text("".join(html.unescape(x) for x in found))
                if page_text:
                    pages.append({"no": i, "text": page_text, "src": "OFD文字"})
                    texts.append(page_text)
    except Exception as exc:
        return _empty_result("OFD 读取失败", str(exc))
    if not pages:
        return _empty_result("OFD（未读到文字）", "OFD 里没有文字层，可能需要转为 PDF/图片后识别")
    return {"pages": pages, "method": "OFD 公文", "warn": "", "pending": [], "images": []}


def _text_pages(path: Path):
    raw = path.read_bytes()[:2_000_000]
    text = ""
    for enc in ("utf-8-sig", "gb18030", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except Exception:
            continue
    if not text:
        text = raw.decode("utf-8", errors="replace")
    return {"pages": [{"no": 1, "text": tidy_text(text), "src": "文本"}],
            "method": "文本文件", "warn": "", "pending": [], "images": []}


def _binary_pages(path: Path, warn: str = ""):
    """二进制文件尽力提取可读字符串，只用于关键词判断。"""
    try:
        data = path.read_bytes()[:8_000_000]
    except Exception as exc:
        return _empty_result("读取失败", str(exc))
    chunks = []
    for enc in ("utf-16-le", "gb18030"):
        try:
            s = data.decode(enc, errors="ignore")
        except Exception:
            continue
        runs = re.findall(rf"[{_CJK}A-Za-z0-9，。、：；（）()《》\-_/\.]{{4,}}", s)
        chunks.extend(runs[:3000])
    text = "\n".join(chunks)
    method = "二进制文本（仅关键词）"
    note = warn or "该格式无法完整读取，只提取了可读字符"
    return {"pages": [{"no": 1, "text": tidy_text(text), "src": method}],
            "method": method, "warn": note, "pending": [], "images": []}


def read_file(path: Path) -> dict:
    """读取一个文件，返回 pages / pending OCR 页 / 内嵌图片。"""
    ext = path.suffix.lower()
    if ext in PDF_EXT:
        return _pdf_pages(path)
    if ext in DOCX_EXT:
        return _docx_pages(path)
    if ext in DOC_EXT:
        return _doc_pages(path)
    if ext in XLSX_EXT:
        return _xlsx_pages(path)
    if ext in XLS_EXT:
        return _xls_pages(path)
    if ext in OFD_EXT:
        return _ofd_pages(path)
    if ext in IMAGE_EXT:
        return {"pages": [{"no": 1, "text": "", "src": ""}], "method": "图片",
                "warn": "", "pending": [1], "images": []}
    if ext in TEXT_EXT:
        return _text_pages(path)
    if ext in PPT_EXT:
        return _empty_result("PPT（未解析）", "PPT 暂未解析文字，只按文件名/目录判断")
    if ext in CAD_EXT:
        return _empty_result("CAD 图纸（未读文字）", "CAD 图纸无法读取文字，只按文件名/目录判断")
    return _empty_result("未识别格式", "该格式无法读取文字，只按文件名/目录判断")


# --------------------------------------------------------------------------
# OCR
# --------------------------------------------------------------------------

def _powershell() -> str:
    for name in ("powershell.exe", "powershell"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",):
        if os.path.exists(p):
            return p
    return ""


def _run_ocr_chunk(ps: str, chunk):
    """一个 PowerShell 进程处理一批 OCR 任务，返回 ({id: {页: 文字}}, 页数)。"""
    token = f"{os.getpid()}-{threading.get_ident()}-{datetime.now():%H%M%S%f}"
    jobs_path = TMP_DIR / f"jobs-{token}.json"
    out_path = TMP_DIR / f"ocr-{token}.json"
    jobs_path.write_text(json.dumps(chunk, ensure_ascii=False), encoding="utf-8")
    mapping = {}
    try:
        proc = subprocess.run(
            [ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(OCR_PS1),
             "-JobsPath", str(jobs_path), "-OutPath", str(out_path), "-Width", str(OCR_WIDTH)],
            capture_output=True, timeout=1800,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if proc.returncode != 0 and not out_path.exists():
            return mapping, sum(len(j["pages"]) for j in chunk)
        data = json.loads(out_path.read_text(encoding="utf-8-sig"))
        n = 0
        for item in data.get("results", []):
            m = mapping.setdefault(item.get("id"), {})
            for pg in item.get("pages", []):
                m[int(pg.get("no") or 1)] = tidy_text(pg.get("text") or "")
                n += 1
        return mapping, n
    except Exception:
        return mapping, 0
    finally:
        for p in (jobs_path, out_path):
            try:
                p.unlink()
            except OSError:
                pass


def _run_ocr_jobs(jobs, on_page=None):
    """jobs: [{id, kind, path, pages}]；返回 {id: {page_no: text}}。

    多个 PowerShell 进程并行 OCR，准确率不变，速度约为单进程的 2-3 倍。
    """
    result = {}
    if not jobs:
        return result
    ps = _powershell()
    if not ps:
        return result

    chunks, cur, cur_pages = [], [], 0
    expanded = []
    for job in jobs:
        pages = job["pages"]
        if len(pages) <= OCR_JOB_PAGES:
            expanded.append(job)
        else:
            for k in range(0, len(pages), OCR_JOB_PAGES):
                part = dict(job)
                part["pages"] = pages[k:k + OCR_JOB_PAGES]
                expanded.append(part)
    for job in expanded:
        n = len(job["pages"])
        if cur and cur_pages + n > 40:
            chunks.append(cur)
            cur, cur_pages = [], 0
        cur.append(job)
        cur_pages += n
    if cur:
        chunks.append(cur)

    workers = min(OCR_WORKERS, len(chunks))
    lock = threading.Lock()

    def merge(mapping, n):
        with lock:
            for jid, pages in mapping.items():
                result.setdefault(jid, {}).update(pages)

    if workers <= 1:
        for chunk in chunks:
            mapping, n = _run_ocr_chunk(ps, chunk)
            merge(mapping, n)
            if on_page:
                for _ in range(n):
                    on_page()
        return result

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_run_ocr_chunk, ps, c) for c in chunks]
        for fut in as_completed(futures):
            try:
                mapping, n = fut.result()
            except Exception:
                continue
            merge(mapping, n)
            if on_page:
                for _ in range(n):
                    on_page()
    return result


def _ocr_group_key(entry) -> str:
    return f"{entry['path']}|{entry['page']}"


def extract_many(files, use_ocr=True, progress=None):
    """读取一批文件的内容。

    files: [{"rel","name","abs","size","ext"}]
    返回 {rel: {"pages":[...], "text":..., "method":..., "warn":..., "chars":...}}
    """
    total = len(files)
    results = {}
    pending_jobs = []      # 待 OCR
    pending_meta = []      # (rel, page_no, job_id)
    loaded = []            # 缓存未命中、需要回写
    ocr_pages_total = 0

    def report(done, phase="读取文件内容", message=""):
        if progress:
            progress(phase, done, total, message)

    # 第一轮：读文字层
    for i, f in enumerate(files, start=1):
        path = Path(f["abs"])
        try:
            key = text_key(path, use_ocr)
            cache_file = TEXT_CACHE / f"{key}.json"
            cached = _read_json(cache_file)
            if cached and isinstance(cached.get("pages"), list):
                cached.setdefault("cache_key", key)
                results[f["rel"]] = cached
                report(i)
                continue
            res = read_file(path)
        except Exception as exc:  # 单个文件被占用/损坏，不影响整批
            results[f["rel"]] = {
                "pages": [{"no": 1, "text": "", "src": ""}],
                "method": "读取失败", "warn": str(exc)[:200], "chars": 0, "text": "",
                "rel": f["rel"], "size": f.get("size", 0), "ext": path.suffix.lower(),
                "cache_key": "",
            }
            report(i)
            continue

        res["rel"] = f["rel"]
        res["size"] = f.get("size", 0)
        res["ext"] = path.suffix.lower()
        res["_cache"] = key
        res["cache_key"] = key
        res["_file"] = str(path)

        if use_ocr:
            job_id = f"f{i}"
            pages = list(res.get("pending") or [])
            # 先查单页 OCR 缓存
            missing = []
            for page_no in pages:
                ck = OCR_CACHE / f"{ocr_key(path, page_no)}.txt"
                if ck.exists():
                    try:
                        txt = ck.read_text(encoding="utf-8")
                        for pg in res["pages"]:
                            if pg["no"] == page_no:
                                pg["text"] = txt
                                pg["src"] = "OCR"
                        continue
                    except Exception:
                        pass
                missing.append(page_no)
            if missing:
                pending_jobs.append({"id": job_id, "kind": "pdf" if is_pdf(path.suffix) else "img",
                                     "path": str(path), "pages": missing})
                for page_no in missing:
                    pending_meta.append((f["rel"], page_no, job_id))
                ocr_pages_total += len(missing)
            if res.get("images"):
                img_jobs = []
                for j, img in enumerate(res["images"][:10], start=1):
                    img_jobs.append({"id": f"f{i}i{j}", "kind": "img",
                                     "path": img["path"], "pages": [1]})
                pending_jobs.extend(img_jobs)
                for job in img_jobs:
                    pending_meta.append((f["rel"], -1, job["id"]))
                ocr_pages_total += len(img_jobs)

        results[f["rel"]] = res
        loaded.append(res)
        report(i)

    # Word 用完就关掉，避免一直占着一个 Word 进程
    try:
        close_word()
    except Exception:
        pass

    # 第二轮：批量 OCR
    if use_ocr and pending_jobs:
        done_pages = 0

        def on_page():
            nonlocal done_pages
            done_pages += 1
            if progress:
                progress("OCR 识别扫描件", done_pages, ocr_pages_total,
                         f"已识别 {done_pages}/{ocr_pages_total} 页")

        ocr_map = _run_ocr_jobs(pending_jobs, on_page=on_page)
        img_texts = {}
        for rel, page_no, job_id in pending_meta:
            text = (ocr_map.get(job_id) or {}).get(1 if page_no == -1 else page_no, "")
            if page_no == -1:
                img_texts.setdefault(rel, []).append(text)
                continue
            res = results.get(rel)
            if not res:
                continue
            for pg in res["pages"]:
                if pg["no"] == page_no:
                    if text:
                        pg["text"] = text
                        pg["src"] = "OCR"
                    break
            if text:
                ck = OCR_CACHE / f"{ocr_key(Path(res['_file']), page_no)}.txt"
                try:
                    ck.write_text(text, encoding="utf-8")
                except OSError:
                    pass
        for rel, texts in img_texts.items():
            res = results.get(rel)
            if res and any(t.strip() for t in texts):
                extra = "\n".join(t for t in texts if t.strip())
                if res["pages"]:
                    base = res["pages"][0].get("text") or ""
                    res["pages"][0]["text"] = tidy_text(base + "\n" + extra)
                    res["pages"][0]["src"] = (res["pages"][0].get("src") or "") + "+图片OCR"

    # 收尾：整理文本、写缓存
    for res in loaded:
        pages = res.get("pages") or []
        for pg in pages:
            pg["text"] = tidy_text(pg.get("text") or "")
        text = "\n".join(pg.get("text") or "" for pg in pages)
        res["text"] = text[:MAX_TEXT]
        res["chars"] = len(normalize_for_match(text))
        if use_ocr:
            n_ocr = sum(1 for pg in pages if pg.get("src") == "OCR")
            if n_ocr:
                res["method"] = (res.get("method") or "") + f" + OCR {n_ocr} 页"
            if res.get("ext") in IMAGE_EXT:
                if res["chars"] > 0:
                    res["method"] = "图片 OCR"
                else:
                    res["method"] = "图片（未识别出文字）"
                    res["warn"] = (res.get("warn") or "") or "图片里没有识别出文字，可能是照片或纯图形"
            elif res.get("pending") and not any(pg.get("text") for pg in pages):
                res["warn"] = (res.get("warn") or "") or "没有读出文字（可能是纯扫描件或空白）"
        for k in ("_file",):
            res.pop(k, None)
        key = res.pop("_cache", "")
        if key:
            try:
                _write_json(TEXT_CACHE / f"{key}.json", res)
            except OSError:
                pass

    return results


def summarize(res: dict, limit: int = 150) -> dict:
    """给网页/Excel 用的精简信息。"""
    pages = res.get("pages") or []
    return {
        "chars": res.get("chars", 0),
        "pages": len(pages),
        "method": res.get("method", ""),
        "warn": res.get("warn", ""),
        "snippet": snippet(res.get("text") or "", limit),
    }


def load_page_texts(cache_key: str):
    data = _read_json(TEXT_CACHE / f"{cache_key}.json")
    return data or {}
