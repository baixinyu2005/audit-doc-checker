# -*- coding: utf-8 -*-
"""比对引擎：扫描送审资料 → 读取内容 → 按模板归类 → 生成核对结果。

所有状态以 JSON 存在 data/projects 下，不依赖数据库。
"""
import hashlib
import json
import os
import re
import shutil
import subprocess
import threading
import time
import uuid
import zipfile
from datetime import datetime, timedelta
from pathlib import Path, PurePosixPath

import classify
import textract
from paths import BASE, RESOURCE

DATA = BASE / "data"
PROJECTS = DATA / "projects"
WORK = DATA / "work"
NESTED = WORK / "nested"
OUT = DATA / "out"

for _d in (PROJECTS, WORK, NESTED, OUT):
    _d.mkdir(parents=True, exist_ok=True)

STATUSES = classify.STATUSES
MERGE_NAME = {"中标通知书": "中标通知书", "甲供物资": "甲供物资相关资料"}
ARCHIVE_EXT = {".zip", ".rar", ".7z", ".tar", ".gz", ".tgz", ".bz2", ".xz"}
EXTRACT_VERSION = 2

RULES_PATH = BASE / "rules.json"
if not RULES_PATH.exists():
    default_rules = RESOURCE / "rules.json"
    if default_rules.exists():
        try:
            shutil.copyfile(default_rules, RULES_PATH)
        except OSError:
            pass
_SAVE_LOCK = threading.Lock()


def load_rules() -> dict:
    return json.loads(RULES_PATH.read_text(encoding="utf-8-sig"))


def save_rules(obj: dict) -> None:
    RULES_PATH.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def cache_key(p: Path) -> str:
    st = p.stat()
    raw = f"{p.resolve()}|{st.st_size}|{st.st_mtime_ns}|v{EXTRACT_VERSION}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]


# --------------------------------------------------------------------------
# 压缩包
# --------------------------------------------------------------------------

def _zip_name(info, flags: int = 0) -> str:
    """修复中文 Windows 打的 zip：文件名按 GBK 存，zipfile 默认按 cp437 解成乱码。"""
    if info.flag_bits & 0x800:
        return info.filename
    try:
        raw = info.filename.encode("cp437")
    except UnicodeEncodeError:
        return info.filename
    for enc in ("gb18030", "utf-8"):
        try:
            name = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        if "\ufffd" not in name:
            return name
    return info.filename


def _safe_parts(name: str):
    name = name.replace("\\", "/")
    if name.startswith("/") or re.match(r"^[A-Za-z]:", name):
        return None
    parts = [p for p in PurePosixPath(name).parts if p not in ("", ".")]
    if not parts or any(p == ".." for p in parts):
        return None
    return parts


def _extract_zip(src: Path, dest: Path) -> None:
    with zipfile.ZipFile(src) as z:
        for info in z.infolist():
            name = _zip_name(info)
            parts = _safe_parts(name.rstrip("/"))
            if not parts:
                continue
            if info.is_dir():
                dest.joinpath(*parts).mkdir(parents=True, exist_ok=True)
                continue
            out = dest.joinpath(*parts)
            out.parent.mkdir(parents=True, exist_ok=True)
            with z.open(info) as rf, open(out, "wb") as wf:
                shutil.copyfileobj(rf, wf)


def _extract_tar(src: Path, dest: Path) -> None:
    exe = shutil.which("tar")
    if not exe:
        raise RuntimeError("系统没有 tar，无法解压 rar/7z")
    r = subprocess.run([exe, "-xf", str(src), "-C", str(dest)],
                       capture_output=True, timeout=900)
    if r.returncode != 0:
        err = (r.stderr or r.stdout or b"").decode("gbk", errors="replace")[:300]
        raise RuntimeError(f"解压失败：{err}")


def _extract_archive(src: Path, dest: Path) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    suffix = src.suffix.lower()
    if suffix == ".zip":
        _extract_zip(src, dest)
    else:
        _extract_tar(src, dest)


def prepare_input(src: str, tag: str) -> Path:
    """接收文件夹或压缩包，统一返回资料所在的文件夹。"""
    p = Path(src.strip().strip('"')).expanduser()
    if not p.exists():
        raise FileNotFoundError(f"路径不存在：{p}")
    if p.is_dir():
        return p

    safe = re.sub(r"[^\w\u4e00-\u9fa5-]+", "_", p.stem)[:24]
    dest = WORK / f"{safe}-{cache_key(p)}" / tag
    if dest.exists() and any(dest.iterdir()):
        return single_root(dest)

    try:
        _extract_archive(p, dest)
    except Exception:
        shutil.rmtree(dest.parent, ignore_errors=True)
        raise
    if not any(dest.iterdir()):
        raise RuntimeError("压缩包是空的，或解压后没有文件")
    return single_root(dest)


def clean_work(days: int = 14) -> int:
    deadline = datetime.now() - timedelta(days=days)
    removed = 0
    for root in (WORK, NESTED):
        if not root.exists():
            continue
        for d in root.iterdir():
            if not d.is_dir() or d.name == "nested":
                continue
            try:
                if datetime.fromtimestamp(d.stat().st_mtime) < deadline:
                    shutil.rmtree(d, ignore_errors=True)
                    removed += 1
            except OSError:
                continue
    return removed


def single_root(d: Path) -> Path:
    items = [x for x in d.iterdir() if x.name not in ("__MACOSX", "Thumbs.db")]
    if len(items) == 1 and items[0].is_dir():
        return items[0]
    return d


# --------------------------------------------------------------------------
# 扫描
# --------------------------------------------------------------------------

SKIP_NAME = re.compile(r"^(~\$|\._|\.DS_Store|Thumbs\.db$)", re.I)


def _walk_files(root: Path):
    for f in sorted(root.rglob("*")):
        if not f.is_file():
            continue
        if SKIP_NAME.search(f.name):
            continue
        if "__MACOSX" in f.parts:
            continue
        yield f


def _extract_nested(archive: Path):
    try:
        key = cache_key(archive)
    except OSError:
        return None
    dest = NESTED / key
    if dest.exists() and any(dest.iterdir()):
        return dest
    try:
        _extract_archive(archive, dest)
    except Exception:
        shutil.rmtree(dest, ignore_errors=True)
        return None
    return dest if dest.exists() and any(dest.iterdir()) else None


def _size_of(path: Path) -> int:
    try:
        return path.stat().st_size
    except OSError:
        return 0


def index_files(root: Path, progress=None) -> list:
    """列出全部文件；压缩包里套的压缩包也展开。"""
    files = []
    archives = []
    for f in _walk_files(root):
        rel = f.relative_to(root).as_posix()
        ext = f.suffix.lower()
        if ext in ARCHIVE_EXT:
            archives.append((f, rel))
        else:
            files.append({"rel": rel, "name": f.name, "size": _size_of(f),
                          "ext": ext, "abs": str(f), "archive": ""})

    for f, rel in archives:
        inner_root = _extract_nested(f)
        if not inner_root:
            files.append({"rel": rel, "name": f.name, "size": _size_of(f),
                          "ext": f.suffix.lower(), "abs": str(f), "archive": ""})
            continue
        for x in _walk_files(inner_root):
            inner_rel = x.relative_to(inner_root).as_posix()
            if x.suffix.lower() in ARCHIVE_EXT:
                continue
            files.append({"rel": f"{rel} → {inner_rel}", "name": x.name,
                          "size": _size_of(x), "ext": x.suffix.lower(),
                          "abs": str(x), "archive": rel})

    files.sort(key=lambda x: x["rel"])
    if progress:
        progress("扫描文件", len(files), len(files), f"共 {len(files)} 个文件")
    return files


def scan(root: Path):
    return [{"rel": f["rel"], "name": f["name"], "size": f["size"]}
            for f in index_files(root)]


def scan_dirs(root: Path):
    out = []
    for d in sorted(root.rglob("*")):
        if d.is_dir():
            n = sum(1 for f in d.rglob("*") if f.is_file())
            out.append({"rel": d.relative_to(root).as_posix(), "files": n})
    return out


# --------------------------------------------------------------------------
# 比对
# --------------------------------------------------------------------------

def _file_rows(files, texts, file_map, rows_by_key):
    out = []
    for f in files:
        res = texts.get(f["rel"]) or {}
        info = file_map.get(f["rel"]) or {}
        assigned = []
        for h in info.get("hits", []):
            row = rows_by_key.get(h.get("req"))
            assigned.append({
                "key": h.get("req", ""),
                "label": row["label"] if row else h.get("req", ""),
                "pages": h.get("pages", []),
                "how": h.get("how", ""),
                "snippet": h.get("snippet", ""),
            })
        method = res.get("method", "未读取内容") if texts else "未读取内容"
        if f.get("ext") in textract.IMAGE_EXT:
            method = "图片 OCR" if res.get("chars") else "图片（未识别出文字）"
        out.append({
            "rel": f["rel"], "name": f["name"], "size": f.get("size", 0),
            "ext": f.get("ext", ""), "archive": f.get("archive", ""),
            "abs": f.get("abs", ""),
            "method": method,
            "chars": res.get("chars", 0), "pages": len(res.get("pages") or []),
            "warn": res.get("warn", ""),
            "cache": res.get("cache_key", ""),
            "snippet": textract.snippet(res.get("text") or "", 180),
            "assigned": assigned,
            "unclassified": not assigned,
        })
    return out


def compare(template_dir, unit_dir, rules: dict, options=None, progress=None):
    """扫描 + 读内容 + 归类 + 逐项判断，返回比对结果。"""
    template_dir, unit_dir = Path(template_dir), Path(unit_dir)
    options = options or {}
    read_content = bool(options.get("read_content", True))
    use_ocr = bool(options.get("ocr", True)) and read_content

    def cb(phase, done, total, message=""):
        if progress:
            progress(phase, done, total, message)

    cb("扫描文件", 0, 1, "正在列出送审资料……")
    files = index_files(unit_dir, progress=cb)

    texts = {}
    if read_content and files:
        texts = textract.extract_many(files, use_ocr=use_ocr, progress=cb)

    cb("按模板归类", 0, 1, "正在根据文件内容归类……")
    reqs, template_info = classify.compile_requirements(rules, template_dir)
    rows, file_map, unclassified = classify.build_rows(files, texts, reqs, rules, template_info)
    rows_by_key = {r["key"]: r for r in rows}
    file_rows = _file_rows(files, texts, file_map, rows_by_key)

    # 未归类文件补上已识别摘要
    for u in unclassified:
        res = texts.get(u["rel"]) or {}
        u.setdefault("method", res.get("method", ""))
        u["snippet"] = textract.snippet(res.get("text") or "", 180)
        u.pop("_", None)

    stats = {s: sum(1 for r in rows if r["status"] == s) for s in STATUSES}
    return {
        "rows": rows, "files": file_rows, "unclassified": unclassified,
        "template_info": template_info, "stats": stats,
        "options": {"read_content": read_content, "ocr": use_ocr},
    }


def build_items(rows: list, rules: dict) -> list:
    """把勾选的条目整理成交接单里的文字（已合并同类项）。"""
    order, groups = [], {}
    for r in rows:
        if not r.get("write"):
            continue
        g = r.get("merge_group")
        if g:
            groups.setdefault(g, [])
            if not r.get("container"):
                groups[g].append(r)
            if g not in order:
                order.append(g)
        else:
            order.append(("row", r["key"], r))

    items = []
    for key in order:
        if isinstance(key, tuple):
            items.append(key[2]["report"])
            continue
        members = groups[key]
        if not members:
            continue
        if len(members) == 1 and not members[0].get("merge_role"):
            items.append(members[0]["report"])
            continue
        labels = [m.get("merge_role") or m["report"] for m in members]
        items.append(f"{MERGE_NAME.get(key, key)}（{'、'.join(labels)}）")
    return items


def clean_name(name: str) -> str:
    """去掉「2026.7.8报审--」这类前缀，留下工程名称。"""
    s = (name or "").strip()
    s = re.sub(
        r"^\s*\d{4}\s*[.\-/年]\s*\d{1,2}\s*[.\-/月]\s*\d{1,2}\s*日?\s*"
        r"(报审|送审|审核|报批)?\s*[-—–－_～~]*\s*", "", s)
    s = re.sub(r"^\s*(报审|送审)[-—–－_～~]*\s*", "", s)
    s = re.sub(r"^\d{1,2}(?=[\u4e00-\u9fa5])", "", s)
    return s.strip() or (name or "").strip()


def new_project(name, template_path, unit_path, docx_path, result=None) -> dict:
    pid = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
    result = result or {}
    return {
        "id": pid,
        "name": name,
        "project_name": clean_name(name),
        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "template_path": template_path,
        "unit_path": unit_path,
        "docx_path": docx_path,
        "rows": result.get("rows", []),
        "files": result.get("files", []),
        "unclassified": result.get("unclassified", []),
        "template_info": result.get("template_info", {}),
        "options": result.get("options", {}),
        "stats": result.get("stats", {}),
        "items": [],
        "outputs": {},
        "status": "done",
        "progress": {"phase": "完成", "done": 1, "total": 1, "message": "", "ts": time.time()},
        "error": "",
    }


def save_project(project: dict) -> None:
    project["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    path = PROJECTS / f"{project['id']}.json"
    payload = json.dumps(project, ensure_ascii=False, indent=1)
    last_exc = None
    with _SAVE_LOCK:
        for _ in range(6):
            tmp = path.with_name(path.name + f".{uuid.uuid4().hex[:6]}.tmp")
            try:
                tmp.write_text(payload, encoding="utf-8")
                os.replace(tmp, path)
                return
            except PermissionError as exc:
                last_exc = exc
                time.sleep(0.2)
            finally:
                try:
                    tmp.unlink()
                except OSError:
                    pass
    if last_exc:
        raise last_exc


def load_project(pid: str) -> dict:
    project = json.loads((PROJECTS / f"{pid}.json").read_text(encoding="utf-8"))
    # 兼容旧版记录：补齐新字段，旧记录仍可打开、改状态、生成交接单
    project.setdefault("status", "done")
    project.setdefault("files", [])
    project.setdefault("unclassified", [])
    project.setdefault("template_info", {})
    project.setdefault("options", {})
    project.setdefault("items", [])
    project.setdefault("outputs", {})
    project.setdefault("stats", {})
    for row in project.get("rows", []):
        row.setdefault("hits", [])
        row.setdefault("candidates", [])
        row.setdefault("assigned", [])
        row.setdefault("note", "")
        row.setdefault("write", row.get("status") == "缺失")
        row.setdefault("report", row.get("label", ""))
        for hit in row.get("hits", []):
            hit.setdefault("pages", [])
            hit.setdefault("page_nos", [])
            hit.setdefault("how", "旧版：文件名/目录")
            hit.setdefault("snippet", "")
            hit.setdefault("req", row.get("key", ""))
            hit.setdefault("size", 0)
    return project


def list_projects() -> list:
    out = []
    for f in sorted(PROJECTS.glob("*.json"), reverse=True):
        try:
            p = json.loads(f.read_text(encoding="utf-8"))
            out.append({"id": p["id"], "name": p["name"], "created": p["created"],
                        "status": p.get("status", "done")})
        except Exception:
            continue
    return out[:20]
