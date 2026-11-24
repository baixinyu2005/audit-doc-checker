# -*- coding: utf-8 -*-
"""报审资料核查系统 - 本地网页版。"""
import json
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import uuid
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import (Flask, flash, jsonify, redirect, render_template, request,
                   send_file, url_for)

import docx_fill
import engine
import exporter
import textract
from paths import BASE, RESOURCE, FROZEN

PORT = int(os.environ.get("AUDIT_PORT", "8765"))
app = Flask(__name__, template_folder=str(RESOURCE / "templates"),
            static_folder=str(RESOURCE / "static"))
app.secret_key = "local-audit-tool"
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024
(BASE / "logs").mkdir(parents=True, exist_ok=True)
(BASE / "data").mkdir(parents=True, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html", projects=engine.list_projects(),
                           rules=engine.load_rules())


def _log(msg: str) -> None:
    try:
        with open(BASE / "logs" / "compare.log", "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except OSError:
        pass


def _thread_excepthook(args):
    _log(f"uncaught thread error: {args.exc_type.__name__}: {args.exc_value!r}")


threading.excepthook = _thread_excepthook


def _worker(project, tpl_src, unit_src, docx_path, options):
    """后台执行：解压 → 扫描 → 读内容（OCR）→ 归类。"""
    last = [0.0]
    _log(f"worker start id={project['id']} tpl={tpl_src!r} unit={unit_src!r}")

    def progress(phase, done, total, message=""):
        project["progress"] = {"phase": phase, "done": done, "total": max(total, 1),
                               "message": message, "ts": time.time()}
        now = time.time()
        if now - last[0] > 0.6 or (total and done >= total):
            last[0] = now
            try:
                engine.save_project(project)
            except OSError:
                pass

    try:
        project["progress"] = {"phase": "解压/准备资料", "done": 0, "total": 1, "message": "",
                               "ts": time.time()}
        engine.save_project(project)
        _log("saved initial progress")
        tdir = engine.prepare_input(tpl_src, "template")
        _log(f"template ready {tdir}")
        udir = engine.prepare_input(unit_src, "unit")
        _log(f"unit ready {udir}")
        project["template_path"] = str(tdir)
        project["unit_path"] = str(udir)
        engine.save_project(project)
        _log("saved paths")

        result = engine.compare(tdir, udir, engine.load_rules(), options, progress)
        _log("compare done")
        project.update(result)
        project["status"] = "done"
        project["error"] = ""
        project["progress"] = {"phase": "完成", "done": 1, "total": 1, "message": "",
                               "ts": time.time()}
    except Exception as exc:  # noqa: BLE001
        import traceback
        _log(f"worker error: {exc!r}\n{traceback.format_exc()}")
        project["status"] = "error"
        project["error"] = str(exc)
        project["progress"] = {"phase": "出错", "done": 0, "total": 1,
                               "message": str(exc)[:200], "ts": time.time()}
    finally:
        try:
            engine.save_project(project)
            _log(f"worker end status={project.get('status')}")
        except Exception as exc:  # noqa: BLE001
            _log(f"worker final save failed: {exc!r}")


def _start(project, tpl_src, unit_src, docx_path, options):
    thread = threading.Thread(target=_worker,
                              args=(project, tpl_src, unit_src, docx_path, options),
                              daemon=True)
    thread.start()


@app.route("/compare", methods=["POST"])
def do_compare():
    tpl = request.form.get("template_path", "").strip()
    unit = request.form.get("unit_path", "").strip()
    docx = request.form.get("docx_path", "").strip()
    name = request.form.get("name", "").strip()
    read_content = request.form.get("read_content", "on") == "on"

    try:
        if not tpl or not Path(tpl.strip('"')).exists():
            raise FileNotFoundError("请选择报审资料模板（文件夹或压缩包）")
        if not unit or not Path(unit.strip('"')).exists():
            raise FileNotFoundError("请选择单位送审资料（文件夹或压缩包）")
        if not docx or not Path(docx.strip('"')).exists():
            raise FileNotFoundError("请选择资料交接单 Word")
        unit_path = Path(unit.strip('"'))
        project = engine.new_project(name or unit_path.stem, "", "", str(Path(docx.strip('"'))))
        project["status"] = "processing"
        project["error"] = ""
        project["options"] = {"read_content": read_content, "ocr": read_content}
        project["source_template"] = tpl
        project["source_unit"] = unit
        engine.save_project(project)
        _start(project, tpl, unit, docx, project["options"])
        return redirect(url_for("review", pid=project["id"]))
    except Exception as exc:  # noqa: BLE001
        flash(str(exc))
        return redirect(url_for("index"))


@app.route("/p/<pid>/recheck", methods=["POST"])
def recheck(pid):
    """按当前规则重新比对（沿用已经解压好的资料目录）。"""
    project = engine.load_project(pid)
    if project.get("status") == "processing":
        return jsonify({"ok": False, "error": "正在处理中"})
    tpl = project.get("source_template") or project.get("template_path")
    unit = project.get("source_unit") or project.get("unit_path")
    if not tpl or not unit:
        return jsonify({"ok": False, "error": "找不到原始路径"})
    project["status"] = "processing"
    project["error"] = ""
    project["rows"] = []
    project["files"] = []
    project["unclassified"] = []
    project["progress"] = {"phase": "准备中", "done": 0, "total": 1, "message": "",
                           "ts": time.time()}
    engine.save_project(project)
    _start(project, tpl, unit, project["docx_path"],
           project.get("options") or {"read_content": True, "ocr": True})
    return jsonify({"ok": True})


@app.route("/p/<pid>")
def review(pid):
    try:
        project = engine.load_project(pid)
    except FileNotFoundError:
        flash("项目不存在")
        return redirect(url_for("index"))

    stats = {s: sum(1 for r in project.get("rows", []) if r["status"] == s) for s in engine.STATUSES}
    files = project.get("files") or []
    unit = Path(project["unit_path"]) if project.get("unit_path") else None
    empty_dirs = []
    if project.get("status") == "done" and unit and unit.is_dir():
        try:
            empty_dirs = [d["rel"] for d in engine.scan_dirs(unit) if d["files"] == 0]
        except OSError:
            empty_dirs = []

    tray = [u for u in (project.get("unclassified") or []) if u.get("rel")]
    assigned = {a for row in project.get("rows", []) for a in row.get("assigned", [])}
    tray = [u for u in tray if u["rel"] not in assigned]
    n_read = sum(1 for f in files if f.get("chars"))
    n_img = sum(1 for f in files if f.get("ext") in textract.IMAGE_EXT)
    n_img_read = sum(1 for f in files if f.get("ext") in textract.IMAGE_EXT and f.get("chars"))

    return render_template("review.html", p=project, stats=stats,
                           statuses=engine.STATUSES, n_files=len(files),
                           n_read=n_read, n_img=n_img, n_img_read=n_img_read,
                           empty_dirs=empty_dirs, tray=tray)


@app.route("/p/<pid>/status")
def status(pid):
    try:
        project = engine.load_project(pid)
    except FileNotFoundError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    if project.get("status") == "processing":
        ts = (project.get("progress") or {}).get("ts") or 0
        if ts and time.time() - ts > 900:
            project["status"] = "error"
            project["error"] = "处理超过 15 分钟没有进展，可能被中断（例如 Word 弹窗、文件被占用）。请重新开始比对。"
            project["progress"] = {"phase": "中断", "done": 0, "total": 1,
                                   "message": project["error"], "ts": time.time()}
            engine.save_project(project)
    stats = {s: sum(1 for r in project.get("rows", []) if r["status"] == s) for s in engine.STATUSES}
    return jsonify({"ok": True, "status": project.get("status", "done"),
                    "progress": project.get("progress", {}),
                    "error": project.get("error", ""), "stats": stats})


@app.route("/p/<pid>/assign", methods=["POST"])
def assign(pid):
    """把某份文件指定给某个条目（拖拽归位）。"""
    project = engine.load_project(pid)
    data = request.get_json(force=True) or {}
    key = data.get("key")
    rels = data.get("files") or []

    row = next((r for r in project.get("rows", []) if r["key"] == key), None)
    if row is None:
        return jsonify({"ok": False, "error": "条目不存在"}), 404

    by_rel = {f["rel"]: f for f in project.get("files") or []}
    hits = []
    for rel in rels:
        f = by_rel.get(rel)
        if not f:
            continue
        hits.append({
            "req": key, "rel": rel, "name": f["name"], "size": f.get("size", 0),
            "pages": [[1, 1]], "page_nos": [1], "snippet": f.get("snippet", ""),
            "how": "手动指定", "score": 99, "ambiguous": False,
        })

    if hits:
        row["hits"] = hits
        row["assigned"] = [h["rel"] for h in hits]
        row["status"] = "已提供"
        row["write"] = False
        row["note"] = "已手动指定"
        if row.get("candidates"):
            row["candidates"] = []
    else:
        row["hits"] = []
        row["assigned"] = []
        row["status"] = "待确认"
        row["note"] = "已取消指定"
    engine.save_project(project)
    return jsonify({"ok": True, "status": row["status"],
                    "hits": [h["name"] for h in hits]})


@app.route("/p/<pid>/text")
def file_text(pid):
    project = engine.load_project(pid)
    rel = request.args.get("rel", "")
    row = next((f for f in project.get("files") or [] if f["rel"] == rel), None)
    if not row:
        return jsonify({"ok": False, "error": "文件不存在"}), 404
    cache = row.get("cache") or ""
    data = textract.load_page_texts(cache) if cache else {}
    pages = data.get("pages") or []
    out, total = [], 0
    for pg in pages:
        text = pg.get("text") or ""
        total += len(text)
        out.append({"no": pg.get("no", 1), "src": pg.get("src", ""), "text": text[:20000]})
        if total > 200000:
            break
    return jsonify({"ok": True, "rel": rel, "name": row["name"],
                    "method": row.get("method", ""), "warn": row.get("warn", ""),
                    "pages": out or [{"no": 1, "src": "", "text": row.get("snippet", "")}]})


@app.route("/p/<pid>/openfile")
def open_file(pid):
    """双击未归类文件时，用系统默认程序打开原文件。"""
    project = engine.load_project(pid)
    rel = request.args.get("rel", "")
    row = next((f for f in project.get("files") or [] if f["rel"] == rel), None)
    if not row:
        return jsonify({"ok": False, "error": "文件不存在"}), 404
    path = Path(row.get("abs") or "")
    if not path.is_file() and project.get("unit_path") and " → " not in rel:
        unit = Path(project["unit_path"])
        try:
            cand = (unit / rel).resolve()
            if unit.resolve() in cand.parents and cand.is_file():
                path = cand
        except (OSError, ValueError):
            pass
    if not path.is_file():
        return jsonify({"ok": False, "error": "原文件已移动或删除"}), 404
    try:
        os.startfile(str(path))  # noqa: S606
    except OSError as exc:
        return jsonify({"ok": False, "error": f"打不开：{exc}"}), 500
    return jsonify({"ok": True})


@app.route("/p/<pid>/generate", methods=["POST"])
def generate(pid):
    project = engine.load_project(pid)
    if project.get("status") != "done":
        return jsonify({"ok": False, "error": "比对还没完成"})
    data = request.get_json(force=True) or {}

    by_key = {r["key"]: r for r in project["rows"]}
    for upd in data.get("rows", []):
        row = by_key.get(upd.get("key"))
        if not row:
            continue
        row["status"] = upd.get("status", row["status"])
        row["write"] = bool(upd.get("write"))
        if "report" in upd:
            row["report"] = upd["report"]

    project["rows"] = [r for r in project["rows"] if not r["key"].startswith("X")]
    for i, extra in enumerate(data.get("extras", [])):
        label = (extra.get("label") or "").strip()
        if not label:
            continue
        project["rows"].append({
            "key": f"X{i + 1}", "dir": "（手动追加）", "label": label,
            "status": "缺失", "note": "手动追加", "hits": [],
            "report": label, "merge_group": "", "merge_role": "",
            "write": True, "sort": len(project["rows"]),
        })

    items = engine.build_items(project["rows"], engine.load_rules())
    if not items:
        return jsonify({"ok": False, "error": "没有勾选任何缺件"})

    name = (data.get("project_name") or project.get("project_name") or project["name"]).strip()
    project["project_name"] = name

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe = "".join(c for c in name if c not in '\\/:*?"<>|')[:40]
    out_docx = engine.OUT / f"{safe}-资料交接单-{stamp}.docx"
    out_xlsx = engine.OUT / f"{safe}-比对明细-{stamp}.xlsx"

    try:
        docx_fill.fill_items(project["docx_path"], items, out_docx, project_name=name)
        exporter.export(project, out_xlsx)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "error": f"生成失败：{exc}"})

    project["items"] = items
    project["outputs"] = {"docx": str(out_docx), "xlsx": str(out_xlsx)}
    engine.save_project(project)

    return jsonify({
        "ok": True,
        "items": items,
        "docx": url_for("download", pid=pid, kind="docx"),
        "xlsx": url_for("download", pid=pid, kind="xlsx"),
        "docx_name": out_docx.name,
    })


@app.route("/p/<pid>/file/<kind>")
def download(pid, kind):
    project = engine.load_project(pid)
    path = project.get("outputs", {}).get(kind)
    if not path or not Path(path).exists():
        flash("文件不存在")
        return redirect(url_for("review", pid=pid))
    return send_file(path, as_attachment=True, download_name=Path(path).name)


@app.route("/p/<pid>/open")
def open_folder(pid):
    project = engine.load_project(pid)
    path = project.get("outputs", {}).get("docx")
    if path and Path(path).exists():
        os.startfile(str(Path(path).parent))  # noqa: S606
    return redirect(url_for("review", pid=pid))


@app.route("/rules", methods=["GET", "POST"])
def rules_page():
    error = None
    if request.method == "POST":
        text = request.form.get("json", "")
        try:
            obj = json.loads(text)
            if "requirements" not in obj:
                raise ValueError("缺少 requirements 字段")
            engine.save_rules(obj)
            flash("规则已保存。已完成的核查可点「按新规则重跑」重新判一次。")
            return redirect(url_for("rules_page"))
        except Exception as exc:  # noqa: BLE001
            error = f"JSON 格式有误：{exc}"
    text = json.dumps(engine.load_rules(), ensure_ascii=False, indent=2)
    if error:
        text = request.form.get("json", text)
    return render_template("rules.html", text=text, error=error)


def _pick_with_powershell(kind: str):
    """打包成 exe 后没有 Python 子进程，用 Windows 自带对话框选文件夹/文件。"""
    ps = shutil.which("powershell.exe") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    if kind == "docx":
        script = (
            "Add-Type -AssemblyName System.Windows.Forms;"
            "$d=New-Object System.Windows.Forms.OpenFileDialog;"
            "$d.Title='选择资料交接单 Word';$d.Filter='Word 文档|*.docx|所有文件|*.*';"
            "if($d.ShowDialog() -eq 'OK'){[Console]::OutputEncoding=[Text.Encoding]::UTF8;Write-Output $d.FileName}"
        )
    else:
        script = (
            "Add-Type -AssemblyName System.Windows.Forms;"
            "$d=New-Object System.Windows.Forms.FolderBrowserDialog;"
            "$d.Description='选择文件夹（会递归读取所有子文件夹和图片）';$d.ShowNewFolderButton=$false;"
            "if($d.ShowDialog() -eq 'OK'){[Console]::OutputEncoding=[Text.Encoding]::UTF8;Write-Output $d.SelectedPath}"
        )
    out = subprocess.run(
        [ps, "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return (out.stdout or "").strip(), (out.stderr or "")[-200:]


@app.route("/pick")
def pick():
    kind = request.args.get("kind", "dir")
    try:
        if FROZEN:
            path, err = _pick_with_powershell(kind)
        else:
            out = subprocess.run(
                [sys.executable, str(BASE / "pick.py"), kind],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
            path, err = (out.stdout or "").strip(), (out.stderr or "")[-200:]
        return jsonify({"ok": bool(path), "path": path, "err": err})
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "err": str(exc)})


@app.template_filter("filesize")
def filesize(n):
    n = n or 0
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.4)
        return s.connect_ex(("127.0.0.1", port)) == 0


def open_browser_when_ready(url: str, port: int) -> None:
    for _ in range(150):
        if port_in_use(port):
            webbrowser.open(url)
            return
        time.sleep(0.2)
    webbrowser.open(url)


if __name__ == "__main__":
    url = f"http://127.0.0.1:{PORT}"

    if port_in_use(PORT):
        print("检测到系统已经在运行，正在打开页面……")
        webbrowser.open(url)
        sys.exit(0)

    (BASE / "data" / "app.pid").write_text(str(os.getpid()), encoding="utf-8")
    removed = engine.clean_work(14)
    if removed:
        print(f"已清理 {removed} 个过期解压缓存")
    with open(BASE / "logs" / "start.log", "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] start pid={os.getpid()}\n")

    if not os.environ.get("AUDIT_NO_BROWSER"):
        threading.Thread(target=open_browser_when_ready, args=(url, PORT), daemon=True).start()
    print(f"服务已启动：{url}")
    print("使用期间请不要关闭这个窗口。")
    try:
        app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)
    except OSError as exc:
        print(f"\n启动失败：{exc}")
        print("可能是端口被其他程序占用。可以运行 stop.bat 后再试。")
        sys.exit(1)
