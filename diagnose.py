# -*- coding: utf-8 -*-
"""环境诊断：把关键信息打印出来，方便排查启动问题。"""
import socket
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent


def line(title):
    print(f"\n[{title}]")


print("=" * 46)
print("  报审资料核查 - 环境诊断")
print("=" * 46)

line("程序目录")
print(f"  {BASE}")

line("Python")
print(f"  {sys.version.split()[0]}  {sys.executable}")
try:
    import docx
    import flask
    import openpyxl
    from importlib.metadata import version
    print(f"  flask {version('flask')} / python-docx {version('python-docx')} / openpyxl {version('openpyxl')}")
except Exception as exc:  # noqa: BLE001
    print(f"  [X] 基础依赖缺失：{exc}")

try:
    import pypdf
    print(f"  pypdf {pypdf.__version__}  OK（PDF 文字层）")
except Exception as exc:  # noqa: BLE001
    print(f"  [X] pypdf 缺失：{exc}（PDF 内容读不出来）")

try:
    import xlrd
    print(f"  xlrd {xlrd.__version__}  OK（.xls 表格）")
except Exception as exc:  # noqa: BLE001
    print(f"  [X] xlrd 缺失：{exc}（.xls 表格读不出来）")

try:
    import win32com.client  # noqa: F401
    print("  pywin32 OK（可调用本机 Word 读 .doc）")
except Exception as exc:  # noqa: BLE001
    print(f"  [!] pywin32 缺失：{exc}（.doc 会退回只提取关键词）")

line("OCR（扫描件识别）")
ps = BASE / "tools" / "win_ocr.ps1"
print(f"  {'OK  ' if ps.exists() else 'MISS'} tools\\win_ocr.ps1")
try:
    out = subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command",
         "[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null;"
         "[Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages | ForEach-Object { $_.LanguageTag }"],
        capture_output=True, text=True, timeout=30)
    langs = [x.strip() for x in (out.stdout or "").splitlines() if x.strip()]
    print(f"  系统 OCR 语言：{', '.join(langs) if langs else '无'}")
    if not any("zh" in x.lower() for x in langs):
        print("  [X] 没有中文 OCR 语言包，扫描件会读不出中文（设置→时间和语言→语言→中文→可选功能→OCR）")
except Exception as exc:  # noqa: BLE001
    print(f"  [!] 无法检测 OCR：{exc}")

line("关键文件")
for name in ("app.py", "engine.py", "textract.py", "classify.py",
             "docx_fill.py", "exporter.py", "rules.json"):
    print(f"  {'OK  ' if (BASE / name).exists() else 'MISS'} {name}")

line("端口 8765")
with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
    s.settimeout(0.5)
    in_use = s.connect_ex(("127.0.0.1", 8765)) == 0
print("  正在监听（服务已启动）" if in_use else "  没有程序监听（服务未启动）")

if in_use:
    out = subprocess.run(["netstat", "-ano"], capture_output=True, text=True,
                         encoding="mbcs", errors="replace").stdout
    for row in out.splitlines():
        if ":8765" in row and "LISTENING" in row:
            print("  " + " ".join(row.split()))

for name in ("start.log", "compare.log", "server.err.log"):
    path = BASE / "logs" / name
    line(name)
    if path.exists():
        rows = path.read_text(encoding="utf-8", errors="replace").splitlines()[-8:]
        for r in rows:
            print("  " + r)
    else:
        print("  （暂无）")

print("\n把本窗口内容截图发给维护人员即可。")
