# -*- coding: utf-8 -*-
"""端到端自测：读内容 → 归类 → 生成交接单。

用法：
    .venv\\Scripts\\python.exe tools\\selftest.py <模板路径> <送审路径> <交接单.docx>

不带参数时使用下面三个默认路径（不存在就跳过）。
"""
import os
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))

import docx_fill  # noqa: E402
import engine  # noqa: E402
import exporter  # noqa: E402

DEFAULT_TPL = Path(r"C:\Users\Kevin\Desktop\报审资料模板")
DEFAULT_UNIT = Path(os.environ.get("AUDIT_TEST_UNIT_DIR", ""))
DEFAULT_DOCX = Path(os.environ.get("AUDIT_TEST_DOCX", ""))


def main():
    tpl = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_TPL
    unit = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_UNIT
    docx = Path(sys.argv[3]) if len(sys.argv) > 3 else DEFAULT_DOCX
    for name, path in (("模板", tpl), ("送审", unit), ("交接单", docx)):
        if not path.exists():
            print(f"[X] {name}不存在：{path}")
            return 1

    def progress(phase, done, total, message=""):
        if phase != "OCR 识别扫描件" or done == total:
            print(f"  {phase} {done}/{total} {message}")

    print("1. 读取 + 归类……")
    result = engine.compare(tpl, unit, engine.load_rules(),
                            {"read_content": True, "ocr": True}, progress)
    print("   统计：", result["stats"], "文件：", len(result["files"]),
          "未归类：", len(result["unclassified"]))
    for row in result["rows"]:
        hit = row["hits"][0] if row["hits"] else None
        print(f"   [{row['status']}] {row['key']:8s} {row['label'][:24]:26s} "
              f"{(hit['name'][:36] if hit else '')}")

    project = engine.new_project("selftest", str(tpl), str(unit), str(docx), result)
    items = engine.build_items(project["rows"], engine.load_rules())
    print("2. 交接单条目：")
    for i, x in enumerate(items, 1):
        print(f"   {i}. 未见{x}；")

    out_docx = engine.OUT / "自测-交接单.docx"
    out_xlsx = engine.OUT / "自测-比对明细.xlsx"
    docx_fill.fill_items(docx, items, out_docx, project_name=project["project_name"])
    exporter.export(project, out_xlsx)
    print("3. 已生成：")
    print("   ", out_docx)
    print("   ", out_xlsx)
    print("4. 交接单回读：")
    for line in docx_fill.read_back(out_docx):
        print("   ", repr(line))
    return 0


if __name__ == "__main__":
    sys.exit(main())
