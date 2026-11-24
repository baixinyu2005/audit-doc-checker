# -*- coding: utf-8 -*-
"""用 Flask test client 走一遍主流程。"""
import sys
from pathlib import Path

sys.path.insert(0, r"D:\报审资料核查系统")
import app as web  # noqa: E402

TPL = r"C:\Users\Kevin\Documents\deepseek-harness\default-workspace\审计资料核查\工作区\模板\报审资料模板"
UNIT = r"C:\Users\Kevin\Documents\deepseek-harness\default-workspace\审计资料核查\工作区\单位资料\2026.7.8报审--国网河北保定供电公司高碑店市10kV肖官营512等23条配电自动化标准化线路改造"
DOCX = r"C:\Users\Kevin\Desktop\国网河北保定供电公司高碑店市10kV肖官营512等23条配电自动化标准化线路改造资料交接单(3).docx"

out = []
c = web.app.test_client()

r = c.get("/")
out.append(f"GET / -> {r.status_code}, 新建核查={'新建核查' in r.get_data(as_text=True)}")

r = c.post("/compare", data={
    "template_path": TPL, "unit_path": UNIT, "docx_path": DOCX,
    "name": "国网河北保定供电公司高碑店市10kV肖官营512等23条配电自动化标准化线路改造",
}, follow_redirects=True)
text = r.get_data(as_text=True)
out.append(f"POST /compare -> {r.status_code}, 含项目名={'肖官营512' in text}, 含表格={'工作事项' in text or '资料名称' in text}")

# 从 URL 拿 pid
pid = r.request.path.split("/p/")[-1] if "/p/" in r.request.path else None
out.append(f"pid={pid}")

if pid:
    r = c.post(f"/p/{pid}/generate", json={"rows": [], "extras": [{"label": "测试追加项"}]})
    j = r.get_json()
    out.append(f"POST generate -> {r.status_code}, ok={j.get('ok')}, items={j.get('items')}")
    if j.get("ok"):
        d = c.get(j["docx"])
        out.append(f"GET docx -> {d.status_code}, bytes={len(d.data)}")
        x = c.get(j["xlsx"])
        out.append(f"GET xlsx -> {x.status_code}, bytes={len(x.data)}")

Path(r"D:\报审资料核查系统\data\webtest.txt").write_text("\n".join(out), encoding="utf-8")
print("done")
