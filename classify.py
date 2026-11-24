# -*- coding: utf-8 -*-
"""按模板要求给文件内容归类，再判断每一项资料是已提供、缺失还是待确认。

分类依据按可信度排序：
1. 文件内容标题（首页前几百字）   —— 最可靠
2. 文件正文关键词
3. 文件名 / 所在目录
4. 角色词（施工 / 设计 / 监理）

扫描件先 OCR 成文字再参与判断；PDF 按页切分，一份 PDF 里装多份资料也能分开。
"""
import re
from pathlib import Path

import textract

STATUSES = ["已提供", "缺失", "不适用", "待确认"]

ROLE_PATTERNS = {
    "施工": r"施工|建设集团|电力建设|建设工程|建筑安装|工程公司|建设有限公司|安装公司",
    "设计": r"设计|勘测设计|设计院|设计研究院|电力设计",
    "监理": r"监理",
}
STRONG = 50
WEAK = 26


def _as_list(value):
    if not value:
        return []
    if isinstance(value, str):
        return [value]
    return [x for x in value if x]


def _rx(pattern):
    try:
        return re.compile(pattern, re.I)
    except re.error:
        return None


def _clean_seg(name: str) -> str:
    s = re.sub(r"^\s*\d{1,2}\s*[.、\-_]\s*", "", (name or "").strip())
    s = re.sub(r"[（(]\s*需?签章齐全\s*[)）]", "", s)
    return s.strip() or (name or "").strip()


def normalize_dir(value: str) -> str:
    return (value or "").replace("/", "\\").strip().strip("\\").casefold()


def template_leaves(root) -> list:
    root = Path(root)
    if not root.exists() or not root.is_dir():
        return []
    leaves = []
    for d in sorted(root.rglob("*")):
        if not d.is_dir():
            continue
        try:
            has_child_dir = any(x.is_dir() for x in d.iterdir())
        except OSError:
            continue
        if not has_child_dir:
            leaves.append(str(d.relative_to(root)).replace("/", "\\"))
    return leaves


def _auto_keywords(label: str) -> str:
    parts = [p for p in re.split(r"[、，,;；/\\]", label) if len(p) >= 2]
    if not parts:
        parts = [label]
    return "|".join(re.escape(p) for p in parts)


def _auto_req(leaf: str, idx: int) -> dict:
    parts = [p for p in leaf.split("\\") if p]
    leaf_name = _clean_seg(parts[-1]) if parts else leaf
    parent = _clean_seg(parts[-2]) if len(parts) > 1 else ""
    generic = leaf_name in ("导出版及扫描件", "扫描件", "扫描件需签章齐全", "扫描件（需签章齐全）")
    if generic and parent:
        label = parent
    elif leaf_name in ("施工", "设计", "监理") and parent:
        label = f"{parent}（{leaf_name}）"
    else:
        label = leaf_name
    keywords = _auto_keywords(label)
    return {"id": f"T{idx}", "dir": leaf, "label": label, "match": keywords,
            "content": keywords, "required": True, "report": label, "auto": True}


def _compile(req: dict) -> dict:
    req = dict(req)
    label = req.get("label", "")
    label_core = re.sub(r"[（(].*?[)）]", "", label).strip() or label
    content = _as_list(req.get("content"))
    if label_core and len(label_core) >= 3:
        content = content + [re.escape(label_core)]
    req["_path"] = [x for x in (_rx(p) for p in _as_list(req.get("match"))) if x]
    req["_title"] = [x for x in (_rx(p) for p in _as_list(req.get("title"))) if x]
    req["_content"] = [x for x in (_rx(p) for p in content) if x]
    req["_title_excl"] = [x for x in (_rx(p) for p in _as_list(req.get("title_exclude"))) if x]
    req["_excl"] = [x for x in (_rx(p) for p in _as_list(req.get("exclude"))) if x]
    req["_content_excl"] = [x for x in (_rx(p) for p in _as_list(req.get("content_exclude"))) if x]
    req["_na"] = [x for x in (_rx(p) for p in _as_list(req.get("na_if"))) if x]
    req["_role_ok"] = _rx(ROLE_PATTERNS.get(req.get("role"), "")) if req.get("role") else None
    req["_role_bad"] = [(_rx(p), k) for k, p in ROLE_PATTERNS.items() if k != req.get("role")]
    return req


def compile_requirements(rules: dict, template_dir=None):
    """规则表 + 模板目录 → 可执行的要求列表。"""
    reqs = [_compile(raw) for raw in rules.get("requirements", [])]

    leaves = template_leaves(template_dir) if template_dir else []
    known = {normalize_dir(r.get("dir")) for r in reqs}
    auto_n = 0
    for leaf in leaves:
        if normalize_dir(leaf) in known:
            continue
        auto_n += 1
        reqs.append(_compile(_auto_req(leaf, auto_n)))

    template_info = {
        "leaves": leaves,
        "leaf_count": len(leaves),
        "auto_count": auto_n,
        "rule_count": len(reqs) - auto_n,
        "unmatched_rules": [r.get("dir") for r in reqs
                            if not r.get("auto") and r.get("dir") and leaves
                            and normalize_dir(r.get("dir")) not in {normalize_dir(x) for x in leaves}],
    }
    return reqs, template_info


def _title_hit(req, title):
    """标题必须出现在页首（前 30 字内），避免正文里提到「施工合同」就被当成合同。"""
    for pat in req["_title"]:
        m = pat.search(title)
        if m and m.start() <= 30:
            return True
    return False


def _score(req, rel: str, head: str, title: str, body: str, has_text: bool):
    """返回 (分数, 依据列表, 命中词)。"""
    score, why, terms = 0, [], []
    path_hit = any(r.search(rel) for r in req["_path"])
    title_hit = False

    if path_hit:
        score += 32
        why.append("文件名/目录")

    if any(r.search(rel) for r in req["_excl"]):
        return -999, [], []

    if has_text:
        if any(r.search(title) for r in req["_title_excl"]):
            return -999, [], []
        if any(r.search(body) for r in req["_content_excl"]):
            return -999, [], []

        if _title_hit(req, title):
            title_hit = True
            score += 75
            why.append("标题")

        hits = 0
        for pat in req["_content"]:
            m = pat.search(body)
            if m:
                hits += 1
                terms.append(m.group(0))
                if hits >= 3:
                    break
        if hits:
            score += 22 * hits
            why.append("正文关键词")

        if req.get("_role_ok"):
            role_name = req.get("role", "")
            role_hit = False
            if req["_role_ok"].search(head):
                score += 55
                why.append(f"{role_name}角色")
                role_hit = True
            elif req["_role_ok"].search(title):
                score += 30
                why.append(f"{role_name}角色")
                role_hit = True
            elif req["_role_ok"].search(body):
                score += 12
                why.append(f"{role_name}角色")
                role_hit = True
            elif req["_role_ok"].search(rel):
                score += 30
                why.append(f"{role_name}角色")
                role_hit = True
            else:
                # 标题里出现了别的角色（如「施工监理」），且本角色未出现，才扣分
                for bad_re, _bad_name in req["_role_bad"]:
                    if bad_re and bad_re.search(head):
                        score -= 45
                        break
            if not role_hit:
                # 中标通知书必须能看出施工/设计/监理角色，否则留给人工指定
                score = min(score, 20)

    # 有明确标题的要求（合同、中标通知书、结算书…）：只有标题或文件名/目录命中才能归类，
    # 防止在别的文件正文里偶然提到「施工合同」「中标金额」就被算作已提供。
    if req["_title"] and req.get("strict", True) and not path_hit and not title_hit:
        score = min(score, 20)

    return score, why, terms


def _evidence(page_text: str, terms, limit: int = 120) -> str:
    for term in terms or []:
        if not term:
            continue
        chars = list(term)[:12]
        raw = re.compile(r"\s*".join(re.escape(c) for c in chars), re.I)
        m = raw.search(page_text)
        if m:
            start = max(0, m.start() - 24)
            end = min(len(page_text), m.end() + 50)
            return textract.snippet(page_text[start:end], limit)
    return textract.snippet(page_text, limit)


def _segments(pages, per_page):
    """按页切分：一份 PDF 里可能装着开工报告 + 竣工报告。"""
    anchors = set()
    for i, scored in enumerate(per_page):
        if not scored:
            continue
        top = scored[0]
        second = scored[1][0] if len(scored) > 1 else 0
        if top[0] >= STRONG and top[0] - second >= 12:
            anchors.add(i)

    if not anchors:
        return []

    segments, cur = [], None
    for i in range(len(pages)):
        scored = per_page[i]
        top = scored[0] if scored else (0, None, [], [])
        is_anchor = i in anchors and top[1] is not None
        if is_anchor and (cur is None or top[1]["id"] != cur["req"]["id"]):
            if cur is not None:
                cur["end"] = i - 1
                segments.append(cur)
                cur = {"req": top[1], "start": i, "end": None, "pages": [i],
                       "score": top[0], "why": top[2], "terms": top[3]}
            else:
                cur = {"req": top[1], "start": 0, "end": None,
                       "pages": list(range(0, i + 1)),
                       "score": top[0], "why": top[2], "terms": top[3]}
        if cur is not None:
            if i not in cur["pages"]:
                cur["pages"].append(i)
            if top[0] > cur["score"]:
                cur["score"] = top[0]
                cur["why"] = top[2]
                cur["terms"] = top[3]
    if cur is not None:
        cur["end"] = len(pages) - 1
        segments.append(cur)
    return segments


def classify_file(f, res, reqs):
    pages = res.get("pages") or []
    if not pages:
        pages = [{"no": 1, "text": "", "src": ""}]
    has_text = any((p.get("text") or "").strip() for p in pages)
    rel = f["rel"]

    per_page = []
    for pg in pages:
        raw = pg.get("text") or ""
        head = textract.normalize_for_match(raw[:140])
        title = textract.normalize_for_match(raw[:80])
        body = textract.normalize_for_match(raw[:12000])
        scored = []
        for req in reqs:
            score, why, terms = _score(req, rel, head, title, body, bool(body))
            if score > -900:
                scored.append((score, req, why, terms))
        scored.sort(key=lambda x: x[0], reverse=True)
        per_page.append(scored)

    segs = _segments(pages, per_page)
    ambiguous = False
    if not segs:
        agg = {}
        for scored in per_page:
            for score, req, why, terms in scored:
                cur = agg.get(req["id"])
                if cur is None or score > cur[0]:
                    agg[req["id"]] = (score, req, why, terms)
        ranked = sorted(agg.values(), key=lambda x: x[0], reverse=True)
        if ranked and ranked[0][0] >= WEAK:
            best = ranked[0]
            segs = [{"req": best[1], "start": 0, "end": len(pages) - 1,
                     "pages": list(range(len(pages))), "score": best[0],
                     "why": best[2], "terms": best[3]}]
            if len(ranked) > 1 and ranked[1][0] >= STRONG and best[0] - ranked[1][0] <= 8:
                ambiguous = True

    by_req = {}
    for seg in segs:
        req_id = seg["req"]["id"]
        item = by_req.setdefault(req_id, {"req": seg["req"], "pages": [],
                                          "score": seg["score"], "why": seg["why"],
                                          "terms": seg["terms"]})
        item["pages"].append([seg["start"] + 1, seg["end"] + 1])
        if seg["score"] > item["score"]:
            item["score"] = seg["score"]
            item["why"] = seg["why"]
            item["terms"] = seg["terms"]

    hits = []
    for req_id, item in by_req.items():
        page_nos = sorted({n for a, b in item["pages"] for n in range(a, b + 1)})
        sample = ""
        for pg in pages:
            if pg.get("no") in page_nos and (pg.get("text") or "").strip():
                sample = pg.get("text") or ""
                break
        how = "、".join(item["why"]) or "按关键词"
        if not has_text:
            how = "文件名/目录（内容未读出）"
        hits.append({
            "req": req_id, "rel": rel, "name": f["name"], "size": f.get("size", 0),
            "pages": sorted(item["pages"]), "page_nos": page_nos,
            "snippet": _evidence(sample, item["terms"]) if sample else "",
            "how": how, "score": item["score"], "ambiguous": ambiguous,
        })

    hits.sort(key=lambda h: (-h["score"], h["rel"]))
    return {"rel": rel, "hits": hits, "ambiguous": ambiguous,
            "unclassified": not hits}


def _search_all(files, texts, pattern, limit=3):
    """清单外必查项：先看文件名，再看文件标题区。"""
    rx = _rx(pattern)
    if not rx:
        return []
    seen_names, uniq = set(), []
    for f in files:
        key = (f["name"], f.get("size", 0))
        if key in seen_names:
            continue
        seen_names.add(key)
        uniq.append(f)
    files = uniq
    found = []
    for f in files:
        if rx.search(f["rel"]):
            found.append({"rel": f["rel"], "name": f["name"], "size": f.get("size", 0),
                          "pages": [[1, 1]], "snippet": "文件名命中",
                          "how": "文件名", "score": 55})
            if len(found) >= limit:
                return found
    if found:
        return found

    for f in files:
        res = texts.get(f["rel"]) or {}
        is_img = Path(f["name"]).suffix.lower() in textract.IMAGE_EXT
        hit = None
        for pg in res.get("pages") or []:
            text = pg.get("text") or ""
            if not text:
                continue
            # 图片也要参与：识别出的文字算线索，但只作「待确认」，不直接算正式资料
            head = textract.normalize_for_match(text[:160])
            if rx.search(head):
                hit = {"rel": f["rel"], "name": f["name"], "size": f.get("size", 0),
                       "pages": [[pg.get("no", 1), pg.get("no", 1)]],
                       "snippet": textract.snippet(text, 160),
                       "how": "图片 OCR" if is_img else "正文标题",
                       "score": 50 if is_img else 60,
                       "weak": is_img}
                break
        if hit:
            found.append(hit)
            if len(found) >= limit:
                break
    return found


def _group_candidates(group_name, group_reqs, hits_by_req, files, texts, limit=12):
    assigned = {h["rel"] for r in group_reqs for h in hits_by_req.get(r["id"], [])}
    out = []
    for f in files:
        if f["rel"] in assigned:
            continue
        res = texts.get(f["rel"]) or {}
        title = textract.normalize_for_match((res.get("text") or "")[:200])
        matched = False
        for req in group_reqs:
            if any(p.search(f["rel"]) for p in req["_path"]):
                matched = True
            elif any(p.search(title) for p in req["_title"]):
                matched = True
            if matched:
                break
        if matched:
            out.append({"rel": f["rel"], "name": f["name"], "size": f.get("size", 0),
                        "snippet": textract.snippet(res.get("text") or "")})
            if len(out) >= limit:
                break
    return out


def build_rows(files, texts, reqs, rules, template_info=None):
    file_map = {f["rel"]: classify_file(f, texts.get(f["rel"]) or {}, reqs) for f in files}

    hits_by_req = {}
    for info in file_map.values():
        for hit in info["hits"]:
            hits_by_req.setdefault(hit["req"], []).append(hit)

    all_norm = "\n".join(
        textract.normalize_for_match(pg.get("text") or "")
        for f in files for pg in ((texts.get(f["rel"]) or {}).get("pages") or []))
    all_norm += "\n" + textract.normalize_for_match(" ".join(f["rel"] for f in files))

    groups = {}
    for req in reqs:
        if req.get("role_group"):
            groups.setdefault(req["role_group"], []).append(req)

    rows = []
    for req in reqs:
        req_hits = hits_by_req.get(req["id"], [])
        # 同一项命中多份时：优先展示文件名/目录也对上的，其次看分数
        req_hits.sort(key=lambda h: (0 if "文件名/目录" in (h.get("how") or "") else 1,
                                    -h["score"]))
        seen_hits, uniq = set(), []
        for h in req_hits:
            k = (h["name"], h["size"])
            if k in seen_hits:
                continue
            seen_hits.add(k)
            uniq.append(h)
        req_hits = uniq
        cands = []
        if req.get("role_group") and not req.get("na_if"):
            cands = _group_candidates(req["role_group"], groups.get(req["role_group"], []),
                                      hits_by_req, files, texts)

        status, note = "", ""
        if any(r.search(all_norm) for r in req["_na"]):
            status, note = "不适用", "单位资料里出现不涉及的说明"
        elif req_hits:
            status = "已提供"
            content_based = any(any(k in (h["how"] or "") for k in ("标题", "正文", "内容"))
                                for h in req_hits)
            if req_hits[0].get("ambiguous"):
                status, note = "待确认", "内容同时像多项资料，请人工确认"
            elif not content_based:
                if req_hits[0]["how"].startswith("文件名/目录（内容未读出）"):
                    note = "按文件名/目录匹配；内容未读出（扫描件/图纸）"
                else:
                    note = "按文件名/目录匹配（内容未命中关键词）"
            if status == "已提供" and not any((h.get("snippet") or "").strip() for h in req_hits) \
                    and any("模板" in (h["name"] or "") for h in req_hits):
                status, note = "待确认", "命中的是模板/样表文件，疑似未填写"
        elif req.get("role_group") and cands:
            status = "待确认"
            note = f"找到 {len(cands)} 份「{req['role_group']}」，但分不清角色，可拖拽指定"
        elif req.get("required", True):
            status, note = "缺失", "模板要求，未找到对应文件或内容"
        else:
            status, note = "待确认", "兜底目录，按需确认"

        rows.append({
            "key": req["id"], "dir": req.get("dir", ""), "label": req["label"],
            "status": status, "note": note, "hits": req_hits[:8],
            "candidates": cands, "assigned": [],
            "report": req.get("report", req["label"]),
            "merge_group": req.get("merge_group", ""),
            "merge_role": req.get("merge_role", ""),
            "container": req.get("container", False),
            "auto": req.get("auto", False),
            "sort": len(rows), "write": status == "缺失",
        })

    for item in rules.get("manual_items", []):
        if not item.get("enabled", True):
            continue
        hits = _search_all(files, texts, item.get("match"))
        if hits and all(h.get("weak") for h in hits):
            status, note = "待确认", "图片里识别到关键词，建议人工确认是否算正式资料"
        elif hits:
            status, note = "已提供", "在文件名或内容中找到"
        elif item.get("default_missing", item.get("confirmed_missing")):
            status, note = "缺失", "清单外必查项，请核对后保留或删除"
        else:
            status, note = "待确认", "清单外必查项，请人工确认"
        rows.append({
            "key": item["id"], "dir": "（清单外人工确认项）", "label": item["label"],
            "status": status, "note": note, "hits": hits, "candidates": [],
            "assigned": [], "report": item["label"],
            "merge_group": item.get("merge_group", ""),
            "merge_role": item.get("merge_role", ""),
            "container": False, "sort": len(rows), "write": status == "缺失",
        })

    unclassified = []
    for f in files:
        if not file_map[f["rel"]]["hits"]:
            res = texts.get(f["rel"]) or {}
            unclassified.append({"rel": f["rel"], "name": f["name"],
                                 "size": f.get("size", 0),
                                 "snippet": textract.snippet(res.get("text") or ""),
                                 "method": res.get("method", "")})
    return rows, file_map, unclassified
