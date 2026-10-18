#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
社团产出质量评分 · 统一入口

用法:
  python review.py --type plan    --file 策划案.docx
  python review.py --type poster  --image 海报.jpg
  python review.py --type ppt     --file 宣讲.pptx --slide 5
  python review.py --type copy    --text "文案正文..."
  python review.py --type plan    --file 策划案.docx --prescreen-only   # 只跑机器可检的一票否决
  python review.py --type plan    --file 策划案.docx --emit-prompt      # 导出提示词，任意模型可用
  python review.py --type plan    --submit 模型输出.json                # 导入模型结果并算档位学分

只依赖标准库。模型可选：配置了 lib/models.json 就自动调用，没配置就配合
--emit-prompt / --submit 人工搬运。这样 rubric 是纯文本标准，换任何模型都能跑。
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import sys
import urllib.error
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
RUBRICS = ROOT / "rubrics"
REFERENCES = ROOT / "references"
IDENTITY_FILE = HERE / "club-identity.json"
MODELS_FILE = HERE / "models.json"

DIMENSIONS = [
    ("completeness", "完整性", "D1"),
    ("accuracy", "准确性", "D2"),
    ("executability", "可执行性", "D3"),
    ("compliance", "规范性", "D4"),
    ("quality", "表达", "D5"),
]

TIER_COEF = {"A": 1.5, "B": 1.0, "C": 0.5, "退回": 0.0}

BASE_CREDIT = {
    ("plan", None): 0.3,
    ("copy", None): 0.1,
    ("poster", None): 0.1,
    ("ppt", "deck"): 0.2,
    ("ppt", "slide"): 0.05,
    ("contest", "entry"): 0.3,
    ("contest", "award"): 0.6,
}

CN_WEEK = {"一": 0, "二": 1, "三": 2, "四": 3, "五": 4, "六": 5, "日": 6, "天": 6}


# ----------------------------------------------------------------------------
# 文档解析（纯标准库）
# ----------------------------------------------------------------------------

def _xml_paras(xml_bytes: bytes) -> list[str]:
    """按段落切分，纯 <w:p>/<a:p> 顺序，保留段落边界。"""
    root = ET.fromstring(xml_bytes)
    out = []
    for node in root.iter():
        if node.tag.split("}")[-1] != "p":
            continue
        line = "".join(t.text or "" for t in node.iter() if t.tag.split("}")[-1] == "t")
        if line.strip():
            out.append(line.strip())
    return out


def _w(tag: str) -> str:
    return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + tag


def parse_docx(path: Path) -> dict:
    """返回 {"text": 全文, "lines": 行列表, "tables": [[行, 行...]], "meta": {...}}"""
    lines: list[str] = []
    tables: list[list[list[str]]] = []

    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        if "word/document.xml" in names:
            root = ET.fromstring(zf.read("word/document.xml"))
            body = root.find(_w("body"))
            if body is not None:
                for child in body:
                    tag = child.tag.split("}")[-1]
                    if tag == "p":
                        line = "".join(
                            t.text or "" for t in child.iter() if t.tag.split("}")[-1] == "t"
                        ).strip()
                        if line:
                            lines.append(line)
                    elif tag == "tbl":
                        rows = []
                        for tr in child.findall(_w("tr")):
                            cells = []
                            for tc in tr.findall(_w("tc")):
                                txt = "".join(
                                    t.text or ""
                                    for t in tc.iter()
                                    if t.tag.split("}")[-1] == "t"
                                ).strip()
                                cells.append(re.sub(r"\s+", " ", txt))
                            if any(cells):
                                rows.append(cells)
                        if rows:
                            tables.append(rows)
                            lines.append("__TABLE__%d__" % (len(tables) - 1))

        for name in sorted(n for n in names if re.search(r"word/(header|footer)\d*\.xml", n)):
            try:
                got = _xml_paras(zf.read(name))
            except ET.ParseError:
                continue
            if got:
                lines.append("[%s]" % name.split("/")[-1])
                lines.extend(got)

        media = [n for n in names if n.startswith("word/media/")]
        meta = {}
        for prop in ("docProps/core.xml", "docProps/app.xml"):
            if prop in names:
                blob = zf.read(prop).decode("utf-8", "ignore")
                for m in re.finditer(r"<(\w+)[^>]*>([^<]{1,120})</\1>", blob):
                    meta.setdefault(m.group(1), m.group(2))

    tables_dump = []
    for i, rows in enumerate(tables):
        block = ["[表格 %d]" % i]
        for r in rows:
            block.append(" | ".join(r))
        tables_dump.append("\n".join(block))

    text_parts = []
    for ln in lines:
        m = re.fullmatch(r"__TABLE__(\d+)__", ln)
        text_parts.append(tables_dump[int(m.group(1))] if m else ln)

    return {
        "text": "\n".join(text_parts),
        "lines": [ln for ln in lines if not re.fullmatch(r"__TABLE__\d+__", ln)],
        "tables": tables,
        "meta": meta,
        "images": len(media),
    }


def parse_pptx(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        slides = sorted(
            (n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)),
        )
        pages = []
        for name in slides:
            idx = int(re.search(r"(\d+)", name.split("/")[-1]).group(1))
            try:
                body = _xml_paras(zf.read(name))
            except ET.ParseError:
                body = []
            note = ""
            nn = "ppt/notesSlides/notesSlide%d.xml" % idx
            if nn in names:
                try:
                    note = "\n".join(
                        x for x in _xml_paras(zf.read(nn)) if not re.fullmatch(r"\d+", x)
                    )
                except ET.ParseError:
                    note = ""
            imgs = []
            rel = "ppt/slides/_rels/slide%d.xml.rels" % idx
            if rel in names:
                rels = zf.read(rel).decode("utf-8", "ignore")
                imgs = re.findall(r'Target="([^"]*media/[^"]+)"', rels)
            pages.append({"index": idx, "text": "\n".join(body), "notes": note, "images": imgs})
        return {"pages": pages, "images": len([n for n in names if n.startswith("ppt/media/")])}


def render_ppt_text(doc: dict, slide: int | None) -> str:
    if slide:
        for p in doc["pages"]:
            if p["index"] == slide:
                head = "===== 第 %d 页（单页评审）=====" % p["index"]
                # PDF 页没有 notes 字段（PDF 无幻灯片备注概念），用 .get 兜住。
                note = "\n[备注]\n%s" % p["notes"] if (p.get("notes") or "").strip() else ""
                return head + "\n" + p["text"] + note
        raise SystemExit("幻灯片不存在：第 %d 页（该 PPT 共 %d 页）" % (slide, len(doc["pages"])))

    out = ["[共 %d 页]" % len(doc["pages"])]
    for p in doc["pages"]:
        out.append("\n===== 第 %d 页 =====\n%s" % (p["index"], p["text"]))
        if (p.get("notes") or "").strip():
            out.append("[备注]\n%s" % p["notes"])
    return "\n".join(out)


def render_ppt_layout(doc: dict) -> str:
    """抽取每页的页脚/页码线索，用于查页码错乱。"""
    out = []
    for p in doc["pages"]:
        found = re.findall(r"(?<![\d.])(\d{1,3})\s*/\s*(\d{1,3})(?![\d.])",
                           re.sub(r"\s*\n\s*", " ", p["text"]))
        tail = re.sub(r"\s*\n\s*", " ", p["text"])[-60:]
        out.append("第 %d 页：声明 %s　尾行「%s」"
                   % (p["index"],
                      "、".join("%s/%s" % f for f in found) if found else "无",
                      tail))
    return "\n".join(out)


def parse_pdf(path: Path) -> dict:
    """PDF 文本层抽取。有 PyMuPDF 就用；没有则报错而不是静默降级。

    返回 {"text", "lines", "pages", "tables", "meta", "images"}，与 parse_docx 同构。
    pages 是 [{index, text}]，让单页模式也能用 --slide 看第 N 页。
    """
    try:
        import pymupdf  # noqa
    except ImportError:
        try:
            import fitz as pymupdf  # type: ignore
        except ImportError:
            raise SystemExit(
                "PDF 需要 PyMuPDF（pymupdf/fitz）才能抽文本层。\n"
                "已装：直接跑；未装：把内容另存为 .docx/.txt 后用 --file，"
                "或用 --text 粘贴。\n"
                "（不做静默降级——PDF 抽不出文本却给出评分，比报错更糟。）"
            )

    pages, meta = [], {}
    n_images = 0
    with pymupdf.open(path) as doc:
        meta = {k: v for k, v in (doc.metadata or {}).items() if v}
        for i, page in enumerate(doc):
            pages.append({"index": i + 1, "text": page.get_text().strip()})
            n_images += len(page.get_images(full=True))

    text_parts, lines = [], []
    for p in pages:
        text_parts.append("===== 第 %d 页 =====\n%s" % (p["index"], p["text"]))
        lines.extend(p["text"].splitlines())
    text = "\n".join(text_parts)

    # PDF 没有表格结构，但 PDF 里最常见的表是「名称 | 规格 | 数量 | 单价」这种竖排文本。
    # 这里不猜表格结构，改由行文本里的算术核对（P0-5）自己解析，避免猜错反而误报。
    return {
        "text": text,
        "lines": [ln for ln in lines if ln.strip()],
        "pages": pages,
        "tables": [],
        "meta": meta,
        "images": n_images,
    }


# ----------------------------------------------------------------------------
# 一票否决预筛
# ----------------------------------------------------------------------------

def load_identity() -> dict:
    return json.loads(IDENTITY_FILE.read_text(encoding="utf-8"))


def _num(s: str):
    s = s.replace(",", "").replace("，", "").strip()
    m = re.fullmatch(r"-?\d+(?:\.\d+)?", s)
    return float(m.group(0)) if m else None


def check_club_names(text: str, ident: dict, findings: list) -> None:
    for name, spec in ident.get("clubs", {}).items():
        code = "P0-1" if spec.get("status") == "confirmed" else "W-pending-club"
        level = "P0" if spec.get("status") == "confirmed" else "warn"
        for bad in spec.get("reject", []):
            if bad["text"] in text and not _is_correct_context(text, bad["text"], name):
                findings.append({
                    "code": code,
                    "level": level,
                    "title": "社团名写法错误" if level == "P0" else "待确认的社团名称",
                    "evidence": bad["reason"] if level == "P0"
                                else "文中出现「%s」，规范状态为 pending" % bad["text"],
                    "detail": "规范写法应为「%s」。%s" % (
                        name, "确认后若为错写将升级为 P0-1；当前仅在 D4 扣 1 分。" if level != "P0" else ""),
                })
    for club, spec in ident.get("clubs", {}).items():
        if spec.get("status") == "pending" and any(
            alt in text for alt in spec.get("accept", []) + [spec.get("canonical", "")]
        ):
            findings.append({
                "code": "W-pending-club",
                "level": "warn",
                "title": "待确认的社团名称",
                "evidence": "文中出现「%s」，其规范状态为 pending" % club,
                "detail": "确认后若为错写将升级为 P0-1；当前仅在 D4 扣 1 分。",
            })


def _is_correct_context(text: str, bad: str, right: str) -> bool:
    """如果同一句里已出现正确全称，视为对协办角色的正当简称，不判错。"""
    for line in text.splitlines():
        if bad in line and right not in line:
            return False
    return True


def check_plan_elements(text: str, findings: list, doc: dict | None = None,
                       doc_form: str = "report") -> None:
    """doc_form=report 是报批策划案（要请示语、签字位、落款、安全预案）；
    doc_form=scheme 是执行方案/技术设计文档（不送审，靠可落库程度说话）。

    P0-2 系列与 P0-3 是「报批公文的体例要求」，只对 report 生效。
    scheme 下它们降为提示：缺了不退回，但要扣完整性/规范性。
    依据：用户给的 95 分锚点样本 plan-lichuang 是执行方案，通篇无「妥否」、
    无签字位、无落款日期、无「主办」，但八章「风险与红线」列了 8 条真风险
    （锂电池过充、充电电流、无机械限位、TF 卡规格、OV5640 发热、蓝牙撞名、
    AI 断网、特价失效），算术全对。这类文档拿 report 标准去判就是误伤。
    """
    report = (doc_form == "report")
    lvl = "P0" if report else "warn"

    # P0-2 公文必备要素
    if "妥否" not in text:
        findings.append({
            "code": "P0-2a", "level": lvl, "title": "缺「妥否，请批示！」",
            "evidence": "全文无「妥否」二字",
            "detail": "报批策划案结尾必备结束语，缺此项老师会直接退回。"
                      if report else "执行方案通常不写请示语；仅当你要同时报批时需补。",
        })
    # P0-2b 指导教师签字位。正文里出现「签字归还」不算，必须是指导教师签字位或签字下划线。
    sig_block = re.search(r"指导\s*(?:教师|老师)[^\n]{0,24}?签\s*[字署]", text) or \
        re.search(r"签\s*[字署]\s*[_＿—\-]{2,}", text)
    if not sig_block:
        findings.append({
            "code": "P0-2b", "level": lvl, "title": "缺指导教师签字位",
            "evidence": "无「指导教师签字」字样，也无签字下划线行（正文中的「签字归还」不算）",
            "detail": "必须保留整行：指导教师签字 ______  日期 ____年__月__日"
                      if report else "执行方案通常不送审，无需签字位；若需报批请加。",
        })
    # 落款：文档中须至少出现一行「只有日期」的行（社团全称可在其上方或同行）
    solo_date = re.search(
        r"(?m)^\s*(?:\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日"
        r"|\d{4}\s*[./\-]\s*\d{1,2}\s*[./\-]\s*\d{1,2})\s*$",
        text,
    )
    if not solo_date:
        findings.append({
            "code": "P0-2c", "level": lvl, "title": "缺落款日期",
            "evidence": "全文找不到「单独成行的日期」",
            "detail": "落款格式：社团全称单独一行 + 年月日单独一行。"
                      if report else "执行方案通常不落款；若要归档请加版本号与日期。",
        })
    if not re.search(r"(主办|承办)", text):
        findings.append({
            "code": "P0-2d", "level": lvl, "title": "缺主办单位声明",
            "evidence": "全文无「主办」或「承办」",
            "detail": "策划案必须写明主办与协办单位。"
                      if report else "执行方案也建议写明责任社团，方便归档追溯。",
        })

    # P0-3 安全预案。注意：P0-3 无论 report 还是 scheme 都是 P0——
    # 缺风险章节是实质硬伤（出事无人知），不是体例问题。
    # 两种形态的差别只在「章节叫什么」和「清单怎么写」：
    # report 查 安全保障/应急预案 + 固定五类覆盖；
    # scheme 查 风险/红线/安全 等章节，且要求逐条给出「后果 + 对策」两列。
    safety_keys = (("安全保障", "应急预案") if report
                   else ("安全保障", "应急预案", "风险", "红线", "安全"))
    has_safety = any(k in text for k in safety_keys)
    if not has_safety:
        findings.append({
            "code": "P0-3", "level": "P0", "title": "缺安全保障与应急预案",
            "evidence": "全文无「%s」章节" % "」「".join(safety_keys[:2]),
            "detail": "策划案必须有该章节，且覆盖用电/消防/物品/秩序/不可抗力。"
                      if report else "执行方案必须列风险章节（后果 + 对策），"
                                         "否则出问题时无人知道怎么办。",
        })
    else:
        if report:
            # 报批策划案的安全预案按固定五类查（用电/消防/物品/秩序/噪音）
            keys = ["用电", "消防", "物品", "秩序", "噪音", "不可抗力"]
            hit = [k for k in keys if k in text]
            if len(hit) < 3:
                findings.append({
                    "code": "P1-safety", "level": "warn", "title": "安全预案覆盖不足",
                    "evidence": "仅覆盖 %s" % ("、".join(hit) if hit else "无"),
                    "detail": "P1：覆盖不足三类，D1 扣 1 分。",
                })
        else:
            # 执行方案的风险清单按自己的体例查：必须给「后果」和「对策」两列，
            # 否则风险清单只是一串名词，出事时无人知道怎么办。
            # 依据 plan-lichuang：八章「风险与红线」表列为 事项 | 后果 | 对策。
            if not ("后果" in text and ("对策" in text or "应对" in text or "缓解" in text)):
                findings.append({
                    "code": "P1-safety", "level": "warn", "title": "风险清单缺「后果 + 对策」两列",
                    "evidence": "有风险章节，但未见「后果」与「对策/应对」成对出现",
                    "detail": "P1：执行方案的风险清单要逐条写清后果与对策，D1/D3 各扣 1 分。",
                })

    # P0-4 预算无明细：出现经费数字但既无单价也无金额列
    has_budget = re.search(r"(预算|经费)", text)
    has_price = re.search(r"单价", text)
    has_amount = re.search(r"(金额|合计|小计)", text)
    if has_budget and not (has_price or has_amount):
        findings.append({
            "code": "P0-4", "level": "P0", "title": "预算无明细",
            "evidence": "文中出现「%s」但全文无「单价」与「金额」" % has_budget.group(0),
            "detail": "预算必须逐项列出名称、规格、单价、数量、金额，不能只写总额。",
        })

    # P0-4b 物料表出现金额数字却没有单价列：说明预算被塞进备注，评审无法核算。
    # 物料表的判据是「有数量列 / 物料名列」，而不是表头出现「预算」二字
    # —— 实测样本的金额是写在「备注」列里的，表头只有 序号/设备/数量/来源/备注。
    for ti, rows in enumerate((doc or {}).get("tables", [])):
        if len(rows) < 2:
            continue
        header = rows[0]
        looks_like_material = any(
            ("数量" in c) or ("物料" in c) or ("设备" in c) or ("用品" in c) or ("用品" in c)
            for c in header
        )
        has_unit_price = any("单价" in c or "价格" in c or "金额" in c for c in header)
        if not looks_like_material or has_unit_price:
            continue
        money_cells = []
        for r in rows[1:]:
            for cell in r:
                if re.search(r"\d+(?:\.\d+)?\s*元", cell):
                    money_cells.append(cell.strip())
        if money_cells:
            findings.append({
                "code": "P0-4b", "level": "P0", "title": "物料/预算表缺单价与金额列",
                "evidence": "表 %d（%s）无「单价」「金额」列，但表体出现金额：%s"
                            % (ti + 1, "|".join(header), "、".join(money_cells[:3])),
                "detail": "预算明细必须能逐行核算：名称、规格、单价、数量、金额，缺一列都不算明细。",
            })
            break


def check_budget_math(doc: dict, findings: list) -> None:
    """P0-5：对含单价/数量/金额列的表做逐行与汇总核对。"""
    for ti, rows in enumerate((doc or {}).get("tables", [])):
        if not rows:
            continue
        header = rows[0]
        try:
            ci_price = next(i for i, c in enumerate(header) if "单价" in c)
            ci_qty = next(i for i, c in enumerate(header) if "数量" in c)
            ci_amt = next(i for i, c in enumerate(header) if c.strip() in ("金额", "金额(元)", "金额（元）"))
        except StopIteration:
            continue

        items = []
        for r in rows[1:]:
            if len(r) <= max(ci_price, ci_qty, ci_amt):
                continue
            name = r[0]
            price, qty, amt = _num(r[ci_price]), _num(r[ci_qty]), _num(r[ci_amt])
            if price is None or qty is None or amt is None:
                continue
            if re.fullmatch(r"(小计|合计|总计|申请经费)", name.strip()):
                items.append({"kind": "summary", "name": name.strip(), "amt": amt})
            else:
                items.append({"kind": "item", "name": name.strip(), "price": price, "qty": qty, "amt": amt})

        line_bad = [it for it in items if it["kind"] == "item" and abs(it["price"] * it["qty"] - it["amt"]) >= 1]
        if line_bad:
            gap = sum(abs(b["price"] * b["qty"] - b["amt"]) for b in line_bad)
            findings.append({
                "code": "P1-budget-line", "level": "warn",
                "title": "表格逐行金额算错（单价×数量≠金额）",
                "evidence": "；".join(
                    "%s: %g×%g=%g 但表里写 %g" % (b["name"], b["price"], b["qty"], b["price"] * b["qty"], b["amt"])
                    for b in line_bad[:5]
                ) + ("　合计偏差 %.2f 元" % gap if len(line_bad) > 5 else ""),
                "detail": "P1：逐行错账通常不易被一眼看出，但仍是硬伤，D2 扣 2 分。"
                          "汇总数（分项之和 / 小计 / 合计）算错才是 P0，见 P0-5b。",
            })

        sums = [it for it in items if it["kind"] == "summary"]
        detail_items = [it for it in items if it["kind"] == "item"]
        if detail_items and sums:
            total = sum(it["amt"] for it in detail_items)
            for s in sums:
                if abs(total - s["amt"]) >= 1:
                    findings.append({
                        "code": "P0-5b", "level": "P0", "title": "分项之和与「%s」不符" % s["name"],
                        "evidence": "逐行相加 = %.2f，表里写 %s = %.2f，差 %.2f"
                                    % (total, s["name"], s["amt"], total - s["amt"]),
                        "detail": "预算表自相矛盾，老师一眼可见。",
                    })
                    break


def check_table_index(doc: dict, findings: list) -> None:
    """P1-1：表格首列序号必须从 1 连续递增。跳号说明行被删了没补，是没做完的痕迹。"""
    for rows in doc.get("tables", []):
        if len(rows) < 3:
            continue
        header = rows[0]
        ci = next((i for i, c in enumerate(header)
                   if c.strip() in ("序号", "编号", "No", "NO", "no")), None)
        if ci is None:
            continue
        seq = []
        for r in rows[1:]:
            if len(r) <= ci:
                continue
            v = r[ci].strip().rstrip(".、）)")
            if re.fullmatch(r"\d{1,3}", v):
                seq.append(int(v))
            elif v:
                break
        if len(seq) < 2:
            continue
        gaps = [n for i, n in enumerate(seq) if i and n != seq[i - 1] + 1]
        if gaps or seq[0] != 1:
            findings.append({
                "code": "P1-1", "level": "warn", "title": "表格序号不连续",
                "evidence": "序号序列 %s…%s，不连续处 %s"
                            % (seq[0], seq[-1],
                               "、".join(str(g) for g in gaps) or "起始不是 1"),
                "detail": "P1：序号跳号说明删过行没补编号，物料表会漏项。D1 扣 1 分。",
            })
            break


DATE_RE = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")


def check_dates(text: str, findings: list) -> None:
    # P0-6 星期核对
    for m in re.finditer(DATE_RE.pattern + r"[^。\n]{0,10}?(?:周|星期)\s*([一二三四五六日天])", text):
        try:
            d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            continue
        want = CN_WEEK.get(m.group(4))
        if want is not None and d.weekday() != want:
            findings.append({
                "code": "P0-6", "level": "P0", "title": "日期与星期不符",
                "evidence": "%s 标注为周%s，实际是周%s"
                            % (m.group(0)[:20], m.group(4), "一二三四五六日"[d.weekday()]),
                "detail": "日历核对不通过，属会被当场发现的硬伤。",
            })

    # P0-7 同一活动的"举办日期"出现两个不同值
    # 只认「紧邻举办时间小节标题 / 举办动词」的日期，向后看会误收正文里的时间描述，
    # 向前看则天然排除成文日期（封面）与落款日期（妥否，请批示之后）。
    hits: dict[str, str] = {}
    for m in re.finditer(DATE_RE.pattern, text):
        before = text[max(0, m.start() - 40): m.start()]
        if re.search(r"(举办时间|活动时间|时间安排|日程|举办|举行|开展|拟于|定于|时间\s*[:：]|日期\s*[:：])", before):
            hits.setdefault(m.group(0), before[-16:].replace("\n", " ").strip())
    if len(hits) >= 2:
        findings.append({
            "code": "P0-7", "level": "P0", "title": "活动日期全文自相矛盾",
            "evidence": "；".join("%s（%s…）" % (d, ctx) for d, ctx in hits.items()),
            "detail": "同一活动的日期必须全文一致（成文日期、落款日期不算）。",
        })


def check_ppt_pages(pages: list, findings: list) -> None:
    """两种页码形态都要抓：
    1) 「04 / 16」式，声明了总页数
    2) 裸数字页码段落（页脚里只有「04」），无法直接比总页数，但能比当前页序
    """
    total = len(pages)
    slash_bad, bare_bad, declared_totals = [], [], set()

    for p in pages:
        idx = p["index"]
        flat = re.sub(r"\s*\n\s*", " ", p["text"])
        pairs = re.findall(r"(?<![\d.])(\d{1,3})\s*/\s*(\d{1,3})(?![\d.])", flat)
        if pairs:
            for cur, tot in pairs:
                declared_totals.add(int(tot))
                if int(cur) != idx:
                    slash_bad.append("第 %d 页声明 %s/%s" % (idx, cur, tot))
            continue
        # 裸页码只在页脚位置才算：取最后一个非空行，且该行必须是纯数字。
        # （正文里的「84 小时」「07」这类数字、章节序号「02」都不能当页码）
        lines = [l.strip() for l in p["text"].split("\n") if l.strip()]
        if not lines:
            continue
        tail = lines[-1]
        m = re.fullmatch(r"(\d{1,3})", tail)
        if m and 1 <= int(m.group(1)) <= total and int(m.group(1)) != idx:
            bare_bad.append("第 %d 页声明 %s" % (idx, m.group(1)))

    if declared_totals and total not in declared_totals:
        slash_bad.append("实际 %d 页 / 声明共 %s 页"
                         % (total, "或".join(str(d) for d in sorted(declared_totals))))

    if slash_bad:
        findings.append({
            "code": "P1-page", "level": "warn", "title": "页码与实际页序不符",
            "evidence": "、".join(slash_bad[:8]),
            "detail": "P1：页码错乱 D4 扣 1 分；总页数与实际不符 D2 扣 2 分。",
        })
    if bare_bad:
        findings.append({
            "code": "P1-page-bare", "level": "warn", "title": "页脚页码与实际页序不符",
            "evidence": "、".join(bare_bad[:10]),
            "detail": "P1：页码未随插入/删除页更新，D4 扣 1 分。"
                      "这类错误放映时一眼可见，属未做交付前检查。",
        })

    flat_all = re.sub(r"\s*\n\s*", " ", "\n".join(p["text"] for p in pages))
    placeholders = re.findall(
        r"(Lorem ipsum|点击此处添加|XXXX|待补充|此处添加文本|在此输入|放入.{0,8}封面)", flat_all)
    if placeholders:
        findings.append({
            "code": "P1-8", "level": "warn", "title": "存在未替换的模板占位文字",
            "evidence": "、".join(sorted(set(placeholders))[:5]),
            "detail": "P1：D5 扣 2 分。",
        })

    # 正文完全空白的页：放映时是一张空页。
    # 只在首末页之外判定，避免把封面/结尾页误报。
    blanks = []
    for p in pages:
        lines = [l for l in p["text"].split("\n") if l.strip()]
        if p["index"] in (1, len(pages)):
            continue
        if not lines:
            hint = (p.get("notes") or "").strip()
            blanks.append("第 %d 页（%s）"
                          % (p["index"],
                             "备注写了内容但正文没放上去" if hint else "完全空白"))
    if blanks:
        findings.append({
            "code": "P1-empty", "level": "warn", "title": "存在正文空白的页",
            "evidence": "、".join(blanks[:6]),
            "detail": "P1：备注里有内容但页面上没放，说明正文页还没填。"
                      "D1 扣 2 分；若同一份超过 2 页如此，D1 扣 3 分。",
        })


def prescreen(kind: str, doc: dict | None, raw_text: str, ident: dict,
              doc_form: str = "report") -> list:
    findings: list = []
    check_club_names(raw_text, ident, findings)
    if kind == "plan":
        check_plan_elements(raw_text, findings, doc, doc_form)
        check_budget_math(doc or {}, findings)
        check_table_index(doc or {}, findings)
        check_dates(raw_text, findings)
    elif kind == "ppt":
        check_ppt_pages((doc or {}).get("pages") or [], findings)
        check_dates(raw_text, findings)
    elif kind == "poster":
        check_dates(raw_text, findings)
    # 视觉类产出（--image / --slide）：机器读不到图内文字，raw_text 为空。
    # 二维码、社团名、校徽配色、字体层级这些要素缺失，只能由视觉模型判断。
    # 这里不产出任何机器结论，避免没做 OCR 就把海报误退回；全部交给 rubric 里的
    # 「先看图再看字」条款与 veto-rules 的 P0-2e 条目。
    if kind in ("poster", "ppt") and not (doc or {}).get("pages"):
        return findings
    if not re.search(r"二维码|扫码|报名|入口", raw_text) and kind == "poster":
        findings.append({
            "code": "W-visual-verify", "level": "W", "title": "待模型核实：海报报名入口",
            "evidence": "文本层未见「二维码/扫码/报名」，但海报文字在图内，机器读不到",
            "detail": "请视觉模型确认海报上是否有二维码或报名方式；确实没有才退回。",
        })
    return findings


# ----------------------------------------------------------------------------
# 提示词渲染
# ----------------------------------------------------------------------------

def read(name: str) -> str:
    return (RUBRICS / name).read_text(encoding="utf-8")


OUTPUT_CONTRACT = """
严格只输出一个 JSON 对象，不要输出任何解释文字、不要用 markdown 代码块包裹。
结构如下：
{
  "veto": {"hit": false, "items": [{"code": "P0-x", "evidence": "原文证据", "detail": "为什么退回"}]},
  "scores": {"completeness": 1-5整数, "accuracy": 1-5整数, "executability": 1-5整数, "compliance": 1-5整数, "quality": 1-5整数},
  "evidence": {"completeness": "该维度得分依据，引用原文或说明缺什么", "accuracy": "...", "executability": "...", "compliance": "...", "quality": "..."},
  "fixes": ["修改指令1", "修改指令2", "修改指令3"],
  "summary": "一句话总评"
}

要求：
1. veto 只在命中 references/veto-rules.md 里标 P0 的条目时置 true。拿不准的一律不否决。
2. scores 必须是 1 到 5 的整数，不给 3.5、不给 0、不给 6。
3. 每个维度的分数必须能由该维度 1-5 分的描述推出，不凭整体印象给分。
4. fixes 必须恰好 3 条，每条指向具体位置并给出可执行动作，按影响从大到小排。
5. 视觉类产出（poster / ppt 单页）必须依据图片内容判断，不许只读图里的文字。
6. 海报类必须逐项核对 references/veto-rules.md 里的「海报三要素」，每项在 evidence 里写明在图的哪个位置找到的；找不到就写「未见」，不要推测给过。
7. 海报的【活动时间与地点没有格式限制】，【配色不限具体色值】——这两项不得作为否决理由，只能进 D1 或 D5 扣分。
"""


def build_prompt(kind: str, subject: str, extra: str = "") -> str:
    parts = [
        "你是高校社团的产出质量评审员。你要按既定标准给一份产出物打分，结果会换算成学分，所以必须严格、可复现——换一个人评同一份东西，结果应一致。",
        "",
        "=== 通用评分标准 ===",
        read("_common.md"),
        "",
        "=== 本类型专属标准（%s）===" % kind,
        read("%s.md" % kind),
        "",
        "=== 一票否决清单 ===",
        (REFERENCES / "veto-rules.md").read_text(encoding="utf-8"),
        "",
        "=== 输出格式 ===",
        OUTPUT_CONTRACT.strip(),
    ]
    if extra:
        parts += ["", "=== 本次评审说明 ===", extra]
    parts += ["", "=== 待评产出物 ===", subject]
    return "\n".join(parts)


ANSWER_SCHEMA = {
    "veto": {"hit": False, "items": [{"code": "P0-x", "evidence": "原文证据", "detail": "为什么退回"}]},
    "scores": {"completeness": 1, "accuracy": 1, "executability": 1, "compliance": 1, "quality": 1},
    "evidence": {"completeness": "该维度得分依据", "accuracy": "", "executability": "",
                 "compliance": "", "quality": ""},
    "fixes": ["修改指令1", "修改指令2", "修改指令3"],
    "summary": "一句话总评",
}

# ----------------------------------------------------------------------------
# 模型调用
# ----------------------------------------------------------------------------

def load_models() -> dict:
    if not MODELS_FILE.exists():
        return {}
    return json.loads(MODELS_FILE.read_text(encoding="utf-8"))


def _extract_json(text: str) -> dict | None:
    if not text:
        return None
    t = text.strip()
    fence = re.search(r"```(?:json)?\s*(.+?)```", t, re.S)
    if fence:
        t = fence.group(1).strip()
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        pass
    depth, start = 0, None
    for i, ch in enumerate(t):
        if ch == "{":
            if start is None:
                start, depth = i, 0
            depth += 1
        elif ch == "}" and start is not None:
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(t[start:i + 1])
                except json.JSONDecodeError:
                    start = None
    return None


def _pick_payload_text(payload: dict) -> str:
    """兼容答案落在 reasoning 字段、content 为空的模型。"""
    if isinstance(payload.get("choices"), list) and payload["choices"]:
        msg = payload["choices"][0].get("message") or {}
        for key in ("content", "reasoning_content", "reasoning", "thinking", "thought"):
            val = msg.get(key)
            if isinstance(val, str) and val.strip():
                if val.strip().startswith("{"):
                    return val
                got = _extract_json(val)
                if got:
                    return val
        text = payload["choices"][0].get("text")
        if isinstance(text, str) and text.strip():
            return text
    content = payload.get("content")
    if isinstance(content, str) and content.strip():
        return content
    if isinstance(content, list):
        joined = "\n".join(
            b.get("text", "") for b in content if isinstance(b, dict) and b.get("type") == "text"
        )
        if joined.strip():
            return joined
    # OpenAI responses 形态：output[].content[].text
    for item in payload.get("output") or []:
        if not isinstance(item, dict):
            continue
        for block in item.get("content") or []:
            if isinstance(block, dict) and isinstance(block.get("text"), str) and block["text"].strip():
                return block["text"]
    for key in ("reasoning", "thinking", "reasoning_content", "output_text"):
        if isinstance(payload.get(key), str) and payload[key].strip():
            return payload[key]
    return None


def call_model(role: dict, prompt: str, image_b64: str | None = None) -> str:
    key = os.environ.get(role.get("api_key_env", ""), "")
    if not key:
        raise SystemExit(
            "缺少环境变量 %s。用 --emit-prompt / --submit 走人工搬运，或先配置该环境变量。"
            % role.get("api_key_env")
        )
    base = role["base_url"].rstrip("/")
    style = role.get("api_style", "openai")
    headers = {"Content-Type": "application/json"}

    if style == "anthropic":
        endpoint = base + "/messages"
        headers["x-api-key"] = key
        headers["anthropic-version"] = "2023-06-01"
        body = {
            "model": role["model"],
            "max_tokens": role.get("max_tokens", 3000),
            "temperature": role.get("temperature", 0),
            "messages": [{
                "role": "user",
                "content": (
                    [{"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                                 "data": image_b64}},
                     {"type": "text", "text": prompt}] if image_b64
                    else prompt
                ),
            }],
        }
    else:
        endpoint = base + "/chat/completions"
        headers["Authorization"] = "Bearer " + key
        content = prompt
        if image_b64:
            content = [
                {"type": "image_url",
                 "image_url": {"url": "data:image/png;base64," + image_b64}},
                {"type": "text", "text": prompt},
            ]
        body = {
            "model": role["model"],
            "messages": [{"role": "user", "content": content}],
            "temperature": role.get("temperature", 0),
            "max_tokens": role.get("max_tokens", 3000),
        }

    req = urllib.request.Request(endpoint, data=json.dumps(body).encode("utf-8"), headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=180) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise SystemExit("模型接口返回 HTTP %s：%s" % (exc.code, exc.read().decode("utf-8", "ignore")[:400]))
    except urllib.error.URLError as exc:
        raise SystemExit("模型接口连接失败：%s" % exc)

    text = _pick_payload_text(payload)
    if not text.strip():
        raise SystemExit("模型返回为空，无法解析。原始响应键：%s" % sorted(payload.keys()))
    return text


# ----------------------------------------------------------------------------
# 评分与换算
# ----------------------------------------------------------------------------

def normalize_scores(raw: dict) -> dict:
    """把模型给的分数规整成 1-5 整数，越界就夹紧并记账。"""
    src = raw.get("scores") if isinstance(raw.get("scores"), dict) else raw
    out, notes = {}, []
    for key, cn, code in DIMENSIONS:
        val = src.get(key)
        if isinstance(val, str):          # 模型常把分数写成 "4" 或 "4分"
            digits = re.search(r"-?\d+(?:\.\d+)?", val)
            if digits:
                val = float(digits.group(0))
        if isinstance(val, (int, float)):
            ival = int(round(float(val)))
        else:
            alt = next((v for k, v in src.items() if k in (cn, code)), None)
            if isinstance(alt, str):
                digits = re.search(r"-?\d+(?:\.\d+)?", alt)
                if digits:
                    alt = float(digits.group(0))
            ival = int(round(float(alt))) if isinstance(alt, (int, float)) else None
        if ival is None:
            notes.append("维度 %s 缺失，按 3 分计" % cn)
            ival = 3
        ival = max(1, min(5, ival))
        if isinstance(val, (int, float)) and (val < 1 or val > 5):
            notes.append("维度 %s 给出越界分 %s，已夹紧为 %d" % (cn, val, ival))
        out[key] = ival
    out["_notes"] = notes
    return out


def compute_tier(scores: dict) -> tuple:
    vals = [scores[k] for k, _, _ in DIMENSIONS]
    mean = sum(vals) / len(vals)
    low = min(vals)
    if low <= 1:
        return "C", round(mean, 2), low
    if mean >= 4.0 and low >= 3:
        return "A", round(mean, 2), low
    if mean < 2.5:
        return "C", round(mean, 2), low
    return "B", round(mean, 2), low


def normalize_veto(raw: dict, kind: str) -> tuple:
    """把模型给的否决判断规整，并强制执行「时间地点/配色不得作为否决理由」。

    返回 (hit, items, notes)。

    m00695 明确了海报的活动时间与地点没有格式限制、配色不限具体色值。
    评审模型很容易按旧版标准把 P0-2f（缺地点）/ P0-2g（时间只写日期）/ P0-2c（配色）
    当成否决理由，所以在这里兜底降级：这些 code 只能扣分，不能否决。
    """
    v = (raw or {}).get("veto")
    if not isinstance(v, dict):
        return False, [], []
    items = [i for i in (v.get("items") or []) if isinstance(i, dict)]
    notes = []
    downgraded = []
    if kind == "poster":
        for it in items:
            code = str(it.get("code") or "")
            # P0-2f 缺地点 / P0-2g 时间只写日期 —— 已降级为 D1 扣分
            if code in ("P0-2f", "P0-2g"):
                it["level"] = "downgraded"
                it["detail"] = (it.get("detail") or "") + \
                    "｜负责人 m00695 明确活动时间与地点无格式限制，此项不否决，改为 D1 扣 1 分。"
                downgraded.append(code)
            # P0-2c 徽章违规里，配色子项同样不否决；徽章位置/拉伸变形仍然否决
            elif code == "P0-2c" and _is_color_only_veto(it):
                it["level"] = "downgraded"
                it["detail"] = (it.get("detail") or "") + \
                    "｜m00695 明确配色不限具体色值、深色系即可，此项不否决，改为 D5 扣分。"
                downgraded.append(code)
    if downgraded:
        notes.append("按 m00695 降级 %d 项否决为扣分：%s"
                     % (len(downgraded), "、".join(sorted(set(downgraded)))))
    # 给每个条目一个显式 level：live 记 P0、降级记 downgraded，缺键的按 P0 处理。
    for it in items:
        it.setdefault("level", "P0")
    live = [i for i in items if i.get("level") != "downgraded"]
    hit = bool(v.get("hit")) and bool(live)
    if v.get("hit") and not hit:
        notes.append("模型原判 veto=true，但命中项全部被降级，最终改为不否决")
    return hit, items, notes


def _is_color_only_veto(item: dict) -> bool:
    """P0-2c 若只因配色被否决（证据里没有徽章问题），按 m00695 降级。"""
    blob = " ".join(str(item.get(k) or "") for k in ("evidence", "detail", "title"))
    color_words = ("配色", "主色", "色值", "颜色", "底色", "强调色", "VI")
    badge_words = ("团徽", "校徽", "徽章", "拉伸", "变形", "改色", "描边", "阴影")
    if any(w in blob for w in badge_words):
        return False   # 徽章本身有问题，仍否决
    return any(w in blob for w in color_words)


def crosscheck(primary: dict, second: dict, threshold: int) -> dict:
    a = normalize_scores(primary)
    b = normalize_scores(second)
    # 复核方没给 scores（或只给了部分维度）时无法逐维比对，直接交人工，不硬凑数字。
    if not isinstance((second or {}).get("scores"), dict) and not any(
            k in (second or {}) for k, _, _ in DIMENSIONS):
        return {
            "primary": a, "second": b, "diff": {}, "max_abs_diff": 0,
            "tier_primary": compute_tier(a)[0], "tier_second": None,
            "needs_human": True,
            "issues": ["复核方未给出维度分数，无法逐维比对，需人工复核。"],
        }
    diffs = {k: a[k] - b[k] for k, _, _ in DIMENSIONS}
    worst = max(abs(v) for v in diffs.values())
    ta, _, _ = compute_tier(a)
    tb, _, _ = compute_tier(b)
    # 主评方已经过 normalize_veto（降级过 P0-2f/g 与配色类否决），优先用降级后的结论，
    # 否则复核方会跟一个模型原始的、超纲的 veto=true 对比，误报「否决判断不一致」。
    pv = (primary.get("_veto_hit") if isinstance(primary.get("_veto_hit"), bool)
          else bool((primary.get("veto") or {}).get("hit")))
    sv = (second.get("_veto_hit") if isinstance(second.get("_veto_hit"), bool)
          else bool((second.get("veto") or {}).get("hit")))
    issues = []
    if worst > threshold:
        issues.append("分维最大差值 %d 超过阈值 %d：%s"
                      % (worst, threshold, "、".join(
                          "%s %d vs %d" % (cn, a[k], b[k]) for k, _, cn in DIMENSIONS if abs(diffs[k]) > 0)))
    if pv != sv:
        issues.append("否决判断不一致：主评 %s，复核 %s" % (pv, sv))
    return {
        "primary": a, "second": b, "diff": diffs, "max_abs_diff": worst,
        "tier_primary": ta, "tier_second": tb, "needs_human": bool(issues), "issues": issues,
    }


def resolve_base(kind: str, base: float | None, slide: bool, contest_award: bool) -> float:
    """返回实际用的基础分值：--base 覆盖，否则查表。报告里要显示这个真值。"""
    if base is not None:
        return base
    if kind == "contest":
        return BASE_CREDIT[("contest", "award" if contest_award else "entry")]
    return BASE_CREDIT[(kind, "slide" if (kind == "ppt" and slide) else "deck" if kind == "ppt" else None)]


def credit(kind: str, tier: str, base: float | None, slide: bool, contest_award: bool) -> float:
    if tier == "退回":
        return 0.0
    return round(resolve_base(kind, base, slide, contest_award) * TIER_COEF[tier], 4)


# ----------------------------------------------------------------------------
# 报告输出
# ----------------------------------------------------------------------------

def print_report(kind: str, subject_name: str, findings: list, result: dict,
                 tier: str, mean: float, low: int, base: float, got: float,
                 cc: dict | None, template_only: bool) -> None:
    print("=" * 62)
    print("社团产出质量评分  类型=%s  对象=%s" % (kind, subject_name))
    print("=" * 62)
    if template_only:
        print("\n[模式] --emit-prompt：以下提示词复制给任意模型，模型输出存成 JSON 后用 --submit 导入。")

    p0 = [f for f in findings if f["level"] == "P0"]
    warn = [f for f in findings if f["level"] != "P0"]
    print("\n一票否决预筛：%d 项 P0，%d 项提示" % (len(p0), len(warn)))
    if not p0 and not warn:
        print("  机器可检项全部通过。")
    for f in p0:
        print("  [P0] %s（%s）%s" % (f["title"], f["code"], f["evidence"]))
    for f in warn:
        print("  [--] %s（%s）%s" % (f["title"], f["code"], f["evidence"]))

    model_veto = result.get("_veto_hit") if result else None
    if result and result.get("veto"):
        items = result["veto"].get("items") or []
        live = [it for it in items if it.get("level") != "downgraded"]
        if live:
            print("\n模型判定的否决项：")
            for it in live:
                print("  [P0] %s：%s" % (it.get("code", "?"), it.get("evidence", it.get("detail", ""))))
        for it in items:
            if it.get("level") == "downgraded":
                print("  [--] %s：%s（按 m00695 不否决）" % (it.get("code", "?"), it.get("evidence", "")))
        for n in (result.get("_veto_notes") or []):
            print("  [!] %s" % n)

    if tier and not p0 and model_veto is not True:
        print("\n维度得分：")
        scores = result.get("_scores", {})
        for key, cn, code in DIMENSIONS:
            print("  %s %s：%s 分" % (code, cn, scores.get(key, "-")))
        print("  均分 %.2f   最低维度 %d" % (mean, low))
        print("\n三条修改指令：")
        for i, fx in enumerate((result.get("fixes") or [])[:3], 1):
            print("  %d) %s" % (i, fx))
        if result.get("summary"):
            print("\n总评：%s" % result["summary"])

    print("\n" + "-" * 62)
    final_tier = "退回" if (p0 or model_veto is True) else tier
    if not final_tier:
        print("未定档（仅预筛，尚未经模型评分）。")
        print("=" * 62)
        return
    print("总评档位：%s     质量系数：%s" % (final_tier, TIER_COEF[final_tier]))
    print("基础分值：%.2f     实得学分：%.2f" % (base, 0.0 if final_tier == "退回" else got))

    if cc:
        print("\n跨族复核（%s vs %s）：" % (cc["tier_primary"] and "主评", "复核"))
        print("  分维最大差值：%d" % cc["max_abs_diff"])
        if cc["needs_human"]:
            print("  ** 标记 needs_human，需人工复核 **")
            for it in cc["issues"]:
                print("     - %s" % it)
            print("  学分暂按 B 档（系数 1.0）挂账：%.2f" % round(base * 1.0, 4))
        else:
            print("  复核通过，分歧在阈值内。")
    print("=" * 62)


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------

AUTO_KIND = {
    ".docx": "plan", ".doc": "plan", ".md": "copy", ".txt": "copy",
    ".pptx": "ppt", ".ppt": "ppt", ".pdf": "plan",
    ".jpg": "poster", ".jpeg": "poster", ".png": "poster", ".webp": "poster",
}

# 报批公文的强标记：出现任一个才按 report（报批策划案）判，否则按 scheme（执行方案）。
REPORT_MARKERS = ("请示", "妥否", "请批示", "签批", "批复", "呈报", "报告如下",
                  "主送", "指导教师签", "指导老师签", "此致敬礼")


def detect_kind(path: Path) -> str:
    """按扩展名猜类型，猜不到就报错让人用 --type 指定。"""
    return AUTO_KIND.get(path.suffix.lower(), "")


def detect_doc_form(text: str) -> tuple:
    """判断这是报批策划案还是执行方案。返回 (form, 判据说明)。

    判错方向不对称：scheme 当成 report 判会让好文档被退回（白挨罚），
    report 当成 scheme 判会让该退的溜过去。所以只认强标记，宁可判 scheme。
    """
    for m in REPORT_MARKERS:
        if m in text:
            return "report", "命中报批公文标记「%s」" % m
    return "scheme", "未见请示/妥否/签字位等报批体例标记，按执行方案判"


def load_subject(args) -> tuple:
    """返回 (kind, 展示名, 原始文本, doc 或 None, 图片 base64 或 None, slide)"""
    kind = args.type
    slide = args.slide

    if args.image:
        p = Path(args.image)
        data = base64.b64encode(p.read_bytes()).decode("ascii")
        return kind, p.name, "[图片产出物，见附件图片]", None, data, slide

    if args.text:
        return kind, "(粘贴文本)", args.text, None, None, slide

    if not args.file:
        raise SystemExit("需要 --file / --text / --image 之一。")

    p = Path(args.file)
    if not p.exists():
        raise SystemExit("文件不存在：%s" % p)
    suffix = p.suffix.lower()

    # 位置参数拖进来的是图片时，按图片处理，不必再手写 --image。
    if suffix in (".jpg", ".jpeg", ".png", ".webp"):
        data = base64.b64encode(p.read_bytes()).decode("ascii")
        return kind, p.name, "[图片产出物，见附件图片]", None, data, slide

    if suffix == ".docx":
        doc = parse_docx(p)
        return kind, p.name, doc["text"], doc, None, slide
    if suffix == ".pptx":
        doc = parse_pptx(p)
        doc["page_count"] = len(doc["pages"])
        text = render_ppt_text(doc, slide)
        if not slide:
            text += "\n\n[页码声明]\n" + (render_ppt_layout(doc) or "（无页码声明）")
        return kind, p.name, text, doc, None, slide
    if suffix in (".md", ".txt"):
        return kind, p.name, p.read_text(encoding="utf-8", errors="ignore"), None, None, slide
    if suffix == ".pdf":
        doc = parse_pdf(p)
        text = doc["text"]
        if slide:
            text = render_ppt_text(doc, slide)
        return kind, p.name, text, doc, None, slide
    raise SystemExit("不支持的类型：%s" % suffix)


BATCH_SUFFIX = (".docx", ".pptx", ".pdf", ".md", ".txt",
                ".jpg", ".jpeg", ".png", ".webp")


def run_batch(args) -> int:
    """批量预筛一个目录：一次看完全部成员交上来的东西。

    只跑机器可检的那部分，不调模型——所以一次也不花钱、不需要 API key。
    目的是让部长先扫一眼谁有硬伤，再决定哪几份值得送模型细评。
    """
    root = Path(args.batch)
    if not root.is_dir():
        raise SystemExit("不是目录：%s" % root)
    ident = load_identity()
    rows = []
    for p in sorted(root.rglob("*")):
        if p.suffix.lower() not in BATCH_SUFFIX or p.name.startswith("~$"):
            continue
        kind = args.type or detect_kind(p) or "copy"
        sub = argparse.Namespace(type=kind, file=str(p), text=None,
                                 image=None, slide=None)
        try:
            _, name, subject, doc, _, slide = load_subject(sub)
            doc_form = args.doc_form or (
                detect_doc_form(subject)[0] if kind == "plan" else "report")
            findings = prescreen(kind, doc, subject, ident, doc_form)
        except SystemExit as e:
            print("[跳过] %s：%s" % (p.name, e))
            continue
        p0 = [f for f in findings if f["level"] == "P0"]
        warn = [f for f in findings if f["level"] != "P0"]
        rows.append((p.name, kind, doc_form, p0, warn))

    if not rows:
        print("目录里没有可评的产出物：%s" % root)
        return 0

    print("=" * 78)
    print("批量预筛：%s（共 %d 份）" % (root, len(rows)))
    print("=" * 78)
    print("%-34s %-8s %-7s %-6s %s" % ("文件", "类型", "形态", "P0", "首要问题"))
    print("-" * 78)
    vetoed = 0
    for name, kind, form, p0, warn in rows:
        if p0:
            vetoed += 1
        first = "；".join(f["title"] for f in p0[:2]) or \
               "；".join(f["title"] for f in warn[:2]) or "机器可检项全部通过"
        print("%-34s %-8s %-7s %-6d %s"
              % (name[:33], kind, form, len(p0), first[:44]))
    print("-" * 78)
    print("命中 P0（该退回）%d 份，其余 %d 份机器可检项通过。"
          % (vetoed, len(rows) - vetoed))
    print("提示：本步只跑机器能查的错别字/日期/预算/页码等。五个维度的打分与档位")
    print("     仍需送模型评审，用 --emit-prompt 导出提示词。")
    print("=" * 78)
    return 1 if vetoed else 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description="社团产出质量评分",
        epilog="最简用法：review.py 你的策划案.docx（类型自动识别）"
               "　|　review.py 你的海报.jpg　|　review.py 结果.json")
    ap.add_argument("path", nargs="?",
                    help="产出物文件（docx/pptx/pdf/jpg/png/md/txt），"
                         "或模型返回的结果 JSON，或直接把文件拖进来")
    ap.add_argument("--type", choices=["plan", "ppt", "copy", "poster", "contest"],
                    help="类型，默认按扩展名自动识别")
    ap.add_argument("--file", help="产出物文件，等同位置参数")
    ap.add_argument("--text", help="直接传入文案正文")
    ap.add_argument("--image")
    ap.add_argument("--slide", type=int, help="PPT 单页评审，1 起")
    ap.add_argument("--doc-form", choices=["report", "scheme"],
                    help="plan 类文档形态：report=报批策划案（要请示语/签字位/"
                         "落款/安全预案）；scheme=执行方案（不送审，P0-2 系列与 P0-3 "
                         "降为提示）。默认按文档内容自动判断")
    ap.add_argument("--base", type=float, help="覆盖基础分值")
    ap.add_argument("--contest-award", action="store_true", help="竞赛材料按获奖档算基础分")
    ap.add_argument("--prescreen-only", action="store_true", help="只跑机器可检的一票否决")
    ap.add_argument("--emit-prompt", action="store_true", help="打印提示词，不调模型")
    ap.add_argument("--submit", help="导入模型返回的 JSON 文件")
    ap.add_argument("--no-crosscheck", action="store_true", help="不跑跨族复核")
    ap.add_argument("--json", help="把完整结果写入该 JSON 文件")
    ap.add_argument("--batch", metavar="目录",
                    help="批量预筛一个目录下的所有产出物，出一张对照表")
    args = ap.parse_args()

    if args.batch:
        return run_batch(args)

    # 位置参数是最短路径：直接给文件就完事，类型和文档形态都不用写。
    if args.path and not (args.file or args.text or args.image or args.submit):
        p = Path(args.path)
        # 结果 JSON 也能拖进来：识别成导入模式，不用切 --submit。
        if p.suffix.lower() == ".json":
            args.submit = str(p)
        else:
            args.file = str(p)

    if not args.type:
        args.type = detect_kind(Path(args.file)) if args.file else ""
        if args.image:
            args.type = detect_kind(Path(args.image)) or "poster"
        if not args.type:
            args.type = "copy" if args.text else ""
        if not args.type:
            raise SystemExit("看不出这是哪类产出物，请用 --type 指定："
                             "plan（策划案）ppt（PPT）copy（文案）"
                             "poster（海报）contest（竞赛材料）")
        print("[自动识别] 类型=%s" % args.type)

    ident = load_identity()
    if args.submit:
        result = json.loads(Path(args.submit).read_text(encoding="utf-8"))
        result = _extract_json(result) if isinstance(result, str) else result
        if result is None:
            raise SystemExit("无法从 %s 解析出 JSON" % args.submit)
        veto_hit, veto_items, veto_notes = normalize_veto(result, args.type)
        result["veto"]["items"] = veto_items
        result["_veto_hit"] = veto_hit
        result["_veto_notes"] = veto_notes
        scores = normalize_scores(result)
        result["_scores"] = scores
        tier, mean, low = compute_tier(scores)
        if veto_hit:
            tier = "退回"
        base = args.base
        real_base = resolve_base(args.type, base, bool(args.slide), args.contest_award)
        got = credit(args.type, tier, base, bool(args.slide), args.contest_award)
        print_report(args.type, args.submit, [], result, tier, mean, low, real_base, got,
                     None, template_only=False)
        if args.json:
            Path(args.json).write_text(json.dumps(
                {"type": args.type, "tier": tier, "mean": mean, "min": low,
                 "credit": got, "result": result}, ensure_ascii=False, indent=2), encoding="utf-8")
        return 0

    kind, name, subject, doc, image_b64, slide = load_subject(args)

    doc_form = args.doc_form
    if not doc_form and kind == "plan":
        doc_form, why = detect_doc_form(subject)
        print("[自动识别] 文档形态=%s（%s）" % (doc_form, why))
    doc_form = doc_form or "report"

    findings = prescreen(kind, doc, subject, ident, doc_form)
    p0 = [f for f in findings if f["level"] == "P0"]

    if args.prescreen_only:
        print_report(kind, name, findings, {}, None, 0, 0, 0, 0, None, template_only=True)
        return 1 if p0 else 0

    if args.emit_prompt:
        extra = []
        if slide:
            extra.append("本次只评审第 %d 页，按单页标准打分。" % slide)
        if kind == "poster" and not slide:
            extra.append(
                "这是海报评审，必须逐项核对海报三要素，缺任一项 veto.hit=true 并填 P0-2：\n"
                "  P0-2e 团徽+校徽都在左上或右上角，且团徽在左、校徽在右（放反也是违规）\n"
                "  P0-2i 有真实可扫的群聊二维码（占位图/示例图算缺）\n"
                "  P0-2h 有主办单位（如「共青团西北大学委员会」，协办不能替代主办）\n"
                "每项都要在 evidence 里写明「在图的哪个位置找到的」；找不到就写「未见」，不要推测给过。\n"
                "注意：活动时间与活动地点【没有格式限制】，只写日期或写得笼统都不否决，"
                "只在 D1 记 1 分差异；配色【不限具体色值】，深色系即可，不要因为配色否决。"
            )
        if doc_form == "scheme":
            extra.append(
                "这份是执行方案（--doc-form scheme），不是报批策划案。"
                "不要因为缺「妥否，请批示」「指导教师签字」「落款日期」「主办单位」而否决，"
                "这些是报批公文体例，不适用于执行方案；但缺风险章节仍是 P0。"
            )
        prompt = build_prompt(kind, subject, "\n".join(extra))
        print(prompt)
        print("\n\n---- 以下为预筛结果，供你对照 ----")
        print_report(kind, name, findings, {}, None, 0, 0, 0, 0, None, template_only=True)
        if args.json:
            Path(args.json).write_text(json.dumps(
                {"type": kind, "prompt": prompt, "prescreen": findings,
                 "answer_schema": ANSWER_SCHEMA}, ensure_ascii=False, indent=2),
                encoding="utf-8")
            print("\n[已写出] 提示词与答案结构 → %s" % args.json)
        return 1 if p0 else 0

    models = load_models()
    route = models.get("路由规则", {}).get(kind)
    if not route:
        raise SystemExit("lib/models.json 缺少 %s 的路由配置，用 --emit-prompt 走人工搬运。" % kind)

    primary_role = models["roles"][route["primary"]]
    cross_role = models["roles"].get(route.get("crosscheck", ""))
    if cross_role and cross_role.get("family") == primary_role.get("family"):
        print("[警告] 主评与复核同族（%s），同源偏差无法规避。" % primary_role.get("family"))
    for role in (primary_role, cross_role):
        if role and role.get("family") and role.get("api_style") is None:
            role.setdefault("api_style", "openai")

    extra = ["本次评审说明：%s" % name]
    if slide:
        extra.append("只评审第 %d 页。" % slide)
    if p0:
        extra.append("机器预筛已命中 P0：%s。若你同意请在 veto.hit 置 true，并在 evidence 中补充模型视角的理由。"
                     % "；".join(f["title"] for f in p0))
    prompt = build_prompt(kind, subject, "\n".join(extra))

    print("[调用] 主评 %s/%s" % (primary_role.get("family"), primary_role["model"]))
    raw_primary = call_model(primary_role, prompt, image_b64)
    result = _extract_json(raw_primary)
    if result is None:
        print("[原始响应]\n%s" % raw_primary[:2000])
        raise SystemExit("主评模型输出无法解析为 JSON。")

    veto_hit, veto_items, veto_notes = normalize_veto(result, kind)
    if result.get("veto"):
        result["veto"]["items"] = veto_items
    result["_veto_hit"] = veto_hit
    result["_veto_notes"] = veto_notes
    scores = normalize_scores(result)
    result["_scores"] = scores
    tier, mean, low = compute_tier(scores)
    if veto_hit or p0:
        tier = "退回"

    cc = None
    if cross_role and not args.no_crosscheck:
        print("[调用] 复核 %s/%s" % (cross_role.get("family"), cross_role["model"]))
        raw_second = call_model(cross_role, prompt, image_b64)
        second = _extract_json(raw_second)
        if second is None:
            print("[警告] 复核模型输出无法解析，跳过复核。")
        else:
            cc = crosscheck(result, second, models.get("分异阈值", {}).get("分维最大差值", 1))

    base = args.base
    final_tier = "退回" if tier == "退回" else ("B" if (cc and cc["needs_human"]) else tier)
    got = credit(kind, final_tier, base, bool(slide), args.contest_award)

    print_report(kind, name, findings, result, tier, mean, low,
                 resolve_base(kind, base, bool(slide), args.contest_award),
                 got, cc, template_only=False)

    if args.json:
        Path(args.json).write_text(json.dumps({
            "type": kind, "object": name, "slide": slide,
            "prescreen": findings, "result": result,
            "scores": scores, "tier": tier, "final_tier": final_tier,
            "mean": mean, "min_dimension": low,
            "crosscheck": cc, "credit": got,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[已写出] %s" % args.json)
    return 1 if final_tier == "退回" else 0


if __name__ == "__main__":
    sys.exit(main())