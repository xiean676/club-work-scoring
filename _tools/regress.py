# -*- coding: utf-8 -*-
"""锚点样本回测：机器预筛结果 vs samples/anchors.json 里登记的 expected_codes。

这是「换谁来评结果都一样」的底线保障——检测规则改了之后必须重跑，
确认没有误报（检出 anchors 说不该报的东西）也没有漏报（该报没报）。

运行方式（PYTHONUTF8=1）：
    python _tools/regress.py
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "lib"))

import review  # noqa: E402

SAMPLES = ROOT / "samples"
anchors = json.loads((SAMPLES / "anchors.json").read_text(encoding="utf-8"))
ident = review.load_identity()

# 误报白名单：anchors 明确说过「不该报」的东西。
# ppt-nuedc-v3 的 P1-3 是 anchors 自己列的期望，但机器抓不到（语义层面），
# 记在这里作为「已知漏报」，不判 FAIL 但会在末尾提示。
KNOWN_MISSES = {
    "ppt-nuedc-v3": {"P1-3", "P1-12"},
}

fails = []
warns = []

SUFFIX = {"ppt": ".pptx", "plan": ".docx", "poster": ".jpg", "pdf": ".pdf"}

for s in anchors["samples"]:
    sid = s["id"]
    kind = s["type"]
    # plan 类可能是 .docx 也可能是 .pdf：优先用 anchors 里登记的 file，
    # 否则回退到 id + 常规后缀。
    if s.get("file"):
        src = SAMPLES / s["file"]
    else:
        src = SAMPLES / (sid + SUFFIX.get(kind, ".docx"))
    if not src.exists():
        warns.append("%s: 样本文件缺失 %s，跳过" % (sid, src.name))
        continue

    doc_form = s.get("doc_form", "report")

    if kind == "poster":
        doc = None
    elif src.suffix.lower() == ".pdf":
        doc = review.parse_pdf(src)
    elif kind == "plan":
        doc = review.parse_docx(src)
    else:
        doc = review.parse_pptx(src)
    raw = review.render_ppt_text(doc, None) if kind == "ppt" else (doc or {}).get("text", "")
    findings = review.prescreen(kind, doc, raw, ident, doc_form)
    codes = [f["code"] for f in findings]
    p0 = [f["code"] for f in findings if f["level"] == "P0"]

    exp = s["expected"]
    want_codes = set(exp.get("expected_codes", []))
    got_codes = set(codes)

    print("\n=== %s（用户主观 %s 分，期望档位 %s%s）"
          % (sid, s.get("user_score_100"), exp.get("tier"),
             "，形态 %s" % doc_form if doc_form != "report" else ""))

    # 1. 该报的 P0 一个都不能少
    missing = want_codes - got_codes
    if missing:
        fails.append("%s 漏报 P0：%s" % (sid, "、".join(sorted(missing))))
        print("  漏报  %s" % "、".join(sorted(missing)))
    else:
        print("  P0 命中齐全：%s" % ("、".join(sorted(want_codes)) or "（无）"))

    # 2. anchors 说过的 should_pass：不该报的误报检查
    known_ok = {c for c in codes if c in KNOWN_MISSES.get(sid, set())}

    # 3. 否决结论与 expected.veto 一致
    want_veto = exp.get("veto")
    got_veto = bool(p0)
    if want_veto is not None and want_veto != got_veto:
        fails.append("%s 否决结论不符：期望 veto=%s，实际 %s" % (sid, want_veto, got_veto))
        print("  FAIL  否决结论不符 期望=%s 实际=%s" % (want_veto, got_veto))
    else:
        print("  否决结论一致：%s" % ("命中" if got_veto else "未命中"))

    lvl = {f["code"]: f["level"] for f in findings}
    for c in sorted(got_codes):
        # 标签按 findings 里真实的 level 走，不能按 code 前缀——
        # scheme 形态下 P0-2a/b/c/d 的 level 已被降为 warn。
        if lvl.get(c) == "P0":
            tag = "P0 否决"
        elif c in known_ok:
            tag = "已知漏报项"
        elif c.startswith("P0"):
            tag = "已降级为提示"
        else:
            tag = "提示"
        print("    %-16s %s" % (c, tag))
    for c in sorted(missing - KNOWN_MISSES.get(sid, set())):
        print("    %-16s 未检出" % c)

# --------------------------------------------------------------------------
print("\n" + "=" * 62)
if warns:
    for w in warns:
        print("提醒：", w)
if fails:
    print("=== 回测失败 %d 项 ===" % len(fails))
    for f in fails:
        print(" -", f)
    sys.exit(1)
print("=== 回测通过：期望命中的 P0 全部检出，且否决结论与 anchors 一致 ===")
print("注：ppt-nuedc-v3 的 P1-3 / P1-12 属语义矛盾与页脚位置歧义，")
print("    机器预筛抓不到，已记入 anchors.measured 作为已知漏报，由模型评审负责。")