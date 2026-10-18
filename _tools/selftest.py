# -*- coding: utf-8 -*-
"""导入链路自测：覆盖 JSON 解析的各种脏格式、档位计算、学分换算、跨族复核分歧。

运行方式（PYTHONUTF8=1）：
    python _tools/selftest.py
"""
import json
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "lib"))

import review  # noqa: E402

PY = sys.executable
FAILS = []


def check(name, got, want):
    if got != want:
        FAILS.append("%s: got %r, want %r" % (name, got, want))
        print("  FAIL %-28s got=%r want=%r" % (name, got, want))
    else:
        print("  ok   %-28s %r" % (name, got))


# --------------------------------------------------------------------------
# 1. JSON 解析：脏格式容忍
# --------------------------------------------------------------------------
print("[1] _extract_json 脏格式容忍")
CLEAN = {"scores": {"completeness": 4, "accuracy": 4, "executability": 4,
                    "compliance": 5, "quality": 4}, "fixes": ["a", "b", "c"]}
check("裸 JSON", review._extract_json(json.dumps(CLEAN, ensure_ascii=False)) is not None, True)
check("markdown 包裹", review._extract_json(
    "```json\n" + json.dumps(CLEAN, ensure_ascii=False) + "\n```") is not None, True)
check("无语言标记的围栏", review._extract_json(
    "```\n" + json.dumps(CLEAN, ensure_ascii=False) + "\n```") is not None, True)
check("前面带解释文字", review._extract_json(
    "好的，我的评审结果如下：\n" + json.dumps(CLEAN, ensure_ascii=False)) is not None, True)
check("纯废话无 JSON", review._extract_json("我无法评审这个文件。"), None)

# 关键坑：有些模型 content 为空，答案落在 reasoning 字段
print("[2] _pick_payload_text：content 空、答案在 reasoning")
PAYLOAD_OK = {"choices": [{"message": {"role": "assistant", "content": json.dumps(CLEAN, ensure_ascii=False)}}]}
PAYLOAD_REASONING = {
    "choices": [{"message": {"reasoning": "思考过程……"}}],   # content 缺失
    "reasoning_content": "思考过程……",
}
PAYLOAD_BOTH = {"choices": [{"message": {
    "reasoning": "这是思考不是答案 [{\"bad\": 1}]",
    "content": json.dumps(CLEAN, ensure_ascii=False)}}]}
# content 为空串，答案在 reasoning
PAYLOAD_EMPTY_CONTENT = {"choices": [{"message": {
    "reasoning_content": "推理：" + json.dumps(CLEAN, ensure_ascii=False),
    "content": ""}}]}

t1 = review._pick_payload_text(PAYLOAD_OK)
check("正常 content", t1 is not None and t1.strip() != "", True)
t2 = review._pick_payload_text(PAYLOAD_REASONING)
check("仅 reasoning_content 有内容也能取到", t2 is not None, True)
t3 = review._pick_payload_text(PAYLOAD_BOTH)
check("content 优先于 reasoning", json.loads(t3)["scores"]["completeness"], 4)
t4 = review._pick_payload_text(PAYLOAD_EMPTY_CONTENT)
check("content 空串回落到 reasoning", t4 is not None and "completeness" in t4, True)
# 顶层字段形态（部分网关把 content 放顶层）
check("顶层 content 形态", review._pick_payload_text({"content": json.dumps(CLEAN)}) is not None, True)
check("OpenAI responses 形态", review._pick_payload_text(
    {"output": [{"content": [{"text": json.dumps(CLEAN)}]}]}) is not None, True)
check("完全空 payload", review._pick_payload_text({}), None)

# --------------------------------------------------------------------------
# 3. 分数归一：脏分数
# --------------------------------------------------------------------------
print("[3] normalize_scores 容忍脏分数")
check("标准", review.normalize_scores(CLEAN)["accuracy"], 4)
check("3.5 截断为 4", review.normalize_scores(
    {"scores": {"accuracy": 3.5}})["accuracy"], 4)
check("7 夹到 5", review.normalize_scores({"scores": {"accuracy": 7}})["accuracy"], 5)
check("0 抬到 1", review.normalize_scores({"scores": {"accuracy": 0}})["accuracy"], 1)
check("\"4\" 字符串转 int", review.normalize_scores(
    {"scores": {"accuracy": "4"}})["accuracy"], 4)
check("缺失维度补 3 并记账", review.normalize_scores({"scores": {}})["accuracy"], 3)
check("带 _notes 字段不串进维度", "_notes" in review.normalize_scores(CLEAN), True)

# --------------------------------------------------------------------------
# 4. 档位规则
# --------------------------------------------------------------------------
print("[4] compute_tier 档位规则")
S = lambda a, b, c, d, e: {"completeness": a, "accuracy": b, "executability": c,
                           "compliance": d, "quality": e}
check("全 5 → A", review.compute_tier(S(5, 5, 5, 5, 5))[0], "A")
check("均分 4.2 最低 3 → A", review.compute_tier(S(4, 5, 4, 4, 4))[0], "A")
check("均分 4.0 最低 2 → 不给 A", review.compute_tier(S(5, 5, 4, 2, 4))[0], "B")
check("全 3 → B", review.compute_tier(S(3, 3, 3, 3, 3))[0], "B")
check("均分 2.4 → C", review.compute_tier(S(2, 2, 2, 3, 3))[0], "C")
check("某维 1 分 → C", review.compute_tier(S(1, 5, 5, 5, 5))[0], "C")

# --------------------------------------------------------------------------
# 5. 学分换算
# --------------------------------------------------------------------------
print("[5] credit 学分换算（策划案基础 0.3）")
check("A 档 ×1.5", review.credit("plan", "A", None, False, False), 0.45)
check("B 档 ×1.0", review.credit("plan", "B", None, False, False), 0.3)
check("C 档 ×0.5", review.credit("plan", "C", None, False, False), 0.15)
check("退回 = 0", review.credit("plan", "退回", None, False, False), 0.0)
check("--base 覆盖基础分值", review.credit("plan", "A", 0.4, False, False), 0.6)
check("PPT 单页", review.credit("ppt", "A", None, True, False), 0.075)
check("PPT 整套", review.credit("ppt", "A", None, False, False), 0.3)
check("竞赛获奖材料 A（0.6×1.5）", review.credit("contest", "A", None, False, True), 0.9)
check("竞赛参赛材料 A（0.3×1.5）", review.credit("contest", "A", None, False, False), 0.45)
check("单件最高 0.9，不破 1 分", review.credit("contest", "A", None, False, True), 0.9)

# --------------------------------------------------------------------------
# 6. 跨族复核：分歧 > 1 标人工
# --------------------------------------------------------------------------
print("[6] crosscheck 分歧标记")
A = {"scores": S(5, 5, 5, 5, 5)}
B = {"scores": S(5, 5, 5, 5, 5)}
check("完全一致 → 无需人工", review.crosscheck(A, B, 1)["needs_human"], False)
C = {"scores": S(2, 5, 5, 5, 5)}
check("差 3 分 → 需人工", review.crosscheck(A, C, 1)["needs_human"], True)
check("差 1 分 → 不需人工", review.crosscheck(
    A, {"scores": S(4, 5, 5, 5, 5)}, 1)["needs_human"], False)
check("恰好差 2 分 → 需人工", review.crosscheck(
    A, {"scores": S(3, 5, 5, 5, 5)}, 1)["needs_human"], True)
cc = review.crosscheck(A, C, 1)
check("分歧明细记录维度", max(cc["diff"], key=lambda k: abs(cc["diff"][k])), "completeness")
check("记录两侧分数", cc["primary"]["completeness"], 5)
check("复核缺 scores 时转人工", review.crosscheck(A, {}, 1)["needs_human"], True)
check("否决判断不一致也需人工", review.crosscheck(
    {"scores": S(4, 4, 4, 4, 4), "veto": {"hit": True}},
    {"scores": S(4, 4, 4, 4, 4), "veto": {"hit": False}}, 1)["needs_human"], True)

# --------------------------------------------------------------------------
# 7. plan 两种文档形态（--doc-form report / scheme）
# --------------------------------------------------------------------------
# 关键用例：同一份执行方案，按 report 判会误吃 5 项 P0，按 scheme 判 0 项 P0。
# P0-2a/b/c/d 与 P0-3 是「报批公文体例要求」，只对 report 生效。
print("[7] --doc-form：report 与 scheme 的否决差异")
EXEC_SCHEME = "\n".join([
    "一、活动概述", "二、活动顺序", "三、五场活动详情", "四、成本与预算",
    "五、物料复用与流转", "六、采购方案", "七、现场执行要点",
    "八、风险与红线",
    "事项 | 后果 | 对策",
    "14500 锂电池 | 无保护板存在过充风险 | 只买带保护板；充电有人值守",
    "写字机限位 | 无机械限位开关，X 轴可能撞机 | 调试盯住行程，先小行程测试再放全图",
    "预算项 | 数量 | 单价 | 金额",
    "主控 ×1 | 17 组 | 19.9 元 | 408 元",
    "合计：单组五场物料 760 元，其中 PCB 付费富余约 145 元。",
])
# 风险清单只列名词，不给后果与对策——这才是该被 P1-safety 拦下的样子。
EXEC_SCHEME_NO_MITIGATION = "\n".join([
    "一、活动概述", "四、成本与预算", "八、风险与红线",
    "风险项",
    "14500 锂电池",
    "写字机限位",
    "TF 卡规格",
    "预算项 | 数量 | 单价 | 金额",
    "合计：单组五场物料 760 元。",
])
f_report = []
review.check_plan_elements(EXEC_SCHEME, f_report, None, "report")
f_scheme = []
review.check_plan_elements(EXEC_SCHEME, f_scheme, None, "scheme")
p0_report = sorted(f["code"] for f in f_report if f["level"] == "P0")
p0_scheme = sorted(f["code"] for f in f_scheme if f["level"] == "P0")
check("report 形态吃 5 项 P0", p0_report,
      ["P0-2a", "P0-2b", "P0-2c", "P0-2d", "P0-3"])
check("scheme 形态 0 项 P0", p0_scheme, [])
check("scheme 仍记提示供扣分", sorted(f["code"] for f in f_scheme),
      ["P0-2a", "P0-2b", "P0-2c", "P0-2d"])
# 风险清单有「后果」「对策」两列 → P1-safety 不该报
check("后果+对策两列不报 P1-safety",
      "P1-safety" in [f["code"] for f in f_scheme], False)
# 只列风险名词、不给对策 → 必须报 P1-safety
f_bare = []
review.check_plan_elements(EXEC_SCHEME_NO_MITIGATION, f_bare, None, "scheme")
check("风险清单缺两列 → 报 P1-safety",
      "P1-safety" in [f["code"] for f in f_bare], True)
# 但「缺风险章节」本身在 scheme 下仍是 P0（连风险都不列 = 出事无人知）
f_none = []
review.check_plan_elements("一、活动概述\n四、成本与预算\n合计 760 元。", f_none, None, "scheme")
check("scheme 缺整个风险章节 → P0-3",
      "P0-3" in [f["code"] for f in f_none if f["level"] == "P0"], True)
# report 形态下有完整公文要素 → 不应误报
FULL_REPORT = "\n".join([
    "一、活动概述", "六、安全保障与应急预案",
    "用电安全：配备漏电保护。", "消防：配备灭火器。", "物品：贵重物品专人看管。",
    "噪音：活动 21:00 前结束。", "不可抗力：遇暴雨顺延一周。",
    "特此报告，妥否，请批示！",
    "指导教师签字 ______    日期 ____年__月__日",
    "万物互联协会",
    "2026年9月29日",
    "主办：万物互联协会  协办：西北大学电子信息学院",
])
f_full = []
review.check_plan_elements(FULL_REPORT, f_full, None, "report")
check("要素齐全的策划案 0 项 P0",
      sorted(f["code"] for f in f_full if f["level"] == "P0"), [])

# --------------------------------------------------------------------------
# 8. 海报三要素与 m00695 降级
# --------------------------------------------------------------------------
# 负责人 m00695：活动时间地点没有格式限制、配色深色即可不限制。
# 所以评审模型若仍按旧版把 P0-2f/P0-2g/配色类 P0-2c 当否决理由，必须在导入时降级，
# 否则规则会把 m00695 明确放过的产出物全部退回。
print("[8] m00695：时间地点/配色不得作为否决理由")
V = lambda code, ev: {"code": code, "evidence": ev, "detail": "按旧版标准判否决"}

hit, items, notes = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2f", "未见活动地点")]}}, "poster")
check("P0-2f 缺地点被降级为不否决", hit, False)
check("P0-2f 标记 downgraded", items[0]["level"], "downgraded")
check("P0-2f 留下降级说明", "m00695" in items[0]["detail"], True)
check("降级有记账", any("降级" in n for n in notes), True)

hit, items, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2g", "活动时间只写 10月19日，无时段")]}}, "poster")
check("P0-2g 只写日期被降级", hit, False)

# 配色类 P0-2c 降级
hit, items, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2c", "配色不符合两会 VI 主色，底色过浅")]}}, "poster")
check("仅因配色的 P0-2c 降级", hit, False)
# 徽章类 P0-2c 不降级（徽章位置/拉伸是真违规）
hit, items, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2c", "校徽被拉伸变形")]}}, "poster")
check("徽章拉伸的 P0-2c 仍否决", hit, True)

# 真正的三要素缺失照常否决
hit, items, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2e", "右上角只有校徽，没有团徽")]}}, "poster")
check("P0-2e 徽章缺失仍否决", hit, True)
hit, _, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2i", "右下角二维码为示例图")]}}, "poster")
check("P0-2i 二维码占位仍否决", hit, True)
hit, _, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2h", "只写协办单位，无主办单位")]}}, "poster")
check("P0-2h 缺主办单位仍否决", hit, True)

# 混合：降级项 + 真否决项，否决仍成立
hit, items, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2f", "未见地点"), V("P0-2e", "缺团徽")]}}, "poster")
check("混合情形仍否决", hit, True)
check("混合情形保留 1 条 live", len([i for i in items if i["level"] != "downgraded"]), 1)

# 非 poster 类型不做海报降级（策划案缺签字位该拒还是拒）
hit, _, _ = review.normalize_veto(
    {"veto": {"hit": True, "items": [V("P0-2d", "缺指导教师签字位")]}}, "plan")
check("策划案缺签字位仍否决", hit, True)

# 跨族复核不能拿未降级的 veto 对比
cc = review.crosscheck(
    {"scores": S(4, 4, 4, 4, 4), "veto": {"hit": True}, "_veto_hit": False},
    {"scores": S(4, 4, 4, 4, 4), "veto": {"hit": False}}, 1)
check("降级后不误报否决不一致", cc["needs_human"], False)

# credit-map：降级后仍按原档位算分，不受否决影响
check("三要素全过仍算 B 档学分",
      review.credit("poster", "B", 0.1, False, False), 0.1)

# --------------------------------------------------------------------------
# 9. 拖入即用：类型与文档形态自动识别
# --------------------------------------------------------------------------
# 这层判错方向不对称，所以两条都要锁死：
#   把执行方案误判成报批策划案 → 好文档被退回（白挨罚）
#   把报批策划案误判成执行方案 → 该退的溜过去
# 因此只认强标记，宁可判 scheme。
print("[9] 类型与文档形态自动识别")
check("docx → plan", review.detect_kind(Path("a.docx")), "plan")
check("pptx → ppt", review.detect_kind(Path("a.pptx")), "ppt")
check("pdf → plan", review.detect_kind(Path("a.pdf")), "plan")
check("jpg → poster", review.detect_kind(Path("a.jpg")), "poster")
check("png → poster", review.detect_kind(Path("a.png")), "poster")
check("md → copy", review.detect_kind(Path("a.md")), "copy")
check("txt → copy", review.detect_kind(Path("a.txt")), "copy")
check("未知后缀识别不出", review.detect_kind(Path("a.xyz")), "")

check("有妥否 → report", review.detect_doc_form("妥否，请批示！")[0], "report")
check("有签字位 → report",
      review.detect_doc_form("指导教师签字 ______")[0], "report")
check("纯执行方案 → scheme",
      review.detect_doc_form("八、风险与红线\n事项 后果 对策")[0], "scheme")
check("普通说明文 → scheme", review.detect_doc_form("本周活动安排")[0], "scheme")
# 「写字机限位」含「签」但不是「签字」——不能被当成报批体例
check("非签字位的「签」不算 report",
      review.detect_doc_form("五、物料规格：注意限位")[0], "scheme")

# resolve_base：报告里显示的基础分值必须是真的，不能显示 0.00
check("plan 基础分值 0.3", review.resolve_base("plan", None, False, False), 0.3)
check("ppt 整套 0.2", review.resolve_base("ppt", None, False, False), 0.2)
check("ppt 单页 0.05", review.resolve_base("ppt", None, True, False), 0.05)
check("contest 报名档 0.3", review.resolve_base("contest", None, False, False), 0.3)
check("contest 获奖档 0.6", review.resolve_base("contest", None, False, True), 0.6)
check("--base 覆盖查表", review.resolve_base("plan", 0.5, False, False), 0.5)

# --------------------------------------------------------------------------
# 10. 位置参数（拖入文件）与旧用法等价
# --------------------------------------------------------------------------
import subprocess  # noqa: E402


def run_cli(*argv):
    p = subprocess.run([PY, str(ROOT / "lib" / "review.py")] + list(argv),
                       capture_output=True, text=True, encoding="utf-8")
    return p.stdout + p.stderr


tvpro = ROOT / "samples" / "plan-tvpro.docx"
out_pos = run_cli(str(tvpro), "--prescreen-only")
out_flag = run_cli("--type", "plan", "--file", str(tvpro), "--prescreen-only")
check("位置参数识别出类型", "[自动识别] 类型=plan" in out_pos, True)
check("位置参数识别出形态", "文档形态=report" in out_pos, True)
check("位置参数与 --type/--file 等价",
      [l for l in out_flag.splitlines() if "自动识别" not in l]
      == [l for l in out_pos.splitlines() if "自动识别" not in l], True)
check("海报警错消失（jpg 可直接拖入）",
      run_cli(str(ROOT / "samples" / "poster-esai.jpg"), "--prescreen-only").count("不支持的类型"), 0)
check("批量模式能跑通", "批量预筛" in run_cli("--batch", str(ROOT / "samples")), True)

# 拖入模型返回的结果 JSON，应自动当作导入模式，不需要再写 --submit
tmp = Path(tempfile.gettempdir()) / "cws_selftest_answer.json"
tmp.write_text(json.dumps({
    "veto": {"hit": False, "items": []},
    "scores": {"completeness": 4, "accuracy": 4, "executability": 3,
               "compliance": 5, "quality": 4},
    "fixes": ["a", "b", "c"], "summary": "ok"}, ensure_ascii=False),
    encoding="utf-8")
out_js = run_cli(str(tmp), "--type", "plan")
check("结果 JSON 拖入 → 自动导入", "实得学分" in out_js, True)
check("结果 JSON 拖入 → 算出学分", "实得学分：0.45" in out_js, True)
check("非 .json 的 .txt 不会误判成导入",
      "实得学分" not in run_cli(str(ROOT / "SKILL.md"), "--type", "copy"), True)

print()
if FAILS:
    print("=== %d 项未通过 ===" % len(FAILS))
    for f in FAILS:
        print(" -", f)
    sys.exit(1)
print("=== 全部通过 ===")