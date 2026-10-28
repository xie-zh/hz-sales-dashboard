# -*- coding: utf-8 -*-
"""
生成「工作空间全量数据摘要」workspace_data.json
覆盖：活动量看板 / 员工三力 / 网销数据分析 / 销售通话录音质检
供 MiniQA 在每个页面 fetch 作为 AI 回答的数据背景。
"""
import json, re, os

BASE = os.path.dirname(os.path.abspath(__file__))

def load(p):
    with open(os.path.join(BASE, p), encoding="utf-8") as f:
        return json.load(f)

def r1(x):
    try: return round(float(x), 1)
    except: return x

def pctf(x):
    try: return str(round(float(x) * 100, 1)) + "%"
    except: return "—"

# ---------- 平衡花括号抽取 const DATA = {...} ----------
def extract_const_data(html):
    i = html.index("const DATA = ") + len("const DATA = ")
    depth = 0; in_str = False; esc = False; end = None
    while i < len(html):
        c = html[i]
        if in_str:
            if esc: esc = False
            elif c == "\\": esc = True
            elif c == '"': in_str = False
        else:
            if c == '"': in_str = True
            elif c == "{": depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0: end = i + 1; break
        i += 1
    return html[i - (end - i) if False else 0: end] if end else "{}"

# ---------- 活动量 ----------
act = load("_deploy/data.json")
meta = act.get("meta", {})
stats = act.get("stats", [])
total = act.get("total", [])
legal = act.get("legal", [])
n_total = len(total)
n_ok = sum(1 for t in total if t.get("是否达标") == "达标")
comp_avg = (sum(float(t.get("完成率") or 0) for t in total) / n_total * 100) if n_total else 0
see_sum = sum(float(t.get("见客合计") or 0) for t in total)
up_sum = sum(float(t.get("当日上门量") or 0) for t in total)
out_sum = sum(float(t.get("当日外出量") or 0) for t in total)
lines = []
lines.append("【一、分公司活动量看板】（数据时间：%s）" % meta.get("data_time", "未知"))
lines.append("考核部门 %d 个，全员明细 %d 人；整体完成率均值 %.2f%%，达标 %d 人（达标率 %s）。"
             % (len(stats), n_total, comp_avg, n_ok, pctf(n_ok / n_total if n_total else 0)))
lines.append("见客合计 %.0f 次，当日上门 %.0f 次，当日外出 %.0f 次。" % (see_sum, up_sum, out_sum))
lines.append("各事业部/部门过程量（总人数 / 考核 / 当天已完成 / 完成率）：")
for s in stats:
    lines.append("  - %s | %s：总 %s 人，考核 %s 人，已完成 %s，完成率 %s"
                 % (s.get("事业部",""), s.get("部门",""), r1(s.get("总人数")), r1(s.get("考核人数")),
                    r1(s.get("当天目标已完成")), pctf(s.get("当天目标完成率"))))
lines.append("法律日达标明细 %d 人。" % len(legal))
activity_text = "\n".join(lines)

# ---------- 三力 ----------
san = load("_deploy/sanli_data.json")
smeta = san.get("meta", {})
bus = san.get("bus", [])
semployees = len(san.get("employees", []))
ssum = san.get("summary", {})
lines = []
lines.append("【二、员工三力看板】（数据时间：%s）" % smeta.get("data_time", "未知"))
tp = ssum.get("门槛破零值") or {}; sp = ssum.get("标准人力值") or {}; hp = ssum.get("英雄人力值") or {}
lines.append("员工总数 %d 人，三力有效人力 %s（门槛破零达成率 %s / 标准人力达成率 %s / 英雄人力达成率 %s）。"
             % (semployees, ssum.get("有效人力"), pctf(tp.get("达成率")), pctf(sp.get("达成率")), pctf(hp.get("达成率"))))
for b in bus:
    rate = b.get("达成率", {})
    lines.append("  - %s：总 %s 人，有效人力 %s；达成率 门槛破零%s / 标准%s / 英雄%s"
                 % (b.get("事业部",""), b.get("总人数"), b.get("有效人力"),
                    pctf(rate.get("门槛破零值")), pctf(rate.get("标准人力值")), pctf(rate.get("英雄人力值"))))
sanli_text = "\n".join(lines)

# ---------- 网销数据分析 ----------
ry = load("_deploy/ry_data.json")
rmeta = ry.get("meta", {})
rsummary = ry.get("summary", {})
rdepartments = ry.get("departments", [])
rfunnel = ry.get("funnel_metrics", [])
rltc = ry.get("ltc_metrics", [])
remployees = len(ry.get("employees", []))
lines = []
lines.append("【三、网销数据分析看板】（数据时间：%s）" % rmeta.get("data_time", "未知"))
lines.append("员工总数 %d 人，涉及部门：%s。" % (remployees, "、".join([str(d) for d in rdepartments]) if rdepartments else "—"))
lines.append("经营概览：%s。" % "；".join("%s=%s" % (k, v) for k, v in rsummary.items() if v is not None))
if rfunnel:
    lines.append("LTC漏斗指标：" + "；".join("%s（%s）" % (f.get("k"), f.get("desc","")) for f in rfunnel))
if rltc:
    lines.append("LTC业务指标：" + "；".join("%s（%s）" % (f.get("k"), f.get("desc","")) for f in rltc))
ry_text = "\n".join(lines)

# ---------- 销售通话录音质检（读取 gen_report 落地的紧凑 call_data.json） ----------
call_text = "【四、销售通话录音质检报告】数据暂不可用。"
try:
    DATA = load("call_data.json")
    cmeta = DATA.get("meta", {})
    kpi = DATA.get("kpi", {})
    rs = DATA.get("recStat", {})
    ags = DATA.get("agentSummary", [])
    n_trans = rs.get("transcribed", 0)
    lines = []
    lines.append("【四、销售通话录音质检报告】（数据日期：%s）" % cmeta.get("date", "未知"))
    lines.append("当日电话拨打 %s 通，接通 %s 通（接通率 %s）；90天加微 %s，月累计利润业绩 ¥%s。"
                 % (r1(kpi.get("call_total")), r1(kpi.get("connect_total")),
                    pctf(kpi.get("connect_rate")), r1(kpi.get("wechat_total")), r1(kpi.get("profit_total"))))
    lines.append("录音库 %s 段，已转写质检 %s 段（≥10秒 %s / ≥30秒 %s），平均时长 %s 秒。"
                 % (rs.get("total"), n_trans, rs.get("ge10"), rs.get("ge30"), r1(rs.get("avgSec"))))
    # 否决项汇总
    veto_all = {}
    for a in ags:
        for v in a.get("vetoes", []):
            txt = v.get("text", "")
            veto_all[txt] = veto_all.get(txt, 0) + (v.get("count") or 0)
    if veto_all:
        lines.append("质检否决项（硬伤）汇总：" + "；".join("%s %s次" % (k, c) for k, c in veto_all.items()))
    # 电邀规划师评分
    if ags:
        ranked = sorted(ags, key=lambda a: a.get("overall", 0), reverse=True)
        lines.append("电邀规划师五维综合评分（综合/5，样本数）：")
        for a in ranked:
            dims = "，".join("%s%s" % (d.get("name"), r1(d.get("score"))) for d in a.get("avg_dims", []))
            lines.append("  - %s：综合 %s（样本 %s 段）；五维[%s]" % (a.get("person"), r1(a.get("overall")), a.get("count"), dims))
            if a.get("strengths"): lines.append("    优势：%s" % "、".join(a["strengths"]))
            if a.get("weaknesses"): lines.append("    短板：%s" % "、".join(a["weaknesses"]))
    call_text = "\n".join(lines)
except Exception as e:
    print("WARN: 读取 call_data.json 失败：", e)

# ---------- 合并为统一上下文 ----------
context_text = "\n\n".join([activity_text, sanli_text, ry_text, call_text])

out = {
    "generated_at": meta.get("server_time", ""),
    "boards": ["activity", "sanli", "ry", "call"],
    "context_text": context_text
}

# 写入两个部署目录
for d in ["_deploy", "report_deploy"]:
    p = os.path.join(BASE, d, "workspace_data.json")
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    print("写入", p, "context_text 长度", len(context_text))

# 打印预览（前 1500 字）
print("\n===== context_text 预览 =====")
print(context_text[:1800])
