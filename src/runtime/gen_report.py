# -*- coding: utf-8 -*-
"""
生成「分公司 · 销售通话录音质检报告」单文件 HTML
=================================================
聚焦：录音质检（主体，真实录音库） + 关键指标概览（KPI）

数据来源：
  - ry_parser.build_meta()       → LTC 指标（KPI） + 电邀规划师→事业部映射
  - dashboard_server.build_meta()→ 活动量（KPI：达标率）
  - sanli_parser.build_meta()    → 三力（KPI：破零达成率）
  - 通话记录.xlsx                → 真实通话录音库（可播放 + 待 FunASR 转写）

输出：单文件 HTML（数据内联为 JSON），可直接打开 / 部署到 CloudStudio。
"""
import os, re, json, datetime
from collections import Counter

import openpyxl
import dashboard_server as ds
import sanli_parser as sp
import ry_parser as rp
import score_calls as sc
import ltc_daily

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, "report_template.html")
CALL_RECORD_PATH = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\通话记录.xlsx"


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("%", ""))
    except Exception:
        return None


def _parse_md(date_str):
    m = re.search(r"(\d{1,2})月(\d{1,2})日", date_str or "")
    if m:
        return int(m.group(1)), int(m.group(2))
    return datetime.date.today().month, datetime.date.today().day


def build_kpi():
    """仅计算关键指标概览所需 KPI（保留原三数据源口径）"""
    act = ds.build_meta()
    sanli = sp.build_meta()
    ry = rp.build_meta()

    call_total = connect_total = s30_contrib = wechat_total = profit_total = 0.0
    for e in ry.get("employees", []):
        ltc = e.get("ltc") or {}
        call = _num(ltc.get("当日全部电话拨打量")) or 0
        conn = _num(ltc.get("当日全部电话接通量")) or 0
        s30r = _num(ltc.get("当日全部电话30秒占比")) or 0
        wec = _num(ltc.get("90天加微量")) or 0
        prof = _num(ltc.get("月累计实算利润业绩")) or 0
        call_total += call
        connect_total += conn
        s30_contrib += conn * s30r
        wechat_total += wec
        profit_total += prof

    s30_rate = (s30_contrib / connect_total) if connect_total else None
    connect_rate = (connect_total / call_total) if call_total else None
    st = act.get("stats_total") or {}
    activity_rate = _num(st.get("当天目标完成率"))
    sanli_rate = (sanli.get("summary", {}).get("门槛破零值", {}).get("达成率"))

    return {
        "call_total": round(call_total),
        "connect_total": round(connect_total),
        "connect_rate": connect_rate,
        "s30_rate": s30_rate,
        "wechat_total": round(wechat_total),
        "profit_total": round(profit_total),
        "activity_rate": activity_rate,
        "sanli_rate": sanli_rate,
    }, ry.get("meta", {}).get("data_time", ""), act.get("meta", {}).get("data_time", ""), sanli.get("meta", {}).get("data_time", "")


def _to_sec(s):
    if not s:
        return None
    m = re.match(r'(\d+):(\d+):(\d+)', str(s))
    if m:
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))
    m = re.match(r'(\d+):(\d+)', str(s))
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def build_recordings():
    """
    从「通话记录.xlsx」提取真实通话录音。
    字段：id, person, dept, std, source, connected, durationSec, durationText,
          time, customer, audioUrl, transcribed(待 FunASR 转写后为 True)
    电邀规划师 -> 事业部 通过 ry_parser 的人员表映射。
    """
    # 电邀规划师 -> 事业部 / 部门标准 映射
    dept_map = {}
    try:
        ry = rp.build_meta()
        for e in ry.get("employees", []):
            dept_map[e.get("姓名")] = (e.get("事业部"), e.get("部门标准"))
    except Exception as ex:
        print("  [warn] 人员映射失败:", ex)

    recs = []
    if not os.path.exists(CALL_RECORD_PATH):
        print("  [warn] 未找到通话记录文件:", CALL_RECORD_PATH)
        return recs

    wb = openpyxl.load_workbook(CALL_RECORD_PATH, read_only=True, data_only=True)
    ws = wb["通话记录"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return recs
    header = rows[0]
    idx = {h: i for i, h in enumerate(header)}
    rec_idx = idx.get("录音地址")
    caller_idx = idx.get("拨打人")
    src_idx = idx.get("通话来源")
    conn_idx = idx.get("接通状态")
    dur_idx = idx.get("通话时长")
    date_idx = idx.get("拨出时间/呼入时间")
    cust_idx = idx.get("客户名称")

    n = 0
    for r in rows[1:]:
        url = r[rec_idx] if rec_idx is not None else None
        if not url or not str(url).startswith("http"):
            continue
        caller = str(r[caller_idx]) if r[caller_idx] else "未知"
        dinfo = dept_map.get(caller, (None, None))
        sec = _to_sec(r[dur_idx])
        recs.append({
            "id": "R%03d" % (n + 1),
            "person": caller,
            "dept": dinfo[0] or "未匹配",
            "std": dinfo[1] or "",
            "source": str(r[src_idx]) if r[src_idx] else "",
            "connected": str(r[conn_idx]) == "已接通",
            "durationSec": sec or 0,
            "durationText": str(r[dur_idx]) if r[dur_idx] else "",
            "time": str(r[date_idx])[:16] if r[date_idx] else "",
            "customer": str(r[cust_idx]) if r[cust_idx] else "未知客户",
            "audioUrl": str(url),
            "overall": None,
            "dims": [],
            "strengths": [], "weaknesses": [], "veto": [], "suggestions": [],
            "transcript": [],
            "transcribed": False,
        })
        n += 1
    return recs


def build_rec_stat(recs):
    if not recs:
        return {"total": 0, "callers": {}, "sources": {}, "depts": {}, "ge30": 0, "avgSec": 0}
    callers = dict(Counter(r["person"] for r in recs))
    sources = dict(Counter(r["source"] for r in recs))
    depts = dict(Counter(r["dept"] for r in recs))
    ge30 = sum(1 for r in recs if r["durationSec"] >= 30)
    ge10 = sum(1 for r in recs if r["durationSec"] >= 10)
    avg = round(sum(r["durationSec"] for r in recs) / len(recs))
    return {"total": len(recs), "callers": callers, "sources": sources,
            "depts": depts, "ge30": ge30, "ge10": ge10, "avgSec": avg}


SCORED_PATH = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\recordings_scored.json"

def merge_scored(recordings):
    """把 score_calls.py 产出的评分/转写结果，按 url 合并到录音库。
    命中的录音标记 transcribed=True，并填充 overall/dims/strengths/weaknesses/suggestions/segments。"""
    if not os.path.exists(SCORED_PATH):
        return recordings, 0
    try:
        scored = json.load(open(SCORED_PATH, encoding="utf-8"))
    except Exception as ex:
        print("  [warn] 读取评分结果失败:", ex)
        return recordings, 0
    by_url = {s.get("url"): s for s in scored}
    n = 0
    for r in recordings:
        s = by_url.get(r.get("audioUrl"))
        if not s or s.get("status") != "ok":
            continue
        r["transcribed"] = True
        r["overall"] = s.get("overall")
        r["dims"] = s.get("dims", [])
        r["strengths"] = s.get("strengths", [])
        r["weaknesses"] = s.get("weaknesses", [])
        r["veto"] = s.get("veto", [])
        r["suggestions"] = s.get("suggestions", [])
        r["segments"] = s.get("segments", [])
        r["rawTranscript"] = s.get("transcript", "")
        n += 1
    print(f"  已合并评分结果 {n} 条（transcribed=True）")
    return recordings, n


DIM_ORDER = list(sc.DIMS.keys())  # 开场介绍 / 痛点价值 / 需求挖掘 / 异议处理 / 加微埋点

def build_agent_summary(recordings):
    """聚合每位电邀规划师（≥30秒已转写样本）的五维均分 + 优势/短板/否决项/改进建议。"""
    from collections import Counter as _C
    by_person = {}
    for r in recordings:
        if not r.get("transcribed"):
            continue
        by_person.setdefault(r.get("person"), []).append(r)
    summary = []
    for p, rs in by_person.items():
        # 各维度均分
        acc = {d: [] for d in DIM_ORDER}
        for r in rs:
            for dm in r.get("dims", []):
                if dm["name"] in acc:
                    acc[dm["name"]].append(dm["score"])
        avg_dims = []
        for d in DIM_ORDER:
            vals = acc[d]
            avg = round(sum(vals) / len(vals), 1) if vals else 0
            avg_dims.append({"name": d, "score": avg})
        overall = round(sum(r["overall"] for r in rs) / len(rs), 1)
        strengths = [d for d in DIM_ORDER if (sum(acc[d]) / len(acc[d]) if acc[d] else 0) >= 4]
        weaknesses = [d for d in DIM_ORDER if (sum(acc[d]) / len(acc[d]) if acc[d] else 0) <= 2.5]
        # 否决项聚合（按类型计数）
        vc = _C()
        for r in rs:
            for v in r.get("veto", []):
                vc[v] += 1
        vetoes = [{"text": k, "count": c} for k, c in vc.most_common()]
        # 改进建议：来自短板维度的 bad 模板
        suggestions = [sc.DIMS[d]["bad"] for d in weaknesses]
        if not suggestions:
            suggestions.append("整体表现较稳，建议持续强化加微与回访埋点，提升转化闭环。")
        summary.append({
            "person": p,
            "count": len(rs),
            "avg_dims": avg_dims,
            "overall": overall,
            "strengths": strengths,
            "weaknesses": weaknesses,
            "vetoes": vetoes,
            "suggestions": suggestions,
        })
    summary.sort(key=lambda x: x["overall"], reverse=True)
    return summary

def build():
    kpi, ltc_time, act_time, sanli_time = build_kpi()
    recordings = build_recordings()
    recordings, n_scored = merge_scored(recordings)
    rec_stat = build_rec_stat(recordings)
    agent_summary = build_agent_summary(recordings)

    md, mday = _parse_md(ltc_time)
    gen = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    meta = {
        "title": "分公司 · 销售通话录音质检报告",
        "date": ltc_time,
        "generated": gen,
        "activity_time": act_time,
        "sanli_time": sanli_time,
        "ltc_time": ltc_time,
    }

    ltc = ltc_daily.build_ltc_daily()

    DATA = {"meta": meta, "kpi": kpi, "recordings": recordings, "recStat": rec_stat,
            "agentSummary": agent_summary, "ltcDaily": ltc}

    # 落一份紧凑 call_data.json，供 MiniQA 全工作空间数据摘要使用（避免解析大 HTML）
    n_trans = sum(1 for r in recordings if r.get("transcribed"))
    call_data = {"meta": meta, "kpi": kpi, "recStat": dict(rec_stat, transcribed=n_trans),
                 "agentSummary": agent_summary}
    with open(os.path.join(BASE_DIR, "call_data.json"), "w", encoding="utf-8") as f:
        json.dump(call_data, f, ensure_ascii=False)

    with open(TEMPLATE, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("__DATA_JSON__", json.dumps(DATA, ensure_ascii=False))

    fname = "分公司销售通话录音质检报告-{:02d}{:02d}.html".format(md, mday)
    out = os.path.join(BASE_DIR, fname)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    print("报告生成完成 →", out)
    print("  数据日期:", meta["date"], "| LTC:", ltc_time)
    print("  KPI 拨打/接通/加微/业绩:", kpi["call_total"], "/", kpi["connect_total"],
          "/", kpi["wechat_total"], "/", kpi["profit_total"])
    print("  录音库:", rec_stat["total"], "段 | ≥30秒:", rec_stat["ge30"],
          "| 电邀规划师:", list(rec_stat["callers"].keys()),
          "| 事业部:", list(rec_stat["depts"].keys()))
    print("  LTC 电邀明细: 天数", len(ltc.get("days", [])),
          "| 有录音", ltc.get("staff_with_recordings"))
    return out


if __name__ == "__main__":
    build()
