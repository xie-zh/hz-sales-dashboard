# -*- coding: utf-8 -*-
"""
LTC 电邀明细数据解析（通话记录.xlsx · 电邀LTC流程指标数据子表）
==============================================================
供 gen_report.py 质检看板「关键指标概览」内嵌电邀明细看板使用。

数据来源：
  - 通话记录.xlsx            「电邀LTC流程指标数据」子表（9 天历史，每天 7 人，161 列）
  - 人员数据分析路径.xlsx     司龄（ry_parser 解析，真实值）
  - recordings_scored.json   有录音人员名单（仅显示这些人的明细）

输出：
  build_ltc_daily() → {
      days: [{date, date_cn, weekday, is_weekend, staff: [{name,dept,tenure,...19字段}]}],
      staff_with_recordings: [names],
      meta: {dept, total_days, workday_days, weekend_days, source}
  }
"""
import os, json, datetime, re
import openpyxl
import ry_parser as rp

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CALL_XLSX = os.path.join(BASE_DIR, "通话记录.xlsx")
SCORED = os.path.join(BASE_DIR, "recordings_scored.json")

YEAR = 2026

# LTC 子表列号（以表头定位，兼容顺序变化）
COL = {
    "date": 0, "dept": 1, "name": 2, "post": 3, "tenure_raw": 4,
    "plan_dial": 10, "plan_exec": 17, "plan_touch": 18, "plan_touch_rate": 19,
    "plan_transfer": 24, "all_dial": 31, "all_conn": 32, "all_duration": 33,
    "all_conn_rate": 39, "dial_30s": 64, "dial_60s": 65, "wechat_bind": 41,
    "visit_out": 67, "visit_door": 69, "visit_online": 71,
    "star2_3_new": 50, "star4_6_new": 51,
    "visit_30d": 74, "sign_30d": 84,
}


def _num(v):
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "").replace("%", ""))
    except Exception:
        return 0


def _parse_md(s):
    """'08-21' 或 '2026-08-21' -> datetime.date(2026,8,21)"""
    s = str(s).strip()
    parts = s.split("-")
    if len(parts) == 3:
        y = int(parts[0])
        m = int(parts[1].lstrip("0"))
        d = int(parts[2].lstrip("0"))
        return datetime.date(y, m, d)
    if len(parts) == 2:
        return datetime.date(YEAR, int(parts[0].lstrip("0")), int(parts[1].lstrip("0")))
    return datetime.date(YEAR, 1, 1)


def _weekday_cn(d):
    return ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][d.weekday()]


def _load_tenure():
    """司龄（月）：从 ry_parser（人员数据分析路径.xlsx）读取，name -> tenure"""
    tenure = {}
    try:
        ry = rp.build_meta()
        for e in ry.get("employees", []):
            name = str(e.get("姓名") or "").strip()
            t = e.get("司龄")
            if name and t is not None:
                tenure[name] = float(t)
    except Exception as ex:
        print("  [warn] 司龄读取失败:", ex)
    return tenure


def _load_with_recordings():
    """有录音质检的人员名单（recordings_scored.json 的 person 字段去重）"""
    if not os.path.exists(SCORED):
        return []
    try:
        recs = json.load(open(SCORED, encoding="utf-8"))
        names = []
        for r in recs:
            p = r.get("person")
            if p and p not in names:
                names.append(str(p).strip())
        return names
    except Exception as ex:
        print("  [warn] 有录音名单读取失败:", ex)
        return []


def build_ltc_daily():
    """解析 通话记录.xlsx LTC 子表 → 每日明细 + 人员过滤名单"""
    tenure = _load_tenure()
    with_rec = _load_with_recordings()

    if not os.path.exists(CALL_XLSX):
        print("  [warn] 未找到 通话记录.xlsx:", CALL_XLSX)
        return {"days": [], "staff_with_recordings": with_rec, "meta": {}}

    wb = openpyxl.load_workbook(CALL_XLSX, read_only=True, data_only=True)
    if "电邀LTC流程指标数据" not in wb.sheetnames:
        print("  [warn] 缺少子表: 电邀LTC流程指标数据")
        return {"days": [], "staff_with_recordings": with_rec, "meta": {}}
    ws = wb["电邀LTC流程指标数据"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 3:
        return {"days": [], "staff_with_recordings": with_rec, "meta": {}}

    # 用表头定位列（兼容表头顺序）
    header = rows[0]
    colmap = {}
    for i, h in enumerate(header):
        if h is None:
            continue
        hs = str(h).strip()
        if hs not in colmap:  # 保留第一次出现（161 列里列名唯一）
            colmap[hs] = i
    def col(*names):
        for n in names:
            if n in colmap:
                return colmap[n]
        return None

    C = {
        "date": col("日期"),
        "dept": col("部门"),
        "name": col("姓名"),
        "post": col("岗位"),
        "tenure_raw": col("司龄"),
        "plan_dial": col("当日计划拨打量"),
        "plan_exec": col("当日计划拨打执行量"),
        "plan_touch": col("当日计划拨打接触量"),
        "plan_touch_rate": col("当日计划拨打接触率"),
        "plan_transfer": col("当日计划拨打转私客量"),
        "all_dial": col("当日全部电话拨打量"),
        "all_conn": col("当日全部电话接通量"),
        "all_duration": col("当日全部通时(分钟)", "当日全部通时（分钟）"),
        "all_conn_rate": col("当日全部电话接通率"),
        "dial_30s": col("当日30S电话量", "当日30秒电话量"),
        "wechat_bind": col("当日全部新增微信绑定量", "当日新增微信量"),
        "visit_out": col("当日外出量"),
        "visit_door": col("当日上门量"),
        "visit_online": col("当日线上见客量"),
        "star2_3_new": col("当日新增2-3星客户量"),
        "star4_6_new": col("当日新增4-6星客户量"),
        "visit_30d": col("近30天见客量"),
        "sign_30d": col("近30天签单总量"),
        "company": col("公司"),
        "biz": col("事业部"),
        "staff_no": col("工号"),
    }

    days_map = {}
    for r in rows[2:]:  # row0=表头 row1=指标释义 row2+=数据
        if r[C["date"]] is None or r[C["name"]] is None:
            continue
        try:
            d = _parse_md(r[C["date"]])
        except Exception:
            continue
        name = str(r[C["name"]]).strip()
        plan_dial = _num(r[C["plan_dial"]])
        plan_exec = _num(r[C["plan_exec"]])
        plan_touch = _num(r[C["plan_touch"]])
        all_dial = _num(r[C["all_dial"]])
        all_conn = _num(r[C["all_conn"]])
        visit_total = _num(r[C["visit_out"]]) + _num(r[C["visit_door"]]) + _num(r[C["visit_online"]])
        star_total = _num(r[C["star2_3_new"]]) + _num(r[C["star4_6_new"]])
        rec = {
            "name": name,
            "staff_no": str(r[C["staff_no"]] or "").strip(),
            "dept": str(r[C["dept"]] or "").strip(),
            "biz": str(r[C["biz"]] or "").strip(),
            "post": str(r[C["post"]] or "").strip(),
            "tenure": tenure.get(name),                       # 真实司龄（月），可能 None
            "tenure_src": "ry" if name in tenure else None,
            "plan_dial": plan_dial,
            "plan_exec": plan_exec,
            "plan_exec_rate": (plan_exec / plan_dial) if plan_dial > 0 else None,
            "plan_touch": plan_touch,
            "plan_touch_rate": (_num(r[C["plan_touch_rate"]]) if C["plan_touch_rate"] is not None else 0),
            "plan_transfer": _num(r[C["plan_transfer"]]),
            "all_dial": all_dial,
            "all_conn": all_conn,
            "all_conn_rate": (_num(r[C["all_conn_rate"]]) if C["all_conn_rate"] is not None else 0),
            "all_duration": _num(r[C["all_duration"]]),
            "dial_30s": _num(r[C["dial_30s"]]),
            "wechat_bind": _num(r[C["wechat_bind"]]),
            "visit_total": visit_total,
            "star_total": star_total,
            "visit_30d": _num(r[C["visit_30d"]]),
            "sign_30d": _num(r[C["sign_30d"]]),
        }
        # 接通率兜底：若表内为 0 但拨打量>0 则计算
        if all_dial > 0 and rec["all_conn_rate"] <= 0:
            rec["all_conn_rate"] = all_conn / all_dial
        # 接触率兜底：若表内为 0 但执行量>0 则计算
        if plan_exec > 0 and rec["plan_touch_rate"] <= 0:
            rec["plan_touch_rate"] = plan_touch / plan_exec
        days_map.setdefault(d, []).append(rec)

    days = []
    for d in sorted(days_map.keys()):
        items = days_map[d]
        items.sort(key=lambda x: (-(x["plan_dial"] or 0), x["name"]))
        days.append({
            "date": d.isoformat(),
            "date_cn": d.strftime("%m-%d"),
            "weekday": _weekday_cn(d),
            "is_weekend": d.weekday() >= 5,
            "staff": items,
        })

    meta = {
        "dept": days[0]["staff"][0]["dept"] if days and days[0]["staff"] else "工商电邀业务六部",
        "biz": days[0]["staff"][0]["biz"] if days and days[0]["staff"] else "",
        "total_days": len(days),
        "workday_days": sum(1 for x in days if not x["is_weekend"]),
        "weekend_days": sum(1 for x in days if x["is_weekend"]),
        "date_range": [days[0]["date"], days[-1]["date"]] if days else [],
        "source": "通话记录.xlsx · 电邀LTC流程指标数据",
    }
    return {"days": days, "staff_with_recordings": with_rec, "meta": meta}


if __name__ == "__main__":
    data = build_ltc_daily()
    print("天数:", len(data["days"]))
    print("有录音:", data["staff_with_recordings"])
    print("meta:", data["meta"])
    if data["days"]:
        d0 = data["days"][-1]
        print("最近一天:", d0["date_cn"], d0["weekday"], "人数", len(d0["staff"]))
        for s in d0["staff"]:
            print("   ", s["name"], "司龄", s["tenure"], "计划", s["plan_dial"], "执行", s["plan_exec"],
                  "执行率", s["plan_exec_rate"], "拨打", s["all_dial"], "接通", s["all_conn"], "接通率", s["all_conn_rate"])
