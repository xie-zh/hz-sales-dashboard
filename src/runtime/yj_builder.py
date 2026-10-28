# -*- coding: utf-8 -*-
"""
业绩数据板块取数引擎（yj.html / sku.html 数据源）
==================================================
数据源（帆软，走 bi_engine 官方导出接口，不改动 bi_engine 本身）：
  1. 业绩含续费   report_yj_d_xf_sr_v5_20260427.cpt
     参数（帆软参数面板）：TYPE(周期月1/自然月2) · SELECT_MONTH(2026-11 …)
                          Q_DATE1 / Q_DATE2（区间）
     预取若干「口径切片」，前端可按 口径+月份+日期区间 切换：
       - 周期月 / 自然月 × 最近 3 个整月（历史月缓存 12h）
       - 本周（周一~今日）· 今日（缓存 30min）
  2. 订单商品明细 crispsv2_order_sku_detail_sr.cpt
     - DATE1 = 周期月起点，DATE2 = 今日，全部分公司
     - 下载后聚合为「分公司×事业部×业态×商品×下单日×付款日」粒度，控制 JSON 体积
     - 另出 catRows：按「分公司×二级分类×下单日×付款日」聚合（订单编号去重）

给 build_snapshot.build() 调用：
  yj_builder.build_yj_meta()  -> dict（业绩含续费，多切片）
  yj_builder.build_sku_meta() -> dict（订单商品明细，聚合行）
内置缓存：_bi_cache/_yj/ 下 xlsx 距上次下载不足 CACHE_SEC 秒则直接复用。
"""
import datetime
import json
import os
import sys

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE_DIR, "Hzdggbot"))
import bi_engine  # noqa: E402

CACHE_DIR = os.path.join(BASE_DIR, "_bi_cache", "_yj")
CACHE_SEC = 30 * 60          # 当期切片：30 分钟刷新一次（与看板更新节奏一致）
HIST_CACHE_SEC = 12 * 60 * 60  # 历史月份切片：数据不再变化，12 小时复用
MONTH_BACK = 3               # 每种口径预取最近几个月
YJ_URL = ("https://bi.example.com/webroot/decision/view/report"
          "?viewlet=dggReport/caiwu/report_yj_d_v4/report_yj_d_xf_sr_v5_20260427.cpt")
SKU_URL = ("https://bi.example.com/webroot/decision/view/report"
           "?viewlet=dggReport/yj/crispsv2_order_sku_detail_sr.cpt")

_R2 = lambda v: round(float(v), 2) if v not in (None, "", "--") else None


# --------------------------------------------------------------------------
# 下载（带缓存）
# --------------------------------------------------------------------------
def _login_no():
    try:
        with open(os.path.join(BASE_DIR, "Hzdggbot", "bi_sources.json"), encoding="utf-8") as f:
            return json.load(f).get("bi_login_no")
    except Exception:
        return None


def _period_month_start(today=None):
    """周期月起点：每月 15 日开启新周期月（10-15 ~ 11-14 为「2026-11」周期月）。"""
    d = today or datetime.date.today()
    if d.day >= 15:
        return d.replace(day=15)
    prev = (d.replace(day=1) - datetime.timedelta(days=1)).replace(day=15)
    return prev


def _grab(sid, url, flt, max_age=CACHE_SEC):
    os.makedirs(CACHE_DIR, exist_ok=True)
    path = os.path.join(CACHE_DIR, "%s.xlsx" % sid)
    if os.path.exists(path) and (time_age(path) < max_age):
        return path
    src = {"id": sid, "url": url, "filter": flt}
    p = bi_engine.download_one(src, CACHE_DIR, _login_no(), retries=2)
    return p if p else (path if os.path.exists(path) else None)


# --------------------------------------------------------------------------
# 月份 / 日期区间切片（对应帆软参数面板：TYPE · SELECT_MONTH · Q_DATE1/2）
# --------------------------------------------------------------------------
def _ym_add(y, m, k):
    t = y * 12 + (m - 1) + k
    return t // 12, t % 12 + 1


def _period_label(d):
    """周期月 label：每月 15 日起进入下一周期月（10-19 → 2026-11）。"""
    y, m = d.year, d.month
    if d.day >= 15:
        y, m = _ym_add(y, m, 1)
    return "%04d-%02d" % (y, m)


def _nature_label(d):
    return "%04d-%02d" % (d.year, d.month)


def _period_span(label):
    """周期月 2026-11 → 2026-10-15 ~ 2026-11-14。"""
    y, m = int(label[:4]), int(label[5:7])
    py, pm = _ym_add(y, m, -1)
    return datetime.date(py, pm, 15), datetime.date(y, m, 14)


def _nature_span(label):
    y, m = int(label[:4]), int(label[5:7])
    ny, nm = _ym_add(y, m, 1)
    return datetime.date(y, m, 1), datetime.date(ny, nm, 1) - datetime.timedelta(days=1)


def slice_defs(today=None):
    """→ [{sid,key,kind,label,start,end,flt,hist}]，前端按 key 取用。"""
    today = today or datetime.date.today()
    out = []
    for kind, lf, sf, tcode, tag in (("周期月", _period_label, _period_span, "1", "P"),
                                     ("自然月", _nature_label, _nature_span, "2", "N")):
        cy, cm = int(lf(today)[:4]), int(lf(today)[5:7])
        for k in range(MONTH_BACK):
            y, m = _ym_add(cy, cm, -k)
            lab = "%04d-%02d" % (y, m)
            s, e = sf(lab)
            hist = k > 0
            end = e if hist else min(e, today)
            if s > today:                       # 未来月份（自然月不会，保险）
                continue
            out.append({
                "sid": "YJ_%s_%s" % (tag, lab), "key": "%s|%s" % (kind, lab),
                "kind": kind, "label": lab,
                "start": s.strftime("%Y-%m-%d"), "end": end.strftime("%Y-%m-%d"),
                "hist": hist,
                "flt": {"TYPE": tcode, "SELECT_MONTH": lab,
                        "Q_DATE1": s.strftime("%Y-%m-%d"), "Q_DATE2": end.strftime("%Y-%m-%d")},
            })
    mon = today - datetime.timedelta(days=today.weekday())
    pl = _period_label(today)
    out.append({"sid": "YJ_WEEK", "key": "本周", "kind": "周", "label": "本周",
                "start": mon.strftime("%Y-%m-%d"), "end": today.strftime("%Y-%m-%d"),
                "hist": False,
                "flt": {"TYPE": "1", "SELECT_MONTH": pl,
                        "Q_DATE1": mon.strftime("%Y-%m-%d"), "Q_DATE2": today.strftime("%Y-%m-%d")}})
    out.append({"sid": "YJ_TODAY", "key": "今日", "kind": "日", "label": "今日",
                "start": today.strftime("%Y-%m-%d"), "end": today.strftime("%Y-%m-%d"),
                "hist": False,
                "flt": {"TYPE": "1", "SELECT_MONTH": pl,
                        "Q_DATE1": today.strftime("%Y-%m-%d"), "Q_DATE2": today.strftime("%Y-%m-%d")}})
    return out


def time_age(path):
    import time as _t
    return _t.time() - os.path.getmtime(path)


# --------------------------------------------------------------------------
# 业绩含续费：解析
# --------------------------------------------------------------------------
def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _row(r):
    """一行 → dict（列序见报表第 8 行表头，r[0] 为空列）。"""
    return {
        "name": (str(r[1]).strip() if r[1] else ""),
        "target": _num(r[2]), "done": _num(r[3]), "today": _num(r[4]),
        "rate": _num(r[5]), "rank": r[6],
        "hv_t": _num(r[7]), "hv_d": _num(r[8]), "hv_rate": _num(r[9]),
        "refund": _num(r[10]), "refund_ly": _num(r[11]), "refund_today": _num(r[12]),
        "ly": _num(r[13]), "yoy": _num(r[14]),
    }


def _parse_yj(path):
    """→ (notes, rows_dict{name:rowdict}, seq[list(name, 按报表出现顺序, 保留重复)])。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    notes, rows, seq = [], {}, []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        cells = [c for c in r if c not in (None, "")]
        if i <= 6:
            if cells:
                notes.append(str(cells[0]))
            continue
        if i == 7:
            continue                      # 表头行（列名硬编码在 _row）
        d = _row(r)
        if not d["name"]:
            continue
        if d["name"] not in rows:
            rows[d["name"]] = d
        seq.append(d["name"])
    return notes, rows, seq


def _split_companies(rows, seq):
    """按出现顺序拆：分公司总计 + 每家分公司的事业部明细 + 其他（后台/供应链等）。"""
    companies, cur, others, seen_units = {}, None, [], set()
    for n in seq:
        d = rows[n]
        if n.endswith("公司"):
            cur = n
            if n not in companies:
                companies[n] = {"total": d, "units": []}
            else:
                companies[n]["total"] = d      # 总览段与明细段数值相同，以明细段为准
            continue
        if n in seen_units:
            continue
        prefix2 = n[:2]
        if cur and (n.startswith(cur.replace("公司", "")) or prefix2 == cur[:2]):
            companies[cur]["units"].append(d)
            seen_units.add(n)
        else:
            others.append(d)
            seen_units.add(n)
    comp_names = list(companies.keys())
    return companies, comp_names, others


def _slice_payload(path):
    """一份 xlsx → {notes, group, overview, companies, company_names}。"""
    notes, rows, seq = _parse_yj(path)
    companies, comp_names, _others = _split_companies(rows, seq)
    overview, seen = [], set()
    for n in seq:
        if n.endswith("公司"):
            break
        if n not in seen:
            seen.add(n)
            overview.append(rows[n])
    return {"notes": notes, "group": rows.get("集团总计"), "overview": overview,
            "companies": companies, "company_names": comp_names}


def build_yj_meta(force=False):
    defs = slice_defs()
    slices, notes, main = {}, [], None
    for d in defs:
        age = 0 if force else (HIST_CACHE_SEC if d["hist"] else CACHE_SEC)
        p = _grab(d["sid"], YJ_URL, d["flt"], max_age=age)
        if not p or not os.path.exists(p):
            continue
        try:
            pl = _slice_payload(p)
        except Exception:
            continue
        if not pl["companies"]:
            continue
        pl["label"] = d["label"]
        pl["kind"] = d["kind"]
        pl["start"] = d["start"]
        pl["end"] = d["end"]
        pl["type"] = "周期月" if d["flt"].get("TYPE") == "1" else "自然月"
        slices[d["key"]] = pl
        if not notes:
            notes = pl["notes"]
        if d["key"] == "周期月|" + _period_label(datetime.date.today()):
            main = d["key"]

    if not slices:
        return None

    # 默认切片 = 当前周期月；缺失则取第一个
    if not main or main not in slices:
        main = "周期月|" + _period_label(datetime.date.today())
        if main not in slices:
            main = sorted(slices.keys())[0]
    d0 = slices[main]

    months = {}
    for k in slices:
        if "|" in k:
            months.setdefault(k.split("|")[0], []).append(k.split("|")[1])
    for v in months.values():
        v.sort(reverse=True)

    return {
        "meta": {
            "data_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "notes": notes,
            "week_ok": "本周" in slices,
            "default_key": main,
            "months": months,
            "span": [d0["start"], d0["end"]],
        },
        "slices": slices,
        # 兼容旧字段（顶层 = 默认切片）
        "group": d0["group"],
        "overview": d0["overview"],
        "company_names": d0["company_names"],
        "companies": d0["companies"],
        "week_companies": slices.get("本周", {}).get("companies"),
    }


def _week_monday(today=None):
    d = today or datetime.date.today()
    return (d - datetime.timedelta(days=d.weekday())).strftime("%Y-%m-%d")


# --------------------------------------------------------------------------
# 订单商品明细：解析（聚合）
# --------------------------------------------------------------------------
def _d10(v):
    s = str(v or "").strip()
    return s[:10] if len(s) >= 10 else ""


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def build_sku_meta(force=False):
    flt = {"DATE1": _period_month_start().strftime("%Y-%m-%d"),
           "DATE2": datetime.date.today().strftime("%Y-%m-%d")}
    p = _grab("SKU_SPAN", SKU_URL, flt, max_age=0) if force else _grab("SKU_SPAN", SKU_URL, flt)
    if not p or not os.path.exists(p):
        return None
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb.active
    agg = {}            # 商品粒度：[分公司,事业部,业态,商品,下单日,付款日]
    cagg = {}           # 二级分类粒度：[分公司,二级分类,下单日,付款日]（订单编号去重）
    total_orders = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        comp = str(r[18] or "未知").strip()
        bu = str(r[19] or "未知").strip()
        cat = str(r[26] or r[27] or "未分类").strip()   # 业态优先，空则退一级分类
        cat2 = str(r[28] or r[27] or r[26] or "未分类").strip()   # 二级分类
        prod = str(r[24] or "未知商品").strip()
        od = _d10(r[7])
        pd_ = _d10(r[14])
        amt = _f(r[30]); paid = _f(r[36]); ref = _f(r[37])
        ono = str(r[0])

        k = (comp, bu, cat, prod, od, pd_)
        a = agg.get(k)
        if a is None:
            agg[k] = a = [0, 0.0, 0.0, 0.0, set()]
        a[0] += 1
        a[1] += amt; a[2] += paid; a[3] += ref
        a[4].add(ono)

        # 分类粒度带订单号 → 前端按分类汇总时「订单编号去重」可精确计算
        ck = (comp, cat, cat2, od, pd_, ono)
        c = cagg.get(ck)
        if c is None:
            cagg[ck] = c = [0, 0.0, 0.0, 0.0]
        c[0] += 1
        c[1] += amt; c[2] += paid; c[3] += ref
        total_orders.add(ono)

    rows = [[c, b, t, pr, od, pdr, a[0], round(a[1], 2), round(a[2], 2),
             round(a[3], 2), len(a[4])]
            for (c, b, t, pr, od, pdr), a in sorted(agg.items(),
                                                    key=lambda kv: (-kv[1][1], kv[0]))]
    # [分公司, 业态, 二级分类, 下单日, 付款日, 订单编号, 子订单数, 金额, 已付, 退款]
    cat_rows = [[c, t, t2, od, pdr, ono, v[0], round(v[1], 2), round(v[2], 2),
                 round(v[3], 2)]
                for (c, t, t2, od, pdr, ono), v in sorted(cagg.items(),
                                                          key=lambda kv: (-kv[1][1], kv[0]))]
    return {
        "meta": {
            "data_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "span": [flt["DATE1"], flt["DATE2"]],
            "order_count": len(total_orders),
            "sku_count": sum(r[6] for r in rows),
            "amount": round(sum(r[7] for r in rows), 2),
            "paid": round(sum(r[8] for r in rows), 2),
            "refund": round(sum(r[9] for r in rows), 2),
            "cat2_count": len(set(r[1] for r in cat_rows)),
        },
        "rows": rows,
        "cat_rows": cat_rows,
    }


if __name__ == "__main__":
    yj = build_yj_meta()
    print("YJ  公司数:", len(yj["company_names"]) if yj else None,
          "| 周度:", bool(yj and yj.get("week_companies")))
    sku = build_sku_meta()
    print("SKU 行数:", len(sku["rows"]) if sku else None,
          "| meta:", json.dumps(sku["meta"], ensure_ascii=False) if sku else None)
