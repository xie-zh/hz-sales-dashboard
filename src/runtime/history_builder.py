# -*- coding: utf-8 -*-
"""
活动量看板 · 每日历史快照（JSON 化，免 COM、免锁文件）
=================================================
把「当天活动量数据」按日期存档为 JSON，供 index.html 历史时间选择器加载；
对外提供 build_history_json() 重新生成 _deploy/history.json。

为什么改方案（2026-09-24）：
  旧逻辑：history.json 仅在 build_snapshot.build() 内、且「数据发生变化」时才重算，
  又依赖 历史数据/*.xlsx 的逐日留存。结果：BI 当日无变化→不重建→不写历史；
  源 Excel 被 WPS 占用/缺失→save_daily 失败→当天 xlsx 缺失。二者叠加 → 历史漏天。
  新逻辑：每天 23:59 由 snapshot_daily.py 跑本模块，直接从「已部署的 _deploy/data.json」
  （永不锁定、永远可读）拍一份当日 JSON 快照，并重新生成完整 history.json，
  不再依赖 xlsx 留存与原生 build 触发，杜绝漏天。
"""
import os
import sys
import json
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(BASE_DIR, "_deploy")
HIST_DIR = os.path.join(BASE_DIR, "历史数据")
SNAP_DIR = os.path.join(HIST_DIR, "_snap")            # 每日 JSON 快照（新方案主力）
DATA_JSON = os.path.join(DEPLOY, "data.json")          # 已部署的活动量数据
HISTORY_JSON = os.path.join(DEPLOY, "history.json")

# build_meta 额外附加的 3 个 meta 键；历史快照保持与 parse_workbook 同形，剥离之
_BUILD_META_KEYS = ("file_mtime", "file_mtime_ts", "server_time")


def _today():
    return datetime.date.today().strftime("%Y-%m-%d")


def _load_data_json():
    if not os.path.exists(DATA_JSON):
        return None
    try:
        with open(DATA_JSON, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def capture_today(date=None):
    """把当前 _deploy/data.json 拍成当日 JSON 快照（覆盖写，幂等）。返回日期或 None。

    来源用已部署的 data.json 而非源 xlsx：data.json 永不锁定、永远可读，
    且已是解析好的标准结构，规避「源 Excel 被 WPS 占用/缺失→快照失败」的老问题。
    """
    date = date or _today()
    data = _load_data_json()
    if data is None:
        print("  [skip] _deploy/data.json 不存在，无法拍今日快照")
        return None
    d = dict(data)
    meta = dict(d.get("meta", {}))
    for k in _BUILD_META_KEYS:
        meta.pop(k, None)
    meta["date"] = date
    d["meta"] = meta
    os.makedirs(SNAP_DIR, exist_ok=True)
    dst = os.path.join(SNAP_DIR, date + ".json")
    tmp = dst + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, dst)
        print("  [ok] 今日活动量快照:", date)
        return date
    except Exception as e:
        print("  [err] 今日快照写入失败:", e)
        return None


def backfill_from_xlsx():
    """一次性：把现有 历史数据/*.xlsx 解析为 JSON 快照，保留历史。"""
    try:
        import dashboard_server as ds
    except Exception as e:
        print("  [err] 无法导入 dashboard_server（缺 openpyxl?）:", e)
        return 0
    os.makedirs(SNAP_DIR, exist_ok=True)
    n = 0
    for fn in sorted(os.listdir(HIST_DIR)):
        if not fn.endswith(".xlsx") or fn.startswith("~$"):
            continue
        date = fn[:-5]
        dst = os.path.join(SNAP_DIR, date + ".json")
        if os.path.exists(dst):
            continue
        try:
            hdata = ds.parse_workbook(os.path.join(HIST_DIR, fn))
            hdata = dict(hdata)
            hdata["meta"] = dict(hdata.get("meta", {}))
            hdata["meta"]["date"] = date
            with open(dst, "w", encoding="utf-8") as f:
                json.dump(hdata, f, ensure_ascii=False)
            n += 1
            print("  [ok] backfill", date)
        except Exception as e:
            print("  [warn] backfill 失败:", fn, e)
    print("  backfill 完成，新增 %d 天" % n)
    return n


def build_history_json(ensure_today=True):
    """聚合所有日快照 → _deploy/history.json（按日期升序）。

    来源优先级：1) _snap/*.json（新方案主力） 2) 历史数据/*.xlsx（遗留兜底）
    3) 今日 _deploy/data.json（极端兜底，保证今天永远在列）。
    """
    os.makedirs(SNAP_DIR, exist_ok=True)
    os.makedirs(DEPLOY, exist_ok=True)
    items = {}

    # 1) JSON 日快照（主力）
    for fn in sorted(os.listdir(SNAP_DIR)):
        if not fn.endswith(".json") or fn.startswith("~$"):
            continue
        date = fn[:-5]
        try:
            with open(os.path.join(SNAP_DIR, fn), encoding="utf-8") as f:
                d = json.load(f)
            d = dict(d)
            d.setdefault("meta", {})
            d["meta"]["date"] = date
            items[date] = d
        except Exception as e:
            print("  [warn] 快照解析失败:", fn, e)

    # 2) xlsx 兜底（历史遗留 / 未 backfill 的日子，best-effort）
    try:
        import dashboard_server as ds
        for fn in sorted(os.listdir(HIST_DIR)):
            if not fn.endswith(".xlsx") or fn.startswith("~$"):
                continue
            date = fn[:-5]
            if date in items:
                continue
            try:
                hdata = ds.parse_workbook(os.path.join(HIST_DIR, fn))
                hdata = dict(hdata)
                hdata["meta"] = dict(hdata.get("meta", {}))
                hdata["meta"]["date"] = date
                items[date] = hdata
            except Exception as e:
                print("  [warn] xlsx 兜底解析失败:", fn, e)
    except Exception as e:
        print("  [warn] xlsx 兜底不可用:", e)

    # 3) 今日兜底（保证今天永远在列，即便 capture 没跑）
    if ensure_today:
        t = _today()
        if t not in items:
            data = _load_data_json()
            if data is not None:
                d = dict(data)
                meta = dict(d.get("meta", {}))
                for k in _BUILD_META_KEYS:
                    meta.pop(k, None)
                meta["date"] = t
                d["meta"] = meta
                items[t] = d
                print("  [ok] 兜底注入今日:", t)

    hist = [{"date": d, "data": items[d]} for d in sorted(items)]
    tmp = HISTORY_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False)
    os.replace(tmp, HISTORY_JSON)
    span = (sorted(items)[0], sorted(items)[-1]) if items else ("-", "-")
    print("history.json 天数:", len(hist), "（", span[0], "~", span[1], "）")
    return len(hist)


if __name__ == "__main__":
    sys.path.insert(0, BASE_DIR)
    backfill_from_xlsx()
    capture_today()
    build_history_json()
    print("done")
