# -*- coding: utf-8 -*-
"""
分公司活动量 · 在线数据看板 服务
- 每次请求实时读取本地 Excel（在线看板.xlsx），文件更新后下次轮询即自动刷新
- 纯标准库实现，无第三方依赖（除读取 xlsx 用的 openpyxl）
"""
import os
import re
import json
import time
import datetime
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import openpyxl
except ImportError:
    openpyxl = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, "分公司活动量-在线看板.xlsx")
INDEX_PATH = os.path.join(BASE_DIR, "index.html")
ECHARTS_PATH = os.path.join(BASE_DIR, "echarts.min.js")
HISTORY_DIR = os.path.join(BASE_DIR, "历史数据")

# ===== 解析缓存（按文件 mtime 失效）=====
_cache_lock = threading.Lock()
_cache = {"mtime": 0.0, "data": None, "error": None}


def _to_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        try:
            return float(v)
        except Exception:
            return None
    s = str(v).strip()
    if s in ("", "-", "None", "#REF!", "#REF", "不计算", "不计算"):
        return None
    # 处理百分比文本如 "80%"
    try:
        return float(s.replace("%", ""))
    except Exception:
        return None


def _to_str(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def _norm(s):
    return re.sub(r"\s+", "", str(s or ""))


def parse_workbook(path=XLSX_PATH):
    if openpyxl is None:
        raise RuntimeError("缺少 openpyxl，请先安装：pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    result = {"meta": {}, "stats": [], "stats_total": None,
              "legal": [], "total": []}

    # ---------- 数据统计（部门过程量汇总）----------
    if "数据统计" in wb.sheetnames:
        ws = wb["数据统计"]
        rows = list(ws.iter_rows(values_only=True))
        # 标题行（含数据时间）
        if rows and rows[0] and rows[0][1]:
            title = str(rows[0][1])
            result["meta"]["sheet_title"] = title
            # 兼容两种标题格式： "8月11日 18点00分" 和 "08月12日-11:26"
            m = re.search(r"(\d{1,2})月(\d{1,2})日[- ]?(\d{1,2})(?::(\d{2})|点(\d{1,2})分?)", title)
            if m:
                mm = (m.group(4) or m.group(5) or "00").zfill(2)
                result["meta"]["data_time"] = "{}月{}日 {}:{}".format(
                    int(m.group(1)), int(m.group(2)), m.group(3), mm)
        # 表头在 index 1：事业部(B) 部门(C) 总人数(D) 考核人数(E) 当天目标已完成(F) 完成率(G)
        cur_bu = None
        for i in range(2, len(rows)):
            r = rows[i]
            if len(r) < 7:
                continue
            bu_raw = r[1]
            dept2 = r[2]  # 部门
            bu = _norm(bu_raw) if (bu_raw is not None and _norm(bu_raw) != "") else cur_bu
            if bu == "合计":
                result["stats_total"] = {
                    "总人数": _to_float(r[3]),
                    "考核人数": _to_float(r[4]),
                    "当天目标已完成": _to_float(r[5]),
                    "当天目标完成率": _to_float(r[6]),
                }
                continue
            if dept2 is None or _norm(dept2) == "":
                continue
            if bu:
                cur_bu = bu
            result["stats"].append({
                "事业部": bu,
                "部门": _norm(dept2),
                "总人数": _to_float(r[3]) or 0,
                "考核人数": _to_float(r[4]) or 0,
                "当天目标已完成": _to_float(r[5]) or 0,
                "当天目标完成率": _to_float(r[6]),
            })

    # ---------- 总表（全员明细）----------
    if "总表" in wb.sheetnames:
        ws = wb["总表"]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = None
        for i, r in enumerate(rows):
            if len(r) > 4 and _norm(r[0]) == "事业部" and _norm(r[4]) == "姓名":
                hdr_idx = i
                break
        if hdr_idx is not None:
            for i in range(hdr_idx + 1, len(rows)):
                r = rows[i]
                if len(r) < 20:
                    continue
                if _norm(r[0]) == "":
                    continue
                result["total"].append({
                    "事业部": _norm(r[0]),
                    "部门": _norm(r[1]),
                    "岗位": _norm(r[2]),
                    "工号": _norm(r[3]),
                    "姓名": _norm(r[4]),
                    "入职日期": _to_str(r[5]),
                    "入职时间": _to_float(r[6]),
                    "英雄人力": _norm(r[7]),
                    "前半月英雄人力": _norm(r[8]),
                    "出勤": _to_float(r[9]),
                    "是否达标": _norm(r[10]),
                    "目标": _to_float(r[11]),
                    "折扣目标": _to_float(r[12]),
                    "完成": _to_float(r[13]),
                    "完成率": _to_float(r[14]),
                    "周活动量": _norm(r[15]),
                    "当日上门量": _to_float(r[16]) or 0,
                    "当日外出量": _to_float(r[17]) or 0,
                    "见客合计": _to_float(r[18]) or 0,
                    "可抵扣": _to_float(r[19]),
                })

    # ---------- 法律（法律事业部明细）----------
    if "法律" in wb.sheetnames:
        ws = wb["法律"]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = None
        for i, r in enumerate(rows):
            if len(r) > 4 and _norm(r[0]) == "事业部" and _norm(r[4]) == "姓名":
                hdr_idx = i
                break
        if hdr_idx is not None:
            for i in range(hdr_idx + 1, len(rows)):
                r = rows[i]
                if len(r) < 28:
                    continue
                if _norm(r[0]) == "":
                    continue
                result["legal"].append({
                    "事业部": _norm(r[0]),
                    "部门": _norm(r[1]),
                    "岗位": _norm(r[2]),
                    "工号": _norm(r[3]),
                    "姓名": _norm(r[4]),
                    "入职日期": _to_str(r[5]),
                    "入职时间": _to_float(r[6]),
                    "英雄人力": _norm(r[7]),
                    "前半月英雄人力": _norm(r[8]),
                    "出勤": _to_float(r[9]),
                    "是否达标": _norm(r[10]),
                    "目标": _to_float(r[11]),
                    "折扣目标": _to_float(r[12]),
                    "今日完成": _to_float(r[13]) or 0,
                    "滚动2日完成": _to_float(r[14]) or 0,
                    "滚动3日完成": _to_float(r[16]) or 0,
                    "完成率": _to_float(r[17]),
                    "周达标情况": _norm(r[18]),
                    "当日上门量": _to_float(r[21]) or 0,
                    "当日外出量": _to_float(r[22]) or 0,
                    "当日线上见客量": _to_float(r[23]) or 0,
                    "总见客时长": _to_float(r[24]),
                    "见客合计": _to_float(r[25]) or 0,
                    "当日新增业绩": _to_float(r[27]) or 0,
                })

    wb.close()
    return result


def get_data():
    """带缓存的解析：文件 mtime 不变则复用上一次结果。"""
    global _cache
    try:
        mtime = os.path.getmtime(XLSX_PATH)
    except OSError:
        mtime = 0.0
    with _cache_lock:
        if _cache["data"] is not None and _cache["mtime"] == mtime and _cache["error"] is None:
            return _cache["data"]
    # 解析（可能抛异常，如文件被 Excel 占用）
    try:
        data = parse_workbook()
    except Exception as e:
        with _cache_lock:
            if _cache["data"] is not None:
                # 解析失败（多为文件被占用），回退到上次成功数据
                _cache["error"] = str(e)
                return _cache["data"]
            raise
    with _cache_lock:
        _cache["mtime"] = mtime
        _cache["data"] = data
        _cache["error"] = None
    return data


def build_meta():
    data = get_data()
    try:
        mtime = os.path.getmtime(XLSX_PATH)
        mtime_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
    except OSError:
        mtime_str = ""
    data = dict(data)
    data["meta"] = dict(data.get("meta", {}))
    data["meta"]["file_mtime"] = mtime_str
    data["meta"]["file_mtime_ts"] = mtime
    data["meta"]["server_time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return data


# ==================== 历史数据（每日快照） ====================
_hist_cache_lock = threading.Lock()
_hist_cache = {"sig": None, "data": []}
_saved_dates = set()  # 本进程已尝试保存的日期（幂等）


def save_daily_snapshot():
    """每天 23:00-23:59 自动把当天「在线看板.xlsx」复制到 历史数据/YYYY-MM-DD.xlsx；
    00:00-00:30 补存昨天。文件已存在则跳过。返回保存路径或 None。"""
    now = datetime.datetime.now()
    target = None
    if now.hour == 23:
        target = now.strftime("%Y-%m-%d")
    elif now.hour == 0 and now.minute < 30:
        target = (now - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    if not target or target in _saved_dates:
        return None
    try:
        os.makedirs(HISTORY_DIR, exist_ok=True)
        dst = os.path.join(HISTORY_DIR, target + ".xlsx")
        if not os.path.exists(dst):
            import shutil
            shutil.copy2(XLSX_PATH, dst)
            print("已保存历史快照:", dst)
        _saved_dates.add(target)
        return dst
    except Exception as e:
        print("[warn] 每日快照保存失败:", e)
        return None


def _daily_saver_loop():
    while True:
        try:
            save_daily_snapshot()
        except Exception:
            pass
        time.sleep(60)


def _hist_signature():
    """历史目录的签名：文件列表 + 各文件 mtime，用于缓存失效。"""
    try:
        files = sorted(
            f for f in os.listdir(HISTORY_DIR)
            if f.endswith(".xlsx") and not f.startswith("~$")
        )
    except OSError:
        return ("-",)
    try:
        mt = max(os.path.getmtime(os.path.join(HISTORY_DIR, f)) for f in files)
    except OSError:
        mt = 0.0
    return (len(files), round(mt, 1), files[-1] if files else "")


def get_history():
    """扫描 历史数据/*.xlsx，解析为 [{date, data:{stats,total,legal,meta}}]，按日期升序。"""
    sig = _hist_signature()
    with _hist_cache_lock:
        if _hist_cache["sig"] == sig:
            return _hist_cache["data"]
    items = []
    try:
        files = sorted(
            f for f in os.listdir(HISTORY_DIR)
            if f.endswith(".xlsx") and not f.startswith("~$")
        )
    except OSError:
        files = []
    for fn in files:
        date = fn[:-5]
        try:
            data = parse_workbook(os.path.join(HISTORY_DIR, fn))
        except Exception as e:
            print("[warn] 历史解析失败:", fn, e)
            continue
        data = dict(data)
        data["meta"] = dict(data.get("meta", {}))
        data["meta"]["date"] = date
        items.append({"date": date, "data": data})
    with _hist_cache_lock:
        _hist_cache["sig"] = sig
        _hist_cache["data"] = items
    return items


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            try:
                with open(INDEX_PATH, "r", encoding="utf-8") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            except FileNotFoundError:
                self._send(500, "index.html 缺失", "text/plain; charset=utf-8")
        elif path == "/api/data":
            try:
                payload = json.dumps(build_meta(), ensure_ascii=False)
                self._send(200, payload)
            except Exception as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
        elif path == "/api/history":
            try:
                payload = json.dumps(get_history(), ensure_ascii=False)
                self._send(200, payload)
            except Exception as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
        elif path == "/static/echarts.min.js":
            try:
                with open(ECHARTS_PATH, "rb") as f:
                    self._send(200, f.read(), "application/javascript; charset=utf-8")
            except FileNotFoundError:
                self._send(404, "echarts.min.js 缺失", "text/plain; charset=utf-8")
        elif path == "/favicon.ico":
            self._send(204, b"")
        else:
            self._send(404, "Not Found", "text/plain; charset=utf-8")

    def log_message(self, fmt, *args):
        pass  # 静默


def main():
    import socket
    port = 8765
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        ip = "127.0.0.1"
    print("=" * 50)
    print("  分公司活动量 · 在线数据看板")
    print("=" * 50)
    print("  本机访问:  http://localhost:%d" % port)
    print("  局域网访问: http://%s:%d" % (ip, port))
    print("  数据文件:  %s" % XLSX_PATH)
    print("  历史快照:  %s（每天23:59自动保存当天）" % HISTORY_DIR)
    print("  按 Ctrl+C 停止服务")
    print("=" * 50)
    # 每日 23:59 自动保存当天数据快照（服务运行期间生效）
    t = threading.Thread(target=_daily_saver_loop, daemon=True)
    t.start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
        server.shutdown()


if __name__ == "__main__":
    main()
