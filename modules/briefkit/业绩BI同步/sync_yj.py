# -*- coding: utf-8 -*-
"""
业绩（含续费）BI源数据 → 企微在线表格 自动同步
================================================
流程：下载BI源数据(周期月口径) → 与在线表逐格比对 → 有差异则写回在线表 →
      写入校验通过后 → 向配置的群里汇报已更新的数据。

用法：
    python sync_yj.py              # 单次同步（有差异才写入+汇报，无差异静默跳过）
    python sync_yj.py --loop 30    # 每30分钟循环同步（Ctrl+C 停止）
    python sync_yj.py --dry        # 只对比不写入不汇报（试运行）

依赖：仅 Python 标准库（urllib/zipfile/xml），无需安装任何第三方包。
调用链：wecom-cli（node 直调 wecom.js，绕开 shell 包装脚本的中文 JSON 篡改坑）。
"""

import argparse
import json
import os
import re
import subprocess
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ============================== 配置区 ==============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "sync_yj.log")
LOCK_FILE = os.path.join(BASE_DIR, "sync_yj.lock")

# 在线表格（业绩追踪表）
DOC_ID = "e3_Ae8A_gYvADACN1gvbrUgvT3SalKT9"
SHEET_TITLE = "业绩（含续费）"

# BI 源数据（帆软直连导出）。口径说明：
#   TYPE=1 周期月 / 2 自然月 —— 在线表用周期月
#   SELECT_MONTH  周期月份   —— 周年庆周期记为 2026-10
#   Q_DATE1       周期起始日 —— 2026-09-15
#   Q_DATE2       截止日     —— 每次运行自动取当天
# ⚠️ 换下一个 PK 周期时：改 SELECT_MONTH / Q_DATE1 两个值即可
BI_URL = ("https://bi.example.com/webroot/decision/view/report"
          "?viewlet=dggReport/caiwu/report_yj_d_v4/report_yj_d_xf_sr_v5_20260427.cpt")
BI_FIXED_PARAMS = {
    "biLoginNo": "100000",
    "format": "excel",
    "extype": "simple",
    "TYPE": "1",
    "SELECT_MONTH": "2026-10",
    "Q_DATE1": "2026-09-15",
}

# 汇报群（按会话名关键字匹配）
GROUPS = ["管理者大群", "分公司经理内部群"]

NCOLS = 15          # 在线表使用列数 A..O
READ_ROWS = 200     # 每次读取在线表行数
BATCH_ROWS = 12     # 每次写入调用的行数（控制单次 JSON 体积）
# 含时间戳的行：只要其他格子有变化就一并刷新，但不作为"是否有更新"的判断依据
TIMESTAMP_ROWS = {2, 4}
# 禁写行：R8 表头含团队手工定制的同期标注（BI 导出里是占位符），永不覆盖、不参与触发
NO_WRITE_ROWS = {8}
# ===================================================================

def _resolve_node_exe():
    """自动探测当前 node.exe：受管运行时升级后旧版本目录会被删除，
    硬编码版本号会导致 subprocess 报 FileNotFoundError(WinError 2)。"""
    import glob
    base = r"<HOME>/.workbuddy/binaries/node/versions"
    cands = sorted(glob.glob(base + "/*/node.exe"), reverse=True)
    for p in cands:
        if os.path.exists(p):
            return p
    # 兜底：PATH 里的 node
    return "node"


NODE_EXE = _resolve_node_exe()
WECOM_JS = (r"<HOME>/.workbuddy/binaries/node/cli-connector-packages"
            r"/node_modules/@wecom/cli/bin/wecom.js")

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def log(msg, level="INFO"):
    line = f"[{datetime.now().strftime('%m-%d %H:%M:%S')}] [{level}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# ------------------------- wecom-cli 调用 -------------------------

def _parse_wecom_resp(text):
    """剥掉 security_notice 等包装，解析业务 JSON。"""
    i, j = text.find("{"), text.rfind("}")
    if i < 0 or j < 0:
        raise RuntimeError(f"wecom-cli 返回非 JSON: {text[:200]}")
    return json.loads(text[i:j + 1])


def wecom_call(args, payload=None, timeout=120):
    """node 直调 wecom.js。payload 为 dict 时以 --json <content> 传入。"""
    cmd = [NODE_EXE, WECOM_JS] + list(args)
    if payload is not None:
        cmd += ["--json", json.dumps(payload, ensure_ascii=False)]
    p = subprocess.run(cmd, capture_output=True, timeout=timeout)
    out = p.stdout.decode("utf-8", errors="replace")
    err = p.stderr.decode("utf-8", errors="replace")
    if p.returncode != 0 and not out.strip():
        raise RuntimeError(f"wecom-cli 退出码{p.returncode}: {err[:300]}")
    d = _parse_wecom_resp(out)
    if isinstance(d, dict) and d.get("error"):
        raise RuntimeError(f"wecom-cli 业务错误: {d['error']}")
    if isinstance(d, dict) and d.get("errcode", 0) not in (0, None):
        raise RuntimeError(f"wecom-cli errcode={d.get('errcode')}: {d.get('errmsg')}")
    return d


# ------------------------- BI 源数据下载与解析 -------------------------

def colnum(ref):
    m = re.match(r"([A-Z]+)", ref)
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n


def fetch_bi_bytes(retries=2):
    today = datetime.now().strftime("%Y-%m-%d")
    params = dict(BI_FIXED_PARAMS)
    params["Q_DATE2"] = today
    qs = "&".join(f"{k}={v}" for k, v in params.items())
    url = BI_URL + "&" + qs
    last_err = None
    for i in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=120) as r:
                data = r.read()
            if len(data) < 1000:
                raise RuntimeError(f"响应过小({len(data)}B)，疑似未登录或被拦截")
            log(f"BI 源数据下载成功: {len(data):,} bytes (Q_DATE2={today})")
            return data
        except Exception as e:
            last_err = e
            log(f"BI 下载失败({i + 1}/{retries + 1}): {e}", "WARN")
            time.sleep(3)
    raise RuntimeError(f"BI 源数据下载失败: {last_err}")


def parse_bi(data):
    """解析 BI xlsx → {(row, col): value}，列号-1（BI 从B列起，在线表从A列起）。"""
    z = zipfile.ZipFile(__import__("io").BytesIO(data))
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            ss.append("".join(t.text or "" for t in
                              si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {rel.get("Id"): rel.get("Target") for rel in rels}
    sh = wb.find("m:sheets", NS)[0]
    tgt = rid2t.get(sh.get(RID))
    path = tgt if tgt.startswith("xl/") else "xl/" + tgt
    sd = ET.fromstring(z.read(path)).find("m:sheetData", NS)
    grid = {}
    for row in sd.findall("m:row", NS):
        rn = int(row.get("r"))
        for c in row.findall("m:c", NS):
            t, v, isn = c.get("t"), c.find("m:v", NS), c.find("m:is", NS)
            if t == "s" and v is not None:
                val = ss[int(v.text)]
            elif isn is not None:
                val = "".join(x.text or "" for x in isn.iter(
                    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            elif v is not None:
                val = v.text
            else:
                val = ""
            val = str(val).strip()
            if val:
                grid[(rn, colnum(c.get("r")) - 1)] = val
    return grid


# ------------------------- 在线表读取 -------------------------

def resolve_sheet_id():
    d = wecom_call(["sheet", "get"], payload={"docid": DOC_ID})
    for s in d.get("sheets", []):
        if s.get("title", "").strip() == SHEET_TITLE:
            return s["sheet_id"]
    raise RuntimeError(f"在线表中未找到 sheet「{SHEET_TITLE}」")


def read_online(sheet_id):
    """读在线表 → {(row, col): value}，value 取 cell_value 里的 text/number。"""
    d = wecom_call(["sheet", "ranges", "get"], payload={
        "docid": DOC_ID, "sheet_id": sheet_id, "mode": "default",
        "range": f"A1:O{READ_ROWS}"}, timeout=180)
    grid = {}
    for ri, row in enumerate(d.get("grid_data", {}).get("rows", []), 1):
        for ci, cell in enumerate(row.get("values", []), 1):
            cvv = cell.get("cell_value") or {}
            if "text" in cvv and str(cvv["text"]).strip():
                grid[(ri, ci)] = str(cvv["text"]).strip()
            elif "number" in cvv:
                num = cvv["number"]
                grid[(ri, ci)] = str(int(num)) if float(num).is_integer() else str(num)
    return grid


# ------------------------- 比对与写入 -------------------------

def _as_num(v):
    try:
        if v in ("-", "--", ""):
            return None
        return float(v)
    except Exception:
        return None


def values_equal(a, b):
    na, nb = _as_num(a), _as_num(b)
    if na is not None and nb is not None:
        return abs(na - nb) <= 0.005 + abs(na + nb) * 1e-7
    # 0 与空等价（接口对零值格的读写回传不一致）
    za = _as_num(a) == 0 or a == ""
    zb = _as_num(b) == 0 or b == ""
    if za and zb:
        return True
    return a == b


def compute_changes(bi, online):
    """返回需要写入的行号集合与每行新值。触发条件：排除时间戳行/禁写行后有差异。"""
    changed_rows = {}
    max_row = max([r for r, _ in bi] + [r for r, _ in online] + [0])
    for rn in range(1, max_row + 1):
        if rn in NO_WRITE_ROWS:
            continue
        new_vals, has_diff = {}, False
        for ci in range(1, NCOLS + 1):
            nv = bi.get((rn, ci), "")
            ov = online.get((rn, ci), "")
            new_vals[ci] = nv
            if not values_equal(nv, ov):
                has_diff = True
        if has_diff:
            changed_rows[rn] = new_vals
    trigger_rows = [r for r in changed_rows if r not in TIMESTAMP_ROWS]
    return changed_rows, bool(trigger_rows)


def _cell_payload(value):
    n = _as_num(value)
    if value and n is not None:
        return {"cell_value": {"number": n}, "data_type": "NUMBER"}
    return {"cell_value": {"text": value}, "data_type": "TEXT"}


def contiguous_chunks(rows, size):
    """把行号列表切成连续行段块（段内必须逐行连续，否则 start_row 会错位）。"""
    chunks, cur = [], []
    for r in rows:
        if cur and r != cur[-1] + 1:
            chunks.extend(cur[i:i + size] for i in range(0, len(cur), size))
            cur = []
        cur.append(r)
    if cur:
        chunks.extend(cur[i:i + size] for i in range(0, len(cur), size))
    return chunks


def write_changes(sheet_id, changed_rows):
    rows = sorted(changed_rows)
    batches = contiguous_chunks(rows, BATCH_ROWS)
    for bi_, batch in enumerate(batches, 1):
        grid_rows = []
        for rn in batch:
            grid_rows.append({"values": [_cell_payload(changed_rows[rn].get(ci, ""))
                                         for ci in range(1, NCOLS + 1)]})
        payload = {"docid": DOC_ID, "sheet_id": sheet_id,
                   "grid_data": {"start_row": batch[0] - 1, "start_column": 0,
                                 "rows": grid_rows}}
        wecom_call(["sheet", "contents", "update"], payload=payload, timeout=180)
        log(f"写入批次 {bi_}/{len(batches)}: 行 {batch[0]}-{batch[-1]} ✅")
        time.sleep(0.4)


# ------------------------- 汇报（简洁版） -------------------------

def fmt_wan(v):
    n = _as_num(v)
    if n is None:
        return v
    return f"{n / 10000:.1f}万"


def fmt_pct(v):
    n = _as_num(v)
    return f"{n * 100:.1f}%" if n is not None else str(v)


def build_report(bi):
    def row_by_first(name, lo=1, hi=200):
        for rn in range(lo, hi):
            if bi.get((rn, 1)) == name:
                return rn
        return None

    period = bi.get((2, 1), "")
    g_rn = row_by_first("集团总计")
    hz_rn = row_by_first("分公司", lo=17, hi=40)
    lines = ["📈 业绩（含续费）已自动同步更新", ""]
    if period:
        lines += [f"📅 {period}", ""]
    if g_rn:
        g = {ci: bi.get((g_rn, ci), "") for ci in range(1, NCOLS + 1)}
        lines.append(f"🏢 集团总计：完成 {fmt_wan(g[3])} / 目标 {fmt_wan(g[2])}"
                     f"（{fmt_pct(g[5])}）")
        lines.append(f"　· 当日业绩 {fmt_wan(g[4])} ｜ 高价值客户 {g[8]}/{g[7]}")
    if hz_rn:
        h = {ci: bi.get((hz_rn, ci), "") for ci in range(1, NCOLS + 1)}
        rank = h[6]
        rank = f"第{int(float(rank))}名" if _as_num(rank) is not None else rank
        lines.append(f"🏙 分公司：完成 {fmt_wan(h[3])}（{fmt_pct(h[5])}）｜ 排名 {rank}")
        lines.append(f"　· 当日 {fmt_wan(h[4])} ｜ 高价值客户 {h[8]}/{h[7]}"
                     f" ｜ 当日退款 {fmt_wan(h[12])}")
    note = bi.get((5, 1), "")
    m = re.search(r"前三的分公司是：(.*?)[；;]。?\s*后三的分公司是：(.*?)[；;。]", note)
    if m:
        lines.append("")
        lines.append(f"🥇 完成率前三：{m.group(1).replace('公司', '')}")
        lines.append(f"🔻 后三：{m.group(2).replace('公司', '')}")
    lines.append("")
    lines.append("数据源：集团BI·业绩(含续费)R-v5 ｜ 已同步至在线表格")
    return "\n".join(lines), None


def send_report(text, img_path=None):
    d = wecom_call(["message", "aibot", "sessions", "list"])
    sessions = d.get("sessions", [])
    ok = []
    for key in GROUPS:
        target = next((s for s in sessions if key in str(s.get("chat_name", ""))), None)
        if not target:
            log(f"会话列表中未找到群「{key}」，跳过", "WARN")
            continue
        payload = {"chat_id": target["chat_id"], "msgtype": "markdown",
                   "markdown": {"content": text}}
        wecom_call(["message", "aibot", "send"], payload=payload)
        ok.append(target["chat_name"])
        log(f"已汇报 → {target['chat_name']} ✅")
        time.sleep(0.5)
    return ok


# ------------------------- 主流程 -------------------------

HEADER_CUSTOM_C13 = "去年同期业绩\n2025-09-15 - 2025-09-28"  # 团队手工定制的同期标注


def full_grid(bi):
    """全量行数据(1..BI最大行)，R8 表头 c13 还原为手工定制文本。"""
    max_row = max(r for r, _ in bi)
    grid = {}
    for rn in range(1, max_row + 1):
        grid[rn] = {ci: bi.get((rn, ci), "") for ci in range(1, NCOLS + 1)}
    if 8 in grid:
        grid[8][13] = HEADER_CUSTOM_C13
    return grid


def one_cycle(dry=False, notify=True):
    t0 = time.time()
    sheet_id = resolve_sheet_id()
    log(f"目标 sheet: {SHEET_TITLE} (id={sheet_id})")

    bi = parse_bi(fetch_bi_bytes())
    online = read_online(sheet_id)
    log(f"BI 数据 {len(bi)} 格 / 在线表 {len(online)} 格，开始比对…")

    changed_rows, trigger = compute_changes(bi, online)
    if not trigger:
        log(f"无数据变化（仅时间戳行差异），跳过写入与汇报。耗时 {time.time() - t0:.0f}s")
        return "SKIP"
    log(f"检测到数据更新：{len(changed_rows)} 行有差异"
        + ("（dry 试运行，不写入）" if dry else ""))

    if dry:
        for rn in sorted(changed_rows)[:10]:
            print(f"  R{rn}: " + " | ".join(
                f"c{ci}:{changed_rows[rn][ci][:20]}" for ci in (1, 2, 3, 5)
                if changed_rows[rn].get(ci)))
        return "DRY"

    write_changes(sheet_id, changed_rows)

    # 写后校验：重新读取，确认除时间戳行外无差异
    online2 = read_online(sheet_id)
    changed2, trigger2 = compute_changes(bi, online2)
    if trigger2:
        log(f"写后校验仍有 {len([r for r in changed2 if r not in TIMESTAMP_ROWS])} 行差异，重写一次…", "WARN")
        write_changes(sheet_id, changed2)
        online2 = read_online(sheet_id)
        _, trigger3 = compute_changes(bi, online2)
        if trigger3:
            raise RuntimeError("写后校验失败（两次写入后仍有差异），本次不汇报")

    if not notify:
        # 静默模式（供 PK 汇报流水线调用）：只同步不汇报，由调用方负责后续汇报
        log(f"同步完成（静默模式，跳过简报）: 写入 {len(changed_rows)} 行,"
            f" 耗时 {time.time() - t0:.0f}s")
        return "SYNCED"

    text, img_path = build_report(bi)
    sent = send_report(text, img_path)
    log(f"同步完成: 写入 {len(changed_rows)} 行, 汇报 {len(sent)} 个群,"
        f" 耗时 {time.time() - t0:.0f}s")
    return "SYNCED"


def repair_cycle():
    """全量重写修复模式：把 BI 数据 1..最大行 全部写入（含表头定制还原），校验后汇报。"""
    t0 = time.time()
    sheet_id = resolve_sheet_id()
    bi = parse_bi(fetch_bi_bytes())
    grid = full_grid(bi)
    log(f"修复模式：全量重写 1-{max(grid)} 行 …")
    write_changes(sheet_id, grid)
    online2 = read_online(sheet_id)
    changed2, trigger2 = compute_changes(bi, online2)
    if trigger2:
        log(f"修复后校验仍有 {len(changed2)} 行差异，再写一次…", "WARN")
        write_changes(sheet_id, changed2)
        online2 = read_online(sheet_id)
        _, t3 = compute_changes(bi, online2)
        if t3:
            raise RuntimeError("修复后校验仍失败，请人工检查在线表")
    log(f"修复校验通过 ✅ 耗时 {time.time() - t0:.0f}s")
    text, img_path = build_report(bi)
    sent = send_report(text, img_path)
    log(f"修复完成并汇报 {len(sent)} 个群")
    return "REPAIRED"


def _pid_alive(pid):
    """Windows 下用 OpenProcess 判活，不依赖已删除的 psutil。"""
    if not pid or pid <= 0:
        return False
    try:
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        STILL_ACTIVE = 259
        k32 = ctypes.windll.kernel32
        h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h:
            return False           # 进程不存在/无权限 → 视为已死
        try:
            code = ctypes.c_ulong()
            if k32.GetExitCodeProcess(h, ctypes.byref(code)):
                return code.value == STILL_ACTIVE
            return False
        finally:
            k32.CloseHandle(h)
    except Exception:
        return True                # 判不出来时保守当作活着，避免误删活锁


def acquire_lock():
    if os.path.exists(LOCK_FILE):
        age = time.time() - os.path.getmtime(LOCK_FILE)
        pid = 0
        try:
            with open(LOCK_FILE, "r", encoding="utf-8") as f:
                pid = int((f.read() or "0").strip() or 0)
        except Exception:
            pass
        # 锁内PID 仍存活 → 真并发，退出；PID 已死（被 SIGTERM/强杀留下的死锁）→ 直接覆盖
        if age < 900 and (pid == os.getpid() or _pid_alive(pid)):
            log(f"另一个同步进程运行中（pid={pid}, 锁 {age:.0f}s），退出", "WARN")
            sys.exit(0)
        log(f"发现过期/死锁（pid={pid}, 锁 {age:.0f}s），覆盖", "WARN")
    with open(LOCK_FILE, "w", encoding="utf-8") as f:
        f.write(str(os.getpid()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--loop", type=int, default=0, help="循环间隔分钟数（0=单次）")
    ap.add_argument("--dry", action="store_true", help="只对比不写入不汇报")
    ap.add_argument("--repair", action="store_true",
                    help="修复模式：全量重写在线表（用于写入错位后恢复）")
    args = ap.parse_args()

    if args.repair:
        acquire_lock()
        try:
            repair_cycle()
        finally:
            if os.path.exists(LOCK_FILE):
                os.remove(LOCK_FILE)
        return

    if args.loop > 0:
        acquire_lock()
        log(f"进入循环模式：每 {args.loop} 分钟同步一次，Ctrl+C 停止")
        while True:
            try:
                one_cycle(dry=args.dry)
            except Exception as e:
                log(f"本轮同步失败: {e}", "ERROR")
            try:
                time.sleep(args.loop * 60)
            except KeyboardInterrupt:
                log("收到停止信号，退出循环")
                break
    else:
        acquire_lock()
        try:
            one_cycle(dry=args.dry)
        finally:
            if os.path.exists(LOCK_FILE):
                os.remove(LOCK_FILE)


if __name__ == "__main__":
    main()
