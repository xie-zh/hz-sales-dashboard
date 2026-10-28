# -*- coding: utf-8 -*-
"""
自动监测「三力 & 有效人力」群文件，出现新日期版本即自动解析 + 推送。

原理
----
企业微信对机器人关闭了群消息/群文件接口（errcode 853006），**读不到群消息体**。
但企微客户端会把群文件自动缓存到本地：
    <DATA_CACHE>/<账号>/Cache/File/<年-月>/<原文件名>
因此改为「本地缓存轮询」：
    扫描缓存 -> 匹配文件名含「三力」且含「有效人力」的 .xlsx
    -> 取文件名中日期戳 (20YYMMDD) 最大的一份
    -> 与 watch_state.json 比对，日期戳更新才处理（同日多版本不重发）
    -> 复制为标准输入 src_latest.xlsx -> 跑 run_hero.sh（渲染 + 发 4 条到 3C 群）

用法
----
  python watch_hero.py            # 检测；有新文件则跑完整流程并发送（默认）
  python watch_hero.py --check    # 只检测不发送（dry-run，用于调试）
  python watch_hero.py --force    # 忽略 state，按当前最新文件强制重跑一次

输出（供自动化/人工读取，单行摘要 + 明细）
  STATUS=NO_NEW   ...            没有比上次更新的文件
  STATUS=SENT     ...            已解析并发送成功
  STATUS=FAILED   ...            处理或发送失败（下次会自动重试，最多 3 次）
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime

DIR = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(DIR)                       # .../数据文件/三力追踪
SRC_LATEST = os.path.join(BASE, "src_latest.xlsx")
META = os.path.join(DIR, "src_latest.meta.json")
STATE = os.path.join(DIR, "watch_state.json")
RUN_SH = os.path.join(DIR, "run_hero.sh")

CACHE_GLOB = "<DATA_CACHE>/*/Cache/File/*"
NAME_RE = re.compile(r"三力.*有效人力|有效人力.*三力")
DATE_RE = re.compile(r"(20\d{2})(\d{2})(\d{2})")
MAX_RETRY = 3


def ts(fmt="%Y-%m-%d %H:%M:%S"):
    return datetime.now().strftime(fmt)


def scan():
    """扫描所有企微账号缓存，返回匹配文件的候选列表"""
    out = []
    for d in glob.glob(CACHE_GLOB):
        if not os.path.isdir(d):
            continue
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for f in names:
            if not f.lower().endswith(".xlsx") or f.startswith("~$"):
                continue
            if not NAME_RE.search(f):
                continue
            m = DATE_RE.search(f)
            if not m:
                continue
            p = os.path.join(d, f)
            try:
                st = os.stat(p)
            except OSError:
                continue
            out.append({
                "path": p, "name": f, "dir": d,
                "ymd": "".join(m.groups()),
                "size": st.st_size, "mtime": st.st_mtime,
                "mtime_str": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            })
    out.sort(key=lambda x: (x["ymd"], x["mtime"]))
    return out


def load_state():
    if os.path.exists(STATE):
        try:
            return json.load(open(STATE, encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_state(st):
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)


def main():
    args = set(sys.argv[1:])
    check_only = "--check" in args
    force = "--force" in args

    cands = scan()
    if not cands:
        print("STATUS=NO_NEW | 缓存中未匹配到任何「三力&有效人力」xlsx")
        print(f"  扫描范围: {CACHE_GLOB}")
        return 0

    best = cands[-1]
    st = load_state()
    last_ymd = str(st.get("last_ymd", "") or "")
    retry = st.get("retry", {}) or {}

    print(f"最新候选 : {best['name']}")
    print(f"  路径   : {best['path']}")
    print(f"  日期戳 : {best['ymd']}  大小 {best['size']} B  落盘 {best['mtime_str']}")
    print(f"上次处理 : last_ymd={last_ymd or '(无)'}  "
          f"last_sent_at={st.get('last_sent_at', '(无)')}")

    is_new = bool(best["ymd"]) and best["ymd"] > last_ymd
    if not is_new and not force:
        if best["ymd"] == last_ymd:
            print(f"STATUS=NO_NEW | 当前最新 {best['ymd']} 已处理过，无需重发")
        else:
            print(f"STATUS=NO_NEW | 缓存中最新 {best['ymd']} 早于已处理 {last_ymd}（老文件回灌，忽略）")
        return 0

    if retry.get(best["ymd"], 0) >= MAX_RETRY:
        print(f"STATUS=FAILED | {best['ymd']} 已连续失败 {MAX_RETRY} 次，已停止自动重试（请人工检查）")
        return 3

    if check_only:
        print(f"STATUS=DRY_RUN | 检测到新文件 {best['name']}，未发送（--check 模式）")
        return 0

    # ---- 1) 复制为标准输入 + 写元数据 ----
    shutil.copy2(best["path"], SRC_LATEST)
    meta = {
        "src_name": best["name"],
        "src_path": best["path"],
        "size": best["size"],
        "src_mtime": best["mtime_str"],
        "copied_at": ts(),
    }
    with open(META, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print(f"[1/3] 已复制源文件 -> {SRC_LATEST}")

    # ---- 2) 渲染 + 发送 ----
    print("[2/3] 执行 run_hero.sh（渲染 + 发送 4 条到 3C 群）...")
    p = subprocess.run(["bash", RUN_SH], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    out = (p.stdout or "") + (p.stderr or "")
    tail = "\n".join(out.strip().splitlines()[-14:])
    print(tail)

    all_ok = "ALL_SUCCESS = True" in out
    if not all_ok:
        retry[best["ymd"]] = retry.get(best["ymd"], 0) + 1
        st.update({"last_ymd": st.get("last_ymd", ""), "retry": retry,
                   "last_error_at": ts(), "last_error_ymd": best["ymd"]})
        save_state(st)
        print(f"[3/3] STATUS=FAILED | 发送未全部成功（第 {retry[best['ymd']]}/{MAX_RETRY} 次），"
              f"将自动重试")
        return 3

    # ---- 3) 记录状态 ----
    try:
        rep = json.load(open(os.path.join(DIR, "report_state.json"), encoding="utf-8"))
    except Exception:
        rep = {}
    st.update({
        "last_ymd": best["ymd"],
        "last_file": best["name"],
        "last_path": best["path"],
        "last_size": best["size"],
        "last_sent_at": ts(),
        "retry": {},
        "last_report": {
            "cycle": rep.get("cycle"),
            "done100": [x["name"] for x in rep.get("done100", [])],
            "done50": [x["name"] for x in rep.get("done50", [])],
            "gap100": rep.get("gap100", []),
            "gap50": [(x["name"], x["gap"]) for x in rep.get("gap50", [])],
        },
    })
    save_state(st)

    print(f"[3/3] 状态已记录 last_ymd={best['ymd']}")
    print(f"STATUS=SENT | {best['name']} | {ts()}")
    print(f"  周期     : {st['last_report']['cycle']}")
    print(f"  达成100% : {st['last_report']['done100']}")
    print(f"  达成50%  : {st['last_report']['done50']}")
    print(f"  差距100  : {st['last_report']['gap100']}")
    print(f"  差距50   : {st['last_report']['gap50']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
