# -*- coding: utf-8 -*-
"""
最简看板更新（一步到位）
========================
用途：用户说"更新看板"时，用最短路跑完：重建快照 → 触发云端部署 → 一句话回报。
不做长轮询等待、不做二次校验（那是最费时的部分）。

逻辑：
  1. 单趟比对 源Excel / 本地快照 / 云端 mtime（仅几个 urllib，约 1-2s）
     - 三者一致 → 直接结束，零积分
  2. 重建 _deploy 快照（build_snapshot.build，一次性重建三看板）
  3. 触发云端部署自动化：直接 UPDATE workbuddy.db 该 ID 的
     next_run_at=now 且 status='ACTIVE'（PAUSED 调度器不接单，必须一并激活）
  4. 打印链接，提示"约 1 分钟生效，刷新即可"

用法：
  python quick_update.py
"""
import os
import sys
import json
import time
import sqlite3
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ⚠️ 2026-09-12 换链：原 CloudStudio 通道（ef627954…bj4.agentos-app.net）的部署工具
# workbuddy_cloudstudio_deploy 已从客户端移除，该链接永久无法更新；改用 sites 通道
# 新建的稳定应用（appId wbapp_Jc5WgPGEeqrk5b5tcLZ3RT）。新链接长期固定。
CLOUD_BASE = "https://dashboard.example.com"
APP_ID = "wbapp_Jc5WgPGEeqrk5b5tcLZ3RT"

DB = os.path.expanduser("~/.workbuddy/workbuddy.db")
AUTOMATION_ID = "automation-1786499389829"
EPS = 2.0  # 秒；mtime 噪声容差

BOARDS = [
    ("活动量", "分公司活动量-在线看板.xlsx", "_deploy/data.json", "data.json"),
    ("三力", "分公司-员工三力明细.xlsx", "_deploy/sanli_data.json", "sanli_data.json"),
    ("人员", "分公司-人员数据分析路径.xlsx", "_deploy/ry_data.json", "ry_data.json"),
]


def meta_ts(rel):
    try:
        with open(os.path.join(BASE_DIR, rel), encoding="utf-8") as f:
            return json.load(f).get("meta", {}).get("file_mtime_ts", 0.0)
    except Exception:
        return None


def cloud_ts(cloud_file):
    url = f"{CLOUD_BASE}/{cloud_file}?cb={int(time.time()*1000)}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "quicksync/1"})
        with urllib.request.urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode("utf-8")).get("meta", {}).get("file_mtime_ts", 0.0)
    except Exception:
        return 0.0  # 云端不可达 → 保守视为需部署


def trigger():
    now = int(time.time() * 1000)
    con = sqlite3.connect(DB)
    cur = con.cursor()
    # 必须一并置 ACTIVE，否则 PAUSED 状态下调度器不接单
    cur.execute(
        "UPDATE automations SET next_run_at=?, updated_at=?, status='ACTIVE' "
        "WHERE id=? AND deleted_at IS NULL",
        (now, now, AUTOMATION_ID),
    )
    rc = cur.rowcount
    con.commit()
    con.close()
    return rc


def main():
    need = False
    for name, src, local, cloud in BOARDS:
        src_path = os.path.join(BASE_DIR, src)
        src_ts = os.path.getmtime(src_path) if os.path.exists(src_path) else 0
        loc = meta_ts(local) or 0
        if src_ts - loc > EPS:
            print(f"  · {name}：源 Excel 有更新，需重建")
            need = True
        else:
            c = cloud_ts(cloud)
            if (loc - c) > EPS:
                print(f"  · {name}：云端滞后，需部署")
                need = True
            else:
                print(f"  · {name}：已是最新，跳过")
    if not need:
        print("✅ 所有看板均已是最新，无需部署（零积分）")
        return 0

    print("↻ 重建 _deploy 快照 …")
    sys.path.insert(0, BASE_DIR)
    import build_snapshot
    build_snapshot.build()

    rc = trigger()
    print(f"🚀 已触发云端部署（自动化 {AUTOMATION_ID}，status=ACTIVE，rowcount={rc}）")
    print(f"   约 1 分钟内生效，刷新即可：{CLOUD_BASE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
