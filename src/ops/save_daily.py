# -*- coding: utf-8 -*-
# @version V0.3  @author <AUTHOR>  @created 2026-09-24
"""
save_daily.py —— 每天 23:59 主看板历史快照（HZDashBoardDailySave 任务负载）

2026-09-24 重构：旧逻辑仅在「源 xlsx 有更新」时才留存当天，且 history.json 只在
「数据变化→重建」时重算，二者叠加导致历史漏天。现改为调用 snapshot_daily.main()：
  1) 原 xlsx 归档（best-effort，保留历史兼容）
  2) capture_today       拍「已部署的 _deploy/data.json」为当日 JSON 快照
  3) build_history_json  重算 _deploy/history.json（含今日兜底，今天必在列）
  4) 发布阶段          走 bi_refresh.stage_publish()（V1.5：WorkBuddy 通道）
全程不依赖 WPS COM、不解析被锁定的源 xlsx，鲁棒性高。

V1.5（2026-10-12）：ROOT 不再写死 —— 交给 _portable 解析，换电脑自动认路。
"""
import os
import sys
import shutil
import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)   # 同目录：_portable / snapshot_daily
sys.path.insert(0, os.path.dirname(BASE))   # 根目录共享模块（迁移前兜底）

try:
    import _portable
except Exception:
    _portable = None

# V1.5：不再写死路径
ROOT = _portable.resolve_root() if _portable else os.path.dirname(BASE)
sys.path.insert(0, ROOT)   # 根目录共享模块：dashboard_server

from dashboard_server import HISTORY_DIR, XLSX_PATH  # noqa: E402

import snapshot_daily  # noqa: E402


def archive_xlsx():
    """原 save_daily 归档：把当前源 xlsx 复制到 历史数据/<date>.xlsx（best-effort）。

    仅作历史兼容留存；主历史已改为 JSON 日快照，源 xlsx 被占用/缺失时静默跳过，
    不影响主流程。
    """
    try:
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        os.makedirs(HISTORY_DIR, exist_ok=True)
        dst = os.path.join(HISTORY_DIR, today + ".xlsx")
        if not os.path.exists(XLSX_PATH):
            print("[warn] 源 xlsx 不存在，跳过归档")
            return
        if os.path.exists(dst) and os.path.getmtime(XLSX_PATH) <= os.path.getmtime(dst):
            return
        shutil.copy2(XLSX_PATH, dst)
        print("已归档源 xlsx:", dst)
    except Exception as e:
        print("[warn] 源 xlsx 归档失败（不影响 JSON 快照）:", e)


def main():
    archive_xlsx()
    return snapshot_daily.main()


if __name__ == "__main__":
    sys.exit(main())
