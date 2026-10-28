# -*- coding: utf-8 -*-
# @version V0.3  @author <AUTHOR>  @created 2026-09-24
"""
snapshot_daily.py —— 每天 23:59 主看板历史快照编排（HZDashBoardDailySave 任务负载）
=============================================================================
替换旧 save_daily.py：旧逻辑只在「源 xlsx 有更新」时才留存当天，且 history.json
只在「数据变化→重建」时重算，二者叠加导致历史漏天。

本脚本每天 23:59 跑一次，全程不依赖 WPS COM、不解析被锁定的源 xlsx：
  1) capture_today       拍「已部署的 _deploy/data.json」为当日 JSON 快照
  2) build_history_json  重算 _deploy/history.json（含今日兜底，保证今天必在列）
  3) hist_snapshot       四模块（ry/sanli/dy/kuaiji）JSON 快照，无变化日也会补录当日
  4) 发布阶段            走 bi_refresh.stage_publish()（V1.5：WorkBuddy 通道；
                        过渡期可镜像到腾讯云）

V1.5（2026-10-12）：
  · ROOT 不再写死 —— 交给 _portable 解析，换电脑/换路径自动认路。
  · 发布阶段复用 bi_refresh，避免两处各写一份发布逻辑（以前就是复制粘贴）。
"""
import os
import sys
import json
import shutil
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)   # 同目录：_portable / bi_refresh
sys.path.insert(0, os.path.dirname(BASE_DIR))   # 根目录共享模块（迁移前的兜底）

try:
    import _portable
except Exception:
    _portable = None

# V1.5：不再写死路径。_portable 解析顺序：_paths.json → 环境变量 → 特征文件 → 兜底
ROOT = _portable.resolve_root() if _portable else os.path.dirname(BASE_DIR)
sys.path.insert(0, ROOT)       # 根目录共享模块：history_builder / hist_snapshot / dashboard_server

import history_builder as hb
import hist_snapshot as hs

DEPLOY = os.path.join(ROOT, "_deploy")
COS_CFG = os.path.join(ROOT, "_cos_config.json")


def stage_publish():
    """发布阶段：直接复用 bi_refresh 的实现（WorkBuddy 通道 + 过渡期镜像）。"""
    try:
        import bi_refresh
        return bi_refresh.stage_publish()
    except Exception as e:
        print("发布阶段失败（不影响历史快照落盘）:", str(e)[:200])
        return False


def mirror_backup():
    """把当日的关键产物镜像到「迁移配置.json」里的 backup_dir。

    腾讯云 COS 备份随账号收回而失效，这里给一个不依赖任何云账号的替代：
    只备份「小而关键」的东西 —— history.json、四个模块数据、_deploy/hist/ 日快照。
    体积小、恢复快；84MB 的源 xlsx 归档不重复搬（本地 历史数据\\ 已有）。
    """
    try:
        import _migrate
        bdir = str(_migrate.load(ROOT).get("backup_dir") or "").strip()
    except Exception:
        bdir = ""
    if not bdir:
        print("[备份] 未配置 backup_dir（迁移配置.json）→ 跳过镜像")
        return False
    day = datetime.datetime.now().strftime("%Y-%m-%d")
    dst = os.path.join(bdir, day)
    try:
        os.makedirs(dst)
    except Exception:
        pass
    n = 0
    for name in ("history.json", "data.json", "sanli_data.json", "ry_data.json",
                 "ry_dy_data.json", "ry_kuaiji_data.json"):
        src = os.path.join(DEPLOY, name)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(dst, name))
            n += 1
    hs_dir = os.path.join(DEPLOY, "hist")
    if os.path.isdir(hs_dir):
        d2 = os.path.join(dst, "hist")
        try:
            shutil.copytree(hs_dir, d2, dirs_exist_ok=True)
        except TypeError:                       # 3.7 及以下
            if not os.path.isdir(d2):
                shutil.copytree(hs_dir, d2)
        except Exception as e:
            print("[备份] hist 目录镜像失败:", str(e)[:80])
        n += 1
    print("[备份] 已镜像 %d 项 → %s" % (n, dst))
    return n > 0


def main():
    print("=== 每日历史快照 %s ===" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    # 1. 拍当日活动量快照
    hb.capture_today()
    # 2. 重算 history.json（含今日兜底）
    hb.build_history_json(ensure_today=True)
    # 3. 四模块 JSON 快照（ry/sanli/dy/kuaiji），无变化日也会补录当日 → 消除漏天
    try:
        hs.snapshot_now()
    except Exception as e:
        print("  [warn] 四模块快照失败:", e)
    # 4. 本地镜像备份（替代原来的腾讯云 COS 备份）
    mirror_backup()
    # 5. 发布（WorkBuddy 通道；过渡期镜像腾讯云）
    stage_publish()
    print("=== 每日历史快照完成 ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
