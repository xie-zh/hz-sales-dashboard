# -*- coding: utf-8 -*-
"""
Excel 文件监听 → 即时同步云端看板（常驻版 · 可选）
==================================================
监控 **全部 3 个源表**（活动量 / 员工三力 / 人员数据分析路径），
任一被保存 → 立即重建 _deploy 五页快照 → 触发云端部署自动化 → 链接不变。

说明：
  · 本脚本是「常驻窗口」版，双击 `启动Excel监听同步.bat` 开启，Ctrl+C 停止；
  · 若不想开窗口，用计划任务版 `check_and_deploy.py`（每 N 分钟单次检查）即可，
    两者可共存，重复触发无害（只会把部署任务提前到当前时刻）。
  · 触发部署按 **自动化 ID** 定位（旧版按名称匹配，名称已变 → 触发无效）。

运行: python watch_and_deploy.py  （保持常驻）
"""
import os
import time
import sqlite3
import sys
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB = os.path.expanduser("~/.workbuddy/workbuddy.db")
# 按 ID 定位（名称可能变，ID 不变）
AUTOMATION_ID = "automation-1786499389829"
LOG_FILE = os.path.join(BASE_DIR, "_watch.log")
CHECK_INTERVAL = 3        # 轮询间隔（秒）
COOLDOWN = 30             # 触发后冷却（秒）

SOURCES = [
    "分公司活动量-在线看板.xlsx",
    "分公司-员工三力明细.xlsx",
    "分公司-人员数据分析路径.xlsx",
]


def log(msg):
    line = "[%s] %s" % (datetime.datetime.now().strftime("%m-%d %H:%M:%S"), msg)
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def mtimes():
    out = {}
    for s in SOURCES:
        p = os.path.join(BASE_DIR, s)
        out[s] = os.path.getmtime(p) if os.path.exists(p) else 0.0
    return out


def trigger_deploy():
    """把云端同步自动化设为立即执行（按 ID）"""
    try:
        conn = sqlite3.connect(DB, timeout=10)
        cur = conn.cursor()
        now_ms = int(time.time() * 1000)
        cur.execute(
            "UPDATE automations SET next_run_at=?, updated_at=? "
            "WHERE id=? AND deleted_at IS NULL",
            (now_ms, now_ms, AUTOMATION_ID),
        )
        conn.commit()
        n = cur.rowcount
        conn.close()
        log("已触发云端部署任务 (rowcount=%d)" % n)
        return n > 0
    except Exception as e:
        log("触发部署失败: %s" % e)
        return False


def main():
    log("=" * 52)
    log("  Excel 监听 → 即时同步云端（全部板块）")
    for s in SOURCES:
        log("  监控: %s" % s)
    log("  轮询: 每 %ds | 冷却 %ds" % (CHECK_INTERVAL, COOLDOWN))
    log("=" * 52)

    missing = [s for s in SOURCES if not os.path.exists(os.path.join(BASE_DIR, s))]
    for s in missing:
        log("⚠️ 找不到源表: %s" % s)

    last = mtimes()
    last_trigger = 0

    while True:
        try:
            cur = mtimes()
            hit = [s for s in SOURCES if cur.get(s, 0) != last.get(s, 0)]
            if hit:
                for s in hit:
                    log("📄 检测到更新: %s" % s)
                last = cur
                now = time.time()
                if now - last_trigger < COOLDOWN:
                    log("  （冷却期内，跳过重建）")
                else:
                    try:
                        sys.path.insert(0, BASE_DIR)
                        import build_snapshot
                        build_snapshot.build()
                    except Exception as e:
                        log("生成快照失败: %s" % e)
                    if trigger_deploy():
                        last_trigger = now
            time.sleep(CHECK_INTERVAL)
        except KeyboardInterrupt:
            log("已停止")
            break
        except Exception as e:
            log("⚠️ 异常: %s" % e)
            time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
