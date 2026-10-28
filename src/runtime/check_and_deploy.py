# -*- coding: utf-8 -*-
"""
看板自动同步检查（单次执行 · 供 Windows 计划任务调用）
=====================================================================
目标：**源 Excel 一更新，线上链接的数据就跟着更新 —— 覆盖全部板块**，
不再像旧版那样只管「活动量」一个模块。

怎么判定「该同步了」
  拿 3 个源表的 mtime 与本地快照 `_deploy/*.json` 里记录的
  `meta.file_mtime_ts`（＝构建那一刻源表的 mtime）对比，任一源表更新即触发。

覆盖的源表 → 5 个板块
  1. 分公司活动量-在线看板.xlsx   → data.json        （活动量）
  2. 分公司-员工三力明细.xlsx      → sanli_data.json  （员工三力）
  3. 分公司-人员数据分析路径.xlsx  → ry_data.json     （网销 / 电邀 / 会计，同一源表）
  4. 分公司健康表/四大模式-重点指标表格.xlsx → 无独立快照  （辅助源表：「网销模式付费客资源
     签单转化率（近30天）」等取自其「网销漏斗」sheet 的 AL 列；与自身记录的 mtime 比较）

与旧版的两处关键修复
  · 旧版只监控「活动量表」→ 三力 / 网销 / 电邀 / 会计 更新后不会同步；
  · 旧版按「自动化名称」定位部署任务，而名称早已改过（→ rowcount=0，触发无效）；
    本版改为按 **自动化 ID** 定位（automation-1786499389829），不受改名影响。

职责边界（为什么本脚本不自己 rebuild）
  重建 + 发布由已存在的自动化任务完成（它的提示词第 2 步就是跑 build_snapshot.py）。
  本脚本**只负责「发现变化 → 触发」**，好处是：
    - 触发成功但自动化尚未跑完时，快照仍为旧值 → 下一轮会**继续重试**，不会漏同步；
    - 自动化不可用时（如桌面端未开）也只是空转，等桌面端一开就会补上。

冷却合并：`COOLDOWN` 秒内若连续保存，只在冷却结束后触发一次，避免频繁消耗算力。
无变化时只做 3 次 stat 即退出，几乎零开销。

用法：
  python check_and_deploy.py            # 计划任务调用
  python check_and_deploy.py --dry      # 只看是否需要同步，不做任何写入
  python check_and_deploy.py --verbose  # 无变化时也打印一行
返回码：0 正常（含无需同步）｜2 需同步(--dry)｜3 触发未生效｜1 出错
"""
import os
import sys
import json
import time
import sqlite3
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(BASE_DIR, "_deploy")
DB = os.path.expanduser("~/.workbuddy/workbuddy.db")
# 按 ID 定位（名称可能变，ID 不变）：看板数据同步到云端（手动触发+每日19点兜底）
AUTOMATION_ID = "automation-1786499389829"
LOG_FILE = os.path.join(BASE_DIR, "_watch.log")
LOG_MAX = 2 * 1024 * 1024        # 日志超过 2MB 轮转一次
STATE_FILE = os.path.join(DEPLOY, "_sync_state.json")
THRESHOLD = 3.0                  # 秒；源 Excel 比快照记录晚这么多才算「有更新」
COOLDOWN = 180                   # 秒；两次触发之间的最小间隔（合并连续保存）

SOURCES = [
    ("分公司活动量-在线看板.xlsx", "data.json", "活动量"),
    ("分公司-员工三力明细.xlsx", "sanli_data.json", "员工三力"),
    ("分公司-人员数据分析路径.xlsx", "ry_data.json", "网销/电邀/会计"),
]

# 辅助源表：影响 ry_* 三页，但主源表是「人员数据分析路径.xlsx」→ 与**自身** mtime 比较
# （该 mtime 由 ry_parser 记进 ry_data.json 的 meta.aux_mtime_ts）。
AUX_LABEL = "网销/电邀/会计（四大模式·网销漏斗AL列）"
AUX_CANDIDATES = [
    os.path.normpath(os.path.join(BASE_DIR, "..", "..", "..", "分公司健康表", "四大模式-重点指标表格.xlsx")),
    os.path.normpath(os.path.join(BASE_DIR, "分公司健康表", "四大模式-重点指标表格.xlsx")),
]


def aux_paths():
    return [p for p in AUX_CANDIDATES if os.path.exists(p)]


def aux_recorded():
    """ry_data.json.meta.aux_mtime_ts（构建时记录的辅助源表 mtime）"""
    try:
        with open(os.path.join(DEPLOY, "ry_data.json"), encoding="utf-8") as f:
            return dict((json.load(f).get("meta") or {}).get("aux_mtime_ts") or {})
    except Exception:
        return {}


def log(msg):
    line = "[%s] %s" % (datetime.datetime.now().strftime("%m-%d %H:%M:%S"), msg)
    print(line, flush=True)
    try:
        if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > LOG_MAX:
            os.replace(LOG_FILE, LOG_FILE + ".1")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def snapshot_ts(json_name):
    """本地快照里记录的「构建时源表 mtime」"""
    try:
        with open(os.path.join(DEPLOY, json_name), encoding="utf-8") as f:
            return float(json.load(f).get("meta", {}).get("file_mtime_ts") or 0.0)
    except Exception:
        return None


def _fmt(ts):
    if not ts:
        return "?"
    return time.strftime("%m-%d %H:%M:%S", time.localtime(ts))


def scan():
    """返回「源表比本地快照新」的板块清单"""
    changed = []
    for src, snap, label in SOURCES:
        sp = os.path.join(BASE_DIR, src)
        if not os.path.exists(sp):
            changed.append("%s：源表缺失(%s)" % (label, src))
            continue
        sts = os.path.getmtime(sp)
        wts = snapshot_ts(snap)
        if wts is None:
            changed.append("%s：快照缺失(%s)" % (label, snap))
        elif sts - wts > THRESHOLD:
            changed.append("%s：源 %s > 快照 %s" % (label, _fmt(sts), _fmt(wts)))
    # 辅助源表（四大模式-重点指标表格.xlsx）→ 与自身记录值比较
    rec = aux_recorded()
    for p in aux_paths():
        base = os.path.basename(p)
        sts = os.path.getmtime(p)
        wts = float(rec.get(base) or 0.0)
        if wts <= 0:
            changed.append("%s：辅助源表 %s 尚未记入快照" % (AUX_LABEL, base))
        elif sts - wts > THRESHOLD:
            changed.append("%s：辅助源表 %s 源 %s > 快照 %s" % (AUX_LABEL, base, _fmt(sts), _fmt(wts)))
    return changed


def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            d = json.load(f)
        return {"last_change_key": str(d.get("last_change_key") or ""),
                "last_trigger_at": float(d.get("last_trigger_at") or 0.0)}
    except Exception:
        return {"last_change_key": "", "last_trigger_at": 0.0}


def save_state(st):
    try:
        tmp = STATE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=2)
        os.replace(tmp, STATE_FILE)
    except Exception:
        pass


def change_key(changed):
    """用「板块 mtime 组合」做变化指纹，避免同一批变化被重复触发"""
    parts = []
    for src, snap, label in SOURCES:
        sp = os.path.join(BASE_DIR, src)
        parts.append("%s=%s" % (label, os.path.getmtime(sp) if os.path.exists(sp) else 0))
    for p in aux_paths():
        parts.append("AUX:%s=%s" % (os.path.basename(p), os.path.getmtime(p)))
    return "|".join(parts)


def trigger():
    """把云端部署自动化设为立即执行（按 ID 定位）"""
    try:
        con = sqlite3.connect(DB, timeout=10)
        cur = con.cursor()
        now_ms = int(time.time() * 1000)
        cur.execute(
            "UPDATE automations SET next_run_at=?, updated_at=? "
            "WHERE id=? AND deleted_at IS NULL",
            (now_ms, now_ms, AUTOMATION_ID),
        )
        n = cur.rowcount
        con.commit()
        con.close()
        return n
    except Exception as e:
        log("触发部署失败: %s" % e)
        return -1


def main():
    dry = "--dry" in sys.argv
    changed = scan()
    if not changed:
        if "--verbose" in sys.argv:
            log("无变化，跳过（本地快照已最新）")
        return 0

    key = change_key(changed)
    st = load_state()
    if key == st["last_change_key"] and (time.time() - st["last_trigger_at"]) < 3600:
        return 0                                  # 同一批变化已触发过，等自动化完成

    if dry:
        log("需同步（--dry 不做写入）：")
        for c in changed:
            log("  - " + c)
        return 2

    now = time.time()
    if now - st["last_trigger_at"] < COOLDOWN:
        log("检测到更新，处于冷却期（%.0fs<%.0fs），合并到下一轮" % (now - st["last_trigger_at"], COOLDOWN))
        return 0

    log("检测到源数据更新：")
    for c in changed:
        log("  - " + c)
    n = trigger()
    if n == 1:
        st["last_change_key"] = key
        st["last_trigger_at"] = now
        save_state(st)
        log("已触发云端部署（自动化 %s）→ 自动化将重建全部五页并发布，约 1~2 分钟生效" % AUTOMATION_ID)
        return 0
    log("触发未生效（rowcount=%s）；本地快照保持旧值，下一轮会重试；"
        "另有每日 19:00 兜底任务" % n)
    return 3


if __name__ == "__main__":
    sys.exit(main())
