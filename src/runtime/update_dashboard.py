# -*- coding: utf-8 -*-
"""
看板三方一致性快速校验 + 条件重建部署
=====================================
对比三层：源 Excel（看板内容） / 本地快照 _deploy/*.json / 云端记录
  - 任一源 Excel 比本地快照新        → 重建快照（build_snapshot.build，一次性重建三看板）
  - 重建后本地比云端新（或刚重建过）  → 需要部署
  - 三者一致                          → 跳过，不消耗任何积分

设计目标：步骤最简、执行最快、适配频繁更新、避免重复部署。
  · 日常最常见路径（无变化）只做本地文件 stat 对比，几乎零开销、零网络；
  · 仅在「源 Excel 比本地快照新」时才重建，且仅在「本地比云端新」时才部署；
  · 单一部署覆盖三看板（共享同一链接），故任何一块变化都走完整重建+部署。

用法：
  python update_dashboard.py              # 完整：校验 → 必要时重建 → 输出部署信号
  python update_dashboard.py --dry        # 仅报告三方状态，不重建、不触发部署（最快检查）
  python update_dashboard.py --local-only # 仅比对本地源Excel与本地快照（不看云端）：
                                           # 本地有变化才重建+部署；无变化直接 SKIP

退出码：
  0   无需部署（无变化 / 三者一致），已跳过（无积分消耗）
  23  需要部署（已重建快照；由调用方执行 CloudStudio 部署）
  1   出错
"""
import os
import sys
import json
import time
import argparse
import urllib.request
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 2026-09-12 换链：原 CloudStudio 通道（ef627954…agentos-app.net）部署工具已被客户端移除，
# 该链接永久无法更新；改用 sites 通道新建的稳定应用 wbapp_Jc5WgPGEeqrk5b5tcLZ3RT。
CLOUD_BASE = "https://dashboard.example.com"
SHARE_LINK = CLOUD_BASE
DEPLOY_DIR = os.path.join(BASE_DIR, "_deploy")
STATUS_FILE = os.path.join(BASE_DIR, "_deploy_status.json")
EPS = 2.0  # 秒；mtime 噪声容差（Excel 保存后 mtime 跳变远大于此值）

# 三看板：源 Excel / 本地快照 / 云端 JSON
BOARDS = [
    {
        "key": "activity",
        "name": "活动量看板",
        "src": "分公司活动量-在线看板.xlsx",
        "local": "_deploy/data.json",
        "cloud": "data.json",
    },
    {
        "key": "sanli",
        "name": "员工三力看板",
        "src": "分公司-员工三力明细.xlsx",
        "local": "_deploy/sanli_data.json",
        "cloud": "sanli_data.json",
    },
    {
        "key": "ry",
        "name": "网销数据分析",
        "src": "分公司-人员数据分析路径.xlsx",
        "local": "_deploy/ry_data.json",
        "cloud": "ry_data.json",
    },
]


def _ts2str(ts):
    if not ts:
        return "—"
    return datetime.fromtimestamp(ts).strftime("%m-%d %H:%M:%S")


def _read_local_ts(rel):
    """读取本地快照 meta.file_mtime_ts"""
    path = os.path.join(BASE_DIR, rel)
    try:
        with open(path, "r", encoding="utf-8") as f:
            meta = json.load(f).get("meta", {})
        return meta.get("file_mtime_ts") or meta.get("file_mtime_ts_ms", 0.0)
    except Exception:
        return None


def _read_cloud_ts(cloud_file):
    """读取云端 JSON 的 meta.file_mtime_ts（带缓存破坏参数）。
    失败时退回到 curl 再试一次；再失败则置 0（保守视为需要部署）。"""
    url = f"{CLOUD_BASE}/{cloud_file}?cb={int(time.time()*1000)}"
    # 第一试：urllib
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "sync-check/1"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            meta = json.loads(resp.read().decode("utf-8")).get("meta", {})
        return meta.get("file_mtime_ts") or meta.get("file_mtime_ts_ms", 0.0)
    except Exception:
        pass
    # 第二试：subprocess 跑 curl（部分环境下 urllib 受限但 curl 可用）
    try:
        import subprocess
        out = subprocess.check_output(
            ["curl", "-sL", "--max-time", "15", "-H", "User-Agent: sync-check/2", url],
            stderr=subprocess.DEVNULL,
        )
        meta = json.loads(out.decode("utf-8", errors="ignore")).get("meta", {})
        return meta.get("file_mtime_ts") or meta.get("file_mtime_ts_ms", 0.0)
    except Exception:
        return 0.0  # 视云端为最旧（最保守：触发部署）


def _rebuild():
    """重建三看板快照（build_snapshot.build 一次性重建全部）"""
    sys.path.insert(0, BASE_DIR)
    import build_snapshot
    build_snapshot.build()


def _write_status(need_deploy, rebuilt, boards):
    payload = {
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "need_deploy": need_deploy,
        "rebuilt": rebuilt,
        "share_link": SHARE_LINK,
        "deploy_dir": DEPLOY_DIR,
        "boards": boards,
    }
    try:
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true",
                    help="仅报告三方状态，不重建、不触发部署")
    ap.add_argument("--local-only", action="store_true",
                    help="仅比对本地源Excel与本地快照（不看云端）：本地有变化才重建+部署，无变化直接 SKIP")
    args = ap.parse_args()

    fetch_cloud = (not args.dry) and (not args.local_only)
    title = ("  看板一致性校验（源Excel / 本地快照）" if args.local_only
             else "  看板三方一致性校验（源Excel / 本地快照 / 云端）")
    print("=" * 64)
    print(title)
    print("=" * 64)

    rows = []
    need_rebuild = False
    for b in BOARDS:
        src_path = os.path.join(BASE_DIR, b["src"])
        src_ts = os.path.getmtime(src_path) if os.path.exists(src_path) else None
        local_ts = _read_local_ts(b["local"])
        cloud_ts = _read_cloud_ts(b["cloud"]) if fetch_cloud else None

        if src_ts is None or local_ts is None:
            state = "???"  # 源或本地缺失
        elif src_ts - local_ts > EPS:
            state = "SRC_NEWER"      # 源比本地新 → 需重建
            need_rebuild = True
        elif (not args.local_only) and (local_ts - (cloud_ts or 0)) > EPS:
            state = "CLOUD_STALE"    # 本地比云端新 → 需部署
        else:
            state = "OK"

        rows.append({
            "name": b["name"],
            "src": src_ts, "local": local_ts, "cloud": cloud_ts,
            "state": state,
        })
        print(f"  {b['name']:<14} 源:{_ts2str(src_ts)}  本地:{_ts2str(local_ts)}"
              f"  云端:{_ts2str(cloud_ts)}  [{state}]")

    rebuilt = False
    if need_rebuild:
        if args.dry:
            print("\n[dry] 检测到源 Excel 有更新，但 --dry 模式不重建。")
        else:
            print("\n↻ 源数据有更新，重建三看板快照 …")
            try:
                _rebuild()
                rebuilt = True
                # 重建后刷新本地时间戳，供后续部署判定
                for r, b in zip(rows, BOARDS):
                    r["local"] = _read_local_ts(b["local"])
            except Exception as e:
                print(f"❌ 重建失败: {e}")
                _write_status(False, False, rows)
                return 1

    # 部署判定：
    #   --local-only 模式 → 只看本地源Excel是否变化（rebuilt），不比较云端；
    #   完整模式        → 刚重建过，或存在本地比云端新的看板。
    need_deploy = rebuilt
    if not need_deploy and (not args.dry) and (not args.local_only):
        for r in rows:
            if (r["local"] or 0) - (r["cloud"] or 0) > EPS:
                need_deploy = True
                break

    _write_status(need_deploy, rebuilt, rows)

    print("-" * 64)
    if args.dry:
        print("RESULT:DRY_REPORT  (--dry 不执行任何写入/部署)")
        return 0
    if need_deploy:
        print("RESULT:DEPLOY  → 已重建快照，需执行 CloudStudio 部署（退出码 23）")
        print(f"           部署目录: {DEPLOY_DIR}")
        print(f"           稳定链接: {SHARE_LINK}（部署后不变）")
        return 23
    print("RESULT:SKIP  → 三者一致，无需部署（本次零积分消耗）")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n已取消。")
        sys.exit(130)
