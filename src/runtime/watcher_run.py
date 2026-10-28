# -*- coding: utf-8 -*-
"""
看板本地监听守护（P3 · 单次执行版）
================================================================
检测三个源 Excel 的修改时间（mtime），与 _deploy 对应 JSON 中记录的
meta.file_mtime_ts 对比：
  - 任一源比本地快照更新  →  运行 build_snapshot 重建  →  运行 verify_artifact 校验
  - 无变化                  →  跳过构建（零消耗）

校验通过则把 _deploy_status.json 的 needs_deploy 标记为 True（本地快照已最新、可部署）；
校验不通过则 needs_deploy 仍为 False 且记录错误，绝不标可部署。

设计边界（防变故）：
  - 本脚本**只负责本地快照的自动构建与校验，不执行公网部署**。
    公网部署仍由「更新看板」（agent）或自动化完成，确保部署动作受控。
  - 写 _deploy_status.json 供上层（自动化 / 人工）决定是否部署。
  - 支持 --force 强制重建（调试用）。

典型用法：
  计划任务 / automation 每 30 分钟调用一次：  python watcher_run.py
"""
import os
import sys
import json
import subprocess
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(BASE_DIR, "_deploy")
PY = sys.executable

# 源 Excel  →  对应 _deploy JSON
SOURCES = [
    ("分公司活动量-在线看板.xlsx", "data.json"),
    ("分公司-员工三力明细.xlsx", "sanli_data.json"),
    ("分公司-人员数据分析路径.xlsx", "ry_data.json"),
]

# mtime 比较容差（秒），规避浮点/写入时序微小差异
MTime_TOLERANCE = 2.0


def src_mtime(name):
    p = os.path.join(BASE_DIR, name)
    if not os.path.exists(p):
        return None
    return os.path.getmtime(p)


def deploy_mtime(json_name):
    p = os.path.join(DEPLOY, json_name)
    if not os.path.exists(p):
        return None
    try:
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
        return d.get("meta", {}).get("file_mtime_ts")
    except Exception:
        return None


def needs_build(force=False):
    if force:
        return True, ["--force 强制重建"]
    changed = []
    for src, j in SOURCES:
        sm = src_mtime(src)
        dm = deploy_mtime(j)
        if sm is None:
            changed.append("源缺失: " + src)
        elif dm is None:
            changed.append("快照缺失: " + j)
        elif sm - (dm or 0) > MTime_TOLERANCE:
            changed.append("{} 更新（源 {} > 快照 {}）".format(src, _fmt(sm), _fmt(dm)))
    return len(changed) > 0, changed


def _fmt(ts):
    if ts is None:
        return "?"
    return datetime.datetime.fromtimestamp(ts).strftime("%m-%d %H:%M:%S")


def run_build():
    r = subprocess.run([PY, "build_snapshot.py"], cwd=BASE_DIR,
                       capture_output=True, text=True)
    return r.returncode == 0, (r.stdout + r.stderr).strip()


def run_verify():
    r = subprocess.run([PY, "verify_artifact.py"], cwd=BASE_DIR,
                       capture_output=True, text=True)
    return r.returncode == 0, (r.stdout + r.stderr).strip()


def write_status(needs_deploy, verified, errors, detail):
    status = {
        "checked_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "needs_deploy": bool(needs_deploy),
        "verified": bool(verified),
        "errors": errors,
        "detail": detail,
    }
    with open(os.path.join(DEPLOY, "_deploy_status.json"), "w", encoding="utf-8") as f:
        json.dump(status, f, ensure_ascii=False, indent=2)
    return status


def main():
    force = "--force" in sys.argv
    need, changes = needs_build(force)
    if not need:
        print("[watcher] 无变化，跳过构建（本地快照已是最新）")
        write_status(needs_deploy=False, verified=True, errors=[],
                     detail="无源文件更新")
        return 0

    print("[watcher] 检测到变化：")
    for c in changes:
        print("  - " + c)
    ok, out = run_build()
    if not ok:
        print("[watcher] 构建失败：\n" + out)
        write_status(needs_deploy=False, verified=False, errors=[out],
                     detail="build 失败")
        return 1
    print("[watcher] 构建完成，开始校验…")
    vok, vout = run_verify()
    print(vout)
    if not vok:
        print("[watcher] ✗ 校验未通过，已阻断部署标记")
        write_status(needs_deploy=False, verified=False, errors=[vout],
                     detail="verify 未通过")
        return 1
    print("[watcher] ✓ 本地快照已更新并通过校验，可部署（由更新看板/自动化执行）")
    write_status(needs_deploy=True, verified=True, errors=[],
                 detail="; ".join(changes))
    return 0


if __name__ == "__main__":
    sys.exit(main())
