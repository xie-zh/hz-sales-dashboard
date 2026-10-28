# -*- coding: utf-8 -*-
"""
看板产物校验（空表 / 脏数据拦截）
================================================================
读取 build_snapshot 生成的 _deploy/*.json，按 config 阈值校验：
  - 人员看板 ry_data.json：总人数 >= 阈值；每位员工岗位必须 ∈ 配置白名单
  - 三力看板 sanli_data.json：总人数 >= 阈值；无「离职」状态残留
  - 活动量看板 data.json：部门数 >= 阈值

用途：
  1) CI 流水线门禁：不达标 exit(1) → 阻断发布（防脏数据上线）
  2) 本地守护 watcher_run.py：build 后调用，校验通过才标记可部署

任何不达标都打印明确原因并以非零码退出，绝不静默放行。
"""
import os
import sys
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(BASE_DIR, "_deploy")

try:
    import config_loader
    CFG = config_loader.load_config()
except Exception:
    CFG = None

PERSONNEL = (CFG or {}).get("personnel", {})
TARGET_POSTS = set(PERSONNEL.get("target_posts", []))
RY_MIN = (CFG or {}).get("verify", {}).get("ry_min_headcount", 10)
SANLI_MIN = (CFG or {}).get("verify", {}).get("sanli_min_headcount", 10)
ACT_MIN = (CFG or {}).get("verify", {}).get("activity_min_depts", 1)
EXCLUDE_STATUS = set((CFG or {}).get("sanli", {}).get("exclude_status", ["离职"]))


def _load(name):
    path = os.path.join(DEPLOY, name)
    if not os.path.exists(path):
        return None, "文件不存在: " + path
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f), None
    except Exception as e:
        return None, "解析失败: " + str(e)


def check_ry(d):
    errs = []
    total = d.get("total", 0)
    if total < RY_MIN:
        errs.append("人员看板总人数 {} < 阈值 {}".format(total, RY_MIN))
    for e in d.get("employees", []):
        post = e.get("岗位")
        if post not in TARGET_POSTS:
            errs.append("人员看板存在非法岗位: {}（{}）".format(post, e.get("姓名")))
    return errs


def check_sanli(d):
    errs = []
    total = d.get("total", 0)
    if total < SANLI_MIN:
        errs.append("三力看板总人数 {} < 阈值 {}".format(total, SANLI_MIN))
    for e in d.get("employees", []):
        if e.get("状态") in EXCLUDE_STATUS:
            errs.append("三力看板残留排除状态人员: {}（{}）".format(e.get("状态"), e.get("姓名")))
    return errs


def check_activity(d):
    errs = []
    depts = d.get("total", [])
    if isinstance(depts, list):
        n = len(depts)
    elif isinstance(depts, dict):
        n = len(depts)
    else:
        n = 0
    if n < ACT_MIN:
        errs.append("活动量看板部门数 {} < 阈值 {}".format(n, ACT_MIN))
    return errs


def main():
    errors = []
    ry, e = _load("ry_data.json")
    if ry is None:
        errors.append("[人员] " + e)
    else:
        errors += ["[人员] " + x for x in check_ry(ry)]

    sanli, e = _load("sanli_data.json")
    if sanli is None:
        errors.append("[三力] " + e)
    else:
        errors += ["[三力] " + x for x in check_sanli(sanli)]

    act, e = _load("data.json")
    if act is None:
        errors.append("[活动量] " + e)
    else:
        errors += ["[活动量] " + x for x in check_activity(act)]

    if errors:
        print("✗ 产物校验未通过，阻断发布：")
        for x in errors:
            print("  - " + x)
        return 1
    print("✓ 产物校验通过（人员 {} / 三力 {} / 活动量部门 {}）".format(
        ry.get("total") if ry else "?", sanli.get("total") if sanli else "?",
        len(act.get("total", [])) if act else "?"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
