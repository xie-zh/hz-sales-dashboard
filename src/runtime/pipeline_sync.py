# -*- coding: utf-8 -*-
"""分公司看板 · 统一流水线编排器

把「BI 取数 → 写 Excel → 重建看板 → 发布云端」四步串成一条链，一次调用跑完：

    帆软 BI  ──导出 xlsx──▶  DGGRe5.0 下载+复制  ──▶  4 个源 Excel
                                                        │
    腾讯云 CloudBase（网页） ◀── 发布 ◀── build_snapshot ◀┘
                                    └─▶ 腾讯云 COS（异地备份）

设计要点：
1. 阶段 1 直接复用 DGGRe5.0 的模块（main.unified_download_files + execute_pipeline_task），
   不改动它一行代码，配置仍然只认 DGGRe5.0\\config.toml 与 tasks\\*.toml。
2. 阶段 1 任一环失败 → 不重建、不发布，线上保持上一轮数据，并把 status.json 标记为
   stale=true（页面顶部挂红条提示「数据未更新」）。这是刻意的选择：
   半截数据比旧数据危险得多。
3. 本脚本用系统 Python 3.15 运行（selenium / tomlkit / pywin32 装在那里），
   构建与 COS 备份用 venv Python（openpyxl / qcloud_cos 装在那里）。
   两个解释器各管一段，互不污染。

用法：
    python pipeline_sync.py            按 08:00-22:00 时段执行（默认）
    python pipeline_sync.py --force    忽略时段，立刻跑一轮
    python pipeline_sync.py --dry      跑完阶段 1/2 但不发布（演练用）
    python pipeline_sync.py --skip-dgg 跳过下载复制，只重建+发布（改了看板模板时用）
退出码：0=成功  1=本轮失败（已跳过发布/或只发布了过期标记）  2=不在时段内
"""

import json
import os
import subprocess
import sys
import time
import traceback
from datetime import datetime

# --------------------------------------------------------------------------
# 路径常量
# --------------------------------------------------------------------------
BASE = os.path.dirname(os.path.abspath(__file__))
DGG_DIR = r"<PROJECT_ROOT>\DGGRe5.0"
PY_VENV = r"<HOME>\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
DEPLOY = os.path.join(BASE, "_deploy")
COS_CFG = os.path.join(BASE, "_cos_config.json")
STATUS_PATH = os.path.join(DEPLOY, "status.json")
STATE_PATH = os.path.join(BASE, "_pipeline_state.json")
LOCK_PATH = os.path.join(BASE, "_pipeline.lock")
LOG_PATH = os.path.join(BASE, "_pipeline.log")

# 更新时段（用户设定：8:00-22:00）
WINDOW_START = "08:00"
WINDOW_END = "22:00"

# 看板的 4 个源 Excel（DGGRe5.0 复制的目标文件，也是 build_snapshot 的输入）
SRC_XLSX = [
    os.path.join(BASE, "分公司活动量-在线看板.xlsx"),
    os.path.join(BASE, "分公司-员工三力明细.xlsx"),
    os.path.join(BASE, "分公司-人员数据分析路径.xlsx"),
    os.path.join(r"<PROJECT_ROOT>\数据文件\分公司健康表",
                 "四大模式-重点指标表格.xlsx"),
]

_LINES = []


def log(msg):
    line = "[%s] %s" % (time.strftime("%H:%M:%S"), msg)
    try:
        print(line)
    except Exception:
        pass
    _LINES.append(line)


def flush_log():
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("=== %s ===\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
            f.write("\n".join(_LINES) + "\n\n")
    except Exception:
        pass


def now_str():
    return time.strftime("%Y-%m-%d %H:%M:%S")


# --------------------------------------------------------------------------
# 阶段 1：跑 DGGRe5.0 一轮（BI 下载 + Excel 复制）
# --------------------------------------------------------------------------
def run_dgg_round():
    """返回 (ok: bool, note: str)。不改动 DGGRe5.0 任何文件，只 import 复用。"""
    cwd0 = os.getcwd()
    try:
        if DGG_DIR not in sys.path:
            sys.path.insert(0, DGG_DIR)
        os.chdir(DGG_DIR)  # DGG 的配置与相对路径都基于它自己的目录

        import utils as U
        from validator import prepare_config
        import main as M

        U.init_logger()
        cfg = U.load_config()
        errs = prepare_config(cfg, log_fn=lambda m, indent=0: log("[配置] %s" % m))
        if errs:
            return False, "配置校验失败: " + "; ".join(str(e) for e in errs[:5])

        sc = cfg.get("scheduler", {})
        bcfg = cfg.get("browser", {})
        dcfg = cfg.get("disk", {})

        M.set_browser_config(bcfg.get("max_browser_restarts", 3),
                             bcfg.get("browser_timeout", 120))
        M._RETRY_BUDGET_SECONDS = sc.get("download_retry_seconds", 20)

        files = {f["id"]: f for f in cfg.get("files", [])}
        tasks = []
        for t in cfg.get("tasks", []):
            tasks.append({
                "name": t["name"],
                "type": "pipeline",
                "enabled": t.get("enabled", True),
                "copies": t.get("copies", []),
                "html": t.get("html", []),
                "interval_minutes": t.get("interval_minutes"),
                "schedule_times": t.get("schedule_times"),
                "end_time": t.get("end_time"),
                "start_time": t.get("start_time"),
                "priority": t.get("priority", 0),
            })
        if not tasks:
            return False, "DGGRe5.0 没有配置任何任务（tasks/*.toml）"

        downloaded, dl = M.unified_download_files(
            tasks, files, cfg["download_dir"], bcfg, cfg.get("bi_login_no", ""),
            {}, dcfg.get("min_free_space_mb", 500), dcfg.get("auto_clean", True),
            dcfg.get("min_file_age_hours", 72), dcfg.get("retention_times", []),
        )

        ok_c = fail_c = 0
        detail = []
        for t in tasks:
            st = M.execute_pipeline_task(
                t, files, cfg["download_dir"],
                cfg.get("html_output_path", ""), cfg.get("html_refresh_seconds", 300),
                bcfg, cfg.get("com", {}).get("engine", "auto"), cfg.get("bi_login_no", ""),
                downloaded, cfg.get("backup", {}), {},
            )
            c = st.get("copy", {})
            ok_c += c.get("success", 0)
            fail_c += c.get("fail", 0)
            detail.append("%s 复制%d/失败%d" % (t["name"], c.get("success", 0), c.get("fail", 0)))

        dl_fail = dl.get("fail", 0)
        note = "下载 成功%d/失败%d/复用%d；%s" % (
            dl.get("success", 0), dl_fail, dl.get("reused", 0), "; ".join(detail))
        ok = (dl_fail == 0 and fail_c == 0 and ok_c > 0)
        return ok, note
    except Exception as e:
        log("阶段1 异常：\n" + traceback.format_exc(limit=4))
        return False, "DGG 执行异常: %s" % (str(e)[:300])
    finally:
        try:
            os.chdir(cwd0)
        except Exception:
            pass


# --------------------------------------------------------------------------
# 阶段 2：等 Excel 落盘稳定 + 重建看板
# --------------------------------------------------------------------------
def wait_stable(paths, timeout=90):
    """WPS/Excel COM 保存后文件可能还在写，等 mtime+size 连续两次一致再放行。"""
    end = time.time() + timeout
    last = {}
    stable_since = None
    while time.time() < end:
        now = {}
        missing = False
        for p in paths:
            if not os.path.exists(p):
                missing = True
                break
            st = os.stat(p)
            now[p] = (st.st_mtime, st.st_size)
        if not missing and now == last:
            if stable_since is None:
                stable_since = time.time()
            elif time.time() - stable_since >= 2:
                return True
        else:
            stable_since = None
        last = now
        time.sleep(2)
    return False


def run_py(args, timeout=1800):
    r = subprocess.run(args, cwd=BASE, capture_output=True, text=True,
                       encoding="utf-8", errors="ignore", timeout=timeout)
    return r.returncode, (r.stdout or ""), (r.stderr or "")


# --------------------------------------------------------------------------
# 阶段 3：状态文件 + 发布
# --------------------------------------------------------------------------
def write_status(stale, last_success, err="", extra=None):
    data = {
        "stale": bool(stale),
        "last_success": last_success or "",
        "updated_at": now_str(),
        "last_error": err or "",
        "window": "%s-%s" % (WINDOW_START, WINDOW_END),
    }
    if extra:
        data.update(extra)
    os.makedirs(DEPLOY, exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    log("status.json: stale=%s last_success=%s" % (data["stale"], data["last_success"]))
    return data


def deploy_tcb():
    """只做 CloudBase 静态托管部署（失败轮次用它把 stale 标记推上线）。"""
    try:
        with open(COS_CFG, encoding="utf-8") as f:
            t = json.load(f).get("tcb") or {}
    except Exception as e:
        log("读取 _cos_config.json 失败：%s" % e)
        return False
    node, cli, env = t.get("node"), t.get("cli"), t.get("env_id")
    if not (node and cli and env) or not os.path.exists(cli):
        log("CloudBase 未配置或 CLI 缺失 → 跳过")
        return False
    cwd = os.path.dirname(os.path.dirname(os.path.dirname(cli)))
    cmd = [node, cli, "hosting", "deploy", DEPLOY, "", "-e", env]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                           errors="ignore", cwd=cwd if os.path.isdir(cwd) else None,
                           timeout=1200)
        out = (r.stdout or "") + (r.stderr or "")
        ok = ("File upload completed" in out) or ("Successfully uploaded" in out)
        log("CloudBase 部署：%s" % ("成功" if ok else "结果存疑"))
        if not ok:
            log("  " + out[-400:].replace("\n", " | "))
        return ok
    except Exception as e:
        log("CloudBase 部署失败：%s" % str(e)[:200])
        return False


def load_state():
    try:
        with open(STATE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(d):
    try:
        with open(STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


# --------------------------------------------------------------------------
def acquire_lock(max_age_min=25):
    """单实例锁：计划任务与手动触发重叠时，后到的直接退出，避免两个进程同时写 Excel。

    锁文件超过 max_age_min 分钟视为上一轮异常退出的残留，自动接管。
    """
    if os.path.exists(LOCK_PATH):
        try:
            age = time.time() - os.path.getmtime(LOCK_PATH)
        except Exception:
            age = 10 ** 9
        if age < max_age_min * 60:
            log("已有实例在运行（锁 %s，%.1f 分钟前创建）→ 本轮退出" % (LOCK_PATH, age / 60))
            return False
        log("发现过期锁（%.1f 分钟）→ 视为残留，接管" % (age / 60))
    try:
        with open(LOCK_PATH, "w", encoding="utf-8") as f:
            f.write("%s %d" % (now_str(), os.getpid()))
    except Exception as e:
        log("写锁文件失败：%s（继续运行）" % e)
    return True


def release_lock():
    try:
        os.remove(LOCK_PATH)
    except Exception:
        pass


def in_window(force):
    now = datetime.now().strftime("%H:%M")
    if force:
        return True, "%s ~ %s（--force 忽略）" % (WINDOW_START, WINDOW_END)
    return (WINDOW_START <= now <= WINDOW_END), "%s ~ %s" % (WINDOW_START, WINDOW_END)


def main():
    argv = set(sys.argv[1:])
    force = "--force" in argv
    dry = "--dry" in argv
    skip_dgg = "--skip-dgg" in argv

    log("=== 统一流水线 %s ===  参数: %s" % (now_str(), " ".join(sorted(argv)) or "无"))

    inwin, winnote = in_window(force)
    if not inwin:
        log("不在更新时段 %s（现在 %s）→ 跳过" % (winnote, datetime.now().strftime("%H:%M")))
        flush_log()
        return 2

    if not acquire_lock():
        flush_log()
        return 3

    state = load_state()
    last_success = state.get("last_success", "")

    # ---- 阶段 1：BI 下载 + 写 Excel ----
    ok, note = True, "已跳过（--skip-dgg）"
    if not skip_dgg:
        # 重试 1 次：Excel 保存偶发被占用（上一轮 COM 尚未释放 / 安全软件短暂锁句柄），
        # 隔 20 秒再来一轮通常就好了；两次都失败才判定本轮失败。
        for attempt in (1, 2):
            t0 = time.time()
            ok, note = run_dgg_round()
            log("阶段1 BI下载+复制（第%d次）：%s（%.1fs）%s"
                % (attempt, "成功" if ok else "失败", time.time() - t0, note))
            if ok:
                break
            if attempt == 1:
                log("  20 秒后重试一轮 ...")
                time.sleep(20)
    else:
        log("阶段1 跳过（--skip-dgg）")

    stale = not ok
    err = "" if ok else note[:200]

    # ---- 阶段 2：等落盘 + 重建看板 ----
    built = False
    if ok:
        if not wait_stable(SRC_XLSX):
            log("警告：源 Excel 90 秒内未稳定，仍继续（可能读到上一瞬间的内容）")
        t0 = time.time()
        rc, out, err2 = run_py([PY_VENV, "build_snapshot.py"])
        built = (rc == 0)
        if built:
            log("阶段2 重建看板完成（%.1fs）" % (time.time() - t0))
            last_success = now_str()
        else:
            stale = True
            err = "build_snapshot 失败 exit=%s" % rc
            log("阶段2 重建失败：" + (out + err2)[-800:].replace("\n", " | "))

    # ---- 状态文件（页面读它决定是否挂「数据未更新」红条）----
    write_status(stale, last_success, err,
                 {"stage1": note[:200], "built": built})

    # ---- 阶段 3：发布 ----
    if dry:
        log("阶段3 跳过发布（--dry）")
    elif built:
        # 已构建过 → cos_sync 判变会认为无变化，只做发布（CloudBase + COS 备份）
        rc, out, e2 = run_py([PY_VENV, "cos_sync.py"])
        log("阶段3 发布（CloudBase+COS）：exit=%s" % rc)
        tail = (out or "").strip().splitlines()[-6:]
        for line in tail:
            log("  | " + line)
    else:
        # 本轮失败：不重建，只把 stale 标记推上线（线上数据仍是上一轮成功的）
        okd = deploy_tcb()
        log("阶段3 仅发布过期标记：%s" % ("成功" if okd else "失败"))

    save_state({
        "last_run": now_str(),
        "last_ok": bool(ok),
        "last_success": last_success,
        "last_error": err,
        "built": built,
        "runs": (state.get("runs", 0) + 1),
    })
    release_lock()
    flush_log()
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
