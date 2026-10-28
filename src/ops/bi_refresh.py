# -*- coding: utf-8 -*-
# @version V0.3  @author <AUTHOR>  @created 2026-09-24
"""某公司看板 · 一键刷新（BI 取数 → 写 Excel → 重建 → 发布）

   帆软 BI ──HTTP直连导出──▶ _bi_cache/*.xlsx ──整块写入──▶ 4 个源 Excel
                                                                │
     WorkBuddy 分享链接（公网看板） ◀── 发布 ◀── build_snapshot ◀┘
                                        （发布由 WorkBuddy 自动化发起）

与旧的差别（这是重写，不是包装 EXAMPLERe5.0）：
  · 不开浏览器、不常驻、不用 Playwright，下载走帆软导出接口，10 张表并发
  · 写 Excel 只开一次后台 WPS/Excel（不可见），整块数组一次性写入，跑完立刻退出
  · 内容没变就秒级跳过：不写 Excel、不重建、不发布
  · 失败保持上一次的数据，并在页面顶部挂「数据未更新」提示，不会白屏

V1.5（2026-10-12）发布通道改造：
  · 公网发布从腾讯云 CloudBase 改为 **WorkBuddy 分享链接**（腾讯云账号要收回）。
  · 脚本负责「本地构建 + 写好发布请求」（_deploy/_publish_request.json）；
    真正的上传由 WorkBuddy 客户端的「发布为应用」完成 —— 脚本无法直连云端，
    所以由一条 WorkBuddy 自动化任务读取该标记后决定是否发布。
  · 过渡期可用「迁移配置.json」的 mirror_tencent 开关，顺带也镜像到腾讯云。

用法（用 venv Python 跑，见 _run_refresh.bat）：
    python bi_refresh.py              按 bi_sources.json 的时段跑，无变化就跳过
    python bi_refresh.py --force      忽略时段 + 忽略判变，强制重建并发布
    python bi_refresh.py --dry        跑完但不发布（演练）
    python bi_refresh.py --skip-bi    跳过取数，只重建+发布（改了看板模板时用）
    python bi_refresh.py --check-online   只核对线上内容与本地是否一致（不发布）
    python bi_refresh.py --mark-published 记录「本次构建已发布」（自动化发布成功后调）
退出码：0=成功或无变化  1=本轮失败  2=不在时段内  3=已有实例在运行
"""

import json
import os
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime

import bi_engine
import _portable

# V1.4.6：署名与完整性的唯一来源（作者署名集中在 _author.py，交付时置只读）
try:
    import _author
except Exception:
    _author = None

BASE = os.path.dirname(os.path.abspath(__file__))
# === Hzexamplebot（V0.3）：脚本已迁入 Hzexamplebot 子目录，但「数据 / 配置 / 构建产物」
# 仍留在根目录（源 Excel、_deploy、历史数据、_cos_config.json 等共享资源），
# 因此用 ROOT 显式指向根目录。
# V0.4（10-08）：ROOT 不再写死 → 走 _portable 解析（_paths.json → 环境变量 →
# 自动探测特征文件 → 兜底父目录），整包换电脑/换路径都能自己找到数据目录。
ROOT = _portable.resolve_root()
CONFIG = os.path.join(BASE, "bi_sources.json")
STATE_PATH = os.path.join(BASE, "_bi_state.json")
LOCK_PATH = os.path.join(BASE, "_bi_refresh.lock")
LOG_PATH = os.path.join(BASE, "_bi_refresh.log")
DEPLOY = os.path.join(ROOT, "_deploy")
STATUS_PATH = os.path.join(DEPLOY, "status.json")
COS_CFG = os.path.join(ROOT, "_cos_config.json")
BACKUP_ROOT = os.path.join(BASE, "_backup", "bi_daily")
# V1.5：本地构建完成、待 WorkBuddy 发布的标记（自动化任务据此判断要不要发）
PUBLISH_REQ = os.path.join(DEPLOY, "_publish_request.json")
# V1.5：线上当前是哪个指纹（自动化发布成功后写）——用于「无变化就不发」
PUBLISHED_PATH = os.path.join(DEPLOY, "_published.json")

# 这几个文件的看板数值由 Excel 公式算出（XLOOKUP/VLOOKUP），写完必须重算，
# 否则读到的还是上一次的缓存值。其余文件不重算 → 保存更快。
DEFAULT_RECALC = ["活动量"]

_LINES = []


def log(msg):
    line = "[%s] %s" % (time.strftime("%H:%M:%S"), msg)
    try:
        print(line)
    except Exception:
        pass
    _LINES.append(line)


def flush_log():
    """滚动保留最近若干轮的日志（排障要看历史，不能每轮都覆盖掉）。"""
    try:
        old = ""
        if os.path.exists(LOG_PATH):
            try:
                raw = open(LOG_PATH, encoding="utf-8", errors="ignore").read()
            except Exception:
                raw = ""
            if len(raw) > 200 * 1024:
                old = "...(更早日志已截断)...\n" + raw[-200 * 1024:]
            else:
                old = raw
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            if old:
                f.write(old)
            f.write("=== %s ===\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
            f.write("\n".join(_LINES) + "\n\n")
    except Exception:
        pass


def now_str():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def author_note():
    """静默看护署名（V1.4.6）：被改动就自动还原 + 重新置只读。
    不打印常规状态、不弹窗；看护失败一律吞掉 —— 看板刷新是生产链路，优先保证能跑。"""
    if _author is None:
        return
    try:
        _author.set_base(BASE)
        info = _author.silent_guard()
        if info.get("restored"):
            log("[维护] 已还原 %d 处改动：%s" % (len(info["restored"]), "、".join(info["restored"])))
    except Exception:
        pass


# --------------------------------------------------------------------------
def load_json(path, default=None):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log("读取 %s 失败：%s" % (os.path.basename(path), e))
        return default if default is not None else {}


def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    except Exception as e:
        log("写 %s 失败：%s" % (os.path.basename(path), e))


def _pid_alive(pid):
    """锁文件里的进程还活着吗？进程被强杀后锁会残留，靠这个识别孤儿锁。"""
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True          # 进程存在，只是没权限发信号
    except OSError:
        return False
    return True


def acquire_lock(max_age_min=25):
    """单实例锁：上一轮没跑完时，新的这一轮直接退出。"""
    if os.path.exists(LOCK_PATH):
        try:
            age = time.time() - os.path.getmtime(LOCK_PATH)
        except Exception:
            age = 10 ** 9
        try:
            pid = int((open(LOCK_PATH, encoding="utf-8").read().split() or ["0"])[-1])
        except Exception:
            pid = 0
        alive = pid and _pid_alive(pid)
        if alive and age < max_age_min * 60:
            log("上一轮还在跑（pid %d，锁 %.1f 分钟前创建）→ 本轮退出" % (pid, age / 60))
            return False
        if not alive:
            log("发现孤儿锁（pid %d 已不存在）→ 接管" % pid)
        else:
            log("发现过期锁（%.1f 分钟）→ 视为残留，接管" % (age / 60))
    try:
        with open(LOCK_PATH, "w", encoding="utf-8") as f:
            f.write("%s %d" % (now_str(), os.getpid()))
    except Exception as e:
        log("写锁失败：%s（继续运行）" % e)
    return True


def release_lock():
    try:
        os.remove(LOCK_PATH)
    except Exception:
        pass


def wait_files_idle(paths, timeout=45, poll=1.0):
    """刚写完的 Excel，WPS/Excel 进程可能还握着句柄，立刻重建会读失败。
    这里轮询到每个文件都能独占打开为止，返回等待秒数（0=本来就可写）。"""
    start = time.time()
    pending = [p for p in paths if os.path.exists(p)]
    while pending and time.time() - start < timeout:
        pending = [p for p in pending if bi_engine.is_locked(p)]
        if pending:
            time.sleep(poll)
    if pending:
        log("  ⚠ 等待 %.0fs 后仍被占用：%s"
            % (timeout, "、".join(os.path.basename(p) for p in pending)))
    return round(time.time() - start, 1)


def in_window(win, force):
    s, e = win.get("start", "00:00"), win.get("end", "23:59")
    now = datetime.now().strftime("%H:%M")
    if force:
        return True, "%s~%s（--force 忽略）" % (s, e)
    return (s <= now <= e), "%s~%s" % (s, e)


def in_peak(peak):
    s, e = peak.get("start", "08:00"), peak.get("end", "22:00")
    now = datetime.now().strftime("%H:%M")
    return s <= now <= e


def throttle_skip(cfg, state, force):
    """按时段节流：高峰用 peak_interval_min，非高峰用 offpeak_interval_min；
    距上次实际运行不足间隔就跳过（计划任务每 5 分钟触发，靠这里降频）。
    间隔设 0 / 不设 = 该时段不节流（每次都跑）。--force 忽略节流。
    对齐更新（throttle.align_min = 30/60，按本机北京时间）：
      30 → 只在每小时 :00/:30 后的第一响放行；60 → 只在 :00 后的第一响放行。
      错过对齐点且距上次实际运行 ≥ 对齐周期+5 分钟 → 兜底放行，避免卡死。"""
    th = cfg.get("throttle") or {}
    peak = th.get("peak") or {}
    if force or not peak:
        return False, "--force" if force else "未配置节流"

    cad = 5   # 计划任务触发节奏（分钟）
    am = int(th.get("align_min") or 0)
    now = datetime.now()
    if am > 0:
        mod = now.minute % am
        ref = state.get("last_run") or state.get("last_success") or ""
        passed = None
        if ref:
            try:
                dt = datetime.strptime(ref, "%Y-%m-%d %H:%M:%S")
                passed = (now - dt).total_seconds() / 60.0
            except Exception:
                passed = None
        if mod < cad:
            if passed is not None and passed < cad + 1:
                return True, "对齐点刚跑过（距上次 %.0f 分钟）→ 跳过" % passed
            return False, "对齐触发点（北京时间 %02d:%02d，每 %d 分钟）→ 放行" % (now.hour, now.minute, am)
        if passed is None or passed >= am + cad:
            return False, "错过对齐点，距上次 %s 分钟 → 兜底放行" % ("?" if passed is None else "%.0f" % passed)
        return True, "未到对齐点（北京时间 %02d:%02d，对齐 %d 分钟）" % (now.hour, now.minute, am)

    is_peak = in_peak(peak)
    if is_peak:
        gap = float(th.get("peak_interval_min") or 0)
        tag = "高峰"
    else:
        gap = float(th.get("offpeak_interval_min") or 30)
        tag = "非高峰"
    if gap <= 0:
        return False, tag + "时段（未设间隔，每次都跑）"
    ref = state.get("last_run") or state.get("last_success") or ""
    if not ref:
        return False, "无上次记录"
    try:
        dt = datetime.strptime(ref, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return False, "上次时间无法解析"
    passed = (datetime.now() - dt).total_seconds() / 60.0
    if passed < gap:
        return True, "%s，距上次 %.0f 分钟（间隔 %d 分钟）" % (tag, passed, gap)
    return False, "%s，距上次 %.0f 分钟 → 放行" % (tag, passed)


# --------------------------------------------------------------------------
def daily_backup(paths):
    """每个源文件每天留一份快照，出事能回滚当天写入前的内容。"""
    day = time.strftime("%Y-%m-%d")
    d = os.path.join(BACKUP_ROOT, day)
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        return
    for p in paths:
        if not os.path.exists(p):
            continue
        dst = os.path.join(d, os.path.basename(p))
        if not os.path.exists(dst):
            try:
                shutil.copy2(p, dst)
            except Exception as e:
                log("  备份 %s 失败：%s" % (os.path.basename(p), str(e)[:60]))


def write_status(stale, last_success, err="", extra=None):
    data = {
        "stale": bool(stale),
        "last_success": last_success or "",
        "updated_at": now_str(),
        "last_error": err or "",
    }
    if extra:
        data.update(extra)
    os.makedirs(DEPLOY, exist_ok=True)
    save_json(STATUS_PATH, data)
    log("status.json → stale=%s  上次成功=%s" % (data["stale"], data["last_success"] or "无"))
    return data


def _resolve_node(node):
    """node 路径失效时（WorkBuddy 升级会换 versions/<ver> 目录名），自动扫描
    versions 目录取版本号最大的 node.exe，并回写 _cos_config.json 自愈。
    V0.4：探测逻辑统一收到 _portable，本机换账号/换盘符也能找到。
    返回可用 node 路径或 None。"""
    if node and os.path.exists(node):
        return node
    best = _portable.resolve_node()
    if not best:
        return None
    log("node 路径失效（%s）→ 自动切换 %s" % (node or "未配置", best))
    try:
        cfg = load_json(COS_CFG)
        if cfg.get("tcb"):
            cfg["tcb"]["node"] = best
            cfg["tcb"]["cli"] = _portable.resolve_tcb() or cfg["tcb"].get("cli")
            save_json(COS_CFG, cfg)
    except Exception:
        pass  # 回写失败不影响本次发布
    return best


# 发布后必须逐一核对的「页面数据文件」：本地非空而线上为空 ⇒ 一定是发布没落地
CRITICAL_DATA = ("data.json", "sanli_data.json", "ry_data.json",
                 "ry_dy_data.json", "ry_kuaiji_data.json")


def _http_get(url, timeout=25):
    import urllib.request
    req = urllib.request.Request(url, headers={
        "User-Agent": "Hzexamplebot-refresh/1.0",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def _has_positive(v, depth=0):
    """递归判断值里是否存在「非零、非空」的实质内容。"""
    if depth > 4:
        return False
    if isinstance(v, dict):
        return any(_has_positive(x, depth + 1) for x in v.values())
    if isinstance(v, (list, tuple)):
        return any(_has_positive(x, depth + 1) for x in v)
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        return v.strip() not in ("", "-", "0", "0.0", "None", "null")
    return v is not None


def _data_has_content(d):
    """粗判一份数据 JSON 是不是「空壳」（结构完整、meta 齐全，业务字段却全空）。

    判据只看明细/汇总块，不看 meta —— 那次线上事故的版本恰恰是 meta 正常、
    employees/bus/departments 全为 [] 且 total=0。字段名对不上时报「空」是
    安全方向：调用方只对「本地非空」的文件要求线上也非空。
    """
    if not isinstance(d, dict):
        return bool(d)
    # 各页的明细块（data.json 用 stats/legal、total 也恰是列表，不能只看数值）
    for k in ("employees", "bus", "departments", "rows", "items", "list",
              "stats", "legal", "total", "count", "人数"):
        v = d.get(k)
        if isinstance(v, list) and len(v) > 0:
            return True
        if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0:
            return True
    v = d.get("summary")
    if isinstance(v, dict) and _has_positive(v):
        return True
    return False


def online_verify(cb=None):
    """发布后独立核对线上内容（不依赖发布工具自己的结论）。

    比对两件事：
      1) 线上 _build_fingerprint.json 的 build_hash 是否等于本地本次构建的；
      2) 本地非空的数据文件，线上不能是空壳。
    返回 (ok, 说明)。网络异常一律算「没核对上」，不当作通过。

    V1.5：线上地址改从「迁移配置.json」的 publish.dash_url 取（WorkBuddy 分享链接，
    应用把 _deploy 目录发布在**根路径**）；dash_url 还没回填时退回老配置的
    腾讯云 tcb.url + /dash/ 子路径，过渡期不至于核对不上。
    """
    import io as _io
    pub = _pub()
    base = (pub.get("dash_url") or "").rstrip("/")
    prefix = ""
    if not base:
        t = load_json(COS_CFG).get("tcb") or {}
        base = (t.get("url") or "").rstrip("/")
        prefix = "/dash"
    if not base:
        return True, "未配置线上地址，跳过核对"
    q = "?cb=%d" % (int(cb if cb is not None else time.time()))
    lh = ""
    try:
        lh = (load_json(os.path.join(DEPLOY, "_build_fingerprint.json")) or {}).get("build_hash") or ""
    except Exception:
        pass
    try:
        rh = (json.loads(_http_get(base + prefix + "/_build_fingerprint.json" + q)
                         .decode("utf-8", "replace")) or {}).get("build_hash") or ""
    except Exception as e:
        return False, "线上指纹取不到（%s）" % str(e)[:70]
    if lh and rh != lh:
        return False, "指纹不一致（线上 %s ≠ 本地 %s）" % (rh or "空", lh)
    for name in CRITICAL_DATA:
        lp = os.path.join(DEPLOY, name)
        if not os.path.isfile(lp):
            continue
        try:
            with _io.open(lp, encoding="utf-8") as f:
                local = json.load(f)
        except Exception:
            continue
        if not _data_has_content(local):
            continue                      # 本地本来就空，不苛求线上
        try:
            remote = json.loads(_http_get(base + prefix + "/" + name + q).decode("utf-8", "replace"))
        except Exception as e:
            return False, "%s 线上取不到（%s）" % (name, str(e)[:60])
        if not _data_has_content(remote):
            return False, "%s 线上是空数据（本地非空）" % name
    return True, "指纹 %s · 5 个数据文件非空" % (lh or "-")


def _pub(root=None):
    """当前发布参数（「迁移配置.json」→ _cos_config.json）。取不到返回空 dict。"""
    try:
        return _portable.publish_info(root or ROOT) or {}
    except Exception:
        return {}


def _tcb_configured():
    """腾讯云那条链是否还配置着（过渡期镜像发布用）。"""
    t = load_json(COS_CFG).get("tcb") or {}
    cli = t.get("cli") or _portable.resolve_tcb()
    return bool(t.get("env_id") and cli and os.path.exists(cli))


def write_publish_request(build_ok=True, note=""):
    """把「本地已构建完成、待发布」落成一个标记文件，供 WorkBuddy 自动化读取。

    WorkBuddy 的发布只能由 WorkBuddy 客户端发起（脚本无法直连），所以脚本这端
    只负责把「该发什么」写清楚：构建指纹、时间、目标目录、入口页、应用名。
    自动化任务只需比 build_hash 与 last_published_hash —— 不同才需要发布，
    整个判断不碰网络、秒级完成。
    """
    fh = ""
    try:
        fh = (load_json(os.path.join(DEPLOY, "_build_fingerprint.json")) or {}).get("build_hash") or ""
    except Exception:
        pass
    last = ""
    try:                                   # 首次运行时该文件还不存在，静默即可
        with open(PUBLISHED_PATH, encoding="utf-8") as f:
            last = (json.load(f) or {}).get("build_hash") or ""
    except Exception:
        last = ""
    pub = _pub()
    data = {
        "built_at": now_str(),
        "build_ok": bool(build_ok),
        "build_hash": fh,
        "last_published_hash": last,
        "need_publish": bool(fh) and fh != last,
        "site_dir": _portable.site_dir(ROOT),
        "entry_html": pub.get("entry_html") or "index.html",
        "app_name": pub.get("app_name") or "某公司看板",
        "domain_prefix": pub.get("domain_prefix") or "",
        "dash_url": pub.get("dash_url") or "",
        "nav_dir": _portable.nav_dir(ROOT),
        "nav_entry_html": "index.html",
        "nav_app_name": pub.get("nav_app_name") or "看板导航页",
        "nav_domain_prefix": pub.get("nav_domain_prefix") or "",
        "nav_url": pub.get("nav_url") or "",
        "note": note or "",
    }
    save_json(PUBLISH_REQ, data)
    return fh


def mark_published():
    """记录「当前构建指纹已经发到线上」——自动化发布成功后调用。

    之后 write_publish_request 就能靠 build_hash 与它比对，判断要不要再发。
    """
    fh = ""
    try:
        fh = (load_json(os.path.join(DEPLOY, "_build_fingerprint.json")) or {}).get("build_hash") or ""
    except Exception:
        pass
    if not fh:
        log("mark_published：读不到 _build_fingerprint.json，跳过")
        return False
    save_json(PUBLISHED_PATH, {"build_hash": fh, "published_at": now_str(),
                               "dash_url": _pub().get("dash_url") or ""})
    log("已记录线上指纹：%s（%s）" % (fh, now_str()))
    return True


def stage_publish():
    """发布阶段（V1.5）。

    channel = workbuddy（当前方案）：本地构建完就到此为止，发布交给 WorkBuddy
        上的自动化任务 —— 脚本把请求写清楚，自动化比对指纹后调「发布为应用」。
    channel = tencent（老配置）：仍走原来的 tcb 全流程。
    mirror_tencent = true：无论上面哪种，都顺带把内容镜像发一份到腾讯云，
        过渡期两条链接都不掉队；腾讯云账号收回后改成 false 即可。
    """
    pub = _pub()
    chan = (pub.get("channel") or "workbuddy").lower()
    mirrored = None

    if chan == "tencent":
        mirrored = deploy_tcb()
        fh = write_publish_request(True, "tencent 通道")
        log("发布通道 tencent：完成（build_hash %s）" % (fh or "-"))
        return mirrored

    if pub.get("mirror_tencent") and _tcb_configured():
        log("过渡期镜像：顺带发一份到腾讯云（mirror_tencent=true）…")
        mirrored = deploy_tcb()
    elif pub.get("mirror_tencent") and not _tcb_configured():
        log("过渡期镜像已开启，但腾讯云配置不可用 → 跳过（不影响 WorkBuddy 发布）")

    fh = write_publish_request(True, "workbuddy 通道待发布")
    log("WorkBuddy 发布：本地已就绪 build_hash=%s → 等自动化发布" % (fh or "-"))
    log("  发布目录 %s（入口 %s）" % (_portable.site_dir(ROOT), pub.get("entry_html") or "index.html"))
    if not (pub.get("dash_url") or ""):
        log("  注意：还没回填线上链接，发布拿到 shareLink 后写进「迁移配置.json」的 dash_url")
    if mirrored is False:
        log("  （镜像到腾讯云失败，但不影响 WorkBuddy 这条链）")
    return True


def deploy_tcb():
    """发布到腾讯云 CloudBase 静态托管。

    V1.5 起它退居「过渡期镜像 / 老配置通道」：主发布已改到 WorkBuddy 分享链接。
    腾讯云账号收回后，把「迁移配置.json」的 mirror_tencent 改成 false 即可停用。

    V1.4.9 起发布后独立核对线上指纹与关键数据文件，不一致就原样补发一次。
    起因：曾出现日志显示「发布成功」、线上却仍是旧内容（会计页整页无数据）
    且长期没人察觉。不用 tcb 自带的 --verify —— 实测它会把 /hist/... 当成
    远端路径去比而误报失败（mismatched=[] 却退出码 5），线上核对才是真判据。
    """
    t = load_json(COS_CFG).get("tcb") or {}
    node, cli, env = t.get("node"), t.get("cli"), t.get("env_id")
    node = _resolve_node(node)
    cli = cli if (cli and os.path.exists(cli)) else _portable.resolve_tcb()
    if not (node and cli and env) or not os.path.exists(cli):
        log("CloudBase 未配置或 CLI 缺失 → 跳过发布")
        return False
    cwd = os.path.dirname(os.path.dirname(os.path.dirname(cli)))

    def _push(extra):
        cmd = [node, cli, "hosting", "deploy", DEPLOY, "/dash", "-e", env] + list(extra)
        t0 = time.time()
        r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                           errors="ignore", cwd=cwd if os.path.isdir(cwd) else None,
                           timeout=1200)
        out = (r.stdout or "") + (r.stderr or "")
        # tcb CLI 的输出文案随版本变化（旧版 "File upload completed"，新版是进度条+访问地址），
        # 故以退出码为主判据，文案仅作辅助，避免把成功误判成"存疑"。
        ok = (r.returncode == 0) or ("File upload completed" in out) or ("Successfully uploaded" in out)
        return ok, out, r.returncode, time.time() - t0

    try:
        ok, out, rc, dt = _push([])
    except Exception as e:
        log("CloudBase 发布失败：%s" % str(e)[:200])
        return False

    vok, vwhy = online_verify()
    if vok:
        log("CloudBase 发布：%s（%.1fs）｜线上核对 √ %s"
            % ("成功" if ok else "退出码 %s" % rc, dt, vwhy))
        return True

    # 线上内容不对：原样补发一次（--safe 会在上传/校验失败时自动回滚）
    # 这行**故意不带**「发布：成功」字样 —— 控制台状态栏取日志里最后一次匹配，
    # 若写成"成功"会把「线上核对没过」这个事实盖掉，等于又回到今天的老问题。
    log("CloudBase 发布：线上核对未通过（%.1fs）｜%s → 原样补发一次" % (dt, vwhy))
    if not ok:
        log("          上传命令退出码 %s" % rc)
        log("  " + out[-300:].replace("\n", " | "))
    dt2 = 0.0
    try:
        ok2, out2, rc2, dt2 = _push(["--safe"])
        if not ok2:
            log("CloudBase 补发：失败（%.1fs，退出码 %s）" % (dt2, rc2))
            log("  " + out2[-300:].replace("\n", " | "))
    except Exception as e:
        log("CloudBase 补发失败：%s" % str(e)[:200])
    vok2, vwhy2 = online_verify()
    if vok2:
        log("CloudBase 发布：成功（含补发 %.1fs）｜线上核对 √ %s" % (dt2, vwhy2))
        return True
    log("CloudBase 发布：失败（补发后线上仍不一致）｜%s" % vwhy2)
    log("  → 到控制台「运行环境 → 🌐 线上核对」排查，或手动跑一次强制全量刷新")
    return False


def run_build():
    rc = subprocess.run([sys.executable, os.path.join(ROOT, "build_snapshot.py")], cwd=ROOT,
                        capture_output=True, text=True, encoding="utf-8",
                        errors="ignore", timeout=1800)
    return rc.returncode, (rc.stdout or ""), (rc.stderr or "")


# --------------------------------------------------------------------------
def main():
    argv = set(a for a in sys.argv[1:] if a.startswith("-"))
    force = "--force" in argv
    dry = "--dry" in argv
    skip_bi = "--skip-bi" in argv

    bi_engine.log = log
    log("=== 看板刷新 %s ===  参数: %s" % (now_str(), " ".join(sorted(argv)) or "无"))
    author_note()

    # 记录「本次构建已发布到线上」：python bi_refresh.py --mark-published
    if "--mark-published" in argv:
        ok = mark_published()
        flush_log()
        return 0 if ok else 1

    # 只做线上核对（不取数、不重建、不发布）：python bi_refresh.py --check-online
    if "--check-online" in argv:
        ok, why = online_verify()
        log("线上核对：%s %s" % ("√" if ok else "×", why))
        flush_log()
        return 0 if ok else 1

    cfg = load_json(CONFIG)
    if not cfg or not cfg.get("sources"):
        log("bi_sources.json 没有数据源，退出")
        flush_log()
        return 1

    okwin, winnote = in_window(cfg.get("window") or {}, force)
    if not okwin:
        log("不在更新时段 %s（现在 %s）→ 跳过" % (winnote, datetime.now().strftime("%H:%M")))
        flush_log()
        return 2

    if not acquire_lock():
        flush_log()
        return 3

    state = load_json(STATE_PATH, {})
    last_success = state.get("last_success", "")
    skip, why = throttle_skip(cfg, state, force)
    if skip:
        log("跳过本轮：%s（时段节流）" % why)
        release_lock()
        flush_log()
        return 0
    if "未配置节流" not in why:
        log("频率检查：%s" % why)
    files = cfg.get("files") or {}
    sources = cfg.get("sources") or []
    recalc_files = cfg.get("recalc_files") or DEFAULT_RECALC
    cache_dir = os.path.join(BASE, cfg.get("cache_dir") or "_bi_cache")

    # --- 环境相关可调项（缺省值与历史行为一致）---
    _http_timeout = int(cfg.get("http_timeout") or 0)
    if _http_timeout > 0:
        bi_engine._TIMEOUT = _http_timeout
    _retries = int(cfg.get("retries") or 2)
    _wait_idle = int(cfg.get("wait_idle_sec") or 45)

    stale = False
    err = ""
    note = ""
    written = 0

    try:
        # ---------- 阶段 1：下载 + 判变 ----------
        paths = {}
        changed = {}
        failed = []
        if skip_bi:
            log("阶段1 跳过（--skip-bi）")
            note = "已跳过取数"
        else:
            t0 = time.time()
            paths = bi_engine.download_all(
                sources, cache_dir, cfg.get("bi_login_no", ""),
                workers=int(cfg.get("workers") or 5), retries=_retries)
            active = [s for s in sources if s.get("enabled", True) and s.get("url")]
            for s in active:
                sid = s.get("id")
                p = paths.get(sid)
                if not p:
                    failed.append(sid)
                    continue
                h = bi_engine.file_md5(p)
                if force or state.get("md5", {}).get(sid) != h:
                    changed[sid] = h
            log("阶段1 完成（%.1fs）：变化 %d / 失败 %d / 共 %d"
                % (time.time() - t0, len(changed), len(failed), len(active)))
            note = "下载 %d，变化 %d，失败 %d" % (len(active), len(changed), len(failed))

            if failed and not changed:
                # 全部拿不到新数据：保持线上不动，只标记
                stale = True
                err = "BI 下载失败：" + "、".join(failed)
                log("阶段1 全部失败 → 保持上次数据")

        # ---------- 阶段 2：写 Excel ----------
        ops_by_file = {}
        if not skip_bi and changed and not (failed and not changed):
            for s in sources:
                sid = s.get("id")
                if sid not in changed or not paths.get(sid):
                    continue
                sheet = s.get("sheet") or ""
                if sheet.lower() in ("", "sheet1"):
                    sheet = None
                try:
                    values = bi_engine.read_sheet_values(paths[sid], sheet)
                except Exception as e:
                    log("  ✗ %s 读取失败：%s" % (sid, str(e)[:120]))
                    failed.append(sid)
                    continue
                if not values:
                    log("  ⚠ %s 下载内容为空 → 跳过" % sid)
                    failed.append(sid)
                    continue
                for t in (s.get("targets") or []):
                    fk = t.get("file")
                    if fk not in files:
                        log("  ⚠ 未知目标文件 %s → 跳过" % fk)
                        continue
                    ops_by_file.setdefault(fk, []).append((t.get("sheet"), values))
                log("  · %-20s %d 行 → %d 个目标表"
                    % (sid, len(values), len(s.get("targets") or [])))

        if ops_by_file:
            targets = [files[fk] for fk in ops_by_file]
            daily_backup(targets)
            locked = [p for p in targets if bi_engine.is_locked(p)]
            if locked:
                log("⚠ 这些文件正被打开，本轮跳过写入：%s"
                    % "、".join(os.path.basename(p) for p in locked))
                err = "源 Excel 被占用：" + "、".join(os.path.basename(p) for p in locked)
                stale = True
                ops_by_file = {}
            else:
                t0 = time.time()
                with bi_engine.ExcelSession(cfg.get("com_engine", "auto")) as xl:
                    for fk, ops in ops_by_file.items():
                        written += xl.write_file(files[fk], ops, recalc=(fk in recalc_files))
                log("阶段2 写 Excel 完成（%.1fs，%d 张表）" % (time.time() - t0, written))
                waited = wait_files_idle(targets, timeout=_wait_idle)
                if waited:
                    log("  等待 Excel 释放句柄 %.1fs" % waited)

        # ---------- 阶段 3：重建 ----------
        built = False
        need_build = (written > 0) or skip_bi or force
        if not need_build:
            log("阶段3 跳过：数据无变化，不重建也不发布")
            save_json(STATE_PATH, dict(state, last_run=now_str(),
                                       last_noop=now_str(),
                                       runs=state.get("runs", 0) + 1))
            release_lock()
            flush_log()
            return 0

        t0 = time.time()
        rc, out, e2 = run_build()
        built = (rc == 0)
        if not built:
            log("阶段3 首次重建失败（exit=%s）→ 等 25s 重试一次" % rc)
            log("  失败摘要：%s" % (out + e2)[-400:].replace("\n", " | "))
            time.sleep(25)
            wait_files_idle([files[fk] for fk in ops_by_file] or [], timeout=30)
            rc, out, e2 = run_build()
            built = (rc == 0)
        if built:
            log("阶段3 重建看板完成（%.1fs）" % (time.time() - t0))
            last_success = now_str()
        else:
            stale = True
            err = "build_snapshot 失败 exit=%s（已重试一次）" % rc
            log("阶段3 重建仍失败：%s" % (out + e2)[-600:].replace("\n", " | "))

        # ---------- 阶段 4：状态 + 发布 ----------
        pub_now = _pub()
        write_status(stale, last_success, err,
                     {"stage1": note[:200], "written": written, "built": built,
                      "partial": failed[:10],
                      "channel": pub_now.get("channel") or "",
                      "dash_url": pub_now.get("dash_url") or "",
                      "nav_url": pub_now.get("nav_url") or ""})

        if dry:
            log("阶段4 跳过发布（--dry）")
        elif not cfg.get("deploy_enabled", True):
            log("阶段4 跳过发布（配置 deploy_enabled=false，仅本地更新）")
            write_publish_request(built, "deploy_enabled=false")
        else:
            if not built:
                # 重建失败：线上保持上一轮数据，只把「数据未更新」标记推上去
                log("阶段4 重建未成功，仅同步「数据未更新」标记")
            stage_publish()

        # 判变基准只在写成功之后才更新，否则下一轮会误判"没变"
        if written > 0 and not failed:
            state["md5"] = dict(state.get("md5", {}), **changed)
        save_json(STATE_PATH, dict(state, last_run=now_str(),
                                   last_success=last_success,
                                   last_ok=(not stale),
                                   last_error=err,
                                   runs=state.get("runs", 0) + 1))
        return 0 if not stale else 1

    except Exception:
        log("未捕获异常：\n" + traceback.format_exc(limit=5))
        write_status(True, last_success, "运行异常", {"stage1": note[:200]})
        return 1
    finally:
        release_lock()
        flush_log()


# --------------------------------------------------------------------------
# 令牌完整性诊断（仅记录，不做任何重启/降权动作）
#   历史背景：计划任务曾以高完整性（提权）令牌运行，此时 WPS 的用户级 COM 注册
#   不可见，Excel COM 无法启动；中等完整性令牌下正常。修复见 _fix_com_hklm.py
#   （把 WPS 的 COM 注册补到机器级 HKLM，使高完整性进程也能激活）。
# --------------------------------------------------------------------------
def _integrity_rid():
    """当前进程令牌的完整性级别 RID（Medium=0x2000, High=0x3000）；探测失败返回 None。"""
    try:
        import win32api
        import win32security
        tok = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32security.TOKEN_QUERY)
        try:
            sid = win32security.GetTokenInformation(tok, win32security.TokenIntegrityLevel)[0]
            return sid.GetSubAuthority(sid.GetSubAuthorityCount() - 1)
        finally:
            win32api.CloseHandle(tok)
    except Exception:
        return None


def main_entry():
    rid = _integrity_rid()
    if rid is not None:
        log('进程完整性级别 IL=0x%X（%s）' % (rid, '高/提权' if rid >= 0x3000 else '中/标准'))
    return main()


if __name__ == "__main__":
    sys.exit(main_entry())
