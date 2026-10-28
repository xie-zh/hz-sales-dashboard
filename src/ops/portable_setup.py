# -*- coding: utf-8 -*-
"""可移植层 V1.0 —— 路径解析 + 环境体检 + 一键修复 + 搬家包

目标：把 8.11 整个目录拷到**任意电脑、任意路径**，程序都能自己找到
数据目录 / Python 解释器 / node / CloudBase CLI，不再依赖写死的
<LOCAL_PATH>\\... 路径。

解析优先级（每一类都一样）：
    1. Hzexamplebot/_paths.json   ← 初始化向导写入的显式配置（最高）
    2. 环境变量 HZEXAMPLE_ROOT / HZEXAMPLE_PY / HZEXAMPLE_NODE / HZEXAMPLE_TCB
    3. 自动探测（本机常见位置 + 目录特征文件）
    4. 兜底（维持旧行为，或返回 None 让调用方给出提示）

被 bi_refresh.py 与 hzexamplebot_gui.py 共用。
"""

import json
import os
import re
import shutil
import subprocess
import sys

# ---------- 自身位置（兼容脚本运行 / PyInstaller 打包） ----------
if getattr(sys, "frozen", False):
    BASE = os.path.dirname(os.path.abspath(sys.executable))   # 打包后 = exe 所在目录
else:
    BASE = os.path.dirname(os.path.abspath(__file__))         # 脚本 = Hzexamplebot/

PATHS_FILE = os.path.join(BASE, "_paths.json")

# 根目录特征文件：build_snapshot.py 必须在；配置文件「迁移配置.json / _cos_config.json」
# 至少有一个（老包只有后者，新包两者都有）。这样即使配置文件名变了也能认路。
ROOT_MARKERS = ("build_snapshot.py",)
ROOT_ALT_MARKERS = ("_cos_config.json", "迁移配置.json")

# 项目必备脚本（体检用）
NEED_SCRIPTS_BOT = ("bi_refresh.py", "bi_engine.py", "bi_sources.json", "hzexamplebot_gui.py")
NEED_SCRIPTS_ROOT = ("build_snapshot.py",)

# 源数据文件（体检用；按后缀在根目录找，不写死文件名）
DATA_PATTERNS = {
    "人员路径": "某公司-人员数据分析路径.xlsx",
    "员工三力": "某公司-员工三力明细.xlsx",
    "活动量": "某公司活动量-在线看板.xlsx",
}

DEP_MODULES = (("openpyxl", "openpyxl"), ("win32com", "pywin32"))

WHEEL_INDEX = "https://pypi.tuna.tsinghua.edu.cn/simple"


# ================= 基础工具 =================

def _load_paths():
    try:
        with open(PATHS_FILE, encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_paths(d):
    cur = _load_paths()
    cur.update(d or {})
    try:
        with open(PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=1)
    except Exception:
        pass
    return PATHS_FILE


def _env(name):
    v = os.environ.get(name)
    return v.strip() if v and v.strip() else None


def _isfile(p):
    return bool(p) and os.path.isfile(p)


def _is_root(d):
    """这个目录是不是 8.11 项目根。"""
    if not d or not os.path.isdir(d):
        return False
    if not all(os.path.exists(os.path.join(d, m)) for m in ROOT_MARKERS):
        return False
    return any(os.path.exists(os.path.join(d, m)) for m in ROOT_ALT_MARKERS)


def _home():
    return os.path.expanduser("~")


def _user_profiles():
    """可能存在的用户目录（同一台机器可能换过账号名）。"""
    out = []
    up = os.environ.get("USERPROFILE") or _home()
    out.append(up)
    drive = os.path.splitdrive(up)[0] + "\\"
    users = os.path.join(drive, "Users")
    if os.path.isdir(users):
        for n in os.listdir(users):
            p = os.path.join(users, n)
            if os.path.isdir(p):
                out.append(p)
    # 去重保序
    seen, res = set(), []
    for p in out:
        k = p.lower()
        if k not in seen:
            seen.add(k)
            res.append(p)
    return res


# ================= 四类路径解析 =================

def resolve_root():
    """8.11 项目根目录（源 Excel / _deploy / 历史数据 / 配置文件所在）。"""
    p = _load_paths().get("root")
    if _is_root(p):
        return p

    ev = _env("HZEXAMPLE_ROOT")
    if ev and os.path.isdir(ev):
        return ev

    # 自动探测：从 BASE 及其父级逐层向上找特征文件
    d = BASE
    for _ in range(5):
        if _is_root(d):
            return d
        nd = os.path.dirname(d)
        if nd == d:
            break
        d = nd

    return os.path.dirname(BASE)   # 兜底：维持旧行为


def resolve_python():
    """Python 解释器。优先本地可移植 venv（<root>\\_env），便于整包搬走。"""
    cands = []
    p = _load_paths().get("python")
    if p:
        cands.append(p)
    ev = _env("HZEXAMPLE_PY")
    if ev:
        cands.append(ev)

    root = resolve_root()
    for base in (root, BASE):
        cands.append(os.path.join(base, "_env", "Scripts", "python.exe"))
        cands.append(os.path.join(base, "_env", "python.exe"))

    # 同机器其它账号下的 WorkBuddy 受管 python
    for up in _user_profiles():
        cands.append(os.path.join(up, ".workbuddy", "binaries", "python",
                                  "envs", "default", "Scripts", "python.exe"))
        vroot = os.path.join(up, ".workbuddy", "binaries", "python", "versions")
        if os.path.isdir(vroot):
            vers = sorted(d for d in os.listdir(vroot) if d[0].isdigit())
            for d in reversed(vers):
                cands.append(os.path.join(vroot, d, "python.exe"))
        # 便携版 python 常见安装位置
        la = os.path.join(up, "AppData", "Local", "Programs", "Python")
        if os.path.isdir(la):
            for d in sorted(os.listdir(la), reverse=True):
                cands.append(os.path.join(la, d, "python.exe"))

    for c in cands:
        if _isfile(c):
            return c

    for name in ("python.exe", "python3.exe", "python"):
        w = shutil.which(name)
        if w:
            return w

    return sys.executable


def resolve_node(prefer=None):
    """node.exe。扫描 versions 目录取版本号最大者。"""
    for c in (prefer, _load_paths().get("node"), _env("HZEXAMPLE_NODE")):
        if _isfile(c):
            return c

    cands = []
    for up in _user_profiles():
        vroot = os.path.join(up, ".workbuddy", "binaries", "node", "versions")
        if os.path.isdir(vroot):
            for d in os.listdir(vroot):
                exe = os.path.join(vroot, d, "node.exe")
                if _isfile(exe):
                    cands.append((d, exe))
    if cands:
        cands.sort()          # 字典序对 "22.22.2-6" 这类目录名够用
        return cands[-1][1]

    return shutil.which("node")


def resolve_tcb(prefer=None):
    """CloudBase CLI 入口（node 脚本 tcb）。"""
    for c in (prefer, _load_paths().get("tcb_cli"), _env("HZEXAMPLE_TCB")):
        if _isfile(c):
            return c

    root = resolve_root()
    rel = os.path.join("node_modules", "@cloudbase", "cli", "bin", "tcb")
    cands = [os.path.join(root, rel), os.path.join(BASE, rel)]
    for up in _user_profiles():
        cands.append(os.path.join(up, ".workbuddy", "binaries", "node",
                                  "workspace", rel))
    for c in cands:
        if _isfile(c):
            return c

    return shutil.which("tcb")


# ================= 依赖体检 =================

def deps_missing(py):
    """返回该解释器缺失的依赖列表（模块名）。"""
    if not _isfile(py):
        return [m for _, m in DEP_MODULES]
    code = ("import importlib,sys\n"
            "miss=[]\n"
            "for m in %r:\n"
            "    try: importlib.import_module(m)\n"
            "    except Exception: miss.append(m)\n"
            "print('|'.join(miss))\n") % ([m for m, _ in DEP_MODULES],)
    try:
        r = subprocess.run([py, "-c", code], capture_output=True, text=True,
                           encoding="utf-8", errors="ignore", timeout=60,
                           creationflags=0x08000000 if os.name == "nt" else 0)
        miss = [x for x in (r.stdout or "").strip().split("|") if x]
        return miss
    except Exception:
        return [m for _, m in DEP_MODULES]


# ================= 环境体检 =================

def publish_info(root=None):
    """发布参数（V1.5：统一走 _migrate 配置层）。取不到时返回空 dict。"""
    root = root or resolve_root()
    try:
        import _migrate
        return _migrate.publish(root)
    except Exception:
        cfg, _ = _read_cloud_cfg(root)
        return cfg.get("publish") or {}


def cloud_url(root=None):
    """线上主看板地址（以 / 结尾）。未配置时返回空串。

    V1.5：配置里的 publish.dash_url 优先；老包只有腾讯云 tcb.url 时兼容读它。
    """
    root = root or resolve_root()
    info = publish_info(root)
    u = (info.get("dash_url") or "").strip()
    if not u:
        cfg, _ = _read_cloud_cfg(root)
        u = ((cfg.get("tcb") or {}).get("url") or "").strip()
    u = u.rstrip("/")
    return (u + "/") if u else ""


def nav_url(root=None):
    """线上导航页地址（以 / 结尾）。"""
    u = (publish_info(root or resolve_root()).get("nav_url") or "").strip()
    return (u.rstrip("/") + "/") if u else ""


def site_dir(root=None):
    """待发布的主看板目录（绝对路径）。"""
    root = root or resolve_root()
    return os.path.join(root, publish_info(root).get("site_dir") or "_deploy")


def nav_dir(root=None):
    root = root or resolve_root()
    d = publish_info(root).get("nav_dir") or os.path.join("deploy_tool", "build")
    return d if os.path.isabs(d) else os.path.join(root, d)


def scan(root=None, py=None):
    """整机环境体检。返回 rows = [{key,label,ok,level,detail,fixable}]。

    level: 'ok' 正常 / 'warn' 可测但存疑 / 'bad' 会导致跑不动
    """
    root = root or resolve_root()
    py = py or resolve_python()
    rows = []

    def add(key, label, ok, detail, level="ok", fixable=False):
        rows.append({"key": key, "label": label, "ok": bool(ok), "level": level,
                     "detail": detail, "fixable": bool(fixable)})

    # 1 项目目录
    root_ok = _is_root(root)
    add("root", "项目目录", root_ok,
        root if root_ok else "未找到（需含 build_snapshot.py 与配置文件）：%s" % root,
        "ok" if root_ok else "bad")

    # 2 脚本齐全
    miss = [f for f in NEED_SCRIPTS_BOT if not _isfile(os.path.join(BASE, f))]
    miss += [f for f in NEED_SCRIPTS_ROOT if not _isfile(os.path.join(root, f))]
    add("scripts", "程序脚本", not miss,
        "齐全" if not miss else "缺失：%s" % "、".join(miss),
        "ok" if not miss else "bad")

    # 3 Python 解释器
    py_ok = _isfile(py)
    ver = ""
    if py_ok:
        try:
            r = subprocess.run([py, "-c", "import sys;print('.'.join(map(str,sys.version_info[:3])))"],
                               capture_output=True, text=True, encoding="utf-8",
                               errors="ignore", timeout=30,
                               creationflags=0x08000000 if os.name == "nt" else 0)
            ver = (r.stdout or "").strip()
        except Exception:
            ver = "?"
    add("python", "Python 解释器", py_ok,
        ("%s（%s）" % (py, ver)) if py_ok else "没找到可用解释器，点「一键修复」自动准备",
        "ok" if py_ok else "bad", fixable=True)

    # 4 必要依赖
    miss_dep = deps_missing(py) if py_ok else [m for _, m in DEP_MODULES]
    add("deps", "依赖库（openpyxl / pywin32）", not miss_dep,
        "已安装" if not miss_dep else "缺：%s，点「一键修复」自动安装" % "、".join(miss_dep),
        "ok" if not miss_dep else "bad", fixable=True)

    # 5 源数据文件
    md = [n for n, fn in DATA_PATTERNS.items() if not _isfile(os.path.join(root, fn))]
    add("data", "源数据 Excel", not md,
        "三张主表齐全" if not md else "缺：%s（放到 %s 根目录）" % ("、".join(md), os.path.basename(root)),
        "ok" if not md else "warn")

    # 6 WPS / Excel COM
    com_wps = _com_registered("KET.Application")
    com_xls = _com_registered("Excel.Application")
    if com_wps or com_xls:
        add("com", "Excel/WPS（COM 写入）", True,
            "WPS=%s  Excel=%s" % ("√" if com_wps else "×", "√" if com_xls else "×"))
    else:
        add("com", "Excel/WPS（COM 写入）", False,
            "未检测到，需自行安装 WPS 或 Microsoft Excel（注意 64 位版）",
            "bad")

    # 7 node + CloudBase CLI
    node = resolve_node()
    tcb = resolve_tcb()
    add("node", "node（发布用）", bool(node), node or "未找到，若只需本地更新可不装",
        "ok" if node else "warn")
    add("tcb", "CloudBase CLI（发布用）", bool(tcb), tcb or "未找到，若不上传云端可不装",
        "ok" if tcb else "warn")

    # 8 发布配置（V1.5：不再看腾讯云密钥，改看 WorkBuddy 发布参数）
    cfg, cfg_err = _read_cloud_cfg(root)
    pub = publish_info(root)
    chan = (pub.get("channel") or "").strip()
    site = site_dir(root)
    has_site = os.path.isdir(site)
    if cfg_err:
        add("cloud", "发布配置", False, cfg_err, "bad")
    elif chan == "workbuddy":
        if has_site:
            add("cloud", "发布配置", True,
                "通道 workbuddy，产物目录 %s" % os.path.basename(site))
        else:
            add("cloud", "发布配置", False,
                "通道已配 workbuddy，但缺产物目录 %s（跑一次构建即可）" % site,
                "warn", fixable=True)
    else:
        add("cloud", "发布配置", False,
            "没有 publish 块（点「填写发布与 BI 参数」补上）", "bad", fixable=True)

    # 9 BI 工号
    login = ""
    try:
        with open(os.path.join(BASE, "bi_sources.json"), encoding="utf-8") as f:
            login = str((json.load(f) or {}).get("bi_login_no") or "")
    except Exception:
        pass
    add("login", "BI 取数工号", bool(login), login or "未填，点「填写云端与 BI 凭据」补上",
        "ok" if login else "bad", fixable=True)

    # 10 计划任务
    ti = task_info("HZBoardRefresh")
    if not ti["exists"]:
        add("task", "自动刷新任务", False, "未注册（点「注册自动刷新任务」一键建）",
            "warn", fixable=True)
    elif ti["state"] == "Disabled":
        add("task", "自动刷新任务", False, "已存在但被禁用（在「任务计划程序」里启用即可）",
            "warn", fixable=True)
    else:
        cmd = ti["command"] or ""
        stale = ""
        mm = re.search(r'"?([^"]*?python\.exe)"?', cmd, re.I)
        if mm and not _isfile(mm.group(1)):
            stale = "，但指向的 Python 已失效（点「注册自动刷新任务」重建）"
        add("task", "自动刷新任务", not stale,
            "运行中，每 %d 分钟唤醒一次%s" % (ti["minute"] or 5, stale),
            "ok" if not stale else "warn", fixable=True)

    # 11 线上可达
    url = cloud_url(root)
    if url:
        reach = _http_ok(url)
        add("online", "线上看板可达", reach,
            ("%s 正常" % url) if reach else "%s 打不开（可能断网，或应用还没发布）" % url,
            "ok" if reach else "warn")
    else:
        add("online", "线上看板可达", False,
            "还没发布过（发布后把链接回填到「迁移配置.json」的 dash_url）", "warn")

    # 12 迁移落地情况（源表路径是否都指向本机）
    try:
        import _migrate
        rows2 = [r for r in _migrate.selfcheck(root) if r["key"] == "files"]
        if rows2:
            r = rows2[0]
            add("migrate", "源表路径（本机）", r["ok"], r["detail"], r["level"])
    except Exception as e:
        add("migrate", "源表路径（本机）", False, "自检不可用：%s" % str(e)[:60], "warn")

    return rows


def summary(rows):
    bad = [r for r in rows if r["level"] == "bad"]
    warn = [r for r in rows if r["level"] == "warn"]
    return len(bad), len(warn)


def _read_cloud_cfg(root):
    p = os.path.join(root, "_cos_config.json")
    if not _isfile(p):
        return {}, "找不到 _cos_config.json（云端配置文件）"
    try:
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
        return (d if isinstance(d, dict) else {}), ""
    except Exception as e:
        return {}, "解析失败：%s" % str(e)[:80]


def _com_registered(progid):
    try:
        import winreg
    except Exception:
        return False
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            k = winreg.OpenKey(hive, r"Software\Classes\%s\CLSID" % progid)
            winreg.CloseKey(k)
            return True
        except Exception:
            continue
    return False


def task_info(name="HZBoardRefresh"):
    """读计划任务定义（不依赖 schtasks，避免被杀软/沙箱拦）。

    返回 dict: {exists, state, command, minute}
    state: 'Ready' / 'Disabled' / ''
    """
    info = {"exists": False, "state": "", "command": "", "minute": 0}
    if os.name != "nt":
        return info
    sr = os.environ.get("SystemRoot", r"C:\WINDOWS")
    p = os.path.join(sr, "System32", "Tasks", name)
    if os.path.isfile(p):
        info["exists"] = True
        try:
            txt = open(p, "r", encoding="utf-16", errors="ignore").read()
        except Exception:
            txt = ""
        info["state"] = "Disabled" if "<Enabled>false</Enabled>" in txt else "Ready"
        m = re.search(r"<Command>(.*?)</Command>", txt)
        if m:
            info["command"] = m.group(1).strip()
        m = re.search(r"<Interval>PT(\d+)M</Interval>", txt)
        if m:
            info["minute"] = int(m.group(1))
        return info

    # 退路：schtasks 查询（部分环境该程序被禁，失败即当未注册）
    r = _run_schtasks(["/query", "/tn", name])
    if r.returncode == 0:
        info["exists"] = True
        out = r.stdout or ""
        if "Disabled" in out or "已禁用" in out:
            info["state"] = "Disabled"
        else:
            info["state"] = "Running" if ("Running" in out or "正在运行" in out) else "Ready"
    return info


def _task_state(name):
    return task_info(name).get("state", "")


def _http_ok(url, timeout=15):
    import urllib.request
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Hzexamplebot/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return 200 <= getattr(r, "status", 200) < 400
    except Exception:
        return False


# ================= 一键修复 =================

def base_python():
    """用来创建 venv 的「底包」解释器（优先系统 Python，其次受管 Python）。"""
    cands = []
    for up in _user_profiles():
        la = os.path.join(up, "AppData", "Local", "Programs", "Python")
        if os.path.isdir(la):
            for d in sorted(os.listdir(la), reverse=True):
                cands.append(os.path.join(la, d, "python.exe"))
    for up in _user_profiles():
        vroot = os.path.join(up, ".workbuddy", "binaries", "python", "versions")
        if os.path.isdir(vroot):
            for d in sorted((x for x in os.listdir(vroot) if x[0].isdigit()), reverse=True):
                cands.append(os.path.join(vroot, d, "python.exe"))
    for name in ("python.exe", "python"):
        w = shutil.which(name)
        if w:
            cands.append(w)
    cands.append(sys.executable)
    for c in cands:
        if _isfile(c):
            return c
    return None


def ensure_venv(root=None, log=print):
    """在 <root>\\_env 建一个可移植 venv 并装依赖。返回 (ok, python_path或错误信息)。

    这样整包搬到别的电脑时，Python 环境跟着走，不再依赖某个账号下的 venv。
    """
    root = root or resolve_root()
    env_dir = os.path.join(root, "_env")
    py = os.path.join(env_dir, "Scripts", "python.exe")

    if not _isfile(py):
        base = base_python()
        if not base:
            return False, "本机没有可用的 Python 底包，请先安装 Python 3.10+ 再点一次"
        log("[初始化] 用 %s 创建可移植环境 → %s" % (base, env_dir))
        try:
            r = subprocess.run([base, "-m", "venv", env_dir], capture_output=True,
                               text=True, encoding="utf-8", errors="ignore",
                               timeout=600,
                               creationflags=0x08000000 if os.name == "nt" else 0)
        except Exception as e:
            return False, "创建 venv 失败：%s" % str(e)[:120]
        if not _isfile(py):
            tail = ((r.stdout or "") + (r.stderr or ""))[-300:].replace("\n", " | ")
            return False, "创建 venv 失败：%s" % tail

    miss = deps_missing(py)
    if miss:
        pkgs = [pkg for mod, pkg in DEP_MODULES if mod in miss]
        log("[初始化] 安装依赖：%s" % "、".join(pkgs))
        wheels = os.path.join(root, "_wheels")
        attempts = []
        if os.path.isdir(wheels) and os.listdir(wheels):
            # 优先离线：包内 _wheels 里放了 openpyxl / pywin32 的多个 Python 版本
            attempts.append(("离线包", [py, "-m", "pip", "install", "--no-index",
                                        "--find-links", wheels] + pkgs))
        attempts.append(("镜像站", [py, "-m", "pip", "install", "-i", WHEEL_INDEX] + pkgs))

        r = None
        for tag, cmd in attempts:
            log("  · 用%s安装…" % tag)
            try:
                r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                                   errors="ignore", timeout=1800,
                                   creationflags=0x08000000 if os.name == "nt" else 0)
            except Exception as e:
                log("  · %s失败：%s" % (tag, str(e)[:100]))
                r = None
                continue
            if not deps_missing(py):
                break
            log("  · %s没装上，换下一种方式" % tag)

        left = deps_missing(py)
        if left:
            tail = (((r.stdout or "") + (r.stderr or ""))[-300:].replace("\n", " | ")
                    if r else "无输出")
            return False, "依赖仍缺 %s（离线包与镜像站都试过）：%s" % ("、".join(left), tail)

    save_paths({"root": root, "python": py,
                "node": resolve_node(), "tcb_cli": resolve_tcb()})
    return True, py


def fix_all(root=None, log=print):
    """一键修复：venv+依赖 → 迁移配置落地 → 写 _paths.json。返回 (ok, 消息列表)。"""
    root = root or resolve_root()
    msgs = []
    ok, res = ensure_venv(root, log=log)
    msgs.append("Python 环境：%s" % (res if ok else "× " + str(res)))
    node = resolve_node()
    save_paths({"root": root, "python": (res if ok else resolve_python()),
                "node": node, "tcb_cli": resolve_tcb()})
    msgs.append("已写入 %s" % os.path.basename(PATHS_FILE))

    # V1.5：把「迁移配置.json」落到 bi_sources.json / _cos_config.json
    try:
        import _migrate
        ok2, m2 = _migrate.apply_all(root, log=log)
        msgs.extend(m2)
        ok = ok and ok2
    except Exception as e:
        msgs.append("迁移配置应用失败：%s" % str(e)[:120])
        ok = False
    return ok, msgs


def write_credentials(root, env_id=None, secret_id=None, secret_key=None,
                      region=None, bucket=None, dash_url=None, bi_login=None,
                      nav_url=None, app_name=None, domain_prefix=None,
                      nav_app_name=None, nav_domain_prefix=None):
    """把「发布参数 + BI 工号」写进配置（留空=保留原值）。返回 (ok, 说明)。

    V1.5：改走 _migrate 配置层。腾讯云相关参数（env_id/secret_*/region/bucket）
    保留在签名里只为兼容旧调用方，传入也不再写入 —— 看板已迁到 WorkBuddy 托管，
    这些凭据随账号一并收回。
    """
    root = root or resolve_root()
    try:
        import _migrate
    except Exception as e:
        return False, "缺少 _migrate 模块：%s" % str(e)[:80]

    cfg = _migrate.load(root)
    pub = cfg.setdefault("publish", {})

    def setif(k, v):
        v = (v or "").strip()
        if v:
            pub[k] = v

    setif("dash_url", dash_url)
    setif("nav_url", nav_url)
    setif("app_name", app_name)
    setif("domain_prefix", domain_prefix)
    setif("nav_app_name", nav_app_name)
    setif("nav_domain_prefix", nav_domain_prefix)
    pub.setdefault("channel", "workbuddy")
    if bi_login and bi_login.strip():
        cfg["bi_login_no"] = bi_login.strip()

    ok, res = _migrate.save(root, cfg)
    if not ok:
        return False, res
    ok2, msgs = _migrate.apply_all(root)
    if (secret_id or secret_key) and not dash_url:
        msgs.append("提示：腾讯云密钥已不再需要（看板改为 WorkBuddy 托管），未写入。")
    return ok2, "已保存到「迁移配置.json」；" + "；".join(msgs)


# ================= 计划任务 =================

def register_task(root=None, py=None, minute=5, name="HZBoardRefresh"):
    """注册每 N 分钟自动刷新（登录态运行，不取最高权限）。

    对齐脚本内部节流，任务本身只是「唤醒器」。已有同名任务则先删后建。
    """
    if os.name != "nt":
        return False, "仅 Windows 支持"
    root = root or resolve_root()
    py = py or resolve_python()
    script = os.path.join(BASE, "bi_refresh.py")
    if not _isfile(script):
        # 脚本在 Hzexamplebot 下；若被搬到根目录也能兜
        script = os.path.join(root, "bi_refresh.py")
    cmd = '"%s" "%s"' % (py, script)
    _run_schtasks(["/delete", "/tn", name, "/f"])
    r = _run_schtasks(["/create", "/tn", name, "/sc", "minute", "/mo", str(int(minute)),
                       "/tr", cmd, "/f", "/rl", "LIMITED"])
    if r.returncode == 0:
        return True, "已注册：每 %d 分钟检查一次（脚本内部按时段节流）" % int(minute)
    return False, ((r.stdout or "") + (r.stderr or ""))[-200:].replace("\n", " | ")


def _run_schtasks(args):
    try:
        return subprocess.run(["schtasks"] + args, capture_output=True, text=True,
                              encoding="utf-8", errors="ignore", timeout=60,
                              creationflags=0x08000000)
    except Exception as e:
        class _E:
            returncode = 1
            stdout = ""
            stderr = str(e)
        return _E()


def register_daily_task(root=None, py=None, hhmm="23:59",
                        name="HZDashBoardDailySave"):
    """注册每日历史快照任务（默认 23:59）。已有同名任务则先删后建。"""
    if os.name != "nt":
        return False, "仅 Windows 支持"
    root = root or resolve_root()
    py = py or resolve_python()
    script = os.path.join(BASE, "save_daily.py")
    if not _isfile(script):
        script = os.path.join(root, "save_daily.py")
    if not _isfile(script):
        return False, "找不到 save_daily.py"
    cmd = '"%s" "%s"' % (py, script)
    _run_schtasks(["/delete", "/tn", name, "/f"])
    r = _run_schtasks(["/create", "/tn", name, "/sc", "daily", "/st", hhmm,
                       "/tr", cmd, "/f", "/rl", "LIMITED"])
    if r.returncode == 0:
        return True, "已注册：每天 %s 拍一次历史快照" % hhmm
    return False, ((r.stdout or "") + (r.stderr or ""))[-200:].replace("\n", " | ")


def remove_task(name):
    """删除计划任务（被别的旧任务占了名字、或要换机器时用）。"""
    if os.name != "nt":
        return False, "仅 Windows 支持"
    r = _run_schtasks(["/delete", "/tn", name, "/f"])
    if r.returncode == 0:
        return True, "已删除 %s" % name
    return False, ((r.stdout or "") + (r.stderr or ""))[-160:].replace("\n", " | ")


def list_tasks():
    """列出本项目的计划任务状态。返回 [(名称, 状态, 分钟)]。"""
    out = []
    for name in ("HZBoardRefresh", "HZDashBoardDailySave", "HZDashboardCosSync"):
        ti = task_info(name)
        out.append((name, ti["state"] or ("未见" if not ti["exists"] else "?"),
                    ti["minute"]))
    return out


# ================= 搬家包 =================

# 打包时一律排除（派生数据 / 缓存 / 旧构建 / 备份 / 临时）
# 注意：「历史数据」不在此列 —— 它由 include_history 开关控制（第 858 行），
# 否则该开关会被这里静默吃掉（曾经的 bug）。「历史源」是临时目录，一律排除。
BUNDLE_SKIP_DIRS = {
    "_bi_cache", "_bi_backups", "_backup", "__pycache__", "_trash",
    "废弃", "scratch", "_portable", "_env", "_build_v142c", "_build_v142c2",
    "_build_v144", "_build_v14b", "work", "_com_debug", "_hzcheck_k_debug",
    "_hzcheck_k", "node_modules", "历史源", "sessions", "skills", "config",
    "_imports", "snapshots",
}
BUNDLE_SKIP_EXT = {".old", ".log", ".pyc", ".genie"}
BUNDLE_SKIP_NAME = {"_paths.json", ".mcp"}

# 开发/自检产物：与运行无关，不进交付包（也让交付件里看不到内部实现细节）
_DEV_EXACT = {
    "_bundle_rebuild.py", "_token_probe.py", "_gen_version_info.py",
}
_DEV_SUFFIX = ("_test.py", "_probe.py", "_audit.py", "_test_out.txt")


def _is_dev_artifact(fname):
    low = fname.lower()
    if low in _DEV_EXACT or low.endswith(_DEV_SUFFIX):
        return True
    # 形如 _selftest_result.txt / _layout_report.txt / _v147_swap.txt 的开发输出
    return low.startswith("_") and low.endswith(".txt")

# 与看板无关的宿主/工具残留（IDE 缓存、会话、发布记录等），一并排除
BUNDLE_SKIP_META = {
    "sessions", "skills", ".codebuddy", ".mcp", ".vscode", ".workbuddy",
    "%USERPROFILE%", "_reg_backup", "_已停用脚本",
}


def make_bundle(root=None, dest=None, include_data=True, include_secrets=True,
                include_history=False, include_wheels=True, log=print):
    """生成「搬家包」zip。返回 (ok, zip路径或错误)。

    include_data    : 是否带源 Excel（不带则新机需自己放数据）
    include_secrets : 是否带账号类配置（BI 工号等；给自己用=True；给别人=False）
    include_history : 是否带历史数据（体积大，默认不带；换机续档要带）
    include_wheels  : 是否带离线依赖 wheel（无网环境也能装 openpyxl/pywin32）
    """
    import time as _t
    import zipfile
    root = root or resolve_root()
    if not os.path.isdir(root):
        return False, "项目目录不存在：%s" % root
    stamp = _t.strftime("%Y%m%d")
    outdir = os.path.join(root, "_portable")
    try:
        os.makedirs(outdir)
    except Exception:
        pass
    dest = dest or os.path.join(outdir, "<CITY>看板_搬家包_%s.zip" % stamp)
    dest_abs = os.path.abspath(dest)

    # ⚠ 关键：产出目录自身绝不能被打进包里 —— 否则 os.walk 会读到正在写的
    # zip 并把它再写进去，文件无限膨胀（实测踩过：10 分钟涨到 10.9 GB）。
    skip_dirs = set(BUNDLE_SKIP_DIRS) | set(BUNDLE_SKIP_META)
    if not include_wheels:
        skip_dirs.add("_wheels")
    try:
        rel_out = os.path.relpath(os.path.abspath(outdir), os.path.abspath(root))
        if not rel_out.startswith(".."):
            skip_dirs.add(rel_out.split(os.sep)[0])
    except Exception:
        pass

    # 「给别人」时要摘掉的账号类配置
    secret_names = {"_cos_config.json", "迁移配置.json", "_paths.json",
                    "bi_sources.json"}

    n_files = n_skip = 0
    total = 0
    per_dir = {}
    try:
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for cur, dirs, files in os.walk(root):
                dirs[:] = [d for d in dirs
                           if d not in skip_dirs and not d.startswith("_build_")]
                rel_dir = os.path.relpath(cur, root)
                if rel_dir == ".":
                    rel_dir = ""
                for fn in files:
                    rel = os.path.join(rel_dir, fn) if rel_dir else fn
                    low = fn.lower()
                    src = os.path.join(cur, fn)
                    if os.path.abspath(src) == dest_abs:
                        n_skip += 1
                        continue
                    if fn in BUNDLE_SKIP_NAME or _is_dev_artifact(fn) or ".probe" in low or ".bak" in low:
                        n_skip += 1
                        continue
                    if os.path.splitext(low)[1] in BUNDLE_SKIP_EXT:
                        n_skip += 1
                        continue
                    if ".old_" in low or low.endswith(".exe.old"):
                        n_skip += 1
                        continue
                    # 历史数据
                    if not include_history and rel.replace("\\", "/").startswith("历史数据/"):
                        n_skip += 1
                        continue
                    # 源 Excel
                    if not include_data and low.endswith((".xlsx", ".xlsm", ".xls")):
                        n_skip += 1
                        continue
                    # 账号类配置
                    if not include_secrets and fn in secret_names:
                        n_skip += 1
                        continue
                    try:
                        # 配置文件一律「净化后」入包：永远不含腾讯云密钥，
                        # 避免账号收回后密钥还跟着包到处跑。
                        if fn in ("_cos_config.json", "迁移配置.json"):
                            import _migrate as _mg
                            data = (_mg.clean_cos_dict(root, include_secrets)
                                    if fn == "_cos_config.json"
                                    else _mg.clean_config_dict(root, include_secrets))
                            z.writestr(rel, json.dumps(data, ensure_ascii=False, indent=2))
                            n_files += 1
                            total += len(json.dumps(data, ensure_ascii=False).encode("utf-8"))
                            per_dir["(根目录)"] = per_dir.get("(根目录)", 0) + 1
                            continue
                        z.write(src, rel)
                        n_files += 1
                        total += os.path.getsize(src)
                        top = rel.split(os.sep)[0] if os.sep in rel else "(根目录)"
                        per_dir[top] = per_dir.get(top, 0) + 1
                    except Exception as e:
                        log("  跳过 %s（%s）" % (rel, str(e)[:60]))
                        n_skip += 1
    except Exception as e:
        try:
            os.remove(dest)
        except Exception:
            pass
        return False, "打包失败：%s" % str(e)[:150]

    size_mb = os.path.getsize(dest) / 1024.0 / 1024.0
    log("[搬家包] %s" % dest)
    log("         文件 %d 个（跳过 %d），原始 %.1f MB，压缩后 %.1f MB"
        % (n_files, n_skip, total / 1048576.0, size_mb))
    top_n = sorted(per_dir.items(), key=lambda kv: -kv[1])[:12]
    log("         主要目录：" + "、".join("%s %d" % (k, v) for k, v in top_n))
    log("         已排除目录：%s" % "、".join(sorted(skip_dirs - {"work"})))
    if not include_secrets:
        log("         注意：未包含账号配置（_cos_config.json / 迁移配置.json / bi_sources.json），"
            "新电脑需自行填写 BI 工号。")
    if not include_data:
        log("         注意：未包含源数据 Excel，新电脑需把源表放进项目根目录。")
    if not include_history:
        log("         注意：未包含历史数据，新电脑的历史趋势将从当天重新积累。")
    return True, dest


if __name__ == "__main__":
    # 命令行入口：python _portable.py scan|fix|bundle|task
    argv = sys.argv[1:]
    act = argv[0] if argv else "scan"
    if act == "scan":
        for r in scan():
            mark = {"ok": "√", "warn": "!", "bad": "×"}[r["level"]]
            print("%s %-22s %s" % (mark, r["label"], r["detail"]))
    elif act == "fix":
        ok, msgs = fix_all()
        print("\n".join(msgs))
        print("结果：", "成功" if ok else "部分未完成")
    elif act == "bundle":
        # 默认 = 全套（给自己换机用）；--no-secrets 给出别人之前务必带上
        full = "--full" in argv or not any(a.startswith("--no-") for a in argv)
        ok, res = make_bundle(
            include_data="--no-data" not in argv,
            include_history=("--no-history" not in argv) and full,
            include_wheels="--no-wheels" not in argv,
            include_secrets="--no-secrets" not in argv)
        print(res)
    elif act == "publish":
        import json as _j
        print(_j.dumps(publish_info(), ensure_ascii=False, indent=2))
        print("主看板目录：", site_dir())
        print("导航页目录：", nav_dir())
        print("线上地址：", cloud_url() or "（未配置）")
    elif act == "migrate":
        import _migrate
        ok, msgs = _migrate.apply_all()
        print("\n".join(msgs))
    elif act == "task":
        ok, msg = register_task()
        print(msg)
    else:
        print("用法：python _portable.py [scan|fix|bundle|publish|migrate|task]")
        print("      bundle 可加 --no-data / --no-history / --no-wheels / --no-secrets")
