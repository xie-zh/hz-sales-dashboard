# -*- coding: utf-8 -*-
"""迁移与发布配置层 V1.0 —— 换电脑时唯一需要改的地方

为什么要单独一层：
    原来「发布到哪 / BI 工号是谁 / 源表在哪」分散在 _cos_config.json、
    bi_sources.json、以及几个脚本里写死的绝对路径中。换一台电脑就要改四处，
    漏一处就静默跑错。本模块把这三类东西收拢成一份可读的配置
    （项目根目录下的「迁移配置.json」），并负责把它落地到各脚本真正认的位置。

职责：
    1. load/save      读写「迁移配置.json」
    2. publish        给 _portable / bi_refresh / GUI 用的发布参数
    3. apply          把配置应用下去：
                        · bi_sources.json 的 files.* 绝对路径 → 指到新电脑的项目根
                        · bi_login_no（BI 取数工号）
                        · _cos_config.json 的 publish 块（并清掉腾讯云密钥）
    4. selfcheck      迁移后的体检：路径是否都落在新电脑上、发布参数是否齐全

被 setup_new_pc.py / _portable.py 共用。不依赖 tkinter，可在任何解释器下跑。
"""

import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.abspath(__file__))        # …\\8.11\\Hzexamplebot

CFG_NAME = "迁移配置.json"

# 源 Excel 在项目根里的固定文件名（迁移时按这个重建绝对路径）
DATA_FILES = {
    "活动量": "某公司活动量-在线看板.xlsx",
    "三力明细": "某公司-员工三力明细.xlsx",
    "人员路径": "某公司-人员数据分析路径.xlsx",
    "四大模式": os.path.join("分公司健康表", "四大模式-重点指标表格.xlsx"),
}

DEFAULT_CFG = {
    "_说明": [
        "某公司看板 · 迁移与发布配置。换电脑 / 换发布账号时，只改这一个文件。",
        "改完保存，双击项目根目录的「一键部署_新电脑.bat」即可全部生效。",
        "dash_url / nav_url 留空表示「还没发布过」；发布拿到链接后回填到这里，",
        "线上核对、控制台的「打开看板」都会跟着走。",
    ],
    "publish": {
        "channel": "workbuddy",
        "mirror_tencent": False,
        "site_dir": "_deploy",
        "nav_dir": os.path.join("deploy_tool", "build"),
        "entry_html": "index.html",
        "app_name": "某公司看板",
        "domain_prefix": "hz-branch-dashboard",
        "nav_app_name": "看板导航页",
        "nav_domain_prefix": "hz-branch-nav",
        "dash_url": "",
        "nav_url": "",
        "nav_report_url": "",
        "nav_title_url": "",
        "visit_api": "",
    },
    "bi_login_no": "000000",
    "task": {
        "refresh_name": "HZBoardRefresh",
        "refresh_minutes": 5,
        "daily_save_name": "HZDashBoardDailySave",
        "daily_save_time": "23:59",
    },
    "backup_dir": "",
}


# ============================ 基础读写 ============================

def root_of(root=None):
    """项目根目录。优先 _portable 的解析（认 _paths.json / 特征文件），退到父目录。"""
    if root:
        return root
    try:
        import _portable
        return _portable.resolve_root()
    except Exception:
        return os.path.dirname(BASE)


def cfg_path(root=None):
    return os.path.join(root_of(root), CFG_NAME)


def _deep_merge(base, over):
    out = dict(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load(root=None):
    """读配置；文件不存在时返回默认值（不写盘）。"""
    p = cfg_path(root)
    try:
        with open(p, encoding="utf-8") as f:
            return _deep_merge(DEFAULT_CFG, json.load(f) or {})
    except Exception:
        return dict(DEFAULT_CFG)


def save(root, data):
    p = cfg_path(root)
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        return False, "写 %s 失败：%s" % (CFG_NAME, str(e)[:100])
    return True, p


def ensure(root=None, log=print):
    """配置不存在就用默认值落一份（新电脑首次运行时调用）。返回 (路径, 是否新建)。"""
    root = root_of(root)
    p = cfg_path(root)
    if os.path.isfile(p):
        return p, False
    save(root, DEFAULT_CFG)
    log("[配置] 已生成 %s（可编辑）" % p)
    return p, True


# ============================ 发布参数 ============================

def publish(root=None):
    """发布参数（合并配置文件与 _cos_config.json，配置文件优先）。"""
    root = root_of(root)
    pub = dict(load(root).get("publish") or {})
    try:
        with open(os.path.join(root, "_cos_config.json"), encoding="utf-8") as f:
            old = (json.load(f) or {}).get("publish") or {}
        for k, v in old.items():
            if v and not pub.get(k):
                pub[k] = v
    except Exception:
        pass
    return pub


def dash_url(root=None):
    """主看板线上地址（以 / 结尾）；未发布过返回空串。"""
    u = (publish(root).get("dash_url") or "").strip()
    return (u.rstrip("/") + "/") if u else ""


def nav_url(root=None):
    u = (publish(root).get("nav_url") or "").strip()
    return (u.rstrip("/") + "/") if u else ""


def site_dir(root=None):
    """待发布的主看板目录（绝对路径）。"""
    return os.path.join(root_of(root), publish(root).get("site_dir") or "_deploy")


def nav_dir(root=None):
    return os.path.join(root_of(root), publish(root).get("nav_dir") or
                        os.path.join("deploy_tool", "build"))


# ============================ 落地到脚本认的位置 ============================

def _find_under(root, basename):
    """在项目根下浅层找同名文件（最多两层深），找不到返回 None。"""
    direct = os.path.join(root, basename)
    if os.path.isfile(direct):
        return direct
    for sub in ("分公司健康表", "data", "源数据"):
        p = os.path.join(root, sub, basename)
        if os.path.isfile(p):
            return p
    try:
        for name in os.listdir(root):
            d = os.path.join(root, name)
            if os.path.isdir(d):
                p = os.path.join(d, basename)
                if os.path.isfile(p):
                    return p
    except Exception:
        pass
    return None


def apply_paths(root=None, log=print, force=False):
    """把 bi_sources.json 里指向**别的电脑**的绝对路径，改写到本机项目根。

    规则（很重要，别改成"一律重写"）：
        · 旧路径在本机**存在** → 原样不动。
          因为有的源表本来就在项目外（例如 分公司健康表\\四大模式-重点指标表格.xlsx
          是另一个看板共用的），贸然改到项目内会让那个看板取不到新数据。
        · 旧路径在本机**不存在** → 说明是从别的电脑拷过来的，按同名文件在本机
          项目根重新定位；找不到就指到 <根>\\<文件名>，并在日志里点名。
        · force=True 时全部改到本机项目根（换机交接后的首次运行用这个）。

    返回 (ok, 说明)。反复运行是安全的。
    """
    root = root_of(root)
    sp = os.path.join(BASE, "bi_sources.json")
    if not os.path.isfile(sp):
        return False, "找不到 bi_sources.json"
    try:
        with open(sp, encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception as e:
        return False, "解析 bi_sources.json 失败：%s" % str(e)[:100]

    files = cfg.get("files") or {}
    changed, kept, missing = [], [], []
    for key, val in list(files.items()):
        old = str(val or "")
        base = os.path.basename(old) or DATA_FILES.get(key) or ""
        if not base:
            continue
        if os.path.isfile(old) and not force:
            kept.append(key)
            continue
        want = os.path.join(root, DATA_FILES[key]) if key in DATA_FILES else \
            (_find_under(root, base) or os.path.join(root, base))
        if os.path.normcase(os.path.normpath(old)) == os.path.normcase(os.path.normpath(want)):
            kept.append(key)
            continue
        files[key] = want
        changed.append("%s → %s" % (key, want))
        if not os.path.isfile(want):
            missing.append(key)

    cfg["files"] = files
    try:
        with open(sp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=1)
    except Exception as e:
        return False, "写 bi_sources.json 失败：%s" % str(e)[:100]

    for c in changed:
        log("[路径] %s" % c)
    if not changed:
        log("[路径] 源表路径本机都有效，无需改动（共 %d 项）" % len(kept))
    if missing:
        log("[路径] ! 这些源表在新路径上不存在，请把文件放到项目根：%s" % "、".join(missing))
    return True, "改写 %d 项，保留 %d 项" % (len(changed), len(kept))


def apply_bi_login(root=None, log=print):
    """把「迁移配置.json」里的 BI 工号写进 bi_sources.json。"""
    root = root_of(root)
    no = str(load(root).get("bi_login_no") or "").strip()
    if not no:
        return False, "配置里没填 BI 工号"
    sp = os.path.join(BASE, "bi_sources.json")
    try:
        with open(sp, encoding="utf-8") as f:
            cfg = json.load(f)
        if str(cfg.get("bi_login_no") or "") == no:
            return True, "BI 工号已是 %s" % no
        cfg["bi_login_no"] = no
        with open(sp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=1)
        log("[账号] BI 取数工号 → %s" % no)
        return True, no
    except Exception as e:
        return False, "写 BI 工号失败：%s" % str(e)[:100]


def apply_cloud_config(root=None, log=print, purge_tencent=False,
                       secret_id=None, secret_key=None):
    """把发布参数写进 _cos_config.json。

    这个文件历史上存腾讯云密钥；迁到 WorkBuddy 后不再需要（账号要收回）。
    但**默认不动**那些字段 —— 老电脑上腾讯云那条链可能还在跑，静默删掉会
    把它打断。要清理请显式调 purge_tencent()（新电脑初始化时默认调）。
    """
    root = root_of(root)
    p = os.path.join(root, "_cos_config.json")
    try:
        with open(p, encoding="utf-8") as f:
            cfg = json.load(f)
        if not isinstance(cfg, dict):
            cfg = {}
    except Exception:
        cfg = {}

    pub = load(root).get("publish") or {}
    cfg["publish"] = _deep_merge(cfg.get("publish") or {}, pub)

    if purge_tencent:
        removed = []
        for k in ("secret_id", "secret_key", "bucket", "prefix"):
            if k in cfg:
                cfg.pop(k, None)
                removed.append(k)
        if "tcb" in cfg:
            cfg.pop("tcb", None)
            removed.append("tcb")
        if removed:
            log("[发布] 已清理腾讯云字段：%s" % "、".join(removed))
    if secret_id:
        cfg["secret_id"] = secret_id
    if secret_key:
        cfg["secret_key"] = secret_key

    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=1)
    except Exception as e:
        return False, "写 _cos_config.json 失败：%s" % str(e)[:100]
    log("[发布] 通道 = %s，主看板链接 = %s"
        % (cfg["publish"].get("channel"), cfg["publish"].get("dash_url") or "（待发布后回填）"))
    return True, "已写入 _cos_config.json"


def purge_tencent(root=None, log=print):
    """彻底移除 _cos_config.json 里的腾讯云密钥/环境/桶（账号收回时调用）。

    返回 (ok, 说明；说明里只报字段名，不回显任何密钥值)。
    """
    root = root_of(root)
    p = os.path.join(root, "_cos_config.json")
    if not os.path.isfile(p):
        return False, "找不到 _cos_config.json"
    try:
        with open(p, encoding="utf-8") as f:
            cfg = json.load(f) or {}
    except Exception as e:
        return False, "解析失败：%s" % str(e)[:80]
    removed = [k for k in ("secret_id", "secret_key", "bucket", "prefix", "tcb")
               if k in cfg]
    for k in removed:
        cfg.pop(k, None)
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=1)
    except Exception as e:
        return False, "写回失败：%s" % str(e)[:100]
    log("[安全] 已移除腾讯云凭据字段：%s" % ("、".join(removed) if removed else "（本来就没有）"))
    return True, "已清理 %d 个字段" % len(removed)


def _blank_machine_publish(pub):
    """把「属于这一台机器/这个发布账号」的字段清掉。

    分享链接是跟着 WorkBuddy 账号走的：换了账号再发一次，链接必然不同，
    把上一台的链接带过去只会让人以为已经通了。所以打包含一律清空，
    等新电脑发布完再回填。
    """
    pub = dict(pub or {})
    pub["channel"] = "workbuddy"
    pub["mirror_tencent"] = False
    for k in ("dash_url", "nav_url", "nav_report_url", "nav_title_url"):
        pub[k] = ""
    return pub


def clean_config_dict(root=None, with_secrets=True):
    """返回一份可安全外发的配置内容（搬家包用）。

    永远不含腾讯云密钥；线上链接一律清空（跟账号走，换机必换）；
    with_secrets=False 时连 BI 工号也留空，由新电脑自行填写。
    """
    root = root_of(root)
    cfg = load(root)
    cfg["publish"] = _blank_machine_publish(cfg.get("publish"))
    if not with_secrets:
        cfg["bi_login_no"] = ""
    return cfg


def clean_cos_dict(root=None, with_secrets=True):
    """返回一份可安全外发的 _cos_config.json 内容（永不含腾讯云密钥）。"""
    root = root_of(root)
    out = {}
    p = os.path.join(root, "_cos_config.json")
    try:
        with open(p, encoding="utf-8") as f:
            old = json.load(f) or {}
        if isinstance(old, dict):
            out = {k: v for k, v in old.items()
                   if k not in ("secret_id", "secret_key", "bucket", "prefix", "tcb")}
    except Exception:
        pass
    pub = load(root).get("publish") or {}
    out["publish"] = _deep_merge({"channel": "workbuddy"},
                                 _blank_machine_publish(pub))
    return out


def apply_all(root=None, log=print, force_paths=False, purge_tencent=False):
    """一次把配置全部落地。返回 (ok, 消息列表)。"""
    root = root_of(root)
    msgs = []
    ensure(root, log=log)
    ok1, m1 = apply_paths(root, log, force=force_paths)
    msgs.append("源表路径：%s" % m1)
    ok2, m2 = apply_bi_login(root, log=log)
    msgs.append("BI 工号：%s" % m2)
    ok3, m3 = apply_cloud_config(root, log=log, purge_tencent=purge_tencent)
    msgs.append("发布配置：%s" % m3)
    return (ok1 and ok3), msgs


# ============================ 体检 ============================

def selfcheck(root=None):
    """迁移体检。返回 rows=[{key,label,ok,level,detail}]。"""
    root = root_of(root)
    rows = []

    def add(key, label, ok, detail, level="ok"):
        rows.append({"key": key, "label": label, "ok": bool(ok),
                     "level": level, "detail": detail})

    p = cfg_path(root)
    add("cfg", "迁移配置", os.path.isfile(p),
        p if os.path.isfile(p) else "缺 %s（双击「一键部署_新电脑.bat」会自动生成）" % CFG_NAME,
        "ok" if os.path.isfile(p) else "bad")

    # 源表路径：不存在=必须处理；存在但在项目外=提醒（换机时要一起打包）
    sp = os.path.join(BASE, "bi_sources.json")
    miss, outside, empty = [], [], []
    inroot, total = 0, 0
    try:
        with open(sp, encoding="utf-8") as f:
            files = (json.load(f) or {}).get("files") or {}
        root_lc = os.path.normcase(os.path.normpath(root))
        for k, v in files.items():
            total += 1
            v = str(v or "")
            if not v:
                empty.append(k)
            elif not os.path.isfile(v):
                miss.append(k)
            else:
                try:
                    if os.path.normcase(os.path.normpath(v)).startswith(root_lc):
                        inroot += 1
                    else:
                        outside.append(k)
                except Exception:
                    outside.append(k)
    except Exception as e:
        empty.append("读取失败:%s" % str(e)[:40])

    det = []
    if empty:
        det.append("未配置：%s" % "、".join(empty))
    if miss:
        det.append("文件不存在（换机后会按同名重新定位，或需自己放进去）：%s" % "、".join(miss))
    if outside:
        det.append("在项目目录之外（换机时记得一起拷）：%s" % "、".join(outside))
    lvl = "bad" if (empty or miss) else ("warn" if outside else "ok")
    add("files", "源表路径", not (empty or miss),
        "；".join(det) if det else "%d/%d 个源表都在本机项目内" % (inroot, total), lvl)

    pub = publish(root)
    add("pub_channel", "发布通道", pub.get("channel") == "workbuddy",
        pub.get("channel") or "未配置",
        "ok" if pub.get("channel") == "workbuddy" else "bad")
    add("pub_site", "主看板产物", os.path.isdir(site_dir(root)),
        site_dir(root) if os.path.isdir(site_dir(root)) else "缺 %s（先跑一次构建）" % site_dir(root),
        "ok" if os.path.isdir(site_dir(root)) else "warn")
    add("pub_nav", "导航页产物", os.path.isdir(nav_dir(root)),
        nav_dir(root) if os.path.isdir(nav_dir(root)) else "缺 %s" % nav_dir(root),
        "ok" if os.path.isdir(nav_dir(root)) else "warn")
    add("pub_url", "线上链接", bool(dash_url(root)),
        dash_url(root) or "尚未发布；发布后把链接回填到「迁移配置.json」的 dash_url",
        "ok" if dash_url(root) else "warn")

    no = str(load(root).get("bi_login_no") or "").strip()
    add("bi", "BI 取数工号", bool(no), no or "未填（迁移配置.json → bi_login_no）",
        "ok" if no else "bad")
    return rows


def format_check(rows):
    mark = {"ok": "√", "warn": "!", "bad": "×"}
    return "\n".join("%s %-14s %s" % (mark[r["level"]], r["label"], r["detail"])
                     for r in rows)


if __name__ == "__main__":
    act = sys.argv[1] if len(sys.argv) > 1 else "check"
    r = root_of()
    if act == "check":
        print("项目根：%s" % r)
        print(format_check(selfcheck(r)))
    elif act == "apply":
        ok, msgs = apply_all(r)
        print("\n".join(msgs))
        print("结果：", "成功" if ok else "部分未完成")
    elif act == "init":
        p, new = ensure(r)
        print(("已新建 " if new else "已存在 ") + p)
    elif act == "purge-tencent":
        ok, msg = purge_tencent(r)
        print(msg)
    elif act == "url":
        # 只打印一行链接 —— 给 .bat 用 for /f 取值（没有就打印占位文案）
        print(dash_url(r) or "（还没发布，链接待回填：迁移配置.json → publish.dash_url）")
    elif act == "publish":
        pub = publish(r)
        print(json.dumps(pub, ensure_ascii=False, indent=2))
        print("主看板目录：", site_dir(r))
        print("导航页目录：", nav_dir(r))
    else:
        print("用法：python _migrate.py [check|apply|init|publish|url|purge-tencent]")
