# -*- coding: utf-8 -*-
"""
cos_sync.py —— 本机 Excel 源数据变动 → 重建看板 → 上传腾讯云 COS（静态网站）

设计要点
  1. 判变只看「源 Excel mtime vs 本地 JSON meta.file_mtime_ts」，绝不用 build_hash
     （build 时注入的 ?cb=<epoch> 会让指纹每次必变）。
  2. 上传是增量的：先 head_object 取 ETag 与本地 md5 比对，相同则跳过。
     首次运行会全量上传。
  3. 看板产物逐个设 public-read，桶本身保持私有 —— 已上传的源 Excel 不会外泄。
  4. 前端零改动：_deploy 内所有 fetch 都是相对路径，搬到 COS 静态网站后同源工作。
     页面默认 30 秒自动刷新，所以 JSON 一换，网页自动变，无需再 deploy。

用法
  python cos_sync.py            # 检测 → 必要时 rebuild → 增量上传
  python cos_sync.py --check    # 只检测，显示「本地/COS 差异」，不上传
  python cos_sync.py --force    # 强制 rebuild + 全量重传（忽略 md5 相同）
  python cos_sync.py --init     # 开静态网站托管 + 全量上传（首次用）
"""
import os
import io
import re
import sys
import json
import time
import mimetypes
import subprocess
import hashlib

BASE = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(BASE, "_deploy")
CFG_PATH = os.path.join(BASE, "_cos_config.json")
STATE_PATH = os.path.join(BASE, "_cos_state.json")
LOG_PATH = os.path.join(BASE, "_cos_sync.log")
PY = sys.executable

# 三个主源（与 _cmp_sync.py 保持一致）
SRC = {
    "act":   os.path.join(BASE, "分公司活动量-在线看板.xlsx"),
    "sanli": os.path.join(BASE, "分公司-员工三力明细.xlsx"),
    "ry":    os.path.join(BASE, "分公司-人员数据分析路径.xlsx"),
}
# 辅助表：只影响 ry 页
AUX = os.path.join(BASE, "..", "分公司健康表", "四大模式-重点指标表格.xlsx")
FILES = {"act": "data.json", "sanli": "sanli_data.json", "ry": "ry_data.json"}

# 上传策略：按扩展名给 Content-Type / Cache-Control
CT_MAP = {
    ".html": "text/html; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".js":   "application/javascript; charset=utf-8",
    ".css":  "text/css; charset=utf-8",
    ".png":  "image/png",
    ".svg":  "image/svg+xml",
}
# no-cache：浏览器每次回源校验 ETag，未变返回 304（省流量），变了立即取新值
CC_HOT = "no-cache, must-revalidate"
# 静态资源可长缓存
CC_COLD = "public, max-age=604800"

_log_lines = []


def log(msg):
    line = "[%s] %s" % (time.strftime("%H:%M:%S"), msg)
    print(line)
    _log_lines.append(line)


# --------------------------------------------------------------------------
# 配置
# --------------------------------------------------------------------------
def load_cfg():
    if not os.path.exists(CFG_PATH):
        sys.exit("缺少配置文件：%s\n请先填入 secret_id / secret_key / region / bucket。" % CFG_PATH)
    cfg = json.load(open(CFG_PATH, encoding="utf-8"))
    for k in ("secret_id", "secret_key", "region", "bucket"):
        if not cfg.get(k):
            sys.exit("配置缺少字段：%s（文件：%s）" % (k, CFG_PATH))
    cfg.setdefault("prefix", "dashboard/")
    if cfg["prefix"] and not cfg["prefix"].endswith("/"):
        cfg["prefix"] += "/"
    return cfg


def make_client(cfg):
    from qcloud_cos import CosConfig, CosS3Client
    conf = CosConfig(
        Region=cfg["region"],
        SecretId=cfg["secret_id"],
        SecretKey=cfg["secret_key"],
        Scheme="https",
    )
    return CosS3Client(conf)


def md5_of(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# --------------------------------------------------------------------------
# 判变：源 Excel mtime vs 本地 JSON meta
# --------------------------------------------------------------------------
def _meta_of(json_name):
    p = os.path.join(DEPLOY, json_name)
    if not os.path.exists(p):
        return None
    try:
        return json.load(open(p, encoding="utf-8")).get("meta") or {}
    except Exception:
        return None


def _aux_val(meta):
    """aux_mtime_ts 可能是 dict {表名: mtime}，取含「四大模式」的那个。"""
    v = (meta or {}).get("aux_mtime_ts")
    if v is None:
        return None
    if isinstance(v, dict):
        for k, x in v.items():
            if "四大模式" in k:
                return x
        return next(iter(v.values())) if v else None
    return v


def check_changed():
    """返回 (need, reasons)。need=True 表示源数据已变，需要 rebuild。"""
    reasons = []
    for k, p in SRC.items():
        if not os.path.exists(p):
            reasons.append("%s 源文件缺失：%s" % (k, os.path.basename(p)))
            continue
        s = os.path.getmtime(p)
        meta = _meta_of(FILES[k])
        l = (meta or {}).get("file_mtime_ts")
        if l is None:
            reasons.append("%s 本地快照无 meta，需重建" % k)
        elif abs(s - l) > 1e-6:
            reasons.append("%s 源表已更新（源 %s ≠ 快照 %s）" % (
                k,
                time.strftime("%m-%d %H:%M:%S", time.localtime(s)),
                time.strftime("%m-%d %H:%M:%S", time.localtime(l)),
            ))
    # 辅助表（只影响 ry）
    if os.path.exists(AUX):
        a = os.path.getmtime(AUX)
        al = _aux_val(_meta_of("ry_data.json"))
        if al is None:
            reasons.append("aux 本地快照无 mtime，需重建")
        elif abs(a - al) > 1e-6:
            reasons.append("aux 四大模式表已更新（源 %s ≠ 快照 %s）" % (
                time.strftime("%m-%d %H:%M:%S", time.localtime(a)),
                time.strftime("%m-%d %H:%M:%S", time.localtime(al)),
            ))
    return (len(reasons) > 0), reasons


# --------------------------------------------------------------------------
# 构建
# --------------------------------------------------------------------------
def rebuild():
    log("运行 build_snapshot.py ...")
    r = subprocess.run([PY, "build_snapshot.py"], cwd=BASE,
                       capture_output=True, text=True, encoding="utf-8", errors="ignore")
    if r.returncode != 0:
        log("构建失败 exit=%s" % r.returncode)
        log((r.stdout or "")[-1500:])
        log((r.stderr or "")[-1500:])
        return False
    log("构建完成")
    return True


# --------------------------------------------------------------------------
# 上传
# --------------------------------------------------------------------------
def collect_files():
    """遍历 _deploy，产出 [(local_path, key)]。"""
    out = []
    for root, dirs, files in os.walk(DEPLOY):
        dirs[:] = [d for d in dirs if not d.startswith("__")]
        for f in files:
            if f.startswith("~$"):
                continue
            lp = os.path.join(root, f)
            rel = os.path.relpath(lp, DEPLOY).replace("\\", "/")
            out.append((lp, rel))
    return out


def deploy_tcb(cfg):
    """上传 _deploy 到腾讯云 CloudBase 静态托管（主链路，网页实际访问地址）。

    与 COS 的关系：CloudBase 是「网页」的落点（默认域名可直接打开、永不休眠），
    COS 是异地备份（dashboard/ 前缀）。两者互不影响。
    CloudBase 的 hosting deploy 是整目录覆盖，无增量；55 文件约 9 秒，可接受。
    """
    t = cfg.get("tcb") or {}
    env, node, cli = t.get("env_id"), t.get("node"), t.get("cli")
    if not (env and node and cli):
        log("CloudBase 未配置 → 跳过")
        return False
    if not os.path.exists(cli):
        log("CloudBase CLI 不存在：%s → 跳过" % cli)
        return False
    if not os.path.exists(node):
        log("node 不存在：%s → 跳过" % node)
        return False
    cwd = os.path.dirname(os.path.dirname(os.path.dirname(cli)))  # node/workspace
    cmd = [node, cli, "hosting", "deploy", DEPLOY, "", "-e", env]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="ignore",
                           cwd=cwd if os.path.isdir(cwd) else None,
                           timeout=1200)
        out = (r.stdout or "") + (r.stderr or "")
        m = re.search(r"Successfully uploaded (\d+) file", out)
        ok = bool(m) or ("File upload completed" in out)
        log("CloudBase：%s%s" % ("部署成功" if ok else "结果存疑",
                                 ("，%s 个文件" % m.group(1)) if m else ""))
        if not ok:
            log("  tcb 输出尾部：" + out[-500:].replace("\n", " | "))
        return ok
    except Exception as e:
        log("CloudBase 部署失败：%s" % str(e)[:200])
        return False


def upload_all(client, cfg, force=False):
    bucket = cfg["bucket"]
    prefix = cfg["prefix"]
    files = collect_files()
    up = skip = fail = 0
    touched = []
    for lp, rel in files:
        key = prefix + rel
        ext = os.path.splitext(lp)[1].lower()
        ct = CT_MAP.get(ext, mimetypes.guess_type(lp)[0] or "application/octet-stream")
        cc = CC_COLD if rel.startswith("static/") else CC_HOT
        try:
            if not force:
                try:
                    head = client.head_object(Bucket=bucket, Key=key)
                    etag = (head.get("ETag") or "").strip('"').lower()
                    if etag == md5_of(lp):
                        skip += 1
                        continue
                except Exception:
                    pass  # 不存在 → 上传
            with open(lp, "rb") as fh:
                body = fh.read()
            client.put_object(
                Bucket=bucket,
                Body=body,
                Key=key,
                ContentType=ct,
                CacheControl=cc,
                ACL="public-read",
            )
            up += 1
            touched.append(rel)
        except Exception as e:
            fail += 1
            log("  上传失败 %s : %s" % (rel, str(e)[:160]))
    log("上传完毕：新增/更新 %d，跳过(未变) %d，失败 %d" % (up, skip, fail))
    return up, skip, fail, touched


def ensure_website(client, cfg):
    """开启静态网站托管（幂等）。"""
    bucket = cfg["bucket"]
    try:
        client.put_bucket_website(
            Bucket=bucket,
            WebsiteConfiguration={
                "IndexDocument": {"Suffix": "index.html"},
                "ErrorDocument": {"Key": "index.html"},
            },
        )
        log("静态网站托管已开启（IndexDocument=index.html）")
        return True
    except Exception as e:
        log("静态网站托管设置失败：%s" % str(e)[:200])
        return False


def website_url(cfg):
    return "https://%s.cos-website.%storage.example.com/%s" % (
        cfg["bucket"], cfg["region"], cfg["prefix"].lstrip("/")
    )


def fingerprint():
    p = os.path.join(DEPLOY, "_build_fingerprint.json")
    if os.path.exists(p):
        try:
            return json.load(open(p, encoding="utf-8")).get("build_hash")
        except Exception:
            return None
    return None


# --------------------------------------------------------------------------
def main():
    args = set(sys.argv[1:])
    cfg = load_cfg()
    client = make_client(cfg)

    log("=== COS 同步 %s ===" % time.strftime("%Y-%m-%d %H:%M:%S"))
    log("桶：%s（%s）　前缀：%s" % (cfg["bucket"], cfg["region"], cfg["prefix"]))

    if "--init" in args:
        ensure_website(client, cfg)

    need, reasons = check_changed()
    if "--check" in args:
        print("\n--- 变更检测 ---")
        for r in reasons:
            print("  * " + r)
        print("NEED_REBUILD =", need)
        print("访问链接：", website_url(cfg) + "index.html")
        return 0 if not need else 23

    if need or "--force" in args or "--init" in args:
        for r in reasons:
            log("  * " + r)
        if "--force" in args and not need:
            log("  * --force：强制重建")
        if not rebuild():
            return 1
    else:
        log("源数据无变化 → 跳过重建")

    # 主链路：CloudBase 静态托管（网页实际访问地址，默认域名可直接打开、永不休眠）
    tcb_ok = deploy_tcb(cfg)
    # 备链路：COS（异地备份；失败不影响主链路）
    up, skip, fail, touched = upload_all(client, cfg, force=("--force" in args))

    fp = fingerprint()
    tcb_url = (cfg.get("tcb") or {}).get("url") or (website_url(cfg) + "index.html")
    state = {
        "synced_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "build_hash": fp,
        "uploaded": up, "skipped": skip, "failed": fail,
        "rebuilt": bool(need or "--force" in args or "--init" in args),
        "url": tcb_url,
        "tcb_ok": tcb_ok,
        "cos_url": website_url(cfg) + "index.html",
        "touched_sample": touched[:8],
    }
    json.dump(state, open(STATE_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    log("build_hash: %s" % fp)
    log("主链接（CloudBase）：%s" % tcb_url)
    log("备份（COS）：%s" % state["cos_url"])
    try:
        with io.open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("\n".join(_log_lines) + "\n")
    except Exception:
        pass
    return 0 if fail == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
