# -*- coding: utf-8 -*-
"""
企微通知抓取 → 导航页「重要通知」/「集团简报TV」分类
=========================================================
数据源（本地解密库，需企微 PC 在线，先 refresh）：
  · 通知群：分公司通知群——重要信息群，请置顶   → cat="notice"
  · 简报群：集团分公司（集团简报TV）   → cat="brief"（仍由 daily_brief.py 写入）

用法：
  python notice_sync.py                    # 抓今天，写入 notices.js
  python notice_sync.py --date 2026-10-19  # 抓指定日期
  python notice_sync.py --from 2026-10-19 --to 2026-10-20
  python notice_sync.py --dry              # 只看会抓什么，不写文件
  python notice_sync.py --normalize        # 只给旧条目补 cat/time 并重排，不抓新消息

写入规则：按「日期 + 标题前 24 字」去重；新条目插到数组顶部；整体按日期时间倒序。
"""
import os
import re
import io
import sys
import json
import time
import sqlite3
import datetime
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NOTICES_JS = os.path.join(BASE_DIR, "deploy_nav", "notices.js")
NOTICE_IMG_DIR = os.path.join(BASE_DIR, "deploy_nav", "notices")
NODE = r"<HOME>\.workbuddy\binaries\node\versions\22.22.2-6\node.exe"

# 企微图片缓存（按文件名存放，非 md5）
IMG_CACHE = r"<DATA_CACHE>\1688857640763810\Cache\Image"

# 复用 daily_brief 的底层能力
import daily_brief as db  # noqa: E402

GROUP_NOTICE = "分公司通知群——重要信息群，请置顶"
CONV_NOTICE = "R:135832800621872"

CJK = re.compile(r"[\u4e00-\u9fff]")
URL_RE = re.compile(r"https?://[^\s\"'<>）)]+")

# 人工精修标题（key = "日期 时间"）；未命中的用正文首句自动生成
TITLE_OVERRIDE = {
    "2026-10-19 08:35": "语音质检违规数据核对（10.9–10.14）",
    "2026-10-19 08:56": "10 月见客数据线下申诉核对（16:00 前提交）",
    "2026-10-19 13:57": "关于 10 月接单核对的时间节点通知",
    "2026-10-19 14:04": "2026 年 11 月全民营销激励",
    "2026-10-19 14:17": "办公室防鼠工作提醒",
    "2026-10-19 18:05": "全民营销资源录入操作手册",
    "2026-10-20 14:24": "企大顺发版通知（10/20、10/22）与发版时间提醒",
}

# 人工整理正文（key = "日期 时间"）：protobuf 拆分导致段落顺序错乱时用它兜底
CONTENT_OVERRIDE = {
    "2026-10-19 14:17": [
        "各位同事，近期办公室发现老鼠活动，请大家配合做好防鼠工作：",
        "1、零食、水果等食物不要留在桌面 / 抽屉过夜，外卖垃圾下班扔至楼梯处大垃圾桶内，及时清理食物碎屑；\n2、最后下班的同事记得关好货梯那边的消防门，防止老鼠进入；\n3、发现鼠迹请联系人事行政部，不要徒手捕捉。",
    ],
}

HEADER = """/* ============================================================
   重要通知 / 集团简报TV —— 维护入口
   ------------------------------------------------------------
   · 只改这个文件，保存后刷新页面即生效（不用动 index.html）
   · 图片放到 deploy_nav/notices/ 目录，写 "notices/文件名"。
     **配图可选**：不写 image 就是纯文字通知，列表里显示 🔔 占位。
   · 每条由 5 部分组成：

       date     日期   "2026-10-20"
       time     时间   "08:35"（可选，同日多条按它排序）
       cat      分类   "notice" = 重要通知 / "brief" = 集团简报TV
       title    标题
       image    图片路径（可选）"notices/1.jpg"
       content  正文，推荐写成数组：一个元素 = 一段

   · 不写 cat 时：标题含「集团简报TV」→ 自动归入简报TV，其余归重要通知。
   · 自动抓取：python notice_sync.py --date 2026-10-19（详见脚本头部）
   ============================================================ */

"""


# ---------------------------------------------------------------- 读取 / 写入
def load_notices():
    """用 node 执行 notices.js，导出 window.__NOTICES__ 的 JSON。"""
    code = ("global.window={};"
            "try{require(%s);}catch(e){process.stderr.write(String(e));process.exit(1);}"
            "process.stdout.write(JSON.stringify(window.__NOTICES__||[]));" % json.dumps(NOTICES_JS))
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    p = subprocess.run([NODE, "-e", code], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=env)
    if p.returncode != 0 or not p.stdout.strip():
        raise SystemExit("读取 notices.js 失败：%s" % (p.stderr or "")[:300])
    return json.loads(p.stdout)


def is_brief(n):
    return n.get("cat") == "brief" or "简报TV" in (n.get("title") or "")


def sort_key(n):
    return (str(n.get("date") or ""), str(n.get("time") or "00:00"))


def save_notices(items, dry=False):
    items = sorted(items, key=sort_key, reverse=True)
    body = "window.__NOTICES__ = " + json.dumps(items, ensure_ascii=False, indent=2) + ";\n"
    if dry:
        return HEADER + body
    with io.open(NOTICES_JS, "w", encoding="utf-8", newline="") as f:
        f.write(HEADER + body)
    return None


# ---------------------------------------------------------------- 抓取
def _strip_wrapper(s):
    """企微 ct=2 的正文外面套了两层 protobuf：<08 00> <12 len> <0a len> 正文。逐层剥掉。"""
    for _ in range(4):
        b = s.encode("utf-8")
        if not b.startswith(b"\x08\x00\x12"):
            break
        i = 3
        ln, i = db._varint(b, i)
        if ln is None or i >= len(b) or b[i] != 0x0A:
            break
        i += 1
        ln2, i = db._varint(b, i)
        if ln2 is None:
            break
        try:
            s = b[i:].decode("utf-8")
        except UnicodeDecodeError:
            break
    return s


def _extract(raw):
    """从 protobuf 里把可读文本段全部捞出来（保真，不丢符号）。"""
    if raw is None:
        return []
    try:
        s = raw.decode("utf-8")
        if db._score(s) > 0.94 and CJK.search(s):
            return [_strip_wrapper(s).strip()]
    except UnicodeDecodeError:
        pass
    outs = []
    try:
        db._walk(raw, outs)
    except Exception:
        return []
    res, seen = [], set()
    for s in outs:
        s = _strip_wrapper(s.strip("\x00").strip()).strip()
        if len(s) < 4 or db._score(s) < 0.88:
            continue
        if not CJK.search(s) and not URL_RE.match(s):
            continue
        # 控制字符残留 / 纯文件名 / 纯 id 之类
        if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", s):
            continue
        if re.fullmatch(r"[\w\-.]+\.(png|jpg|jpeg|xlsx|docx|pdf|pptx|eml)", s.strip(), re.I):
            continue
        if re.match(r"^企业微信截图[_0-9A-Za-z.\-]*$", s.strip()):
            continue
        if re.search(r"please view it in WeCom|^Email from\b", s.strip(), re.I):
            continue
        if re.fullmatch(r"[\w.\-]+@[\w.\-]+", s.strip()):
            continue
        if re.search(r"@dgg\.net", s, re.I) and len(s) < 60:
            continue
        if s in seen:
            continue
        seen.add(s)
        res.append(s)
    return res


def _clean(t):
    t = t.replace("\r\n", "\n").replace("\u2005", " ").replace("\xa0", " ")
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def _auto_title(text):
    first = _clean(text).split("\n")[0]
    first = re.sub(r"^(Dear\s*All|各位同事|各位|大家好|@All)[：:，,\s]*", "", first, flags=re.I)
    first = first.strip(" ，,。.！!：:")
    return (first[:38] + "…") if len(first) > 38 else (first or "群通知")


def _file_table_rows(day):
    """查解密库 file.db：该群当天的图片记录（extension_type=4 → 图片）。"""
    p = os.path.join(db.DEC, "file.db")
    if not os.path.exists(p):
        return []
    t0 = int(time.mktime(datetime.datetime.combine(day, datetime.time(0, 0)).timetuple()))
    t1 = t0 + 86400
    c = db._connect(p)
    try:
        rows = c.execute(
            "select receive_time, name, md5, size from file_table4 "
            "where conversation_id=? and receive_time>=? and receive_time<? "
            "and extension_type=4 order by receive_time",
            (CONV_NOTICE, t0, t1)).fetchall()
    finally:
        c.close()
    return rows


def _find_cache_image(name, ts):
    """在企微图片缓存 Cache/Image/<YYYY-MM>/ 里按文件名找图（含 (1)(2) 副本），取体积最大的。"""
    if not name:
        return None
    stem = os.path.splitext(name)[0]
    dt = datetime.datetime.fromtimestamp(ts)
    best = None
    for k in range(0, 3):                       # 当月 + 往前两个月（跨月转发的截图）
        ym = (dt - datetime.timedelta(days=31 * k)).strftime("%Y-%m")
        d = os.path.join(IMG_CACHE, ym)
        if not os.path.isdir(d):
            continue
        try:
            files = os.listdir(d)
        except OSError:
            continue
        for f in files:
            fs = os.path.splitext(f)[0]
            if fs == stem or fs.startswith(stem + "("):
                p = os.path.join(d, f)
                try:
                    sz = os.path.getsize(p)
                except OSError:
                    continue
                if not best or sz > best[1]:
                    best = (p, sz)
    return best[0] if best else None


_IMG_MD5_INDEX = None


def _md5_file(p, buf=1 << 20):
    import hashlib
    h = hashlib.md5()
    with open(p, "rb") as f:
        while True:
            b = f.read(buf)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _same_image_exists(src):
    """已有同内容图片（手工加过的接单核对图等）直接复用，避免重复占空间。"""
    global _IMG_MD5_INDEX
    if not os.path.isdir(NOTICE_IMG_DIR):
        return None
    if _IMG_MD5_INDEX is None:
        _IMG_MD5_INDEX = {}
        for f in os.listdir(NOTICE_IMG_DIR):
            if f.lower().endswith((".png", ".jpg", ".jpeg")):
                p = os.path.join(NOTICE_IMG_DIR, f)
                try:
                    _IMG_MD5_INDEX.setdefault(_md5_file(p), p)
                except OSError:
                    pass
    return _IMG_MD5_INDEX.get(_md5_file(src))


def _save_image(src, date, hhmm, idx):
    """复制/压缩到 deploy_nav/notices/，返回站内相对路径。"""
    from PIL import Image
    os.makedirs(NOTICE_IMG_DIR, exist_ok=True)
    base = "%s-%s-%d" % (date, hhmm.replace(":", ""), idx)
    ext = os.path.splitext(src)[1].lower()
    size = os.path.getsize(src)
    try:
        im = Image.open(src)
        im.load()
        w, h = im.size
        need = size > 520 * 1024 or w > 1200
        if not need:
            same = _same_image_exists(src)
            if same:
                return "notices/" + os.path.basename(same)
        if need:
            if w > 1200:
                im = im.resize((1200, max(1, round(h * 1200.0 / w))), Image.LANCZOS)
            dst = os.path.join(NOTICE_IMG_DIR, base + ".jpg")
            im.convert("RGB").save(dst, "JPEG", quality=86, optimize=True)
        else:
            dst = os.path.join(NOTICE_IMG_DIR, base + (".jpg" if ext in (".jpg", ".jpeg") else ".png"))
            with open(src, "rb") as fi, open(dst, "wb") as fo:
                fo.write(fi.read())
    except Exception:
        dst = os.path.join(NOTICE_IMG_DIR, base + (ext if ext in (".jpg", ".jpeg", ".png") else ".png"))
        with open(src, "rb") as fi, open(dst, "wb") as fo:
            fo.write(fi.read())
    return "notices/" + os.path.basename(dst)


def _group_images(day, groups):
    """把当天的图片按时间挂到最近的 5 分钟窗口上。"""
    out = {}
    files = _file_table_rows(day)
    for ts, name, md5, size in files:
        if not groups:
            break
        gi = min(range(len(groups)),
                 key=lambda i: 0 if groups[i]["ts"] - 60 <= ts <= groups[i]["last"] + 300
                 else min(abs(ts - groups[i]["ts"]), abs(ts - groups[i]["last"])))
        g = groups[gi]
        if min(abs(ts - g["ts"]), abs(ts - g["last"])) > 900:
            continue
        if "imgs" not in g:
            g["imgs"] = []
        src = _find_cache_image(name, ts)
        if src:
            g["imgs"].append(src)
        else:
            print("   [warn] 缓存里没找到图片：%s" % name)
    return out


def fetch_day(day):
    """抓通知群某一天的消息，按 5 分钟窗口聚成若干条通知。"""
    conv = CONV_NOTICE
    t0 = int(time.mktime(datetime.datetime.combine(day, datetime.time(0, 0)).timetuple()))
    t1 = t0 + 86400
    p = os.path.join(db.DEC, "message.db")
    if not os.path.exists(p):
        raise SystemExit("找不到解密库 message.db，请先跑 daily_brief.py refresh 或 read.py refresh")

    c = db._connect(p)
    try:
        rows = c.execute(
            "select send_time, content_type, sender_id, content from message_table "
            "where conversation_id=? and send_time>=? and send_time<? "
            "and content_type in (0,2,13,14,20,123) order by send_time",
            (conv, t0, t1)).fetchall()
    finally:
        c.close()

    # 5 分钟聚类（同一批图文/附件算一条通知）
    groups = []
    for st, ct, sd, content in rows:
        txts = _extract(content)
        if not txts:
            continue
        if groups and st - groups[-1]["last"] <= 300:
            g = groups[-1]
        else:
            g = {"ts": st, "last": st, "parts": []}
            groups.append(g)
        g["last"] = st
        for t in txts:
            if t not in g["parts"]:
                g["parts"].append(t)

    _group_images(day, groups)

    items = []
    for g in groups:
        parts = [q for q in g["parts"] if len(_clean(q)) >= 6]
        if not parts:
            continue
        urls = [u for u in parts if u.startswith("http")]
        parts = [p for p in parts if not p.startswith("http")]
        if not parts:
            continue
        # 最长的一段当正文主体，其余（如附件名）附在后面
        parts.sort(key=lambda s: len(s), reverse=True)
        body = []
        for q in parts:
            q = _clean(q)
            if q and q not in body:
                body.append(q)
        # 正文里出现多条编号时拆段，便于详情页悬挂缩进
        dt = datetime.datetime.fromtimestamp(g["ts"])
        key = dt.strftime("%Y-%m-%d %H:%M")
        override = CONTENT_OVERRIDE.get(key)
        if override:
            body = override
        title = TITLE_OVERRIDE.get(key) or _auto_title(body[0])
        item = {
            "date": dt.strftime("%Y-%m-%d"),
            "time": dt.strftime("%H:%M"),
            "cat": "notice",
            "title": title,
            "content": body,
        }
        if urls:
            # 邮件卡片里的 oauth 授权跳转没有阅读价值，丢弃
            usable = [u for u in urls if "open.weixin.qq.com/connect/oauth2" not in u]
            if usable:
                item["link"] = usable[0]
        if g.get("imgs"):
            paths = []
            for i, src in enumerate(g["imgs"][:9]):
                try:
                    paths.append(_save_image(src, item["date"], item["time"], i + 1))
                except Exception as e:
                    print("   [warn] 图片保存失败 %s：%s" % (src, e))
            if paths:
                item["images"] = paths
        items.append(item)
    return items


# ---------------------------------------------------------------- 合并
def norm_title(t):
    """去空格与常见标点后再比对，避免「10 月」与「10月」被判成两条。"""
    t = re.sub(r"[\s\-—–_、,，。.：:；;（）()【】\[\]「」\"“”]", "", str(t or ""))
    return t[:24]


def merge(old, new, dry=False):
    old = [dict(x) for x in old]
    for n in old:                      # 旧条目补分类
        if "cat" not in n:
            n["cat"] = "brief" if is_brief(n) else "notice"
    old_map = {(n.get("date"), norm_title(n.get("title"))): n for n in old}
    added, updated = [], []
    for it in new:
        k = (it["date"], norm_title(it["title"]))
        if k in old_map:               # 已存在 → 补截图（/补正文）
            o = old_map[k]
            touched = False
            if it.get("images") and o.get("images") != it["images"]:
                o["images"] = it["images"]
                touched = True
            if it.get("link") and not o.get("link"):
                o["link"] = it["link"]
                touched = True
            if touched:
                updated.append(o)
            continue
        old.append(it)
        added.append(it)
    save_notices(old, dry=dry)
    return added, updated


def main():
    args = sys.argv[1:]
    dry = "--dry" in args
    if "--normalize" in args:
        items = load_notices()
        save_notices(items, dry=dry)
        print("已规范 %d 条（补 cat / 排序）" % len(items))
        return 0

    days = []
    if "--date" in args:
        i = args.index("--date")
        days = [datetime.datetime.strptime(args[i + 1], "%Y-%m-%d").date()]
    elif "--from" in args and "--to" in args:
        a = datetime.datetime.strptime(args[args.index("--from") + 1], "%Y-%m-%d").date()
        b = datetime.datetime.strptime(args[args.index("--to") + 1], "%Y-%m-%d").date()
        d = a
        while d <= b:
            days.append(d)
            d += datetime.timedelta(days=1)
    else:
        days = [datetime.date.today()]

    if "--no-refresh" not in args:
        ok, msg = db._refresh_db()
        print("刷新解密库：", msg)
        if not ok:
            return 2

    new = []
    for d in days:
        got = fetch_day(d)
        print("%s 抓到 %d 条：" % (d, len(got)))
        for g in got:
            print("   %s  %s" % (g["time"], g["title"]))
        new.extend(got)

    added, updated = merge(load_notices(), new, dry=dry)
    print("\n新增 %d 条，补截图 %d 条（%s）" % (
        len(added), len(updated), "dry-run 未写入" if dry else "已写入 notices.js"))
    for a in added:
        print("   + %s %s %s%s" % (a["date"], a.get("time", ""), a["title"],
                                   "  [图 %d]" % len(a["images"]) if a.get("images") else ""))
    for u in updated:
        print("   ~ %s %s %s  [图 %d]" % (u["date"], u.get("time", ""), u["title"], len(u.get("images") or [])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
