#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
每日「集团简报TV」自动抓取 → 写入导航页重要通知
==================================================
流程：
  1) 刷新企微本地解密库（需要企微 PC 客户端在线登录）
  2) 在「集团分公司」群里找当天含「集团简报TV」的纯文本消息
  3) 解析出 日期 / 标题 / 正文分段
  4) 按「第X期」去重后插入 deploy_nav/notices.js 顶部
  5) 写 deploy_nav/_brief_sync.json 记录本次结果，供发布环节判断

用法：
  python daily_brief.py                 # 抓今天
  python daily_brief.py --date 2026-10-19
  python daily_brief.py --dry           # 只解析不写入
  python daily_brief.py --no-refresh    # 跳过解密库刷新（已有新鲜快照时省时间）

⚠️ 直接读解密后的 message.db（content_type=0 存的就是干净 UTF-8 原文），
   不走 wecom_reader 的 _parse_content —— 那条正则会吃掉 ℃、- 等字符。
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
SKILL = r"<HOME>\.workbuddy\skills\wecom-reader"
PY = r"<HOME>\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
DB_SRC = r"<DATA_CACHE>\1688857640763810\Data"          # 企微源库（真身，勿改）
DEC = r"<HOME>\WorkBuddy\2026-10-16-16-24-27\wxwork_decrypted"

GROUP_NAME = "集团分公司"
KEYWORD = "集团简报TV"
NOTICES_JS = os.path.join(BASE_DIR, "deploy_nav", "notices.js")
STATE_JSON = os.path.join(BASE_DIR, "deploy_nav", "_brief_sync.json")


# ---------------------------------------------------------------- 底层工具
def _refresh_db():
    """刷新解密库：从 WXWork.exe 内存取密钥重新解密。返回 (ok, msg)。"""
    env = dict(os.environ)
    env["PYTHONPATH"] = SKILL
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [PY, "-m", "wecom_reader.cli", "--db-dir", DB_SRC,
           "--decrypted-dir", DEC, "init"]
    try:
        p = subprocess.run(cmd, cwd=SKILL, env=env, capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=900)
    except Exception as e:
        return False, "刷新异常: %s" % e
    out = (p.stdout or "") + (p.stderr or "")
    m = re.search(r"\{[\s\S]*\}", out)
    if not m:
        return False, "刷新失败（无返回）: %s" % out[:300]
    try:
        d = json.loads(m.group(0))
    except Exception:
        return False, "刷新返回无法解析: %s" % out[:300]
    if not d.get("success"):
        return False, "刷新失败: %s" % d.get("error")
    return True, "解密库已刷新（decrypted=%s copied=%s）" % (d.get("decrypted"), d.get("copied"))


def _connect(path):
    return sqlite3.connect("file:%s?mode=ro" % path.replace("\\", "/"), uri=True)


def _find_group_conv():
    """按群名在 session.db 里查会话 ID。"""
    p = os.path.join(DEC, "session.db")
    if not os.path.exists(p):
        return None
    c = _connect(p)
    try:
        row = c.execute(
            "select id from conversation_table where name=? limit 1", (GROUP_NAME,)
        ).fetchone()
        return row[0] if row else None
    finally:
        c.close()


def _varint(b, i):
    n, s = 0, 0
    while i < len(b):
        c = b[i]
        i += 1
        n |= (c & 0x7F) << s
        if not (c & 0x80):
            return n, i
        s += 7
        if s > 63:
            break
    return None, i


def _walk(buf, out, depth=0):
    """极简 protobuf 遍历：收集所有能严格解成 UTF-8 的 length-delimited 载荷。"""
    i, n = 0, len(buf)
    while i < n:
        key, i = _varint(buf, i)
        if key is None:
            return
        wt = key & 7
        if wt == 2:
            ln, i = _varint(buf, i)
            if ln is None or ln < 0 or i + ln > n:
                return
            payload = buf[i:i + ln]
            i += ln
            try:
                out.append(payload.decode("utf-8"))
            except UnicodeDecodeError:
                if depth < 4 and len(payload) > 4:
                    _walk(payload, out, depth + 1)
        elif wt == 0:
            _, i = _varint(buf, i)
        elif wt == 5:
            i += 4
        elif wt == 1:
            i += 8
        else:
            return


_OK = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef\u2000-\u206f"
                 r"\u2100-\u214f\u2190-\u21ff\u2460-\u24ff\w\s"
                 r".,;:!?'\"()\[\]{}<>《》【】、，。！？；：·—…％%#@&*/\\|+^-]")


def _score(s):
    if not s:
        return 0.0
    good = sum(1 for ch in s if _OK.match(ch))
    return good / float(len(s))


def _text(raw):
    """content 是 protobuf 包装的，解包后取出正文原文（保真，不丢 ℃ - 等字符）。"""
    if raw is None:
        return ""
    if isinstance(raw, str):
        return raw
    try:
        s = raw.decode("utf-8")
        if _score(s) > 0.92:
            return s
    except UnicodeDecodeError:
        pass
    out = []
    _walk(raw, out)
    best, bs = "", 0.0
    for s in out:
        sc = _score(s)
        # 只在「几乎全是正常字符」的候选里挑最长的
        if sc > 0.92 and len(s) > len(best):
            best, bs = s, sc
    return best


# ---------------------------------------------------------------- 抓取
def fetch_brief(day):
    """返回 (msg_dict, hint)。day 为 datetime.date。"""
    conv = _find_group_conv()
    if not conv:
        return None, "解密库里找不到群「%s」" % GROUP_NAME
    p = os.path.join(DEC, "message.db")
    if not os.path.exists(p):
        return None, "解密库缺少 message.db，请先刷新"

    t0 = int(time.mktime(datetime.datetime.combine(day, datetime.time(0, 0)).timetuple()))
    t1 = t0 + 86400
    c = _connect(p)
    try:
        rows = c.execute(
            "select send_time, content from message_table "
            "where conversation_id=? and content_type=0 and send_time>=? and send_time<? "
            "order by send_time desc", (conv, t0, t1)
        ).fetchall()
    finally:
        c.close()

    for send_time, content in rows:
        t = _text(content)
        if KEYWORD in t:
            return {"send_time": send_time, "text": t,
                    "at": datetime.datetime.fromtimestamp(send_time).strftime("%H:%M")}, None

    # 当天没找到 → 回看最近一条，告诉调用方它还没发
    c = _connect(p)
    try:
        rows = c.execute(
            "select send_time, content from message_table "
            "where conversation_id=? and content_type=0 order by send_time desc limit 400",
            (conv,)
        ).fetchall()
    finally:
        c.close()
    for send_time, content in rows:
        t = _text(content)
        if KEYWORD in t:
            last = datetime.datetime.fromtimestamp(send_time).strftime("%Y-%m-%d %H:%M")
            return None, "当天（%s）未找到简报TV；群里最近一期是 %s" % (day, last)
    return None, "群里从未出现「%s」" % KEYWORD


# ---------------------------------------------------------------- 解析
def parse_brief(text):
    """→ dict(date,title,content)。text 为完整原文。"""
    lines = [l.rstrip() for l in text.replace("\r\n", "\n").split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)

    head = lines[0].strip()
    # 标题行形如：2026年10月19日集团简报TV，星期一，第1714期微新闻，6O秒微新闻，众览天下事；
    title = head.rstrip("；;。 ").replace("6O秒", "60秒").replace("6Ｏ秒", "60秒")

    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", head)
    if m:
        date = "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    else:
        date = datetime.date.today().strftime("%Y-%m-%d")

    body = "\n".join(lines[1:]).strip("\n")
    blocks = []
    for b in re.split(r"\n\s*\n", body):
        b = b.strip("\n").strip()
        if b:
            blocks.append(re.sub(r"[ \t]+\n", "\n", b))

    return {"date": date, "title": title, "content": blocks}


# ---------------------------------------------------------------- 写入
def _js_literal(item):
    """生成与现有风格一致的 JS 对象字面量（键不带引号）。"""
    parts = ['    date: %s,' % json.dumps(item["date"], ensure_ascii=False)]
    if item.get("time"):
        parts.append('    time: %s,' % json.dumps(item["time"], ensure_ascii=False))
    parts.append('    cat: %s,' % json.dumps(item.get("cat", "brief"), ensure_ascii=False))
    parts.append('    title: %s,' % json.dumps(item["title"], ensure_ascii=False))
    parts.append('    content: [')
    inner = [",\n\n      ".join(json.dumps(x, ensure_ascii=False) for x in item["content"])]
    parts.append("      " + inner[0])
    parts.append("    ]")
    return "  {\n" + "\n".join(parts) + "\n  }"


def merge_notices(item, dry=False):
    """按期号去重后插到数组顶部。返回 (changed, msg)。"""
    item.setdefault("cat", "brief")      # 简报单独板块展示，不混进重要通知
    with io.open(NOTICES_JS, "r", encoding="utf-8") as f:
        src = f.read()

    m = re.search(r"第\s*(\d+)\s*期", item["title"])
    issue = m.group(1) if m else None
    if issue and re.search(r"第\s*%s\s*期" % issue, src):
        return False, "第%s期已存在，跳过" % issue

    anchor = re.search(r"(window\.__NOTICES__\s*=\s*\[)", src)
    if not anchor:
        return False, "notices.js 结构异常：找不到 window.__NOTICES__ 数组"

    block = _js_literal(item)
    new = src[:anchor.end()] + "\n" + block + ",\n" + src[anchor.end():]
    if dry:
        return True, "（dry-run，未写入）将插入：%s" % item["title"]

    with io.open(NOTICES_JS, "w", encoding="utf-8", newline="") as f:
        f.write(new)
    return True, "已写入第%s期：%s（正文 %d 段）" % (issue or "?", item["title"], len(item["content"]))


# ---------------------------------------------------------------- 主流程
def main():
    args = sys.argv[1:]
    dry = "--dry" in args
    no_refresh = "--no-refresh" in args
    day = datetime.date.today()
    if "--date" in args:
        i = args.index("--date")
        if i + 1 < len(args):
            day = datetime.datetime.strptime(args[i + 1], "%Y-%m-%d").date()

    print("目标日期: %s" % day)
    if not no_refresh:
        ok, msg = _refresh_db()
        print("刷新解密库:", msg)
        if not ok:
            _state(False, "解密库刷新失败", msg)
            return 2
    else:
        print("刷新解密库: 已跳过（--no-refresh）")

    msg, hint = fetch_brief(day)
    if not msg:
        print("抓取结果:", hint)
        _state(False, "未抓到简报TV", hint or "")
        return 1

    print("抓到消息: %s 发送" % msg["at"])
    item = parse_brief(msg["text"])
    item["time"] = msg.get("at", "")
    item["cat"] = "brief"
    print("标题: %s" % item["title"])
    print("日期: %s | 正文 %d 段" % (item["date"], len(item["content"])))
    if dry:
        for i, b in enumerate(item["content"][:4]):
            print("  [%d] %s" % (i, b[:80].replace("\n", " / ")))

    changed, note = merge_notices(item, dry=dry)
    print("写入:", note)
    _state(changed and not dry, note, item["title"])
    return 0 if changed else 1


def _state(changed, note, title):
    d = {"changed": bool(changed), "note": note, "title": title,
         "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    try:
        with io.open(STATE_JSON, "w", encoding="utf-8") as f:
            f.write(json.dumps(d, ensure_ascii=False, indent=1))
    except Exception as e:
        print("状态文件写入失败:", e)


if __name__ == "__main__":
    sys.exit(main())
