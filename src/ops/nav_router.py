# -*- coding: utf-8 -*-
"""_repoint_nav.py —— 把导航页里指向腾讯云的链接，改指到新的 WorkBuddy 分享链接

背景：导航页（顶部那张卡片列表）原先三个卡片全部指向腾讯云 CloudBase 的
/dash/、/report/、/title/ 子路径。腾讯云账号收回后这些链接会整体失效，
需要按「迁移配置.json」里的新地址重写。

改哪个文件：<根>\\deploy_tool\\build\\index.html（导航页的构建产物）
改完还要**重新发布导航页**（在 WorkBuddy 里把该目录发布为应用）才会生效。

映射规则（「迁移配置.json」→ publish 块）：
    /dash/    → publish.dash_url         主看板，必填
    /report/  → publish.nav_report_url   录音质检看板，没配就置灰
    /title/   → publish.nav_title_url    称号看板，没配就置灰
    （其它 /xxx/ 一律按 nav_report_url 同法处理，见 publish.nav_map 自定义）
    访问记录接口 publish.visit_api       没配就把访问记录整体关掉，不报错

置灰 = 去掉 href、加 off 样式、文案标注「暂未迁移」，避免点进去是死链。

用法：
    python Hzexamplebot\\_repoint_nav.py           按配置改写
    python Hzexamplebot\\_repoint_nav.py --dry     只报告会改什么，不落盘
退出码：0 = 已改写或无改动；1 = 失败
"""

import os
import re
import sys
import time
import shutil

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, BASE)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import _portable                       # noqa: E402
import _migrate                        # noqa: E402

MARK = "hz-repoint:v1"
TB_RE = re.compile(r"https://[a-z0-9-]+\.tcloudbaseapp\.com", re.I)
CARD_RE = re.compile(
    r'<a class="card" href="(?P<href>[^"]*)"(?P<attrs>[^>]*)>(?P<body>.*?)</a>', re.S)
URLSPAN_RE = re.compile(r'(<span class="url">)(.*?)(</span>)', re.S)

OFF_CSS = (
    "  <style id=\"hz-repoint-off\">\n"
    "    /* 由 _repoint_nav.py 注入：暂未迁移到新平台的卡片置灰 */\n"
    "    a.card.off{opacity:.45;filter:grayscale(.5)}\n"
    "    a.card.off .btn{background:#c9ced6}\n"
    "  </style>\n"
)


def _short(url, keep=46):
    u = re.sub(r"^https?://", "", url or "")
    return u if len(u) <= keep else u[:keep - 1] + "…"


def main(argv):
    dry = "--dry" in argv
    root = _portable.resolve_root()
    nav_dir = _portable.nav_dir(root)
    f = os.path.join(nav_dir, "index.html")
    pub = _portable.publish_info(root)

    print("项目根：%s" % root)
    print("导航页：%s" % f)
    if not os.path.isfile(f):
        print("× 找不到导航页文件。请先确认 deploy_tool\\build\\index.html 存在。")
        return 1

    html = open(f, encoding="utf-8").read()

    dash = (pub.get("dash_url") or "").strip()
    mapping = {
        "/dash/": dash,
        "/report/": (pub.get("nav_report_url") or "").strip(),
        "/title/": (pub.get("nav_title_url") or "").strip(),
    }
    for k, v in (pub.get("nav_map") or {}).items():
        mapping[k] = (v or "").strip()

    bases = TB_RE.findall(html)
    old_base = bases[0].rstrip("/") if bases else ""
    print("检测到的旧地址前缀：%s" % (old_base or "（没有，可能已经改过）"))
    print("目标地址：")
    for k, v in mapping.items():
        print("   %-10s → %s" % (k, v or "（未配置 → 置灰）"))

    if not old_base and MARK in html:
        print("\n导航页已经改写过了（标记 %s 在），无需重复处理。" % MARK)
        return 0
    if not dash:
        print("\n× 还没配置主看板链接（publish.dash_url）。")
        print("  请在 WorkBuddy 里把 _deploy 发布为应用，拿到分享链接后回填，再跑本脚本。")
        return 1

    changes = []
    offs = []

    def fix_card(m):
        href = m.group("href")
        attrs = m.group("attrs")
        body = m.group("body")
        if old_base and href.startswith(old_base):
            suffix = href[len(old_base):] or "/"
            if not suffix.endswith("/"):
                suffix += "/"
            tgt = mapping.get(suffix, "")
            if tgt:
                new_href = tgt.rstrip("/") + "/"
                body2 = URLSPAN_RE.sub(
                    lambda mm: mm.group(1) + _short(new_href) + mm.group(3), body)
                changes.append("  %-10s → %s" % (suffix, new_href))
                return '<a class="card" href="%s"%s>%s</a>' % (new_href, attrs, body2)
            # 没配目标 → 置灰
            body2 = URLSPAN_RE.sub(
                lambda mm: mm.group(1) + "暂未迁移到新平台" + mm.group(3), body)
            body2 = body2.replace('class="btn">进入看板 →',
                                  'class="btn">暂未迁移')
            offs.append("  %s" % suffix)
            return '<a class="card off" href="javascript:void(0)"%s>%s</a>' % (attrs, body2)
        return m.group(0)

    new_html = CARD_RE.sub(fix_card, html)

    # 访问记录接口
    visit_new = (pub.get("visit_api") or "").strip()
    old_visit = ""
    mv = re.search(r"var\s+API\s*=\s*'([^']*)'", new_html)
    if mv:
        old_visit = mv.group(1)
        if visit_new:
            new_html = new_html.replace(old_visit, visit_new)
            changes.append("  访问记录接口 → %s" % visit_new)
        else:
            new_html = new_html.replace(
                "var API='" + old_visit + "'", "var API='' /* 访问记录已停用（腾讯云云函数随账号收回）*/")
            changes.append("  访问记录接口 → 已停用")

    if MARK not in new_html:
        new_html = new_html.replace("</head>", OFF_CSS + "</head>", 1)
        new_html = new_html.replace("<head>", "<head>\n<!-- %s -->" % MARK, 1)

    print("\n将发生的改动：")
    print("\n".join(changes) if changes else "  （没有可替换的链接）")
    if offs:
        print("置灰的卡片（还没迁到新平台）：")
        print("\n".join(sorted(set(offs))))

    if new_html == html:
        print("\n内容没有变化，不落盘。")
        return 0
    if dry:
        print("\n--dry：未写入。")
        return 0

    # 备份 + 写入
    snap = os.path.join(os.path.dirname(nav_dir), "snapshots")
    try:
        os.makedirs(snap)
    except Exception:
        pass
    bak = os.path.join(snap, "nav_before_repoint_%s.html" % time.strftime("%Y%m%d_%H%M%S"))
    try:
        shutil.copy2(f, bak)
        print("\n已备份原文件 → %s" % bak)
    except Exception as e:
        print("\n! 备份失败（继续写入）：%s" % str(e)[:80])

    with open(f, "w", encoding="utf-8") as fh:
        fh.write(new_html)
    print("已写入 %s" % f)
    print("\n下一步：到 WorkBuddy 里把这个目录发布为应用（导航页），")
    print("       拿到链接后回填到「迁移配置.json」的 publish.nav_url。")
    print("       目录：%s" % nav_dir)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
