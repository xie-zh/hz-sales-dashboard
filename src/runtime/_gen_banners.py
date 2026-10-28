# -*- coding: utf-8 -*-
"""扫描 deploy_nav/banners/ 下的板块文件夹，生成轮播图集清单并写回 index.html。

规则（与用户约定一致）：
  · banners/ 下**每个子文件夹 = 一个板块**，文件夹名就是板块名
  · 文件夹内的图片按文件名排序，**第一张作轮播封面**
  · 点封面 → 弹出该文件夹全部图片
  · 文件夹名可带排序前缀（`1.分公司激励` / `2、经理激励`），显示时自动去掉前缀
  · 没有图片的空文件夹会被跳过（会打印提示）

用法：
  python _gen_banners.py          扫描并写回 deploy_nav/index.html
  python _gen_banners.py --dry    只打印将要生成的内容，不写文件

生成位置：index.html 里
  /* ==== ALBUMS:BEGIN ==== */ ... /* ==== ALBUMS:END ==== */
之间的 `window.__ALBUMS__ = [...]`。
"""
import os
import re
import sys
import json
import io
from urllib.parse import quote

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
NAV_DIR = os.path.join(ROOT, "deploy_nav")
BANNER_DIR = os.path.join(NAV_DIR, "banners")
HTML = os.path.join(NAV_DIR, "index.html")

IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".avif"}
PREFIX_RE = re.compile(r"^\s*\d+\s*[.、．_\-\s]\s*")   # 1.xxx / 2、xxx / 3_xxx / 4 xxx


def display_name(folder: str) -> str:
    return PREFIX_RE.sub("", folder).strip() or folder


def sort_key(folder: str):
    """带数字前缀的按前缀排；没前缀的按文件夹修改时间（＝创建先后）排，并排在最后。"""
    m = re.match(r"^\s*(\d+)\s*[.、．_\-\s]", folder)
    if m:
        return (0, int(m.group(1)), folder)
    try:
        mt = os.path.getmtime(os.path.join(BANNER_DIR, folder))
    except OSError:
        mt = 0
    return (1, mt, folder)


def scan():
    albums, skipped = [], []
    if not os.path.isdir(BANNER_DIR):
        return albums, skipped
    for name in sorted(os.listdir(BANNER_DIR)):
        p = os.path.join(BANNER_DIR, name)
        if not os.path.isdir(p):
            continue
        files = [f for f in os.listdir(p)
                 if os.path.isfile(os.path.join(p, f)) and os.path.splitext(f)[1].lower() in IMG_EXT]
        files.sort()
        if not files:
            skipped.append(name)
            continue
        urls = ["banners/%s/%s" % (quote(name), quote(f)) for f in files]
        albums.append({"name": display_name(name), "folder": name,
                       "count": len(files), "images": urls})
    albums.sort(key=lambda a: sort_key(a["folder"]))
    return albums, skipped


def build_js(albums):
    lines = ["window.__ALBUMS__ = ["]
    for a in albums:
        lines.append("  { name: %s, images: [" % json.dumps(a["name"], ensure_ascii=False))
        for u in a["images"]:
            lines.append("    %s," % json.dumps(u, ensure_ascii=False))
        lines.append("  ] },")
    lines.append("];")
    return "\n".join(lines)


def main(argv):
    dry = "--dry" in argv
    albums, skipped = scan()

    print("板块目录：%s" % BANNER_DIR)
    print("识别到 %d 个板块：" % len(albums))
    for i, a in enumerate(albums, 1):
        print("  %d) %-12s %2d 张   封面 = %s" % (i, a["name"], a["count"], a["images"][0]))
    if skipped:
        print("跳过空文件夹：%s" % "、".join(skipped))

    block = build_js(albums)
    if dry:
        print("\n=== 将写入的 JS ===")
        print(block)
        return 0

    if not os.path.isfile(HTML):
        print("× 找不到 %s" % HTML)
        return 1
    html = io.open(HTML, encoding="utf-8").read()
    a = html.find("/* ==== ALBUMS:BEGIN ==== */")
    b = html.find("/* ==== ALBUMS:END ==== */")
    if a < 0 or b < 0 or b < a:
        print("× index.html 里找不到 ALBUMS 标记块，无法写入")
        return 1
    a += len("/* ==== ALBUMS:BEGIN ==== */")
    new = html[:a] + "\n" + block + "\n" + html[b:]
    io.open(HTML, "w", encoding="utf-8").write(new)
    print("\n√ 已写回 index.html（%d 个板块 / %d 张图）" %
          (len(albums), sum(a["count"] for a in albums)))
    if skipped:
        print("  提示：%s 里还没有图片，本次未收录" % "、".join(skipped))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
