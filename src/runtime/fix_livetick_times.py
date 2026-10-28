# -*- coding: utf-8 -*-
"""修复 livetick「数据时间」与「同步时间」对不上的问题（2026-09-23）。

根因：实时条里的「数据 <live-dt>」直接复制了静态的源数据截止时间（meta-dt，
      即 data_time，如 9月23日 07:30），却和实时时钟（07:57）、“同步于X秒前”
      并列展示 —— 给人“刚同步却是旧数据”的矛盾观感，且数据时间根本没动。

修复：
  1) 实时条「数据」→「数据同步」，显示“最近一次成功同步(fetch)的时刻” CD.last，
     每次刷新就跳到当前时刻 —— 数据时间真正“动起来”，且与“同步于X秒前”自洽。
  2) 顶部原有「数据快照 / 数据日期」→「数据截至」，明确它是“数据内容截止时间”
     （真实、静态），与实时条的“同步时刻”区分开，不再有两个“数据时间”互相打架。

幂等：每段都有存在性检查，重复运行只改一次。
"""
import os, shutil

BASE = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11"
PAGES = ["index.html", "sanli.html", "ry.html"]
BACKUP = os.path.join(BASE, "_backup", "20260923_pre_fixtime")

# --- 实时条 HTML：把「数据」改名「数据同步」，初始值改为时钟格式 ---
OLD_LIVE = '<span>数据 <b id="live-dt">--</b></span>'
NEW_LIVE = '<span>数据同步 <b id="live-dt">--:--:--</b></span>'

# --- JS：live-dt 改为显示最近同步时刻，而非复制 meta-dt ---
OLD_JS = (
    "    if((el=document.getElementById('live-dt'))){\n"
    "      var md=document.getElementById('meta-dt')||document.getElementById('meta-dt2');\n"
    "      if(md){ var t=(md.textContent||'').trim(); if(t) el.textContent=t; }\n"
    "    }"
)
NEW_JS = (
    "    if((el=document.getElementById('live-dt'))){\n"
    "      el.textContent = hms(new Date(CD.last));\n"
    "    }"
)

# --- 顶部标签：明确为“数据截至”（真实内容截止时间） ---
OLD_SNAP = '<span>数据快照 <b id="meta-dt">--</b></span>'
NEW_SNAP = '<span>数据截至 <b id="meta-dt">--</b></span>'
OLD_DATE = '<span>数据日期 <b id="meta-dt">--</b></span>'
NEW_DATE = '<span>数据截至 <b id="meta-dt">--</b></span>'


def patch(path):
    name = os.path.basename(path)
    s = open(path, encoding="utf-8", errors="ignore").read()
    # 备份一次
    os.makedirs(BACKUP, exist_ok=True)
    dst = os.path.join(BACKUP, name)
    if not os.path.exists(dst):
        shutil.copy2(path, dst)
    cnt = 0
    if OLD_LIVE in s:
        s = s.replace(OLD_LIVE, NEW_LIVE, 1); cnt += 1
    else:
        print("  [warn]%s 未找到实时条「数据」HTML" % name)
    if OLD_JS in s:
        s = s.replace(OLD_JS, NEW_JS, 1); cnt += 1
    else:
        print("  [warn]%s 未找到 live-dt JS 块" % name)
    if OLD_SNAP in s:
        s = s.replace(OLD_SNAP, NEW_SNAP, 1); cnt += 1
    if OLD_DATE in s:
        s = s.replace(OLD_DATE, NEW_DATE, 1); cnt += 1
    if cnt == 0:
        print("  %s 无改动（可能已修过）" % name)
        return cnt
    open(path, "w", encoding="utf-8", newline="").write(s)
    print("  %s 改动 %d 处" % (name, cnt))
    return cnt


if __name__ == "__main__":
    for p in PAGES:
        fp = os.path.join(BASE, p)
        print("=" * 60)
        print(p)
        if not os.path.exists(fp):
            print("  缺失"); continue
        patch(fp)
    # 同步把 patch_livetick.py 的常量也更新，保证将来重跑一致
    pl = os.path.join(BASE, "_hzcheck_k", "patch_livetick.py")
    if os.path.exists(pl):
        t = open(pl, encoding="utf-8", errors="ignore").read()
        if OLD_LIVE in t:
            t = t.replace(OLD_LIVE, NEW_LIVE, 1)
        if OLD_JS in t:
            t = t.replace(OLD_JS, NEW_JS, 1)
        # patch_livetick 里顶部标签不在其注入范围，无需改
        open(pl, "w", encoding="utf-8", newline="").write(t)
        print("=" * 60)
        print("已同步更新 patch_livetick.py 常量")
