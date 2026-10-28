# -*- coding: utf-8 -*-
"""
导航页「整体版式」样板生成器（版式 A / 版式 B）
================================================
配色沿用已定稿的「方案一 · 晨雾杂志版」（_make_nav_themes.V1），标题统一为「分公司·看板导航」。
区别于主题脚本：这里会**重构 body 结构**（不只是换颜色），数据与交互脚本从 index.html 原样继承。

  deploy_nav/layout-a.html  版式 A · 侧栏导航控制台（SaaS 后台范式：固定侧栏 + 滚动定位 + 主内容区）
  deploy_nav/layout-b.html  版式 B · 全幅 Hero 门户（官网范式：渐变 Hero + 统计 + 焦点大卡 + 分区流）

用法：python _make_nav_layouts.py
"""
import os, re, io

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE_DIR, "deploy_nav", "index.html")
SITE_TITLE = "分公司·看板导航"

from _make_nav_themes import COMMON, V1   # 公共骨架 + 方案一配色

# ----------------------------------------------------------------- 公共片段
CAROUSEL_NOTICE = """
    <div class="notice-area">
      <div class="notice-carousel" id="bannerBox">
        <div class="nc-viewport"><div class="nc-track" id="ncTrack"></div></div>
        <button type="button" class="nc-arrow nc-prev" id="ncPrev" aria-label="上一张">‹</button>
        <button type="button" class="nc-arrow nc-next" id="ncNext" aria-label="下一张">›</button>
        <div class="nc-dots" id="ncDots"></div>
      </div>
      <div class="notice-scroll">
        <div class="ns-head">📢 重要通知<span class="more" id="nsCount"></span></div>
        <ul class="ns-list" id="nsList"></ul>
      </div>
    </div>
"""

# 集团简报TV（独立板块，位于公告区下方）· 单次最多 4 期，左右翻页看往期
BRIEF_PANEL = """
    <section class="brief-panel" id="briefPanel">
      <div class="bp-head">
        <span class="bp-title">📺 集团简报TV</span>
        <span class="bp-sub">每日 60 秒 · 众览天下事</span>
        <span class="bp-more" id="bpCount"></span>
        <div class="bp-nav">
          <button type="button" class="bp-arrow" id="bpPrev" aria-label="上一组（较新）" title="上一组（较新）">‹</button>
          <span class="bp-page" id="bpPage">1 / 1</span>
          <button type="button" class="bp-arrow" id="bpNext" aria-label="下一组（更早）" title="下一组（更早）">›</button>
        </div>
      </div>
      <div class="bp-grid" id="bpGrid"></div>
    </section>
"""

NO_RESULT = """
    <div class="no-results" id="noResults">
      <span class="icon">🔍</span>
      <div class="text">没有匹配的看板</div>
      <div class="sub">换个关键词试试，或按 Esc 清空</div>
    </div>
"""

LIGHTBOX = """
<button class="back-top" id="backTop" title="回到顶部">↑</button>

<!-- ===== 灯箱：点板块看图集 ===== -->
<div class="lb-mask" id="lbMask">
  <div class="lb-head">
    <span class="lb-title" id="lbTitle"></span>
    <span class="lb-count" id="lbCount"></span>
    <button type="button" class="lb-close" id="lbClose" title="关闭（Esc）">✕</button>
  </div>
  <button type="button" class="lb-arrow lb-prev" id="lbPrev" aria-label="上一张">‹</button>
  <img class="lb-img" id="lbImg" alt="">
  <button type="button" class="lb-arrow lb-next" id="lbNext" aria-label="下一张">›</button>
  <div class="lb-thumbs" id="lbThumbs"></div>
</div>

<!-- ===== 通知详情弹窗 ===== -->
<div class="nt-mask" id="ntMask" role="dialog" aria-modal="true" aria-labelledby="ntTitle">
  <div class="nt-box">
    <div class="nt-head">
      <div style="flex:1;min-width:0;">
        <div class="nt-title" id="ntTitle"></div>
        <div class="nt-date" id="ntDate"></div>
      </div>
      <button type="button" class="nt-close" id="ntClose" title="关闭（Esc）">✕</button>
    </div>
    <div class="nt-body" id="ntBody"></div>
  </div>
</div>
"""

SEARCH_BOX = """
      <div class="search-box">
        <span class="search-icon">🔍</span>
        <input type="text" id="searchInput" placeholder="搜索看板…" autocomplete="off">
        <button class="clear-btn" id="clearBtn" title="清空">✕</button>
      </div>
"""

# ----------------------------------------------------------------- 简报TV 样式（版式无关，随骨架统一注入）
BRIEF_CSS = r"""
/* ===== 集团简报TV（独立板块，位于公告区下方）· 单次最多 4 期，左右翻页看往期 ===== */
.brief-panel{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);padding:14px 16px 16px;margin:0 0 18px}
.bp-head{display:flex;align-items:center;gap:10px;margin-bottom:12px;flex-wrap:wrap}
.bp-title{font-size:15px;font-weight:700;color:var(--ink);display:flex;align-items:center;gap:8px}
.bp-title::before{content:'';width:4px;height:15px;border-radius:2px;background:var(--bar);flex-shrink:0}
.bp-sub{font-size:12px;color:var(--ink3)}
.bp-more{margin-left:auto;font-size:12px;color:var(--ink3);font-weight:500}
.bp-nav{display:flex;align-items:center;gap:5px;flex-shrink:0}
.bp-arrow{width:26px;height:26px;padding:0;border:1px solid var(--line);border-radius:8px;background:var(--panel2);color:var(--ink2);font-size:15px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .15s,color .15s,border-color .15s,opacity .15s}
.bp-arrow:hover:not(:disabled){background:var(--bar);border-color:var(--bar);color:#fff}
.bp-arrow:disabled{opacity:.3;cursor:default}
.bp-page{min-width:40px;text-align:center;font-size:11.5px;color:var(--ink3);font-variant-numeric:tabular-nums}
.bp-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
.bp-grid.bp-in{animation:bpIn .24s ease}
@keyframes bpIn{from{opacity:0;transform:translateX(12px);}to{opacity:1;transform:none;}}
@media (max-width:1240px){.bp-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:560px){.bp-grid{grid-template-columns:1fr}.bp-sub{display:none}}
.bp-card{position:relative;background:var(--panel2);border:1px solid var(--line);border-radius:12px;padding:11px 13px 12px;cursor:pointer;transition:transform .16s,border-color .16s,box-shadow .16s;display:flex;flex-direction:column;gap:6px;min-width:0}
.bp-card:hover{transform:translateY(-2px);border-color:var(--bar);box-shadow:0 8px 20px rgba(17,24,39,.10)}
.bp-card:hover .bp-go{color:var(--ns-hover-ink)}
.bp-issue{align-self:flex-start;font-size:11px;font-weight:700;color:#fff;background:var(--bar);border-radius:999px;padding:1px 9px;font-variant-numeric:tabular-nums}
.bp-date{font-size:11.5px;color:var(--ink3);font-variant-numeric:tabular-nums}
.bp-tt{font-size:13px;font-weight:600;color:var(--ink2);line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.bp-go{margin-top:auto;font-size:11.5px;color:var(--ink3);transition:color .15s}
"""

# ----------------------------------------------------------------- 版式 A
LAYOUT_A_CSS = r"""
/* ============================================================
   版式 A · 侧栏导航控制台
   固定侧栏（品牌 + 搜索 + 分区锚点导航）+ 主内容区（轮播/通知 + 分区卡片流）
   ============================================================ */
.layout{display:flex;align-items:flex-start;max-width:1460px;margin:0 auto}
.side{position:sticky;top:0;flex:0 0 252px;width:252px;height:100vh;background:#fff;border-right:1px solid var(--line);padding:18px 14px 14px;display:flex;flex-direction:column;overflow:hidden}
.side-brand{display:flex;align-items:center;gap:10px;padding:0 6px 16px}
.side-brand .brand-logo{width:36px;height:36px;border-radius:11px;font-size:17px;box-shadow:0 4px 12px rgba(91,115,150,.32)}
.brand-txt{min-width:0}
.brand-txt b{display:block;font-size:14.5px;font-weight:800;color:var(--ink);line-height:1.35;letter-spacing:.3px}
.brand-txt em{display:block;font-style:normal;font-size:10px;color:var(--ink3);letter-spacing:1px;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;margin-top:3px}
.side-search{padding:0 4px 12px}
.side-search .search-box{width:100%}
.side-search .search-box input{width:100%;background:#f3f6fa;border-color:var(--line);color:var(--ink)}
.side-search .search-box input::placeholder{color:var(--ink3)}
.side-search .search-box input:focus{background:#fff;border-color:var(--sc-default);box-shadow:0 0 0 3px rgba(91,115,150,.14)}
.side-nav{flex:1;min-height:0;overflow-y:auto;padding:2px 0 8px}
.nv{display:flex;align-items:center;gap:10px;padding:9px 10px;margin:1px 0;border-radius:9px;font-size:13.5px;font-weight:600;color:var(--ink2);text-decoration:none;transition:background .15s,color .15s,box-shadow .15s}
.nv:hover{background:#f2f5f9;color:var(--ink)}
.nv.on{background:var(--sc,var(--sc-default));color:#fff;box-shadow:0 4px 12px rgba(27,36,48,.18)}
.nv .nv-ic{width:20px;text-align:center;font-size:14px;flex-shrink:0}
.nv .nv-tx{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.side-foot{padding-top:11px;border-top:1px solid var(--line);font-size:11.5px;color:var(--ink3);line-height:1.7}
.main{flex:1;min-width:0;padding:20px 26px 44px}
.main-head{margin:0 0 16px}
.main-head h1{font-size:23px;font-weight:800;letter-spacing:.6px;color:var(--ink)}
.main-head .sub{margin-top:6px;font-size:12.5px;color:var(--ink3)}
html{scroll-behavior:smooth}
.footer{margin-top:14px}
@media (max-width:1000px){
  .layout{flex-direction:column}
  .side{position:sticky;top:0;z-index:250;flex:none;width:100%;height:auto;border-right:none;border-bottom:1px solid var(--line);padding:10px 12px}
  .side-brand{padding:0 2px 10px}
  .side-search{padding:0 0 10px}
  .side-nav{display:flex;gap:6px;overflow-x:auto;padding:2px 0}
  .nv{margin:0;white-space:nowrap;flex:0 0 auto}
  .side-foot{display:none}
  .main{padding:14px 14px 34px}
  .main-head h1{font-size:19px}
}
"""

BODY_A = """
<div class="layout">
  <aside class="side">
    <div class="side-brand">
      <img class="brand-logo" src="logo.png" alt="看板导航 logo">
      <div class="brand-txt"><b>分公司·看板导航</b><em>HANGZHOU · DASHBOARD</em></div>
    </div>
    <div class="side-search">
%s
    </div>
    <nav class="side-nav" id="sideNav"></nav>
    <div class="side-foot"><span id="sideCount"></span></div>
  </aside>

  <main class="main">
    <div class="main-head">
      <h1>看板导航</h1>
      <p class="sub">数据由 BI 自动同步 · <span id="footVer"></span></p>
    </div>
%s
    <div id="cardContainer"></div>
%s
    <footer class="footer">分公司 · 看板导航<span class="dot">·</span> / 集团数据看板统一入口</footer>
  </main>
</div>
%s
""" % (SEARCH_BOX, CAROUSEL_NOTICE + BRIEF_PANEL, NO_RESULT, LIGHTBOX)

JS_A = r"""
<script>
/* 版式 A：侧栏分区导航 + 滚动定位高亮 */
(function () {
  var titles = [].slice.call(document.querySelectorAll('#cardContainer .section-title'));
  var nav = document.getElementById('sideNav');
  if (!nav || !titles.length) return;
  nav.innerHTML = '';            /* 幂等：重复执行也不会叠加出两份导航 */
  titles.forEach(function (t, i) {
    t.id = 'sec-' + i;
    var raw = ((t.childNodes[0] && t.childNodes[0].nodeValue) || t.textContent || '').trim();
    var ch = Array.from(raw);
    var a = document.createElement('a');
    a.className = 'nv';
    a.href = '#sec-' + i;
    a.innerHTML = '<span class="nv-ic"></span><span class="nv-tx"></span>';
    a.querySelector('.nv-ic').textContent = ch[0] || '';
    a.querySelector('.nv-tx').textContent = ch.slice(1).join('').trim() || raw;
    var sc = getComputedStyle(t).getPropertyValue('--sc');
    if (sc && sc.trim()) a.style.setProperty('--sc', sc.trim());
    nav.appendChild(a);
  });
  var cnt = document.getElementById('sideCount');
  if (cnt) {
    cnt.textContent = titles.length + ' 个分区 · ' +
      document.querySelectorAll('#cardContainer .file-card').length + ' 个入口';
  }
  var items = [].slice.call(nav.children);
  function spy() {
    var y = window.scrollY + 170, idx = 0;
    titles.forEach(function (t, i) {
      var top = t.getBoundingClientRect().top + window.scrollY;
      if (top <= y) idx = i;
    });
    items.forEach(function (a, i) { a.classList.toggle('on', i === idx); });
  }
  window.addEventListener('scroll', spy, { passive: true });
  window.addEventListener('resize', spy);
  spy();
})();
</script>
"""

# ----------------------------------------------------------------- 版式 B
LAYOUT_B_CSS = r"""
/* ============================================================
   版式 B · 全幅 Hero 门户
   细顶栏 + 渐变 Hero（大标题 / 统计数字）+ 悬浮面板（轮播+通知）+ 焦点大卡 + 分区流
   ============================================================ */
.topbar{position:sticky;top:0;z-index:220;background:rgba(255,255,255,.92);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);border-bottom:1px solid var(--line)}
.topbar-inner{max-width:1280px;margin:0 auto;display:flex;align-items:center;gap:16px;padding:11px 24px}
.topbar .brand{display:flex;align-items:center;gap:9px;font-size:15px;font-weight:800;color:var(--ink);letter-spacing:.3px}
.topbar .brand .brand-logo{width:30px;height:30px;border-radius:9px;font-size:15px}
.topbar .search-box{margin-left:auto}
.topbar .search-box input{background:#f3f6fa;border-color:var(--line);color:var(--ink);width:268px}
.topbar .search-box input::placeholder{color:var(--ink3)}
.topbar .search-box input:focus{background:#fff;border-color:var(--sc-default);box-shadow:0 0 0 3px rgba(91,115,150,.14)}
.hero{position:relative;background:linear-gradient(118deg,#3a5573 0%,#5b7396 52%,#7490b0 100%);color:#fff;padding:46px 24px 76px;overflow:hidden}
.hero::before{content:'';position:absolute;inset:0;background:radial-gradient(720px 330px at 78% 8%,rgba(255,255,255,.17),transparent 62%)}
.hero::after{content:'';position:absolute;right:-60px;bottom:-90px;width:340px;height:340px;border-radius:50%;background:rgba(255,255,255,.07)}
.hero-inner{position:relative;z-index:1;max-width:1280px;margin:0 auto}
.hero h1{font-size:34px;font-weight:800;letter-spacing:1.6px;line-height:1.25}
.hero .lead{margin-top:11px;font-size:14px;opacity:.9;letter-spacing:.3px}
.hero-stats{display:flex;flex-wrap:wrap;gap:38px;margin-top:26px}
.stat b{display:block;font-size:27px;font-weight:800;line-height:1.1;font-variant-numeric:tabular-nums}
.stat span{display:block;margin-top:5px;font-size:12px;opacity:.85;letter-spacing:.5px}
.wrap{position:relative;z-index:2;max-width:1280px;margin:-44px auto 0;padding:0 24px 40px}
.panel{background:#fff;border:1px solid var(--line);border-radius:16px;box-shadow:0 10px 30px rgba(27,36,48,.09);padding:14px;margin-bottom:24px}
.feat-head{display:flex;align-items:center;gap:9px;font-size:16px;font-weight:800;color:var(--ink);margin:2px 2px 13px;letter-spacing:.3px}
.feat-head::before{content:'';width:5px;height:18px;border-radius:3px;background:linear-gradient(180deg,#c8952e,#e2b862)}
.feat-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(234px,1fr));gap:14px;margin:0 2px 26px}
.file-card.big{padding:16px 18px;border-left-width:5px}
.file-card.big .card-title{font-size:15.5px;gap:11px}
.file-card.big .card-title .icon{width:44px;height:44px;border-radius:13px;font-size:21px}
.file-card.big .card-desc{margin-top:10px;font-size:13px}
.section-title{margin-top:28px}
.footer{border-top:1px solid var(--footer-line);margin-top:18px}
html{scroll-behavior:smooth}
@media (max-width:768px){
  .hero{padding:30px 16px 64px}
  .hero h1{font-size:24px;letter-spacing:1px}
  .hero .lead{font-size:13px}
  .hero-stats{gap:22px;margin-top:20px}
  .stat b{font-size:22px}
  .wrap{padding:0 14px 30px;margin-top:-36px}
  .panel{padding:12px}
  .topbar-inner{padding:10px 14px;gap:10px}
  .topbar .brand{font-size:13.5px}
  .topbar .search-box input{width:150px}
}
"""

BODY_B = """
<header class="topbar">
  <div class="topbar-inner">
    <div class="brand"><span class="brand-logo">杭</span>分公司·看板导航</div>
%s
  </div>
</header>

<section class="hero">
  <div class="hero-inner">
    <h1>分公司 · 看板导航</h1>
    <p class="lead">一个入口，直达 / 集团全部数据看板、业绩报表与常用工具</p>
    <div class="hero-stats">
      <div class="stat"><b id="statEnt">-</b><span>看板入口</span></div>
      <div class="stat"><b id="statCat">-</b><span>分区</span></div>
      <div class="stat"><b id="statAlb">-</b><span>激励图集</span></div>
    </div>
  </div>
</section>

<div class="wrap">
  <div class="panel">
%s
  </div>

  <div class="feat-head" id="featHead">⭐ 常用看板</div>
  <div class="feat-row" id="featRow"></div>

  <div id="cardContainer"></div>
%s
  <footer class="footer">分公司 · 看板导航<span class="dot">·</span>数据由 BI 自动同步<br><span id="footVer"></span></footer>
</div>
%s
""" % (SEARCH_BOX, CAROUSEL_NOTICE + BRIEF_PANEL, NO_RESULT, LIGHTBOX)

JS_B = r"""
<script>
/* 版式 B：Hero 统计 + 焦点大卡（克隆自分区，随搜索同步显隐） */
(function () {
  function set(id, v) { var e = document.getElementById(id); if (e) e.textContent = v; }
  var cards = [].slice.call(document.querySelectorAll('#cardContainer .file-card'));
  set('statEnt', cards.length);
  set('statCat', document.querySelectorAll('#cardContainer .section-title').length);
  set('statAlb', (window.__ALBUMS__ || []).length);

  var FEAT = ['分公司数据看板（主看板）', '分公司业绩含续费', '分公司订单商品明细', '当日集团业绩大屏'];
  var box = document.getElementById('featRow');
  if (box) {
    FEAT.forEach(function (name) {
      var hit = cards.filter(function (c) {
        var sp = c.querySelectorAll('.card-title span');
        return sp[1] && sp[1].textContent === name;
      })[0];
      if (!hit) return;
      var cl = hit.cloneNode(true);
      cl.classList.add('big');
      box.appendChild(cl);
    });
    var inp = document.getElementById('searchInput');
    function sync() {
      var q = (inp.value || '').trim().toLowerCase(), shown = 0;
      [].forEach.call(box.querySelectorAll('.file-card'), function (c) {
        var ok = !q || (c.getAttribute('data-k') || '').indexOf(q) >= 0;
        c.style.display = ok ? '' : 'none';
        if (ok) shown++;
      });
      box.style.display = shown ? '' : 'none';
      var h = document.getElementById('featHead');
      if (h) h.style.display = shown ? '' : 'none';
    }
    inp.addEventListener('input', sync);
    document.getElementById('clearBtn').addEventListener('click', function () { setTimeout(sync, 0); });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') setTimeout(sync, 0); });
    sync();
  }
})();
</script>
"""

LAYOUTS = [
    ("layout-a.html", "版式 A · 侧栏导航控制台", LAYOUT_A_CSS, BODY_A, JS_A),
    ("layout-b.html", "版式 B · 全幅 Hero 门户", LAYOUT_B_CSS, BODY_B, JS_B),
]


def build_one(src, layout_css, body, extra_js, label, tip_html=""):
    """把某个版式的 CSS + BODY + 附加 JS 拼进页面骨架，返回完整 HTML。"""
    scripts = re.findall(r"<script\b[\s\S]*?</script>", src)
    if len(scripts) < 3:
        raise SystemExit("index.html 脚本块数量异常：%d" % len(scripts))
    # 幂等：正式页里可能已含上一轮追加的版式脚本，按内容去重后再拼回去
    seen, uniq = set(), []
    for s in scripts:
        if s not in seen:
            seen.add(s)
            uniq.append(s)
    scripts = uniq

    style_re = re.compile(r"<style>.*?</style>", re.S)
    title_re = re.compile(r"<title>.*?</title>", re.S)
    body_re = re.compile(r"<body>[\s\S]*?</body>", re.S)

    from _make_nav_themes import COMMON, V1
    css = (COMMON.strip() + "\n" + V1.strip() + "\n" + layout_css.strip()
           + "\n" + BRIEF_CSS.strip())
    tail = "" if extra_js.strip() in scripts else extra_js.strip()   # 已存在则不重复追加
    new_body = ("<body>\n" + (tip_html + "\n" if tip_html else "")
                + body.strip() + "\n\n" + "\n".join(scripts)
                + ("\n\n" + tail if tail else "") + "\n</body>")
    out = style_re.sub("<style>\n" + css + "\n</style>", src, count=1)
    out = title_re.sub("<title>%s%s</title>" % (SITE_TITLE, label), out, count=1)
    if 'rel="icon"' not in out:      # 站点图标与侧栏 logo 同一张图
        out = title_re.sub(
            '<title>%s%s</title>\n<link rel="icon" type="image/png" href="logo.png">'
            % (SITE_TITLE, label), out, count=1)
    out = body_re.sub(lambda m: new_body, out, count=1)
    return out


def apply_layout(key):
    """把选定版式固化进正式 deploy_nav/index.html（不带预览提示条）。"""
    table = {fname.split(".")[0]: (css, body, js, label)
             for fname, label, css, body, js in LAYOUTS}
    if key not in table:
        raise SystemExit("未知版式：%s（可选 %s）" % (key, "/".join(table)))
    css, body, js, label = table[key]
    with io.open(SRC, "r", encoding="utf-8") as f:
        src = f.read()
    out = build_one(src, css, body, js, "")
    with io.open(SRC, "w", encoding="utf-8", newline="") as f:
        f.write(out)
    print("applied:", SRC, len(out), "chars ·", label)


def main():
    with io.open(SRC, "r", encoding="utf-8") as f:
        src = f.read()

    for fname, label, layout_css, body, extra_js in LAYOUTS:
        out = build_one(src, layout_css, body, extra_js, "（%s）" % label)
        tip = ('<div class="theme-tip" style="position:sticky;top:0;z-index:400;'
               'background:#111827;color:#fff;font-size:12.5px;padding:7px 14px;'
               'text-align:center;letter-spacing:.4px">版式预览：%s　·　配色沿用方案一　·　'
               '正式页面 index.html 暂未套用</div>' % label)
        out = out.replace("<body>", "<body>\n" + tip, 1)
        path = os.path.join(BASE_DIR, "deploy_nav", fname)
        with io.open(path, "w", encoding="utf-8", newline="") as f:
            f.write(out)
        print("written:", fname, len(out), "chars ·", label)


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "apply":
        apply_layout(sys.argv[2])
    else:
        main()
