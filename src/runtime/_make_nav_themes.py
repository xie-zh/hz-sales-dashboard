# -*- coding: utf-8 -*-
"""
导航页三版样式样板生成器
=========================
读 deploy_nav/index.html（唯一数据源 / 交互逻辑不变），替换 <style> 块与 <title>，
产出三份可独立预览/发布的样板：

  deploy_nav/theme-v1.html  晨雾杂志版（亮色 · 分区配色 · 悬浮卡片）
  deploy_nav/theme-v2.html  深空极光版（深色 · 玻璃拟态 · 霓虹描边）
  deploy_nav/theme-v3.html  极简控制台版（白底网格 · 等宽编号 · 高密度）

用法：python _make_nav_themes.py
数据/链接有变动后重跑本脚本即可刷新三版样板。
"""
import os, re, io, sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE_DIR, "deploy_nav", "index.html")
SITE_TITLE = "分公司·看板导航"      # 全站统一标题（用户 10-19 指定）

# ---------------------------------------------------------------- 公共骨架
# 只含「结构尺寸 + 颜色变量引用」，具体配色与卡片形态由各主题的 DECOR 决定。
COMMON = r"""
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:var(--bg);background-image:var(--bg-img);color:var(--ink);min-height:100vh}
body.lb-open{overflow:hidden}
.container{max-width:1220px;margin:0 auto;padding:22px 24px 0}

/* ===== 顶栏 ===== */
.header{position:sticky;top:0;z-index:200;background:var(--hdr-bg);border-bottom:1px solid var(--hdr-line);backdrop-filter:var(--hdr-blur);-webkit-backdrop-filter:var(--hdr-blur)}
.header-inner{max-width:1220px;margin:0 auto;display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:14px;padding:12px 24px}
.header-brand{display:flex;align-items:center;gap:10px;font-size:19px;font-weight:700;color:var(--hdr-ink);letter-spacing:.3px}
.brand-logo{width:32px;height:32px;border-radius:9px;background:var(--logo-bg);color:var(--logo-ink);border:1px solid var(--logo-line);display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:700;flex-shrink:0}
img.brand-logo{object-fit:cover;display:block;background:#fff}
.header-actions{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.search-box{position:relative;display:flex;align-items:center}
.search-box input{padding:9px 34px 9px 36px;border:1px solid var(--input-line);border-radius:999px;font-size:14px;background:var(--input-bg);color:var(--input-ink);width:250px;outline:none;transition:border-color .18s,background .18s,box-shadow .18s}
.search-box input::placeholder{color:var(--input-ph)}
.search-box input:focus{border-color:var(--focus);background:var(--input-bg-f);box-shadow:0 0 0 3px var(--focus-ring)}
.search-box .search-icon{position:absolute;left:13px;color:var(--ink3);font-size:15px}
.search-box .clear-btn{position:absolute;right:12px;background:none;border:none;color:var(--ink3);font-size:16px;cursor:pointer;display:none}
.search-box .clear-btn.visible{display:block}

/* ===== 顶部：板块轮播 + 重要通知 ===== */
.notice-area{display:grid;grid-template-columns:1.35fr 1fr;gap:16px;margin-bottom:6px;align-items:start}
/* ⚠️ 必须 align-self:start + width:100% + aspect-ratio：否则 aspect-ratio 会把宽度算大压到右侧通知栏 */
.notice-carousel{position:relative;display:flex;flex-direction:column;width:100%;min-width:0;height:auto;align-self:start;aspect-ratio:1.82 / 1;background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden}
.nc-viewport{flex:1;overflow:hidden}
.nc-track{display:flex;height:100%;transition:transform .5s ease}
.nc-slide{position:relative;flex:0 0 100%;min-width:0;overflow:hidden;background:var(--panel2);cursor:pointer}
.nc-blur{position:absolute;inset:-12%;background-size:cover;background-position:center;filter:blur(26px) saturate(1.05);display:none}
.nc-slide.is-contain .nc-blur{display:block}
.nc-slide img{position:relative;z-index:1;width:100%;height:100%;object-fit:cover;display:block;background:transparent;transition:transform .5s ease}
.nc-slide:hover img{transform:scale(1.03)}
.nc-slide img.is-contain{object-fit:contain}
.nc-slide:hover img.is-contain{transform:none}
.nc-mask{position:absolute;inset:0;z-index:2;background:linear-gradient(to top,rgba(0,0,0,.78) 0%,rgba(0,0,0,.55) 38%,rgba(0,0,0,.12) 70%,rgba(0,0,0,0) 100%);opacity:0;transition:opacity .25s ease;display:flex;flex-direction:column;justify-content:flex-end;padding:22px;color:#fff}
.nc-slide:hover .nc-mask{opacity:1}
.nc-name{font-size:22px;font-weight:700;letter-spacing:.5px;text-shadow:0 2px 6px rgba(0,0,0,.45)}
.nc-hint{font-size:13px;margin-top:7px;opacity:.9;display:flex;align-items:center;gap:6px}
.nc-hint .pill{background:rgba(255,255,255,.22);border-radius:999px;padding:2px 10px;font-size:12px;font-weight:600}
.nc-corner{position:absolute;right:14px;top:14px;z-index:2;background:rgba(17,24,39,.55);color:#fff;font-size:12px;font-weight:600;border-radius:999px;padding:4px 11px;backdrop-filter:blur(4px);transition:opacity .25s}
.nc-slide:hover .nc-corner{opacity:0}
.nc-arrow{position:absolute;top:50%;transform:translateY(-50%);width:34px;height:34px;padding:0;border:none;border-radius:999px;background:rgba(17,24,39,.32);color:#fff;font-size:20px;line-height:1;cursor:pointer;opacity:0;transition:opacity .2s,background .2s;z-index:3}
.notice-carousel:hover .nc-arrow{opacity:1}
.nc-arrow:hover{background:rgba(17,24,39,.6)}
.nc-prev{left:10px}
.nc-next{right:10px}
.nc-dots{position:absolute;left:0;right:0;bottom:10px;display:flex;justify-content:center;gap:6px;z-index:3}
.nc-dot{width:8px;height:8px;padding:0;border:none;border-radius:999px;background:rgba(255,255,255,.6);cursor:pointer;transition:background .2s,width .2s}
.nc-dot.active{background:#fff;width:20px}
.nc-empty{position:absolute;inset:14px;border:2px dashed var(--line);border-radius:12px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;text-align:center;padding:20px;background:var(--panel2)}
.nc-empty .ic{font-size:40px;color:var(--ink3)}
.nc-empty .t{font-size:15px;font-weight:700;color:var(--ink2)}
.nc-empty .s{font-size:12.5px;color:var(--ink3);line-height:1.7;max-width:360px}

/* 重要通知 */
.notice-scroll{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);display:flex;flex-direction:column;height:auto;overflow:hidden;aspect-ratio:1.348 / 1}
.ns-head{display:flex;align-items:center;gap:8px;padding:13px 16px;border-bottom:1px solid var(--line);font-size:15px;font-weight:700;color:var(--ink)}
.ns-head::before{content:'';width:4px;height:15px;border-radius:2px;background:var(--bar)}
.ns-head .more{margin-left:auto;font-size:12px;font-weight:500;color:var(--ink3)}
.ns-list{list-style:none;margin:0;padding:6px 0;flex:1;min-height:0;overflow-y:auto}
.ns-item{display:flex;gap:10px;padding:9px 16px;cursor:default;transition:background .15s;align-items:flex-start}
.ns-item[data-i],.ns-item[data-url]{cursor:pointer}
.ns-item[data-i]:hover,.ns-item[data-url]:hover{background:var(--ns-hover)}
.ns-item[data-i]:hover .ns-title,.ns-item[data-url]:hover .ns-title{color:var(--ns-hover-ink)}
.ns-thumb{width:54px;height:38px;flex-shrink:0;border-radius:6px;overflow:hidden;background:var(--panel2);border:1px solid var(--line);display:flex;align-items:center;justify-content:center;font-size:15px;color:var(--ink3)}
.ns-thumb img{width:100%;height:100%;object-fit:cover;display:block}
.ns-main{flex:1;min-width:0;display:flex;flex-direction:column;gap:3px}
.ns-date{color:var(--ink3);font-size:11.5px;font-variant-numeric:tabular-nums}
.ns-title{color:var(--ink2);font-size:13px;line-height:1.45;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;transition:color .15s}
.ns-empty{color:var(--ink3);font-size:13px;text-align:center;padding:36px 16px;line-height:1.9}
.ns-empty .ic{display:block;font-size:30px;margin-bottom:8px;color:var(--line)}

/* ===== 通知详情弹窗 ===== */
.nt-mask{position:fixed;inset:0;z-index:9998;background:var(--mask-bg);backdrop-filter:blur(5px);display:none;align-items:center;justify-content:center;padding:34px 20px}
.nt-mask.visible{display:flex}
.nt-box{width:100%;max-width:760px;max-height:86vh;background:var(--modal-bg);border-radius:16px;box-shadow:0 24px 64px rgba(0,0,0,.3);display:flex;flex-direction:column;overflow:hidden;animation:ntIn .22s ease}
@keyframes ntIn{from{opacity:0;transform:translateY(14px) scale(.985);}to{opacity:1;transform:none;}}
.nt-head{display:flex;align-items:flex-start;gap:14px;padding:20px 24px 15px;border-bottom:1px solid var(--modal-line);background:var(--modal-head-bg)}
.nt-title{font-size:19px;font-weight:700;color:var(--modal-ink);line-height:1.45}
.nt-date{margin-top:7px;font-size:12.5px;color:var(--modal-ink2);font-variant-numeric:tabular-nums}
.nt-close{margin-left:auto;flex-shrink:0;width:32px;height:32px;border:none;border-radius:999px;background:var(--close-bg);color:var(--close-ink);font-size:15px;line-height:1;cursor:pointer;transition:background .18s,color .18s}
.nt-close:hover{background:var(--close-bg-hv);color:var(--modal-ink)}
.nt-body{flex:1;overflow-y:auto;padding:20px 24px 26px}
.nt-body img{max-width:100%;height:auto;display:block;border-radius:10px;margin:0 0 16px;border:1px solid var(--modal-line)}
.nt-fig{display:block;margin:0 0 16px;cursor:zoom-in}
.nt-fig img{margin:0}
.nt-body.nt-photo{padding:16px}
.nt-body.nt-photo .nt-fig{margin:0}
.nt-body.nt-photo img{max-height:calc(86vh - 132px);width:auto;max-width:100%;margin:0 auto;object-fit:contain}
.nt-p{margin:0 0 12px;font-size:14.5px;line-height:1.9;color:var(--modal-ink2);white-space:pre-wrap;word-break:break-word}
.nt-p:last-child{margin-bottom:0}
.nt-list{margin:0 0 12px}
.nt-li{margin:0 0 6px;padding-left:1.95em;text-indent:-1.95em;line-height:1.85}
.nt-li:last-child{margin-bottom:0}
.nt-h{font-size:15.5px;font-weight:700;color:var(--modal-ink);line-height:1.5;margin:20px 0 11px;padding-left:10px;border-left:3px solid var(--bar)}
.nt-h:first-child{margin-top:0}
.nt-body a{color:var(--link)}

/* ===== 分区 / 卡片 ===== */
.section-title{display:flex;align-items:center;gap:9px;font-size:16px;font-weight:700;color:var(--ink);margin:26px 0 12px}
.section-title::before{content:'';width:4px;height:16px;border-radius:2px;background:var(--sc)}
.section-title .badge{font-size:12px;font-weight:500;color:var(--ink2);background:var(--panel2);padding:2px 10px;border-radius:999px;margin-left:auto}
/* 六个分区各自的强调色（顺序：业绩数据 / 集团看板 / 分公司 / 企微表格 / 企微文档 / 工具栏） */
#cardContainer > div:nth-child(1),#cardContainer > div:nth-child(2){--sc:var(--c1)}
#cardContainer > div:nth-child(3),#cardContainer > div:nth-child(4){--sc:var(--c2)}
#cardContainer > div:nth-child(5),#cardContainer > div:nth-child(6){--sc:var(--c3)}
#cardContainer > div:nth-child(7),#cardContainer > div:nth-child(8){--sc:var(--c4)}
#cardContainer > div:nth-child(9),#cardContainer > div:nth-child(10){--sc:var(--c5)}
#cardContainer > div:nth-child(11),#cardContainer > div:nth-child(12){--sc:var(--c6)}
#cardContainer > div:nth-child(13),#cardContainer > div:nth-child(14){--sc:var(--c6)}
.file-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:12px;margin:0 0 18px}
.file-card{text-decoration:none;color:inherit;display:block;cursor:pointer}
.file-card .src{margin-left:auto;font-size:11px;font-weight:600;border-radius:999px;padding:1px 8px;flex-shrink:0}
.src-hz{color:var(--tag-hz-ink);background:var(--tag-hz-bg)}
.src-group{color:var(--tag-gp-ink);background:var(--tag-gp-bg)}
.tool-grid{grid-template-columns:repeat(auto-fill,minmax(174px,1fr));gap:10px}

/* ===== 灯箱（三版统一深色） ===== */
.lb-mask{position:fixed;inset:0;z-index:9999;background:rgba(10,14,20,.92);backdrop-filter:blur(6px);display:none;flex-direction:column;align-items:center;justify-content:center;padding:56px 78px 96px}
.lb-mask.visible{display:flex}
.lb-head{position:absolute;top:0;left:0;right:0;height:56px;display:flex;align-items:center;gap:12px;padding:0 22px;color:#fff}
.lb-title{font-size:16px;font-weight:700}
.lb-count{font-size:13px;color:rgba(255,255,255,.62);font-variant-numeric:tabular-nums}
.lb-close{margin-left:auto;width:36px;height:36px;border:none;border-radius:999px;background:rgba(255,255,255,.14);color:#fff;font-size:17px;cursor:pointer;transition:background .18s}
.lb-close:hover{background:rgba(255,255,255,.28)}
.lb-img{max-width:100%;max-height:100%;object-fit:contain;border-radius:10px;box-shadow:0 24px 70px rgba(0,0,0,.55);background:#fff;cursor:zoom-out}
.lb-arrow{position:absolute;top:50%;transform:translateY(-50%);width:48px;height:48px;border:none;border-radius:999px;background:rgba(255,255,255,.14);color:#fff;font-size:26px;line-height:1;cursor:pointer;transition:background .18s;z-index:2}
.lb-arrow:hover{background:rgba(255,255,255,.3)}
.lb-arrow[disabled]{opacity:.25;cursor:default}
.lb-prev{left:18px}
.lb-next{right:18px}
.lb-thumbs{position:absolute;left:0;right:0;bottom:16px;display:flex;gap:8px;justify-content:center;overflow-x:auto;padding:0 22px}
.lb-thumb{width:66px;height:46px;flex:0 0 auto;border-radius:7px;overflow:hidden;border:2px solid transparent;cursor:pointer;opacity:.55;transition:opacity .18s,border-color .18s;background:#1a2330}
.lb-thumb img{width:100%;height:100%;object-fit:cover;display:block}
.lb-thumb:hover{opacity:.9}
.lb-thumb.on{border-color:#fff;opacity:1}

/* ===== 无结果 / 页脚 / 返回顶部 ===== */
.no-results{display:none;text-align:center;padding:50px 16px;color:var(--ink3)}
.no-results .icon{font-size:44px;display:block;margin-bottom:14px}
.no-results .text{font-size:17px;font-weight:500;color:var(--ink2)}
.no-results .sub{font-size:13px;margin-top:6px}
.no-results.visible{display:block}
.footer{text-align:center;padding:22px 0;color:var(--footer-ink);font-size:13px;border-top:1px solid var(--footer-line);margin-top:10px;line-height:1.8}
.footer .dot{color:var(--ink3);margin:0 7px}
.back-top{position:fixed;bottom:30px;right:30px;width:46px;height:46px;border-radius:999px;background:var(--bt-bg);border:none;color:#fff;font-size:18px;cursor:pointer;display:none;align-items:center;justify-content:center;box-shadow:var(--bt-shadow);z-index:100}
.back-top:hover{background:var(--bt-hv)}
.back-top.visible{display:flex}

@media (max-width:900px){.notice-area{grid-template-columns:1fr;}.notice-carousel{max-width:620px;width:100%;justify-self:center;}.notice-scroll{aspect-ratio:auto;height:auto;min-height:260px;}}
@media (max-width:768px){
  .header-inner{flex-direction:column;align-items:stretch;gap:10px;padding:12px 16px;}
  .header-brand{font-size:17px;justify-content:center;}
  .header-actions{justify-content:center;}
  .search-box{width:100%;}
  .search-box input{width:100%;}
  .container{padding:16px 14px 0;}
  .file-grid{grid-template-columns:1fr 1fr;}
  .lb-mask{padding:52px 10px 86px;}
  .lb-arrow{width:38px;height:38px;font-size:20px;}
  .lb-prev{left:6px;}.lb-next{right:6px;}
}
@media (max-width:480px){.file-grid:not(.tool-grid){grid-template-columns:1fr;}}
"""

# ---------------------------------------------------------------- 方案一
V1 = r"""
/* ============================================================
   方案一 · 晨雾杂志版
   亮色 / 每个分区一个主色 / 卡片悬浮上浮 / 彩色图标块
   ============================================================ */
:root{
  --bg:#eef2f7;--bg-img:none;
  --ink:#1b2430;--ink2:#5b6675;--ink3:#93a0b2;
  --panel:#fff;--panel2:#f2f5f9;--line:#e2e8f0;--line2:#eef1f6;
  --shadow:0 2px 8px rgba(27,36,48,.06);--shadow-hv:0 16px 30px rgba(27,36,48,.14);
  --radius:16px;
  --hdr-bg:linear-gradient(96deg,#3d5878 0%,#5b7396 52%,#7189ab 100%);
  --hdr-ink:#fff;--hdr-line:transparent;--hdr-blur:none;
  --logo-bg:rgba(255,255,255,.20);--logo-ink:#fff;--logo-line:rgba(255,255,255,.42);
  --input-bg:rgba(255,255,255,.16);--input-line:rgba(255,255,255,.30);--input-ink:#fff;--input-ph:rgba(255,255,255,.72);
  --input-bg-f:rgba(255,255,255,.26);--focus:#fff;--focus-ring:rgba(255,255,255,.22);
  --sc-default:#5b7396;--bar:#5b7396;
  /* emoji 图标块底色：淡青蓝（用户定稿，统一浅色，不跟分区主色走） */
  --icon-bg:#b1d9e5;--icon-bg-hv:#c4e4ed;--icon-line:rgba(47,99,118,.20);
  --c1:#c8952e;--c2:#4a78b5;--c3:#4e8a6b;--c4:#3d8f96;--c5:#7a6bb0;--c6:#8a94a6;
  --ns-hover:#f3f7fb;--ns-hover-ink:#3d5878;
  --modal-bg:#fff;--modal-ink:#1b2430;--modal-ink2:#48566b;--modal-head-bg:#fbfcfe;--modal-line:#e8edf4;
  --close-bg:#f0f3f7;--close-ink:#6b7280;--close-bg-hv:#e4e9f0;
  --footer-ink:#93a0b2;--footer-line:#e2e8f0;
  --bt-bg:#5b7396;--bt-hv:#4a5e7d;--bt-shadow:0 6px 18px rgba(91,115,150,.35);
  --mask-bg:rgba(16,24,40,.52);--link:#5b7396;
  --tag-hz-ink:#4a6b58;--tag-hz-bg:#e8f1ec;--tag-gp-ink:#5b6b8a;--tag-gp-bg:#eaeff7;
}
.header{box-shadow:0 6px 20px rgba(61,88,120,.24)}
.header-brand{font-size:20px;letter-spacing:.6px}
.section-title{font-size:17px;font-weight:800;letter-spacing:.3px;margin:30px 0 14px;padding-bottom:11px;border-bottom:1px solid var(--line)}
.section-title::before{width:5px;height:20px;border-radius:3px}
.section-title .badge{background:var(--sc);color:#fff;font-weight:700}
.file-grid{background:transparent;padding:0;margin:0 0 26px;border:none;box-shadow:none}
.file-card{background:var(--panel);border:1px solid var(--line);border-left:4px solid var(--sc);border-radius:var(--radius);padding:14px 16px;box-shadow:var(--shadow);transition:transform .2s,box-shadow .2s}
.file-card:hover{transform:translateY(-4px);box-shadow:var(--shadow-hv)}
.file-card .card-title{font-size:14.5px;font-weight:700;display:flex;align-items:center;gap:10px}
.file-card .card-title .icon{width:36px;height:36px;border-radius:11px;background:var(--icon-bg);
  border:1px solid var(--icon-line);color:var(--ink);display:flex;align-items:center;
  justify-content:center;font-size:18px;flex-shrink:0;transition:background .2s,border-color .2s}
.file-card:hover .card-title .icon{background:var(--icon-bg-hv);border-color:rgba(47,99,118,.34)}
.file-card .card-desc{font-size:12.5px;color:var(--ink2);margin-top:8px;line-height:1.6;min-height:2.4em}
.tool-grid{grid-template-columns:repeat(auto-fill,minmax(178px,1fr));gap:12px}
.tool-grid .file-card{padding:11px 13px}
.tool-grid .file-card .card-title{font-size:13.5px;gap:9px}
.tool-grid .file-card .card-title .icon{width:29px;height:29px;border-radius:9px;font-size:15px}
.tool-grid .file-card .card-desc{display:none}
@media (max-width:768px){.file-grid:not(.tool-grid){grid-template-columns:1fr 1fr;gap:10px;}}
"""

# ---------------------------------------------------------------- 方案二
V2 = r"""
/* ============================================================
   方案二 · 深空极光版
   深色底 / 极光光斑 / 玻璃拟态卡片 / 霓虹描边 / 发光分区条
   ============================================================ */
:root{
  --bg:#0b1020;
  --bg-img:radial-gradient(1100px 560px at 8% -12%,rgba(110,168,255,.20),transparent 62%),radial-gradient(900px 500px at 94% 0%,rgba(167,139,250,.18),transparent 60%);
  --ink:#e9eef8;--ink2:#a7b3c8;--ink3:#77869f;
  --panel:linear-gradient(180deg,rgba(255,255,255,.075),rgba(255,255,255,.032));
  --panel2:rgba(255,255,255,.05);--line:rgba(255,255,255,.10);--line2:rgba(255,255,255,.06);
  --shadow:0 8px 26px rgba(0,0,0,.32);--shadow-hv:0 20px 42px rgba(0,0,0,.5);
  --radius:16px;
  --hdr-bg:rgba(11,16,32,.78);--hdr-ink:#eaf0fb;--hdr-line:rgba(255,255,255,.08);--hdr-blur:blur(14px);
  --logo-bg:linear-gradient(135deg,#6ea8ff,#a78bfa);--logo-ink:#0b1020;--logo-line:transparent;
  --input-bg:rgba(255,255,255,.07);--input-line:rgba(255,255,255,.14);--input-ink:#e9eef8;--input-ph:rgba(233,238,248,.55);
  --input-bg-f:rgba(255,255,255,.12);--focus:#6ea8ff;--focus-ring:rgba(110,168,255,.25);
  --sc-default:#6ea8ff;--bar:#6ea8ff;
  --c1:#f0b64b;--c2:#6ea8ff;--c3:#5fd39a;--c4:#4fd0d6;--c5:#a78bfa;--c6:#8b96ac;
  --ns-hover:rgba(255,255,255,.07);--ns-hover-ink:#fff;
  --modal-bg:#111827;--modal-ink:#f1f5fb;--modal-ink2:#c3cee0;--modal-head-bg:rgba(255,255,255,.04);--modal-line:rgba(255,255,255,.10);
  --close-bg:rgba(255,255,255,.10);--close-ink:#cbd5e1;--close-bg-hv:rgba(255,255,255,.2);
  --footer-ink:#77869f;--footer-line:rgba(255,255,255,.08);
  --bt-bg:#6ea8ff;--bt-hv:#4b8ef0;--bt-shadow:0 8px 22px rgba(110,168,255,.4);
  --mask-bg:rgba(4,8,18,.74);--link:#7fb4ff;
  --tag-hz-ink:#7fe0b0;--tag-hz-bg:rgba(95,211,154,.14);--tag-gp-ink:#9cc2ff;--tag-gp-bg:rgba(110,168,255,.16);
}
body{background-attachment:fixed}
.header-brand{font-size:20px;letter-spacing:.6px}
.section-title{color:#fff;font-size:17px;font-weight:800;margin:30px 0 14px;letter-spacing:.4px}
.section-title::before{width:5px;height:19px;border-radius:3px;background:linear-gradient(180deg,var(--sc),rgba(255,255,255,.18));box-shadow:0 0 12px var(--sc)}
.section-title .badge{background:rgba(255,255,255,.07);color:var(--ink2)}
.file-grid{background:transparent;padding:0;margin:0 0 26px;border:none;box-shadow:none}
.file-card{position:relative;overflow:hidden;background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);padding:14px 16px;backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);transition:transform .2s,border-color .2s,box-shadow .2s}
.file-card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--sc);opacity:.9}
.file-card:hover{transform:translateY(-4px);border-color:var(--sc);box-shadow:var(--shadow-hv),0 0 0 1px var(--sc)}
.file-card .card-title{font-size:14.5px;font-weight:700;display:flex;align-items:center;gap:10px}
.file-card .card-title .icon{width:34px;height:34px;border-radius:11px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0}
.file-card:hover .card-title .icon{background:rgba(255,255,255,.14)}
.file-card .card-desc{font-size:12.5px;color:var(--ink2);margin-top:8px;line-height:1.6;min-height:2.4em}
.tool-grid{grid-template-columns:repeat(auto-fill,minmax(178px,1fr));gap:12px}
.tool-grid .file-card{padding:11px 13px}
.tool-grid .file-card .card-title{font-size:13.5px;gap:9px}
.tool-grid .file-card .card-title .icon{width:29px;height:29px;border-radius:9px;font-size:15px}
.tool-grid .file-card .card-desc{display:none}
.nc-slide{background:#141b2e}
.ns-thumb{border-color:rgba(255,255,255,.10)}
@media (max-width:768px){.file-grid:not(.tool-grid){grid-template-columns:1fr 1fr;gap:10px;}}
"""

# ---------------------------------------------------------------- 方案三
V3 = r"""
/* ============================================================
   方案三 · 极简控制台版
   白底网格 / 等宽编号分区 / 方形黑图标块 / 高密度 / hover 出箭头
   ============================================================ */
:root{
  --bg:#fafbfc;--bg-img:linear-gradient(#f0f3f7 1px,transparent 1px),linear-gradient(90deg,#f0f3f7 1px,transparent 1px);
  --ink:#111827;--ink2:#5b6570;--ink3:#9aa4b2;
  --panel:#fff;--panel2:#f4f6f9;--line:#e4e8ee;--line2:#eef1f5;
  --shadow:none;--shadow-hv:0 2px 10px rgba(17,24,39,.09);
  --radius:8px;
  --hdr-bg:rgba(255,255,255,.93);--hdr-ink:#111827;--hdr-line:#e4e8ee;--hdr-blur:blur(10px);
  --logo-bg:#111827;--logo-ink:#fff;--logo-line:transparent;
  --input-bg:#f4f6f9;--input-line:#e4e8ee;--input-ink:#111827;--input-ph:#9aa4b2;
  --input-bg-f:#fff;--focus:#111827;--focus-ring:rgba(17,24,39,.10);
  --sc-default:#111827;--bar:#111827;
  --c1:#111827;--c2:#111827;--c3:#111827;--c4:#111827;--c5:#111827;--c6:#111827;
  --ns-hover:#f4f6f9;--ns-hover-ink:#111827;
  --modal-bg:#fff;--modal-ink:#111827;--modal-ink2:#48566b;--modal-head-bg:#fafbfc;--modal-line:#e8ebf0;
  --close-bg:#f0f2f5;--close-ink:#6b7280;--close-bg-hv:#e4e8ee;
  --footer-ink:#9aa4b2;--footer-line:#e4e8ee;
  --bt-bg:#111827;--bt-hv:#000;--bt-shadow:0 6px 16px rgba(17,24,39,.28);
  --mask-bg:rgba(16,24,40,.48);--link:#111827;
  --tag-hz-ink:#4a6b58;--tag-hz-bg:#eef4f0;--tag-gp-ink:#5b6b8a;--tag-gp-bg:#eef1f7;
}
body{background-size:26px 26px}
.header-brand{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:17px;letter-spacing:.6px}
.brand-logo{border-radius:6px}
.search-box input{border-radius:6px;padding-left:34px}
#cardContainer{counter-reset:sec}
.section-title{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:14.5px;font-weight:700;letter-spacing:.7px;margin:24px 0 10px;border-left:3px solid var(--ink);padding-left:11px;gap:8px}
.section-title::before{content:counter(sec,decimal-leading-zero);counter-increment:sec;width:auto;height:auto;background:none;border-radius:0;color:var(--ink3);font-size:13px;font-weight:700;order:-1;margin-right:1px}
.section-title .badge{background:#f0f2f5;color:var(--ink2);font-family:inherit}
.file-grid{background:#fff;border:1px solid var(--line);border-radius:8px;padding:12px;margin:0 0 18px}
.file-card{position:relative;background:#fff;border:1px solid var(--line);border-radius:7px;padding:11px 13px;transition:border-color .16s,box-shadow .16s,transform .16s}
.file-card:hover{border-color:#111827;box-shadow:var(--shadow-hv);transform:translateY(-2px)}
.file-card::after{content:'↗';position:absolute;right:11px;bottom:8px;font-size:12px;color:#111827;opacity:0;transition:opacity .16s}
.file-card:hover::after{opacity:1}
.file-card .card-title{font-size:13.5px;font-weight:700;display:flex;align-items:center;gap:9px}
.file-card .card-title .icon{width:28px;height:28px;border-radius:6px;background:#111827;color:#fff;display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0}
.file-card .card-desc{font-size:12px;color:var(--ink2);margin-top:7px;line-height:1.55;min-height:2.2em}
.notice-carousel,.notice-scroll{border-radius:8px}
.tool-grid{grid-template-columns:repeat(auto-fill,minmax(148px,1fr));gap:8px;padding:10px}
.tool-grid .file-card{padding:9px 11px}
.tool-grid .file-card .card-title{font-size:12.5px;gap:8px}
.tool-grid .file-card .card-title .icon{width:24px;height:24px;border-radius:5px;font-size:13px}
.tool-grid .file-card .card-desc{display:none}
.tool-grid .file-card::after{display:none}
@media (max-width:768px){.file-grid:not(.tool-grid){grid-template-columns:1fr 1fr;gap:8px;padding:10px;}}
@media (max-width:480px){.tool-grid{grid-template-columns:repeat(2,minmax(0,1fr));}}
"""

THEMES = [
    ("theme-v1.html", "方案一 · 晨雾杂志版", V1),
    ("theme-v2.html", "方案二 · 深空极光版", V2),
    ("theme-v3.html", "方案三 · 极简控制台版", V3),
]


def apply_theme(key):
    """把选定主题套用到正式导航页 deploy_nav/index.html（不带样板提示条）。"""
    decor = dict((t[0].replace("theme-", "").replace(".html", ""), t[1:]) for t in THEMES)[key]
    label = decor[0]
    with io.open(SRC, "r", encoding="utf-8") as f:
        src = f.read()
    style_re = re.compile(r"<style>.*?</style>", re.S)
    title_re = re.compile(r"<title>.*?</title>", re.S)
    brand_re = re.compile(r'(<div class="header-brand">).*?(</div>)', re.S)
    out = style_re.sub("<style>\n" + COMMON.strip() + "\n" + decor[1].strip() + "\n</style>", src, count=1)
    out = title_re.sub("<title>%s</title>" % SITE_TITLE, out, count=1)
    out = brand_re.sub(r'\1<span class="brand-logo">杭</span> ' + SITE_TITLE + r"\2", out, count=1)
    with io.open(SRC, "w", encoding="utf-8", newline="") as f:
        f.write(out)
    print("applied:", label, "->", SRC)


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "apply":
        apply_theme(sys.argv[2])
        return

    with io.open(SRC, "r", encoding="utf-8") as f:
        src = f.read()

    style_re = re.compile(r"<style>.*?</style>", re.S)
    title_re = re.compile(r"<title>.*?</title>", re.S)
    if not style_re.search(src) or not title_re.search(src):
        raise SystemExit("index.html 结构异常：找不到 <style> 或 <title>")

    for fname, label, decor in THEMES:
        # 主题装饰在后：同特异性时覆盖公共骨架（如方案三的分区编号 ::before）
        css = COMMON.strip() + "\n" + decor.strip()
        out = style_re.sub("<style>\n" + css + "\n</style>", src, count=1)
        out = title_re.sub("<title>%s（%s）</title>" % (SITE_TITLE, label), out, count=1)
        # 顶部加一条样板提示条（仅样板，正式 index.html 不带）
        tip = ('<div class="theme-tip" style="position:sticky;top:0;z-index:300;'
               'background:#111827;color:#fff;font-size:12.5px;padding:7px 14px;'
               'text-align:center;letter-spacing:.4px">样板预览：%s　·　'
               '正式页面 index.html 暂未套用</div>' % label)
        out = out.replace('<header class="header">', tip + '\n<header class="header">', 1)
        path = os.path.join(BASE_DIR, "deploy_nav", fname)
        with io.open(path, "w", encoding="utf-8", newline="") as f:
            f.write(out)
        print("written:", fname, len(out), "chars ·", label)


if __name__ == "__main__":
    main()
