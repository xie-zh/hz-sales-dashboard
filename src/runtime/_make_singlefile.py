# -*- coding: utf-8 -*-
"""把演示站打包成「单文件离线版」。

产物：一个 .html 文件，内部内嵌
  - ECharts 库（本地副本，不走CDN）
  - 各页面 HTML
  - 各数据 JSON
用 iframe + srcdoc 承载页面，父页面用 blob/data 提供资源。

特点：**双击即可打开，不依赖网络、不依赖服务器、不依赖任何本地文件**。
适合直接发给面试官，或放进任何网盘/邮件/二维码。
"""
import os, re, json, base64, shutil

SRC = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\_publish_github"
OUT = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\_portfolio_site"
DEMO = os.path.join(SRC, "demo")

PAGES = [
    ("index.html", "活动量看板", "核心看板：KPI 总览 · 部门过程量 · 全员明细（总表/法律/会计）· 图表 · 历史趋势"),
    ("sanli.html", "员工三力", "人力结构四层分级：英雄人力 / 标准人力 / 破零人力 / 无效人力"),
    ("ry.html", "网销数据分析", "网销漏斗转化分析，逐人逐环节"),
    ("ry_dy.html", "电邀数据分析", "电话邀约过程量追踪"),
    ("ry_kuaiji.html", "会计数据分析", "财务侧活动量考核 · 票据抽检与个人出错率"),
]

ACCENT = "#5b7396"
GREEN = "#5e8a70"
AMBER = "#bf9350"
BLUE = "#5b7ea6"
BRICK = "#b05a52"
COLORS = [ACCENT, GREEN, BLUE, AMBER, BRICK]


def read(p):
    with open(p, "r", encoding="utf-8") as f:
        return f.read()


def main():
    os.makedirs(OUT, exist_ok=True)

    # 1. ECharts（本地，不走 CDN）
    echarts = read(os.path.join(DEMO, "static", "echarts.min.js"))
    print("ECharts: %.0f KB" % (len(echarts) / 1024))

    # 2. 各页面的数据 JSON（含 hist/ 下的历史快照，按文件名索引）
    data_files = {}
    for f in os.listdir(DEMO):
        if f.endswith(".json") and f != "workspace_data.json":
            data_files[f] = read(os.path.join(DEMO, f))
    hist_dir = os.path.join(DEMO, "hist")
    n_hist = 0
    if os.path.isdir(hist_dir):
        for dp, dn, fn in os.walk(hist_dir):
            for f in fn:
                if f.endswith(".json"):
                    key = "hist/" + f
                    if key not in data_files:
                        data_files[key] = read(os.path.join(dp, f))
                        n_hist += 1
    print("数据文件: %d 个（其中历史快照 %d）" % (len(data_files), n_hist))

    # 3. 页面 HTML：把 /static/ 引用改掉，注入数据与echarts
    pages = {}
    for fn, title, desc in PAGES:
        h = read(os.path.join(DEMO, fn))
        # echarts 由父页面注入（iframe 内用srcdoc，无法访问父级 script）
        h = h.replace('<script src="static/echarts.min.js"></script>',
                      "<script>/*ECHARTS_PLACEHOLDER*/</script>")
        # fetch 相对路径 → 父页面注入数据后已不需要（页面本就有内嵌兜底）
        pages[fn] = (h, title, desc)

    # 4. 打包
    # ⚠ 关键：payload 里全是 HTML/JS，必然含 `</script>`；
    #   直接塞进 <script> 会提前闭合脚本块导致整页白屏 ⇒ 必须转义 `<`。
    payload = {
        "echarts": echarts,
        "data": data_files,
        "pages": {k: v[0] for k, v in pages.items()},
    }
    blob = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")

    # 5. 生成单文件 HTML
    cards = ""
    for i, (fn, title, desc) in enumerate(PAGES):
        cards += (
            '<button class="card" data-p="%s" style="border-left-color:%s">'
            '<h3>%s</h3><p>%s</p></button>\n'
            % (fn, COLORS[i % len(COLORS)], title, desc)
        )

    html = HTML_TPL.replace("/*__PAYLOAD__*/", "window.__P=" + blob + ";")
    html = html.replace("<!--CARDS-->", cards)
    dst = os.path.join(OUT, "销售看板作品集_单文件离线版.html")
    with open(dst, "w", encoding="utf-8", newline="\n") as f:
        f.write(html)
    size = os.path.getsize(dst) / 1048576
    print("\n单文件离线版已生成: %s (%.1f MB)" % (dst, size))
    print("双击即可打开，不需网络 / 服务器 / 本地文件")


HTML_TPL = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>销售团队活动量数据看板系统 · 作品集（离线单文件版）</title>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:"PingFang SC","Microsoft YaHei",system-ui,sans-serif;
     background:#f4f6f9;color:#1f2937;line-height:1.7;height:100vh;display:flex;flex-direction:column}
.top{background:linear-gradient(135deg,#1e293b,#334155);color:#e2e8f0;padding:14px 18px;flex:0 0 auto}
.top h1{margin:0;font-size:17px;letter-spacing:.4px}
.top .sub{margin:3px 0 0;font-size:12.5px;opacity:.85}
.warn{background:#fffbeb;color:#92400e;padding:7px 18px;font-size:12.5px;
      border-bottom:1px solid #fcd34d;flex:0 0 auto}
.bar{display:flex;gap:8px;padding:11px 18px;overflow-x:auto;background:#fff;
     border-bottom:1px solid #e2e8f0;flex:0 0 auto}
.card{flex:0 0 auto;min-width:186px;background:#fff;border:1px solid #e2e8f0;
      border-left:4px solid #5b7396;border-radius:9px;padding:9px 13px;cursor:pointer;
      text-align:left;font-family:inherit;transition:.15s}
.card:hover{box-shadow:0 4px 14px rgba(0,0,0,.09);transform:translateY(-1px)}
.card.on{background:#eef2f7;border-color:#94a3b8}
.card h3{margin:0 0 2px;font-size:13.5px;color:#1f2937}
.card p{margin:0;font-size:11.5px;color:#64748b;line-height:1.5}
.frame{flex:1 1 auto;border:0;width:100%;background:#fff;min-height:0}
.hint{padding:8px 18px;font-size:11.5px;color:#94a3b8;border-top:1px solid #e2e8f0;flex:0 0 auto}
</style>
</head>
<body>
<div class="top">
  <h1>销售团队活动量数据看板系统</h1>
  <p class="sub">从线下 Excel 到线上自动看板的完整工程实践 · 5 个数据板块 · 全自动更新流水线 · 1.6 万行 Python + 原生JS</p>
</div>
<div class="warn">⚠ 演示环境：全部人名、工号、组织名与数值均为虚构，仅用于展示系统功能与界面效果。</div>
<div class="bar">
<!--CARDS-->
<button class="card" id="btnHome" style="border-left-color:#1f2937">
  <h3>🏠 项目总览</h3><p>技术架构 · 数据流水线 · 工程难点</p>
</button>
</div>
<iframe class="frame" id="frame" title="看板预览"></iframe>
<div class="hint">全部内容已内嵌于本文件 —— 无需网络、无需服务器、无需本地依赖，可离线打开。</div>

<script>
/*__PAYLOAD__*/
var P=window.__P;
var fr=document.getElementById('frame');
var cur=null;

/* 把 fetch 拦截掉，用内嵌数据应答 —— 让页面在 srcdoc 环境里正常读数据 */
function buildDoc(page){
  var html=P.pages[page]||P.pages['index.html'];
  html=html.replace('/*ECHARTS_PLACEHOLDER*/',P.echarts);
  /* 用一个完整的 Response 形状 mock fetch —— 页面会读 ok/status/headers，
     缺字段会在 .then 里抛 undefined。*/
  var inject=
    '<script>'+
    'window.__D='+JSON.stringify(P.data).replace(/</g,"\\u003c")+';'+
    'function __key(u){var p=String(u).split("?")[0];var base=p.split("/").pop();'+
      'if(Object.prototype.hasOwnProperty.call(window.__D,p))return p;'+
      'if(Object.prototype.hasOwnProperty.call(window.__D,base))return base;'+
      'var m=p.match(/(?:^|\\/)hist\\/.+$/);if(m&&Object.prototype.hasOwnProperty.call(window.__D,m[0]))return m[0];'+
      'return null;}'+
    'var __of=window.fetch;window.fetch=function(u,o){'+
      'var k=__key(u);if(k){var t=window.__D[k];'+
        'return Promise.resolve(new Response(t,'+
        '{status:200,statusText:"OK",headers:{"Content-Type":"application/json"}}));}'+
      'return __of?__of.apply(window,arguments):Promise.reject(new Error("offline:"+u));};'+
    '<'+'/script>';
  /* 把注入脚本放到 </head> 前，确保在页面自身 script 之前执行 */
  html=html.replace('</head>', inject+'\n</head>');
  return html;
}

function open(page,btn){
  cur=page;
  document.querySelectorAll('.card').forEach(function(c){c.classList.remove('on')});
  if(btn)btn.classList.add('on');
  fr.srcdoc=buildDoc(page);
}

document.querySelectorAll('.card[data-p]').forEach(function(b){
  b.addEventListener('click',function(){open(b.dataset.p,b)});
});
document.getElementById('btnHome').addEventListener('click',function(){
  open('index.html',this);
});
open('index.html',document.querySelector('.card[data-p]'));
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()