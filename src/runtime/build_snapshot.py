# -*- coding: utf-8 -*-
"""
生成云端看板快照 → _deploy/
================================
部署到 CloudStudio 后得到稳定链接，数据更新后重新部署，
链接不变，用户刷新即看到最新数据。

快照包含：
  index.html        活动量看板（动态读取 data.json，含内嵌数据兜底 + 顶部导航）
  data.json         活动量数据（从 分公司活动量-在线看板.xlsx 解析）
  sanli.html        员工三力看板（动态读取 sanli_data.json，含内嵌数据兜底 + 顶部导航）
  sanli_data.json   三力数据（从 分公司-员工三力明细.xlsx 解析）
  ry.html           网销数据分析看板（动态读取 ry_data.json，含内嵌数据兜底 + 顶部导航）
  ry_data.json      网销人员数据（从 分公司-人员数据分析路径.xlsx 解析：网销漏斗+LTC）
  static/echarts.min.js
"""
import json, os, shutil, time, hashlib, csv
import dashboard_server as ds
import sanli_parser as sp
import ry_parser as rp
import hist_snapshot  # 历史快照：每次构建自动留存当日四模块数据
import yj_builder     # 业绩数据板块（业绩含续费 + 订单商品明细）

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEPLOY_DIR = os.path.join(BASE_DIR, "_deploy")
INDEX_PATH = os.path.join(BASE_DIR, "index.html")
SANLI_PATH = os.path.join(BASE_DIR, "sanli.html")
RY_PATH = os.path.join(BASE_DIR, "ry.html")
YJ_PATH = os.path.join(BASE_DIR, "yj.html")
SKU_PATH = os.path.join(BASE_DIR, "sku.html")

NAV_CSS = (
    "/* board nav */\n"
    ".board-nav{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}\n"
    ".board-nav a{display:inline-flex;align-items:center;gap:6px;padding:8px 18px;border-radius:8px;border:1px solid var(--border);background:var(--card);color:var(--muted);text-decoration:none;font-size:14px;font-weight:600;transition:all .15s}\n"
    ".board-nav a:hover{background:var(--accent);color:#fff;border-color:var(--accent)}\n"
    ".board-nav a.act{background:var(--accent);color:#fff;border-color:var(--accent);box-shadow:0 2px 6px rgba(79,70,229,.3)}\n"
)

NAV_LINKS = [
    ("index.html", "📊 活动量看板"),
    ("sanli.html", "🏆 员工三力看板"),
    ("ry.html", "🧭 网销数据分析"),
    ("ry_dy.html", "📞 电邀数据分析"),
    ("ry_kuaiji.html", "📒 会计数据分析"),
    ("yj.html", "💰 业绩含续费"),
    ("sku.html", "🛒 订单商品明细"),
]

# 导航缓存破坏参数：每次 build() 开头写入「?cb=<unix时间戳>」，让首页跳转到四个
# 模块的链接每次构建都是新 URL，浏览器不会命中旧缓存 HTML（平台静态托管不返回
# Cache-Control，启发式缓存会把旧页缓存并回放，导致「部署了新版却看不到」）。
# 用户只要从任意带 ?cb= 的链接进入一次，之后点导航都是新 URL，彻底跳出旧缓存。
NAV_CB = ""


def _nav_html(active):
    """生成五入口导航条；active=当前页 href。"""
    items = []
    for href, label in NAV_LINKS:
        h = href + NAV_CB
        cls = ' class="act"' if href == active else ''
        items.append('  <a href="%s"%s>%s</a>' % (h, cls, label))
    return '<nav class="board-nav">\n' + "\n".join(items) + "\n</nav>\n"


def _inject_nav(html, nav_html):
    """注入顶部导航条（在 .container 之后）"""
    if 'class="board-nav"' not in html:
        html = html.replace(
            '<div class="container">',
            '<div class="container">\n' + nav_html,
            1,
        )
    else:
        # 替换已有导航为三看板导航（保持 act 类由 nav_html 决定）
        import re
        html = re.sub(
            r'<nav class="board-nav">.*?</nav>',
            nav_html.replace('\\', '\\\\').replace('</nav>', '</nav>'),
            html,
            count=1,
            flags=re.S,
        )
    return html


def _ver_fetch(html, json_name, ver):
    """给数据 fetch 加「数据版本号」查询参数。

    背景（2026-09-14）：曾出现「新 HTML + CDN/浏览器缓存的旧 JSON」被混用，
    导致依赖新字段（dept_charts.km8_dept）的「🎯 8 大重点指标组合图」整块空白。
    版本号只由数据内容决定 → 数据未变则 URL 不变（仍可缓存），数据一变 URL 即变。

    兼容两种写法：
      - 源文件直接 fetch('json_name')（index / 旧写法）
      - 历史选择器引入的 fetch(LIVE_URL)（ry.html / sanli.html 模板，派生页由
        _derive_ry_page 把 LIVE_URL 换成各自数据文件）→ 注入为 fetch(LIVE_URL+'?v=ver')
    """
    old = "fetch('%s')" % json_name
    new = "fetch('%s?v=%s')" % (json_name, ver)
    if old in html:
        return html.replace(old, new)
    alt = "fetch(LIVE_URL)"
    alt_new = "fetch(LIVE_URL+'?v=%s')" % ver
    if alt in html:
        return html.replace(alt, alt_new)
    print("  [warn] 未找到数据 fetch 锚点:", json_name)
    return html


# --------------------------------------------------------------------------
# 数据过期警示条（配合统一编排器 pipeline_sync.py）
# 编排器每轮同步后写 _deploy/status.json（stale / last_success / last_error），
# 页面加载时拉一次：stale=true 就在顶部挂红条，明确告诉看板的人
# 「当前数据是上一轮成功的，本轮同步失败了」，而不是让人误以为是实时数据。
# --------------------------------------------------------------------------
STATUS_BANNER_JS = r"""(function () {
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c];
    });
  }
  function show(s) {
    if (!s || !s.stale) return;
    var b = document.createElement('div');
    b.id = 'stale-banner';
    b.style.cssText = 'position:fixed;left:0;right:0;top:0;z-index:99999;' +
      'background:#b05a52;color:#fff;padding:7px 14px;text-align:center;' +
      'font:13px/1.7 -apple-system,"Microsoft YaHei",sans-serif;' +
      'box-shadow:0 2px 10px rgba(0,0,0,.28)';
    b.innerHTML = '⚠ 数据未更新：最近一次同步失败' +
      (s.last_error ? '（' + esc(s.last_error) + '）' : '') +
      ' · 当前显示 ' + esc(s.last_success || '上一轮') + ' 的数据' +
      (s.updated_at ? ' · 检查于 ' + esc(s.updated_at) : '');
    document.body.appendChild(b);
    document.body.style.paddingTop = (b.offsetHeight + 4) + 'px';
  }
  function load() {
    fetch('status.json?cb=' + Date.now(), { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(show).catch(function () {});
  }
  if (document.readyState !== 'loading') load();
  else document.addEventListener('DOMContentLoaded', load);
})();
"""

_BANNER_TAG = '<script src="status_banner.js"></script>'


def _inject_banner(html):
    """给页面挂上 status_banner.js（幂等）。"""
    if 'status_banner.js' in html:
        return html
    if '</body>' in html:
        return html.replace('</body>', _BANNER_TAG + '\n</body>', 1)
    return html + '\n' + _BANNER_TAG


# 防缓存：平台静态托管不返回 Cache-Control / ETag（只有 Last-Modified），
# 浏览器会按「启发式缓存」把 HTML 直接缓存并在有效期内**不回源校验** →
# 部署了新版本，用户普通刷新仍看到旧页面（典型症状："明明改了，怎么还是旧的"）。
# 显式声明 no-cache 后，浏览器每次都必须回源校验（304 很便宜）。
NO_CACHE_META = (
    "<!-- 防缓存（build_snapshot 注入）：看板更新后普通刷新即可看到最新版本 -->\n"
    '<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">\n'
    '<meta http-equiv="Pragma" content="no-cache">\n'
    '<meta http-equiv="Expires" content="0">\n'
)

# 页头 meta 行的统一锚点（index.html / sanli.html / ry.html 三份源文件都有）
_META_ANCHOR = ('<span style="font-size:11px;color:var(--muted)" '
                'id="meta-status">加载中...</span>')


def _no_cache(html):
    """给 <head> 注入 no-cache 指令（幂等）。"""
    if "must-revalidate" in html:
        return html
    if "<head>" not in html:
        print("  [warn] 未找到 <head>，防缓存指令未注入")
        return html
    return html.replace("<head>", "<head>\n" + NO_CACHE_META, 1)


def _stamp(html, ver, at):
    """在页头 meta 行后插入「版本号 · 构建时间」徽标（幂等）。

    作用：用户/排查者一眼就能看出当前页面是哪一次构建，
    不用再靠"感觉还是旧的"来判断刷新是否生效。
    """
    if 'class="build-badge"' in html:
        return html
    anchor = _META_ANCHOR
    badge = (anchor
             + '\n    <span class="build-badge" style="font-size:11px;color:#94a3b8">'
               '版本 <b>%s</b> · %s</span>' % (ver, at))
    if anchor not in html:
        print("  [warn] 未找到版本号锚点（meta-status）")
        return html
    return html.replace(anchor, badge, 1)


def _embed(html, json_str):
    """注入内嵌数据兜底脚本（id=embedded-data）"""
    embedded = (
        '<script id="embedded-data" type="application/json">'
        + json_str.replace("</", "<\\/")
        + "</script>"
    )
    return html.replace("</head>", embedded + "\n</head>")


def _write_text(path, content):
    """带重试的写入：先删后写（释放可能的句柄锁），失败重试几次再放弃。"""
    for _attempt in range(8):
        try:
            try:
                os.remove(path)  # 释放部署工具/安全删除遗留的句柄锁
            except OSError:
                pass
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            return
        except PermissionError:
            time.sleep(0.8)
    raise PermissionError("写入失败(多次重试): " + path)


def _copy_file(src, dst):
    for _attempt in range(8):
        try:
            try:
                os.remove(dst)
            except OSError:
                pass
            shutil.copy(src, dst)
            return
        except PermissionError:
            time.sleep(0.8)
    raise PermissionError("复制失败(多次重试): " + dst)


def scope_ry(ry, dept, funnel_is_dy):
    """从完整 ry 数据中提取单一部门，生成独立看板数据（各自未来可接独立数据源）。"""
    emp = [e for e in ry["employees"] if e.get("部门标准") == dept]
    bus = [b for b in ry["bus"] if b.get("部门") == dept]
    total_sanli = {"英雄人力": 0, "标准人力": 0, "破零人力": 0, "未破零": 0, "无三力": 0, "其他": 0}
    total_real = 0.0
    total_check = 0.0
    total_today = 0.0
    for e in emp:
        sl = e.get("三力达标")
        if sl in total_sanli:
            total_sanli[sl] += 1
        else:
            total_sanli["其他"] += 1
        total_real += e.get("月累计实算利润业绩") or 0
        total_check += e.get("月累计核算利润业绩") or 0
        total_today += e.get("当日新增业绩") or 0
    colors = ry["meta"].get("dept_colors") or {}
    color = colors.get(dept, "#2563eb")
    meta = dict(ry["meta"])
    meta["sheet_title"] = "%s数据分析-%s" % (dept, ry["meta"].get("data_time", ""))
    meta["dept_order"] = [dept]
    meta["dept_colors"] = {dept: color}
    meta["dy_count"] = len(emp) if funnel_is_dy else 0
    # 重点指标：按派生页对应模块注入；主看板由 build() 单独处理（meta.key_module="net"）
    if dept == "电邀":
        meta["key_module"] = "dy"
        scoped_km = {"dy": ry["key_metrics"]["dy"]}
    elif dept == "会计":
        meta["key_module"] = "kuaiji"
        scoped_km = {"kuaiji": ry["key_metrics"]["kuaiji"]}
    else:
        meta["key_module"] = "net"
        scoped_km = {"net": ry["key_metrics"]["net"]}
    return {
        "meta": meta,
        "total": len(emp),
        "departments": [dept],
        "summary": {
            "三力": total_sanli,
            "月实算业绩": total_real,
            "月核算业绩": total_check,
            "当日新增业绩": total_today,
        },
        "bus": bus,
        "employees": emp,
        "funnel_metrics": ry["dy_funnel_metrics"] if funnel_is_dy else ry["funnel_metrics"],
        "dy_funnel_metrics": ry["dy_funnel_metrics"],
        "ltc_metrics": ry["ltc_metrics"],
        "key_metrics": scoped_km,
        "dept_charts": ry.get("dept_charts"),
    }


def _filter_net_ry(ry):
    """人员数据分析主看板：剔除 会计 / 电邀 部门的数据与人员。

    二者各有独立看板（ry_dy.html / ry_kuaiji.html），由 scope_ry 从「原始 ry」派生，
    不受本函数影响。本函数只产出主看板用的净数据副本（人员/业务表/部门图表/
    公司级汇总/部门列表均排除 会计、电邀），不修改传入的 ry。
    """
    hide = {"会计", "电邀"}
    emp = [e for e in ry.get("employees", []) if e.get("部门标准") not in hide]
    bus = [b for b in ry.get("bus", []) if b.get("部门") not in hide]
    depts = [d for d in ry.get("departments", []) if d not in hide]
    # 元信息部门顺序
    meta = dict(ry.get("meta", {}))
    if "dept_order" in meta:
        meta["dept_order"] = [d for d in meta["dept_order"] if d not in hide]
    # 重新计算公司级汇总（仅含主看板部门）
    total_sanli = {"英雄人力": 0, "标准人力": 0, "破零人力": 0, "未破零": 0, "无三力": 0, "其他": 0}
    total_real = 0.0
    total_check = 0.0
    total_today = 0.0
    for e in emp:
        sl = e.get("三力达标")
        if sl in total_sanli:
            total_sanli[sl] += 1
        else:
            total_sanli["其他"] += 1
        total_real += e.get("月累计实算利润业绩") or 0
        total_check += e.get("月累计核算利润业绩") or 0
        total_today += e.get("当日新增业绩") or 0
    # 部门图表：各子项按部门名过滤（depts 列表 + 各 dict 子项）
    dc = ry.get("dept_charts") or {}
    fdc = {}
    for k, v in dc.items():
        if k == "depts":
            fdc[k] = [d for d in (v or []) if d not in hide]
        elif k == "health":
            h = dict(v or {})
            h["metrics"] = {d: m for d, m in (h.get("metrics") or {}).items() if d not in hide}
            h["summary"] = {d: m for d, m in (h.get("summary") or {}).items() if d not in hide}
            fdc[k] = h
        elif isinstance(v, dict):
            fdc[k] = {d: m for d, m in v.items() if d not in hide}
        else:
            fdc[k] = v
    out = dict(ry)
    out["meta"] = meta
    out["total"] = len(emp)
    out["departments"] = depts
    out["summary"] = {
        "三力": total_sanli,
        "月实算业绩": total_real,
        "月核算业绩": total_check,
        "当日新增业绩": total_today,
    }
    out["bus"] = bus
    out["employees"] = emp
    out["dept_charts"] = fdc
    return out


# ---------- 会计「活动量完成情况」（sheet「会计」，2026-10-15 新版考核表） ----------
def build_kj_activity(xlsx_path=None):
    """解析「活动量-在线看板.xlsx」sheet「会计」的活动量考核数据。

    表结构：r1/r2 多级表头（考核项 / 周考核），r3 列名，r4 起逐人数据。
    固定列（r3 无表头）：idx6=入职时间(天)、idx9=出勤、idx10=是否达标、idx20=考核类型。
    只读不写（该文件必须 COM 重算，绝不能用 openpyxl 改写）。
    返回 dict；sheet 不存在 / 无数据 / 结构不符时返回 None（前端板块自动隐藏）。
    """
    from openpyxl import load_workbook
    path = xlsx_path or os.path.join(BASE_DIR, "分公司活动量-在线看板.xlsx")
    if not os.path.exists(path):
        return None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    if "会计" not in wb.sheetnames:
        wb.close()
        return None
    try:
        ws = wb["会计"]
        rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 400),
                                 values_only=True))
    finally:
        wb.close()
    if len(rows) < 4:
        return None
    hdr = [(str(v).strip() if v is not None else "") for v in rows[2]]

    def col(name):
        for i, h in enumerate(hdr):
            if h == name:
                return i
        return None

    c = {k: col(k) for k in ("部门", "岗位", "工号", "姓名", "见客量", "全民营销",
                             "当日全部通时(分钟)", "当日新增业绩", "周活动量",
                             "当日上门量", "当日外出量", "见客合计", "可抵扣")}
    if any(c[k] is None for k in ("部门", "岗位", "工号", "姓名", "周活动量")):
        return None

    def num(v):
        try:
            f = float(v)
            return int(f) if f == int(f) else round(f, 2)
        except (TypeError, ValueError):
            return 0

    out = []
    for r in rows[3:]:
        if c["姓名"] >= len(r):
            continue
        name = r[c["姓名"]]
        if name is None or not str(name).strip():
            continue

        def g(k):
            i = c[k]
            return r[i] if (i is not None and i < len(r)) else None

        def fixed(idx):
            return r[idx] if idx < len(r) else None

        out.append({
            "部门": str(g("部门") or "").strip(),
            "岗位": str(g("岗位") or "").strip(),
            "工号": str(g("工号") or "").strip(),
            "姓名": str(name).strip(),
            "入职天数": num(fixed(6)),
            "出勤": num(fixed(9)),
            "是否达标": str(fixed(10) or "").strip(),
            "见客量": num(g("见客量")),
            "全民营销": num(g("全民营销")),
            "通时": num(g("当日全部通时(分钟)")),
            "当日新增业绩": num(g("当日新增业绩")),
            "周活动量": str(g("周活动量") or "").strip(),
            "见客合计": num(g("见客合计")),
            "考核类型": str(fixed(20) or "").strip() or "正常考核",
        })
    if not out:
        return None
    summary = {
        "正常考核": sum(1 for x in out if x["考核类型"] == "正常考核"),
        "免考核": sum(1 for x in out if x["考核类型"] == "不考核活动量"),
        "周活动量合格": sum(1 for x in out
                        if x["考核类型"] == "正常考核" and x["周活动量"] == "合格"),
        "周活动量不合格": sum(1 for x in out
                          if x["考核类型"] == "正常考核" and x["周活动量"] == "不合格"),
    }
    return {
        "data_time": time.strftime("%Y-%m-%d %H:%M"),
        "total": len(out),
        "summary": summary,
        "rows": out,
    }


# ---------- 资质「活动量完成情况」（sheet「资质」，2026-10-22 新增） ----------
def build_zz_activity(xlsx_path=None):
    """解析「活动量-在线看板.xlsx」sheet「资质」的活动量考核数据。

    表结构：r1/r2 多级表头（基本信息(资质) / 入职时间 / 英雄人力 / 前半月英雄人力 /
    出勤 / 是否达标 / 考核项(当日全部通时+货架上架情况)），r3 列名，r4 起逐人数据。
    固定列（r3 无表头）：idx6=入职天数、idx7=英雄人力、idx8=前半月英雄人力、
    idx9=出勤、idx10=是否达标、idx15=上架个数。
    命名列：事业部/部门/岗位/工号/姓名/入职日期/目标/折扣目标/完成/完成率。
    只读不写（该文件必须 COM 重算，绝不能用 openpyxl 改写）。
    返回 dict；sheet 不存在 / 无数据 / 结构不符时返回 None（前端板块自动隐藏）。
    """
    from openpyxl import load_workbook
    path = xlsx_path or os.path.join(BASE_DIR, "分公司活动量-在线看板.xlsx")
    if not os.path.exists(path):
        return None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    if "资质" not in wb.sheetnames:
        wb.close()
        return None
    try:
        ws = wb["资质"]
        rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 400),
                                 values_only=True))
    finally:
        wb.close()
    if len(rows) < 4:
        return None
    hdr = [(str(v).strip() if v is not None else "") for v in rows[2]]

    def col(name):
        for i, h in enumerate(hdr):
            if h == name:
                return i
        return None

    c = {k: col(k) for k in ("事业部", "部门", "岗位", "工号", "姓名", "入职日期",
                             "目标", "折扣目标", "完成", "完成率")}
    if any(c[k] is None for k in ("部门", "工号", "姓名")):
        return None

    def num(v):
        try:
            f = float(v)
            return int(f) if f == int(f) else round(f, 2)
        except (TypeError, ValueError):
            return 0

    out = []
    for r in rows[3:]:
        if c["姓名"] >= len(r):
            continue
        name = r[c["姓名"]]
        if name is None or not str(name).strip():
            continue

        def g(k):
            i = c[k]
            return r[i] if (i is not None and i < len(r)) else None

        def fixed(idx):
            return r[idx] if idx < len(r) else None

        out.append({
            "事业部": str(g("事业部") or "").strip(),
            "部门": str(g("部门") or "").strip(),
            "岗位": str(g("岗位") or "").strip(),
            "工号": str(g("工号") or "").strip(),
            "姓名": str(name).strip(),
            "入职日期": str(g("入职日期") or "").strip()[:10],
            "入职天数": num(fixed(6)),
            "英雄人力": str(fixed(7) or "").strip(),
            "前半月英雄人力": str(fixed(8) or "").strip(),
            "出勤": num(fixed(9)),
            "是否达标": str(fixed(10) or "").strip(),
            "目标": num(g("目标")),
            "折扣目标": num(g("折扣目标")),
            "完成": num(g("完成")),
            "完成率": num(g("完成率")),
            "上架个数": num(fixed(15)),
        })
    if not out:
        return None
    summary = {
        "达标": sum(1 for x in out if x["是否达标"] == "达标"),
        "不达标": sum(1 for x in out if x["是否达标"] == "不达标"),
        "上架总数": sum(x["上架个数"] for x in out),
    }
    return {
        "data_time": time.strftime("%Y-%m-%d %H:%M"),
        "total": len(out),
        "summary": summary,
        "rows": out,
    }


def _derive_ry_page(src_html, fetch_json, name, nav_html, embed_str):
    """由 ry.html 源派生一个单部门看板页面（框架一致，数据源独立）。"""
    h = src_html
    # 数据文件指向该部门独立 JSON
    h = h.replace("fetch('ry_data.json')", "fetch('%s')" % fetch_json)
    # 历史快照目录也随数据文件切换（ry_dy_data.json -> hist/ry_dy，ry_kuaiji_data.json -> hist/ry_kuaiji）
    h = h.replace("LIVE_URL='ry_data.json'", "LIVE_URL='%s'" % fetch_json)
    # 标题 / 页头 / 导出文件名 改为部门名
    h = h.replace("分公司网销数据分析 · 看板", "分公司%s · 看板" % name)
    h = h.replace("分公司网销数据分析 · 在线看板", "分公司%s · 在线看板" % name)
    h = h.replace("网销数据分析_", "%s_" % name)
    # 注入导航 + 内嵌数据
    h = _inject_nav(h, nav_html)
    h = _embed(h, embed_str)
    return h


def _yj_prev(path):
    """读取上次生成的 JSON 文本（帆软新源取数失败时的兜底）。"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return '{"meta":{"data_time":"","notes":[],"week_ok":false},"company_names":[],"companies":{},"week_companies":null,"overview":[],"group":null,"rows":[]}'


def build():
    global NAV_CB
    # 导航缓存破坏：每次构建用新时间戳，使首页→四个模块的链接总是新 URL，
    # 规避平台静态托管无 Cache-Control 导致的旧 HTML 被浏览器永久缓存问题。
    NAV_CB = "?cb=%d" % int(time.time())
    # ---------- 活动量数据 ----------
    data = ds.build_meta()
    # 总表剔除交易事业部（2026-10-22）：该板块人员已挪至 sheet「资质」，
    # 由主看板「资质达标」Tab 单独跟踪，不再计入全员明细（总表），避免两处重复。
    # 注：部门过程量总览（数据统计）/ 法律 / 会计 / 资质 各板块不受影响。
    data["total"] = [r for r in (data.get("total") or [])
                     if r.get("事业部") != "交易事业部"]
    # 会计「活动量完成情况」：sheet「会计」逐人达标明细。
    # 总表 sheet 只含交易/跨境/工商/数字科技 22 人，会计不在其中 ⇒ 主看板必须单独挂载，
    # 否则会计员工在活动量看板上查不到自己的达标情况。
    _kj_act_main = build_kj_activity()
    if _kj_act_main:
        data["kj_activity"] = _kj_act_main
    # 资质「活动量完成情况」：sheet「资质」逐人达标明细（2026-10-22 新增）。
    # 资质板块人员不在总表 sheet ⇒ 主看板单独挂载，Tab 排在「会计达标」右侧。
    _zz_act_main = build_zz_activity()
    if _zz_act_main:
        data["zz_activity"] = _zz_act_main
    json_str = json.dumps(data, ensure_ascii=False)

    # ---------- 三力数据 ----------
    sanli = sp.build_meta()
    sanli_str = json.dumps(sanli, ensure_ascii=False)

    # ---------- 活动量看板 index.html ----------
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    # 1. 注入导航 CSS
    if ".board-nav" not in html:
        html = html.replace("</style>", NAV_CSS + "</style>")
    # 2. 注入顶部导航条（在 .container 之后）
    html = _inject_nav(html, _nav_html("index.html"))

    # 3. 内嵌数据兜底
    html = _embed(html, json_str)

    # 4. 修改 fetchData：优先请求 data.json，失败则用内嵌数据
    inject = (
        "var EMBEDDED_DATA=null;\n"
        "(function(){var el=document.getElementById('embedded-data');"
        "if(el){try{EMBEDDED_DATA=JSON.parse(el.textContent)}catch(e){}}})();\n"
    )
    html = html.replace(
        "function fetchData(cb){",
        inject
        + "function fetchData(cb){\n"
        + "  function applyData(d){RAW=d;d._loaded=Date.now();var fresh=d.meta&&d.meta.file_mtime_ts!==LAST_MTIME;if(fresh&&LAST_MTIME)toast('数据已刷新 ✓');if(d.meta&&d.meta.file_mtime_ts)LAST_MTIME=d.meta.file_mtime_ts;if(cb)cb();}\n"
        + "  fetch('data.json').then(function(r){return r.json()}).then(applyData).catch(function(e){if(EMBEDDED_DATA)applyData(EMBEDDED_DATA);else console.error(e)});\n"
        + "  return;\n"
        + "}"
    )
    # 移除原有的 fetch('/api/data') 实现（已被上方新函数替代）
    html = html.replace(
        "  fetch('/api/data').then(function(r){return r.json()}).then(function(d){\n"
        "    d._loaded=Date.now();\n"
        "    var fresh = d.meta.file_mtime_ts !== LAST_MTIME;\n"
        "    LAST_MTIME = d.meta.file_mtime_ts;\n"
        "    RAW = d;\n"
        "    if(fresh) toast('数据已刷新 ✅');\n"
        "    if(cb) cb();\n"
        "  }).catch(function(e){console.error(e)});\n"
        "}",
        "",
    )

    # ---------- 三力看板 sanli.html ----------
    with open(SANLI_PATH, "r", encoding="utf-8") as f:
        sanli_html = f.read()
    # 注入三看板导航（sanli 页为当前页）
    sanli_html = _inject_nav(sanli_html, _nav_html("sanli.html"))
    # sanli.html 已自带 fetch('sanli_data.json') 与 EMBEDDED_DATA 兜底逻辑，仅需注入内嵌数据
    sanli_html = _embed(sanli_html, sanli_str)

    # ---------- 网销数据分析 ry.html（全部门） + 电邀/会计 独立看板 ----------
    ry = rp.build_meta()
    ry["meta"]["key_module"] = "net"  # 主看板默认展示「网销模式」重点指标
    # 网销数据分析主看板：剔除 会计 / 电邀（二者各有独立看板 ry_dy.html / ry_kuaiji.html），
    # 仅隐藏主看板数据；电邀/会计独立看板由下方 scope_ry 从「原始 ry」派生、不受影响。
    ry_net = _filter_net_ry(ry)
    ry_str = json.dumps(ry_net, ensure_ascii=False)
    with open(RY_PATH, "r", encoding="utf-8") as f:
        ry_html_base = f.read()
    # 主看板：网销数据分析（含全部部门）
    ry_main_html = _inject_nav(ry_html_base, _nav_html("ry.html"))
    ry_main_html = _embed(ry_main_html, ry_str)
    # 电邀 / 会计 独立看板（各自数据源，框架与网销数据分析一致）
    ry_dy = scope_ry(ry, "电邀", True)
    # 会计明细：数据直接来自 sheet 名带「（会计）」的表（ry_parser 已构建）
    ry_kj = ry.get("kuaiji_detail") or scope_ry(ry, "会计", False)
    # 会计「活动量完成情况」：来自「活动量」xlsx 的「会计」sheet（2026-10-15 新增考核表）
    kj_act = build_kj_activity()
    if kj_act:
        ry_kj["kj_activity"] = kj_act
    ry_dy_str = json.dumps(ry_dy, ensure_ascii=False)
    ry_kj_str = json.dumps(ry_kj, ensure_ascii=False)
    ry_dy_html = _derive_ry_page(ry_html_base, "ry_dy_data.json", "电邀数据分析", _nav_html("ry_dy.html"), ry_dy_str)
    ry_kj_html = _derive_ry_page(ry_html_base, "ry_kuaiji_data.json", "会计数据分析", _nav_html("ry_kuaiji.html"), ry_kj_str)

    # ---------- 业绩数据板块（帆软新源：业绩含续费 + 订单商品明细） ----------
    # yj_builder 内置 15 分钟文件缓存；下载失败时回退用上次生成的 JSON，保证页面不空。
    yj_str = _yj_prev(os.path.join(DEPLOY_DIR, "yj_data.json"))
    try:
        _yj = yj_builder.build_yj_meta()
        if _yj:
            yj_str = json.dumps(_yj, ensure_ascii=False)
        else:
            print("  [warn] 业绩含续费数据为空，沿用上次 yj_data.json")
    except Exception as _e:
        print("  [warn] 业绩含续费取数失败（不影响主构建）:", _e)
    sku_str = _yj_prev(os.path.join(DEPLOY_DIR, "sku_data.json"))
    try:
        _sku = yj_builder.build_sku_meta()
        if _sku:
            sku_str = json.dumps(_sku, ensure_ascii=False)
        else:
            print("  [warn] 订单商品明细数据为空，沿用上次 sku_data.json")
    except Exception as _e:
        print("  [warn] 订单商品明细取数失败（不影响主构建）:", _e)
    with open(YJ_PATH, "r", encoding="utf-8") as f:
        yj_html = f.read()
    yj_html = _inject_nav(yj_html, _nav_html("yj.html"))
    yj_html = _embed(yj_html, yj_str)
    with open(SKU_PATH, "r", encoding="utf-8") as f:
        sku_html = f.read()
    sku_html = _inject_nav(sku_html, _nav_html("sku.html"))
    sku_html = _embed(sku_html, sku_str)

    # ---------- 写入部署目录 ----------
    # 重建前保留 MiniQA 的全工作空间数据摘要，避免被一并清掉
    _ws_backup = None
    _ws_path = os.path.join(DEPLOY_DIR, "workspace_data.json")
    if os.path.exists(_ws_path):
        try:
            with open(_ws_path, encoding="utf-8") as _f:
                _ws_backup = _f.read()
        except Exception:
            _ws_backup = None
    # 注意：不再 rmtree 整目录（沙盒安全删除钩子会把目录置为"待删除+锁句柄"状态，
    # 导致随后同名文件新建/写入持续 PermissionError）；改为逐文件覆盖。
    os.makedirs(DEPLOY_DIR, exist_ok=True)
    os.makedirs(os.path.join(DEPLOY_DIR, "static"), exist_ok=True)
    if _ws_backup is not None:
        try:
            _write_text(_ws_path, _ws_backup)
        except Exception:
            pass

    # ---------- 部署指纹（供自动化判断"内容是否变化"，与文件 mtime 无关） ----------
    # 仅纳入确定性输出（ry 系数据 + 派生 HTML）；活动量/三力含集合序不确定性，
    # 其变化由自动化里的 file_mtime_ts 比对覆盖。
    # ⚠️ 必须在「注入 ?v= 数据版本号」与「版本徽标」之前计算：
    #    活动量/三力 JSON 跨进程不稳定（PYTHONHASHSEED），若其版本号/时间戳进入指纹，
    #    指纹会每次构建都变，"无变化则跳过"的判断就永远失效（2026-09-14 实测踩坑）。
    _fp = "|".join([ry_str, ry_dy_str, ry_kj_str,
                    ry_main_html, ry_dy_html, ry_kj_html,
                    yj_str, sku_str, yj_html, sku_html])
    build_hash = hashlib.md5(_fp.encode("utf-8")).hexdigest()[:16]
    _built_at = time.strftime("%Y-%m-%d %H:%M:%S")

    # ---------- 数据版本号（写进各页 fetch URL，防「新 HTML + 旧缓存 JSON」混用） ----------
    # 版本号取「本页自己的数据」，且优先确定性来源；活动量/三力 JSON 不稳定，
    # 只会影响各自页面的 URL（index/sanli 不参与 build_hash，故不影响"跳过"判断）。
    html = _ver_fetch(html, "data.json",
                      hashlib.md5(json_str.encode("utf-8")).hexdigest()[:10])
    sanli_html = _ver_fetch(sanli_html, "sanli_data.json",
                            hashlib.md5(sanli_str.encode("utf-8")).hexdigest()[:10])
    ry_main_html = _ver_fetch(ry_main_html, "ry_data.json",
                              hashlib.md5(ry_str.encode("utf-8")).hexdigest()[:10])
    ry_dy_html = _ver_fetch(ry_dy_html, "ry_dy_data.json",
                            hashlib.md5(ry_dy_str.encode("utf-8")).hexdigest()[:10])
    ry_kj_html = _ver_fetch(ry_kj_html, "ry_kuaiji_data.json",
                            hashlib.md5(ry_kj_str.encode("utf-8")).hexdigest()[:10])
    yj_html = _ver_fetch(yj_html, "yj_data.json",
                         hashlib.md5(yj_str.encode("utf-8")).hexdigest()[:10])
    sku_html = _ver_fetch(sku_html, "sku_data.json",
                          hashlib.md5(sku_str.encode("utf-8")).hexdigest()[:10])

    # ---------- 防缓存 + 页头版本徽标（09-14） ----------
    # 平台不返回 Cache-Control，浏览器启发式缓存会吞掉新版本 → 统一注入 no-cache。
    _sv = build_hash[:8]
    html = _stamp(_no_cache(html), _sv, _built_at)
    sanli_html = _stamp(_no_cache(sanli_html), _sv, _built_at)
    ry_main_html = _stamp(_no_cache(ry_main_html), _sv, _built_at)
    ry_dy_html = _stamp(_no_cache(ry_dy_html), _sv, _built_at)
    ry_kj_html = _stamp(_no_cache(ry_kj_html), _sv, _built_at)
    yj_html = _stamp(_no_cache(yj_html), _sv, _built_at)
    sku_html = _stamp(_no_cache(sku_html), _sv, _built_at)

    # 数据过期警示条：js 落地 + 5 个页面注入（幂等）
    _write_text(os.path.join(DEPLOY_DIR, "status_banner.js"), STATUS_BANNER_JS)
    _write_text(os.path.join(DEPLOY_DIR, "index.html"), _inject_banner(html))
    _write_text(os.path.join(DEPLOY_DIR, "sanli.html"), _inject_banner(sanli_html))
    _write_text(os.path.join(DEPLOY_DIR, "ry.html"), _inject_banner(ry_main_html))
    _write_text(os.path.join(DEPLOY_DIR, "ry_dy.html"), _inject_banner(ry_dy_html))
    _write_text(os.path.join(DEPLOY_DIR, "ry_kuaiji.html"), _inject_banner(ry_kj_html))
    _write_text(os.path.join(DEPLOY_DIR, "yj.html"), _inject_banner(yj_html))
    _write_text(os.path.join(DEPLOY_DIR, "sku.html"), _inject_banner(sku_html))

    _copy_file(os.path.join(BASE_DIR, "echarts.min.js"),
               os.path.join(DEPLOY_DIR, "static", "echarts.min.js"))

    _write_text(os.path.join(DEPLOY_DIR, "data.json"), json_str)
    _write_text(os.path.join(DEPLOY_DIR, "sanli_data.json"), sanli_str)
    _write_text(os.path.join(DEPLOY_DIR, "ry_data.json"), ry_str)
    _write_text(os.path.join(DEPLOY_DIR, "ry_dy_data.json"), ry_dy_str)
    _write_text(os.path.join(DEPLOY_DIR, "ry_kuaiji_data.json"), ry_kj_str)
    _write_text(os.path.join(DEPLOY_DIR, "yj_data.json"), yj_str)
    _write_text(os.path.join(DEPLOY_DIR, "sku_data.json"), sku_str)

    # ---------- 历史快照：把当日 4 个模块数据按日期存档（供历史时间选择器） ----------
    # 注意：hist_snapshot.backfill 补录历史时会临时把源表换成历史版本再调 build()，
    #       此时 SUPPRESS_AUTO_SNAPSHOT=True，必须跳过自动快照，否则「当天」会被写脏。
    try:
        if getattr(hist_snapshot, "SUPPRESS_AUTO_SNAPSHOT", False):
            print("  [skip] 历史快照自动写入已抑制（backfill 模式）")
        else:
            hist_snapshot.snapshot_now()
    except Exception as _e:
        print("  [warn] 历史快照写入失败（不影响主构建）:", _e)

    _write_text(os.path.join(DEPLOY_DIR, "_build_fingerprint.json"),
                json.dumps({"build_hash": build_hash, "built_at": _built_at},
                           ensure_ascii=False))
    print("build_hash:", build_hash, "| built_at:", _built_at)

    # ---------- 历史数据快照（history_builder 聚合 JSON 日快照 → history.json） ----------
    # 2026-09-24 起改为 JSON 日快照方案：每天 23:59 由 snapshot_daily.py 拍当日快照，
    # 不再依赖 xlsx 留存与原生 build 触发，杜绝历史漏天。build() 内也走同一入口保证一致。
    #
    # 2026-10-16 加固：每次构建也拍一份「当日主看板快照」（覆盖写、幂等）。
    # 起因：10-15 机器在 23:59 前就关机了，snapshot_daily 那一枪没打出去
    #       → 整个 10-15 在主看板历史里漏天（_snap/ 里没有 2026-10-15.json）。
    # 白天构建每半小时就在跑，让它兜底后：只要当天开过机，这一天就不会漏。
    # ⚠️ 必须与 hist_snapshot 一样尊重 backfill 抑制位，否则补录历史时「当天」会被写脏。
    import history_builder
    try:
        if getattr(hist_snapshot, "SUPPRESS_AUTO_SNAPSHOT", False):
            print("  [skip] 当日主看板快照已抑制（backfill 模式）")
        else:
            history_builder.capture_today()
    except Exception as _e:
        print("  [warn] 当日主看板快照写入失败（不影响主构建）:", _e)
    hcount = history_builder.build_history_json(ensure_today=True)
    print("history.json 天数:", hcount)

    print("快照生成完成 →", DEPLOY_DIR)
    print("活动量 data_time:", data["meta"].get("data_time"),
          "| total:", len(data.get("total", [])),
          "| legal:", len(data.get("legal", [])),
          "| stats:", len(data.get("stats", [])))
    print("三力   data_time:", sanli["meta"].get("data_time"),
          "| 员工:", sanli.get("total"),
          "| 事业部:", len(sanli.get("bus", [])))
    print("人员   data_time:", ry["meta"].get("data_time"),
          "| 员工:", ry.get("total"),
          "| 部门:", len(ry.get("bus", [])))


if __name__ == "__main__":
    build()
