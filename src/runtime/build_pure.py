# -*- coding: utf-8 -*-
"""
build_pure.py v3 —— 纯净看板：完整复刻「四大模式健康度综合管理看板」(#modes) 设计。
- 视觉/交互：参照 重分数据看板-纯净版.html#modes（靛紫主题、岗位模式 tabs、综合数据子页签、
  两级部门筛选树、KPI 卡＋标准线、红字预警表、点击表头排序）。
- 数据：9月考核月-四大模式健康度综合管理表.xlsx（用户 09-18 直发）；岗位表逐列全量保留（指标不能少），
  综合视图的事业部/部门表由「所有规划师明细数据」聚合生成（xlsx 自带该两表为重庆模板空值）。
- 综合视图：构建期生成 SVG 图表（事业部柱线组合图+业绩占比环形图、部门业绩条形图）。
- 口径：目标/定义仍取 四大模式-重点指标表格.xlsx · 目标（9月 xlsm 为空壳，填好后自动切换）。
- 页面最底部：🤖 AI 数据解析 + 📖 口径说明（全页签可见）。
- 输出：_deploy_pure/index.html（自包含单页）。
"""
import os, json, sys, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import ry_parser as rp  # noqa: E402
import openpyxl  # noqa: E402

OUT_DIR = os.path.join(HERE, "_deploy_pure")
AUX_DIR = os.path.normpath(os.path.join(HERE, "..", "..", "..", "分公司健康表"))
NEW_XLSM = os.path.join(AUX_DIR, "9月考核月-四大模式健康度综合管理表.xlsm")
OLD_XLSX = os.path.join(AUX_DIR, "四大模式-重点指标表格.xlsx")
REF_CSS = os.path.join(HERE, "_hzcheck_k", "_ref_bundle.css")

# ----------------------------------------------------------------------------
# 1) 三大模块完整数据（计算逻辑与路径＝原链接）
# ----------------------------------------------------------------------------
RY = rp.build_meta()

# ----------------------------------------------------------------------------
# 2) 口径（目标/定义）
# ----------------------------------------------------------------------------

def parse_mode_metric_target_defs(ws):
    rws = list(ws.iter_rows(values_only=True))
    if len(rws) < 5:
        return []
    modes, metrics, targets, defs = rws[1], rws[2], rws[3], rws[4]
    maxcol = ws.max_column or 0
    cols = []
    for c in range(1, maxcol):
        m = modes[c] if c < len(modes) else None
        if m in (None, ""):
            continue
        mt = metrics[c] if c < len(metrics) else None
        tg = targets[c] if c < len(targets) else None
        df = defs[c] if c < len(defs) else None
        if (mt in (None, "")) and (tg in (None, "")) and (df in (None, "")):
            continue
        cols.append([m, mt if mt is not None else "", tg if tg is not None else "",
                     df if df is not None else ""])
    grouped = []
    for col in cols:
        if grouped and grouped[-1][0] == col[0] and grouped[-1][1] == col[1]:
            grouped[-1][2].append(col[2])
        else:
            grouped.append([col[0], col[1], [col[2]], col[3]])
    out = []
    for mode, metric, tgs, df in grouped:
        tgs = [str(t) for t in tgs if t not in (None, "")]
        ttxt = " / ".join(tgs) if tgs else "—"
        out.append({"mode": mode, "metric": metric, "target": ttxt, "definition": df or "—"})
    return out


TARGET = []
TARGET_SRC = ""
try:
    if os.path.exists(NEW_XLSM):
        wb_new = openpyxl.load_workbook(NEW_XLSM, read_only=True, data_only=True)
        if "DLP" in wb_new.sheetnames:
            cand = parse_mode_metric_target_defs(wb_new["DLP"])
            non_empty = sum(1 for row in wb_new["DLP"].iter_rows(values_only=True)
                            for c in row if c is not None)
            if cand and non_empty > 5:
                TARGET, TARGET_SRC = cand, "9月考核月-四大模式健康度综合管理表.xlsm · DLP"
except Exception as e:  # noqa: BLE001
    print("WARN: read new xlsm failed:", repr(e))
if not TARGET and os.path.exists(OLD_XLSX):
    wb_old = openpyxl.load_workbook(OLD_XLSX, read_only=True, data_only=True)
    TARGET = parse_mode_metric_target_defs(wb_old["目标"])
    TARGET_SRC = "四大模式-重点指标表格.xlsx · 目标"

# ----------------------------------------------------------------------------
# 3) MODES 数据整形（移植参考件 build_modes.js 的 colMeta/isRed/fmtCell 口径）
# ----------------------------------------------------------------------------
BIZ_THRESH = {"工商": 0.15, "跨境": 0.08, "法律": 0.07, "智能体": 0.04, "AI": 0.04,
              "交易": 0.03, "知识产权": 0.05}
# v14：业态下限在页面上的展示名（键须与 BIZ_THRESH 一致，顺序即展示顺序）
BIZ_LABEL = {"工商": "工商", "跨境": "跨境", "法律": "法律", "智能体": "AI智能体",
             "AI": "AI智能体", "交易": "交易", "知识产权": "知识产权"}
# 页面展示顺序（「智能体」只是匹配别名，与 "AI" 同值，不重复展示）
BIZ_SHOW = ["工商", "跨境", "法律", "AI", "交易", "知识产权"]


def to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    is_pct = "%" in s
    s = s.replace("¥", "").replace(",", "").replace("%", "").replace(" ", "")
    if s in ("", "—", "-"):
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    return n / 100 if is_pct else n


def col_meta(col, mode):
    m = {"key": col, "type": "num", "unit": "", "std": None, "card": None, "isStd": False}
    if "工号" in col:
        m["type"] = "id"
        return m
    if "人员编制" in col:
        # v17：主看板「员工档案」第 8 项（岗位 / 归属）＝文本列
        m["type"] = "id"
        return m
    if "转化率" in col or "签单率" in col:
        m.update(type="pct", unit="%", card="avg", isStd=True, std=0.10 if
                 ("网销模式付费客资源" in col or "近30天付费客资源转化率" in col) else 0.30)
        if ("付费客资源" in col and "转化率" in col) or col in (
                "见客签单转化率", "见客签单转化率（主面谈）"):
            # v14/v16：目标按业态下限（工商15/跨境8/法律7/AI 4/交易3）
            m["buStd"] = True
        return m
    if "当月到期" in col and "家数" in col:      # 会计到期家数（09-20）
        m.update(type="num", unit="家", card="sum")
        return m
    if "续费率" in col or "到期续率" in col:      # 「到期续率」(团队表) / 「到期续费率」(明细表)
        m.update(type="pct", unit="%", card="avg", isStd=True, std=0.85)
        return m
    if "出错率" in col:
        m.update(type="pct", unit="%", card="avg")
        return m
    if "人均产值" in col:
        thr = 30000 if mode == "dakehu" else (10000 if mode == "dianyao" else 20000)
        if "主看板" in col:
            thr = 20000          # v17：员工档案口径与主看板一致（统一 20000）
        m.update(type="money", unit="元", card="avg", isStd=True, std=thr)
        return m
    if "人均月度高质量见客量" in col or "人均月度高价值见客量" in col:
        m.update(type="num", card="avg", isStd=True, std=3)
        return m
    if "人均山头立项" in col:
        m.update(type="num", card="avg", isStd=True, std=0.5)
        return m
    if col in ("考核模式", "考核结果"):
        m.update(type="id")                       # 文本列（左对齐、不按数值格式化）
        return m
    if col == "达标率":                           # v16：达成指标数 ÷ 考核指标数
        m.update(type="pct", unit="%", card="avg")
        return m
    if col in ("达标数", "未达标数"):            # v12 旧列（v16 已被考核结果/达标率取代）
        m.update(type="num", unit="人", card="sum")
        return m
    if "人数" in col and "人均" not in col:
        m.update(type="num", unit="人", card="sum")
        return m
    if "业绩" in col and "人均" not in col:
        m.update(type="money", unit="元", card="sum")
        return m
    if ("见客量" in col or "签单量" in col or "接通量" in col or "加微量" in col
            or "资源量" in col or "通时" in col or "流失户数" in col or "达标项" in col):
        m.update(type="num", card="sum")
        return m
    return m


def is_red(col, val, mode, row):
    n = to_num(val)
    if n is None:
        return False
    if "人均产值" in col:
        thr = 30000 if mode == "dakehu" else (10000 if mode == "dianyao" else 20000)
        if "主看板" in col:
            thr = 20000          # v17：与主看板一致
        return n < thr
    if "转化率" in col or "签单率" in col:
        if ("网销模式付费客资源签单转化率" in col or "近30天付费客资源转化率" in col
                or col in ("见客签单转化率", "见客签单转化率（主面谈）")):
            biz = ""
            for k in ("事业部", "三级部门", "业态", "部门"):
                if row.get(k):
                    biz = str(row[k])
                    break
            t = None
            for key, tv in BIZ_THRESH.items():
                if key in biz:
                    t = tv
                    break
            if t is None:
                t = 0.10
            return n < t
        return n < 0.30
    if "续费率" in col or "到期续率" in col:
        return n < 0.85
    if "人均月度高质量见客量" in col or "人均月度高价值见客量" in col:
        gl = row.get("司龄月")
        gv = to_num(gl) if gl is not None else None
        th2 = 3 if (gv is not None and gv <= 3) else 5
        if row.get("岗龄"):
            gls = str(row["岗龄"])
            th2 = 3 if ("3个月" in gls and "以上" not in gls) else 5
        return n < th2
    return False


def fmt_cell(v, ctype):
    if v is None:
        return ""
    if ctype == "id":
        return str(v)
    if isinstance(v, (int, float)):
        if ctype == "pct":
            return "{:.2f}%".format(v * 100)
        if ctype == "money":
            return "{:,.0f}".format(round(v))
        r = round(v)
        if abs(v - r) < 1e-4:
            return "{:,}".format(r)
        return "{:,.2f}".format(v)
    if ctype == "pct":
        return str(v)
    return str(v)


def make_sheet(rows_dicts, mode):
    """rows_dicts: [{col: raw}] → {headers, rows(display), redFlags, cols}"""
    headers, seen = [], {}
    for rd in rows_dicts:
        for k in rd:
            if k not in seen:
                seen[k] = 0
                headers.append(k)
    out_rows, out_reds = [], []
    for rd in rows_dicts:
        disp, reds = [], []
        for i, h in enumerate(headers):
            v = rd.get(h)
            ctype = col_meta(h, mode)["type"]
            disp.append(fmt_cell(v, ctype))
            if is_red(h, v, mode, rd):
                reds.append(i)
        out_rows.append(disp)
        out_reds.append(reds)
    return {"headers": headers, "rows": out_rows, "redFlags": out_reds,
            "cols": [col_meta(h, mode) for h in headers]}


# ---- 数据源：9月考核月 四大模式健康度综合管理表.xlsx（用户直发） ----
XLSX_LOCAL = os.path.join(HERE, "9月考核月-四大模式健康度综合管理表.xlsx")
XLSX_WXWORK = r"<DATA_CACHE>/1688857640763810/Cache/File/2026-09/四大模式健康度综合管理表.xlsx"
if not os.path.exists(XLSX_LOCAL) and os.path.exists(XLSX_WXWORK):
    import shutil
    shutil.copy2(XLSX_WXWORK, XLSX_LOCAL)
if not os.path.exists(XLSX_LOCAL):
    raise SystemExit("找不到 9月考核月-四大模式健康度综合管理表.xlsx（本地与 WXWork 缓存均无）")
XLSX_SRC = XLSX_LOCAL

BUS = (RY.get("kuaiji_detail") or {}).get("bus") or []  # ai_insights 仍用

TEXT_COLS = {"事业部", "部门", "岗位", "负责人", "姓名", "入职日期"}
ERR_VALS = {"#DIV/0!", "#REF!", "#N/A", "#VALUE!", "#NAME?", "#NULL!", "#NUM!", "#SPILL!"}


def _clean_cell(v):
    if v is None:
        return ""
    if isinstance(v, str):
        v = v.strip()
        return "" if v in ERR_VALS else v
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return int(v)
        return v
    return v


def read_xlsx_sheet(wb, sn):
    """返回 (headers(重名去重), rows(list[dict]))；连续 3 空行截断（防 1048575 行幽灵）。"""
    ws = wb[sn]
    raw, blank = [], 0
    for row in ws.iter_rows(values_only=True):
        vals = [_clean_cell(c) for c in row]
        while vals and vals[-1] == "":
            vals.pop()
        if not vals:
            blank += 1
            if raw and blank >= 3:
                break
            continue
        blank = 0
        raw.append(vals)
    headers, seen = [], {}
    for i, h in enumerate(raw[0]):
        h = str(h).strip() if h not in (None, "") else "列%d" % (i + 1)
        if h in seen:
            seen[h] += 1
            h = h + "②" if seen[h] == 2 else "%s(%d)" % (h, seen[h])
        else:
            seen[h] = 1
        headers.append(h)
    rows = []
    for vals in raw[1:]:
        rd = {}
        for i, h in enumerate(headers):
            v = vals[i] if i < len(vals) else ""
            if h in TEXT_COLS and (v == 0 or v == "0"):
                v = ""
            rd[h] = v
        rows.append(rd)
    return headers, rows


WB = openpyxl.load_workbook(XLSX_SRC, read_only=True, data_only=True)


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def _f(v):
    """宽松数值解析（BI 导出可能是字符串带千分位）"""
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("¥", "").strip() or 0)
    except ValueError:
        return 0.0


# ----------------------------------------------------------------------------
# 业绩口径（用户 09-20）
#   ① 分公司 / 事业部业绩 ← 《公司和部门业绩》sheet（分公司 与 各事业部「业绩完成」）
#   ② 个人业绩            ← 《人员业绩汇总》sheet；工作簿内无该表时回退最新「个人业绩汇总」导出
#   ③ 法律（事业部名含「法律」）取「利润」列，其余取「业绩总和」
#   用途：人均业绩（人均产值）、人员明细、事业部/部门业绩
# ----------------------------------------------------------------------------
def _raw_grid(ws):
    """原始二维表（连续 3 空行截断，防 1048575 幽灵行）"""
    out, blank = [], 0
    for row in ws.iter_rows(values_only=True):
        vals = [_clean_cell(c) for c in row]
        while vals and vals[-1] == "":
            vals.pop()
        if not vals:
            blank += 1
            if out and blank >= 3:
                break
            continue
        blank = 0
        out.append(vals)
    return out


def _find_header(grid, must, limit=12):
    """在前 limit 行中查找同时含 must 关键词的表头行（-1 = 未找到）"""
    for i, r in enumerate(grid[:limit]):
        cs = [str(c).strip() for c in r]
        if all(any(m == c or m in c for c in cs) for m in must):
            return i
    return -1


def _load_company_perf():
    """《公司和部门业绩》→ ({事业部: 业绩完成}, 分公司合计, 源描述)"""
    if "公司和部门业绩" not in WB.sheetnames:
        return {}, None, ""
    grid = _raw_grid(WB["公司和部门业绩"])
    hi = _find_header(grid, ("事业部", "业绩完成"))
    if hi < 0:
        return {}, None, ""
    hs = [str(c).strip() for c in grid[hi]]
    ib = hs.index("事业部")
    ip = next(i for i, c in enumerate(hs) if "业绩完成" in c)
    per, total = {}, None
    for r in grid[hi + 1:]:
        if len(r) <= ip:
            continue
        b = str(r[ib]).strip()
        if not b.startswith(""):
            continue
        v = _f(r[ip])
        if b == "分公司":
            total = v
        elif "事业部" in b or "其他" in b:
            per[b] = v
    return per, total, "公司和部门业绩"


def _load_emp_perf():
    """《人员业绩汇总》→ {姓名: {code,biz,dept,sum,pro,val}}；法律取利润，其余取业绩总和"""
    grid, src = None, ""
    _sn = next((s for s in WB.sheetnames if "人员业绩" in str(s)), None)
    if _sn:                                                   # ① 工作簿内同名 sheet
        _wbf = openpyxl.load_workbook(XLSX_SRC, data_only=True)   # 非 read_only：规避维度缺失
        grid, src = _raw_grid(_wbf[_sn]), "工作簿·" + str(_sn)
        _wbf.close()
    if grid is None:                                          # ② 最新 BI 导出兜底
        import glob
        cand = []
        for _d in (os.path.join(os.environ.get("USERPROFILE") or "", "Downloads"),
                   os.path.join(os.path.expanduser("~"), "Downloads"),
                   r"<HOME>\Downloads"):
            if _d and os.path.isdir(_d):
                cand += glob.glob(os.path.join(_d, "个人业绩汇总*.xlsx"))
        cand = [c for c in sorted(set(cand), key=os.path.getmtime)
                if not os.path.basename(c).startswith("~$")]
        if cand:
            _wbx = openpyxl.load_workbook(cand[-1], data_only=True)
            grid = _raw_grid(_wbx.worksheets[0])
            src = os.path.basename(cand[-1])
    if not grid:
        return {}, ""
    hi = _find_header(grid, ("姓名", "业绩"), limit=8)
    if hi < 0:
        return {}, ""
    hs = [str(c).strip() for c in grid[hi]]

    def col(*names):
        for n in names:
            if n in hs:
                return hs.index(n)
        for n in names:
            for i, c in enumerate(hs):
                if n in c:
                    return i
        return -1

    iN, iC, iB, iD = col("姓名"), col("工号"), col("事业部"), col("部门")
    iSum = col("业绩总和", "业绩")
    iPro = col("利润总和")
    if iPro < 0:
        for i, c in enumerate(hs):
            if "利润" in c and "续费" not in c and "核算" not in c:
                iPro = i
                break
    if iN < 0 or iSum < 0:
        return {}, ""
    out = {}
    for r in grid[hi + 1:]:
        nm = str(r[iN]).strip() if iN < len(r) else ""
        if not nm or nm in ("合计", "总计"):
            continue
        biz = str(r[iB]).strip() if 0 <= iB < len(r) else ""
        s = _f(r[iSum]) if iSum < len(r) else 0.0
        pr = _f(r[iPro]) if 0 <= iPro < len(r) else 0.0
        out[nm] = {"code": str(r[iC]).strip() if 0 <= iC < len(r) else "", "biz": biz,
                   "dept": str(r[iD]).strip() if 0 <= iD < len(r) else "",
                   "sum": s, "pro": pr, "val": pr if "法律" in biz else s}
    return out, src


PERF_HZ, PERF_TOTAL, PERF_CO_SRC = _load_company_perf()
EMP_PERF, EMP_PERF_SRC = _load_emp_perf()


# 口径（用户 09-20）：以下职能/支撑类事业部的人员
#   —— 不显示在表格中，也不参与任何计算（部门/事业部聚合、合计、KPI、图表）
EXCLUDE_BIZ = ("消化中心", "职能", "业务管理中心", "业务支持部")


def _bz(r):
    return str(r.get("事业部") or "").strip()


def _keep_row(r):
    b = _bz(r)
    return bool(b) and b not in EXCLUDE_BIZ


# 编制人数口径（用户 09-20 晚）：只统计「在职」∧「一线业务岗」的人。
#   · 在职       = 出现在当期在职花名册《所有规划师明细数据》（业绩汇总里的离场人员不算）
#   · 一线业务岗  = 企服/电邀规划师、法务主任、提成律师、会计顾问/客服会计/会计主管、企服（储备）主管
#   · 不计入     = 管理者（经理/总监/总/副总/负责人）、中台老师与解决方案岗（专家/专员）、
#                  职能支撑（HRBP/PMO/交付）、流程会计（会计板块剔除）
STAFF_POSTS = ("企服规划师", "电邀规划师", "法务主任", "提成律师", "会计顾问",
               "客服会计", "会计主管", "企服主管", "企服储备主管")
NON_STAFF_KW = ("经理", "总监", "副总", "负责人", "专家", "专员",
                "HRBP", "PMO", "交付", "流程会计")
ONBOARD_FLAG = "_onboard"


def _is_staff(r):
    """该行是否计入「编制人数」（在职一线业务岗）。"""
    if not r.get(ONBOARD_FLAG):
        return False
    p = str(r.get("岗位") or "").strip()
    if not p or any(k in p for k in NON_STAFF_KW):
        return False
    return p in STAFF_POSTS or p.endswith("规划师")


# ---- 会计到期三指标（用户 09-20）：源＝历史源/2026-09-14「四大模式-重点指标表格.xlsx」· 会计续费明细表
#   口径与主看板「会计续费明细」板块（by_advisor 顾问口径）逐人对齐：
#     当月到期家数     = 是否应续费 = 是 的家数
#     当月到期实收家数 = 应续 且 是否续费 = 是 的家数
#     到期续费率       = 当月到期实收家数 ÷ 当月到期家数（标准 85%，见「标准」sheet c35）
HIST_SNAP_DIR = os.path.join(HERE, "历史源", "2026-09-14")
KJ_DUE_HEADER = ("当月到期实收家数", "当月到期家数")


def _kj_due_load():
    """→ ({"person": {姓名: {...}}, "dept": {顾问部门: {...}}, "all": {...}}, src)"""
    empty = {"person": {}, "dept": {}, "all": {}}
    cands = [os.path.join(HIST_SNAP_DIR, "四大模式-重点指标表格.xlsx"),
             os.path.join(AUX_DIR, "四大模式-重点指标表格.xlsx")]
    for p in cands:
        if not os.path.exists(p):
            continue
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            if "会计续费明细表" not in wb.sheetnames:
                continue
            rows = list(wb["会计续费明细表"].iter_rows(values_only=True))
        except Exception:                                          # noqa: BLE001
            continue
        if len(rows) < 3:
            continue
        head = [str(c).strip() if c is not None else "" for c in rows[1]]
        ix = {h: i for i, h in enumerate(head)}
        need = ("公司名称", "会计顾问", "会计顾问部门", "是否续费", "是否应续费")
        if any(k not in ix for k in need):
            continue

        def _g(r, k):
            i = ix[k]
            return str(r[i]).strip() if (i < len(r) and r[i] is not None) else ""

        def _nm(v):
            import re as _re
            m = _re.match(r"^(.*?)(\d{4,})$", v)
            return (m.group(1) if m else v).strip()

        recs = []
        for r in rows[2:]:
            if not _g(r, "公司名称"):
                continue
            recs.append({"顾问": _nm(_g(r, "会计顾问")), "顾问部": _g(r, "会计顾问部门"),
                         "续": _g(r, "是否续费"), "应续": _g(r, "是否应续费")})
        if not recs:
            continue

        def _agg(items, key):
            g = {}
            for d in items:
                g.setdefault(d[key] or "（空）", []).append(d)
            out = {}
            for k, v in g.items():
                due = [d for d in v if d["应续"] == "是"]
                recv = [d for d in due if d["续"] == "是"]
                out[k] = {"到期": len(due), "实收": len(recv), "记录": len(v)}
            return out

        allrec = {"到期": len([d for d in recs if d["应续"] == "是"]),
                  "实收": len([d for d in recs if d["应续"] == "是" and d["续"] == "是"]),
                  "记录": len(recs)}
        _lab = os.path.relpath(p, HERE).replace("\\", "/")
        return ({"person": _agg(recs, "顾问"), "dept": _agg(recs, "顾问部"),
                 "all": allrec}, _lab + " · 会计续费明细表")
    return empty, ""


KJ_DUE, KJ_DUE_SRC = _kj_due_load()


def _kj_due_rate(d):
    if not d:
        return None
    due = d.get("到期") or 0
    return (d.get("实收", 0) / due) if due else None


def _kj_due_apply_person(rows):
    """会计岗明细：按姓名填三列（查不到 → 0 家 / 空率，不臆造）。"""
    mp = KJ_DUE.get("person") or {}
    for r in rows:
        d = mp.get(str(r.get("姓名") or "").strip())
        r["当月到期实收家数"] = d["实收"] if d else 0
        r["当月到期家数"] = d["到期"] if d else 0
        r["到期续费率"] = _kj_due_rate(d)
    return rows


def _kj_due_apply_team(rows, detail_rows):
    """会计岗团队数据：按本行内出现的「部门」归集明细行求和（部门合计口径）。"""
    dept = {}
    for x in detail_rows:
        k = str(x.get("部门") or "").strip()
        if k:
            dept.setdefault(k, []).append(x)
    for r in rows:
        mem = []
        for v in r.values():
            k = str(v or "").strip()
            if k in dept:
                mem = dept[k]
                break
        due = sum(int(x.get("当月到期家数") or 0) for x in mem)
        recv = sum(int(x.get("当月到期实收家数") or 0) for x in mem)
        r["当月到期实收家数"] = recv
        r["当月到期家数"] = due
        r["到期续率"] = (recv / due) if due else None
    return rows


# 会计岗团队数据（用户 09-20 晚）：会计团队＝在职一线会计岗（会计顾问/客服会计/会计主管），
#   流程会计是后台票据岗 ⇒ 不计入团队人数。源表只给了「会计顾问」一组 → 按本部门在职业务岗
#   重算团队人数及随之变化的业绩/人均/签单/高端/山头；到期·流失·票据列沿用客户与流程口径不动。
KJ_STAFF_POSTS = ("会计顾问", "客服会计", "会计主管")


def _kj_team_headcount(team_rows, detail_rows):
    from collections import defaultdict
    g = defaultdict(list)
    for d in detail_rows:
        if str(d.get("岗位") or "").strip() in KJ_STAFF_POSTS:
            g[str(d.get("部门") or "").strip()].append(d)
    for t in team_rows:
        dep = ""
        for k in ("部门", "岗位", "岗位②"):
            v = str(t.get(k) or "").strip()
            if v in g:
                dep = v
                break
        mem = g.get(dep) or []
        if not mem:
            continue
        n = len(mem)
        incl = sum(num(x.get("业绩（含续费）")) for x in mem)
        excl = sum(num(x.get("业绩（不含续费）")) for x in mem)
        t["团队人数"] = n
        t["业绩(含续费）"] = incl
        t["业绩(不含续费）"] = excl
        t["人均业绩（含续费）"] = incl / n
        t["人均业绩（不含续费）"] = excl / n
        t["总签单数"] = sum(num(x.get("总签单数")) for x in mem)
        hl = sum(num(x.get("高端账签单量")) for x in mem)
        hp = sum(num(x.get("高端账签单业绩")) for x in mem)
        t["高端账签单量"] = hl
        t["高端账签单业绩"] = hp
        t["高端账签单占比"] = (hp / incl) if incl else 0
        t["高端账签单客均"] = (hp / hl) if hl else 0
        li = sum(num(x.get("山头立项")) for x in mem)
        sg = sum(num(x.get("山头签单")) for x in mem)
        t["人均山头立项"] = li / n
        t["山头转化率"] = (sg / li) if li else 0
    return team_rows


# ---- 岗位六表：逐列全量保留（指标不能少） ----
SHEET_MODE = {"网销岗团队数据": "wangxiao", "网销岗明细": "wangxiao",
              "电邀岗团队数据": "dianyao", "电邀岗明细": "dianyao",
              "会计岗团队数据": "kuaiji", "会计岗明细": "kuaiji"}
_RAW = {}
for _sn, _md in SHEET_MODE.items():
    _h, _r = read_xlsx_sheet(WB, _sn)
    if "事业部" in _h:                      # v6：职能/支撑类事业部人员不出现在表中
        _r = [x for x in _r if _keep_row(x)]
    if _sn == "会计岗团队数据" and "事业部" not in _h:
        # 该表只有 部门 层级 → 补一列 事业部（用于左侧冻结列对齐）
        _r = [dict([("事业部", "财务会计事业部")] + list(x.items())) for x in _r]
    _RAW[_sn] = _r

_kj_due_apply_person(_RAW.get("会计岗明细") or [])
_kj_due_apply_team(_RAW.get("会计岗团队数据") or [], _RAW.get("会计岗明细") or [])
_kj_team_headcount(_RAW.get("会计岗团队数据") or [], _RAW.get("会计岗明细") or [])
# ---- v15：事业部负责人（源表「部门」列为空）按指定归入其分管部门 ----
#   示例负责人A同时挂两个部门 ⇒ 该人在明细表中出现两次。4 人各指标均为 0，仅影响分组显示，
#   不参与聚合（AGG_ROWS 在下方由 U_CLEAN 的原对象构建，override 只作用在落表副本上）。
DEPT_OVERRIDE = {
    "示例负责人A": ["工商网销业务二部", "工商电邀业务六部"],
    "示例负责人B": ["法律咨询三部"],
    "示例负责人C": ["跨境业务二部"],
    "示例负责人D": ["数字科技二部"],
}


def _dept_override(rows):
    out = []
    for _r in rows:
        _ds = DEPT_OVERRIDE.get(str(_r.get("姓名") or "").strip())
        if not _ds:
            out.append(_r)
            continue
        for _d in _ds:
            _c = dict(_r)
            _c["部门"] = _d
            out.append(_c)
    return out


# ---------- v16：部门考核（用户 09-20 晚） ----------
#   网销模式：只考核「付费客资源签单转化率（近30天）」，目标＝业态下限
#             （工商 15% / 跨境 8% / 法律 7% / AI智能体 4% / 交易 3%）
#   电邀模式：人均月度高价值见客量 —— 入职 0-3 月 ≥3；入职 3 个月以上 ≥5
#   会计模式：到期续费率 ≥85% 且 非正常流失率 ≤1%
#   部门级：任一考核指标未达标 ⇒ 该部门未达标；达标率 ＝ 达成指标数 ÷ 考核指标数
EXAM = {}


def _biz_thr(biz):
    for _k, _v in BIZ_THRESH.items():
        if _k in str(biz or ""):
            return _v
    return None


def _exam_mode(ps):
    _ms = []
    for _p in ps:
        if _p and _p["mode"] not in _ms:
            _ms.append(_p["mode"])
    return "+".join(_ms)


def _exam_pack(items, mode, biz):
    if not items:
        return None
    _okn = sum(1 for it in items if it["ok"])
    return {"mode": mode, "biz": biz, "items": items, "okn": _okn, "tot": len(items),
            "ok": _okn == len(items), "rate": _okn / len(items)}


def _exam_build():
    _out = {}
    # ① 会计：到期续费率 ≥85% 且 非正常流失率 ≤1%
    _g = {}
    for _r in (_RAW.get("会计岗明细") or []):
        _d = str(_r.get("部门") or "").strip()
        if _d:
            _g.setdefault(_d, []).append(_r)
    for _d, _rs in _g.items():
        _items = []
        _due = sum(num(r.get("当月到期家数")) for r in _rs)
        _recv = sum(num(r.get("当月到期实收家数")) for r in _rs)
        if _due:
            _v = _recv / _due
            _items.append({"k": "到期续费率", "v": _v, "t": 0.85, "dir": ">=",
                           "ok": _v >= 0.85})
        _held = sum(num(r.get("保有量")) for r in _rs)
        _lost = sum(num(r.get("非正常流失家数")) for r in _rs)
        if _held:
            _lv = _lost / _held
        else:
            _vs = [num(r.get("非正常流失率")) for r in _rs
                   if r.get("非正常流失率") not in (None, "")]
            _lv = (sum(_vs) / len(_vs)) if _vs else None
        if _lv is not None:
            _items.append({"k": "非正常流失率", "v": _lv, "t": 0.01, "dir": "<=",
                           "ok": _lv <= 0.01})
        _p = _exam_pack(_items, "会计", str(_rs[0].get("事业部") or "").strip())
        if _p:
            _out[_d] = _p
    # ② 电邀：人均月度高价值见客量（按入职时长分档：0-3月≥3 / 3个月以上≥5）
    _g = {}
    for _r in (_RAW.get("电邀岗明细") or []):
        _d = str(_r.get("部门") or "").strip()
        if _d:
            _g.setdefault(_d, []).append(_r)
    for _d, _rs in _g.items():
        if _d in _out:
            continue
        _items = []
        for _lbl, _lo, _hi, _thr in (("入职0-3月", 0.0, 3.0, 3),
                                     ("入职3个月以上", 3.0000001, 1e9, 5)):
            _mem = [r for r in _rs if _lo <= num(r.get("司龄月")) <= _hi]
            if not _mem:
                continue
            _v = sum(num(r.get("人均月度高价值见客量")) for r in _mem) / len(_mem)
            _items.append({"k": "人均月度高价值见客量（%s）" % _lbl, "v": _v,
                           "t": _thr, "dir": ">=", "ok": _v >= _thr})
        _p = _exam_pack(_items, "电邀", str(_rs[0].get("事业部") or "").strip())
        if _p:
            _out[_d] = _p
    # ③ 网销：付费客资源签单转化率 ≥ 业态下限（唯一考核指标）
    _g = {}
    for _r in (_RAW.get("网销岗明细") or []):
        _b = str(_r.get("事业部") or "").strip()
        _d = str(_r.get("部门") or "").strip()
        if _b and _d:
            _g.setdefault((_b, _d), []).append(_r)
    for (_b, _d), _rs in _g.items():
        if _d in _out:
            continue
        _thr = _biz_thr(_b)
        _res = sum(num(r.get("近30天付费客资源总量")) for r in _rs)
        if not _res or _thr is None:
            continue
        _v = sum(num(r.get("近30天付费客资源签单量")) for r in _rs) / _res
        _p = _exam_pack([{"k": "付费客资源签单转化率", "v": _v, "t": _thr,
                          "dir": ">=", "ok": _v >= _thr}], "网销", _b)
        if _p:
            _out[_d] = _p
    return _out


EXAM = _exam_build()


def _exam_for_biz(biz):
    _ps = [_p for _p in EXAM.values() if _p["biz"] == biz]
    return _exam_pack([it for _p in _ps for it in _p["items"]], _exam_mode(_ps), biz)


def _exam_all():
    _ps = list(EXAM.values())
    return _exam_pack([it for _p in _ps for it in _p["items"]], _exam_mode(_ps), "")


def _exam_cols(p):
    """→ (考核模式, 考核结果, 达标率)"""
    if not p:
        return "", "", ""
    return p["mode"], ("达标" if p["ok"] else "未达标"), p["rate"]


def _exam_team_rows():
    """模式团队表逐行判定（用团队行自身的指标值），写「考核结果」。"""
    for _r in (_RAW.get("网销岗团队数据") or []):
        _thr = _biz_thr(_r.get("事业部"))
        _v = num(_r.get("网销模式付费客资源签单转化率（近30天）"))
        _r["考核模式"] = "网销"
        _r["考核结果"] = ("达标" if _v >= _thr else "未达标") if _thr is not None else ""
    for _r in (_RAW.get("电邀岗团队数据") or []):
        _v = num(_r.get("人均月度高价值见客量"))
        _thr = 3 if "3个月内" in str(_r.get("岗龄") or "") else 5
        _r["考核模式"] = "电邀"
        _r["考核结果"] = "达标" if _v >= _thr else "未达标"
    for _r in (_RAW.get("会计岗团队数据") or []):
        _seen, _ok = False, True
        _due = num(_r.get("当月到期家数"))
        if _due:
            _seen = True
            _ok = _ok and (num(_r.get("当月到期实收家数")) / _due >= 0.85)
        if _r.get("非正常流失率") not in (None, ""):
            _seen = True
            _ok = _ok and (num(_r.get("非正常流失率")) <= 0.01)
        _r["考核模式"] = "会计"
        _r["考核结果"] = ("达标" if _ok else "未达标") if _seen else ""
    return


# ---- v17：主看板「员工档案」8 大指标（ry_parser km8 口径）· 按姓名匹配追加 ----
RY_EMP = {}
for _e in (RY.get("employees") or []):
    _nm = str(_e.get("姓名") or "").strip()
    if _nm:
        RY_EMP[_nm] = _e

KM8_KEYS = ("付费客资源签单转化率", "员工见客签单转化率", "业务经理陪同见客签单转化率",
            "中台老师陪同见客签单转化率", "事业部总陪同见客签单转化率",
            "当月见客签单转化率", "人均产值", "人员编制")
KM8_SUF = "（主看板）"
KM8_HIT = 0


def _km8_apply(rows):
    """给明细行追加主看板 8 大指标列（源：ry_parser km8；未命中＝「—」）。"""
    global KM8_HIT
    for _r in rows:
        _nm = str(_r.get("姓名") or "").strip()
        _e = RY_EMP.get(_nm)
        _k = ((_e or {}).get("km8") or {}) if _e else {}
        if _e:
            KM8_HIT += 1
        for _c in KM8_KEYS:
            if _c == "人员编制":
                _v = ((_k.get("detail") or {}).get("人员编制") or {}).get("pos")
                _r[_c + KM8_SUF] = _v or ((_e or {}).get("岗位") or "—")
            else:
                _v = _k.get(_c)
                _r[_c + KM8_SUF] = _v if isinstance(_v, (int, float)) else "—"
    return rows

_exam_team_rows()

SHEETS = {}
for _sn, _md in SHEET_MODE.items():
    _rows = _dept_override(_RAW[_sn]) if _sn.endswith("明细") else _RAW[_sn]
    if _sn.endswith("明细"):        # v17：主看板员工档案 8 大指标
        _rows = _km8_apply(_rows)
    SHEETS[_sn] = make_sheet(_rows, _md)

# ---- 综合视图：由「所有规划师明细数据」聚合（xlsx 自带事业/部门两表为模板空值） ----
_, U_ROWS = read_xlsx_sheet(WB, "所有规划师明细数据")

BIZ_HEADERS = ["事业部", "编制人数", "业绩", "人均产值", "所有见客", "见客(主面谈）",
               "见客签单数", "见客签单转化率", "考核模式", "考核结果", "达标率",
               "山头立项", "山头签单", "山头签单转化率",
               "近30天付费资源量（不包含三方）", "近30天付费客资源签单量", "近30天付费客资源转化率"]
DEPT_HEADERS = ["事业部", "部门", "编制人数", "编制应增应减"] + \
               [h for h in BIZ_HEADERS[1:] if h != "编制人数"] + \
               ["网销人数", "大客户人数", "电邀人数"]     # 去重：编制人数 只保留 1 列


def _agg(rows, fields, perf_map=None, order=None):
    """按 fields 聚合。
    perf_map：仅单字段分组时生效，用权威业绩口径覆盖「业绩 / 人均产值」（如事业部←公司和部门业绩）。
    order：分组顺序（缺项自动补在末尾）。
    """
    groups = {}
    for r in rows:
        k = tuple(str(r.get(f) or "").strip() for f in fields)
        if len(fields) > 1 and not k[-1]:
            continue
        if not any(k):
            continue
        groups.setdefault(k, []).append(r)
    if order:
        # 指定顺序的分组即使暂无人员也保留（业绩由 perf_map 权威给出），保证合计=各行之和
        keys = [k for k in order if (k in groups or perf_map is not None)] + \
               [k for k in sorted(groups) if k not in order]
    else:
        keys = sorted(groups)
    out = []
    for k in keys:
        rs = groups.get(k, [])
        perf = sum(num(r.get("业绩")) for r in rs)
        if perf_map is not None and len(fields) == 1 and k[0] in perf_map:
            perf = perf_map[k[0]]
        cnt = sum(1 for r in rs if _is_staff(r))      # 编制人数＝在职一线业务岗
        visits = sum(num(r.get("所有见客")) for r in rs)
        sign = sum(num(r.get("见客签单数")) for r in rs)
        st_li = sum(num(r.get("山头立项")) for r in rs)
        st_sg = sum(num(r.get("山头签单")) for r in rs)
        res = sum(num(r.get("近30天付费资源量（不包含三方）")) for r in rs)
        rsg = sum(num(r.get("近30天付费客资源签单量")) for r in rs)
        pos = lambda kw: sum(1 for r in rs if kw in str(r.get("岗位") or ""))  # noqa: E731
        d = dict(zip(fields, k))
        d["编制人数"] = cnt
        d["业绩"] = perf
        d["人均产值"] = (perf / cnt) if cnt else ""
        d["所有见客"] = visits
        # v12：见客三项＝源《近30天付费客资源转化》按工号取；
        #      见客数 / 签单数＝求和；转化率＝有该表记录成员的平均（率类不适用 Σ签单÷Σ见客）
        _jkv = sum(num(r.get("见客(主面谈）")) for r in rs)
        _jks = sum(num(r.get("见客签单数")) for r in rs)
        _jkr = [num(JK_ROWS[_k]["rate"]) for _k in
                (str(r.get("工号") or "").strip() for r in rs) if _k in JK_ROWS]
        _jko = sum(1 for r in rs if _is_staff(r)
                   and num(r.get("见客签单转化率（主面谈）")) >= JK_STD)
        d["见客(主面谈）"] = _jkv
        d["见客签单数"] = _jks
        d["见客签单转化率"] = (sum(_jkr) / len(_jkr)) if _jkr else ""
        # v16：达标数/未达标数 → 考核模式/考核结果/达标率（＝达成指标数 ÷ 考核指标数）
        _exp = EXAM.get(str(d.get("部门") or "").strip()) if len(fields) > 1 else _exam_for_biz(k[0])
        for _cv, _v in zip(("考核模式", "考核结果", "达标率"), _exam_cols(_exp)):
            d[_cv] = _v
        d["山头立项"] = st_li
        d["山头签单"] = st_sg
        d["山头签单转化率"] = (st_sg / st_li) if st_li else 0
        d["近30天付费资源量（不包含三方）"] = res
        d["近30天付费客资源签单量"] = rsg
        d["近30天付费客资源转化率"] = (rsg / res) if res else 0
        d["编制应增应减"] = ""
        d["网销人数"] = pos("企服") + pos("网销")
        d["大客户人数"] = pos("大客户")
        d["电邀人数"] = pos("电邀")
        out.append(d)
    return out


# 口径（用户 09-20）：EXCLUDE_BIZ（消化中心/职能/业务管理中心/业务支持部）
#   ＝职能/支撑岗，不显示在表格中，也不参与任何计算。
# 咨询中心＝中台老师：保留在「个人明细数据」，但不进事业部/部门表（非经营事业部）。
# 另：事业部为空的管理行（如分公司总经理）与 _agg 口径一致，不参与聚合。
CALC_ROWS = []
for _r in U_ROWS:
    if not _keep_row(_r):
        continue
    _r[ONBOARD_FLAG] = True        # 在职花名册；业绩汇总补充的人无此标记 ⇒ 不计入编制
    CALC_ROWS.append(_r)

# ---- 人员明细＝源「所有规划师明细数据」 ∪ 《人员业绩汇总》 ----
#   业绩一律以个人业绩汇总为准（法律取「利润」列，其余取「业绩总和」）；
#   个人业绩汇总有、规划师源表没有的人（如交易三部部分成员）一并补进明细表。
U_CLEAN, _uni_names = [], set()
for _r in CALC_ROWS:
    _d = dict(_r)
    _e = EMP_PERF.get(str(_d.get("姓名") or "").strip())
    if _e:
        _d["业绩"] = _e["val"]
    U_CLEAN.append(_d)
    _uni_names.add(str(_d.get("姓名") or "").strip())
for _nm, _e in EMP_PERF.items():
    if _nm in _uni_names or not _e["biz"] or _e["biz"] in EXCLUDE_BIZ:
        continue
    U_CLEAN.append({"工号": _e["code"], "姓名": _nm, "事业部": _e["biz"], "部门": _e["dept"],
                    "岗位": "", "业绩": _e["val"]})

NON_BIZ = ("咨询中心",)                          # 中台老师，不进事业部/部门表
AGG_ROWS = [r for r in U_CLEAN if str(r.get("事业部") or "").strip() not in NON_BIZ]

# ---- 见客三项（用户 09-20）：员工级数值按工号匹配《近30天付费客资源转化》（端口）
#   见客数(主面谈)  ← 成单或见客量·30天
#   见客签单数      ← 成单量·30天
#   见客签单转化率  ← 30天及时性转化率（剔除三方）
#   达标判定：转化率 ≥ 30%（源表《标准》页「员工见客签单转化率（考核月）」目标 30%）
JK_STD = 0.30
JK_HEADER_NAME = "近30天付费客资源转化"


def _load_jk():
    """《近30天付费客资源转化》为双表头（成单量 / 成单或见客量 各带 1天…30天），故前 2 行跳过。"""
    try:
        ws = WB[JK_HEADER_NAME]
    except KeyError:
        return {}
    raw, blank = [], 0
    for row in ws.iter_rows(values_only=True):
        vals = [_clean_cell(c) for c in row]
        if all(v == "" for v in vals):
            blank += 1
            if raw and blank >= 3:
                break
            continue
        blank = 0
        raw.append(vals)
        if len(raw) > 20000:
            break
    out = {}
    for r in raw[2:]:
        if len(r) < 39 or not r[1]:
            continue
        if "" not in str(r[3] or ""):
            continue
        out[str(r[0]).strip()] = {"visit": num(r[38]), "sign": num(r[33]), "rate": num(r[28])}
    return out


JK_ROWS = _load_jk()
JK_HIT = 0
JK_KEYS = ("见客(主面谈）", "见客签单数", "见客签单转化率（主面谈）")
for _r in U_CLEAN:
    for _c in JK_KEYS:
        _r.setdefault(_c, 0)               # 业绩汇总补充行补键，避免落表多出列
    _v = JK_ROWS.get(str(_r.get("工号") or "").strip())
    if _v:
        JK_HIT += 1
        _r["见客(主面谈）"] = _v["visit"]
        _r["见客签单数"] = _v["sign"]
        _r["见客签单转化率（主面谈）"] = _v["rate"]
    else:
        _r["见客(主面谈）"] = 0
        _r["见客签单数"] = 0
        _r["见客签单转化率（主面谈）"] = 0
print("     见客三项：命中 %d 人 / 表内 %d 人" % (JK_HIT, len(JK_ROWS)))


def _total_row(rows):
    perf = sum(num(r.get("业绩")) for r in rows)
    visits = sum(num(r.get("所有见客")) for r in rows)
    sign = sum(num(r.get("见客签单数")) for r in rows)
    st_li = sum(num(r.get("山头立项")) for r in rows)
    st_sg = sum(num(r.get("山头签单")) for r in rows)
    res = sum(num(r.get("近30天付费资源量（不包含三方）")) for r in rows)
    rsg = sum(num(r.get("近30天付费客资源签单量")) for r in rows)
    pos = lambda kw: sum(1 for r in rows if kw in str(r.get("岗位") or ""))  # noqa: E731
    cnt = sum(1 for r in rows if _is_staff(r))
    _jkv = sum(num(r.get("见客(主面谈）")) for r in rows)
    _jks = sum(num(r.get("见客签单数")) for r in rows)
    _jkr = [num(JK_ROWS[_k]["rate"]) for _k in
            (str(r.get("工号") or "").strip() for r in rows) if _k in JK_ROWS]
    _jko = sum(1 for r in rows if _is_staff(r)
               and num(r.get("见客签单转化率（主面谈）")) >= JK_STD)
    _exa = _exam_cols(_exam_all())
    return {"编制人数": cnt, "业绩": perf,
            "人均产值": (perf / cnt) if cnt else 0,
            "所有见客": visits,
            "见客(主面谈）": _jkv,
            "见客签单数": _jks,
            "见客签单转化率": (sum(_jkr) / len(_jkr)) if _jkr else "",
            "考核模式": _exa[0], "考核结果": _exa[1], "达标率": _exa[2],
            "山头立项": st_li, "山头签单": st_sg,
            "山头签单转化率": (st_sg / st_li) if st_li else 0,
            "近30天付费资源量（不包含三方）": res,
            "近30天付费客资源签单量": rsg,
            "近30天付费客资源转化率": (rsg / res) if res else 0,
            "编制应增应减": "", "网销人数": pos("企服") + pos("网销"),
            "大客户人数": pos("大客户"), "电邀人数": pos("电邀")}


# ---- 事业部表：行＝《公司和部门业绩》各事业部（业绩/人均以该表为准，与「分公司」合计严格相等）
_BIZ_ORDER = [(b,) for b in PERF_HZ]
_biz = _agg(AGG_ROWS, ["事业部"], perf_map=PERF_HZ, order=_BIZ_ORDER)
_biz_total = []
if _biz:
    _total = _total_row(AGG_ROWS)
    _total["业绩"] = PERF_TOTAL if PERF_TOTAL is not None else _total["业绩"]
    _total["编制人数"] = sum(num(x.get("编制人数")) for x in _biz)
    _total["人均产值"] = (_total["业绩"] / _total["编制人数"]) if _total["编制人数"] else ""
    _total["事业部"] = "分公司合计"
    _biz_total = [_total]

# ---- 部门表：部门/个人业绩＝《人员业绩汇总》按（事业部+部门）聚合（法律取利润）
DEPT_RAW = _agg(AGG_ROWS, ["事业部", "部门"])
_DEPT_ORDER = {b: i for i, b in enumerate(PERF_HZ)}
DEPT_RAW.sort(key=lambda d: (_DEPT_ORDER.get(str(d.get("事业部")), 99), -num(d.get("业绩"))))
_dept_total = []
if _biz:
    _dt = _total_row(AGG_ROWS)
    _dt["业绩"] = PERF_TOTAL if PERF_TOTAL is not None else _dt["业绩"]
    _dt["编制人数"] = sum(num(x.get("编制人数")) for x in DEPT_RAW)
    _dt["人均产值"] = (_dt["业绩"] / _dt["编制人数"]) if _dt["编制人数"] else ""
    _dt["事业部"] = "分公司合计"
    _dt["部门"] = ""
    _dept_total = [_dt]

for _r in U_CLEAN:                 # 内部标记不落表
    _r.pop(ONBOARD_FLAG, None)
SHEETS["所有规划师明细数据"] = make_sheet(_km8_apply(_dept_override(U_CLEAN)), "wangxiao")  # v6：剔除职能/支撑岗
print(" 主看板 8 大指标（员工档案口径）：命中 %d 人次 / 网销员工 %d 人" % (KM8_HIT, len(RY_EMP)))
SHEETS["所有事业部明细数据"] = make_sheet(_biz + _biz_total, "wangxiao")
SHEETS["所有部门明细数据"] = make_sheet(DEPT_RAW + _dept_total, "wangxiao")
# 列顺序固定为表头全集（聚合 dict 缺列补空）
for _sn, _hs in (("所有事业部明细数据", BIZ_HEADERS), ("所有部门明细数据", DEPT_HEADERS)):
    sh = SHEETS[_sn]
    idx = {h: sh["headers"].index(h) for h in _hs if h in sh["headers"]}
    missing = [h for h in _hs if h not in sh["headers"]]
    sh["headers"] = _hs + missing
    sh["rows"] = [[(r[idx[h]] if h in idx else "") for h in sh["headers"]] for r in sh["rows"]]
    sh["redFlags"] = [[] for _ in sh["rows"]]
    # 重算 redFlags（列序变了）
    mode = "wangxiao"
    sh["cols"] = [col_meta(h, mode) for h in sh["headers"]]
    sh["redFlags"] = []
    for r in sh["rows"]:
        rd = dict(zip(sh["headers"], r))
        reds = [i for i, h in enumerate(sh["headers"]) if is_red(h, rd.get(h), mode, rd)]
        sh["redFlags"].append(reds)

# ---- 综合视图 SVG 图表（构建期生成，零依赖） ----
C_BAR1, C_BAR2, C_LINE, C_TXT, C_GRID = "#6366f1", "#818cf8", "#f59e0b", "#6b7280", "#e5e7eb"
PIE_COLORS = ["#6366f1", "#8b5cf6", "#06b6d4", "#10b981", "#f59e0b", "#fb7185", "#3b82f6", "#a78bfa"]


def _esc(sx):
    return str(sx).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def fmt_wan(v):
    v = float(v or 0)
    if abs(v) >= 1e8:
        return "%.1f亿" % (v / 1e8)
    if abs(v) >= 1e4:
        return "%.1f万" % (v / 1e4)
    return "{:,.0f}".format(v)


def svg_combo(biz):
    rows = [r for r in biz if (max(num(r.get("业绩")), 0) or num(r.get("所有见客")))
            and "合计" not in str(r.get("事业部") or "")]
    if not rows:
        return ""
    W, H, PL, PR, PT, PB = 680, 300, 66, 60, 26, 66
    pw, ph = W - PL - PR, H - PT - PB
    maxv = max(num(r.get("业绩")) for r in rows) or 1
    maxv *= 1.12
    maxp = max(num(r.get("见客签单转化率")) for r in rows) or 1
    maxp *= 1.25
    n = len(rows)
    slot = pw / n
    bw = min(44, slot * 0.52)
    s = ['<svg viewBox="0 0 %d %d" xmlns="http://www.w3.org/2000/svg">'
         '<defs><linearGradient id="gBar" x1="0" y1="0" x2="0" y2="1">'
         '<stop offset="0" stop-color="%s"/><stop offset="1" stop-color="%s"/>'
         '</linearGradient></defs>' % (W, H, C_BAR1, C_BAR2)]
    for i in range(5):
        y = PT + ph - ph * i / 4
        s.append('<line x1="%d" y1="%.1f" x2="%d" y2="%.1f" stroke="%s" stroke-width="1"/>'
                 % (PL, y, PL + pw, y, C_GRID))
        s.append('<text x="%d" y="%.1f" text-anchor="end" font-size="10.5" fill="%s">%s</text>'
                 % (PL - 6, y + 3.5, C_TXT, fmt_wan(maxv * i / 4)))
        s.append('<text x="%d" y="%.1f" text-anchor="start" font-size="10.5" fill="%s">%.0f%%</text>'
                 % (PL + pw + 8, y + 3.5, C_LINE, maxp * i / 4 * 100))
    pts = []
    for i, r in enumerate(rows):
        cx = PL + slot * i + slot / 2
        v = num(r.get("业绩"))
        hgt = ph * v / maxv * 0.95
        if v > 0:
            s.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="4" fill="url(#gBar)"/>'
                     % (cx - bw / 2, PT + ph - hgt, bw, hgt))
            s.append('<text x="%.1f" y="%.1f" text-anchor="middle" font-size="10.5" font-weight="700" '
                     'fill="#4338ca">%s</text>' % (cx, PT + ph - hgt - 5, fmt_wan(v)))
        elif v < 0:
            s.append('<rect x="%.1f" y="%.1f" width="%.1f" height="7" rx="2.5" fill="#e11d48"/>'
                     % (cx - bw / 2, PT + ph + 1, bw))
            s.append('<text x="%.1f" y="%.1f" text-anchor="middle" font-size="10" font-weight="700" '
                     'fill="#e11d48">%s</text>' % (cx, PT + ph + 17, fmt_wan(v)))
        p = num(r.get("见客签单转化率"))
        pts.append((cx, PT + ph - ph * (p / maxp)))
    s.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="2.5" '
             'stroke-linejoin="round"/>' % (" ".join("%.1f,%.1f" % pt for pt in pts), C_LINE))
    for cx, py in pts:
        s.append('<circle cx="%.1f" cy="%.1f" r="4" fill="#fff" stroke="%s" stroke-width="2"/>'
                 % (cx, py, C_LINE))
    for i, r in enumerate(rows):
        cx = PL + slot * i + slot / 2
        nm = str(r.get("事业部") or "")
        nm = nm.replace("", "").replace("事业部", "")
        s.append('<text x="%.1f" y="%.1f" text-anchor="middle" font-size="10.5" fill="%s" '
                 'transform="rotate(-24 %.1f %.1f)">%s</text>'
                 % (cx, H - PB + 16, C_TXT, cx, H - PB + 16, _esc(nm)))
    s.append('<text x="%.1f" y="%.1f" text-anchor="middle" font-size="10.5" fill="%s">各事业部</text>'
             % (PL + pw / 2, H - 4, C_TXT))
    s.append('</svg>')
    return "".join(s)


def svg_donut(biz):
    rows = [r for r in biz if num(r.get("业绩")) > 0
            and "合计" not in str(r.get("事业部") or "")]
    if not rows:
        return ""
    total = sum(num(r.get("业绩")) for r in rows)
    cx, cy, R, R0 = 150, 150, 88, 54
    s = ['<svg viewBox="0 0 680 300" xmlns="http://www.w3.org/2000/svg">']
    ang = -90.0
    for i, r in enumerate(rows):
        v = num(r.get("业绩"))
        sweep = v / total * 360.0
        a0, a1 = ang, ang + sweep
        large = 1 if sweep > 180 else 0
        import math
        x0, y0 = cx + R * math.cos(math.radians(a0)), cy + R * math.sin(math.radians(a0))
        x1, y1 = cx + R * math.cos(math.radians(a1)), cy + R * math.sin(math.radians(a1))
        x2, y2 = cx + R0 * math.cos(math.radians(a1)), cy + R0 * math.sin(math.radians(a1))
        x3, y3 = cx + R0 * math.cos(math.radians(a0)), cy + R0 * math.sin(math.radians(a0))
        s.append('<path d="M %.1f %.1f A %.1f %.1f 0 %d 1 %.1f %.1f L %.1f %.1f A %.1f %.1f 0 %d 0 '
                 '%.1f %.1f Z" fill="%s" stroke="#fff" stroke-width="1.5"/>'
                 % (x0, y0, R, R, large, x1, y1, x2, y2, R0, R0, large, x3, y3,
                    PIE_COLORS[i % len(PIE_COLORS)]))
        ang = a1
    s.append('<text x="%d" y="%d" text-anchor="middle" font-size="11" fill="%s">业绩总量</text>'
             % (cx, cy - 8, C_TXT))
    s.append('<text x="%d" y="%d" text-anchor="middle" font-size="16" font-weight="800" fill="#312e81">%s</text>'
             % (cx, cy + 14, fmt_wan(total)))
    ly = 40
    for i, r in enumerate(rows):
        v = num(r.get("业绩"))
        nm = str(r.get("事业部") or "").replace("", "").replace("事业部", "")
        s.append('<rect x="330" y="%d" width="11" height="11" rx="2.5" fill="%s"/>'
                 % (ly - 9, PIE_COLORS[i % len(PIE_COLORS)]))
        s.append('<text x="348" y="%d" font-size="11.5" fill="#374151">%s</text>' % (ly, _esc(nm)))
        s.append('<text x="660" y="%d" text-anchor="end" font-size="11.5" font-weight="700" '
                 'fill="#374151">%s · %.1f%%</text>' % (ly, fmt_wan(v), v / total * 100))
        ly += 26
    s.append('</svg>')
    return "".join(s)


def svg_hbar(depts):
    rows = [r for r in depts if "合计" not in str(r.get("事业部") or "") and str(r.get("部门") or "")]
    rows = sorted(rows, key=lambda r: -num(r.get("业绩")))[:10]
    if not rows:
        return ""
    maxv = max(num(r.get("业绩")) for r in rows) or 1
    rowh = 30
    W = 680
    H = 30 + rowh * len(rows)
    LX, BX, BW = 158, 166, W - 166 - 92
    s = ['<svg viewBox="0 0 %d %d" xmlns="http://www.w3.org/2000/svg">' % (W, H)]
    for i, r in enumerate(rows):
        y = 18 + i * rowh
        v = num(r.get("业绩"))
        w = BW * v / maxv if v > 0 else 0
        nm = str(r.get("部门") or "")
        nm = nm.replace("", "").replace("业务", "")
        s.append('<text x="%d" y="%.1f" text-anchor="end" font-size="11" fill="#374151">%s</text>'
                 % (LX - 8, y + 14, _esc(nm)))
        s.append('<rect x="%d" y="%.1f" width="%.1f" height="17" rx="4" fill="url(#gH)"/>'
                 % (BX, y + 3, max(w, 1)))
        s.append('<text x="%.1f" y="%.1f" text-anchor="start" font-size="11" font-weight="700" '
                 'fill="#4338ca">%s</text>' % (BX + max(w, 1) + 6, y + 15.5, fmt_wan(v)))
    s.append('<defs><linearGradient id="gH" x1="0" y1="0" x2="1" y2="0">'
             '<stop offset="0" stop-color="%s"/><stop offset="1" stop-color="%s"/>'
             '</linearGradient></defs></svg>' % (C_BAR2, C_BAR1))
    return "".join(s)


# ---- v17：分公司资源明细（主看板口径：网销漏斗 + LTC 近30天） ----
RES_AGG_KEYS = ("total", "valid", "mgr", "wx", "seen", "seenSign", "sign",
                "hand", "handDen", "ltRecv", "ltExec")
RES_TG = {"经理分配占比": 0.60, "资源加微率": 0.85, "见客转化率": 0.30}
RES_NET = {}
for _e in (RY.get("employees") or []):
    if _e.get("funnel_source") != "net":
        continue
    _f, _l = (_e.get("funnel") or {}), (_e.get("ltc") or {})
    RES_NET[str(_e.get("姓名") or "").strip()] = {
        "total": num(_f.get("资源总量")), "valid": num(_f.get("有效资源量")),
        "mgr": num(_f.get("经理手动分配资源")), "wx": num(_f.get("资源加微量")),
        "seen": num(_f.get("接收资源30天内见客量")),
        "seenSign": num(_f.get("接收资源30天内见客后签单量")),
        "sign": num(_f.get("付费客30天签单量")),
        "hand": num(_f.get("付费客7天后未添加微信资源移交量")),
        "handDen": num(_f.get("接收资源7天内未加上微的量")),
        "ltRecv": num(_l.get("近30天移交接收资源量")),
        "ltExec": num(_l.get("近30天移交接收资源执行量")),
    }


def _res_sum(names):
    a = dict((k, 0.0) for k in RES_AGG_KEYS)
    hit = 0
    for _n in names:
        x = RES_NET.get(_n)
        if not x:
            continue
        hit += 1
        for _k in RES_AGG_KEYS:
            a[_k] += (x.get(_k) or 0)
    return a, hit


def _res_rate(x, y):
    return (x / y) if y else None


def _res_metrics(a, hit):
    ks = ("有效资源量", "资源加微量", "经理分配占比", "资源加微率", "付费客30天签单转化率",
          "见客转化率", "付费客7天未加微移交量", "付费客7天未加微移交率", "移交接收执行率")
    if not hit:
        return dict((k, None) for k in ks)
    return {
        "有效资源量": a["valid"], "资源加微量": a["wx"],
        "经理分配占比": _res_rate(a["mgr"], a["valid"]),
        "资源加微率": _res_rate(a["wx"], a["valid"]),
        "付费客30天签单转化率": _res_rate(a["sign"], a["valid"]),
        "见客转化率": _res_rate(a["seenSign"], a["seen"]),
        "付费客7天未加微移交量": a["hand"],
        "付费客7天未加微移交率": _res_rate(a["hand"], a["handDen"]),
        "移交接收执行率": _res_rate(a["ltExec"], a["ltRecv"]),
    }


RES_COLS = ["名称", "人数", "有效资源量", "资源加微量", "经理分配占比", "资源加微率",
            "付费客30天签单转化率", "见客转化率", "付费客7天未加微移交量",
            "付费客7天未加微移交率", "移交接收执行率"]
RES_PCT = {"经理分配占比", "资源加微率", "付费客30天签单转化率", "见客转化率",
           "付费客7天未加微移交率", "移交接收执行率"}
RES_METRIC_COLS = RES_COLS[2:]


def _fi(v):
    return "—" if v is None else "{:,.0f}".format(round(v))


def _fp(v):
    return "—" if v is None else "{:.1f}%".format(v * 100)


def _res_row(name, lvl, names):
    a, hit = _res_sum(names)
    m = _res_metrics(a, hit)
    _cnt = str(len(names))
    if lvl < 3 and hit != len(names):
        _cnt = "%d（有数据 %d）" % (len(names), hit)   # v17：区分全表人数与有资源数据人数
    cells = [name, ("" if lvl == 3 else _cnt)]
    for _c in RES_METRIC_COLS:
        cells.append(_fp(m[_c]) if _c in RES_PCT else _fi(m[_c]))
    reds = []
    for _c in RES_METRIC_COLS:
        _v = m[_c]
        if _c in RES_TG and _v is not None and _v < RES_TG[_c]:
            reds.append(RES_COLS.index(_c))
    return cells, reds, hit


RES_ROWS, RES_REDS, RES_LVL, RES_HAS = [], [], [], []
# v19：只纳入「有资源数据」的员工 —— 资源/过程数据全为 0 的人员（含离职）整行剔除
_ro_rows = _dept_override(U_CLEAN)
_res_has = lambda _n: _n in RES_NET
_res_all = [str(r.get("姓名") or "").strip() for r in _ro_rows
            if _res_has(str(r.get("姓名") or "").strip())]
_rr, _rd, _hit = _res_row("🏢 分公司", 0, _res_all)
RES_ROWS.append(_rr)
RES_REDS.append(_rd)
RES_LVL.append(0)
RES_HAS.append(1)

_res_tree = {}
for _r in _ro_rows:
    if not _res_has(str(_r.get("姓名") or "").strip()):
        continue
    _b = str(_r.get("事业部") or "").strip()
    _b = "" if _b in ("0", "None") else _b
    _b = _b or "未分配事业部"
    _d = str(_r.get("部门") or "").strip()
    _d = "" if _d in ("0", "None") else _d
    _d = _d or "未分配部门"
    _res_tree.setdefault(_b, {}).setdefault(_d, []).append(
        str(_r.get("姓名") or "").strip())

for _b, _dmap in _res_tree.items():
    _bnames = []
    for _v in _dmap.values():
        _bnames.extend(_v)
    _rr, _rd, _hit = _res_row(_b, 1, _bnames)
    RES_ROWS.append(_rr)
    RES_REDS.append(_rd)
    RES_LVL.append(1)
    RES_HAS.append(1 if _hit else 0)
    for _d, _names in _dmap.items():
        _rr, _rd, _hit = _res_row(_d, 2, _names)
        RES_ROWS.append(_rr)
        RES_REDS.append(_rd)
        RES_LVL.append(2)
        RES_HAS.append(1 if _hit else 0)
        for _n in _names:
            _rr, _rd, _hit = _res_row(_n, 3, [_n])
            RES_ROWS.append(_rr)
            RES_REDS.append(_rd)
            RES_LVL.append(3)
            RES_HAS.append(1 if _hit else 0)

SHEETS["分公司资源明细"] = {
    "headers": RES_COLS,
    "rows": RES_ROWS,
    "redFlags": RES_REDS,
    "cols": [{"key": h, "type": ("id" if h == "名称" else "num"), "unit": "",
              "std": RES_TG.get(h), "card": None, "isStd": bool(RES_TG.get(h))}
             for h in RES_COLS],
    "tree": RES_LVL,
    "hasData": RES_HAS,
}
print(" 分公司资源明细：%d 行（分公司1 / 事业部%d / 部门 / 员工），网销有数据 %d 人"
      % (len(RES_ROWS), len(_res_tree), len(RES_NET)))

CHARTS = {}
_c = svg_combo(_biz)
_d = svg_donut(_biz)
CHARTS["biz"] = ""
if _c:
    CHARTS["biz"] += '<div class="chart-card"><h4>各事业部 业绩 × 见客签单转化率</h4>' + _c + '</div>'
if _d:
    CHARTS["biz"] += '<div class="chart-card"><h4>业绩占比（有业绩事业部）</h4>' + _d + '</div>'
if CHARTS["biz"]:
    CHARTS["biz"] = '<div class="charts">' + CHARTS["biz"] + '</div>'
_hb = svg_hbar(DEPT_RAW)
if _hb:
    CHARTS["dept"] = '<div class="charts single"><div class="chart-card"><h4>各部门 业绩 Top10</h4>' + _hb + '</div></div>'

MODES = {
    "updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "source": "9月考核月 · 四大模式健康度综合管理表.xlsx（）｜业绩口径：《公司和部门业绩》+《人员业绩汇总》",
    "kj_due_src": KJ_DUE_SRC,
    "charts": CHARTS,
    "groups": [
        {"label": "网销模式", "subTabs": [
            {"key": "team", "label": "团队数据", "sheet": "网销岗团队数据"},
            {"key": "detail", "label": "明细数据", "sheet": "网销岗明细"}]},
        {"label": "电邀模式", "subTabs": [
            {"key": "team", "label": "团队数据", "sheet": "电邀岗团队数据"},
            {"key": "detail", "label": "明细数据", "sheet": "电邀岗明细"}]},
        {"label": "会计模式", "subTabs": [
            {"key": "team", "label": "团队数据", "sheet": "会计岗团队数据"},
            {"key": "detail", "label": "明细数据", "sheet": "会计岗明细"}]},
        {"label": "资源明细", "subTabs": [
            {"key": "res", "label": "资源明细", "sheet": "分公司资源明细"}]},
        {"label": "综合数据", "subTabs": [
            {"key": "biz", "label": "事业部", "sheet": "所有事业部明细数据"},
            {"key": "dept", "label": "部门", "sheet": "所有部门明细数据"},
            {"key": "user", "label": "规划师", "sheet": "所有规划师明细数据"}]},
    ],
    "sheets": SHEETS,
    "buTargets": {
        "metric": "付费客资源签单转化率（近30天）",
        "fallback": 0.10,
        "rows": [{"k": _bk, "l": BIZ_LABEL.get(_bk, _bk), "v": BIZ_THRESH[_bk]}
                 for _bk in BIZ_SHOW],
    },
}

# ----------------------------------------------------------------------------
# 4) AI 数据解析（构建期生成）
# ----------------------------------------------------------------------------

def ai_insights(RY):
    items = []
    net_rows = (RY.get("company_summary") or {}).get("rows") or []
    dy = ((RY.get("key_metrics") or {}).get("dy") or {}).get("metrics") or []
    kj = ((RY.get("key_metrics") or {}).get("kuaiji") or {}).get("metrics") or []
    depts = (RY.get("biz_health") or {}).get("departments") or {}

    def stat(ms):
        return (sum(1 for m in ms if m.get("status") == "pass"),
                sum(1 for m in ms if m.get("status") == "fail"),
                sum(1 for m in ms if m.get("status") == "warn"))

    np_, nf, nw = stat(net_rows)
    dp_, df_, dw_ = stat(dy)
    kp_, kf, kw = stat(kj)
    total = len(net_rows) + len(dy) + len(kj)
    passed = np_ + dp_ + kp_
    items.append({"sev": "info", "title": "整体达标面",
                  "text": "三大模式共 {} 项重点指标，达标 {} 项、预警 {} 项、不达标 {} 项，综合达标率 {}%。".format(
                      total, passed, nw + dw_ + kw, nf + df_ + kf,
                      round(passed * 100.0 / total, 1) if total else 0)})
    lv_map = {}
    for dname, dv in depts.items():
        lv_map.setdefault(dv.get("level") or "—", []).append(dname)
    for lv, sev in (("危险", "bad"), ("预警", "warn")):
        if lv_map.get(lv):
            items.append({"sev": sev, "title": "网销部门健康度",
                          "text": "{}事业部：{}（建议重点辅导）。".format(lv, "、".join(lv_map[lv]))})
    if lv_map.get("良好"):
        items.append({"sev": "good", "title": "网销标杆部门",
                      "text": "健康等级良好：{}，可沉淀打法横向推广。".format("、".join(lv_map["良好"]))})
    for r in [x for x in net_rows if x.get("status") == "fail"][:2]:
        items.append({"sev": "bad", "title": "网销差距",
                      "text": "{}：当前 {}，{}。".format(r.get("name", "—"), r.get("display", "—"),
                                                       r.get("gap_text", "") or "与目标存在差距")})
    dy_fails = [m for m in dy if m.get("status") == "fail"]
    if dy_fails:
        items.append({"sev": "bad", "title": "电邀短板",
                      "text": "15 项指标中 {} 项不达标：{}{}。建议按 0-3 月 / 3 个月以上分层抓见客量与转化。".format(
                          len(dy_fails), "、".join(m.get("name", "") for m in dy_fails[:3]),
                          "等" if len(dy_fails) > 3 else "")})
    for m in [x for x in kj if x.get("status") == "fail"][:2]:
        items.append({"sev": "bad", "title": "会计短板",
                      "text": "{}：当前 {}，目标 {}。".format(m.get("name", "—"), m.get("display", "—"),
                                                       m.get("target_display") or "")})
    neg = [b for b in BUS if isinstance(b.get("月实算业绩"), (int, float)) and b["月实算业绩"] < 0]
    if neg:
        items.append({"sev": "warn", "title": "会计异常提醒",
                      "text": "{}月实算业绩为负（{} 元），多为退款冲抵，需核对签单与退费明细。".format(
                          "、".join(str(b["部门"]) for b in neg),
                          "、".join(str(int(b["月实算业绩"])) for b in neg))})
    return items


AI = ai_insights(RY)

# ----------------------------------------------------------------------------
# 5) HTML 模板（复刻参考件 modes 页）
# ----------------------------------------------------------------------------
CSS_BUNDLE = ""
if os.path.exists(REF_CSS):
    CSS_BUNDLE = open(REF_CSS, encoding="utf-8").read()

TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>四大模式健康度综合管理看板</title>
<style>
/* 主题变量兜底（参考件同款靛紫基调） */
:root{
  --bg:#F7F9FD; --bg-soft:#F1F4FA; --card:#ffffff; --border:#E5E9F2; --border-soft:#EEF1F7;
  --line:#E5E9F2; --line-soft:#EEF1F7;
  --text:#1F2430; --text-sub:#3D4660; --text-weak:#8A93A6;
  --accent:#4F46E5; --primary:#4F46E5; --primary-deep:#4338CA; --primary-light:#6366F1;
  --primary-soft:#EEF0FE; --primary-tint:#F5F6FF;
  --pass:#16A34A; --ok:#16A34A; --fail:#DC2626; --warn:#D97706;
  --radius:12px; --radius-sm:8px;
  --shadow:0 4px 16px rgba(49,46,129,.08); --shadow-sm:0 2px 8px rgba(49,46,129,.06);
  --shadow-card:0 2px 10px rgba(49,46,129,.08); --head-height:64px;
}
/* ====== 参考件 bundle CSS ====== */
__CSS__
/* ====== 本页补充 ====== */
__EXTRA__

    .res-box{margin-top:8px}
    .res-tbl th:first-child,.res-tbl td:first-child{min-width:220px;white-space:nowrap}
    .res-tbl th:nth-child(2),.res-tbl td:nth-child(2){white-space:nowrap}
    .res-bar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin:2px 0 10px}
    .res-n{font-size:12.5px;color:#475569}
    .res-only{font-size:12.5px;color:#334155;display:inline-flex;align-items:center;gap:5px;cursor:pointer}
    .res-tbl tr.res-l0 td{background:#e8eef7;font-weight:700}
    .res-tbl tr.res-l1 td{background:#f2f6fc;font-weight:600}
    .res-tbl tr.res-l2 td{background:#f8fafc}
    .res-tbl tr.res-l3 td{background:#fff}
    .res-tbl tr.res-tr{cursor:pointer}
    .res-tbl tr.res-tr:hover td{background:#eef4fb}
    .res-arw{display:inline-block;width:15px;color:#64748b}
    .res-tbl tr.res-l1 td:first-child{padding-left:18px}
    .res-tbl tr.res-l2 td:first-child{padding-left:34px}
    .res-tbl tr.res-l3 td:first-child{padding-left:50px;color:#475569}
    .res-tbl td.bad{background:#fdeeee!important;color:#c0392b}
    .res-tbl tr.res-na td{color:#94a3b8}
    .res-note{margin-top:10px;font-size:12px;color:#64748b;line-height:1.75}
    .date-pick{display:flex;align-items:center;gap:6px;color:rgba(255,255,255,.8);font-size:12px;white-space:nowrap}
    .date-pick select{background:rgba(255,255,255,.14);color:#fff;border:1px solid rgba(255,255,255,.28);
      border-radius:7px;padding:4px 8px;font-size:12px;cursor:pointer;outline:none}
    .date-pick select option{color:#1e293b;background:#fff}
    .res-latest{font-size:12px;color:#94a3b8;margin-left:6px}
  </style>
</head>
<body class="modes-app" data-page="modes">
<div class="app">
  <header class="header">
    <div class="brand">
      <div class="brand-mark">模式</div>
      <div class="brand-text"><h1>四大模式健康度综合管理看板</h1></div>
      <span class="sub" id="srcNote"></span>
    </div>
    <div class="header-right"><label class="date-pick" id="datePickWrap" style="display:none">历史数据<select id="snapSel" title="切换历史数据快照"><option value="">最新数据</option></select></label><span class="upd-tip" id="updTip"></span></div>
  </header>
  <div class="body">
    <aside class="sidebar">
      <div class="side-title">岗位模式</div>
      <div class="mode-tabs" id="modeTabs"></div>
      <div class="side-title" style="margin-top:14px">汇总视图</div>
      <div class="all-dept-tab" id="allDataTab">综合数据</div>
      <div class="mode-tabs" id="allSubTabs" style="display:none"></div>
      <div class="side-title dept-title" id="deptTitle">部门筛选 <span class="muted" id="filterHint"></span></div>
      <ul class="tree" id="tree"></ul>
    </aside>
    <main class="main" id="main"></main>
  </div>
</div>
<div class="bottom-blocks">
  <div class="bb-card" id="aiBlock"></div>
  <div class="bb-card collapsible" id="caliberBlock"></div>
</div>

<script>
window.MODES_DATA = __MODES__;
window.AI_DATA = __AI__;
window.TARGET = __TARGET__;
window.TARGET_SRC = __SRC__;
</script>
<script>
(function(){
  "use strict";
  var data = window.MODES_DATA, AI = window.AI_DATA || [], TARGET = window.TARGET || [];

  function toNum(v){
    if(v===null||v===undefined) return null;
    if(typeof v==="number") return v;
    var raw=String(v).trim(), isPct=raw.indexOf("%")>-1;
    var s=raw.replace(/[¥,\s]/g,"").replace(/%/g,"");
    if(s===""||s==="—"||s==="-") return null;
    var n=parseFloat(s); if(isNaN(n)) return null;
    return isPct? n/100 : n;
  }
  function fmtCard(val,type){
    if(val==null||isNaN(val)) return "—";
    if(type==="money") return "¥"+Math.round(val).toLocaleString("zh-CN");
    if(type==="pct") return (val*100).toLocaleString("zh-CN",{minimumFractionDigits:2,maximumFractionDigits:2})+"%";
    return val.toLocaleString("zh-CN",{maximumFractionDigits:2});
  }
  function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(c){
    return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c];});}

  /* ===== v12：表头「目标值」行 + 悬停显示目标值 / 当前值 / 差距 ===== */
  function stdText(c){
    if(!c||c.std==null) return "";
    if(c.type==="pct") return "\u2265"+(c.std*100).toFixed(c.std<0.1?1:0)+"%";
    if(c.type==="money") return "\u2265\u00a5"+Math.round(c.std).toLocaleString("zh-CN");
    return "\u2265"+c.std;
  }
  function gapTxt(c,x){
    if(c.type==="pct") return (x*100).toFixed(2)+"%";
    if(c.type==="money") return "\u00a5"+Math.round(x).toLocaleString("zh-CN");
    return String(Math.round(x*100)/100);
  }
  function signTxt(x){ return (x>=0?"+":"")+gapTxt.apply(null,[this, x]); }
  function stdTip(sh,c,i,rows){
    if(!c) return "";
    if(c.buStd){
      var _k=String(c.key||(sh&&sh.headers?sh.headers[i]:"")||"");
      return buTipAll()+"\n口径："+(_k.indexOf("见客签单转化率")>-1
        ? "见客签单转化率＝30 天及时性转化率（剔除三方）；考核按该事业部业态下限判定"
        : "近30天付费客资源签单量 ÷ 近30天付费客资源总量（剔除三方、剔除无效资源）")
        +"\n考核：部门级任一考核指标未达标即判未达标";
    }
    if(c.std==null) return "";
    var t=stdText(c), bi=sh.headers.indexOf("事业部"), vals=[];
    (rows||[]).forEach(function(r){
      if(bi>-1&&String(r[bi]||"").indexOf("\u5408\u8ba1")>-1) return;
      var n=toNum(r[i]); if(n!=null) vals.push(n);
    });
    if(!vals.length) return "\u76ee\u6807\u503c "+t;
    var sum=vals.reduce(function(a,b){return a+b;},0);
    var cur=(c.card==="sum")?sum:sum/vals.length;
    var lab=(c.card==="sum")?"\u5408\u8ba1":"\u5747\u503c";
    var gap=cur-c.std;
    return "\u76ee\u6807\u503c "+t+" \uff5c "+lab+" "+gapTxt(c,cur)+" \uff5c \u5dee\u8ddd "
      +(gap>=0?"+":"")+gapTxt(c,gap);
  }
  function cellTip(sh,i,cell,stdOv){
    var c=sh.cols[i]; if(!c) return "";
    if(c.buStd&&stdOv==null) return buTipAll();
    if(!c.buStd&&c.std==null) return "";
    var base=c.buStd?stdOv:c.std;
    var t=c.buStd?("≥"+pctTxt(base)):stdText(c), n=toNum(cell);
    if(n==null) return "目标值 "+t;
    var gap=n-base;
    return "目标值 "+t+" ｜ 实际 "+gapTxt(c,n)
      +" ｜ 差距 "+(gap>=0?"+":"")+gapTxt(c,gap);
  }

  function stdRowHtml(sh,rows){
    var h='<tr class="std-row">';
    sh.headers.forEach(function(x,i){
      var c=sh.cols[i]||{}, t=c.buStd?"按业态":stdText(c), tip=stdTip(sh,c,i,rows);
      h+='<th class="std-th'+(t?'':' std-none')+'" data-i="'+i+'"'
        +(tip?' title="'+esc(tip)+'"':'')+'>'+(t?esc(t):"\u2014")+'</th>';
    });
    return h+'</tr>';
  }

  var groups=data.groups, sheets=data.sheets;

  /* ===== v14：付费客资源转化率「目标按业态」（工商15/跨境8/法律7/AI 4/交易3） ===== */
  var BU_T=(data.buTargets&&data.buTargets.rows)||[];
  function pctTxt(v){ var p=v*100; return (Math.round(p*10)/10===Math.round(p))?Math.round(p)+"%":p.toFixed(1)+"%"; }
  function buOfRow(sh,row){
    if(!sh||!row) return null;
    var v="", i=sh.headers.indexOf("事业部");
    if(i>-1&&row[i]!=null&&String(row[i]).trim()) v=String(row[i]).trim();
    if(!v){ i=sh.headers.indexOf("部门"); if(i>-1&&row[i]!=null) v=String(row[i]).trim(); }
    if(!v) return null;
    for(var k=0;k<BU_T.length;k++){ if(v.indexOf(BU_T[k].k)>-1) return BU_T[k]; }
    return null;
  }
  function buTargetOfRow(sh,row){ var t=buOfRow(sh,row); return t?t.v:null; }
  function buTxtAll(){ return BU_T.map(function(t){return t.l+" ≥"+pctTxt(t.v);}).join(" ｜ "); }
  function buTipAll(){ return "目标（按业态）："+buTxtAll(); }
  function buMetricLabels(sh){
    var out=[];
    if(!sh||!sh.cols) return out;
    for(var i=0;i<sh.headers.length;i++){
      var c=sh.cols[i];
      if(c&&c.buStd){
        out.push(String(sh.headers[i]).replace(/（[^）]*）/g,"").replace("网销模式","").replace("近30天",""));
      }
    }
    return out;
  }
  function cardBuStd(rows,sh){
    var m={},ks=[];
    (rows||[]).forEach(function(r){ var t=buOfRow(sh,r);
      if(t&&m[t.k]===undefined){ m[t.k]=t.v; ks.push(t.k); } });
    return ks.length===1?{k:ks[0],v:m[ks[0]]}:null;
  }
  var modeGroups=groups.filter(function(g){return g.label!=="综合数据";});
  var allGroup=groups.filter(function(g){return g.label==="综合数据";})[0]||groups[groups.length-1];
  var state={groupIdx:0, all:false, filter:null, allTab:"biz", expanded:{}, panel:{}, resColl:{}, resOnly:false, resInit:false};

  var modeTabs=document.getElementById("modeTabs");
  var allDataTab=document.getElementById("allDataTab");
  var tree=document.getElementById("tree");
  var allSub=document.getElementById("allSubTabs");
  var deptTitle=document.getElementById("deptTitle");

  function currentGroup(){ return state.all? allGroup : modeGroups[state.groupIdx]; }
  function treeSheet(){
    /* v18(09-20)：取「同时含事业部+部门列 且行数最多」的子表构建部门树。
       旧逻辑取第一个命中的子表＝团队聚合表（行数少、部门列大量为空、且漏事业部，
       如网销团队表无 AI 事业部 / 会计树子项错位成事业部名），改为明细表后覆盖完整。 */
    var g=currentGroup(), sh=null, best=-1;
    g.subTabs.forEach(function(sub){
      var s=sheets[sub.sheet]; if(!s) return;
      var tc=getTreeCols(s.headers);
      if(!tc||tc.p<0||tc.c<0) return;
      if(s.rows.length>best){ best=s.rows.length; sh=s; }
    });
    if(!sh) sh=currentSheet();
    return sh;
  }
  function currentSheet(){
    var g=currentGroup(), key=state.all?(state.allTab||"biz"):g.subTabs[0].key, sub=null;
    g.subTabs.forEach(function(s){ if(s.key===key) sub=s; });
    if(!sub) sub=g.subTabs[0];
    return sheets[sub.sheet]||{headers:[],rows:[],redFlags:[],cols:[]};
  }
  function getTreeCols(headers){
    var si=headers.indexOf("事业部"), bi=headers.indexOf("部门");
    if(si>-1&&bi>-1) return {p:si,c:bi};
    if(bi>-1) return {p:-1,c:bi};
    return {p:-1,c:-1};
  }

  /* ===== 左侧模式 tab ===== */
  function renderSidebar(){
    modeTabs.innerHTML="";
    modeGroups.forEach(function(g,i){
      var first=sheets[g.subTabs[0].sheet];
      var count=first?first.rows.length:0;
      var label=g.label.replace("模式","");
      var tab=document.createElement("div");
      tab.className="mode-tab"+(!state.all&&state.groupIdx===i?" active":"");
      tab.innerHTML='<span class="mt-icon">'+esc(label.slice(0,1))+'</span>'
        +'<span class="mt-name">'+esc(label)+'</span>'
        +'<span class="mt-count">'+count+'</span>';
      tab.addEventListener("click",function(){state.all=false;state.groupIdx=i;state.filter=null;renderAll();});
      modeTabs.appendChild(tab);
    });
    allDataTab.className="all-dept-tab"+(state.all?" active":"");
    if(state.all){
      allSub.style.display="flex"; allSub.innerHTML="";
      var subLabels={biz:"事业部数据",dept:"部门数据",user:"个人明细数据"};
      allGroup.subTabs.forEach(function(sub){
        var t=document.createElement("div");
        t.className="mode-tab"+((state.allTab||"biz")===sub.key?" active":"");
        t.innerHTML='<span class="mt-name">'+esc(subLabels[sub.key]||sub.label)+'</span>';
        t.addEventListener("click",function(){state.allTab=sub.key;renderAll();});
        allSub.appendChild(t);
      });
      deptTitle.style.display="none"; tree.style.display="none";
    }else{
      allSub.style.display="none"; allSub.innerHTML="";
      deptTitle.style.display=""; tree.style.display="";
    }
  }
  allDataTab.addEventListener("click",function(){state.all=true;state.groupIdx=0;state.filter=null;renderAll();});

  /* ===== 部门筛选树 ===== */
  function renderTree(){
    var sh=treeSheet(), tc=getTreeCols(sh.headers);
    tree.innerHTML="";
    var hint=document.getElementById("filterHint");
    if(hint) hint.textContent="";
    /* v18：当前板块无「部门」列（如资源明细按姓名四级树，不走部门筛选）→ 隐藏筛选区 */
    if(tc.c<0){ deptTitle.style.display="none"; tree.style.display="none"; return; }
    deptTitle.style.display=""; tree.style.display="";
    var treeData={}, totalChildren=0;
    if(tc.c>-1){
      sh.rows.forEach(function(r){
        var p=tc.p>-1?(String(r[tc.p]||"").trim()):"全部";
        var c=String(r[tc.c]||"").trim();
        if(!c) return;
        if(!treeData[p]) treeData[p]={children:{}};
        if(!treeData[p].children[c]){treeData[p].children[c]=true;totalChildren++;}
      });
    }
    var liAll=document.createElement("li");
    var divAll=document.createElement("div");
    divAll.className="node all-node"+(state.filter?"":" active");
    divAll.innerHTML='<span class="arrow leaf"></span><span class="node-name">全部部门</span>'
      +'<span class="node-count">'+totalChildren+'</span>';
    divAll.addEventListener("click",function(){state.filter=null;renderAll();});
    liAll.appendChild(divAll); tree.appendChild(liAll);
    if(tc.p<0){
      Object.keys(treeData).forEach(function(p){
        Object.keys(treeData[p].children).forEach(function(c){
          var li=document.createElement("li"), d=document.createElement("div");
          d.className="node child-node"+(state.filter===c?" active":"");
          d.innerHTML='<span class="arrow leaf"></span><span class="node-name">'+esc(c)+'</span>';
          d.addEventListener("click",function(){state.filter=c;renderAll();});
          li.appendChild(d); tree.appendChild(li);
        });
      });
      if(hint) hint.textContent="（"+totalChildren+" 个）";
      return;
    }
    Object.keys(treeData).sort().forEach(function(p){
      var children=Object.keys(treeData[p].children).sort();
      if(children.indexOf(state.filter)>-1) state.expanded[p]=true;
      var isExp=!!state.expanded[p];
      var li=document.createElement("li"); li.className="parent-li";
      var d=document.createElement("div");
      d.className="node parent-node"+(state.filter===p?" active":"");
      d.innerHTML='<span class="arrow '+(isExp?"expanded":"collapsed")+'"></span>'
        +'<span class="node-name">'+esc(p)+'</span>'
        +'<span class="node-count">'+children.length+'</span>';
      d.addEventListener("click",function(e){
        if(e.target&&e.target.classList&&e.target.classList.contains("arrow")){
          state.expanded[p]=!isExp; renderTree(); return;
        }
        state.expanded[p]=true; state.filter=p; renderAll();
      });
      li.appendChild(d); tree.appendChild(li);
      if(isExp){
        children.forEach(function(c){
          var li2=document.createElement("li"), d2=document.createElement("div");
          d2.className="node child-node"+(state.filter===c?" active":"");
          d2.innerHTML='<span class="arrow leaf"></span><span class="node-name">'+esc(c)+'</span>';
          d2.addEventListener("click",function(){state.filter=c;renderAll();});
          li2.appendChild(d2); tree.appendChild(li2);
        });
      }
    });
    if(hint) hint.textContent="（"+totalChildren+" 个）";
  }

  /* ===== KPI 卡 ===== */
  function pickCardRows(rows,headers){
    var si=headers.indexOf("事业部"), bi=headers.indexOf("部门");
    if(si<0) return rows;
    var total=[],biz=[],detail=[];
    rows.forEach(function(r){
      var bv=String(r[si]==null?"":r[si]).trim();
      var dv=bi>-1?String(r[bi]==null?"":r[bi]).trim():"";
      if(bv.indexOf("合计")>-1) total.push(r);
      else if(bi>-1&&bv&&!dv) biz.push(r);
      else detail.push(r);
    });
    if(total.length===1) return total;
    if(bi>-1&&biz.length===1&&detail.length) return biz;
    return detail.length?detail:rows;
  }
  function aggregate(cardCols,rows,totRow){
    return cardCols.map(function(col){
      /* v16：「达标率」是行级判定结果 → 取合计行，不做算术平均（否则与表格合计口径冲突） */
      if(totRow&&col.key==="达标率"){
        var tv=toNum(totRow[col.idx]);
        if(tv!=null) return {col:col,val:tv};
      }
      var vals=[];
      rows.forEach(function(r){
        var n=toNum(r[col.idx]);
        if(n!=null) vals.push(n);
      });
      var val=null;
      if(vals.length){
        val=vals.reduce(function(a,b){return a+b;},0);
        if(col.card!=="sum") val=val/vals.length;
      }
      return {col:col,val:val};
    });
  }
  function renderCards(sh,rows){
    var cardCols=[],tenureIdx=sh.headers.indexOf("司龄月");
    sh.cols.forEach(function(c,i){
      if(!c.card) return;
      cardCols.push({idx:i,card:c.card,type:c.type,std:c.std,isStd:c.isStd,key:c.key,buStd:c.buStd});
    });
    if(!cardCols.length&&tenureIdx<0) return "";
    rows=pickCardRows(rows,sh.headers);
    var totRow=null;
    (function(){
      var xi=sh.headers.indexOf("事业部");
      if(xi<0) return;
      sh.rows.forEach(function(r){ if(String(r[xi]||"").indexOf("合计")>-1) totRow=r; });
    })();
    var html='<div class="cards">';
    aggregate(cardCols,rows,totRow).forEach(function(a){
      var meta=a.col, val=a.val, bus=meta.buStd?cardBuStd(rows,sh):null;
      var cls="";
      if(val==null||isNaN(val)) cls="na";
      else if(meta.buStd) cls=(bus!=null)?(val<bus.v?"bad":"ok"):"";
      else if(meta.isStd) cls=val<meta.std?"bad":"ok";
      var stdTxt="";
      if(meta.buStd) stdTxt=(bus!=null)?("按业态 ≥ "+pctTxt(bus.v)):"按业态";
      else if(meta.isStd&&meta.std!=null){
        if(meta.type==="pct") stdTxt="≥ "+(meta.std*100).toFixed(0)+"%";
        else if(meta.type==="money") stdTxt="≥ ¥"+Math.round(meta.std).toLocaleString("zh-CN");
        else stdTxt="≥ "+meta.std;
      }
      html+='<div class="card '+cls+'" title="'+esc(meta.key+(meta.buStd?" ｜ "+buTxtAll():""))+'">'
        +'<div class="k">'+esc(meta.key)+'</div>'
        +'<div class="v">'+fmtCard(val,meta.type)+'</div>'
        +(stdTxt?'<div class="std">标准 '+stdTxt+'</div>':'')
        +'</div>';
    });
    if(tenureIdx>-1){
      var buckets=[0,0,0,0],total=0;
      rows.forEach(function(r){
        var n=toNum(r[tenureIdx]); if(n==null) return;
        total++;
        if(n<=3) buckets[0]++; else if(n<=6) buckets[1]++; else if(n<=12) buckets[2]++; else buckets[3]++;
      });
      var items=[{l:"3个月以内",v:buckets[0],c:"t1"},{l:"3-6个月",v:buckets[1],c:"t2"},
                 {l:"6-12个月",v:buckets[2],c:"t3"},{l:"12个月以上",v:buckets[3],c:"t4"}];
      html+='<div class="card card-tenure"><div class="k">司龄分布<span class="tenure-total">共 '+total+' 人</span></div><div class="tenure-grid">';
      items.forEach(function(it){
        html+='<div class="t-item '+it.c+'"><div class="t-val">'+it.v+'</div><div class="t-label">'+it.l+'</div></div>';
      });
      html+='</div></div>';
    }
    html+='</div>';
    return html;
  }

  /* ===== 表格 ===== */
  function filtRows(sh){
    var tc=getTreeCols(sh.headers);
    return sh.rows.filter(function(r){
      if(!state.filter) return true;
      if(tc.p>-1&&String(r[tc.p]||"").trim()===state.filter) return true;
      if(tc.c>-1&&String(r[tc.c]||"").trim()===state.filter) return true;
      return false;
    });
  }
  function tableHtml(sh,rows){
    var tc=getTreeCols(sh.headers);
    var filtered=filtRows(sh);
    var numFlags={};
    sh.cols.forEach(function(c,i){ if(c.type!=="id") numFlags[i]=true; });
    var html='<div class="tbl-wrap"><table><thead><tr>';
    sh.headers.forEach(function(h,i){
      html+='<th data-i="'+i+'">'+esc(h)+'</th>';
    });
    html+='</tr>'+stdRowHtml(sh,filtered)+'</thead><tbody>';
    if(!filtered.length){
      html+='<tr><td colspan="'+sh.headers.length+'" style="text-align:center;color:var(--text-weak);padding:26px">当前筛选下暂无数据</td></tr>';
    }
    filtered.forEach(function(r){
      var redSet={};
      (sh.redFlags[sh.rows.indexOf(r)]||[]).forEach(function(ci){redSet[ci]=true;});
      var _bu=buTargetOfRow(sh,r);
      var rowCls="";
      var si=sh.headers.indexOf("事业部"), bi=sh.headers.indexOf("部门");
      if(si>-1){
        var bv=String(r[si]||"").trim(), dv=bi>-1?String(r[bi]||"").trim():"";
        if(bv.indexOf("合计")>-1) rowCls=" row-total";
        else if(bv&&!dv) rowCls=" row-biz";
      }
      html+='<tr class="'+rowCls.trim()+'">';
      r.forEach(function(cell,i){
        if(isExamCol(sh,i)){ html+=examTd(sh,i,cell); return; }
        var empty=cell==null||String(cell).trim()===""||String(cell).trim()==="—";
        var cm=sh.cols[i];
        var _ov=(cm&&cm.buStd)?_bu:null;
        var _base=(cm&&cm.buStd)?_ov:(cm?cm.std:null);
        var cls=(numFlags[i]&&!empty?"num ":"")+(i===0?"t-left ":"");
        if(redSet[i]) cls+="bad ";
        else if(empty) cls+="na ";
        else if(cm&&cm.isStd&&_base!=null&&toNum(cell)!=null) cls+="ok ";
        var _tip=empty?"":cellTip(sh,i,cell,_ov);
        html+='<td class="'+cls.trim()+'"'+(_tip?' title="'+esc(_tip)+'"':'')+'>'
          +(empty?"":esc(cell))+'</td>';
      });
      html+='</tr>';
    });
    html+='</tbody></table></div>';
    return html;
  }
  function bindSort(tbl){
    if(!tbl) return;
    Array.prototype.forEach.call(tbl.querySelectorAll("thead tr:first-child th"),function(th){
      th.addEventListener("click",function(){
        var i=parseInt(th.getAttribute("data-i"),10);
        if(tbl._sortCol===i) tbl._sortAsc=!tbl._sortAsc;
        else{tbl._sortCol=i;tbl._sortAsc=true;}
        var body=tbl.querySelector("tbody");
        var rows=Array.prototype.slice.call(body.querySelectorAll("tr"));
        var cmp=function(a,b){
          var av=a.cells[i]?a.cells[i].textContent:"", bv=b.cells[i]?b.cells[i].textContent:"";
          var an=toNum(av),bn=toNum(bv);
          var r=(an!=null&&bn!=null)?(an-bn):String(av).localeCompare(String(bv),"zh-CN");
          return tbl._sortAsc?r:-r;
        };
        if(tbl._grouped){                       /* v9：分组表 ⇒ 组内排序，组头保持原位 */
          var gs=[],cur=null;
          rows.forEach(function(tr){
            if(tr.classList.contains("dept-group")){ cur={head:tr,items:[]}; gs.push(cur); }
            else if(cur) cur.items.push(tr);
          });
          gs.forEach(function(g){
            g.items.sort(cmp);
            body.appendChild(g.head);
            g.items.forEach(function(r){ body.appendChild(r); });
          });
        }else{
          rows.sort(cmp);
          rows.forEach(function(r){body.appendChild(r);});
        }
        Array.prototype.forEach.call(tbl.querySelectorAll("thead tr:first-child th"),function(x){x.classList.remove("sort-asc","sort-desc");});
        th.classList.add(tbl._sortAsc?"sort-asc":"sort-desc");
      });
    });
  }

  /* ===== 冻结列：工号 / 姓名 / 事业部（无则首列），横向滚动时固定 ===== */
  var STK_KEYS=["工号","姓名","事业部"];
  function applySticky(tbl){
    if(!tbl) return;
    var tr0=tbl.querySelector("thead tr");
    var ths=tr0?tr0.querySelectorAll("th"):tbl.querySelectorAll("thead th");
    if(!ths.length) return;
    var i, prefix=[];
    for(i=0;i<ths.length&&i<3;i++){
      var t=ths[i].textContent.replace(/\s/g,"");
      if(STK_KEYS.indexOf(t)>-1) prefix.push(i); else break;
    }
    if(!prefix.length) prefix=[0];          /* 兜底：冻结首列 */
    var left=0;
    prefix.forEach(function(ci,order){
      var w=ths[ci].getBoundingClientRect().width;
      var last=(order===prefix.length-1);
      Array.prototype.forEach.call(tbl.rows,function(tr){
        if(tr.classList&&tr.classList.contains("dept-group")) return;   /* v9：分组行不参与冻结列 */
        var c=tr.cells[ci]; if(!c) return;
        c.classList.add("stk");
        if(last) c.classList.add("stk-last");
        c.style.left=left+"px";
      });
      left+=w;
    });
  }
  function applyStickyAll(){
    Array.prototype.forEach.call(document.querySelectorAll(".main table"),applySticky);
  }
  window.addEventListener("resize",function(){applyStickyAll();});

  /* ===== v9：明细数据（个人）＝ 部门筛选 + 按部门分组折叠（09-20 用户要求） ===== */
  var TITLE_MAP={team:"团队数据（事业部）",detail:"明细数据（个人）",
                 biz:"事业部明细数据",dept:"部门明细数据",user:"个人明细数据"};
  var DETAIL_KEYS=["detail","user"];
  function isDetailPanel(key){ return DETAIL_KEYS.indexOf(key)>-1; }
  function pState(key){
    if(!state.panel[key]) state.panel[key]={sel:null,collapsed:{}};
    return state.panel[key];
  }
  function deptOf(sh,r){
    var i=sh.headers.indexOf("部门");
    var v=i>-1?String(r[i]==null?"":r[i]).trim():"";
    if(v==="0"||v==="—"||v==="-") v="";
    return v||"未分配部门";
  }
  function isExamCol(sh,i){ var h=(sh&&sh.headers&&sh.headers[i])||""; return h==="考核结果"||h==="考核模式"; }
  function examInner(cell){
    var v=cell==null?"":String(cell).trim();
    if(v===""||v==="—"||v==="-") return "";
    var cls=(v==="达标")?"ex-ok":((v==="未达标")?"ex-no":"");
    return cls?('<span class="ex-badge '+cls+'">'+esc(v)+'</span>'):esc(v);
  }
  function examTd(sh,i,cell){
    var h=(sh&&sh.headers&&sh.headers[i])||"";
    var tip=(h==="考核结果")
      ? "考核结果：该行按所属模式考核（网销·付费客资源转化率按业态 / 电邀·人均月度高价值见客量 / 会计·到期续费率+非正常流失率），任一考核指标未达标即未达标"
      : "考核模式：该部门/事业部适用的考核口径";
    return '<td class="ex-td" title="'+esc(tip)+'">'+examInner(cell)+'</td>';
  }
  function cellTd(sh,cell,i,numFlags,redSet,buT){
    if(isExamCol(sh,i)) return examTd(sh,i,cell);
    var empty=cell==null||String(cell).trim()===""||String(cell).trim()==="—";
    var cm=sh.cols[i];
    var _ov=(cm&&cm.buStd)?buT:null;
    var _base=(cm&&cm.buStd)?_ov:(cm?cm.std:null);
    var cls=(numFlags[i]&&!empty?"num ":"")+(i===0?"t-left ":"");
    if(redSet[i]) cls+="bad ";
    else if(empty) cls+="na ";
    else if(cm&&cm.isStd&&_base!=null&&toNum(cell)!=null) cls+="ok ";
    var tip=empty?"":cellTip(sh,i,cell,_ov);
    return '<td class="'+cls.trim()+'"'+(tip?' title="'+esc(tip)+'"':'')+'>'
      +(empty?"":esc(cell))+'</td>';
  }

  function detailPanelHtml(sh,rows,key){
    var ps=pState(key);
    /* v16：部门分组头的考核结果取自「部门明细数据」（与部门表一致） */
    var DK_EXAM={};
    (function(){
      var dsh=sheets["所有部门明细数据"];
      if(!dsh) return;
      var di=dsh.headers.indexOf("部门"), ri=dsh.headers.indexOf("考核结果"),
          mi=dsh.headers.indexOf("考核模式"), ti=dsh.headers.indexOf("达标率");
      if(di<0||ri<0) return;
      dsh.rows.forEach(function(r){
        var dv=String(r[di]==null?"":r[di]).trim();
        if(!dv||dv==="0"||dv==="—") return;
        DK_EXAM[dv]={res:String(r[ri]==null?"":r[ri]).trim(),
                     mode:mi>-1?String(r[mi]==null?"":r[mi]).trim():"",
                     rate:ti>-1?String(r[ti]==null?"":r[ti]).trim():""};
      });
    })();
    var order=[],cnt={};
    rows.forEach(function(r){
      var d=deptOf(sh,r);
      if(cnt[d]===undefined){cnt[d]=0;order.push(d);}
      cnt[d]++;
    });
    var sel=ps.sel, hasSel=!!(sel&&sel.length);
    var isOn=function(d){ return !hasSel||sel.indexOf(d)>-1; };
    var shown=rows.filter(function(r){return isOn(deptOf(sh,r));});
    var bar='<div class="dt-bar"><span class="dt-lb">部门筛选</span>'
      +'<span class="dt-chip'+(hasSel?"":" on")+'" data-dt-all="1">全部部门<span class="n">'+rows.length+'</span></span>';
    order.forEach(function(d){
      bar+='<span class="dt-chip'+(hasSel&&isOn(d)?" on":"")+'" data-dt="'+esc(d)+'">'
        +esc(d)+'<span class="n">'+cnt[d]+'</span></span>';
    });
    bar+='<span class="dt-sum">已显示 '+shown.length+' / '+rows.length+' 人'
      +(hasSel?'<a href="javascript:void(0)" class="dt-clr" data-dt-clr="1">清除筛选</a>':'')+'</span></div>';
    var numFlags={};
    sh.cols.forEach(function(c,i){ if(c.type!=="id") numFlags[i]=true; });
    var html=bar+'<div class="tbl-wrap"><table data-detail="1"><thead><tr>';
    sh.headers.forEach(function(h,i){ html+='<th data-i="'+i+'">'+esc(h)+'</th>'; });
    html+='</tr>'+stdRowHtml(sh,shown)+'</thead><tbody>';
    if(!shown.length){
      html+='<tr><td colspan="'+sh.headers.length+'" style="text-align:center;color:var(--text-weak);padding:26px">当前筛选下暂无数据</td></tr>';
    }
    order.forEach(function(d){
      var mem=shown.filter(function(r){return deptOf(sh,r)===d;});
      if(!mem.length) return;
      var col=!!ps.collapsed[d];
      var _ex=DK_EXAM[d], _bdg="";
      if(_ex&&_ex.res){
        var _isOk=(_ex.res==="达标");
        _bdg='<span class="dg-badge '+(_isOk?"is-ok":"is-no")+'"'
          +' title="考核结果：'+esc(_ex.res)+' ｜ 模式 '+esc(_ex.mode||"—")
          +' ｜ 达标率 '+esc(_ex.rate||"—")+'（任一考核指标未达标即未达标）">'
          +'<i class="'+(_isOk?"ok":"no")+'">'+esc(_ex.res)+'</i></span>';
      }
      /* v13：未分配部门的成员是各事业部负责人，姓名直接写在可展开的分组头上 */
      var _nm="";
      if(d==="未分配部门"){
        var _ni=sh.headers.indexOf("姓名");
        if(_ni>-1&&mem.length){
          _nm='<span class="dg-names">'+mem.map(function(r){return esc(r[_ni]);}).join(" · ")+'</span>';
        }
      }
            /* v14：该组的付费客资源转化率目标（按业态；同组跨业态时给区间） */
      var _bt="";
      buMetricLabels(sh).forEach(function(_lab){
        var _bset={},_bk=[];
        mem.forEach(function(r){ var t=buOfRow(sh,r);
          if(t&&_bset[t.k]===undefined){ _bset[t.k]=t.v; _bk.push(t.k); } });
        if(_bk.length===1){
          _bt+='<span class="dg-std" title="'+esc(_lab+" 目标 ≥"+pctTxt(_bset[_bk[0]])
            +"（"+_bk[0]+"事业部）｜"+buTipAll())+'">'+esc(_lab)+" ≥"+pctTxt(_bset[_bk[0]])+'</span>';
        }else if(_bk.length>1){
          var _vs=_bk.map(function(k){return _bset[k];});
          _bt+='<span class="dg-std" title="'+esc(_lab+" 目标按业态 ｜ "+buTipAll())+'">'
            +esc(_lab)+" 按业态 ≥"+pctTxt(Math.min.apply(null,_vs))+"~"+pctTxt(Math.max.apply(null,_vs))+'</span>';
        }
      });
      html+='<tr class="dept-group" data-dept="'+esc(d)+'"><td colspan="'+sh.headers.length+'">'
        +'<span class="dg-ar">'+(col?"▸":"▾")+'</span>'+esc(d)
        +'<span class="dg-n">'+mem.length+' 人</span>'+_nm+_bt+_bdg+'</td></tr>';
      mem.forEach(function(r){
        var redSet={};
        (sh.redFlags[sh.rows.indexOf(r)]||[]).forEach(function(ci){redSet[ci]=true;});
        html+='<tr class="row-grp'+(col?" row-dept-hide":"")+'">';
        var _bu=buTargetOfRow(sh,r);
        r.forEach(function(cell,i){ html+=cellTd(sh,cell,i,numFlags,redSet,_bu); });
        html+='</tr>';
      });
    });
    html+='</tbody></table></div>';
    return html;
  }
  function renderPanel(key){
    var g=currentGroup();
    var subs=state.all?allGroup.subTabs:g.subTabs;
    var sub=subs.filter(function(s){return s.key===key;})[0];
    if(!sub) return;
    var sh=sheets[sub.sheet];
    var el=document.getElementById("main").querySelector("#panel_"+key);
    if(!el||!sh) return;
    var ps=pState(key);
    el.innerHTML='<div class="panel-head"><h2>'+esc(TITLE_MAP[key]||sub.label)+'</h2></div>';
    var ch=(data.charts||{})[key];
    if(state.all&&ch) el.insertAdjacentHTML("beforeend",ch);
    el.insertAdjacentHTML("beforeend",detailPanelHtml(sh,filtRows(sh),key));
    var tbl=el.querySelector("table");
    if(tbl) tbl._grouped=true;
    Array.prototype.forEach.call(el.querySelectorAll(".dt-chip"),function(chip){
      chip.addEventListener("click",function(){
        if(chip.getAttribute("data-dt-all")){ ps.sel=null; renderPanel(key); return; }
        var d=chip.getAttribute("data-dt");
        var cur=(ps.sel&&ps.sel.length)?ps.sel.slice():[d];
        if(ps.sel&&ps.sel.length){
          var i=cur.indexOf(d);
          if(i>-1) cur.splice(i,1); else cur.push(d);
        }
        ps.sel=cur.length?cur:null;
        renderPanel(key);
      });
    });
    var clr=el.querySelector(".dt-clr");
    if(clr) clr.addEventListener("click",function(){ ps.sel=null; renderPanel(key); });
    Array.prototype.forEach.call(el.querySelectorAll("tr.dept-group"),function(tr){
      tr.addEventListener("click",function(){
        var d=tr.getAttribute("data-dept");
        ps.collapsed[d]=!ps.collapsed[d];
        renderPanel(key);
      });
    });
    bindSort(tbl);
    applySticky(tbl);
  }

  function jkBadgeHtml(){
    var bsh=sheets["所有事业部明细数据"];
    if(!bsh) return "";
    var xi=bsh.headers.indexOf("事业部"), ri=bsh.headers.indexOf("考核结果"),
        mi=bsh.headers.indexOf("考核模式"), ti=bsh.headers.indexOf("达标率");
    if(xi<0||ri<0) return "";
    var tot=null;
    bsh.rows.forEach(function(r){
      if(String(r[xi]||"").indexOf("合计")>-1) tot=r;
    });
    if(!tot) return "";
    var res=String(tot[ri]==null?"":tot[ri]).trim();
    return '<div class="jk-badge"><span class="jbb-t">分公司考核（全部考核指标）</span>'
      +'<span class="jbb '+(res==="达标"?"ok":"no")+'">'+esc(res||"—")+'</span>'
      +'<span class="jbb-note">模式 '+esc(String(tot[mi]==null?"—":tot[mi]))
      +' · 达标率 '+esc(String(tot[ti]==null?"—":tot[ti]))
      +'（任一考核指标未达标即未达标）</span></div>';
  }
  /* ===== v17：分公司资源明细（四级树：分公司 → 事业部 → 部门 → 员工） ===== */
  function resChildren(lv,i){
    var out=[];
    for(var j=i+1;j<lv.length;j++){ if(lv[j]<=lv[i]) break; out.push(j); }
    return out;
  }
  function renderRes(sheetName,el){
    var sh=sheets[sheetName];
    if(!sh){ return; }
    var lv=sh.tree||[], hd=sh.hasData||[], rows=sh.rows||[], reds=sh.redFlags||[];
    if(!state.resInit){
      for(var a=0;a<lv.length;a++){ if(lv[a]===2) state.resColl[a]=true; }  /* 部门级默认收起 */
      state.resInit=true;
    }
    var net=0, empN=0;
    for(var b=0;b<lv.length;b++){ if(lv[b]===3){ empN++; if(hd[b]) net++; } }
    var box=document.createElement("div");
    box.className="res-box";
    var head='<tr>';
    sh.headers.forEach(function(h,ix){ head+='<th>'+esc(h)+'</th>'; });
    head+='</tr><tr class="std-row">';
    sh.headers.forEach(function(h,ix){
      var t=(ix===4?"目标 60%":(ix===5?"目标 85%":(ix===7?"目标 30%":"—")));
      head+='<th class="std-th'+(t==="—"?" std-none":"")+'">'+esc(t)+'</th>';
    });
    head+='</tr>';
    box.innerHTML='<div class="res-bar">'
      +'<span class="res-n">数据源：网销漏斗 sheet（近30天）+ LTC流程数据 · <b>'+net+'</b>'
      +' 名有资源数据的员工 · 点击分组行展开 / 收起</span>'
      +'<span class="res-latest" title="资源明细不参与历史快照，始终展示最新数据">· 该板块不随历史日期切换，恒为最新</span></div>'
      +'<div class="tbl-wrap"><table class="res-tbl"><thead>'+head
      +'</thead><tbody></tbody></table></div>'
      +'<div class="res-note"><b>口径说明：</b>① 有效资源量 / 资源加微量 / 付费客7天未加微移交量'
      +'取自「网销漏斗」sheet 近30天明细，按员工汇总后重算；'
      +'② 经理分配占比＝经理手动分配资源 ÷ 有效资源量（目标 60%）；'
      +'资源加微率＝资源加微量 ÷ 有效资源量（目标 85%）；'
      +'付费客30天签单转化率＝付费客30天签单量 ÷ 有效资源量；'
      +'见客转化率＝接收资源30天内见客后签单量 ÷ 接收资源30天内见客量（目标 30%）；'
      +'③ 付费客7天未加微移交率＝移交量 ÷ 接收资源7天内未加上微的量；'
      +'④ 移交接收执行率＝LTC 近30天移交接收资源执行量 ÷ 接收资源量；'
      +'⑤ 比率一律「先汇总再相除」，不是员工比率的简单平均；'
      +'⑥ <b style="color:#c0392b">未达标＝红字</b>；⑦ 非网销岗位无资源数据，显示「—」。</div>';
    el.appendChild(box);
    var tb=box.querySelector("tbody");
    function visible(){
      var out=[], stop=-1;
      for(var i=0;i<lv.length;i++){
        if(i<=stop) continue;
        if(state.resOnly && !hd[i] && lv[i]>=1){
          var last=i;
          for(var j=i+1;j<lv.length;j++){ if(lv[j]<=lv[i]) break; last=j-1; }
          stop=(last<i)?i:last;
          continue;
        }
        out.push(i);
        if(state.resColl[i]){
          var ch=resChildren(lv,i);
          if(ch.length) stop=ch[ch.length-1];
        }
      }
      return out;
    }
    function paint(){
      var vis=visible(), html="";
      vis.forEach(function(i){
        var ch=resChildren(lv,i);
        html+='<tr class="res-tr res-l'+lv[i]+(hd[i]?"":" res-na")+'" data-i="'+i+'">';
        (rows[i]||[]).forEach(function(c,ix){
          var bad=(reds[i]||[]).indexOf(ix)>-1;
          var arw='';
          if(ix===0&&ch.length) arw='<span class="res-arw">'+(state.resColl[i]?"▸":"▾")+'</span>';
          html+='<td class="'+(ix===0?"t-left ":"")+(ix>1?"num ":"")+(bad?"bad ":"")+'">'
            +arw+esc(c==null?"":c)+'</td>';
        });
        html+='</tr>';
      });
      tb.innerHTML=html;
      Array.prototype.forEach.call(tb.querySelectorAll("tr.res-tr"),function(tr){
        tr.addEventListener("click",function(){
          var i=parseInt(tr.getAttribute("data-i"),10);
          if(!resChildren(lv,i).length) return;
          state.resColl[i]=!state.resColl[i];
          paint();
        });
      });
    }
    paint();
    var chk=box.querySelector("#resOnlyChk");
    if(chk) chk.addEventListener("change",function(){ state.resOnly=chk.checked; paint(); });
  }

  function renderMain(){
    var g=currentGroup();
    /* 综合数据：按左侧子页签只显示对应子表；模式视图：团队+明细两张表同屏（与参考件一致） */
    var activeSubs = state.all
      ? allGroup.subTabs.filter(function(s){return s.key===(state.allTab||"biz");})
      : g.subTabs;
    var titleMap=TITLE_MAP;
    var teamSheet=sheets[activeSubs[0].sheet];
    var cardRows=teamSheet?filtRows(teamSheet):[];
    if(state.all&&teamSheet){
      /* 综合 KPI 卡只按有效事业部计算（剔除合计行与零业绩/零见客支撑行） */
      var bi=teamSheet.headers.indexOf("事业部"), pi=teamSheet.headers.indexOf("业绩"),
          vi=teamSheet.headers.indexOf("所有见客");
      cardRows=cardRows.filter(function(r){
        if(bi>-1&&String(r[bi]).indexOf("合计")>-1) return false;
        var pv=pi>-1?(toNum(r[pi])||0):0, vv=vi>-1?(toNum(r[vi])||0):0;
        return pv!==0||vv!==0;   /* v6b: 仅剔除零业绩且零见客的支撑行，保留负业绩（退款冲抵） */
      });
    }
    var cardsHTML=teamSheet?renderCards(teamSheet,cardRows):"";
    var crumb='<div class="crumb"><b>'+esc(state.all?"综合数据":g.label)+'</b>'
      +(state.filter?'<span class="crumb-filter">已筛选：'+esc(state.filter)+' <a href="javascript:void(0)" id="clearFilter">清除</a></span>':'')
      +'</div>';
    var panels="";
    activeSubs.forEach(function(sub){
      panels+='<section class="panel" id="panel_'+sub.key+'"><div class="panel-head"><h2>'+esc(titleMap[sub.key]||sub.label)+'</h2></div></section>';
    });
    if(state.all) crumb+=jkBadgeHtml();
    if(state.all) crumb+='<div class="cal-note">业绩口径：'
      +'① 分公司 / 事业部业绩＝《公司和部门业绩》（分公司及各事业部「业绩完成」）；'
      +'② 部门 / 个人业绩＝《人员业绩汇总》（<b>法律取「利润业绩」，其余取「业绩总和」</b>）；'
      +'③ 人均业绩（人均产值）＝业绩 ÷ 编制人数；<b>编制人数只统计「在职 ∧ 一线业务岗」</b>'
      +'（不含管理者、中台老师、职能支撑与流程会计），故「部门页」由个人汇总、与事业部页存在口径差属正常；'
      +'④ 咨询中心（中台老师）仅出现在「个人明细数据」；'
      +'⑤ 见客三项（见客(主面谈）/见客签单数/见客签单转化率）按工号匹配《近30天付费客资源转化》，各列目标按业态下限；'
      +'⑥ 部门考核：网销＝付费客资源签单转化率（按业态）／电邀＝人均月度高价值见客量／会计＝到期续费率＋非正常流失率，任一指标未达标即该部门未达标，达标率＝达成指标数 ÷ 考核指标数；'
      +'⑦ 职能/支撑类事业部人员（共 11 人）不显示在表格中、不参与任何指标计算，也不计入人员编制。具体名单见页底「口径说明」。</div>';
    main.innerHTML=crumb+cardsHTML+panels;
    activeSubs.forEach(function(sub){
      var sh=sheets[sub.sheet];
      var el=main.querySelector("#panel_"+sub.key);
      if(!el||!sh) return;
      if(sub.key==="res"){ renderRes(sub.sheet, el); return; }   /* v17：分公司资源明细（四级树） */
      if(isDetailPanel(sub.key)){ renderPanel(sub.key); return; }   /* v9：明细表走「部门筛选 + 分组折叠」渲染 */
      var ch=(data.charts||{})[sub.key];
      if(state.all&&ch) el.insertAdjacentHTML("beforeend",ch);
      el.insertAdjacentHTML("beforeend",tableHtml(sh,sh.rows));
      var tbl=el.querySelector("table");
      bindSort(tbl);
      applySticky(tbl);
    });
    var cf=document.getElementById("clearFilter");
    if(cf) cf.addEventListener("click",function(){state.filter=null;renderAll();});
  }

  /* ===== 页面最底部：AI 解析 + 口径说明 ===== */
  function renderBottom(){
    var h='<h3 class="bb-h">🤖 AI 数据解析</h3><p class="bb-sub">基于当前数据周期的自动解析（构建时按固定业务规则生成，口径与原链接一致）</p><div class="ai-grid">';
    AI.forEach(function(a){
      h+='<div class="ai-card ai-'+esc(a.sev)+'"><div class="ai-t">'+esc(a.title)+'</div><div class="ai-x">'+esc(a.text)+'</div></div>';
    });
    h+='</div>';
    document.getElementById("aiBlock").innerHTML=h;

    /* 口径说明：点击展开（09-20 用户要求） */
    var c='<div class="bb-head" id="calHead"><h3 class="bb-h">📖 口径说明</h3>'
      +'<span class="bb-toggle" id="calToggle">点击展开 ▾</span></div>'
      +'<div class="bb-body" id="calBody" hidden>'
      +'<p class="bb-sub" style="padding-top:14px">数据来源：'+esc(window.TARGET_SRC||"—")+'</p>'
      +'<p class="bb-sub">剔除口径：职能/支撑类事业部共 11 人已从全部表格与全部计算中剔除'
      +'（消化中心 6 人、职能 2 人、业务管理中心 2 人、业务支持部 1 人）；'
      +'「咨询中心」为中台老师，仅在「个人明细数据」中出现。</p>'
      +'<p class="bb-sub">编制人数口径（09-20）：只统计<b>在职 ∧ 一线业务岗</b>的人 —— 在职＝出现在当期在职花名册'
      +'《所有规划师明细数据》；一线业务岗＝企服/电邀规划师、法务主任、提成律师、会计顾问/客服会计/会计主管、'
      +'企服（储备）主管；<b>管理者</b>（业务经理/事业部总·副总/分公司总/总监）、<b>中台老师</b>与解决方案岗'
      +'（专家/专员）、<b>职能支撑</b>（HRBP/PMO/交付）、<b>流程会计</b>均不计入。'
      +'因此业务岗以外人员（含已离场但期间有业绩者）只贡献业绩、不摊入人均分母；'
      +'会计板块「团队人数」＝本部门在职会计业务岗（会计顾问＋客服会计，不含流程会计）。</p>'
      +'<p class="bb-sub">见客三项口径（09-20）：员工级数值按工号匹配《近30天付费客资源转化》'
      +'（端口，31 人）——「见客(主面谈）」＝成单或见客量·30 天，「见客签单数」＝成单量·30 天，'
      +'「见客签单转化率」＝30 天及时性转化率（剔除三方）；部门/公司层转化率＝有该表记录成员的平均值。'
      +'该列与「付费客资源签单转化率」<b>目标均按业态下限</b>（工商 ≥15%、跨境 ≥8%、法律 ≥7%、AI智能体 ≥4%、交易 ≥3%）。'
      +'表头下方一行「目标值」为该列标准线，悬停可看目标值 / 当前值 / 差距。</p>'
      +'<p class="bb-sub">部门考核口径（09-20）：<b>一个部门只要有任一考核指标未达标，该部门即判未达标</b>；「达标率」＝达成指标数 ÷ 考核指标数。'
      +'网销部门只考核「付费客资源签单转化率（近30天）」，目标＝该事业部业态下限（工商15%/跨境8%/法律7%/AI智能体4%/交易3%）；'
      +'电邀部门考核「人均月度高价值见客量」——入职 0-3 月 ≥3、入职 3 个月以上 ≥5；'
      +'会计部门考核「到期续费率 ≥85%」且「非正常流失率 ≤1%」。'
      +'综合数据「事业部数据 / 部门数据」表与三个模式页的团队表均给出考核结果；个人明细数据的部门分组头上也会显示该组考核结果（折叠时仍可见）。</p>'
      +'<p class="bb-sub">会计到期三指标（当月到期实收家数 / 当月到期家数 / 到期续费率）取自 '
      +esc(data.kj_due_src||"历史源·2026-09-14 · 四大模式-重点指标表格.xlsx · 会计续费明细表")
      +'：当月到期家数＝「是否应续费＝是」的家数，当月到期实收家数＝应续且「是否续费＝是」的家数，'
      +'到期续费率＝当月到期实收家数 ÷ 当月到期家数（标准 85%）；个人按「会计顾问」归口，'
      +'部门/团队按本部门明细求和，口径与主看板「会计续费明细」板块一致。</p>'
      +'<p class="bb-sub">付费客资源转化率口径（09-20）：「付费客资源签单转化率（近30天）」<b>目标按业态下限、各事业部不一致</b> —— 工商 ≥15%、跨境 ≥8%、法律 ≥7%、AI智能体 ≥4%、交易 ≥3%、知识产权 ≥5%（源表《标准》页）。该列表头「目标值」单元格显示为「按业态」，<b>悬停可看到各业态目标值</b>；部门分组行（可展开的那一行）按该组所属事业部显示对应目标，同组跨业态时显示区间。口径＝近30天付费客资源签单量 ÷ 近30天付费客资源总量（剔除三方、剔除无效资源）；分公司 / 多业态合计不套用单一目标，故不做达标着色。该列同为<b>网销部门的唯一考核指标</b>，「考核结果」列据此判定。</p>'
      +'<p class="bb-sub">员工档案口径对齐（09-20 v17）：明细表（网销 / 电邀 / 会计 / 所有规划师）末尾追加 8 列主看板「员工档案」指标，列名带「（主看板）」后缀 —— '
      +'付费客资源签单转化率 / 员工见客签单转化率 / 业务经理 · 中台老师 · 事业部总陪同见客签单转化率 / 当月见客签单转化率 / 人均产值 / 人员编制；'
      +'数值取自主看板网销漏斗口径（近30天 / 考核月），按姓名匹配，非网销岗位显示「—」。'
      +'该 8 列与表内原有的同名列<b>口径不同、并存供对比</b>（例：「人均产值（考核月）」＝源表口径，「人均产值（主看板）」＝漏斗口径）。'
      +'目标线：转化率类 ≥30%、付费客资源签单转化率按业态下限、人均产值 ≥¥20,000。</p>'
      +'<p class="bb-sub">分公司资源明细（09-20 v17）：新增顶级板块，四级树「分公司 → 事业部 → 部门 → 员工」，点击分组行展开 / 收起，'
      +'可勾选「仅显示有资源数据的人员」。9 项指标与主看板「分公司资源明细」完全同口径（网销漏斗 + LTC 近30天）：'
      +'有效资源量 / 资源加微量 / 经理分配占比(<b>&lt;60% 判红</b>) / 资源加微率(<b>&lt;85% 判红</b>) / 付费客30天签单转化率 / '
      +'见客转化率(<b>&lt;30% 判红</b>) / 付费客7天未加微移交量 / 付费客7天未加微移交率 / 移交接收执行率；'
      +'比率一律「先汇总再相除」，不是员工比率的简单平均。目前仅网销 35 人有资源数据，其余人员显示「—」，'
      +'分组行的「人数」列同时给出「有数据 N」。</p>'
      +'<p class="bb-sub">部门归属（09-20）：事业部负责人（源表「部门」列为空）按指定归入其分管部门 —— 示例负责人A → 工商网销业务二部 / 工商电邀业务六部（两组均出现），示例负责人B → 法律咨询三部，示例负责人C → 跨境业务二部，示例负责人D → 数字科技二部；4 人均为管理者，不计入编制人数，仅影响明细表的分组显示。</p>';
    var modes=[],map={};
    TARGET.forEach(function(t){
      if(!map[t.mode]){map[t.mode]=[];modes.push(t.mode);}
      map[t.mode].push(t);
    });
    modes.forEach(function(mode){
      c+='<div class="sub-title">'+esc(mode)+'</div><div class="tbl-wrap"><table><thead><tr><th>指标</th><th>目标</th><th>定义</th></tr></thead><tbody>';
      map[mode].forEach(function(t){
        c+='<tr><td class="t-left">'+esc(t.metric)+'</td><td class="t-left"><b>'+esc(t.target)+'</b></td>'
          +'<td class="t-left"><div class="def">'+esc(t.definition)+'</div></td></tr>';
      });
      c+='</tbody></table></div>';
    });
    c+='</div>';
    document.getElementById("caliberBlock").innerHTML=c;
    var _body=document.getElementById("calBody"), _tg=document.getElementById("calToggle");
    document.getElementById("calHead").addEventListener("click",function(){
      if(_body.hasAttribute("hidden")){_body.removeAttribute("hidden");_tg.textContent="收起 ▴";}
      else{_body.setAttribute("hidden","");_tg.textContent="点击展开 ▾";}
    });
  }

  function renderAll(){ renderSidebar(); renderTree(); renderMain(); }
  /* ===== v19：历史日期快照（资源明细不进快照，恒为最新） ===== */
  var RES_SHEET="分公司资源明细";
  var ORIG_RES=sheets[RES_SHEET], ORIG_DATA=data;
  function headText(){
    var _s=document.getElementById("srcNote"), _u=document.getElementById("updTip");
    if(_s) _s.textContent="数据源："+(data.source||"—");
    if(_u) _u.textContent="数据更新 · "+(data.updated||"—");
  }
  function applySnap(d){
    data=d; groups=d.groups||[]; sheets=d.sheets||{};
    if(ORIG_RES) sheets[RES_SHEET]=ORIG_RES;
    modeGroups=groups.filter(function(g){return g.label!=="综合数据";});
    allGroup=groups.filter(function(g){return g.label==="综合数据";})[0]||groups[groups.length-1]||allGroup;
    BU_T=(data.buTargets&&data.buTargets.rows)||[];
    state.groupIdx=0; state.all=false; state.filter=null; state.expanded={};
    state.panel={}; state.allTab="biz"; state.resColl={}; state.resInit=false;
    headText(); renderAll();
  }
  function loadSnap(date){
    var sel=document.getElementById("snapSel");
    if(!date){ applySnap(ORIG_DATA); return; }
    fetch("snapshots/"+date+".json",{cache:"no-store"}).then(function(r){
      if(!r.ok) throw new Error("HTTP "+r.status);
      return r.json();
    }).then(function(d){ applySnap(d); }).catch(function(e){
      alert("无法加载 "+date+" 的历史快照（"+e.message+"）。\n历史数据需通过云端链接访问（本地文件方式不支持）。");
      if(sel) sel.value="";
    });
  }
  function initSnap(){
    var sel=document.getElementById("snapSel"), wrap=document.getElementById("datePickWrap");
    if(!sel||!wrap) return;
    fetch("snapshots/manifest.json",{cache:"no-store"}).then(function(r){
      if(!r.ok) throw new Error("no manifest");
      return r.json();
    }).then(function(mf){
      var dates=(mf&&mf.dates)||[], latest=mf&&mf.latest;
      dates.forEach(function(d){
        if(d===latest) return;   /* 最新那档＝页面自带数据，不进下拉 */
        var o=document.createElement("option"); o.value=d; o.textContent=d; sel.appendChild(o);
      });
      if(sel.options.length>1) wrap.style.display="";
      sel.addEventListener("change",function(){ loadSnap(sel.value); });
    }).catch(function(){ /* 离线/本地打开：不显示日期选择器 */ });
  }
  headText(); renderBottom(); renderAll(); initSnap();
})();
</script>
</body>
</html>
"""

EXTRA_CSS = r"""
/* 表头可读性（09-18）：参考件 .modes-app th 的 min-width:0+break-word 会把窄列压成一字一行 */
.modes-app table thead th{min-width:88px}
.modes-app table thead th:first-child,.modes-app table thead th:nth-child(2){min-width:60px}

/* ===== 冻结列（09-20）：左侧 工号 / 姓名 / 事业部（缺则首列）横向滚动时固定 ===== */
.modes-app table thead th.stk{position:sticky;z-index:7;background:var(--bg-soft);
  box-shadow:inset -1px 0 0 var(--border),inset 0 -1px 0 var(--border)}
.modes-app table tbody td.stk{position:sticky;z-index:4;background:#fff;
  box-shadow:inset 0 -1px 0 var(--line-soft)}
.modes-app table tbody tr.row-biz td.stk{background:var(--primary-soft)}
.modes-app table tbody tr.row-total td.stk{background:#F3EFFE}
.modes-app table tbody tr:hover td.stk{background:var(--primary-tint)}
.modes-app table thead th.stk-last,.modes-app table tbody td.stk-last{
  box-shadow:3px 0 6px -2px rgba(31,36,48,.16),inset 0 -1px 0 var(--line-soft)}
.modes-app table thead th.stk-last{
  box-shadow:3px 0 6px -2px rgba(31,36,48,.16),inset -1px 0 0 var(--border)}
.tbl-wrap.pure-scroll{max-height:none}

/* ===== 口径说明：点击展开框（09-20） ===== */
.bb-card.collapsible{padding:0;overflow:hidden}
.bb-head{display:flex;align-items:center;justify-content:space-between;gap:12px;
  padding:16px 20px;cursor:pointer;user-select:none;transition:.15s}
.bb-head:hover{background:var(--primary-tint)}
.bb-head .bb-h{margin:0}
.bb-toggle{font-size:12.5px;font-weight:600;color:var(--primary-deep);background:var(--primary-soft);
  border-radius:8px;padding:4px 12px;white-space:nowrap;flex-shrink:0}
.bb-body{padding:0 20px 18px;border-top:1px solid var(--border-soft)}
.bb-body[hidden]{display:none}

.cal-note{font-size:12px;color:#475569;background:#EEF2FF;border:1px solid #E0E7FF;
  border-left:3px solid #6366F1;border-radius:8px;padding:7px 12px;margin:0 0 12px;line-height:1.7}

/* ===== v9：明细数据（个人）部门筛选 + 部门分组折叠 ===== */
.dt-bar{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin:8px 0 10px;padding:8px 10px;
  background:#F6F8FC;border:1px solid #E3E9F2;border-radius:9px}
.dt-lb{font-size:12px;color:#5B7396;font-weight:700;margin-right:2px}
.dt-chip{font-size:12px;line-height:1;padding:6px 10px;border-radius:14px;border:1px solid #D6DEEA;
  background:#fff;color:#475569;cursor:pointer;user-select:none;white-space:nowrap;transition:all .12s}
.dt-chip:hover{border-color:#5B7396;color:#33507A}
.dt-chip.on{background:#5B7396;border-color:#5B7396;color:#fff;font-weight:700}
.dt-chip .n{opacity:.7;margin-left:5px;font-size:11px}
.dt-sum{margin-left:auto;font-size:12px;color:#64748B;white-space:nowrap}
.dt-sum .dt-clr{color:#5B7396;text-decoration:none;border-bottom:1px dashed #5B7396;margin-left:4px}
.modes-app table tbody tr.dept-group>td, .dept-group>td{background:#E8EEF8 !important;color:#33507A !important;
  font-weight:800 !important;text-align:left !important;cursor:pointer;padding:8px 12px !important;
  border-top:1px solid #D3DEF0 !important;border-bottom:1px solid #D3DEF0 !important;position:static !important;left:auto !important}
.dept-group:hover>td{background:#DEE7F5 !important}
.dept-group .dg-ar{display:inline-block;width:14px;color:#5B7396;font-weight:700}
.dept-group .dg-n{font-weight:400;color:#64748B;font-size:12px;margin-left:10px}
.dept-group .dg-names{margin-left:12px;color:#33507A;font-weight:600;font-size:12.5px;letter-spacing:.2px}
tr.row-dept-hide{display:none !important}
tr.row-grp>td:first-child{border-left:2px solid #CBD8EC}

/* ===== v12（09-20）：表头「目标值」行 + 悬停提示 + 达标徽标 ===== */
.modes-app table thead tr.std-row th{background:#EDF1F8;color:#4A6382;font-weight:600;
  font-size:11.5px;text-align:center;padding:3px 8px;white-space:nowrap;cursor:help;letter-spacing:0;
  border-bottom:1px solid var(--border)}
.modes-app table thead tr.std-row th.std-none{color:#B4BECC;font-weight:400;cursor:default}
.modes-app table thead tr.std-row th.stk{background:#E4EAF4}
.modes-app table tbody td[title]{cursor:help}
.dg-badge{margin-left:12px;display:inline-flex;gap:6px;vertical-align:middle}
.dg-badge i{font-style:normal;font-size:11.5px;font-weight:700;border-radius:9px;padding:1px 9px}
.dg-badge i.ok{background:#DFF0E4;color:#37704F;border:1px solid #B9DCC6}
.dg-badge i.no{background:#FBE4E2;color:#9E4A43;border:1px solid #F0C3BF}
.dg-std{margin-left:12px;display:inline-block;font-size:11.5px;font-weight:700;color:#8A6520;background:#FDF3E0;border:1px solid #EFD9AE;border-radius:9px;padding:1px 9px;cursor:help}
.jk-badge{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:0 0 12px;padding:9px 14px;
  background:#F4F7FC;border:1px solid #DFE6F1;border-left:3px solid #5B7396;border-radius:8px}
.jk-badge .jbb-t{font-size:12.5px;font-weight:700;color:#33507A}
.jk-badge .jbb{font-size:12.5px;color:#475569;border-radius:10px;padding:1px 10px;background:#fff;
  border:1px solid #DFE6F1}
.jk-badge .jbb.ok{border-color:#B9DCC6;background:#F1F9F4;color:#37704F}
.jk-badge .jbb.no{border-color:#F0C3BF;background:#FDF3F2;color:#9E4A43}
.jk-badge .jbb b{font-size:13.5px}
.jk-badge .jbb-note{font-size:11.5px;color:#6B7A8D;margin-left:auto}

/* 兜底：modes 页核心组件（bundle css 未覆盖时保证可用） */
body{margin:0;background:var(--bg);color:var(--text);
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;font-size:13.5px}
.app{display:flex;flex-direction:column;min-height:100vh}
.header{display:flex;align-items:center;justify-content:space-between;padding:0 24px;height:64px;gap:20px;
  background:linear-gradient(90deg,#2563EB 0%,#4F46E5 45%,#7C3AED 100%);border-bottom:1px solid #312E81;flex-shrink:0}
.brand{display:flex;align-items:center;gap:12px;min-width:0}
.brand-mark{width:38px;height:38px;border-radius:10px;background:rgba(255,255,255,.18);color:#fff;
  display:flex;align-items:center;justify-content:center;font-weight:800;font-size:15px;flex-shrink:0}
.brand-text h1{margin:0;font-size:17px;color:#fff;letter-spacing:.5px;white-space:nowrap}
.brand .sub{color:rgba(255,255,255,.75);font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.header-right{display:flex;align-items:center;gap:12px}
.upd-tip{color:rgba(255,255,255,.8);font-size:12px;white-space:nowrap}
.body{display:flex;flex:1 1 auto}
.sidebar{width:230px;flex:0 0 230px;background:#ffffff;border-right:1px solid var(--border);
  padding:16px 12px;position:sticky;top:0;height:calc(100vh - 64px);overflow:auto}
.side-title{font-size:12px;font-weight:700;color:var(--text-weak);padding:4px 8px 8px;letter-spacing:1px}
.mode-tabs{display:flex;flex-direction:column;gap:4px}
.mode-tab{display:flex;align-items:center;gap:9px;padding:9px 10px;border-radius:9px;cursor:pointer;
  color:var(--text-sub);user-select:none;transition:.15s}
.mode-tab:hover{background:var(--primary-tint)}
.mode-tab.active{background:var(--primary-soft);color:var(--primary-deep);font-weight:700}
.mt-icon{width:24px;height:24px;border-radius:7px;background:var(--primary-soft);color:var(--primary-deep);
  display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;flex-shrink:0}
.mode-tab.active .mt-icon{background:var(--primary);color:#fff}
.mt-name{flex:1;font-size:13.5px}
.mt-count{font-size:11.5px;color:var(--text-weak);background:var(--bg-soft);border-radius:10px;padding:1px 8px}
.all-dept-tab{display:flex;align-items:center;gap:9px;padding:9px 10px;border-radius:9px;cursor:pointer;
  color:var(--text-sub);user-select:none;font-size:13.5px;transition:.15s}
.all-dept-tab:hover{background:var(--primary-tint)}
.all-dept-tab.active{background:var(--primary);color:#fff;font-weight:700}
.tree{list-style:none;margin:6px 0 0;padding:0}
.tree li{margin:1px 0}
.node{display:flex;align-items:center;gap:6px;padding:7px 8px;border-radius:8px;cursor:pointer;
  color:var(--text-sub);font-size:13px;user-select:none}
.node:hover{background:var(--primary-tint)}
.node.active{background:var(--primary-soft);color:var(--primary-deep);font-weight:700}
.node .arrow{width:0;height:0;border-left:5px solid var(--text-weak);border-top:4px solid transparent;
  border-bottom:4px solid transparent;transition:.15s;flex-shrink:0}
.node .arrow.expanded{transform:rotate(90deg)}
.node .arrow.leaf{border-left:4px dotted var(--text-weak)}
.child-node{margin-left:20px}
.node-count{margin-left:auto;font-size:11px;color:var(--text-weak);background:var(--bg-soft);
  border-radius:9px;padding:0 7px}
.main{flex:1;padding:20px 26px 10px;min-width:0}
.main-h{margin:0 0 14px;font-size:18px;color:var(--text)}
.crumb{display:flex;align-items:center;gap:10px;margin-bottom:14px;font-size:15px;color:var(--text)}
.crumb b{font-weight:800}
.crumb-filter{font-size:12.5px;color:var(--primary-deep);background:var(--primary-soft);
  border-radius:8px;padding:4px 10px;font-weight:400}
.crumb-filter a{color:var(--primary);margin-left:4px}
.panel{background:var(--card);border:1px solid var(--border);border-radius:12px;
  box-shadow:var(--shadow-sm);margin-bottom:16px;overflow:hidden}
.panel-head{padding:12px 16px;border-bottom:1px solid var(--border-soft)}
.panel-head h2{margin:0;font-size:14.5px;color:var(--text)}
.panel .tbl-wrap{border:0;border-radius:0;box-shadow:none;margin:0}
.tbl-wrap{background:var(--card);border:1px solid var(--border);border-radius:12px;
  overflow:auto;box-shadow:var(--shadow-sm);margin-bottom:8px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th,td{padding:8px 11px;border-bottom:1px solid var(--line-soft);white-space:nowrap}
th{background:var(--bg-soft);color:var(--text-sub);font-weight:600;font-size:12px;cursor:pointer;
  position:sticky;top:0;user-select:none}
th.sort-asc::after{content:" ▲";color:var(--primary)}
th.sort-desc::after{content:" ▼";color:var(--primary)}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover{background:var(--primary-tint)}
/* 字体居中（用户要求） */
.main table th,.main table td{text-align:center!important}
.charts{display:grid;grid-template-columns:2fr 1fr;gap:14px;margin:12px 0}
.charts.single{grid-template-columns:1fr;max-width:860px}
.chart-card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:12px 14px}
.chart-card h4{margin:0 0 8px;font-size:13px;color:var(--text);text-align:center}
.chart-card svg{width:100%;height:auto;display:block}
td.num{text-align:right;font-variant-numeric:tabular-nums}
td.t-left{text-align:left}
td.bad{color:var(--fail);font-weight:700}
td.ok{color:var(--pass)}
td.na{color:var(--text-weak)}
tr.row-biz{background:var(--primary-soft)}
tr.row-total{background:#F3EFFE}
.card.bad .v{color:var(--fail)}
.card.ok .v{color:var(--pass)}
.card.na .v{color:var(--text-weak)}
.card .std{font-size:10.5px;color:var(--text-weak);margin-top:3px}
.card .k{font-size:11.5px;color:var(--text-weak);margin-bottom:5px}
.card .v{font-size:17px;font-weight:800;color:var(--text);font-variant-numeric:tabular-nums}
.card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:12px 14px;
  width:150px;flex-shrink:0;box-shadow:var(--shadow-card)}
.cards{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:stretch}
.card-tenure{width:auto;min-width:220px}
.tenure-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-top:8px}
.t-item{background:var(--primary-tint);border-radius:8px;padding:6px 4px;text-align:center}
.t-item .t-val{font-weight:800;font-size:15px}
.t-item .t-label{font-size:10.5px;color:var(--text-weak);margin-top:2px}
.t-item.t1{background:#EEF0FE}.t-item.t2{background:#E4E7FD}.t-item.t3{background:#DDE1FC}.t-item.t4{background:#F3EFFE}
.tenure-total{margin-left:8px;font-size:11px;color:var(--text-weak);font-weight:400}
.sub-title{font-size:13.5px;font-weight:700;color:var(--primary-deep);margin:14px 0 8px;
  padding-left:9px;border-left:3px solid var(--primary)}
/* 底部 AI 解析 + 口径说明 */
.bottom-blocks{padding:6px 26px 46px}
.bb-card{background:var(--card);border:1px solid var(--border);border-radius:14px;
  box-shadow:var(--shadow-sm);padding:18px 20px;margin-top:18px}
.bb-h{font-size:16px;font-weight:700;color:var(--text);margin:0 0 4px}
.bb-sub{font-size:12px;color:var(--text-weak);margin:0 0 14px}
.ai-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px}
.ai-card{background:var(--primary-tint);border:1px solid var(--border);border-left:4px solid var(--accent);
  border-radius:10px;padding:12px 14px}
.ai-card .ai-t{font-weight:700;font-size:13.5px;margin-bottom:4px;color:var(--primary-deep)}
.ai-card .ai-x{font-size:12.5px;color:var(--text-sub);line-height:1.65}
.ai-good{border-left-color:var(--pass)}.ai-good .ai-t{color:var(--pass)}
.ai-pass{border-left-color:var(--pass)}.ai-pass .ai-t{color:var(--pass)}
.ai-warn{border-left-color:var(--warn)}.ai-warn .ai-t{color:var(--warn)}
.ai-bad{border-left-color:var(--fail)}.ai-bad .ai-t{color:var(--fail)}
.ai-info{border-left-color:var(--accent)}
.def{font-size:12px;color:var(--text-weak);line-height:1.6;white-space:pre-wrap}

  /* v16：考核结果 / 考核模式 列 + 徽标 */
  .ex-td{text-align:center;white-space:nowrap}
  .ex-badge{display:inline-block;padding:1px 8px;border-radius:9px;font-size:11px;
            font-weight:600;line-height:17px;white-space:nowrap}
  .ex-badge.ex-ok{background:#e4efe8;color:#3f7357;border:1px solid #bfd9c9}
  .ex-badge.ex-no{background:#f7e7e4;color:#a2463c;border:1px solid #e9c6c0}
  .dg-badge.is-ok i.ok{background:#e4efe8;color:#3f7357}
  .dg-badge.is-no i.no{background:#f7e7e4;color:#a2463c}
"""

html = (TEMPLATE
        .replace("__CSS__", CSS_BUNDLE)
        .replace("__EXTRA__", EXTRA_CSS)
        .replace("__MODES__", json.dumps(MODES, ensure_ascii=False).replace("</", "<\\/"))
        .replace("__AI__", json.dumps(AI, ensure_ascii=False).replace("</", "<\\/"))
        .replace("__TARGET__", json.dumps(TARGET, ensure_ascii=False).replace("</", "<\\/"))
        .replace("__SRC__", json.dumps(TARGET_SRC, ensure_ascii=False)))

os.makedirs(OUT_DIR, exist_ok=True)
out_path = os.path.join(OUT_DIR, "index.html")
with open(out_path, "w", encoding="utf-8") as f:
    f.write(html)
print("WROTE", out_path, "size=", len(html))

# ---- v19：历史快照（资源明细不进快照：该板块恒为最新数据） ----
SNAP_DIR = os.path.join(OUT_DIR, "snapshots")
os.makedirs(SNAP_DIR, exist_ok=True)
SNAP_RES = "分公司资源明细"
_snap = dict(MODES)
_snap["sheets"] = dict((k, v) for k, v in SHEETS.items() if k != SNAP_RES)
_snap["_resSnapshot"] = False
_snap_date = str(MODES.get("updated") or "")[:10] or datetime.datetime.now().strftime("%Y-%m-%d")
_snap_path = os.path.join(SNAP_DIR, _snap_date + ".json")
with open(_snap_path, "w", encoding="utf-8") as f:
    json.dump(_snap, f, ensure_ascii=False)
_snap_dates = sorted([x[:-5] for x in os.listdir(SNAP_DIR) if x.endswith(".json")], reverse=True)
with open(os.path.join(SNAP_DIR, "manifest.json"), "w", encoding="utf-8") as f:
    json.dump({"latest": _snap_date, "dates": _snap_dates}, f, ensure_ascii=False)
print("SNAPSHOT", _snap_path, "size=", os.path.getsize(_snap_path), "| dates=", _snap_dates)
for name, sh in SHEETS.items():
    print("  sheet", name, "rows=", len(sh["rows"]), "cols=", len(sh["headers"]),
          "redrows=", sum(1 for r in sh["redFlags"] if r))
print("target_rows=", len(TARGET), "src=", TARGET_SRC, "ai=", len(AI))
print("EXAM depts=", len(EXAM))
for _k, _v in sorted(EXAM.items()):
    print("   %-30s %-4s %d/%d ok=%s  " % (_k, _v["mode"], _v["okn"], _v["tot"], _v["ok"]),
          " | ".join("%s=%.4f(>=%.4f)" % (it["k"], it["v"], it["t"]) for it in _v["items"]))
print("kj_due src=", KJ_DUE_SRC, "all=", KJ_DUE.get("all"),
      "person=", {k: (v["到期"], v["实收"]) for k, v in (KJ_DUE.get("person") or {}).items()})
print("口径源：公司/事业部业绩 =", PERF_CO_SRC, "| 分公司合计 =", PERF_TOTAL)
print("        个人业绩     =", EMP_PERF_SRC, "| 人数 =", len(EMP_PERF),
      "| 人员明细行 =", len(U_CLEAN))
print("        事业部行小计 =", round(sum(num(x.get("业绩")) for x in _biz), 2),
      "（应与分公司合计严格相等）")
