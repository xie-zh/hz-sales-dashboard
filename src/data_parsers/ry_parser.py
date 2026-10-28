# -*- coding: utf-8 -*-
"""
人员数据分析路径解析器 → 员工数据（网销漏斗 + LTC 流程指标）
================================================================
读取「某公司-人员数据分析路径.xlsx」：
  1. 网销漏斗 表：近30天漏斗指标（36人）
  2. LTC流程数据 表：当日全量 LTC 指标（78人，含六部门）
仅保留六个部门：工商、会计、法律、跨境、数字科技、交易。
按姓名关联两表，生成每位员工的完整画像 + 针对性诊断建议。
"""
import json, os, math, re
from collections import OrderedDict
from datetime import datetime, timedelta

import openpyxl
import config_loader

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE_DIR, "某公司-人员数据分析路径.xlsx")

# 部门/岗位规则改由 config/dashboard.config.json 驱动（配置即代码）。
# 配置文件缺失或解析失败时，config_loader 自动回退到与现状一致的内置默认值，行为不变。
CFG = config_loader.load_config()
_P = CFG["personnel"]
DEPT_RULES = [tuple(x) for x in _P["dept_rules"]]
DEPT_ORDER = list(_P["dept_order"])
COMPANY_KEYWORD = _P["company_keyword"]
TARGET_POSTS = set(_P["target_posts"])

DEPT_COLORS = {
    "工商": "#5b7ea6",
    "会计": "#6fa3b8",
    "法律": "#8f7fb8",
    "跨境": "#c2a15e",
    "数字科技": "#6da88b",
    "交易": "#c07f76",
}

# 电邀漏斗相关（2026-09-11 新增）
# 部门包含此关键字的人员，funnel 字段使用「电邀漏斗」sheet 而非「网销漏斗」
DY_DEPT_KEYWORD = "电邀业务"
DY_DEPT_STD = "电邀"  # 这类人员的标准部门名（用于聚合分组）
DY_DEPT_ORDER = 0  # 标准部门列表中的排序位置（最前）

# 电邀漏斗核心展示指标（12 个关键列）：名称, 是否比率, 说明
DY_FUNNEL_METRICS = [
    ("计划拨打数", False, "当日计划拨打总数"),
    ("计划完成", False, "实际完成拨打量"),
    ("计划拨打完成率", True, "完成 / 计划"),
    ("接通数", False, "接通呼叫数"),
    ("接触数", False, "有效接触客户数"),
    ("30S以上", False, "30秒以上有效沟通数"),
    ("通时", False, "总通时(分钟)"),
    ("转化私客数", False, "当日转化为私客的数量"),
    ("加微量", False, "当日加微总数"),
    ("有效沟通加微率", True, "有效沟通转加微的比例"),
    ("过去30天见客接触签单率（不能低于20%）", True, "见客→签单转化率"),
    ("过去30天万产额", False, "近30天万产额"),
    ("过去30天客单价", False, "近30天客单价"),
]


def _is_dy_dept(dept):
    """某部门是否归类为电邀（即从电邀漏斗取数）"""
    if not dept:
        return False
    return DY_DEPT_KEYWORD in dept


def _classify_dept(dept):
    """把 LTC 部门名归入六部门。电邀业务六部单独归到「电邀」"""
    if not dept:
        return None
    if _is_dy_dept(dept):
        return DY_DEPT_STD
    for kw, std in DEPT_RULES:
        if kw in dept:
            return std
    return None


def _to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "--"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _pct(v):
    """0-1 比例 → 显示为百分数（保留1位）或 None"""
    if v is None:
        return None
    return round(v * 100, 1)



def _load_funnel():
    """网销漏斗：返回 (by_id, by_name)，每项 = {指标名: 值}，仅保留数值列。

    源表为多分公司混合表（约 541 人、13 家分公司），故同时按「工号」与「姓名」
    双索引；匹配时工号优先（防跨分公司重名串号，源表已存在重名「王玲」），
    姓名仅作兜底。
    """
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["网销漏斗"]
    rows = list(ws.iter_rows(values_only=True))
    headers = []
    for j in range(len(rows[1])):
        h = rows[1][j]
        headers.append(str(h).strip() if h else "")
    by_id = {}
    by_name = {}
    for r in rows[3:]:
        if not r[5]:
            continue
        name = str(r[5]).strip()
        emp_id = str(r[4]).strip() if len(r) > 4 and r[4] is not None else ""
        emp = {}
        for j in range(6, len(headers)):
            h = headers[j]
            if h:
                emp[h] = _to_num(r[j])
        by_name.setdefault(name, emp)
        if emp_id:
            by_id.setdefault(emp_id, emp)
    return by_id, by_name


def _load_dy_funnel():
    """电邀漏斗：返回 (by_id, by_name, people)。

    电邀漏斗 sheet 是层级结构：第 0 列为层级标签（分公司 / 事业部 / 部门 / "姓名(工号)"）。
    列 1..52 是漏斗指标（行 2 表头）。
    仅提取员工行（"姓名(工号)"），用于与 LTC 员工按工号/姓名匹配。
    """
    import re
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    if "电邀漏斗" not in wb.sheetnames:
        return {}, {}, []
    ws = wb["电邀漏斗"]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 4:
        return {}, {}, []
    headers = []
    for j in range(len(rows[1])):
        h = rows[1][j]
        headers.append(str(h).strip() if h else "")
    by_id = {}
    by_name = {}
    people = []
    pat = re.compile(r"^(.+?)\((\d+)\)\s*$")
    for r in rows[3:]:
        if not r or r[0] is None:
            continue
        label = str(r[0]).strip()
        m = pat.match(label)
        if not m:
            continue
        name = m.group(1).strip()
        emp_id = m.group(2).strip()
        emp = {}
        for j in range(1, len(headers)):
            h = headers[j]
            if h:
                emp[h] = _to_num(r[j])
        people.append({"工号": emp_id, "姓名": name, "部门": label})
        by_id[emp_id] = emp
        by_name[name] = emp
    return by_id, by_name, people


def _load_ltc():
    """LTC流程数据：每行员工 → 全部列，返回 (rows, data_date)"""
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["LTC流程数据"]
    rows = list(ws.iter_rows(values_only=True))
    headers = []
    for j in range(len(rows[0])):
        h = rows[0][j]
        headers.append(str(h).strip() if h else "")
    result = []
    data_date = None
    for r in rows[2:]:
        if not r[2]:
            continue
        row = {}
        for j in range(min(len(headers), len(r))):
            h = headers[j]
            if h:
                row[h] = r[j]
        result.append(row)
        if data_date is None and row.get("日期"):
            data_date = str(row["日期"]).strip()
    return result, data_date


# ---------- 面谈明细（陪同见客签单）数据源 ----------
# 数据源：数据文件/分公司健康表/四大模式-重点指标表格.xlsx 内的「面谈明细」sheet
#        （表名可能为 Sheet1；表头在第 2 行，数据自第 3 行；AZ 列 = 岗位）。
INTERVIEW_XLSX_CANDIDATES = [
    os.path.normpath(os.path.join(BASE_DIR, "..", "..", "..", "分公司健康表", "四大模式-重点指标表格.xlsx")),
    os.path.normpath(os.path.join(BASE_DIR, "分公司健康表", "四大模式-重点指标表格.xlsx")),
]
# 岗位(AZ) → 指标名（用户口径：管理者=事业部总陪谈；业务经理=业务经理陪谈；中台=中台陪谈）
INTERVIEW_ROLE_MAP = [
    ("业务经理", "业务经理陪同见客签单转化率"),
    ("中台", "中台老师陪同见客签单转化率"),
    ("管理者", "事业部总陪同见客签单转化率"),
]


# ---------- 网销漏斗「分公司/事业部级」汇总行：付费客30天签单转化率（AL 列） ----------
# 数据源：数据文件/分公司健康表/四大模式-重点指标表格.xlsx 内的「网销漏斗」sheet。
# 用户口径（2026-09-15）：「演示用户付费客资源签单转化率（近30天）」取自该 sheet 中
#   A 列名称匹配的行 —— A 列同时含 分公司 / 事业部 / 部门 / 员工 四级名称，
#   例如 A 列 =「某公司」的那一行即分公司级数据；指标值取该行 AL 列
#   「付费客30天签单转化率」（= 付费客30天签单量 ÷ 有效资源量）。
# 列按表头名定位（不写死列号），表头在第 1~2 行、数据自第 4 行（第 3 行为指标说明）。
# ⚠️ 该 sheet 含全国各分公司数据，名称匹配必须限定在本看板所属分公司内，否则
#    「工商」会误命中「上海工商事业部」。
NET_SUMMARY_BRANCH = "<CITY>"
_NET_SUMMARY_CACHE = {}


def _net_summary_path():
    for p in INTERVIEW_XLSX_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def _aux_mtimes():
    """辅助数据源（不产生独立快照、但会影响网销三页）的 mtime 清单。

    写入 ry_*.json 的 meta.aux_mtime_ts，供 check_and_deploy.py 判定是否需要重新同步。
    ⚠️ 必须与「自身记录的 mtime」比较：若拿它去和主源表的 mtime 比，会因主源表更旧
    而每轮都触发（永久同步循环）。
    """
    out = {}
    for p in (_net_summary_path(),):
        if p and os.path.exists(p):
            try:
                out[os.path.basename(p)] = os.path.getmtime(p)
            except OSError:
                pass
    return out


def _load_net_summary():
    """四大模式「网销漏斗」→ {A列名称: {"rate","sign","valid","total"}}，附带 path。"""
    if _NET_SUMMARY_CACHE.get("loaded"):
        return _NET_SUMMARY_CACHE.get("data") or {}, _NET_SUMMARY_CACHE.get("path")
    by_name, path = {}, None
    p = _net_summary_path()
    if p:
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            if "网销漏斗" in wb.sheetnames:
                ws = wb["网销漏斗"]
                allrows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
                ncol = max([len(r) for r in allrows[:3]] or [0])
                head = []
                for i in range(ncol):
                    v = None
                    for r in allrows[:2]:          # 第 2 行优先于第 1 行
                        if i < len(r) and r[i] is not None and str(r[i]).strip():
                            v = str(r[i]).strip()
                    head.append(v)

                def _find(*subs):
                    for i, nm in enumerate(head):
                        if nm and all(x in nm for x in subs):
                            return i
                    return None

                i_rate = _find("付费客30天签单转化率")   # AL
                i_sign = _find("付费客30天签单量")       # AK
                i_valid = _find("有效资源量")           # F
                i_total = _find("演示用户")             # E

                # 该表 A1 形如 '分公司\n(2026-09-11)' → 记录导出/数据日期，便于前端标注时效
                _a1 = str(allrows[0][0]) if allrows and allrows[0] else ""
                _dm = re.search(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", _a1)
                if _dm:
                    _NET_SUMMARY_CACHE["date"] = "%s-%02d-%02d" % (
                        _dm.group(1), int(_dm.group(2)), int(_dm.group(3)))

                for r in allrows[3:]:
                    a = r[0] if len(r) else None
                    if a is None:
                        continue
                    nm = str(a).strip()
                    if (not nm) or ("\n" in nm) or nm == "指标说明":
                        continue

                    def _g(idx, _r=r):
                        if idx is None or idx >= len(_r):
                            return None
                        v = _r[idx]
                        return v if isinstance(v, (int, float)) else None

                    if nm not in by_name:
                        by_name[nm] = {"rate": _g(i_rate), "sign": _g(i_sign),
                                       "valid": _g(i_valid), "total": _g(i_total)}
                path = p
        except Exception:
            by_name, path = {}, p
    _NET_SUMMARY_CACHE.update({"loaded": True, "data": by_name, "path": path})
    return by_name, path


def _net_row_lookup(name, by_name=None):
    """按 A 列名称定位「网销漏斗」行，返回 (匹配到的行名, meta)。

    先精确匹配；未命中时在「<CITY>…事业部」级名称中按关键字包含匹配
    （如 工商 → <CITY>工商事业部、数字科技 → <CITY>数字科技事业部）。
    返回 (None, None) 表示未找到。
    """
    if by_name is None:
        by_name, _ = _load_net_summary()
    key = str(name or "").strip()
    if (not by_name) or (not key):
        return None, None
    if key in by_name:
        return key, by_name[key]
    cands = sorted([n for n in by_name
                    if n.startswith(NET_SUMMARY_BRANCH)
                    and n.endswith("事业部") and key in n], key=len)
    if cands:
        return cands[0], by_name[cands[0]]
    return None, None


def _net_rate_lookup(name, by_name=None):
    """按 A 列名称取 AL「付费客30天签单转化率」。返回 (rate, meta)。"""
    _row, meta = _net_row_lookup(name, by_name)
    if (not meta) or meta.get("rate") is None:
        return None, None
    return meta.get("rate"), meta


def _net_rate_for_dept(dept):
    """部门（工商/法律/跨境/数字科技/交易）→ 对应事业部行 AL；未命中 None。"""
    return _net_rate_lookup(dept)[0]


def _net_rate_for_bizline(biz):
    """业务线（工商/跨境/法律/AI智能体/交易）→ 对应事业部行 AL；未命中 None。"""
    return _net_rate_for_dept(BIZ_TO_DEPT.get(biz, biz))


def _net_rate_for_emp(e):
    """员工级：A 列「姓名(工号)」行 → AL「付费客30天签单转化率」。

    用户口径（2026-09-15 M-3）：本指标在**四个级别**（分公司 / 事业部 / 部门 / 员工）
    一律取自「网销漏斗」sheet 的 AL 列，按 A 列名称匹配行。
    员工级以「姓名(工号)」精确匹配为主（工号唯一，最稳）；未命中时按姓名唯一匹配兜底；
    仍未命中返回 None（调用方回退原「付费客资源 sheet」口径）。
    """
    by_name, _ = _load_net_summary()
    if not by_name:
        return None
    if isinstance(e, dict):
        nm = str(e.get("姓名") or "").strip()
        gid = str(e.get("工号") or "").strip()
    else:
        nm, gid = str(e or "").strip(), ""
    if not nm:
        return None
    if gid:
        for k in ("%s(%s)" % (nm, gid), "%s（%s）" % (nm, gid)):
            meta = by_name.get(k)
            if meta and meta.get("rate") is not None:
                return meta.get("rate")
    cands = [k for k in by_name
             if k == nm or k.startswith(nm + "(") or k.startswith(nm + "（")]
    if len(cands) == 1:
        return by_name[cands[0]].get("rate")
    return None


def _interview_path():
    for p in INTERVIEW_XLSX_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def _load_interview():
    """读取面谈明细 sheet（按表头名解析，不写死列号）。

    返回 (rows, path)：rows = [{列名: 值}]。
    表头定位：每个 sheet 取前 3 行中非空最多的一行作表头，要求同时含「岗位」与含「面谈」的列。
    （用户已确认 AZ 列 = 岗位；此处按名称匹配以防列位漂移。）
    """
    p = _interview_path()
    if not p:
        return [], None
    try:
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception:
        return [], p
    ws = None
    hdr_idx = None
    for sn in wb.sheetnames:
        w = wb[sn]
        rws = list(w.iter_rows(min_row=1, max_row=3, values_only=True))
        if not rws:
            continue
        hi = max(range(len(rws)),
                 key=lambda i: sum(1 for c in rws[i] if c is not None and str(c).strip()))
        hs = set(str(c).strip() for c in rws[hi] if c is not None and str(c).strip())
        if "岗位" in hs and any("面谈" in h for h in hs):
            ws, hdr_idx = w, hi
            break
    if ws is None:
        return [], p
    rows = list(ws.iter_rows(values_only=True))
    if hdr_idx is None or hdr_idx >= len(rows):
        return [], p
    names = [str(c).strip() if c is not None else "" for c in rows[hdr_idx]]
    # 表头归一化：去掉换行/空白（源表头含 "私客所属人\n姓名" 之类），便于按名取值
    import re as _re
    names = [_re.sub(r"\s+", "", n) for n in names]
    out = []
    for r in rows[hdr_idx + 1:]:
        if not r or all(v is None or str(v).strip() == "" for v in r):
            continue
        d = {}
        for j, v in enumerate(r):
            if j < len(names) and names[j]:
                d[names[j]] = v
        out.append(d)
    return out, p


# ---------- 源表「见客」sheet：员工见客签单转化率（考核月 · 不含陪谈） ----------
# 用户口径（2026-09-21，源表「目标」sheet 第 5 行「定义」原文）：
#   ② 员工见客签单转化率（考核月）= 考核月见客签单量 ÷ 考核月的演示用户（**不含陪谈**）目标 30%
#   ⑥ 当月见客签单转化率（考核月）= 分公司网销漏斗「接收资源30天内见客后的转化率」（另行维护，见下方）
# 数据源：源表 XLSX 的「见客」sheet（SR-薯片邀约上门明细），表头在第 2 行、数据自第 3 行。
#   · 「记录类型」= 主面谈 / 陪面谈 ⇒「不含陪谈」= 仅取「主面谈」行
#   · 「岗位」∈ {网销岗, 综法岗}（与命中表 I / J 口径一致：den=51, num=2 ⇒ 3.92%）
#   · 签单判定 =「首次收款时间」非空
# ⚠️ 该 sheet 当前仅覆盖 09-13 ~ 09-20（缺 09-19），即 5 天口径 —— 用户已知悉并接受，
#    看板按「注明口径」处理，不等整月导出。分母为 0 显示「-」且不判定达标。
VISIT_SHEET_NAME = "见客"
VISIT_ROLE_INCLUDE = ("网销岗", "综法岗")
VISIT_MAIN_TYPE = "主面谈"
VISIT_SIGN_COL = "首次收款时间"
# 见客表部门列 → 看板部门短名（「数字科技」须先匹配，避免被其他关键词截胡）
VISIT_DEPT_KEYS = (
    ("数字科技", ("数字科技",)),
    ("法律", ("法律",)),
    ("跨境", ("跨境",)),
    ("交易", ("交易",)),
    ("工商", ("工商",)),
    ("会计", ("会计",)),
)
VISIT_WINDOW_NOTE = "见客 sheet 窗口 09-13~09-20（5 天口径）"
_VISIT_CACHE = {}


def _visit_dept_short(row):
    for col in ("事业部", "部门", "三级部门", "四级部门"):
        s = str(row.get(col) or "")
        if not s:
            continue
        for short, keys in VISIT_DEPT_KEYS:
            for k in keys:
                if k in s:
                    return short
    return None


def _visit_is_signed(row):
    v = row.get(VISIT_SIGN_COL)
    return v is not None and str(v).strip() not in ("", "无", "None", "none", "nan", "-")


def _load_visit_rows():
    """读源表「见客」sheet → 行字典列表。失败 / 缺 sheet → []（安全降级）。"""
    if not os.path.exists(XLSX):
        return []
    try:
        wb = openpyxl.load_workbook(XLSX, data_only=True)
    except Exception:
        return []
    try:
        if VISIT_SHEET_NAME not in wb.sheetnames:
            return []
        raw = list(wb[VISIT_SHEET_NAME].iter_rows(values_only=True))
    except Exception:
        return []
    finally:
        try:
            wb.close()
        except Exception:
            pass
    if len(raw) < 3:
        return []
    import re as _re
    hi = max(range(min(3, len(raw))),
             key=lambda i: sum(1 for c in raw[i] if c is not None and str(c).strip()))
    names = [_re.sub(r"\s+", "", str(c)) if c is not None else "" for c in raw[hi]]
    out = []
    for r in raw[hi + 1:]:
        if not r or all(v is None or str(v).strip() == "" for v in r):
            continue
        d = {}
        for j, v in enumerate(r):
            if j < len(names) and names[j]:
                d[names[j]] = v
        if d.get("姓名") is None:
            continue
        out.append(d)
    return out


def visit_agg():
    """「见客」sheet 三级聚合：主面谈 ∧ 岗位∈{网销岗,综法岗}。

    返回 {"company": {den,num,rate}, "dept": {部门短名: {...}}, "emp": {姓名: {...}},
          "all": {den,num,rate}}；"all" = 全部记录（不分记录类型/岗位），供「演示用户」展示。
    """
    if _VISIT_CACHE:
        return _VISIT_CACHE
    rows = _load_visit_rows()
    comp = {"den": 0, "num": 0}
    allc = {"den": 0, "num": 0}
    dept, emp = {}, {}
    for r in rows:
        allc["den"] += 1
        sgn = 1 if _visit_is_signed(r) else 0
        allc["num"] += sgn
        if str(r.get("记录类型") or "").strip() != VISIT_MAIN_TYPE:
            continue
        if str(r.get("岗位") or "").strip() not in VISIT_ROLE_INCLUDE:
            continue
        comp["den"] += 1
        comp["num"] += sgn
        d = _visit_dept_short(r)
        if d:
            o = dept.setdefault(d, {"den": 0, "num": 0})
            o["den"] += 1
            o["num"] += sgn
        nm = str(r.get("姓名") or "").strip()
        if nm:
            o = emp.setdefault(nm, {"den": 0, "num": 0})
            o["den"] += 1
            o["num"] += sgn

    def _fill(o):
        o["rate"] = (o["num"] / o["den"]) if o["den"] else None

    _fill(comp)
    _fill(allc)
    for m in (dept, emp):
        for o in m.values():
            _fill(o)
    _VISIT_CACHE.update({"company": comp, "dept": dept, "emp": emp, "all": allc,
                         "row_count": len(rows), "src": XLSX})
    return _VISIT_CACHE


# ---------- 资源指标改由「网销漏斗」逐人列派生（2026-09-21 用户口径） ----------
# 用户口径（2026-09-21）：「资源以漏斗数据中的有效资源量即可」。
# 源表 2026-09-21 15:12 版起移除「付费客资源」sheet，故资源类指标（演示用户 / 有效资源量 /
# 加微量 / 加微率 / 付费客30天签单量·转化率 / 见客量·见客转化率 / 移交量·移交率 /
# 经理手动分配占比）一律取「网销漏斗」sheet 的逐人员工行（列名与表头完全同名），
# 再由 _funnel_res_agg() 聚合到 个人 / 部门 / 公司 三级。
FUNNEL_RES_KEYS = (
    "演示用户", "有效资源量", "资源加微量", "资源加微率",
    "付费客30天签单量", "付费客30天签单转化率",
    "接收资源30天内见客量", "接收资源30天内见客后签单量", "接收资源30天内见客后的转化率",
    "经理手动分配资源占比",
    "付费客7天后未添加微信资源移交量", "付费客7天后未添加微信资源移交率",
    "销售额",
)
# 需在 部门 / 公司 级求和的计数型字段（**英文键**，与 _funnel_res_of() 的输出一致；
# 率类字段不求和，改由 Σ分子 ÷ Σ分母 重算）
FUNNEL_RES_SUM = ("resources", "total", "wx", "signs", "visits", "signed_visit", "mv_qty")


def _funnel_num(e, key):
    """取某员工「网销漏斗」逐人列的值（非数值→None）。"""
    v = (e.get("funnel") or {}).get(key)
    return v if isinstance(v, (int, float)) else None


def _funnel_res_of(e):
    """单个员工的资源指标（直接取「网销漏斗」逐人列；率缺失时用分子÷分母兜底）。"""
    g = lambda k: _funnel_num(e, k)  # noqa: E731

    def _r(a, b):
        return round(a / b, 4) if (a is not None and b) else None

    res = g("有效资源量")
    wx = g("资源加微量")
    signs = g("付费客30天签单量")
    visits = g("接收资源30天内见客量")
    sv = g("接收资源30天内见客后签单量")
    wxr = g("资源加微率")
    sgr = g("付费客30天签单转化率")
    vtr = g("接收资源30天内见客后的转化率")
    return {
        "resources": res,
        "total": g("演示用户"),
        "wx": wx,
        "signs": signs,
        "visits": visits,
        "signed_visit": sv,
        "wx_rate": wxr if wxr is not None else _r(wx, res),
        "sign_rate": sgr if sgr is not None else _r(signs, res),
        "visit_rate": vtr if vtr is not None else _r(sv, visits),
        "mgr_share": g("经理手动分配资源占比"),
        "mv_qty": g("付费客7天后未添加微信资源移交量"),
        "mv_rate": g("付费客7天后未添加微信资源移交率"),
        "perf": g("销售额"),
    }


def _funnel_res_agg(employees, funnel_src="net"):
    """把「网销漏斗」资源指标聚合为 个人 / 部门 / 公司 三级。

    返回 (per_person, by_dept, company)；by_dept/company 的率类字段用 Σ分子 ÷ Σ分母 重算
    （与原「付费客资源」口径的 Σ成单 ÷ Σ资源 保持一致）。
    """
    pool = [e for e in employees if e.get("funnel_source") == funnel_src]
    per, by_dept = {}, {}
    comp = {"n": 0}
    for k in FUNNEL_RES_SUM:
        comp[k] = 0.0
    for e in pool:
        m = _funnel_res_of(e)
        nm = str(e.get("姓名") or "").strip()
        if nm:
            per[nm] = m
        d = e.get("部门标准")
        a = by_dept.get(d)
        if a is None:
            a = {"n": 0}
            for k in FUNNEL_RES_SUM:
                a[k] = 0.0
            by_dept[d] = a
        a["n"] += 1
        comp["n"] += 1
        for k in FUNNEL_RES_SUM:
            v = m.get(k)
            if isinstance(v, (int, float)):
                a[k] += v
                comp[k] += v

    def _rate(num, den):
        return round(num / den, 4) if den else None

    for a in list(by_dept.values()) + [comp]:
        a["wx_rate"] = _rate(a["wx"], a["resources"])
        a["sign_rate"] = _rate(a["signs"], a["resources"])
        a["visit_rate"] = _rate(a["signed_visit"], a["visits"])
    return per, by_dept, comp


# ---------- 资源指标按「人」聚合（付费客资源 AS=当前私客所属人；剔除无效资源） ----------
PAY_EXCL_DESC = "①有效资源量=0 ②资源编码为空 ③是否无效=是 ④是否反无效=是"


def _aggregate_pay_by_person(pay_rows, pay_max_date):
    """付费客资源 → 按「当前私客所属人」(AS 列) 聚合，近30天窗口，已剔除无效资源。

    返回 {姓名: {resources, visits, signs, sign_rate, visit_rate, wx, wx_rate, sign_perf}}
    口径：演示用户=有效资源条数；资源签单率=成单数÷演示用户；资源加微信=「是否加微」为是的条数。
    """
    from datetime import timedelta
    wstart = (pay_max_date - timedelta(days=30)) if pay_max_date else None
    acc = {}
    for r in pay_rows:
        if _pay_excluded(r):
            continue
        if wstart and r["date"] < wstart:
            continue
        nm = str(r.get("当前私客所属人") or "").strip()
        if not nm:
            continue
        a = acc.setdefault(nm, {"resources": 0, "visits": 0, "signs": 0, "wx": 0, "sign_perf": 0.0})
        a["resources"] += 1
        if _is_yes(r.get("见客")):
            a["visits"] += 1
        if str(r.get("signed") or "").strip() == "是":
            a["signs"] += 1
        if _is_yes(r.get("加微")):
            a["wx"] += 1
        a["sign_perf"] += _to_num(r.get("成单业绩")) or 0
    out = {}
    for nm, a in acc.items():
        n = a["resources"]
        out[nm] = {
            "resources": n,
            "visits": a["visits"],
            "signs": a["signs"],
            "sign_rate": round(a["signs"] / n, 4) if n else None,
            "visit_rate": round(a["visits"] / n, 4) if n else None,
            "wx": a["wx"],
            "wx_rate": round(a["wx"] / n, 4) if n else None,
            "sign_perf": round(a["sign_perf"], 2),
        }
    return out


def _aggregate_interview_by_person():
    """面谈明细 → 按「私客所属人姓名」聚合三项陪同见客签单率（签单=首次收款时间非空）。

    返回 {姓名: {"业务经理"|"中台"|"管理者": {n, signed, rate}}}
    """
    rows, _p = _load_interview()
    out = {}
    if not rows:
        return out

    def _nz(v):
        return v is not None and str(v).strip() != ""

    cols = set(rows[0].keys())
    sign_col = "首次收款时间" if "首次收款时间" in cols else None
    for d in rows:
        nm = str(d.get("私客所属人姓名") or "").strip()
        if not nm:
            continue
        role = str(d.get("岗位") or "").strip()
        if role not in ("业务经理", "中台", "管理者"):
            continue
        signed = _nz(d.get(sign_col)) if sign_col else (_nz(d.get("客单号")) or _nz(d.get("订单号")))
        s = out.setdefault(nm, {}).setdefault(role, {"n": 0, "signed": 0, "rate": None})
        s["n"] += 1
        if signed:
            s["signed"] += 1
    for nm, e in out.items():
        for role, s in e.items():
            s["rate"] = round(s["signed"] / s["n"], 4) if s["n"] else None
    return out


# 漏斗关键指标（展示用）：名称, 是否比率, 说明
FUNNEL_METRICS = [
    ("演示用户", False, "近30天获得付费资源总数"),
    ("有效资源量", False, "有效+反无效成功的资源量"),
    ("日人均资源量", False, "日均资源量"),
    ("反无效资源占比", True, "反无效成功量/演示用户"),
    ("8分钟跟进率(含备注与接通）", True, "收到资源8分钟内响应率"),
    ("资源24小时内接触率（仅计算接通与加微）", True, "24小时内接通或加微比例"),
    ("3日连跟率", True, "首联后连续3日跟进客户占比"),
    ("资源加微率", True, "加微量/演示用户"),
    ("30s有效沟通占比", True, "有效沟通≥30秒占比"),
    ("付费客30天及时转化率", True, "30天内见客或成单占比"),
    ("接收资源30天内见客量", False, "30天内见客数"),
    ("接收资源30天内见客后的转化率", True, "见客后签单转化率"),
    ("付费客30天签单转化率", True, "30天签单转化率"),
    ("销售额", False, "近30天销售额"),
    ("成交顾客数", False, "近30天成交顾客数"),
    ("客单价", False, "客单价"),
    ("投产比", False, "投产比"),
]

# LTC 关键指标（展示用）
LTC_METRICS = [
    ("三力达标", False, "当月三力达标情况"),
    ("月累计实算利润业绩", False, "月至今累计实算利润业绩"),
    ("月累计核算利润业绩", False, "月至今累计核算利润业绩"),
    ("当日新增业绩", False, "当日新增业绩"),
    ("当日计划拨打量", False, "当日常规计划拨打量"),
    ("当日计划拨打执行量", False, "当日计划拨打执行量"),
    ("当日计划拨打接触量", False, "当日计划拨打接触量"),
    ("当日计划拨打接触率", True, "当日接触量/执行量"),
    ("近30天计划拨打(180天内)首次拨打占比", True, "近30天首拨占比"),
    ("30天计划拨打转私客量", False, "近30天计划拨打转私客量"),
    ("当日全部电话拨打量", False, "当日全部电话拨打量"),
    ("当日全部电话接通量", False, "当日全部电话接通量"),
    ("当日全部电话接通率", True, "当日接通量/拨打量"),
    ("当日全部电话30秒占比", True, "当日接通≥30秒占比"),
    ("当日平均通时（秒）", False, "当日平均通时(秒)"),
    ("近30天电话接通量", False, "近30天电话接通量"),
    ("近30天全部通时（分钟）", False, "近30天通时(分钟)"),
    ("90天加微量", False, "近90天加微量"),
    ("90天内新增微信绑定见客或签单率", True, "加微后见客/签单率"),
    ("累计客户量", False, "累计客户量"),
    ("累计成单客户量", False, "累计成单客户量"),
    ("累计优质客户量(4-6星)", False, "累计4-6星优质客户"),
    ("近30天新增4-6星客户量", False, "近30天新增4-6星客户"),
    ("近30天高质量微信反馈量", False, "近30天高质量微信反馈"),
    ("近30天1-3星掉库率", True, "近30天1-3星掉库率"),
    ("近30天公共库拾回掉库率", True, "近30天公共库拾回掉库率"),
    ("近30天见客量", False, "近30天见客量"),
    ("近30天合格高价值见客", False, "近30天合格高价值见客"),
    ("近30天见客转化率", True, "近30天见客转化率"),
    ("近90天见客转化率", True, "近90天见客转化率"),
    ("近30天签单总量", False, "近30天签单总量"),
    ("近30天付费资源分配量", False, "近30天付费资源分配量"),
    ("付费客30天及时转化率", True, "付费客30天及时转化率"),
    ("近30天付费资源加微率", True, "近30天付费资源加微率"),
]


def _sug(level, area, text):
    return {"level": level, "area": area, "text": text}


def build_advice(emp):
    """按员工数据生成针对性建议。emp 已含 funnel 与 ltc 字典。"""
    sug = []
    f = emp.get("funnel") or {}
    l = emp.get("ltc") or {}
    sanli = emp.get("三力达标") or ""
    months = emp.get("司龄") or 0
    post = emp.get("岗位") or ""

    def v(d, k):
        x = d.get(k)
        return x if isinstance(x, (int, float)) else None

    def is_manager():
        for kw in ("总经理", "副总", "经理", "主管"):
            if kw in post:
                return True
        return False

    # ---------- 管理岗：管理视角 ----------
    if is_manager():
        # 团队负责人：关注团队破零与人均产出
        if "主管" in post or "经理" in post or "副总" in post or "总经理" in post:
            sug.append(_sug("good", "管理", "作为{}，建议把精力聚焦在团队漏斗管理与员工陪访：抓团队8分钟跟进率、3日连跟率两项关键动作，并针对未破零员工安排1对1陪跑。".format(post)))
            return sug

    # ---------- 三力 ----------
    if sanli in ("未破零",):
        if months is not None and months >= 3:
            sug.append(_sug("warn", "三力", "入职已 {} 个月仍未破零，需优先突破首单：聚焦高意向付费客，用「3日连跟+见客逼单」跑通第一单。".format(round(months, 1))))
        else:
            sug.append(_sug("info", "三力", "当前未破零。建议先保证每日基础活动量（拨打/加微/见客），用最短路径跑通首单建立信心。"))
    elif sanli in ("破零人力",):
        sug.append(_sug("good", "三力", "已达成破零人力，保持当前节奏，冲刺标准人力（提升付费客30天及时转化率）。"))
    elif sanli in ("标准人力",):
        sug.append(_sug("good", "三力", "已达成标准人力，建议深挖高价值客户（4-6星）与转介绍，向英雄人力冲刺。"))
    elif sanli in ("英雄人力",):
        sug.append(_sug("good", "三力", "已达成英雄人力，发挥标杆作用：沉淀打法并带教新员工。"))
    elif sanli in ("无三力",):
        sug.append(_sug("info", "三力", "当前无三力考核口径，建议参照部门标杆的漏斗动作保持活跃度。"))

    # ---------- 漏斗动作 ----------
    def rate(d, k):
        x = v(d, k)
        return x if x is not None and 0 <= x <= 1.5 else None

    r8 = rate(f, "8分钟跟进率(含备注与接通）")
    if r8 is not None and r8 < 0.8:
        sug.append(_sug("warn", "漏斗", "8分钟跟进率仅 {}%，低于80%目标。资源到手即行动：收到后立即备注+电话，抢黄金8分钟。".format(round(r8 * 100, 1))))
    r24 = rate(f, "资源24小时内接触率（仅计算接通与加微）")
    if r24 is not None and r24 < 0.8:
        sug.append(_sug("warn", "漏斗", "24小时接触率 {}%，资源老化风险高。当天资源当天必须接通或加微，避免转入掉库。".format(round(r24 * 100, 1))))
    r3 = rate(f, "3日连跟率")
    if r3 is not None and r3 < 0.6:
        sug.append(_sug("warn", "漏斗", "3日连跟率 {}%，首联后跟进断裂。建议为每个首联客户排未来3天跟进计划，每天至少1次有效沟通。".format(round(r3 * 100, 1))))
    rw = rate(f, "资源加微率")
    if rw is not None and rw < 0.6:
        sug.append(_sug("warn", "漏斗", "加微率 {}%，微信触达弱。优化开场话术，用利益点+快速响应引导加微。".format(round(rw * 100, 1))))
    rs30 = rate(f, "30s有效沟通占比")
    if rs30 is not None and rs30 < 0.6:
        sug.append(_sug("warn", "漏斗", "30s有效沟通占比 {}%，通话质量待提升。开场先探需求再讲产品，控制通时≥30秒。".format(round(rs30 * 100, 1))))
    rt30 = rate(f, "付费客30天及时转化率")
    if rt30 is not None and rt30 < 0.2:
        sug.append(_sug("warn", "漏斗", "付费客30天及时转化率 {}%，线索变现慢。优先跟进已接触未成交客户，30天内推进见客。".format(round(rt30 * 100, 1))))
    rv = rate(f, "接收资源30天内见客后的转化率")
    rv_visit = v(f, "接收资源30天内见客量") or 0
    if rv_visit > 0 and rv is not None and rv < 0.3:
        sug.append(_sug("warn", "漏斗", "见客转化率 {}%，见客后逼单不足。见客后24小时内必须出方案/报价并约定下一步。".format(round(rv * 100, 1))))

    # ---------- LTC 活动量 ----------
    dcall = v(l, "当日全部电话拨打量") or 0
    dexec = v(l, "当日计划拨打执行量") or 0
    dcontact = rate(l, "当日计划拨打接触率")
    if dcall == 0 and dexec == 0:
        sug.append(_sug("info", "LTC", "当日电话拨打量为0，活动量不足。建议设定每日最低拨打量目标并自检完成。"))
    elif dexec > 0 and dcontact is not None and dcontact < 0.5:
        sug.append(_sug("warn", "LTC", "计划拨打接触率 {}%，名单质量或通话策略待优化，增加名单量并提高有效拨打量。".format(round(dcontact * 100, 1))))
    d30rate = rate(l, "当日全部电话30秒占比")
    d30call = v(l, "当日全部电话接通量") or 0
    if d30call > 0 and d30rate is not None and d30rate < 0.5:
        sug.append(_sug("warn", "LTC", "当日30秒通话占比 {}%，有效沟通偏少。提升话术质量，先建立信任再切入业务。".format(round(d30rate * 100, 1))))
    trans = v(l, "30天计划拨打转私客量")
    if trans is not None and trans == 0:
        sug.append(_sug("info", "LTC", "近30天计划拨打转私客量为0，线索未沉淀。通话后及时标记意向，转私客建立长期触达。"))

    # ---------- 业绩 ----------
    month_profit = v(l, "月累计实算利润业绩")
    if month_profit is not None and month_profit <= 0:
        sug.append(_sug("warn", "业绩", "本月实算利润业绩为 {}，尚未产生正业绩。需加速推进在谈客户成单，或复盘近期流失原因。".format(month_profit)))

    # 无任何建议的兜底
    if not sug:
        sug.append(_sug("good", "综合", "各核心指标表现良好，无明显短板。建议保持节奏，向更高三力等级冲刺。"))
    return sug


def build_dept_advice(b):
    """按部门聚合数据生成部门级建议"""
    sug = []
    n = b["人数"]
    sanli = b["三力"]
    unbroken = sanli.get("未破零", 0)
    no_sanli = sanli.get("无三力", 0)
    real = b.get("月实算业绩") or 0
    fa = b.get("funnel_avg") or {}
    la = b.get("ltc_avg") or {}

    if n > 0 and unbroken / n >= 0.6:
        sug.append(_sug("warn", "三力", "部门 {}/{} 员工尚未破零，破零率偏低。建议管理层建立「首单陪跑」机制：对入职超3个月未破零员工一对一辅导，集中高意向线索优先分配。".format(unbroken, n)))
    if no_sanli > 0:
        sug.append(_sug("info", "三力", "有 {} 人无三力考核口径（多为提成/交付/管理层），建议重点关注有考核指标员工的实际漏斗动作。".format(no_sanli)))

    r8 = fa.get("8分钟跟进率(含备注与接通）")
    if isinstance(r8, (int, float)) and 0 < r8 < 0.8:
        sug.append(_sug("warn", "漏斗", "部门平均8分钟跟进率 {}%，低于80%目标。建议例会通报个人跟进率，营造即时响应文化。".format(round(r8 * 100, 1))))
    r3 = fa.get("3日连跟率")
    if isinstance(r3, (int, float)) and 0 < r3 < 0.6:
        sug.append(_sug("warn", "漏斗", "部门平均3日连跟率 {}%，首联后跟进断层明显。建议强制排「未来3天跟进计划」，主管抽查跟进记录。".format(round(r3 * 100, 1))))
    rw = fa.get("资源加微率")
    if isinstance(rw, (int, float)) and 0 < rw < 0.6:
        sug.append(_sug("warn", "漏斗", "部门平均加微率 {}%，微信触达偏弱。建议统一开场白+加微钩子，午会演练加微话术。".format(round(rw * 100, 1))))
    rt30 = fa.get("付费客30天及时转化率")
    if isinstance(rt30, (int, float)) and 0 < rt30 < 0.2:
        sug.append(_sug("warn", "漏斗", "部门付费客30天及时转化率 {}%，线索变现慢。建议聚焦「30天内见客」专项，周度复盘漏斗各环节。".format(round(rt30 * 100, 1))))

    dcall = la.get("当日全部电话拨打量")
    if isinstance(dcall, (int, float)) and dcall == 0:
        sug.append(_sug("info", "LTC", "部门当日电话拨打量平均为0，整体活动量偏低。建议设定部门最低拨打量底线并每日公示完成率。"))

    if real <= 0 and n > 0:
        sug.append(_sug("warn", "业绩", "部门月累计实算业绩为 {}，整体产出待提升。建议优先抓高意向客户的见客与逼单，同时复盘线索分配效率。".format(real)))

    if not sug:
        sug.append(_sug("good", "综合", "部门整体指标表现良好，无显著短板。建议保持并沉淀最佳实践，向更高目标冲刺。"))
    return sug


# ===== 重点指标三大模块（网销 / 电邀 / 会计） =====
# 网销 13 项 / 电邀 7 项 / 会计 5 项
# 口径约定：
#   比例类保留 1 位小数；金额类保留整数；分母为 0 显示 "-" 且不参与达标判定。
#   正向指标：实际 ≥ 目标 → 绿；目标*(1-0.05) ≤ 实际 < 目标 → 黄；更低 → 红。
#   区间型（业务线签单转化率）：<区间下限 → 红；区间内 → 绿；>区间上限 → 绿但标注"超上限"。
#   反向指标（非正常流失率）：≤目标 → 绿；>目标 → 红。

BIZ_LINE_RULES = [
    ("工商",    {"一级业态": "工商服务"}),
    ("跨境",    {"一级业态": "跨境业务"}),
    ("法律",    {"一级业态": "法律服务"}),
    ("AI智能体", {"三级业态": "AI数智化"}),
    ("交易",    {"产品归类": "资质交易"}),
]
BIZ_TARGETS = {
    "工商":    {"kind": "interval", "lower": 0.15, "upper": 0.24, "display": "15%~24%（基准15%）"},
    "跨境":    {"kind": "interval", "lower": 0.08, "upper": 0.15, "display": "8%~15%（基准8%）"},
    "法律":    {"kind": "interval", "lower": 0.07, "upper": 0.10, "display": "7%~10%（基准7%）"},
    "AI智能体": {"kind": "interval", "lower": 0.04, "upper": 0.07, "display": "4%~7%（基准4%）"},
    "交易":    {"kind": "interval", "lower": 0.03, "upper": 0.05, "display": "3%~5%（基准3%）"},
}


def _kv(d, *subs):
    """从会计 sheet 行字典中查找 header 同时包含全部子串的字段值。"""
    for k, v in d.items():
        if all(s in k for s in subs):
            return v
    return None


def _status_pos(value, target, band=0.05):
    if value is None or target is None:
        return ("na", "不参与")
    if value >= target:
        return ("pass", "达标")
    if value >= target * (1 - band):
        return ("warn", "预警")
    return ("fail", "不达标")


def _status_rev(value, target, band=0.05):
    if value is None or target is None:
        return ("na", "不参与")
    if value <= target:
        return ("pass", "达标")
    if value <= target * (1 + band):
        return ("warn", "预警")
    return ("fail", "不达标")


def _status_interval(value, lower, upper):
    if value is None:
        return ("na", "不参与")
    if value < lower:
        return ("fail", "不达标·低于下限")
    if value <= upper:
        return ("pass", "达标·区间内")
    return ("pass", "达标·超上限")


def _fmt_ratio(v):
    """比率 → 百分比字符串。四舍五入（半值向上），与前端 (v*100).toFixed(1) 保持一致。

    注意：不能用内置 round()/「%.1f」，它们是 banker's rounding（6.25→6.2%），
    会与前端 toFixed(1)（6.25→6.3%）出现同一指标两个数字。
    """
    if v is None:
        return "-"
    import math
    return "%.1f%%" % (math.floor(v * 1000 + 0.5) / 10.0)


def _fmt_amount(v):
    if v is None:
        return "-"
    return "%d" % round(v)


def _fmt_count(v):
    if v is None:
        return "-"
    return "%d" % round(v)


def _sum_funnel(emps, key):
    s = 0.0
    n = 0
    for e in emps:
        v = (e.get("funnel") or {}).get(key)
        if isinstance(v, (int, float)):
            s += v
            n += 1
    return s, n


def _load_pay_resources():
    """付费客资源 → 解析后行列表（已剥离排除规则后剩余的也保留，便于窗口统计）；
    返回 (rows, max_date)。每行含 date/valid/code/invalid/anti_invalid/signed/b1/b2/b3/prod/
    当前私客所属人(AS)/员工所属部门(AX)。

    ⚠️ 2026-09-21：源表 15:12 版起**移除「付费客资源」sheet**。用户当天明确口径：
    「资源以漏斗数据中的有效资源量即可」⇒ 该 sheet 缺失时返回 ([], None)，资源类指标
    一律改由 _funnel_res_agg()（源表「网销漏斗」逐人列）派生，不再抛 KeyError。
    """
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    if "付费客资源" not in wb.sheetnames:
        return [], None
    ws = wb["付费客资源"]
    rows = list(ws.iter_rows(values_only=True))
    headers = [(str(c).strip() if c is not None else "") for c in rows[1]]
    idx = {h: i for i, h in enumerate(headers) if h}
    out = []
    max_date = None
    for r in rows[2:]:
        if not r or r[0] is None:
            continue
        d = r[0]
        if not isinstance(d, datetime):
            continue
        if max_date is None or d > max_date:
            max_date = d
        out.append({
            "date": d,
            "valid": r[idx.get("有效资源量", 5)],
            "code": r[idx.get("资源编码", 7)],
            "invalid": r[idx.get("是否无效", 37)],
            "anti_invalid": r[idx.get("是否反无效", 38)],
            "signed": r[idx.get("是否成单", 26)],
            "一级业态": r[idx.get("一级业态", 9)],
            "二级业态": r[idx.get("二级业态", 10)],
            "三级业态": r[idx.get("三级业态", 11)],
            "产品归类": r[idx.get("产品归类", 48)],
            "当前私客所属人": r[idx.get("当前私客所属人", 44)],
            "员工所属部门": r[idx.get("员工所属部门", 49)],
            "见客": r[idx.get("是否见客", 41)],
            "加微": r[idx.get("是否加微", 42)],
            "邀约": r[idx.get("是否邀约", 40)],
            "高质量": r[idx.get("高质量资源", 29)],
            "转介绍": r[idx.get("是否转介绍", 39)],
            "拨打次数": r[idx.get("拨打次数", 30)],
            "接通次数": r[idx.get("接通次数", 32)],
            "s30接通": r[idx.get("30秒以上接通次数", 34)],
            "通时": r[idx.get("通时（单位：min）", 36)],
            "成单业绩": r[idx.get("成单业绩", 27)],
            "利润业绩": r[idx.get("利润业绩", 25)],
        })
    return out, max_date


def _pay_excluded(row):
    """四条排除规则；满足任一条即排除。"""
    if _to_num(row["valid"]) == 0:
        return True
    if row["code"] is None or str(row["code"]).strip() == "":
        return True
    if str(row.get("invalid") or "").strip() == "是":
        return True
    if str(row.get("anti_invalid") or "").strip() == "是":
        return True
    return False


def _is_yes(v):
    """付费客资源中"是/否"类字段判定：兼容 int 1 / '1' / '是' / bool，其余视为否。"""
    if v is None:
        return False
    if v is True:
        return True
    s = str(v).strip()
    return s in ("1", "1.0", "是", "True", "true", "Y", "y")


def _bizline_of(row):
    """业务线匹配（命中 BIZ_LINE_RULES 中第一条）"""
    for name, rule in BIZ_LINE_RULES:
        ok = True
        for col, val in rule.items():
            if str(row.get(col) or "").strip() != val:
                ok = False
                break
        if ok:
            return name
    return None


def _load_dy_aggregates():
    """电邀漏斗层级行（分公司/事业部/部门）的指标字典，key 为 col[0] 标签。"""
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    if "电邀漏斗" not in wb.sheetnames:
        return {}
    ws = wb["电邀漏斗"]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 4:
        return {}
    headers = [(str(c).strip() if c is not None else "") for c in rows[1]]
    import re
    pat = re.compile(r"^(.+?)\((\d+)\)\s*$")
    out = {}
    for r in rows[3:]:
        if not r or r[0] is None:
            continue
        label = str(r[0]).strip()
        if pat.match(label):
            continue
        emp = {}
        for j in range(1, len(headers)):
            h = headers[j]
            if h:
                emp[h] = _to_num(r[j])
        out[label] = emp
    return out


# 会计三表（多级表头）锚点判据：每列从锚点行向上取第一个非空单元格作为表头名。
KJ_ANCHORS = {
    "S0(会计)":     ["8月账务工作价值量积分", "9月月度实算业绩(含续费)"],
    "规划师(会计)": ["非正常流失户数(8月)", "到期续费率(家数)(8月)"],
    "主管(会计)":   ["个人非正常流失户数(8月)", "团队月人均账务工作价值量积分(8月)"],
}
KJ_SHEET_NAMES = ("S0(会计)", "规划师(会计)", "主管(会计)")


def _parse_kj_sheet(ws, name):
    """解析单张会计表（多级表头）。

    表头结构：行 4 分组（如「基础信息」「各项指标实际值」）+ 行 5 子表头
    （工号/姓名/.../8月/9月）+ 行 6 完整指标名（非空/含括号/8月后缀）。
    合并策略：每个列位置，从锚点行向上找第一个非空单元格作为 header。
    """
    rows = list(ws.iter_rows(values_only=True))
    anchor_idx = None
    for i, r in enumerate(rows[:8]):
        cells = {str(c).strip() for c in r if c is not None}
        if any(a in cells for a in KJ_ANCHORS.get(name, [])):
            anchor_idx = i
            break
    if anchor_idx is None:
        # fallback：非空最多的行
        best_count = 0
        for i, r in enumerate(rows[:8]):
            non_empty = [c for c in r if c is not None and str(c).strip()]
            if len(non_empty) > best_count:
                best_count = len(non_empty)
                anchor_idx = i
    if anchor_idx is None:
        return []
    # 合并表头：从锚点行向上取每个列位置第一个非空
    n_cols = max((len(r) for r in rows[:anchor_idx + 1] if r), default=0)
    merged = [""] * n_cols
    for col in range(n_cols):
        for ri in range(anchor_idx, -1, -1):
            if ri < len(rows) and col < len(rows[ri]) and rows[ri][col] is not None and str(rows[ri][col]).strip():
                merged[col] = str(rows[ri][col]).strip()
                break
    data = []
    for r in rows[anchor_idx + 1:]:
        if not r or r[0] is None:
            continue
        d = {}
        for j, v in enumerate(r):
            if j < len(merged) and merged[j]:
                d[merged[j]] = v
        data.append(d)
    return data


def _load_kuaiji():
    """会计三表（S0/规划师/主管）解析为 dict 列表。

    ⚠️ 2026-09-21（源表改版）：《某公司-人员数据分析路径.xlsx》换版后**不再包含**
    这三张表（12 sheet 仅剩 会计高端账签单 / 会计监控表 / 会计续费明细表），
    导致会计页「员工档案（10 人）」与「客服/顾问/主管 演示用户」全部落空
    （total=0 / departments=[] / employees=[]）。

    改为 **源表优先 · 辅助表兜底**：
      ① 源表 XLSX 含 S0/规划师/主管(会计) 时按原逻辑读；
      ② 缺失的表由辅助表《分公司健康表/四大模式-重点指标表格.xlsx》补齐 ——
         该簿同样含这三张表，历史核对为**逐格相同**（见 `数据源对照梳理…md` §3.1
         「内容完全相同，纯冗余」），且它本就是会计板块（票据抽检 / 流程会计个人出错率 /
         会计续费明细）的既有数据源，口径天然一致。
    """
    result = {}
    for path in (XLSX, _net_summary_path()):
        if not path or not os.path.exists(path):
            continue
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception:
            continue
        for name in KJ_SHEET_NAMES:
            if result.get(name) or name not in wb.sheetnames:
                continue
            data = _parse_kj_sheet(wb[name], name)
            if data:
                result[name] = data
        del wb
    return result


def _kj_sheet_rows(kuaiji, keyword):
    """从动态加载的会计 sheets 中，取 sheet 名含 keyword 的表的数据行。"""
    for name, rows in (kuaiji or {}).items():
        if keyword in name:
            return rows or []
    return []


def _is_hangzhou(row):
    port = str(row.get("端口") or "")
    return "<CITY>" in port


# ---------- 会计新增三表（2026-09-16 需求 N）：票据抽检 / 个人出错率 / 续费明细 ----------
# 数据源与「网销漏斗」同一本工作簿（四大模式-重点指标表格.xlsx），故复用 _net_summary_path()；
# 该表 mtime 已由 _aux_mtimes() 写进 ry_*.json 的 meta.aux_mtime_ts → 自动同步链天然覆盖。
# 列一律按表头名定位，不写死列号（换导出批次列位漂移也不会错）。
_KJ_EXTRA_CACHE = {}


def _kj_num(v):
    """会计新表数值清洗：'' / '#DIV/0!' / 文本一律 None。"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("%", "")
    if (not s) or s.startswith("#"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _kj_txt(v):
    return "" if v is None else str(v).strip()


def _kj_person(v):
    """'演示用户 <ID>' → ('演示用户', '154958')；无工号 → (姓名, '')。"""
    s = _kj_txt(v)
    m = re.match(r"^(.*?)[\s\-_]*(\d{4,})$", s)
    if m:
        return m.group(1).strip(), m.group(2)
    return s, ""


def _kj_head_index(head):
    """表头行 → {列名: 列序号}（重名取首个）。"""
    idx = {}
    for i, h in enumerate(head):
        if h and h not in idx:
            idx[h] = i
    return idx


def _kj_parse_audit(wb):
    """「<CITY>会计抽检」→ 事业部级票据抽检汇总。

    表头跨第 2~3 行（第 3 行是子表头，如 col6 = '人工+AI' 补全上一行的 '抽检总量'），
    故每列把两行文本拼起来再按关键字匹配。
    """
    if "<CITY>会计抽检" not in wb.sheetnames:
        return None
    rows = list(wb["<CITY>会计抽检"].iter_rows(values_only=True))
    if len(rows) < 4:
        return None
    top = [_kj_txt(c) for c in rows[1]]
    sub = [_kj_txt(c) for c in rows[2]]
    ncol = max(len(top), len(sub))
    idx = {}
    for i in range(ncol):
        j = (sub[i] if i < len(sub) else "") + (top[i] if i < len(top) else "")
        if not j:
            continue
        if j == "AI家数":
            idx.setdefault("ai", i)
        elif j == "人工":                      # ⚠️ 必须精确等于，否则会命中「人工+AI」
            idx.setdefault("manual", i)
        elif "总家数" in j:
            idx.setdefault("total_families", i)
        elif "抽检总量" in j:
            idx.setdefault("check_total", i)
        elif "覆盖率" in j:
            idx.setdefault("coverage", i)
        elif "出错家数" in j:
            idx.setdefault("err_families", i)
        elif "出错率" in j:
            idx.setdefault("err_rate", i)
        elif j == "事业部":
            idx.setdefault("biz", i)
        elif j == "负责人":
            idx.setdefault("owner", i)

    def _g(r, key):
        i = idx.get(key)
        return r[i] if (i is not None and i < len(r)) else None

    items = []
    for r in rows[3:]:
        if not r or not _kj_txt(_g(r, "biz")):
            continue
        items.append({
            "事业部": _kj_txt(_g(r, "biz")),
            "负责人": _kj_txt(_g(r, "owner")),
            "AI家数": _kj_num(_g(r, "ai")),
            "人工家数": _kj_num(_g(r, "manual")),
            "总家数": _kj_num(_g(r, "total_families")),
            "抽检总量": _kj_num(_g(r, "check_total")),
            "覆盖率": _kj_num(_g(r, "coverage")),
            "出错家数": _kj_num(_g(r, "err_families")),
            "出错率": _kj_num(_g(r, "err_rate")),
        })
    if not items:
        return None

    def _s(key):
        return sum(x[key] for x in items if x[key] is not None)

    tot_f, tot_c, tot_e = _s("总家数"), _s("抽检总量"), _s("出错家数")
    title = _kj_txt(rows[0][0]) if rows[0] else ""
    mp = re.search(r"[\d.]+\s*[-—~至]\s*[\d.]+", title)
    return {
        "title": title,
        "period": mp.group(0) if mp else "",
        "items": items,
        "total": {
            "AI家数": _s("AI家数"), "人工家数": _s("人工家数"), "总家数": tot_f,
            "抽检总量": tot_c, "出错家数": tot_e,
            # 覆盖率 = 抽检总量 ÷ 总家数；出错率 = 出错家数 ÷ 抽检总量（与源表同口径）
            "覆盖率": (tot_c / tot_f) if tot_f else None,
            "出错率": (tot_e / tot_c) if tot_c else None,
        },
    }


def _kj_parse_err(wb):
    """「流程会计个人出错率」→ 逐人 检查量 / 出错量 / 出错率（含「新人」标记）。"""
    if "流程会计个人出错率" not in wb.sheetnames:
        return None
    rows = list(wb["流程会计个人出错率"].iter_rows(values_only=True))
    if len(rows) < 3:
        return None
    idx = _kj_head_index([_kj_txt(c) for c in rows[1]])
    i_name = idx.get("姓名")
    if i_name is None:
        return None
    i_new = 7   # 第 8 列无表头，值 = 「新人」标记

    def _g(r, key):
        i = idx.get(key)
        return r[i] if (i is not None and i < len(r)) else None

    recs = []
    for r in rows[2:]:
        if not r or i_name >= len(r) or not _kj_txt(r[i_name]):
            continue
        nm, gid = _kj_person(r[i_name])
        if not nm:
            continue
        check, err = _kj_num(_g(r, "检查量")), _kj_num(_g(r, "出错量"))
        recs.append({
            "姓名": nm, "工号": gid,
            "事业部": _kj_txt(_g(r, "事业部")),
            "部门": _kj_txt(_g(r, "部门")),
            "岗位": _kj_txt(_g(r, "岗位")),
            "检查量": check, "出错量": err,
            "出错率": _kj_num(_g(r, "出错率")),
            # 源表分母为 0 时是 '#DIV/0!' → 与「出错率真的是 0」区分开
            "出错率无分母": (_kj_num(_g(r, "出错率")) is None and (check or 0) == 0),
            "新人": bool(_kj_txt(r[i_new])) if i_new < len(r) else False,
        })
    if not recs:
        return None
    tc = sum(x["检查量"] or 0 for x in recs)
    te = sum(x["出错量"] or 0 for x in recs)
    return {
        "title": _kj_txt(rows[0][0]) if rows[0] else "",
        "records": recs,
        "total": {"检查量": tc, "出错量": te, "出错率": (te / tc) if tc else None,
                  "人数": len(recs), "新人": sum(1 for x in recs if x["新人"])},
    }


# 会计续费明细表字段（客户级一行一条）
KJ_RENEW_FIELDS = [
    "续费月份", "公司名称", "公司状态", "事业部", "会计主管", "会计主管部门",
    "会计顾问", "会计顾问部门", "订单编号", "服务开始时间", "服务结束时间",
    "月代理费（参考）", "付款方式", "应收续费金额", "实际收续费金额",
    "是否续费", "是否应续费", "是否往月欠费",
]
KJ_RENEW_AMT_KEY = "月代理费"      # 表头可能写作「月代理费（参考）」→ 用包含匹配


def _kj_parse_renew(wb):
    """「会计续费明细表」→ 客户级明细 + 按公司状态 / 是否续费 / 是否应续费的分类统计。

    口径：
      · 是否流失  = 「公司状态」含「流失」（如 流失中）
      · 是否续费  = 源表「是否续费」列
      · 是否应续费 = 源表「是否应续费」列
      · 续费率（全量）= 续费=是 的家数 ÷ 全部家数；应续续费率 = 应续=是 且 续费=是 ÷ 应续=是
    """
    if "会计续费明细表" not in wb.sheetnames:
        return None
    rows = list(wb["会计续费明细表"].iter_rows(values_only=True))
    if len(rows) < 3:
        return None
    head = [_kj_txt(c) for c in rows[1]]
    idx = _kj_head_index(head)
    i_amt = None
    for i, h in enumerate(head):
        if h and KJ_RENEW_AMT_KEY in h:
            i_amt = i
            break

    def _g(r, key):
        i = idx.get(key)
        return r[i] if (i is not None and i < len(r)) else None

    recs = []
    for r in rows[2:]:
        if not r or not _kj_txt(_g(r, "公司名称")):
            continue
        recs.append({
            "续费月份": _kj_txt(_g(r, "续费月份")),
            "公司名称": _kj_txt(_g(r, "公司名称")),
            "公司状态": _kj_txt(_g(r, "公司状态")),
            "事业部": _kj_txt(_g(r, "事业部")),
            "会计主管": _kj_txt(_g(r, "会计主管")),
            "会计主管部门": _kj_txt(_g(r, "会计主管部门")),
            "会计顾问": _kj_txt(_g(r, "会计顾问")),
            "会计顾问部门": _kj_txt(_g(r, "会计顾问部门")),
            "订单编号": _kj_txt(_g(r, "订单编号")),
            "服务开始时间": _kj_txt(_g(r, "服务开始时间")),
            "服务结束时间": _kj_txt(_g(r, "服务结束时间")),
            "月代理费": _kj_num(r[i_amt]) if (i_amt is not None and i_amt < len(r)) else None,
            "付款方式": _kj_txt(_g(r, "付款方式")),
            "应收续费金额": _kj_num(_g(r, "应收续费金额")),
            "实际收续费金额": _kj_num(_g(r, "实际收续费金额")),
            "是否续费": _kj_txt(_g(r, "是否续费")),
            "是否应续费": _kj_txt(_g(r, "是否应续费")),
            "是否往月欠费": _kj_txt(_g(r, "是否往月欠费")),
        })
    if not recs:
        return None

    def _yes(d, k):
        return _kj_txt(d.get(k)) == "是"

    def _agg(items):
        return {
            "家数": len(items),
            "应收": round(sum(d["应收续费金额"] or 0 for d in items), 2),
            "实收": round(sum(d["实际收续费金额"] or 0 for d in items), 2),
        }

    total = len(recs)
    statuses = []
    for d in recs:
        s = d["公司状态"] or "（空）"
        if s not in statuses:
            statuses.append(s)

    by_status = []
    for s in statuses:
        g = [d for d in recs if (d["公司状态"] or "（空）") == s]
        a = _agg(g)
        a["公司状态"] = s
        a["占比"] = (a["家数"] / total) if total else None
        by_status.append(a)

    lost = [d for d in recs if "流失" in (d["公司状态"] or "")]
    alive = [d for d in recs if "流失" not in (d["公司状态"] or "")]
    renewed = [d for d in recs if _yes(d, "是否续费")]
    due = [d for d in recs if _yes(d, "是否应续费")]
    due_renewed = [d for d in due if _yes(d, "是否续费")]

    cross = []
    for s in statuses:
        for rf in ("是", "否"):
            for df in ("是", "否"):
                g = [d for d in recs if (d["公司状态"] or "（空）") == s
                     and d["是否续费"] == rf and d["是否应续费"] == df]
                if not g:
                    continue
                a = _agg(g)
                a.update({"公司状态": s, "是否续费": rf, "是否应续费": df,
                          "占比": (a["家数"] / total) if total else None})
                cross.append(a)

    def _by_person(key, role):
        groups = OrderedDict()
        for d in recs:
            nm, gid = _kj_person(d.get(key))
            if not nm:
                continue
            groups.setdefault((nm, gid), []).append(d)
        out = []
        for (nm, gid), g in groups.items():
            due_g = [d for d in g if _yes(d, "是否应续费")]
            ren_g = [d for d in g if _yes(d, "是否续费")]
            due_ren = [d for d in due_g if _yes(d, "是否续费")]
            a = _agg(g)
            a.update({
                "姓名": nm, "工号": gid, "角色": role,
                "应续家数": len(due_g), "已续家数": len(ren_g), "应续已续": len(due_ren),
                "续费率": (len(ren_g) / len(g)) if g else None,
                "应续续费率": (len(due_ren) / len(due_g)) if due_g else None,
                "流失家数": len([d for d in g if "流失" in (d["公司状态"] or "")]),
            })
            out.append(a)
        out.sort(key=lambda x: (-x["家数"], x["姓名"]))
        return out

    ar = round(sum(d["应收续费金额"] or 0 for d in recs), 2)
    ac = round(sum(d["实际收续费金额"] or 0 for d in recs), 2)
    mm = re.search(r"(\d{4}-\d{1,2})", _kj_txt(rows[0][0]) if rows[0] else "")
    return {
        "title": _kj_txt(rows[0][0]) if rows[0] else "",
        "month": mm.group(1) if mm else "",
        "scripts": recs,
        "stats": {
            "total": total,
            "lost": dict(_agg(lost), **{"占比": (len(lost) / total) if total else None}),
            "alive": dict(_agg(alive), **{"占比": (len(alive) / total) if total else None}),
            "renew": {"是": len(renewed), "否": total - len(renewed),
                      "率": (len(renewed) / total) if total else None},
            "due": {"是": len(due), "否": total - len(due),
                    "占比": (len(due) / total) if total else None,
                    "应续已续": len(due_renewed),
                    "应续未续": len(due) - len(due_renewed),
                    "应续续费率": (len(due_renewed) / len(due)) if due else None},
            "amount": {"应收": ar, "实收": ac, "收现率": (ac / ar) if ar else None},
            "by_status": by_status,
            "cross": cross,
            "status_values": [s for s in statuses if s != "（空）"],
            # 源表该列整列为空 → 前端呈现「源表未提供」，不臆造
            "nodata": (["是否往月欠费"] if not any(d.get("是否往月欠费") for d in recs) else []),
        },
        "by_advisor": _by_person("会计顾问", "会计顾问"),
        "by_manager": _by_person("会计主管", "会计主管"),
    }


def _load_kj_extra():
    """读取三张新增会计 sheet（带缓存）：{"audit", "err", "renew"}。"""
    if _KJ_EXTRA_CACHE.get("loaded"):
        return _KJ_EXTRA_CACHE.get("data") or {}
    out = {"audit": None, "err": None, "renew": None}
    p = _net_summary_path()
    if p:
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            out["audit"] = _kj_parse_audit(wb)
            out["err"] = _kj_parse_err(wb)
            out["renew"] = _kj_parse_renew(wb)
        except Exception:
            out = {"audit": None, "err": None, "renew": None}
    _KJ_EXTRA_CACHE.update({"loaded": True, "data": out})
    return out


_KJ_MONITOR_CACHE = {}


def _load_kj_monitor():
    """「会计监控表」（主簿 XLSX）· 官方口径（用户 2026-09-28 确认）：

    非正常流失率 = 非正常流失家数 ÷ 最新保有量（源表即按此公式逐行计算）。
    返回行 dict 列表；试用/离职行的保有量与流失家数多为空 → None，合计时自动跳过。
    表头固定在 Excel 第 2 行（row index 1），列按表头名定位，不写死列号。
    """
    if _KJ_MONITOR_CACHE.get("loaded"):
        return _KJ_MONITOR_CACHE.get("rows") or []
    rows = []
    try:
        if XLSX and os.path.exists(XLSX):
            wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
            if "会计监控表" in wb.sheetnames:
                grid = list(wb["会计监控表"].iter_rows(values_only=True))
                head = [_kj_txt(c) for c in (grid[1] if len(grid) > 1 else [])]
                ix = {h: i for i, h in enumerate(head) if h}
                if all(k in ix for k in ("姓名", "最新保有量", "非正常流失家数")):

                    def _cell(r, k):
                        j = ix.get(k)
                        if j is None or j >= len(r):
                            return None
                        if k in ("最新保有量", "正常流失家数", "非正常流失家数", "非正常流失率"):
                            return _kj_num(r[j])
                        return _kj_txt(r[j])

                    for r in grid[2:]:
                        nm = _cell(r, "姓名")
                        if not nm:
                            continue
                        rows.append({
                            "姓名": nm,
                            "岗位": _cell(r, "岗位") or "",
                            "状态": _cell(r, "状态") or "",
                            "最新保有量": _cell(r, "最新保有量"),
                            "正常流失家数": _cell(r, "正常流失家数"),
                            "非正常流失家数": _cell(r, "非正常流失家数"),
                            "非正常流失率": _cell(r, "非正常流失率"),
                        })
            wb.close()
    except Exception:
        rows = []
    _KJ_MONITOR_CACHE.update({"loaded": True, "rows": rows})
    return rows


# ---------- 公司及部门业绩（用户 2026-09-28 新增 sheet，网销页展示<CITY>部分） ----------
_BIZ_PERF_COLS = ("集团业绩目标", "业绩完成", "当日业绩", "集团目标完成率", "排名",
                  "高价值客户目标", "高价值客户完成", "高价值客户完成率",
                  "月累计退款业绩", "去年同期退款业绩", "当日退款业绩",
                  "去年同期业绩", "业绩同比")
_BIZ_PERF_CACHE = {}


def _load_biz_perf():
    """「公司及部门业绩」sheet · 只取<CITY>公司与<CITY>各部门（网销数据分析页用）。

    sheet 结构：头部说明行 →「分公司数据」区块（总计/战区/各分公司行）→ 空行 →
    「分公司+部门明细」区块（公司行后紧跟其事业部行，遇下一公司行切换归属）。
    「公司行」识别：名字以「公司」结尾且不是「…公司其他」。
    返回 {"period", "company": {...}, "depts": [...]}；无数据返回 None。
    """
    if _BIZ_PERF_CACHE.get("loaded"):
        return _BIZ_PERF_CACHE.get("data")
    data = None
    try:
        if XLSX and os.path.exists(XLSX):
            wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
            if "公司及部门业绩" in wb.sheetnames:
                grid = list(wb["公司及部门业绩"].iter_rows(values_only=True))
                notes = [_kj_txt(r[1]) for r in grid[:6]
                         if len(r) > 1 and _kj_txt(r[1])]
                period = next((n[len("数据周期："):] for n in notes
                               if n.startswith("数据周期：")), "")
                head_i = None
                head = []
                for i, r in enumerate(grid):
                    cells = [_kj_txt(c).split("\n")[0] for c in (r or [])]
                    if "集团业绩目标" in cells:
                        head_i, head = i, cells
                        break
                if head_i is not None:
                    ix = {h: i for i, h in enumerate(head) if h}
                    nm_j = ix.get("事业部", 1)

                    def _row(r):
                        out = {"事业部": _kj_txt(r[nm_j]) if nm_j < len(r) else ""}
                        for h in _BIZ_PERF_COLS:
                            j = ix.get(h)
                            out[h] = (_kj_num(r[j]) if (j is not None and j < len(r)) else None)
                        return out

                    hz_company = None
                    hz_depts = []
                    cur = None
                    for r in grid[head_i + 1:]:
                        nm = _kj_txt(r[nm_j]) if nm_j < len(r) else ""
                        if not nm:
                            continue
                        if nm.endswith("公司") and not nm.endswith("公司其他"):
                            cur = nm
                            if nm == "<CITY>公司":
                                hz_company = _row(r)   # 出现两次时取部门明细区块那行
                            continue
                        if cur == "<CITY>公司":
                            hz_depts.append(_row(r))
                    if hz_company is not None:
                        data = {"period": period, "company": hz_company, "depts": hz_depts}
            wb.close()
    except Exception:
        data = None
    _BIZ_PERF_CACHE.update({"loaded": True, "data": data})
    return data


def _kj_err_lookup(name, code=""):
    """按「姓名+工号」取个人抽检出错率（工号优先；未命中按姓名唯一匹配）。"""
    err = (_load_kj_extra() or {}).get("err") or {}
    recs = err.get("records") or []
    if not recs or not name:
        return None
    nm, gid = str(name).strip(), str(code or "").strip()
    for r in recs:
        if r["姓名"] == nm and gid and r["工号"] == gid:
            return r
    cands = [r for r in recs if r["姓名"] == nm]
    return cands[0] if len(cands) == 1 else None


def _kj_renew_lookup(name, code=""):
    """按「姓名+工号」取个人续费聚合（会计顾问口径）。"""
    rn = (_load_kj_extra() or {}).get("renew") or {}
    rows = rn.get("by_advisor") or []
    if not rows or not name:
        return None
    nm, gid = str(name).strip(), str(code or "").strip()
    for r in rows:
        if r["姓名"] == nm and gid and r["工号"] == gid:
            return r
    cands = [r for r in rows if r["姓名"] == nm]
    return cands[0] if len(cands) == 1 else None


# ---------- 会计高端账签单占比 ----------
HIGH_END_CATS = ("星级财务外包新签", "星级财务外包续签")
HIGH_END_INCLUDE = ("税务管家", "财务外包")
HIGH_END_EXCLUDE = ("赠品",)
# 「认退款」列含「赠品」⇒ 该行同样按赠品剔除（用户 2026-09-21：赠品不计算在内）
HIGH_END_GIFT_FLAGS = ("赠品",)


def _load_high_end_account():
    """会计高端账签单占比。

    数据源：sheet「会计高端账签单」（业绩明细表 · 表头在第 1 行）。
    判定规则（用户 2026-09-21 口径 —— **并集 OR**，任一满足即算高端账）：
      ① 分类名称（三级分类）∈ {星级财务外包新签, 星级财务外包续签}
      ② 产品名称 / 商品名称 含「税务管家」或「财务外包」
      ③ **赠品不计算在内**（优先级最高）—— 两个判定位：
         · 产品名称 / 商品名称 含「赠品」字样
         · 「认退款」列含「赠品」（= 赠品增加）⇒ 该行同样剔除。
           实测演示用户 <CONTRACT_NO> 的 6 行虽分类为「星级财务外包新签」，
           但认退款全为「赠品增加」、业绩与核款全 0 ⇒ 整单不算高端账。
    ⚠️ 原实现为 ① AND ② 的交集，实测仅 1/20 = 5%；改为并集后为 2/20 = 10%，
       但按用户「赠品不算」口径剔除演示用户赠品单后为 **1/20 = 5%**。
    计算口径：以「认款账单编号」去重，同一账单编号只计一次；
             占比 = 高端账去重账单数 ÷ 全部去重账单数。
             客均（AQ 高端账签单客均）= 高端账「业绩」合计 ÷ 高端账单数
             （用户 2026-09-21：演示用户按业绩计算）；核款口径保留备查。
    端口：区域=某公司（与会计模块其他口径一致）。
    """
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    if "会计高端账签单" not in wb.sheetnames:
        return None
    ws = wb["会计高端账签单"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None
    headers = [(str(c).strip() if c is not None else "") for c in rows[0]]
    idx = {h: i for i, h in enumerate(headers) if h}

    def _cell(r, name, default):
        i = idx.get(name)
        if i is None:
            i = default
        return r[i] if i < len(r) else None

    all_bills = set()
    he_bills = set()
    he_perf = {}      # 账单号 → Σ高端账「业绩」
    he_credit = {}    # 账单号 → Σ高端账「核款金额」
    gift_bills = set()
    detail = []
    for r in rows[1:]:
        if not r:
            continue
        bill = str(_cell(r, "认款账单编号", 5) or "").strip()
        if not bill:
            continue
        region = str(_cell(r, "区域", 13) or "").strip()
        if region and region != "某公司":
            continue
        all_bills.add(bill)
        cat = str(_cell(r, "分类名称", 7) or "").strip()
        prod = (str(_cell(r, "产品名称", 9) or "") + "|" + str(_cell(r, "商品名称", 10) or ""))
        refund = str(_cell(r, "认退款", 11) or "")
        # 高端账 = 「三级分类 ∈ 星级财务外包新签/续签」OR「产品/商品名含 税务管家/财务外包」
        #          再排除赠品行（产品/商品名含「赠品」 **或** 认退款含「赠品」）
        if not (cat in HIGH_END_CATS or any(k in prod for k in HIGH_END_INCLUDE)):
            continue
        if any(k in prod for k in HIGH_END_EXCLUDE) or any(k in refund for k in HIGH_END_GIFT_FLAGS):
            gift_bills.add(bill)
            continue
        perf = _to_num(_cell(r, "业绩", 23)) or 0.0
        credit = _to_num(_cell(r, "核款金额", 25)) or 0.0
        he_perf[bill] = he_perf.get(bill, 0.0) + perf
        he_credit[bill] = he_credit.get(bill, 0.0) + credit
        if bill not in he_bills:
            he_bills.add(bill)
            detail.append({
                "bill": bill,
                "cat": cat,
                "product": str(_cell(r, "商品名称", 10) or _cell(r, "产品名称", 9) or "").strip(),
                "planner": str(_cell(r, "规划师", 19) or "").strip(),
                "signer": str(_cell(r, "签单人", 2) or "").strip(),
                "time": str(_cell(r, "业绩时间", 12) or "").strip()[:10],
            })
    total = len(all_bills)
    he = len(he_bills)
    perf_total = round(sum(he_perf.values()), 2)
    credit_total = round(sum(he_credit.values()), 2)
    for d in detail:                    # 每单金额回填（供前端明细列展示）
        d["perf"] = round(he_perf.get(d["bill"], 0.0), 2)
        d["credit"] = round(he_credit.get(d["bill"], 0.0), 2)
    return {
        "total_bills": total,
        "high_end_bills": he,
        "ratio": (round(he / total, 4) if total else None),
        "high_end_perf": perf_total,
        "high_end_credit": credit_total,
        # 客均（AQ）：用户 2026-09-21 「演示用户按业绩来计算」⇒ 主口径 = 业绩 ÷ 单数
        "per_bill_perf": (round(perf_total / he, 2) if he else None),
        "per_bill_credit": (round(credit_total / he, 2) if he else None),
        "gift_bills": sorted(gift_bills),
        "detail": detail,
        "cats": list(HIGH_END_CATS),
        "include": list(HIGH_END_INCLUDE),
        "exclude": list(HIGH_END_EXCLUDE),
    }


def _staff_split(emps):
    """把员工分成 管理岗 / 员工岗（用于演示用户 "管理+员工" 结构）

    M-29（用户 2026-09-17）：**「主管」不算管理者** —— 只有「经理」级及以上才算管理岗，
    故 mgr_keys 去掉「主管」（企服主管 / 企服储备主管 ⇒ 员工岗）。
    """
    mgr_keys = ("经理", "负责人", "总经理", "副总")
    mgrs = []
    staffs = []
    for e in emps:
        post = str(e.get("岗位") or "")
        if any(k in post for k in mgr_keys):
            mgrs.append(e)
        else:
            staffs.append(e)
    return mgrs, staffs


def _avg_num(vals):
    """求均值（None 自动忽略）。返回 (avg, count)。"""
    vs = [v for v in vals if isinstance(v, (int, float))]
    if not vs:
        return None, 0
    return sum(vs) / len(vs), len(vs)


def build_key_metrics(employees, pay_rows, max_date, kuaiji):
    """计算三大模块的重点指标。返回 {net, dy, kuaiji}。"""
    if max_date:
        wstart = max_date - timedelta(days=30)
    else:
        wstart = None

    # ===== 网销模块 =====
    net_emps = [e for e in employees if e.get("funnel_source") == "net"]
    net_biz = {"name": "演示用户", "sub": "考核月", "metrics": [], "notes": []}
    # 1-5 业务线签单转化率
    lines = {n: {"total": 0, "signed": 0} for n, _ in BIZ_LINE_RULES}
    for r in pay_rows:
        if _pay_excluded(r):
            continue
        if wstart and r["date"] < wstart:
            continue
        bl = _bizline_of(r)
        if bl is None:
            continue
        lines[bl]["total"] += 1
        if str(r.get("signed") or "").strip() == "是":
            lines[bl]["signed"] += 1
    if not pay_rows:
        # 2026-09-21：源表移除「付费客资源」sheet ⇒ 业态级回退值改由「网销漏斗」逐人列派生
        #（用户口径：资源以漏斗数据中的有效资源量即可）
        _dept2biz = {v: k for k, v in BIZ_TO_DEPT.items()}
        _fr_bl = {n: [0.0, 0.0] for n, _ in BIZ_LINE_RULES}
        for e in net_emps:
            _bl = _dept2biz.get(e.get("部门标准"))
            if _bl is None:
                continue
            _fr_bl[_bl][0] += _funnel_num(e, "有效资源量") or 0
            _fr_bl[_bl][1] += _funnel_num(e, "付费客30天签单量") or 0
        for n, _ in BIZ_LINE_RULES:
            lines[n]["total"] = int(round(_fr_bl[n][0]))
            lines[n]["signed"] = int(round(_fr_bl[n][1]))
    for name, _ in BIZ_LINE_RULES:
        total = lines[name]["total"]
        signed = lines[name]["signed"]
        rate = (signed / total) if total > 0 else None
        # 用户口径（2026-09-15 M-4）：「重点指标 · 演示用户」里的业态级卡片
        #   「付费客资源签单转化率-业态名称（近30天）」与分公司/事业部/员工三级同源
        #   —— 同样取「网销漏斗」A 列名称匹配行的 AL 列「付费客30天签单转化率」
        #   （业态 → 对应事业部行，如 工商 → <CITY>工商事业部）。
        #   未匹配到该行时才回退「付费客资源」sheet 聚合，保证不留空白。
        row_name, row_meta = _net_row_lookup(BIZ_TO_DEPT.get(name, name))
        if row_meta and row_meta.get("rate") is not None:
            rate = row_meta.get("rate")
            note = "来源：网销漏斗「%s」行 AL 列｜付费客30天签单量 %s ÷ 有效资源量 %s" % (
                row_name, _fmt_count(row_meta.get("sign")), _fmt_count(row_meta.get("valid")))
            src = "网销漏斗「%s」行 AL 列" % row_name
        else:
            note = ("签单 %d / 有效资源 %d" % (signed, total)) if total > 0 else "无数据"
            src = "网销漏斗逐人列聚合（未匹配到网销漏斗同名行）"
        tgt = BIZ_TARGETS[name]
        status, stext = _status_interval(rate, tgt["lower"], tgt["upper"])
        net_biz["metrics"].append({
            "name": "付费客资源签单转化率-%s（近30天）" % name,
            "value": rate,
            "display": _fmt_ratio(rate),
            "type": "ratio",
            "target": tgt,
            "target_display": tgt["display"],
            "status": status,
            "status_text": stext,
            "note": note,
            "source": src,
            "scope": name,
        })
    # 6-11 网销漏斗聚合
    rsrc_total, _ = _sum_funnel(net_emps, "演示用户")
    # 当月见客签单转化率口径（2026-09-13 / 2026-09-21 用户确认）：
    #   = 网销漏斗「接收资源30天内见客后的转化率」
    #   = 接收资源30天内见客后签单量 ÷ 接收资源30天内见客量。
    visit_total, visit_n = _sum_funnel(net_emps, "接收资源30天内见客量")
    sign_total, sign_n = _sum_funnel(net_emps, "接收资源30天内见客后签单量")
    sales_total, sales_n = _sum_funnel(net_emps, "销售额")
    # 员工见客签单转化率口径（2026-09-21 用户口径，源表「目标」sheet 定义原文）：
    #   ② 员工见客签单转化率（考核月）= 考核月见客签单量 ÷ 考核月的演示用户（**不含陪谈**）
    #      → 源表「见客」sheet：记录类型=主面谈 ∧ 岗位∈{网销岗,综法岗}
    #   ⑥ 当月见客签单转化率（考核月）仍走网销漏斗「接收资源30天内见客后的转化率」（上面的 sign/visit）。
    try:
        _vg = visit_agg()
    except Exception:
        _vg = {}
    _v_comp = _vg.get("company") or {"den": 0, "num": 0, "rate": None}
    _v_all = _vg.get("all") or {"den": 0, "num": 0, "rate": None}
    # 演示用户（仅展示）
    net_biz["metrics"].append({
        "name": "演示用户", "value": rsrc_total,
        "display": _fmt_count(rsrc_total), "type": "count",
        "target": {"kind": "display"}, "target_display": "仅展示",
        "status": "display", "status_text": "—",
        "note": "网销漏斗近30天聚合",
    })
    # 演示用户（= 见客 sheet 全部记录，含陪谈）
    net_biz["metrics"].append({
        "name": "演示用户", "value": _v_all["den"],
        "display": _fmt_count(_v_all["den"]), "type": "count",
        "target": {"kind": "display"}, "target_display": "仅展示",
        "status": "display", "status_text": "—",
        "note": "源表「见客」sheet 全部记录（含陪谈）｜%s" % VISIT_WINDOW_NOTE,
    })
    # 员工见客量（= 主面谈，不含陪谈）
    net_biz["metrics"].append({
        "name": "员工见客量", "value": _v_comp["den"],
        "display": _fmt_count(_v_comp["den"]), "type": "count",
        "target": {"kind": "display"}, "target_display": "仅展示",
        "status": "display", "status_text": "—",
        "note": "见客 sheet 主面谈（不含陪谈）· 岗位∈{网销岗,综法岗}｜%s" % VISIT_WINDOW_NOTE,
    })
    # 员工见客签单转化率（不含陪谈）
    visit_rate = _v_comp.get("rate")
    status, stext = _status_pos(visit_rate, 0.30)
    net_biz["metrics"].append({
        "name": "员工见客签单转化率", "value": visit_rate,
        "display": _fmt_ratio(visit_rate), "type": "ratio",
        "target": {"kind": "positive", "target": 0.30},
        "target_display": "30%",
        "status": status, "status_text": stext,
        "note": ("见客签单 %d / 主面谈见客 %d（不含陪谈）｜%s" % (_v_comp["num"], _v_comp["den"], VISIT_WINDOW_NOTE)
                 if _v_comp.get("den") else "源表「见客」sheet 无主面谈记录"),
    })
    # 当月见客量（= 见客 sheet 全部记录）
    net_biz["metrics"].append({
        "name": "当月见客量", "value": _v_all["den"],
        "display": _fmt_count(_v_all["den"]), "type": "count",
        "target": {"kind": "display"}, "target_display": "仅展示",
        "status": "display", "status_text": "—",
        "note": "见客 sheet 全部记录（按面谈结束时间归属）｜%s" % VISIT_WINDOW_NOTE,
    })
    # 当月见客签单转化率（考核月）：2026-09-21 用户指定走「分公司网销漏斗」
    #   c36「接收资源30天内见客后的转化率」⇒ 与「员工见客签单转化率」不再是同口径，正常展示。
    month_rate = (sign_total / visit_total) if visit_total > 0 else None
    status, stext = _status_pos(month_rate, 0.30)
    net_biz["metrics"].append({
        "name": "当月见客签单转化率", "value": month_rate,
        "display": _fmt_ratio(month_rate), "type": "ratio",
        "target": {"kind": "positive", "target": 0.30},
        "target_display": "30%",
        "status": status, "status_text": stext,
        "note": ("见客签单 %d / 见客 %d（分公司网销漏斗 · 近30天块）"
                 % (round(sign_total), round(visit_total))) if visit_total else "无见客",
    })
    # 演示用户
    n_net = len(net_emps)
    per_capita = (sales_total / n_net) if n_net > 0 else None
    status, stext = _status_pos(per_capita, 20000)
    net_biz["metrics"].append({
        "name": "演示用户", "value": per_capita,
        "display": _fmt_amount(per_capita), "type": "amount",
        "target": {"kind": "positive", "target": 20000},
        "target_display": "20,000元",
        "status": status, "status_text": stext,
        "note": "总销售额 %s / 网销规划师 %d 人" % (_fmt_amount(sales_total), n_net) if n_net else "无员工",
    })
    # 演示用户
    mgrs, staffs = _staff_split(net_emps)
    struct_str = "%d+%d" % (len(mgrs), len(staffs))
    struct_ok = (len(mgrs) >= 1 and 5 <= len(staffs) <= 6)
    net_biz["metrics"].append({
        "name": "演示用户", "value": struct_str,
        "display": struct_str, "type": "struct",
        "target": {"kind": "range", "lower": "1+5", "upper": "1+6"},
        "target_display": "1+5~6（1 名管理岗带 5~6 名员工）",
        "status": "pass" if struct_ok else "warn",
        "status_text": "达标" if struct_ok else "需调整",
        "note": "管理岗 %d · 员工 %d" % (len(mgrs), len(staffs)),
    })
    net_biz["notes"] = [
        "「付费客资源签单转化率-×（近30天）」5 张业态卡片与分公司/事业部/员工三级同源：取「四大模式-重点指标表格.xlsx → 网销漏斗」sheet 中 A 列名称匹配行（业态→对应「<CITY>XX事业部」行，工商→<CITY>工商事业部、AI智能体→<CITY>数字科技事业部）的 AL 列「付费客30天签单转化率」＝付费客30天签单量÷有效资源量。",
        "回退口径（仅当网销漏斗未匹配到该业态对应行时启用）：按「付费客资源」sheet 的 一级/三级业态与产品归类归口，Σ签单÷Σ资源。",
        "付费客资源已剔除：①有效资源量=0 ②资源编码为空 ③是否无效='是' ④是否反无效='是'。",
        "业务线口径（回退时使用）：工商=一级业态'工商服务'；跨境=一级业态'跨境业务'；法律=一级业态'法律服务'；AI智能体=三级业态'AI数智化'；交易=产品归类'资质交易'。",
        "「员工见客签单转化率（考核月）」口径（2026-09-21 用户口径，源表「目标」sheet 定义原文）＝ 考核月见客签单量 ÷ 考核月的演示用户（**不含陪谈**）；取源表「见客」sheet：记录类型=主面谈 ∧ 岗位∈{网销岗,综法岗}，签单=「首次收款时间」非空。",
        "「当月见客签单转化率（考核月）」口径＝ 网销漏斗「接收资源30天内见客后的转化率」= Σ接收资源30天内见客后签单量 ÷ Σ接收资源30天内见客量（含陪谈，窗口近30天）。与上一行的「员工见客签单转化率」口径不同，勿互相替代。",
        "「见客」sheet 当前覆盖 09-13~09-20（缺 09-19），即 5 天口径；分母为 0 的部门/员工显示「-」且不判定达标。",
        "正向指标 5 个百分点缓冲；分母为 0 显示 '-' 且不参与达标判定。",
    ]

    # ===== 电邀模块 =====
    # 用户 2026-09-14 指定：按官方「电邀模式（考核月）」重点指标表（共 15 项）逐项建卡；
    #   —— 有数据源的按同口径计算；无数据源的显示「-」（不臆造数值）。
    #   数据源三类：① 电邀漏斗（funnel_source='dy'） ② LTC 报表 ③ 面谈明细（陪同见客）
    dy_emps = [e for e in employees if e.get("部门标准") == "电邀"]
    dy_biz = {"name": "电邀模式（考核月）",
              "sub": "考核月 · 某公司 + <CITY>工商电邀六部 ｜ 数据源：电邀漏斗 · LTC · 面谈明细",
              "metrics": [], "notes": []}

    def _dy_funnel(emps, key):
        s = 0.0
        n = 0
        for e in emps:
            v = (e.get("funnel") or {}).get(key)
            if isinstance(v, (int, float)):
                s += v
                n += 1
        return s, n

    def _dy_nodata(name, target_display, note, scope="—"):
        """无数据源指标：统一显示「-」（用户口径：没有数据显示 - 即可）。"""
        return {"name": name, "value": None, "display": "-", "type": "na",
                "target": {"kind": "display"}, "target_display": target_display,
                "status": "na", "status_text": "无数据", "note": note, "scope": scope}

    # 新人/老人按司龄分
    new_emps = [e for e in dy_emps if (e.get("司龄") is not None and e["司龄"] <= 3)]
    old_emps = [e for e in dy_emps if (e.get("司龄") is not None and e["司龄"] > 3)]
    _, new_hq = _dy_funnel(new_emps, "过去30天高价值客户人均见客量")
    _, old_hq = _dy_funnel(old_emps, "过去30天高价值客户人均见客量")
    hq_new = (new_hq / len(new_emps)) if new_emps else None
    hq_old = (old_hq / len(old_emps)) if old_emps else None
    visit_sum, _ = _dy_funnel(dy_emps, "近30天30S沟通见客数")
    sign_sum, _ = _dy_funnel(dy_emps, "近30天高价值线索见客成单")
    dy_rate = (sign_sum / visit_sum) if visit_sum > 0 else None
    sales_sum, _ = _dy_funnel(dy_emps, "近30天高价值客户业绩")
    n_dy = len(dy_emps)
    per_capita = (sales_sum / n_dy) if n_dy > 0 else None
    mgrs, staffs = _staff_split(dy_emps)
    struct_str = "%d+%d" % (len(mgrs), len(staffs))
    struct_ok = (len(mgrs) >= 1 and 10 <= len(staffs) <= 15)

    # ---------- 陪同类：面谈明细「岗位」×「私客所属人姓名」∈电邀在职名单 归口 ----------
    # （与 8 大重点指标的部门归口同一套口径；签单 = 「首次收款时间」非空）
    _iv_rows, _ = _load_interview()
    _dy_names = set(str(e.get("姓名") or "").strip() for e in dy_emps)
    _dy_iv = {}
    for _r in _iv_rows:
        if str(_r.get("私客所属人姓名") or "").strip() not in _dy_names:
            continue
        _role = str(_r.get("岗位") or "").strip()
        if _role not in ("业务经理", "中台", "管理者"):
            continue
        _s = _dy_iv.setdefault(_role, {"n": 0, "signed": 0})
        _s["n"] += 1
        if _r.get("首次收款时间") not in (None, ""):
            _s["signed"] += 1

    def _iv_pair(role, cn, name_visit, name_rate):
        """输出「陪同见客量」+「陪同见客签单转化率」两张卡；无记录 → 显示「-」。"""
        st = _dy_iv.get(role) or {"n": 0, "signed": 0}
        if st["n"] > 0:
            dy_biz["metrics"].append({
                "name": name_visit, "value": st["n"], "display": _fmt_count(st["n"]),
                "type": "count", "target": {"kind": "display"}, "target_display": "仅展示",
                "status": "display", "status_text": "", "show": True, "scope": "面谈明细",
                "note": "口径：考核月%s陪同演示用户。本页归口到电邀在职名单的陪谈记录共 %d 条。" % (cn, st["n"]),
            })
            _rate = st["signed"] / st["n"]
            _st, _stx = _status_pos(_rate, 0.30)
            dy_biz["metrics"].append({
                "name": name_rate, "value": _rate, "display": _fmt_ratio(_rate),
                "type": "ratio", "target": {"kind": "positive", "target": 0.30},
                "target_display": "30%", "status": _st, "status_text": _stx, "scope": "面谈明细",
                "note": "口径：考核月%s陪同见客签单量 ÷ %s陪同演示用户。本页 签单 %d / 陪谈 %d（签单=「首次收款时间」非空）。" % (cn, cn, st["signed"], st["n"]),
            })
        else:
            _src = "面谈明细中「岗位=%s」且「私客所属人姓名」∈电邀在职名单" % role
            dy_biz["metrics"].append(_dy_nodata(
                name_visit, "仅展示",
                "口径：考核月%s陪同演示用户。%s 的记录数为 0 → 无数据。" % (cn, _src), "面谈明细"))
            dy_biz["metrics"].append(_dy_nodata(
                name_rate, "30%",
                "口径：考核月%s陪同见客签单量 ÷ %s陪同演示用户。同左，当前无记录（分母为 0 不判定达标）。" % (cn, cn), "面谈明细"))

    # ===== 1) 人均月度高质量见客量（0-3月新人） =====
    status, stext = _status_pos(hq_new, 3, band=0.20) if hq_new is not None else ("na", "无新人")
    dy_biz["metrics"].append({
        "name": "人均月度高质量见客量（0-3月）",
        "value": hq_new, "display": _fmt_count(hq_new), "type": "count",
        "target": {"kind": "positive", "target": 3},
        "target_display": "0-3月月目标 3 个（宽松 3-5 个）",
        "status": status, "status_text": stext, "scope": "电邀漏斗",
        "note": "口径：0-3个月新人考核月高质量演示用户 ÷ 人员总量。本页：高质量见客合计 %.0f ÷ 新人 %d 人。" % (new_hq, len(new_emps)) if new_emps else "无新人（司龄≤3月）",
    })
    # ===== 2) 人均月度高质量见客量（3个月以上） =====
    status, stext = _status_pos(hq_old, 6, band=0.20) if hq_old is not None else ("na", "无老人")
    dy_biz["metrics"].append({
        "name": "人均月度高质量见客量（3个月以上）",
        "value": hq_old, "display": _fmt_count(hq_old), "type": "count",
        "target": {"kind": "positive", "target": 6},
        "target_display": "3个月以上月目标 6 个（宽松 5-9 个）",
        "status": status, "status_text": stext, "scope": "电邀漏斗",
        "note": "口径：3个月以上人员考核月高质量演示用户 ÷ 人员总量。本页：高质量见客合计 %.0f ÷ 老人 %d 人。" % (old_hq, len(old_emps)) if old_emps else "无老人（司龄>3月）",
    })
    # ===== 3) 员工见客量 =====
    dy_biz["metrics"].append({
        "name": "员工见客量", "value": visit_sum,
        "display": _fmt_count(visit_sum), "type": "count",
        "target": {"kind": "display"}, "target_display": "仅展示",
        "status": "display", "status_text": "", "show": True, "scope": "电邀漏斗",
        "note": "口径：考核月员工演示用户。本页：电邀漏斗「近30天30S沟通见客数」全员求和。",
    })
    # ===== 4) 员工见客签单转化率（考核月） =====
    status, stext = _status_pos(dy_rate, 0.30)
    dy_biz["metrics"].append({
        "name": "员工见客签单转化率（考核月）", "value": dy_rate,
        "display": _fmt_ratio(dy_rate), "type": "ratio",
        "target": {"kind": "positive", "target": 0.30},
        "target_display": "30%（不含陪谈）",
        "status": status, "status_text": stext, "scope": "电邀漏斗",
        "note": "口径：考核月见客签单量 ÷ 考核月演示用户（不含陪谈）。本页：见客签单 %s / 见客 %s。" % (round(sign_sum), round(visit_sum)) if visit_sum else "无见客",
    })
    # ===== 5) KA岗陪同见客量 =====
    dy_biz["metrics"].append(_dy_nodata(
        "KA岗陪同见客量", "仅展示",
        "口径：考核月KA岗陪同演示用户。面谈明细的「岗位」列现仅有 综法岗/网销岗/中台/业务经理/会计岗/管理者/电邀岗，无「KA岗」分类 → 无数据源。", "无数据源"))
    # ===== 6) KA岗见客签单转化率（考核月） =====
    dy_biz["metrics"].append(_dy_nodata(
        "KA岗见客签单转化率（考核月）", "30%",
        "口径：KA岗考核月见客签单量 ÷ 考核月演示用户。同上，面谈明细无「KA岗」分类 → 无数据源。", "无数据源"))
    # ===== 7/8) 中台老师陪同 =====
    _iv_pair("中台", "中台老师", "中台老师陪同见客量", "中台老师陪同见客签单转化率")
    # ===== 9/10) 事业部总陪同 =====
    _iv_pair("管理者", "事业部总", "事业部总陪同见客量", "事业部总陪同见客签单转化率")
    # ===== 11) 当月见客量 =====
    dy_biz["metrics"].append(_dy_nodata(
        "当月见客量", "仅展示",
        "口径：考核月当月演示用户。系统现有口径仅「近30天 / 考核月块」，无独立「当月」见客字段 → 暂显示「-」；如需以某字段映射请指定。", "无数据源"))
    # ===== 12) 当月见客签单转化率（考核月） =====
    dy_biz["metrics"].append(_dy_nodata(
        "当月见客签单转化率（考核月）", "30%",
        "口径：考核月见客签单量 ÷ 考核月演示用户。同上，无独立「当月」字段 → 暂显示「-」。", "无数据源"))
    # ===== 13) KA转化岗与电邀岗人员比例配置 =====
    dy_biz["metrics"].append(_dy_nodata(
        "KA转化岗与电邀岗人员比例配置", "KA : 电邀岗 = 1 : 5",
        "口径：KA 转化岗 : 电邀岗 = 1 : 5 比例配置。当前在职电邀规划师 %d 人，但系统无「KA转化岗」人员数据 → 无法计算实际比例。" % n_dy, "无数据源"))
    # ===== 14) 演示用户 =====
    dy_biz["metrics"].append({
        "name": "演示用户", "value": struct_str,
        "display": struct_str, "type": "struct",
        "target": {"kind": "range", "lower": "1+10", "upper": "1+15"},
        "target_display": "编制 1+10~15（1 名管理岗带 10~15 名员工）",
        "status": "pass" if struct_ok else "warn",
        "status_text": "达标" if struct_ok else "需调整", "scope": "在职名单",
        "note": "管理岗 %d · 员工 %d（结构性指标，不计入达成率图）。" % (len(mgrs), len(staffs)),
    })
    # ===== 15) 演示用户（考核月） =====
    status, stext = _status_pos(per_capita, 10000)
    dy_biz["metrics"].append({
        "name": "演示用户（考核月）", "value": per_capita,
        "display": _fmt_amount(per_capita), "type": "amount",
        "target": {"kind": "positive", "target": 10000},
        "target_display": "10,000元",
        "status": status, "status_text": stext, "scope": "电邀漏斗",
        "note": "口径：在职电邀规划师业绩 ÷ 在职电邀规划师总人数。本页：总业绩 %s / %d 人。" % (_fmt_amount(sales_sum), n_dy) if n_dy else "无员工",
    })
    dy_biz["notes"] = [
        "数据范围：某公司 + <CITY>工商电邀六部及其员工（funnel_source='dy'）。",
        "指标来源：官方「电邀模式（考核月）」重点指标表（共 15 项），逐项建卡；**无数据源的指标显示「-」，不臆造数值**。",
        "数据源三类：① 电邀漏斗（见客/成单/业绩等，个人明细层字段）② LTC 报表 ③ 面谈明细（陪同见客，按「岗位」×「私客所属人姓名」归口到电邀在职名单）。",
        "新人/老人按司龄分：新人 ≤3 月、老人 >3 月；高质量见客量 = 电邀漏斗「过去30天高价值客户人均见客量」。",
        "陪同类：中台老师＝「岗位=中台」、事业部总＝「岗位=管理者」，签单＝「首次收款时间」非空；KA 岗在当前面谈明细中无对应「岗位」分类 → 无数据。",
        "「当月见客量」「当月见客签单转化率」：系统现有口径为近30天/考核月块，无独立「当月」字段 → 显示「-」。",
    ]

    # ===== 会计模块 =====（sheet 名动态含「会计」，按关键字取：S0 / 规划师 / 主管）
    s0_rows = [r for r in _kj_sheet_rows(kuaiji, "S0") if _is_hangzhou(r)]
    pl_rows = [r for r in _kj_sheet_rows(kuaiji, "规划师") if _is_hangzhou(r)]
    zg_rows = [r for r in _kj_sheet_rows(kuaiji, "主管") if _is_hangzhou(r)]
    kj_biz = {"name": "演示用户", "sub": "考核月 · 数据周期 2026/8/15—2026/9/8（口径：自然月续费率 / 其余 15-14 周期）",
              "metrics": [], "notes": []}
    # 1) 会计到期续费率（家数）
    renew_vals = [_to_num(_kv(r, "到期续费率", "家数", "9月")) for r in pl_rows]
    renew_vals = [v for v in renew_vals if v is not None]
    renew_avg = (sum(renew_vals) / len(renew_vals)) if renew_vals else None
    status, stext = _status_pos(renew_avg, 0.85)
    kj_biz["metrics"].append({
        "name": "会计到期续费率", "value": renew_avg,
        "display": _fmt_ratio(renew_avg), "type": "ratio",
        "target": {"kind": "positive", "target": 0.85},
        "target_display": "85%",
        "status": status, "status_text": stext,
        "note": "规划师「到期续费率(家数)(9月)」平均 n=%d" % len(renew_vals) if renew_vals else "无数据",
    })
    # 2) 客服会计演示用户（含续费）
    kf_vals = [_to_num(_kv(r, "含续费", "9月")) for r in (s0_rows + pl_rows)
               if str(r.get("岗位") or "").strip() == "客服会计"]
    kf_avg, kf_n = _avg_num(kf_vals)
    status, stext = _status_pos(kf_avg, 15000)
    kj_biz["metrics"].append({
        "name": "客服会计演示用户（含续费）", "value": kf_avg,
        "display": _fmt_amount(kf_avg), "type": "amount",
        "target": {"kind": "positive", "target": 15000},
        "target_display": "15,000元",
        "status": status, "status_text": stext,
        "note": "客服会计 %d 人 · 来自 S0 + 规划师" % kf_n if kf_n else "无客服会计",
    })
    # 3) 顾问会计演示用户（不含续费）
    gw_vals = [_to_num(_kv(r, "不含续费", "9月")) for r in pl_rows
               if str(r.get("岗位") or "").strip() == "会计顾问"]
    gw_avg, gw_n = _avg_num(gw_vals)
    status, stext = _status_pos(gw_avg, 30000)
    kj_biz["metrics"].append({
        "name": "顾问会计演示用户（不含续费）", "value": gw_avg,
        "display": _fmt_amount(gw_avg), "type": "amount",
        "target": {"kind": "positive", "target": 30000},
        "target_display": "30,000元",
        "status": status, "status_text": stext,
        "note": "顾问会计 %d 人 · 来自规划师" % gw_n if gw_n else "无顾问会计",
    })
    # 4) 会计主管演示用户（不含续费）
    zg_vals = [_to_num(_kv(r, "不含续费", "9月")) for r in zg_rows
               if str(r.get("岗位") or "").strip() == "会计主管"]
    zg_avg, zg_n = _avg_num(zg_vals)
    status, stext = _status_pos(zg_avg, 40000)
    kj_biz["metrics"].append({
        "name": "会计主管演示用户（不含续费）", "value": zg_avg,
        "display": _fmt_amount(zg_avg), "type": "amount",
        "target": {"kind": "positive", "target": 40000},
        "target_display": "40,000元",
        "status": status, "status_text": stext,
        "note": "会计主管 %d 人 · 团队月人均实算业绩(不含续费)" % zg_n if zg_n else "无会计主管",
    })
    # 5) 非正常流失率（反向指标）—— 2026-09-28 口径修正（用户确认）：
    #    非正常流失率 = 非正常流失家数 ÷ 最新保有量（源表「会计监控表」，合计口径 Σ流失 ÷ Σ保有）。
    #    旧口径「S0/规划师 非正常流失户数 ÷ 维护家数」与源表官方公式不一致，已弃用；
    #    员工明细里的「非正常流失」列（个人户数过程量）保留不动。
    mon_rows = _load_kj_monitor()
    lost_sum = sum([r["非正常流失家数"] or 0 for r in mon_rows])
    held_sum = sum([r["最新保有量"] or 0 for r in mon_rows])
    lost_rate = (lost_sum / held_sum) if held_sum > 0 else None
    status, stext = _status_rev(lost_rate, 0.01)
    kj_biz["metrics"].append({
        "name": "非正常流失率", "value": lost_rate,
        "display": _fmt_ratio(lost_rate), "type": "ratio",
        "target": {"kind": "reverse", "target": 0.01},
        "target_display": "≤1%（反向）",
        "status": status, "status_text": stext,
        "note": "非正常流失家数 %d ÷ 最新保有量 %d（源表「会计监控表」合计口径 · %d 人有保有量）" % (
            round(lost_sum), round(held_sum),
            len([r for r in mon_rows if r["最新保有量"] is not None]),
        ) if held_sum else "监控表无保有量数据",
    })
    # 6~8) 新增（2026-09-16 需求 N）：票据抽检覆盖率 / 票据质检出错率 / 续费明细应续续费率
    #    三张新增 sheet 暂无官方目标值 → status='display' + show=True（出卡但不进「达成率」图，
    #    避免臆造目标线）；数值与下方「🔎 会计票据抽检与个人出错率」「🔄 会计续费明细」两板块同源。
    _kje = _load_kj_extra() or {}
    _audit = _kje.get("audit") or {}
    _atot = _audit.get("total") or {}
    _err_tot = (_kje.get("err") or {}).get("total") or {}
    _rstats = ((_kje.get("renew") or {}).get("stats") or {})
    _due = _rstats.get("due") or {}
    if _atot:
        kj_biz["metrics"].append({
            "name": "票据抽检覆盖率", "value": _atot.get("覆盖率"),
            "display": _fmt_ratio(_atot.get("覆盖率")), "type": "ratio",
            "target": {"kind": "display"}, "target_display": "仅展示（无官方目标）",
            "status": "display", "status_text": "仅展示", "show": True,
            "note": "抽检总量 %s（人工+AI）÷ 总家数 %s｜周期 %s" % (
                _fmt_count(_atot.get("抽检总量")), _fmt_count(_atot.get("总家数")),
                _audit.get("period") or "—"),
        })
        kj_biz["metrics"].append({
            "name": "票据质检出错率", "value": _atot.get("出错率"),
            "display": _fmt_ratio(_atot.get("出错率")), "type": "ratio",
            "target": {"kind": "display"}, "target_display": "仅展示（无官方目标）",
            "status": "display", "status_text": "仅展示", "show": True,
            "note": "出错家数 %s ÷ 抽检总量 %s（人工+AI）｜个人出错率合计 %s" % (
                _fmt_count(_atot.get("出错家数")), _fmt_count(_atot.get("抽检总量")),
                _fmt_ratio(_err_tot.get("出错率"))),
        })
    if _due:
        kj_biz["metrics"].append({
            "name": "续费明细应续续费率", "value": _due.get("应续续费率"),
            "display": _fmt_ratio(_due.get("应续续费率")), "type": "ratio",
            "target": {"kind": "display"}, "target_display": "仅展示（无官方目标）",
            "status": "display", "status_text": "仅展示", "show": True,
            "note": "应续已续 %d ÷ 应续 %d｜全量续费率 %s（当月续费明细表 %s）" % (
                _due.get("应续已续") or 0, _due.get("是") or 0,
                _fmt_ratio((_rstats.get("renew") or {}).get("率")),
                (_kje.get("renew") or {}).get("month") or "—"),
        })
    # 9) 流程会计个人出错率（源表「流程会计个人出错率」合计口径）
    #    与「票据质检出错率」是两张表的两个口径：前者按个人抽检（273→8.4%），
    #    后者按事业部票据抽检（275→8.4%），刻意并列呈现、不做合并。
    # 10) 流失公司数（源表「会计续费明细表」，流失＝公司状态含「流失」）
    #    ⚠️ 与第 5 项「非正常流失率」口径不同：那是监控表「非正常流失家数÷最新保有量」（合计口径），
    #    这里是续费明细表的「流失家数」，故两张卡并存、各自标注口径。
    if _err_tot:
        kj_biz["metrics"].append({
            "name": "流程会计个人出错率", "value": _err_tot.get("出错率"),
            "display": _fmt_ratio(_err_tot.get("出错率")), "type": "ratio",
            "target": {"kind": "display"}, "target_display": "仅展示（无官方目标）",
            "status": "display", "status_text": "仅展示", "show": True,
            "note": "出错量 %s ÷ 检查量 %s｜抽检 %s 人（其中新人 %s）｜源表「流程会计个人出错率」" % (
                _fmt_count(_err_tot.get("出错量")), _fmt_count(_err_tot.get("检查量")),
                _fmt_count(_err_tot.get("人数")), _fmt_count(_err_tot.get("新人"))),
        })
    _lost = _rstats.get("lost") or {}
    if _lost:
        kj_biz["metrics"].append({
            "name": "流失公司数", "value": _lost.get("家数"),
            "display": "%s 家" % _fmt_count(_lost.get("家数")), "type": "count",
            "target": {"kind": "display"}, "target_display": "仅展示（无官方目标）",
            "status": "display", "status_text": "仅展示", "show": True,
            "note": "流失 %s 家 ÷ 合计 %s 家（占比 %s）｜源表「会计续费明细表」·流失＝公司状态含「流失」" % (
                _fmt_count(_lost.get("家数")), _fmt_count(_rstats.get("total")),
                _fmt_ratio(_lost.get("占比"))),
        })
    kj_biz["notes"] = [
        "数据来源：表格中带集团公布职级表，按端口=某公司筛选。",
        "续费率采用「家数」口径，与「金额」主管口径分离。",
        "保有量口径（2026-09-28 确认）：非正常流失率＝「会计监控表」非正常流失家数合计 ÷ 最新保有量合计；"
        "员工明细中的「非正常流失」列是个人户数过程量，与该率分子同源不同层。",
        "正向指标 5 个百分点缓冲；反向指标同等缓冲；分母为 0 显示 '-' 且不参与判定。",
        "新增 3 项（2026-09-16）来自同工作簿新增 sheet：①票据抽检覆盖率＝抽检总量(人工+AI)÷总家数；"
        "②票据质检出错率＝出错家数÷抽检总量（源表「<CITY>会计抽检」）；③续费明细应续续费率＝应续已续÷应续家数"
        "（源表「会计续费明细表」，全量续费率＝续费=是÷全部家数）。三者暂无官方目标值 → 仅展示、不参与达标判定。",
        "补充 2 项（2026-09-16 需求 N-2）：④流程会计个人出错率＝出错量÷检查量（源表「流程会计个人出错率」"
        "个人抽检口径）；⑤流失公司数＝公司状态含「流失」的家数（源表「会计续费明细表」，合计 109 家）。"
        "均无官方目标值 → 仅展示、不参与达标判定。",
        "⚠️ 口径并列说明（三组易混，均刻意各自标注、不做合并）："
        "①「票据质检出错率」按事业部票据抽检 275 计，与「流程会计个人出错率」按个人检查量 273 计，两表两口径；"
        "②「非正常流失率」＝监控表「非正常流失家数÷最新保有量」（合计口径），与「流失公司数」＝续费明细表家数，两表两口径；"
        "③「会计到期续费率」（规划师 9 月家数口径）与「续费明细应续续费率」（续费明细表应续口径）亦不同源。",
        "⚠️ 「<CITY>会计抽检」抽检总量 275 与「流程会计个人出错率」检查量合计 273 略有差异：前者按事业部票据口径、"
        "后者按个人抽检口径，两张表独立取数，看板按各自源表原值呈现、不做强行对齐。",
    ]
    # 会计高端账签单占比（独立口径，不进入"重点指标达成率"图，单列展示）
    high_end = _load_high_end_account()
    if high_end:
        kj_biz["high_end"] = high_end

    return {"net": net_biz, "dy": dy_biz, "kuaiji": kj_biz}


# 会计明细展示字段（员工详情「漏斗指标」区）
KJ_FUNNEL_METRICS = [
    ("到期续费率(家数)(9月)", True, "当月到期实收家数 ÷ 当月到期家数"),
    ("到期续费率(金额)(9月)", True, "当月到期实收金额 ÷ 当月到期金额"),
    ("月度实算业绩(含续费)(9月）", False, "含续费月度实算业绩（元）"),
    ("月度实算业绩(不含续费)(9月）", False, "不含续费月度实算业绩（元）"),
    ("账务工作价值量积分新版(9月)", False, "账务工作价值量积分"),
    ("非正常流失户数(9月)", False, "当月非正常流失户数"),
]
KJ_POST_ORDER = ["客服会计", "会计顾问", "流程会计", "会计主管"]


def _kj_pick(row, *subs):
    """会计行按子串取数值：优先含续费口径，回退不含续费 / 团队人均。"""
    return _to_num(_kv(row, *subs))


def _kuaiji_employees(kuaiji, dept_std_mode="post"):
    """从（会计）sheet 构建【<CITY>端口】员工列表（会计口径，不走 LTC）。

    dept_std_mode:
      "post" → 部门标准=岗位（会计分析页按岗位聚合：客服会计/会计顾问/流程会计）
      "会计" → 部门标准统一="会计"（主页面并入单一"会计"部门）
    会计人员不参与三力评价，三力达标记"无三力"；业绩取月度实算业绩(含续费)→(不含续费)→团队人均。
    """
    emps = []
    seen = set()
    _perf = _kj_signer_perf()
    _kj_perf_total = _perf["total"]
    _kj_perf_today = _perf["today"]
    for sheet, rows in (kuaiji or {}).items():
        for r in rows:
            if not _is_hangzhou(r):
                continue
            name = str(r.get("姓名") or "").strip()
            if not name:
                continue
            code = str(r.get("工号") or "").strip()
            key = code or name
            if key in seen:
                continue
            seen.add(key)
            post = str(r.get("岗位") or "").strip() or "会计"
            deptname = str(r.get("部门") or "").strip() or post
            perf = _kj_pick(r, "月度实算业绩", "含续费", "9月")
            if perf is None:
                perf = _kj_pick(r, "月度实算业绩", "不含续费", "9月")
            if perf is None:
                perf = _kj_pick(r, "团队月人均实算业绩", "不含续费", "9月")
            funnel = {
                "到期续费率(家数)(9月)": _kj_pick(r, "到期续费率", "家数", "9月"),
                "到期续费率(金额)(9月)": _kj_pick(r, "到期续费率", "金额", "9月"),
                "月度实算业绩(含续费)(9月）": _kj_pick(r, "月度实算业绩", "含续费", "9月"),
                "月度实算业绩(不含续费)(9月）": _kj_pick(r, "月度实算业绩", "不含续费", "9月"),
                "账务工作价值量积分新版(9月)": _kj_pick(r, "账务工作价值量积分", "9月"),
                "非正常流失户数(9月)": _kj_pick(r, "非正常流失户数", "9月"),
            }
            rec = {
                "姓名": name,
                "工号": code,
                "部门": deptname,
                "部门标准": post if dept_std_mode == "post" else "会计",
                "岗位": post,
                "事业部": r.get("三级部门") or r.get("端口") or "",
                "端口": r.get("端口") or "",
                "职级": r.get("9月职级") or "",
                "司龄": _to_num(r.get("司龄")),
                "三力达标": "无三力",
                "月累计实算利润业绩": (_kj_perf_total.get(name) or 0.0),
                "月累计核算利润业绩": (_kj_perf_total.get(name) or 0.0),
                "当日新增业绩": (_kj_perf_today.get(name) or 0.0),
                "funnel_source": "kuaiji",
                "funnel": funnel,
                "ltc": {},
                "advice": [],
            }
            # 用户 2026-09-16 需求 N：会计人员明细要「数据齐全」——
            #   补两处新增 sheet 的人员级数据：
            #   · kj_err   ← 「流程会计个人出错率」（检查量/出错量/出错率/新人）
            #   · kj_renew ← 「会计续费明细表」按「会计顾问」归口的续费家数
            _err = _kj_err_lookup(name, code)
            rec["kj_err"] = {
                "检查量": _err.get("检查量") if _err else None,
                "出错量": _err.get("出错量") if _err else None,
                "出错率": _err.get("出错率") if _err else None,
                "出错率无分母": bool(_err.get("出错率无分母")) if _err else False,
                "新人": bool(_err.get("新人")) if _err else False,
                "来源岗位": (_err or {}).get("岗位") or "",
            }
            _rn = _kj_renew_lookup(name, code)
            rec["kj_renew"] = {
                "记录数": _rn.get("家数") if _rn else None,
                "应续家数": _rn.get("应续家数") if _rn else None,
                "已续家数": _rn.get("已续家数") if _rn else None,
                "应续已续": _rn.get("应续已续") if _rn else None,
                "续费率": _rn.get("续费率") if _rn else None,
                "应续续费率": _rn.get("应续续费率") if _rn else None,
                "流失家数": _rn.get("流失家数") if _rn else None,
                "应收": _rn.get("应收") if _rn else None,
                "实收": _rn.get("实收") if _rn else None,
            }
            emps.append(rec)
    return emps


def _build_kuaiji_detail(kuaiji, kj_key, data_time=""):
    """会计明细：直接取自 sheet 名带「（会计）」的数据（仅<CITY>端口），
    产出与人员数据分析页同构的结构（employees / bus / summary / meta / key_metrics）。

    说明：会计人员是「当前在册明细」，不参与三力评价，故 三力达标 记"无三力"、
    funnel 用会计自身指标（续费率 / 实算业绩 / 积分 / 流失户数）。"""
    emps = _kuaiji_employees(kuaiji, dept_std_mode="post")

    def _post_rank(p):
        return KJ_POST_ORDER.index(p) if p in KJ_POST_ORDER else 99
    emps.sort(key=lambda e: (_post_rank(e["部门标准"]), -(e["月累计实算利润业绩"] or 0)))

    # 岗位聚合
    groups = OrderedDict()
    for e in emps:
        groups.setdefault(e["部门标准"], []).append(e)
    bus = []
    for post, g in groups.items():
        perf_sum = sum([x["月累计实算利润业绩"] or 0 for x in g])
        renew_vals = [x["funnel"]["到期续费率(家数)(9月)"] for x in g
                      if x["funnel"]["到期续费率(家数)(9月)"] is not None]
        renew_avg = (sum(renew_vals) / len(renew_vals)) if renew_vals else None
        lost_sum = sum([x["funnel"]["非正常流失户数(9月)"] or 0 for x in g])
        bus.append({
            "部门": post,
            "人数": len(g),
            "月实算业绩": perf_sum,
            "funnel_source": "kuaiji",
            "三力": {"英雄人力": 0, "标准人力": 0, "破零人力": 0, "未破零": 0, "无三力": len(g)},
            "funnel_avg": {
                "到期续费率(家数)(9月)": renew_avg,
                "非正常流失户数(9月)": lost_sum,
            },
        })
    bus.sort(key=lambda b: _post_rank(b["部门"]))

    perf_sum = sum([e["月累计实算利润业绩"] or 0 for e in emps])
    # 源表全空检测：某指标在所有会计员工中均无数值 → 前端呈现「源表未提供」
    _kj_nodata = set()
    for k, _r, _d in KJ_FUNNEL_METRICS:
        if not any(isinstance((e.get("funnel") or {}).get(k), (int, float)) for e in emps):
            _kj_nodata.add(k)
    return {
        "meta": {
            "data_time": data_time,
            "sheet_title": "会计数据分析 · 数据来自带（会计）的 sheet",
            "dept_order": [b["部门"] for b in bus],
            "key_module": "kuaiji",
            "kuaiji_detail": True,
        },
        "total": len(emps),
        "departments": [b["部门"] for b in bus],
        "summary": {
            "三力": {"英雄人力": 0, "标准人力": 0, "破零人力": 0, "未破零": 0, "无三力": len(emps), "其他": 0},
            "月实算业绩": perf_sum,
            "月核算业绩": perf_sum,
            "当日新增业绩": sum([(e.get("当日新增业绩") or 0) for e in emps]),
        },
        "bus": bus,
        "employees": emps,
        "funnel_metrics": [{"k": k, "ratio": r, "desc": d, "nodata": (k in _kj_nodata)} for k, r, d in KJ_FUNNEL_METRICS],
        "dy_funnel_metrics": [],
        "ltc_metrics": [],
        "key_metrics": {"kuaiji": kj_key},
        # 用户 2026-09-16 需求 N：三张新增会计 sheet（抽检 / 个人出错率 / 续费明细）
        "kj_extra": _load_kj_extra(),
    }


# ---------- 各部门新增图表聚合（演示用户 / 见客量 / 见客签单率 / 演示用户[剔除无效]） ----------
# 业务线名称 → 部门短名（AI智能体 在部门表里叫 数字科技）
BIZ_TO_DEPT = {name: (name if name != "AI智能体" else "数字科技") for name, _ in BIZ_LINE_RULES}


def build_dept_charts(employees, bus_list, pay_rows, pay_max_date):
    """为人员数据分析新增四类部门维度图准备聚合数据。

    维度均为「部门」；各字段统计口径见注释。
    """
    depts = [b["部门"] for b in bus_list]
    head = {b["部门"]: (b.get("人数") or 0) for b in bus_list}
    by_dept = {}
    for e in employees:
        by_dept.setdefault(e.get("部门标准"), []).append(e)

    per_capita = {}   # 演示用户：Σ漏斗「销售额」÷ 部门人数（元/人）；电邀/会计无销售额→None
    visit = {}        # 见客量：网销=Σ「接收资源30天内见客量」，电邀=近30天30S沟通见客数，会计→None
    signrate = {}     # 见客签单率：网销=Σ「接收资源30天内见客后签单量」÷Σ「接收资源30天内见客量」（=接收资源30天内见客后的转化率）；会计→None
    for d in depts:
        emps = by_dept.get(d, [])
        src = emps[0].get("funnel_source") if emps else None
        if src == "net":
            # 2026-09-13 用户口径：「见客签单转化率」= 网销漏斗「接收资源30天内见客后的转化率」
            #   = 接收资源30天内见客后签单量 ÷ 接收资源30天内见客量。④图见客签单率与之一致。
            sales = sum((e.get("funnel") or {}).get("销售额") or 0 for e in emps)
            v = sum((e.get("funnel") or {}).get("接收资源30天内见客量") or 0 for e in emps)
            sv = sum((e.get("funnel") or {}).get("接收资源30天内见客后签单量") or 0 for e in emps)
            h = head.get(d) or 0
            per_capita[d] = round(sales / h, 2) if h else None
            visit[d] = round(v, 1) if v else 0
            signrate[d] = round(sv / v, 4) if v else None
        elif src == "dy":
            v = sum((e.get("funnel") or {}).get("近30天30S沟通见客数") or 0 for e in emps)
            sv = sum((e.get("funnel") or {}).get("近30天高价值线索见客成单") or 0 for e in emps)
            per_capita[d] = None
            visit[d] = round(v, 1) if v else 0
            signrate[d] = round(sv / v, 4) if v else None
        else:
            per_capita[d] = None
            visit[d] = None
            signrate[d] = None

    # 演示用户（剔除无效资源后）+ 付费客资源分析指标：付费客资源明细，近30天窗口，按「员工所属部门」(AX列) 归到标准部门。
    # 用户明确要求：付费客资源数量表必须使用付费客资源本身的数据，并以 AX 列看资源当前所属部门。
    from datetime import timedelta
    wstart = (pay_max_date - timedelta(days=30)) if pay_max_date else None
    res = {d: 0 for d in depts}
    _DEPT_KEYWORDS = ("电邀", "工商", "会计", "法律", "跨境", "数字科技", "交易")

    def _blank_pay():
        return {
            "resources": 0,   # 有效资源数（剔除无效后分母基准）
            "visits": 0,      # 见客数（是否见客=是）
            "signs": 0,       # 成单数（是否成单=是）
            "wx": 0,          # 加微数（是否加微=是）
            "timely": 0,      # 及时转化数（见客=是 或 成单=是）
            "invites": 0,     # 邀约数（是否邀约=是）
            "hq": 0,          # 高质量资源数（高质量资源=1）
            "referrals": 0,   # 转介绍数（是否转介绍=1）
            "_dials": 0.0,    # Σ拨打次数
            "_connects": 0.0,  # Σ接通次数
            "_s30": 0.0,      # Σ30秒以上接通次数
            "_duration": 0.0,  # Σ通时(min)
            "sign_perf": 0.0,  # Σ成单业绩
            "profit_perf": 0.0,  # Σ利润业绩
        }

    pay_acc = {d: _blank_pay() for d in depts}
    res_unmapped = 0
    for r in pay_rows:
        if _pay_excluded(r):
            continue
        if wstart and r["date"] < wstart:
            continue
        ax = str(r.get("员工所属部门") or "").strip()
        short = None
        for kw in _DEPT_KEYWORDS:
            if kw in ax:
                short = kw
                break
        if short is None or short not in pay_acc:
            res_unmapped += 1
            continue
        res[short] += 1
        a = pay_acc[short]
        a["resources"] += 1
        if _is_yes(r.get("见客")):
            a["visits"] += 1
        if str(r.get("signed") or "").strip() == "是":
            a["signs"] += 1
        if _is_yes(r.get("加微")):
            a["wx"] += 1
        if _is_yes(r.get("见客")) or (str(r.get("signed") or "").strip() == "是"):
            a["timely"] += 1
        if _is_yes(r.get("邀约")):
            a["invites"] += 1
        if _is_yes(r.get("高质量")):
            a["hq"] += 1
        if _is_yes(r.get("转介绍")):
            a["referrals"] += 1
        a["_dials"] += _to_num(r.get("拨打次数")) or 0
        a["_connects"] += _to_num(r.get("接通次数")) or 0
        a["_s30"] += _to_num(r.get("s30接通")) or 0
        a["_duration"] += _to_num(r.get("通时")) or 0
        a["sign_perf"] += _to_num(r.get("成单业绩")) or 0
        a["profit_perf"] += _to_num(r.get("利润业绩")) or 0

    def _rate(num, den):
        return round(num / den, 4) if den else None

    pay_metrics = {}
    for d in depts:
        a = pay_acc[d]
        n = a["resources"]
        pay_metrics[d] = {
            "resources": n,
            "visits": a["visits"],
            "visit_rate": _rate(a["visits"], n),
            "signs": a["signs"],
            "sign_rate": _rate(a["signs"], n),
            "visit_sign_rate": _rate(a["signs"], a["visits"]),          # 资源级：成单数÷见客数（保留，可能 >100%）
            # 见客签单转化率（2026-09-13 用户确认统一口径）= 网销漏斗「接收资源30天内见客后的转化率」(net) / 电邀漏斗口径(dy)
            "visit_sign_rate_funnel": signrate.get(d),
            "wx": a["wx"],
            "wx_rate": _rate(a["wx"], n),
            "timely": a["timely"],
            "timely_rate": _rate(a["timely"], n),
            "invites": a["invites"],
            "invite_rate": _rate(a["invites"], n),
            "hq": a["hq"],
            "hq_rate": _rate(a["hq"], n),
            "referrals": a["referrals"],
            "dials": round(a["_dials"], 1),
            "avg_dials": round(a["_dials"] / n, 2) if n else None,
            "connects": round(a["_connects"], 1),
            "connect_rate": _rate(a["_connects"], a["_dials"]),
            "s30": round(a["_s30"], 1),
            "s30_rate": _rate(a["_s30"], a["_dials"]),
            "duration": round(a["_duration"], 1),
            "avg_duration": round(a["_duration"] / n, 2) if n else None,
            "sign_perf": round(a["sign_perf"], 2),
            "per_sign_perf": round(a["sign_perf"] / a["signs"], 2) if a["signs"] else None,
            "per_res_perf": round(a["sign_perf"] / n, 2) if n else None,
        }

    # 2026-09-21：源表移除「付费客资源」sheet ⇒ pay_rows 为空时，部门级资源类字段
    # 改由「网销漏斗」逐人列派生（用户口径：资源以漏斗数据中的有效资源量即可）。
    # 漏斗无对应列的字段（拨打/接通/通时/利润业绩/邀约/高质量/转介绍）保持 0/None。
    if not pay_rows:
        try:
            _fr_per_d, _fr_dept_d, _fr_comp_d = _funnel_res_agg(employees)
        except Exception:
            _fr_dept_d = {}
        for d in depts:
            s = _fr_dept_d.get(d)
            if not s:
                continue
            m = pay_metrics.setdefault(d, {})
            m["resources"] = s.get("resources")
            m["visits"] = s.get("visits")
            m["visit_rate"] = s.get("visit_rate")
            m["signs"] = s.get("signs")
            m["sign_rate"] = s.get("sign_rate")
            m["wx"] = s.get("wx")
            m["wx_rate"] = s.get("wx_rate")
            m["mv_qty"] = s.get("mv_qty")
            _rt = (s.get("visits") or 0) + (s.get("signs") or 0)
            m["timely"] = _rt
            m["timely_rate"] = _rate(_rt, s.get("resources") or 0)
            if d in res:
                res[d] = s.get("resources") or 0

    # ---------- 部门健康度诊断（领导视角） ----------
    # 设计：健康目标优先采用「用户给定的具体业务考核规则」（如 员工见客签单转化率 ≥ 30%），
    #       无具体考核标准的指标才以「公司平均水平」（各指标全员均值）为兜底基准
    #       （用户原始口径：不知道具体健康目标的以公司平均水平为准）。
    # 比率类指标“越高越好”(forward)；核心健康指标 = 见客成单转化率（成单数÷见客数，即用户要求保留的「员工见客签单转化率」）。
    HEALTH_METRICS = [
        ("见客率", "visit_rate"),
        ("成单率", "sign_rate"),
        ("见客签单转化率", "visit_sign_rate_funnel"),
        ("加微率", "wx_rate"),
        ("邀约率", "invite_rate"),
        ("高质量资源占比", "hq_rate"),
        ("接通率", "connect_rate"),
        ("30秒接通率", "s30_rate"),
    ]
    # 用户给定的具体健康考核目标（明确值，非均值）；未列出的指标以公司均值兜底。
    # —— 例：见客成单转化率 30% 取自「重点指标·网销/电邀模式」业务考核线。
    # 如需新增其它指标的具体规则，在此字典追加即可（如 "wx_rate": 0.40）。
    HEALTH_TARGETS = {
        "visit_sign_rate_funnel": 0.30,
    }
    health_depts = [d for d in depts if (pay_metrics.get(d) or {}).get("resources")]
    # 公司平均目标（各指标全员均值，剔除 None）—— 兜底基准
    health_avg = {}
    for _label, _key in HEALTH_METRICS:
        _vals = [pay_metrics[d][_key] for d in health_depts if pay_metrics[d].get(_key) is not None]
        health_avg[_key] = round(sum(_vals) / len(_vals), 4) if _vals else None
    health_metrics = {}
    health_summary = {}
    for d in health_depts:
        m = pay_metrics[d]
        per = {}
        below = []
        for _label, _key in HEALTH_METRICS:
            v = m.get(_key)
            # 优先用具体考核规则目标，否则用公司均值兜底
            if _key in HEALTH_TARGETS:
                t = HEALTH_TARGETS[_key]
                tsrc = "rule"
            else:
                t = health_avg.get(_key)
                tsrc = "avg"
            if v is None or t is None:
                per[_key] = {"label": _label, "value": v, "target": t, "target_source": tsrc, "gap": None, "status": "na"}
            else:
                gap = round(v - t, 4)
                status = "healthy" if v >= t else "unhealthy"
                per[_key] = {"label": _label, "value": v, "target": t, "target_source": tsrc, "gap": gap, "status": status}
                if status == "unhealthy":
                    below.append({"label": _label, "key": _key, "gap": gap})
        # 整体健康判定：以「见客签单转化率」（漏斗口径，用户指定保留指标）为核心指标，按业务规则 30% 判定
        kp = per.get("visit_sign_rate_funnel")
        overall = "healthy" if (kp and kp["status"] == "healthy") else "unhealthy"
        health_summary[d] = {
            "status": overall,
            "below_count": len(below),
            "below": below,
            "key_metric": "visit_sign_rate_funnel",
        }
        health_metrics[d] = per
    health = {
        "metrics_def": [{"label": l, "key": k} for l, k in HEALTH_METRICS],
        "targets": {k: HEALTH_TARGETS.get(k) for _, k in HEALTH_METRICS},  # 具体规则目标（命中则为数值，否则 None）
        "avg_targets": health_avg,                                        # 公司均值兜底
        "metrics": health_metrics,
        "summary": health_summary,
    }

    # ---------- 部门级 8 项重点指标（组合图 · 口径对齐 build_company_summary） ----------
    # 用户 2026-09-14 指定：组合图横轴＝部门；三项「陪同见客签单转化率」按
    #   「私客所属人姓名」映射到「员工所属部门」归口（与人员明细口径一致）。
    from datetime import timedelta as _td2
    _w30b = (pay_max_date - _td2(days=30)) if pay_max_date else None
    net_depts = [b["部门"] for b in bus_list
                 if any((e.get("funnel_source") == "net") for e in by_dept.get(b["部门"], []))]

    # 1) 付费客资源签单转化率（近30天）：按 AX 部门归口，与公司级同口径（仅计可归类业务线的资源）
    _prd = {d: [0, 0] for d in net_depts}   # [资源数, 签单数]
    if pay_rows:
        for r in pay_rows:
            if _pay_excluded(r):
                continue
            if _w30b and r["date"] < _w30b:
                continue
            if _bizline_of(r) is None:
                continue
            ax = str(r.get("员工所属部门") or "").strip()
            short = None
            for kw in _DEPT_KEYWORDS:
                if kw in ax:
                    short = kw
                    break
            if short not in _prd:
                continue
            _prd[short][0] += 1
            if str(r.get("signed") or "").strip() == "是":
                _prd[short][1] += 1
    if not pay_rows:
        # 2026-09-21：源表移除「付费客资源」sheet ⇒ 部门级回退值改由「网销漏斗」逐人列派生
        for e in employees:
            if e.get("funnel_source") != "net":
                continue
            _d = e.get("部门标准")
            if _d not in _prd:
                continue
            _prd[_d][0] += int(round(_funnel_num(e, "有效资源量") or 0))
            _prd[_d][1] += int(round(_funnel_num(e, "付费客30天签单量") or 0))

    # 2) 三项陪同：面谈明细 → 私客所属人姓名 → 员工所属部门
    _person_dept = {}
    for e in employees:
        _nm = e.get("姓名")
        if _nm:
            _person_dept[_nm] = e.get("部门标准")
    _iv_rows, _iv_p = _load_interview()
    _iv_cols = set(_iv_rows[0].keys()) if _iv_rows else set()
    _sign_cols = ["首次收款时间"] if any("首次收款时间" in k for k in _iv_cols) else ["客单号", "订单号"]
    _roles = [r for r, _ in INTERVIEW_ROLE_MAP]
    _iv_acc = {d: {role: [0, 0] for role in _roles} for d in net_depts}   # [记录数, 签单数]
    _iv_total = _iv_mapped = 0
    for r in _iv_rows:
        role = str(r.get("岗位") or "").strip()
        if role not in _roles:
            continue
        _iv_total += 1
        dep = _person_dept.get(str(r.get("私客所属人姓名") or "").strip())
        if dep not in _iv_acc:
            continue
        _iv_mapped += 1
        signed = any(r.get(c) is not None and str(r.get(c)).strip() != "" for c in _sign_cols)
        a = _iv_acc[dep][role]
        a[0] += 1
        if signed:
            a[1] += 1

    def _r2(a):
        return round(a[1] / a[0], 4) if a[0] else None

    km_struct = {}
    km_vals = {}
    # 用户口径（2026-09-21）：员工见客签单转化率（不含陪谈）→ 源表「见客」sheet 主面谈口径（部门级）
    try:
        _visit_dept = visit_agg().get("dept") or {}
    except Exception:
        _visit_dept = {}
    for d in net_depts:
        emps = [e for e in by_dept.get(d, []) if e.get("funnel_source") == "net"]

        def _ds(key, _emps=emps):
            t = 0.0
            for e in _emps:
                v = (e.get("funnel") or {}).get(key)
                if isinstance(v, (int, float)):
                    t += v
            return t

        # 用户口径（2026-09-21）：当月见客签单转化率改取「分公司网销漏斗」c36
        #   「接收资源30天内见客后的转化率」= Σ接收资源30天内见客后签单量 ÷ Σ接收资源30天内见客量。
        #   实测「分公司网销漏斗」c36 与「网销漏斗」c37 <CITY> 35 人逐人值 0 差异 ⇒ 用逐人列聚合等价。
        v30 = _ds("接收资源30天内见客量")
        s30 = _ds("接收资源30天内见客后签单量")
        sales = _ds("销售额")
        h = head.get(d) or 0
        mg, stf = _staff_split(emps)
        km_struct[d] = {"mgr": len(mg), "staff": len(stf), "text": "%d+%d" % (len(mg), len(stf))}
        _net_rate_d = _net_rate_for_dept(d)      # 用户口径：A列名称匹配 → AL 列
        km_vals[d] = {
            "付费客资源签单转化率": (_net_rate_d if _net_rate_d is not None
                                     else _r2(_prd.get(d, [0, 0]))),
            # 用户口径（2026-09-21）：员工见客签单转化率（考核月）
            #   = 考核月见客签单量 ÷ 考核月的演示用户（**不含陪谈**）→ 源表「见客」sheet 主面谈
            "员工见客签单转化率": (_visit_dept.get(d) or {}).get("rate"),
            "业务经理陪同见客签单转化率": _r2(_iv_acc[d]["业务经理"]),
            "中台老师陪同见客签单转化率": _r2(_iv_acc[d]["中台"]),
            "事业部总陪同见客签单转化率": _r2(_iv_acc[d]["管理者"]),
            # 当月见客签单转化率（考核月）：保持 c36 口径（含陪谈的见客口径另有其表，此处不混用）
            "当月见客签单转化率": (s30 / v30) if v30 else None,
            "演示用户": (sales / h) if h else None,
        }

    # 指标顺序 / 名称 / 单位 / 轴（0=比率% 左轴，1=金额元 右轴）；target=None 表示无统一目标线
    _IND = [
        ("付费客资源签单转化率", "付费客资源签单转化率（近30天）", "付费客资源签单转化率", "%", 0, None),
        ("员工见客签单转化率", "员工见客签单转化率（考核月）", "员工见客签单转化率", "%", 0, CS_TARGET_EMP_VISIT),
        ("业务经理陪同见客签单转化率", "业务经理陪同见客签单转化率", "业务经理陪同转化率", "%", 0, CS_TARGET_IV_VISIT),
        ("中台老师陪同见客签单转化率", "中台老师陪同见客签单转化率", "中台老师陪同转化率", "%", 0, CS_TARGET_IV_VISIT),
        ("事业部总陪同见客签单转化率", "事业部总陪同见客签单转化率", "事业部总陪同转化率", "%", 0, CS_TARGET_IV_VISIT),
        ("当月见客签单转化率", "当月见客签单转化率（考核月）", "当月见客签单转化率", "%", 0, CS_TARGET_MONTH_VISIT),
        ("演示用户", "演示用户（考核月）", "演示用户（右轴·元）", "元", 1, CS_TARGET_PER_CAPITA),
    ]
    km8_dept = {
        "depts": net_depts,
        "indicators": [
            {
                "key": k, "name": nm, "short": short, "unit": unit, "axis": axis,
                "data": [km_vals[d].get(k) for d in net_depts],
                "target": tgt,
            }
            for (k, nm, short, unit, axis, tgt) in _IND
        ],
        "biz_targets": BIZ_TARGETS,
        "struct": km_struct,
        "iv_coverage": {"total": _iv_total, "mapped": _iv_mapped, "unmapped": _iv_total - _iv_mapped},
    }

    return {
        "depts": depts,
        "per_capita": per_capita,
        "visit": visit,
        "signrate": signrate,
        "resource_after_excl": res,
        "resource_unmapped": res_unmapped,
        "pay_metrics": pay_metrics,
        "health": health,
        "km8_dept": km8_dept,
    }


def build_biz_health(employees, bus_list, pay_rows, pay_max_date, dept_charts):
    """业务健康度评估（目标采用项目启动时已确认目标）。

    三项指标独立评定，再按「达标计数」综合：3/3 优秀 · 2/3 良好 · 1/3 预警 · 0/3 危险。
    范围：按部门 + 公司 两级。
      M1 付费客资源签单转化率（近30天）：按业务线（工商/跨境/法律/AI智能体/交易）对比 BIZ_TARGETS 区间目标。
      M2 当月见客签单转化率（考核月=近30天窗口）：= 网销漏斗「接收资源30天内见客后的转化率」，
         部门级 = Σ接收资源30天内见客后签单量 ÷ Σ接收资源30天内见客量，vs 30%。
      M3 各部门演示用户：部门人均销售额 vs 20,000元。
    仅对网销5部门（有业务线映射者）生成；电邀/会计不在此健康体系内。
    """
    if pay_max_date:
        wstart = pay_max_date - timedelta(days=30)
    else:
        wstart = None

    # ---- M1：付费客资源签单转化率（近30天），按业务线 ----
    m1_lines = {name: {"total": 0, "signed": 0} for name, _ in BIZ_LINE_RULES}
    for r in pay_rows:
        if _pay_excluded(r):
            continue
        if wstart and r["date"] < wstart:
            continue
        bl = _bizline_of(r)
        if bl is None:
            continue
        m1_lines[bl]["total"] += 1
        if str(r.get("signed") or "").strip() == "是":
            m1_lines[bl]["signed"] += 1
    m1_by_line = {}
    for name, _ in BIZ_LINE_RULES:
        t = BIZ_TARGETS[name]
        total = m1_lines[name]["total"]
        signed = m1_lines[name]["signed"]
        # 用户口径（2026-09-15）：优先取「网销漏斗」A列名称匹配行的 AL 列
        rate = _net_rate_for_bizline(name)
        if rate is None:
            rate = (signed / total) if total > 0 else None
        st, stext = _status_interval(rate, t["lower"], t["upper"])
        m1_by_line[name] = {
            "value": rate, "display": _fmt_ratio(rate),
            "total": total, "signed": signed,
            "target": t, "target_display": t["display"],
            "status": st, "status_text": stext,
        }
    dept_to_biz = {v: k for k, v in BIZ_TO_DEPT.items()}

    # ---- M2：见客签单转化率（= 网销漏斗「接收资源30天内见客后的转化率」），按部门 ----
    # 用户口径（2026-09-13 确认）：见客签单转化率 = 源表「网销漏斗」sheet 的
    #   「接收资源30天内见客后的转化率」= 接收资源30天内见客后签单量 ÷ 接收资源30天内见客量。
    # 部门级 = Σ「接收资源30天内见客后签单量」÷ Σ「接收资源30天内见客量」。
    visit30_by_dept = {}
    sign30_by_dept = {}
    for e in employees:
        d = e.get("部门标准")
        f = e.get("funnel") or {}
        v = f.get("接收资源30天内见客量")
        s = f.get("接收资源30天内见客后签单量")
        if isinstance(v, (int, float)):
            visit30_by_dept[d] = visit30_by_dept.get(d, 0) + v
        if isinstance(s, (int, float)):
            sign30_by_dept[d] = sign30_by_dept.get(d, 0) + s

    per_capita = dept_charts.get("per_capita", {})

    # ---- 部门级三指标 + 综合 ----
    dept_health = {}
    for b in bus_list:
        d = b["部门"]
        bl = dept_to_biz.get(d)
        if not bl:  # 仅网销5部门进入健康体系
            continue
        # M1
        if bl in m1_by_line:
            m1 = dict(m1_by_line[bl])
        else:
            m1 = {"value": None, "display": "-", "target_display": "-",
                  "status": "na", "status_text": "无数据", "total": 0, "signed": 0}
        # M2 见客签单转化率 = 接收资源30天内见客后的转化率（Σ签单量÷Σ见客量）
        vis = visit30_by_dept.get(d, 0)
        sgn = sign30_by_dept.get(d, 0)
        m2v = (sgn / vis) if vis > 0 else None
        m2st, m2text = _status_pos(m2v, 0.30)
        m2 = {"value": m2v, "display": _fmt_ratio(m2v),
              "visit": vis, "sign": sgn,
              "target": {"kind": "positive", "target": 0.30}, "target_display": "30%",
              "status": m2st, "status_text": m2text}
        # M3
        pc = per_capita.get(d)
        m3st, m3text = _status_pos(pc, 20000)
        m3 = {"value": pc, "display": _fmt_amount(pc),
              "target": {"kind": "positive", "target": 20000}, "target_display": "20,000元",
              "status": m3st, "status_text": m3text}
        metrics = {"pay_sign_rate": m1, "visit_sign_rate": m2, "per_capita": m3}
        pass_count = sum(1 for m in metrics.values() if m["status"] == "pass")
        level = {3: "优秀", 2: "良好", 1: "预警", 0: "危险"}[pass_count]
        dept_health[d] = {"metrics": metrics, "pass_count": pass_count, "level": level}

    # ---- 公司级（仅网销口径）----
    # M1 公司：全部业务线达标才算达标
    m1_all_pass = all(m["status"] == "pass" for m in m1_by_line.values())
    m1_pass_n = sum(1 for m in m1_by_line.values() if m["status"] == "pass")
    m1_company = {
        "status": "pass" if m1_all_pass else "fail",
        "status_text": "全线达标" if m1_all_pass else "存在未达标业务线",
        "display": "%d/%d 业务线达标" % (m1_pass_n, len(m1_by_line)),
        "by_line": {k: {"display": v["display"], "target_display": v["target_display"],
                        "status": v["status"], "status_text": v["status_text"]}
                    for k, v in m1_by_line.items()},
    }
    # M2 公司：按部门去重后合并「接收资源30天内见客量/见客后签单量」
    net_emps = [e for e in employees if e.get("funnel_source") == "net"]
    net_depts = []
    for e in net_emps:
        d = e.get("部门标准")
        if d and d not in net_depts:
            net_depts.append(d)
    c_vis = sum(visit30_by_dept.get(d, 0) for d in net_depts)
    c_sgn = sum(sign30_by_dept.get(d, 0) for d in net_depts)
    c_m2v = (c_sgn / c_vis) if c_vis > 0 else None
    c_m2st, c_m2text = _status_pos(c_m2v, 0.30)
    m2_company = {"value": c_m2v, "display": _fmt_ratio(c_m2v),
                  "target": {"kind": "positive", "target": 0.30}, "target_display": "30%",
                  "status": c_m2st, "status_text": c_m2text}
    # M3 公司：Σ网销月实算业绩 / 网销人数
    c_sales = sum(e.get("月累计实算利润业绩") or 0 for e in net_emps)
    c_pc = (c_sales / len(net_emps)) if net_emps else None
    c_m3st, c_m3text = _status_pos(c_pc, 20000)
    m3_company = {"value": c_pc, "display": _fmt_amount(c_pc),
                  "target": {"kind": "positive", "target": 20000}, "target_display": "20,000元",
                  "status": c_m3st, "status_text": c_m3text}
    comp_metrics = {"pay_sign_rate": m1_company, "visit_sign_rate": m2_company, "per_capita": m3_company}
    comp_pass = sum(1 for m in comp_metrics.values() if m["status"] == "pass")
    comp_level = {3: "优秀", 2: "良好", 1: "预警", 0: "危险"}[comp_pass]
    dist = {"优秀": 0, "良好": 0, "预警": 0, "危险": 0}
    for v in dept_health.values():
        dist[v["level"]] = dist.get(v["level"], 0) + 1
    return {
        "metrics_def": [
            {"key": "pay_sign_rate", "label": "付费客资源签单转化率（近30天）",
             "target_note": "业务线区间目标：工商15-24% / 跨境8-15% / 法律7-10% / AI智能体4-7% / 交易3-5%"},
            {"key": "visit_sign_rate", "label": "当月见客签单转化率（考核月=近30天窗口）",
             "target_note": "目标 30%（= 网销漏斗「接收资源30天内见客后的转化率」，部门级=Σ见客后签单量÷Σ见客量）"},
            {"key": "per_capita", "label": "各部门演示用户",
             "target_note": "目标 20,000元"},
        ],
        "levels": {"3": "优秀", "2": "良好", "1": "预警", "0": "危险"},
        "composite_rule": "三项指标独立评定，按达标计数综合：3/3 优秀 · 2/3 良好 · 1/3 预警 · 0/3 危险",
        "company": {"metrics": comp_metrics, "pass_count": comp_pass, "level": comp_level},
        "departments": dept_health,
        "dist": dist,
    }


# ---------- 某公司汇总 · 8 项重点指标（演示用户） ----------
# 数据源：源 Excel「网销漏斗」sheet（含某公司人员行，逐人聚合到分公司级）。
#   · 近30天块（G:AM，idx6-38）  ：资源/见客/付费客等
#   · 当前业务月块（=考核月，AN:BB，idx39-53）：接通/30s沟通/见客/销售额等
# 「业务经理 / 中台老师 / 事业部总 陪同见客签单转化率」暂无数据源 → 置空，
# 待用户新增 sheet 后接入（只需在 _CS_ROWS / build_company_summary 中补数据项即可）。
CS_TARGET_EMP_VISIT = 0.30      # 员工见客签单转化率目标
CS_TARGET_MONTH_VISIT = 0.30    # 当月见客签单转化率目标
CS_TARGET_PER_CAPITA = 20000    # 演示用户目标（元）
CS_TARGET_IV_VISIT = 0.30       # 业务经理/中台老师/事业部总 三档陪同见客签单转化率目标（2026-09-16 图定）


def build_company_summary(employees, pay_rows, pay_max_date):
    """某公司（网销）8 项重点指标汇总。

    返回 {"rows": [...], "notes": [...], "total_net": n}；每行：
      name/display/value/unit/source/target/status/note/available
    available=False 表示当前无数据源（先留空，待接入新 sheet）。
    """
    net = [e for e in employees if e.get("funnel_source") == "net"]
    n = len(net)

    def _sum(key):
        t = 0.0
        for e in net:
            v = (e.get("funnel") or {}).get(key)
            if isinstance(v, (int, float)):
                t += v
        return t

    # —— 近30天块：见客/签单（供「当月见客签单转化率」= 分公司网销漏斗 c36 口径）——
    visit30 = _sum("接收资源30天内见客量")
    sign30 = _sum("接收资源30天内见客后签单量")

    # 用户口径（2026-09-21，源表「目标」sheet 第 5 行「定义」原文）：
    #   ② 员工见客签单转化率（考核月）= 考核月见客签单量 ÷ 考核月的演示用户（**不含陪谈**）
    #      → 源表「见客」sheet：记录类型=主面谈 ∧ 岗位∈{网销岗,综法岗}，签单=首次收款时间非空
    #   ⑥ 当月见客签单转化率（考核月）=「分公司网销漏斗」「接收资源30天内见客后的转化率」
    #      （= Σ见客后签单量 ÷ Σ见客量），与本行的 rate_emp 不再是同口径。
    #   ⚠️ 旧实现误用「当前业务月块 · 30s有效沟通后的见客转化 ÷ 30s有效沟通后的见客」，
    #      该列<CITY>侧本期全 0（全国仅 16 行非 0）⇒ 页面恒显示 0.00%。已废弃，勿回退。
    try:
        _vg = visit_agg()
        _v_comp = _vg.get("company") or {"den": 0, "num": 0, "rate": None}
    except Exception:
        _vg, _v_comp = {}, {"den": 0, "num": 0, "rate": None}
    rate_emp = _v_comp.get("rate")
    rate_month = (sign30 / visit30) if visit30 else None

    # —— 考核月销售额 → 演示用户 ——
    sales_m = _sum("销售额")
    per_capita = (sales_m / n) if n else None

    # —— 演示用户（考核月）——
    mgrs, staffs = _staff_split(net)
    struct_str = "%d+%d" % (len(mgrs), len(staffs))

    # —— 付费客资源签单转化率（近30天，公司级 = Σ签单 ÷ Σ资源）——
    pay_total = pay_signed = 0
    if pay_rows:
        wstart = (pay_max_date - timedelta(days=30)) if pay_max_date else None
        for r in pay_rows:
            if _pay_excluded(r):
                continue
            if wstart and r["date"] < wstart:
                continue
            if _bizline_of(r) is None:
                continue
            pay_total += 1
            if str(r.get("signed") or "").strip() == "是":
                pay_signed += 1
    pay_rate = (pay_signed / pay_total) if pay_total else None

    # —— 资源三项（全量有效资源；不按业务线过滤，用于「一二三级」的 1 级=分公司）——
    res_total = res_sign = res_wx = 0
    if pay_rows:
        _ws2 = (pay_max_date - timedelta(days=30)) if pay_max_date else None
        for r in pay_rows:
            if _pay_excluded(r):
                continue
            if _ws2 and r["date"] < _ws2:
                continue
            res_total += 1
            if str(r.get("signed") or "").strip() == "是":
                res_sign += 1
            if _is_yes(r.get("加微")):
                res_wx += 1
    res_sign_rate = (res_sign / res_total) if res_total else None
    res_wx_rate = (res_wx / res_total) if res_total else None
    RES_SRC = "付费客资源 sheet · 近30天 · 已剔除无效资源（%s）" % PAY_EXCL_DESC

    # 2026-09-21：源表移除「付费客资源」sheet ⇒ 公司级资源类指标改由「网销漏斗」逐人列派生
    #（用户口径：资源以漏斗数据中的有效资源量即可）。
    if not pay_rows:
        try:
            _fr_per_c, _fr_dept_c, _fr_comp_c = _funnel_res_agg(employees)
        except Exception:
            _fr_comp_c = None
        if _fr_comp_c:
            res_total = int(_fr_comp_c.get("resources") or 0)
            res_sign = int(_fr_comp_c.get("signs") or 0)
            res_wx = int(_fr_comp_c.get("wx") or 0)
            res_sign_rate = _fr_comp_c.get("sign_rate")
            res_wx_rate = _fr_comp_c.get("wx_rate")
            pay_total, pay_signed = res_total, res_sign
            pay_rate = _fr_comp_c.get("sign_rate")
            RES_SRC = ("网销漏斗 · 有效资源量（用户口径 2026-09-21）"
                       "｜Σ付费客30天签单量 %d ÷ Σ有效资源量 %d" % (res_sign, res_total))

    def _row(name, value, unit, source, target=None, note="", available=True, target_val=None, status=None):
        st = "na"
        if available and value is not None and target_val is not None:
            st = _status_pos(value, target_val)[0]
        if status is not None:
            st = status
        # 未完成目标差距：实际值 − 目标（比率→pp；金额→元）
        gap = None
        gap_txt = "—"
        if available and value is not None and target_val is not None:
            d = value - target_val
            gap = d
            if unit == "%":
                gap_txt = ("%+.1fpp" % round(d * 100, 1))
            elif unit == "元":
                gap_txt = ("%+d元" % round(d))
            else:
                gap_txt = "%+g" % d
        health = {"pass": "健康", "warn": "预警", "fail": "不健康"}.get(st, "—") if available else "—"
        return {
            "name": name,
            "value": value,
            "display": _fmt_ratio(value) if unit == "%" else (
                _fmt_amount(value) if unit == "元" else (str(value) if value is not None else "—")),
            "unit": unit,
            "source": source,
            "target": target,
            "target_val": target_val,
            "gap": gap,
            "gap_text": gap_txt,
            "health": health,
            "status": st if available else "na",
            "note": note,
            "available": available,
        }

    struct_ok = (len(mgrs) >= 1 and 5 <= len(staffs) <= 6)

    # —— 陪同见客签单转化率（面谈明细：按「岗位」分组的陪谈记录签单率）——
    # 口径：分母 = 该岗位陪谈记录数（面谈明细中 AZ=岗位 的行）；
    #       分子 = 其中「已签单」的记录数。是否签单以「首次收款时间」(AB 列) 非空为准。
    interview, ipath = _load_interview()

    def _iv_nz(v):
        return v is not None and str(v).strip() != ""

    # 判定列：优先「首次收款时间」(AB)；若该表头缺失（表结构变动）则回退「客单号/订单号」
    _iv_cols = set(interview[0].keys()) if interview else set()
    if any("首次收款时间" in k for k in _iv_cols):
        SIGN_COLS, SIGN_DESC = ["首次收款时间"], "首次收款时间非空"
    else:
        SIGN_COLS, SIGN_DESC = ["客单号", "订单号"], "客单号/订单号非空"

    def _is_signed(d):
        return any(_iv_nz(d.get(c)) for c in SIGN_COLS)

    role_stat = {}
    for role, _name in INTERVIEW_ROLE_MAP:
        sub = [d for d in interview if str(d.get("岗位") or "").strip() == role]
        denom = len(sub)
        numer = sum(1 for d in sub if _is_signed(d))
        role_stat[role] = {"n": denom, "signed": numer,
                           "rate": (numer / denom) if denom else None}
    IV_SRC = "面谈明细 · 岗位=%s（陪谈记录） · 签单(%s)÷陪谈见客" % ("%s", SIGN_DESC)

    NA_SRC = "源数据暂缺（待新增 sheet 接入）"
    # —— 用户口径（2026-09-15）：本指标 =「网销漏斗」sheet 中 A 列 =「某公司」行的
    #    AL 列「付费客30天签单转化率」（=付费客30天签单量 ÷ 有效资源量）。
    _net_sum, _net_sum_path = _load_net_summary()
    _comp_rate, _comp_meta = _net_rate_lookup("某公司", _net_sum)
    if _comp_rate is None:                      # 未命中该行 → 回退原付费客资源口径
        _comp_rate = pay_rate
    _net_date = _NET_SUMMARY_CACHE.get("date")  # 该表 A1 标注的数据日期
    _comp_src = ("网销漏斗 · A列「某公司」行 · AL「付费客30天签单转化率」"
                 + ("（源表标注 %s）" % _net_date if _net_date else "")
                 if _comp_meta else "付费客资源 sheet · 近30天（Σ签单÷Σ资源）")
    _comp_note = (("付费客30天签单量 %s / 有效资源 %s" % (_comp_meta.get("sign"), _comp_meta.get("valid")))
                  if _comp_meta else
                  ("签单 %d / 资源 %d" % (round(pay_signed), round(pay_total)) if pay_total else "无资源"))
    rows = [
        _row("演示用户付费客资源签单转化率（近30天）", _comp_rate, "%",
             _comp_src, note=_comp_note),
        _row("员工见客签单转化率（考核月）", rate_emp, "%",
             "见客 sheet · 记录类型=主面谈（不含陪谈）· 岗位∈{网销岗,综法岗} · 签单(首次收款时间非空)÷见客量",
             target="30%", target_val=CS_TARGET_EMP_VISIT,
             note=("见客签单 %d / 主面谈见客 %d ｜ %s" % (_v_comp["num"], _v_comp["den"], VISIT_WINDOW_NOTE)
                   if _v_comp.get("den") else "源表「见客」sheet 无主面谈记录")),
    ]
    for role, rname in INTERVIEW_ROLE_MAP:
        stt = role_stat.get(role, {"n": 0, "signed": 0, "rate": None})
        rows.append(_row(
            rname, stt["rate"], "%", IV_SRC % role,
            note=("签单 %d / 陪谈见客 %d" % (stt["signed"], stt["n"])) if stt["n"] else NA_SRC,
            available=(stt["n"] > 0),
        ))
    rows += [
        _row("当月见客签单转化率（考核月）", rate_month, "%",
             "分公司网销漏斗 · 「接收资源30天内见客后的转化率」（= 接收资源30天内见客后签单量÷见客量）",
             target="30%", target_val=CS_TARGET_MONTH_VISIT,
             note="见客签单 %d / 见客 %d" % (round(sign30), round(visit30)) if visit30 else "无见客"),
        _row("演示用户（考核月）", per_capita, "元",
             "网销漏斗 · 当前业务月块 · 销售额 ÷ 网销人数",
             target="20,000元", target_val=CS_TARGET_PER_CAPITA,
             note="销售额 %s / %d 人" % (_fmt_amount(sales_m), n)),
        _row("演示用户（考核月）", struct_str, "结构",
             "网销漏斗 · 网销岗位人数结构",
             target="1+5~6", status=("pass" if struct_ok else "warn"),
             note="管理岗 %d · 员工 %d（合计 %d 人）" % (len(mgrs), len(staffs), n)),
    ]
    # —— 资源三项（1 级 = 某公司；全量有效资源）——
    resource_rows = [
        _row("演示用户", (res_total if res_total else None), "条", RES_SRC,
             note="有效资源 %d 条（AX 部门 / AS 所属人归口）" % res_total),
        _row("资源签单率", res_sign_rate, "%", RES_SRC + " · 成单数 ÷ 演示用户",
             note="成单 %d / 资源 %d" % (res_sign, res_total) if res_total else "无资源"),
        _row("资源加微信", (res_wx if res_total else None), "条",
             RES_SRC + " · 「是否加微」=是 的条数",
             note=("加微 %d 条（占比见下行）" % res_wx) if res_total else "无资源"),
        # 用户 2026-09-14 要求：资源指标在每个级别都要有「资源加微率」，1 级＝某公司
        _row("资源加微率", res_wx_rate, "%", RES_SRC + " · 加微数 ÷ 演示用户",
             note=("加微 %d / 资源 %d" % (res_wx, res_total)) if res_total else "无资源"),
    ]
    iv_note = ("面谈明细已接入：%s（共 %d 条记录）" % (os.path.basename(ipath), len(interview))
               if interview else "面谈明细未找到（%s）" % (ipath or "四大模式-重点指标表格.xlsx"))
    return {
        "title": "某公司汇总 · 8 项重点指标",
        "subtitle": "口径：网销漏斗（近30天块 / 当前业务月块＝考核月）· 面谈明细（陪谈）· 按某公司聚合",
        "total_net": n,
        "rows": rows,
        "resource_rows": resource_rows,
        "notes": [
            "数据源：源表「网销漏斗」sheet（逐人聚合为分公司级）+「四大模式-重点指标表格.xlsx」的「面谈明细」sheet +「付费客资源」sheet。",
            "「演示用户付费客资源签单转化率（近30天）」取自「四大模式-重点指标表格.xlsx」→「网销漏斗」sheet 中 A 列名称匹配的行（A 列 =「某公司」即分公司级）的 AL 列「付费客30天签单转化率」＝付费客30天签单量 ÷ 有效资源量。**该口径覆盖四个级别**：A 列 =「某公司」→ 分公司级、「<CITY>XX事业部」→ 部门级、「<CITY>XX业务X部」→ 业务部级、「姓名(工号)」→ 员工级（员工明细表「🎯付费客资源签单转化率」列）。",
            "「考核月」= 网销漏斗「当前业务月」块（AM:BC）；「近30天」= 其余块。",
            "「健康状况」= 达标(≥目标,健康) / 目标5%内(预警) / 低于目标(不健康)；「未完成目标差距」= 实际值 − 目标（正=领先，负=未完成）。",
            "资源指标（演示用户/资源签单率/资源加微信/资源加微率）只统计全量有效资源，按 AX「部门」归部门、AS「当前私客所属人」归人；「资源加微率」= 加微数 ÷ 演示用户。",
            "三项「陪同见客签单转化率」按面谈明细的「岗位」列分组：管理者=事业部总陪谈 / 业务经理=业务经理陪谈 / 中台=中台陪谈；比率 = 该岗位陪谈记录中「已签单」条数 ÷ 该岗位陪谈记录数；「已签单」以「首次收款时间」(AB 列) 非空为准。",
            iv_note,
        ],
    }


def build_meta():
    net_by_id, net_by_name = _load_funnel()
    dy_by_id, dy_by_name, dy_people = _load_dy_funnel()
    ltc_rows, data_date = _load_ltc()
    pay_rows, pay_max_date = _load_pay_resources()
    kuaiji = _load_kuaiji()

    # 日期格式化：支持 "08-21" 文本 或 Excel 序列号(如 46258)
    date_str = ""
    if data_date:
        ds = str(data_date).strip()
        try:
            mm, dd = ds.split("-")
            date_str = f"{int(mm)}月{int(dd)}日"
        except Exception:
            try:
                from datetime import datetime as _dt, timedelta as _td
                n = float(ds)
                dt = _dt(1899, 12, 30) + _td(days=n)
                date_str = f"{dt.month}月{dt.day}日"
            except Exception:
                date_str = ds

    # 部门排序：电邀置顶（电邀业务的漏斗来源不同，单独成一组）
    dept_order_eff = [DY_DEPT_STD] + [d for d in DEPT_ORDER if d != DY_DEPT_STD]

    employees = []
    seen = set()
    for row in ltc_rows:
        name = str(row.get("姓名") or "").strip()
        if not name:
            continue
        # 仅保留指定分公司（防御性筛选，防其他分公司数据混入）
        comp = str(row.get("公司") or "").strip()
        if COMPANY_KEYWORD not in comp:
            continue
        # 仅保留指定岗位的人员
        post = str(row.get("岗位") or "").strip()
        if post not in TARGET_POSTS:
            continue
        dept = str(row.get("部门") or "").strip()
        # 部门归类失败（如「<CITY>知识产权邀约一部」）时保留原始部门名，不丢弃人员
        std_dept = _classify_dept(dept) or dept or "未知部门"
        if name in seen:
            continue
        seen.add(name)
        emp_id = row.get("工号")
        emp_id_str = str(emp_id).strip() if emp_id is not None else ""

        # 漏斗数据匹配：电邀部门优先用「电邀漏斗」，其余用「网销漏斗」
        # 匹配口径：工号优先，姓名兜底
        funnel_data = {}
        funnel_source = None
        if std_dept == DY_DEPT_STD:
            if emp_id_str and emp_id_str in dy_by_id:
                funnel_data = dy_by_id[emp_id_str]
                funnel_source = "dy"
            elif name in dy_by_name:
                funnel_data = dy_by_name[name]
                funnel_source = "dy"
            else:
                # 电邀部门但电邀漏斗无此员工：回退到网销漏斗（兜底，工号优先），source=dy_missing
                funnel_data = (net_by_id.get(emp_id_str) if emp_id_str else None) or net_by_name.get(name, {})
                funnel_source = "dy_missing" if funnel_data else None
        else:
            # 网销岗：工号优先匹配，姓名兜底（源表为多分公司混合表，防跨公司重名串号）
            funnel_data = (net_by_id.get(emp_id_str) if emp_id_str else None) or net_by_name.get(name, {})
            if funnel_data:
                funnel_source = "net"

        emp = {
            "姓名": name,
            "工号": emp_id,
            "部门": dept,
            "事业部": row.get("事业部") or "",
            "部门标准": std_dept,
            "岗位": row.get("岗位") or "",
            "司龄": _to_num(row.get("司龄")),
            "职级": row.get("职级") or "",
            "三力达标": row.get("三力达标") or "",
            "月累计实算利润业绩": _to_num(row.get("月累计实算利润业绩")),
            "月累计核算利润业绩": _to_num(row.get("月累计核算利润业绩")),
            "当日新增业绩": _to_num(row.get("当日新增业绩")),
            "ltc": {k: _to_num(row.get(k)) for k, _, _ in LTC_METRICS if k != "三力达标"},
            "funnel": funnel_data,
            "funnel_source": funnel_source,
        }
        # 纯数值 LTC 子集用于图表
        emp["ltc"] = {k: _to_num(row.get(k)) for k, _, _ in LTC_METRICS if k != "三力达标"}
        # M-35：资源明细「近30天移交接收资源执行率」三件套（LTC流程数据 sheet；率=小数）
        for _k in ("近30天移交接收资源量", "近30天移交接收资源执行量", "近30天移交接收资源执行率"):
            emp["ltc"][_k] = _to_num(row.get(_k))
        emp["advice"] = build_advice(emp)
        employees.append(emp)

    # 会计部门改用（会计）sheet 数据：剔除 LTC 口径的会计员工（防残留），注入会计sheet员工
    employees = [e for e in employees if e.get("部门标准") != "会计"]
    employees.extend(_kuaiji_employees(kuaiji, dept_std_mode="会计"))

    # 按部门顺序+司龄排序（含动态部门兜底）
    def _dept_rank(d):
        return dept_order_eff.index(d) if d in dept_order_eff else 999
    employees.sort(key=lambda e: (_dept_rank(e["部门标准"]), e["部门标准"], -(e["司龄"] or 0)))

    # ---------- 人员级：资源指标(AS) / 陪同见客指标(面谈明细·私客所属人姓名) / 8 项重点指标 ----------
    try:
        _pay_person = _aggregate_pay_by_person(pay_rows, pay_max_date)
    except Exception:
        _pay_person = {}
    # 2026-09-21：源表已移除「付费客资源」sheet ⇒ 员工级资源指标改由「网销漏斗」逐人列派生
    #（用户口径：资源以漏斗数据中的有效资源量即可）
    if not _pay_person:
        try:
            _pay_person, _fr_dept, _fr_comp = _funnel_res_agg(employees)
        except Exception:
            _pay_person, _fr_dept, _fr_comp = {}, {}, {}
    else:
        _fr_dept, _fr_comp = {}, {}
    try:
        _iv_person = _aggregate_interview_by_person()
    except Exception:
        _iv_person = {}

    # 用户口径（2026-09-21）：员工级「员工见客签单转化率（考核月）」改用源表「见客」sheet
    #   主面谈（不含陪谈）口径；「当月见客签单转化率（考核月）」保持网销漏斗 c36 口径。
    try:
        _visit_emp = visit_agg().get("emp") or {}
    except Exception:
        _visit_emp = {}

    def _km8_of(e):
        f = e.get("funnel") or {}

        def _n(k):
            v = f.get(k)
            return v if isinstance(v, (int, float)) else None

        def _r(a, b):
            return round(a / b, 4) if (a is not None and b) else None

        pay = e.get("pay") or {}
        iv = e.get("iv") or {}
        # 见客 sheet 主面谈（不含陪谈）人员级明细：{den, num, rate}
        _v_emp = _visit_emp.get(str(e.get("姓名") or "").strip()) or {}
        # 用户口径（2026-09-21）：当月见客签单转化率 = 「分公司网销漏斗」c36
        #   「接收资源30天内见客后的转化率」（= 见客后签单量 ÷ 见客量）；
        #   逐人取值与「网销漏斗」c37 完全一致，故仍走 funnel 逐人列。
        v30 = _n("接收资源30天内见客量")
        s30 = _n("接收资源30天内见客后签单量")
        sale = _n("销售额")
        role_r = lambda r: (iv.get(r) or {}).get("rate")  # noqa: E731
        # 用户 2026-09-15 M-3：员工级（3 级）同样取「网销漏斗」AL 列（A 列「姓名(工号)」匹配）；
        # 未命中才回退原「付费客资源 sheet」口径，保证任何情况下不出现空值。
        _al_emp = _net_rate_for_emp(e)

        def _ivc(r):
            o = iv.get(r) or {}
            return {"den": o.get("n"), "num": o.get("signed")}
        return {
            "付费客资源签单转化率": (_al_emp if _al_emp is not None else pay.get("sign_rate")),
            # 员工见客签单转化率（考核月）：见客 sheet 主面谈（不含陪谈）｜分母 0 → None
            "员工见客签单转化率": _v_emp.get("rate"),
            "业务经理陪同见客签单转化率": role_r("业务经理"),
            "中台老师陪同见客签单转化率": role_r("中台"),
            "事业部总陪同见客签单转化率": role_r("管理者"),
            "当月见客签单转化率": _r(s30, v30),
            "演示用户": sale,
            "演示用户": pay.get("resources"),
            "资源签单率": pay.get("sign_rate"),
            "资源加微信": pay.get("wx"),
            # 用户 2026-09-14：员工级（3 级）资源指标同样要带「资源加微率」= 加微数 ÷ 演示用户
            "资源加微率": pay.get("wx_rate"),
            # 用户 2026-09-16 M-16：健康度总览「点击部门展开员工子表」需「每指标 5 子列
            #（资源量/见客量 · 签单量 · 转化率 · 目标 · 目标完成率）」中的分子/分母，此处一并给出；
            # 转化率本身沿用上方 km8 值（口径不变），detail 仅补充计数明细。演示用户给出岗位。
            "detail": {
                "付费客资源签单转化率": {"den": _n("有效资源量"), "num": _n("付费客30天签单量")},
                "员工见客签单转化率": {"den": _v_emp.get("den"), "num": _v_emp.get("num")},
                "业务经理陪同见客签单转化率": _ivc("业务经理"),
                "中台老师陪同见客签单转化率": _ivc("中台"),
                "事业部总陪同见客签单转化率": _ivc("管理者"),
                "当月见客签单转化率": {"den": v30, "num": s30},
                "演示用户": {"val": sale},
                "演示用户": {"pos": e.get("岗位")},
            },
        }

    for e in employees:
        e["pay"] = _pay_person.get(e["姓名"])
        e["iv"] = _iv_person.get(e["姓名"])
        e["km8"] = _km8_of(e)


    # ---------- 部门聚合（预设六部门+电邀置顶，动态部门追加） ----------
    def _new_bus(d):
        return {"部门": d, "人数": 0, "三力": {"英雄人力": 0, "标准人力": 0, "破零人力": 0, "未破零": 0, "无三力": 0, "其他": 0},
                "月实算业绩": 0, "月核算业绩": 0, "当日新增业绩": 0,
                "ltc_acc": {}, "funnel_acc": {}, "ltc_cnt": {}, "funnel_cnt": {}}
    bus_map = OrderedDict()
    for d in dept_order_eff:
        bus_map[d] = _new_bus(d)
    for e in employees:
        d = e["部门标准"]
        if d not in bus_map:
            bus_map[d] = _new_bus(d)
        b = bus_map[d]
        b["人数"] += 1
        sl = e["三力达标"]
        if sl in ("英雄人力", "标准人力", "破零人力", "未破零", "无三力"):
            b["三力"][sl] += 1
        else:
            b["三力"]["其他"] += 1
        b["月实算业绩"] += e["月累计实算利润业绩"] or 0
        b["月核算业绩"] += e["月累计核算利润业绩"] or 0
        b["当日新增业绩"] += e["当日新增业绩"] or 0
        for k in e["ltc"]:
            val = e["ltc"][k]
            if isinstance(val, (int, float)):
                b["ltc_acc"][k] = b["ltc_acc"].get(k, 0) + val
                b["ltc_cnt"][k] = b["ltc_cnt"].get(k, 0) + 1
        for k in e["funnel"]:
            val = e["funnel"][k]
            if isinstance(val, (int, float)):
                b["funnel_acc"][k] = b["funnel_acc"].get(k, 0) + val
                b["funnel_cnt"][k] = b["funnel_cnt"].get(k, 0) + 1

    bus_list = []
    for d in bus_map:
        b = bus_map[d]
        n = b["人数"] or 1
        ltc_avg = {}
        for k, acc in b["ltc_acc"].items():
            cnt = b["ltc_cnt"].get(k, 0) or 1
            ltc_avg[k] = round(acc / cnt, 2)
        funnel_avg = {}
        for k, acc in b["funnel_acc"].items():
            cnt = b["funnel_cnt"].get(k, 0) or 1
            funnel_avg[k] = round(acc / cnt, 2)
        # 部门漏斗来源：电邀部门用 dy，其他部门用 net
        bus_funnel_source = "dy" if d == DY_DEPT_STD else ("net" if any(
            e["部门标准"] == d and e["funnel_source"] == "net" for e in employees
        ) else (any(
            e["部门标准"] == d and e["funnel_source"] == "dy_missing" for e in employees
        ) and "dy_missing" or None))
        # 比率类指标重新取平均（仅对有值员工）
        dept_bus = {
            "部门": d,
            "人数": n,
            "三力": b["三力"],
            "月实算业绩": b["月实算业绩"],
            "月核算业绩": b["月核算业绩"],
            "当日新增业绩": b["当日新增业绩"],
            "ltc_avg": ltc_avg,
            "funnel_avg": funnel_avg,
            "funnel_source": bus_funnel_source,
        }
        dept_bus["advice"] = build_dept_advice(dept_bus)
        # 附加部门重点指标（按模块，未配置模块的部门为 None）
        bus_list.append(dept_bus)

    # ---------- 全局汇总 ----------
    total_sanli = {"英雄人力": 0, "标准人力": 0, "破零人力": 0, "未破零": 0, "无三力": 0, "其他": 0}
    total_real = 0
    total_check = 0
    total_today = 0
    for e in employees:
        sl = e["三力达标"]
        if sl in total_sanli:
            total_sanli[sl] += 1
        else:
            total_sanli["其他"] += 1
        total_real += e["月累计实算利润业绩"] or 0
        total_check += e["月累计核算利润业绩"] or 0
        total_today += e["当日新增业绩"] or 0

    # 部门颜色扩展：给「电邀」标准部门也配一个颜色
    extended_colors = dict(DEPT_COLORS)
    extended_colors[DY_DEPT_STD] = "#b05a52"  # 砖红（低饱和）：突出电邀为新型口径

    mtime = os.path.getmtime(XLSX)

    key_metrics = build_key_metrics(employees, pay_rows, pay_max_date, kuaiji)
    kuaiji_detail = _build_kuaiji_detail(kuaiji, key_metrics["kuaiji"], date_str or "")
    dept_charts = build_dept_charts(employees, bus_list, pay_rows, pay_max_date)
    biz_health = build_biz_health(employees, bus_list, pay_rows, pay_max_date, dept_charts)
    company_summary = build_company_summary(employees, pay_rows, pay_max_date)

    # 源表全空检测：某漏斗指标在所有对应员工中均无任何数值 → 标记 nodata，
    # 前端呈现为「源表未提供」，避免误导为解析缺失。（如「日人均资源量」「投产比」）
    def _nodata_keys(metrics, emps, src):
        keys = []
        pool = [e for e in emps if e.get("funnel_source") == src]
        for k, _r, _d in metrics:
            if not any(isinstance((e.get("funnel") or {}).get(k), (int, float)) for e in pool):
                keys.append(k)
        return set(keys)

    net_nodata = _nodata_keys(FUNNEL_METRICS, employees, "net")
    dy_nodata = _nodata_keys(DY_FUNNEL_METRICS, employees, "dy")

    return {
        "meta": {
            "file_mtime_ts": mtime,
            "aux_mtime_ts": _aux_mtimes(),
            "data_time": date_str or "8月21日",
            "sheet_title": f"人员数据分析-{date_str}",
            "dept_order": dept_order_eff,
            "dept_colors": extended_colors,
            "dy_count": len(dy_people),
        },
        "total": len(employees),
        "departments": list(bus_map.keys()),
        "summary": {
            "三力": total_sanli,
            "月实算业绩": total_real,
            "月核算业绩": total_check,
            "当日新增业绩": total_today,
        },
        "bus": bus_list,
        "employees": employees,
        "dept_charts": dept_charts,
        "biz_health": biz_health,
        "company_summary": company_summary,
        "funnel_metrics": [{"k": k, "ratio": r, "desc": d, "nodata": (k in net_nodata)} for k, r, d in FUNNEL_METRICS],
        "dy_funnel_metrics": [{"k": k, "ratio": r, "desc": d, "nodata": (k in dy_nodata)} for k, r, d in DY_FUNNEL_METRICS],
        # 会计口径漏斗指标：主页面「会计」部门里的员工详情要用它，否则会拿网销口径去套 → 全 '--'
        "kuaiji_funnel_metrics": [{"k": k, "ratio": r, "desc": d} for k, r, d in KJ_FUNNEL_METRICS],
        "ltc_metrics": [{"k": k, "ratio": r, "desc": d} for k, r, d in LTC_METRICS],
        "key_metrics": key_metrics,
        "kuaiji_detail": kuaiji_detail,
        # 用户 2026-09-16 需求 N：三张新增会计 sheet（抽检 / 个人出错率 / 续费明细）
        "kj_extra": _load_kj_extra(),
        # 用户 2026-09-28：公司及部门业绩（网销页展示<CITY>公司 + <CITY>各部门）
        "biz_perf": _load_biz_perf(),
    }


if __name__ == "__main__":
    d = build_meta()
    print("data_time:", d["meta"]["data_time"])
    print("总人数:", d["total"])
    for b in d["bus"]:
        print(f"  {b['部门']}: {b['人数']}人 | 三力:{b['三力']} | 月实算:{b['月实算业绩']:.0f}")
    # 示例：第一个员工建议
    e = d["employees"][0]
    print("\n示例员工:", e["姓名"], "| 部门:", e["部门标准"], "| 三力:", e["三力达标"])
    for s in e["advice"]:
        print("  [{}][{}] {}".format(s["level"], s["area"], s["text"]))

# ---------- 会计业绩：来自「会计高端账签单」sheet 的「业绩」列（按签单人归集，<CITY>端口） ----------
_KJ_PERF_CACHE = None

def _kj_signer_perf():
    """从「会计高端账签单」sheet 的「业绩」列，按「签单人」(<CITY>端口) 聚合。

    用户 2026-09-18 需求：会计数据分析的业绩指标统一以该 sheet 的「业绩」列计算，
    取代原 S0/规划师/主管(会计) 的「月度实算利润业绩」(利润口径)。
    返回 {"total": {签单人: 业绩合计}, "today": {签单人: 最新签单日业绩}, "max_date": "YYYY-MM-DD"}。
    """
    global _KJ_PERF_CACHE
    if _KJ_PERF_CACHE is not None:
        return _KJ_PERF_CACHE
    out = {}
    entries = []
    max_date = None
    try:
        wb = openpyxl.load_workbook(XLSX, data_only=True)
        if "会计高端账签单" in wb.sheetnames:
            ws = wb["会计高端账签单"]
            rows = list(ws.iter_rows(values_only=True))
            if rows:
                headers = [str(c).strip() if c is not None else "" for c in rows[0]]
                idx = {h: i for i, h in enumerate(headers) if h}
                si = idx.get("签单人"); ei = idx.get("业绩")
                ri = idx.get("区域"); ti = idx.get("业绩时间")
                if si is not None and ei is not None:
                    for r in rows[1:]:
                        if not r:
                            continue
                        region = str(r[ri] or "").strip() if ri is not None else ""
                        if region and region != "某公司":
                            continue
                        sgn = str(r[si] or "").strip()
                        if not sgn:
                            continue
                        v = r[ei]
                        try:
                            v = float(v or 0)
                        except Exception:
                            v = 0.0
                        d = ""
                        if ti is not None and r[ti] is not None:
                            d = (r[ti].strftime("%Y-%m-%d") if isinstance(r[ti], datetime) else str(r[ti])[:10])
                        out[sgn] = out.get(sgn, 0.0) + v
                        entries.append((sgn, v, d))
                        if d:
                            if max_date is None or d > max_date:
                                max_date = d
    except Exception:
        pass
    today = {}
    if max_date:
        for sgn, v, d in entries:
            if d == max_date:
                today[sgn] = today.get(sgn, 0.0) + v
    _KJ_PERF_CACHE = {"total": out, "today": today, "max_date": max_date}
    return _KJ_PERF_CACHE

