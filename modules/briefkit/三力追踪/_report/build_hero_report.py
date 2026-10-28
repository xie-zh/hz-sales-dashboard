# -*- coding: utf-8 -*-
"""
分公司 三力追踪 —— 英雄人力「完成情况」+「差距榜」

用法:
  python build_hero_report.py [源文件.xlsx]
  不传参则默认读 _report/../src_latest.xlsx

数据源: sheet 企服顾问明细
口径（已与 9月周期 20260827 版逐值校验一致）:
  - 人员范围 : 部门(I列) 含 ""
  - 利润口径 : R = 当月签单实算利润业绩(第18列)
  - 英雄100% : 标准 Y(第25列) / 是否达成 Z(第26列) / 差距 AA(第27列)
  - 英雄 50% : 标准 AF(第32列) / 是否达成 AG(第33列) / 差距 AH(第34列)
  - 差距     : 目标值 - 利润
  - 完成名单 : 是否达成 == "是"（50%名单要去掉已达成100%的人）
  - 差距榜（**统一口径，2026-09-24 起**）:
      未达成 且 **当月签单实算业绩 Q(第17列) > 0**（业绩为 0 或负的一律排除）
      且 `0 < 差距 < GAP_LIMIT(3000)`，按差距升序 —— 四个追踪全部适用
  - 破零追踪 : 门槛破零标准 S(第19列) / 是否达成 T(第20列)，差距 = S - 利润，同上口径
  - 有效人力 : 是否达成 AN(第40列) / 标准 AO(第41列)，差距 = AO - 利润，同上口径
周期推导（从文件名日期戳）:
  日 >= 15 → 属「下月」周期；日 <= 14 → 属「当月」周期
  例: 20260923 → 10月周期(9/15-10/14) ; 20260914 → 9月周期(8/15-9/14)

产出:
  report_img/hero_done.png       —— 英雄人力目标完成情况（100% + 50%）
  report_img/hero_gap50.png      —— 距英雄人力50%目标 差距榜
  report_img/hero_breakzero.png  —— 破零人力重点催办名单
  report_img/hero_valid.png      —— 有效人力差距追踪
  report_text1..4.txt / report_state.json
"""
import os
import re
import sys
import json
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import openpyxl

DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, DIR)
import mention_util  # 真实 @单人（<@userid>），缺失时回退纯文本 @姓名


IMG_DONE = os.path.join(DIR, "report_img", "hero_done.png")
IMG_GAP = os.path.join(DIR, "report_img", "hero_gap50.png")
IMG_ZERO = os.path.join(DIR, "report_img", "hero_breakzero.png")
IMG_VALID = os.path.join(DIR, "report_img", "hero_valid.png")
IMG_SUMMARY = os.path.join(DIR, "report_img", "hero_summary.png")
OUT_TXT = os.path.join(DIR, "report_text.txt")
OUT_JSON = os.path.join(DIR, "report_state.json")

FONT = "C:/Windows/Fonts/msyh.ttc"
FONTB = "C:/Windows/Fonts/msyhbd.ttc"

C_BG = (255, 255, 255)
C_TITLE = (17, 24, 39)
C_SUB = (110, 117, 128)
C_HEAD = (79, 70, 229)
C_HEADTX = (255, 255, 255)
C_ROW = (255, 255, 255)
C_ROWALT = (249, 250, 251)
C_LINE = (226, 229, 234)
C_TXT = (31, 41, 55)
C_GAPTX = (190, 18, 60)
C_GAPBG = (255, 241, 242)
C_OKTX = (21, 128, 61)
C_OKBG = (231, 250, 236)
C_STD = (146, 64, 14)
C_BAN100 = (13, 110, 104)
C_BAN50 = (176, 82, 10)

GAP_LIMIT = 3000     # 统一口径：当月签单实算业绩 Q>0 且 0 < 差距 < 该值（四个追踪共用）


def resolve_src():
    """源文件：命令行参数 > 环境变量 > _report/../src_latest.xlsx"""
    if len(sys.argv) > 1 and str(sys.argv[1]).strip():
        return os.path.abspath(sys.argv[1])
    env = os.environ.get("HERO_SRC", "").strip()
    if env:
        return os.path.abspath(env)
    return os.path.join(DIR, "..", "src_latest.xlsx")


SRC = resolve_src()


DATE_RE = re.compile(r"(20\d{2})[^\d]?(\d{2})[^\d]?(\d{2})")


def cycle_info(src_path):
    """从源文件名中的日期戳推导周期信息。
    日 >= 15 → 属「下月」周期；日 <= 14 → 属「当月」周期。

    文件名无日期戳（如 src_latest.xlsx）时，回退读
    _report/src_latest.meta.json 里记录的原始文件名（由 watch 落盘时写入）。
    """
    def _find(text):
        return list(DATE_RE.finditer(os.path.basename(str(text))))

    m = _find(src_path)
    if not m:
        meta = os.path.join(DIR, "src_latest.meta.json")
        if os.path.exists(meta):
            try:
                nm = json.load(open(meta, encoding="utf-8")).get("src_name", "")
                m = _find(nm)
            except Exception:
                m = []
    if m:
        y, mo, da = (int(x) for x in m[-1].groups())
    else:
        t = datetime.now()
        y, mo, da = t.year, t.month, t.day
    data_date = f"{y}/{mo:02d}/{da:02d}"
    if da >= 15:
        sy, sm = y, mo
        pm = mo + 1
        if pm > 12:
            pm = 1
    else:
        pm = mo
        sy, sm = (y - 1, 12) if mo == 1 else (y, mo - 1)
    return {
        "cycle": f"{sy}/{sm:02d}/15-{data_date}",
        "data_date": data_date,
        "data_ymd": f"{y}{mo:02d}{da:02d}",
        "period_month": pm,
        "period_end": f"{pm}月14日",
        "period_half": f"{pm}月1日-{pm}月14日",
    }


CI = cycle_info(SRC)
CYCLE = CI["cycle"]


def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def money(v):
    if v is None:
        return "-"
    if abs(v - round(v)) < 1e-6:
        return f"{int(round(v)):,}"
    s = f"{v:,.2f}".rstrip("0").rstrip(".")
    return s


def score_txt(v):
    """差异值展示：整数不带小数，否则最多两位并去掉尾零"""
    if v is None:
        return "-"
    if abs(v - round(v)) < 5e-3:
        return f"{int(round(v)):,}"
    return f"{v:,.1f}"


def load_rows():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["企服顾问明细"]
    out = []
    for r in range(4, ws.max_row + 1):
        dept = ws.cell(r, 9).value
        if dept and "" in str(dept):
            out.append({
                "no": ws.cell(r, 1).value,
                "name": ws.cell(r, 2).value,
                "dept": str(dept).replace("", ""),
                "Q": fnum(ws.cell(r, 17).value),
                "R": fnum(ws.cell(r, 18).value),
                "S": fnum(ws.cell(r, 19).value),
                "T": str(ws.cell(r, 20).value).strip(),
                "Y": fnum(ws.cell(r, 25).value),
                "Z": str(ws.cell(r, 26).value).strip(),
                "AF": fnum(ws.cell(r, 32).value),
                "AG": str(ws.cell(r, 33).value).strip(),
                "AN": str(ws.cell(r, 40).value).strip(),
                "AO": fnum(ws.cell(r, 41).value),
                "W": str(ws.cell(r, 23).value).strip(),
                "AZ": str(ws.cell(r, 52).value).strip(),
            })
    return out


def has_q(p):
    """统一口径：当月签单实算业绩 Q > 0（业绩为 0 或负数一律不进榜）"""
    return p["Q"] is not None and p["Q"] > 0


def gap_of(std, p):
    """差距 = 标准 − 当月签单实算利润业绩；任一为空则返回 None"""
    if std is None or p["R"] is None:
        return None
    return std - p["R"]


def analyze():
    rows = load_rows()
    d100 = [p for p in rows if p["Z"] == "是"]
    d50 = [p for p in rows if p["AG"] == "是"]
    done = {p["name"] for p in d100} | {p["name"] for p in d50}
    d50_only = [p for p in d50 if p["name"] not in {x["name"] for x in d100}]

    g100, g50 = [], []
    for p in rows:
        if p["name"] in done or not has_q(p):
            continue
        g = gap_of(p["Y"], p)
        if g is not None and 0 < g < GAP_LIMIT:
            g100.append((g, p))
        g = gap_of(p["AF"], p)
        if g is not None and 0 < g < GAP_LIMIT:
            g50.append((g, p))
    g100.sort(key=lambda x: x[0])
    g50.sort(key=lambda x: x[0])

    # 破零人力：未达成破零 且 业绩>0 且 0 < 差距 < GAP_LIMIT
    bzero = []
    for p in rows:
        if p["T"] == "是" or not has_q(p):
            continue
        g = gap_of(p["S"], p)
        if g is not None and 0 < g < GAP_LIMIT:
            bzero.append((g, p))
    bzero.sort(key=lambda x: x[0])

    # 有效人力：已达成 = AN == "是"；差距榜 = AN == "否" 且 业绩>0 且 0 < 差距 < GAP_LIMIT
    vdone = [p for p in rows if p["AN"] == "是"]
    vgap = []
    for p in rows:
        if p["AN"] != "否" or not has_q(p):
            continue
        g = gap_of(p["AO"], p)
        if g is not None and 0 < g < GAP_LIMIT:
            vgap.append((g, p))
    vgap.sort(key=lambda x: x[0])

    return d100, d50_only, g100, g50, bzero, vdone, vgap


# ---------------- 绘图 ----------------
def draw_head(d, x, y, w, cols, head_h=62):
    """画表头，返回 xs 累计宽度"""
    ft = ImageFont.truetype(FONTB, 23)
    inner = w
    xs, acc = [0], 0
    for _, ratio, _, _ in cols:
        acc += inner * ratio
        xs.append(acc)
    d.rectangle([x, y, x + inner, y + head_h], fill=C_HEAD)
    for i, (label, _, align, _) in enumerate(cols):
        cx0, cx1 = x + xs[i], x + xs[i + 1]
        lines = label.split("\n")
        ny = y + (head_h - len(lines) * 28) / 2 + 1
        for ln in lines:
            tw = d.textlength(ln, font=ft)
            if align == "r":
                d.text((cx1 - 12 - tw, ny), ln, font=ft, fill=C_HEADTX)
            elif align == "c":
                d.text((cx0 + (cx1 - cx0 - tw) / 2, ny), ln, font=ft, fill=C_HEADTX)
            else:
                d.text((cx0 + 12, ny), ln, font=ft, fill=C_HEADTX)
            ny += 28
    for i in range(1, len(cols)):
        d.line([x + xs[i], y, x + xs[i], y + head_h], fill=(129, 122, 235), width=1)
    return xs


def draw_grid(d, x, y, w, cols, data_rows, row_h=50, highlight_last=False):
    """data_rows: list of list of str（列数与 cols 一致）；返回底边 y。
    highlight_last=True 时最后一行（合计行）用表头紫底白字突出。"""
    head_h = 62
    xs = draw_head(d, x, y, w, cols)
    ftn = ImageFont.truetype(FONT, 23)
    ftb = ImageFont.truetype(FONTB, 23)
    ry = y + head_h
    for idx, vals in enumerate(data_rows):
        last = highlight_last and idx == len(data_rows) - 1
        if last:
            base = C_HEAD
        else:
            base = C_ROWALT if idx % 2 else C_ROW
        d.rectangle([x, ry, x + w, ry + row_h], fill=base)
        for ci, (label, ratio, align, kind) in enumerate(cols):
            cx0, cx1 = x + xs[ci], x + xs[ci + 1]
            txt = str(vals[ci])
            if last:
                fnt, col = ftb, C_HEADTX
            elif kind == "gap":
                d.rectangle([cx0, ry, cx1, ry + row_h], fill=C_GAPBG)
                fnt, col = ftb, C_GAPTX
            elif kind == "badge":
                d.rectangle([cx0, ry, cx1, ry + row_h], fill=C_OKBG)
                fnt, col = ftb, C_OKTX
            elif kind == "std":
                fnt, col = ftn, C_STD
            elif kind == "name":
                fnt, col = ftb, C_TXT
            else:
                fnt, col = ftn, C_TXT
            lines = str(txt).split("\n")
            lh = 26
            th = len(lines) * lh
            y0 = ry + (row_h - th) / 2 + 2
            for ln in lines:
                tw = d.textlength(ln, font=fnt)
                if align == "r":
                    d.text((cx1 - 12 - tw, y0), ln, font=fnt, fill=col)
                elif align == "c":
                    d.text((cx0 + (cx1 - cx0 - tw) / 2, y0), ln, font=fnt, fill=col)
                else:
                    d.text((cx0 + 12, y0), ln, font=fnt, fill=col)
                y0 += lh
        d.line([x, ry + row_h, x + w, ry + row_h], fill=C_LINE, width=1)
        ry += row_h
    d.rectangle([x, y, x + w, ry], outline=C_LINE, width=1)
    for i in range(1, len(cols)):
        d.line([x + xs[i], y, x + xs[i], ry], fill=C_LINE, width=1)
    return ry


def banner(d, x, y, w, text, color, note=""):
    h = 52
    ft = ImageFont.truetype(FONTB, 25)
    d.rounded_rectangle([x, y, x + w, y + h], radius=10, fill=color)
    d.text((x + 18, y + 11), text, font=ft, fill=(255, 255, 255))
    if note:
        ftn = ImageFont.truetype(FONT, 21)
        tw = d.textlength(text, font=ft)
        d.text((x + 18 + tw + 18, y + 16), note, font=ftn, fill=(255, 255, 255))
    return y + h + 10


COLS_DONE = [
    ("工号", 0.10, "c", ""),
    ("姓名", 0.10, "c", "name"),
    ("部门", 0.22, "l", ""),
    ("当月签单\n实算业绩", 0.15, "r", ""),
    ("当月签单\n实算利润业绩", 0.17, "r", ""),
    ("英雄人力值标准", 0.14, "r", "std"),
    ("是否达成", 0.12, "c", "badge"),
]
COLS_GAP = [
    ("工号", 0.10, "c", ""),
    ("姓名", 0.10, "c", "name"),
    ("部门", 0.20, "l", ""),
    ("当月签单\n实算业绩", 0.15, "r", ""),
    ("当月签单\n实算利润业绩", 0.17, "r", ""),
    ("英雄人力值标准", 0.14, "r", "std"),
    ("差距", 0.14, "r", "gap"),
]
COLS_ZERO = [
    ("工号", 0.10, "c", ""),
    ("姓名", 0.10, "c", "name"),
    ("部门", 0.20, "l", ""),
    ("当月签单\n实算业绩", 0.15, "r", ""),
    ("当月签单\n实算利润业绩", 0.17, "r", ""),
    ("门槛破零标准值", 0.14, "r", "std"),
    ("差距", 0.14, "r", "gap"),
]
COLS_VALID = [
    ("工号", 0.10, "c", ""),
    ("姓名", 0.10, "c", "name"),
    ("部门", 0.20, "l", ""),
    ("当月签单\n实算业绩", 0.15, "r", ""),
    ("当月签单\n实算利润业绩", 0.17, "r", ""),
    ("实算利润\n业绩标准", 0.14, "r", "std"),
    ("差距", 0.14, "r", "gap"),
]
COLS_SUMMARY = [
    ("事业部", 0.22, "l", "name"),
    ("月初\n人数", 0.10, "c", ""),
    ("考核\n人数", 0.10, "c", ""),
    ("英雄\n人力", 0.145, "c", "badge"),
    ("标准\n人力", 0.145, "c", "badge"),
    ("破零\n人力", 0.145, "c", "badge"),
    ("有效\n人力", 0.145, "c", "badge"),
]


def render_done(d100, d50, now_str, out_path):
    W, P = 1500, 34
    inner = W - P * 2
    rows100 = [[str(p["no"]), p["name"], p["dept"], money(p["Q"]), money(p["R"]),
                money(p["Y"]), "已达成"] for p in d100]
    rows50 = [[str(p["no"]), p["name"], p["dept"], money(p["Q"]), money(p["R"]),
               money(p["AF"]), "已达成"] for p in d50]

    h = 150 + 62 + 1 * 50 + 20                               # 100% 区
    h += 62 + 62 + max(1, len(rows50)) * 50 + 20             # 50% 区
    h += 46 + P
    img = Image.new("RGB", (W, h), C_BG)
    d = ImageDraw.Draw(img)

    ft_t = ImageFont.truetype(FONTB, 40)
    ft_s = ImageFont.truetype(FONT, 22)
    d.text((P, 26), "分公司 · 英雄人力目标完成情况", font=ft_t, fill=C_TITLE)
    d.text((P, 80), f"数据周期：{CYCLE}（截至 24 点）　｜　口径：部门含「」　｜　制图 {now_str}",
           font=ft_s, fill=C_SUB)
    d.text((P, 110), "达成英雄人力目标(全月) → 本月全月活动量免考核；达成 50% → 下半月免考核",
           font=ft_s, fill=C_SUB)

    y = 150
    y = banner(d, P, y, inner, "① 达成英雄人力目标（全月）",
               C_BAN100, f"{len(d100)} 人 · 本月全月免考核")
    y = draw_grid(d, P, y, inner, COLS_DONE, rows100) if rows100 else y
    y += 26

    y = banner(d, P, y, inner, "② 达成英雄人力目标（50%）",
               C_BAN50, f"{len(d50)} 人 · 下半月（{CI['period_half']}）免考核")
    y = draw_grid(d, P, y, inner, COLS_DONE, rows50) if rows50 else y
    y += 40

    if not rows100:
        d.text((P, y - 10), "（本期暂无达成全月英雄人力目标的伙伴）",
               font=ft_s, fill=C_SUB)
    img = img.crop((0, 0, W, min(h, y + 10)))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    img.save(out_path, "PNG", optimize=True)
    return img.size


def render_gap(g50, now_str, out_path):
    W, P = 1500, 34
    inner = W - P * 2
    rows = [[str(p["no"]), p["name"], p["dept"], money(p["Q"]), money(p["R"]),
             money(p["AF"]), score_txt(g)] for g, p in g50]

    ft_t = ImageFont.truetype(FONTB, 40)
    ft_s = ImageFont.truetype(FONT, 22)
    h = 150 + 62 + max(1, len(rows)) * 50 + 46 + P
    img = Image.new("RGB", (W, h), C_BG)
    d = ImageDraw.Draw(img)
    d.text((P, 26), "距英雄人力目标（50%）差距较小的人员", font=ft_t, fill=C_TITLE)
    d.text((P, 80), f"数据周期：{CYCLE}（截至 24 点）　｜　共 {len(rows)} 人　｜　"
                    f"差距 = 英雄人力值标准 − 当月签单实算利润业绩（< {GAP_LIMIT} 元）",
           font=ft_s, fill=C_SUB)
    d.text((P, 110), "统计口径：当月签单实算业绩 > 0 且尚未达成，按差距升序排列（越靠前越接近达标）",
           font=ft_s, fill=C_SUB)
    y = draw_grid(d, P, 150, inner, COLS_GAP, rows) if rows else 150
    img = img.crop((0, 0, W, min(h, y + P)))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    img.save(out_path, "PNG", optimize=True)
    return img.size


def render_zero(bzero, now_str, out_path):
    W, P = 1500, 34
    inner = W - P * 2
    rows = [[str(p["no"]), p["name"], p["dept"], money(p["Q"]), money(p["R"]),
             money(p["S"]), score_txt(g)] for g, p in bzero]

    ft_t = ImageFont.truetype(FONTB, 40)
    ft_s = ImageFont.truetype(FONT, 22)
    h = 150 + 62 + max(1, len(rows)) * 50 + 46 + P
    img = Image.new("RGB", (W, h), C_BG)
    d = ImageDraw.Draw(img)
    d.text((P, 26), "破零人力 · 重点催办名单", font=ft_t, fill=C_TITLE)
    d.text((P, 80), f"数据周期：{CYCLE}（截至 24 点）　｜　共 {len(rows)} 人　｜　"
                    f"差距 = 门槛破零标准值 − 当月签单实算利润业绩（< {GAP_LIMIT} 元）",
           font=ft_s, fill=C_SUB)
    d.text((P, 110), "统计口径：当月签单实算业绩 > 0 且尚未破零，按差距升序排列（越靠前越接近破零）",
           font=ft_s, fill=C_SUB)
    y = draw_grid(d, P, 150, inner, COLS_ZERO, rows) if rows else 150
    img = img.crop((0, 0, W, min(h, y + P)))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    img.save(out_path, "PNG", optimize=True)
    return img.size


def render_valid(vgap, now_str, out_path):
    W, P = 1500, 34
    inner = W - P * 2
    rows = [[str(p["no"]), p["name"], p["dept"], money(p["Q"]), money(p["R"]),
             money(p["AO"]), score_txt(g)] for g, p in vgap]

    ft_t = ImageFont.truetype(FONTB, 40)
    ft_s = ImageFont.truetype(FONT, 22)
    h = 150 + 62 + max(1, len(rows)) * 50 + 46 + P
    img = Image.new("RGB", (W, h), C_BG)
    d = ImageDraw.Draw(img)
    d.text((P, 26), "有效人力 · 差距追踪", font=ft_t, fill=C_TITLE)
    d.text((P, 80), f"数据周期：{CYCLE}（截至 24 点）　｜　共 {len(rows)} 人　｜　"
                    f"差距 = 实算利润业绩标准 − 当月签单实算利润业绩（< {GAP_LIMIT} 元）",
           font=ft_s, fill=C_SUB)
    d.text((P, 110), "统计口径：当月签单实算业绩 > 0 且尚未达成有效人力，按差距升序排列（越靠前越接近达标）",
           font=ft_s, fill=C_SUB)
    y = draw_grid(d, P, 150, inner, COLS_VALID, rows) if rows else 150
    img = img.crop((0, 0, W, min(h, y + P)))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    img.save(out_path, "PNG", optimize=True)
    return img.size


def _pct(rate, cnt=None, base=None):
    """率展示：优先用官方 sheet 给的率(小数)；缺失时用 cnt/base 兜底计算。"""
    if isinstance(rate, (int, float)):
        return "%.1f%%" % (rate * 100)
    if cnt is not None and base:
        return "%.1f%%" % (cnt / base * 100) if base else "-"
    return "-"


def load_summary_official():
    """读官方「事业部汇总」sheet 的各行（含「」合计行），人数口径与官方一致：
    月初企服顾问人数 = 三力（英雄/标准/破零）率的分母；考核人数 = 有效人力占比的分母。
    返回 (rows, total)；sheet 缺失 / 无行 / 表头对不上 时返回 None（调用方回退明细聚合）。"""
    try:
        wb = openpyxl.load_workbook(SRC, data_only=True)
        if "事业部汇总" not in wb.sheetnames:
            return None
        ws = wb["事业部汇总"]
        head = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(2, c).value
            if v is not None:
                head[str(v).strip()] = c
        need = ["事业部", "月初企服顾问人数", "门槛破零人力", "门槛破零人力率",
                "标准人力", "标准人力率", "英雄人力", "英雄人力率",
                "考核人数", "有效人力", "有效人力占比"]
        if any(k not in head for k in need):
            return None

        def num(col):
            v = ws.cell(ws_r, head[col]).value
            return v if isinstance(v, (int, float)) else None

        rows, total = [], None
        for ws_r in range(3, ws.max_row + 1):
            name = str(ws.cell(ws_r, head["事业部"]).value or "").strip()
            if not name.startswith(""):
                continue
            item = {
                "dept": "分公司" if name == "" else name[2:],
                "month0": num("月初企服顾问人数"),
                "assessed": num("考核人数"),
                "h100": num("英雄人力"), "h100_rate": num("英雄人力率"),
                "std": num("标准人力"), "std_rate": num("标准人力率"),
                "zero": num("门槛破零人力"), "zero_rate": num("门槛破零人力率"),
                "valid": num("有效人力"), "valid_rate": num("有效人力占比"),
            }
            if name == "":
                total = item
            else:
                rows.append(item)
        if not rows:
            return None
        rows.sort(key=lambda x: -(x["month0"] or 0))
        if total is None:
            m0 = sum(x["month0"] or 0 for x in rows)
            ac = sum(x["assessed"] or 0 for x in rows)
            total = {"dept": "分公司", "month0": m0, "assessed": ac,
                     "h100": sum(x["h100"] or 0 for x in rows),
                     "std": sum(x["std"] or 0 for x in rows),
                     "zero": sum(x["zero"] or 0 for x in rows),
                     "valid": sum(x["valid"] or 0 for x in rows)}
            total["h100_rate"] = total["h100"] / m0 if m0 else None
            total["std_rate"] = total["std"] / m0 if m0 else None
            total["zero_rate"] = total["zero"] / m0 if m0 else None
            total["valid_rate"] = total["valid"] / ac if ac else None
        return rows, total
    except Exception:
        return None


def division_summary():
    """事业部三力汇总：优先直读官方「事业部汇总」sheet（人数/率与官方完全一致）；
    sheet 不可用时回退按企服顾问明细聚合（月初人数≈明细行数、考核人数=AZ(是否考核有效人力)=是）。"""
    off = load_summary_official()
    if off:
        return off[0] + [off[1]], "官方事业部汇总sheet"
    rows = load_rows()
    divs = {}
    for p in rows:
        d = (p["dept"] or "未知").strip() or "未知"
        e = divs.setdefault(d, {"total": 0, "h100": 0, "std": 0, "zero": 0,
                                "valid": 0, "assessed": 0})
        e["total"] += 1
        e["h100"] += p["Z"] == "是"
        e["std"] += str(p.get("W", "")) == "是"
        e["zero"] += p["T"] == "是"
        e["valid"] += p["AN"] == "是"
        e["assessed"] += str(p.get("AZ", "")) == "是"
    items = sorted(divs.items(), key=lambda kv: (-kv[1]["total"], kv[0]))
    summary = []
    for k, v in items:
        summary.append({"dept": k, "month0": v["total"], "assessed": v["assessed"],
                        "h100": v["h100"], "h100_rate": v["h100"] / v["total"] if v["total"] else None,
                        "std": v["std"], "std_rate": v["std"] / v["total"] if v["total"] else None,
                        "zero": v["zero"], "zero_rate": v["zero"] / v["total"] if v["total"] else None,
                        "valid": v["valid"], "valid_rate": v["valid"] / v["assessed"] if v["assessed"] else None})
    m0 = sum(v["total"] for v in divs.values())
    ac = sum(v["assessed"] for v in divs.values())
    total = {"dept": "分公司", "month0": m0, "assessed": ac,
             "h100": sum(v["h100"] for v in divs.values()),
             "std": sum(v["std"] for v in divs.values()),
             "zero": sum(v["zero"] for v in divs.values()),
             "valid": sum(v["valid"] for v in divs.values()),
             "h100_rate": sum(v["h100"] for v in divs.values()) / m0 if m0 else None,
             "std_rate": sum(v["std"] for v in divs.values()) / m0 if m0 else None,
             "zero_rate": sum(v["zero"] for v in divs.values()) / m0 if m0 else None,
             "valid_rate": sum(v["valid"] for v in divs.values()) / ac if ac else None}
    summary.append(total)
    return summary, "明细聚合(回退)"


def render_summary(summary, now_str, out_path):
    """渲染「分公司各事业部三力汇总」表（末行分公司合计高亮）。
    人数口径与官方「事业部汇总」sheet 一致：月初人数 / 考核人数两列；
    每指标格「N人 / 率%」两行——三力(英雄/标准/破零)率÷月初人数，有效占比÷考核人数。"""
    W, P = 1500, 34
    inner = W - P * 2
    rows = []
    for s in summary:
        rows.append([
            s["dept"], str(s["month0"]), str(s["assessed"]),
            "%d人\n%s" % (s["h100"], _pct(s["h100_rate"], s["h100"], s["month0"])),
            "%d人\n%s" % (s["std"], _pct(s["std_rate"], s["std"], s["month0"])),
            "%d人\n%s" % (s["zero"], _pct(s["zero_rate"], s["zero"], s["month0"])),
            "%d人\n%s" % (s["valid"], _pct(s["valid_rate"], s["valid"], s["assessed"])),
        ])
    ft_t = ImageFont.truetype(FONTB, 40)
    ft_s = ImageFont.truetype(FONT, 22)
    h = 150 + 62 + len(rows) * 60 + 46 + P
    img = Image.new("RGB", (W, h), C_BG)
    d = ImageDraw.Draw(img)
    d.text((P, 26), "分公司各事业部 · 三力汇总", font=ft_t, fill=C_TITLE)
    d.text((P, 80), f"数据周期：{CYCLE}（截至 24 点）　｜　制图 {now_str}　｜　"
                    f"数据源：官方「事业部汇总」sheet",
           font=ft_s, fill=C_SUB)
    d.text((P, 110), "各事业部按月初人数降序；末行为分公司合计　｜　"
                     "英雄/标准/破零 率÷月初人数，有效占比÷考核人数",
           font=ft_s, fill=C_SUB)
    y = draw_grid(d, P, 150, inner, COLS_SUMMARY, rows, row_h=60, highlight_last=True)
    img = img.crop((0, 0, W, min(h, y + P)))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    img.save(out_path, "PNG", optimize=True)
    return img.size


def main():
    d100, d50, g100, g50, bzero, vdone, vgap = analyze()
    # 真实 @单人：把本期出现的姓名解析成 <@userid>（通讯录授权后生效，未授权回退纯文本）
    all_names = set()
    for p in d100 + d50:
        all_names.add(p["name"])
    for _, p in g100 + g50 + bzero + vgap:
        all_names.add(p["name"])
    for p in vdone:
        all_names.add(p["name"])
    mention_util.resolve_names(sorted(all_names))
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    s1 = render_done(d100, d50, now_str, IMG_DONE)
    s2 = render_gap(g50, now_str, IMG_GAP)
    s3 = render_zero(bzero, now_str, IMG_ZERO)
    s4 = render_valid(vgap, now_str, IMG_VALID)
    summary, sum_src = division_summary()
    s0 = render_summary(summary, now_str, IMG_SUMMARY)

    def pnum(v):
        """文案里的数字：不带千分位（对齐群里既有文案写法）"""
        if v is None:
            return "-"
        if abs(v - round(v)) < 5e-3:
            return str(int(round(v)))
        return f"{v:.1f}"

    t1, t2 = [], []

    # ---- 第一段：恭喜 ----
    if d100:
        t1.append("🎉🎉恭喜" + " ".join(mention_util.mention(p["name"]) for p in d100) +
                  f" 已达成英雄人力目标，今日起本月全月（今日-{CI['period_end']}）活动量免考核~")
    if d50:
        t1.append("🎉🎉恭喜" + " ".join(mention_util.mention(p["name"]) for p in d50) +
                  f" 已达成英雄人力目标50%，下半月（{CI['period_half']}）活动量免考核~")
    if not t1:
        t1.append("本期暂无伙伴达成英雄人力目标，大家加油！")

    # ---- 第二段：差距 ----
    if g100:
        seg = "；".join(f"{mention_util.mention(p['name'])} 距离英雄人力差距{pnum(g)}元" for g, p in g100)
        t2.append(seg + "，这几位伙伴即将达成英雄人力目标！")
    else:
        t2.append(f"距离英雄人力目标（全月）差距<{GAP_LIMIT}元：本期暂无")
    t2.append("")
    if g50:
        seg = "；".join(f"{mention_util.mention(p['name'])} 距离英雄人力50%差距{pnum(g)}元" for g, p in g50)
        t2.append(seg + "，这几位伙伴即将达成英雄人力50%！")
    else:
        t2.append(f"距离英雄人力目标（50%）差距<{GAP_LIMIT}元：本期暂无")

    # ---- 第三段：破零人力重点催办 ----
    t3 = []
    if bzero:
        seg = "；".join(f"{mention_util.mention(p['name'])} 距离破零差距{pnum(g)}元" for g, p in bzero)
        t3.append(seg + "，这几位伙伴即将达成破零，冲刺一把！")
    else:
        t3.append(f"破零人力重点催办（差距<{GAP_LIMIT}元）：本期暂无")

    # ---- 第四段：有效人力 ----
    t4 = []
    if vdone:
        t4.append("🎉🎉恭喜" + " ".join(mention_util.mention(p["name"]) for p in vdone) +
                  " 已达成有效人力目标~")
    else:
        t4.append("本期暂无伙伴达成有效人力目标。")
    t4.append("")
    if vgap:
        seg = "；".join(f"{mention_util.mention(p['name'])} 距离有效人力差距{pnum(g)}元" for g, p in vgap)
        t4.append(seg + "，这几位伙伴即将达成有效人力，加油！")
    else:
        t4.append(f"有效人力差距追踪（差距<{GAP_LIMIT}元）：本期暂无")

    text1 = "\n".join(t1) + "\n"
    text2 = "\n".join(t2) + "\n"
    text3 = "\n".join(t3) + "\n"
    text4 = "\n".join(t4) + "\n"
    total = summary[-1]
    text0 = ("📊 分公司各事业部三力汇总（数据时间 %s）\n" % now_str +
             "英雄人力达成 %d 人(%s) / 标准人力达成 %d 人(%s) / 破零达成 %d 人(%s) / 有效人力达成 %d 人(%s)"
             "（分公司：月初 %s 人 · 考核 %s 人，详见下图）\n"
             % (total["h100"], _pct(total["h100_rate"], total["h100"], total["month0"]),
                total["std"], _pct(total["std_rate"], total["std"], total["month0"]),
                total["zero"], _pct(total["zero_rate"], total["zero"], total["month0"]),
                total["valid"], _pct(total["valid_rate"], total["valid"], total["assessed"]),
                total["month0"], total["assessed"]))
    with open(OUT_TXT, "w", encoding="utf-8") as f:
        f.write(text0 + "\n" + text1 + "\n" + text2 + "\n" + text3 + "\n" + text4)
    with open(os.path.join(DIR, "report_text_summary.txt"), "w", encoding="utf-8") as f:
        f.write(text0)
    with open(os.path.join(DIR, "report_text1.txt"), "w", encoding="utf-8") as f:
        f.write(text1)
    with open(os.path.join(DIR, "report_text2.txt"), "w", encoding="utf-8") as f:
        f.write(text2)
    with open(os.path.join(DIR, "report_text3.txt"), "w", encoding="utf-8") as f:
        f.write(text3)
    with open(os.path.join(DIR, "report_text4.txt"), "w", encoding="utf-8") as f:
        f.write(text4)
    text = text0 + "\n" + text1 + "\n" + text2 + "\n" + text3 + "\n" + text4

    state = {
        "generated_at": now_str,
        "source": os.path.basename(SRC),
        "cycle": CYCLE,
        "period_month": CI["period_month"],
        "gap_limit": GAP_LIMIT,
        "done100": [{"no": p["no"], "name": p["name"], "dept": p["dept"], "Q": p["Q"],
                     "R": p["R"], "std": p["Y"]} for p in d100],
        "done50": [{"no": p["no"], "name": p["name"], "dept": p["dept"], "Q": p["Q"],
                    "R": p["R"], "std": p["AF"]} for p in d50],
        "gap100": [{"name": p["name"], "gap": round(g, 2)} for g, p in g100],
        "gap50": [{"no": p["no"], "name": p["name"], "dept": p["dept"], "Q": p["Q"],
                   "R": p["R"], "std": p["AF"], "gap": round(g, 2)} for g, p in g50],
        "breakzero": [{"no": p["no"], "name": p["name"], "dept": p["dept"], "Q": p["Q"],
                       "R": p["R"], "std": p["S"], "gap": round(g, 2)} for g, p in bzero],
        "valid_done": [{"no": p["no"], "name": p["name"], "dept": p["dept"],
                        "Q": p["Q"], "R": p["R"], "std": p["AO"]} for p in vdone],
        "valid_gap": [{"no": p["no"], "name": p["name"], "dept": p["dept"], "Q": p["Q"],
                       "R": p["R"], "std": p["AO"], "gap": round(g, 2)} for g, p in vgap],
        "image_done": {"path": IMG_DONE, "w": s1[0], "h": s1[1]},
        "image_gap": {"path": IMG_GAP, "w": s2[0], "h": s2[1]},
        "image_zero": {"path": IMG_ZERO, "w": s3[0], "h": s3[1]},
        "image_valid": {"path": IMG_VALID, "w": s4[0], "h": s4[1]},
        "image_summary": {"path": IMG_SUMMARY, "w": s0[0], "h": s0[1]},
        "summary_source": sum_src,
        "summary": [{"dept": s["dept"], "month0": s["month0"], "assessed": s["assessed"],
                     "h100": s["h100"], "h100_rate": s["h100_rate"],
                     "std": s["std"], "std_rate": s["std_rate"],
                     "zero": s["zero"], "zero_rate": s["zero_rate"],
                     "valid": s["valid"], "valid_rate": s["valid_rate"]}
                    for s in summary],
    }
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

    print(f"SRC  : {SRC}")
    print(f"CYCLE: {CYCLE}  周期月={CI['period_month']}  免考核={CI['period_end']} / {CI['period_half']}")
    print(f"达成100%: {len(d100)} -> {[p['name'] for p in d100]}")
    print(f"达成50% : {len(d50)} -> {[p['name'] for p in d50]}")
    print(f"差距100 <{GAP_LIMIT}: {len(g100)} -> {[(p['name'], round(g,2)) for g,p in g100]}")
    print(f"差距50 <{GAP_LIMIT}: {len(g50)} -> {[(p['name'], round(g,2)) for g,p in g50]}")
    print(f"破零  <{GAP_LIMIT}: {len(bzero)} -> {[(p['name'], round(g,2)) for g,p in bzero]}")
    print(f"有效人力达成: {len(vdone)} -> {[p['name'] for p in vdone]}")
    print(f"有效人力差距<{GAP_LIMIT}: {len(vgap)} -> {[(p['name'], round(g,2)) for g,p in vgap]}")
    print(f"IMG0 {s0}  IMG1 {s1}  IMG2 {s2}  IMG3 {s3}  IMG4 {s4}")
    print("---- 事业部汇总（口径：%s）----" % sum_src)
    for s in summary:
        print(f"  {s['dept']}: 月初{s['month0']} 考核{s['assessed']} "
              f"英雄{s['h100']}({_pct(s['h100_rate'])}) 标准{s['std']}({_pct(s['std_rate'])}) "
              f"破零{s['zero']}({_pct(s['zero_rate'])}) 有效{s['valid']}({_pct(s['valid_rate'])})")
    print("---- 文字 ----")
    print(text)


if __name__ == "__main__":
    main()
