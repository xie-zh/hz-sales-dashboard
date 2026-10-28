# -*- coding: utf-8 -*-
"""活动量汇报（新版数据源 · 日档）—— 渲染 4 张图（只统计 · 网销 / 电邀 / 综法）

数据源：WorkBuddy 资料库发布页「活动量工作台 · 七岗排行榜」
        https://dashboard.example.com/p/LLT6AK0VUUvnPu12Cc8PKZ   （取数见 wb_source.py）

⚠️ 口径是页面顶部的【日】档（单日），不是【全周期】累计：
   页面用内嵌引擎 RBE 按选定日期区间重算，所以这里走 `wb_source.fetch_day()`
   （调同一套引擎，逐格与页面一致）。

四张图：
  report_img/act1_overview_{track|result}.png   ① 整体完成情况（ 三岗 + 合计）
  report_img/act2_wangxiao_{track|result}.png   ② 网销岗（部门 + 人员，两级）
  report_img/act3_dianxiao_{track|result}.png   ③ 电邀岗（部门 + 人员，两级）
  report_img/act4_zongfa_{track|result}.png     ④ 综法岗（部门 + 人员，两级）

口径（2026-10-09 用户确认，仅 + 仅这三岗 + 日档；2026-10-10 增补电邀达标口径）：
  · 只取（部门名含「」）
  · 只统计 网销岗 / 电邀岗 / 综法岗
  · 数据周期 = **严格「当天」（today）**：只取今天这一天的记录，历史与其他日期的数据一律排除
    （2026-10-10 用户口径）。实现：`render_all()` 调 `wb_source.fetch_day(day=<今天YYYYMMDD>)`，
    三岗统一用当天，不再按岗位各自锚点回溯。
    ⚠️ 若某岗当天没有考核记录（典型：综法岗经常几天才有一条有效考核记录），该岗出一张
    「当日无考核数据」的图，而不是跳过——保证 4 张图固定产出。
  · 达标口径：
      网销岗  24:00 达标率 ≥ 100%（源站口径）
      电邀岗  6 项中「任意一项」达标（≥1）即算达标
              （2026-10-10 用户口径；网页默认是「6 项全中」）
      综法岗  19:30 达标率 ≥ 100%（源站口径）
  · track 只列未达标人员；result 列全部考核人员

模式：track（追踪）/ result（结果）
"""
import io
import json
import os
import sys
from datetime import datetime

from PIL import Image, ImageDraw, ImageFont

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
BASE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(BASE, 'report_img')
sys.path.insert(0, BASE)
import wb_source  # noqa: E402

FONT = r'C:\Windows\Fonts\msyh.ttc'
FONTB = r'C:\Windows\Fonts\msyhbd.ttc'

C_TEXT, C_MUTED, C_BORDER = '#1a1a2e', '#6b7280', '#e5e7eb'
C_ACCENT, C_PAGE, C_SUB = '#4f46e5', '#f4f6f9', '#8b93a1'
C_DEPT_BG, C_ROW, C_ROWALT = '#eef2ff', '#ffffff', '#fafbff'
C_PASS_BG, C_PASS_FG = '#dcfce7', '#166534'
C_FAIL_BG, C_FAIL_FG = '#fee2e2', '#991b1b'
C_NA_BG, C_NA_FG = '#f3f4f6', '#6b7280'
C_TOTAL_BG, C_TOTAL_FG = '#4f46e5', '#ffffff'
GREEN, AMBER, RED = '#16a34a', '#f59e0b', '#dc2626'

FS_TITLE, FS_SUB, FS_TH, FS_TD, FS_FOOT, FS_BADGE = 30, 17, 18, 18, 15, 15
ROW_H, TH_H, PAD, RADIUS, W = 40, 46, 34, 12, 1500

FT_TITLE = ImageFont.truetype(FONTB, FS_TITLE)
FT_SUB = ImageFont.truetype(FONT, FS_SUB)
FT_TH = ImageFont.truetype(FONTB, FS_TH)
FT_TD = ImageFont.truetype(FONT, FS_TD)
FT_TDB = ImageFont.truetype(FONTB, FS_TD)
FT_FOOT = ImageFont.truetype(FONT, FS_FOOT)
FT_BADGE = ImageFont.truetype(FONTB, FS_BADGE)

# 只统计这三岗（顺序即展示顺序）
POST_ORDER = ['网销岗', '电邀岗', '综法岗']
# 每岗配置：达标率字段 / 考核字段 / 在职字段 / 口径文案 / 单位
POST_CFG = {
    '网销岗': {'rate': 'r2400', 'kaohe': 'kaohe', 'onjob': 'onJob',
               'metric': '24:00 达标率', 'unit': '日均'},
    '电邀岗': {'rate': 'n6rate', 'kaohe': 'kaohe', 'onjob': 'onJob',
               'metric': '6 项任一达标', 'unit': '日均'},
    '综法岗': {'rate': 'r1930', 'kaohe': 'kaoheN', 'onjob': 'onjob',
               'metric': '19:30 达标率', 'unit': '人天'},
}
POST_FNAME = {'网销岗': ('act2_wangxiao', '网销岗'),
              '电邀岗': ('act3_dianxiao', '电邀岗（电销）'),
              '综法岗': ('act4_zongfa', '综法岗')}
# 未达标名单用到的人数占比字段（用于图1「达标」列）
POST_PASSFIELD = {'网销岗': 'p2400', '电邀岗': 'n6pass', '综法岗': 'p1930'}
N6_KEYS = ('n6wx', 'n6jzds', 'n6yj', 'n6jk', 'n6yx', 'n6jc')
# 6 项中文名（与页面红框列一致）：加微 / 精准审核 / 业绩 / 见客 / 意向 / 45s
N6_LABEL = {'n6wx': '加微', 'n6jzds': '精准审核', 'n6yj': '业绩',
            'n6jk': '见客', 'n6yx': '意向', 'n6jc': '45s'}
N6_MIN = 1          # 用户 2026-10-10 口径：任意一项达标即算达标


# ---------------- 格式化 ----------------
def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def f1(v):
    v = num(v)
    return '--' if v is None else ('%.1f' % v)


def fpct(v):
    v = num(v)
    return '--' if v is None else ('%.1f%%' % v)


def fmoney(v):
    v = num(v)
    return '--' if v is None else format(int(round(v)), ',')


def rate_color(v):
    v = num(v)
    if v is None:
        return C_MUTED
    if v >= 100:
        return GREEN
    return AMBER if v >= 50 else RED


def clean_name(s):
    if not s:
        return ''
    return str(s).split('（')[0].strip()


def clean_dept(s):
    """'/法律咨询三部' -> '法律咨询三部'"""
    s = str(s or '')
    if '/' in s:
        s = s.split('/')[-1]
    return s.strip()


def d2(s):
    """20261009 -> 10/09"""
    s = str(s or '')
    return '%s/%s' % (s[4:6], s[6:8]) if len(s) == 8 else s


def cn_date(s):
    """20261009 -> 2026-10-09"""
    s = str(s or '')
    return '%s-%s-%s' % (s[0:4], s[4:6], s[6:8]) if len(s) == 8 else s


def today_str():
    return datetime.now().strftime('%Y-%m-%d')


def today_ymd():
    """20261010 —— 用于强制「当天」取数。"""
    return datetime.now().strftime('%Y%m%d')


# ---------------- 绘图原语 ----------------
def draw_badge(d, cx, cy, text, kind='na'):
    fill, fg = {'pass': (C_PASS_BG, C_PASS_FG),
                'fail': (C_FAIL_BG, C_FAIL_FG),
                'na': (C_NA_BG, C_NA_FG)}[kind]
    tw = d.textlength(text, font=FT_BADGE)
    bw, bh = tw + 20, 26
    d.rounded_rectangle([cx - bw / 2, cy - bh / 2, cx + bw / 2, cy + bh / 2],
                        radius=9, fill=fill)
    d.text((cx, cy), text, font=FT_BADGE, fill=fg, anchor='mm')


def draw_cell(d, x, cy, cw, cell, bold_default=False):
    if isinstance(cell, str):
        cell = {'t': cell}
    t = cell.get('t', '')
    bold = cell.get('b', bold_default)
    f = FT_TDB if bold else FT_TD
    if cell.get('badge'):
        draw_badge(d, x + cw / 2, cy, t, cell.get('badge'))
        return
    col = cell.get('c') or C_TEXT
    align = cell.get('a', 'r')
    ind = cell.get('indent', 0)
    bar = cell.get('bar')
    ty = cy - 6 if bar is not None else cy
    if align == 'l':
        ax, anc = x + 14 + ind, 'lm'
    elif align == 'c':
        ax, anc = x + cw / 2, 'mm'
    else:
        ax, anc = x + cw - 14, 'rm'
    if bar is not None:
        v = num(bar)
        pad = 14
        bx0, bx1 = x + pad, x + cw - pad
        d.rounded_rectangle([bx0, cy + 10, bx1, cy + 16], radius=3, fill='#eceef3')
        if v is not None:
            bw = max(3, min(v, 100) / 100.0 * (bx1 - bx0))
            d.rounded_rectangle([bx0, cy + 10, bx0 + bw, cy + 16], radius=3,
                                fill=rate_color(v))
    d.text((ax, ty), t, font=f, fill=col, anchor=anc)


def draw_table(d, x0, y0, w, cols, rows):
    d.rectangle([x0, y0, x0 + w, y0 + TH_H], fill=C_ACCENT)
    cx = x0
    for (label, ratio, align) in cols:
        cw = w * ratio
        ax = cx + 14 if align == 'l' else (cx + cw - 14 if align == 'r' else cx + cw / 2)
        d.text((ax, y0 + TH_H / 2), label, font=FT_TH, fill='#ffffff',
               anchor={'l': 'lm', 'r': 'rm', 'c': 'mm'}[align])
        cx += cw
    y = y0 + TH_H
    ri = 0
    for row in rows:
        kind = row.get('kind', 'person')
        h = row.get('h', ROW_H)
        if kind == 'dept':
            bg = C_DEPT_BG
        elif kind == 'total':
            bg = C_TOTAL_BG
        elif kind == 'empty':
            bg = C_ROW
        else:
            bg = C_ROW if ri % 2 == 0 else C_ROWALT
            ri += 1
        d.rectangle([x0, y, x0 + w, y + h], fill=bg)
        d.line([x0, y, x0 + w, y], fill=C_BORDER, width=1)
        cx = x0
        cy = y + h / 2
        for i, (label, ratio, align) in enumerate(cols):
            cw = w * ratio
            cell = row['cells'][i] if i < len(row.get('cells', [])) else ''
            if kind == 'total' and not isinstance(cell, dict):
                cell = {'t': cell, 'c': C_TOTAL_FG}
            elif kind == 'total':
                cell = dict(cell)
                cell.setdefault('c', C_TOTAL_FG)
            draw_cell(d, cx, cy, cw, cell, bold_default=(kind in ('dept', 'total')))
            cx += cw
        d.line([x0, y + h, x0 + w, y + h], fill=C_BORDER, width=1)
        y += h
    return y


def panel(title, subtitle, note, cols, rows, out_path, foot):
    H = PAD * 2 + 78 + TH_H + sum(r.get('h', ROW_H) for r in rows) + 30 + 22 * len(note) + 34
    img = Image.new('RGB', (W, H), C_PAGE)
    d = ImageDraw.Draw(img)
    inner = W - PAD * 2
    d.rounded_rectangle([PAD, PAD, W - PAD, H - PAD], radius=RADIUS, fill='#ffffff',
                        outline=C_BORDER, width=1)
    d.text((PAD + 26, PAD + 34), title, font=FT_TITLE, fill=C_TEXT, anchor='lm')
    d.text((PAD + 26, PAD + 68), subtitle, font=FT_SUB, fill=C_SUB, anchor='lm')
    y = PAD + 96
    y = draw_table(d, PAD + 20, y, inner - 40, cols, rows)
    yy = y + 14
    for line in note:
        d.text((PAD + 26, yy), line, font=FT_FOOT, fill=C_MUTED, anchor='lm')
        yy += 22
    d.text((W - PAD - 26, H - PAD - 18), foot, font=FT_FOOT, fill=C_SUB, anchor='rm')
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    img.save(out_path, 'PNG', optimize=True)
    print('已生成 %s  %dx%d  行数=%d' % (out_path, img.size[0], img.size[1], len(rows)))
    return out_path


# ---------------- 口径 ----------------
def kaohe_of(post, c):
    return num(c.get(POST_CFG.get(post, {}).get('kaohe', 'kaohe'))) or 0.0


def rate_of(post, c):
    return num(c.get(POST_CFG.get(post, {}).get('rate', 'r2400')))


def n6_count(c):
    """已达标的项数（0~6）。"""
    return sum(1 for k in N6_KEYS if (num(c.get(k)) or 0) >= 1)


def dx_any(v, dept):
    """电邀岗按用户口径重算部门级的「达标人数 / 考核人数 / 达标率」。

    用户口径（2026-10-10）：6 项中任意一项达标即算达标（网页默认 6 项全中）。
    返回 (达标人数, 考核人数, 达标率)；无人员明细时返回 None（调用方回退引擎值）。
    """
    ppl = (v.get('people') or {}).get(dept)
    if not ppl:
        return None
    graded = [p for p in ppl if kaohe_of('电邀岗', p['cells']) > 0]
    if not graded:
        return (0, 0, None)
    n = sum(1 for p in graded if n6_count(p['cells']) >= 1)
    return (n, len(graded), n / len(graded) * 100)


def person_status(post, c, n6=None):
    """返回 ('pass'|'fail'|'na', 说明文本)。"""
    if post == '电邀岗':
        n6 = n6_count(c) if n6 is None else n6
        if kaohe_of(post, c) <= 0:
            return 'na', '%d/6' % n6
        return ('pass' if n6 >= N6_MIN else 'fail'), '%d/6' % n6
    rate = rate_of(post, c)
    if kaohe_of(post, c) <= 0 and rate is None:
        return 'na', '--'
    if rate is None:
        return 'na', '--'
    return ('pass' if rate >= 100 else 'fail'), fpct(rate)


STATUS_BADGE = {'pass': ('达标', 'pass'), 'fail': ('未达标', 'fail'), 'na': ('未考核', 'na')}


def post_totals(post, v):
    g = lambda c, k: num(c.get(k)) or 0.0
    onjob = kaohe = p = perf = 0.0
    for dept, r in v['depts'].items():
        c = r['cells']
        onjob += g(c, POST_CFG[post]['onjob'])
        kaohe += g(c, POST_CFG[post]['kaohe'])
        perf += g(c, 'dayPerf') + g(c, 'payAll')
        if post == '电邀岗':
            st = dx_any(v, dept)                 # 用户口径：6 项任一项达标
            p += st[0] if st else g(c, POST_PASSFIELD[post])
        else:
            p += g(c, POST_PASSFIELD[post])
    return {'onjob': onjob, 'kaohe': kaohe, 'pass': p, 'perf': perf}


# ---------------- 面板 ----------------
def build_overview(data, mode):
    cols = [('岗位', 0.11, 'l'), ('数据日期', 0.11, 'c'), ('在职', 0.08, 'r'),
            ('考核', 0.08, 'r'), ('达标', 0.08, 'r'), ('达标率', 0.15, 'r'),
            ('达标口径', 0.15, 'l'), ('当日业绩', 0.14, 'r'), ('人均业绩', 0.10, 'r')]
    anchors = data.get('anchor') or {}
    rows = []
    agg = {'onjob': 0.0, 'kaohe': 0.0, 'pass': 0.0, 'perf': 0.0}
    for post in POST_ORDER:
        v = data['hz'].get(post)
        if not v:
            # 当天无该岗数据：出一行占位，保持三岗结构完整
            rows.append({'kind': 'person', 'cells': [
                {'t': post, 'a': 'l', 'b': True},
                {'t': d2(anchors.get(post)) or '--', 'a': 'c', 'c': C_SUB},
                '--', '--', '--', {'t': '当日无数据', 'c': C_MUTED},
                POST_CFG[post]['metric'], '--', '--']})
            continue
        s = post_totals(post, v)
        rate = (s['pass'] / s['kaohe'] * 100) if s['kaohe'] else None
        peravg = (s['perf'] / s['kaohe']) if s['kaohe'] else None
        rows.append({'kind': 'person', 'cells': [
            {'t': post, 'a': 'l', 'b': True},
            {'t': d2(anchors.get(post)) or '--', 'a': 'c', 'c': C_SUB},
            f1(s['onjob']), f1(s['kaohe']), f1(s['pass']),
            {'t': fpct(rate), 'c': rate_color(rate), 'bar': rate},
            POST_CFG[post]['metric'], fmoney(s['perf']), fmoney(peravg)]})
        for k in agg:              # 日档三岗单位一致（人），可合计
            agg[k] += s[k]
    if agg['kaohe']:
        rate = agg['pass'] / agg['kaohe'] * 100
        rows.append({'kind': 'total', 'cells': [
            '合计（三岗）', d2(data.get('data_date')) or '--',
            f1(agg['onjob']), f1(agg['kaohe']), f1(agg['pass']),
            fpct(rate), '网销24:00 + 电邀任一 + 综法19:30', fmoney(agg['perf']),
            fmoney(agg['perf'] / agg['kaohe'])]})
    return cols, rows


def _dept_sort_key(post):
    f = POST_CFG[post]['rate']

    def k(item):
        dept, r = item
        return -(num(r['cells'].get(f)) or -1)
    return k


# 各岗明细表的列定义（第二个元素 = 该列取值的函数）
def _wx_cells(c):
    return [f1(c.get('kaohe')), f1(c.get('avgT1730')), fpct(c.get('r1730')),
            f1(c.get('avgT2400')),
            {'t': fpct(c.get('r2400')), 'c': rate_color(c.get('r2400')), 'bar': c.get('r2400')},
            fmoney(c.get('dayPerf'))]


def _dx_cells(c):
    return [f1(c.get('kaohe')), f1(c.get('avgContact')), f1(c.get('n6pass')),
            {'t': fpct(c.get('n6rate')), 'c': rate_color(c.get('n6rate')), 'bar': c.get('n6rate')},
            '%d/6' % n6_count(c), fmoney(c.get('dayPerf'))]


def _zf_cells(c):
    return [f1(c.get('onjob')), f1(c.get('kaoheN')), f1(c.get('meetAll')),
            f1(c.get('meetAvg')), fmoney(c.get('payAll')),
            {'t': fpct(c.get('r1930')), 'c': rate_color(c.get('r1930')), 'bar': c.get('r1930')}]


COLS_POST = {
    '网销岗': [('部门 / 姓名', 0.20, 'l'), ('考核(人)', 0.08, 'r'),
               ('通时17:30', 0.12, 'r'), ('达标率17:30', 0.14, 'r'),
               ('通时24:00', 0.12, 'r'), ('达标率24:00', 0.14, 'r'),
               ('当日业绩', 0.12, 'r'), ('状态', 0.08, 'c')],
    '电邀岗': [('部门 / 姓名', 0.22, 'l'), ('考核(人)', 0.09, 'r'),
               ('人均接触量', 0.13, 'r'), ('达标人数', 0.11, 'r'),
               ('达标率', 0.14, 'r'), ('达标项数', 0.10, 'c'),
               ('当日业绩', 0.13, 'r'), ('状态', 0.08, 'c')],
    '综法岗': [('部门 / 姓名', 0.22, 'l'), ('在职(人)', 0.10, 'r'),
               ('考核(人)', 0.10, 'r'), ('合格见客', 0.13, 'r'),
               ('人均见客', 0.10, 'r'), ('当日收款', 0.14, 'r'),
               ('达标率19:30', 0.13, 'r'), ('状态', 0.08, 'c')],
}
CELLS_POST = {'网销岗': _wx_cells, '电邀岗': _dx_cells, '综法岗': _zf_cells}


def build_post_panel(data, post, mode):
    cols = COLS_POST[post]
    v = data['hz'].get(post)
    if not v:
        # 严格按「当天」取数时，某岗当天可能没有考核记录（如综法）——出一张「无数据」图
        row = {'kind': 'empty', 'cells': [
            {'t': '　当日（%s）无考核数据' % cn_date(data.get('data_date')),
             'a': 'l', 'c': C_MUTED}] + [''] * (len(cols) - 1)}
        return cols, [row], []
    cells_of = CELLS_POST[post]
    rf = POST_CFG[post]['rate']

    undone = []
    rows = []
    for dept, r in sorted(v['depts'].items(), key=_dept_sort_key(post)):
        c = dict(r['cells'])
        if post == '电邀岗':
            # 用户口径：部门达标人数/达标率也按「6 项任一项」重算
            st_any = dx_any(v, dept)
            if st_any:
                c['n6pass'], c['n6rate'] = st_any[0], st_any[2]
                if st_any[1] <= 0:
                    st = 'na'
                else:
                    st = 'pass' if st_any[0] >= st_any[1] else 'fail'
            else:
                st, _ = person_status(post, c)
        else:
            st, _ = person_status(post, c)
        rows.append({'kind': 'dept', 'cells': [
            {'t': clean_dept(dept), 'a': 'l', 'b': True}] + cells_of(c) +
            [{'t': STATUS_BADGE[st][0], 'badge': STATUS_BADGE[st][1]}]})
        graded = []
        for p in v['people'].get(dept, []):
            stp, txtp = person_status(post, p['cells'])
            if stp == 'na' and kaohe_of(post, p['cells']) <= 0:
                continue
            graded.append((stp, txtp, p))
        if post == '电邀岗':
            graded.sort(key=lambda x: (x[0] != 'fail', -n6_count(x[2]['cells'])))
        else:
            graded.sort(key=lambda x: (x[0] != 'fail', -(num(x[2]['cells'].get(rf)) or 0)))
        shown = [x for x in graded if mode == 'result' or x[0] == 'fail']
        if mode == 'track':
            undone.extend([(clean_dept(dept), clean_name(p['lead']), txtp)
                           for stp, txtp, p in shown])
        if not shown:
            rows.append({'kind': 'empty', 'cells': [
                {'t': '　└ 本期全部达标 ✓' if graded else '　└ 本期无考核数据',
                 'a': 'l', 'c': GREEN if graded else C_MUTED}]})
            continue
        for stp, txtp, p in shown:
            rows.append({'kind': 'person', 'cells': [
                {'t': clean_name(p['lead']), 'a': 'l', 'indent': 22}] +
                cells_of(p['cells']) +
                [{'t': STATUS_BADGE[stp][0], 'badge': STATUS_BADGE[stp][1]}]})
    return cols, rows, undone


def render_all(mode='track', data=None):
    os.makedirs(OUTDIR, exist_ok=True)
    if data is None:
        data = wb_source.fetch_day(day=today_ymd())      # 严格取「当天」
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    anchors = data.get('anchor') or {}
    day = data.get('data_date') or anchors.get('网销岗') or today_ymd()
    diff = [p for p in POST_ORDER if anchors.get(p) and anchors.get(p) != day]
    dstr = cn_date(day) + (('（%s）' % ' · '.join('%s %s' % (p, d2(anchors[p])) for p in diff))
                           if diff else '')
    sub = ('数据周期：单日 %s（当天）　｜　仅 · 网销/电邀/综法　｜　'
           '数据源：活动量工作台·七岗排行榜（页面【日】档口径）　｜　制图 %s' % (dstr, now))

    c1, r1 = build_overview(data, mode)
    p1 = panel('活动量 · 整体完成情况（单日）', sub,
               ['口径：达标率 = 达标 ÷ 考核；网销 24:00 达标率、综法 19:30 达标率（源站口径）；'
                '电邀为「6 项任一项达标」（用户口径）。',
                '各岗锚定自身最新有效日：网销/电邀=全局最新日，综法=本岗最新有效考核日（通常落后几天）。'],
               c1, r1, os.path.join(OUTDIR, 'act1_overview_%s.png' % mode),
               '分公司 · 整体')

    out = {'mode': mode, 'source': 'workbuddy-rank-day', 'cycle': data.get('cycle_key'),
           'data_date': data.get('data_date'), 'anchor': anchors,
           'built_at': now, 'panels': [], 'undone': {}}
    for post in POST_ORDER:
        fname, label = POST_FNAME[post]
        built = build_post_panel(data, post, mode)
        if not built:
            print('跳过（无数据）:', post)
            continue
        cols, rows, undone = built
        notes = []
        vv = data['hz'].get(post)
        tot = post_totals(post, vv) if vv else {'onjob': 0.0, 'kaohe': 0.0, 'pass': 0.0, 'perf': 0.0}
        if post == '综法岗':
            if data['hz'].get(post):
                notes.append('口径：达标率 = 19:30 达标人数 ÷ 考核人数。')
            else:
                notes.append('说明：综法岗当日（%s）没有有效考核记录，故本日无数据。'
                             % cn_date(day))
        if post == '电邀岗':
            notes.append('达标口径（用户 2026-10-10）：%s 六项中「任意一项」达标即算达标'
                         '（网页默认是 6 项全中，故本图达标率与网页可能不同）。'
                         % '/'.join(N6_LABEL[k] for k in N6_KEYS))
            notes.append('说明：电邀岗当日仅 %.0f 人参与考核（在职 %.0f 人）。'
                         % (tot['kaohe'], tot['onjob']))
        if mode == 'track':
            notes.append('本图只列「未达标」人员；部门行为源站聚合值，含该部门全部考核人员。')
        p = panel('活动量 · %s 完成情况（%s）' % (label, d2(anchors.get(post))), sub,
                  notes, cols, rows,
                  os.path.join(OUTDIR, '%s_%s.png' % (fname, mode)), '分公司 · ' + label)
        out['panels'].append(os.path.basename(p))
        out['undone'][post] = undone
    out['panels'].insert(0, os.path.basename(p1))
    with open(os.path.join(BASE, 'act_state.json'), 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    print('\n模式:', '未达标追踪' if mode == 'track' else '全员结果')
    print('数据周期: 单日 %s（当天，严格过滤）' % dstr)
    for post in POST_ORDER:
        u = out['undone'].get(post, [])
        print('未达标 %s: %d 人 %s' % (post, len(u),
                                      '、'.join('%s %s' % (n, t) for _, n, t in u)))
    return out


if __name__ == '__main__':
    render_all((sys.argv[1] if len(sys.argv) > 1 else 'track').strip().lower())
