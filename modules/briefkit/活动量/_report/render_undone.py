# -*- coding: utf-8 -*-
"""
活动量看板「未达标人员」报表图渲染
- 数据源：看板实时 data.json（页面本身也用它渲染）
- 输出：report_img/panel1_total_undone.png（全员明细·总表）
        report_img/panel2_legal_undone.png（法律日达标）
- 只保留「是否达标 == 不达标」的人员，免考核不展示
样式对齐看板自身 CSS（--accent #4f46e5 / rateColor / fmtNum / badge）
"""
import json, os, sys, io, time, urllib.request
from PIL import Image, ImageDraw, ImageFont

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(BASE, 'report_img')
DATA_URL = ('https://dashboard.example.com'
            '/dash/data.json')
FACE = r'C:\Windows\Fonts\msyh.ttc'
FACEB = r'C:\Windows\Fonts\msyhbd.ttc'

C_TEXT, C_MUTED, C_BORDER = '#1a1a2e', '#6b7280', '#e5e7eb'
C_ACCENT, C_PAGE = '#4f46e5', '#f4f6f9'
C_FAIL_BG, C_FAIL_FG = '#fee2e2', '#991b1b'
C_PASS_BG, C_PASS_FG = '#dcfce7', '#166534'   # 达标徽章（结果模式才会用到）
FS_TITLE, FS_TAB, FS_TH, FS_TD, FS_FOOT, FS_BADGE = 22, 20, 21, 21, 17, 18
ROW_H, TH_H, TITLE_H, TAB_H, FOOT_H, M, RADIUS = 44, 48, 54, 56, 40, 18, 14


def fmt_num(v):
    if v is None:
        return '--'
    if isinstance(v, (int, float)):
        if v >= 100:
            return '%.0f' % v
        if v >= 10:
            return '%.1f' % v
        return '%.2f' % v
    return str(v)


def rate_color(v):
    if v is None:
        return '#aaaaaa'
    return '#16a34a' if v >= 1 else ('#f59e0b' if v >= 0.5 else '#dc2626')


def gap_minutes(r):
    """距达标还需要多少分钟 = 目标 - 完成（活动量单位即分钟，目标 60）。
    完成/目标缺失返回 None；已达标或超额返回 0。"""
    done = r.get('完成')
    tgt = r.get('折扣目标')
    if tgt is None:
        tgt = r.get('目标')
    if done is None or tgt is None:
        return None
    try:
        return max(0, int(round(float(tgt) - float(done))))
    except (TypeError, ValueError):
        return None


def fetch_data():
    req = urllib.request.Request(DATA_URL + '?t=%d' % int(time.time()),
                                 headers={'Cache-Control': 'no-cache'})
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.loads(r.read().decode('utf-8'))


# ---------- 列定义：(表头, 宽度, 对齐, 类型, 取值字段) ----------
COLS1 = [('姓名', 100, 'lm', 'text', '姓名'),
         ('事业部', 205, 'lm', 'text', '事业部'),
         ('部门', 245, 'lm', 'text', '部门'),
         ('目标', 95, 'rm', 'num', '目标'),
         ('完成量', 105, 'rm', 'num', '完成'),
         ('完成率', 330, 'rm', 'rate', '完成率'),
         ('周活动量', 120, 'lm', 'text', '周活动量'),
         ('见客', 95, 'rm', 'num', '见客合计'),
         ('是否达标', 135, 'cm', 'badge', '是否达标')]

COLS2 = [('日达标', 115, 'cm', 'badge', '是否达标'),
         ('姓名', 105, 'lm', 'text', '姓名'),
         ('部门', 230, 'lm', 'text', '部门'),
         ('今日完成', 115, 'rm', 'num', '今日完成'),
         ('滚动2日', 115, 'rm', 'num', '滚动2日完成'),
         ('滚动3日', 115, 'rm', 'num', '滚动3日完成'),
         ('完成率', 135, 'rm', 'rate_text', '完成率'),
         ('上门', 95, 'rm', 'num', '当日上门量'),
         ('外出', 95, 'rm', 'num', '当日外出量'),
         ('线上', 95, 'rm', 'num', '当日线上见客量'),
         ('新增业绩', 120, 'rm', 'num', '当日新增业绩')]


def draw_badge(d, cx, cy, text):
    fill, fg = (C_FAIL_BG, C_FAIL_FG) if text == '未达标' else \
               ((C_PASS_BG, C_PASS_FG) if text == '达标' else ('#f3f4f6', C_MUTED))
    f = ImageFont.truetype(FACEB, FS_BADGE)
    tw = d.textlength(text, font=f)
    bw, bh = tw + 26, 30
    d.rounded_rectangle([cx - bw / 2, cy - bh / 2, cx + bw / 2, cy + bh / 2],
                        radius=13, fill=fill)
    d.text((cx, cy), text, font=f, fill=fg, anchor='mm')


def draw_rate(d, x_right, cy, rate, bar_max=130, with_bar=True):
    pct = '--' if rate is None else '%.1f%%' % (rate * 100)
    f = ImageFont.truetype(FACE, FS_TD)
    col = rate_color(rate)
    d.text((x_right, cy), pct, font=f, fill=col, anchor='rm')
    if with_bar and rate is not None:
        tw = d.textlength(pct, font=f)
        bar_w = max(3, min(rate, 1.0) * bar_max)
        bx2 = x_right - tw - 14
        bx1 = bx2 - bar_max
        d.rounded_rectangle([bx1, cy - 4, bx2, cy + 4], radius=4, fill='#eceef3')
        d.rounded_rectangle([bx1, cy - 4, bx1 + bar_w, cy + 4], radius=4, fill=col)


def render(cols, rows, title, total_n, tab_on_idx, active_tab, data_time, out_path,
           foot_fmt='数据时间 %s · 仅显示未达标人员（免考核人员不列入）', tabs=None):
    if tabs is None:
        tabs = ['全员明细（总表）', '法律日达标']
    tbl_w = sum(c[1] for c in cols)
    W = tbl_w + 2 * M
    H = 2 * M + TITLE_H + TAB_H + TH_H + max(len(rows), 1) * ROW_H + FOOT_H
    img = Image.new('RGB', (W, H), C_PAGE)
    d = ImageDraw.Draw(img)

    card_top = M
    card_bot = H - M
    d.rounded_rectangle([M, card_top, M + tbl_w, card_bot], radius=RADIUS,
                        fill='#ffffff', outline=C_BORDER, width=1)

    f_title = ImageFont.truetype(FACEB, FS_TITLE)
    f_tab = ImageFont.truetype(FACEB, FS_TAB)
    f_th = ImageFont.truetype(FACEB, FS_TH)
    f_td = ImageFont.truetype(FACE, FS_TD)
    f_tdb = ImageFont.truetype(FACEB, FS_TD)
    f_foot = ImageFont.truetype(FACE, FS_FOOT)

    # ---- 标题行 ----
    ty = card_top + TITLE_H / 2
    d.text((M + 20, ty), title, font=f_title, fill=C_TEXT, anchor='lm')
    tw = d.textlength(title, font=f_title)
    d.text((M + 20 + tw + 16, ty), '显示 %d / %d 人' % (len(rows), total_n),
           font=f_tab, fill=C_MUTED, anchor='lm')

    # ---- tab 行 ----
    x = M + 20
    for i, t in enumerate(tabs):
        twi = d.textlength(t, font=f_tab)
        bw = twi + 44
        yy = card_top + TITLE_H
        if i == tab_on_idx:
            d.rounded_rectangle([x, yy + 6, x + bw, yy + TAB_H - 6], radius=10,
                                fill=C_ACCENT)
            d.text((x + bw / 2, yy + TAB_H / 2), t, font=f_tab, fill='#ffffff',
                   anchor='mm')
        else:
            d.rounded_rectangle([x, yy + 6, x + bw, yy + TAB_H - 6], radius=10,
                                fill='#ffffff', outline=C_BORDER, width=1)
            d.text((x + bw / 2, yy + TAB_H / 2), t, font=f_tab, fill=C_MUTED,
                   anchor='mm')
        x += bw + 12

    # ---- 表头 ----
    hy = card_top + TITLE_H + TAB_H
    d.rectangle([M + 1, hy, M + tbl_w - 1, hy + TH_H], fill=C_ACCENT)
    cx = M
    for (label, w, align, _, _) in cols:
        ax = cx + 16 if align == 'lm' else (cx + w - 16 if align == 'rm' else cx + w / 2)
        d.text((ax, hy + TH_H / 2), label, font=f_th, fill='#ffffff',
               anchor={'lm': 'lm', 'rm': 'rm', 'cm': 'mm'}[align])
        cx += w

    # ---- 数据行 ----
    y = hy + TH_H
    if not rows:
        d.rectangle([M + 1, y, M + tbl_w - 1, y + ROW_H], fill='#ffffff')
        d.text((M + tbl_w / 2, y + ROW_H / 2), '全部达标 ✓ 当前无未达标人员',
               font=f_tdb, fill='#16a34a', anchor='mm')
        y += ROW_H
    for ri, row in enumerate(rows):
        cy = y + ri * ROW_H + ROW_H / 2
        d.rectangle([M + 1, y + ri * ROW_H, M + tbl_w - 1, y + (ri + 1) * ROW_H],
                    fill='#ffffff' if ri % 2 == 0 else '#fafbff')
        d.line([M + 1, y + ri * ROW_H, M + tbl_w - 1, y + ri * ROW_H],
               fill=C_BORDER, width=1)
        cx = M
        for (label, w, align, typ, key) in cols:
            val = row.get(key)
            if typ == 'text':
                ax = cx + 16
                fnt = f_tdb if (label == '姓名') else f_td
                d.text((ax, cy), '' if val is None else str(val), font=fnt,
                       fill=C_TEXT, anchor='lm')
            elif typ == 'num':
                ax = cx if align == 'rm' else cx + 16
                anc = 'rm' if align == 'rm' else 'lm'
                off = -16 if align == 'rm' else 0
                d.text((ax + w + off if align == 'rm' else ax, cy), fmt_num(val),
                       font=f_td, fill=C_TEXT, anchor=anc)
            elif typ == 'badge':
                draw_badge(d, cx + w / 2, cy,
                           '未达标' if val == '不达标' else str(val))
            elif typ == 'rate':
                draw_rate(d, cx + w - 16, cy, val, with_bar=True)
            elif typ == 'rate_text':
                draw_rate(d, cx + w - 16, cy, val, with_bar=False)
            cx += w
        d.line([M + 1, y + (ri + 1) * ROW_H, M + tbl_w - 1, y + (ri + 1) * ROW_H],
               fill=C_BORDER, width=1)

    # ---- 页脚 ----
    fy = card_bot - FOOT_H / 2
    d.text((M + 20, fy), foot_fmt % data_time,
           font=f_foot, fill=C_MUTED, anchor='lm')

    img.save(out_path)
    print('已生成 %s  %dx%d  行数=%d' % (out_path, W, H, len(rows)))
    return out_path


def main():
    # 模式：track（默认，11:00/14:30/16:00 追踪，仅未达标）| result（18:00，全员结果总表）
    mode = (sys.argv[1] if len(sys.argv) > 1 else 'track').strip().lower()
    result_mode = mode in ('result', 'res', 'r')
    os.makedirs(OUTDIR, exist_ok=True)
    data = fetch_data()
    dt = data['meta']['data_time']

    if result_mode:
        rows1 = list(data['total'])
        rows2 = list(data['legal'])
        title = '当日结果总表'
        foot = '数据时间 %s · 当日全员结果（达标 / 未达标 / 免考核）'
        suffix = 'result'
    else:
        rows1 = [r for r in data['total'] if r.get('是否达标') == '不达标']
        rows2 = [r for r in data['legal'] if r.get('是否达标') == '不达标']
        title = '人员明细'
        foot = '数据时间 %s · 仅显示未达标人员（免考核人员不列入）'
        suffix = 'undone'

    tabs1 = None   # 两种模式都显示「全员明细（总表）/ 法律日达标」两个 tab（法律同样公示）
    p1 = render(COLS1, rows1, title, len(data['total']), 0, '总表',
                dt, os.path.join(OUTDIR, 'panel1_total_%s.png' % suffix), foot_fmt=foot,
                tabs=tabs1)
    p2 = render(COLS2, rows2, title, len(data['legal']), 1, '法律',
                dt, os.path.join(OUTDIR, 'panel2_legal_%s.png' % suffix), foot_fmt=foot)

    def names(rows):
        return ', '.join('%s(%s)' % (r['姓名'], r.get('是否达标')) for r in rows) or '无'

    print('\n模式:', '结果总表' if result_mode else '未达标追踪')
    print('数据时间:', dt)
    print('板块1（总表 %d 人）:' % len(rows1), names(rows1))
    print('板块2（法律 %d 人）:' % len(rows2), names(rows2))
    def _rows(rs):
        return [{'姓名': r['姓名'], '部门': r.get('部门'), '是否达标': r.get('是否达标'),
                 '距达标分钟': gap_minutes(r) if not result_mode else None} for r in rs]

    with open(os.path.join(BASE, 'report_state.json'), 'w', encoding='utf-8') as f:
        json.dump({'mode': 'result' if result_mode else 'track', 'data_time': dt,
                   'panel1': _rows(rows1), 'panel2': _rows(rows2)}, f, ensure_ascii=False)


if __name__ == '__main__':
    main()
