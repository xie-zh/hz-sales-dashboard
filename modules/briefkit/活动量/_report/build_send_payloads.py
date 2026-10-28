# -*- coding: utf-8 -*-
"""组装企微群汇报负载（新版数据源 · 仅 · 网销/电邀/综法）

  track  —— 活动量追踪（11:00/14:30/16:00）：文案（未达标人员 + 当前达标率）+ 4 张图
  result —— 活动量结果（18:00）：文案 + 4 张图

用法：python build_send_payloads.py [track|result] [可选文案后缀]
环境变量 REPORT_GROUP 可覆盖目标群（默认「数据点检群」），用于测试群试发。

依赖同目录：sessions.json、up1.json~up4.json（四图）、act_state.json（由 render_wb.py 写入）
"""
import io
import json
import os
import sys
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import mention_util  # noqa: E402  纯文本 @姓名（aibot 无法真@）

GROUP_NAME = os.environ.get('REPORT_GROUP', '数据点检群')
# 对外链接：活动量工作台 · 七岗排行榜（2026-10-09 起改用新数据源，文案同步换此链接）
LINK = 'https://dashboard.example.com/p/LLT6AK0VUUvnPu12Cc8PKZ'

MODE = (sys.argv[1] if len(sys.argv) > 1 else 'track').strip().lower()
RESULT = MODE in ('result', 'res', 'r')
SUFFIX = sys.argv[2] if len(sys.argv) > 2 else ''
TAG = 'result' if RESULT else 'track'

IMG1 = 'report_img/act1_overview_%s.png' % TAG
IMG2 = 'report_img/act2_wangxiao_%s.png' % TAG
IMG3 = 'report_img/act3_dianxiao_%s.png' % TAG
IMG4 = 'report_img/act4_zongfa_%s.png' % TAG

# 未达标名单分组标题（岗位 -> 文案前缀）
GROUP_TITLE = {'网销岗': '网销（24:00 达标率）',
               '电邀岗': '电邀（6 项任一达标）',
               '综法岗': '综法（19:30 达标率）'}


def load(name):
    with open(os.path.join(BASE, name), encoding='utf-8') as f:
        return json.load(f)


def img(media_id):
    return {'chat_id': chat_id, 'msg_type': 'image', 'image': {'media_id': media_id}}


sessions = load('sessions.json')
chat_id = None
for s in sessions.get('sessions', []):
    if GROUP_NAME in str(s.get('chat_name', '')):
        chat_id = s['chat_id']
        break
if not chat_id:
    print('ERROR: 本次会话列表中未找到群「%s」，请先在群里 @ 机器人' % GROUP_NAME)
    sys.exit(2)

m1 = load('up1.json')['media_id']
m2 = load('up2.json')['media_id']
m3 = load('up3.json')['media_id']
m4 = load('up4.json')['media_id']

now = datetime.now()
kind = '活动量结果数据' if RESULT else '活动量追踪数据'
note = ('当日全员结果如下，已达标人员可安排下班'
        if RESULT else
        '特别提醒!!👉若截止17:30活动量未达标,则什么时候达成什么时候下班，最晚需在19:31以后打卡☀️')
# 注意：@ 全体成员由「群机器人 Webhook」负责（见 run_report.sh 的 send_atall.py）；
# 链接后面必须留一个空格，否则企业微信会把紧跟的「(」吞进超链接，点开链接会错。

st = load('act_state.json')
anchors = st.get('anchor') or {}


def _cn(s):
    s = str(s or '')
    return '%s-%s-%s' % (s[0:4], s[4:6], s[6:8]) if len(s) == 8 else (s or '--')


P_SHORT = {'网销岗': '网销', '电邀岗': '电邀', '综法岗': '综法'}
main_day = _cn(st.get('data_date') or anchors.get('网销岗'))
_extra = ' · '.join('%s %s' % (P_SHORT.get(p, p), _cn(d))
                    for p, d in anchors.items() if _cn(d) != main_day)
# 例：单日 2026-10-09 · 综法 2026-09-29 —— 综法通常落后几天，必须显式标注
period = '单日 %s' % main_day + ((' · %s' % _extra) if _extra else '')

if RESULT:
    caption = ('今日{}点{}分{}{}（数据周期：{}）：{} ({}\n'
               '全员完成情况见下四图（图①整体 / 图②网销岗 / 图③电邀岗 / 图④综法岗）').format(
        now.hour, now.minute, kind, SUFFIX, period, LINK, note)
else:
    undone = st.get('undone', {})
    names = [n for post in ('网销岗', '电邀岗', '综法岗') for _, n, _ in undone.get(post, [])]
    mention_util.resolve_names(names)
    n_wx = len(undone.get('网销岗', []))
    n_dx = len(undone.get('电邀岗', []))
    n_zf = len(undone.get('综法岗', []))
    lines = ['今日{}点{}分{}{}（数据周期：{}）：{} ({}\n'
             '未达标 {} 人（网销 {} · 电邀 {} · 综法 {}）· 当前达标率：'.format(
                 now.hour, now.minute, kind, SUFFIX, period, LINK, note,
                 n_wx + n_dx + n_zf, n_wx, n_dx, n_zf)]
    for post in ('网销岗', '电邀岗', '综法岗'):
        rows = undone.get(post, [])
        if rows:
            lines.append(GROUP_TITLE[post] + '：' + ' ｜ '.join(
                '%s %s' % (mention_util.mention(n), t) for _, n, t in rows))
    if not (n_wx or n_dx or n_zf):
        lines.append('本期全员达标 🎉')
    caption = '\n'.join(lines)

payloads = {
    'p_caption.json': {'chat_id': chat_id, 'msg_type': 'markdown',
                       'markdown': {'content': caption}},
    'p_img1.json': img(m1),
    'p_img2.json': img(m2),
    'p_img3.json': img(m3),
    'p_img4.json': img(m4),
}

for name, obj in payloads.items():
    with open(os.path.join(BASE, name), 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False)
print('目标群:', GROUP_NAME, chat_id)
print('数据周期:', period)
print('图:', IMG1, '|', IMG2, '|', IMG3, '|', IMG4)
print('文案:', caption.replace('\n', '  ⏎  '))
