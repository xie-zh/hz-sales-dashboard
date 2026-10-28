# -*- coding: utf-8 -*-
"""汇总 run_report.sh 的发送结果（新版：文案 + 4 图 = 5 条）
用法：python show_result.py [track|result]
"""
import io
import json
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
BASE = os.path.dirname(os.path.abspath(__file__))

LABELS = [('send1.json', '文案'),
          ('send2.json', '图①整体完成情况'),
          ('send3.json', '图②网销岗'),
          ('send4.json', '图③电邀岗'),
          ('send5.json', '图④综法岗')]

all_ok = True
for fn, label in LABELS:
    p = os.path.join(BASE, fn)
    if not os.path.exists(p):
        print('  %s -> 未执行（缺 %s）' % (label, fn))
        all_ok = False
        continue
    try:
        d = json.load(open(p, encoding='utf-8'))
    except Exception as e:
        print('  %s -> 解析失败: %s' % (label, e))
        all_ok = False
        continue
    ok = d.get('success')
    print('  %s -> success=%s errcode=%s %s'
          % (label, ok, d.get('errcode'), d.get('errmsg', '') or ''))
    if not ok:
        all_ok = False

print('结果:', '全部发送成功' if all_ok else '存在失败，请检查')
sys.exit(0 if all_ok else 1)
