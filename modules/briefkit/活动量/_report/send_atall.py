# -*- coding: utf-8 -*-
"""活动量战报 @ 全体（群机器人 webhook）。用法：python send_atall.py [群关键字] [文案]"""
import sys
sys.path.insert(0, r'<PROJECT_ROOT>/数据文件/_common')
from atall import send_at_all, load_webhooks

KEY = sys.argv[1] if len(sys.argv) > 1 else "数据点检群"
TXT = (sys.argv[2] if len(sys.argv) > 2 else
       "📢 分公司活动量战报已更新，请各位伙伴查看下方报表 👇")
send_at_all(load_webhooks().get(KEY, ""), TXT)
