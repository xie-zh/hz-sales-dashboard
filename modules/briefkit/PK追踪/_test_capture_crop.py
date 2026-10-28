# -*- coding: utf-8 -*-
"""离线验证：只跑「截图 + 裁剪 + 画面校验」，**不发送任何消息**。
用于确认修复后产出的图中确实包含下方两张活动海报。"""
import os
import sys

sys.path.insert(0, r'<PROJECT_ROOT>/数据文件/PK追踪')
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import pk_report  # noqa: E402

from PIL import Image  # noqa: E402

print('IMG_OUT =', pk_report.IMG_OUT)
bottom = pk_report.capture_and_crop()
pk_report.verify_shot_is_table(pk_report.IMG_OUT)

img = Image.open(pk_report.IMG_OUT).convert('RGB')
W, H = img.size
print(f'\n产物: {W} x {H}  {os.path.getsize(pk_report.IMG_OUT)} B')
print('裁剪底缘 bottom =', bottom)

# 统计「表格底(812)-261 = 551 之后」的海报区是否有内容
px = img.load()
poster_rows = 0
for y in range(600, H, 4):
    cnt = 0
    for x in range(0, W, 8):
        p = px[x, y]
        if not (p[0] >= 252 and p[1] >= 252 and p[2] >= 252):
            cnt += 1
    if cnt > 60:
        poster_rows += 1
print(f'裁剪图中「表格下方」有内容的采样行数 = {poster_rows} '
      f'(=0 表示海报没截进来)')
print('结论:', '✅ 海报已包含' if poster_rows > 5 else '❌ 海报仍缺失')
