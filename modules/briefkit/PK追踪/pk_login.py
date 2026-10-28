# -*- coding: utf-8 -*-
"""弹浏览器扫码登录,保存 auth_state.json。成功后即可跑 pk_report.py。"""
import asyncio
import os
import sys

PROJ = r'<PROJECT_ROOT>/数据文件/20260924_tencent-doc-export'
sys.path.insert(0, os.path.join(PROJ, 'src'))
import core

LINK = ('https://doc.weixin.qq.com/sheet/e3_Ae8A_gYvADACN1gvbrUgvT3SalKT9'
        '?scode=ADwAKQdcAAcl4y7c4pAaIAAQbBAM8&tab=be48i0')

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

ok = asyncio.run(core.interactive_login(LINK))
print('LOGIN_RESULT =', 'OK' if ok else 'FAIL')
print('state_file =', core.state_file(),
      'exists =', core.state_file().exists())
sys.exit(0 if ok else 1)