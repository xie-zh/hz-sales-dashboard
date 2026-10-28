# -*- coding: utf-8 -*-
"""校验 auth_state.json 是否真的有效(存在 != 有效)。
用 dop-api 探测: 命中 check userInfo failed / retcode 30004 即失效。
只读校验，不写任何文件。
"""
import asyncio
import datetime
import os
import sys
import time

PROJ = r'<PROJECT_ROOT>/数据文件/20260924_tencent-doc-export'
sys.path.insert(0, os.path.join(PROJ, 'src'))
import core  # noqa: E402

LINK = ('https://doc.weixin.qq.com/sheet/e3_Ae8A_gYvADACN1gvbrUgvT3SalKT9'
        '?scode=ADwAKQdcAAcl4y7c4pAaIAAQbBAM8&tab=be48i0')

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass


def main() -> bool:
    st = core.state_file()
    print('state_file =', st)
    print('exists     =', st.exists())
    if not st.exists():
        print('LOGIN_VALID = False  (无登录态文件)')
        return False
    mt = st.stat().st_mtime
    print('mtime      =', datetime.datetime.fromtimestamp(mt).strftime('%Y-%m-%d %H:%M:%S'))
    print('age_min    = %.1f' % ((time.time() - mt) / 60))

    info = core.parse_link(LINK)
    probe = core._api_url(info)
    edge = core.find_edge()
    print('edge       =', edge)
    if not edge:
        print('LOGIN_VALID = False  (未找到浏览器)')
        return False

    async def run():
        from playwright.async_api import async_playwright
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                executable_path=edge, headless=True,
                args=['--disable-blink-features=AutomationControlled'])
            try:
                ctx = await browser.new_context(
                    storage_state=str(st), user_agent=core.UA,
                    viewport={'width': 1440, 'height': 900})
                page = await ctx.new_page()
                await page.goto(info['url'], wait_until='domcontentloaded',
                                timeout=core.PAGE_TIMEOUT_S * 1000)
                ok = await core._check_login(page, probe)
                print('probe_url  =', probe[:110] + '...')
                return bool(ok)
            finally:
                try:
                    await browser.close()
                except Exception:
                    pass

    ok = asyncio.run(run())
    print('LOGIN_VALID =', ok)
    return ok


if __name__ == '__main__':
    sys.exit(0 if main() else 3)
