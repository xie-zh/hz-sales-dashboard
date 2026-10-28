# -*- coding: utf-8 -*-
"""用已登录 storage_state 打开 PK sheet(用户提供的最新链接)，整页截图。

⚠️ 海报渲染时序（2026-10-08 踩坑，务必保留这个等待逻辑）：
在线表下方那两张「活动海报」是浮动图片对象，页面刚打开 / 刚同步完时可能还没渲染，
此时截图会得到「表格 + 一大片空白」，下游裁剪就找不到海报。
实测同一天同一脚本：14:35 那次 raw = 327,963 B（海报空白），重跑一次 = 601,220 B
（海报正常）。所以截图前必须等海报真的加载出来，不能只固定等几秒。
"""
import asyncio, os, sys, time
PROJ = r'<PROJECT_ROOT>/数据文件/20260924_tencent-doc-export'
sys.path.insert(0, os.path.join(PROJ, 'src'))
import core

LINK = ('https://doc.weixin.qq.com/sheet/e3_Ae8A_gYvADACN1gvbrUgvT3SalKT9'
        '?scode=ADwAKQdcAAcl4y7c4pAaIAAQbBAM8&tab=be48i0')
RAW = os.path.join(PROJ, 'scratch', 'pk_raw2.png')

# 海报 = 页面里的大图(<img>)：已加载完成(naturalWidth>0) 且渲染高度 >150px 才算就绪
POSTER_JS = r"""() => {
    const imgs = [...document.images];
    const big = imgs.filter(i => i.complete && i.naturalWidth > 0
                                 && i.getBoundingClientRect().height > 150);
    return { n: imgs.length, big: big.length,
             detail: imgs.map(i => ({c: i.complete, nw: i.naturalWidth,
                                     h: Math.round(i.getBoundingClientRect().height)})) };
}"""


async def wait_for_posters(page, timeout_s=30, need=1):
    """轮询等海报加载完。返回 (ok, 说明)；超时不抛异常，由调用方决定后续。"""
    t0 = time.time()
    last = None
    while time.time() - t0 < timeout_s:
        try:
            last = await page.evaluate(POSTER_JS)
            if int(last.get('big', 0)) >= need:
                return True, (f'{last["big"]} 张大图已就绪'
                              f'(等了 {time.time() - t0:.1f}s)')
        except Exception as e:
            last = {'err': f'{type(e).__name__}: {e}'[:120]}
        await page.wait_for_timeout(800)
    return False, f'等 {timeout_s}s 仍未就绪 last={last}'


async def main():
    edge = core.find_edge()
    if not edge or not core.login_state_exists():
        print('NO_EDGE_OR_NO_LOGIN'); return
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=edge, headless=True,
            args=["--disable-blink-features=AutomationControlled"])
        context = await browser.new_context(
            storage_state=str(core.state_file()), user_agent=core.UA,
            viewport={"width": 1600, "height": 1450}, device_scale_factor=1.5)
        page = await context.new_page()
        await page.goto(LINK, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(6000)
        ok, info = await wait_for_posters(page)      # ← 关键：等海报再截图
        print("POSTER_WAIT:", "OK" if ok else "TIMEOUT", "-", info)
        title = (await page.title() or "").strip()
        await page.screenshot(path=RAW, full_page=True)
        print("TITLE:", title)
        print("RAW:", RAW, os.path.getsize(RAW), "bytes")
        await browser.close()

asyncio.run(main())
