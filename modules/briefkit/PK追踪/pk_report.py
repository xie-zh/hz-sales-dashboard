# -*- coding: utf-8 -*-
"""
PK 战报自动汇报流水线（业绩联动版）
====================================
流程：
  1) 调 sync_yj.one_cycle(notify=False)：拉取 BI 源数据 → 替换在线表【业绩（含续费）】→ 变化检测
  2) 有数据变化（或 --force）：
     a. 重截 PK sheet（be48i0）整页截图，裁剪 = 纯表格 + 下方两张活动海报
     b. 从 be48i0 直读最新数据生成战报文案（数据时间取自【业绩（含续费）】R2）
     c. 发送 markdown 文案 + 截图 到指定群（默认 3C认证团队 = 实验群）
  3) 无数据变化：静默跳过（不发消息、不截图，省部署/算力）

用法：
    python pk_report.py                          # 常规：有变化才发 3C 群
    python pk_report.py --force                  # 强制发送（用于 3C 群确认格式）
    python pk_report.py --force --group 管理者大群   # 确认后发指定群（可多次传）

依赖：sync_yj.py（同套 wecom-cli 调用链）、playwright+PIL（default env）、core.py 登录态。
"""

import argparse
import json
import os
import subprocess
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

PK_DIR = r"<PROJECT_ROOT>/数据文件/PK追踪"
SYNC_DIR = r"<PROJECT_ROOT>/数据文件/业绩BI同步"
SCRATCH_DIR = r"<PROJECT_ROOT>/数据文件/20260924_tencent-doc-export/scratch"
PY_EXE = r"<HOME>/.workbuddy/binaries/python/envs/default/Scripts/python.exe"

DOC_ID = "e3_Ae8A_gYvADACN1gvbrUgvT3SalKT9"
PK_SHEET_ID = "be48i0"          # 事业部PK追踪
YJ_SHEET_ID = "kf6k91"# 业绩（含续费）
IMG_OUT = os.path.join(PK_DIR, "pk_事业部PK追踪_10月.png")

# 表格底缘(实测 812)下方留一点余量，作为「没截到海报时」的退化裁剪底缘，
# 避免发出一张「表格 + 一大片空白」的图。
TABLE_ONLY_BOTTOM = 900

PK_LINK = ("https://doc.weixin.qq.com/sheet/e3_Ae8A_gYvADACN1gvbrUgvT3SalKT9"
           "?scode=ADwAKQdcAAcl4y7c4pAaIAAQbBAM8&tab=be48i0")

# 业绩目标线（汇报口径）。改这里即可切换目标，如 0.65 = 65%。
# 注意：PK sheet 的 J/K 列仍是表内按 60% 写的公式，不能直接引用，
#故本脚本一律用 C 列(系统目标) × TARGET_RATE - D 列(完成) 自行重算差距。
TARGET_RATE = 0.65

EXIT_NEED_LOGIN = 3              # 需人工扫码登录，未截图未发送

EXPORTER_DIR = r"<PROJECT_ROOT>/数据文件/20260924_tencent-doc-export"
sys.path.insert(0, os.path.join(EXPORTER_DIR, "src"))
sys.path.insert(0, SYNC_DIR)
import sync_yj  # noqa: E402


class NeedLogin(RuntimeError):
    """企微文档登录态不可用 —— 需人工扫码，本次不截图不发送。"""


# ------------------------- 第 0 步：登录态预检 -------------------------

def check_login():
    """截图前预检企微文档登录态。返回 (ok: bool, msg: str)。

    登录态不可用时抛 NeedLogin —— 调用方必须中止且**不发送**任何群消息，
    等用户扫码登录后再重跑。
    """
    import core

    edge = core.find_edge()
    if not edge:
        raise NeedLogin("未找到 Edge/Chrome 浏览器，无法截图。"
                        "请确认浏览器已安装到标准路径。")

    if not core.login_state_exists():
        raise NeedLogin(f"登录态文件不存在：{core.state_file()}。"
                        "请先扫码登录（运行 export_sheet.py 或用「腾讯文档导出工具」扫码）。")

    # 用 dop-api 探测登录态真伪（storage_state 存在 ≠ Cookie 还有效）
    import asyncio
    from playwright.async_api import async_playwright

    async def _probe():
        info = core.parse_link(PK_LINK)
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                executable_path=edge, headless=True,
                args=["--disable-blink-features=AutomationControlled"])
            try:
                ctx = await browser.new_context(
                    storage_state=str(core.state_file()), user_agent=core.UA,
                    viewport={"width": 1280, "height": 800})
                page = await ctx.new_page()
                await page.goto(PK_LINK, wait_until="domcontentloaded", timeout=45000)
                await page.wait_for_timeout(2500)
                return await core._check_login(page, core._api_url(info))
            finally:
                await browser.close()

    try:
        ok = asyncio.run(_probe())
    except Exception as e:
        raise NeedLogin(f"登录态探测失败：{type(e).__name__}: {e}。"
                        "建议重新扫码登录。")

    if not ok:
        raise NeedLogin(
            f"登录态已失效（{core.login_expired_hint()}）。"
            "请重新扫码登录后再重跑本脚本。")

    return True, core.login_expired_hint()


# ------------------------- 第 1 步：业绩同步 -------------------------

def run_sync():
    sync_yj.acquire_lock()
    try:
        return sync_yj.one_cycle(notify=False)
    finally:
        if os.path.exists(sync_yj.LOCK_FILE):
            os.remove(sync_yj.LOCK_FILE)


# ------------------------- 第 2 步：截图 + 裁剪 -------------------------

def capture_raw():
    """复用 capture_pk2.py 整页截图 → pk_raw2.png"""
    r = subprocess.run([PY_EXE, os.path.join(SCRATCH_DIR, "capture_pk2.py")],
                       capture_output=True, timeout=180)
    out = r.stdout.decode("utf-8", errors="replace")
    if r.returncode != 0 or "NO_EDGE_OR_NO_LOGIN" in out:
        raise NeedLogin(f"截图失败（登录态不可用）：{out[-300:]} "
                        f"{r.stderr.decode('utf-8', 'replace')[-300:]}")
    print(out.strip())


def crop_with_posters(raw_path, table_top=261, table_bottom=812, right=2400):
    """裁剪 = 表格区 + 下方两张活动海报（全宽2400，保右海报完整；纯表格版才用2245）。
    海报底缘检测：从表格底(812)向下扫描到倒数 100px（排除底部标签栏），
    找最后一条「高度≥15px 的宽幅非白色带」（孤立 1px 噪声行/小色块不算），
    底缘 + 20px 余量。

    返回 (bottom, poster_found)：
      poster_found=False 表示表格下方**一条海报带都没检测到**。实测这几乎都是
      「海报浮动图片还没渲染就截图」造成的（2026-10-08 14:35 漏图事故：raw 328KB
      空白 vs 正常 601KB 有图），此时调用方必须重截，而不是静默套一个回退值裁出
      「表格 + 大片空白」就发出去。退化时只裁表格区(到 TABLE_ONLY_BOTTOM)。"""
    from PIL import Image
    img = Image.open(raw_path).convert("RGB")
    W, H = img.size
    px = img.load()
    bottom_limit = H - 100

    def row_on(y):
        cnt = sum(1 for x in range(0, W, 2)
                  if (lambda p: (p[0] + p[1] + p[2]) < 700
                      or max(p) - min(p) > 25)(px[x, y]))
        return cnt > 100

    bands, prev, start = [], False, None
    for y in range(table_bottom + 8, bottom_limit):
        on = row_on(y)
        if on and not prev:
            start = y
        if not on and prev:
            bands.append((start, y - 1))
        prev = on
    if prev:
        bands.append((start, bottom_limit - 1))
    big = [b for b in bands if b[1] - b[0] + 1 >= 15]
    poster_found = bool(big)
    bottom = (big[-1][1] + 20) if big else TABLE_ONLY_BOTTOM
    bottom = min(bottom, bottom_limit)
    img.crop((0, table_top, right, bottom)).save(IMG_OUT)
    print(f"裁剪完成: (0,{table_top},{right},{bottom}) → {IMG_OUT} "
          f"({right}x{bottom - table_top}) "
          f"{'含活动海报 ✅' if poster_found else '⚠ 未检测到活动海报'}")
    return bottom, poster_found


def capture_and_crop(tries=3):
    """截图 + 裁剪，并**校验表格下方确实截到了两张活动海报**。

    海报是浮动图片对象，页面刚打开（尤其刚同步完在线表）时可能还没渲染：
    实测 2026-10-08 14:35 那次截图 raw=327,963B（海报区全空白），
    同一脚本重跑一次 raw=601,220B（海报正常）——属渲染时序问题，重开页面即可恢复。
    所以这里检测不到海报就重截，最多 tries 次。

    连续失败**不中止发送**：战报正文与表格仍是完整的，只缺海报，
    故退化为「表格版」继续发（并发声告警），避免因缺图导致整份战报发不出去。
    返回裁剪底缘 bottom，或 None（一直没截到海报）。
    """
    raw = os.path.join(SCRATCH_DIR, "pk_raw2.png")
    for i in range(1, tries + 1):
        capture_raw()
        bottom, poster_found = crop_with_posters(raw)
        if poster_found:
            return bottom
        print(f"[WARN] 第 {i}/{tries} 次截图：表格下方未检测到活动海报"
              f"（多为海报尚未渲染），2s 后重新截图…")
        time.sleep(2)
    print(f"[WARN] 连续 {tries} 次都没截到活动海报 —— 仍继续发送（仅表格），"
          f"请人工确认在线表里的两张活动海报是否还在。")
    return None


def verify_shot_is_table(path):
    """校验截图里真有表格，而不是登录二维码/登录页。阈值来自实测样本：

        【坏】登录二维码页  整体白占比 14% / 11%
        【好】真实 PK 表格   整体白占比 55% / 61% / 69%
    取 35% 阈值，分离度40 个百分点，远大于噪声。
    只用这一条特征，不加「中央白块」副规则——表格局部放大图中央本就是白底单元格，
    加了会误杀（pk_table_zoom 实测中央 87% 却是真表格）。

    背景：登录态失效时页面渲染成「企业微信扫码登录」页，而该页 title 同样是
    「企业微信文档」，旧的 capture_pk2.py 只看title 非空 → 会把二维码当战报发群。
    此处按画面特征判定，与登录态预检（探 Cookie 真伪）形成两道互补防线。
    """
    from PIL import Image
    img = Image.open(path).convert("L")
    W, H = img.size
    px = img.load()

    x0, x1 = int(W * 0.02), int(W * 0.98)
    y0, y1 = int(H * 0.05), int(H * 0.95)
    tot = white = 0
    for y in range(y0, y1, 2):
        for x in range(x0, x1, 2):
            tot += 1
            if px[x, y] > 245:
                white += 1
    overall = (white / tot if tot else 0)

    # 表格 = 大面积白底 + 文字；登录页 = 灰底 + 中央一个白框二维码
    if overall < 0.35:
        raise NeedLogin(
            f"截图内容是登录页/二维码，不是表格"
            f"（整体白占比 {overall:.0%}，正常表格应≥35%）。"
            "登录态已失效，请重新扫码登录。")

    print(f"画面校验通过：确为表格（整体白占比 {overall:.0%}）")


# ------------------------- 第 3 步：读数据生成文案 -------------------------

def _num(v):
    try:
        return float(v)
    except Exception:
        return None


def _read_range(sheet_id, rng):
    d = sync_yj.wecom_call(["sheet", "ranges", "get"], payload={
        "docid": DOC_ID, "sheet_id": sheet_id, "mode": "default", "range": rng},
        timeout=180)
    rows = []
    for row in d.get("grid_data", {}).get("rows", []):
        vals = []
        for c in row.get("values", []):
            cvv = c.get("cell_value") or {}
            if "text" in cvv and str(cvv["text"]).strip():
                vals.append(str(cvv["text"]).strip())
            elif "number" in cvv:
                n = cvv["number"]
                vals.append(str(int(n)) if float(n).is_integer() else str(n))
            else:
                vals.append("")
        rows.append(vals)
    return rows


def fmt_amt(v):
    n = _num(v)
    return f"{n:,.0f}" if n is not None else str(v)


def fmt_pct(v):
    n = _num(v)
    return f"{n * 100:.2f}%" if n is not None else str(v)


def build_text():
    rows = _read_range(PK_SHEET_ID, "A1:P9")
    if len(rows) < 9:
        raise RuntimeError(f"PK sheet 读取行数不足: {len(rows)}")

    depts, branch = [], None
    for vals in rows[1:]:  # R2 起
        name = (vals[0] if len(vals) > 0 else "").strip()
        if not name:
            continue
        rec = {i: (vals[i] if i < len(vals) else "") for i in range(10)}
        if name == "分公司":
            branch = rec
        elif name != "知识产权事业部" and rec[3] not in ("-", ""):
            depts.append(rec)

    if not depts or not branch:
        raise RuntimeError("PK sheet 解析失败：事业部或分公司行缺失")

    # 数据时间：取【业绩（含续费）】R2，去掉「数据周期：」前缀
    yj_rows = _read_range(YJ_SHEET_ID, "A2:A2")
    data_time = ""
    if yj_rows and yj_rows[0]:
        data_time = yj_rows[0][0].replace("数据周期：", "").strip()

    depts.sort(key=lambda d: _num(d[4]) or 0, reverse=True)
    icons = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣"]
    pct_label = f"{TARGET_RATE * 100:.0f}%"
    lines = ["📊 分公司事业部PK追踪 · 战报", ""]
    if data_time:
        lines += [f"📅 数据时间：{data_time}", ""]
    lines += [f"🏆 完成率 & 距 {pct_label} 目标差距"]

    def gap_of(rec):
        """按TARGET_RATE 自算差距：系统目标×目标率 - 已完成。
        表内 J/K 列仍是 60% 公式，不能直接引用。"""
        g, dn = _num(rec[2]), _num(rec[3])
        if g is None or dn is None:
            return None
        return g * TARGET_RATE - dn

    for i, d in enumerate(depts):
        g = gap_of(d)
        gap_txt = f"｜差 {fmt_amt(g)}" if g is not None else ""
        lines.append(f"{icons[i]} {d[0]} · {d[1]}｜{fmt_pct(d[4])}"
                     f"（{fmt_amt(d[3])} / {fmt_amt(d[2])}）{gap_txt}")

    b_goal, b_done = _num(branch[2]), _num(branch[3])
    b_line = b_goal * TARGET_RATE if b_goal is not None else None
    gap = b_line - b_done if (b_line is not None and b_done is not None) else None
    lines += ["", "🧮 分公司（汇总）",
              f"· 完成率 {fmt_pct(branch[4])}（{fmt_amt(branch[3])} / {fmt_amt(branch[2])}）"]
    if gap is not None:
        lines.append(f"· 距 {pct_label} 目标差额约 {gap:,.0f}"
                     f"（{pct_label} 目标线 {b_line:,.0f}）")
    return "\n".join(lines)


# ------------------------- 第 4 步：发送 -------------------------

def send_all(group_keys, text):
    d = sync_yj.wecom_call(["message", "aibot", "sessions", "list"])
    sessions = d.get("sessions", [])
    sent = []
    for key in group_keys:
        target = next((s for s in sessions if key in str(s.get("chat_name", ""))), None)
        if not target:
            print(f"[WARN] 未找到群「{key}」，跳过")
            continue
        up = sync_yj.wecom_call(["media", "upload"], payload={
            "file_path": IMG_OUT, "type": "image"}, timeout=180)
        mid = up.get("media_id") if isinstance(up, dict) else None
        sync_yj.wecom_call(["message", "aibot", "send"], payload={
            "chat_id": target["chat_id"], "msgtype": "markdown",
            "markdown": {"content": text}})
        if mid:
            sync_yj.wecom_call(["message", "aibot", "send"], payload={
                "chat_id": target["chat_id"], "msgtype": "image",
                "image": {"media_id": mid}})
        sent.append(target["chat_name"])
        print(f"已发送 → {target['chat_name']} ✅ (media_id={bool(mid)})")
        time.sleep(0.5)
    return sent


# ------------------------- 主流程 -------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="无视变化检测，强制发送")
    ap.add_argument("--group", action="append", default=None,
                    help="目标群关键字（可多次传；默认 3C认证团队）")
    args = ap.parse_args()
    groups = args.group if args.group else ["3C认证"]

    t0 = time.time()
    try:
        # 第 0 步：登录态预检 —— 放在同步之前。
        # 若登录态失效，此时数据尚未同步，下次重跑仍是 SYNCED，不会漏发战报。
        print("=== 0) 登录态预检（截图前置条件）===")
        ok, msg = check_login()
        print(f"登录态: OK（{msg}）")
    except NeedLogin as e:
        print("")
        print("=" * 60)
        print("❌ 需要扫码登录 —— 本次已中止，未截图、未发送任何群消息")
        print(f"原因：{e}")
        print("")
        print("请按以下步骤处理：")
        print(f"  1) 打开链接并扫码登录：{PK_LINK}")
        print("     或运行「腾讯文档导出工具」点【扫码登录】")
        print("  2) 登录完成后重跑本脚本（数据未同步过，不会漏发）")
        print("=" * 60)
        sys.exit(EXIT_NEED_LOGIN)

    try:
        print("=== 1) 业绩同步（静默模式）===")
        status = run_sync()
        print(f"同步状态: {status}")

        if status == "SKIP" and not args.force:
            print("无数据变化 → 跳过截图与发送（省算力）")
            return
        if status not in ("SYNCED",) and not args.force:
            print(f"同步状态异常({status}) → 不发送")
            return

        print("=== 2) 截图 + 裁剪（含两张活动海报）===")
        capture_and_crop()                # 内含海报就绪校验 + 未截到就重截
        verify_shot_is_table(IMG_OUT)     # 发送前最后一道闸：图里必须是表格

        print("=== 3) 生成战报文案 ===")
        text = build_text()
        print(text)

        print("=== 4) 发送 ===")
        sent = send_all(groups, text)
        print(f"=== 完成: {status} → 发 {len(sent)} 个群, 耗时 {time.time() - t0:.0f}s ===")
    except NeedLogin as e:
        # 同步已完成但截图阶段登录态失效：明确告知，且不发送。
        print("")
        print("=" * 60)
        print("❌ 截图阶段登录态失效 —— 已中止，未发送任何群消息")
        print(f"原因：{e}")
        print(f"注意：本次数据已同步，下次重跑可能判定为「无变化」而跳过。")
        print(f"如需立即出图，请扫码登录后用 --force 重跑：")
        print(f"  {PK_LINK}")
        print("=" * 60)
        sys.exit(EXIT_NEED_LOGIN)


if __name__ == "__main__":
    main()
