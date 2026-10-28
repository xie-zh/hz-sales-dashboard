"""core.py —— 核心导出逻辑(无 GUI 依赖,便于单测与打包)

职责:
1. 解析分享链接(单条/批量)
2. Playwright 打开文档(复用系统 Edge,优先 headless)
3. 多策略导出: D) 接口直取(免验证码,V0.1.3 起优先)
   → A) UI 菜单导出 → B) dop-api 数据解析 → C) 人工滑块后重导
4. 批量导出(多链接并发,按文档名自动命名,V0.1.3)
5. 定时导出调度(纯函数,V0.1.3)
6. 替换目标 xlsx 的 sheet 源(含外部引用刷新)
"""

import asyncio
import base64
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse, parse_qs

from version import AUTHOR  # V0.1.1: 导出 xlsx 文件属性署名(隐藏水印)

# ---- 常量 ----
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36 Edg/125.0.0.0")
EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
LOGIN_TIMEOUT_S = 300          # 扫码等待 5 分钟
PAGE_TIMEOUT_S = 60


def app_dir() -> Path:
    """exe/脚本所在目录(打包后为 exe 同目录;开发时为项目根)。"""
    if getattr(sys, "frozen", False):  # PyInstaller
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def data_dir() -> Path:
    """登录态等数据目录: exe 旁边的 .tdexporter 目录(绿色便携)。"""
    d = app_dir() / ".tdexporter"
    d.mkdir(parents=True, exist_ok=True)
    return d


def state_file() -> Path:
    return data_dir() / "auth_state.json"


def find_edge() -> str | None:
    for p in EDGE_PATHS:
        if os.path.exists(p):
            return p
    return None


def parse_link(link: str) -> dict:
    """解析分享链接 → {id, scode, tab, url}。不合法抛 ValueError。"""
    link = (link or "").strip()
    if not link:
        raise ValueError("链接为空")
    if not link.startswith("http"):
        link = "https://" + link
    u = urlparse(link)
    host = u.netloc.lower()
    if not ("weixin.qq.com" in host or "docs.qq.com" in host):
        raise ValueError(f"不是腾讯文档/企微文档链接: {host}")
    segments = [seg for seg in u.path.split("/") if seg]
    doc_id = segments[-1] if segments else ""
    if not doc_id:
        raise ValueError("链接里没有文档 ID")
    qs = parse_qs(u.query)
    scode = (qs.get("scode") or [""])[0]
    tab = (qs.get("tab") or [""])[0]
    return {"id": doc_id, "scode": scode, "tab": tab, "url": link}


# ================= 登录态管理 =================

def login_state_exists() -> bool:
    return state_file().exists()


def login_expired_hint() -> str:
    if not login_state_exists():
        return "未登录"
    ts = state_file().stat().st_mtime
    age_days = (time.time() - ts) / 86400
    if age_days > 30:
        return f"登录态较旧({age_days:.0f} 天前),建议重新扫码"
    return f"已登录({age_days:.1f} 天前扫码)"


def _api_url(info: dict) -> str:
    return ("https://doc.weixin.qq.com/dop-api/opendoc"
            f"?u=&id={info['id']}&normal=1&outformat=1&scode={info['scode']}"
            f"&noEscape=1&commandsFormat=1&tab={info['tab']}")


async def _check_login(page, api_url: str) -> bool:
    """True=登录态有效。以 dop-api 探测为准(匿名 Cookie 也会有 TOK,不可靠)。

    页面/浏览器已关闭等调用异常 → False,由调用方兜底处理。
    """
    try:
        res = await page.evaluate(
            """async (url) => {
                try {
                    const resp = await fetch(url, { credentials: 'include' });
                    const text = await resp.text();
                    return !(text.includes('check userInfo failed')
                          || text.includes('"retcode":30004')
                          || text.includes('"errcode":"30004'));
                } catch (e) { return false; }
            }""", api_url)
    except Exception:
        return False
    return bool(res)


def login_cookie_names() -> set:
    """腾讯文档登录态特征 Cookie 名集合(用于状态展示,非判定依据)。"""
    return {"TOK", "ES2", "UIN", "vid", "open_id", "ww_code"}


async def interactive_login(link: str, log=print, headless: bool = False) -> bool:
    """弹浏览器扫码,成功后保存登录态。返回是否成功。

    link 可以是具体文档链接,也可以是根地址(无文档 ID,仅用于扫码)。
    """
    from playwright.async_api import async_playwright

    try:
        info = parse_link(link)
        url, probe_url = info["url"], _api_url(info)
    except ValueError:
        url = link if link.startswith("http") else "https://doc.weixin.qq.com/"
        # 根页无文档 ID: 用该域名已知文档作探测地址,仅判断登录,不导出
        probe_info = {"id": "e3_AfUA6QYZAA4CNPM3wfpScRMemcm20",
                      "scode": "ADwAKQdcAAcCYu0lhDAaIAAQbBAM8", "tab": "BB08J2"}
        probe_url = _api_url(probe_info)
    edge = find_edge()
    if not edge:
        raise RuntimeError("未找到 Edge/Chrome 浏览器,请先安装 Microsoft Edge")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=edge, headless=headless,
            args=["--disable-blink-features=AutomationControlled"])
        try:
            context = await browser.new_context(
                viewport={"width": 1440, "height": 900}, user_agent=UA)
            page = await context.new_page()
            log("[登录] 打开文档页面…")
            await page.goto(url, wait_until="domcontentloaded",
                            timeout=PAGE_TIMEOUT_S * 1000)

            ok = await _check_login(page, probe_url)
            if ok:  # 已有登录态(理论上首次不会)
                await context.storage_state(path=str(state_file()))
                await browser.close()
                return True

            log("[登录] 请在弹出的浏览器中扫码登录(最长等待 5 分钟)…")
            deadline = time.time() + LOGIN_TIMEOUT_S
            ok = False
            while time.time() < deadline:
                ok = await _check_login(page, probe_url)
                if ok:
                    break
                if not browser.is_connected():
                    log("[登录] 浏览器被关闭,未完成扫码。")
                    return False
                try:
                    await page.wait_for_timeout(3000)
                except Exception:
                    if not browser.is_connected():
                        log("[登录] 浏览器被关闭,未完成扫码。")
                        return False

            if not ok:
                log("[登录] 超时未检测到登录成功。")
                try:
                    await browser.close()
                except Exception:
                    pass
                return False

            await context.storage_state(path=str(state_file()))
            log(f"[登录] 登录态已保存: {state_file()}")
            try:
                await browser.close()
            except Exception:
                pass
            return True
        except Exception:
            try:
                await browser.close()
            except Exception:
                pass
            raise


# ================= V0.1.3: 直取链 / 批量 / 定时 =================

# 阻断的冗余资源后缀(图片/字体/媒体): 直取与加载提速,不影响 canvas 表格
BLOCK_EXT = ("png", "jpg", "jpeg", "gif", "svg", "webp", "ico", "bmp",
             "woff", "woff2", "ttf", "otf", "eot", "mp4", "webm", "mp3")


async def _block_resources(context) -> None:
    """V0.1.3 性能优化: context 级阻断图片/字体/媒体等冗余资源下载。"""
    async def _route(route):
        try:
            url = route.request.url.lower()
            if any(url.endswith("." + e) or ("." + e + "?") in url
                   for e in BLOCK_EXT):
                await route.abort()
                return
        except Exception:
            pass
        try:
            await route.continue_()
        except Exception:
            pass

    await context.route("**/*", _route)


def _capture_sid(page) -> str:
    """V0.1.3: 从已监听的页面请求中拿 sid(同步读,配合 install_sid_listener)。
    返回 '' 表示尚未捕获。"""
    holder = getattr(page, "_sid_holder", None)
    return holder["sid"] if holder else ""


def _install_sid_listener(page) -> None:
    """V0.1.3: 安装请求监听捕获 sid(页面常规 API 请求 URL 均带 sid= 参数,
    如 /cgi/grayconfig/queryKeys?sid=xxx&wedoc_xsrf=1;wxf=1 时为空需跳过)。"""
    holder = {"sid": ""}
    page._sid_holder = holder

    def on_req(req):
        try:
            if holder["sid"]:
                return
            u = req.url
            if "sid=" in u:
                m = re.search(r"[?&]sid=([^&]+)", u)
                if m and m.group(1):
                    holder["sid"] = m.group(1)
        except Exception:
            pass

    try:
        page.on("request", on_req)
    except Exception:
        pass


class _DirectFailedError(RuntimeError):
    """V0.1.3: 策略 D(接口直取)失败,由 export_sheet 捕获后降级回 UI 链。"""


def _doc_version(info: dict) -> str:
    """V0.1.3: sheet 文档导出类型码——doc_id 前缀决定 version:
    e3_(智能表格)=2;其他前缀默认 2(导出链对该类文档稳定适用)。"""
    return "2"


async def _export_direct(page, info: dict, out_file: str,
                         log=print) -> str:
    """V0.1.3 策略 D: 已登录页面上下文直取导出接口(免验证码)。

    链路: getbannerinfo 前置(真实前端导出前的场景声明)
    → POST /v1/export/export_office (form: docId & version,
      credentials include, 不带 captchaTicket)
    → 轮询 /v1/export/query_progress 拿 file_url(COS 签名直链)
    → Python urllib 下载 → PK 魔数校验 → 水印。
    任一步 ret≠0(如 520112 captcha required)抛 _DirectFailedError,
    由调用方自动降级回原 UI 菜单导出方案并打印降级提醒。

    返回服务端建议文件名(file_name,如"每周待办.xlsx"),可用于自动命名。
    """
    sid = _capture_sid(page)
    if not sid:
        raise _DirectFailedError("sid 未捕获")

    out = Path(out_file)
    out.parent.mkdir(parents=True, exist_ok=True)

    # ① getbannerinfo 前置(0.3s,真实前端导出前必调,声明验证码场景)
    try:
        ban = await page.evaluate(
            """async (sid) => {
                const url = '/disk/getbannerinfo?func=3&captcha_sence=1'
                          + '&sid=' + sid;
                const r = await fetch(url, {credentials: 'include'});
                return {status: r.status, text: (await r.text()).slice(0, 400)};
            }""", sid)
    except Exception as e:
        raise _DirectFailedError(f"getbannerinfo 异常: {e}")

    # ② export_office(form 编码,不带 ticket)
    try:
        res = await page.evaluate(
            """async ([sid, docId, version]) => {
                const url = '/v1/export/export_office?sid=' + sid +
                            '&wedoc_xsrf=1';
                const body = 'docId=' + encodeURIComponent(docId) +
                             '&version=' + version;
                const r = await fetch(url, {
                    method: 'POST',
                    credentials: 'include',
                    headers: {'Content-Type':
                              'application/x-www-form-urlencoded'},
                    body: body});
                const text = await r.text();
                return {status: r.status, text: text.slice(0, 1500)};
            }""", [sid, info["id"], _doc_version(info)])
    except Exception as e:
        raise _DirectFailedError(f"export_office 异常: {e}")

    try:
        data = json.loads(res["text"])
    except Exception:
        raise _DirectFailedError(
            f"export_office 响应非 JSON[{res['status']}]: {res['text'][:120]}")
    if data.get("ret") != 0 or not data.get("operationId"):
        raise _DirectFailedError(
            f"export_office ret={data.get('ret')}({data.get('msg', '')[:60]})")

    # ③ 轮询 query_progress(1s×60,Done 时含 file_url+file_name)
    oid = data["operationId"]
    file_url, file_name = "", ""
    for i in range(60):
        try:
            pr = await page.evaluate(
                """async ([oid]) => {
                    const url = '/v1/export/query_progress?operationId='
                        + oid + '&timestamp=' + Date.now();
                    const r = await fetch(url, {credentials: 'include'});
                    return await r.text();
                }""", [oid])
            d = json.loads(pr) if pr.startswith("{") else {}
        except Exception as e:
            if i >= 3:
                raise _DirectFailedError(f"query_progress 异常: {e}")
            await page.wait_for_timeout(1000)
            continue
        if d.get("ret") != 0:
            raise _DirectFailedError(
                f"query_progress ret={d.get('ret')}({d.get('msg', '')[:60]})")
        if d.get("status") == "Done" and d.get("file_url"):
            file_url = d["file_url"]
            file_name = d.get("file_name", "")
            break
        await page.wait_for_timeout(1000)
    if not file_url:
        raise _DirectFailedError("轮询 60s 未拿到 file_url")

    # ④ Python 下载 COS 签名直链(带 UA,30 分钟签名有效)
    try:
        req = urllib.request.Request(file_url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=120) as resp:
            blob = resp.read()
    except Exception as e:
        raise _DirectFailedError(f"COS 下载异常: {e}")
    if not blob or blob[:2] != b"PK":
        raise _DirectFailedError(
            f"下载内容不是 xlsx(PK 校验失败,{len(blob)} bytes)")
    out.write_bytes(blob)

    stamp_xlsx_author(str(out), log=log)  # V0.1.1 水印
    return file_name or ""


def parse_links(text: str) -> list:
    """V0.1.3: 批量粘贴框文本(一行一链接)→ [parse_link 结果,...]。
    忽略空行/注释行;某行非法不中断整体,由返回值外的 errors 报告。

    返回 (links, errors): links=合法 info 列表,
    errors=[(行号, 原文, 错误信息),...]。
    """
    links, errors = [], []
    for ln, raw in enumerate((text or "").splitlines(), start=1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        try:
            links.append(parse_link(line))
        except ValueError as e:
            errors.append((ln, line, str(e)))
    return links, errors


_INVALID_FS = re.compile(r'[\\/:*?"<>|]')


def sanitize_filename(name: str) -> str:
    """V0.1.3: 文档名 → 合法文件名(去非法字符/控制符/首尾空格点,限 80 字符)。
    空结果返回空串(调用方用 doc_id 兜底)。"""
    name = (name or "").strip()
    name = _INVALID_FS.sub("_", name)
    name = re.sub(r"[\x00-\x1f]", "", name).strip().strip(".")
    if len(name) > 80:
        name = name[:80]
    return name


def unique_path(directory: Path, stem: str, suffix: str = ".xlsx") -> Path:
    """V0.1.3: 目标目录下重名自动加序号: a.xlsx → a (2).xlsx → a (3).xlsx。"""
    p = directory / f"{stem}{suffix}"
    n = 2
    while p.exists():
        p = directory / f"{stem} ({n}){suffix}"
        n += 1
    return p


async def _doc_title(info: dict, log=print) -> str:
    """V0.1.3: 无头+阻断资源快速拿文档标题(批量自动命名用,~2s/链接)。
    失败返回空串(调用方用 doc_id 兜底)。"""
    from playwright.async_api import async_playwright

    edge = find_edge()
    if not edge or not login_state_exists():
        return ""
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=edge, headless=True,
            args=["--disable-blink-features=AutomationControlled"])
        try:
            context = await browser.new_context(
                storage_state=str(state_file()), user_agent=UA,
                viewport={"width": 1440, "height": 900})
            await _block_resources(context)
            page = await context.new_page()
            await page.goto(info["url"], wait_until="domcontentloaded",
                            timeout=PAGE_TIMEOUT_S * 1000)
            for _ in range(10):
                try:
                    t = (await page.title() or "").strip()
                except Exception:
                    t = ""
                if t and "腾讯文档" not in t[:3]:
                    return t
                await page.wait_for_timeout(500)
            try:
                return (await page.title() or "").strip()
            except Exception:
                return ""
        except Exception:
            return ""
        finally:
            try:
                await browser.close()
            except Exception:
                pass


async def export_batch(links_text: str, out_dir: str, log=print,
                       headless: bool = True, concurrency: int = 3) -> dict:
    """V0.1.3: 多链接批量导出(2~3 并发,按文档名自动命名存目标目录)。

    - 一行一链接;无效行逐条报错不影响其他行
    - 文件名 = 文档标题清洗(兜底 doc_id);重名自动加 (2)(3)
    - 每链接独立 browser 实例(策略 D 直取优先→A→B 降级;
      无人值守场景跳过 C 人工滑块,失败记入结果继续下一链接)
    - 返回 {"ok": 成功数, "fail": 失败数, "files": [路径,...],
      "errors": [(链接, 错误),...]}
    """
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    links, line_errors = parse_links(links_text)
    results = {"ok": 0, "fail": 0, "files": [], "errors": []}
    for ln, raw, msg in line_errors:
        log(f"[批量] 第 {ln} 行不是有效链接,已跳过: {raw[:50]}({msg})")
        results["fail"] += 1
        results["errors"].append((raw, msg))
    if not links:
        return results

    sem = asyncio.Semaphore(max(1, min(concurrency, 3)))
    used_names = set()

    async def one(info: dict, idx: int) -> None:
        tag = f"[批量 {idx}/{len(links)}]"
        async with sem:
            try:
                title = await _doc_title(info, log=log)
            except Exception:
                title = ""
            stem = sanitize_filename(title) or sanitize_filename(info["id"])
            base = stem
            n = 2
            while base.lower() in used_names:
                base = f"{stem} ({n})"
                n += 1
            used_names.add(base.lower())
            target = unique_path(out, base)
            log(f"{tag} 开始: {info['id']} → {target.name}")
            try:
                await export_sheet(info["url"], str(target), log=log,
                                   headless=headless, manual_captcha=False)
                results["ok"] += 1
                results["files"].append(str(target))
                log(f"{tag} 完成 ✓ {target.name}")
            except Exception as e:
                results["fail"] += 1
                results["errors"].append((info["url"], f"{e}"))
                log(f"{tag} 失败 ✗ {info['id']}: {e}")

    await asyncio.gather(*(one(info, i + 1)
                           for i, info in enumerate(links)))
    log(f"[批量] 全部完成: 成功 {results['ok']}, 失败 {results['fail']}"
        f"(目录: {out})")
    return results


# ---------- 定时导出(纯函数,GUI after 轮询调用) ----------

SCHEDULE_FILE = "schedule.json"


def schedule_state_file() -> Path:
    return data_dir() / SCHEDULE_FILE


def load_schedule_state() -> dict:
    """V0.1.3: 读 .tdexporter/schedule.json({} 容错)。"""
    try:
        return json.loads(
            schedule_state_file().read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_schedule_state(state: dict) -> None:
    """V0.1.3: 写定时任务状态(next_run ISO 字符串等)。"""
    schedule_state_file().parent.mkdir(parents=True, exist_ok=True)
    schedule_state_file().write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def compute_next_run(cfg: dict, now: datetime) -> "datetime | None":
    """V0.1.3: 计算下次运行时刻。

    cfg: {"mode": "off"|"interval"|"daily"|"once",
          "interval_hours": int(>0),
          "daily_hhmm": "HH:MM",
          "once_dt": iso 字符串或 datetime}
    - interval: now + N 小时(从启用/上次触发点起算)
    - daily: 今天 HH:MM(已过则明天)
    - once: 指定时刻(已过则返回 None,由错过策略处理)
    - off/非法: None
    """
    mode = (cfg or {}).get("mode", "off")
    if mode == "off":
        return None
    if mode == "interval":
        try:
            hours = max(1, int(cfg.get("interval_hours", 0)))
        except (TypeError, ValueError):
            return None
        return now + timedelta(hours=hours)
    if mode == "daily":
        hhmm = str(cfg.get("daily_hhmm", "09:00"))
        m = re.match(r"^(\d{1,2}):(\d{1,2})$", hhmm.strip())
        if not m:
            return None
        hh, mm = int(m.group(1)), int(m.group(2))
        if not (0 <= hh < 24 and 0 <= mm < 60):
            return None
        nxt = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
        if nxt <= now:
            nxt += timedelta(days=1)
        return nxt
    if mode == "once":
        raw = cfg.get("once_dt")
        if isinstance(raw, datetime):
            return raw
        try:
            return datetime.fromisoformat(str(raw))
        except (TypeError, ValueError):
            return None
    return None


def schedule_should_fire(state: dict, now: datetime) -> bool:
    """V0.1.3: 到点判定: state 含未过期的 next_run 且 now ≥ next_run。"""
    nr = state.get("next_run")
    if not nr:
        return False
    try:
        nxt = datetime.fromisoformat(str(nr))
    except (TypeError, ValueError):
        return False
    return now >= nxt


# ================= 导出 =================

async def export_sheet(link: str, out_xlsx: str, log=print,
                       headless: bool = True,
                       manual_captcha: bool = True) -> str:
    """导出在线表格 → out_xlsx。返回实际使用的策略名。异常向上抛。

    V0.1.3 多策略链(速度优先):
    D) 接口直取(免验证码: 已登录页面上下文直调导出接口,跳过 UI 流程)
       → 失败自动降级回原 UI 导出方案,日志打印降级提醒;
    A) UI 菜单导出 → B) dop-api 数据解析
       → C) 人工滑块后重导(manual_captcha=False 时跳过,
           批量/定时等无人值守场景避免弹窗阻塞)。

    V0.1.3 性能优化: 复用已登录 storage_state(免扫码)、
    context 级阻断图片/字体/媒体冗余资源、压缩固定等待。
    """
    from playwright.async_api import async_playwright

    info = parse_link(link)
    if not login_state_exists():
        raise RuntimeError("尚未登录: 请先点击「扫码登录」")

    edge = find_edge()
    if not edge:
        raise RuntimeError("未找到 Edge/Chrome 浏览器")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=edge, headless=headless,
            args=["--disable-blink-features=AutomationControlled"])
        context = await browser.new_context(
            storage_state=str(state_file()), user_agent=UA,
            viewport={"width": 1440, "height": 900}, accept_downloads=True)
        await _block_resources(context)  # V0.1.3: 阻断冗余资源
        page = await context.new_page()
        _install_sid_listener(page)      # V0.1.3: 捕获 sid(直取链用)

        log("[导出] 打开文档页面…")
        await page.goto(info["url"], wait_until="domcontentloaded",
                        timeout=PAGE_TIMEOUT_S * 1000)

        # 登录态校验
        if not await _check_login(page, _api_url(info)):
            await browser.close()
            raise RuntimeError("登录态已失效: 请重新「扫码登录」")

        # ---- 策略 D: 接口直取(免验证码,V0.1.3) ----
        if headless:
            try:
                # V0.1.3 提速: sid 请求先于 canvas 渲染完成,
                # 等 sid 到手即可发起直取,不再固定等 3s 渲染
                for _ in range(15):
                    if _capture_sid(page):
                        break
                    await page.wait_for_timeout(200)
                fname = await _export_direct(page, info, out_xlsx, log=log)
                try:
                    await browser.close()
                except Exception:
                    pass
                log(f"[导出] 完成(策略 D·接口直取,免验证码): {out_xlsx}"
                    + (f" 服务端文件名: {fname}" if fname else ""))
                return "接口直取(免验证码)"
            except _DirectFailedError as e:
                log(f"[导出] 策略 D 直取失败({e})——"
                    "自动降级回原 UI 菜单导出方案…")

        # 等表格渲染(V0.1.2: 加新版智能表格的 ribbon 工具栏与文件菜单按钮)
        try:
            await page.wait_for_selector(
                ".sheet-view, .sheet-container, canvas, "
                ".ribbon-toolbar_tablist__3z3Z3, .menu-button-file",
                timeout=30000)
        except Exception:
            pass
        # V0.1.3 提速: UI 链固定等待 3000→1500(菜单渲染足够)
        await page.wait_for_timeout(1500)

        # 策略 A: UI 菜单导出
        err_a = None
        used_captcha = False
        try:
            out = Path(out_xlsx)
            out.parent.mkdir(parents=True, exist_ok=True)
            try:
                await _export_via_menu(page, str(out),
                                       headless=headless, log=log)
            except _CaptchaNeededError:
                if not headless or not manual_captcha:
                    raise  # 有头模式已等过/无人值守场景,直接报错走后续策略
                # V0.1.2b: 无头模式命中滑块验证(频繁导出风控) →
                # 关闭无头浏览器,弹可见窗口由用户手动拖滑块后重走导出链
                log("[导出] 触发滑块验证,正在弹出浏览器窗口,"
                    "请手动拖动滑块完成验证(约 2 分钟)…")
                try:
                    await browser.close()
                except Exception:
                    pass
                browser = await p.chromium.launch(
                    executable_path=edge, headless=False,
                    args=["--disable-blink-features=AutomationControlled"])
                context = await browser.new_context(
                    storage_state=str(state_file()), user_agent=UA,
                    viewport={"width": 1440, "height": 900},
                    accept_downloads=True)
                page = await context.new_page()
                log("[导出] 打开文档页面(可见窗口)…")
                await page.goto(info["url"], wait_until="domcontentloaded",
                                timeout=PAGE_TIMEOUT_S * 1000)
                try:
                    await page.wait_for_selector(
                        ".sheet-view, .sheet-container, canvas, "
                        ".ribbon-toolbar_tablist__3z3Z3, .menu-button-file",
                        timeout=30000)
                except Exception:
                    pass
                await page.wait_for_timeout(1500)
                used_captcha = True
                await _export_via_menu(page, str(out),
                                       headless=False, log=log)
            await browser.close()
            stamp_xlsx_author(str(out), log=log)  # V0.1.1 水印
            log(f"[导出] 完成(策略 A·菜单导出): {out}")
            return "UI菜单导出(滑块验证)" if used_captcha else "UI菜单导出"
        except Exception as e:
            err_a = e
            log(f"[导出] 策略 A 失败({e}),尝试策略 B(API)…")

        # 策略 B: API 解析
        err_b = None
        try:
            await _export_via_api(page, _api_url(info), out_xlsx)
            await browser.close()
            log(f"[导出] 完成(策略 B·API): {out_xlsx}")
            return "API解析"
        except Exception as e:
            err_b = e
            log(f"[导出] 策略 B 失败({e})")

        # 方案 C(V0.1.2b): A/B 均失败,大概率是滑块验证码拦截
        # (频繁导出触发风控,无头模式下无人完成滑块)。
        # 直接弹浏览器窗口,用户手动完成滑块后自动继续全流程导出。
        # V0.1.3: manual_captcha=False(批量/定时无人值守)时跳过,
        # 避免弹窗阻塞后续链接;报错注明可用单链接模式手动重试。
        if not manual_captcha:
            try:
                await browser.close()
            except Exception:
                pass
            raise RuntimeError(
                f"导出失败(D/A/B 策略均未成功,无人值守模式跳过人工滑块)。"
                f"A:{err_a}; B:{err_b}。可用单链接模式重试(支持人工滑块)。")
        try:
            try:
                await browser.close()
            except Exception:
                pass
            log("[导出] A/B 均失败,疑似滑块验证拦截——"
                "正在弹出浏览器窗口,请手动拖动滑块完成验证,"
                "验证通过后导出将自动继续…")
            browser = await p.chromium.launch(
                executable_path=edge, headless=False,
                args=["--disable-blink-features=AutomationControlled"])
            context = await browser.new_context(
                storage_state=str(state_file()), user_agent=UA,
                viewport={"width": 1440, "height": 900},
                accept_downloads=True)
            page = await context.new_page()
            log("[导出] 打开文档页面(可见窗口)…")
            await page.goto(info["url"], wait_until="domcontentloaded",
                            timeout=PAGE_TIMEOUT_S * 1000)
            try:
                await page.wait_for_selector(
                    ".sheet-view, .sheet-container, canvas, "
                    ".ribbon-toolbar_tablist__3z3Z3, .menu-button-file",
                    timeout=30000)
            except Exception:
                pass
            await page.wait_for_timeout(1500)
            out = Path(out_xlsx)
            await _export_via_menu(page, str(out), headless=False, log=log)
            await browser.close()
            stamp_xlsx_author(str(out), log=log)  # V0.1.1 水印
            log(f"[导出] 完成(方案 C·人工滑块验证后导出): {out}")
            return "UI菜单导出(人工验证)"
        except Exception as e:
            try:
                await browser.close()
            except Exception:
                pass
            raise RuntimeError(
                f"三种导出方式均失败。A:{err_a}; B:{err_b}; C(人工验证):{e}")


_LEAVES_JS = """() => {
    const out = [];
    for (const el of document.querySelectorAll('*')) {
        if (el.childElementCount > 0) continue;
        const t = (el.textContent || '').trim();
        if (!t || t.length > 60) continue;
        const r = el.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0 || r.x < -9000) continue;
        out.push({text: t, x: Math.round(r.x), y: Math.round(r.y),
                  w: Math.round(r.width), h: Math.round(r.height)});
    }
    return out;
}"""


async def _click_leaf_text(page, keywords, skip=(), min_w=0):
    """按文本定位可见叶子元素并坐标点击(先精确等值,再包含匹配)。

    V0.1.2: 新版智能表格界面控件无稳定 class,菜单项用
    「可见叶子元素文本 + getBoundingClientRect 坐标」定位。
    返回被点击的文本;找不到抛 RuntimeError。
    """
    leaves = await page.evaluate(_LEAVES_JS)
    for exact in (True, False):
        for kw in keywords:
            for it in leaves:
                t = it.get("text", "")
                if not t or it.get("w", 0) < min_w:
                    continue
                if any(s in t for s in skip):
                    continue
                if (t == kw) if exact else (kw in t):
                    x = it["x"] + it["w"] // 2
                    y = it["y"] + it["h"] // 2
                    await page.mouse.move(x, y)
                    await page.wait_for_timeout(150)
                    await page.mouse.click(x, y)
                    return t
    raise RuntimeError(f"未找到可点击项: {keywords[0]!r}")


class _CaptchaNeededError(RuntimeError):
    """V0.1.2b: 导出触发腾讯滑块验证(频繁导出风控),需人工完成。"""


_CAPTCHA_TEXTS = ("尝试太多", "刷新验证码", "安全验证", "拖动滑块",
                  "请完成验证", "人机验证", "身份验证")


async def _captcha_present(page) -> bool:
    """检测页面是否出现滑块验证码弹窗(按可见文本特征)。"""
    try:
        txt = await page.evaluate(
            "() => (document.body ? document.body.innerText : '')"
            ".slice(0, 200000)")
    except Exception:
        return False
    return isinstance(txt, str) and any(k in txt for k in _CAPTCHA_TEXTS)


async def _watch_captcha(page, out_file: str, headless: bool,
                         log=print) -> None:
    """V0.1.2b: 点击导出后监测滑块验证码(频繁导出会触发腾讯服务端风控,
    弹窗文案"尝试太多/2s后自动为您刷新验证码",网络层走 t.captcha.qq.com)。

    - 无头模式: 用户无法操作滑块 → 抛 _CaptchaNeededError,
      由 export_sheet 自动切换可见窗口重走导出链;
    - 有头模式: log 提示用户在窗口中手动拖滑块,验证码消失后返回,
      下载事件由外层 expect_download 继续等待捕获。
    首次检测到时保存现场截图 captcha.diag.png。
    """
    t0, prompted = time.time(), False
    while time.time() - t0 < 100:  # 留在 expect_download 120s 预算内
        if not await _captcha_present(page):
            if prompted:
                log("[导出] 滑块验证已通过,继续等待下载…")
            return
        if not prompted:
            prompted = True
            try:
                await page.screenshot(
                    path=str(Path(out_file).parent / "captcha.diag.png"))
            except Exception:
                pass
            if headless:
                raise _CaptchaNeededError(
                    "导出触发滑块验证(频繁导出风控),需人工完成")
            log("[导出] 检测到滑块验证码! "
                "请在弹出的浏览器窗口中拖动滑块完成验证(约 2 分钟内)…")
        await page.wait_for_timeout(2000)
    if prompted:
        raise _CaptchaNeededError("滑块验证等待超时(未完成)")


async def _export_via_menu(page, out_file: str, headless: bool = True,
                           log=print):
    """策略 A: 模拟 文件菜单→导出→导出Excel表格(.xlsx),捕获下载。

    V0.1.2 适配新版智能表格(smartsheet): 文件菜单按钮是纯图标
    (.menu-button-file,无文本),菜单项按叶子文本+坐标点击;
    「导出」会展开格式子菜单,再点「导出Excel表格 (.xlsx)」触发下载。
    旧版界面(文本按钮/直接导出项)由兜底分支兼容。

    V0.1.2a: 失败时保存现场截图与叶子元素清单到导出目录旁的
    *.diag.png/*.diag.json,便于偶发失败定位(下载超时/菜单未展开)。
    """
    _diag_dir = Path(out_file).parent

    def _snap(tag: str):
        try:
            asyncio.get_running_loop().create_task(
                page.screenshot(path=str(_diag_dir / f"{tag}.diag.png")))
        except Exception:
            pass

    # 1) 打开文件菜单(带重试: 首次点击可能被页面状态吞掉)
    menu_btn_sels = ('.menu-button-file', '[class*="menu-button-file"]',
                     '.more-menu-btn', '[class*="menu"][class*="more"]',
                     '.title-bar .menu-btn')
    opened = False
    for _attempt in range(3):
        clicked_btn = False
        for sel in menu_btn_sels:
            try:
                btn = page.locator(sel).first
                if await btn.count() > 0 and await btn.is_visible():
                    await btn.click(timeout=5000)
                    clicked_btn = True
                    break
            except Exception:
                continue
        if not clicked_btn:
            try:  # 旧版界面: 文本菜单按钮
                await _click_leaf_text(page, ("文件", "菜单"))
                clicked_btn = True
            except RuntimeError:
                pass
        if not clicked_btn:
            _snap("menu_button_missing")
            raise RuntimeError("未找到文件菜单按钮(工具栏未就绪?)")
        await page.wait_for_timeout(1800)
        # 验证菜单已展开(出现常见菜单项文本)
        try:
            leaves = await page.evaluate(_LEAVES_JS)
        except Exception:
            leaves = []
        texts = [l.get("text", "") for l in leaves]
        if any(any(k in t for k in ("导出", "打印", "另存为", "版本历史"))
               for t in texts):
            opened = True
            break
    if not opened:
        _snap("menu_not_open")
        raise RuntimeError("文件菜单点击后未展开(重试 3 次)")

    # 2) 菜单里点「导出」(新版展开格式子菜单;旧版可能直接触发下载)
    fmt_kws = (".xlsx", "Excel表格", "导出Excel", "导出为 Excel", "导出为Excel")
    try:
        # V0.1.2a: 服务器生成 xlsx 偶发偏慢(大文档),超时 60→120s
        async with page.expect_download(timeout=120000) as dl_info:
            await _click_leaf_text(page, ("导出", "导出为 Excel", "导出为Excel"),
                                   skip=("权限", "记录", "模板", "图片"))
            await page.wait_for_timeout(1800)
            # 新版: 点「导出Excel表格 (.xlsx)」格式项(旧版此时已直接下载)
            try:
                await _click_leaf_text(page, fmt_kws,
                                       skip=("图片", ".zip", ".csv", "模板"))
                await page.wait_for_timeout(1500)
                # 若弹出导出设置/确认对话框,点确认(无弹窗则找不到,跳过)
                try:
                    await _click_leaf_text(page, ("开始导出", "确 定", "确定"),
                                           min_w=30)
                except RuntimeError:
                    pass
                # V0.1.2b: 频繁导出会触发滑块验证码,阻塞下载事件;
                # 有头模式等用户拖滑块,无头模式抛异常切有头窗口重试
                await _watch_captcha(page, out_file, headless, log)
            except RuntimeError:
                pass  # 旧版菜单: 点导出项即直接下载,无格式子菜单
            await page.wait_for_timeout(2000)
        dl = await dl_info.value
    except _CaptchaNeededError:
        raise
    except Exception as e:
        _snap("download_timeout")
        raise RuntimeError(f"点击导出菜单项后未触发下载: {e}") from e

    await dl.save_as(out_file)
    p = Path(out_file)
    if not p.exists() or p.stat().st_size == 0:
        raise RuntimeError("下载文件未保存成功")
    # 校验是 xlsx(zip 容器,魔数 PK),防止下到网页/错误内容
    with open(out_file, "rb") as fh:
        if fh.read(2) != b"PK":
            raise RuntimeError("下载的文件不是 xlsx 格式(可能导出失败)")


async def _export_via_api(page, api_url: str, out_file: str):
    """策略 B: fetch opendoc,解析 sheet 数据块写 xlsx。"""
    import openpyxl

    raw = await page.evaluate(
        """async (url) => {
            const resp = await fetch(url, { credentials: 'include' });
            return await resp.text();
        }""", api_url)

    try:
        data = json.loads(raw)
    except Exception:
        m = re.search(r"\((\{.*\})\)\s*;?\s*$", raw, re.S)
        if not m:
            # 保存现场便于排查
            dbg = Path(out_file).with_suffix(".debug.txt")
            dbg.write_text(raw[:500000], encoding="utf-8")
            raise RuntimeError("接口响应不是 JSON(已保存 debug 文件)")
        data = json.loads(m.group(1))

    blocks = []

    def find_sheets(node, depth=0):
        if node is None or not isinstance(node, (dict, list)) or depth > 8:
            return
        if isinstance(node, list):
            for n in node:
                find_sheets(n, depth + 1)
            return
        keys = list(node.keys())
        if any(k in ("rowData", "cellData", "cells", "sheetName")
               for k in keys) and len(keys) < 60:
            blocks.append(node)
        for k in keys:
            find_sheets(node[k], depth + 1)

    find_sheets(data.get("clientVars") or data)

    seen, uniq = set(), []
    for b in blocks:
        key = f"{len(json.dumps(b, ensure_ascii=False))}|{b.get('sheetName', '')}"
        if key not in seen:
            seen.add(key)
            uniq.append(b)
    if not uniq:
        dbg = Path(out_file).with_suffix(".debug.json")
        dbg.write_text(json.dumps(data, ensure_ascii=False, indent=2)[:800000],
                       encoding="utf-8")
        raise RuntimeError("响应中未找到表格数据块(已保存 debug 文件)")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    cell_count = 0
    for i, block in enumerate(uniq):
        name = str(block.get("sheetName") or f"Sheet{i+1}")[:31]
        ws = wb.create_sheet(name)
        rows = _block_rows(block)
        for ri, row in enumerate(rows, start=1):
            for ci, val in enumerate(row, start=1):
                if val is None or val == "":
                    continue
                if isinstance(val, (dict, list)):
                    val = json.dumps(val, ensure_ascii=False)
                ws.cell(row=ri, column=ci, value=val)
                cell_count += 1

    if cell_count == 0:
        raise RuntimeError("数据块解析结果为空(0 个单元格)")
    # V0.1.1 隐藏署名: 写入文件属性(Excel 详情页"创建者")
    wb.properties.creator = AUTHOR
    wb.properties.lastModifiedBy = AUTHOR
    wb.save(out_file)


def _set_core_xml_author(xlsx_path: str, log=print) -> bool:
    """V0.1.2: zip 级写署名——直接改写 xlsx 内 docProps/core.xml 的
    dc:creator / cp:lastModifiedBy,不经 openpyxl 加载。

    服务器导出的 xlsx 样式表 openpyxl 3.1.5 无法解析(Fill 兼容问题),
    但 Excel 本身可正常打开,所以水印走 zip 级方案,对任何 xlsx 都成立。
    V0.1.2a: 失败原因经 log 回调输出(默认 print),便于诊断水印缺失。
    V0.1.2b: 实测刚下载的文件会被 Edge/杀软以"无删除共享"句柄锁住
    2~3 分钟,os.replace 全程 WinError 5(8×1.5s 重试不够)。
    改为三级链: ①内存构建新 zip 并校验 → ②os.replace 短重试 →
    ③原位 r+b 改写(先备份旧字节,写坏回滚),绕过被锁的目标文件。
    """
    import zipfile
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            names = zin.namelist()
            if "docProps/core.xml" not in names:
                return False
            core = zin.read("docProps/core.xml").decode("utf-8", "ignore")
            others = {n: zin.read(n) for n in names if n != "docProps/core.xml"}
        # 替换/插入 creator 与 lastModifiedBy
        for tag in ("dc:creator", "cp:lastModifiedBy"):
            if f"<{tag}>" in core:
                core = re.sub(f"<{tag}>[^<]*</{tag}>",
                              f"<{tag}>{AUTHOR}</{tag}>", core)
            else:
                core = core.replace(
                    "</cp:coreProperties>",
                    f"<{tag}>{AUTHOR}</{tag}></cp:coreProperties>")
        # ① 在内存构建新 zip(不落盘,可先行校验,避免写一半被锁)
        import io
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                if n == "docProps/core.xml":
                    zout.writestr(n, core)
                else:
                    zout.writestr(n, others[n])
        data = buf.getvalue()
        # 校验: 新 zip 必须完好且 core.xml 已含署名,否则不动原文件
        with zipfile.ZipFile(io.BytesIO(data)) as zchk:
            if zchk.testzip() is not None:
                return False
            if AUTHOR not in zchk.read("docProps/core.xml").decode(
                    "utf-8", "ignore"):
                return False
        # ② os.replace(文件未被锁的正常路径,保留原文件属性)
        tmp = xlsx_path + ".tmp"
        last_err = None
        for _i in range(4):
            try:
                Path(tmp).write_bytes(data)
                os.replace(tmp, xlsx_path)
                return True
            except OSError as e:
                last_err = e
                time.sleep(1.0)
        # ③ 原位改写: 目标被无删除共享句柄锁住时 replace 永远失败,
        # 但 r+b 写句柄通常可获得;写前备份,失败/写坏则回滚原字节
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        backup = None
        try:
            with open(xlsx_path, "rb") as f:
                backup = f.read()
            with open(xlsx_path, "r+b") as f:
                f.write(data)
                f.truncate()
            with zipfile.ZipFile(xlsx_path) as zchk:
                if zchk.testzip() is not None:
                    raise RuntimeError("原位改写后 zip 校验失败")
            return True
        except Exception as e:
            try:  # 回滚,宁可水印缺失也不交付损坏文件
                if backup is not None:
                    with open(xlsx_path, "wb") as f:
                        f.write(backup)
                log(f"[水印][原位改写] 失败已回滚: {e}; "
                    f"replace 侧: {last_err}")
            except Exception:
                pass
        return False
    except Exception as e:
        try:
            log(f"[水印][zip级] 异常: {type(e).__name__}: {str(e)[:120]}")
        except Exception:
            pass
        return False


def stamp_xlsx_author(xlsx_path: str, log=print) -> bool:
    """V0.1.1 水印: 给已存在的 xlsx 打文件属性署名(策略 A 下载文件的后处理)。

    V0.1.2: 优先 openpyxl(样式可解析时);失败回退 zip 级改写 core.xml
    (服务器导出文件的样式表 openpyxl 解析不了,但 zip 方案总能写入)。
    失败不影响导出结果(仅水印缺失),返回是否成功。
    """
    import openpyxl
    try:
        wb = openpyxl.load_workbook(xlsx_path)
        wb.properties.creator = AUTHOR
        wb.properties.lastModifiedBy = AUTHOR
        wb.save(xlsx_path)
        return True
    except Exception:
        pass  # 落到 zip 级方案
    if _set_core_xml_author(xlsx_path, log=log):
        return True
    try:
        log("[水印] 写入文件属性失败(不影响导出结果)")
    except Exception:
        pass
    return False


def _block_rows(block: dict) -> list:
    """把 sheet 数据块规整成二维数组。"""
    rows = []
    if isinstance(block.get("rowData"), list):
        rows = block["rowData"]
    elif isinstance(block.get("cells"), list):
        rows = block["cells"]
    elif isinstance(block.get("cellData"), dict):
        cd = block["cellData"]
        try:
            max_r = max(int(k) for k in cd.keys())
        except ValueError:
            return []
        for r in range(max_r + 1):
            rc = cd.get(str(r)) or cd.get(r) or {}
            if not isinstance(rc, dict):
                rows.append([])
                continue
            try:
                max_c = max(int(k) for k in rc.keys())
            except ValueError:
                rows.append([])
                continue
            arr = []
            for c in range(max_c + 1):
                cell = rc.get(str(c)) or rc.get(c)
                v = cell.get("v") if isinstance(cell, dict) else None
                arr.append(v)
            rows.append(arr)
    return rows


# ================= 替换目标文件 sheet 源 =================

def first_sheet_name(xlsx_path: str, log=print) -> str:
    """V0.1.2: 取 xlsx 第一个 sheet 名。openpyxl 失败(服务器导出文件
    样式表兼容问题)时回退 zip 级读 workbook.xml。"""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True)
        name = wb.sheetnames[0]
        wb.close()
        return name
    except Exception as e:
        try:
            log(f"[导出] 文件样式较新,openpyxl 读取 sheet 名失败"
                f"({str(e)[:50]}),改用 zip 级读取…")
        except Exception:
            pass
    names = _xlsx_sheet_names_zip(xlsx_path)
    if not names:
        raise RuntimeError("无法读取导出文件的 sheet 列表")
    return names[0]


def _xlsx_sheet_names_zip(xlsx_path: str) -> list:
    """V0.1.2: zip 级读 sheet 名(服务器导出文件样式表 openpyxl 解析失败时的兜底)。"""
    import zipfile
    with zipfile.ZipFile(xlsx_path) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8", "ignore")
    return re.findall(r'<sheet[^>]*name="([^"]+)"', xml)


def _xlsx_sheet_rows_zip(xlsx_path: str, sheet_name: str) -> list:
    """V0.1.2: zip 级读指定 sheet 的二维数据(兜底路径,只取值不做样式)。

    依次解析 sharedStrings 与目标 sheet 的 <c r= s= t=> 单元格;
    返回 [[v,...],...](1 起行号对齐,空行为 [])。
    """
    import zipfile
    import xml.etree.ElementTree as ET
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(xlsx_path) as z:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        # rid → target
        rid_map = {r.get("Id"): r.get("Target") for r in rels}
        target = None
        for sh in wb_xml.iter(f"{NS}sheet"):
            if sh.get("name") == sheet_name:
                rid = (sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
                target = rid_map.get(rid)
                break
        if not target:
            raise ValueError(f"zip 级未找到 sheet「{sheet_name}」")
        if target.startswith("/"):
            target = target[1:]
        elif not target.startswith("xl/"):
            target = "xl/" + target
        sheet_xml = ET.fromstring(z.read(target))
        # 共享字符串
        strs = []
        if "xl/sharedStrings.xml" in z.namelist():
            sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sst.iter(f"{NS}si"):
                strs.append("".join(t.text or "" for t in si.iter(f"{NS}t")))
    rows = {}
    for row in sheet_xml.iter(f"{NS}row"):
        r = int(row.get("r", "0"))
        arr = {}
        for c in row.iter(f"{NS}c"):
            ref = c.get("r", "")
            m = re.match(r"([A-Z]+)(\d+)", ref)
            if not m:
                continue
            col = 0
            for ch in m.group(1):
                col = col * 26 + (ord(ch) - 64)
            v = None
            vnode = c.find(f"{NS}v")
            is_node = c.find(f"{NS}is")
            if c.get("t") == "s" and vnode is not None:
                idx = int(vnode.text)
                v = strs[idx] if 0 <= idx < len(strs) else None
            elif c.get("t") == "inlineStr" and is_node is not None:
                v = "".join(t.text or "" for t in is_node.iter(f"{NS}t"))
            elif vnode is not None:
                v = vnode.text
                if v is not None:
                    try:
                        v = float(v) if "." in v or "e" in v.lower() else int(v)
                    except ValueError:
                        pass
            if v is not None:
                arr[col] = v
        rows[r] = arr
    if not rows:
        return []
    max_r = max(rows.keys())
    out = []
    for r in range(1, max_r + 1):
        arr = rows.get(r)
        if not arr:
            out.append([])
            continue
        max_c = max(arr.keys())
        out.append([arr.get(c) for c in range(1, max_c + 1)])
    return out


def replace_sheet(src_xlsx: str, dst_xlsx: str, sheet_name: str,
                  log=print) -> str:
    """把 src 的同名 sheet 写入/替换到 dst;若 dst 不存在则直接复制 src 为 dst。

    保持 dst 其他 sheet 不动;新 sheet 追加到末尾。
    返回 dst 中新 sheet 的名字(重名冲突时自动加序号)。

    V0.1.2: 服务器导出的源文件 openpyxl 可能解析不了(样式表兼容问题),
    源读取提供 zip 级兜底;目标文件仍用 openpyxl(本地文件无此问题)。
    """
    import openpyxl

    src, dst = Path(src_xlsx), Path(dst_xlsx)
    if not src.exists():
        raise FileNotFoundError(f"源文件不存在: {src}")
    if not dst.exists():
        # 目标不存在: 直接复制(相当于"落地"导出)
        import shutil
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        log(f"[替换] 目标不存在,已复制为: {dst}")
        return sheet_name

    # 源文件: 先 openpyxl,失败走 zip 级兜底
    zip_mode = False
    try:
        wb_src = openpyxl.load_workbook(src, data_only=True)
        if sheet_name not in wb_src.sheetnames:
            raise ValueError(f"源文件中没有 sheet「{sheet_name}」,"
                             f"现有: {wb_src.sheetnames}")
        ws_src = wb_src[sheet_name]
        src_rows = [(c.row, c.column, c.value) for row in ws_src.iter_rows()
                    for c in row if c.value is not None]
    except ValueError:
        raise
    except Exception as e:
        log(f"[替换] 源文件样式较新,openpyxl 读取失败({str(e)[:60]}),"
            "改用 zip 级数据提取…")
        zip_mode = True
        names = _xlsx_sheet_names_zip(str(src))
        if sheet_name not in names:
            raise ValueError(f"源文件中没有 sheet「{sheet_name}」,现有: {names}")
        src_rows = []
        for ri, row in enumerate(_xlsx_sheet_rows_zip(str(src), sheet_name),
                                 start=1):
            for ci, v in enumerate(row, start=1):
                if v is not None:
                    src_rows.append((ri, ci, v))

    # V0.1.2: 空 sheet 也允许写入(清空目标 sheet 是合法操作),
    # 不因源 sheet 无数据而报错(与 V0.1.0 行为保持一致)

    wb_dst = openpyxl.load_workbook(dst)

    # 重名时替换,否则追加
    final_name = sheet_name
    if sheet_name in wb_dst.sheetnames:
        log(f"[替换] 目标已存在 sheet「{sheet_name}」,将被覆盖")
        del wb_dst[sheet_name]
    else:
        n = 1
        while final_name in wb_dst.sheetnames:
            final_name = f"{sheet_name} ({n})"
            n += 1

    ws_dst = wb_dst.create_sheet(final_name)
    for ri, ci, v in src_rows:
        ws_dst.cell(row=ri, column=ci, value=v)

    wb_dst.save(dst)
    log(f"[替换] 已写入「{final_name}」到 {dst}"
        f"({'zip 级提取' if zip_mode else 'openpyxl'}, "
        f"{len(src_rows)} 个单元格)")
    return final_name
