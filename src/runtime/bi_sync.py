# -*- coding: utf-8 -*-
"""
分公司活动量 · BI自动同步脚本
==================================
功能：
  1. 打开 BI 系统 (bi.example.com)
  2. 登录态复用 Chrome 已登录的 Cookie
  3. 点击「明细导出」→ 筛选今日 + 分公司 → Excel原样导出
  4. 下载的文件替换到「分公司活动量-在线看板.xlsx」的【当日活动量LTC】B2开始
  5. 每 30 分钟执行一次
  6. 在工作目录放置 stop_sync.txt 即可安全停止

用法：
  python bi_sync.py

首次运行前请确保：
  - 已安装 playwright：pip install playwright && playwright install chromium
  - Chrome 中已登录 bi.example.com
"""

import os
import sys
import time
import glob
import shutil
import datetime
import threading
import subprocess

# ======================== 配置 ========================
BI_URL = "https://bi.example.com/#/home/index"
DOWNLOAD_DIR = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\_downloads"
TARGET_XLSX = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\分公司活动量-在线看板.xlsx"
SHEET_NAME = "当日活动量LTC"
STOP_FILE = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\stop_sync.txt"
INTERVAL_MINUTES = 30
CHROME_PROFILE = r"<HOME>\AppData\Local\Google\Chrome\User Data"

# Playwright
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ======================== 工具函数 ========================
def log(msg):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def is_stopped():
    return os.path.exists(STOP_FILE)

def find_latest_download():
    """返回下载目录中最新的文件"""
    files = glob.glob(os.path.join(DOWNLOAD_DIR, "*"))
    if not files:
        return None
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]

def clear_download_dir():
    """清空下载目录"""
    if os.path.exists(DOWNLOAD_DIR):
        for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*")):
            try:
                os.remove(f)
            except Exception:
                pass
    else:
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def close_chrome():
    """关闭所有 Chrome 进程（以便用 Playwright 接管用户配置）"""
    try:
        subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"],
                       capture_output=True, timeout=10)
        time.sleep(2)
    except Exception:
        pass

# ======================== BI 下载 ========================
def download_bi_report():
    """用 Playwright 打开 BI 系统 + 复用 Chrome Cookie → 下载今日分公司报表"""
    close_chrome()  # 释放用户配置锁定
    clear_download_dir()

    with sync_playwright() as p:
        log("启动 Chrome（使用已有配置）...")
        context = p.chromium.launch_persistent_context(
            user_data_dir=CHROME_PROFILE,
            channel="chrome",     # 使用系统安装的 Chrome
            headless=False,       # 非无头，便于调试
            accept_downloads=True,
            no_viewport=True,
            args=["--start-maximized"],
        )
        page = context.new_page()

        try:
            # 1. 打开 BI 首页
            log("打开 BI 系统...")
            page.goto(BI_URL, wait_until="networkidle", timeout=30000)
            time.sleep(3)  # 等待 SPA 渲染

            # 检查是否需要手动登录（未登录会有登录页）
            if "login" in page.url.lower() or "登录" in page.content():
                log("⚠️ 需要手动登录！请在弹出的 Chrome 窗口中完成登录后按回车...")
                input(">>> 登录完成后按回车继续...")
                page.goto(BI_URL, wait_until="networkidle", timeout=30000)
                time.sleep(3)

            # 2. 等待页面加载，找到「明细导出」按钮
            log("查找「明细导出」按钮...")
            page.wait_for_load_state("networkidle", timeout=15000)
            time.sleep(2)

            # 尝试多种方式定位导出按钮
            export_btn = None
            selectors = [
                "text=明细导出",
                "button:has-text('明细导出')",
                "a:has-text('明细导出')",
                "span:has-text('明细导出')",
                "div:has-text('明细导出')",
                "[title*='明细导出']",
                "[aria-label*='明细导出']",
            ]
            for sel in selectors:
                try:
                    export_btn = page.locator(sel).first
                    if export_btn.is_visible(timeout=2000):
                        log(f"  找到按钮: {sel}")
                        break
                except Exception:
                    continue

            if not export_btn:
                log("⚠️ 自动定位失败，请手动点击「明细导出」按钮（60秒超时）")
                try:
                    page.wait_for_selector("text=明细导出", timeout=60000)
                    export_btn = page.locator("text=明细导出").first
                except PWTimeout:
                    log("❌ 超时，请确认页面是否有「明细导出」按钮")
                    return False

            # 3. 点击导出按钮
            export_btn.click()
            time.sleep(2)

            # 4. 设置筛选条件：日期=今天
            log("设置筛选条件...")
            # 日期选择器 - 通常导出的弹窗里有日期选择
            try:
                today_btn = page.locator("text=今天").first
                if today_btn.is_visible(timeout=3000):
                    today_btn.click()
                    log("  日期 → 今天")
            except Exception:
                log("  日期筛选跳过（可能默认已是今天）")

            # 分公司选择
            try:
                branch_selector = page.locator("text=分公司").first
                if not branch_selector.is_visible(timeout=2000):
                    # 尝试点击分公司下拉框
                    for branch_trigger in [
                        "text=分公司",
                        "text=选择分公司",
                        "[placeholder*='分公司']",
                        ".branch-select",
                    ]:
                        try:
                            trigger = page.locator(branch_trigger).first
                            if trigger.is_visible(timeout=1000):
                                trigger.click()
                                time.sleep(1)
                                break
                        except Exception:
                            continue
                branch_btn = page.locator("text=分公司").first
                if branch_btn.is_visible(timeout=3000):
                    branch_btn.click()
                    log("  分公司 → 分公司")
            except Exception:
                log("  分公司筛选跳过")

            # 导出格式：Excel-原样导出
            try:
                excel_btn = page.locator("text=原样导出").first
                if not excel_btn.is_visible(timeout=2000):
                    excel_btn = page.locator("text=Excel").first
                if excel_btn.is_visible(timeout=3000):
                    excel_btn.click()
                    log("  格式 → Excel原样导出")
            except Exception:
                log("  Excel格式选择跳过")

            # 5. 点击确认导出
            log("执行导出...")
            confirm_selectors = [
                "button:has-text('导出')",
                "button:has-text('确认')",
                "button:has-text('确定')",
                "text=导出",
                "text=确认导出",
            ]
            confirmed = False
            for sel in confirm_selectors:
                try:
                    btn = page.locator(sel).first
                    if btn.is_visible(timeout=2000):
                        # 监听下载
                        with page.expect_download(timeout=60000) as download_info:
                            btn.click()
                        download = download_info.value
                        save_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename or "ltc_report.xlsx")
                        download.save_as(save_path)
                        log(f"✅ 下载完成: {save_path}")
                        confirmed = True
                        break
                except Exception:
                    continue

            if not confirmed:
                log("⚠️ 自动点击导出失败，尝试等待手动下载（60秒）...")
                time.sleep(60)
                latest = find_latest_download()
                if latest:
                    log(f"✅ 检测到下载: {latest}")
                else:
                    log("❌ 未检测到下载文件")
                    return False

            time.sleep(2)
            return True

        except Exception as e:
            log(f"❌ BI 下载异常: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            log("关闭 Chrome...")
            context.close()
            time.sleep(1)


# ======================== Excel 替换 ========================
def merge_to_xlsx():
    """将下载的报表替换到目标 Excel 的当日活动量LTC 子表（从 B2 开始粘贴，含表头）"""
    import openpyxl
    from openpyxl.utils import get_column_letter

    latest = find_latest_download()
    if not latest:
        log("❌ 没有找到下载文件，跳过 Excel 替换")
        return False

    log(f"读取下载文件: {os.path.basename(latest)}")
    try:
        src_wb = openpyxl.load_workbook(latest, read_only=True, data_only=True)
        src_ws = src_wb.active
        src_data = list(src_ws.iter_rows(values_only=True))
        src_rows = len(src_data)
        src_cols = len(src_data[0]) if src_data else 0
        log(f"  下载文件: {src_rows} 行 × {src_cols} 列")
        src_wb.close()
    except Exception as e:
        log(f"❌ 读取下载文件失败: {e}")
        return False

    if src_rows == 0:
        log("❌ 下载文件为空")
        return False

    log(f"打开目标文件: {os.path.basename(TARGET_XLSX)}")
    try:
        dst_wb = openpyxl.load_workbook(TARGET_XLSX)
        dst_ws = dst_wb[SHEET_NAME]
    except Exception as e:
        log(f"❌ 打开目标文件失败: {e}")
        return False

    # 清空旧数据 + 粘贴新数据：从 B2（Excel row 2, col B）开始
    # 先计算需要的列范围（取下载文件和目标文件中较大的列数）
    max_row = dst_ws.max_row
    need_rows = max(max_row, 2 + src_rows)
    need_cols = min(dst_ws.max_column, 2 + src_cols)

    log(f"清空旧数据 (B2 → {need_rows}行 × {need_cols-1}列)...")
    # A列保留不动，清空 B..need_cols，行 2..need_rows
    for row in range(2, need_rows + 1):
        for col in range(2, need_cols + 1):
            dst_ws.cell(row=row, column=col).value = None

    # 粘贴新数据：从 B2 开始
    log(f"粘贴新数据 ({src_rows} 行 × {src_cols} 列) → B2...")
    for src_row_idx, src_row in enumerate(src_data):
        dst_excel_row = 2 + src_row_idx
        for src_col_idx, value in enumerate(src_row):
            dst_excel_col = 2 + src_col_idx
            if dst_excel_col > dst_ws.max_column:
                break  # 超出目标列数
            # 跳过公式错误值
            if isinstance(value, str) and (value.startswith("#") or value == "None"):
                continue
            dst_ws.cell(row=dst_excel_row, column=dst_excel_col).value = value

    # 保存
    log("保存目标文件...")
    try:
        dst_wb.save(TARGET_XLSX)
        log("✅ Excel 替换完成")
        return True
    except PermissionError:
        log("⚠️ 文件被占用（可能在 Excel 中打开），请关闭 Excel 后重试")
        return False
    except Exception as e:
        log(f"❌ 保存失败: {e}")
        return False
    finally:
        dst_wb.close()


# ======================== 主循环 ========================
def run_once():
    """执行一次完整的 下载 → 替换 流程"""
    log("=" * 50)
    log("开始同步任务")

    if is_stopped():
        log("检测到 stop_sync.txt，任务终止")
        return "stopped"

    success = download_bi_report()
    if success:
        merge_to_xlsx()
    else:
        log("下载失败，跳过 Excel 替换")

    log(f"本轮完成，{INTERVAL_MINUTES} 分钟后执行下一轮")
    return "ok"


def main_loop():
    log("=" * 50)
    log("  分公司 BI → Excel 自动同步")
    log(f"  间隔: {INTERVAL_MINUTES} 分钟")
    log(f"  停止: 在脚本目录创建 stop_sync.txt")
    log("=" * 50)

    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    while True:
        result = run_once()
        if result == "stopped":
            log("⏹ 同步已停止（stop_sync.txt 存在）")
            break

        # 每 10 秒检查一次停止标志，共 INTERVAL_MINUTES 分钟
        for _ in range(INTERVAL_MINUTES * 6):
            if is_stopped():
                log("⏹ 检测到 stop_sync.txt，停止循环")
                return
            time.sleep(10)

    log("程序退出")


if __name__ == "__main__":
    main_loop()
