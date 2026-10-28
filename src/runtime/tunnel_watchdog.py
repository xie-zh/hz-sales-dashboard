# -*- coding: utf-8 -*-
"""
隧道守护 v2 —— PipeKeepalive + 健康检查
========================================
• 管道维持 SSH 永不断开
• 每 30 秒健康探测 → 连续失败 2 次即重建
• URL 写入 外网访问地址.txt

运行: python tunnel_watchdog.py
或双击 启动隧道守护.bat
"""
import subprocess, re, time, os, signal
import urllib.request, json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
URL_FILE = os.path.join(BASE_DIR, "外网访问地址.txt")
TUNNEL_LOG = os.path.join(BASE_DIR, "_tunnel.log")
PORT = 8765

current_url = None
current_proc = None
shutdown = False


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def save_url(url):
    global current_url
    current_url = url
    try:
        with open(URL_FILE, "w", encoding="utf-8") as f:
            f.write(f"{url}\n（隧道守护自动维护 · 每30秒健康探测）\n")
    except Exception:
        pass


def health_check():
    if not current_url:
        return False
    try:
        req = urllib.request.Request(f"{current_url}/api/data",
                                     headers={"User-Agent": "TunnelWatchdog/2"})
        resp = urllib.request.urlopen(req, timeout=8)
        data = json.loads(resp.read().decode("utf-8"))
        return "meta" in data and "stats" in data
    except Exception:
        return False


def kill_tunnel():
    global current_proc
    if current_proc and current_proc.poll() is None:
        try:
            current_proc.terminate()
            current_proc.wait(timeout=5)
        except Exception:
            try:
                current_proc.kill()
            except Exception:
                pass
    current_proc = None


def start_tunnel():
    """通过 bash 管道保活启动 SSH 隧道，从日志中提取 URL"""
    global current_proc
    kill_tunnel()

    # 清空日志
    with open(TUNNEL_LOG, "w") as f:
        f.write("")

    log("启动隧道 (管道保活)...")

    cmd = (
        '(while true; do sleep 86400; done) | '
        'ssh -tt '
        '-o StrictHostKeyChecking=no '
        '-o UserKnownHostsFile=/dev/null '
        '-o ServerAliveInterval=15 '
        '-o ServerAliveCountMax=3 '
        '-o TCPKeepAlive=yes '
        '-o ConnectTimeout=10 '
        f'-R 80:localhost:{PORT} '
        'nokey@localhost.run '
        f'>"{TUNNEL_LOG}" 2>&1'
    )
    try:
        proc = subprocess.Popen(
            ["bash", "-c", cmd],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    except Exception as e:
        log(f"  ❌ 启动失败: {e}")
        return None

    current_proc = proc

    # 轮询日志等待 URL
    deadline = time.time() + 30
    while time.time() < deadline:
        if shutdown:
            kill_tunnel()
            return None
        if proc.poll() is not None:
            log(f"  ❌ SSH 进程退出 (code={proc.returncode})")
            return None
        try:
            with open(TUNNEL_LOG, "r", errors="replace") as f:
                content = f.read()
            m = re.search(r"([a-z0-9]+)\.lhr\.life", content)
            if m:
                return f"https://{m.group(0)}"
        except Exception:
            pass
        time.sleep(0.5)

    log("  ❌ 超时未获取 URL")
    kill_tunnel()
    return None


def main():
    global current_url, shutdown

    log("=" * 45)
    log("  隧道守护 v2 · PipeKeepalive + 健康检查")
    log(f"  {PORT}→80 | 探测间隔 30s | 失败 2 次重连")
    log("=" * 45)

    def on_signal(sig, frame):
        global shutdown
        shutdown = True
        log("⏹ 停止信号")
    signal.signal(signal.SIGINT, on_signal)
    signal.signal(signal.SIGTERM, on_signal)

    fail = 0
    while not shutdown:
        if current_proc is None or current_proc.poll() is not None:
            fail = 99

        if fail >= 2:
            log(f"🔴 重建隧道 (fail={fail})")
            for _ in range(3):
                if shutdown:
                    break
                url = start_tunnel()
                if url:
                    save_url(url)
                    time.sleep(2)
                    if health_check():
                        log(f"  ✅ 就绪: {url}")
                        fail = 0
                        break
                time.sleep(3)
            else:
                log("  ❌ 重试耗尽，90秒后重试")
                for _ in range(90):
                    if shutdown:
                        break
                    time.sleep(1)
                fail = 0
                continue

        for _ in range(30):
            if shutdown:
                break
            time.sleep(1)

        if shutdown:
            break

        if health_check():
            if fail > 0:
                log("🟢 已恢复")
            fail = 0
        else:
            fail += 1
            if fail == 1:
                log("⚠️ 探测失败 (1/2)")

    kill_tunnel()
    log("守护已退出")


if __name__ == "__main__":
    main()
