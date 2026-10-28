# -*- coding: utf-8 -*-
"""
手动部署脚本：对比本地/云端 → 有差异才触发云端部署
====================================================
由「手动部署.bat」调用。逻辑：
  1. 读取本地 _deploy/data.json 的 file_mtime_ts
  2. 请求云端 data.json 的 file_mtime_ts
  3. 相同        → 提示「已是最新」，结束（不消耗任何积分）
  4. 本地较新    → 重新生成快照 → 确认后触发 WorkBuddy 云端部署（消耗 1 次）
  5. 云端不可达  → 提示，由用户决定是否强制触发
"""
import os, sys, json, time, sqlite3, datetime
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(BASE_DIR, "_deploy", "data.json")
DB = os.path.expanduser("~/.workbuddy/workbuddy.db")
AUTOMATION_ID = "automation-1786499389829"  # 与实际自动化 ID 一致（原 AUTOMATION_NAME 名称不符导致 UPDATE rowcount=0，部署永不被触发）
CLOUD_DATA_URL = "https://dashboard.example.com/data.json"
LOG_FILE = os.path.join(BASE_DIR, "_manual_deploy.log")
TIMEOUT = 15


def log(msg):
    line = f"[{datetime.datetime.now().strftime('%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def read_local_meta():
    """返回 (file_mtime_ts, data_time)"""
    try:
        with open(DATA_JSON, "r", encoding="utf-8") as f:
            meta = json.load(f).get("meta", {})
        return meta.get("file_mtime_ts", 0), meta.get("data_time", "?")
    except Exception as e:
        log(f"⚠️ 读取本地快照失败: {e}")
        return None, None


def read_cloud_meta():
    """返回 (file_mtime_ts, data_time)；失败返回 None"""
    try:
        req = urllib.request.Request(
            CLOUD_DATA_URL, headers={"User-Agent": "manual-deploy/1"})
        resp = urllib.request.urlopen(req, timeout=TIMEOUT)
        meta = json.loads(resp.read().decode("utf-8")).get("meta", {})
        return meta.get("file_mtime_ts", 0), meta.get("data_time", "?")
    except Exception as e:
        log(f"⚠️ 云端不可达: {e}")
        return None, None


def rebuild_snapshot():
    """重新生成 _deploy 快照"""
    try:
        sys.path.insert(0, BASE_DIR)
        import build_snapshot
        build_snapshot.build()
        return True
    except Exception as e:
        log(f"❌ 生成快照失败: {e}")
        return False


def trigger_deploy():
    """把云端同步自动化设为立即执行（WorkBuddy 调度器随即跑一次 AI 部署）"""
    try:
        conn = sqlite3.connect(DB, timeout=10)
        cur = conn.cursor()
        now_ms = int(time.time() * 1000)
        # 按自动化 ID 精确匹配（原按 name 匹配因名称不符导致 rowcount=0）
        cur.execute(
            "UPDATE automations SET next_run_at=?, updated_at=?, status='ACTIVE' "
            "WHERE id=? AND deleted_at IS NULL",
            (now_ms, now_ms, AUTOMATION_ID),
        )
        conn.commit()
        n = cur.rowcount
        conn.close()
        log(f"✅ 已触发云端部署任务 (rowcount={n})")
        return n > 0
    except Exception as e:
        log(f"❌ 触发部署失败: {e}")
        return False


def main():
    log("=" * 50)
    log("  看板手动部署工具")
    log("=" * 50)

    local_ts, local_time = read_local_meta()
    if local_ts is None:
        log("请先运行一次自动同步生成快照，或检查 _deploy/data.json")
        return

    cloud_ts, cloud_time = read_cloud_meta()

    if cloud_ts is None:
        log(f"\n本地快照时间: {local_time} (ts={local_ts:.0f})")
        log("云端暂时不可达，无法对比。")
        ans = input("\n是否仍强制触发一次部署？(y/N): ").strip().lower()
        if ans not in ("y", "yes"):
            log("已取消。")
            return
        if not rebuild_snapshot():
            return
        trigger_deploy()
        log("已触发。约 1-2 分钟后云端更新，链接不变。")
        return

    diff = abs(local_ts - cloud_ts)
    if local_ts <= cloud_ts and diff < 1:
        log(f"\n✅ 已是最新，无需部署。")
        log(f"   本地: {local_time}")
        log(f"   云端: {cloud_time}")
        log("（本次未消耗任何积分）")
        return

    # 本地较新（或差异 > 1s）
    log(f"\n📄 检测到本地有新数据:")
    log(f"   本地: {local_time} (ts={local_ts:.0f})")
    log(f"   云端: {cloud_time} (ts={cloud_ts:.0f})")
    ans = input("\n是否生成快照并触发云端部署？(y/N): ").strip().lower()
    if ans not in ("y", "yes"):
        log("已取消。")
        return

    if not rebuild_snapshot():
        return
    if trigger_deploy():
        log("\n✅ 已触发云端部署，约 1-2 分钟后生效。")
        log("   访问链接: https://dashboard.example.com")
        log("   （本次消耗 1 次 AI 部署积分）")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已取消。")
