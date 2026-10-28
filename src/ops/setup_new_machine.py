# -*- coding: utf-8 -*-
# @version V1.5  @author <AUTHOR>  @created 2026-10-12
"""setup_new_pc.py —— 新电脑一键部署（配合「一键部署_新电脑.bat」使用）

把整个目录拷到新电脑后，双击项目根目录的「一键部署_新电脑.bat」，
本脚本按顺序做完下面这些事，每一步都打印进度和结果：

  1  定位项目根、读取「迁移配置.json」（没有就生成一份）
  2  清理腾讯云凭据（账号要收回；看板已改为 WorkBuddy 托管，这些不再需要）
  3  准备 Python 环境：<根>\\_env（可移植 venv）+ 装 openpyxl / pywin32
       · 优先用包内 _wheels 离线安装，新电脑没外网也能装好
  4  写 Hzexamplebot\\_paths.json（本机的位置档案）
  5  把「迁移配置.json」落地到脚本真正认的位置：
       源表绝对路径改指本机 / BI 取数工号 / 发布参数
  6  注册两个计划任务：
       HZBoardRefresh       每 5 分钟唤醒（真正跑不跑由脚本内节流决定）
       HZDashBoardDailySave 每天 23:59 拍历史快照
  7  体检（_portable.scan + _migrate.selfcheck），把结果写成一份报告
  8  首次构建（仅当 _deploy 还没有产物时），并写好发布请求标记
  9  打印「接下来要做的三件事」

参数（bat 已经替你带好了，手工跑也可以）：
    --yes           不问问题，全部按默认走
    --no-task       不注册计划任务
    --no-venv       跳过建 venv，沿用当前解释器
    --keep-tencent  保留腾讯云凭据（默认清理）
    --force-build   不管有没有产物，都重建一次
    --no-build      完全跳过构建
    --report <路径> 报告写到指定文件

退出码：0=全部就绪　1=有未完成项（看报告）
"""

import os
import sys
import time
import traceback
import datetime

BASE = os.path.dirname(os.path.abspath(__file__))          # …\\Hzexamplebot
ROOT = os.path.dirname(BASE)
sys.path.insert(0, BASE)

# 让中文在任何解释器下都能正确打印（子进程 stdout 接管道时默认是 GBK）
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import _portable                                            # noqa: E402
import _migrate                                             # noqa: E402

LOG = []


def say(msg=""):
    line = str(msg)
    LOG.append(line)
    try:
        print(line)
    except Exception:
        print(line.encode("utf-8", "replace").decode("utf-8", "replace"))


def step(n, title):
    say("")
    say("=" * 58)
    say(" %d. %s" % (n, title))
    say("=" * 58)


def main(argv):
    yes = "--yes" in argv
    do_task = "--no-task" not in argv
    do_venv = "--no-venv" not in argv
    keep_tencent = "--keep-tencent" in argv
    do_build = "--no-build" not in argv
    force_build = "--force-build" in argv

    t0 = time.time()
    say("某公司看板 · 新电脑部署")
    say("时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    say("脚本目录：%s" % BASE)

    # ---------- 1 项目根与配置 ----------
    step(1, "检查项目目录与配置")
    root = _portable.resolve_root()
    say("项目根：%s" % root)
    if not _portable._is_root(root):
        say("× 这不是有效的项目目录：需要同时存在 build_snapshot.py 与配置文件。")
        say("  请确认整个文件夹是**完整解压**的（不要只复制 Hzexamplebot 子目录）。")
        return 1
    cfg_path, created = _migrate.ensure(root, log=say)
    say("配置文件：%s%s" % (cfg_path, "（新建）" if created else ""))
    cfg = _migrate.load(root)
    say("发布通道：%s" % (cfg.get("publish", {}).get("channel") or "未设置"))

    # ---------- 2 清理腾讯云凭据 ----------
    step(2, "清理腾讯云凭据（账号收回）")
    if keep_tencent:
        say("- 按参数要求保留腾讯云凭据，跳过。")
    else:
        ok, msg = _migrate.purge_tencent(root, log=say)
        say("- %s" % msg)

    # ---------- 3 Python 环境 ----------
    step(3, "准备 Python 环境（可移植 venv + 依赖）")
    py = None
    if do_venv:
        wheels = os.path.join(root, "_wheels")
        if os.path.isdir(wheels) and os.listdir(wheels):
            say("- 检测到离线依赖包 _wheels（%d 个文件），将优先离线安装"
                % len(os.listdir(wheels)))
        else:
            say("- 未发现离线依赖包，将从镜像站安装（需要外网）")
        ok, res = _portable.ensure_venv(root, log=say)
        if ok:
            py = res
            say("- Python 环境就绪：%s" % py)
        else:
            say("- × 环境准备未完成：%s" % res)
            py = _portable.resolve_python()
            say("- 退回使用现有解释器：%s" % py)
    else:
        py = _portable.resolve_python()
        say("- 按参数跳过建 venv，使用：%s" % py)

    if not py or not os.path.isfile(py):
        say("- × 本机没有可用的 Python。请先装一个 Python 3.10+")
        say("  （python.org 下载，安装时务必勾选 Add python.exe to PATH），")
        say("  或者装好 WorkBuddy 后重跑本脚本（它会带来 Python 运行时）。")

    # ---------- 4 位置档案 ----------
    step(4, "写入位置档案 _paths.json")
    p = _portable.save_paths({"root": root, "python": py or "",
                              "node": _portable.resolve_node(),
                              "tcb_cli": _portable.resolve_tcb()})
    say("- %s" % p)
    say("- 说明：这是本机专用的路径记录，换机器删掉它会自动重新认路。")

    # ---------- 5 配置落地 ----------
    step(5, "把配置落地到脚本认的位置")
    # 新电脑：源表路径一律强制指到本机项目根（旧机上"存在即不动"的规则在这里不适用，
    # 因为老路径本来就该在这台机器上失效）
    ok5, msgs = _migrate.apply_all(root, log=say, force_paths=True,
                                   purge_tencent=False)
    for m in msgs:
        say("- %s" % m)

    # ---------- 5b 导航页链接校正 ----------
    if (_migrate.load(root).get("publish") or {}).get("dash_url"):
        step("5b", "按新地址校正导航页链接")
        try:
            import _repoint_nav
            rc = _repoint_nav.main([])
            say("- 导航页改写结束，退出码 %s" % rc)
        except Exception as e:
            say("- ! 导航页改写未完成：%s" % str(e)[:150])
            say("  （不影响主看板；回填 dash_url 后手工跑一次 Hzexamplebot\\_repoint_nav.py 即可）")
    else:
        step("5b", "导航页链接校正")
        say("- 主看板链接还没回填，跳过。发布 _deploy 拿到链接、填进配置后再跑：")
        say("    python Hzexamplebot\\_repoint_nav.py")

    # ---------- 6 计划任务 ----------
    step(6, "注册计划任务")
    if not do_task:
        say("- 按参数跳过计划任务注册。")
    else:
        tk = cfg.get("task") or {}
        ok1, m1 = _portable.register_task(
            root, py, minute=int(tk.get("refresh_minutes") or 5),
            name=tk.get("refresh_name") or "HZBoardRefresh")
        say("- %s %s" % ("√" if ok1 else "×", m1))
        ok2, m2 = _portable.register_daily_task(
            root, py, hhmm=tk.get("daily_save_time") or "23:59",
            name=tk.get("daily_save_name") or "HZDashBoardDailySave")
        say("- %s %s" % ("√" if ok2 else "×", m2))
        if not keep_tencent:
            ok3, m3 = _portable.remove_task("HZDashboardCosSync")
            say("- 腾讯云备份任务：%s" % (m3 if ok3 else "（本来就没有，忽略）"))

    # ---------- 7 体检 ----------
    step(7, "环境体检")
    rows = []
    try:
        rows = _portable.scan(root, py)
    except Exception as e:
        say("- × 体检失败：%s" % str(e)[:120])
    mark = {"ok": "√", "warn": "!", "bad": "×"}
    bad = warn = 0
    for r in rows:
        say("  %s %-20s %s" % (mark[r["level"]], r["label"], r["detail"]))
        bad += 1 if r["level"] == "bad" else 0
        warn += 1 if r["level"] == "warn" else 0
    say("- 合计：%d 项需要处理，%d 项提醒" % (bad, warn))

    # ---------- 8 首次构建 ----------
    step(8, "首次构建本地看板")
    site = _portable.site_dir(root)
    need = force_build or not os.path.isfile(os.path.join(site, "index.html"))
    if not do_build:
        say("- 按参数跳过构建。")
    elif not need:
        say("- 已有构建产物（%s），跳过。要强制重建请加 --force-build" % site)
    else:
        say("- 开始构建（取数已完成时约 1 分钟，首次可能要 2~3 分钟）…")
        try:
            import bi_refresh
            rc = bi_refresh.main_entry()
            say("- 构建+发布阶段结束，退出码 %s" % rc)
        except Exception:
            say("- × 构建异常：")
            say(traceback.format_exc(limit=4))

    # ---------- 9 接下来 ----------
    step(9, "接下来要做的三件事")
    pub = _migrate.publish(root)
    dash = pub.get("dash_url") or ""
    nav = pub.get("nav_url") or ""
    say("① 发布到 WorkBuddy（脚本发不了，必须由客户端发起）")
    say("   在本机打开 WorkBuddy，对它说一句：")
    say("     「把 %s 发布为应用」" % site)
    say("   拿到分享链接后回填到 %s 的 publish.dash_url" % os.path.basename(cfg_path))
    if nav:
        say("   导航页同法：目录 %s" % _portable.nav_dir(root))
    say("")
    say("② 回填链接（回填后线上核对、控制台「打开看板」都会跟上）")
    say("   现在：dash_url=%s  nav_url=%s" % (dash or "（空）", nav or "（空）"))
    say("")
    say("③ 建一条 WorkBuddy 自动化，让看板自己更新")
    say("   提示词直接抄同目录的「新电脑_自动化任务.txt」")
    say("   （本机需保持开机 + WorkBuddy 登录，发布才发得出去）")

    # ---------- 报告 ----------
    say("")
    say("=" * 58)
    if bad == 0:
        say(" 部署完成 —— 全部 %d 项就绪（%d 项提醒）" % (len(rows), warn))
    else:
        say(" 部署基本完成 —— 有 %d 项需要处理（见上）" % bad)
    say(" 用时 %.1f 秒" % (time.time() - t0))
    say("=" * 58)

    rp = ""
    for i, a in enumerate(argv):
        if a == "--report" and i + 1 < len(argv):
            rp = argv[i + 1]
    if not rp:
        rp = os.path.join(root, "_portable", "新电脑部署报告.txt")
    try:
        os.makedirs(os.path.dirname(rp))
    except Exception:
        pass
    try:
        with open(rp, "w", encoding="utf-8") as f:
            f.write("\n".join(LOG) + "\n")
        print("报告已写入：%s" % rp)
    except Exception as e:
        print("报告写入失败：%s" % str(e)[:80])

    if not yes:
        try:
            input("\n按回车关闭…")
        except Exception:
            pass
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except Exception:
        print(traceback.format_exc(limit=6))
        try:
            input("\n出错了，按回车关闭…")
        except Exception:
            pass
        sys.exit(1)
