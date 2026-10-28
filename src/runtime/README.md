# 运行目录脚本（仅存档，不执行）

这里是从实际运行目录原样收进来的运维脚本，用于说明系统在真实环境里是怎么跑起来的。
**它们不会在演示环境中执行**：路径、内网地址、凭证都已移除，业务数据也不在仓库内。

## 目录定位

| 分组 | 文件 | 作用 |
|---|---|---|
| 调度入口 | `0_一键刷新看板.bat` | 一键跑完整条链路 |
| 主流水线 | `pipeline_sync.py` / `check_and_deploy.py` / `watch_and_deploy.py` | 取数 → 构建 → 判变 → 发布 |
| 取数 | `bi_sync.py` / `cos_sync.py` | BI 报表下载、对象存储同步 |
| 解析 | `dashboard_server.py` / `ry_parser.py` / `sanli_parser.py` | 活动量 / 网销 / 三力三套解析器 |
| 构建 | `build_snapshot.py` / `build_pure.py` / `hist_snapshot.py` / `history_builder.py` | 生成页面与历史快照 |
| 单文件化 | `_make_singlefile.py` | 把多文件产物压成自包含 HTML |
| 导航页 | `_make_nav_layouts.py` / `_make_nav_themes.py` / `_gen_banners.py` | 导航页版式与轮播素材 |
| 质检 | `asr_tencent.py` / `asr_local_whisper.py` / `score_calls.py` / `gen_report.py` | 录音转写与五维评分 |
| 简报 | `daily_brief.py` / `notice_sync.py` / `gen_workspace_context.py` | 每日简报与公告同步 |
| 部署 | `manual_deploy.py` / `quick_update.py` / `update_dashboard.py` | 手动发布与增量更新 |
| 体检 | `_pre_build_check.py` / `verify_artifact.py` / `config_loader.py` | 产物校验与环境自检 |
| 隧道 | `tunnel_watchdog.py` | 本地服务外网隧道的存活看护 |
| 前端 | `miniqa.js` / `spark_proxy.js` | 页面内 AI 问答与流式代理 |

## 为什么放在这里

面试或交接时经常会被问「调度是怎么串起来的」「判变逻辑在哪」。
把脚本原样保留，配合 `docs/架构说明.md` 就能对照阅读，不必回原始环境翻文件。
