# 项目记忆：HZ 战报项目

> 供 WorkBuddy 在项目会话中自动加载。记架构、路径、约定、常用命令。

## 架构

三条独立流水线，共用一套「渲染图片 → 装配负载 → 企业微信发送 → 汇总结果」的模式。

```
数据源 ──► 渲染器 (Pillow) ──► 图片
                              ├─► 上传取 media_id
文案（含口径计算）────────────┴─► markdown 负载 ──► aibot send ──► show_*result*.py 汇总
```

| 链路 | 目录（运行目录相对路径） | 关键脚本 |
|---|---|---|
| 活动量 | `活动量/_report/` | `render_undone.py`、`build_send_payloads.py`、`run_report.sh`、`show_result.py` |
| 三力 | `三力追踪/_report/` | `build_hero_report.py`、`assemble_hero.py`、`run_hero.sh`、`show_hero_result.py` |
| PK | `PK追踪/` | `pk_report.py`（串起 `业绩BI同步/sync_yj.py` + `scratch/capture_pk2.py`） |
| 公共 | `_common/` | `atall.py`（群机器人 webhook @全体）、`webhooks.json` |

## 路径

- 运行目录：`<PROJECT_ROOT>/数据文件/`
- 本项目（协作副本）：`<PROJECT_ROOT>/数据文件/HZ战报项目/`
- 路径一律用 **Windows 风格** `C:/...`；Git Bash 的 `/c/...` 传给 Windows 程序会被解析成 `C:\c\...`
- Python：`<HOME>/.workbuddy/binaries/python/envs/default/Scripts/python.exe`（含 openpyxl/Pillow/playwright）
- Node：`<HOME>/.workbuddy/binaries/node/versions/*/node.exe`（**glob 探测，勿硬编码版本**）
- 源文件缓存：`<DATA_CACHE>/1688857640763810/Cache/File/<年-月>/`

## 常用命令

```bash
# 活动量
bash "活动量/_report/run_report.sh" track          # 追踪（默认）
bash "活动量/_report/run_report.sh" result         # 结果
REPORT_GROUP="3C认证" bash "活动量/_report/run_report.sh" track   # 试发测试群

# 三力（源 xlsx 用环境变量指定）
HERO_SRC="<DATA_CACHE>/.../10月三力&有效人力追踪数据YYYYMMDD 24点.xlsx" bash "三力追踪/_report/run_hero.sh"

# PK
python "PK追踪/pk_report.py" --group "管理者大群"
```

## 项目约定（团队）

1. `源码/` 是协作副本，运行目录是执行环境；改完用 `python tools/sync_to_runtime.py` 同步（先 `--dry`）。
2. 企业微信 aibot **不支持任何 @ 语法**：文案只能写纯文本 `@姓名`。真@需群机器人 webhook（`_common/webhooks.json`，当前为空）。
3. 发送失败**不重试**；发送成功与否以 `show_*result*.py` 的 `ALL_SUCCESS` 为准。
4. 群发前先试发测试群。用户说「不发」就只改不发并明确回报。
5. 文案链接规则（活动量，已定稿）：不带结尾斜杠、**链接后留一个空格**、不带结尾右括号。
6. 不在对外回复中暴露 `chat_id` / `userid` / `media_id`。

## 定时任务（自动化是**个人级**的，不随项目共享）

| 时间 | 链路 |
|---|---|
| 9:03 / 11:03 / 14:03 / 16:03 | PK 战报 |
| 11:03 / 14:33 / 16:03 | 活动量追踪 |
| 18:03 | 活动量结果 |

→ 别的成员加入项目后**不会**看到这些自动化，需要各自按需创建。
