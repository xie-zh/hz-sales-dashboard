---
name: hz-branch-wecom-reports
description: 分公司企业微信「战报」流水线运维技能，覆盖三力/有效人力战报、活动量战报、PK 战报三条链路：触发方式、脚本路径、数据口径、发送顺序，以及 wecom-cli 授权过期（errcode 850003）、node 运行时路径失效、同步锁等故障排查。当用户丢来「三力&有效人力追踪数据」xlsx、要求发/更新战报、发活动量追踪或结果、发 PK 战报，或排查企微定时任务与 wecom-cli 报错时使用。
agent_created: true
---

# 分公司企微战报运维

负责分公司三条企业微信战报流水线的日常运维：**三力战报**（三力/有效人力追踪）、**活动量战报**（追踪/结果）、**PK 战报**。核心约定：单文件脚本一键跑通、发送到指定群、**无变化则跳过**、失败不重试。

## 何时使用

- 用户直接丢来 `10月三力&有效人力追踪数据YYYYMMDD 24点.xlsx`（企微缓存路径）→ 走「三力战报」
- 用户说「发活动量」「活动量追踪/结果」「更新看板文案」→ 走「活动量战报」
- 用户说「发 PK 战报」「同步业绩」→ 走「PK 战报」
- 用户报告定时任务没发、或 wecom-cli 报 `850003`/`FileNotFoundError` → 走「故障排查」
- 用户要求「检测/体检定时任务」→ 按「定时任务」一节做**不发送**的 dry-run

## 环境事实（重要，先记住）

| 项 | 值 |
|---|---|
| Python（含 openpyxl/Pillow/playwright） | `<HOME>/.workbuddy/binaries/python/envs/default/Scripts/python.exe` |
| Node（受管运行时） | `<HOME>/.workbuddy/binaries/node/versions/<版本>/node.exe` — **版本号会随升级变化，禁止硬编码**，用 glob 探测 |
| wecom-cli | 直接 `wecom-cli ...`（bash 里在 PATH；**Python subprocess 里必须走 node 直调 wecom.js**，见故障排查） |
| 企微账号缓存目录 | `<DATA_CACHE>/1688857640763810/Cache/File/<年-月>/` |
| 路径风格 | **必须 Windows 风格** `C:/...`；`/c/...` 会被 Windows 程序解析成 `C:\c\...` |

## 通用发送规则（三条链路共同约定）

- **aibot 通道无法真正 @ 任何人**（2026-10-09 实测定论，别再尝试）：markdown 只有 `content`、
  没有 `mentioned_list`；正文写 `<@userid>` 不会被渲染，而是**原样显示成乱码**
  `<@woPClcBwAA…>`（真实事故）；`wecom-cli contact users search` 返回的 `userid` 是**框架加密句柄**
  （连授权真人「王梦洁」也是同一形态），写进正文不会被解密；`message send` 仅纯文本、同样无提及字段。
  **真@全体** 只能靠「群机器人 Webhook」（text + `mentioned_list`），webhook 未配置则静默跳过（`|| true`），不阻塞主流程。
- **@单人**：`mention_util.py` 直接产出**纯文本 `@姓名`**——群里可读、无乱码，但**不会触发对方提醒**。
  真@ 需 webhook + **真实 userid 或手机号**（加密句柄无效）。`mention_util.py` 的模块注释已写明原委，
  **不要再改回 `<@…>`**（会产生乱码事故）。
- **chat_id 每次重新取**（`wecom-cli message aibot sessions list`），不要硬编码复用
- 发送后必须看结果汇总（`ALL_SUCCESS`）；**有失败只报告 errcode/errmsg，不重试**
- 幂等靠状态文件的 `last_ymd` 严格大于才发；同日多版本不重发，确需重发用 `--force`

参考群（仅参考，实跑以 sessions list 为准）：数据点检群、管理者大群、管理者群（勇闯奖金包…）、分公司经理内部群、🤡集团3C认证团队🤡

---

## 一、三力战报（三力 & 有效人力追踪）

**触发**：用户丢来 `<DATA_CACHE>/.../10月三力&有效人力追踪数据YYYYMMDD 24点.xlsx`，即按惯例发「数据点检群」。**每次由人工触发，无定时任务。**

**一键命令**：

```bash
HERO_SRC="<源文件.xlsx>" bash "<PROJECT_ROOT>/数据文件/三力追踪/_report/run_hero.sh"
```

**流程**：`build_hero_report.py`（取数+渲染 5 图+5 段文案）→ sessions list → 上传 6 图 → `assemble_hero.py` 装配 11 条 → webhook @全体 → 发送 11 条 → `show_hero_result.py` 汇总。

**发送顺序（固定，不可乱）**：① 事业部汇总文案 ② 事业部汇总图 ③ 恭喜英雄人力 ④ 英雄完成情况图 ⑤ 英雄差距文案 ⑥ 英雄差距榜图 ⑦ 破零催办文案 ⑧ 破零名单图 ⑨ 有效人力文案 ⑩ 有效人力差距图 ⑪ **英雄人力特权海报（永远最后）**。

### 数据口径（易错，必读）

- 源 sheet = **企服顾问明细**（第 3 行表头，第 4 行起数据），按 `部门(I列) 含""` 取人
- 周期由文件名日期戳推导：**日 ≥ 15 → 属"下月"周期；日 ≤ 14 → 属"当月"周期**（如 `20261008` → 10月周期 9/15–10/14）
- **列位会随周期整体位移**（10月版比9月版右移一列）→ 换周期文件**必须先核对表头行**，禁止沿用硬编码列号
- 关键列（10月周期）：`S/T` 门槛破零标准/达成、`V/W` 标准人力、`Y/Z` 英雄人力(全月)、`AF/AG` 英雄50%、`AN/AO` 有效人力、`AZ` 是否考核有效人力、`Q/R` 当月签单实算业绩/利润
- 统一筛选：**差距 = 目标值 − 利润业绩(R)**，且 **业绩 Q > 0** 且 **0 < 差距 < 3000（严格小于）**，按差距升序

### 事业部汇总（重点，用户明确要求过）

汇总图/文案**直读官方 `事业部汇总` sheet**（不自己聚合明细，明细行数含非月初在职行会偏大）：

- sheet 表头在第 2 行，**按列名定位**（防列位位移）；取 `事业部` 名称以「」开头的行 + 名称恰为「」的合计行
- 人数口径**与官方一致**：`月初企服顾问人数`（三力率的分母）、`考核人数`（有效占比的分母）
- 展示 7 列：`事业部 / 月初人数 / 考核人数 / 英雄人力 / 标准人力 / 破零人力 / 有效人力`，每格「N人 / 率%」两行
- 率：**英雄/标准/破零 ÷ 月初人数；有效占比 ÷ 考核人数**（直接取 sheet 的率列）
- ⚠️ 官方「标准人力」= 明细 **W 列**，**不等于**「英雄50%」(AG 列)；正文恭喜文案的 50% 仍按 AG，两者不同指标属正常
- sheet 缺失/表头对不上 → 自动回退明细聚合（月初≈明细行数、考核人数 = `AZ=="是"`），state 里 `summary_source` 会标明

### 文案规则

- 数字不带千分位；用真实 emoji（`🎉`），不用 `[庆祝]` 短代码
- 表内不用 emoji（微软雅黑无字形会变方框 □），改用「已达成」文本徽章
- `report_text_summary.txt` 为汇总文案（最前一条），`report_text1..4.txt` 为四段汇报文案

**产出**：`report_img/hero_summary.png`（汇总）、`hero_done.png`、`hero_gap50.png`、`hero_breakzero.png`、`hero_valid.png`；`report_state.json`；发完更新 `watch_state.json`

---

## 二、活动量战报（追踪 / 结果）· 2026-10-09 换源版

**数据源已换**（旧看板 `hzdash-*dashboard.example.com/dash/data.json` 口径**作废**）：
`https://dashboard.example.com/p/LLT6AK0VUUvnPu12Cc8PKZ`（活动量工作台 · 七岗排行榜）。
取数见 `wb_source.py`：先取分享页里的 `publishVersion`，再下载静态资源 `七岗排行榜.html`（约 16 MB，内嵌 `var RB` / `var DASH` / **`var RBE`**）；
**版本号每次重新发布会变，禁止硬编码**。

**汇报范围（用户确认，勿擅自扩大）**：只统计****的 **网销岗 / 电邀岗（=电销）/ 综法岗** 三岗。
代码开关：`wb_source.py` 顶部 `POSTS`、`render_wb.py` 的 `POST_ORDER`。

**⚠️ 口径 = 页面顶部的【日】档，且严格取「当天」（today）**（用户 2026-10-09 定档、2026-10-10 收紧）：
页面顶部有「日期区间 + 全周期/周期/日」三档。

- **只取今天这一天（today）的记录，历史与其他日期的数据一律排除**。实现：`render_all()` 调
  `wb_source.fetch_day(day=<今天 YYYYMMDD>)`，**三岗统一用当天**，不再按岗位各自回溯锚点。
- 某岗当天**没有**考核记录时（综法岗常见，可能几天才有一条有效考核）：该岗出一张
  「当日（YYYY-MM-DD）无考核数据」占位图，整体表对应行显示「当日无数据」——**保证 4 张图固定产出**。
  `build_post_panel` 对 `hz[岗]` 缺失已做兜底；`wb_source.day_anchors()` 的旧「各岗各自锚点」逻辑只在
  不传 `day` 时才会走到（保留供手动排查，正常链路不会用）。
- 页面数值**不是静态的**——不同档位由页面内嵌引擎 `RBE`（工作台 `gen_rank_all/gen_rank5` 的 JS 复刻）按区间重算，
  所以必须调同一套引擎，不能自己按人天原料拍脑袋聚合。实现：`wb_source.fetch_day()` 把 `RB/DASH` 落盘、
  把 `RBE` 源码抽到 `_cache/rbe.js`，再用受管 node 跑 `wb_engine.js`（网销/电邀→`RBE.build`，综法→`RBE.build4`），
  输出 `children['分公司']['']` + `children['部门'][部门]` 并整理成与全周期版相同的 `hz[岗]` 结构。

> 校验基准（2026-10-09，页面 v492）：引擎输出与页面截图逐格一致 —— 网销岗 在职 269 / 考核 146 /
> 24:00 达标率 58.2% / 冲刺率 1.6% / 日均通时 47.8 分。改引擎调用后**必须这样对一次**再上线。

**排障**：`fetch_day` 报「未找到 `var RBE`」/「未找到 RBE 导出尾」= 页面结构变了（抽取锚点失效）；
报「引擎无输出」= node 路径或 `_cache/rbe.js` 有问题；某岗返回 `ok:false` = 该区间该岗没有有效考核日
（`RBE.build` 会返回 null，属正常，取数侧会跳过该岗）。**改完务必打印三岗 `ok/vd` 自查。**

**有定时任务**（见下）。一键命令：

```bash
bash "<PROJECT_ROOT>/数据文件/活动量/_report/run_report.sh" track   # 追踪：未达标人员 + 4 图
bash "<PROJECT_ROOT>/数据文件/活动量/_report/run_report.sh" result  # 结果：当日全员 + 4 图
```

**流程**：`render_wb.py <mode>` 渲染 **4 图** → sessions list → 上传 4 图 → `build_send_payloads.py <mode> [后缀]` 装配 → webhook @全体 → 发送 **5 条** → `show_result.py <mode>`。

- 4 图：`act1_overview_*`（整体：三岗 + 合计，含**各岗数据日期**列）、`act2_wangxiao_*`（网销）、`act3_dianxiao_*`（电邀）、`act4_zongfa_*`（综法）；部门行 + 其下人员行两级。
- **track**：只列**未达标**人员；文案按岗位分组：`网销（24:00 达标率）… / 电邀（6 项任一达标）… / 综法（19:30 达标率）…`
- **result**：列全员（含达标/未达标/未考核徽章）。
- 单日口径下三岗单位一致（人），整体表合计行把三岗一起加；各岗日期不同时（综法落后）单独标注。

**达标口径**：网销 24:00 达标率 ≥100%、综法 19:30 达标率 ≥100%（源站口径）；
**电邀 = 6 项中「任意一项」达标即算达标**（用户口径，2026-10-10）；考核为 0 = 未考核（不计入未达标）。

> **电邀 6 项**（页面引擎 `RBE` 的 `N6` 常量，各为 0/1 标记）：加微≥5 ｜ 精准审核≥1 ｜ 业绩≥1000 ｜
> 见客≥1 ｜ 意向≥3 ｜ 45s≥20（2026-10-08 起第 6 项由「接触数≥135」改为「45s≥20」）。
> **网页默认是「6 项全中」**，我们按用户要求改成「任一项 ≥1」，所以**不要**直接用引擎给的
> `n6pass` / `n6rate`——用 `render_wb.dx_any()` 按人员明细重算（部门级与整体表都重算），
> 图③脚注会写明这条口径，避免和网页对不上时被误判成脚本错误。
> 实证（2026-10-09，电邀）：旧口径「6 项全中」达标 0 人 → 新口径「任一项」达标 1 人（徐伶 1/6）。

⚠️ 新源**没有**「距达标还差 N 分钟」这个概念；旧脚本 `render_undone.py`（读旧看板）**已停用**，只作历史参照，**勿回退**。

**试发到其它群**：`REPORT_GROUP="3C认证" bash run_report.sh track`（环境变量覆盖目标群，默认「数据点检群」）。

**文案格式（已按用户要求改定，勿回退）**：

```
今日{H}点{M}分{活动量追踪|结果}数据（数据周期：单日 YYYY-MM-DD）：https://dashboard.example.com/p/LLT6AK0VUUvnPu12Cc8PKZ ({说明}
未达标 {N} 人（网销 {a} · 电邀 {b} · 综法 {c}）· 当前达标率：
网销（24:00 达标率）：@张玄琳 0.0% ｜ @王玉 0.0% ｜ …
电邀（6 项任一达标）：@庄灿佳 0/6 ｜ @张奕鑫 0/6 ｜ …
综法（19:30 达标率）：@… （综法当日无数据/全员达标时该行不出现）
```

- **数据周期写「单日 + 当天日期」**（如 `单日 2026-10-10`）；仅当某岗日期与主日期不同才追加「· 综法 MM/DD」（统一当天后一般不会出现）。

- 链接**后必须留一个空格**（`host (`）——否则企业微信会把紧跟的 `(` 吞进超链接，点开链接错误。
- 追踪说明：`未达成活动量人员最晚需要19:30下班，如提前达成则可以提前下班`
- 结果说明：`当日全员结果如下，已达标人员可安排下班`（2026-10-10 用户要求**删掉**「未达标人员最晚19:30下班」）
- 链接常量在 `build_send_payloads.py` 的 `LINK`；改链接只改这里（脚本每次会重新生成 `p_caption.json`）
- ⚠️ `活动量/8月/8.11/deploy_tool/` 里的旧 host 是**部署可达性校验 URL**，与公告文案无关，勿混改

---

## 三、PK 战报

**有定时任务**（4 条）。一键命令：

```bash
"<HOME>/.workbuddy/binaries/python/envs/default/Scripts/python.exe" \
  "<PROJECT_ROOT>/数据文件/PK追踪/pk_report.py" \
  --group "管理者大群" --group "管理者群（勇闯奖金包" --group "分公司经理内部群"
```

**流程**：① `sync_yj.one_cycle(notify=False)` 静默拉 BI 源数据、替换在线表【业绩（含续费）】并做变化检测 → ② 有变化才截图（`capture_pk2.py` 整页 → `crop_with_posters()` 裁出含双海报的 2400×宽图）+ 生成文案 → ③ 发送到三个群。**无变化自动跳过，属正常，不算故障。**

**跨目录依赖**（改路径时注意）：

- `sync_yj.py` 在 `业绩BI同步/`（PK 脚本把该目录 insert 进 sys.path）
- `capture_pk2.py` 与 `core.py` 在 `20260924_tencent-doc-export/`（`scratch/` 与 `src/`）
- 截图依赖 Edge + 登录态 `20260924_tencent-doc-export/.tdexporter/auth_state.json`；登录态失效则截图拿不到内容

**裁剪口径**：扫描表格底部以下的宽幅非白色带，取最后一条高度≥15px 的带 +20px 为底缘；宽度保持全宽 2400 以保住右侧海报。

---

## 四、故障排查（实战踩坑，优先级高）

### 1. wecom-cli 授权过期（`errcode 850003`）

**症状**：`sessions list` 返回 `{"errcode":850003,"errmsg":"authorization expired"}`；上层表现为 `assemble` 报「未在会话列表找到含 xxx 的群」。**先怀疑授权，别急着改 GROUP_KEY。**

**关键陷阱**：`wecom-cli auth show` 可能仍显示 `Status: authorized`——那只是**本地凭据存在**，不代表上游授权有效。以 `sessions list` 的 errcode 为准。

**修复**（需用户手机扫码）：

```bash
wecom-cli auth init --noninteractive --output-qrcode auth_reauth.png
```

后台运行（会阻塞等待扫码约数分钟），二维码生成后用 present_files 给用户扫。扫码成功输出 `✔ 扫码成功！Bot ID 和 Secret 已自动获取`。扫完先验证 `sessions list` 正常再继续发送。

### 2. node 运行时路径失效（`FileNotFoundError: [WinError 2]`）

**症状**：Python 脚本（非 bash）调 wecom 时崩 `FileNotFoundError`。**bash 里直调 `wecom-cli` 正常，极易误判为授权问题。**

**根因**：脚本里硬编码了 `node/versions/<旧版本>/node.exe`，而受管运行时升级后旧版本目录被删除。检查方式：`grep -rn "node/versions/2" <项目>`。

**修复**：改为自动探测（已落地在 `sync_yj.py` 的 `_resolve_node_exe()`）：

```python
def _resolve_node_exe():
    import glob
    base = r"<HOME>/.workbuddy/binaries/node/versions"
    for p in sorted(glob.glob(base + "/*/node.exe"), reverse=True):
        if os.path.exists(p):
            return p
    return "node"   # 兜底走 PATH
```

**前提**：`WECOM_JS` 指向 `node/cli-connector-packages/node_modules/@wecom/cli/bin/wecom.js`，Python 里以 `subprocess.run([NODE_EXE, WECOM_JS] + args)` 调用。

### 3. 同步锁「另一个同步进程运行中」

`sync_yj.lock` 在 **900s 内**存在即拒绝并发并 `sys.exit(0)`。**这是正常并发保护，不是故障**；超 900s 会自动覆盖。手动连跑多次会互相挡住，间隔开即可。**不要因为看到这条就当故障去删锁。**

### 4. Python subprocess 跑不了 `wecom-cli`（.CMD shim）

`shutil.which("wecom-cli")` 在 Python 侧返回的是 `wecom-cli.CMD`，而 Windows `CreateProcess`
**不能直接执行 .CMD/.BAT** → `subprocess.run(["wecom-cli", ...])` 抛 `FileNotFoundError/WinError`。
表现是「功能静默失效」（如 `mention_util` 把 `_disabled` 置 True、@ 全部退化成纯文本）。

**修复**：Python 里一律 **node 直调** `wecom.js`（glob 探测 node 版本 + 兜底 shell）：
```python
NODE_BASE = "<HOME>/.workbuddy/binaries/node"
WECOM_JS = os.path.join(NODE_BASE, "cli-connector-packages",
                        "node_modules", "@wecom", "cli", "bin", "wecom.js")
# 见 mention_util._base_cmd() / sync_yj._resolve_node_exe()
```

### 5. 半发风险控制

跑多步流水线时，**先失败先停**：装配（找群/取 media_id）失败应中止，避免只发了一部分。三力/活动量脚本用 `|| exit N` 顺序编排，天然满足；自定义流程要自己保证。

### 6. 沙箱/工具限制（本机实测）

- Bash 工具缺 `sleep`/`dirname`，需要时用 Python（`time.sleep`/`os.path`）
- 某些 `python -c` 命令与 PowerShell 的 `Get-Process` 会被沙箱误拦（报 `sandbox-center cmd decisionRecord missing actual resource subject`）→ 改用 Read/Glob/`ls` 绕开
- 输出乱码多为 `tail -c` 截断多字节 UTF-8 的假象，用 Python 真读文件为准

---

## 五、定时任务清单（8 条 ACTIVE）

| 任务 | 频率 | 命令要点 |
|---|---|---|
| PK 战报 9:03 / 11:03 / 14:03 / 16:03 | 每天 4 次 | `pk_report.py --group ...`（3 群） |
| 活动量追踪 11:03 / 14:33 / 16:03 | 每天 3 次 | `run_report.sh track` |
| 活动量结果 18:03 | 每天 1 次 | `run_report.sh result` |

（另有看板类任务处于 PAUSED，不在本技能范围。）

**体检（不发送）**：静态核验脚本存在性 + `py_compile` + Python 依赖 import；动态核验 `auth show`/`sessions list`、目标群名匹配、截图链路、渲染产物。**跳过会改动线上数据的环节**（如完整 BI 拉数），并明确告知用户哪些没测。

---

## 交付与沟通约定

- 用户偏好中文、简洁；给「文件清单 + 操作流程 + 结果摘要」
- 群发类操作**谨慎**：默认不重复群发；用户说「仅补发新增的」就只发增量
- 改动脚本后**只改不发**时，务必在回复里写明「未发送，下次触发自动生效」
- 每次发完更新 `watch_state.json`（`last_ymd`/`last_file`/`last_sent_at`）
