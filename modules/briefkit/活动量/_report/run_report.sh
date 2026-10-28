#!/usr/bin/env bash
# 活动量看板企微群汇报（默认「数据点检群」）—— 一键执行
# 用法：bash run_report.sh [track|result] ["可选文案后缀"]
#   track  追踪模式（11:00 / 14:30 / 16:00）：文案（未达标人员 + 当前达标率）+ 4 张图（整体 / 网销 / 电邀 / 综法）
#   result 结果模式（18:00）：文案 + 4 张图（全员）
# 数据源：WorkBuddy 资料库发布页「活动量工作台 · 七岗排行榜」
#         https://dashboard.example.com/p/LLT6AK0VUUvnPu12Cc8PKZ
#         只统计的 网销岗 / 电邀岗 / 综法岗；数据日期 = 当天
# 环境变量 REPORT_GROUP 可覆盖目标群（例：REPORT_GROUP="3C认证" 先试发测试群）
set -euo pipefail

DIR="<PROJECT_ROOT>/数据文件/活动量/_report"
PY="<HOME>/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
GROUP="${REPORT_GROUP:-数据点检群}"
MODE="${1:-track}"
SUFFIX="${2:-}"
case "$MODE" in
  result|R|res|r) SUF=result; MODE=result ;;
  *)              SUF=track;  MODE=track  ;;
esac
cd "$DIR"

echo "[1/4] 取数 + 渲染 4 张报表图（模式: $MODE）..."
"$PY" render_wb.py "$MODE"

echo "[2/4] 拉取群会话列表 ..."
wecom-cli message aibot sessions list > sessions.json
"$PY" -c "import json,io,sys;sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8',errors='replace');d=json.load(open('sessions.json',encoding='utf-8'));print('  可发送会话数:',d.get('sessions_count'),[s['chat_name'] for s in d.get('sessions',[])])"

echo "[3/4] 上传 4 张图获取 media_id ..."
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/act1_overview_$SUF.png\",\"type\":\"image\"}" > up1.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/act2_wangxiao_$SUF.png\",\"type\":\"image\"}" > up2.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/act3_dianxiao_$SUF.png\",\"type\":\"image\"}" > up3.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/act4_zongfa_$SUF.png\",\"type\":\"image\"}" > up4.json

echo "[4/4] 组装文案并发送（目标群: $GROUP）..."
REPORT_GROUP="$GROUP" "$PY" build_send_payloads.py "$MODE" "$SUFFIX"

echo "  → 群机器人 Webhook @ 全体（aibot 无法真@全体，单独用 webhook 发）"
"$PY" send_atall.py "$GROUP" "📢 分公司活动量战报已更新，请各位伙伴查看下方报表 👇" || true

wecom-cli message aibot send --json "$(cat p_caption.json)" > send1.json
wecom-cli message aibot send --json "$(cat p_img1.json)"    > send2.json
wecom-cli message aibot send --json "$(cat p_img2.json)"    > send3.json
wecom-cli message aibot send --json "$(cat p_img3.json)"    > send4.json
wecom-cli message aibot send --json "$(cat p_img4.json)"    > send5.json
"$PY" show_result.py "$MODE"
