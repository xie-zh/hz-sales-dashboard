#!/usr/bin/env bash
# 发送 PK 追踪报表到指定群：事业部文案(markdown) + 截图(image) + webhook @全体
#   ./run_pk.sh [full|text] [群关键字]
#   例：bash run_pk.sh full 管理者群   (发「管理者大群」)
#       bash run_pk.sh full 分公司经理内部群
#       bash run_pk.sh text 管理者群   (仅文案，不重发图)
# 注：aibot 无法真@全体，@全体由「群机器人 webhook」负责（webhooks.json 缺则跳过）。
DIR="<PROJECT_ROOT>/数据文件/PK追踪"
PY="<HOME>/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
cd "$DIR" || exit 1

MODE="${1:-full}"
GROUP="${2:-分公司经理内部群}"

echo "=== 1) 会话列表 ==="
wecom-cli message aibot sessions list > sessions.json

if [ "$MODE" != "text" ]; then
  echo "=== 2) 上传截图 ==="
  wecom-cli media upload --json "{\"file_path\":\"$DIR/pk_事业部PK追踪_10月.png\",\"type\":\"image\"}" > up_img.json
fi

echo "=== 3) 装配负载（群=$GROUP）==="
"$PY" build_pk_payloads.py "$GROUP" || exit 1

echo "=== 4) 发送 事业部文案 ==="
wecom-cli message aibot send --json "$(cat p_bu.json)" > s_bu.json

if [ "$MODE" != "text" ]; then
  echo "=== 5) 发送 截图 ==="
  wecom-cli message aibot send --json "$(cat p_img.json)" > s_img.json
fi

echo "=== 6) webhook @全体（缺则跳过）==="
"$PY" send_atall.py "$GROUP" "📢 分公司事业部PK战报已更新，请各位伙伴查看下方战报 👇" || true

echo "=== 完成（文案/截图已发到 $GROUP；@全体需群机器人 webhook）==="
