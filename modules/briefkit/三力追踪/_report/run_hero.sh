#!/usr/bin/env bash
# 一键：取数 → 渲染五张图(+事业部汇总图) → 取会话 → 上传六张图 → 发送到「数据点检群」
#   0 群机器人 Webhook @ 全体（aibot 无法真@全体，单独用 webhook 发，指向汇总）
#   1 事业部汇总文案  2 事业部汇总表图  3 恭喜英雄人力文案  4 英雄完成情况图
#   5 英雄差距文案    6 英雄差距榜图    7 破零催办文案    8 破零名单图
#   9 有效人力文案   10 有效人力差距图  11 英雄人力特权海报（永远保持在最后）
DIR="<PROJECT_ROOT>/数据文件/三力追踪/_report"
PY="<HOME>/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
cd "$DIR" || exit 1

echo "=== 1) 取数 + 渲染（含事业部汇总图）==="
"$PY" build_hero_report.py || exit 2

echo "=== 2) 会话列表 ==="
wecom-cli message aibot sessions list > sessions.json

echo "=== 3) 上传六张图（up0=事业部汇总图，up1..up5=原五图）==="
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/hero_summary.png\",\"type\":\"image\"}" > up0.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/hero_done.png\",\"type\":\"image\"}" > up1.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/hero_gap50.png\",\"type\":\"image\"}" > up2.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/hero_breakzero.png\",\"type\":\"image\"}" > up3.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/report_img/hero_valid.png\",\"type\":\"image\"}" > up4.json
wecom-cli media upload --json "{\"file_path\":\"$DIR/assets/hero_privilege.jpg\",\"type\":\"image\"}" > up5.json

echo "=== 4) 装配负载（11 条，汇总在最前，特权海报在最后）==="
"$PY" assemble_hero.py || exit 3

echo "=== 4.5) 群机器人 Webhook 真 @ 全体（指向本期汇总，aibot 无法真@全体）==="
"$PY" send_atall.py 数据点检群 "📢 分公司三力/有效人力战报已更新（含各事业部汇总），请各位伙伴查看 👇" || true

echo "=== 5) 发送 11 条（aibot 富文本+图片）==="
wecom-cli message aibot send --json "$(cat p_t0.json)" > s1.json
wecom-cli message aibot send --json "$(cat p_i0.json)" > s2.json
wecom-cli message aibot send --json "$(cat p_t1.json)" > s3.json
wecom-cli message aibot send --json "$(cat p_i1.json)" > s4.json
wecom-cli message aibot send --json "$(cat p_t2.json)" > s5.json
wecom-cli message aibot send --json "$(cat p_i2.json)" > s6.json
wecom-cli message aibot send --json "$(cat p_t3.json)" > s7.json
wecom-cli message aibot send --json "$(cat p_i3.json)" > s8.json
wecom-cli message aibot send --json "$(cat p_t4.json)" > s9.json
wecom-cli message aibot send --json "$(cat p_i4.json)" > s10.json
wecom-cli message aibot send --json "$(cat p_i5.json)" > s11.json

echo "=== 6) 结果 ==="
"$PY" show_hero_result.py
