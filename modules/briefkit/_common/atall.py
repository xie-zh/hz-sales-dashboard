# -*- coding: utf-8 -*-
"""
企业微信群机器人(Webhook) @ 全体成员 工具
============================================
背景：本企业的「智能机器人(aibot)」无法通过 API 真正 @ 全体成员
      （markdown 不支持 @all；text 类型被服务端拒绝；普通 message send 被禁用 853006）。
      唯一可靠的 @ 全体方式是用「群机器人 Webhook」发 text + mentioned_list=['@all']。

用法：
    from atall import send_at_all, load_webhooks
    send_at_all(load_webhooks()['数据点检群'], '📢 数据已更新，请查看下方战报 👇')

webhooks.json 结构（由用户填写各群 webhook 地址）：
    {
      "数据点检群": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxx",
      "分公司经理内部群": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=yyyy",
      "3C认证团队":     "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=zzzz"
    }
"""
import json
import os
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
WEBHOOKS_FILE = os.path.join(HERE, "webhooks.json")


def load_webhooks():
    """读取 webhooks.json -> {群关键字: webhook_url}"""
    try:
        with open(WEBHOOKS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def send_at_all(webhook_url, text="📢 数据已更新，请各位伙伴查看下方汇报 👇"):
    """
    用群机器人 webhook 发送一条 text 消息并真正 @ 全体成员。
    返回 True / False。webhook 未配置时返回 False 并打印警告（不报错中断主流程）。
    """
    if not webhook_url:
        print("WARN: 该群 webhook 未配置，跳过 @ 全体（请在企业微信建群机器人并填入 webhooks.json）")
        return False
    body = json.dumps(
        {"msgtype": "text", "text": {"content": text, "mentioned_list": ["@all"]}},
        ensure_ascii=False,
    ).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=body, headers={"Content-Type": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            resp = json.loads(r.read().decode("utf-8"))
        print("webhook @全体 ->", resp)
        return resp.get("errcode") == 0
    except Exception as e:
        print("webhook @全体 失败:", e)
        return False


if __name__ == "__main__":
    import sys
    key = sys.argv[1] if len(sys.argv) > 1 else "数据点检群"
    txt = sys.argv[2] if len(sys.argv) > 2 else "📢 数据已更新，请各位伙伴查看下方汇报 👇"
    ok = send_at_all(load_webhooks().get(key, ""), txt)
    sys.exit(0 if ok else 1)
