# -*- coding: utf-8 -*-
"""从 sessions.json / up_img.json 读 chat_id 与 media_id，把【事业部文案】装配成 send 负载。
用法：build_pk_payloads.py [群关键字]   默认=分公司经理内部群
"""
import json
import os
import sys

DIR = os.path.dirname(os.path.abspath(__file__))
GROUP = sys.argv[1] if len(sys.argv) > 1 else "分公司经理内部群"


def jload(p):
    if not os.path.exists(p):
        return {}
    raw = open(p, encoding="utf-8").read()
    try:
        return json.loads(raw[raw.find("{"): raw.rfind("}") + 1])
    except Exception:
        return {}


def main():
    obj = jload(os.path.join(DIR, "sessions.json"))
    cid = cname = None
    for s in obj.get("sessions", []):
        if GROUP in str(s.get("chat_name", "")):
            cid, cname = s["chat_id"], s["chat_name"]
            break
    if not cid:
        raise RuntimeError(f"未在会话列表找到含 '{GROUP}' 的群")

    bu = open(os.path.join(DIR, "事业部文案.txt"), encoding="utf-8").read()
    json.dump({"chat_id": cid, "msgtype": "markdown", "markdown": {"content": bu}},
              open(os.path.join(DIR, "p_bu.json"), "w", encoding="utf-8"), ensure_ascii=False)

    up = jload(os.path.join(DIR, "up_img.json"))
    mid = up.get("media_id") if isinstance(up, dict) else None
    if mid:
        json.dump({"chat_id": cid, "msgtype": "image", "image": {"media_id": mid}},
                  open(os.path.join(DIR, "p_img.json"), "w", encoding="utf-8"), ensure_ascii=False)
    print("目标群:", cname, "文案已装配;", ("含截图" if mid else "仅文本"))


if __name__ == "__main__":
    main()
