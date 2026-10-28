# -*- coding: utf-8 -*-
"""装配发送负载：定位目标群 chat_id + 读 6 张图的 media_id + 5 段文案 -> 11 条消息负载
发送顺序（汇总永远在战报最前，特权海报永远在最后）：
  1 p_t0 事业部汇总文案   2 p_i0 事业部汇总表图   3 p_t1 恭喜英雄人力   4 p_i1 英雄完成情况图
  5 p_t2 英雄差距文案     6 p_i2 英雄差距榜图     7 p_t3 破零催办文案   8 p_i3 破零名单图
  9 p_t4 有效人力文案    10 p_i4 有效人力差距图  11 p_i5 英雄特权海报
注：真实 @全体 由「群机器人 webhook」负责（见 run_hero.sh 的 send_atall.py 步骤），
aibot markdown 的 <@all> 不可靠，故 p_t0 仅放汇总文案，不写字面 <@all>。"""
import os
import json

DIR = os.path.dirname(os.path.abspath(__file__))
GROUP_KEY = "数据点检群"    # 目标群关键字
# 注意：智能机器人(aibot)无法真正 @ 全体成员（markdown 不支持 @all、text 被服务端拒绝）。
# @ 全体改由「群机器人 Webhook」负责，见 run_hero.sh 中的 send_atall.py 步骤。


def jload(path):
    raw = open(path, encoding="utf-8").read()
    s = raw[raw.find("{"): raw.rfind("}") + 1]
    return json.loads(s)


def walk_media(o):
    if isinstance(o, dict):
        for k, v in o.items():
            if k.lower() == "media_id" and isinstance(v, str) and v:
                return v
            r = walk_media(v)
            if r:
                return r
    elif isinstance(o, list):
        for v in o:
            r = walk_media(v)
            if r:
                return r
    return None


def main():
    obj = jload(os.path.join(DIR, "sessions.json"))
    cid = cname = None
    for s in obj.get("sessions", []):
        if GROUP_KEY in str(s.get("chat_name", "")):
            cid, cname = s["chat_id"], s["chat_name"]
            break
    if not cid:
        raise RuntimeError("未在会话列表找到含 %r 的群" % GROUP_KEY)

    mid = {}
    for i in range(0, 6):
        mid[i] = walk_media(jload(os.path.join(DIR, "up%d.json" % i)))
    if not all(mid.values()):
        raise RuntimeError("media_id 缺失: %r" % (mid,))

    t0 = open(os.path.join(DIR, "report_text_summary.txt"), encoding="utf-8").read()
    t1 = open(os.path.join(DIR, "report_text1.txt"), encoding="utf-8").read()
    t2 = open(os.path.join(DIR, "report_text2.txt"), encoding="utf-8").read()
    t3 = open(os.path.join(DIR, "report_text3.txt"), encoding="utf-8").read()
    t4 = open(os.path.join(DIR, "report_text4.txt"), encoding="utf-8").read()

    def md(text):
        return {"chat_id": cid, "msgtype": "markdown", "markdown": {"content": text}}

    def img(i):
        return {"chat_id": cid, "msgtype": "image", "image": {"media_id": mid[i]}}

    payloads = {
        "p_t0.json": md(t0),                    # 1 事业部汇总文案（永远在首条；@全体由 webhook 负责）
        "p_i0.json": img(0),                    # 2 事业部汇总表图
        "p_t1.json": md(t1),                    # 3 恭喜英雄人力
        "p_i1.json": img(1),                    # 4 英雄人力完成情况图
        "p_t2.json": md(t2),                    # 5 英雄人力差距文案
        "p_i2.json": img(2),                    # 6 英雄人力差距榜图
        "p_t3.json": md(t3),                    # 7 破零人力催办文案
        "p_i3.json": img(3),                    # 8 破零人力名单图
        "p_t4.json": md(t4),                    # 9 有效人力文案
        "p_i4.json": img(4),                    # 10 有效人力差距榜图
        "p_i5.json": img(5),                    # 11 英雄人力特权海报（永远保持在最后）
    }
    for fn, obj2 in payloads.items():
        with open(os.path.join(DIR, fn), "w", encoding="utf-8") as f:
            json.dump(obj2, f, ensure_ascii=False)

    with open(os.path.join(DIR, "chat_id.txt"), "w", encoding="utf-8") as f:
        f.write(cid)

    print("目标群 :", cname, cid)
    for i in range(0, 6):
        print("media%d :" % i, mid[i][:20])
    print("text0  :", len(t0), "字 | text1:", len(t1), "字 | text2:", len(t2),
          "字 | text3:", len(t3), "字 | text4:", len(t4), "字")


if __name__ == "__main__":
    main()
