# -*- coding: utf-8 -*-
"""汇总 11 条发送结果的 success 字段（汇总在最前，特权海报在最后）"""
import os
import json

DIR = os.path.dirname(os.path.abspath(__file__))
LABEL = [("s1.json", "文案·事业部汇总"), ("s2.json", "图·事业部汇总表"),
         ("s3.json", "文案·恭喜英雄人力"), ("s4.json", "图·英雄完成情况"),
         ("s5.json", "文案·英雄差距"), ("s6.json", "图·英雄差距榜"),
         ("s7.json", "文案·破零催办"), ("s8.json", "图·破零名单"),
         ("s9.json", "文案·有效人力"), ("s10.json", "图·有效人力差距"),
         ("s11.json", "图·特权海报")]


def ok(path):
    try:
        raw = open(path, encoding="utf-8").read()
        s = raw[raw.find("{"): raw.rfind("}") + 1]
        o = json.loads(s)

        def walk(x):
            if isinstance(x, dict):
                if "success" in x:
                    return x["success"]
                if "errcode" in x:
                    return x["errcode"] == 0
                for v in x.values():
                    r = walk(v)
                    if r is not None:
                        return r
            elif isinstance(x, list):
                for v in x:
                    r = walk(v)
                    if r is not None:
                        return r
            return None
        return walk(o)
    except Exception as e:
        return "ERR:%s" % e


def main():
    allok = True
    for fn, label in LABEL:
        p = os.path.join(DIR, fn)
        r = ok(p) if os.path.exists(p) else "缺失"
        if r is not True:
            allok = False
        print(f"  {label:<12} {fn:<10} success={r}")
    print("ALL_SUCCESS =", allok)


if __name__ == "__main__":
    main()
