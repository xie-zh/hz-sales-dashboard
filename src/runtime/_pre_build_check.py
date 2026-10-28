# -*- coding: utf-8 -*-
"""构建前预检 —— 确认四个源表「体积稳定 + 是完好的 xlsx」。

为什么要预检：WPS/Excel 正在写盘时体积还在变，此时构建会读到半个文件；
或者文件其实是损坏的（崩溃中断过）。两种都会让看板出现莫名奇妙的空值，
而且很难回查。这个脚本只读、不写，跑一遍 6 秒，稳了就构建。

V1.5（2026-10-12）：路径不再写死 —— 源表定位交给 _portable，
整个目录搬到别的电脑照样能用。

用法：python _pre_build_check.py
退出码：0 = 可以构建；1 = 有表不稳定/损坏
"""
import os
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Hzdggbot"))

try:
    import _portable
    BASE = _portable.resolve_root()
except Exception:
    BASE = HERE

# 四个源表：三个在项目根，辅助表统一收进 <根>\分公司健康表\（换机自包含）
PATHS = [
    os.path.join(BASE, "分公司活动量-在线看板.xlsx"),
    os.path.join(BASE, "分公司-员工三力明细.xlsx"),
    os.path.join(BASE, "分公司-人员数据分析路径.xlsx"),
    os.path.join(BASE, "分公司健康表", "四大模式-重点指标表格.xlsx"),
]


def main():
    print("项目根：%s" % BASE)
    missing = [p for p in PATHS if not os.path.isfile(p)]
    if missing:
        for p in missing:
            print("× 缺文件：%s" % p)
        print("\nSAFE_TO_BUILD = False")
        return 1

    s1 = [(p, os.path.getsize(p)) for p in PATHS]
    time.sleep(6)
    s2 = [(p, os.path.getsize(p)) for p in PATHS]

    ok = True
    for (p1, a), (p2, b) in zip(s1, s2):
        stable = (a == b)
        try:
            z = zipfile.ZipFile(p1)
            bad = z.testzip()
            zi = len(z.namelist())
            z.close()
        except Exception as e:
            bad, zi = str(e), -1
        good = stable and (bad is None) and zi > 0
        ok = ok and good
        print("%-34s size=%-9d stable=%-5s entries=%-4d testzip=%s -> %s" % (
            os.path.basename(p1), b, stable, zi, bad, "OK" if good else "FAIL"))
    print("\nSAFE_TO_BUILD =", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
