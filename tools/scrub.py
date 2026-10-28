# -*- coding: utf-8 -*-
"""脱敏脚本骨架（公开版）。

真实姓名清单不随仓库分发，需放在 `names.json`（已被 .gitignore 排除）：
    { "employees": ["张三", "李四", ...],
      "vendors":   ["某某科技有限公司", ...] }

用法：
    python tools/scrub.py --deploy <真实产物目录> --out <脱敏输出目录>

流程：结构化映射 → 全文本替换 → 资源修补 → 残留校验。
"""
import argparse
import hashlib
import json
import os
import re
import shutil

XING = list("赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜")
MING = list("嘉睿思彦宇航欣怡泽轩梦琪晓然明辉志强建国婷丽雪倩雅涵鑫文浩子墨诺沐宸清歌")
HERE = os.path.dirname(os.path.abspath(__file__))
NAME_FILE = os.path.join(HERE, "names.json")

_name_map = {}
NAME_PAIRS = []


# ------------------------------------------------------------------ 映射
def fake_name(real):
    """确定性映射：同一真实姓名恒得到同一虚构姓名，保证跨页面关联不断。"""
    real = str(real or "").strip()
    if not real:
        return "—"
    if real not in _name_map:
        h = hashlib.md5(("nm" + real).encode()).hexdigest()
        x = XING[int(h[:4], 16) % len(XING)]
        n = 2 + (int(h[4:6], 16) % 2)
        m = "".join(MING[int(h[6 + i * 2:8 + i * 2], 16) % len(MING)] for i in range(n))
        _name_map[real] = x + m
    return _name_map[real]


def fake_id(real):
    """工号重新编号，避开原号段。"""
    h = int(hashlib.md5(("gh" + str(real)).encode()).hexdigest()[:8], 16)
    if h % 97 == 0:
        h += 1
    return str(h % 900000 + 100000)


def load_names():
    """从外部文件载入真实姓名全集；文件不入库。"""
    global NAME_PAIRS
    if not os.path.exists(NAME_FILE):
        raise SystemExit("缺少 names.json —— 真实姓名清单不随仓库分发，请自行准备")
    d = json.load(open(NAME_FILE, encoding="utf-8"))
    names = [n for n in d.get("employees", []) if is_person(n)]
    NAME_PAIRS = [(n, fake_name(n)) for n in sorted(names, key=len, reverse=True)]
    # 人员类局部替换只在长姓名优先的前提下做全局替换
    return len(names)


# 业务词黑名单：避免把「网销模式」「开场介绍」这类词当成人名替换掉
NOT_NAME = (
    "模式", "事业部", "业务部", "分公司", "公司", "中心", "团队", "部门", "资源",
    "业绩", "看板", "总表", "明细", "汇总", "分析", "指标", "考核", "达标", "合格",
    "人力", "英雄", "破零", "标准", "无效", "有效", "网销", "电邀", "会计", "法律",
    "跨境", "工商", "交易", "数智", "付费", "客户", "主管", "经理", "规划师", "顾问",
    "开场", "介绍", "需求", "挖掘", "异议", "处理", "加微", "埋点", "引导", "沟通",
)


def is_person(s):
    s = (s or "").strip()
    if not (2 <= len(s) <= 4):
        return False
    if re.search(r"[0-9A-Za-z（）()·\-—/、，,。\s]", s):
        return False
    return not any(w in s for w in NOT_NAME)


# ------------------------------------------------------------------ 文本
ORG_RE = re.compile(r"(?:<CITY>)?(?=[\u4e00-\u9fa5A-Za-z0-9]{1,14}(?:事业部|业务部|部|中心))")
ENDPOINT_RULES = [
    (r"https?://[\w.-]*<INTERNAL_DOMAIN>[^\"'\s)]*", "https://bi.example.com"),
    (r"https?://[\w.-]*tcloudbaseapp\.com[^\"'\s)]*", "https://dashboard.example.com"),
    (r"https?://[\w.-]*workbuddy\.(?:host|link)[^\"'\s)]*", "https://dashboard.example.com"),
]


def scrub_text(s):
    """最后一道防线：所有文本都过这里，保证深层嵌套里的人名也被替换。"""
    if not isinstance(s, str):
        return s
    for real, fake in NAME_PAIRS:
        if real in s:
            s = s.replace(real, fake)
    for pat, rep in ENDPOINT_RULES:
        s = re.sub(pat, rep, s)
    s = re.sub(r"[A-Za-z]:\\\\?Users\\\\?[^\\\"'\s)]+", "<LOCAL_PATH>", s)
    s = re.sub(r"\b[A-Z]{2,6}\d{8,}\b", "<SERIAL_NO>", s)      # 合同号 / 订单号
    return s


def scrub_obj(o):
    if isinstance(o, dict):
        return {scrub_text(k): scrub_obj(v) for k, v in o.items()}
    if isinstance(o, list):
        return [scrub_obj(v) for v in o]
    if isinstance(o, str):
        return scrub_text(o)
    return o


def scrub_records(rows):
    out = []
    for r in rows:
        if not isinstance(r, dict):
            out.append(scrub_obj(r))
            continue
        n = dict(r)
        for k in list(n.keys()):
            if k in ("工号", "员工工号", "id"):
                n[k] = fake_id(n[k])
            elif ("姓名" in k or k in ("name", "person", "员工")) and isinstance(n[k], str):
                n[k] = fake_name(n[k]) if is_person(n[k]) else n[k]
            elif k == "入职日期":
                n[k] = "20XX-01-01"
        out.append(scrub_obj(n))
    return out


# 需要置空而非删除的敏感字段（删除会让前端读 undefined 抛错，整页空白）
BLANK_KEYS = {"audioUrl": "", "voiceUrl": "", "audio": "", "rawTranscript": "",
              "transcript": [], "segments": [], "dialog": [], "conversation": []}


def blank_sensitive(recs):
    n = 0
    for r in recs:
        for k, v in BLANK_KEYS.items():
            if k in r:
                r[k] = "" if isinstance(r[k], str) else v
                n += 1
    return n


# ------------------------------------------------------------------ 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--deploy", required=True, help="真实产物目录")
    ap.add_argument("--out", required=True, help="脱敏输出目录")
    a = ap.parse_args()

    cnt = load_names()
    print("姓名映射：%d 个" % cnt)

    demo = os.path.join(a.out, "demo")
    if os.path.isdir(demo):
        shutil.rmtree(demo)
    os.makedirs(demo)

    for f in os.listdir(a.deploy):
        p = os.path.join(a.deploy, f)
        if f.endswith(".html"):
            h = scrub_text(open(p, encoding="utf-8", errors="ignore").read())
            h = h.replace('src="/static/', 'src="static/')     # 绝对路径改相对
            open(os.path.join(demo, f), "w", encoding="utf-8").write(h)
        elif f.endswith(".json") and f not in ("status.json", "_publish_request.json"):
            try:
                d = json.load(open(p, encoding="utf-8"))
            except Exception:
                continue
            if isinstance(d, dict) and isinstance(d.get("recordings"), list):
                blank_sensitive(d["recordings"])
            d = scrub_obj(d)
            open(os.path.join(demo, f), "w", encoding="utf-8").write(
                json.dumps(d, ensure_ascii=False))
            print("  %s" % f)

    # 静态资源必须一起带走，否则图表全空
    st = os.path.join(a.deploy, "static")
    if os.path.isdir(st):
        shutil.copytree(st, os.path.join(demo, "static"))
        print("  static/ 已复制")

    print("完成 → %s" % demo)


if __name__ == "__main__":
    main()
