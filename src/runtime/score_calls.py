# -*- coding: utf-8 -*-
"""
五维质检评分器 —— 基于腾讯云 ASR 转写文本，按长沙方法论自动打分
五维：开场介绍 / 痛点价值 / 需求挖掘 / 异议处理 / 加微埋点
输入：recordings_transcript.json（asr_tencent.py 产出）
输出：recordings_scored.json（供 gen_report.py 注入报告）
- 纯正则 + 关键词启发式，零依赖
- 每维 0~5 分；综合=五维均值（加权）
- 自动生成 优势(strengths) / 短板(weaknesses) / 改进建议(suggestions)
"""
import os, re, json

TRANS = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\recordings_transcript.json"
OUT = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\recordings_scored.json"

# 段落解析：[mm:ss.xxx,mm:ss.xxx] 文本
SEG_RE = re.compile(r'\[(\d+):([\d.]+),(\d+):([\d.]+)\]\s*(.*)')

def parse_segments(text):
    segs = []
    if not text:
        return segs
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = SEG_RE.search(line)
        if m:
            sh, sms, eh, ems = int(m.group(1)), float(m.group(2)), int(m.group(3)), float(m.group(4))
            start = sh * 60 + sms
            end = eh * 60 + ems
            content = m.group(5).strip()
            if content:
                segs.append({"start": round(start, 2), "end": round(end, 2), "text": content})
        else:
            # 无时间戳的整段（兜底）
            segs.append({"start": None, "end": None, "text": line})
    return segs

# ---------- 五维关键词库（贴合财税/工商电销话术） ----------
DIMS = {
    "开场介绍": {
        "weight": 1.0,
        "signals": {
            "问候": ["您好", "喂，你好", "你好", "上午好", "下午好"],
            "自报家门": ["我这边是", "我是", "我姓", "这边是给企业", "给您来电", "打电话"],
            "说明来意/由头": ["看您名下", "名下有一家", "有一家", "了解到您", "关注到您", "给您做个", "做个回访", "税务合规", "工商"],
        },
        "good": "开场清晰：问候+自报身份+说明来意/由头，让客户快速理解来电目的。",
        "bad": "开场偏弱：未清晰自报身份或说明来意，客户可能一头雾水。建议先问候、再报公司/姓名、再点出来电由头（如'看您名下有一家XX公司'）。",
    },
    "痛点价值": {
        "weight": 1.3,
        "signals": {
            "风险痛点": ["金税四期", "查", "补税", "成本票", "风险", "异常", "合规", "税务", "罚款", "多缴税", "抵扣", "挂靠", "虚开发票", "稽查"],
            "价值呈现": ["辅导", "服务", "解决", "专业", "帮您", "规避", "规划", "优化", "省心", "我们上周", "同行业", "案例"],
            "紧迫/稀缺": ["近期", "比较严重", "马上", "政策", "抓紧", "限时", "名额"],
            "公司背书": ["我们是企业", "集团", "公司", "平台", "团队", "服务过", "上市公司"],
        },
        "good": "痛点与价值呈现到位：既点出财税风险（金税四期/成本票/补税），又给出专业解决价值，有同行业案例增强信任。",
        "bad": "痛点/价值不够：未有效激发客户对财税风险（金税四期、成本票、补税）的焦虑，也缺少专业价值与案例背书。建议'风险钩子+价值承接'组合表达。",
    },
    "需求挖掘": {
        "weight": 1.1,
        "signals": {
            "提问": ["有没有", "遇到", "您目前", "您这边", "平时", "情况呢", "是吗", "对吧", "怎么", "多少", "做什么", "经营"],
            "倾听引导": ["方便说下", "您看", "您说", "具体", "详细", "了解下"],
        },
        "good": "需求挖掘有互动：通过提问了解客户实际经营与财税状况，而非单向宣讲。",
        "bad": "需求挖掘偏弱：多为单向介绍，缺少针对性提问，难以定位客户真实痛点。建议用'您目前成本票够不够用？''有没有遇到补税/异常？'等开放问题引导。",
    },
    "异议处理": {
        "weight": 1.0,
        "signals": {
            "识别异议": ["不需要", "不用", "考虑", "没时间", "忙", "已经", "太贵", "再说", "不用了", "暂时", "有会计", "自己弄"],
            "应对化解": ["理解", "明白", "没关系", "其实", "只是", "不妨", "先", "不打扰", "正好", "也行", "您看这样", "哪怕"],
        },
        "good": "异议处理有应对：客户有顾虑时能共情并给出化解话术，不急于强推。",
        "bad": "异议处理缺失：若客户有拒绝/顾虑，未见有效共情与化解，容易直接流失。建议'先共情（理解/没关系）+轻化解（哪怕先加个微信发资料）'。",
    },
    "加微埋点": {
        "weight": 1.2,
        "signals": {
            "提出加微": ["加微信", "加您微信", "加个微信", "微信号", "扫我", "加我"],
            "给出理由": ["发资料", "发一些", "相关资料", "案例", "政策", "发您", "方便看"],
            "约时间/埋点": ["方便吗", "您看方便", "后续", "跟进", "到时候", "过两天", "回头", "约", "时间"],
        },
        "good": "加微+埋点到位：明确提出加微信、给出'发资料'理由，并埋下后续跟进触点。",
        "bad": "加微/埋点不足：未有效引导加微信或约定后续跟进，错失沉淀客户的机会。建议'加个微信发XX资料，平时不打扰您，方便吗？'并约个回访时间。",
    },
}

def score_dimension(name, dim, full_text):
    """返回 0~5 分 与 命中的分组"""
    hits = {}
    groups = dim["signals"]
    for gname, kws in groups.items():
        found = [kw for kw in kws if kw in full_text]
        if found:
            hits[gname] = found
    # 评分逻辑：先看是否核心组命中，再按命中组数/关键词丰富度加成
    total_groups = len(groups)
    hit_groups = len(hits)
    if hit_groups == 0:
        return 0, hits
    # 基础分：命中1组=2，2组=3，3组=4，全部组=5
    base = {0: 0, 1: 2, 2: 3, 3: 4, 4: 5}.get(min(hit_groups, total_groups), 2)
    # 关键词丰富度微调：命中词越多略加
    total_kw = sum(len(v) for v in hits.values())
    if total_kw >= 4 and base < 5:
        base += 1
    return min(5, base), hits

def evaluate(rec):
    text = rec.get("text") or ""
    segs = parse_segments(text)
    full = text
    dims = []
    scores = []
    for name, dim in DIMS.items():
        sc, hits = score_dimension(name, dim, full)
        dims.append({"name": name, "score": sc, "hits": hits, "weight": dim["weight"]})
        scores.append(sc * dim["weight"])
    wsum = sum(d["weight"] for d in dims)
    overall = round(sum(scores) / wsum, 1)

    strengths = [d["name"] for d in dims if d["score"] >= 4]
    weaknesses = [d["name"] for d in dims if d["score"] <= 2]
    # 改进建议：短板维度给出对应 bad 文案；若全优给总体肯定
    suggestions = []
    for d in dims:
        if d["score"] <= 2:
            suggestions.append(DIMS[d["name"]]["bad"])
    if not suggestions:
        suggestions.append("整体话术较完整，建议保持开场清晰、痛点前置，并强化每次通话的加微与回访埋点。")

    # 否决项（一票否决式硬伤）：核心维度得 0 分即触发
    def _sc(name):
        return next(dd["score"] for dd in dims if dd["name"] == name)
    vetoes = []
    if _sc("开场介绍") == 0:
        vetoes.append("未自报身份/说明来意，违反外呼开场规范")
    if _sc("痛点价值") == 0:
        vetoes.append("未触达任何财税痛点，通话价值缺失")
    if _sc("加微埋点") == 0:
        vetoes.append("全程未引导加微信/约回访，客户资源未沉淀")

    return {
        "id": rec.get("id"),
        "person": rec.get("person"),
        "customer": rec.get("customer"),
        "url": rec.get("url"),
        "durationSec": rec.get("durationSec"),
        "time": rec.get("time"),
        "source": rec.get("source"),
        "dept": rec.get("dept", "工商业务事业部"),
        "status": rec.get("status", "ok"),
        "transcript": text,
        "segments": segs,
        "dims": dims,
        "overall": overall,
        "strengths": strengths,
        "weaknesses": weaknesses,
        "veto": vetoes,
        "suggestions": suggestions,
    }

def main():
    if not os.path.exists(TRANS):
        print("ERROR: 未找到", TRANS, "请先运行 asr_tencent.py 完成转写")
        raise SystemExit(1)
    raw = json.load(open(TRANS, encoding="utf-8"))
    out = [evaluate(r) for r in raw]
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"已评分 {len(out)} 条 -> {OUT}")
    # 简单统计
    avg = round(sum(x["overall"] for x in out) / len(out), 2)
    print("综合均分:", avg)
    for d in DIMS:
        s = round(sum(x["overall"] if False else next(dd["score"] for dd in x["dims"] if dd["name"] == d) for x in out) / len(out), 2)
        print(f"  {d} 平均分: {s}")

if __name__ == "__main__":
    main()
