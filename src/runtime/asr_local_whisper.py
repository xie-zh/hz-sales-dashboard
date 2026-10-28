# -*- coding: utf-8 -*-
"""
本地 Whisper 转写 —— 替代腾讯云 ASR（腾讯云已无使用权限）
======================================================
- 数据源：通话记录.xlsx「通话记录」表（http 录音地址）
- 范围：仅转写「拨出时间/呼入时间 >= 2026-09-15」且 通话时长 >= MIN_SEC 的新增录音
- 模型：faster-whisper-small（本地路径 call_qc/scratch/whisper-small-local，CPU int8）
- 增量化：读取已有 recordings_transcript.json，按 url 跳过 status=="ok" 条目；
  转完一条立刻落盘，防中断丢进度。不会重复转写、不动旧数据。
- 输出：recordings_transcript.json（与 asr_tencent.py 同构：
  person/customer/url/durationSec/time/source/text/status + segments）
"""
import os, re, json, time, datetime
import sys
# fw_standalone 自包含环境优先（python-libs 目录枚举已损坏）
_FW = os.path.join(os.environ.get("AIOS_RUNTIME_LIBS", r"<HOME>\AppData\Roaming\aios-app\runtime-libs"), "fw_standalone")
if os.path.isdir(_FW):
    sys.path.insert(0, _FW)
sys.path[:] = [p for p in sys.path if p and "runtime-libs\\python-libs" not in p and "runtime-libs/python-libs" not in p]
os.environ["PYTHONPATH"] = ""
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE_DIR, "通话记录.xlsx")
OUT = os.path.join(BASE_DIR, "recordings_transcript.json")

# call_qc 项目里的本地模型路径（9/23 已验证 94/94 成功）
MODEL_DIR = r"<HOME>\AppData\Roaming\aios-app\workspaces\56b357e25bef434ba9d506065cc2fbaa\59e8f8cd43b44b62a0a000af8e9ab233\20260923_call_qc\scratch\whisper-small-local"
# 引入 call_qc 已转写的 transcript（94 条），避免重复转写
QC_TRANS_DIR = r"<HOME>\AppData\Roaming\aios-app\workspaces\56b357e25bef434ba9d506065cc2fbaa\59e8f8cd43b44b62a0a000af8e9ab233\20260923_call_qc\transcripts"

MIN_SEC = 10  # 与旧口径一致：≥10 秒才转写
NEW_AFTER = datetime.datetime(2026, 9, 15)  # 只处理 9/15 之后新增

def parse_sec(s):
    if not s:
        return None
    s = str(s)
    m = re.match(r'(\d+):(\d+):(\d+)', s)
    if m:
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))
    m = re.match(r'(\d+):(\d+)', s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    try:
        return int(float(s))
    except Exception:
        return None

def parse_time(v):
    if not v:
        return None
    v = str(v).strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M'):
        try:
            return datetime.datetime.strptime(v[:19] if len(v) >= 19 else v, fmt)
        except Exception:
            pass
    return None

def load_new_records():
    """读取 xlsx 中 9/15 之后、≥MIN_SEC、有 http 录音地址的行"""
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb["通话记录"]
    rows = list(ws.iter_rows(values_only=True))
    idx = {h: i for i, h in enumerate(rows[0])}
    recs = []
    for r in rows[1:]:
        url = r[idx["录音地址"]]
        if not url or not str(url).startswith("http"):
            continue
        sec = parse_sec(r[idx["通话时长"]])
        t = parse_time(r[idx["拨出时间/呼入时间"]])
        if not t or t < NEW_AFTER or not sec or sec < MIN_SEC:
            continue
        recs.append({
            "person": str(r[idx["拨打人"]]) if r[idx["拨打人"]] else "未知",
            "customer": str(r[idx["客户名称"]]) if r[idx["客户名称"]] else "",
            "url": str(url),
            "durationSec": sec,
            "time": str(r[idx["拨出时间/呼入时间"]]),
            "source": str(r[idx["通话来源"]]) if r[idx["通话来源"]] else "",
            "status": None,
            "text": None,
        })
    return recs

def load_qc_transcripts():
    """读取 call_qc 已转写的 transcripts，按 url 索引，转成与 OUT 同构的记录"""
    qc = {}
    if not os.path.isdir(QC_TRANS_DIR):
        return qc
    for f in os.listdir(QC_TRANS_DIR):
        if not f.endswith(".json"):
            continue
        try:
            d = json.load(open(os.path.join(QC_TRANS_DIR, f), encoding="utf-8"))
        except Exception:
            continue
        url = d.get("录音地址")
        if not url:
            continue
        segs = []
        for s in (d.get("分段") or []):
            t = (s.get("text") or "").strip()
            if t:
                segs.append({"start": s.get("start"), "end": s.get("end"), "text": t})
        # 文本按与腾讯云一致的时间戳行拼接，保证 score_calls.parse_segments 能还原逐句气泡
        full = "\n".join(seg_to_line(s) for s in segs)
        if not full:
            full = d.get("转写文本") or ""
        qc[url] = {
            "person": d.get("拨打人") or "",
            "customer": d.get("客户名称") or "",
            "url": url,
            "durationSec": parse_sec(d.get("通话时长")),
            "time": d.get("拨出时间") or "",
            "source": d.get("通话来源") or "",
            "status": "ok",
            "text": full,
            "segments": segs,
        }
    return qc

def fmt_time(sec):
    """秒 -> m:ss.xxx（与腾讯云 ASR 一致的段头格式，如 0:2.260）"""
    if sec is None:
        return "0:0.000"
    m = int(sec // 60)
    s = sec - m * 60
    return "%d:%.3f" % (m, s)

def _segget(seg, key, default=None):
    """兼容 dict（call_qc json）与 faster_whisper Segment 对象两种分段"""
    if isinstance(seg, dict):
        return seg.get(key, default)
    return getattr(seg, key, default)

def seg_to_line(seg):
    return "[%s,%s]%s" % (fmt_time(_segget(seg, "start")), fmt_time(_segget(seg, "end")), (_segget(seg, "text") or "").strip())

def wh_transcribe(model, audio_bytes_or_path):
    """用 faster_whisper 转写本地 wav 文件路径
    返回 (全文, 分段, 语言)；全文按 [m:ss.xxx,m:ss.xxx]文本 逐行拼接，
    与 score_calls.parse_segments 兼容，能还原逐句气泡。"""
    segments, info = model.transcribe(audio_bytes_or_path, language="zh", beam_size=5)
    seg_list = []
    text_lines = []
    for seg in segments:
        t = (seg.text or "").strip()
        if not t:
            continue
        seg_list.append({"start": round(seg.start, 2), "end": round(seg.end, 2), "text": t})
        text_lines.append(seg_to_line(seg))
    return "\n".join(text_lines), seg_list, info.language

def download(url, path, timeout=60):
    """下载录音到本地（断点续传：已存在且 >1KB 跳过）"""
    if os.path.exists(path) and os.path.getsize(path) > 1000:
        return True
    import urllib.request
    try:
        urllib.request.urlretrieve(url, path)
        return os.path.getsize(path) > 1000
    except Exception as e:
        print(f"   [下载失败] {url}: {e}", flush=True)
        return False

def main():
    from faster_whisper import WhisperModel  # 延迟导入，确保依赖就绪

    recs = load_new_records()
    print(f"9/15后新增录音(>= {MIN_SEC}s): {len(recs)} 条", flush=True)

    # 已转写（from 现有 OUT）
    done = {}
    if os.path.exists(OUT):
        try:
            for x in json.load(open(OUT, encoding="utf-8")):
                done[x["url"]] = x
        except Exception:
            done = {}
    print(f"recordings_transcript.json 已有 {len(done)} 条（含旧8月数据）", flush=True)

    # call_qc 已转写（94 条 9/20~9/23）
    qc = load_qc_transcripts()
    print(f"call_qc transcripts 可复用: {len(qc)} 条", flush=True)

    # 合并去重：优先 call_qc，其次 OUT
    pooled = dict(qc)
    pooled.update(done)  # done 在后，覆盖同 url（以现有 OUT 为准）
    skipped_qc = 0
    for rec in recs:
        if rec["url"] in pooled and pooled[rec["url"]].get("status") == "ok":
            skipped_qc += 1
    print(f"其中可跳过（已转写）: {skipped_qc} 条，真正需转写: {len(recs) - skipped_qc} 条", flush=True)

    # 组装输出：先保留 OUT 全部（旧数据不丢），再补 call_qc 中属于新范围、但 OUT 没有的
    out = []
    seen = set()
    for x in json.load(open(OUT, encoding="utf-8")) if os.path.exists(OUT) else []:
        out.append(x)
        seen.add(x.get("url"))
    for rec in recs:
        if rec["url"] in seen:
            continue
        if rec["url"] in qc:
            out.append(qc[rec["url"]])
            seen.add(rec["url"])
            print(f"  复用 call_qc: {rec['person']} | {rec['customer']}", flush=True)

    # 真正需要转写的
    todo = [r for r in recs if r["url"] not in seen and r["url"] not in pooled or (r["url"] in pooled and pooled[r["url"]].get("status") != "ok")]
    todo = [r for r in recs if r["url"] not in seen]
    # --limit N：本批最多转写 N 条（分批跑，避免超时被杀；断点续传按 url 跳过）
    if "--limit" in sys.argv:
        try:
            _n = int(sys.argv[sys.argv.index("--limit") + 1])
            todo = todo[:_n]
        except Exception:
            pass
    print(f"待转写: {len(todo)} 条", flush=True)
    if not todo:
        print("没有需要转写的新录音。DONE ->", OUT, flush=True)
        return

    print("加载 Whisper 模型 (本地 faster-whisper-small, CPU int8)...", flush=True)
    t0 = time.time()
    model = WhisperModel(MODEL_DIR, device="cpu", compute_type="int8")
    print(f"模型加载完成 {time.time()-t0:.1f}s", flush=True)

    import urllib.request, tempfile
    tmpdir = os.path.join(BASE_DIR, "scratch")
    os.makedirs(tmpdir, exist_ok=True)

    for i, rec in enumerate(todo, 1):
        name = re.sub(r'[\\/:*?"<>|]', "_", str(rec["customer"] or "NA"))[:30] or "NA"
        wav = os.path.join(tmpdir, f"new_{name}_{i}.wav")
        print(f"[{i}/{len(todo)}] 下载+转写: {rec['person']} | {rec['customer']} | {rec['durationSec']}s", flush=True)
        if not download(rec["url"], wav):
            rec["status"] = "fail"
            rec["text"] = None
            out.append(rec)
            seen.add(rec["url"])
            with open(OUT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
            print(f"   [下载失败] {rec['customer']}", flush=True)
            continue
        t_start = time.time()
        try:
            text, segs, lang = wh_transcribe(model, wav)
            rec["text"] = text
            rec["segments"] = segs
            rec["status"] = "ok"
            print(f"   完成 {time.time()-t_start:.1f}s, 文本{len(text)}字符", flush=True)
        except Exception as e:
            rec["status"] = "fail"
            rec["text"] = None
            print(f"   [转写失败] {e}", flush=True)
        out.append(rec)
        seen.add(rec["url"])
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"   [{i}/{len(todo)}] {rec['person']} | {rec['customer']} | {rec['status']} | {(rec['text'] or '')[:24]}", flush=True)

    print("DONE ->", OUT, flush=True)

if __name__ == "__main__":
    main()