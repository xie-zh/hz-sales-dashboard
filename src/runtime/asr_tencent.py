# -*- coding: utf-8 -*-
"""
腾讯云 ASR 录音文件识别 —— 纯标准库实现（零第三方依赖，绕过沙盒 pip 限制）
- 数据源：通话记录.xlsx（工作区）
- 筛选：通话时长 >= 30 秒 且 有录音地址
- 调用：腾讯云「录音文件识别」CreateRecTask + 轮询 DescribeTaskStatus
- 签名：自实现 TC3-HMAC-SHA256（纯 stdlib）
- 密钥：从环境变量 TENCENT_SECRET_ID / TENCENT_SECRET_KEY 读取，不落盘
- 输出：recordings_transcript.json
"""
import os, re, json, time, hashlib, hmac, datetime, urllib.request, urllib.error
import openpyxl

XLSX = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\通话记录.xlsx"
OUT = r"<PROJECT_ROOT>\数据文件\活动量\8月\8.11\recordings_transcript.json"
MIN_SEC = 10  # 扩展：≥10秒全部转写（含原≥30秒 53 条 + 10~30秒 151 条）
ENGINE = "8k_zh"  # 电话录音默认 8k；若识别异常改为 16k_zh

# ---------- TC3 签名（纯 stdlib） ----------
def _sha256(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8")).hexdigest()

def _hmac(key: bytes, msg: str) -> bytes:
    return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()

def _sign_tc3(secret_id, secret_key, service, action, payload, version="2019-06-14", region=""):
    host = f"{service}.tencentcloudapi.com"
    algorithm = "TC3-HMAC-SHA256"
    t = int(time.time())
    timestamp = str(t)
    date = datetime.datetime.utcfromtimestamp(t).strftime("%Y-%m-%d")
    payload_str = json.dumps(payload)
    canonical_headers = f"content-type:application/json; charset=utf-8\nhost:{host}\n"
    signed_headers = "content-type;host"
    canonical_request = "\n".join([
        "POST", "/", "", canonical_headers, signed_headers, _sha256(payload_str)
    ])
    credential_scope = f"{date}/{service}/tc3_request"
    string_to_sign = "\n".join([
        algorithm, timestamp, credential_scope, _sha256(canonical_request)
    ])
    secret_date = _hmac(("TC3" + secret_key).encode("utf-8"), date)
    secret_service = _hmac(secret_date, service)
    secret_signing = _hmac(secret_service, "tc3_request")
    signature = hmac.new(secret_signing, string_to_sign.encode("utf-8"), hashlib.sha256).hexdigest()
    authorization = (f"{algorithm} Credential={secret_id}/{credential_scope}, "
                    f"SignedHeaders={signed_headers}, Signature={signature}")
    headers = {
        "Authorization": authorization,
        "Content-Type": "application/json; charset=utf-8",
        "Host": host,
        "X-TC-Action": action,
        "X-TC-Timestamp": timestamp,
        "X-TC-Version": version,
    }
    if region:
        headers["X-TC-Region"] = region
    return headers, payload_str, host

def call_tc3(secret_id, secret_key, service, action, payload, version="2019-06-14", region=""):
    headers, payload_str, host = _sign_tc3(secret_id, secret_key, service, action, payload, version, region)
    url = f"https://{host}/"
    req = urllib.request.Request(url, data=payload_str.encode("utf-8"), headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8")
    return json.loads(body)

# ---------- 数据读取 ----------
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

def load_records():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb["通话记录"]
    rows = list(ws.iter_rows(values_only=True))
    idx = {h: i for i, h in enumerate(rows[0])}
    recs = []
    for r in rows[1:]:
        url = r[idx["录音地址"]]
        sec = parse_sec(r[idx["通话时长"]])
        if url and str(url).startswith("http") and sec and sec >= MIN_SEC:
            recs.append({
                "person": r[idx["拨打人"]],
                "customer": r[idx["客户名称"]],
                "url": str(url),
                "durationSec": sec,
                "time": str(r[idx["拨出时间/呼入时间"]]),
                "source": r[idx["通话来源"]],
            })
    return recs

# ---------- ASR ----------
def asr_transcribe(sid, skey, url):
    create = call_tc3(sid, skey, "asr", "CreateRecTask", {
        "EngineModelType": ENGINE, "ChannelNum": 1, "ResTextFormat": 0,
        "SourceType": 0, "Url": url,
    }, version="2019-06-14")
    if "Response" not in create or "Error" in create["Response"]:
        return None, f"CreateRecTask_ERR: {create.get('Response', {}).get('Error', create)}"
    task_id = create["Response"]["Data"]["TaskId"]
    for _ in range(80):  # 最多 4 分钟
        time.sleep(3)
        stat = call_tc3(sid, skey, "asr", "DescribeTaskStatus", {"TaskId": task_id}, version="2019-06-14")
        d = stat.get("Response", {}).get("Data", {})
        st = d.get("Status")
        if st == 2:
            return d.get("Result"), "ok"
        if st == 3:
            return None, f"Task_FAIL: {d.get('ErrorMsg', 'unknown')}"
    return None, "TIMEOUT"

def main():
    sid = os.environ.get("TENCENT_SECRET_ID")
    skey = os.environ.get("TENCENT_SECRET_KEY")
    if not sid or not skey:
        print("ERROR: 未设置 TENCENT_SECRET_ID / TENCENT_SECRET_KEY")
        raise SystemExit(1)
    recs = load_records()
    print(f"待转写(>= {MIN_SEC}s): {len(recs)} 条", flush=True)

    done = {}
    if os.path.exists(OUT):
        try:
            for x in json.load(open(OUT, encoding="utf-8")):
                done[x["url"]] = x
        except Exception:
            done = {}

    out = []
    for i, rec in enumerate(recs):
        if rec["url"] in done and done[rec["url"]].get("status") == "ok":
            out.append(done[rec["url"]])
            print(f"[{i+1}/{len(recs)}] 跳过已完成: {rec['person']} | {rec['customer']}", flush=True)
            continue
        text, status = asr_transcribe(sid, skey, rec["url"])
        rec["text"] = text
        rec["status"] = status
        out.append(rec)
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"[{i+1}/{len(recs)}] {rec['person']} | {rec['customer']} | {status} | {(text or '')[:24]}", flush=True)
        time.sleep(0.5)
    print("DONE ->", OUT, flush=True)

if __name__ == "__main__":
    main()
