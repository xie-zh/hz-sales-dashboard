# -*- coding: utf-8 -*-
"""活动量新数据源：WorkBuddy 资料库发布页「活动量工作台 · 七岗排行榜」

来源页：https://dashboard.example.com/p/LLT6AK0VUUvnPu12Cc8PKZ

取数分两步（指针会随每次「重新发布」变化，所以每次都要现取，不能硬编码版本号）：
  1) GET /p/<nodeId>              → 外壳页，含 window.__PUBLISH_BOOTSTRAP__ = {... "publishVersion": N ...}
  2) GET <static>/<nodeId>/<N>/七岗排行榜.html
        该 HTML 内嵌两块 JSON：
          var RB   = {...}  —— 排行榜基础：fields / posts / thrMain / cityDir / deptMgr / bench / cycles
          var DASH = {...}  —— 周期数据：cycles[] = [{key, span, range, validDays, data{岗位:表}, overview, dianshen}]

岗位表 data[岗位] 的结构：
    titles  : {字段key: 中文名}
    keys    : 字段顺序（与 cols 分组表头一一对应）
    children: {'分公司': {城市: [部门行...]}, '业态': {业态: [部门行...]}, '部门': {部门名: [人员行...]}}
    行      : {'lead': [显示名], 'cells': {字段key: 值}, 'n': 人数/天数}
    total   : 全局合计行的 cells

用法：
    import wb_source
    d = wb_source.fetch()                  # 有本地缓存，默认 30 分钟内不重复下载
    d = wb_source.fetch(force=True)        # 强制刷新
    d['cycle_key'], d['valid_days'], d['range']
    d['hz']            # {岗位: {'depts': [...], 'people': {部门: [...]}, 'titles': {}, 'keys': [...]}}

注意：整页 HTML 约 16 MB，解析要几秒；解析结果缓存在 _cache/dash.json。
"""
import json
import os
import re
import time
import urllib.parse
import urllib.request

PAGE_ID = 'LLT6AK0VUUvnPu12Cc8PKZ'
SHARE_URL = 'https://dashboard.example.com/p/%s' % PAGE_ID
STATIC_BASE = 'https://workbuddy-space-static.codebuddy.work/page/%s' % PAGE_ID
PAGE_FILE = '七岗排行榜.html'

# 只要这三岗（2026-10-09 与用户确认：只统计 网销 / 电邀 / 综法，其余岗位不再统计）
POSTS = ('网销岗', '电邀岗', '综法岗')

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(HERE, '_cache')
CACHE_JSON = os.path.join(CACHE_DIR, 'dash.json')
CACHE_HTML = os.path.join(CACHE_DIR, 'page.html')
DEFAULT_MAX_AGE = 30 * 60          # 秒


def _get(url, timeout=90):
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
        'Cache-Control': 'no-cache',
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def get_version():
    """返回 (nodeId, publishVersion)。"""
    html = _get(SHARE_URL).decode('utf-8', 'replace')
    m = re.search(r'window\.__PUBLISH_BOOTSTRAP__\s*=\s*(\{.*?\})\s*;', html, re.S)
    if not m:
        raise RuntimeError('未在分享页找到 __PUBLISH_BOOTSTRAP__，页面结构可能已变')
    cfg = json.loads(m.group(1))
    art = (cfg.get('artifact') or {})
    return cfg.get('nodeId') or PAGE_ID, art.get('publishVersion')


def _grab_var(text, name):
    """从 JS 源码里按大括号配平抓出 `var <name> = {...}` 的 JSON。"""
    i = text.find('var %s = ' % name)
    if i < 0:
        raise RuntimeError('未找到 `var %s`' % name)
    j = text.find('{', i)
    depth = 0
    k = j
    instr = False
    esc = False
    while k < len(text):
        c = text[k]
        if instr:
            if esc:
                esc = False
            elif c == '\\':
                esc = True
            elif c == '"':
                instr = False
        else:
            if c == '"':
                instr = True
            elif c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0:
                    break
        k += 1
    return json.loads(text[j:k + 1])


def _download_page(force=False, max_age=None):
    os.makedirs(CACHE_DIR, exist_ok=True)
    age = DEFAULT_MAX_AGE if max_age is None else max_age
    if not force and os.path.isfile(CACHE_HTML) and \
            time.time() - os.path.getmtime(CACHE_HTML) < age:
        return open(CACHE_HTML, encoding='utf-8', errors='replace').read()
    node_id, ver = get_version()
    url = '%s/%s/%s' % (STATIC_BASE, ver, urllib.parse.quote(PAGE_FILE))
    raw = _get(url, timeout=180).decode('utf-8', 'replace')
    with open(CACHE_HTML, 'w', encoding='utf-8') as f:
        f.write(raw)
    print('[wb_source] 已下载 %s (version=%s, %.1f MB)' % (PAGE_FILE, ver, len(raw) / 1048576))
    return raw


def _match_dept(lead, dept_keys):
    """把一条「分公司层」聚合行配到某个部门 key 上；配不上返回 None。

    规则（顺序很关键）：
      1) 精确：lead 里某个元素**就是**部门键（部门键为完整部门名，如「工商网销业务二部」）
      2) 包含：lead 元素长度 ≥ 4 且被部门键包含。
    ⚠️ 不能只用「包含」——各岗 lead 结构不同（电邀/网销/大促=[城市,业态,部门]、
       中台=[业态]、综法=[城市,部门,"-"]），城市名「」这种短名会命中**任意**
       部门，导致多行取到同一行。
    """
    cands = [x for x in lead if x and x != '-']
    if not cands:
        return None
    for c in sorted(cands, key=len, reverse=True):
        if c in dept_keys:
            return c
    for c in sorted(cands, key=len, reverse=True):
        if len(c) < 4:
            continue
        for k in dept_keys:
            if c in k:
                return k
    return None


def _normalize(RB, DASH):
    """整理成「」视角的易用结构。"""
    cycles = DASH.get('cycles') or []
    if not cycles:
        raise RuntimeError('DASH.cycles 为空')
    # 取最后一个周期（最新）
    cyc = cycles[-1]
    post_range = cyc.get('postRange') or {}
    ends = [v[1] for v in post_range.values() if isinstance(v, list) and len(v) > 1]
    out = {
        'cycle_key': cyc.get('key'),
        'range': cyc.get('range'),
        'span': cyc.get('span'),
        'valid_days': cyc.get('validDays'),
        'asof': max(ends) if ends else (cyc.get('rangeAll') or [None, None])[-1],
        'post_range': post_range,
        'posts_all': list((cyc.get('data') or {}).keys()),
        'rb': {
            'posts': RB.get('posts'),
            'thr_main': RB.get('thrMain'),
            'bench': RB.get('bench'),
        },
        'hz': {},
    }
    for post, d in (cyc.get('data') or {}).items():
        if post not in POSTS:
            continue
        ch = d.get('children') or {}
        if not isinstance(ch, dict):
            continue
        dept_people = ch.get('部门') or {}
        if not isinstance(dept_people, dict):
            continue
        hz_depts = [k for k in dept_people if '' in str(k)]
        if not hz_depts:
            continue
        # 部门聚合行：取「分公司」层里的部门行（权威聚合值）
        # 注意：不同岗位 lead 结构不同（电邀/网销/会计= [城市,业态,部门]，中台= [业态]，综法= [城市,部门,"-"]），
        # 所以不能一律取 lead[-1]，改用「lead 元素与部门名互相包含」来配对。
        cm = ch.get('分公司') or {}
        hzrows = cm.get('') or []
        agg = {}
        for r in hzrows:
            lead = [str(x) for x in (r.get('lead') or [])]
            name = _match_dept(lead, hz_depts)
            if name:
                agg[name] = {'lead': lead, 'cells': r.get('cells') or {}, 'n': r.get('n')}
        people = {}
        for dept in hz_depts:
            rows = dept_people.get(dept) or []
            people[dept] = [{'lead': (r.get('lead') or [None])[0],
                             'cells': r.get('cells') or {}, 'n': r.get('n')}
                            for r in rows]
        out['hz'][post] = {
            'titles': d.get('titles') or {},
            'keys': d.get('keys') or [],
            'depts': agg,
            'people': people,
        }
    return out


# =====================================================================================
#  「日」档取数：调用页面内嵌引擎 RBE 按选定日期区间重算
# =====================================================================================
# 源站数值是「按选定日期区间重算」的（页面顶部：日期区间 + 全周期/周期/日 三档）：
#   全周期 → 预构建的周期累计（本模块 fetch() 给的就是这一档）
#   日     → 单日，必须用页面自己的引擎 RBE 重算，否则数值对不上页面。
# RBE 是工作台 Python 侧（gen_rank_all / gen_rank5）的 JS 复刻，随页面一起发布，
# 所以这里把它原样抽出来、用 Node 跑（比开浏览器轻，且逐格与页面一致）。
#
# 【日】档锚点（与页面 v101/v162/v163d 口径一致）：
#   网销岗 / 电邀岗 → 全局日轴上的最新有效日（源表每日齐备）
#   综法岗          → 本岗「自身最新有效考核日」（该日须存在考核状态记录 r[1]≠0，
#                     仅为「源表缺行但有收款」补的合成记录不算）——综法数据通常落后几天
# =====================================================================================
CACHE_RB = os.path.join(CACHE_DIR, 'RB.json')
CACHE_DASH = os.path.join(CACHE_DIR, 'DASH.json')
CACHE_RBE = os.path.join(CACHE_DIR, 'rbe.js')
CACHE_DAY = os.path.join(CACHE_DIR, 'day.json')
ENGINE_JS = os.path.join(HERE, 'wb_engine.js')
RBE_START = 'var RBE = '
RBE_TAIL = "if (typeof module !== 'undefined' && module.exports) { module.exports = RBE; }"


def _node_exe():
    """定位受管 node（版本目录会随升级变化，禁止硬编码版本号）。"""
    import glob
    import shutil
    pat = os.path.join(os.path.expanduser('~'), '.workbuddy', 'binaries', 'node',
                       'versions', '*', 'node.exe')

    def key(p):
        name = os.path.basename(os.path.dirname(p))
        return [int(x) if x.isdigit() else 0 for x in re.split(r'[^0-9]+', name) if x != '']
    cands = sorted([p for p in glob.glob(pat) if os.path.isfile(p)], key=key, reverse=True)
    if cands:
        return cands[0]
    return shutil.which('node')


def _grab_rbe(html):
    i = html.find(RBE_START)
    if i < 0:
        raise RuntimeError('未在页面里找到 `var RBE`')
    e = html.find(RBE_TAIL, i)
    if e < 0:
        raise RuntimeError('未找到 RBE 导出尾（页面结构可能已变）')
    return html[i:e + len(RBE_TAIL)]


def day_anchors(RB, cycle_key):
    """返回 (anchors, latest_global)：每岗【日】档锚点日期。"""
    c = (RB.get('cycles') or {}).get(cycle_key)
    if not c:
        raise RuntimeError('RB.cycles 里没有周期 %s' % cycle_key)
    axis = c.get('days') or []
    valid = [d['d'] for d in axis if d.get('x') != 1]
    latest = max(valid) if valid else ''
    anchors = {}
    for post in POSTS:
        if post == '综法岗':
            node = ((c.get('posts2') or {}).get(post) or {})
            best = ''
            for p in (node.get('ppl') or []):
                recs = p[5] if isinstance(p, list) and len(p) > 5 else []
                for rec in (recs or []):
                    di = rec[0] if rec else None
                    if di is None or di >= len(axis):
                        continue
                    md = axis[di]
                    if md.get('x') == 1 or not rec[1]:
                        continue          # 剔除日 / 无考核状态的合成记录
                    if md['d'] > best:
                        best = md['d']
            anchors[post] = best or latest
        else:
            anchors[post] = latest
    return anchors, latest


def fetch_day(cycle=None, force=False, max_age=600, day=None):
    """按「日」档取数（每岗各自锚点，可用 day 强制指定 YYYYMMDD）。

    返回结构与 fetch() 兼容：d['hz'][岗位] = {titles, keys, depts, people}
    另加：d['mode']='day'，d['anchor'][岗位]，d['data_date']。
    """
    os.makedirs(CACHE_DIR, exist_ok=True)
    html = _download_page(force=force, max_age=max_age)
    RB = _grab_var(html, 'RB')
    DASH = _grab_var(html, 'DASH')
    with open(CACHE_RB, 'w', encoding='utf-8') as f:
        json.dump(RB, f, ensure_ascii=False)
    with open(CACHE_DASH, 'w', encoding='utf-8') as f:
        json.dump(DASH, f, ensure_ascii=False)
    with open(CACHE_RBE, 'w', encoding='utf-8') as f:
        f.write(_grab_rbe(html) + '\n')

    ck = cycle or (DASH['cycles'][-1]['key'])
    anchors, latest = day_anchors(RB, ck)
    if day:
        anchors = {p: day for p in POSTS}

    node = _node_exe()
    if not node:
        raise RuntimeError('未找到 node 运行时可执行文件')
    spec = {'cycle': ck,
            'items': [{'post': p, 'lo': anchors[p], 'hi': anchors[p]} for p in POSTS]}
    import subprocess
    p = subprocess.run([node, ENGINE_JS, json.dumps(spec, ensure_ascii=False)],
                       cwd=HERE, capture_output=True, timeout=300)
    raw = p.stdout.decode('utf-8', errors='replace')
    if not raw.strip():
        raise RuntimeError('引擎无输出: %s' % p.stderr.decode('utf-8', errors='replace')[:400])
    res = json.loads(raw[raw.find('{'):raw.rfind('}') + 1])

    hz = {}
    info = {}
    for it in res.get('items', []):
        post = it['post']
        info[post] = {'ok': it['ok'], 'vd': it.get('vd'), 'err': it.get('err', '')}
        if not it.get('ok'):
            continue
        dev = it.get('depts') or {}
        cm = (it.get('cm') or {}).get('') or []
        keys = list(dev.keys())
        depts = {}
        for key in keys:
            agg = None
            for r in cm:
                lead = [str(x) for x in (r.get('lead') or [])]
                if _match_dept(lead, [key]) == key:
                    agg = r
                    break
            depts[key] = {'lead': (agg or {}).get('lead') or [key],
                          'cells': (agg or {}).get('cells') or {},
                          'n': (agg or {}).get('n')}
        people = {}
        for key in keys:
            people[key] = [{'lead': (r.get('lead') or [None])[0],
                            'cells': r.get('cells') or {}, 'n': r.get('n')}
                           for r in (dev.get(key) or [])]
        hz[post] = {'titles': it.get('titles') or {}, 'keys': it.get('keys') or [],
                    'depts': depts, 'people': people}

    out = {
        'mode': 'day',
        'cycle_key': ck,
        'anchor': anchors,
        'latest_day': latest,
        'data_date': anchors.get('网销岗') or latest,
        'range': [latest, latest],
        'valid_days': 1,
        'asof': latest,
        'posts_all': list((DASH['cycles'][-1].get('data') or {}).keys()),
        'rb': {'posts': RB.get('posts'), 'thr_main': RB.get('thrMain'), 'bench': RB.get('bench')},
        'engine': info,
        'hz': hz,
    }
    json.dump(out, open(CACHE_DAY, 'w', encoding='utf-8'), ensure_ascii=False)
    return out


def fetch(force=False, max_age=DEFAULT_MAX_AGE):
    """返回整理后的结构 dict；带本地缓存。"""
    global DEFAULT_MAX_AGE
    os.makedirs(CACHE_DIR, exist_ok=True)
    if not force and os.path.isfile(CACHE_JSON) and \
            time.time() - os.path.getmtime(CACHE_JSON) < max_age:
        return json.load(open(CACHE_JSON, encoding='utf-8'))
    html = _download_page(force=force)
    RB = _grab_var(html, 'RB')
    DASH = _grab_var(html, 'DASH')
    data = _normalize(RB, DASH)
    json.dump(data, open(CACHE_JSON, 'w', encoding='utf-8'), ensure_ascii=False)
    return data


if __name__ == '__main__':
    import io
    import sys
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    d = fetch(force=True)
    print('周期:', d['cycle_key'], d['range'], '有效天数', d['valid_days'])
    print('全部岗位:', d['posts_all'])
    print('有数据的岗位:')
    for post, v in d['hz'].items():
        nppl = sum(len(x) for x in v['people'].values())
        print('  %-8s 部门 %d 个 / 人数 %d  部门=%s' % (post, len(v['depts']), nppl, list(v['depts'].keys())))
    print()
    print('--- 网销岗 ·  部门聚合行（字段: 值）---')
    for dept, r in (d['hz'].get('网销岗', {}).get('depts') or {}).items():
        t = (d['hz']['网销岗']['titles'])
        print('  %s (n=%s)' % (dept, r['n']))
        for k in ['onJob', 'kaohe', 'avgT1730', 'r1730', 'avgT2400', 'r2400', 'dayPerf', 'perfAvg']:
            print('      %-10s %-20s %s' % (k, t.get(k, ''), r['cells'].get(k)))
