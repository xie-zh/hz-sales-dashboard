# -*- coding: utf-8 -*-
# @version V0.3  @author <AUTHOR>  @created 2026-09-24
"""看板取数引擎 —— BI 下载（帆软直连）+ Excel 写入（极简 COM）

与 EXAMPLERe5.0 的关系：完全独立，不 import 它任何模块，也不启动它的常驻程序。
下载走帆软官方导出接口（URL 追加 &format=excel&extype=simple），不打开浏览器。
写 Excel 走一次后台 COM 会话（不可见），整块数组一次性写入，跑完立即退出进程。

依赖：openpyxl（读下载文件）+ pywin32（写目标 Excel）+ 标准库（下载）
"""

import datetime
import json
import os
import re
import time
import urllib.parse
import urllib.request
import urllib.error
from concurrent.futures import ThreadPoolExecutor

import openpyxl

# --------------------------------------------------------------------------
# 日志：主脚本可替换成自己的 log（带时间戳、可写文件）
# --------------------------------------------------------------------------
log = print

_UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36')
_TIMEOUT = 90
_XLSX_MAGIC = b'PK\x03\x04'

_SKIP_WIDGET_TYPES = {'label', 'button', 'submit', 'image', 'link', 'para'}
_SKIP_NAME_PREFIX = ('LABEL', 'FORMSUBMIT', 'FR_', 'PARA', 'BUTTON')

_WIDGET_CACHE = {}        # url -> (ts, defaults)
_WIDGET_TTL = 1800        # 30 分钟：日期类默认值每天会变
_DICT_CACHE = {}          # widgetUrl -> (ts, {显示名: 代码值})


# --------------------------------------------------------------------------
# 占位符：{today} {yesterday} {today-7} {yyyy} {mm} {dd} {month_start}
# --------------------------------------------------------------------------
def resolve_alias(text):
    if not isinstance(text, str) or '{' not in text:
        return text

    def _one(expr):
        expr = expr.strip()
        m = re.match(r'^(today|yesterday|tomorrow)(?:([+-])(\d+))?$', expr, re.I)
        base = datetime.date.today()
        if m:
            kind, sign, num = m.group(1).lower(), m.group(2), int(m.group(3) or 0)
            d = {'today': base, 'yesterday': base - datetime.timedelta(days=1),
                 'tomorrow': base + datetime.timedelta(days=1)}[kind]
            if sign == '-':
                d -= datetime.timedelta(days=num)
            elif sign == '+':
                d += datetime.timedelta(days=num)
            return d.strftime('%Y-%m-%d')
        if expr.lower() == 'month_start':
            return base.replace(day=1).strftime('%Y-%m-%d')
        if expr.lower() == 'month_end':
            nxt = (base.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
            return (nxt - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        if expr.lower() == 'yyyy':
            return '%04d' % base.year
        if expr.lower() == 'mm':
            return '%02d' % base.month
        if expr.lower() == 'dd':
            return '%02d' % base.day
        return None

    def _rep(m):
        v = _one(m.group(1))
        return v if v is not None else m.group(0)

    return re.sub(r'\{([^}]+)\}', _rep, text)


# --------------------------------------------------------------------------
# 帆软：抓取参数控件默认值 + 远程字典
# --------------------------------------------------------------------------
def _abs_url(url, referer):
    if url.startswith('http'):
        return url
    if url.startswith('/'):
        m = re.match(r'(https?://[^/]+)', referer or '')
        if m:
            return m.group(1) + url
    return url


def _fetch_remote_dict(widget_url, referer):
    """remote 控件（如分公司下拉）提交的是代码值，不是显示名。"""
    url = _abs_url(widget_url, referer)
    cached = _DICT_CACHE.get(url)
    if cached and time.time() - cached[0] < _WIDGET_TTL:
        return cached[1]
    try:
        req = urllib.request.Request(url, headers={'User-Agent': _UA, 'Referer': referer})
        text = urllib.request.urlopen(req, timeout=15).read().decode('utf-8', 'replace')
        data = json.loads(text)
        dmap = {}
        if isinstance(data, list):
            for it in data:
                if isinstance(it, dict) and 'value' in it and 'text' in it:
                    dmap[str(it.get('text'))] = str(it.get('value'))
        _DICT_CACHE[url] = (time.time(), dmap)
        return dmap
    except Exception:
        return {}


def fetch_widget_defaults(url):
    """从报表页面 HTML 里解析参数控件的默认值，用来补全我们没配置的控件。

    不这么做的话，后端会用它自己的默认值（可能是别的分公司/别的日期），导出来就是空的。
    """
    cached = _WIDGET_CACHE.get(url)
    if cached and time.time() - cached[0] < _WIDGET_TTL:
        return cached[1]
    try:
        req = urllib.request.Request(url, headers={'User-Agent': _UA})
        text = urllib.request.urlopen(req, timeout=20).read().decode('utf-8', 'replace')
    except Exception as e:
        log('  ⚠ 读取报表参数失败（%s），只用配置里的筛选条件' % str(e)[:80])
        return {}

    out = {}
    pos = 0
    while True:
        i = text.find('"widgetName"', pos)
        if i < 0:
            break
        pos = i + len('"widgetName"')
        start = text.rfind('{', 0, i)
        depth = 0
        end = None
        for k in range(start, len(text)):
            if text[k] == '{':
                depth += 1
            elif text[k] == '}':
                depth -= 1
                if depth == 0:
                    end = k
                    break
        if end is None or end < i:
            continue
        try:
            w = json.loads(text[start:end + 1])
        except Exception:
            w = None
        if not w:
            continue
        if w.get('type') == 'parameter' or 'children' in w:
            continue
        name = w.get('widgetName')
        if not isinstance(name, str):
            continue
        wtype = w.get('type')
        if (wtype in _SKIP_WIDGET_TYPES or name.startswith(_SKIP_NAME_PREFIX)
                or w.get('needSubmit') is False):
            continue
        val = w.get('value')
        if isinstance(val, dict) and 'date_milliseconds' in val:
            try:
                val = datetime.datetime.fromtimestamp(
                    val['date_milliseconds'] / 1000).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                val = None
        if not isinstance(val, (str, int, float)):
            continue
        dmap = {}
        if w.get('mode') == 'remote' and w.get('widgetUrl'):
            dmap = _fetch_remote_dict(w.get('widgetUrl'), url)
        sval = str(val)
        fmt = w.get('format')
        # 日期控件内部值带时间，后端只认显示格式（纯日期就去掉时间部分）
        if ' ' in sval and not (fmt and any(t in fmt for t in ('HH', 'mm', 'ss'))):
            sval = sval.split(' ')[0]
        out[name] = {'value': sval, 'dict_map': dmap}
        pos = max(pos, end + 1)
    _WIDGET_CACHE[url] = (time.time(), out)
    return out


def build_export_url(src, bi_login_no, defaults=None):
    base = src.get('url', '')
    sep = '&' if '?' in base else '?'
    params = [('format', 'excel'), ('extype', 'simple')]
    added = {'format', 'extype'}
    shown = []
    for key, val in (src.get('filter') or {}).items():
        val = resolve_alias(val)
        dmap = (defaults or {}).get(key, {}).get('dict_map') or {}
        real = dmap.get(str(val), val)
        params.append((key, real))
        added.add(key)
        shown.append('%s=%s%s' % (key, val, '' if str(real) == str(val) else '(→%s)' % real))
    if defaults:
        for key, info in defaults.items():
            if key in added:
                continue
            val = info.get('value', '')
            dmap = info.get('dict_map') or {}
            real = dmap.get(str(val), val)
            params.append((key, real))
            added.add(key)
    if bi_login_no:
        params.append(('biLoginNo', bi_login_no))
    if shown:
        log('    筛选: ' + '  '.join(shown))
    return base + sep + urllib.parse.urlencode(params)


# --------------------------------------------------------------------------
# 下载
# --------------------------------------------------------------------------
def download_one(src, out_dir, bi_login_no, retries=2):
    """下载单个报表，返回文件路径；失败返回 None。"""
    sid = src.get('id', '?')
    url = src.get('url', '')
    defaults = fetch_widget_defaults(url) if url else {}
    full = build_export_url(src, bi_login_no, defaults)
    path = os.path.join(out_dir, '%s.xlsx' % sid)
    tmp = path + '.tmp'

    last_err = ''
    for attempt in range(retries + 1):
        if attempt:
            log('  ↻ %s 第 %d 次重试' % (sid, attempt + 1))
            time.sleep(3 * attempt)
        try:
            req = urllib.request.Request(full, headers={
                'User-Agent': _UA,
                'Accept': 'text/html,application/xhtml+xml,*/*;q=0.8',
                'Referer': url,
            })
            with urllib.request.urlopen(req, timeout=_TIMEOUT) as resp:
                if resp.status != 200:
                    last_err = 'HTTP %s' % resp.status
                    continue
                data = resp.read()
        except urllib.error.HTTPError as e:
            last_err = 'HTTP %s %s' % (e.code, e.reason)
            break                      # 4xx/5xx 重试没意义
        except Exception as e:
            last_err = str(e)[:120]
            continue
        if data[:4] != _XLSX_MAGIC:
            last_err = '响应不是 Excel（疑似登录页）'
            continue
        try:
            with open(tmp, 'wb') as f:
                f.write(data)
            os.replace(tmp, path)      # 原子替换：不会留下半截文件
            return path
        except Exception as e:
            last_err = '写盘失败 %s' % str(e)[:80]
            try:
                os.remove(tmp)
            except Exception:
                pass
    log('  ✗ %s 下载失败：%s' % (sid, last_err))
    return None


def download_all(sources, out_dir, bi_login_no, workers=5, retries=2):
    """并发下载，返回 {id: 路径或 None}。"""
    os.makedirs(out_dir, exist_ok=True)
    srcs = [s for s in sources if s.get('enabled', True) and s.get('url')]
    log('  并发下载 %d 张报表（%d 线程）' % (len(srcs), workers))

    def _task(s):
        t0 = time.time()
        p = download_one(s, out_dir, bi_login_no, retries=retries)
        return s.get('id'), p, time.time() - t0

    result = {}
    ok = fail = 0
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for sid, p, sec in ex.map(_task, srcs):
            result[sid] = p
            if p:
                ok += 1
                log('  ✓ %-20s %6.1fs  %6.0f KB' % (sid, sec, os.path.getsize(p) / 1024))
            else:
                fail += 1
    log('  下载完成：成功 %d / 失败 %d（%.1fs）' % (ok, fail, time.time() - t0))
    return result


# --------------------------------------------------------------------------
# 读取下载文件
# --------------------------------------------------------------------------
def read_sheet_values(path, sheet=None):
    """读出二维值（第一行通常是表头）。sheet=None 取第一个工作表。"""
    # 不能用 read_only=True：帆软导出的文件 dimension 信息不准，
    # read_only 模式下 iter_rows() 只会返回 1 行。文件都很小（<200KB），全量读没问题。
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb[sheet] if (sheet and sheet in wb.sheetnames) else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    # 去掉尾部全空行，避免把大量空行写进目标表
    while rows and all(v is None or v == '' for v in rows[-1]):
        rows.pop()
    return rows


def file_md5(path, block=1 << 20):
    import hashlib
    h = hashlib.md5()
    with open(path, 'rb') as f:
        while True:
            b = f.read(block)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


# --------------------------------------------------------------------------
# 机器级 COM 注册自愈
#   高完整性（提权）进程按 Windows 安全模型**不使用** HKCU\Software\Classes 下的
#   用户级 COM 注册，而 WPS 只把 Excel 的 COM 类注册在 HKCU —— 于是提权进程激活
#   时报 CO_E_CLASSSTRING / REGDB_E_CLASSNOTREG（计划任务正是这种情况）。
#   这里把 HKCU 中的注册补写到机器级 HKLM（64 位 + 32 位两个视图），使提权进程
#   同样能激活 WPS 的 Excel COM。幂等：已存在直接返回；无权限写入时安静跳过。
# --------------------------------------------------------------------------
COM_CLASSES = (
    ('KET.Application',   '{45540001-5750-5300-4B49-4E47534F4655}'),   # WPS 表格
    ('Excel.Application', '{00024500-0000-0000-C000-000000000046}'),   # 兼容 Excel 的 ProgID
)


def machine_com_ready():
    """HKLM 中是否已具备**可用**的 WPS Excel COM 注册（64 位视图）。

    除了键存在，还校验 LocalServer32 指向的 exe 是否仍在 —— WPS 升级会改版本目录，
    旧注册会指向不存在的路径，此时视为"未就绪"，让自愈逻辑用 HKCU 的新值覆盖。
    """
    import winreg
    try:
        for _prog, clsid in COM_CLASSES:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                r'SOFTWARE\Classes\CLSID\%s\LocalServer32' % clsid,
                                0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as k:
                srv = (winreg.QueryValueEx(k, '')[0] or '').strip()
            if not srv:
                return False
            exe = srv.split('"')[1] if srv.startswith('"') else srv.split(' /')[0].strip()
            if not os.path.exists(exe):
                return False
        return True
    except Exception:
        return False


def ensure_machine_com_registration():
    """把 HKCU 中 WPS 的 Excel COM 注册补写到 HKLM。返回可读状态字符串。"""
    import winreg
    if machine_com_ready():
        return 'ok(已存在)'
    src = {}
    for prog, clsid in COM_CLASSES:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r'Software\Classes\CLSID\%s\LocalServer32' % clsid) as k:
                src[clsid] = winreg.QueryValueEx(k, '')[0]
        except Exception as e:
            return 'skip(读 HKCU 失败 %r)' % e
    # 两个视图：\SOFTWARE\Classes = 64 位；\SOFTWARE\WOW6432Node\Classes = 32 位
    for base, wow in ((r'SOFTWARE\Classes', winreg.KEY_WOW64_64KEY),
                      (r'SOFTWARE\WOW6432Node\Classes', 0)):
        for prog, clsid in COM_CLASSES:
            for sub, val in ((prog, prog),
                             (prog + r'\CLSID', clsid),
                             (r'CLSID\%s' % clsid, prog),
                             (r'CLSID\%s\LocalServer32' % clsid, src[clsid])):
                k = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, base + '\\' + sub, 0,
                                       winreg.KEY_ALL_ACCESS | wow)
                try:
                    winreg.SetValueEx(k, '', 0, winreg.REG_SZ, val)
                finally:
                    winreg.CloseKey(k)
    return 'ok(已写入 HKLM)'


# --------------------------------------------------------------------------
# Excel 写入：一次后台 COM 会话搞定所有文件
# --------------------------------------------------------------------------
class ExcelSession:
    """不可见的 Excel/WPS 会话。

    用法：
        with ExcelSession() as xl:
            xl.write_file(path, [(sheet名, 二维值), ...], recalc=True)
    """

    XL_AUTOMATIC = -4105
    XL_MANUAL = -4135

    def __init__(self, engine='auto'):
        self.engine = engine
        self.app = None
        self.prog = None
        self._coinit = False

    def __enter__(self):
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        self._coinit = True
        # 提权进程看不到 HKCU 的用户级 COM 注册 → 先把注册补到机器级 HKLM（幂等）
        try:
            _r = ensure_machine_com_registration()
            if '已存在' not in _r:
                log('    COM 机器级注册：%s' % _r)
        except Exception as _e:
            log('    COM 机器级注册跳过（%s）' % _e)
        order = (['KET.Application', 'Excel.Application'] if self.engine == 'auto'
                 else ([self.engine] if '.' in self.engine else
                       ['KET.Application' if self.engine == 'wps' else 'Excel.Application']))
        last = None
        for prog in order:
            # 主路径：DispatchEx（历史成功路径，行为不变）
            try:
                self.app = win32.DispatchEx(prog)
                self.prog = prog
                break
            except Exception as e:
                last = e
                log('    COM 启动失败 [%s via DispatchEx]：%r' % (prog, e))
            # 备用路径：显式 CoCreateInstance，换不同 clsctx（高完整性令牌下行为可能不同）
            iid = self._progid_to_clsid(prog)
            if iid is None:
                continue
            for label, ctx in (('LOCAL_SERVER', pythoncom.CLSCTX_LOCAL_SERVER),
                               ('LOCAL|32BIT', pythoncom.CLSCTX_LOCAL_SERVER | 0x40000),
                               ('ALL', pythoncom.CLSCTX_ALL)):
                try:
                    obj = pythoncom.CoCreateInstance(iid, None, ctx, pythoncom.IID_IDispatch)
                    self.app = win32.Dispatch(obj)
                    self.prog = prog
                    log('    COM 备用激活成功：%s via %s' % (prog, label))
                    break
                except Exception as e:
                    last = e
                    log('    COM 备用失败 [%s via %s]：%r' % (prog, label, e))
            if self.app is not None:
                break
        if self.app is None:
            self._log_com_diag(order)
            raise RuntimeError('无法启动 Excel/WPS COM：%s' % last)
        a = self.app
        a.Visible = False
        try:
            a.DisplayAlerts = False
            a.ScreenUpdating = False
            a.EnableEvents = False
            a.AskToUpdateLinks = False
            a.AlertBeforeOverwriting = False
        except Exception:
            pass
        log('  COM 引擎：%s' % ('WPS' if 'KET' in self.prog else 'Excel'))
        return self

    @staticmethod
    def _progid_to_clsid(prog):
        """ProgID → PyIID。pywin32 未提供 CLSIDFromProgID，改从注册表读。"""
        try:
            import winreg
            import pywintypes
            s = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, prog + '\\CLSID')
            return pywintypes.IID(s)
        except Exception as e:
            log('    解析 %s 的 CLSID 失败：%r' % (prog, e))
            return None

    def _log_com_diag(self, order):
        """COM 启动失败时的环境快照，用于定位是权限 / 注册表 / apartment 哪一类问题。"""
        log('  --- COM 诊断 ---')
        try:
            import sys as _sys
            import ctypes
            log('    python=%d 位  pid=%d  用户=%s' % (
                64 if _sys.maxsize > 2 ** 32 else 32, os.getpid(),
                os.environ.get('USERNAME', '?')))
            for k in ('USERPROFILE', 'APPDATA', 'LOCALAPPDATA', 'TEMP', 'SESSIONNAME', 'PATH'):
                v = os.environ.get(k)
                if k == 'PATH' and v:
                    v = v[:120] + '...'
                log('    %s=%s' % (k, v if v else '(未设置)'))
            try:
                _k = ctypes.windll.kernel32
                _u = ctypes.windll.user32
                buf = ctypes.create_unicode_buffer(256)
                need = ctypes.c_ulong()
                if _u.GetUserObjectInformationW(_u.GetProcessWindowStation(), 2, buf, 512, ctypes.byref(need)):
                    log('    窗口站=%s' % buf.value)
                if _u.GetUserObjectInformationW(_u.GetThreadDesktop(_k.GetCurrentThreadId()), 2, buf, 512, ctypes.byref(need)):
                    log('    桌面=%s' % buf.value)
            except Exception as e:
                log('    窗口站探测失败：%s' % e)
            try:
                import winreg
                for prog in order:
                    try:
                        kk = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog + '\\CLSID')
                        clsid, _ = winreg.QueryValueEx(kk, '')
                        winreg.CloseKey(kk)
                        srv = '(无 LocalServer32)'
                        try:
                            sk = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'CLSID\\' + clsid + '\\LocalServer32')
                            srv, _ = winreg.QueryValueEx(sk, '')
                            winreg.CloseKey(sk)
                        except Exception:
                            pass
                        log('    HKCR\\%s → %s  %s' % (prog, clsid, srv))
                    except Exception as e:
                        log('    HKCR\\%s 读取失败：%s' % (prog, e))
            except Exception as e:
                log('    注册表探测失败：%s' % e)
            # 令牌画像（对比交互会话进程，看组/完整性级别差异）
            try:
                import win32api
                import win32security
                tok = win32security.OpenProcessToken(win32api.GetCurrentProcess(),
                                                     win32security.TOKEN_QUERY)
                names = []
                for sid, _attr in win32security.GetTokenInformation(tok, win32security.TokenGroups):
                    try:
                        names.append(win32security.LookupAccountSid(None, sid)[0])
                    except Exception:
                        names.append(str(sid))
                log('    令牌组(%d): %s' % (len(names), ','.join(names)))
                try:
                    log('    完整性级别: %s' % (win32security.GetTokenInformation(
                        tok, win32security.TokenIntegrityLevel)[0],))
                except Exception:
                    pass
                try:
                    log('    TokenElevation=%s' % win32security.GetTokenInformation(
                        tok, win32security.TokenElevation))
                except Exception:
                    pass
            except Exception as e:
                log('    令牌探测失败：%s' % e)
            # 子进程对照：同一台机器上另起干净进程，能否启动 COM
            try:
                import subprocess
                here = os.path.dirname(os.path.abspath(__file__))
                probe = os.path.join(here, '_probe_reg.py')
                if os.path.exists(probe):
                    r = subprocess.run([_sys.executable, probe], cwd=here,
                                       capture_output=True, text=True, timeout=150)
                    hit = False
                    for ln in (r.stdout or '').splitlines():
                        s = ln.strip()
                        if s.startswith('OK') or s.startswith('FAIL'):
                            log('    子进程对照: %s' % s)
                            hit = True
                    if not hit:
                        log('    子进程对照: 无结果 rc=%s %s' % (r.returncode, (r.stderr or '')[:200]))
            except Exception as e:
                log('    子进程对照失败：%s' % e)
        except Exception as e:
            log('    诊断输出失败：%s' % e)

    def __exit__(self, *exc):
        try:
            if self.app is not None:
                self.app.Quit()
                del self.app
        except Exception:
            pass
        if self._coinit:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return False

    # ---------- 值转换 ----------
    @staticmethod
    def _com_value(v):
        if v is None:
            return ''
        if isinstance(v, datetime.datetime):
            import pywintypes
            try:
                return pywintypes.Time(v)
            except Exception:
                return v.strftime('%Y-%m-%d %H:%M:%S')
        if isinstance(v, datetime.date):
            import pywintypes
            try:
                return pywintypes.Time(datetime.datetime(v.year, v.month, v.day))
            except Exception:
                return v.strftime('%Y-%m-%d')
        if isinstance(v, datetime.time):
            return v.strftime('%H:%M:%S')
        return v

    def _get_sheet(self, wb, name):
        for s in wb.Worksheets:
            if s.Name == name:
                return s
        ws = wb.Worksheets.Add()
        ws.Name = name
        return ws

    def write_file(self, path, ops, recalc=False):
        """ops = [(sheet名, 二维值), ...]；recalc=True 时保存前强制重算。"""
        if not ops:
            return 0
        t0 = time.time()
        # 重算开关：只有依赖公式的文件才需要（看板读的是公式算出来的值）
        try:
            self.app.Calculation = self.XL_AUTOMATIC if recalc else self.XL_MANUAL
            self.app.CalculateBeforeSave = bool(recalc)
        except Exception:
            pass
        wb = self.app.Workbooks.Open(path)
        try:
            for name, values in ops:
                ws = self._get_sheet(wb, name)
                try:
                    ws.UsedRange.ClearContents()
                except Exception:
                    ws.Cells.ClearContents()
                if not values:
                    continue
                ncols = max(len(r) for r in values)
                # 分块写入：一次整块赋值最稳，块太大才拆
                CHUNK = 3000
                for start in range(0, len(values), CHUNK):
                    part = values[start:start + CHUNK]
                    block = tuple(
                        tuple(self._com_value(v) for v in (row + [''] * ncols)[:ncols])
                        for row in part
                    )
                    r1 = start + 1
                    r2 = start + len(part)
                    ws.Range(ws.Cells(r1, 1), ws.Cells(r2, ncols)).Value = block
            if recalc:
                try:
                    self.app.CalculateFullRebuild()
                except Exception:
                    pass
            wb.Save()
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        log('  ✓ %s 写入 %d 张表%s（%.1fs）'
            % (os.path.basename(path), len(ops), ' + 重算' if recalc else '',
               time.time() - t0))
        return len(ops)


def is_locked(path):
    """文件是否被占用（用户正打开着）。"""
    try:
        f = open(path, 'r+b')
        f.close()
        return False
    except Exception:
        return True
