# -*- coding: utf-8 -*-
"""
四模块历史快照（员工三力 / 网销 / 电邀 / 会计）
=====================================================================
把当前 _deploy 下 4 个实时数据 JSON 按「日期」存档，供前端历史时间选择器加载。

数据流：
  构建 build_snapshot.py 每跑一次都会调用 snapshot_now()，把当天最新的 4 个 JSON
  复制成 hist/<module>/<YYYY-MM-DD>.json，并刷新 hist/manifest.json（模块→可用日期列表）。

  模块 -> 实时数据文件映射（与 build_snapshot.py 写出的 4 个文件一一对应）：
    sanli    -> sanli_data.json   （员工三力）
    ry       -> ry_data.json      （网销数据分析）
    ry_dy    -> ry_dy_data.json   （电邀数据分析）
    ry_kuaiji-> ry_kuaiji_data.json（会计数据分析）

用法：
  python hist_snapshot.py snapshot_now            # 把当前 _deploy JSON 存为「今天」快照
  python hist_snapshot.py snapshot_now 2026-08-20 # 指定日期（供 backfill 调用）
  python hist_snapshot.py backfill <src_root>     # 用历史源 Excel 补录过去日期（后期使用）
"""
import os
import sys
import json
import shutil
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(BASE_DIR, "_deploy")
HIST = os.path.join(DEPLOY, "hist")
MANIFEST = os.path.join(HIST, "manifest.json")

# 模块 -> 对应的实时数据 JSON 文件名（务必与 build_snapshot.py 写出的一致）
STREAMS = {
    "sanli": "sanli_data.json",
    "ry": "ry_data.json",
    "ry_dy": "ry_dy_data.json",
    "ry_kuaiji": "ry_kuaiji_data.json",
}

# 保留天数；默认 9999 = 全部保留。若部署体积过大可调小（如 180）。
RETENTION_DAYS = 9999

# 抑制开关：backfill 临时替换源表后重建，此时绝不能让 build() 里的自动快照把「当天」
# 覆盖成历史日期的数据。backfill 在调用 build() 前把它置 True，build 内自动快照即跳过。
SUPPRESS_AUTO_SNAPSHOT = False


def _today():
    return datetime.date.today().strftime("%Y-%m-%d")


def _load_manifest():
    if os.path.exists(MANIFEST):
        try:
            m = json.load(open(MANIFEST, encoding="utf-8"))
            m.setdefault("modules", {k: [] for k in STREAMS})
            m.setdefault("data_time", {})
            return m
        except Exception:
            pass
    return {"updated": None, "modules": {k: [] for k in STREAMS}, "data_time": {}}


def _save_manifest(m):
    os.makedirs(HIST, exist_ok=True)
    tmp = MANIFEST + ".tmp"
    json.dump(m, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    os.replace(tmp, MANIFEST)


def _month_of(date):
    """'2026-09-14' -> '2026-09'（快照按月度目录归档）。"""
    return date[:7]


def _snap_path(mod, date):
    """快照文件路径：hist/<module>/<YYYY-MM>/<YYYY-MM-DD>.json（按月度目录）。"""
    return os.path.join(HIST, mod, _month_of(date), date + ".json")


def _legacy_path(mod, date):
    """旧版扁平路径 hist/<module>/<date>.json（用于一次性迁移/兼容读取）。"""
    return os.path.join(HIST, mod, date + ".json")


def snapshot_now(date=None, pinned=False):
    """把当前 _deploy 的 4 个 JSON 复制为指定日期的快照（按月度目录归档）。"""
    date = date or _today()
    m = _load_manifest()
    changed = False
    os.makedirs(HIST, exist_ok=True)
    for mod, fname in STREAMS.items():
        src = os.path.join(DEPLOY, fname)
        if not os.path.exists(src):
            print("  [skip] %s 缺失 %s" % (mod, fname))
            continue
        # 源 JSON 与已存快照逐字节相同则跳过（避免无谓写入）
        dstdir = os.path.join(HIST, mod, _month_of(date))
        os.makedirs(dstdir, exist_ok=True)
        dst = _snap_path(mod, date)
        # 旧版扁平文件存在则迁到月度目录（平滑升级，不丢历史）
        _migrate_legacy(mod, date)
        if os.path.exists(dst) and _same_file(src, dst):
            # 仍刷新 data_time 记录
            try:
                dt = (json.load(open(src, encoding="utf-8")).get("meta") or {}).get("data_time")
            except Exception:
                dt = None
            m.setdefault("data_time", {}).setdefault(mod, {})[date] = dt
            continue
        shutil.copyfile(src, dst)
        try:
            dt = (json.load(open(src, encoding="utf-8")).get("meta") or {}).get("data_time")
        except Exception:
            dt = None
        m.setdefault("data_time", {}).setdefault(mod, {})[date] = dt
        dates = m.setdefault("modules", {}).setdefault(mod, [])
        if date not in dates:
            dates.append(date)
        changed = True
        print("  [ok] %s/%s/%s.json" % (mod, _month_of(date), date))
    m["updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if pinned:
        pinned_set = set(m.get("pinned", []))
        pinned_set.add(date)
        m["pinned"] = sorted(pinned_set)
    _save_manifest(m)
    _prune(m)
    print("snapshot %s -> %s" % (date, "updated" if changed else "no-change"))
    return changed


def _migrate_legacy(mod, date):
    """把旧扁平快照 hist/<mod>/<date>.json 移到 hist/<mod>/<YYYY-MM>/<date>.json。"""
    old = _legacy_path(mod, date)
    if not os.path.exists(old):
        return
    new = _snap_path(mod, date)
    if not os.path.exists(new):
        try:
            os.makedirs(os.path.dirname(new), exist_ok=True)
            shutil.move(old, new)
        except Exception:
            return
    else:
        try:
            os.remove(old)
        except Exception:
            pass


def _same_file(a, b):
    try:
        return os.path.getsize(a) == os.path.getsize(b) and _md5(a) == _md5(b)
    except Exception:
        return False


def _md5(p):
    import hashlib
    h = hashlib.md5()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _prune(m):
    """仅修剪非 pinned 的过期快照（按 RETENTION_DAYS）。pinned 日期常驻（补录的历史）。"""
    if RETENTION_DAYS >= 9999:
        return
    cutoff = (datetime.date.today() -
              datetime.timedelta(days=RETENTION_DAYS)).strftime("%Y-%m-%d")
    pinned = set(m.get("pinned", []))
    for mod in STREAMS:
        dates = m.get("modules", {}).get(mod, [])
        keep = [d for d in dates if d >= cutoff or d in pinned]
        removed = [d for d in dates if d not in keep]
        for d in removed:
            try:
                os.remove(_snap_path(mod, d))
            except Exception:
                pass
            dt = m.get("data_time", {}).get(mod, {})
            dt.pop(d, None)
        m["modules"][mod] = keep


def backfill(src_root):
    """用历史源 Excel 补录过去日期（后期使用）。

    src_root 下每个子目录以日期命名（YYYY-MM-DD），内含该日的三张源 Excel：
      - 四大模式-重点指标表格.xlsx
      - 某公司-人员数据分析路径.xlsx
      - 某公司-员工三力明细.xlsx
      （文件名须与当前线上源文件完全一致，脚本会临时覆盖再还原）

    流程：逐日期 备份当前源 -> 覆盖为历史源 -> build_snapshot.build() 重建 ->
          snapshot_now(date, pinned=True) -> 还原源。
    """
    if not os.path.isdir(src_root):
        print("源根目录不存在: %s" % src_root)
        return 1
    import build_snapshot as bs
    src_map = {
        "四大模式-重点指标表格.xlsx": os.path.join(BASE_DIR, "四大模式-重点指标表格.xlsx"),
        "某公司-人员数据分析路径.xlsx": os.path.join(BASE_DIR, "某公司-人员数据分析路径.xlsx"),
        "某公司-员工三力明细.xlsx": os.path.join(BASE_DIR, "某公司-员工三力明细.xlsx"),
    }
    # 四大模式实际路径在 分公司健康表/ 下，按 build_snapshot 的候选定位
    aux_candidates = [
        os.path.normpath(os.path.join(BASE_DIR, "..", "..", "..", "分公司健康表", "四大模式-重点指标表格.xlsx")),
        os.path.normpath(os.path.join(BASE_DIR, "分公司健康表", "四大模式-重点指标表格.xlsx")),
    ]
    aux_paths = [p for p in aux_candidates if os.path.exists(p)]
    if aux_paths:
        src_map["四大模式-重点指标表格.xlsx"] = aux_paths[0]

    subdirs = sorted(d for d in os.listdir(src_root)
                     if os.path.isdir(os.path.join(src_root, d)) and _is_date(d))
    if not subdirs:
        print("未找到 YYYY-MM-DD 子目录: %s" % src_root)
        return 1
    print("待补录日期: %s" % ", ".join(subdirs))
    global SUPPRESS_AUTO_SNAPSHOT
    for d in subdirs:
        folder = os.path.join(src_root, d)
        backups = {}
        ok = True
        for name, live in src_map.items():
            hist_file = os.path.join(folder, name)
            if not os.path.exists(hist_file):
                print("  [skip] %s 缺少 %s" % (d, name))
                ok = False
                break
            if os.path.exists(live):
                backups[live] = live + ".bak"
                shutil.copyfile(live, backups[live])
            shutil.copyfile(hist_file, live)
        if not ok:
            # 还原已备份的
            for live, bak in backups.items():
                if os.path.exists(bak):
                    shutil.copyfile(bak, live)
                    os.remove(bak)
            continue
        try:
            # 关键：build() 内部会自动 snapshot_now()（存「当天」）。补录历史时源表已被换成
            # 历史版本，若不禁用该自动快照，会把「当天」快照污染成历史数据。
            SUPPRESS_AUTO_SNAPSHOT = True
            bs.build()
            SUPPRESS_AUTO_SNAPSHOT = False
            snapshot_now(d, pinned=True)
            print("  [done] %s 补录完成" % d)
        except Exception as e:
            SUPPRESS_AUTO_SNAPSHOT = False
            print("  [err] %s: %s" % (d, e))
        finally:
            SUPPRESS_AUTO_SNAPSHOT = False
            for live, bak in backups.items():
                if os.path.exists(bak):
                    shutil.copyfile(bak, live)
                    os.remove(bak)
    return 0


def _is_date(s):
    try:
        datetime.datetime.strptime(s, "%Y-%m-%d")
        return True
    except Exception:
        return False


def main():
    if len(sys.argv) < 2 or sys.argv[1] == "snapshot_now":
        snapshot_now(sys.argv[2] if len(sys.argv) > 2 else None)
    elif sys.argv[1] == "backfill":
        if len(sys.argv) < 3:
            print("用法: python hist_snapshot.py backfill <src_root>")
            return 2
        return backfill(sys.argv[2])
    else:
        print("未知命令: %s" % sys.argv[1])
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
