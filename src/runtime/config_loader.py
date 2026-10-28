# -*- coding: utf-8 -*-
"""
看板配置加载器（配置即代码）
================================================================
所有易变规则（岗位白名单 / 部门映射 / 筛选 / 离职排除 / 校验阈值）
集中到 config/dashboard.config.json。

设计原则（防止"变故"）：
  - 本加载器**永远返回完整配置**：配置文件缺失、格式损坏、字段缺失，
    均自动回退到与现状完全一致的内置 DEFAULT_CONFIG，解析器行为不变。
  - 仅使用标准库 json，零第三方依赖。
  - 修改规则 = 编辑 config 文件并重新构建，无需改任何 parser 代码。
"""
import os
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config", "dashboard.config.json")

# 与现状完全一致的默认值。即使配置文件整份丢失，行为也不变。
DEFAULT_CONFIG = {
    "personnel": {
        # 仅保留公司名含该关键词的分公司（当前：分公司）
        "company_keyword": "",
        # 人员数据分析看板仅保留的岗位（含本次新增的「企服规划师」）
        "target_posts": [
            "电邀规划师", "企服规划师", "企服主管", "企服储备主管",
            "流程会计", "会计顾问", "客服会计", "法务主任",
        ],
        # 部门关键词 → 标准部门名（用于 LTC 部门归类）
        "dept_rules": [
            ["工商", "工商"], ["会计", "会计"], ["法律", "法律"],
            ["提成律师", "法律"], ["跨境", "跨境"],
            ["数字科技AI", "数字科技AI"], ["数字科技", "数字科技AI"],
            ["交易", "交易"],
        ],
        # 部门展示顺序
        "dept_order": ["工商", "会计", "法律", "跨境", "数字科技AI", "交易"],
        # 非空防护：解析后人数低于此值判定为编辑中间态（脏数据）
        "min_headcount": 10,
    },
    "sanli": {
        # 三力看板排除的状态（离职人员不计入）
        "exclude_status": ["离职"],
    },
    "verify": {
        # 部署前产物校验阈值（任一不达标则阻断发布）
        "ry_min_headcount": 10,
        "sanli_min_headcount": 10,
        "activity_min_depts": 1,
    },
}

_cache = None


def _deep_merge(default, override):
    """递归合并：override 中的值覆盖 default 的同名键，缺字段用 default。"""
    out = dict(default)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config():
    """加载配置；任何异常都回退到 DEFAULT_CONFIG，保证解析器行为不变。"""
    global _cache
    if _cache is not None:
        return _cache
    cfg = DEFAULT_CONFIG
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            user = json.load(f)
        cfg = _deep_merge(DEFAULT_CONFIG, user)
    except FileNotFoundError:
        print("[config] 未找到配置文件，使用内置默认值（行为不变）")
    except Exception as e:
        print("[config] 配置解析失败，使用内置默认值：", e)
    _cache = cfg
    return cfg


if __name__ == "__main__":
    import pprint
    pprint.pprint(load_config())
