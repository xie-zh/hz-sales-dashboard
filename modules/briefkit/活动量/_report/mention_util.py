# -*- coding: utf-8 -*-
"""
@ 提及工具（企业微信智能机器人 aibot 场景）
============================================
结论（2026-10-09 实测，务必不要再改成 <@userid>）：

- aibot 的 markdown **不支持任何 @ 语法**，`<@userid>` 不会被渲染，
  而是把尖括号内容原样显示出来（真实事故：群里出现 `<@<WECOM_USERID>>`）。
- `wecom-cli contact users search` 返回的 `userid` 是**框架加密句柄**（形如 `woPClcBwAA…`），
  连授权真人（王梦洁）也是同一形态 —— 它不是企业微信真实 userid，写进正文也不会被解密。
- `message aibot send` 的 markdown 只有 `content` 字段，没有 `mentioned_list`；
  `message send` 仅支持纯文本，同样没有提及字段。

因此本模块统一产出**纯文本 `@姓名`**：群里可读、不会出现乱码，
但**不会触发对方提醒**（仅视觉上的 @）。

如需真正触发提醒，唯一可靠通道是「群机器人 Webhook」（text + mentioned_list），
且必须使用**真实 userid 或手机号**（框架加密句柄无效）。
把各群 webhook 地址填进 `数据文件/_common/webhooks.json` 后另行实现。
"""


def resolve_names(names):
    """兼容旧调用：不再联网解析（加密句柄对 @ 无用）。返回空映射。"""
    return {}


def mention(name):
    """返回纯文本 @姓名（aibot 无法真@，详见模块说明）。"""
    return '@' + str(name)
