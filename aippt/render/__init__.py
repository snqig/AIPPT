"""
aippt.render 子包：渲染引擎抽象层与实现

模块说明：
    - base_renderer.py：抽象基类 BaseRenderer，定义统一渲染接口
    - autolayout_renderer.py：无模板自动布局渲染器（T005）
    - 现有 ppt_renderer.py 中的 PptRenderer 继承 BaseRenderer（T001）

设计目标：
    1. 双引擎共用一套 outline.json / 动画 / 转场 / 校验规范
    2. 100% 向后兼容，旧调用方无需修改
    3. 上层根据 mode 动态选择渲染实现（template / auto）
"""
from aippt.render.base_renderer import BaseRenderer, RenderArgs, RenderResult
from aippt.render.autolayout_renderer import AutoLayoutRenderer

__all__ = ["BaseRenderer", "RenderArgs", "RenderResult", "AutoLayoutRenderer"]
