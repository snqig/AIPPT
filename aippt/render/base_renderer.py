"""
BaseRenderer 渲染基类（T001）

抽象层目标：
    1. 定义统一渲染接口 render(outline_data, output_path, render_args)
    2. 双引擎共用入参规范：outline 原始字典 + 输出路径 + 渲染参数
    3. 统一动画 / 转场 / 主题 / 备注 / 自适应等可选参数封装为 RenderArgs
    4. 现有 PptRenderer 与新增 AutoLayoutRenderer 均继承本基类

设计原则：
    - 100% 向后兼容：现有 PptRenderer.render(slot_data, ...) 旧签名保留
    - 抽象接口与具体实现解耦：上层调度按 mode 选择渲染器
    - 公共能力（动画 / 转场 / 备注）由基类提供默认实现，子类可覆盖
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class RenderArgs:
    """渲染参数容器（双引擎统一入参规范）

    封装所有可选渲染配置，避免 render() 方法签名无限膨胀。
    所有字段均有默认值，缺省时行为与原版完全一致。

    :param remove_copyright: 是否自动删除版权页（仅 template 模式生效）
    :param auto_fit: 是否启用长文本字号自适应
    :param transitions: 全局转场配置 "auto"/"none"/dict/None
    :param animations: 全局动画配置 "auto"/"none"/dict/None
    :param notes_map: 演讲者备注 {page_id: 备注文本}（page_id 从 1 开始）
    :param animation_theme: 动画主题名（business/tech/formal）
    :param theme: 视觉主题名（仅 auto 模式生效，如 商务蓝/极简灰/科技青）
    :param extra: 引擎专属参数扩展位，避免基类频繁变更
    """
    remove_copyright: bool = True
    auto_fit: bool = True
    transitions: Optional[Any] = None
    animations: Optional[Any] = None
    notes_map: Optional[dict[int, str]] = None
    animation_theme: Optional[str] = None
    theme: Optional[str] = None
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass
class RenderResult:
    """渲染结果结构（统一输出规范）

    :param output_path: 实际输出 PPTX 路径
    :param mode: 渲染模式（template / auto）
    :param total_pages: 生成 PPT 总页数
    :param replaced: 替换/绘制元素数（template 模式为槽位替换数，auto 模式为元素数）
    :param missed: 未匹配槽位数（auto 模式恒为 0）
    :param skipped: 跳过槽位数（auto 模式恒为 0）
    :param removed_pages: 已删除版权页列表
    :param warnings: 渲染过程中的警告信息列表
    :param meta: 引擎专属输出元数据（如 shape_id 映射表等）
    """
    output_path: str
    mode: str
    total_pages: int
    replaced: int = 0
    missed: int = 0
    skipped: int = 0
    removed_pages: list[int] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    meta: dict[str, Any] = field(default_factory=dict)


class BaseRenderer(ABC):
    """渲染引擎抽象基类

    所有渲染引擎实现（PptRenderer / AutoLayoutRenderer）必须继承本类，
    并实现 render_outline() 抽象方法。

    接口设计说明：
        - render_outline(outline_data, output_path, render_args)：统一抽象接口
        - render(...)：各子类可保留自有签名（如 PptRenderer.render(slot_data, ...)）
          以保证 100% 向后兼容，不强制覆盖基类
        - 上层调度代码统一调用 render_outline()，由子类转换为内部逻辑
    """

    #: 引擎模式标识，子类必须覆盖（"template" / "auto"）
    MODE: str = "base"

    @abstractmethod
    def render_outline(
        self,
        outline_data: dict[str, Any],
        output_path: str,
        render_args: Optional[RenderArgs] = None,
    ) -> RenderResult:
        """渲染 PPT 入口（抽象方法，子类必须实现）

        :param outline_data: outline.json 原始字典
            - template 模式：从中提取 scene / pages / sections 等业务数据
            - auto 模式：直接消费 pages 数组生成布局
        :param output_path: 输出 PPTX 路径
        :param render_args: 渲染参数容器，None 时使用默认 RenderArgs()
        :return: RenderResult 渲染结果
        """
        raise NotImplementedError

    @staticmethod
    def normalize_args(render_args: Optional[RenderArgs]) -> RenderArgs:
        """规范化 RenderArgs，None 转为默认实例

        :param render_args: 入参 RenderArgs 或 None
        :return: 非 None 的 RenderArgs 实例
        """
        if render_args is None:
            return RenderArgs()
        return render_args

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__} mode={self.MODE!r}>"
