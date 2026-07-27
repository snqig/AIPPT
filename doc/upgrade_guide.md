# AIPPT 升级指引

> 适用范围：基于 `python-pptx 1.0.2` 的 PPT 自动生成系统
> 架构基础：槽位替换 + 五步工作流
> 当前版本：v2.1
> 兼容策略：100% 向后兼容，所有新增能力均为可选参数

---

## 一、版本演进概览

### 1.1 v1.0 —— 初始版本

- 10 类标准场景（工作总结 / 述职汇报 / 商业计划书 / 教学课件 / 产品发布 / 培训资料 / 年终汇报 / 项目复盘 / 竞聘演讲 / 通用模板）。
- 确立「槽位替换」核心架构：模板占位符 `{{title}}` / `{{content}}` / `{{subtitle}}` 等通过 outline.json 注入。
- 确立「五步工作流」：理解需求 → 选模板 → 填大纲 → 渲染 → 校验输出。
- 基础动画与转场（淡入 / 推入 / 擦除等若干种）。

### 1.2 v2.0 —— 能力扩展

- 新增 `aippt` 核心包：`aippt/config`、`aippt/constants`、`aippt/logger`，统一配置与日志。
- 转场扩展至 **38 种**，覆盖 PPT 全部内置转场类别。
- 动画扩展至 **20+ 种**，含进入 / 强调 / 退出 / 路径四类。
- 支持模板批量导入（`import_templates.py`）。
- 支持多页缩略图生成，模板预览从单图升级为多页拼图。

### 1.3 v2.1 —— 质量与治理

- **六层防御校验体系**：从格式 / 结构 / 模板 / 动画 / 数据契约 / 输出前自检六层保证大纲质量。
- **SmartArt 替换**：`ppt_smartart.py` 自动识别并替换 SmartArt 文本节点。
- **演讲者备注**：渲染时按 `notes_map` 注入每页备注。
- **图表 / 表格动态扩展**：`chart_data` / `table_data` 自动填充，保留模板样式。
- **3 套动画预设主题**：`business` / `tech` / `formal`，开箱即用。
- **模板自动标注**（`auto-annotate`）：一键生成 `.meta.json` + 缩略图 + 索引更新 + 质量报告。
- **Morph 平滑切换**：转场总数提升至 **39 种**。
- **安全教育场景**：场景扩展至 **11 类**，新增安全教育。
- **质量门禁**：模板入库前强制质量评分，低于阈值拒绝入库。
- **模板标签体系**：`style_tags` / `color_scheme` / `industry` / `page_range` / `quality_score` 五维标签。

---

## 二、从 v1.x 升级到 v2.1

### 2.1 破坏性变更（无）

> 强调：v2.1 实现 **100% 向后兼容**，所有新增能力均为可选参数，不传则行为与 v1.x 完全一致。

| 新增能力 | 默认值 | 不传时行为 |
| --- | --- | --- |
| `notes_map` | `None` | 不渲染备注，与旧版一致 |
| `animation_theme` | `None` | 不套用主题，与旧版一致 |
| `--transitions` / `--animations` | `auto` | 自动选择，与旧版一致 |
| `removable_pages` | 未配置 | 不删除任何页 |
| `chart_data` / `table_data` | 未提供 | 不触发动态扩展 |
| SmartArt 替换 | 自动识别 | 无 SmartArt 时静默跳过 |

### 2.2 新增依赖

| 依赖包 | 版本建议 | 是否必选 | 用途 |
| --- | --- | --- | --- |
| `jsonschema` | ≥ 4.0 | 必选 | 六层防御校验的 JSON Schema 引擎 |
| `pywin32` | ≥ 306 | 可选 | Windows 平台通过 COM 调用 PowerPoint 生成模板预览图 |

安装命令：

```bash
pip install jsonschema
pip install pywin32   # 仅 Windows 需要，且仅用于模板预览图
```

### 2.3 配置迁移

v2.1 对模板元数据做了字段扩展，但 **旧版 `meta.json` 无需手动修改**，迁移工具会自动补全：

1. **新增字段自动补全**：`style_tags` / `color_scheme` / `industry` / `page_range` / `quality_score` 由 `ppt_meta_tool.py` 在校验时自动补全默认值。
2. **索引文件迁移**：为旧版 `templates_index.json` 补充新标签字段。

```bash
# 1. 迁移索引文件，补充新标签字段
python migrate_index.py

# 2. 重新校验模板质量并补全 meta.json
python ppt_meta_tool.py check --dir models

# 3.（可选）批量重新生成缩略图
python import_templates.py refresh-thumbnails --dir models
```

---

## 三、新增能力启用指南

### 3.1 启用六层防御校验

校验在渲染前执行，输出结构化 JSON 报告，便于程序化消费。

```bash
python aippt_outline.py validate --outline outline.json
```

输出结构：

```json
{
  "validate_pass": false,
  "errors":   [{ "code": "S01", "field": "pages", "message": "...", "suggestion": "...", "auto_fix": false }],
  "warnings": [{ "code": "A02", "field": "animations", "message": "...", "suggestion": "..." }],
  "fixes":    [{ "code": "F01", "field": "title", "from": "...", "to": "...", "applied": true }]
}
```

错误码体系：

| 前缀 | 类别 | 含义 |
| --- | --- | --- |
| `F0xx` | 格式 | JSON 格式 / 字段类型 / 编码问题 |
| `S0xx` | 结构 | 大纲层级 / 页面顺序 / 槽位缺失 |
| `T0xx` | 模板 | 模板不存在 / 槽位不匹配 / 场景冲突 |
| `A0xx` | 动画 | 动画名非法 / 转场名非法 / 主题冲突 |

> 校验失败 **不阻断渲染**：`errors` 会阻断，`warnings` 仅提示不影响生成。

### 3.2 启用动画预设主题

通过命令行参数一键套用预设主题：

```bash
python aippt_outline.py render --outline outline.json --animation-theme business
```

可选主题：`business`（商务稳重）/ `tech`（科技灵动）/ `formal`（正式克制）。

**优先级规则**（高 → 低）：

1. 单页显式配置（`page.animations` / `page.transition`）
2. 主题预设（`--animation-theme`）
3. 全局默认（`auto`）

> 主题未传时，行为与旧版完全一致，不会引入任何动画变化。

### 3.3 启用演讲者备注

通过 Python API 的 `notes_map` 参数注入：

```python
from ppt_renderer import Renderer

renderer = Renderer(template_path="models/工作总结/template.pptx")
renderer.render(
    outline_path="outline.json",
    output_path="output.pptx",
    notes_map={
        1: "封面页：开场问候，介绍汇报人身份与主题背景。",
        2: "目录页：概述本次汇报的四大板块，控制在30秒内。",
        3: "数据页：重点强调同比增长 23% 的关键指标。",
    },
)
```

> `notes_map` 的键为 `page_id`（从 1 开始，对应 outline 中页面的顺序索引）。未提供的页面不注入备注。

### 3.4 启用图表 / 表格动态扩展

在 `outline.json` 中为对应页面提供数据即可：

```json
{
  "pages": [
    {
      "page_id": 3,
      "type": "chart",
      "chart_data": {
        "chart_type": "bar",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
          { "name": "营收", "values": [120, 150, 180, 210] },
          { "name": "利润", "values": [30, 45, 60, 75] }
        ]
      }
    },
    {
      "page_id": 4,
      "type": "table",
      "table_data": {
        "headers": ["项目", "负责人", "进度", "状态"],
        "rows": [
          ["需求调研", "张三", "100%", "已完成"],
          ["方案设计", "李四", "80%",  "进行中"],
          ["开发实施", "王五", "30%",  "进行中"]
        ]
      }
    }
  ]
}
```

渲染引擎自动识别 `chart_data` / `table_data` 并调用 `replace_data` 替换模板占位图表 / 表格，**保留模板原有样式**（字体 / 配色 / 边框 / 图例位置）。

### 3.5 启用 SmartArt 替换

当模板包含 SmartArt 图形时，`ppt_smartart.py` 会自动识别并按节点替换文本：

```python
from ppt_smartart import replace_smartart_text

replace_smartart_text(
    slide=slide,
    smartart_index=0,
    text_map={
        "node_1": "战略目标",
        "node_2": "关键举措",
        "node_3": "预期成果",
    },
)
```

> SmartArt 替换为 **自动触发**：渲染引擎检测到模板含 SmartArt 且 outline 提供对应 `smartart_data` 时自动调用，无需手动干预。

### 3.6 启用模板自动标注

一键完成模板入库全流程：

```bash
python import_templates.py auto-annotate \
  --input my.pptx \
  --scene 工作总结 \
  --output models/工作总结/
```

`auto-annotate` 自动完成：

1. 生成 `.meta.json`（含 `style_tags` / `color_scheme` / `industry` / `page_range` / `quality_score`）。
2. 生成多页缩略图（首页 + 中间页 + 末页拼图）。
3. 更新 `templates_index.json` 索引。
4. 输出质量报告（不符合门禁的项会标记 `quality_gate: fail`）。

---

## 四、SKILL.md 升级要点

v2.1 的 SKILL.md 相较 v1.x 新增以下章节，**旧的提示词与示例全部保留**，仅做增量扩展：

| 新增章节 | 作用 |
| --- | --- |
| 数据契约（7 条铁则） | 明确 outline.json 字段类型 / 取值范围 / 必填项的硬约束 |
| 高频错误对照表 | 常见错误码 → 原因 → 修复方案的快速索引 |
| 12 类页面最小示例 | 每类页面给出最小可渲染的 outline 片段 |
| 输出前强制 10 项自检清单 | 渲染前必须逐项确认的检查点 |
| 动画转场枚举 | 39 种转场 + 20 种动画的完整枚举表 |
| 六层防御体系 | 校验流程 / 错误码 / 修复策略的总览 |
| 高级渲染能力 | SmartArt / 图表扩展 / 表格扩展 / 备注 / 主题的启用说明 |

---

## 五、回滚方案

v2.1 的所有新特性均为可选，回滚成本低：

1. **参数级回滚**：删除新增参数（`notes_map` / `animation_theme` / `chart_data` / `table_data` / `removable_pages`），行为立即回退到 v1.x。
2. **校验不阻断**：`warnings` 不影响生成；`errors` 可通过 `--skip-validate` 跳过。
3. **完全回滚**：保留 v1.x 的 `SKILL.md` 与 `aippt_outline.py`，移除 `aippt` 核心包与 `ppt_smartart.py`，即可恢复 v1.x 全部行为。

> 回滚不影响已生成的 PPT 文件，仅影响后续渲染流程。

---

## 六、常见问题

**Q1：升级后旧模板还能用吗？**
A：可以。v2.1 100% 向后兼容，旧模板的 `meta.json` 缺失新字段时由 `ppt_meta_tool.py` 自动补全默认值，渲染行为不变。

**Q2：校验报错怎么办？**
A：查看错误码（`F0xx` / `S0xx` / `T0xx` / `A0xx`），按 `suggestion` 字段修正。对于非原则性错误（如缺失可选字段），可使用 `auto_fix` 自动修复；或加 `--skip-validate` 跳过校验直接渲染（不推荐）。

**Q3：动画主题和单页配置冲突时如何处理？**
A：单页显式配置优先，主题作为该页的默认值。即：页面未显式声明 `animations` / `transition` 时才套用主题。

**Q4：图表替换后样式会变吗？**
A：不会。`replace_data` 仅更新数值与分类，保留模板原有的字体 / 配色 / 边框 / 图例 / 坐标轴样式。

**Q5：SmartArt 替换失败会中断渲染吗？**
A：不会。SmartArt 替换失败时记录 `warning` 并跳过该图形，继续渲染其他内容。

**Q6：`pywin32` 不安装会影响核心功能吗？**
A：不会。`pywin32` 仅用于 Windows 平台的模板预览图生成（`auto-annotate` 的缩略图环节）。核心渲染 / 校验 / 动画 / 备注等能力均不依赖它。

**Q7：v1.x 的 outline.json 需要改写吗？**
A：不需要。v1.x 的 outline.json 可直接用于 v2.1 渲染。新增字段（`chart_data` / `table_data` / `smartart_data` / `notes` 等）均为可选，不提供则不触发对应能力。

---

## 附录：版本对照速查表

| 能力 | v1.0 | v2.0 | v2.1 |
| --- | --- | --- | --- |
| 场景数量 | 10 | 10 | 11 |
| 转场数量 | 基础若干 | 38 | 39 |
| 动画数量 | 基础若干 | 20+ | 20+ |
| aippt 核心包 | ✗ | ✓ | ✓ |
| 六层防御校验 | ✗ | ✗ | ✓ |
| SmartArt 替换 | ✗ | ✗ | ✓ |
| 演讲者备注 | ✗ | ✗ | ✓ |
| 图表 / 表格动态扩展 | ✗ | ✗ | ✓ |
| 动画预设主题 | ✗ | ✗ | ✓（3 套） |
| 模板自动标注 | ✗ | 批量导入 | ✓（auto-annotate） |
| Morph 平滑切换 | ✗ | ✗ | ✓ |
| 质量门禁 | ✗ | ✗ | ✓ |
| 模板标签体系 | ✗ | ✗ | ✓（5 维） |
| 多页缩略图 | ✗ | ✓ | ✓ |
| 向后兼容 | — | ✓ | ✓ |
