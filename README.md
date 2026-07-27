# AIPPT — AI 驱动的 PPT 自动生成系统

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![python-pptx 1.0.2](https://img.shields.io/badge/python--pptx-1.0.2-green)](https://pypi.org/project/python-pptx/)
[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](LICENSE)
[![opencode skill](https://img.shields.io/badge/opencode-skill-purple)](SKILL.md)

将用户描述或结构化大纲一键转化为成品 PPT。采用「模板槽位替换 + 自动布局」双引擎架构，100% 保留原模板字体、配色与版式，支持 **11 类商务场景、39 种转场效果（含 Morph）、20+ 动画效果、SmartArt 文本替换、演讲者备注注入、图表/表格动态扩展、3 套动画预设主题、5 套视觉主题（含 Swiss/杂志风设计令牌）、无模板自动布局生成、六层防御校验体系**。

---

## ✨ 核心能力

| 能力 | 说明 |
|------|------|
| **双引擎架构** | Template 引擎（槽位替换，保留模板 100% 样式）+ AutoLayout 引擎（无模板自动生成，基于 Design Token） |
| **11 类场景** | 工作总结 / 年终总结 / 工作汇报 / 工作计划 / 述职报告 / 个人简历 / 自我介绍 / 开题报告 / 公司简介 / 职业规划 / 安全教育 |
| **样式 100% 保留** | 字体、颜色、字号、粗斜体与原模板完全一致，仅替换槽位文本 |
| **文字自适应** | 超长文本自动缩小字号，避免溢出文本框 |
| **版权页清理** | 生成时自动删除模板中版权/广告页（`--keep-copyright` 可保留）|
| **39 种转场效果** | 含 Morph 平滑切换、fade/push/wipe 等商务转场、vortex/flip 等创意转场 |
| **20+ 动画效果** | 入场 / 退场 / 强调三类，支持 `by_bullet` 逐段播放与段间延迟控制 |
| **3 套动画主题** | `business` 简约商务 / `tech` 活力科技 / `formal` 沉稳正式，一键切换 |
| **SmartArt 替换** | 操作 dgm 命名空间 XML，精准替换 SmartArt 文本节点，保留结构与配色 |
| **演讲者备注** | 每页可传入备注文本，自动写入演讲者备注栏 |
| **图表动态扩展** | 柱状/折线/饼/雷达 4 类图表数据源替换，多系列自动适配 |
| **表格动态行扩展** | 传入 N 行数据自动追加/删除行，继承表头样式与列宽 |
| **模板自动标注** | `auto-annotate` 命令输入 PPTX 自动生成完整元数据 + 多页缩略图 |
| **自动布局生成** | 12 列栅格系统 + 5 类标准页面布局，无需模板从零生成可编辑 PPTX |
| **设计令牌系统** | Swiss 风格 / 杂志风格 / 经典商务 3 大类主题预设，6 套开箱即用主题 |
| **CV 设计解析** | 上传设计稿图片自动提取色板、字号层级、间距规范，生成 Design Token |
| **六层防御校验** | JSON Schema + 分步校验 + 模板匹配 + 自动修复 + 错误码体系 |
| **批量渲染** | 同一份内容一键生成同分类下所有模板 PPT |

**性能**：单份生成 0.15~0.92 秒，替换准确率 100%（实测 2000+ 槽位）。

---

## 🚀 快速开始

### 安装

```bash
pip install python-pptx>=1.0.2 jsonschema
# Windows 截图需 pywin32
pip install pywin32  # 可选，用于模板预览图生成
# CV 设计解析需额外依赖
pip install opencv-python pillow numpy  # 可选，用于 design_parser
```

### 一句话生成

```bash
python aippt_outline.py step4-generate \
  --outline my_outline.json \
  --output my_ppt.pptx \
  --transitions auto \
  --animations auto \
  --animation-theme business
```

### 五步工作流

```bash
# Step 1: 理解需求 → 识别场景与参数
python aippt_outline.py step1-understand --text "帮我做一个年终总结PPT"

# Step 2: 构建大纲 → 生成 outline.json（纯 JSON，过 Schema 校验）
python aippt_outline.py step2-outline \
  --scene 年终总结 --purpose "2025年度业绩汇报" \
  --audience 管理层 --length 12 --output outline.json

# Step 3: 视觉匹配 → 推荐模板
python aippt_outline.py step3-visuals --outline outline.json

# Step 4: 生成 PPT（支持动画主题、备注、版权页清理）
python aippt_outline.py step4-generate \
  --outline outline.json \
  --template-id 年终总结_年终总结 \
  --output final.pptx \
  --animation-theme business

# 校验大纲格式（六层防御）
python aippt_outline.py validate --outline outline.json
```

### Python API — Template 引擎

```python
from ppt_scene_adapter import SceneAdapter
from ppt_renderer import PptRenderer

adapter = SceneAdapter("models")
meta, _ = adapter.get_template_meta(template_id="工作总结_工作总结")
slot_data = adapter.adapt("工作总结", business_data, meta)

renderer = PptRenderer("models/工作总结/工作总结.pptx",
                       "models/工作总结/工作总结.meta.json")
renderer.render(slot_data, "output.pptx",
                remove_copyright=True, auto_fit=True,
                transitions="auto", animations="auto",
                animation_theme="business",
                notes_map={1: "封面备注", 4: "KPI页备注"})
```

### Python API — AutoLayout 引擎

```python
from aippt.render import AutoLayoutRenderer

renderer = AutoLayoutRenderer()
renderer.render(
    outline_data={
        "scene": "工作总结",
        "theme": "商务蓝",
        "pages": [
            {"page_type": "cover", "title": "工作总结", "subtitle": "2025年度"},
            {"page_type": "catalog", "title": "目录", "items": ["工作回顾", "成果展示"]},
            {"page_type": "numbered_list", "title": "重点项目",
             "items": [{"title": "项目A", "desc": "完成率120%"}]},
            {"page_type": "kpi", "title": "核心指标",
             "kpi_items": [{"label": "营收", "value": "500万"}, {"label": "增长", "value": "30%"}]}
        ]
    },
    output_path="autolayout_output.pptx",
    render_args={"animation_theme": "tech"}
)
```

### 图表 + 表格 + SmartArt

```bash
# 列出 PPTX 内所有图表/表格
python insert_tables.py list --input template.pptx

# 测试图表数据源替换
python insert_tables.py test-chart --input template.pptx \
  --data chart_data.json --output out.pptx

# 测试表格动态行扩展
python insert_tables.py test-table --input template.pptx \
  --data table_data.json --output out.pptx
```

### 自定义模板接入

```bash
# 自动标注：输入 PPTX 自动生成 .meta.json + 多页缩略图 + 更新索引
python import_templates.py auto-annotate \
  --input my_template.pptx --scene 工作总结 --output models/工作总结/

# 质量门禁：校验元数据完整性、渲染测试、样式检查
python ppt_meta_tool.py check --dir models
```

### 主题系统

```bash
# 列出可用主题
python -c "from aippt.theme_loader import list_themes; print(list_themes())"

# 从预设生成主题（无需图片）
python -c "from aippt.theme_generator import generate_theme_from_preset, save_theme; save_theme(generate_theme_from_preset('swiss_klein_blue'), 'themes/')"

# 从设计稿图片解析生成主题
python -c "from aippt.theme_generator import generate_theme_from_image, save_theme; save_theme(generate_theme_from_image('design.png', '我的主题'), 'themes/')"
```

---

## 📂 项目结构

```
├── aippt/                            # 共享核心包
│   ├── config.py                     # 集中配置（路径/关键词/默认参数）
│   ├── constants.py                  # 共享常量（去重占位文本等）
│   ├── logger.py                     # 统一日志
│   ├── validators.py                 # 六层防御校验引擎
│   ├── animation_themes.py           # 3 套动画预设主题
│   ├── profile_layouts.py            # 母版与版式深度解析
│   ├── ppt_element_classifier.py     # 元素角色识别（标题/正文/KPI）
│   ├── text_replacer.py              # 文本替换子模块（递归遍历+槽位定位+自适应）
│   ├── chart_replacer.py             # 图表数据替换子模块（4类图表+多系列适配）
│   ├── table_filler.py               # 表格动态扩展子模块（行增删+样式继承）
│   ├── design_parser.py              # CV 设计稿解析器（色板/字号/间距抽取）
│   ├── design_tokens.py              # 设计令牌直接抽取（Swiss/杂志风预设）
│   ├── theme_generator.py            # 主题生成器（解析结果→标准主题JSON）
│   ├── theme_loader.py              # 主题加载器（发现/缓存/加载/降级）
│   ├── layout/                       # 自动布局引擎子包
│   │   ├── __init__.py               #   12列栅格系统 + 原子组件 + 页面分发入口
│   │   ├── ppt_auto_layout.py        #   栅格坐标计算、原子绘制组件（622行）
│   │   └── page_layouts.py           #   5类标准页面布局（cover/catalog/divider/...）
│   └── render/                       # 渲染引擎抽象子包
│       ├── __init__.py               #   渲染器注册入口
│       ├── base_renderer.py          #   渲染基类（统一接口 RenderArgs/RenderResult）
│       └── autolayout_renderer.py    #   无模板自动布局渲染器（354行）
├── schemas/                          # JSON Schema 校验规则
│   ├── requirement_params.schema.json
│   └── outline.schema.json
├── tests/                            # 测试体系
│   ├── test_schema.py                # Schema 校验测试
│   ├── test_validators.py            # 校验引擎测试
│   ├── test_smoke_all.py             # 全量冒烟测试
│   ├── test_edge_cases.py            # 边界场景测试
│   ├── test_chart_table.py           # 图表表格专项测试
│   └── test_performance.py           # 性能基准测试
├── themes/                           # 视觉主题配置
│   ├── 商务蓝.json                   # 经典商务主题
│   ├── 科技青.json                   # 科技风主题
│   ├── 极简灰.json                   # 极简风格主题
│   ├── guizang-瑞士风-克莱因蓝.json  # Swiss 风格（克莱因蓝）
│   └── guizang-杂志风-靛蓝瓷.json     # 杂志风格（靛蓝瓷）
├── aippt_outline.py                  # 五步工作流 CLI
├── ppt_renderer.py                   # Template 引擎渲染器（文本/图表/表格/备注/动画）
├── ppt_scene_adapter.py              # 场景适配器（业务字段→槽位映射）
├── ppt_meta_tool.py                  # 模板元数据解析 + 质量门禁
├── ppt_animations.py                 # 动画注入（20+ 效果 + by_bullet）
├── ppt_transitions.py                # 转场注入（39 种效果 + Morph）
├── ppt_smartart.py                   # SmartArt 文本替换
├── insert_tables.py                  # 图表/表格测试工具
├── import_templates.py               # 模板批量导入 + auto-annotate
├── process_safety_templates.py       # 安全教育模板批量处理器
├── migrate_index.py                  # 模板索引迁移工具
├── models/                           # 模板库（按场景分类）
│   ├── 工作总结/ 年终总结/ ... 安全教育/
│   ├── templates_index.json          # 模板总索引（含标签/色系/质量分）
│   └── preview_manifest.json         # 预览图清单
├── SKILL.md                          # opencode skill 定义（六层防御）
├── SKILL_USAGE.md                    # 完整使用手册
└── doc/                              # 文档
    ├── upgrade_guide.md              # 升级指引
    ├── benchmark_reference.md        # 对标项目吸收说明
    └── template_contribution_guide.md # 模板贡献指南
```

---

## 🏗️ 双引擎架构

```
用户描述 / outline.json
        │
        ▼
aippt_outline.py  ── step1 理解 → step2 大纲 → step3 视觉 → step4 生成
        │                  ↓ validate（六层防御校验）
        │
        ├── Template 引擎 ─────────────────────────────── AutoLayout 引擎
        │                                                    │
        ▼                                                    ▼
ppt_scene_adapter.py                                   aippt/render/
  ── validate → adapt                                 AutoLayoutRenderer
  （业务字段→槽位映射）                                    │
        │                                                ▼
        ▼                                          aippt/layout/
ppt_renderer.py                                      栅格系统 + 页面布局
  ── text_replacer（文本替换）                              │
  ── chart_replacer（图表）                                ▼
  ── table_filler（表格）                             Design Token
  ── SmartArt 替换                                     (themes/*.json)
  ── 演讲者备注注入                                         │
  ── 动画/转场注入                                        ▼
        │                                          原生可编辑 PPTX 元素
        ▼                                               （无模板）
保留模板 100% 样式
        │
        ▼
    final.pptx ✅
```

**Template 引擎**：适合需要严格遵循企业品牌模板的场景，100% 保留原模板字体、配色、版式、背景。

**AutoLayout 引擎**：适合从零快速生成 PPT 的场景，基于 12 列栅格系统和 Design Token，无需任何模板即可生成原生可编辑 PPTX。

---

## 🎨 设计令牌与主题体系

### Design Token 架构

```
                    ┌─────────────────────┐
                    │   design_parser.py   │ ← CV 解析设计稿图片
                    │ (OpenCV + Pillow)    │
                    └─────────┬───────────┘
                              │ 色板/字号/间距
                              ▼
┌─────────────────┐   ┌──────────────┐
│ design_tokens.py│ ← │ 主题生成器    │ → themes/*.json
│ (Swiss/杂志预设) │   │ theme_generator│
└─────────────────┘   └──────┬───────┘
                             │ 标准主题 JSON
                             ▼
                      ┌──────────────┐
                      │ theme_loader │ → get_style(role, attr)
                      │ (缓存+降级)   │
                      └──────┬───────┘
                             │ Design Token 查询接口
                             ▼
                      layout / render 引擎
```

### 预置主题

| 主题文件 | 风格 | 色系 |
|----------|------|------|
| `商务蓝.json` | 经典商务 | 蓝色系 |
| `科技青.json` | 科技风 | 青色系 |
| `极简灰.json` | 极简主义 | 灰色系 |
| `guizang-瑞士风-克莱因蓝.json` | Swiss Design | 克莱因蓝 + 无衬线 |
| `guizang-杂志风-靛蓝瓷.json` | 杂志风格 | 靛蓝瓷 + 衬线体 |

可通过 `theme_generator.py` 从设计稿图片或预设名称自动生成新主题。

---

## 📐 OOXML 技术实现

本系统深度使用 **OOXML（Office Open XML）标准** 进行 PPTX 底层操作：

| OOXML 规范 | 实现模块 | 操作内容 |
|-----------|---------|---------|
| **p:dgm** (SmartArt) | `ppt_smartart.py` | 操作 dgm 命名空间 XML，精准替换 SmartArt 文本节点 |
| **c:chart** (图表) | `aippt/chart_replacer.py` | ChartData 数据源替换，保留样式/轴格式/图例位置 |
| **a:tbl** (表格) | `aippt/table_filler.py` | 行复制/删除，单元格样式继承（填充/边框/字体）|
| **p:timing** (动画) | `ppt_animations.py` | 构建动画 XML 树，控制缓动/延迟/触发条件 |
| **p:transition** (转场) | `ppt_transitions.py` | 39 种转场 XML 注入，含 Morph 平滑切换参数 |
| **p:spTree** (形状树) | `aippt/text_replacer.py` | 递归遍历 group shape，3 层槽位定位匹配 |
| **a:rPr** (文本样式) | `aippt/text_replacer.py` | 保留首个 run 的字体/颜色/字号/粗斜体 |
| **p:notes** (备注) | `ppt_renderer.py` | 演讲者备注 XML 节点注入，支持多页 |
| **p:sldLayout** (版式) | `aippt/profile_layouts.py` | 母版与版式深度解析，分析占位符层级 |

---

## 🎬 动画 & 转场

- **39 种转场**：`fade` / `push` / `wipe` / `dissolve` / `zoom` / `morph`（平滑切换）/ `flip` / `vortex` 等
- **20+ 动画效果**：入场（`fade` / `fly_in` / `zoom` / `wipe`）+ 退场 + 强调（`pulse` / `spin`）
- **3 套预设主题**：`--animation-theme business/tech/formal` 一键切换全套动画风格
- **两级配置**：全局 `--transitions` / `--animations` + 单页 `transition` / `animations` 字段，单页优先
- **逐段动画**：`by_bullet` 按段落逐步显示，支持段间延迟、播放顺序自定义
- **自动推荐**：`auto` 模式按页面类型自动匹配动画方案

**优先级**（高 → 低）：单页显式配置 > 动画主题 > 全局参数

---

## 🛡️ 六层防御体系

| 层级 | 防御点 | 实现 |
|------|--------|------|
| 第一层 | 模型端自检 | SKILL.md 指令引导（铁则/正例/反例/自检清单）|
| 第二层 | JSON Schema 机器化校验 | `schemas/*.schema.json` + `jsonschema` |
| 第三层 | 分步校验流程 | `validate_requirement` / `validate_outline` |
| 第四层 | 模板槽位匹配校验 | `validate_template_match`（结合 meta.json）|
| 第五层 | 运行时兜底自动修复 | `auto_fix_outline`（page_id重排/截断/枚举标准化）|
| 第六层 | 标准化错误反馈 | 错误码体系 F0xx/S0xx/T0xx/A0xx + 修正建议 |

错误码覆盖：基础格式（F）、字段规则（F1）、结构逻辑（S）、模板匹配（T）、动画转场（A001-A006）。

---

## 🧪 测试

```bash
# 单元测试（快速，不渲染 PPTX）
python -m pytest tests/ -m "not slow" -v

# 全量冒烟测试（渲染所有模板）
python -m pytest tests/test_smoke_all.py -v

# 图表表格专项测试（动态构造 PPTX，不依赖外部模板）
python -m pytest tests/test_chart_table.py -v

# 性能基准测试
python -m pytest tests/test_performance.py -v

# 元数据质量校验
python ppt_meta_tool.py check --dir models
```

---

## 📖 完整文档

- **[SKILL.md](SKILL.md)** — opencode skill 定义，含六层防御、12 类页面示例、动画转场枚举
- **[SKILL_USAGE.md](SKILL_USAGE.md)** — 完整使用手册（CLI / API / 数据格式 / 模板接入）
- **[doc/upgrade_guide.md](doc/upgrade_guide.md)** — 从旧版本升级指引
- **[doc/benchmark_reference.md](doc/benchmark_reference.md)** — 对标项目吸收说明
- **[doc/template_contribution_guide.md](doc/template_contribution_guide.md)** — 模板贡献指南

---

## 📜 许可

GPL-3.0
