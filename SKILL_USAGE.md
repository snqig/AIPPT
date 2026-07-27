# PPT 生成 Skill 使用手册

## 一、概述

PPT 生成 Skill 是一套基于 python-pptx 1.0.2 的自动化 PPT 生成工具，支持 **11 类商务场景**（工作总结 / 年终总结 / 工作汇报 / 工作计划 / 述职报告 / 个人简历 / 自我介绍 / 开题报告 / 公司简介 / 职业规划 / 安全教育），实现「输入结构化 JSON → 自动匹配模板 → 一键生成 PPT」的全流程自动化。采用「槽位替换」架构，100% 保留模板样式。

### 核心能力

- **11 类场景**：预定义 11 类业务场景的 Schema，覆盖常见商务 PPT 需求
- **236 套模板**：内置 236 套模板元数据，含安全教育 157 套，自动匹配章节结构与槽位
- **样式 100% 还原**：字体、颜色、布局、形状与原模板完全一致
- **39 种转场效果**：含 Morph 平滑切换、fade/push/wipe 等商务转场
- **20+ 动画效果**：入场/退场/强调三类，支持 `by_bullet` 逐段播放
- **3 套动画主题**：`business`/`tech`/`formal` 一键切换全套动画风格
- **SmartArt 替换**：操作 dgm 命名空间 XML，精准替换 SmartArt 文本节点
- **演讲者备注注入**：每页可传入备注文本，自动写入演讲者备注栏
- **图表动态扩展**：柱状/折线/饼/雷达 4 类图表数据源替换，保留模板样式
- **表格动态行扩展**：传入 N 行数据自动追加/删除行，继承表头样式
- **模板自动标注**：`auto-annotate` 命令输入 PPTX 自动生成元数据 + 缩略图
- **六层防御校验**：JSON Schema + 分步校验 + 模板匹配 + 自动修复 + 错误码
- **版权页自动清理**：生成时自动删除标记的广告 / 版权页
- **长文本自适应**：超长文本自动缩小字号，避免溢出
- **批量生成**：同一份内容一键生成同分类所有模板 PPT

### 性能指标

| 指标 | 目标 | 实测 |
|------|------|------|
| 单份生成耗时 | ≤ 3 秒 | 0.15 - 0.92 秒 |
| 替换准确率 | ≥ 98% | 100%（0 未匹配） |
| 样式还原度 | 100% | 100% |
| 批量 12 套耗时 | ≤ 20 秒 | 3.92 秒 |

---

## 二、安装与依赖

### 2.1 环境要求

- Python 3.8+
- python-pptx 1.0.2

### 2.2 安装依赖

```bash
pip install python-pptx==1.0.2
```

### 2.3 项目结构

```
Ppt_work/
├── aippt/                          # 共享核心包
│   ├── config.py                   # 集中配置
│   ├── constants.py                # 共享常量
│   ├── logger.py                   # 统一日志
│   ├── validators.py               # 六层防御校验引擎
│   ├── animation_themes.py         # 3 套动画预设主题
│   ├── profile_layouts.py          # 母版与版式深度解析
│   ├── ppt_element_classifier.py   # 元素角色识别
│   ├── text_replacer.py            # 文本替换 / shape 定位 / 字号自适应
│   ├── chart_replacer.py           # 图表数据源替换（4 类图表）
│   └── table_filler.py             # 表格动态行扩展
├── schemas/                        # JSON Schema 校验规则
│   ├── requirement_params.schema.json
│   └── outline.schema.json
├── tests/                          # 测试体系
│   ├── test_schema.py              # Schema 校验测试
│   ├── test_validators.py          # 校验引擎测试
│   ├── test_smoke_all.py           # 全量冒烟测试
│   ├── test_edge_cases.py          # 边界场景测试
│   ├── test_chart_table.py         # 图表表格专项测试
│   └── test_performance.py         # 性能基准测试
├── aippt_outline.py                # 五步工作流 CLI
├── ppt_renderer.py                 # 核心渲染引擎（调度子模块）
├── ppt_scene_adapter.py            # 场景适配层（Skill 主入口）
├── ppt_meta_tool.py                # 元数据处理 + 质量门禁
├── ppt_animations.py               # 动画注入（20+ 效果）
├── ppt_transitions.py              # 转场注入（39 种效果）
├── ppt_smartart.py                 # SmartArt 文本替换
├── insert_tables.py                # 图表/表格测试工具
├── import_templates.py             # 模板批量导入 + auto-annotate
├── models/                         # 模板资产库
│   ├── 工作总结/ 年终总结/ ... 安全教育/  # 11 个分类目录
│   ├── templates_index.json        # 模板总索引（含标签/色系/质量分）
│   └── preview_manifest.json       # 预览图清单
├── SKILL.md                        # opencode skill 定义
├── SKILL_USAGE.md                  # 本手册
└── doc/                            # 补充文档
    ├── upgrade_guide.md            # 升级指引
    ├── benchmark_reference.md      # 对标项目吸收说明
    └── template_contribution_guide.md  # 模板贡献指南
```

### 2.4 依赖安装

```bash
pip install python-pptx==1.0.2 jsonschema
# Windows 模板预览图生成（可选）
pip install pywin32
```

---

## 三、CLI 命令手册

Skill 主入口为 `ppt_scene_adapter.py`，支持 7 个子命令。

### 3.1 查看支持的场景

```bash
python ppt_scene_adapter.py scenes
```

输出所有支持的 10 类场景及模板数量统计。

### 3.2 查看场景 Schema

```bash
python ppt_scene_adapter.py schema --scene 工作总结
```

输出指定场景的字段定义，包括封面字段（cover_fields）、章节定义（chapter_sections）、结束字段（end_fields）。

### 3.3 列出模板

```bash
# 列出所有模板
python ppt_scene_adapter.py templates

# 按分类筛选
python ppt_scene_adapter.py templates --category 工作总结
```

### 3.4 查看模板详情

```bash
python ppt_scene_adapter.py detail --template 工作总结_worksummary_01
```

输出模板的总页数、章节数、槽位总数、章节结构、各页槽位分布。

### 3.5 校验业务数据

```bash
python ppt_scene_adapter.py validate --scene 工作总结 --data business_worksummary.json
```

校验 JSON 数据是否符合场景 Schema，输出缺失字段或不符格式的提示。

### 3.6 单模板渲染

```bash
# 使用分类下第一个模板
python ppt_scene_adapter.py render --scene 工作总结 --data business_worksummary.json --output output.pptx

# 指定模板
python ppt_scene_adapter.py render --scene 工作总结 --data business_worksummary.json --template 工作总结_worksummary_01 --output output.pptx
```

### 3.7 批量渲染

```bash
python ppt_scene_adapter.py batch --scene 工作总结 --data business_worksummary.json --output-dir output_batch
```

对同分类下所有模板批量生成 PPT，输出到指定目录。

---

## 四、API 调用手册

### 4.1 初始化

```python
from ppt_scene_adapter import SceneAdapter

adapter = SceneAdapter(templates_root="models")
```

### 4.2 查询场景与模板

```python
# 列出所有场景
scenes = adapter.list_scenes()

# 查看场景Schema
schema = adapter.get_scene_schema("工作总结")

# 按条件筛选模板
templates = adapter.list_templates(category="工作总结", min_pages=15, max_pages=30)

# 查看模板详情
detail = adapter.get_template_detail("工作总结_worksummary_01")
```

### 4.3 校验业务数据

```python
import json

with open("business_worksummary.json", "r", encoding="utf-8") as f:
    data = json.load(f)

is_valid, issues = adapter.validate_business_data("工作总结", data)
if not is_valid:
    print("数据问题:", issues)
```

### 4.4 单模板渲染

```python
# 使用分类下第一个模板
adapter.render("工作总结", data, output_path="output.pptx")

# 指定模板
adapter.render("工作总结", data, template_id="工作总结_worksummary_01", output_path="output.pptx")
```

### 4.5 批量渲染

```python
results = adapter.render_batch("工作总结", data, output_dir="output_batch")
for r in results:
    print(f"  {r['template_id']}: {r['status']}")
```

### 4.6 底层渲染引擎（直接调用）

```python
from ppt_renderer import PptRenderer

renderer = PptRenderer("models/工作总结/WorkSummary_01.pptx", "models/工作总结/WorkSummary_01.meta.json")
renderer.render(
    slot_data={"1": {"title": "新标题"}, "4": {"desc": "新内容"}},
    output_path="output.pptx",
    remove_copyright=True,
    auto_fit=True
)
```

---

## 五、业务数据格式

### 5.1 通用结构

所有场景的业务数据遵循统一的三层结构：

```json
{
  "cover": {
    "title": "标题",
    "reporter": "汇报人",
    "period": "时间周期"
  },
  "sections": {
    "section_key_1": [
      {"title": "项目标题1", "desc": "项目描述1"},
      {"title": "项目标题2", "desc": "项目描述2"}
    ],
    "section_key_2": [
      {"title": "项目标题", "desc": "项目描述"}
    ]
  },
  "end": {
    "thanks": "致谢语"
  }
}
```

### 5.2 字段说明

| 层级 | 字段 | 类型 | 说明 |
|------|------|------|------|
| cover | title | string | 封面主标题 |
| cover | reporter | string | 汇报人姓名 |
| cover | period | string | 汇报周期 / 时间 |
| cover | department | string | 所属部门（部分场景） |
| sections | - | object | 章节内容，key 对应 chapter_sections[].key |
| sections.* | title | string | 内容项标题（建议 ≤ 15 字） |
| sections.* | desc | string | 内容项描述（建议 ≤ 60 字） |
| end | thanks | string | 结束页致谢语 |

### 5.3 各场景 section key 对照表

| 场景 | section 1 | section 2 | section 3 | section 4 |
|------|-----------|-----------|-----------|-----------|
| 工作总结 | work_content | project_progress | issues | plan |
| 年终总结 | annual_review | achievements | experience | next_year |
| 工作汇报 | progress | results | challenges | next_steps |
| 工作计划 | objectives | tasks | schedule | resources |
| 述职报告 | duty_performance | achievements | problems | improvement |
| 个人简历 | basic_info | competency | job_awareness | career_plan |
| 自我介绍 | basic_info | student_work | other_works | self_assessment |
| 开题报告 | background | literature | methodology | conclusion |
| 公司简介 | overview | honors | leadership | business（+5个扩展） |
| 职业规划 | self_analysis | career_goal | action_plan | risk_assessment |

> 使用 `schema --scene <场景名>` 命令可查看各场景的完整字段定义。

### 5.4 数据样例

每个场景都有对应的 `business_*.json` 样例文件：

| 场景 | 样例文件 |
|------|----------|
| 工作总结 | business_worksummary.json |
| 年终总结 | business_annual.json |
| 工作汇报 | business_report.json |
| 工作计划 | business_plan.json |
| 述职报告 | business_duty.json |
| 个人简历 | business_resume.json |
| 自我介绍 | business_intro.json |
| 开题报告 | business_thesis.json |
| 公司简介 | business_company.json |
| 职业规划 | business_career.json |

---

## 六、新增模板接入流程

### 6.1 准备模板文件

1. 将 `.pptx` 模板文件放入 `models/<分类名>/` 目录（如 `models/工作总结/`）
2. 命名建议：使用有意义的名称，如 `WorkSummary_03.pptx`

### 6.2 生成元数据

```bash
# 生成单个模板的 meta.json（跳过已存在的）
python ppt_meta_tool.py generate --dir models

# 强制重新生成所有模板的 meta.json
python ppt_meta_tool.py generate --dir models --force
```

### 6.3 校验元数据质量

```bash
python ppt_meta_tool.py check --dir models
```

检查项：
- 必填字段完整性（template_id / category / total_pages / chapters / page_slots）
- 章节数量（< 2 会告警）
- 槽位数量（< 5 会告警）
- 版权页识别

### 6.4 更新模板索引

```bash
python ppt_meta_tool.py index --dir models
```

重新生成 `models/templates_index.json` 总索引文件。

### 6.5 验证新模板

```bash
# 查看模板详情，确认章节结构正确
python ppt_scene_adapter.py detail --template <新模板ID>

# 使用该模板渲染测试
python ppt_scene_adapter.py render --scene <分类名> --data business_<场景>.json --template <新模板ID> --output test.pptx
```

### 6.6 常见问题排查

| 问题 | 原因 | 解决方案 |
|------|------|----------|
| 章节数为 0 或 1 | 章节页未识别 | 检查模板章节页是否含 PART / 第N章 等关键词，必要时调整 `ppt_meta_tool.py` 中 `detect_page_type` 的关键词 |
| 槽位数过少 | 占位文本未识别 | 检查模板占位符是否匹配 `SLOT_MATCH_KEYWORDS`，必要时补充关键词 |
| 替换数为 0 | 模板 ID 不匹配 | 确认 `template_id` 唯一，检查 `templates_index.json` |
| 版权页未删除 | 版权关键词未命中 | 检查末页是否含版权关键词，必要时手动编辑 meta 的 `removable_pages` |
| 文本溢出 | 内容过长 | 启用 `auto_fit=True`（默认开启），或缩短文本 |

---

## 七、渲染引擎参数

### 7.1 PptRenderer.render() 参数

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| slot_data | dict | - | 槽位数据，格式为 `{"页码": {"槽位名": "值"}}`；含 `chart_data`/`table_data` 时自动触发图表/表格替换 |
| output_path | str | - | 输出文件路径 |
| remove_copyright | bool | True | 是否自动删除版权页 |
| auto_fit | bool | True | 是否启用长文本字号自适应 |
| transitions | str/dict/None | None | 全局转场：`auto`/`none`/枚举名/页码字典 |
| animations | str/dict/None | None | 全局动画：`auto`/`none`/枚举名/页码字典 |
| notes_map | dict[int,str] | None | 演讲者备注，键为 page_id（从1开始），如 `{1: "封面备注"}` |
| animation_theme | str | None | 动画主题：`business`/`tech`/`formal`，None 不启用 |

### 7.2 SceneAdapter.render() 参数

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| scene | str | - | 场景名（如"工作总结"） |
| business_data | dict | - | 业务数据（符合场景 Schema） |
| template_id | str | None | 模板 ID（不指定则取该分类第一个） |
| output_path | str | None | 输出路径（不指定则自动生成） |
| auto_fit | bool | True | 是否启用字号自适应 |
| transitions | str/None | None | 全局转场效果 |
| animations | str/None | None | 全局动画效果 |
| animation_theme | str | None | 动画主题 |

### 7.3 动画主题优先级

优先级从高到低：
1. outline.json 单页显式 `transition` / `animations` 字段
2. `--animation-theme` 主题的 `page_overrides`（按 page_type 匹配）
3. `--animation-theme` 主题的 `global_transition`
4. `--transitions` / `--animations` 全局参数

主题未传入时（默认 None）行为与原版完全一致，100% 向后兼容。

---

## 八、测试与验证

### 8.1 单元测试（快速，不渲染 PPTX）

```bash
python -m pytest tests/ -m "not slow" -v
```

覆盖 Schema 校验、校验引擎、图表表格识别等，单次运行 < 1 秒。

### 8.2 全量冒烟测试

```bash
python -m pytest tests/test_smoke_all.py -v
```

对所有模板逐一执行渲染，验证页数正确性和生成成功率。标记为 slow，耗时较长。

### 8.3 边界场景测试

```bash
python -m pytest tests/test_edge_cases.py -v
```

测试长文本溢出、空字段、部分缺失、内容溢出、空字符串、None 值、极短文本等异常场景。

### 8.4 图表表格专项测试

```bash
python -m pytest tests/test_chart_table.py -v
```

动态构造含图表/表格的 PPTX，覆盖 4 类图表替换、多系列适配、表格动态行扩展/缩减/空数据/大数据量。

### 8.5 性能基准测试

```bash
python -m pytest tests/test_performance.py -v
```

统计单份生成耗时、内存占用、批量渲染平均耗时、校验性能，基线 0.10~3.50 秒/100MB。

### 8.6 元数据质量校验

```bash
python ppt_meta_tool.py check --dir models
```

批量校验所有 meta 文件的字段完整性、章节识别、槽位数量、样式规范性。

---

## 九、附录

### 9.1 模板清单

当前模板库共 236 套模板元数据，按 11 类场景分布：

| 分类 | 数量 | 说明 |
|------|------|------|
| 安全教育 | 157 | 安全培训/消防教育等 |
| 工作总结 | 39 | 含商务风系列 |
| 工作汇报 | 12 | 含商务风系列 |
| 年终总结 | 8 | 含商务风系列 |
| 述职报告 | 8 | 含商务风系列 |
| 工作计划 | 5 | 含商务风系列 |
| 公司简介 | 2 | 公司介绍 + 商务风 |
| 开题报告 | 2 | 含商务风 |
| 个人简历 | 1 | 岗位竞聘 |
| 自我介绍 | 1 | 简约学术风研究生复试 |
| 职业规划 | 1 | 职业规划 |

可通过 `python ppt_meta_tool.py index` 查看完整清单，或通过 `SceneAdapter.list_templates(category, style_tag, min_pages, max_pages)` 筛选。

### 9.2 页面类型识别规则（12 类）

| 类型 | 识别规则 |
|------|----------|
| cover | 首页强制识别 |
| catalog | 含 CONTENTS / 目录，文本框数 ≤ 25 |
| divider | 含 PART ONE / 第N章 / Part.01，或数字+短中文标题模式 |
| numbered_list | 序号 + 标题 + 描述组合，序号为纯数字 |
| kpi | 含百分比/数值+标签，数值字号 ≥ 40pt |
| timeline | 含年份 + 事件描述，年份为纯数字 |
| two_column | 左右对称布局，两侧各有标题+列表 |
| skill_percent | 含技能名 + 百分比，百分比为 0-100 |
| preset_titles | 预设标题列表，多个短文本并排 |
| chart | 含 GraphicFrame + chart 对象（柱/折线/饼/雷达） |
| table | 含 GraphicFrame + table 对象 |
| ending | 末 5 页内含 感谢聆听 / THANK YOU 等 |

### 9.3 槽位语义化命名规则

| 槽位名 | 匹配规则 |
|--------|----------|
| title / title_N | 短文本（< 15 字）或含"标题"关键词 |
| desc / desc_N | 含"内容/录入/输入"或长文本（> 20 字） |
| number / number_N | 纯数字（如 01、02） |
| year / year_N | 纯年份（如 2019、2024） |
| percent / percent_N | 百分比（如 67%） |
| reporter | 含"汇报人/姓名" |
| period | 含"年度/202X" |
| chart_data | 图表数据槽位（特殊：触发 `_replace_chart_data`） |
| table_data | 表格数据槽位（特殊：触发 `_fill_dynamic_table`） |

---

## 十、图表与表格数据格式

### 10.1 图表数据格式（chart_data）

outline.json 中 `page_type: chart` 的页面可传入 `chart_data` 字段，渲染引擎自动识别 PPTX 中的图表并替换数据源，100% 保留模板样式（字体/颜色/坐标轴/图例）。

```json
{
  "page_id": 10, "page_type": "chart", "title": "季度营收趋势",
  "chart_type": "bar",
  "chart_data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [
      {"name": "营收", "data": [1200, 1500, 1800, 2100]},
      {"name": "利润", "data": [300, 450, 600, 800]}
    ]
  }
}
```

**支持图表类型**：`bar`（柱状/条形）、`line`（折线）、`pie`（饼图）、`radar`（雷达）

**多系列自动适配**：
- 模板 M 系列，传入 N 系列
- M > N：多余系列清空数据（置零），保留系列结构
- M < N：仅替换前 M 系列，多余数据忽略
- M == N：直接替换

### 10.2 表格数据格式（table_data）

`page_type: table` 的页面传入 `table_data`，渲染引擎自动识别表格并动态扩展行数，继承表头样式与列宽。

```json
{
  "page_id": 11, "page_type": "table", "title": "产品对比",
  "headers": ["产品", "价格", "销量"],
  "rows": [
    ["产品A", "99元", "1.2万"],
    ["产品B", "199元", "0.8万"],
    ["产品C", "299元", "0.6万"]
  ]
}
```

**动态行扩展**：
- 模板 R 行（含表头），传入 N 数据行
- N+1 > R：自动追加行（克隆最后一行保持样式）
- N+1 < R：自动删除多余行（从末尾删除，保留表头）
- 自动行高适配：短内容 0.4 英寸，长内容（>20字符）0.6 英寸

### 10.3 图表/表格测试工具

```bash
# 列出 PPTX 内所有图表/表格
python insert_tables.py list --input template.pptx

# 测试图表替换（输出前后对比报告）
python insert_tables.py test-chart --input template.pptx \
  --data chart_data.json --output out.pptx

# 测试表格扩展
python insert_tables.py test-table --input template.pptx \
  --data table_data.json --output out.pptx
```

---

## 十一、SmartArt 文本替换

模板含 SmartArt 元素（组织架构图/流程图/关系图）时，`ppt_smartart.py` 自动识别并替换文本节点，100% 保留 SmartArt 结构、布局、配色、形状层级。

### 11.1 工作原理

直接操作 SmartArt 的 XML 节点（dgm 命名空间），遍历 `<dgm:pt>` 节点，按 `{old_text: new_text}` 映射替换 `val` 属性，不改动结构与样式。

### 11.2 Python API

```python
from ppt_smartart import replace_smartart_text, list_smartart_text

# 列出幻灯片所有 SmartArt 文本节点（调试用）
texts = list_smartart_text(slide)
print(texts)  # ["节点1文本", "节点2文本", ...]

# 按 {旧文本: 新文本} 替换
replace_smartart_text(slide, {
    "原节点1文本": "新节点1文本",
    "原节点2文本": "新节点2文本",
})
```

### 11.3 使用场景

- 组织架构图：替换部门/职位名称
- 流程图：替换步骤描述
- 关系图：替换关系标签
- 金字塔图：替换层级标题

> SmartArt 替换由渲染引擎在渲染时自动触发，无需额外配置。模板中的 SmartArt 文本会被识别为槽位，业务数据中提供对应新文本即可。

---

## 十二、动画与转场配置

### 12.1 全局配置（CLI 参数）

```bash
python aippt_outline.py step4-generate \
  --outline outline.json \
  --output final.pptx \
  --transitions auto \
  --animations auto \
  --animation-theme business
```

| 参数 | 合法值 | 说明 |
|------|--------|------|
| `--transitions` | `auto` / `none` / 枚举名 | 全局转场效果 |
| `--animations` | `auto` / `none` / 枚举名 | 全局动画效果 |
| `--animation-theme` | `business` / `tech` / `formal` | 动画预设主题 |

### 12.2 单页配置（outline.json）

outline.json 单页对象可添加 `transition` 与 `animations` 字段，优先级高于全局参数。

```json
{
  "page_id": 5, "page_type": "numbered_list", "title": "核心成果",
  "items": [{"subtitle": "增长", "desc": "提升35%"}],
  "transition": "push",
  "animations": {"entry": "fly_in", "by_bullet": true}
}
```

### 12.3 动画预设主题

| 主题 | 风格 | 适用场景 |
|------|------|----------|
| `business` | 简约商务：fade/push 为主 | 工作汇报、年终总结 |
| `tech` | 活力科技：zoom/flip 为主 | 产品发布、技术分享 |
| `formal` | 沉稳正式：fade/wipe 为主 | 政府汇报、学术答辩 |

### 12.4 转场效果枚举（39 种）

**一级推荐**（商务通用）：`fade` / `push` / `wipe` / `dissolve` / `cut` / `morph`

**二级推荐**（适度动感）：`zoom` / `flip` / `conveyor` / `split` / `reveal`

**三级特效**（谨慎使用）：`vortex` / `switch` / `gallery` / `doors` / `window`

> 完整列表：`python ppt_transitions.py list`

### 12.5 动画效果枚举（20+ 种）

**入场**：`fade` / `fly_in` / `zoom` / `wipe` / `slide_in` / `bounce` / `spin`

**退场**：`fade_out` / `fly_out` / `zoom_out` / `slide_out`

**强调**：`pulse` / `spin` / `shake` / `grow_shrink` / `color_blast`

> `by_bullet: true` 仅适用于 `numbered_list` / `catalog` / `timeline` 等多段落页面；`cover` / `divider` / `kpi` / `ending` 禁用。

---

## 十三、六层防御校验体系

### 13.1 校验命令

```bash
# 校验大纲格式
python aippt_outline.py validate --outline outline.json

# 输出结构化 JSON
{
  "validate_pass": true,
  "error_count": 0,
  "warning_count": 2,
  "fixed_count": 1,
  "errors": [],
  "warnings": [...],
  "fixes": ["已自动重排 page_id"]
}
```

### 13.2 六层防御架构

| 层级 | 防御点 | 实现 |
|---|---|---|
| 第一层 | 模型端自检 | SKILL.md 指令引导（铁则/正例/反例/自检清单）|
| 第二层 | JSON Schema 机器化校验 | `schemas/*.schema.json` + `jsonschema` |
| 第三层 | 分步校验流程 | `validate_requirement` / `validate_outline` |
| 第四层 | 模板槽位匹配校验 | `validate_template_match` |
| 第五层 | 运行时兜底自动修复 | `auto_fix_outline` |
| 第六层 | 标准化错误反馈 | 错误码 F0xx/S0xx/T0xx/A0xx |

### 13.3 错误码体系

| 前缀 | 类型 | 示例 |
|------|------|------|
| F0xx | 基础格式错误 | JSON 解析失败、类型不匹配 |
| F1xx | 字段规则错误 | 字段缺失、枚举非法、长度超限 |
| S0xx | 结构逻辑错误 | 页数不匹配、ID 不连续 |
| T0xx | 模板匹配错误 | 槽位不匹配、页面类型不兼容 |
| A0xx | 动画转场错误 | 转场/动画名称非法、by_bullet 类型错 |

### 13.4 自动修复 vs 刚性拦截

| 类型 | 处理方式 | 示例 |
|------|----------|------|
| 可自动修复 | 静默处理 + 日志 | page_id 重排、数组截断、枚举标准化、by_bullet 自动关闭 |
| 刚性拦截 | 返回错误，由模型修正 | 必填字段缺失、枚举值非法、JSON 语法错误 |

---

## 十四、模板自动标注与质量门禁

### 14.1 自动标注命令

```bash
python import_templates.py auto-annotate \
  --input my_template.pptx \
  --scene 工作总结 \
  --output models/工作总结/
```

自动完成：
1. 母版与版式解析（`profile_layouts`）
2. 元素角色识别（`ppt_element_classifier`）
3. 页面模式分类（`_detect_page_pattern`）
4. 生成 `.meta.json` 元数据
5. 生成 2x2 多页缩略图（封面/目录/内容/结尾）
6. 更新 `templates_index.json` 索引
7. 输出质量报告与低置信度 warnings

### 14.2 质量门禁校验

```bash
# 批量校验
python ppt_meta_tool.py check --dir models

# 单模板校验
python ppt_meta_tool.py check --template-id 工作总结_蓝色商务
```

校验项：
- **meta_required**：必填字段完整性（template_id/category/total_pages/chapters/page_slots）
- **rendering_test**：渲染测试（能否正常打开、页数匹配）
- **style_check**：样式检查（主色调数量 ≤ 3、字体一致性）
- **meta_completeness**：元数据完整性（chapters 结构、page_slots 覆盖率）

### 14.3 模板标签体系

`templates_index.json` 为每个模板维护以下标签：

| 字段 | 说明 | 示例 |
|------|------|------|
| `style_tags` | 风格标签 | `["商务", "16:9"]` |
| `color_scheme` | 色系 | `蓝色系` / `红色系` / `绿色系` |
| `industry` | 适用行业 | `["通用"]` 或 `["金融", "教育"]` |
| `page_range` | 页数范围 | `21-25页` |
| `quality_score` | 质量评分 | `80` |

筛选示例：
```python
adapter = SceneAdapter()
# 按场景筛选
templates = adapter.list_templates(category="工作总结")
# 按风格+页数筛选
templates = adapter.list_templates(category="工作总结", style_tag="商务", min_pages=20, max_pages=30)
```

---

## 十五、五步工作流 CLI

AIPPT v2.1 采用五步闭环工作流，每步输出纯 JSON，适配 opencode 模型调用。

```bash
# Step 1: 需求拆解 → 识别场景与参数
python aippt_outline.py step1-understand --text "帮我做一个年终总结PPT"

# Step 2: 大纲生成 → 输出 outline.json（过 Schema 校验）
python aippt_outline.py step2-outline \
  --scene 年终总结 --purpose "2025年度业绩汇报" \
  --audience 管理层 --length 12 --output outline.json

# Step 3: 视觉匹配 → 推荐模板
python aippt_outline.py step3-visuals --outline outline.json

# Step 4: 渲染生成 → 输出 .pptx
python aippt_outline.py step4-generate \
  --outline outline.json \
  --template-id 年终总结_年终总结 \
  --output final.pptx \
  --transitions auto --animations auto \
  --animation-theme business

# 校验（可选，六层防御）
python aippt_outline.py validate --outline outline.json
```

**数据契约**：所有结构化数据（需求参数、大纲、槽位数据）必须输出纯 JSON，严禁 markdown 代码块、解释性文字。详见 [SKILL.md](SKILL.md) 第一章「数据契约」。
