# PPT 生成 Skill 使用手册

## 一、概述

PPT 生成 Skill 是一套基于 python-pptx 的自动化 PPT 生成工具，支持 10 类商务场景（工作总结 / 年终总结 / 工作汇报 / 工作计划 / 述职报告 / 个人简历 / 自我介绍 / 开题报告 / 公司简介 / 职业规划），实现「输入结构化 JSON → 自动匹配模板 → 一键生成 PPT」的全流程自动化。

### 核心能力

- **10 类场景**：预定义 10 类业务场景的 Schema，覆盖常见商务 PPT 需求
- **12 套模板**：内置 12 套商务模板，自动匹配章节结构与槽位
- **样式 100% 还原**：字体、颜色、布局、形状与原模板完全一致
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
├── ppt_meta_tool.py          # 元数据批量处理工具
├── ppt_renderer.py           # 核心渲染引擎
├── ppt_scene_adapter.py      # 场景适配层（Skill 主入口）
├── models/                   # 模板资产库
│   ├── 工作总结/
│   │   ├── WorkSummary_01.pptx
│   │   ├── WorkSummary_01.meta.json
│   │   └── ...
│   ├── 年终总结/
│   ├── ...（10 个分类目录）
│   └── templates_index.json  # 模板总索引
├── business_*.json           # 10 类场景业务数据样例
├── smoke_test_all.py         # 全量冒烟测试
├── edge_test.py              # 边界场景测试
└── verify_all_scenes.py      # 场景校验脚本
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
| slot_data | dict | - | 槽位数据，格式为 `{"页码": {"槽位名": "值"}}` |
| output_path | str | - | 输出文件路径 |
| remove_copyright | bool | True | 是否自动删除版权页 |
| auto_fit | bool | True | 是否启用长文本字号自适应 |

### 7.2 SceneAdapter.render() 参数

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| scene | str | - | 场景名（如"工作总结"） |
| business_data | dict | - | 业务数据（符合场景 Schema） |
| template_id | str | None | 模板 ID（不指定则取该分类第一个） |
| output_path | str | None | 输出路径（不指定则自动生成） |
| auto_fit | bool | True | 是否启用字号自适应 |

---

## 八、测试与验证

### 8.1 全量冒烟测试

```bash
python smoke_test_all.py
```

对 12 套模板逐一执行渲染，验证页数正确性和生成成功率。

### 8.2 边界场景测试

```bash
python edge_test.py
```

测试 7 类异常场景：长文本溢出、空字段、部分缺失、内容溢出、空字符串、None 值、极短文本。

### 8.3 场景校验

```bash
python verify_all_scenes.py
```

校验 10 类场景的输出文件，统计页数、章节、槽位、替换数。

### 8.4 元数据质量校验

```bash
python ppt_meta_tool.py check --dir models
```

批量校验所有 meta 文件的字段完整性、章节识别、槽位数量。

---

## 九、附录

### 9.1 模板清单（12 套）

| 分类 | 模板 | 页数 | 章节数 |
|------|------|------|--------|
| 工作总结 | WorkSummary_01 | 20 | 7 |
| 工作总结 | WorkSummary_02 | 22 | 6 |
| 工作总结 | 工作总结 | 20 | 6 |
| 年终总结 | 年终总结 | 20 | 6 |
| 工作汇报 | 工作汇报 | 20 | 7 |
| 工作计划 | 工作计划 | 23 | 6 |
| 述职报告 | 述职报告 | 20 | 6 |
| 个人简历 | 岗位竞聘 | 21 | 7 |
| 自我介绍 | 简约学术风研究生复试 | 27 | 9 |
| 开题报告 | 开题报告 | 24 | 6 |
| 公司简介 | 公司介绍 | 35 | 12 |
| 职业规划 | 职业规划 | 20 | 6 |

### 9.2 页面类型识别规则

| 类型 | 识别规则 |
|------|----------|
| cover | 首页强制识别 |
| catalog | 含 CONTENTS / 目录，文本框数 ≤ 25 |
| chapter | 含 PART ONE / 第N章 / Part.01，或数字+短中文标题模式 |
| end | 末 5 页内含 感谢聆听 / THANK YOU 等 |
| copyright | 末 3 页含版权关键词，或末页与封面样式重复 |
| content | 以上均不匹配时的默认类型 |

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
