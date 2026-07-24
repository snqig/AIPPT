---
name: aippt-generator
description: Transform user ideas or source documents into presentation-ready PPT through a four-step workflow (Understand → Outline → Visuals → Iterate). Use when user mentions "PPT", "演示文稿", "幻灯片", "汇报材料", or provides documents for presentation generation.
license: MIT
compatibility: opencode
metadata:
  audience: knowledge-workers
  workflow: ai-presentation
  scenes: 工作总结/年终总结/工作汇报/工作计划/述职报告/个人简历/自我介绍/开题报告/公司简介/职业规划
---

# AIPPT Generator Skill

将用户想法或源文档转化为成品 PPT，通过四步智能工作流：

1. **Understand & Parse** — 澄清意图，解析输入
2. **Build Outline** — 生成层级化大纲骨架
3. **Match Visuals** — 匹配场景与模板
4. **Iterate & Refine** — 迭代完善并生成 PPT

## ⚠️ 强制流程约束

**四步必须依次执行，不可跳过、不可乱序。** 每一步都有明确的产出物和用户确认 gate：

| 步骤 | 命令 | 产出物 | 确认 gate |
|---|---|---|---|
| Step 1 | `step1-understand` | 场景识别结果 + 澄清问题清单 | 用户回答澄清问题 |
| Step 2 | `step2-outline` | outline.json（大纲骨架） | 用户审阅大纲逻辑 |
| Step 3 | `step3-visuals` | 视觉匹配建议 + 模板推荐 | 用户确认模板选择 |
| Step 4 | `step4-generate` | 成品 .pptx 文件 | 用户验收 PPT |

**禁止行为**：
- ❌ 跳过 Step 1 直接生成大纲
- ❌ 跳过 Step 2 直接生成 PPT
- ❌ Step 3 未确认模板就进入 Step 4
- ❌ 在用户未确认前进入下一步

**正确做法**：每步执行后，向用户呈现产出物并显式询问"是否确认进入下一步？"，得到肯定答复后才继续。

## 何时使用

**仅当**用户请求涉及以下情况时使用：
- 从零创建演示文稿或 PPT
- 将文档（Word/PDF/markdown）转换为幻灯片
- 为演示生成大纲
- 改进或重构现有演示内容

**不要用于**：
- 与演示无关的通用文档摘要
- 无内容生成需求的纯设计任务

## 工作流

### Step 1: Understand & Parse — "听懂需求"  [`step1-understand`]

**命令**：`python aippt_outline.py step1-understand --text "用户描述"`

**产出**：场景识别结果 + 待澄清问题清单（JSON 输出）

**意图澄清**：生成任何内容前，必须先澄清以下要素：
- **Goal**：用途（学术汇报 / 商业路演 / 项目复盘 / 培训材料 / 公司介绍 / 述职等）
- **Audience**：受众
- **Tone & Style**：正式 / 轻松 / 说服 / 信息型
- **Length**：期望页数
- **Key messages**：2-3 条必须传达的核心信息

**场景识别**：根据用户描述映射到本项目支持的 10 类场景之一：

| 用户描述关键词 | 推荐场景 |
|---|---|
| 工作总结、年度回顾、工作成绩 | `工作总结` |
| 年终、年度汇报、去年回顾 | `年终总结` |
| 工作汇报、项目汇报、阶段汇报 | `工作汇报` |
| 工作计划、年度规划、下步计划 | `工作计划` |
| 述职、履职、岗位汇报 | `述职报告` |
| 竞聘、简历、岗位申请 | `个人简历` |
| 自我介绍、复试、面试 | `自我介绍` |
| 开题、答辩、论文 | `开题报告` |
| 公司介绍、企业简介、公司概况 | `公司简介` |
| 职业规划、职业生涯、发展规划 | `职业规划` |

**触发问**（必要时询问，避免穷举）：
> "这次 PPT 的主要用途是什么？受众是谁？"
> "有偏好的风格吗（学术 / 商务 / 创意）？"
> "大概多少页？"
> "有必须包含的关键信息或数据吗？"

若用户提供文档，解析时需：
- 识别标题层级与段落主题
- 提取关键数据点与统计数字
- 识别章节间逻辑关系

### Step 2: Build Outline — "搭建骨架"  [`step2-outline`]

**命令**：`python aippt_outline.py step2-outline --scene 工作汇报 --purpose "..." --audience "..." --length 14 --keys "关键信息" --output outline.json`

**产出**：outline.json（含 scene/cover/sections/end 的完整大纲）

**确认 gate**：向用户展示大纲结构，询问"大纲逻辑是否需要调整？"，确认后才进入 Step 3

理解意图后，生成结构化大纲。

**大纲结构规则**：
- 首页为标题页：明确主题与副标题
- 逻辑流遵循演示最佳实践：背景 → 现状 → 问题 → 机会 → 结论
- 每页含明确标题与 2-4 条要点
- 层级深度适配内容复杂度

**场景化大纲模板**：不同场景的章节结构对应 `ppt_scene_adapter.py` 中 `SCENE_SCHEMAS` 的 `chapter_sections`：

| 场景 | 章节结构（固定 4-9 章） |
|---|---|
| 工作总结 | 工作内容 / 项目进度 / 问题不足 / 下步计划 |
| 年终总结 | 年度回顾 / 业绩成果 / 经验不足 / 新年规划 |
| 工作汇报 | 工作进展 / 阶段成果 / 困难挑战 / 后续安排 |
| 工作计划 | 工作目标 / 重点任务 / 进度安排 / 资源保障 |
| 述职报告 | 履职情况 / 工作业绩 / 问题不足 / 改进方向 |
| 个人简历 | 个人信息 / 胜任能力 / 岗位认知 / 职业规划 |
| 自我介绍 | 基本信息 / 教育背景 / 实习经历 / 科研 / 获奖 / 自评 / 读研展望 |
| 开题报告 | 选题背景 / 研究现状 / 研究方法 / 论文结论 |
| 公司简介 | 概况 / 荣誉 / 团队 / 业务 / 产品 / 营销 / 市场分析 / 发展 / 未来 |
| 职业规划 | 自我分析 / 职业目标 / 行动计划 / 风险评估 |

**输出格式**：先以可编辑格式呈现给用户确认，再进入下一步。

> **示例大纲**（项目复盘 → 工作汇报场景）：
> ```
> 1. 标题页：Q3 项目复盘报告
> 2. 工作进展：Q3 三大里程碑节点
> 3. 阶段成果：核心指标与交付物
> 4. 困难挑战：风险与应对
> 5. 后续安排：Q4 行动计划
> ```

### Step 3: Match Visuals — "穿上外衣"  [`step3-visuals`]

**命令**：`python aippt_outline.py step3-visuals --outline outline.json`

**产出**：模板候选列表 + 版式匹配建议 + 风格适配决策（JSON 输出）

**确认 gate**：向用户展示模板候选的**首页截图**和视觉建议，询问"选用哪个模板？动画/转场配置？"，确认后才进入 Step 4

大纲确认后，匹配视觉处理。

**模板截图浏览**：每个模板在导入时已通过 PowerPoint COM 生成首页 PNG 截图，路径记录在 `models/preview_manifest.json`。Step 3 输出中会包含每个候选模板的 `screenshot` 字段（相对路径），用户可：

1. **查看截图选择模板**：打开 `models/<分类>/<模板名>.png` 预览模板风格
2. **按截图对比**：将同分类的多个模板截图并排对比，挑选最匹配的视觉风格
3. **参考模板特征**：结合 meta 中的 `total_pages`、`has_chart`、`has_picture` 等字段综合判断

```bash
# 查看某分类下所有模板截图
ls models/工作总结/*.png

# 用默认图片查看器打开某模板截图
start models/工作总结/商务风_001.pptx.png
```

**场景 → 模板匹配**：调用 `ppt_scene_adapter.py` 的 `SceneAdapter.list_templates(category=场景名)` 列出可用模板，按以下规则推荐：

| 内容特征 | 推荐模板特征 |
|---|---|
| 数据密集 | 含图表页的模板（meta 中 `page_meta` 含 `has_chart`） |
| 时间线叙事 | 含 timeline 页面模式的模板 |
| KPI 展示 | 含 KPI 页面模式的模板 |
| 对比论述 | 含双栏页面的模板 |
| 图文并茂 | 含 picture 的模板 |

**模板导入与截图生成**：新模板通过 `import_templates.py` 导入时会自动：
1. 根据首页文本关键词分类到 10 大类目
2. 复制到 `models/<分类>/` 并重命名
3. 生成 meta.json 元数据
4. 用 PowerPoint COM 生成首页 PNG 截图
5. 支持 `--removable-tail N` 标记末尾 N 页为可删除（版权页/致谢页）

```bash
# 导入新模板（自动分类+截图+末尾2页标记删除）
python import_templates.py --src "新模板目录" --prefix 新模板 --removable-tail 2

# 无 PowerPoint 环境时跳过截图
python import_templates.py --src "新模板目录" --prefix 新模板 --no-screenshot
```

**页面模式自动识别**：渲染时 `SceneAdapter._detect_page_pattern` 自动识别 8 类页面模式（divider / numbered_list / timeline / preset_titles / skill_percent / kpi / two_column / chart / table / content），无需手动指定。

**风格建议**：
- 学术汇报 → 简约、数据导向、低饱和度
- 商业路演 → 大胆、品牌化、视觉密集
- 项目复盘 → 时间线或里程碑视觉
- 培训材料 → 步骤化、教学式布局

### Step 4: Iterate & Refine — "共同完善并生成"  [`step4-generate`]

**命令**：`python aippt_outline.py step4-generate --outline outline.json --template-id 工作汇报_工作汇报 --output final.pptx --transitions auto --animations auto`

**产出**：成品 .pptx 文件 + 替换统计 + 残留校验报告

**确认 gate**：向用户呈现 PPT 路径和质量指标，询问"是否需要迭代修改？"

**交互式迭代**：
- 支持"边聊边改"：通过对话修改大纲或内容
- 接受结构、措辞、视觉建议的反馈
- 按用户方向重新生成特定章节

**质量检查**：
- 章节间逻辑流畅
- 关键信息突出
- 长度符合预期
- 每条 desc 控制在 30-60 字（适配模板容量，避免触发缩字号）

**生成 PPT**：大纲确认后，执行以下流程：

```bash
# 1. 大纲 → business_data JSON
python aippt_outline.py --scene 工作汇报 --outline outline.json --output business_custom.json

# 2. business_data → PPT
python ppt_scene_adapter.py generate \
  --scene 工作汇报 \
  --template-id 工作汇报_工作汇报 \
  --data business_custom.json \
  --output final.pptx \
  --transitions auto \
  --animations auto
```

或通过 Python API：

```python
from ppt_scene_adapter import SceneAdapter
from ppt_renderer import PptRenderer

adapter = SceneAdapter("models")
meta, _ = adapter.get_template_meta(template_id="工作汇报_工作汇报")
slot_data = adapter.adapt("工作汇报", business_data, meta)

renderer = PptRenderer("models/工作汇报/工作汇报.pptx",
                       "models/工作汇报/工作汇报.meta.json")
renderer.render(slot_data, "final.pptx",
                remove_copyright=True, auto_fit=True,
                transitions="auto", animations="auto")
```

## business_data JSON 格式

四步工作流最终产出符合 `SCENE_SCHEMAS` 的 JSON：

```json
{
  "cover": {
    "title": "主标题",
    "reporter": "汇报人",
    "period": "汇报周期"
  },
  "sections": {
    "<section_key>": [
      {"title": "要点标题", "desc": "30-60字描述"},
      {"title": "要点标题", "desc": "30-60字描述"}
    ]
  },
  "end": {
    "thanks": "结束致谢语"
  }
}
```

`section_key` 必须对应场景 schema 的 `chapter_sections[].key`（见 Step 2 表格）。

## 不会做的事

- ❌ 编造数据或统计：缺失信息时询问用户
- ❌ 跳过大纲确认直接生成 PPT
- ❌ 在未澄清场景前猜测 section key
- ✅ 每条 desc 控制在 30-60 字（适配模板容量，避免触发缩字号）
- ✅ 生成前用 `SceneAdapter.validate_business_data` 校验

## 输出格式

最终输出包含：
1. **结构化大纲**（slide-by-slide）
2. **视觉风格建议**（按页面类型）
3. **business_data JSON**（符合场景 schema）
4. **成品 PPT 文件路径**
5. **关键信息摘要**（供汇报人参考）

## 示例交互

**用户**："帮我做一个项目复盘的 PPT"

**Agent**：
> 好的，我来帮你生成项目复盘 PPT。先确认几个问题：
> 1. 这个项目是做什么的？
> 2. 复盘的主要受众是谁？
> 3. 期望多少页？
> 4. 有需要重点突出的数据或成果吗？

**用户**："AI 教育产品 Q3 复盘，给管理层看，10 页左右，突出用户增长和营收数据"

**Agent**：识别为 `工作汇报` 场景 → 生成 5 章节大纲（工作进展/阶段成果/困难挑战/后续安排 + 封面）→ 用户确认 → 生成 business_data JSON → 调用渲染引擎产出 PPT

## 项目结构

```
Ppt_work/
├── aippt/                          # 共享核心包（v2.0 新增）
│   ├── __init__.py
│   ├── config.py                   # 集中配置（路径/关键词/默认参数）
│   ├── constants.py                # 共享常量（去重后的占位文本/关键词）
│   └── logger.py                   # 统一日志模块
├── aippt_outline.py                # 四步工作流 CLI + 大纲转换
├── ppt_renderer.py                 # 渲染引擎（模板 → 成品 PPT）
├── ppt_scene_adapter.py            # 场景适配（业务字段 → 模板槽位）
├── ppt_meta_tool.py                # 模板元数据解析工具
├── ppt_animations.py               # 动画效果注入
├── ppt_transitions.py              # 转场效果注入
├── import_templates.py             # 模板批量导入
├── models/                         # 模板库（10 场景分类）
├── pyproject.toml                  # 依赖管理 / 类型检查 / lint 配置
└── SKILL.md
```

## 与本项目的对接

| Skill 步骤 | 项目模块 | 输入 | 输出 |
|---|---|---|---|
| Step 1 Understand | `SceneAdapter.list_scenes()` | 用户描述 | 场景名 |
| Step 2 Outline | `SCENE_SCHEMAS[scene].chapter_sections` | 场景名 | 大纲 JSON |
| Step 3 Visuals | `SceneAdapter.list_templates(category=scene)` | 场景名 | 模板 ID |
| Step 4 Generate | `aippt_outline.py` + `PptRenderer.render()` | 大纲 + 模板 | .pptx 文件 |

## 依赖

- Python 3.10+
- python-pptx >=1.0.2
- lxml（动画/转场注入）
- 可选: pywin32（截图功能）

```bash
# 安装核心依赖
pip install -e .

# 安装全部依赖（含截图、开发工具）
pip install -e ".[all]"
```
