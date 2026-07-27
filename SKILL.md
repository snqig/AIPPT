---
name: aippt-generator
description: Transform user ideas or source documents into presentation-ready PPT through a five-step closed-loop workflow (Understand → Outline → Expand → Match → Generate). Use when user mentions "PPT", "演示文稿", "幻灯片", "汇报材料", "安全教育", or provides documents for presentation generation.
license: MIT
compatibility: opencode
metadata:
  audience: knowledge-workers
  workflow: ai-presentation
  scenes: 工作总结/年终总结/工作汇报/工作计划/述职报告/个人简历/自我介绍/开题报告/公司简介/职业规划/安全教育
---

# AIPPT Generator Skill

依托 AIPPT 渲染引擎，提供模板保真式 PPT 一键生成能力。采用「槽位替换」架构，100% 保留模板字体、配色与版式，支持 11 类商务场景、38 种转场效果、20+ 动画效果，可将用户自然语言需求转化为可下载的成品 .pptx 文件。

**核心分工**：宿主大模型负责需求理解、内容创作、大纲编排、参数组装与流程调度；AIPPT 引擎负责模板管理、槽位渲染、效果注入与文件输出。

---

## 一、数据契约（铁则，违反则渲染必然失败）

> ⚠️ **重要提示**：你输出的 JSON 将直接作为渲染命令的输入，程序不会做语义理解，只会严格按字段名读取。任何多余文字、格式偏差都会直接导致程序报错，无法生成 PPT。

### 1.1 格式铁则（必须严格遵守，不可商量）

1. **纯输出原则**：所有结构化数据（需求参数、大纲、槽位数据）必须输出**纯 JSON 文本**，严禁包裹 markdown 代码块，严禁前后添加解释性文字、备注、说明。
2. **字段零自创原则**：所有字段名必须与本规范完全一致，拼写、大小写、下划线严格对齐，禁止自创字段、禁止用近义词替代。
3. **枚举封闭原则**：`scene`、`page_type` 等枚举字段，必须从给定列表中取值，禁止中文别名、大小写变体、自定义新值。
4. **结构匹配原则**：数组嵌套层级、对象结构必须与示例完全一致，同数组内对象结构必须统一。
5. **真实渲染原则**：所有 PPT 必须通过调用 AIPPT 命令行工具生成真实 .pptx 文件，禁止以文本描述、Markdown 模拟、代码块伪代码等方式代替成品文件交付。
6. **样式保真原则**：仅修改模板槽位内的文本、图片、数据内容，不得要求或尝试修改模板本身的版式、字体、配色与布局。
7. **动画转场枚举封闭原则**：所有转场（`transition`）、动画（`animations.entry/exit/emphasis`）效果名称必须从官方枚举列表中选择，禁止中文别名、自定义名称、大小写变体。`animations` 必须为对象结构，`by_bullet` 必须为布尔值。无特殊需求时全局统一使用 `auto`，由渲染引擎按页面类型自动匹配，禁止逐页随机分配效果。

### 1.2 优先级与后果关联

**格式合规 > 内容详实**。宁可精简内容保证格式正确，也不能为了内容丰富破坏格式。

| 行为 | 后果 |
|---|---|
| 格式错误 | 渲染命令执行失败 = 用户拿不到成品 PPT |
| 内容略简陋但格式正确 | 可正常生成 = 用户可获得可用成品 |

**分工边界**：你只负责产出符合契约的结构化数据，渲染引擎只会按槽位机械替换，不会自动修正你的格式错误，也不会补全缺失字段。格式有任何偏差，都会直接失败。

### 1.3 术语 100% 统一

整个 SKILL 文档内，所有字段名、页面类型、参数名必须和代码、JSON Schema 完全同名，杜绝任何别名、俗称、简称。

- ✅ 全程统一使用 `kpi` 作为页面类型名称，所有地方提及都写 `page_type: kpi`
- ❌ 一会写 "KPI 页"，一会写 "指标卡片页"，一会写 `kpi_card`

---

## 二、完整正例（可直接运行的基准参照）

以下是完整的、可直接通过校验、可直接渲染的 outline.json 示例，作为模型的基准参照。生成大纲时必须严格对照此结构。

```json
{
  "scene": "年终总结",
  "purpose": "2025年度业绩复盘与2026规划汇报",
  "audience": "部门管理层",
  "total_pages": 6,
  "pages": [
    {"page_id": 1, "page_type": "cover", "title": "2025年度工作总结", "subtitle": "复盘·前行"},
    {"page_id": 2, "page_type": "catalog", "title": "目录", "items": ["业绩概览", "工作成果", "明年规划"]},
    {"page_id": 3, "page_type": "divider", "section_no": "01", "title": "业绩概览"},
    {"page_id": 4, "page_type": "kpi", "title": "核心指标", "kpi_items": [
      {"label": "用户量", "value": "128万", "trend": "+35%"},
      {"label": "营收", "value": "8600万", "trend": "+28%"},
      {"label": "留存率", "value": "72%", "trend": "+5pct"}
    ],
     "animations": {"entry": "fade", "emphasis": "pulse", "by_bullet": false}},
    {"page_id": 5, "page_type": "numbered_list", "title": "核心成果", "items": [
      {"subtitle": "用户增长", "desc": "新增用户同比提升35%"}
    ],
     "transition": "push",
     "animations": {"entry": "fly_in", "by_bullet": true}},
    {"page_id": 6, "page_type": "ending", "title": "感谢观看", "subtitle": "欢迎指正"}
  ]
}
```

### 2.1 动画与转场字段规范（单页覆盖，可选）

大纲单页对象可添加 `transition` 与 `animations` 字段，**仅对当前页生效，优先级高于全局 `--transitions` / `--animations`**。无特殊需求时省略该字段，继承全局 `auto`。

| 字段 | 类型 | 必填 | 合法值 | 说明 |
|---|---|---|---|---|
| `transition` | string | 否 | 转场枚举列表 / `none` | 单页转场效果，不填则继承全局 |
| `animations.entry` | string | 否 | 入场动画枚举 / `null` | 元素入场动画 |
| `animations.exit` | string | 否 | 退场动画枚举 / `null` | 元素退场动画，商务场景不推荐 |
| `animations.emphasis` | string | 否 | 强调动画枚举 / `null` | 仅重点数据页使用，每页不超过 1 处 |
| `animations.by_bullet` | boolean | 否 | `true` / `false` | 是否按段落逐条播放，默认 `false` |

**使用约束**：
- `animations` 必须为对象结构，禁止直接写字符串（如 `"animations": "fly_in"` ❌）
- `by_bullet` 必须为布尔值，禁止写字符串（如 `"by_bullet": "true"` ❌）
- `by_bullet: true` 仅适用于 `numbered_list` / `catalog` / `timeline` / `preset_titles` 等多段落页面；`cover` / `divider` / `kpi` / `ending` 禁用
- 商务场景优先使用 `fade` / `push` / `wipe` 等简约效果，禁用 `vortex` / `fling` 等炫技特效

---

## 三、高频错误对照表（精准命中常见错误）

| 错误类型 | 错误写法 | 正确写法 | 错误原因 |
|---|---|---|---|
| 页面类型用中文 | `"page_type": "KPI指标页"` | `"page_type": "kpi"` | 枚举值必须用规定英文，程序无法识别中文 |
| 页面类型大小写错 | `"page_type": "KPI"` | `"page_type": "kpi"` | 枚举值全小写，大小写敏感 |
| 字段名自创 | `"kpi_list": [...]` | `"kpi_items": [...]` | 字段名必须严格对齐，程序只读 `kpi_items` |
| 多余包裹代码块 | ```json {...}``` | 直接输出 `{...}` | 程序直接解析文本，代码块标记会导致 JSON 解析失败 |
| 前后加解释文字 | `以下是生成的大纲：{...}` | 直接输出 `{...}` | 多余文字会导致 JSON 解析失败 |
| 同数组结构不一致 | `[{"label":"xx"}, {"name":"yy"}]` | `[{"label":"xx"}, {"label":"yy"}]` | 同数组内对象字段必须统一 |
| page_id 从 0 开始 | `"page_id": 0` | `"page_id": 1` | 页码必须从 1 开始连续递增 |
| page_id 不连续 | `1, 2, 5, 6` | `1, 2, 3, 4` | 页码必须连续递增 |
| total_pages 不匹配 | `total_pages: 5` 但 pages 有 6 个 | `total_pages: 6` | total_pages 必须等于 pages 数组长度 |
| 必填字段缺失 | KPI 页缺 `kpi_items` | 补全 `kpi_items` | 每种页面类型有必填字段，不可省略 |
| 转场用中文 | `"transition": "淡入效果"` | `"transition": "fade"` | 必须使用英文枚举值，程序无法识别中文 |
| 动画字段层级错 | `"animations": "fly_in"` | `"animations": {"entry": "fly_in"}` | `animations` 必须是对象，不能直接写字符串 |
| by_bullet 类型错 | `"by_bullet": "true"` | `"by_bullet": true` | 必须是布尔值，字符串会导致解析失败 |
| 自定义效果名 | `"transition": "smooth_fade"` | `"transition": "fade"` | 枚举值封闭，禁止自定义新效果 |
| 大小写错误 | `"transition": "Fade"` | `"transition": "fade"` | 枚举值严格小写，大小写敏感 |
| 封面用逐段动画 | cover 页 `"by_bullet": true` | `"by_bullet": false` 或省略 | `by_bullet` 仅适用于列表/目录/时间轴页 |

---

## 四、枚举值定义（封闭列表，禁止自创）

### 4.1 场景枚举（共 11 类）

`工作总结` / `年终总结` / `工作汇报` / `工作计划` / `述职报告` / `个人简历` / `自我介绍` / `开题报告` / `公司简介` / `职业规划` / `安全教育`

**场景关键词映射**：

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
| 安全教育、安全培训、安全生产、消防教育 | `安全教育` |

### 4.2 页面类型枚举（共 12 类）

`cover` / `catalog` / `divider` / `numbered_list` / `kpi` / `timeline` / `two_column` / `skill_percent` / `preset_titles` / `chart` / `table` / `ending`

### 4.3 转场效果枚举（38 种，按推荐度分级）

**一级推荐**（商务通用，`auto` 模式优先选用）：
`fade`（淡入） / `push`（推进） / `wipe`（擦除） / `dissolve`（溶解） / `cut`（切出）

**二级推荐**（适度动感）：
`zoom`（缩放） / `flip`（翻转） / `conveyor`（传送带） / `split`（分割） / `reveal`（揭开） / `random`（随机）

**三级特效**（谨慎使用，仅创意场景）：
`vortex`（漩涡） / `switch`（切换） / `fling`（抛出） / `gallery`（画廊） / `cube`（立方体） / `doors`（开门） / `window`（开窗）

> 完整 38 种效果可通过 `python ppt_transitions.py list` 查询。商务汇报类场景禁止使用三级特效。

### 4.4 元素动画枚举（20+ 种）

**入场动画**（`animations.entry`，最常用）：
`fade` / `fly_in` / `zoom` / `wipe` / `slide_in` / `bounce` / `spin`

**退场动画**（`animations.exit`，商务场景不推荐）：
`fade_out` / `fly_out` / `zoom_out` / `slide_out`

**强调动画**（`animations.emphasis`，仅重点数据页使用）：
`pulse` / `spin` / `shake` / `grow_shrink` / `color_blast`

> 全局命令行参数 `--transitions` / `--animations` 合法值：`auto` / `none` / 对应枚举列表中的具体名称。

---

## 五、12 类页面类型最小示例（生成时直接对照填空）

每种 page_type 的标准结构，生成对应页面时必须严格对照。

### 5.1 cover（封面页）
```json
{"page_id": 1, "page_type": "cover", "title": "2025年度工作总结", "subtitle": "复盘·前行"}
```

### 5.2 catalog（目录页）
```json
{"page_id": 2, "page_type": "catalog", "title": "目录", "items": ["业绩概览", "工作成果", "明年规划"]}
```

### 5.3 divider（章节分隔页）
```json
{"page_id": 3, "page_type": "divider", "section_no": "01", "title": "业绩概览"}
```

### 5.4 numbered_list（数字列表页）
```json
{
  "page_id": 4, "page_type": "numbered_list", "title": "核心工作成果",
  "items": [
    {"subtitle": "用户增长体系搭建", "desc": "完成全链路获客体系建设，新增用户同比提升35%"},
    {"subtitle": "营收结构优化", "desc": "增值服务占比提升至42%，盈利能力显著增强"}
  ]
}
```

### 5.5 kpi（KPI 卡片页）
```json
{
  "page_id": 5, "page_type": "kpi", "title": "核心指标",
  "kpi_items": [
    {"label": "用户量", "value": "128万", "trend": "+35%"},
    {"label": "营收", "value": "8600万", "trend": "+28%"}
  ]
}
```

### 5.6 timeline（时间轴页）
```json
{
  "page_id": 6, "page_type": "timeline", "title": "年度里程碑",
  "timeline_items": [
    {"time": "2025年3月", "event": "核心产品上线，首月用户破10万"},
    {"time": "2025年7月", "event": "完成B轮融资，金额5000万美元"}
  ]
}
```

### 5.7 two_column（双栏对比页）
```json
{
  "page_id": 7, "page_type": "two_column", "title": "竞品对比",
  "left_title": "我方优势", "left_items": ["技术领先", "成本可控"],
  "right_title": "竞品现状", "right_items": ["技术滞后", "成本高昂"]
}
```

### 5.8 skill_percent（技能百分比页）
```json
{
  "page_id": 8, "page_type": "skill_percent", "title": "能力雷达",
  "skills": [
    {"name": "数据分析", "percent": 90},
    {"name": "项目管理", "percent": 85}
  ]
}
```

### 5.9 preset_titles（预设标题列表页）
```json
{"page_id": 9, "page_type": "preset_titles", "title": "核心模块", "items": ["用户增长", "营收提升", "团队建设"]}
```

### 5.10 chart（图表页）
```json
{
  "page_id": 10, "page_type": "chart", "title": "季度营收趋势",
  "chart_type": "bar",
  "chart_data": {"categories": ["Q1","Q2","Q3","Q4"], "series": [{"name":"营收","data":[1200,1500,1800,2100]}]}
}
```

### 5.11 table（表格页）
```json
{
  "page_id": 11, "page_type": "table", "title": "产品对比",
  "headers": ["产品", "价格", "销量"],
  "rows": [["产品A", "99元", "1.2万"], ["产品B", "199元", "0.8万"]]
}
```

### 5.12 ending（结尾页）
```json
{"page_id": 12, "page_type": "ending", "title": "感谢观看", "subtitle": "欢迎指正"}
```

---

## 六、输出前强制自检清单（10 项）

在每个输出节点（需求参数、大纲、槽位数据）输出前，必须逐项完成以下自检，全部通过后方可输出：

- ✅ 所有 `page_type` 都在枚举列表内，无中文、无大小写错误
- ✅ 所有必填字段都存在，无缺失、无拼写错误
- ✅ `total_pages` 数值与 `pages` 数组长度完全相等
- ✅ 所有数组内对象结构统一，字段名一致
- ✅ 纯 JSON 输出，无代码块、无多余文字、无解释说明
- ✅ 所有 `transition`、`animations.entry/exit/emphasis` 名称均来自官方枚举列表，无中文、无自定义、无大小写错误
- ✅ `animations` 为对象结构，包含 `entry` / `exit` / `emphasis` / `by_bullet` 子字段，未直接写字符串
- ✅ `by_bullet` 为布尔值 `true` / `false`，而非字符串
- ✅ 仅 `numbered_list` / `catalog` / `timeline` / `preset_titles` 等多段落页面开启 `by_bullet`，`cover` / `divider` / `kpi` / `ending` 未开启
- ✅ 无特殊需求时优先使用全局 `auto`，未随意自定义每页效果

---

## 七、六层防御体系总览

模型生成的内容会经过六道关卡，格式合规率接近 100%：

| 层级 | 防御点 | 实现方式 |
|---|---|---|
| 第一层 | 模型端自检 | SKILL 指令引导（第一至六章） |
| 第二层 | JSON Schema 机器化校验 | `schemas/*.schema.json` + `jsonschema` 库 |
| 第三层 | 分步校验流程 | `aippt/validators.py` 的 `validate_requirement` / `validate_outline` |
| 第四层 | 模板槽位匹配校验 | `validate_template_match`（结合 meta.json） |
| 第五层 | 运行时兜底自动修复 | `auto_fix_outline`（page_id重排/截断/枚举标准化/动画字段降级） |
| 第六层 | 标准化错误反馈 | 错误码体系（F0xx/S0xx/T0xx/A0xx）+ 修正建议 |

### 7.1 错误码分类

| 错误码前缀 | 错误类型 | 示例 |
|---|---|---|
| F0xx | 基础格式错误 | JSON 解析失败、类型不匹配 |
| F1xx | 字段规则错误 | 字段缺失、枚举非法、长度超限 |
| S0xx | 结构逻辑错误 | 页数不匹配、ID 不连续 |
| T0xx | 模板匹配错误 | 槽位不匹配、页面类型不兼容 |
| A0xx | 动画转场错误 | 转场/动画名称非法、`by_bullet` 类型错、字段层级错 |

### 7.2 可自动修复 vs 刚性拦截

| 类型 | 处理方式 | 示例 |
|---|---|---|
| 可自动修复 | 静默处理 + 日志记录 | page_id 重排、数组截断、枚举大小写标准化、移除空值字段、`by_bullet` 自动关闭、动画字段降级为 `auto` |
| 刚性拦截 | 返回错误，由模型修正 | 必填字段缺失、枚举值非法、页数严重偏差、JSON 语法错误、转场/动画名称非法、`animations` 类型错 |

**动画转场错误码明细**：

| 错误码 | 级别 | 错误说明 | 修正建议 |
|---|---|---|---|
| A001 | error | 转场效果名称非法 | 请从枚举列表选择：`fade`, `push`, `wipe`, `dissolve`, `zoom`, `flip`, `cut`, `split`, `reveal`, `random`, `vortex`, `switch`, `fling`, `gallery`, `cube`, `doors`, `window`, `conveyor` |
| A002 | error | 动画效果名称非法 | 入场动画合法值：`fade`, `fly_in`, `zoom`, `wipe`, `slide_in`, `bounce`, `spin`；退场：`fade_out`, `fly_out`, `zoom_out`, `slide_out`；强调：`pulse`, `spin`, `shake`, `grow_shrink`, `color_blast` |
| A003 | error | `animations` 字段类型错误 | `animations` 必须为对象结构，如 `{"entry": "fly_in", "by_bullet": true}`，不能直接写字符串 |
| A004 | error | `by_bullet` 类型错误 | `by_bullet` 必须是布尔值 `true` / `false`，不能写字符串 `"true"` |
| A005 | warning | 页面类型不支持逐段动画 | 当前页面类型（如 `cover` / `divider` / `kpi` / `ending`）不建议开启 `by_bullet`，已自动关闭 |
| A006 | warning | 不推荐使用高动态特效 | 商务场景建议使用 `fade` / `push` / `wipe` 等简约转场，避免 `vortex` / `fling` 等炫技特效 |

### 7.3 错误反馈三要素（修正精准可落地）

校验失败时，返回结构化错误信息，包含三个核心信息，模型可直接定位修改：

1. **精准定位**：用 JSON Path 指明错误位置，例如 `pages[2].kpi_items[1].label`
2. **明确原因**：说明违反了哪条规则，例如 "字段缺失：缺少必填字段 label"
3. **修正指引**：给出正确写法或可选值，例如 "可选枚举值：cover, catalog, divider, kpi..."

**错误反馈示例**：
```json
{
  "validate_pass": false,
  "errors": [
    {
      "code": "F102",
      "level": "error",
      "path": "pages[2].page_type",
      "message": "页面类型枚举值非法: 数据指标页",
      "suggestion": "请从以下列表选择: cover, catalog, divider, numbered_list, kpi, timeline, two_column, skill_percent, preset_titles, chart, table, ending"
    }
  ]
}
```

### 7.4 修正原则：最小改动

格式错误修正时必须遵循：

- 只修改报错位置的字段，不要重写整个大纲
- 优先从枚举列表、示例中复制正确写法，不要自己发明新写法
- 修正后必须重新执行第六章自检清单，确认无误再重试
- 同一错误连续 2 次修正失败，主动向用户说明问题，请求确认需求

---

## 八、标准工作流（五步闭环）

### 8.1 步骤一：需求拆解与参数提取

**输入**：用户自然语言描述
**执行动作**：从用户描述中提取生成参数，输出标准化需求参数字典

**命令**：`python aippt_outline.py step1-understand --text "用户描述"`

**输出格式**：
```json
{
  "scene": "年终总结",
  "audience": "部门管理层",
  "purpose": "2025年度业绩复盘与2026规划汇报",
  "page_count": 12,
  "tone": "正式专业",
  "focus_points": ["业绩达成", "问题复盘", "明年规划"]
}
```

**约束规则**：
- `scene` 必须从「4.1 场景枚举」中选择，不匹配时自动归类到最接近的场景
- `page_count` 必须为正整数，默认值：工作总结类 10-12 页，简历类 4-6 页
- `focus_points` 提取用户明确强调的核心内容，无则留空数组

**自检（步骤一后）**：校验场景枚举、页数范围、字段完整性

### 8.2 步骤二：结构化大纲生成（核心）

**输入**：步骤一生成的需求参数字典
**执行动作**：基于场景与受众，创作符合商务逻辑的完整 PPT 大纲，输出标准 outline.json

**命令**：`python aippt_outline.py step2-outline --scene 年终总结 --purpose "..." --audience "..." --length 12 --keys "关键信息" --output outline.json`

**输出要求**：页数严格等于 `page_count`，必须包含 1 页封面、1 页目录、1 页结尾页，其余为内容页；每页必须指定合法 `page_type`。

**约束规则**：
- 每页字段必须与「第五章 12 类页面类型最小示例」完全对应，不得缺省必填字段
- `numbered_list` 每页要点控制在 3-5 个；`kpi` 每页 3-4 个指标；`timeline` 每页 4-6 个节点
- 内容逻辑符合对应场景的专业汇报结构，层级清晰，要点具象化
- 每条 desc 控制在 30-60 字（适配模板容量，避免触发缩字号）
- title 控制在 15 字以内

**确认 gate**：向用户展示大纲结构，询问"大纲内容和逻辑是否需要调整？"，确认后才进入 Step 3

**自检（步骤二后）**：校验页面类型、必填字段、数组长度、整体结构

### 8.3 步骤三：内容扩写与槽位适配（可选）

**适用场景**：用户要求内容详实、或模板槽位需要完整正文时执行
**执行动作**：将大纲要点扩写为完整正文段落，确保内容长度适配模板槽位容量
**输出**：更新后的 outline.json，将短要点替换为完整表述

**约束规则**：
- 正文内容控制在合理长度，优先保证可读性，避免过度堆砌
- 自动适配槽位数量，内容不足则补充维度，超出则精简提炼
- 数据类内容优先使用「结论 + 数据 + 对比」的结构化表达

### 8.4 步骤四：模板智能匹配

**命令**：`python aippt_outline.py step3-visuals --outline outline.json`

**执行动作**：调用模板列表查询命令，获取对应场景下的全部可用模板，基于内容风格自动选择最优模板

**模板截图浏览**：每个模板在导入时已通过 PowerPoint COM 生成 2x2 多页缩略图（封面/目录/内容/结尾），路径记录在 `models/preview_manifest.json`。

**匹配优先级**：
1. 场景完全匹配
2. 页数与大纲页数接近
3. 风格与受众适配（正式商务、简约科技、清新活力等）

**确认 gate**：向用户展示模板候选的**首页截图**和视觉建议，询问"选用哪个模板？动画/转场配置？"，确认后才进入 Step 4

**自检（步骤三后）**：校验模板 ID 格式、场景匹配性

### 8.5 步骤五：渲染生成与结果交付

**命令**：`python aippt_outline.py step4-generate --outline outline.json --template-id 模板ID --output final.pptx --transitions auto --animations auto`

**渲染前必做参数二次校验**：
- ✅ `--outline` 指向的文件是合法 JSON
- ✅ `--template-id` 是从模板列表中获取的真实 ID
- ✅ 所有参数名拼写正确，无多余空格与特殊字符

**默认参数**：
- `--transitions auto`：自动匹配页面转场效果
- `--animations auto`：自动匹配入场 / 强调动画
- `remove_copyright=True`：自动清理模板版权页
- `auto_fit=True`：开启超长文本字号自适应

**参数合法取值表**：

| 参数 | 合法值 | 说明 |
|---|---|---|
| `--transitions` | `auto` / `none` / 具体转场名称（如 `fade`、`push`） | 全局转场效果，默认 `auto`；具体名称须来自 4.3 转场枚举 |
| `--animations` | `auto` / `none` / 具体动画名称（如 `fade`、`fly_in`） | 全局入场动画，默认 `auto`；具体名称须来自 4.4 动画枚举 |

**单页覆盖优先级**：outline.json 单页对象的 `transition` / `animations` 字段优先于全局 `--transitions` / `--animations`；未指定单页字段时继承全局配置。

**可选参数**：
- `--trim-pages 6,10,11`：裁剪指定页面
- `--insert-tables`：自动插入表格页（对比表+报价表）

**内置渲染前终检**：step4-generate 内置六层防御体系校验（Layer 3-5），格式错误会被自动拦截，不会执行渲染。

**交付内容**：
1. 成品 PPT 文件路径
2. 基本信息：页数、使用模板、生成耗时
3. 核心内容摘要
4. 残留占位文本校验报告

**确认 gate**：向用户呈现 PPT 路径和质量指标，询问"是否需要迭代修改？"

---

## 九、各场景章节结构

对应 `ppt_scene_adapter.py` 中 `SCENE_SCHEMAS` 的 `chapter_sections`：

| 场景 | 章节结构 |
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
| 安全教育 | 安全概述 / 危险源识别 / 安全措施 / 安全培训 / 应急预案 |

---

## 十、可用工具命令全集

### 10.1 核心生成命令

**1. 全链路一键生成**
```bash
python aippt_outline.py auto-generate \
  --prompt "用户原始需求文本" \
  --output 输出文件路径.pptx \
  [--template-id 模板ID] \
  [--transitions auto] \
  [--animations auto]
```
说明：自动完成需求理解、大纲生成、模板匹配、渲染全流程；未指定模板时自动选择最优模板。生成的是骨架 PPT，建议用 step2-outline 填充真实内容后用 step4-generate 重渲染。

**2. 基于大纲渲染生成**（最常用）
```bash
python aippt_outline.py step4-generate \
  --outline 大纲文件路径.json \
  --template-id 模板ID \
  --output 输出文件路径.pptx \
  --transitions auto \
  --animations auto \
  [--trim-pages 6,10,11] \
  [--insert-tables]
```
说明：输入标准 outline.json，输出成品 PPT，并附残留占位文本校验报告。step4-generate 内置渲染前终检（六层防御体系 Layer 3-5），格式错误会被自动拦截。

**3. 四步工作流独立子命令**
```bash
# Step 1: 理解与拆解
python aippt_outline.py step1-understand --text "用户描述"

# Step 2: 构建大纲
python aippt_outline.py step2-outline --scene 工作汇报 --purpose "..." --audience "..." --length 14 --keys "关键信息" --output outline.json

# Step 3: 视觉匹配
python aippt_outline.py step3-visuals --outline outline.json

# Step 4: 渲染生成
python aippt_outline.py step4-generate --outline outline.json --template-id 模板ID --output final.pptx
```

### 10.2 格式校验命令（六层防御体系）

**1. 大纲格式校验**
```bash
python aippt_outline.py validate \
  --outline outline.json \
  [--template-id 模板ID] \
  [--auto-fix] \
  [--output result.json]
```
说明：执行六层防御体系校验。`--template-id` 启用模板槽位匹配校验（Layer 4）；`--auto-fix` 自动修复非原则性问题并回写文件（Layer 5）；输出标准化校验结果（Layer 6）。

**2. 分步校验时机**

| 校验时机 | 校验内容 | 命令 |
|---|---|---|
| Step 1 后 | 场景枚举、页数范围、核心参数 | `validate --outline` (需求参数) |
| Step 2 后 | 大纲结构、页面类型、字段完整性 | `validate --outline outline.json` |
| Step 3 后 | 模板槽位匹配、页面类型兼容性 | `validate --outline outline.json --template-id 模板ID` |
| Step 4 前 | 渲染前终检（自动执行） | step4-generate 内置，无需手动调用 |

### 10.3 模板管理命令

**1. 查询模板列表**
```bash
python aippt_outline.py list-templates [--scene 场景名] [--style 风格] [--min-pages N] [--max-pages N]
```
输出：结构化 JSON 列表，包含 template_id、场景、风格标签、页数。

**2. 查看模板元数据**
```bash
python ppt_meta_tool.py info --template-id 模板ID
```
输出：模板详细信息，包括每页类型、槽位数量、可删除页、章节结构。

**3. 模板质量校验（单模板）**
```bash
python ppt_meta_tool.py check --template-id 模板ID
```
输出：槽位完整性、章节置信度、渲染兼容性校验结果。

**4. 模板批量校验**
```bash
python ppt_meta_tool.py check --dir models
```
输出：全量 meta 质量校验报告。

### 10.4 模板接入与校验命令

**1. 批量导入模板**
```bash
python import_templates.py import --src "源目录" --prefix 模板前缀 --removable-tail 2 [--force] [--no-screenshot]
```
说明：自动分类+复制+生成 meta+2x2 多页缩略图+更新索引。`--removable-tail N` 标记末尾 N 页为可删除。

**2. 自定义模板自动标注**
```bash
python import_templates.py auto-annotate --input 模板文件路径.pptx --scene 所属场景 --output 元数据输出目录
```
说明：自动解析模板结构、识别槽位、分类页面类型，生成标准 .meta.json 元数据文件。

**3. 全量模板索引更新**
```bash
python import_templates.py rebuild-index [--models-dir models]
```
说明：重新生成 templates_index.json 全局模板索引。

### 10.5 辅助工具命令

```bash
# 裁剪指定页面
python trim_ppt.py --input input.pptx --output output.pptx --pages 1,3,5-8

# 批量插入表格
python insert_tables.py --template template.pptx --data data.json --output output.pptx

# 多页缩略图生成（2x2 网格）
python generate_thumbnails.py --models-dir models --layout 2x2 [--force]
```

---

## 十一、错误处理与降级策略

| 错误场景 | 处理方式 |
|---|---|
| 模板不存在 | 列出对应场景下所有可用模板，引导用户选择 |
| 页数不匹配 | 自动调整大纲页数，增删过渡页或内容页，优先保证结构完整 |
| 槽位数量不匹配 | 自动精简 / 补充内容要点，确保每页要点数与模板槽位数一致 |
| 渲染执行失败 | 先调用校验命令排查问题，根据错误提示修正数据后重试，最多重试 2 次 |
| 场景无法匹配 | 自动归类到最相近的通用场景，同时告知用户适配情况 |
| 内容溢出文本框 | 开启 auto_fit 自动缩小字号；严重溢出时自动精简文案 |

---

## 十二、最佳实践

1. **内容密度控制**：每页正文要点不超过 5 条，单条描述不超过 2 行，兼顾可读性与美观度
2. **页面类型搭配**：合理混合使用列表、KPI、时间轴、双栏等多种页面类型，避免版式单调
3. **动画默认自动**：无特殊要求时统一使用 `auto` 模式，渲染引擎会根据页面类型匹配最优效果
4. **数据具象化**：优先使用「数字 + 比例 + 对比」的表达，避免空泛描述
5. **交付简洁化**：生成完成后直接给出文件路径与核心信息，无需冗余铺垫
6. **四步确认 gate**：每步执行后向用户呈现产出物并显式询问"是否确认进入下一步？"
7. **格式优先**：格式合规 > 内容详实，宁可精简内容保证格式正确
8. **最小改动修正**：校验失败时只改报错字段，不重写整个大纲
9. **动画默认 auto**：90% 以上场景直接使用 `--transitions auto --animations auto`，渲染引擎按页面类型自动匹配专业效果，无需手动定制
10. **转场风格统一**：同一套 PPT 转场风格不超过 2 种，以 `fade` / `push` 为主，避免花哨特效喧宾夺主
11. **逐段动画慎用**：仅在 `numbered_list` / `catalog` / `timeline` 页使用 `by_bullet: true` 配合讲解节奏；`cover` / `divider` / `kpi` / `ending` 禁用
12. **强调动画克制**：`emphasis` 仅用于核心 KPI 数值、关键结论等需突出的内容，每页不超过 1 处
13. **降级兼容**：若不确定效果是否支持，统一使用 `fade`，兼容性最好、风格最稳妥

---

## 十三、不会做的事

- ❌ 不讲占位文本"待补充具体内容"放入大纲 — AI 必须生成真实、可用的业务内容
- ❌ 编造具体数字或统计 — 缺失精确数据时使用合理估算并标注"约"，否则询问用户
- ❌ 跳过大纲确认直接生成 PPT
- ❌ 在未澄清场景前猜测 section key
- ❌ 跳过 Step 1 直接生成大纲
- ❌ Step 3 未确认模板就进入 Step 4
- ❌ 用 markdown 代码块包裹 JSON 输出
- ❌ 在 JSON 前后添加解释性文字
- ❌ 自创字段名或枚举值
- ❌ 自创转场/动画效果名称，或使用中文别名（如 `"transition": "淡入"`）
- ❌ 将 `animations` 写成字符串（如 `"animations": "fly_in"`），必须为对象结构
- ❌ 将 `by_bullet` 写成字符串（如 `"by_bullet": "true"`），必须为布尔值
- ❌ 在 `cover` / `divider` / `kpi` / `ending` 页开启 `by_bullet: true`
- ❌ 商务场景使用 `vortex` / `fling` / `switch` 等炫技转场
- ✅ 每条 desc 控制在 30-60 字（适配模板容量，避免触发缩字号）
- ✅ title 控制在 15 字以内
- ✅ 生成前用 `SceneAdapter.validate_business_data` 校验
- ✅ 输出前完成第六章 10 项自检清单

---

## 十四、项目结构

```
Ppt_work/
├── aippt/                          # 共享核心包
│   ├── __init__.py
│   ├── config.py                   # 集中配置（路径/关键词/默认参数）
│   ├── constants.py                # 共享常量（占位文本/关键词）
│   ├── logger.py                   # 统一日志模块
│   └── validators.py               # 六层防御体系校验引擎
├── aippt_outline.py                # 五步工作流 CLI + 大纲转换
├── ppt_renderer.py                 # 渲染引擎（模板 → 成品 PPT）
├── ppt_scene_adapter.py            # 场景适配（业务字段 → 模板槽位）
├── ppt_meta_tool.py                # 模板元数据解析工具
├── ppt_animations.py               # 动画效果注入
├── ppt_transitions.py              # 转场效果注入
├── import_templates.py             # 模板批量导入 + auto-annotate + rebuild-index
├── generate_thumbnails.py          # 多页缩略图生成（2x2 网格）
├── trim_ppt.py                     # 页面裁剪工具
├── insert_tables.py                # 表格批量插入工具
├── process_safety_templates.py     # 安全教育模板批量处理
├── schemas/                        # JSON Schema 定义（六层防御 Layer 2）
│   ├── requirement_params.schema.json
│   └── outline.schema.json
├── models/                         # 模板库（11 场景分类）
│   ├── templates_index.json        # 全局模板索引
│   └── preview_manifest.json       # 截图预览清单
└── SKILL.md
```

---

## 十五、与本项目的对接

| Skill 步骤 | 项目模块 | 输入 | 输出 |
|---|---|---|---|
| Step 1 Understand | `SceneAdapter.list_scenes()` + `detect_scene()` | 用户描述 | 场景名 |
| Step 2 Outline | `SCENE_SCHEMAS[scene].chapter_sections` | 场景名 | 大纲 JSON |
| Step 3 Visuals | `SceneAdapter.list_templates(category=scene)` | 场景名 | 模板 ID + 截图 |
| Step 4 Generate | `aippt_outline.py step4-generate` + `PptRenderer.render()` | 大纲 + 模板 | .pptx 文件 |

---

## 十六、依赖

- Python 3.10+
- python-pptx >=1.0.2
- lxml（动画/转场注入）
- jsonschema >=4.0（六层防御体系格式校验）
- 可选: pywin32（截图功能）
- 可选: Pillow（多页缩略图拼接）

```bash
# 安装核心依赖
pip install -e .

# 安装全部依赖（含截图、开发工具）
pip install -e ".[all]"
```
