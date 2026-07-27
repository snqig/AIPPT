# AIPPT 产品需求文档（PRD）

> **文档版本**：v2.2  
> **创建日期**：2026-07-24  
> **最近更新**：2026-07-27（v2.2 双引擎 + 设计稿解析 + 主题生成器）  
> **产品名称**：AIPPT（AI-driven PPT Auto-Generation System）  
> **当前版本**：v2.2.0  
> **产品定位**：开源 AI 驱动的 PPT 自动生成系统，通过「模板槽位替换 + 无模板自动布局」双引擎架构将结构化大纲一键转化为成品 PPT

---

## 一、文档修订记录

| 版本 | 日期 | 修订人 | 修订内容 |
|------|------|--------|---------|
| v1.0 | 2026-07-24 | 产品团队 | 初版，基于 v2.0.0 现有能力 + 需求规划输出 |
| v2.2 | 2026-07-27 | 产品团队 | 新增双引擎架构、设计稿 CV 解析、主题生成器、5 套视觉主题；模板扩展至 236 套 |

---

## 二、产品背景与目标

### 2.1 背景

在商务办公场景中，PPT 制作是高频刚需。然而传统 PPT 制作流程存在以下问题：
- **耗时长**：从零做一份专业 PPT 平均耗时 2-4 小时，排版美化占 60% 时间
- **设计门槛高**：大多数职场人士缺乏专业设计能力
- **重复劳动多**：同一内容在不同模板间切换需要重复操作
- **品牌一致性差**：多人协作时难以保持统一视觉规范

市场上已有多款 AI PPT 工具（ChatPPT、Gamma、Beautiful.ai 等），但全部为云端 SaaS，存在数据安全风险、网络依赖、无法定制等问题。

AIPPT 定位为**开源、本地运行、双引擎驱动**的 PPT 自动生成引擎，面向有技术能力的用户和需要私有化部署的企业。v2.2 引入无模板自动布局引擎与设计稿 CV 解析能力，从「模板驱动」升级为「模板 + 设计令牌」双轨架构。

### 2.2 产品目标

| 目标维度 | 目标描述 | 当前达成 |
|---------|---------|---------|
| 核心渲染（模板引擎） | 模板槽位替换准确率 ≥ 98% | ✅ 100%（实测 2000+ 槽位） |
| 核心渲染（自动引擎） | 无模板自动布局生成原生可编辑 PPTX | ✅ v2.2 新增，T503 验证通过 |
| 生成速度 | 单份 PPT 生成 ≤ 3 秒 | ✅ 0.15-0.92 秒 |
| 场景覆盖 | 支持 ≥ 10 类常见商务场景 | ✅ 11 类 |
| 模板数量 | ≥ 50 套可用模板 | ✅ 236 套 |
| 视觉主题 | 支持多套主题切换 | ✅ 5 套（3 商务 + 2 设计令牌） |
| 样式还原 | 100% 保留原模板字体/颜色/布局 | ✅ 100% |
| 易用性 | 支持 CLI 和 Python API 两种调用方式 | ✅ 已支持 |
| 扩展性 | 支持用户自行导入模板 + 自定义主题 | ✅ 已支持 |
| 设计稿注入 | 从设计稿图片自动提取主题令牌 | ✅ T501 原型完成 |

---

## 三、用户故事

| 编号 | 角色 | 需求 | 目的 | 对应功能 |
|------|------|------|------|---------|
| US-01 | 职场白领 | 我希望能输入年终总结的关键数据和要点，自动生成一份专业的年终总结 PPT | 节省排版时间，获得专业视觉效果 | 四步工作流 + 年终总结场景 |
| US-02 | 项目经理 | 我希望将项目复盘的结构化大纲一键生成不同风格的 PPT，方便选择最佳方案 | 快速对比不同视觉风格 | 批量渲染 |
| US-03 | 开发者 | 我希望能通过 API 将 PPT 生成集成到我的系统中 | 实现自动化 PPT 生成流水线 | Python API |
| US-04 | 求职者 | 我希望能用简洁大方的模板生成竞聘 PPT | 在竞聘中展示专业形象 | 个人简历场景 |
| US-05 | 研究生 | 我希望能生成符合学术规范的复试/开题 PPT | 满足学术答辩格式要求 | 自我介绍 + 开题报告场景 |
| US-06 | 企业 IT | 我希望能在内网部署 AIPPT，确保业务数据不外泄 | 满足数据安全合规要求 | 本地部署 |
| US-07 | 模板设计师 | 我希望能将自己设计的 PPT 模板导入系统并自动分类 | 扩展模板库 | 模板导入工具 |
| US-08 | 产品经理 | 我希望能用一句话描述需求，自动生成完整的 PPT | 极致降低 PPT 制作门槛 | **P0 待实现：LLM 集成** |
| US-09 | 业务人员 | 我希望能上传 Word 文档直接生成 PPT | 无需手动整理大纲 | **P0 待实现：文档导入** |

---

## 四、功能清单

### 4.1 功能架构总览

```
AIPPT 功能架构（v2.2）
├── 1. 渲染引擎层（render/）
│   ├── 1.1 BaseRenderer 抽象基类（统一接口 render_outline）
│   ├── 1.2 PptRenderer（模板槽位替换引擎）
│   │   ├── 模板槽位替换 / 文字自适应 / 版权页清理
│   │   └── 子模块：text_replacer / chart_replacer / table_filler
│   ├── 1.3 AutoLayoutRenderer（无模板自动布局引擎）★ v2.2 新增
│   │   ├── 12 列网格 + 安全区 + 分区计算
│   │   ├── 5 类核心页面布局（cover/catalog/divider/numbered_list/kpi）
│   │   └── 文本 auto_fit（10pt 下限）
│   ├── 1.4 转场效果注入（39 种，含 Morph）
│   └── 1.5 动画效果注入（20+ 种，含 by_bullet）
├── 2. 主题系统（themes/）★ v2.2 新增
│   ├── 2.1 theme_loader（含 fallback 机制）
│   ├── 2.2 5 套视觉主题 JSON
│   │   ├── 商务蓝 / 极简灰 / 科技青
│   │   └── guizang-瑞士风-克莱因蓝 / guizang-杂志风-靛蓝瓷
│   └── 2.3 ppt_auto_layout（页面元素生成器）
├── 3. 设计稿解析与主题生成 ★ v2.2 新增
│   ├── 3.1 design_tokens.py（guizang 设计令牌直接抽取）
│   ├── 3.2 design_parser.py（CV 解析：K-means 配色 + 轮廓字号 + 空白带间距）
│   └── 3.3 theme_generator.py（解析器输出 → 标准 theme JSON + overrides 微调）
├── 4. 场景适配层（ppt_scene_adapter.py）
│   ├── 4.1 11 类场景 Schema
│   ├── 4.2 业务字段 → 槽位映射
│   ├── 4.3 页面模式自动识别（8 类）
│   └── 4.4 业务数据校验
├── 5. 大纲转换与校验（aippt_outline.py）
│   ├── 5.1 四步工作流 CLI（step1~step4）
│   ├── 5.2 场景关键词识别
│   ├── 5.3 outline.json ↔ business_data.json
│   └── 5.4 六层防御校验（validate 子命令）
├── 6. 模板管理
│   ├── 6.1 模板库（236 套 / 11 分类）
│   ├── 6.2 模板导入工具
│   ├── 6.3 元数据解析与质量门禁
│   └── 6.4 模板索引管理
├── 7. 批量与扩展
│   ├── 7.1 批量渲染
│   ├── 7.2 PPT 裁剪工具
│   ├── 7.3 表格插入工具
│   └── 7.4 SmartArt 替换
└── 8. [待实现] 智能输入
    ├── 8.1 LLM 集成生成大纲
    └── 8.2 文档导入（Word/PDF/Markdown）
```

### 4.2 详细功能需求（EARS 原则）

---

#### F-01：模板槽位替换

**Ubiquitous**：
- The system shall 在渲染时 100% 保留原模板的字体、颜色、字号、粗体、斜体等格式属性
- The system shall 支持三种匹配策略：shape_name + match_text 双匹配（最精确）→ match_text 文本匹配 → shape_name 匹配（兜底）
- The system shall 对每个槽位的替换操作进行统计，输出 replaced / missed / skipped 计数

**Event-driven**：
- When 渲染引擎收到 slot_data（格式为 `{"页码": {"槽位名": "值"}}`），the system shall 按页遍历 page_slots，逐一执行文本替换
- When 槽位匹配的 shape 包含多个 run，the system shall 保留第一个 run 的格式，替换其文本，删除其余 run
- When 槽位匹配的 shape 包含多个 paragraph，the system shall 保留首段，删除其余段落

**Unwanted**：
- If 模板文件不存在，then the system shall 抛出 `FileNotFoundError` 并给出明确路径信息
- If meta 文件不存在，then the system shall 抛出 `FileNotFoundError` 并给出明确路径信息
- If 槽位未匹配到 shape，then the system shall 记录 warning 日志并计入 missed 统计，不中断整体渲染

---

#### F-02：文字自适应

**State-driven**：
- While 启用 auto_fit（默认开启），the system shall 对每个替换后的文本框计算容量并判断是否需要缩小字号
- While 文本长度超过容量阈值，the system shall 按 `(capacity / text_len)^0.5` 比例缩小字号，下限 8pt

**Event-driven**：
- When meta 中的 slot_info 包含 `capacity.total_chars` 字段，the system shall 优先使用该值作为容量阈值（策略1：事前预警）
- When meta 中无 capacity 字段，the system shall 使用运行时几何估算（策略2：基于 shape 宽高和字号计算）

---

#### F-03：版权页自动清理

**Event-driven**：
- When 启用 remove_copyright（默认开启）且 meta 中存在 `removable_pages` 字段，the system shall 自动删除标记的版权/广告页面
- When 删除页面时，the system shall 按页码倒序删除以避免索引偏移

**Unwanted**：
- If removable_pages 中的页码超出实际页数范围，then the system shall 忽略该页码不报错

---

#### F-04：转场效果注入

**Optional**：
- Where 用户指定 transitions 参数，the system shall 为指定页面注入对应的转场效果
- Where transitions 设为 "auto"，the system shall 为所有页面注入默认 fade 转场（speed=med）

**Ubiquitous**：
- The system shall 支持 ECMA-376 核心 19 种 + PowerPoint 2010+ 扩展 19 种，共 38 种转场效果
- The system shall 对 p14 转场使用 `mc:AlternateContent` 包裹，提供 fade 降级方案

---

#### F-05：动画效果注入

**Optional**：
- Where 用户指定 animations 参数，the system shall 为指定页面的 shape 注入对应动画
- Where animations 设为 "auto"，the system shall 根据页面类型（COVER/CHAPTER/CONTENT/KPI/TIMELINE/CHART/TABLE/END）自动匹配推荐动画方案

**Ubiquitous**：
- The system shall 支持入场（entrance）、退场（exit）、强调（emphasis）三类共 20+ 种动画效果
- The system shall 支持四种触发类型：on_load / on_click / after_prev / with_prev
- The system shall 支持 `text_build: by_bullet` 按段落逐步显示

---

#### F-06：场景 Schema管理

**Ubiquitous**：
- The system shall 预定义 11 类商务场景的 Schema：工作总结、年终总结、工作汇报、工作计划、述职报告、个人简历、自我介绍、开题报告、公司简介、职业规划、安全教育
- The system shall 每个 Schema 包含 `cover_fields`（封面字段）、`chapter_sections`（章节定义）、`end_fields`（结束字段）

**Event-driven**：
- When 调用 `SceneAdapter.adapt(scene, business_data, meta)`，the system shall 依次填充封面页、目录页、各章节、结束页
- When 章节页识别为 divider 模式，the system shall 自动填充 PART.0N 编号和章节名称

---

#### F-07：页面模式自动识别

**Ubiquitous**：
- The system shall 自动识别 8 类页面布局模式：divider（章节分隔页）、numbered_list（数字列表）、timeline（时间轴）、preset_titles（预设标题列表）、skill_percent（技能百分比）、kpi（KPI 卡片）、two_column（双栏对比）、content（标准内容页）

**Event-driven**：
- When 页面包含 chart 形状（`has_chart=true`），the system shall 优先识别为 chart 模式
- When 页面包含 table 形状（`has_table=true`），the system shall 优先识别为 table 模式
- When 页面包含 "PART" 或 "第N章" 关键词且槽位数 ≤ 6，the system shall 识别为 divider 模式
- When 页面包含 percent 槽位，the system shall 识别为 skill_percent 模式
- When 页面包含 year 槽位或 2 个以上年份格式文本，the system shall 识别为 timeline 模式

---

#### F-08：批量渲染

**Event-driven**：
- When 调用 `SceneAdapter.render_batch(scene, business_data, output_dir)`，the system shall 对同分类下所有模板逐一渲染并输出到指定目录
- When 某模板渲染失败，the system shall 记录失败状态并继续处理剩余模板，不中断批量流程

---

#### F-09：模板导入

**Event-driven**：
- When 用户通过 `import_templates.py --src <目录> --prefix <前缀>` 导入新模板，the system shall 自动：1) 根据首頁文本关键词分类到 10 大目录；2) 复制并重命名；3) 生成 meta.json；4) 生成首頁 PNG 截图；5) 标记可删除尾部页面

**Optional**：
- Where 使用 `--removable-tail N` 参数，the system shall 标记末尾 N 页为可删除（版权页/致谢页）
- Where 使用 `--no-screenshot` 参数，the system shall 跳过截图生成（无 PowerPoint 环境时）

---

#### F-10：四步工作流 CLI

**Ubiquitous**：
- The system shall 强制四步依次执行：step1-understand → step2-outline → step3-visuals → step4-generate，不可跳过、不可乱序
- The system shall 每步执行后设置确认 gate，向用户呈现产出物并等待确认后才进入下一步

**Event-driven**：
- When 执行 step1-understand --text "用户描述"，the system shall 基于关键词识别场景并生成澄清问题清单
- When 执行 step2-outline --scene --purpose --audience --length，the system shall 基于场景 Schema 生成结构化大纲 outline.json
- When 执行 step3-visuals --outline，the system shall 推荐匹配模板并按得分排序
- When 执行 step4-generate --outline --template-id，the system shall 完成大纲→business_data→PPT 的全链路生成

---

#### F-13：无模板自动布局引擎（v2.2 新增）

**Ubiquitous**：
- The system shall 提供独立的 `AutoLayoutRenderer` 类，继承 `BaseRenderer`，与 `PptRenderer` 共用统一接口 `render_outline(outline_data, output_path, render_args)`
- The system shall 基于 12 列网格 + 安全区 + 分区计算生成布局，所有坐标尺寸使用 inches（对齐 python-pptx 原生坐标系）
- The system shall 支持 5 类核心页面布局：cover（封面）、catalog（目录）、divider（章节分隔）、numbered_list（数字列表）、kpi（KPI 卡片）
- The system shall 所有样式读取自主题 Design Token 配置，禁止在代码中硬编码颜色、字号、间距

**Optional**：
- Where 用户指定 `--mode auto` 参数，the system shall 调用 AutoLayoutRenderer 进行无模板渲染
- Where 用户指定 `--theme <主题名>` 参数，the system shall 加载对应主题 JSON；主题不可用时 fallback 到默认主题

**Event-driven**：
- When 生成每个元素后，the system shall 收集 shape_id + role 列表，用于动画模块匹配
- When 文本超出文本框容量，the system shall 自动缩小字号，下限 10pt

**Unwanted**：
- If 指定主题文件不存在，then the system shall 记录 warning 并使用默认主题，不中断渲染
- If outline 中包含不支持的 page_type，then the system shall 跳过该页并记录 warning

---

#### F-14：设计稿 CV 解析（v2.2 新增，T501）

**Event-driven**：
- When 用户调用 `design_parser.parse_design_image(image_path)`，the system shall 使用 OpenCV/Pillow 分析图片，输出设计令牌字典
- When 执行配色提取，the system shall 使用 HSV 色彩空间过滤低饱和度/亮度像素，再进行 K-means 聚类，输出主色/辅助色/文本色/背景色
- When 执行字体预估，the system shall 通过轮廓检测识别文本区域，按像素高度预估字号
- When 执行间距测量，the system shall 检测空白带（horizontal/vertical projection），对齐 8px 刻度原则

**Ubiquitous**：
- The system shall 同时提供 `design_tokens.py` 模块，直接抽取 guizang-ppt-skill 的 CSS 变量作为主要方案，CV 解析作为辅助

---

#### F-15：主题生成器（v2.2 新增，T502）

**Event-driven**：
- When 用户调用 `theme_generator.generate_theme(parsed_tokens, theme_name, overrides)`，the system shall 将解析器输出转换为标准 theme JSON 格式
- When 生成主题文件，the system shall 校验 WCAG 对比度，确保文本色与背景色对比度 ≥ 4.5

**Optional**：
- Where 用户提供 `overrides` 参数，the system shall 用用户指定值覆盖解析器提取的值，支持手动微调

**Ubiquitous**：
- The system shall 预置 9 套 guizang 主题预设（4 套瑞士风 + 5 套杂志风）供选择
- The system shall 生成的主题文件可直接被 `theme_loader.py` 加载，无需人工修改

---

#### F-16：渲染引擎模块化（v2.2 新增）

**Ubiquitous**：
- The system shall 将 `ppt_renderer.py` 的渲染子逻辑拆分为独立子模块：`text_replacer.py`（文本替换/shape 定位/字号自适应）、`chart_replacer.py`（4 类图表数据替换）、`table_filler.py`（动态表格行扩展）
- The system shall 主引擎 `PptRenderer` 仅负责调度，子模块函数为无状态纯函数，便于独立测试与复用
- The system shall 对调用方零侵入：现有 `renderer.render(slot_data, output_path, ...)` 调用方式保持不变

---

#### F-11：[待实现] LLM 集成生成大纲

**Event-driven**：
- When 用户输入自然语言描述（如"帮我做一个 Q3 项目复盘 PPT"），the system shall 通过 LLM 理解意图并自动生成 outline.json
- When LLM 返回结果后，the system shall 校验 outline 结构完整性，确保符合场景 Schema

**Unwanted**：
- If LLM 返回的 section key 不在场景 Schema 中，then the system shall 标记警告并使用空数组占位，不中断流程
- If LLM 调用超时或失败，then the system shall 返回明确的错误信息并提示用户重试或手动填写大纲

**Ubiquitous**：
- The system shall 支持可配置的 LLM 后端（OpenAI API / Claude API / 本地模型）

---

#### F-12：[待实现] 文档导入生成 PPT

**Event-driven**：
- When 用户上传 Word/PDF/Markdown 文档，the system shall 自动解析文档结构（标题层级、段落主题、关键数据点）并生成 outline.json
- When 文档解析完成后，the system shall 识别文档类型并推荐匹配的场景

**Unwanted**：
- If 文档格式不支持，then the system shall 返回明确错误提示并列出支持的格式列表
- If 文档解析后无有效内容，then the system shall 提示用户检查文档内容

---

## 五、业务流程

### 5.1 四步工作流（当前）

```
用户输入描述文本
        │
        ▼
┌──────────────────────┐
│ Step 1: Understand   │  场景识别 + 澄清问题
│ (step1-understand)   │
└──────────┬───────────┘
           │ 用户确认场景
           ▼
┌──────────────────────┐
│ Step 2: Outline      │  生成 outline.json
│ (step2-outline)      │
└──────────┬───────────┘
           │ 用户审阅大纲
           ▼
┌──────────────────────┐
│ Step 3: Visuals      │  模板推荐 + 版式匹配
│ (step3-visuals)      │
└──────────┬───────────┘
           │ 用户选定模板
           ▼
┌──────────────────────┐
│ Step 4: Generate     │  渲染生成 PPT
│ (step4-generate)     │
└──────────┬───────────┘
           │
           ▼
      final.pptx ✅
```

### 5.2 渲染引擎内部流程

```
business_data.json + 模板 .pptx + .meta.json
        │
        ▼
┌──────────────────────────┐
│ SceneAdapter.adapt()     │
│ ├─ _fill_cover()         │  封面字段映射
│ ├─ _fill_catalog()       │  目录页填充
│ ├─ _fill_chapters()      │  章节内容适配
│ │   ├─ _detect_page_pattern()  │  页面模式识别
│ │   ├─ _fill_divider_page()    │
│ │   ├─ _fill_numbered_list_page()│
│ │   ├─ _fill_timeline_page()   │
│ │   ├─ _fill_kpi_page()        │
│ │   ├─ _fill_two_column_page() │
│ │   └─ _fill_content_page()    │
│ ├─ _fill_orphan_pages()  │  兜底填充
│ └─ _fill_end()           │  结束页填充
└──────────┬───────────────┘
           │ slot_data: {"页码": {"槽位名": "值"}}
           ▼
┌──────────────────────────┐
│ PptRenderer.render()     │
│ ├─ _find_shape()         │  定位 shape
│ ├─ _replace_text()       │  文本替换
│ ├─ _auto_fit()           │  字号自适应
│ ├─ _remove_slides()      │  版权页删除
│ └─ _inject_effects()     │  转场+动画注入
└──────────┬───────────────┘
           │
           ▼
      output.pptx ✅
```

---

## 六、交互说明

### 6.1 CLI 交互

当前产品主要交互方式为命令行（CLI），面向技术用户：

| 命令 | 用途 | 示例 |
|------|------|------|
| `python ppt_scene_adapter.py scenes` | 列出所有场景 | - |
| `python ppt_scene_adapter.py schema --scene 工作总结` | 查看场景 Schema | - |
| `python ppt_scene_adapter.py templates --category 工作总结` | 列出模板 | - |
| `python ppt_scene_adapter.py detail --template xxx` | 查看模板详情 | - |
| `python ppt_scene_adapter.py validate --scene xxx --data xxx` | 校验数据 | - |
| `python ppt_scene_adapter.py render --scene xxx --data xxx --output xxx` | 单模板渲染 | - |
| `python ppt_scene_adapter.py batch --scene xxx --data xxx` | 批量渲染 | - |

### 6.2 Python API 交互

面向开发者集成：

```python
# 初始化
adapter = SceneAdapter("models")

# 查询
scenes = adapter.list_scenes()
templates = adapter.list_templates(category="工作总结")

# 校验
is_valid, issues = adapter.validate_business_data("工作总结", data)

# 渲染
adapter.render("工作总结", data, template_id="xxx", output_path="output.pptx")

# 批量
results = adapter.render_batch("工作总结", data, output_dir="output_batch")
```

### 6.3 待实现的交互方式

| 交互方式 | 目标用户 | 优先级 |
|---------|---------|--------|
| Web UI | 非技术用户 | P0 |
| 桌面 GUI | 非技术用户 | P2 |
| REST API | 系统集成商 | P1 |
| 自然语言对话 | 所有用户 | P0 |

---

## 七、数据指标

### 7.1 性能指标

| 指标 | 目标值 | 当前实测 | 是否达标 |
|------|--------|---------|---------|
| 单份生成耗时 | ≤ 3 秒 | 0.15 - 0.92 秒 | ✅ |
| 替换准确率 | ≥ 98% | 100%（2000+ 槽位） | ✅ |
| 样式还原度 | 100% | 100% | ✅ |
| 批量 12 套耗时 | ≤ 20 秒 | 3.92 秒 | ✅ |
| 模板槽位识别率 | ≥ 95% | 待验证 | - |

### 7.2 质量指标（待上线后追踪）

| 指标 | 说明 | 目标值 |
|------|------|--------|
| 生成成功率 | 成功生成 / 总生成次数 | ≥ 99% |
| 用户满意度 | NPS 评分 | ≥ 40 |
| 模板使用率 | 被使用模板 / 总模板数 | ≥ 60% |
| 场景覆盖率 | 被使用场景 / 总场景数 | ≥ 80% |

### 7.3 埋点需求（待 Web UI 上线后实施）

| 埋点事件 | 说明 | 属性 |
|---------|------|------|
| `ppt_generate_start` | 开始生成 PPT | scene, template_id, input_type |
| `ppt_generate_complete` | PPT 生成完成 | scene, template_id, duration_ms, page_count |
| `ppt_generate_error` | PPT 生成失败 | scene, error_type, error_message |
| `template_import` | 导入新模板 | category, page_count |
| `scene_select` | 选择场景 | scene_name |
| `batch_render` | 批量渲染 | scene, template_count, success_count |

---

## 八、验收标准

### 8.1 渲染引擎验收

| 编号 | 验收项 | 验收标准 | 测试方法 |
|------|--------|---------|---------|
| AC-01 | 槽位替换 | 所有可识别槽位 100% 替换，无残留占位文本 | 全量冒烟测试 |
| AC-02 | 样式保留 | 生成后的 PPT 字体/颜色/字号/粗斜体与原模板一致 | 视觉对比 |
| AC-03 | 长文本自适应 | 超长文本自动缩小字号，不溢出文本框，字号 ≥ 8pt | 边界场景测试 |
| AC-04 | 版权页清理 | 标记的版权页被删除，其余页面不受影响 | 元数据校验 |
| AC-05 | 转场注入 | 指定转场效果在 PowerPoint 中正常播放 | 手动验收 |
| AC-06 | 动画注入 | 指定动画效果在 PowerPoint 中正常播放 | 手动验收 |

### 8.2 场景适配验收

| 编号 | 验收项 | 验收标准 | 测试方法 |
|------|--------|---------|---------|
| AC-07 | 11 类场景覆盖 | 11 类场景均有可用模板并能成功生成 | 场景校验脚本 |
| AC-08 | 业务数据校验 | 不符合 Schema 的数据返回明确的问题清单 | validate CLI |
| AC-09 | 页面模式识别 | 8 类页面模式均能正确识别 | 元数据校验 |
| AC-10 | 批量渲染 | 同分类所有模板均生成成功 | smoke_test_all.py |
| AC-10a | 无模板自动布局 | `--mode auto` 能生成 5 类核心页面（cover/catalog/divider/numbered_list/kpi） | T503 集成测试 |
| AC-10b | 主题切换 | `--theme` 参数能切换 5 套主题，封面/正文颜色随之变化 | T503 主题对比测试 |
| AC-10c | 设计稿解析 | `design_parser.parse_design_image()` 能从设计稿图片提取配色/字号/间距 | T501 验证脚本 |
| AC-10d | 主题生成器 | `theme_generator.generate_theme()` 能输出符合 schema 的 theme JSON | T502 单元测试 |

### 8.3 边界场景验收

| 编号 | 验收项 | 验收标准 |
|------|--------|---------|
| AC-11 | 空数据 | sections 为空时不崩溃，生成仅有封面+结束页的 PPT |
| AC-12 | 超长文本 | desc 超过 200 字时自动缩字号，不溢出 |
| AC-13 | 特殊字符 | 含 emoji/特殊符号的文本正常渲染 |
| AC-14 | 缺失字段 | cover 缺失非必填字段时不报错 |
| AC-15 | 模板缺失 | 场景无可用模板时给出明确错误提示 |

---

## 九、附录

### 9.1 技术栈

| 层级 | 技术 | 版本要求 |
|------|------|---------|
| 语言 | Python | ≥ 3.10 |
| PPT 操作 | python-pptx | ≥ 1.0.2 |
| XML 处理 | lxml | ≥ 5.0 |
| 类型检查 | mypy | ≥ 1.10 |
| 代码规范 | ruff | ≥ 0.5 |

### 9.2 场景 Schema 速查

| 场景 | 章节数 | 章节 key |
|------|--------|---------|
| 工作总结 | 4 | work_content / project_progress / issues / plan |
| 年终总结 | 4 | annual_review / achievements / experience / next_year |
| 工作汇报 | 4 | progress / results / challenges / next_steps |
| 工作计划 | 4 | objectives / tasks / schedule / resources |
| 述职报告 | 4 | duty_performance / achievements / problems / improvement |
| 个人简历 | 4 | basic_info / competency / job_awareness / career_plan |
| 自我介绍 | 7 | basic_info / education / internship / research / awards / self_assessment / future_plan |
| 开题报告 | 4 | background / literature / methodology / conclusion |
| 公司简介 | 9 | overview / honors / leadership / business / main_products / marketing / market_analysis / development / future |
| 职业规划 | 4 | self_analysis / career_goal / action_plan / risk_assessment |
| 安全教育 | 视模板而定 | safety_awareness / hazards / prevention / emergency 等 |

### 9.3 模板分布统计（v2.2）

| 分类 | 模板数量 |
|------|---------|
| 安全教育 | 157 |
| 工作总结 | 39 |
| 工作汇报 | 12 |
| 年终总结 | 8 |
| 述职报告 | 8 |
| 工作计划 | 5 |
| 公司简介 | 2 |
| 开题报告 | 2 |
| 个人简历 | 1 |
| 自我介绍 | 1 |
| 职业规划 | 1 |
| **合计** | **236** |

---

---
## 十、文档关联

本文档为阶段二（PRD）产出。请继续参阅：
- **01_需求规划文档.md** — 阶段一需求规划，定义目标用户、痛点与优先级
- **03_功能评审报告.md** — 阶段三功能评审，包含架构评估与代码质量分析

---

> 📌 **流转提醒**：PRD 完成后，建议将文档上传到项目资料库，创建评审事项分配给技术负责人、设计师和测试同学。评审通过后进入阶段三（设计与研发评审）。
