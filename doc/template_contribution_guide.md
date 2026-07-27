# AIPPT 模板贡献指南

> **适用版本**：v2.2+（双引擎架构）
> **当前模板库**：236 套 / 11 类场景

## 一、前言

本指南面向希望向 AIPPT 模板库贡献 PPT 模板的设计者与开发者。AIPPT v2.2 采用「模板槽位替换 + 无模板自动布局」双引擎架构：

- **模板引擎（PptRenderer）**：需要符合规范的 PPTX 模板，本指南适用
- **自动引擎（AutoLayoutRenderer）**：无需模板，基于 Design Token 自动生成，本指南不适用

本指南针对模板引擎路径。模板需符合特定规范才能被自动标注（auto-annotate）和渲染引擎正确处理。遵循本指南可确保模板快速接入、高质量渲染。

## 二、模板基本要求

### 2.1 文件格式
- 格式：.pptx（PowerPoint 2007+）
- 比例：16:9 优先（也支持 4:3）
- 字体：使用常见中文字体（微软雅黑/思源黑体/苹方），嵌入字体需测试兼容性
- 页数：建议 15-35 页（少于10页信息量不足，多于40页渲染慢）

### 2.2 场景分类
模板需归入以下 11 类场景之一：
工作总结 / 年终总结 / 工作汇报 / 工作计划 / 述职报告 / 个人简历 / 自我介绍 / 开题报告 / 公司简介 / 职业规划 / 安全教育

### 2.3 页面结构建议
完整模板应包含以下页面类型（顺序可调）：
1. 封面页（cover）：标题 + 副标题
2. 目录页（catalog）：章节列表
3. 章节分隔页（divider）：PART.01 / 第N章 等编号 + 标题
4. 内容页（numbered_list/content）：序号 + 标题 + 描述
5. KPI 页（kpi）：标签 + 数值 + 趋势
6. 时间轴页（timeline）：时间 + 事件
7. 双栏对比页（two_column）：左右对比
8. 图表页（chart）：含柱状/折线/饼/雷达图
9. 表格页（table）：含表格
10. 结尾页（ending）：感谢语

## 三、模板制作规范

### 3.1 占位符文本规范（重要）
模板中的占位文本将被自动识别为槽位并替换，请遵循：
- 标题占位：使用「添加标题」「请输入标题」等明确提示文本
- 正文占位：使用「请输入内容」「点击编辑」等
- 避免使用纯英文装饰文本（如 LOREM IPSUM），会被识别为装饰元素跳过
- 避免使用纯数字（如 12345），会被识别为图表坐标轴数据跳过

### 3.2 字号与位置规范（影响角色识别）
自动标注引擎通过字号和位置判定元素角色：
- 标题：字号 ≥ 36pt，位于页面上 30% 区域
- 副标题：字号 24-32pt，位于页面上 40% 区域
- 正文：字号 14-22pt，位于页面中部
- KPI 数值：字号 ≥ 40pt，居中或突出位置
- 备注：字号 ≤ 12pt，位于页面底部

### 3.3 配色规范
- 主色调不超过 3 种（质量门禁会检查，超过 3 种会 warning）
- 使用主题色（theme color）而非硬编码 RGB，便于适配
- 文字与背景对比度 ≥ 4.5:1（WCAG AA 标准）

### 3.4 版权页标记
如模板含版权/广告页：
- 在 meta.json 的 `removable_pages` 字段标注页码（从1开始）
- 渲染时默认自动删除，`--keep-copyright` 可保留
- 未标记的页面不会被删除

### 3.5 图表与表格规范
- 图表：使用 PowerPoint 原生图表（非图片），支持柱状/折线/饼/雷达
- 表格：使用 PowerPoint 原生表格，表头在第 0 行，预留至少 2 数据行
- 图表/表格的数据将被动态替换，模板中可放任意示例数据

### 3.6 SmartArt 规范
- 使用 PowerPoint 原生 SmartArt（组织架构图/流程图/关系图）
- SmartArt 文本节点将被精准替换，结构与配色保留
- 避免使用图片化的 SmartArt（无法替换文本）

## 四、模板接入流程（5步）

### Step 1: 准备模板文件
将 .pptx 文件放入临时目录，命名为「场景_模板名.pptx」（如 `工作总结_蓝色商务.pptx`）

### Step 2: 自动标注
```bash
python import_templates.py auto-annotate \
  --input 工作总结_蓝色商务.pptx \
  --scene 工作总结 \
  --output models/工作总结/
```
自动完成：
- 母版与版式解析（profile_layouts）
- 元素角色识别（ppt_element_classifier）
- 页面模式分类（_detect_page_pattern）
- 生成 .meta.json 元数据
- 生成 2x2 多页缩略图
- 更新 templates_index.json 索引
- 输出质量报告与低置信度 warnings

### Step 3: 人工复核
检查自动标注输出的 warnings 列表，对低置信度元素进行人工修正：
- 打开生成的 .meta.json
- 核对 page_slots 中每个槽位的 slot/type/match_text 是否准确
- 修正错误的页面类型分类
- 补充 removable_pages（版权页页码）

### Step 4: 质量门禁校验
```bash
python ppt_meta_tool.py check --template-id 工作总结_蓝色商务
```
校验项：
- meta_required：必填字段完整性（template_id/category/total_pages/chapters/page_slots）
- rendering_test：渲染测试（能否正常打开、页数是否匹配）
- style_check：样式检查（主色调数量、字体一致性）
- meta_completeness：元数据完整性（chapters 结构、page_slots 覆盖率）

校验通过后模板才能进入索引库。

### Step 5: 冒烟测试
```bash
python -m pytest tests/test_smoke_all.py -v -k "工作总结_蓝色商务"
```
使用场景对应的业务数据渲染模板，验证：
- 渲染无报错
- 输出页数与预期一致（total_pages - len(removable_pages)）
- 槽位替换率 ≥ 90%

## 五、质量标准

### 5.1 必须通过（check_pass = True）
- meta_required：所有必填字段完整
- rendering_test：模板可正常渲染
- meta_completeness：元数据结构完整

### 5.2 建议通过（warning，不阻断）
- style_check：主色调 ≤ 3 种
- 槽位覆盖率 ≥ 80%
- 低置信度元素 ≤ 5 个

### 5.3 质量评分
模板入库后会获得 quality_score（0-100），评分维度：
- 元数据完整性（30分）
- 渲染稳定性（30分）
- 样式规范性（20分）
- 槽位覆盖率（20分）

## 六、模板标签体系

templates_index.json 为每个模板维护以下标签，支持多维度筛选：
- style_tags：风格标签（如 ["商务","16:9"]）
- color_scheme：色系（蓝色系/红色系/绿色系/灰色系/多彩）
- industry：适用行业（如 ["通用"] 或 ["金融","教育"]）
- page_range：页数范围（如 "21-25页"）
- quality_score：质量评分（0-100）

可通过 `SceneAdapter.list_templates(category, style_tag, min_pages, max_pages)` 筛选。

## 七、常见问题

### Q1: 自动标注识别错误怎么办？
A: 检查 warnings 列表，手动修正 .meta.json 中对应槽位的 match_text 和 type。常见原因：装饰文本与占位文本混淆、字号异常。

### Q2: 模板含特殊元素（视频/音频）能接入吗？
A: 不建议。当前渲染引擎仅处理文本、图片、图表、表格、SmartArt。特殊元素会被跳过，可能导致布局错位。

### Q3: 如何贡献多套模板？
A: 每套模板独立执行 auto-annotate，建议每类场景至少 3 套模板覆盖不同风格（商务蓝/简约灰/科技风）。

### Q4: 模板更新后如何重新标注？
A: 删除旧的 .meta.json，重新执行 auto-annotate。templates_index.json 会自动更新。

### Q5: 缩略图生成失败怎么办？
A: 缩略图依赖 pywin32（Windows COM）。安装 `pip install pywin32`。缩略图失败不阻断 meta 生成，可手动截图补充。
