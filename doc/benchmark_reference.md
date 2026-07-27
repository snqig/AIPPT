# AIPPT 对标项目吸收说明

## 一、对标项目概览

本项目在迭代过程中对标了三个开源项目，吸收其设计理念，结合自身「槽位替换 + 五步工作流」架构进行差异化实现。

| 对标项目 | 定位 | 借鉴核心 | AIPPT 差异化实现 |
|---|---|---|---|
| pptx-from-layouts-skill | 模板版式解析 | 母版/版式深度解析、元素角色识别 | 增加场景适配层，业务字段自动映射槽位 |
| python-office-templates | 富媒体渲染 | 图表数据源替换、表格动态扩展 | 保留模板样式100%，多系列自动适配 |
| ppt-master | XML 底层方案 | SmartArt 操作、动画时间线控制 | 封装为 Python API，集成到渲染引擎 |

## 二、pptx-from-layouts-skill 借鉴明细

### 2.1 借鉴点
1. **母版与版式深度解析**：遍历 slide master 下所有 layout，提取占位符类型、索引、位置、尺寸、默认样式
   - AIPPT 实现：`aippt/profile_layouts.py` 的 `profile_layouts(prs)` 函数
2. **元素角色识别**：对非占位符文本框，通过字号、位置、格式特征判定角色（标题/正文/KPI数值/备注），输出置信度
   - AIPPT 实现：`aippt/ppt_element_classifier.py` 的 `classify_element/classify_page` 函数
3. **页面模式自动分类**：基于元素构成判定页面类型（cover/divider/numbered_list/kpi/timeline/two_column/table/chart）
   - AIPPT 实现：`ppt_scene_adapter.py` 的 `_detect_page_pattern` 方法，识别8种页面模式
4. **低置信度标注提示**：对识别不确定的槽位标红，输出人工复核建议
   - AIPPT 实现：`auto_annotate` 命令输出 warnings 列表

### 2.2 差异化优势
- pptx-from-layouts-skill 仅做模板解析，AIPPT 在解析基础上增加「场景适配层」，将业务字段自动映射到模板槽位
- AIPPT 的 `_detect_page_pattern` 不仅识别页面类型，还针对性填充内容（title/desc配对、number自动序号、timeline事件映射）
- AIPPT 集成质量门禁（`ppt_meta_tool.py check`），模板接入后自动校验完整性

## 三、python-office-templates 借鉴明细

### 3.1 借鉴点
1. **图表数据源替换**：基于 python-pptx 的 chart 对象，支持传入结构化数据更新图表数值
   - AIPPT 实现：`ppt_renderer.py` 的 `_replace_chart_data()` 方法，使用 `chart.replace_data(ChartData)` API
2. **多系列图表适配**：柱状图、折线图、饼图、雷达图 4 类，自动适配系列数量
   - AIPPT 实现：M>N 仅替换前N系列；M<N 多余系列清空数据；M==N 直接替换
3. **表格动态行扩展**：模板预设表头样式，传入N行数据自动追加行，继承表头样式与列宽
   - AIPPT 实现：`_fill_dynamic_table()` 方法，克隆最后一行保持样式，自动行高适配
4. **槽位类型体系扩展**：新增 chart、table 槽位类型
   - AIPPT 实现：`schemas/outline.schema.json` 新增 chart_type/chart_data/headers/rows 字段

### 3.2 差异化优势
- python-office-templates 替换图表时可能丢失样式，AIPPT 使用 `replace_data` 100% 保留模板配色、字体、坐标轴样式
- AIPPT 表格扩展时自动行高适配（短内容0.4英寸，长内容0.6英寸），避免溢出
- AIPPT 提供 `insert_tables.py` 独立测试工具，可单独验证图表/表格替换效果，输出前后对比报告
- AIPPT 图表/表格与六层防御校验集成，chart_data/table_data 需过 Schema 校验

## 四、ppt-master 借鉴明细

### 4.1 借鉴点
1. **SmartArt 文本节点替换**：直接操作 SmartArt 的 XML 节点（dgm 命名空间），实现节点文本精准替换
   - AIPPT 实现：`ppt_smartart.py` 的 `replace_smartart_text(slide, replacements)` 函数
2. **演讲者备注注入**：支持每页传入备注文本，自动写入演讲者备注栏
   - AIPPT 实现：`ppt_renderer.py` 的 `_inject_notes(slide, notes_text)` 方法，通过 `notes_map` 参数传入
3. **动画时间线精细控制**：按段落（by_bullet）时间间隔设置、动画顺序调整、触发方式配置
   - AIPPT 实现：`ppt_animations.py` 的 `_build_by_bullet_nodes()` 函数，支持 bullet_delay_ms 段间延迟、bullet_order 播放顺序
4. **平滑切换效果**：Morph 等高级转场的 XML 注入
   - AIPPT 实现：`ppt_transitions.py` 的 `_build_morph_xml()` 函数，p159 命名空间 AlternateContent 结构

### 4.2 差异化优势
- ppt-master 偏底层 XML 操作，学习曲线陡峭；AIPPT 封装为 Python API，集成到渲染引擎，一行参数即可启用
- AIPPT 动画转场枚举封闭（39转场+20动画），通过 JSON Schema 校验，避免非法名称导致渲染失败
- AIPPT 提供3套动画预设主题（business/tech/formal），一键切换全套风格，ppt-master 需逐页配置
- AIPPT 的 by_bullet 智能适配：非列表页自动关闭并给出警告（A005错误码）

## 五、架构决策记录

### 5.1 为何选择「槽位替换」而非「内容生成」
- 槽位替换 100% 保留模板样式，内容生成（如 python-pptx 从零构建）无法保证视觉一致性
- 槽位替换性能高（0.15-0.92秒），内容生成需处理布局算法，耗时数倍
- 槽位替换与模板解耦，同一模板可服务多个场景

### 5.2 为何自研校验体系而非用通用 Schema
- 通用 Schema 仅校验语法，AIPPT 需校验业务规则（页面类型与模板兼容性、槽位匹配度）
- 六层防御体系含自动修复（auto_fix_outline），通用 Schema 无此能力
- 错误码体系（F/S/T/A）提供精准修正指引，降低模型重试成本

### 5.3 为何选择 OOXML 注入而非 python-pptx 原生 API
- python-pptx 1.0.2 不支持动画、转场、Morph 等高级效果
- OOXML 注入直接操作 XML，可实现 python-pptx 无法覆盖的所有效果
- 封装为 Python 函数，对外暴露简洁 API，隐藏 XML 复杂性

## 六、后续迭代方向
- 模板自动 Profile 进一步强化：元素角色识别准确率提升至 98%+
- 图表类型扩展：支持组合图、