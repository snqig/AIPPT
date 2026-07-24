# AIPPT — AI 驱动的 PPT 自动生成系统

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](LICENSE)
[![opencode skill](https://img.shields.io/badge/opencode-skill-purple)](SKILL.md)

将用户描述或结构化大纲一键转化为成品 PPT。采用「模板槽位替换」架构，保留原模板样式与布局，支持 10 类商务场景、38 种转场效果、20+ 动画效果。

---

## ✨ 核心能力

| 能力 | 说明 |
|------|------|
| **10 类场景** | 工作总结 / 年终总结 / 工作汇报 / 工作计划 / 述职报告 / 个人简历 / 自我介绍 / 开题报告 / 公司简介 / 职业规划 |
| **样式 100% 保留** | 字体、颜色、字号、粗斜体与原模板完全一致 |
| **文字自适应** | 超长文本自动缩小字号，避免溢出文本框 |
| **版权页清理** | 生成时自动删除模板中版权/广告页 |
| **转场 & 动画** | 38 种转场效果 + 入场/退场/强调动画 |
| **页面类型识别** | 自动识别封面/目录/章节/内容/KPI/时间轴/图表/表格页 |
| **批量渲染** | 同一份内容一键生成同分类下所有模板 PPT |

**性能**：单份生成 0.15~0.92 秒，替换准确率 100%（实测 2000+ 槽位）。

---

## 🚀 快速开始

### 安装

```bash
pip install python-pptx>=1.0.2
```

### 一句话生成

```bash
python aippt_outline.py step4-generate \
  --outline my_outline.json \
  --output my_ppt.pptx \
  --transitions auto \
  --animations auto
```

### 四步工作流

```bash
# Step 1: 理解需求 → 识别场景
python aippt_outline.py step1-understand --text "帮我做一个年终总结PPT"

# Step 2: 构建大纲 → 生成 outline.json
python aippt_outline.py step2-outline \
  --scene 年终总结 \
  --purpose "2025年度业绩汇报" \
  --audience "管理层" \
  --length 12 \
  --output outline.json

# Step 3: 视觉匹配 → 推荐模板
python aippt_outline.py step3-visuals --outline outline.json

# Step 4: 生成 PPT
python aippt_outline.py step4-generate \
  --outline outline.json \
  --template-id 年终总结_年终总结 \
  --output final.pptx
```

### Python API

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
                transitions="auto", animations="auto")
```

---

## 📂 项目结构

```
├── aippt/                      # 共享核心包
│   ├── config.py               # 集中配置（路径/关键词/默认参数）
│   ├── constants.py            # 共享常量（去重占位文本等）
│   └── logger.py               # 统一日志
├── aippt_outline.py            # 四步工作流 CLI
├── ppt_renderer.py             # 渲染引擎
├── ppt_scene_adapter.py        # 场景适配器
├── ppt_meta_tool.py            # 模板元数据解析
├── ppt_animations.py           # 动画注入（20+ 效果）
├── ppt_transitions.py          # 转场注入（38 种效果）
├── import_templates.py         # 模板批量导入
├── models/                     # 模板库（按场景分类）
│   ├── 工作总结/
│   ├── 年终总结/
│   ├── 工作汇报/
│   ├── 工作计划/
│   ├── 述职报告/
│   ├── 个人简历/
│   ├── 自我介绍/
│   ├── 开题报告/
│   ├── 公司简介/
│   ├── 职业规划/
│   └── templates_index.json   # 模板总索引
├── SKILL.md                    # opencode skill 定义
├── SKILL_USAGE.md              # 完整使用手册（437 行）
└── pyproject.toml              # 工程配置
```

---

## 📖 完整文档

所有详细用法、API 参考、数据格式、模板接入流程、测试方法均在 **[SKILL_USAGE.md](SKILL_USAGE.md)** 中，包含：

- 7 条 CLI 命令及示例
- Python API 完整调用手册
- 业务数据 JSON 格式与字段说明
- 10 场景 section key 对照表
- 新增模板接入流程（5 步）
- 渲染引擎参数说明
- 测试与验证方法
- 常见问题排查

---

## 🧠 架构

```
用户描述 / outline.json
        │
        ▼
aippt_outline.py  ──── step1 理解 → step2 大纲 → step3 视觉 → step4 生成
        │
        ▼
ppt_scene_adapter.py  ──── validate → adapt（业务字段 → 槽位映射）
        │
        ▼
ppt_renderer.py  ──── _find_shape → _replace_text → _auto_fit → _inject_effects
        │
        ▼
    final.pptx ✅
```

### 页面模式自动识别

`SceneAdapter._detect_page_pattern` 自动识别 8 类页面布局：

| 模式 | 说明 |
|------|------|
| divider | 章节分隔页（PART.01 / 第N章） |
| numbered_list | 数字列表（序号 + 标题 + 描述） |
| timeline | 时间轴（年份 + 事件） |
| preset_titles | 预设标题列表 |
| skill_percent | 技能百分比 |
| kpi | KPI 卡片页 |
| two_column | 双栏对比 |
| chart / table | 图表 / 表格页 |

---

## 🎬 动画 & 转场

- **38 种转场**：fade / push / wipe / dissolve / zoom / conveyor / flip / vortex 等
- **20+ 动画效果**：入场（fade / fly_in / zoom / wipe）+ 退场 + 强调（pulse / spin）
- **自动推荐**：基于页面类型（COVER / CHAPTER / CONTENT / KPI / TIMELINE）自动匹配动画方案
- **逐段动画**：支持 `by_bullet` 按段落逐步显示

---

## 🧪 测试

```bash
# 全量冒烟测试
python smoke_test_all.py

# 边界场景测试
python edge_test.py

# 元数据质量校验
python ppt_meta_tool.py check --dir models
```

---

## 📜 许可

GPL-3.0