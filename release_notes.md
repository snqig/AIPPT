## AIPPT v2.0.0 - AI驱动的PPT自动生成系统

### 核心功能
- 全新核心包 aippt/ (config + constants + logger)
- 6 个核心模块重构：logging + 类型注解 + 共享常量
- 38 种转场 + 20+ 动画效果
- 10 类场景覆盖，79 套模板元数据

### 增强功能
- SKILL.md 实现"一句话生成"：用户说一句话 → AI 自动填充真实大纲 → CLI 渲染 → 成品 PPT
- 79 套模板元数据完整入库 (meta.json + templates_index)

### 性能指标
- Slot 替换准确率 100%
- 生成速度 0.15-0.92s
- 100% 保留原模板样式

### Bug 修复
- 修复双 .pptx 后缀问题 (get_template_pptx)
- 兼容 xxx.pptx.meta.json / xxx.meta.json 两种命名自动识别

### 使用方式
```bash
git clone https://github.com/snqig/AIPPT.git
pip install -e .
```