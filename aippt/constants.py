"""
共享常量模块：集中管理各处复用的占位文本/关键词列表
避免 ppt_meta_tool.py 与 aippt_outline.py 等模块中的重复定义
"""

# ==================== 可替换槽位识别关键词 ====================
# 用于 ppt_meta_tool.extract_slots 和渲染后残留校验
SLOT_MATCH_KEYWORDS: list[str] = [
    "在此录入", "请输入", "添加标题", "输入标题", "请替换", "您的内容打在这里",
    "在此处输入", "添加文字", "输入文字", "请在此处添加", "单击此处输入",
    "PROJECT –", "TITLE HERE", "YOUR CONTENT", "ENTER HERE", "预设标题", "添加内容",
    "点击输入简要文字", "请在此添加文字", "单击此处添加", "在这里说点什么",
    "点击添加标题", "添加您的标题", "请输入姓名", "可通过右键选择只保留文本",
    "复制您的文本后", "在此框中选择粘贴", "Text Here",
    "某某省", "某某市", "约翰·史密斯", "请在此添加文字说明",
    "此处添加您的文本内容", "在此处添加您的文本内容", "添加您的文本内容",
    "添加文本内容", "预设标题文本", "Add your text content",
    "添加小标题", "标题内容", "关键词标题", "点击输入简要文字解说",
    "输入你的正文", "您的文本内容",
    "单击输入标题", "在此输入标题", "输入您的主要标题", "单击此处输入文本内容",
    "单击此处添加合适文字", "请您单击此处添加", "请在此处添加具体内容",
    "您的内容", "点击输入", "请在此", "单击此处", "在这里说", "添加您的", "预设标题",
    "某某省", "某某市", "添加描述", "输入标题", "输入文字", "在此录入", "请替换",
    "TITLE HERE", "YOUR CONTENT", "添加内容", "点击添加", "请在此添加",
    "单击此处输入", "复制您的文本后", "在此框中选择粘贴",
    "可通过右键", "添加小标题", "关键词标题", "点击输入简要文字",
    "输入你的正文", "您的文本内容", "单击输入标题",
    "在此输入标题", "输入您的主要标题", "单击此处输入文本内容",
    "请您单击此处添加", "请在此处添加具体内容", "单击此处添加合适文字",
]

# ==================== 占位提示文本关键词 ====================
# 用于 SceneAdapter._is_placeholder_text
PLACEHOLDER_KEYWORDS: list[str] = [
    "点击添加", "添加标题", "请在此处添加", "请在此录入", "在此录入",
    "请替换文字", "请输入", "单击此处", "添加描述", "添加内容",
    "输入标题", "输入文字", "添加文字", "标题文字内容", "预设标题",
    "添加标题内容", "输入标题内容", "您的内容打在这里", "复制您的文本后",
    "在此框中选择粘贴", "可通过右键", "在这里说点什么", "添加您的标题",
    "点击输入简要文字", "Text Here",
]

# ==================== 图表装饰文本关键词 ====================
# 用于 SceneAdapter._is_chart_decoration
CHART_DECORATION_KEYWORDS: list[str] = [
    "单位：", "销售额",
]

# ==================== Copyright 检测关键词 ====================
COPYRIGHT_KEYWORDS: list[str] = [
    "版权", "授权", "素材来源", "包图网", "1ppt", "仅供学习", "禁止商用",
    "模板来源", "ibaotu", "ppt模板", "素材授权",
]

# ==================== 辅助函数 ====================
def is_placeholder(text: str) -> bool:
    if not text:
        return False
    return any(kw in text for kw in PLACEHOLDER_KEYWORDS)


def is_chart_decoration_text(text: str) -> bool:
    if not text:
        return False
    cleaned = text.strip()
    if cleaned.isdigit():
        val = int(cleaned)
        if val == 0 or (50 <= val <= 1000 and val % 50 == 0):
            return True
        return False
    for kw in CHART_DECORATION_KEYWORDS:
        if kw in text:
            return True
    if cleaned == "添加文字":
        return True
    return False
