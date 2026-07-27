from pathlib import Path

# ==================== 路径配置 ====================
MODELS_ROOT = Path(__file__).resolve().parent.parent / "models"
TEMPLATES_INDEX = MODELS_ROOT / "templates_index.json"
PREVIEW_MANIFEST = MODELS_ROOT / "preview_manifest.json"

# ==================== 渲染默认参数 ====================
DEFAULT_REMOVE_COPYRIGHT = True
DEFAULT_AUTO_FIT = True
DEFAULT_TRANSITIONS = "auto"
DEFAULT_ANIMATIONS = "auto"

# ==================== 场景配置 ====================
SCENE_NAMES = [
    "工作总结", "年终总结", "工作汇报", "工作计划", "述职报告",
    "个人简历", "自我介绍", "开题报告", "公司简介", "职业规划",
    "安全教育",
]

# 场景关键词映射（用户输入 → 场景名）
SCENE_KEYWORDS: dict[str, list[str]] = {
    "工作总结": ["工作总结", "年度回顾", "工作成绩", "工作概况"],
    "年终总结": ["年终", "年度汇报", "去年回顾", "全年总结"],
    "工作汇报": ["工作汇报", "项目汇报", "阶段汇报", "项目复盘"],
    "工作计划": ["工作计划", "年度规划", "下步计划", "工作安排"],
    "述职报告": ["述职", "履职", "岗位汇报", "岗位述职"],
    "个人简历": ["竞聘", "简历", "岗位申请", "求职"],
    "自我介绍": ["自我介绍", "复试", "面试", "研究生复试"],
    "开题报告": ["开题", "答辩", "论文", "毕业论文"],
    "公司简介": ["公司介绍", "企业简介", "公司概况", "公司简介"],
    "职业规划": ["职业规划", "职业生涯", "发展规划", "职业发展"],
    "安全教育": ["安全教育", "安全培训", "安全生产", "安全知识", "消防教育", "应急演练"],
}

# 页面类型识别关键词
PAGE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "cover": [
        "汇报人", "姓名", "年度", "WORK REPORT", "RESUME", "个人简历",
        "公司简介", "开题报告", "自我介绍",
    ],
    "catalog": ["目录", "CONTENTS", "目 录", "目录页", "CONTENT"],
    "chapter": [
        "PART ", "第", "章", "PART ONE", "PART TWO", "PART 01", "PART 02",
        "COMPETENCY", "JOB AWARENESS", "INFORMATION", "CAREER PLANNING",
        "岗位认知", "胜任能力", "个人信息", "职业规划", "自我评价",
    ],
    "end": ["感谢", "THANK YOU", "谢谢聆听", "结束", "致谢", "谢谢观看"],
    "copyright": [
        "版权", "授权", "素材来源", "包图网", "1ppt", "仅供学习", "禁止商用",
        "模板来源", "ibaotu", "ppt模板", "素材授权",
        "Lorem Ipsum", "simply dummy text", "printing and typeset",
        "PPT下载", "1ppt.com", "www.1ppt",
    ],
}

# ==================== 装饰文本配置 ====================
# 纯英文装饰大写关键词
EN_DECORATIVE_KEYWORDS: list[str] = [
    "BUSINESS", "WORK REPORT", "PPT TEMPLATE", "YOUR CONTENT",
    "THANK YOU", "CONTENTS", "CONTENT", "RESUME", "POSTGRADUATE",
    "PART.0", "TITLE HERE", "LOGO", "MARKETING PLAN",
    "A DREAM", "DREAM NEED", "SUBTITLE TEXT", "ENTER HERE",
    "PROJECT –", "YOUR MARKETING",
]

# 中文装饰/水印词
CN_DECORATIVE_PHRASES: list[str] = [
    "工作总结汇报", "百分比", "数值", "数据分布一", "数据分布二",
    "工作成绩展示", "未来工作规划", "X业务收入", "X业务支出",
    "设计业务", "线上促销",
]

# ==================== 动画推荐配置 ====================
ANIMATION_SPEED_MAP: dict[str, int] = {
    "slow": 1500,
    "med": 800,
    "fast": 400,
}
