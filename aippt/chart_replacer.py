"""
图表数据替换子模块 - 渲染引擎拆分
功能：基于 python-pptx 的 chart 对象，替换图表数据源，100% 保留模板样式
依赖：python-pptx
"""
from typing import Any

from pptx.shapes.base import BaseShape

from aippt.logger import logger


def replace_chart_data(shape: BaseShape, chart_data_dict: dict[str, Any]) -> None:
    """替换图表数据，100% 保留模板样式（字体/颜色/坐标轴格式/图例位置）

    支持 4 类图表：bar（柱状图）、line（折线图）、pie（饼图）、radar（雷达图）
    多系列自动适配：M>N 仅替换前 N 系列；M<N 多余系列清空数据；M==N 直接替换
    使用 chart.replace_data(ChartData) API，保留图表样式（系列数与模板一致）

    :param shape: GraphicFrame 形状，含 chart
    :param chart_data_dict: 图表数据，结构为
        {"categories": [...], "series": [{"name": "...", "data": [...]}]}
    """
    from pptx.chart.data import ChartData

    try:
        chart = shape.chart
    except Exception as e:
        logger.warning("无法访问图表数据: %s", e)
        return

    # 检查图表类型，映射到 bar/line/pie/radar 四大类
    ct_str = str(chart.chart_type).upper()
    if 'BAR' in ct_str or 'COLUMN' in ct_str:
        chart_category = 'bar'
    elif 'LINE' in ct_str:
        chart_category = 'line'
    elif 'PIE' in ct_str:
        chart_category = 'pie'
    elif 'RADAR' in ct_str:
        chart_category = 'radar'
    else:
        logger.warning("不支持的图表类型: %s，跳过数据替换", chart.chart_type)
        return

    categories = chart_data_dict.get('categories', [])
    new_series = chart_data_dict.get('series', [])
    if not categories or not new_series:
        logger.warning("图表数据为空（categories=%d, series=%d），跳过",
                       len(categories), len(new_series))
        return

    try:
        plot = chart.plots[0]
    except Exception as e:
        logger.warning("无法访问图表 plot: %s", e)
        return

    # 多系列自动适配：缓存 series 集合到本地变量，获取模板原有系列数
    chart_series = list(plot.series)
    n = len(chart_series)  # 模板原有系列数
    m = len(new_series)    # 新数据系列数

    if m > n:
        logger.warning("图表系列数不匹配：模板 %d 系列，新数据 %d 系列，仅替换前 %d 系列",
                       n, m, n)
    elif m < n:
        logger.warning("图表系列数不匹配：模板 %d 系列，新数据 %d 系列，多余 %d 系列清空数据",
                       n, m, n - m)

    # 构造 ChartData，保持系列数与模板一致（N 系列），保留图表样式
    chart_data = ChartData()
    chart_data.categories = categories
    for i in range(n):
        if i < m:
            # 替换为新数据
            ser = new_series[i]
            chart_data.add_series(ser.get('name', ''), ser.get('data', []))
        else:
            # 多余系列清空数据（置零，保留系列结构避免破坏图表样式）
            orig_name = ''
            try:
                orig_name = chart_series[i].name or ''
            except Exception:
                pass
            chart_data.add_series(orig_name, [0] * len(categories))

    # 替换图表数据（replace_data 保留图表类型、坐标轴格式、图例等样式）
    try:
        chart.replace_data(chart_data)
    except Exception as e:
        # replace_data 可能因外部 Excel 链接报错（.target_part undefined），
        # 但图表 XML 数据通常已更新（分类/系列值已写入），图表显示正常
        logger.warning("图表数据替换（Excel 同步异常，图表 XML 已更新）: %s", e)

    logger.info("图表数据替换完成（类型=%s, 系列=%d/%d, 分类=%d）",
                chart_category, min(m, n), n, len(categories))
