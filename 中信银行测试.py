from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


def split_date_ranges(start_date_str, end_date_str, max_months=3):
    """
    将日期范围拆分为多个不超过 max_months 个月的子范围。

    参数:
        start_date_str (str): 开始日期，格式为 'YYYY-MM-DD'
        end_date_str (str): 结束日期，格式为 'YYYY-MM-DD'
        max_months (int): 最大月份跨度，默认为3个月

    返回:
        List[Tuple[str, str]]: 包含多个 (start_date, end_date) 的列表，日期格式为 'YYYY-MM-DD'
    """
    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
    except ValueError as e:
        raise ValueError(f"日期格式错误，应为 YYYY-MM-DD: {str(e)}")

    if start_date > end_date:
        raise ValueError("开始日期不能晚于结束日期")

    date_ranges = []
    current_start = start_date

    while current_start <= end_date:
        # 计算当前范围的结束日期：开始日期 + max_months 月 - 1 天
        current_end = min(
            current_start + relativedelta(months=max_months) - timedelta(days=1),
            end_date
        )
        # 添加当前范围
        date_ranges.append((
            current_start.strftime("%Y-%m-%d"),
            current_end.strftime("%Y-%m-%d")
        ))
        # 更新下一次循环的开始日期
        current_start = current_end + timedelta(days=1)

    return date_ranges

ranges = split_date_ranges("2025-07-04", "2025-11-15")
print(ranges)