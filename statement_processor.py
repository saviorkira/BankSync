import pandas as pd
import os
from datetime import datetime
import re
import warnings

# 忽略 openpyxl 的默认样式警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 定义银行规则：数据起始行、列位置、关键词和银行名称
bank_rules = {
    "hangzhou_bank": {
        "start_row": 1,  # 第 2 行（索引从 0 开始）
        "amount_column": "C",  # 利息金额列
        "balance_column": "E",  # 余额列
        "description_column": "I",  # 摘要内容列
        "keyword": "应付利息",  # 利息相关关键词
        "bank_name": "杭州银行"
    },
    "ningbo_bank": {
        "start_row": 7,  # 第 8 行
        "amount_column": "E",  # 利息金额列
        "balance_column": "G",  # 余额列
        "description_column": "I",  # 内容列
        "keyword": "利息收入",
        "bank_name": "宁波银行"
    },
    "zhongxin_bank": {
        "start_row": 0,  # 第 1 行
        "amount_column": "G",  # 利息金额列
        "balance_column": "H",  # 余额列
        "description_column": "L",  # 内容列
        "keyword": "批量结息",
        "bank_name": "中信银行"
    }
}

def get_quarterly_settlement_date(target_month):
    """将 YYYY-MM 转换为对应的季度结息日 YYYYMMDD，仅支持 3、6、9、12 月"""
    try:
        year, month = map(int, target_month.split('-'))
        # 季度结息日映射
        quarter_months = {3: '03', 6: '06', 9: '09', 12: '12'}
        # 检查月份是否为结息月
        if month not in quarter_months:
            return None  # 非结息月（1、2、4、5、7、8、10、11）返回 None
        return f"{year}{quarter_months[month]}21"
    except ValueError:
        return None

def extract_interest_data(file_path, bank_type, target_month):
    """从指定银行的 Excel 文件中提取利息数据"""
    try:
        # 获取银行规则
        rule = bank_rules.get(bank_type)
        if not rule:
            return None, f"未知银行类型: {bank_type}"

        # 读取 Excel 文件，跳过前几行
        df = pd.read_excel(file_path, skiprows=rule["start_row"])
        if df.empty:
            return None, f"文件 {file_path} 为空 (行数: 0)"

        # 获取列索引
        amount_col = rule["amount_column"]
        balance_col = rule["balance_column"]
        desc_col = rule["description_column"]
        keyword = rule["keyword"]
        bank_name = rule["bank_name"]

        amount_idx = ord(amount_col.upper()) - ord('A')
        balance_idx = ord(balance_col.upper()) - ord('A')
        desc_idx = ord(desc_col.upper()) - ord('A')

        # 确保列存在
        if amount_idx >= len(df.columns) or balance_idx >= len(df.columns) or desc_idx >= len(df.columns):
            return None, f"文件 {file_path} 列索引超出范围 (总列数: {len(df.columns)})"

        # 查找包含关键词的行
        interest_rows = df[df.iloc[:, desc_idx].astype(str).str.contains(keyword, na=False)]
        if interest_rows.empty:
            desc_content = df.iloc[:, desc_idx].astype(str).unique()[:5].tolist()
            return None, f"文件 {file_path} 未找到包含 '{keyword}' 的数据 (行数: {len(df)}, 描述列内容: {desc_content})"

        # 从文件名提取信息
        filename = os.path.basename(file_path)
        parts = filename.split('_')
        if len(parts) < 4:
            return None, f"文件 {file_path} 名称格式错误"

        product_id = parts[0]  # 24196, 24311 等
        product_name = parts[1]  # 重庆信托·余杭发展集合资金信托计划

        # 计算季度结息日
        settlement_date = get_quarterly_settlement_date(target_month)
        if not settlement_date:
            return None, f"无效的月份: {target_month}，仅支持 3、6、9、12 月"

        # 提取数据
        data = []
        for idx, row in interest_rows.iterrows():
            date = row.iloc[0] if pd.notna(row.iloc[0]) else None  # 假设日期在第一列
            amount = row.iloc[amount_idx] if pd.notna(row.iloc[amount_idx]) else 0
            balance = row.iloc[balance_idx] if pd.notna(row.iloc[balance_idx]) else 0
            data.append({
                "日期": settlement_date,  # 使用季度结息日 YYYYMMDD
                "产品编号": product_id,
                "产品名称": product_name,
                "利息金额": amount,
                "余额": balance,
                "银行": bank_name
            })
        return data, None
    except Exception as e:
        return None, f"处理 {file_path} 失败: {e}"

def process_bank_statements(root_dir, target_month, export_path, update_log, dropdown_value="全部"):
    """处理所有银行流水文件，生成统一 Excel"""
    all_data = []

    # 解析目标月份
    try:
        target_date = datetime.strptime(target_month + "-01", "%Y-%m-%d")
        target_year_month = target_date.strftime("%Y-%m")
    except ValueError:
        update_log("错误: 月份格式不正确，应为 YYYY-MM")
        return False

    # 使用正则表达式匹配日期范围
    date_pattern = re.compile(r'(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})\.xlsx$')

    # 遍历目录下的 Excel 文件
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith(".xlsx") and not filename.startswith('~$'):  # 跳过临时文件
                file_path = os.path.join(dirpath, filename)
                # 判断银行类型
                bank_type = None
                if "杭州银行流水" in filename:
                    bank_type = "hangzhou_bank"
                elif "宁波银行" in filename:
                    bank_type = "ningbo_bank"
                elif "中信银行" in filename:
                    bank_type = "zhongxin_bank"
                else:
                    update_log(f"跳过文件 {file_path}: 无法识别银行类型")
                    continue

                # 提取日期范围
                match = date_pattern.search(filename)
                if not match:
                    update_log(f"跳过文件 {file_path}: 日期范围格式错误 (无法匹配日期)")
                    continue

                start_date, end_date = match.groups()
                update_log(f"解析文件 {file_path}: 日期范围 {start_date} 至 {end_date}")

                # 判断日期范围是否包含目标月份
                try:
                    start = datetime.strptime(start_date, "%Y-%m-%d")
                    end = datetime.strptime(end_date, "%Y-%m-%d")
                    start_ym = start.strftime("%Y-%m")
                    end_ym = end.strftime("%Y-%m")
                    # 检查目标月份是否在日期范围内（包含开始或结束月份）
                    if not (start_ym <= target_year_month <= end_ym):
                        update_log(
                            f"跳过文件 {file_path}: 日期范围 {start_date} 至 {end_date} 不包含 {target_year_month}")
                        continue
                except ValueError as e:
                    update_log(f"跳过文件 {file_path}: 日期格式错误 ({e})")
                    continue

                update_log(f"处理文件: {file_path}")
                data, error = extract_interest_data(file_path, bank_type, target_month)
                if error:
                    update_log(error)
                elif data:
                    all_data.extend(data)

    # 创建 DataFrame
    if all_data:
        df_result = pd.DataFrame(all_data)

        # 转换为标准格式
        df_result["利息金额"] = pd.to_numeric(df_result["利息金额"], errors='coerce')
        df_result["余额"] = pd.to_numeric(df_result["余额"], errors='coerce')
        # 不再对“日期”列应用 pd.to_datetime，保持 YYYYMMDD 字符串格式

        # 删除重复数据（日期、产品编号、产品名称、利息金额 均相同）
        df_result = df_result.drop_duplicates(subset=["日期", "产品编号", "产品名称", "利息金额"], keep='first')

        # 按日期排序
        df_result = df_result.sort_values("日期")

        # 生成动态文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{target_month}_{dropdown_value}_{timestamp}.xlsx"
        output_file = os.path.join(export_path, output_filename)

        # 保存到 Excel，从第 1 行开始
        df_result.to_excel(output_file, index=False, startrow=0)
        update_log(f"数据已保存到 {output_file}")
        return True
    else:
        update_log("未找到任何利息数据")
        return False