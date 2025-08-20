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
    },
    "pingan_bank": {
        "start_row": 0,  # 第 1 行
        "amount_column": "D",  # 利息金额列
        "balance_column": "G",  # 余额列
        "description_column": "J",  # 内容列
        "keyword": "结息",
        "bank_name": "平安银行"
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
        return False, None

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
                elif "平安银行" in filename:
                    bank_type = "pingan_bank"
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
        df_result.to_excel(output_file, index=False, startrow=0, engine='openpyxl')
        update_log(f"数据已保存到 {output_file}")
        return True, output_file
    else:
        update_log("未找到任何利息数据")
        return False, None

def update_with_match_data(output_file, match_file_path, update_log):
    """将生成的 Excel 文件与匹配的 Excel 文件进行合并，添加托管账户和入息账套列，按银行 sheet 匹配"""
    try:
        # 读取生成的 Excel
        df_generated = pd.read_excel(output_file)

        # 初始化新列
        df_generated['托管账户'] = None
        df_generated['入息账套'] = None

        # 获取独特银行列表
        unique_banks = df_generated['银行'].unique()
        update_log(f"找到的银行列表: {unique_banks.tolist()}")

        for bank in unique_banks:
            if pd.isna(bank):
                update_log(f"警告: 跳过银行为空的行")
                continue

            # 读取匹配 Excel 的对应 sheet，将托管账户和产品编号强制为字符串
            try:
                df_match = pd.read_excel(
                    match_file_path,
                    sheet_name=bank,
                    header=0,
                    dtype={"产品编号": str, "托管账户": str, "入息账套": str}  # 强制为字符串，避免数值精度丢失
                )
                update_log(f"成功读取 '{bank}' sheet，行数: {len(df_match)}")
            except ValueError as ve:
                update_log(f"错误: 未找到 '{bank}' 工作表：{str(ve)}")
                continue
            except Exception as ex:
                update_log(f"读取 '{bank}' 工作表失败：{str(ex)}")
                continue

            # 清理列名中的空格
            df_match.columns = df_match.columns.str.strip()

            # 检查匹配 sheet 的列（假设 A:产品编号, C:托管账户, D:入息账套，使用列名）
            required_cols = ["产品编号", "托管账户", "入息账套"]
            if not all(col in df_match.columns for col in required_cols):
                update_log(f"错误: '{bank}' sheet 缺少必要列（产品编号、托管账户、入息账套），实际列: {df_match.columns.tolist()}")
                continue

            # 确保产品编号和托管账户为字符串并去除空格
            df_match['产品编号'] = df_match['产品编号'].astype(str).str.strip()
            df_match['托管账户'] = df_match['托管账户'].astype(str).str.strip()
            df_generated['产品编号'] = df_generated['产品编号'].astype(str).str.strip()

            # 调试：记录匹配 sheet 的产品编号和托管账户的前几行
            update_log(f"'{bank}' sheet 前5行数据 (产品编号, 托管账户, 入息账套):")
            for i, row in df_match.head(5)[['产品编号', '托管账户', '入息账套']].iterrows():
                update_log(f"  产品编号: {row['产品编号']}, 托管账户: {row['托管账户']}, 入息账套: {row['入息账套']}")

            # 过滤生成的 df 中该银行的行
            df_bank = df_generated[df_generated['银行'] == bank]
            if df_bank.empty:
                update_log(f"警告: 在生成的 Excel 中未找到银行 '{bank}' 的数据")
                continue

            # 匹配
            df_merged = df_bank.merge(
                df_match[["产品编号", "托管账户", "入息账套"]],
                left_on="产品编号",  # 生成的 B列: 产品编号
                right_on="产品编号",
                how="left",
                suffixes=('', '_new')
            )

            # 调试：记录 df_merged 的列名
            update_log(f"银行 '{bank}' df_merged columns: {df_merged.columns.tolist()}")

            # 调试：记录匹配结果
            update_log(f"银行 '{bank}' 匹配结果，匹配行数: {len(df_merged)}")
            for idx, row in df_merged.head(5).iterrows():
                update_log(f"  产品编号: {row['产品编号']}, 托管账户: {row.get('托管账户', 'N/A')}, 入息账套: {row.get('入息账套', 'N/A')}")

            # 更新原 df_generated 的对应行
            for idx in df_merged.index:
                original_idx = df_bank.index[df_bank['产品编号'] == df_merged.at[idx, '产品编号']].tolist()[0]  # 假设产品编号唯一
                df_generated.at[original_idx, '托管账户'] = df_merged.at[idx, '托管账户_new'] if '托管账户_new' in df_merged.columns else df_merged.at[idx, '托管账户']
                df_generated.at[original_idx, '入息账套'] = df_merged.at[idx, '入息账套_new'] if '入息账套_new' in df_merged.columns else df_merged.at[idx, '入息账套']

        # 重新排序列
        # 假设原列: A:日期, B:产品编号, C:产品名称, D:利息金额, E:余额, F:银行
        # 新: G:托管账户, H:入息账套
        df_generated = df_generated[[
            "日期", "产品编号", "产品名称", "利息金额", "余额", "银行",
            "托管账户", "入息账套"
        ]]
        # 保存回原文件（覆盖），使用 openpyxl 引擎，确保字符串格式
        df_generated.to_excel(output_file, index=False, startrow=0, engine='openpyxl')
        update_log(f"匹配完成，已更新 {output_file} 添加托管账户和入息账套列")

        # 第二步：基于匹配完成的 df_generated 生成“其他交易导入模板” sheet
        template_data = []
        for _, row in df_generated.iterrows():
            # 第一条记录
            template_data.append({
                "业务日期": row["日期"],
                "记账日期": row["日期"],
                "账套号": row["入息账套"],
                "摘要": f"收到银行利息收入（{row['产品编号']}-{row['产品名称']}）",
                "科目代码": "1002",
                "借贷方向（0借，1贷）": "0",
                "金额": row["利息金额"],
                "银行/资金账号": row["托管账户"],
                "受益凭证号": None,
                "证券代码": None,
                "凭证号": None,
                "分录号": "1",
                "备注": row["产品名称"],
                "是否支付日（0是，1否）": "0",
                "是否全额冲销（0全部，1部分）": "0"
            })
            # 第二条记录
            template_data.append({
                "业务日期": row["日期"],
                "记账日期": row["日期"],
                "账套号": row["入息账套"],
                "摘要": f"收到银行利息收入（{row['产品编号']}-{row['产品名称']}）",
                "科目代码": "601101",
                "借贷方向（0借，1贷）": "1",
                "金额": row["利息金额"],
                "银行/资金账号": row["托管账户"],
                "受益凭证号": None,
                "证券代码": None,
                "凭证号": None,
                "分录号": "2",
                "备注": row["产品名称"],
                "是否支付日（0是，1否）": "0",
                "是否全额冲销（0全部，1部分）": "0"
            })

        # 创建 DataFrame
        df_template = pd.DataFrame(template_data)
        update_log(f"生成‘其他交易导入模板’ sheet，包含 {len(df_template)} 条记录（{len(df_generated)} 条原始记录，每条生成 2 条）")

        # 保存到 Excel，指定 sheet 顺序
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 先写入“其他交易导入模板” sheet
            df_template.to_excel(writer, sheet_name="其他交易导入模板", index=False, startrow=0)
            # 再写入“Sheet1”
            df_generated.to_excel(writer, sheet_name="Sheet1", index=False, startrow=0)

        update_log(f"匹配完成，已更新 {output_file}，包含 sheet：其他交易导入模板（第一）、Sheet1（第二）")
        return True
    except Exception as ex:
        update_log(f"匹配失败：{str(ex)}")
        return False