import re
import pandas as pd
import akshare as ak
from datetime import datetime

def extract_stock_code(text):
    """
    从文本中提取股票代码，并补全为6位数字。
    支持格式：
        - '招商银行 (600036)'  → 600036
        - '中大力德-2896'      → 002896（补零到6位）
        - '双环传动-2472'      → 002472
        - '600036'            → 600036
    """
    # 先尝试匹配6位数字
    match = re.search(r'\b(\d{6})\b', str(text))
    if match:
        return match.group(1)
    # 再尝试匹配4位数字（常见于深市）
    match = re.search(r'\b(\d{4})\b', str(text))
    if match:
        code4 = match.group(1)
        # 补零到6位：深市通常为 000、002、300 开头，这里简单前面加 '00'
        return '00' + code4
    # 可选：匹配3位数字（如2050 → 002050）
    match = re.search(r'\b(\d{3})\b', str(text))
    if match:
        code3 = match.group(1)
        return '00' + code3
    return None

def get_realtime_price_akshare(code):
    """
    通过 AkShare 获取实时股价（最新价）
    返回 float 类型股价，失败返回 None
    """
    try:
        # 获取沪深京A股实时行情（东方财富数据源）
        df = ak.stock_zh_a_spot_em()
        # 筛选目标代码
        row = df[df['代码'] == code]
        if not row.empty:
            # 最新价列名通常为 '最新价'
            price = float(row['最新价'].values[0])
            return price
        else:
            print(f"  未找到代码 {code} 的行情数据")
            return None
    except Exception as e:
        print(f"  AkShare 请求失败: {e}")
        return None

def is_trading_time():
    """简单判断是否在交易时段（周一至周五 9:30-11:30, 13:00-15:00）"""
    now = datetime.now()
    if now.weekday() >= 5:  # 周六周日
        return False
    current_time = now.time()
    morning_start = datetime.strptime("09:30", "%H:%M").time()
    morning_end = datetime.strptime("11:30", "%H:%M").time()
    afternoon_start = datetime.strptime("13:00", "%H:%M").time()
    afternoon_end = datetime.strptime("15:00", "%H:%M").time()
    if (morning_start <= current_time <= morning_end) or (afternoon_start <= current_time <= afternoon_end):
        return True
    return False

def main():
    file_path = r'C:\XSR\实用信息\选股.xlsx'
    
    # 提示交易时间
    if not is_trading_time():
        print("警告：当前不是A股交易时段（9:30-11:30, 13:00-15:00），获取的价格可能为昨日收盘价或无效。")
    
    try:
        df = pd.read_excel(file_path, dtype={'公司名称 (代码)': str})
    except Exception as e:
        print(f"读取文件失败: {e}")
        return

    code_col = '公司名称 (代码)'
    price_col = '当前股价'
    if code_col not in df.columns:
        print(f"错误：Excel中缺少 '{code_col}' 列")
        return
    if price_col not in df.columns:
        df[price_col] = None

    # 遍历每一行
    for idx, row in df.iterrows():
        raw_text = row[code_col]
        stock_code = extract_stock_code(raw_text)
        if not stock_code:
            print(f"第 {idx+2} 行无法提取股票代码：'{raw_text}'，跳过")
            continue
        
        print(f"正在获取 {stock_code} ({raw_text}) ...")
        price = get_realtime_price_akshare(stock_code)
        if price is not None:
            df.at[idx, price_col] = price
            print(f"  -> 成功: {price}")
        else:
            print(f"  -> 获取失败，保持原值")

    # 保存回原文件
    try:
        df.to_excel(file_path, index=False)
        print(f"更新完成，已保存至 {file_path}")
    except Exception as e:
        print(f"保存文件失败: {e}")

if __name__ == '__main__':
    main()