import pandas as pd
import numpy as np
import os

def transform_excel_data(input_file, output_file=None):
    """
    将原表格数据转换为品牌+店铺维度的数据
    """
    # 读取Excel文件
    df = pd.read_excel(input_file, header=None, dtype=str)
    
    # 数据清洗：去除逗号和百分号，转换为数值
    def clean_value(value):
        if pd.isna(value) or value == '':
            return None
        value = str(value).replace(',', '').replace('%', '')
        try:
            return float(value)
        except:
            return value
    
    # 品牌和店铺的映射关系
    brand_shop_mapping = [
        ("优思明", "阿里健康大药房", 1, 7),
        ("优思明", "阿里健康医药旗舰店", 10, 16),
        ("优思明", "优思明官方旗舰店", 19, 25),
        ("优思明", "健康福利社", 28, 34),
        ("优思悦", "阿里健康大药房", 37, 43),
        ("优思悦", "阿里健康医药旗舰店", 46, 52),
        ("优思悦", "优思明官方旗舰店", 55, 61),
        ("优思悦", "健康福利社", 64, 70),
        ("唯散宁", "阿里健康大药房", 73, 79),
        ("唯散宁", "阿里健康医药旗舰店", 82, 88),
        ("唯散宁", "优思明官方旗舰店", 91, 97),
        ("唯散宁", "健康福利社", 100, 106),
    ]
    
    metrics_mapping = {0: "访客数", 1: "浏览量", 2: "支付金额", 3: "支付人数", 4: "客单价", 5: "支付件数", 6: "支付订单数"}
    result_data = []
    
    for brand, shop, start_row, end_row in brand_shop_mapping:
        row_indices = list(range(start_row-1, end_row))
        row_data = {"品牌": brand, "店铺名": shop}
        
        for i, row_idx in enumerate(row_indices):
            metric_name = metrics_mapping.get(i, f"指标{i+1}")
            if row_idx < len(df):
                value = df.iloc[row_idx, 1] if len(df.columns) > 1 else None
                cleaned_value = clean_value(value)
                if cleaned_value is None or cleaned_value == '':
                    row_data[metric_name] = '---'
                else:
                    if isinstance(cleaned_value, (int, float)):
                        row_data[metric_name] = int(cleaned_value) if cleaned_value == int(cleaned_value) else cleaned_value
                    else:
                        row_data[metric_name] = cleaned_value
            else:
                row_data[metric_name] = '---'
        
        result_data.append(row_data)
    
    result_df = pd.DataFrame(result_data, columns=["品牌", "店铺名", "访客数", "浏览量", "支付金额", "支付人数", "客单价", "支付件数", "支付订单数"])
    
    if output_file:
        result_df.to_excel(output_file, index=False)
        print(f"数据已保存到: {output_file}")
    
    return result_df

# ===================== 主程序逻辑 =====================
if __name__ == "__main__":
    # 1. 重新定义的输入和输出路径
    input_file = r"D:\rpa\前端支付与用户\DateDaily.xlsx"
    output_file = r"D:\rpa\前端支付与用户\frontpay0507.xlsx"
    
    try:
        # 2. 执行转换
        print(f"正在读取并转换: {input_file}")
        result = transform_excel_data(input_file, output_file)
        print(f"转换成功！总记录数: {len(result)}")
        
        # 3. 转换成功后抹除内容
        print(f"\n正在清空原始输入文件内容...")
        # 写入一个空的 DataFrame 以抹除内容
        pd.DataFrame().to_excel(input_file, index=False, header=False)
        print(f"✓ {input_file} 内容已抹除 (保留空文件)。")
        
    except FileNotFoundError:
        print(f"错误: 找不到文件 {input_file}")
    except Exception as e:
        print(f"程序运行出错: {str(e)}")
        import traceback
        traceback.print_exc()
