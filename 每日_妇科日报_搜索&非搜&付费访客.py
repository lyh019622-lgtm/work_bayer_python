import pandas as pd
import re
import os

class TrafficDataTransformer:
    """流量数据转换器，用于合并相同来源渠道的数据"""
    
    def transform_traffic_data(self, input_df: pd.DataFrame) -> pd.DataFrame:
        """
        将底表2格式的数据转换为合并后的宽表格式
        """
        # 存储结果
        result_rows = []
        current_brand = ""
        current_store = ""
        current_data_block = []
        
        # 遍历原始数据行
        for idx, row in input_df.iterrows():
            first_cell = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
            
            # 1. 检查是否为品牌行
            if first_cell in ["优思明", "优思悦", "唯散宁"]:
                if current_brand and current_store and current_data_block:
                    self._process_data_block(current_brand, current_store, current_data_block, result_rows)
                
                current_brand = first_cell
                current_store = ""
                current_data_block = []
                continue
            
            # 2. 检查是否为店铺名称行
            if current_brand and not current_store and first_cell:
                if first_cell not in ["优思明", "优思悦", "唯散宁"]:
                    current_store = first_cell
                continue
            
            # 3. 检查是否为数据行
            if current_brand and current_store and first_cell:
                if first_cell == "暂无数据":
                    continue
                if len(row) >= 2 and not pd.isna(row.iloc[1]):
                    current_data_block.append(row)
        
        # 处理最后一个数据块
        if current_brand and current_store and current_data_block:
            self._process_data_block(current_brand, current_store, current_data_block, result_rows)
        
        columns = ["品牌", "店铺", "来源渠道", "指标1", "指标2", "指标3", "指标4", "指标5"]
        result_df = pd.DataFrame(result_rows, columns=columns)
        return result_df
    
    def _process_data_block(self, brand: str, store: str, data_rows: list, result_rows: list):
        source_dict = {}
        for row in data_rows:
            source = str(row.iloc[0]).strip()
            values = []
            for i in range(1, min(6, len(row))):
                val = row.iloc[i]
                if pd.isna(val):
                    values.append(0)
                else:
                    try:
                        clean_val = str(val).replace(',', '').strip()
                        values.append(float(clean_val) if clean_val != '' else 0)
                    except:
                        values.append(0)
            
            if source in source_dict:
                existing_values = source_dict[source]
                new_values = [existing_values[i] + values[i] for i in range(len(values))]
                source_dict[source] = new_values
            else:
                source_dict[source] = values
        
        for source, values in source_dict.items():
            while len(values) < 5:
                values.append(0)
            result_row = [brand, store, source] + values[:5]
            result_rows.append(result_row)

# ===================== 修改后的执行逻辑 =====================
if __name__ == "__main__":
    # 1. 设置输入和输出路径
    input_file = r"C:\Users\dell\Desktop\extractedDateable.xlsx"
    output_file = r"D:\rpa\搜索&非搜&付费访客\visitor0507.xlsx"
    
    try:
        # 读取原始数据
        df = pd.read_excel(input_file, header=None)
        print(f"成功读取文件: {input_file}")
        
        # 创建转换器并执行转换
        transformer = TrafficDataTransformer()
        print("正在转换数据...")
        result_df = transformer.transform_traffic_data(df)
        
        # 2. 保存转换后的结果
        result_df.to_excel(output_file, index=False)
        print(f"转换完成！结果已保存到: {output_file}")
        
        # ===================== 3. 清空原始输入文件内容 =====================
        print(f"\n正在清空输入文件内容: {input_file}")
        # 写入一个空的 DataFrame 以抹除原有数据，保留空文件
        pd.DataFrame().to_excel(input_file, index=False, header=False)
        print("✓ 原始文件内容已成功抹除。")
        # =================================================================
        
    except FileNotFoundError:
        print(f"错误：找不到输入文件 {input_file}")
    except Exception as e:
        print(f"转换过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()
