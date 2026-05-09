import pandas as pd
import re
from typing import List, Dict, Any
import os
import gc  # 强制垃圾回收，释放文件句柄

class DataTransformer:
    """数据转换器类，用于将原始数据底表转换为目标宽表格式"""
    
    def __init__(self, possible_brands: List[str] = None, filename: str = ""):
        self.possible_brands = possible_brands or ["优思明", "优思悦", "唯散宁", "安今益"]
        filename_lower = filename.lower()
        
        # 1. 品牌判断逻辑
        if filename_lower.startswith("yaz"):
            self.default_brand = "优思悦"
        elif filename_lower.startswith("visanne"):
            self.default_brand = "唯散宁"
        elif filename_lower.startswith("yasmin"):
            self.default_brand = "优思明"
        elif filename_lower.startswith("anjy"):
            self.default_brand = "安今益"
        else:
            self.default_brand = self.possible_brands[0] if self.possible_brands else "优思明"
        
        # 2. 日期解析逻辑：从文件名中提取4位数字(月日)，并固定年份为2026
        self.output_date = "2026/01/01" # 默认值
        date_match = re.search(r'(\d{2})(\d{2})', filename)
        if date_match:
            month, day = date_match.groups()
            self.output_date = f"2026/{month}/{day}"
            
        print(f"根据输入 '{filename}' 确定：品牌 = {self.default_brand}, 日期 = {self.output_date}")
        
    def transform_dataframe(self, input_df: pd.DataFrame) -> pd.DataFrame:
        result_rows = []
        current_block_info = {}
        current_sku_rows = []
        
        for idx, row in input_df.iterrows():
            if pd.isna(row.iloc[0]):
                continue
            first_cell = str(row.iloc[0]).strip()
            
            if self._is_store_name(first_cell):
                if "商品ID" in current_block_info and "店铺名称" not in current_block_info:
                    current_block_info["店铺名称"] = first_cell
                    self._process_block(current_block_info, current_sku_rows, result_rows)
                    current_block_info = {}
                    current_sku_rows = []
                continue
                
            if re.match(r'^\d+(\.\d+E\+\d+)?$', first_cell.replace(',', '')):
                item_id = self._convert_to_standard_id(first_cell)
                current_block_info["商品ID"] = item_id
                current_block_info["品牌名称"] = self.default_brand
                continue
                
            if first_cell in ["支付金额", "访客数", "支付用户", "转化率", "客单价"]:
                self._extract_metrics(first_cell, row, current_block_info)
                continue
                
            # 4. 检查SKU行：只要这一行不是“暂无数据”，其他的全部抓取
            if first_cell != "暂无数据":
                current_sku_rows.append(row)
        
        if current_block_info and current_sku_rows:
            self._process_block(current_block_info, current_sku_rows, result_rows)
        
        result_df = pd.DataFrame(result_rows, columns=self._get_output_columns())
        return self._fill_missing_values(result_df)

    def _is_store_name(self, text: str) -> bool:
        return text in ["阿里健康大药房", "阿里健康医药旗舰店", "健康福利社", "优思明官方旗舰店"]

    def _convert_to_standard_id(self, id_str: str) -> str:
        try:
            clean_id = id_str.replace(',', '')
            return str(int(float(clean_id))) if 'E' in clean_id.upper() else str(int(clean_id))
        except:
            return id_str

    def _extract_metrics(self, indicator_name: str, row: pd.Series, block_info: Dict[str, Any]):
        def clean_value(value):
            if pd.isna(value): return ""
            str_val = str(value).replace(',', '').replace('%', '').strip()
            try: return float(str_val)
            except: return str_val
        
        def format_percent(value):
            if pd.isna(value) or value == "": return ""
            return f"{value}%" if "%" not in str(value) else value

        mapping = {
            "支付金额": [("支付金额", 1), ("新客金额", 3), ("支付订单", 5)],
            "访客数": [("访客数量", 1), ("新访客数量", 3), ("老访客数量", 5)],
            "支付用户": [("支付用户", 1), ("支付用户_新客", 3), ("支付用户_老客", 5)],
            "转化率": [("转化率", 1), ("转化率_新客", 3), ("转化率_老客", 5)],
            "客单价": [("客单价", 1), ("客单价_新客", 3), ("客单价_老客", 5)]
        }
        
        for key, col_idx in mapping.get(indicator_name, []):
            val = clean_value(row.iloc[col_idx])
            block_info[key] = format_percent(val) if "转化率" in key else val

    def _process_block(self, block_info: Dict[str, Any], sku_rows: List[pd.Series], result_rows: List[Dict[str, Any]]):
        total_pieces = 0
        for sku_row in sku_rows:
            try:
                val = sku_row.iloc[3] if len(sku_row) > 3 else 0
                total_pieces += float(str(val).replace(',', '')) if not pd.isna(val) else 0
            except: pass
        
        for sku_row in sku_rows:
            def clean_sku(idx):
                if idx >= len(sku_row) or pd.isna(sku_row.iloc[idx]): return 0
                try: return float(str(sku_row.iloc[idx]).replace(',', ''))
                except: return 0

            result_rows.append({
                "日期": self.output_date,
                "店铺名称": block_info.get("店铺名称", ""),
                "品牌名称": block_info.get("品牌名称", ""),
                "商品ID": block_info.get("商品ID", ""),
                "支付件数": total_pieces,
                "支付金额": block_info.get("支付金额", 0),
                "新客金额": block_info.get("新客金额", 0),
                "支付订单": block_info.get("支付订单", 0),
                "访客数量": block_info.get("访客数量", 0),
                "新访客数量": block_info.get("新访客数量", 0),
                "老访客数量": block_info.get("老访客数量", 0),
                "免费流量": 0, "付费流量": 0,
                "客单价": block_info.get("客单价", 0),
                "客单价_新客": block_info.get("客单价_新客", 0),
                "客单价_老客": block_info.get("客单价_老客", 0),
                "转化率": block_info.get("转化率", "0%"),
                "转化率_新客": block_info.get("转化率_新客", "0%"),
                "转化率_老客": block_info.get("转化率_老客", "0%"),
                "30天复购率": 0, "90天复购率": 0, "180天复购率": 0,
                "支付用户": block_info.get("支付用户", 0),
                "支付用户_新客": block_info.get("支付用户_新客", 0),
                "支付用户_老客": block_info.get("支付用户_老客", 0),
                "sku_名称": str(sku_row.iloc[0]).strip(),
                "SKUID": str(sku_row.iloc[1]).strip() if len(sku_row) > 1 else "",
                "sku_规格简称": 0,
                "sku_支付金额": clean_sku(2),
                "sku_支付件数": clean_sku(3),
                "sku_支付用户数": clean_sku(4),
                "sku_支付新客数": clean_sku(5),
                "sku_支付老客户数": clean_sku(6),
                "sku_规格盒数": 0, "sku_规格": 0, "sku_规格_l": 0
            })

    def _get_output_columns(self) -> List[str]:
        return ["日期", "店铺名称", "品牌名称", "商品ID", "支付件数", "支付金额", "新客金额", "支付订单",
                "访客数量", "新访客数量", "老访客数量", "免费流量", "付费流量", "客单价", "客单价_新客", 
                "客单价_老客", "转化率", "转化率_新客", "转化率_老客", "30天复购率", "90天复购率", 
                "180天复购率", "支付用户", "支付用户_新客", "支付用户_老客", "sku_名称", "SKUID", 
                "sku_规格简称", "sku_支付金额", "sku_支付件数", "sku_支付用户数", "sku_支付新客数", 
                "sku_支付老客户数", "sku_规格盒数", "sku_规格", "sku_规格_l"]

    def _fill_missing_values(self, df: pd.DataFrame) -> pd.DataFrame:
        cols_to_fix = ["支付件数", "支付金额", "新客金额", "支付订单", "访客数量", "新访客数量", "老访客数量", 
                       "免费流量", "付费流量", "客单价", "客单价_新客", "客单价_老客", "支付用户", 
                       "支付用户_新客", "支付用户_老客", "sku_支付金额", "sku_支付件数", "sku_支付用户数", 
                       "sku_支付新客数", "sku_支付老客户数"]
        for col in cols_to_fix:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].apply(lambda x: 0 if x == "" or pd.isna(x) else x), errors='coerce').fillna(0)
        return df

if __name__ == "__main__":
    input_file = r"C:\Users\dell\Desktop\extractedDateable.xlsx"
    output_path = r"D:\rpa\GMV\total.xlsx"

    print("请输入参考文件名（例如 'visanne0416'，四品为yasmin,yaz,visanne,anjy，日期正常输入0606）用于品牌和日期判断：")
    filename_input = input(">>> ").strip() 
    
    try:
        # ===================== 修复点1：读取后立即释放文件 =====================
        df = pd.read_excel(input_file, header=None)
        gc.collect()  # 强制回收，关闭文件句柄

        transformer = DataTransformer(filename=filename_input)
        new_data = transformer.transform_dataframe(df)

        # 用完立即删除变量，强制释放
        del df
        gc.collect()
        
        # ===================== 修复点2：安全追加写入 =====================
        if os.path.exists(output_path):
            existing_df = pd.read_excel(output_path)
            final_df = pd.concat([existing_df, new_data], ignore_index=True)
            del existing_df  # 释放
            print(f"检测到已有文件，已追加 {len(new_data)} 行数据。")
        else:
            final_df = new_data
            print(f"未检测到已有文件，已创建新文件。")

        del new_data
        gc.collect()
        
        # 写入目标文件
        final_df.to_excel(output_path, index=False)
        del final_df
        gc.collect()
        print(f"✓ 数据已成功写入 {output_path}")
        
        # ===================== 修复点3：安全清空原始文件 =====================
        print(f"正在清空原始输入文件: {input_file}")
        empty_df = pd.DataFrame()
        empty_df.to_excel(input_file, index=False, header=False)
        del empty_df
        gc.collect()
        
        print("✓ 全部处理完成，文件已释放！")
        
    except Exception as e:
        print(f"错误: {str(e)}")
        gc.collect()
