import pandas as pd
import numpy as np
import re
import os
from datetime import datetime

# ============================================================================
# 第一部分：数据清洗（去重、删除海外店铺、价格清洗）
# ============================================================================

def clean_price(value):
    """清洗价格数据"""
    if pd.isna(value):
        return np.nan

    if isinstance(value, (int, float)):
        return float(value)

    value_str = str(value).strip()
    value_str = value_str.replace('\n', '').replace(' ', '')
    value_str = value_str.replace('¥', '').replace('￥', '')
    value_str = value_str.replace(',', '')

    if not value_str:
        return np.nan   

    try:
        return float(value_str)
    except ValueError:
        cleaned = re.sub(r'[^\d.-]', '', value_str)
        if cleaned:
            try:
                return float(cleaned)
            except ValueError:
                return np.nan
        else:
            return np.nan


def get_drug_name(file_name):
    """从文件名中提取药品名称"""
    file_lower = file_name.lower()

    if '优思明' in file_lower:
        return '优思明'
    elif '优思悦' in file_lower:
        return '优思悦'
    elif '唯散宁' in file_lower:
        return '唯散宁'
    else:
        return '未知药品'


def extract_info_from_filename(input_file):
    """从输入文件名中提取药品名称和日期"""
    file_name = os.path.basename(input_file)
    name_part, ext = os.path.splitext(file_name)

    # 提取药品名称
    drug_name = get_drug_name(name_part)

    # 提取日期信息，如0420
    import re
    date_match = re.search(r'\d{4}', name_part)
    if date_match:
        date_str = date_match.group()
    else:
        date_match = re.search(r'\d{2,}', name_part)
        date_str = date_match.group() if date_match else '0000'

    return drug_name, date_str, ext


def step1_data_cleaning(input_file):
    """第一步：数据清洗（不再输出中间文件）"""
    drug_name = get_drug_name(input_file)

    df = pd.read_excel(input_file)

    required_columns = ['到手_优惠价']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        return None, None
    else:
        original_count = len(df)

        link_column = '商品链接'

        if link_column in df.columns:
            duplicate_counts = df[link_column].value_counts()
            duplicate_links = duplicate_counts[duplicate_counts > 1]

            df = df.drop_duplicates(subset=link_column, keep='first')
            removed_count = original_count - len(df)

        shop_column = '店铺名称'

        if shop_column in df.columns:
            df_before_overseas = len(df)
            overseas_mask = df[shop_column].astype(str).str.contains('海外', case=False, na=False)
            overseas_count = overseas_mask.sum()

            df = df[~overseas_mask]
            removed_overseas_count = df_before_overseas - len(df)

        df['到手_优惠价_清洗后'] = df['到手_优惠价'].apply(clean_price)

        success_count = df['到手_优惠价_清洗后'].notna().sum()
        fail_count = df['到手_优惠价_清洗后'].isna().sum()

        df_sorted_all = df.sort_values('到手_优惠价_清洗后', ascending=True, na_position='last')

        price_stats = df_sorted_all['到手_优惠价_清洗后'].describe()

    return None, df_sorted_all


# ============================================================================
# 第二部分：KA店铺筛选
# ============================================================================

def filter_ka_stores(df_sorted_all, input_file, sheet_name, candidate_stores):
    """筛选包含候选店铺名的数据行"""

    df = df_sorted_all

    if '店铺名称' not in df.columns:
        return None

    original_count = len(df)

    condition = pd.Series(False, index=df.index)

    for store in candidate_stores:
        condition = condition | df['店铺名称'].astype(str).str.contains(store, case=False, na=False)

    filtered_df = df[condition]

    if len(filtered_df) == 0:
        return None

    # 保留必要的列（包括商品名称，因为后续处理需要）
    required_columns = ['关键词', '商品链接', '店铺名称', '商品名称', '到手_优惠价_清洗后']
    available_columns = [col for col in required_columns if col in filtered_df.columns]
    filtered_df = filtered_df[available_columns]

    # 动态生成输出文件名和路径
    drug_name, date_str, ext = extract_info_from_filename(input_file)

    # 确保输出目录存在
    output_dir = r"D:\work"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    output_file_path = os.path.join(output_dir, f"京东{drug_name}KA{date_str}{ext}")

    try:
        # 输出KA文件时只保留指定的四列，并在商品名称列后面添加来源
        output_columns = ['关键词', '商品链接', '店铺名称', '到手_优惠价_清洗后']
        output_available_columns = [col for col in output_columns if col in filtered_df.columns]
        output_df = filtered_df[output_available_columns].copy()
        # 在第二列添加来源列，值为京东
        output_df.insert(1, '来源', '京东')
        output_df.to_excel(output_file_path, index=False)
        return output_file_path
    except Exception as e:
        return None


# ============================================================================
# 第三部分：综合筛选（优思明/优思悦商品筛选+最低价格）
# ============================================================================

class KAStoreFilter:
    def __init__(self):
        self.candidate_stores = [
            "海王星辰", "益丰", "老百姓", "健之佳",
            "漱玉", "广药", "大参林", "国大", "佳康君安"
        ]
        self.price_column = '到手_优惠价_清洗后'

    def get_store_group(self, store_name):
        """根据店铺名称获取店铺分组"""
        if pd.isna(store_name):
            return None
        store_name_str = str(store_name).lower()
        for store_keyword in self.candidate_stores:
            if store_keyword.lower() in store_name_str:
                return store_keyword
        return None

    def filter_by_product_pattern(self, df, product_pattern):
        """按商品规格筛选"""
        if product_pattern == '21片':
            def contains_both_patterns_21(text):
                if pd.isna(text):
                    return False
                text_str = str(text)
                if '+' in text_str:
                    return False
                if '63片' in text_str:
                    return False
                has_21_per_piece = re.search(r'21\s*[/／]\s*片|21\s*片|21/片|21／片', text_str, re.IGNORECASE) is not None
                has_1_box = re.search(r'1\s*盒\s*装|1\s*盒|1盒装|1盒', text_str, re.IGNORECASE) is not None
                return has_21_per_piece and has_1_box
            condition = df['商品名称'].apply(contains_both_patterns_21)
        elif product_pattern == '28片':
            def contains_both_patterns_28(text):
                if pd.isna(text):
                    return False
                text_str = str(text)
                if '+' in text_str:
                    return False
                if '63片' in text_str:
                    return False
                has_28_per_piece = re.search(r'28\s*[/／]\s*片|28\s*片|28/片|28／片', text_str, re.IGNORECASE) is not None
                has_1_box = re.search(r'1\s*盒\s*装|1\s*盒|1盒装|1盒', text_str, re.IGNORECASE) is not None
                return has_28_per_piece and has_1_box
            condition = df['商品名称'].apply(contains_both_patterns_28)
        return df[condition]

    def find_lowest_price_per_store(self, df):
        """找出每家店的最低价格商品记录"""
        if len(df) == 0:
            return pd.DataFrame()

        if self.price_column not in df.columns:
            return pd.DataFrame()

        df_with_group = df.copy()
        df_with_group['店铺分组'] = df_with_group['店铺名称'].apply(self.get_store_group)

        df_with_group = df_with_group[df_with_group['店铺分组'].notna()]

        try:
            df_with_group[self.price_column] = df_with_group[self.price_column].astype(str).str.replace('[^\d.]', '', regex=True)
            df_with_group[self.price_column] = pd.to_numeric(df_with_group[self.price_column], errors='coerce')
        except Exception as e:
            return pd.DataFrame()

        lowest_price_records = []

        for store_keyword in self.candidate_stores:
            store_data = df_with_group[df_with_group['店铺分组'] == store_keyword]
            if len(store_data) > 0:
                store_data_valid = store_data[store_data[self.price_column].notna()]
                if len(store_data_valid) > 0:
                    min_price = store_data_valid[self.price_column].min()
                    min_price_records = store_data_valid[store_data_valid[self.price_column] == min_price]
                    if len(min_price_records) > 0:
                        lowest_price_record = min_price_records.iloc[0]
                        lowest_price_records.append(lowest_price_record)

        if lowest_price_records:
            result_df = pd.DataFrame(lowest_price_records)
            result_df = result_df.sort_values('店铺分组')
            return result_df
        else:
            return pd.DataFrame()

    def process_file(self, input_file_path, sheet_name):
        """处理Excel文件"""
        try:
            df = pd.read_excel(input_file_path, sheet_name=sheet_name)
        except Exception as e:
            return False, f"读取文件失败: {e}", None

        required_columns = ['店铺名称', '商品名称']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            error_msg = f"Excel文件中没有找到以下列: {', '.join(missing_columns)}"
            return False, error_msg, None

        store_condition = pd.Series(False, index=df.index)
        for store in self.candidate_stores:
            store_condition = store_condition | df['店铺名称'].astype(str).str.contains(store, case=False, na=False)

        store_filtered_df = df[store_condition]

        if len(store_filtered_df) == 0:
            return False, "没有筛选到任何店铺数据", None

        filtered_21_df = self.filter_by_product_pattern(store_filtered_df, '21片')

        filtered_28_df = self.filter_by_product_pattern(store_filtered_df, '28片')

        lowest_price_21_df = self.find_lowest_price_per_store(filtered_21_df)

        lowest_price_28_df = self.find_lowest_price_per_store(filtered_28_df)

        filtered_dfs = {
            '优思明筛选结果': filtered_21_df,
            '优思悦筛选结果': filtered_28_df,
            '优思明最低价格': lowest_price_21_df,
            '优思悦最低价格': lowest_price_28_df
        }

        return True, "处理完成", filtered_dfs

    def save_results(self, input_file_path, sheet_name, filtered_dfs):
        """保存结果到Excel文件"""
        if not filtered_dfs:
            print("错误: filtered_dfs 为空")
            return None

        # 动态生成输出文件名和路径
        drug_name, date_str, ext = extract_info_from_filename(input_file_path)

        # 确保输出目录存在
        output_dir = r"D:\work"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        output_file_path = os.path.join(output_dir, f"京东{drug_name}8KA{date_str}{ext}")

        try:
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                # 只保留指定的四列，并在第二列添加来源
                required_columns = ['关键词', '商品链接', '店铺名称', '到手_优惠价_清洗后']

                if len(filtered_dfs['优思明筛选结果']) > 0:
                    df = filtered_dfs['优思明筛选结果']
                    available_columns = [col for col in required_columns if col in df.columns]
                    output_df = df[available_columns].copy()
                    output_df.insert(1, '来源', '京东')
                    output_df.to_excel(writer, sheet_name='优思明筛选结果', index=False)
                else:
                    pd.DataFrame({'说明': ['无符合条件的数据']}).to_excel(writer, sheet_name='优思明筛选结果', index=False)

                if len(filtered_dfs['优思悦筛选结果']) > 0:
                    df = filtered_dfs['优思悦筛选结果']
                    available_columns = [col for col in required_columns if col in df.columns]
                    output_df = df[available_columns].copy()
                    output_df.insert(1, '来源', '京东')
                    output_df.to_excel(writer, sheet_name='优思悦筛选结果', index=False)
                else:
                    pd.DataFrame({'说明': ['无符合条件的数据']}).to_excel(writer, sheet_name='优思悦筛选结果', index=False)

                if len(filtered_dfs['优思明最低价格']) > 0:
                    df = filtered_dfs['优思明最低价格']
                    available_columns = [col for col in required_columns if col in df.columns]
                    output_df = df[available_columns].copy()
                    output_df.insert(1, '来源', '京东')
                    output_df.to_excel(writer, sheet_name='优思明最低价格', index=False)
                else:
                    pd.DataFrame({'说明': ['无符合条件的数据']}).to_excel(writer, sheet_name='优思明最低价格', index=False)

                if len(filtered_dfs['优思悦最低价格']) > 0:
                    df = filtered_dfs['优思悦最低价格']
                    available_columns = [col for col in required_columns if col in df.columns]
                    output_df = df[available_columns].copy()
                    output_df.insert(1, '来源', '京东')
                    output_df.to_excel(writer, sheet_name='优思悦最低价格', index=False)
                else:
                    pd.DataFrame({'说明': ['无符合条件的数据']}).to_excel(writer, sheet_name='优思悦最低价格', index=False)

            print(f"成功保存文件: {output_file_path}")
            return output_file_path

        except Exception as e:
            print(f"保存文件时出错: {e}")
            return None


# ============================================================================
# 主程序
# ============================================================================

def main():
    input_file = input("请输入Excel文件路径: ").strip()

    input_file = input_file.strip('"').strip("'")

    if not os.path.exists(input_file):
        return

    complete_output_file, df_sorted_all = step1_data_cleaning(input_file)

    if df_sorted_all is None:
        return

    candidate_stores = [
        "养和堂", "国大", "华氏", "好药师", "崇明第一医药", "第一医药商店", "得一大药房", "海王星辰",
        "益丰", "雷允上", "一心堂", "老百姓", "东飞药业", "人民大药房", "为了你健康药房", "昊邦",
        "白药大药房", "百姓仁药业", "医药新特药", "玉溪医药", "立之康", "靓桐", "健之佳", "淞茂",
        "滇西", "益尔健", "得云药业", "新特新", "仁和堂", "北域健康", "厚道医药", "康中福", "德医堂",
        "方圆", "惠生", "泽强", "达明堂", "东升天保堂", "日月七天", "金象", "东辽医药", "福春永",
        "保康", "江城大药房", "福聚康", "敖东", "一笑堂", "万家星火", "义立方", "光明医药", "养天和康赛",
        "华安大药房", "同仁大药房", "天诺", "恒爱", "房联", "新药特药", "春萍", "永新堂", "百姓润",
        "百年红星", "禹成", "邱健安康", "金牛大药房", "骅禧", "神农", "益和", "中联保健", "天一堂",
        "正道", "时珍", "恒爱", "光明万家", "万和堂", "信康", "汇丰", "长新", "义善堂", "三好名典",
        "安顺堂", "好医生", "嘉宝堂正红", "东升", "健之佳福利", "全泰堂", "华鼎", "圣杰", "天天康",
        "太极", "杏林", "正和祥", "万盛", "华安堂", "巴中怡和", "百欣堂", "芙蓉", "贝尔康", "高济",
        "齐力堂", "龙一", "大参林", "广恩堂", "心健", "寿亨堂", "泰和", "芝心", "大成家人", "普济文军",
        "万华", "健尔堂", "宁丰堂", "东升", "乐源堂", "华杏", "成毅康缘", "本草堂", "泉源堂", "百信",
        "科伦", "敬仁堂", "东升", "国健", "眉州药房", "安康", "众益堂", "宇豪", "天泰同安", "卫康",
        "京华壹零贰", "瑞澄", "德立信", "百姓乐", "惠仁堂", "美合泰", "丰原", "佳百姓", "华康", "四方百信",
        "国春堂", "国胜", "天星", "尚本堂", "广济", "百十信", "邻加医康复", "立方", "达嘉维康", "韵天",
        "高济敬贤堂", "鸿兴", "人民大药房", "仁和", "养生堂", "淮金益寿堂", "天平", "佳佳恒康", "新诚",
        "华巨百姓缘", "庆瑞祥堂", "博利", "中山", "元初", "易安堂", "赵泰和堂", "江南", "天地仁和",
        "杨博", "第一", "博爱", "曼迪新", "华仁堂", "同春", "花园", "康美", "仁德", "仁和堂", "立健",
        "信宏仁", "国民医保", "德信堂", "民泰", "燕喜堂", "百年康润", "益寿堂", "联众", "葆春堂", "誉天",
        "幸福人", "平民", "润康广场", "润生堂", "天天好", "百姓药业", "天和堂", "国康", "平嘉", "方成",
        "广联", "仁合堂", "天和堂景泰成", "众康", "漱玉", "德信行", "金通", "万民阳光", "中医世家",
        "惠仁", "健民", "五洲", "益民", "方正", "医保城", "同方", "国风", "紫光", "东方", "益源",
        "竹林", "仁人和", "同人众鑫", "同仁康", "国大万民", "国邦", "思迈乐", "百姓药房", "百汇",
        "荣华", "吉达惠康", "天诚", "新长城", "亚宝", "滋生堂", "永康堂", "中智", "佛心", "百姓堂",
        "顺德", "北京同仁堂", "德信行", "国控国大", "健丰", "大源堂", "五福堂", "健药", "品健",
        "好万家", "家之和", "恒青", "新合健", "春天一百", "深华", "爱心", "百源堂", "益康", "东莞国药",
        "邦健", "东兴堂", "健民", "宝芝林", "林芝参", "正和", "民信", "集和堂", "广药大药房", "广药", "康泽",
        "卫康", "百姓缘", "立丰", "民德", "二天堂", "万泽", "中港", "南北药行", "民心", "瑞草堂",
        "北湖", "益民", "健春堂", "一朝心阳", "新兴堂", "宝和堂", "康全济惠", "康全", "桂中",
        "百姓人家", "梧州百姓", "康是美", "康之源", "安康", "惠生堂", "新特", "塔城", "新特药",
        "库尔勒", "一心康达", "万佳康", "乐盈众康", "好又多", "康宁", "康泰东方", "普济堂", "汇泰",
        "济康", "百草堂", "维之康", "达康", "绿珠", "腾龙", "礼安", "上元堂", "健桥", "天颐", "百信",
        "恒玖信", "聚力堂", "天天乐", "京典", "江海", "普泽", "济生堂", "颐丹堂", "永泰", "大德生",
        "天健", "一盛", "万仁", "中诚", "人寿天", "金坛百姓", "恒心远", "建发", "恒泰", "百禾",
        "中健", "保庆堂", "广济", "恩华统一", "汉邦", "泰隆", "恒泰人民", "众成堂", "同康泰", "三品堂",
        "九州", "山禾", "天禾堂", "汇华强盛", "烨辉", "双鹤同德堂", "万家春", "同慈", "万家康", "万泽",
        "九鼎", "仁源生", "仁爱", "众佳康", "众药师", "健康家园", "华鑫源", "卫民", "同正", "启扬",
        "嘉伦光彩", "国瑞堂", "大众", "大德隆", "奇冠", "市民", "康济", "海邦", "海鹏", "润天",
        "百佳惠瑞丰", "百佳惠苏禾", "盐淮百信", "竹宸国瑞堂", "筑康", "联环健康", "舒健", "赛邦",
        "兴药", "为你想", "平民", "张氏", "德新元", "芝林", "洪医", "济生", "广济", "顺泰", "先锋",
        "正泰", "东方红", "百草堂", "健生源", "天一", "采芝灵", "开开心心", "粤海永熙堂", "雷允上",
        "南山", "天泰", "东康", "时代", "福仁堂", "龙腾盛世", "存仁堂", "华联", "江南", "心连心",
        "再康", "昌盛", "康每乐", "开心人", "汇仁堂", "洪兴", "黄庆仁栈", "尧乡", "旭康", "益民",
        "盛世华兴", "乐仁堂", "崇德", "百和一笑堂", "德医堂", "颈复康", "中亚", "仁泰", "唐人",
        "华佗", "康仁", "新兴", "狮城百姓", "益友", "石药", "神威", "诚仁堂", "新正", "百姓康",
        "百草堂", "天宇平价", "医药大厦", "宜生", "神农百草堂", "华为", "伊康", "信锐", "大参林百姓福",
        "恒生", "百姓健康", "隆泰仁", "天博", "仙鹤", "千年健", "益寿堂", "晟尚", "百氏康", "同心堂",
        "大参林健康", "林州大药房", "瑞康", "一心堂康健", "东宸", "东森", "丽康", "仁爱健民", "佐今明",
        "叮当智慧", "同和堂", "名药师", "大中元", "大参林", "天伦", "天康仁", "宜致", "张仲景", "恒信",
        "普生源", "杏一", "永乐", "润禾", "百事康健", "百家好一生", "中圆", "传仁堂", "华泰宜和堂",
        "发扬", "国信", "康丰", "康辉", "新越人", "花城", "美锐", "隆祥", "鲲鹏", "一杆秤", "宝神鹿",
        "心连心", "华辉", "华隆仁康", "晟和", "柏海宏济", "爱心", "蓝十字", "保元堂", "新东方",
        "胖东来", "九恒", "新向民", "禧康", "德良方", "福郎中", "保济新", "正大", "老百姓", "百姓缘",
        "大德", "四明", "正源", "彩虹", "易心堂", "为诚人家", "九洲", "华东", "武林", "百合楚济堂",
        "益万家", "邻家", "丽水便民", "健一行", "华通", "友禾", "国药", "大丛林", "天天好", "张济堂",
        "惠仁", "振和", "洪福堂", "海港", "华圣", "统一药房", "诚心", "长相宜", "长红", "震元",
        "一正", "修正堂", "宁康", "布衣", "延生堂", "张和堂", "长生堂", "平家健康", "华枫", "里肯",
        "英特一洲", "养生堂", "太和堂", "尖峰", "惠民", "民众", "广药晨菲", "鑫九州", "国康",
        "汉口大药房", "天济", "元昌", "为民天寿", "健康人", "叶开泰", "同仁美康", "聚荣璟", "佳源堂",
        "惠好", "隆泰益丰", "马应龙", "万众森林", "吴都", "好药师", "孝药", "宏泰", "宜草堂", "寿延堂",
        "心连心", "晓琳", "普恩堂", "柴氏荣盛", "用心人", "百佳和", "益生天济", "济公", "高济明联",
        "黎民", "福元堂", "普天独活", "康泰", "神州", "远志", "养天和", "健康南门", "楚济堂", "丹桂园",
        "九芝堂", "千金", "恒康", "津湘", "益汉", "诚益信", "诺舟", "达嘉维康", "雅馨", "佛慈",
        "新生", "德生堂", "亚欣", "众友", "心连心", "快康", "至仁同济", "祝强", "鹭燕", "康佰家",
        "东南", "国大", "榕参", "宜又佳", "恒生", "惠百姓", "新永惠", "永惠", "海华", "康利达",
        "扬祖惠民", "民心", "燕煌", "百泰", "聚善堂", "聚芝林", "一品", "一树", "正和祥", "健一生",
        "老天祥", "振安", "百盛新药特药", "泽康", "大厦", "安东新特", "世纪", "亚安", "新立群",
        "阳光", "敖东", "漱玉", "康源", "东北", "利安德", "九宇康展", "博爱", "康芝悦", "福聚和",
        "百年东和堂", "万家康", "五洲通", "人民康泰", "北药家", "华诺", "博大维康", "嘉和", "天一优市美",
        "天士力", "天益堂", "巨力", "康益堂", "建联", "心向民", "成大方圆", "新益康", "施福堂",
        "旺福营盘", "沛芝堂", "一明", "百和堂", "太星", "福缘堂", "襄元堂", "雪松", "安康", "益春",
        "弘大惠诚", "华星堂", "万鑫", "和平", "唐氏", "万和", "万家燕", "桐君阁", "泉源堂", "润药渝康",
        "西部医药商城", "鑫斛", "健生", "中和堂", "同心", "广济堂", "派林", "德厚丰", "怡康", "欣康",
        "泰生", "东亚", "乐榕融", "乡亲", "京兆", "众信", "五星", "孙思邈", "博爱", "百姓乐", "康德",
        "朋成", "百家惠", "美致惠", "广济", "长健", "桢州", "九康", "新富康", "万益", "新绿洲",
        "华伟", "七星堂", "人民同泰", "北秀", "宏腾", "宝丰", "建国", "海晖", "苮丰", "灵峰",
        "福斯特", "三元至盛", "鑫世一", "利达", "天利", "一辰", "华辰", "安李泰", "大众平安", "安泰", "齐泰", "佳康君安"
    ]

    # 不再需要从Excel读取，直接使用清洗后的DataFrame
    sheet_name = 'Sheet1'

    ka_output_file = filter_ka_stores(df_sorted_all, input_file, sheet_name, candidate_stores)

    if ka_output_file:
        print(f"KA文件已生成: {ka_output_file}")

    filter_tool = KAStoreFilter()

    # 直接使用 df_sorted_all 进行后续处理（包含所有必需的列）
    # 创建临时文件用于后续处理，保留所有必需的列
    temp_file = r"d:\AIwork\work\temp_ka.xlsx"

    # 筛选KA店铺数据（不保存到文件，只用于后续处理）
    if '店铺名称' in df_sorted_all.columns:
        condition = pd.Series(False, index=df_sorted_all.index)
        for store in candidate_stores:
            condition = condition | df_sorted_all['店铺名称'].astype(str).str.contains(store, case=False, na=False)
        ka_df = df_sorted_all[condition]
    else:
        ka_df = df_sorted_all

    ka_df.to_excel(temp_file, index=False)

    # 使用 with 语句确保文件正确关闭
    with pd.ExcelFile(temp_file) as xl:
        sheet_names = xl.sheet_names
        sheet_name = sheet_names[0]

        success, message, filtered_dfs = filter_tool.process_file(temp_file, sheet_name)

        if success and filtered_dfs:
            final_output_file = filter_tool.save_results(input_file, sheet_name, filtered_dfs)

            if final_output_file:
                print(f"8KA文件已生成: {final_output_file}")
            else:
                print("8KA文件生成失败")
        else:
            print(f"处理失败: {message}")

    # 删除临时文件
    if os.path.exists(temp_file):
        try:
            os.remove(temp_file)
        except:
            pass  # 如果删除失败，忽略错误


if __name__ == "__main__":
    main()
