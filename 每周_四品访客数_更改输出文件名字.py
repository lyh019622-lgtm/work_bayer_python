import pandas as pd

# ===================== 你的文件路径 =====================
extractdate_file = r"D:\rpa\优思明每日访客数\商品明细.xlsx"
content_file = r"D:\rpa\优思明每日访客数\商品360.xlsx"

# 1. 加载数据
# 由于你提到第一行可以去掉，这里用 skiprows=1
df_extract = pd.read_excel(extractdate_file, header=None)
df_content = pd.read_excel(content_file, header=None)

# 商品明细表：恢复 支付金额 列
df_extract.columns = ['商品名称', '商品ID', '支付金额']

# 商品360表：你需要的 8 列完整字段
df_content.columns = [
    '流量来源',
    '访客数',
    '浏览量',
    '加购人数',
    '支付人数',
    '支付转化率',
    '直接引导支付人数'
]

# 强制 ID 为字符串
df_extract['商品ID'] = df_extract['商品ID'].astype(str).str.strip()

# 2. 处理 content 表的 ID 向上填充逻辑
def is_id(val):
    s = str(val).strip()
    return s if (s.isdigit() and len(s) > 8) else None

df_content['商品ID'] = df_content['流量来源'].apply(is_id)
df_content['商品ID'] = df_content['商品ID'].bfill()

# 清理 content 表：删掉纯ID行，只保留数据行
df_content_clean = df_content[df_content['流量来源'].astype(str).str.strip() != df_content['商品ID']].copy()

# 3. 合并数据
merged_df = pd.merge(df_extract, df_content_clean, on='商品ID', how='left')

# 4. 重复商品信息置空（只保留第一行）
group_cols = ['商品名称', '商品ID', '支付金额']
merged_df.loc[merged_df.duplicated(subset=group_cols), group_cols] = ""

# 5. 保存结果
output_path = r"D:\rpa\优思明每日访客数\唯散宁访客数0501-0507.xlsx"
merged_df.to_excel(output_path, index=False)

print(f"处理完成！结果已保存至: {output_path}")
# 在 print(f"处理完成...") 之后添加：

for file_path in [extractdate_file, content_file]:
    # 创建一个空的 DataFrame 并覆盖原文件
    pd.DataFrame().to_excel(file_path, index=False)

print("原始文件内容已清空（保留空文件）。")
