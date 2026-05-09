import pandas as pd

# 1. 设置文件路径
file_path = r'D:\AIwork\work\出库统计1.xlsx'

try:
    # 2. 读取Excel文件
    df = pd.read_excel(file_path)

    # 3. 统一商品名称（建立映射字典，方便维护）
    name_mapping = {
        "【原研药】拜耳 优思明 屈螺酮炔雌醇片 21片 短效避孕药 非避孕贴 女性口服": "优思明(小盒装)",
        "【原研药】拜耳 优思明 屈螺酮炔雌醇片 21片 短效避孕药 非紧急避孕药 女性口服": "优思明(小盒装)",
        "【原研药】拜耳 优思明 屈螺酮炔雌醇片": "优思明(小盒装)",
        "【原研药】优思悦 屈螺酮炔雌醇片(II)28片": "优思悦(小盒装)",
        "【原研药】唯散宁 地诺孕素片28片": "唯散宁(小盒装)",
        "【原研药】唯散宁 地诺孕素片84片": "唯散宁(大盒装)",
        "【原研药】安今益 雌二醇屈螺酮片 1mg:2mg*28片": "安今益(小盒装)",
        "【原研药】拜耳 优思明 屈螺酮炔雌醇片 21片*3板 非紧急避孕药 短效避孕药  女性口服": "优思明(大盒装)",
        "【原研药】优思悦 屈螺酮炔雌醇片(Ⅱ) 28片/板*3板": "优思悦(大盒装)",
    }
    
    # 使用 map 或 replace 替换名称
    df['商品名称'] = df['商品名称'].replace(name_mapping)

    # 4. 转换日期格式
    # 先转为 datetime，再格式化为 yy/mm/dd 字符串
    df['时间'] = pd.to_datetime(df['时间']).dt.strftime('%Y/%m/%d')

    # 5. 分组汇总
    # 核心：必须把 ['时间', '品牌', '商品名称'] 都放在 groupby 里，才能在结果中全部显示
    result = df.groupby(['时间', '品牌', '商品名称'])['昨日出库商品件数'].sum().reset_index()

    # 6. 查看结果或导出
    print("--- 汇总结果 ---")
    print(result.head()) # 打印前几行预览

    # 保存到新的 Excel
    output_path = r'D:\AIwork\work\汇总统计1_结果.xlsx'
    result.to_excel(output_path, index=False)
    print(f"\n处理完成！结果已保存至: {output_path}")

except Exception as e:
    print(f"处理过程中出现错误: {e}")