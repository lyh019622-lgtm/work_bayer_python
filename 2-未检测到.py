import pandas as pd
import re
import os

def check_no_agent_conversations(input_file, output_sheet='未检测到客服'):
    print("正在检查客服ID为空的行...")

    try:
        # 读取整合后的Excel文件
        xl = pd.ExcelFile(input_file)

        all_no_agent_data = []

        for sheet_name in xl.sheet_names:
            # 跳过目标输出sheet，防止重复处理
            if sheet_name == output_sheet:
                continue

            # 读取每个sheet的数据
            df = pd.read_excel(input_file, sheet_name=sheet_name)

            print(f"\n处理: {sheet_name}")
            print(f"原始行数: {len(df)}")

            # 找到客服姓名为空的行
            no_agent_mask = df['客服姓名'].isna() | (df['客服姓名'].astype(str).str.strip() == '')
            no_agent_df = df[no_agent_mask].copy()

            print(f"客服ID为空的行数: {len(no_agent_df)}")

            # 检查这些行的对话内容是否包含两个人的对话
            valid_no_agent_data = []

            for idx, row in no_agent_df.iterrows():
                dialog_content = str(row['对话内容'])

                # 提取所有说话人
                speakers = []
                lines = dialog_content.split('\n')

                for line in lines:
                    line = line.strip()
                    if line and ': ' in line:
                        # 提取说话人（格式：说话人 时间: 内容）
                        speaker_part = line.split(': ', 1)[0]
                        # 去除时间戳
                        if ' 20' in speaker_part:
                            speaker = speaker_part.split(' 20')[0].strip()
                        else:
                            speaker = speaker_part.strip()

                        if speaker and speaker not in speakers and len(speaker) > 0 and not speaker.startswith('['):
                            speakers.append(speaker)

                # 如果有2个或更多不同的说话人，这可能是一个有效的对话
                if len(speakers) >= 2:
                    valid_no_agent_data.append(row)
                    print(f"找到有效对话: 用户ID={row['用户ID']}, 包含说话人: {speakers}")

            if valid_no_agent_data:
                print(f"找到 {len(valid_no_agent_data)} 个有效的对话")
                all_no_agent_data.extend(valid_no_agent_data)

    except Exception as e:
        print(f"错误: {e}")
        return False

    if not all_no_agent_data:
        print("没有找到符合条件的有效对话")
        return True

    # 将结果写入新的sheet中
    try:
        # 使用ExcelWriter保留原有数据
        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a') as writer:
            # 删除已存在的输出sheet（如果存在）
            if output_sheet in writer.book.sheetnames:
                idx = writer.book.sheetnames.index(output_sheet)
                writer.book.remove(writer.book.worksheets[idx])

            # 写入新数据
            pd.DataFrame(all_no_agent_data).to_excel(writer, sheet_name=output_sheet, index=False)

        print(f"\n成功保存到: {input_file} -> sheet '{output_sheet}'")
        print(f"总有效行数: {len(all_no_agent_data)}")

        return True

    except Exception as e:
        print(f"写入文件错误: {e}")
        return False

if __name__ == "__main__":
    input_file = input("请输入Excel文件路径: ").strip().strip('"')
    output_sheet = "未检测到客服"

    if not os.path.exists(input_file):
        print("错误: 文件不存在")
    else:
        if check_no_agent_conversations(input_file, output_sheet):
            print("\n检查完成")
