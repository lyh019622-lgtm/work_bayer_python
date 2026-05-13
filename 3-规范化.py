import pandas as pd
import os

def count_speakers(dialog):
    """计算对话中的说话人数"""
    speakers = set()
    lines = str(dialog).split('\n')
    for line in lines:
        line = line.strip()
        if line and ': ' in line:
            speaker_part = line.split(': ', 1)[0]
            if ' 20' in speaker_part:
                speaker = speaker_part.split(' 20')[0].strip()
            else:
                speaker = speaker_part.strip()
            if speaker and len(speaker) > 0 and not speaker.startswith('['):
                speakers.add(speaker)
    return len(speakers)

def remove_single_speaker_lines(input_file, output_file):
    print("正在处理文件...")

    try:
        xl = pd.ExcelFile(input_file)
        sheet_names = list(xl.sheet_names)
        xl.close()  # 立即关闭，释放文件锁

        all_sheets = {}

        for sheet_name in sheet_names:
            print(f"\n处理: {sheet_name}")
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            print(f"原始行数: {len(df)}")

            if sheet_name == '未检测到客服':
                all_sheets[sheet_name] = df
                print(f"保留 '未检测到客服' 数据")
                continue

            speakers_count = df['对话内容'].apply(count_speakers)
            valid_mask = speakers_count >= 2
            valid_df = df[valid_mask].copy()
            invalid_count = len(df) - len(valid_df)

            all_sheets[sheet_name] = valid_df
            print(f"有效行数（≥2人）: {len(valid_df)}")
            print(f"去除无效行数（≤1人）: {invalid_count}")

        # 保存处理后的文件（先写临时文件再覆盖，避免文件锁定）
        temp_file = output_file.replace('.xlsx', '_temp.xlsx')
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except:
                pass
        with pd.ExcelWriter(temp_file, engine='openpyxl') as writer:
            for sheet_name, data in all_sheets.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)

        try:
            os.replace(temp_file, output_file)
        except:
            if os.path.exists(output_file):
                os.remove(output_file)
            os.rename(temp_file, output_file)

        print(f"\n处理完成！结果保存到: {output_file}")
        return True

    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        return False

if __name__ == "__main__":
    input_file = input("请输入Excel文件路径: ").strip().strip('"')
    output_file = input_file

    if not os.path.exists(input_file):
        print("错误: 文件不存在")
    else:
        if remove_single_speaker_lines(input_file, output_file):
            print("\n处理成功")
