import pandas as pd
import re
import os


def process_sheet(file_path, sheet_name):
    print(f"\n{'='*50}\n处理: {sheet_name}\n{'='*50}")

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"读取 {len(df)} 行数据")
    except Exception as e:
        print(f"读取失败: {e}")
        return []

    chat_data = df.iloc[:, 0].tolist()
    conversations = []
    current_conversation = None

    speaker_pattern = re.compile(r'^(.*?)\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})$')
    agent_keywords = ['拜耳-', 'jimi_', 'sqjk']

    has_started = False
    for line in chat_data:
        if pd.isna(line):
            continue

        line_str = str(line).strip()
        if not line_str:
            continue

        if '会话结束' in line_str:
            if current_conversation is not None:
                time_match = re.search(r'时间:(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', line_str)
                if time_match:
                    current_conversation['end_time'] = time_match.group(1)
                conversations.append(current_conversation)
                current_conversation = None
            continue

        if '以下为一通会话' in line_str:
            current_conversation = {
                'user_id': None, 'agent_name': None,
                'start_time': None, 'end_time': None, 'messages': []
            }
            has_started = True
            continue

        if not has_started:
            sp_match = speaker_pattern.match(line_str)
            if sp_match:
                current_conversation = {
                    'user_id': None, 'agent_name': None,
                    'start_time': None, 'end_time': None, 'messages': []
                }
                has_started = True

        if current_conversation:
            sp_match = speaker_pattern.match(line_str)
            if sp_match:
                speaker = sp_match.group(1).strip()
                time_str = sp_match.group(2)

                is_agent = any(kw in speaker for kw in agent_keywords)

                if is_agent:
                    if current_conversation['agent_name'] is None:
                        current_conversation['agent_name'] = speaker
                    if current_conversation['start_time'] is None:
                        current_conversation['start_time'] = time_str
                    else:
                        current_conversation['end_time'] = time_str
                else:
                    if current_conversation['user_id'] is None:
                        current_conversation['user_id'] = speaker
                    if current_conversation['start_time'] is None:
                        current_conversation['start_time'] = time_str
                    else:
                        current_conversation['end_time'] = time_str

                current_conversation['messages'].append(
                    {'speaker': speaker, 'time': time_str, 'type': 'text', 'content': ''}
                )
                continue

            if line_str.startswith('http'):
                current_conversation['messages'].append(
                    {'speaker': None, 'time': None, 'type': 'link', 'content': line_str}
                )
                continue

            if line_str.startswith('<') and line_str.endswith('>'):
                current_conversation['messages'].append(
                    {'speaker': None, 'time': None, 'type': 'html', 'content': line_str}
                )
                continue

            if current_conversation['messages'] and current_conversation['messages'][-1]['type'] == 'text':
                last = current_conversation['messages'][-1]
                if last['content']:
                    last['content'] += '\n' + line_str
                else:
                    last['content'] = line_str
            else:
                current_conversation['messages'].append(
                    {'speaker': None, 'time': None, 'type': 'text', 'content': line_str}
                )

    print(f"解析完成: {len(conversations)} 段通话")

    # 构建输出
    output_data = []
    for conv in conversations:
        message_text = []
        for msg in conv['messages']:
            if msg['type'] == 'text' and msg['content']:
                if msg['speaker']:
                    message_text.append(f"{msg['speaker']} {msg['time'] if msg['time'] else ''}: {msg['content']}")
                else:
                    message_text.append(msg['content'])
            elif msg['type'] == 'link':
                message_text.append(f"[链接: {msg['content']}]")
            elif msg['type'] == 'html':
                message_text.append(f"[HTML: {msg['content'][:100]}...]")

        merged_text = '\n'.join(message_text)

        final_agent = conv['agent_name']
        final_user = conv['user_id']

        lines = merged_text.split('\n')
        user_empty = final_user is None or (isinstance(final_user, float) and pd.isna(final_user))
        agent_empty = final_agent is None or (isinstance(final_agent, float) and pd.isna(final_agent))

        if user_empty or agent_empty:
            if user_empty:
                for line in lines:
                    if '拜耳-' in line and ': ' in line:
                        parts = line.split(': ', 1)
                        sp = parts[0]
                        final_agent = sp.split(' 20')[0] if ' 20' in sp else sp
                        break
            if agent_empty and (not user_empty or not final_agent):
                for line in lines:
                    if (('京东大药房' in line or '您好' in line or '⭐' in line) and ': ' in line):
                        parts = line.split(': ', 1)
                        sp = parts[0]
                        candidate = sp.split(' 20')[0] if ' 20' in sp else sp
                        if candidate:
                            final_agent = candidate
                            break

        if user_empty and final_agent:
            for line in lines:
                if ': ' in line:
                    parts = line.split(': ', 1)
                    sp = parts[0]
                    speaker = sp.split(' 20')[0] if ' 20' in sp else sp
                    if speaker and speaker != final_agent and not speaker.startswith('['):
                        final_user = speaker
                        break

        output_data.append({
            '用户ID': final_user,
            '客服姓名': final_agent,
            '开始时间': conv['start_time'],
            '结束时间': conv['end_time'],
            '对话内容': merged_text
        })

    if not output_data:
        return []

    out_df = pd.DataFrame(output_data)
    out_df['开始时间'] = pd.to_datetime(out_df['开始时间'], errors='coerce')
    out_df['结束时间'] = pd.to_datetime(out_df['结束时间'], errors='coerce')
    out_df.loc[out_df['结束时间'] < out_df['开始时间'], '结束时间'] = pd.NaT
    out_df.loc[pd.isna(out_df['结束时间']), '结束时间'] = out_df['开始时间']

    final_data = []
    for _, row in out_df.iterrows():
        start_dt = row['开始时间']
        end_dt = row['结束时间']

        start_date = start_dt.strftime('%Y/%m/%d') if pd.notna(start_dt) else ""
        start_time = start_dt.strftime('%H:%M:%S') if pd.notna(start_dt) else ""
        end_date = end_dt.strftime('%Y/%m/%d') if pd.notna(end_dt) else ""
        end_time = end_dt.strftime('%H:%M:%S') if pd.notna(end_dt) else ""

        dialog = str(row['对话内容'])
        dlines = dialog.split('\n')

        entry_link = ""
        cleaned_lines = []
        for dline in dlines:
            link_match = re.match(r'^\[链接: (https?://[^\]]+)\]$', dline.strip())
            if link_match and not entry_link:
                entry_link = link_match.group(1)
            else:
                cleaned_lines.append(dline)

        final_dialog = []
        for dline in cleaned_lines:
            dline = re.sub(r'^(.*?)(?: \d{4}-\d{2}-\d{2})?(?: \d{2}:\d{2}:\d{2})?: ', r'\1: ', dline)
            final_dialog.append(dline)

        final_data.append({
            '用户ID': row['用户ID'],
            '客服姓名': row['客服姓名'],
            '开始日期': start_date,
            '开始时间': start_time,
            '结束日期': end_date,
            '结束时间': end_time,
            '进线链接': entry_link,
            '对话内容': '\n'.join(final_dialog)
        })

    return final_data


def main():
    print("=" * 60)
    print("京东聊天记录预处理")
    print("=" * 60)

    # 让用户输入文件路径
    file_path = input("请输入Excel文件路径: ").strip().strip('"')

    if not os.path.exists(file_path):
        print(f"错误: 文件不存在 {file_path}")
        return

    # 输出到输入文件同目录
    output_dir = os.path.dirname(file_path)
    output_file = os.path.join(output_dir, "客服数据.xlsx")

    # 查找聊天记录sheet
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    xl.close()

    chat_sheets = [s for s in sheet_names if '聊天记录' in s]
    if not chat_sheets:
        print("错误: 未找到包含'聊天记录'的sheet")
        return

    print(f"\n找到 {len(chat_sheets)} 个聊天记录 sheet:")
    for s in chat_sheets:
        print(f"  - {s}")

    # 处理所有sheet，整合写入一个xlsx
    total = 0
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet in chat_sheets:
            data = process_sheet(file_path, sheet)
            if data:
                sdf = pd.DataFrame(data)
                sheet_name_clean = sheet.replace('聊天记录_', '')[:31]  # Excel sheet名最长31字符
                sdf.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                total += len(data)

    print(f"\n{'='*60}")
    print(f"处理完成! 共 {total} 条记录")
    print(f"输出文件: {output_file}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
