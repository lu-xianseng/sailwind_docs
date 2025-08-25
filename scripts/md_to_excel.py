# _*_ coding:utf-8 _*_
"""
:Author:Hugh
:Date  :2025/05/12
"""

import os
import re
from pathlib import Path

import pandas as pd


def parse_md_file(file_path, ignore_flag=None):
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 提取所有标题信息
    headings = []
    for line_num, line in enumerate(lines):
        match = re.match(r'^(#+)\s*(.*)$', line.strip())
        if match:
            level = len(match.group(1))
            title = match.group(2).strip()
            title = re.sub(r" \\\{.+\}","", title)

            # 将一级标题中的"🚧"替换为""
            # if level == 1:
            #     title = title.replace("🚧", "") if "🚧" in title else title

            headings.append({
                'level': level,
                'title': title,
                'start_line': line_num,
                'end_line': None
            })

    # 确定每个标题的结束行
    for i in range(len(headings)):
        if i < len(headings) - 1:
            headings[i]['end_line'] = headings[i + 1]['start_line'] - 1
        else:
            headings[i]['end_line'] = len(lines) - 1

    # 判断是否为末级标题
    # 遍历所有标题，检查后续是否存在更高级或同级标题
    for i in range(len(headings)):
        current_level = headings[i]['level']
        is_last = True
        # 向后查找后续标题
        for j in range(i + 1, len(headings)):
            if headings[j]['level'] > current_level:
                # 发现子标题，说明当前标题不是末级
                is_last = False
                break
            elif headings[j]['level'] <= current_level:
                # 遇到同级或上级标题，表示当前标题已结束
                break
        headings[i]['is_last'] = is_last

    # 构建每个标题的路径
    current_levels = {}
    for heading in headings:
        level = heading['level']
        # 清除同级或更高层级
        keys_to_remove = [k for k in current_levels if k >= level]
        for k in keys_to_remove:
            del current_levels[k]
        current_levels[level] = heading['title']
        # 生成路径
        sorted_levels = sorted(current_levels.keys())
        path = '-'.join([current_levels[l] for l in sorted_levels])
        heading['path'] = path.replace(" ", "_").replace(">", "_").replace("/", "_").replace("\\", "_").replace("__",
                                                                                                                "_")

    # 收集末级标题的内容
    entries = []
    for heading in headings:

        # 如果有忽略标志且标题中包含忽略标志，则跳过
        if ignore_flag and any(flag in heading['path'] for flag in ignore_flag):
            continue

        # 跳过一级标题
        if heading['level'] == 1:
            continue

        if heading['is_last']:
            start = heading['start_line'] + 1
            end = heading['end_line'] + 1  # 切片需要包含最后一行
            content_lines = lines[start:end]
            content = ''.join(content_lines).strip().replace("**", '"')

            entries.append({
                '需求名称': heading['title'],
                '所属模块': heading['path'],
                '需求说明': content,
                # '测试点设计要求': ""
            })

    return entries


def process_directory(file_path, output_excel, ignore_flag=None):
    all_entries = []
    for root, _, files in os.walk(file_path):
        md_files = [f for f in files if f.endswith('.md')]
        sorted_files = sorted(md_files, key=lambda x: int(x.split('_')[0]))
        for file in sorted_files:
            if file.endswith('.md'):
                full_path = os.path.join(root, file)
                entries = parse_md_file(full_path, ignore_flag)
                all_entries.extend(entries)

    # 创建DataFrame并保存Excel
    df = pd.DataFrame(all_entries, columns=['需求名称', '所属模块', '需求说明', '测试点设计要求'])
    # 判断 output_excel 目录是否存在，如果不存在则创建
    output_excel.parent.mkdir(parents=True, exist_ok=True)
    # 保存为Excel文件
    df.to_excel(output_excel, index=False, engine='openpyxl')


if __name__ == "__main__":
    file_path = Path(r"D:\hugh\code\sailwind3.0_docs\docs\logic\guide")
    output_file = Path(__file__).parent / "product_demand" / "logic_guide.xlsx"

    ignore_flag = ["🚧", "ಠ_ಠ", "❌"]

    process_directory(file_path, output_file, ignore_flag)

    print(f"处理完成，结果已保存至：{output_file}")
