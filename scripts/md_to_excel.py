# _*_ coding:utf-8 _*_
"""
:Author:Hugh
:Date  :2025/05/12
"""

import os
import re
from pathlib import Path

import pandas as pd


def get_md_headings(file_path):
    """
    提取 Markdown 文件中的所有标题信息，包括父级标题
    :param file_path: Markdown文件路径
    :return: 包含标题信息的列表，每个标题包含level、title、start_line、end_line和parent_title
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    headings = []
    parent_stack = []  # 用于跟踪各级标题的父级关系

    for line_num, line in enumerate(lines):
        match = re.match(r'^(#+)\s*(.*)$', line.strip())
        if match:
            level = len(match.group(1))
            title = re.sub(r" \\\{.+", "", match.group(2).strip())  # 清理标题中的特殊标记

            # 确定父级标题
            parent_title = ""
            if parent_stack:
                # 找到当前级别的前一个同级或更高级标题作为父级
                for i in range(len(parent_stack) - 1, -1, -1):
                    if parent_stack[i]['level'] < level:
                        parent_title = parent_stack[i]['title']
                        break

            headings.append({
                'level': level,
                'title': title,
                'start_line': line_num,
                'end_line': None,
                'parent_title': parent_title
            })

            # 更新父级栈
            # 移除同级或更低级的标题
            while parent_stack and parent_stack[-1]['level'] >= level:
                parent_stack.pop()
            parent_stack.append({
                'level': level,
                'title': title
            })

    # 确定每个标题的结束行
    for i in range(len(headings)):
        if i < len(headings) - 1:
            headings[i]['end_line'] = headings[i + 1]['start_line'] - 1
        else:
            headings[i]['end_line'] = len(lines) - 1

    return headings


def parse_md_file(file_path, ignore_flag=None):
    # 提取 Markdown 文件中的所有标题信息
    headings = get_md_headings(file_path)

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
        heading['path'] = re.sub(r'[ >/\\\\]+', '_', path).strip('_')

    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

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


def find_files_by_extension(directory, extensions):
    """
    在指定目录中递归查找指定扩展名的文件
    """
    if isinstance(extensions, str):
        extensions = [extensions]

    found_files = []
    for root, _, files in os.walk(directory):
        found_files = [os.path.join(root, file) for file in files if any(file.endswith(ext) for ext in extensions)]

    found_files = sorted(found_files, key=lambda x: int(os.path.basename(x).split('_')[0]))
    return found_files


def process_directory(directory, output_excel, ignore_flag=None):
    all_entries = []
    found_files = find_files_by_extension(directory, r".md")
    for file in found_files:
        entries = parse_md_file(file, ignore_flag)
        all_entries.extend(entries)

    # 创建DataFrame并保存Excel
    df = pd.DataFrame(all_entries, columns=['需求名称', '所属模块', '需求说明', '测试点设计要求'])
    # 判断 output_excel 目录是否存在，如果不存在则创建
    output_excel.parent.mkdir(parents=True, exist_ok=True)
    # 保存为Excel文件
    df.to_excel(output_excel, index=False, engine='openpyxl')


if __name__ == "__main__":
    directory = Path(r"D:\hugh\code\sailwind3.0_docs\docs\logic\guide")
    output_file = Path(__file__).parent / "product_demand" / "logic_guide.xlsx"

    ignore_flag = ["🚧", "ಠ_ಠ", "❌"]

    process_directory(directory, output_file, ignore_flag)

    print(f"处理完成，结果已保存至：{output_file}")
