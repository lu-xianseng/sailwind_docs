import re


def add_empty_lines_to_markdown(markdown_text):
    """
    在Markdown文本的每行之间添加空行，但：
    1. 跳过已经有空行的情况
    2. 不对表格内容进行处理
    """
    lines = markdown_text.split('\n')
    new_lines = []
    in_table = False

    for i, line in enumerate(lines):
        # 检查是否进入或离开表格区域
        if re.match(r'^\|.*\|$', line.strip()):
            if not in_table:
                in_table = True
            new_lines.append(line)
            continue
        elif in_table and line.strip() == '':
            in_table = False

        # 如果不是表格区域
        if not in_table:
            new_lines.append(line)
            # 检查是否需要添加空行
            if i < len(lines) - 1 and lines[i + 1].strip() != '' and line.strip() != '':
                new_lines.append('')

    return '\n'.join(new_lines)

if __name__ == '__main__':
    import os

    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docs_dir = os.path.join(root_dir, "docs")

    for file in os.listdir(os.path.join(docs_dir, "logic", "scripts")):
        if not file.endswith("zh.md"):
            continue
        md_files = os.path.join(docs_dir, "logic", "scripts", file)
        with  open(md_files, 'r', encoding='utf-8') as f:
            markdown_text = f.read()

        new_markdown_text = add_empty_lines_to_markdown(markdown_text)

        with open(md_files, 'w', encoding='utf-8') as f:
            f.write(new_markdown_text)