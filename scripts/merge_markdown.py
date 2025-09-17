import os
import glob
import re


def merge_markdown_files(directory, output_filename="merged.md"):
    """
    合并目录中按序列编号的Markdown文件

    :param directory: 包含Markdown文件的目录
    :param output_filename: 合并后的输出文件名
    """
    # 获取所有符合 x_数字_zh.md 模式的文件
    pattern = re.compile(r'^(.*?)_(\d+)_zh\.md$')
    files = []

    for filename in os.listdir(directory):
        match = pattern.match(filename)
        if match:
            base_name = match.group(1)
            seq_num = int(match.group(2))
            files.append((seq_num, filename, base_name))

    if not files:
        print("未找到符合命名规则的文件")
        return

    # 按序号排序文件
    files.sort()

    # 确定基础名称（取第一个文件的基础名）
    base_name = files[0][2]

    # 构建输出文件路径
    output_path = os.path.join(directory, f"{base_name}_zh.md")

    # 合并文件内容
    with open(output_path, 'w', encoding='utf-8') as outfile:
        for i, (seq_num, filename, _) in enumerate(files):
            filepath = os.path.join(directory, filename)

            with open(filepath, 'r', encoding='utf-8') as infile:
                content = infile.read().strip()  # 去除首尾空白

                # 如果不是第一个文件，在前面加一个空行
                if i > 0:
                    outfile.write("\n\n")

                outfile.write(content)

    print(f"合并完成！共合并 {len(files)} 个文件到 {output_path}")


if __name__ == "__main__":
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    target_directory = os.path.join(root_dir, "scripts", "3")

    if os.path.isdir(target_directory):
        merge_markdown_files(target_directory)
    else:
        print("目录不存在，请检查路径")