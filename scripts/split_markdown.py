import os


def split_markdown_by_heading(input_file, lines_per_chunk=250):
    """
    拆分Markdown文件，每100行之后的第一个#号开头的标题处拆分

    :param input_file: 输入的Markdown文件路径
    :param lines_per_chunk: 每个块的行数阈值
    """
    # 读取原始文件内容
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 获取文件名和目录名
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_dir = base_name  # 输出目录与文件名相同

    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)

    # 初始化变量
    chunks = []
    current_chunk = []
    chunk_count = 1
    line_count_since_last_split = 0
    pending_heading = None  # 用于暂存待处理的标题行

    for line in lines:
        # 先检查是否是标题行且满足拆分条件
        if (line_count_since_last_split >= lines_per_chunk
                and line.strip().startswith('#')):
            # 保存当前块（不包含这个标题行）
            chunks.append((chunk_count, current_chunk))
            chunk_count += 1
            current_chunk = [line]  # 新块从这个标题行开始
            line_count_since_last_split = 1  # 重置计数器（当前标题行算作1）
            continue

        current_chunk.append(line)
        line_count_since_last_split += 1

    # 添加最后一个块
    if current_chunk:
        chunks.append((chunk_count, current_chunk))

    # 写入拆分后的文件
    for chunk_num, chunk_content in chunks:
        output_filename = f"{base_name}_{chunk_num}.md"
        output_path = os.path.join(output_dir, output_filename)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.writelines(chunk_content)

    print(f"拆分完成，共拆分为 {len(chunks)} 个文件，保存在 '{output_dir}' 目录中")

# 使用示例
if __name__ == "__main__":
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docs_dir = os.path.join(root_dir, "docs")

    input_file = os.path.join(docs_dir, "logic", "scripts", "4.md")
    if os.path.isfile(input_file):
        split_markdown_by_heading(input_file)
    else:
        print("文件不存在，请检查路径")