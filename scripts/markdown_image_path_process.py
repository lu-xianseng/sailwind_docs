# _*_ coding:utf-8 _*_
"""
:Author:Hugh
:Date  :2025/05/20
"""

import os
import re
import shutil
import argparse
from pathlib import Path

project_path = Path(__file__).parent.parent


def process_image_links(input_dir):
    input_path = Path(input_dir)
    for md_file in input_path.rglob('*.md'):
        relative_path = md_file.parent.relative_to(input_path)
        content = md_file.read_text(encoding='utf-8')

        # 替换图片路径
        def replace_image(match):
            img_path = match.group(2)
            if img_path.startswith(('http://', 'https://', '/')):
                return match.group(0)
            base_dir = input_dir.parent.parent
            _path = input_dir.relative_to(base_dir)
            new_path = f'/{_path}/{relative_path}/{img_path}'.replace('\\', '/')
            return f'![{match.group(1)}]({new_path})'

        new_content = re.sub(r'!\[(.*?)\]\((.*?)\)', replace_image, content)
        md_file.write_text(new_content, encoding='utf-8')


def copy_and_clean_public(input_dir):
    # 获取基目录（docs目录）
    base_dir = input_dir.parent.parent
    # 获取相对于v4的子路径（layout/guide）
    relative_path = input_dir.relative_to(base_dir)
    # 构造新路径
    image_path = base_dir / 'public' / relative_path

    if image_path.exists():
        shutil.rmtree(image_path)
    shutil.copytree(input_dir, image_path)

    # 删除非图片文件并清理空目录
    for root, dirs, files in os.walk(image_path, topdown=False):
        for file in files:
            if not file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                os.remove(os.path.join(root, file))
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            if not os.listdir(dir_path):
                os.rmdir(dir_path)


def move_md_files_and_clean(input_dir):
    input_path = Path(input_dir)
    for md_file in input_path.rglob('*.md'):
        if md_file.parent != input_path:
            dest = input_path / md_file.name
            if dest.exists():
                dest.unlink()
            shutil.move(md_file, dest)

    # 删除所有子目录
    for child in input_path.iterdir():
        if child.is_dir():
            shutil.rmtree(child)


def main(input_dir):
    input_dir = Path(input_dir)
    process_image_links(input_dir)
    copy_and_clean_public(input_dir)
    move_md_files_and_clean(input_dir)


if __name__ == "__main__":
    # parser = argparse.ArgumentParser(description="处理文档目录")
    # parser.add_argument(
    #     "--document_dir",
    #     type=str,
    #     default="docs/v4/layout/guide",  # 默认已替换为 public
    #     help="文档相对路径（默认：docs/v4/layout/guide）"
    # )
    #
    # args = parser.parse_args()
    # document_dir = args.document_dir

    document_dir = r"D:\hugh\code\sailwind3.0_docs\docs\logic\scripts"
    input_dir = project_path / document_dir
    # input_dir = r"D:\hugh\code\sailwind3.0_docs\docs\v4\lpcreator"
    main(input_dir)
