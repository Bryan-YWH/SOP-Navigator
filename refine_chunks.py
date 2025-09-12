#!/usr/bin/env python3
"""
文本块精炼脚本 - 用于优化RAG系统的数据质量

该脚本读取包含大文本块的CSV文件，根据Markdown标题进行拆分，
生成更细粒度的知识块，提升RAG系统的检索精度。

"""

import sys
import re
import pandas as pd
from typing import List, Dict, Any, Tuple


def clean_text_content(text: str) -> str:
    """
    清理和标准化文本内容。
    
    将常见的非标准列表符号替换为标准的Markdown无序列表符号。
    
    参数:
        text: 原始文本内容
        
    返回:
        清理后的文本内容
    """
    if not text:
        return text
    
    # 替换各种非标准列表符号为标准的Markdown列表符号
    # 处理 `· ` 符号
    text = re.sub(r'^·\s+', '* ', text, flags=re.MULTILINE)
    
    # 处理 `--\t` 符号
    text = re.sub(r'^--\s+', '* ', text, flags=re.MULTILINE)
    
    # 处理中文顿号用于列表的情况（需要上下文判断）
    # 这里简化处理，将行首的顿号替换为星号
    text = re.sub(r'^\s*、\s+', '* ', text, flags=re.MULTILINE)
    
    # 处理其他常见的列表符号
    text = re.sub(r'^[-•]\s+', '* ', text, flags=re.MULTILINE)
    
    return text


def split_text_by_headers(text: str) -> List[Tuple[str, str]]:
    """
    根据标题拆分文本内容。
    支持多种标题格式：
    1. Markdown格式：## 标题 或 ### 标题
    2. 数字编号格式：3.1 标题、3.2.1 标题 等
    3. 括号编号格式：1) 标题、2) 标题 等
    
    参数:
        text: 原始文本内容
        
    返回:
        包含(标题, 内容)元组的列表
    """
    if not text:
        return []
    
    lines = text.split('\n')
    chunks = []
    current_chunk = []
    current_header = None
    
    for line in lines:
        line_stripped = line.strip()
        
        # 检查是否为标题
        is_header = False
        header_text = None
        
        # 1. 检查Markdown格式标题
        markdown_match = re.match(r'^(#{2,3})\s+(.+)$', line_stripped)
        if markdown_match:
            is_header = True
            header_text = markdown_match.group(2).strip()
        
        # 2. 检查数字编号格式标题 (如 3.1、3.2.1、7.1.11 等)
        elif re.match(r'^\d+\.\d+(\.\d+)*\s+', line_stripped):
            is_header = True
            header_text = line_stripped
        
        # 3. 检查括号编号格式标题 (如 1)、2)、3) 等)
        elif re.match(r'^\d+\)\s+', line_stripped):
            is_header = True
            header_text = line_stripped
        
        if is_header:
            # 如果当前有积累的内容，先保存
            if current_chunk:
                chunk_text = '\n'.join(current_chunk).strip()
                if chunk_text:
                    chunks.append((current_header, chunk_text))
            
            # 开始新的块
            current_header = header_text
            current_chunk = [line]  # 包含标题行
        else:
            # 普通内容行
            current_chunk.append(line)
    
    # 处理最后一个块
    if current_chunk:
        chunk_text = '\n'.join(current_chunk).strip()
        if chunk_text:
            chunks.append((current_header, chunk_text))
    
    return chunks


def process_single_row(row: pd.Series) -> List[Dict[str, Any]]:
    """
    处理单行数据，根据Markdown标题进行拆分。
    
    参数:
        row: pandas Series，包含一行CSV数据
        
    返回:
        拆分后的多行数据列表
    """
    text_content = str(row['text']) if pd.notna(row['text']) else ''
    
    # 拆分文本
    chunks = split_text_by_headers(text_content)
    
    if not chunks:
        # 如果没有找到Markdown标题，返回原始行
        return [row.to_dict()]
    
    result_rows = []
    
    for i, (header, chunk_text) in enumerate(chunks):
        # 创建新行数据
        new_row = row.to_dict()
        
        # 更新文本内容
        new_row['text'] = clean_text_content(chunk_text)
        
        # 更新section_path
        if i == 0:
            # 第一个块保持原始section_path
            new_row['section_path'] = row['section_path']
        else:
            # 后续块使用对应的Markdown标题作为section_path
            new_row['section_path'] = header
        
        result_rows.append(new_row)
    
    return result_rows


def refine_csv_chunks(input_file: str, output_file: str) -> None:
    """
    主要的CSV处理函数。
    
    参数:
        input_file: 输入CSV文件路径
        output_file: 输出CSV文件路径
    """
    try:
        # 读取输入CSV文件
        print(f"正在读取输入文件: {input_file}")
        df = pd.read_csv(input_file, encoding='utf-8-sig')
        print(f"成功读取 {len(df)} 行数据")
        
        # 处理每一行
        all_processed_rows = []
        
        for index, row in df.iterrows():
            print(f"处理第 {index + 1}/{len(df)} 行...")
            
            # 拆分当前行
            processed_rows = process_single_row(row)
            all_processed_rows.extend(processed_rows)
            
            if len(processed_rows) > 1:
                print(f"  -> 拆分为 {len(processed_rows)} 个块")
        
        # 创建新的DataFrame
        result_df = pd.DataFrame(all_processed_rows)
        
        # 保存到输出文件
        print(f"正在保存到输出文件: {output_file}")
        result_df.to_csv(output_file, index=False, encoding='utf-8-sig', quoting=1)  # quoting=1 确保所有字段都被引号包围
        
        print(f"处理完成！")
        print(f"原始行数: {len(df)}")
        print(f"处理后行数: {len(result_df)}")
        print(f"平均每行拆分为: {len(result_df) / len(df):.2f} 个块")
        
    except FileNotFoundError:
        print(f"错误: 找不到输入文件 '{input_file}'")
        sys.exit(1)
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        sys.exit(1)


def main():
    """主函数，处理命令行参数并调用核心处理函数。"""
    if len(sys.argv) != 3:
        print("用法: python refine_chunks.py <输入CSV文件> <输出CSV文件>")
        print("示例: python refine_chunks.py input.csv output.csv")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    print("=" * 60)
    print("CSV文本块精炼工具")
    print("=" * 60)
    
    refine_csv_chunks(input_file, output_file)


if __name__ == "__main__":
    main()
