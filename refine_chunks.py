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


def identify_table_section(table_content: str) -> str:
    """
    根据表格内容识别表格应该归属的章节。
    
    参数:
        table_content: 表格内容
        
    返回:
        表格应该归属的章节名称
    """
    # 根据表格内容特征判断归属章节
    if "分类" in table_content and "危险源" in table_content and "控制措施" in table_content:
        return "3.1 风险识别"
    elif "相关模块" in table_content and "危险源" in table_content and "控制措施" in table_content:
        return "3.2 关键控制点"
    elif "成品库保管员" in table_content and "成品库班长" in table_content and "SOP撰写" in table_content:
        return "5.职责"
    elif "本SOP涉及到的主要KPI" in table_content and "PI" in table_content:
        return "6.定义和缩写"
    elif "版本" in table_content and "作者" in table_content and "日期" in table_content:
        return "8.历史文件记录"
    elif "仓库利用率" in table_content and "劳动生产率" in table_content:
        return "6.定义和缩写"
    elif "PPE矩阵" in table_content and "风险评估" in table_content:
        return "3.2 关键控制点"
    elif "应急方案" in table_content and "成品酒高空坠落" in table_content:
        return "3.2 关键控制点"
    else:
        return "未知章节"


def split_text_by_headers_and_tables(text: str, section_path: str = "") -> List[Tuple[str, str]]:
    """
    根据标题和表格拆分文本内容。
    支持多种标题格式和表格识别：
    1. Markdown格式：## 标题 或 ### 标题
    2. 数字编号格式：3.1 标题、3.2.1 标题 等
    3. 括号编号格式：1) 标题、2) 标题 等
    4. 纯数字标题：8.历史文件记录
    5. 表格识别：以 | 开头的Markdown表格
    
    参数:
        text: 原始文本内容
        section_path: 当前文本所在的section路径
        
    返回:
        包含(标题, 内容)元组的列表
    """
    if not text:
        return []
    
    # 检查是否整个文本就是一个表格
    if re.match(r'^\|.*\|', text.strip()):
        # 整个文本就是一个表格
        table_section = identify_table_section(text)
        table_header = table_section + " - 表格"
        return [(table_header, text)]
    
    # 检查是否包含表格内容
    if '|' in text and '---' in text:
        # 包含表格内容，需要拆分
        # 使用更简单的方法：直接按表格拆分
        result_chunks = split_tables_simple(text)
        return result_chunks
    
    # 如果不是纯表格，则按标题拆分
    chunks = split_text_by_headers_only(text)
    
    # 检查拆分后的chunks中是否有表格，如果有则重新处理
    result_chunks = []
    for header, content in chunks:
        if re.match(r'^\|.*\|', content.strip()):
            # 这是一个表格chunk，根据内容识别归属章节
            table_section = identify_table_section(content)
            table_header = table_section + " - 表格"
            result_chunks.append((table_header, content))
        else:
            result_chunks.append((header, content))
    
    return result_chunks


def split_tables_simple(text: str) -> List[Tuple[str, str]]:
    """
    使用改进的方法拆分包含多个表格的文本。
    确保每个表格都被识别为一个完整的chunk。
    """
    chunks = []
    
    # 首先按标题拆分
    header_chunks = split_text_by_headers_only(text)
    
    for header, content in header_chunks:
        # 检查内容是否包含表格
        if '|' in content and '---' in content:
            # 使用更简单的方法：按连续表格行分组
            lines = content.split('\n')
            tables = []
            current_table = []
            remaining_lines = []
            
            for line in lines:
                line_stripped = line.strip()
                # 检查是否为表格行
                is_table_line = (line_stripped.startswith('|') and 
                               line_stripped.endswith('|') and 
                               line_stripped.count('|') >= 2)
                
                if is_table_line:
                    current_table.append(line)
                else:
                    if current_table:
                        # 完成当前表格
                        table_text = '\n'.join(current_table).strip()
                        if table_text:
                            tables.append(table_text)
                        current_table = []
                    remaining_lines.append(line)
            
            # 处理最后的表格
            if current_table:
                table_text = '\n'.join(current_table).strip()
                if table_text:
                    tables.append(table_text)
            
            if tables:
                # 添加剩余内容（如果有）
                if remaining_lines:
                    remaining_content = '\n'.join(remaining_lines).strip()
                    if remaining_content:
                        chunks.append((header, remaining_content))
                
                # 为每个表格创建独立的chunk
                for table in tables:
                    table_section = identify_table_section(table)
                    table_header = table_section + " - 表格"
                    chunks.append((table_header, table))
            else:
                chunks.append((header, content))
        else:
            chunks.append((header, content))
    
    return chunks


def split_tables_by_regex(text: str) -> List[Tuple[str, str]]:
    """
    使用正则表达式拆分包含多个表格的文本。
    """
    chunks = []
    
    # 首先按标题拆分
    header_chunks = split_text_by_headers_only(text)
    
    for header, content in header_chunks:
        # 检查内容是否包含表格
        if '|' in content and '---' in content:
            # 使用更精确的正则表达式匹配完整的表格
            # 表格模式：以|开头和结尾的行，包含分隔符行，直到遇到非表格行
            lines = content.split('\n')
            current_table = []
            in_table = False
            remaining_lines = []
            
            for line in lines:
                line_stripped = line.strip()
                # 检查是否为表格行
                is_table_line = (line_stripped.startswith('|') and 
                               line_stripped.endswith('|') and 
                               line_stripped.count('|') >= 2)
                
                if is_table_line:
                    if not in_table:
                        # 开始新表格
                        in_table = True
                        current_table = []
                    current_table.append(line)
                else:
                    if in_table:
                        # 表格结束，保存表格
                        if current_table:
                            table_text = '\n'.join(current_table).strip()
                            if table_text:
                                table_section = identify_table_section(table_text)
                                table_header = table_section + " - 表格"
                                chunks.append((table_header, table_text))
                        in_table = False
                        current_table = []
                    remaining_lines.append(line)
            
            # 处理最后的表格
            if in_table and current_table:
                table_text = '\n'.join(current_table).strip()
                if table_text:
                    table_section = identify_table_section(table_text)
                    table_header = table_section + " - 表格"
                    chunks.append((table_header, table_text))
            
            # 添加剩余内容（如果有）
            if remaining_lines:
                remaining_content = '\n'.join(remaining_lines).strip()
                if remaining_content:
                    chunks.append((header, remaining_content))
        else:
            chunks.append((header, content))
    
    return chunks


def split_text_with_tables(text: str) -> List[Tuple[str, str]]:
    """
    拆分包含表格的文本，将每个表格作为独立的chunk。
    """
    lines = text.split('\n')
    chunks = []
    current_chunk = []
    current_header = None
    in_table = False
    table_lines = []
    
    for line in lines:
        line_stripped = line.strip()
        
        # 检查是否为表格行
        is_table_line = (line_stripped.startswith('|') and 
                        line_stripped.endswith('|') and 
                        line_stripped.count('|') >= 2)
        
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
        
        # 4. 检查纯数字标题 (如 8.历史文件记录)
        elif re.match(r'^\d+\.\s+', line_stripped):
            is_header = True
            header_text = line_stripped
        
        # 5. 检查纯数字标题（无空格，如 8.历史文件记录）
        elif re.match(r'^\d+\.', line_stripped):
            is_header = True
            header_text = line_stripped
        
        # 处理表格逻辑
        if is_table_line:
            if not in_table:
                # 开始新表格，先保存当前chunk
                if current_chunk:
                    chunk_text = '\n'.join(current_chunk).strip()
                    if chunk_text:
                        chunks.append((current_header, chunk_text))
                    current_chunk = []
                in_table = True
                table_lines = []
            table_lines.append(line)
        else:
            if in_table:
                # 表格结束，保存表格作为独立chunk
                if table_lines:
                    table_text = '\n'.join(table_lines).strip()
                    if table_text:
                        # 根据表格内容识别归属章节
                        table_section = identify_table_section(table_text)
                        table_header = table_section + " - 表格"
                        chunks.append((table_header, table_text))
                in_table = False
                table_lines = []
            
            if is_header:
                # 如果当前有积累的内容，先保存
                if current_chunk:
                    chunk_text = '\n'.join(current_chunk).strip()
                    if chunk_text:
                        chunks.append((current_header, chunk_text))
                    current_chunk = []
                
                # 设置新的标题
                current_header = header_text
                current_chunk.append(line)
            else:
                # 普通内容行
                current_chunk.append(line)
    
    # 处理最后一块内容
    if in_table and table_lines:
        # 处理最后的表格
        table_text = '\n'.join(table_lines).strip()
        if table_text:
            table_section = identify_table_section(table_text)
            table_header = table_section + " - 表格"
            chunks.append((table_header, table_text))
    elif current_chunk:
        # 处理最后的普通内容
        chunk_text = '\n'.join(current_chunk).strip()
        if chunk_text:
            chunks.append((current_header, chunk_text))
    
    return chunks


def split_text_by_headers_only(text: str) -> List[Tuple[str, str]]:
    """
    仅根据标题拆分文本内容（不处理表格）。
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
        
        # 4. 检查纯数字标题 (如 8.历史文件记录)
        elif re.match(r'^\d+\.\s+', line_stripped):
            is_header = True
            header_text = line_stripped
        
        # 5. 检查纯数字标题（无空格，如 8.历史文件记录）
        elif re.match(r'^\d+\.', line_stripped):
            is_header = True
            header_text = line_stripped
        
        if is_header:
            # 如果当前有积累的内容，先保存
            if current_chunk:
                chunk_text = '\n'.join(current_chunk).strip()
                if chunk_text:
                    chunks.append((current_header, chunk_text))
                current_chunk = []
            
            # 设置新的标题
            current_header = header_text
            current_chunk.append(line)
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
    
    # 拆分文本，传递section_path用于表格标题生成
    section_path = str(row['section_path']) if pd.notna(row['section_path']) else ""
    chunks = split_text_by_headers_and_tables(text_content, section_path)
    
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
