#!/usr/bin/env python3
"""
将 .docx 文档解析为显式的嵌套 JSON，支持混合识别逻辑。

功能要点：
- 使用 sys.argv 接收输入文件路径并进行基本检查。
- 使用 python-docx 读取段落与样式，实现"数字编号为主，样式为辅"的混合识别。
- 普通段落归属到其上方最近的标题的 content 列表中。
- 自动识别并处理文档中的表格，将表格内容转换为Markdown格式。
- 输出 JSON: { sop_id, sop_name, sections }，sections 为顶层标题列表，
  每个标题包含 { title, level, content, subsections }。

混合识别逻辑：
1. 第一优先级 - 检查文本：使用正则表达式匹配数字编号格式推断标题级别
   - 一级：^\\d+\\.\\s+ (如 "1. 目的")
   - 二级：^\\d+\\.\\d+\\s+ (如 "2.1 适用范围") 
   - 三级：^\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1 具体内容")
   - 四级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1 详细说明")
   - 五级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1 子详细说明")
   - 六级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1 更详细说明")
   - 七级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1 最详细说明")
   - 八级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1.1 超详细说明")
   - 九级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1.1.1 极详细说明")
   - 十级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1.1.1.1 终极详细说明")
2. 第二优先级 - 检查样式：paragraph.style.name 匹配 'Heading X' 或 '标题 X'（作为备用）

注意：
- 文档标题(样式名 'Title' 或中文 '标题')不会作为章节加入，而是作为 sop_name 的优先来源。
- 若无法检测到文档标题，则 sop_name 取文档第一段非空文本；若仍不可得，则回退为 sop_id。
"""

from __future__ import annotations

import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Tuple

try:
    from docx2python import docx2python  # type: ignore
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "[ERROR] 未找到 docx2python 库，请先安装：pip install docx2python\n"
    )
    sys.exit(1)


HeadingNode = Dict[str, Any]


def extract_heading_level_from_style(style_name: Optional[str]) -> Optional[int]:
    """根据段落样式名推断标题级别（第二优先级，作为备用）。

    支持示例：
    - 'Heading 1', 'Heading 2', ..., 'Heading 10'
    - '标题 1', '标题 2', ..., '标题 10'（中文界面）
    返回对应的整数级别（1-10）；若非标题样式则返回 None。
    """
    if not style_name:
        return None

    # 统一去除两端空白，便于匹配
    normalized = style_name.strip()

    # 英文样式：Heading 1/2/3 ...
    m = re.match(r"^Heading\s+([1-9][0-9]*)$", normalized, re.IGNORECASE)
    if m:
        return int(m.group(1))

    # 中文样式：标题 1/2/3 ...
    m = re.match(r"^标题\s*([1-9][0-9]*)$", normalized)
    if m:
        return int(m.group(1))

    return None


def extract_heading_level_from_text(text: str) -> Optional[int]:
    """根据段落文本内容推断标题级别（第一优先级）。

    使用正则表达式匹配数字编号格式：
    - 一级：^\\d+\\.\\s+ (如 "1. 目的")
    - 二级：^\\d+\\.\\d+\\s+ (如 "2.1 适用范围") 
    - 三级：^\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1 具体内容")
    - 四级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1 详细说明")
    - 五级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1 子详细说明")
    - 六级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1 更详细说明")
    - 七级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1 最详细说明")
    - 八级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1.1 超详细说明")
    - 九级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1.1.1 极详细说明")
    - 十级：^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s+ (如 "2.1.1.1.1.1.1.1.1.1 终极详细说明")
    
    返回对应的整数级别（1-10）；若文本不符合编号格式则返回 None。
    """
    if not text:
        return None
    
    text = text.strip()
    
    # 十级标题：数字.数字.数字.数字.数字.数字.数字.数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\s+", text):
        return 10
    
    # 九级标题：数字.数字.数字.数字.数字.数字.数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\s+", text):
        return 9
    
    # 八级标题：数字.数字.数字.数字.数字.数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\s+", text):
        return 8
    
    # 七级标题：数字.数字.数字.数字.数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\.\d+\.\d+\.\d+\.\d+\s+", text):
        return 7
    
    # 六级标题：数字.数字.数字.数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\.\d+\.\d+\.\d+\s+", text):
        return 6
    
    # 五级标题：数字.数字.数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\.\d+\.\d+\s+", text):
        return 5
    
    # 四级标题：数字.数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\.\d+\s+", text):
        return 4
    
    # 三级标题：数字.数字.数字 空格
    if re.match(r"^\d+\.\d+\.\d+\s+", text):
        return 3
    
    # 二级标题：数字.数字 空格
    if re.match(r"^\d+\.\d+\s+", text):
        return 2
    
    # 一级标题：数字. 空格
    if re.match(r"^\d+\.\s+", text):
        return 1
    
    return None


def extract_heading_level(style_name: Optional[str], text: str) -> Optional[int]:
    """混合识别逻辑：数字编号为主，样式为辅。
    
    第一优先级：检查文本内容的数字编号格式
    第二优先级：检查段落样式名（作为备用）
    
    返回标题级别（1-10），若非标题则返回 None。
    """
    # 第一优先级：检查文本编号格式
    level = extract_heading_level_from_text(text)
    if level is not None:
        return level
    
    # 第二优先级：检查样式（作为备用）
    level = extract_heading_level_from_style(style_name)
    if level is not None:
        return level
    
    return None


def is_document_title_style(style_name: Optional[str]) -> bool:
    """判断样式是否为文档标题(非 Heading 层级)。

    常见样式名：'Title' 或中文界面中的 '标题'（注意与 '标题 1' 区分）。
    """
    if not style_name:
        return False
    name = style_name.strip()
    if name.lower() == "title":
        return True
    # 精确等于“标题”视为文档标题，避免与“标题 1/2...”混淆
    if name == "标题":
        return True
    return False


def convert_table_to_markdown(table_data: List[List[str]]) -> str:
    """将一个二维列表（表格数据）转换为Markdown表格格式的字符串。
    
    参数:
        table_data: 二维列表，每个子列表代表表格的一行
        
    返回:
        Markdown格式的表格字符串
    """
    if not table_data:
        return ""

    # 构建Markdown表格的字符串
    markdown_string = ""

    # 1. 处理表头
    header = "| " + " | ".join(str(cell).strip() for cell in table_data[0]) + " |"
    markdown_string += header + "\n"

    # 2. 处理分隔符行
    separator = "| " + " | ".join(["---"] * len(table_data[0])) + " |"
    markdown_string += separator + "\n"

    # 3. 处理表格内容行
    for row in table_data[1:]:
        # 确保行数据长度与表头一致
        while len(row) < len(table_data[0]):
            row.append("")
        content_row = "| " + " | ".join(str(cell).strip() for cell in row[:len(table_data[0])]) + " |"
        markdown_string += content_row + "\n"

    return markdown_string


def extract_images_from_text(text: str, image_mapping: Dict[str, str]) -> Tuple[str, List[str]]:
    """从文本中提取图片占位符并返回清理后的文本和图片列表。
    
    参数:
        text: 包含图片占位符的文本
        image_mapping: 原始文件名到新文件名的映射
        
    返回:
        Tuple[清理后的文本, 图片文件名列表]
    """
    if not text:
        return text, []
    
    # 匹配图片占位符格式：----media/filename.ext----
    image_pattern = r'----media/([^/]+\.(?:png|jpg|jpeg|gif|bmp|tiff?))----'
    matches = re.findall(image_pattern, text, re.IGNORECASE)
    
    # 移除占位符，保留清理后的文本
    cleaned_text = re.sub(image_pattern, '', text).strip()
    
    # 将原始文件名转换为新的文件名
    new_image_names = []
    for img_name in matches:
        if img_name in image_mapping:
            new_image_names.append(image_mapping[img_name])
    
    return cleaned_text, new_image_names


def rename_image_files(images_dict: Dict[str, bytes], sop_id: str, image_folder: str) -> Dict[str, str]:
    """重命名图片文件并返回原始文件名到新文件名的映射。
    
    参数:
        images_dict: docx2python返回的图片字典
        sop_id: SOP文档ID
        image_folder: 图片保存目录
        
    返回:
        原始文件名到新文件名的映射字典
    """
    if not images_dict:
        return {}
    
    # 确保图片目录存在
    os.makedirs(image_folder, exist_ok=True)
    
    renamed_mapping = {}
    
    for original_name, image_data in images_dict.items():
        # 生成新的唯一文件名
        name, ext = os.path.splitext(original_name)
        new_name = f"{sop_id}_{name}{ext}"
        
        # 保存重命名后的图片文件
        new_path = os.path.join(image_folder, new_name)
        with open(new_path, 'wb') as f:
            f.write(image_data)
        
        # 记录映射关系
        renamed_mapping[original_name] = new_name
    
    return renamed_mapping


def make_node(title: str, level: int) -> HeadingNode:
    """创建一个标题节点，将标题同时作为第一条内容。"""
    return {
        "title": title,
        "level": level,
        "content": [title],  # 将标题作为第一条内容
        "images": [],  # type: List[str]
        "subsections": [],  # type: List[HeadingNode]
    }


def docx_to_nested_json(input_path: str) -> Dict[str, Any]:
    """解析给定 .docx 文档，返回目标 JSON 结构的 Python 字典。"""
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"文件不存在: {input_path}")
    if not input_path.lower().endswith(".docx"):
        raise ValueError("输入文件必须为 .docx 格式")

    sop_id = os.path.splitext(os.path.basename(input_path))[0]
    image_folder = "sop_images"
    
    try:
        # 使用 docx2python 解析文档
        doc_result = docx2python(input_path, image_folder=image_folder)
    except Exception as exc:  # pragma: no cover
        raise ValueError(f"无法读取 Word 文档: {exc}")

    # 处理图片文件重命名
    image_mapping = rename_image_files(doc_result.images, sop_id, image_folder)
    
    # 获取文档文本内容和表格数据
    document_text = doc_result.text
    document_body = doc_result.body
    
    # 提取文档标题(sop_name)
    sop_name: Optional[str] = None
    lines = document_text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 检查是否为文档标题（通常在第一行或前几行）
        if not re.match(r'^\d+\.', line):  # 不是数字编号开头
            sop_name = line
            break
    
    if not sop_name:
        sop_name = sop_id

    sections: List[HeadingNode] = []
    # 栈保存 (level, node)
    heading_stack: List[Tuple[int, HeadingNode]] = []

    # 按行处理文档内容
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 检查是否为标题
        level = extract_heading_level(None, line)  # docx2python不提供样式信息

        if level is not None:
            # 新的标题节点。维护单调栈，使其父子关系正确。
            new_node = make_node(line, level)

            # 弹出栈中当前级别及更深的标题，直到找到父级(level 更小)
            while heading_stack and heading_stack[-1][0] >= level:
                heading_stack.pop()

            if not heading_stack:
                # 顶级标题
                sections.append(new_node)
            else:
                # 作为最近上级标题的子节点
                parent_node = heading_stack[-1][1]
                parent_node["subsections"].append(new_node)

            # 将当前标题压栈
            heading_stack.append((level, new_node))
        else:
            # 普通段落，处理图片占位符
            cleaned_text, image_names = extract_images_from_text(line, image_mapping)
            
            if cleaned_text:  # 只有非空文本才添加
                if heading_stack:
                    heading_stack[-1][1]["content"].append(cleaned_text)
                    
                    # 添加关联的图片
                    for img_name in image_names:
                        if img_name not in heading_stack[-1][1]["images"]:
                            heading_stack[-1][1]["images"].append(img_name)
                else:
                    # 若在任何标题出现之前出现正文，则不纳入 sections。
                    # 可根据业务需要改为挂入一个"前言"节点。
                    pass

    # 处理表格数据 - 从body中提取表格并添加到对应标题下
    def extract_tables_from_body(body_data):
        """从docx2python的body数据中提取表格"""
        tables = []
        if isinstance(body_data, list):
            for item in body_data:
                if isinstance(item, list):
                    # 检查是否为表格（二维列表结构）
                    if len(item) > 0 and isinstance(item[0], list):
                        # 检查是否所有子项都是列表（表格行）
                        is_table = all(isinstance(row, list) for row in item)
                        if is_table and len(item) > 1:  # 确保表格有多行
                            # 清理表格数据，移除方括号和多余格式
                            cleaned_table = []
                            for row in item:
                                cleaned_row = []
                                for cell in row:
                                    if isinstance(cell, str):
                                        # 移除方括号和引号，处理嵌套的字符串表示
                                        clean_cell = cell.strip("[]'\"")
                                        # 如果仍然包含方括号，进一步清理
                                        if '[' in clean_cell and ']' in clean_cell:
                                            # 使用正则表达式移除所有方括号和引号
                                            import re
                                            clean_cell = re.sub(r'[\[\]\'"]', '', clean_cell)
                                        cleaned_row.append(clean_cell.strip())
                                    else:
                                        cleaned_row.append(str(cell).strip())
                                cleaned_table.append(cleaned_row)
                            tables.append(cleaned_table)
                    else:
                        # 递归处理嵌套结构
                        tables.extend(extract_tables_from_body(item))
        return tables
    
    document_tables = extract_tables_from_body(document_body)
    
    # 将表格添加到最后一个标题下（简化处理）
    if document_tables and heading_stack:
        for table_data in document_tables:
            if table_data:  # 确保表格不为空
                # 将表格转换为Markdown格式
                markdown_table = convert_table_to_markdown(table_data)
                if markdown_table:
                    # 表格归属到最近的标题
                    heading_stack[-1][1]["content"].append(markdown_table)

    return {
        "sop_id": sop_id,
        "sop_name": sop_name,
        "sections": sections,
    }


def main(argv: List[str]) -> int:
    if len(argv) < 2:
        sys.stderr.write(
            "用法: python word_to_json.py <输入文件.docx>\n"
        )
        return 1

    input_path = argv[1]
    try:
        data = docx_to_nested_json(input_path)
    except (FileNotFoundError, ValueError) as exc:
        sys.stderr.write(f"[ERROR] {exc}\n")
        return 1
    except Exception as exc:  # pragma: no cover
        # 捕获未知错误，便于定位问题
        sys.stderr.write(f"[ERROR] 未预期的异常: {exc}\n")
        return 1

    # 输出至同名 .json
    output_path = os.path.splitext(input_path)[0] + ".json"
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as exc:  # pragma: no cover
        sys.stderr.write(f"[ERROR] 写入 JSON 失败: {exc}\n")
        return 1

    print(f"已生成: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
