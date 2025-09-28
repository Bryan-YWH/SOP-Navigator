#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SOP文档智能解析脚本 - 集成图片处理版本
实现精确的上下文追踪、表格归属、动态标题生成和图片处理
"""

import sys
import re
import json
import pandas as pd
import os
import shutil
from collections import defaultdict
from typing import List, Dict, Tuple
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import parse_xml


def iter_block_items(doc: Document):
    """
    按文档实际顺序依次返回段落和表格。
    用于在不改动主处理流程的情况下，预先为表格记录其出现位置对应的最近标题路径。
    """
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)

def extract_images_with_captions_from_docx(docx_path: str, output_dir: str = "sop_images") -> Dict[str, Dict[str, str]]:
    """
    从Word文档中提取图片及其caption信息并保存到指定目录
    
    参数:
        docx_path: Word文档路径
        output_dir: 图片输出目录
        
    返回:
        图片信息字典 {图片ID: {"filename": 图片文件名, "caption": 图片标题}}
    """
    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 获取文档名称（不含扩展名）
    doc_name = os.path.splitext(os.path.basename(docx_path))[0]
    
    # 打开Word文档
    doc = Document(docx_path)
    
    # 图片信息字典
    image_info = {}
    image_counter = 1
    
    # 遍历文档中的所有段落，查找图片和其caption
    for i, paragraph in enumerate(doc.paragraphs):
        # 检查段落是否包含图片
        has_image = False
        image_rel_id = None
        
        for run in paragraph.runs:
            # 查找图片关系ID
            for drawing in run._element.xpath('.//a:blip'):
                # 获取图片关系ID
                embed_id = drawing.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed_id:
                    image_rel_id = embed_id
                    has_image = True
                    break
            if has_image:
                break
        
        if has_image and image_rel_id:
            # 查找图片的caption（通常在图片下方或上方的段落中）
            caption = ""
            
            # 检查当前段落是否包含caption
            if paragraph.text.strip():
                caption = paragraph.text.strip()
            else:
                # 检查下一个段落是否包含caption
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    if next_para.text.strip():
                        caption = next_para.text.strip()
            
            # 根据关系ID找到对应的图片关系
            target_rel = None
            for rel in doc.part.rels.values():
                if rel.rId == image_rel_id:
                    target_rel = rel
                    break
            
            if target_rel and "image" in target_rel.target_ref:
                # 获取图片数据
                image_data = target_rel.target_part.blob
                
                # 确定图片扩展名
                if target_rel.target_ref.endswith('.png'):
                    ext = '.png'
                elif target_rel.target_ref.endswith('.jpg') or target_rel.target_ref.endswith('.jpeg'):
                    ext = '.jpg'
                elif target_rel.target_ref.endswith('.gif'):
                    ext = '.gif'
                elif target_rel.target_ref.endswith('.bmp'):
                    ext = '.bmp'
                else:
                    ext = '.png'  # 默认扩展名
                
                # 生成图片文件名，添加image_id前缀
                image_filename = f"{doc_name}_image_id____image{image_counter}{ext}"
                image_path = os.path.join(output_dir, image_filename)
                
                # 保存图片
                with open(image_path, 'wb') as f:
                    f.write(image_data)
                
                # 记录图片信息
                image_info[f"image_{image_counter}"] = {
                    "filename": image_filename,
                    "caption": caption
                }
                
                print(f"提取图片: {image_filename}")
                if caption:
                    print(f"  图片标题: {caption}")
                
                image_counter += 1
    
    return image_info


def create_enhanced_image_chunk_content(filename: str, target_section: str, caption: str) -> str:
    """
    创建增强的图片chunk内容
    
    参数:
        filename: 图片文件名
        target_section: 目标章节路径
        caption: 图片标题/描述
        
    返回:
        格式化的图片chunk内容
    """
    # 提取纯文字内容（去掉数字）
    clean_caption = ""
    if caption:
        # 使用正则表达式提取中文字符和英文字母，去掉数字和特殊符号
        import re
        clean_caption = re.sub(r'[\d\.\-\s]+', '', caption).strip()
        if clean_caption:
            clean_caption = f"图片内容：{clean_caption}"
    
    # 构建增强的chunk内容
    content_parts = [
        f"图片名称：{filename}",
        f"图片所在SOP位置：{target_section}"
    ]
    
    if clean_caption:
        content_parts.append(clean_caption)
    
    return "\n".join(content_parts)


def find_image_references_in_text(text: str) -> List[str]:
    """
    在文本中查找图片引用
    
    参数:
        text: 文本内容
        
    返回:
        图片引用列表
    """
    # 查找可能的图片引用模式
    image_patterns = [
        r'图\s*\d+',  # 图1, 图 1
        r'图片\s*\d+',  # 图片1, 图片 1
        r'附图\s*\d+',  # 附图1, 附图 1
        r'Figure\s*\d+',  # Figure 1
        r'Fig\s*\d+',  # Fig 1
        r'见下图',  # 见下图
        r'如图所示',  # 如图所示
        r'参考图',  # 参考图
    ]
    
    references = []
    for pattern in image_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        references.extend(matches)
    
    return references


def is_heading_paragraph(paragraph: Paragraph) -> bool:
    """
    判断段落是否为标题
    只识别真正的标题，其他所有文本都作为普通文本处理
    """
    # 第一优先级：检查Word样式
    style_name = paragraph.style.name
    if 'Heading' in style_name or '标题' in style_name:
        return True
    # 兼容简写样式名：H1/H2/H3...
    if re.match(r'^H\d+$', style_name):
        return True
    
    text = paragraph.text.strip()
    
    # 第二优先级：检查多级数字编号 (如 3.1, 8.2.1)
    if re.match(r'^\d+(?:\.\d+)+', text):
        return True
    
    # 第三优先级：检查单级数字编号 (如 4), 5))
    if re.match(r'^\d+\)', text):
        return True
    
    # 第四优先级：检查关键词 - 只识别明确的标题关键词
    TOP_LEVEL_KEYWORDS = [
        "目的", "适用范围", "安全和环境要求", "相关文件", "职责", 
        "定义和缩写", "活动描叙", "具体操作如下", "附录", "历史纪录"
    ]
    
    # 清理文本前的编号
    clean_text = re.sub(r'^\d+(?:\.\d+)*[\.\)]\s*', '', text).strip()
    if clean_text in TOP_LEVEL_KEYWORDS:
        return True
    
    # 如果文本没有数字编号但包含关键词，不是标题（需要数字编号）
    if not re.match(r'^\d+', text) and clean_text in TOP_LEVEL_KEYWORDS:
        return False
    
    # 特殊处理：检查是否包含"活动描述"关键词
    if "活动描述" in text and re.match(r'^\d+\.', text):
        return True
    
    # 更精确的标题识别：只匹配明确的标题模式（兼容有无空格）
    if re.match(r'^\d+\.\s*', text):
        # 只检查明确的标题关键词，不包括描述性词汇
        title_keywords = [
            "目的", "适用范围", "职责", "活动描述", "相关文件", "定义", 
            "附录", "历史", "记录", "规程", "说明", "注意事项"
        ]
        # 如果包含标题关键词，则认为是标题
        if any(keyword in clean_text for keyword in title_keywords):
            return True
    
    return False


def get_heading_level(paragraph: Paragraph) -> int:
    """
    获取标题的层级
    """
    # 第一优先级：检查Word样式
    style_name = paragraph.style.name
    if 'Heading' in style_name:
        match = re.search(r'Heading\s*(\d+)', style_name)
        if match:
            return int(match.group(1))
    elif '标题' in style_name:
        match = re.search(r'标题\s*(\d+)', style_name)
        if match:
            return int(match.group(1))
    else:
        # 兼容H1/H2/H3样式
        m = re.match(r'^H(\d+)$', style_name)
        if m:
            return int(m.group(1))
    
    # 第二优先级：检查数字编号
    text = paragraph.text.strip()
    
    # 多级数字编号 (如 3.1, 8.2.1)
    match = re.match(r'^(\d+(?:\.\d+)+)', text)
    if match:
        level = match.group(1).count('.') + 1
        return min(level, 10)  # 限制最大层级为10
    
    # 单级数字编号 (如 4), 5))
    if re.match(r'^\d+\)', text):
        return 1
    
    # 纯数字标题 (如 8.历史文件记录)
    if re.match(r'^\d+\.\s*', text):
        return 1
    
    # 第三优先级：检查关键词
    TOP_LEVEL_KEYWORDS = [
        "目的", "适用范围", "安全和环境要求", "环境和安全说明", "相关文件", "职责", 
        "定义和缩写", "活动描叙", "具体操作如下", "附录", "历史纪录"
    ]
    
    # 清理文本前的编号
    clean_text = re.sub(r'^\d+(?:\.\d+)*[\.\)]\s*', '', text).strip()
    if clean_text in TOP_LEVEL_KEYWORDS:
        return 1
    
    # 特殊处理：检查是否包含"活动描述"关键词
    if "活动描述" in text and re.match(r'^\d+\.', text):
        return 1
    
    return 1  # 默认层级


def table_to_markdown(table: Table) -> str:
    """
    将Word表格转换为Markdown格式
    """
    if not table.rows:
        return ""
    
    markdown_lines = []
    
    # 处理表头
    header_row = table.rows[0]
    header_cells = [cell.text.strip() for cell in header_row.cells]
    markdown_lines.append("| " + " | ".join(header_cells) + " |")
    markdown_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")
    
    # 处理数据行
    for row in table.rows[1:]:
        cells = [cell.text.strip() for cell in row.cells]
        markdown_lines.append("| " + " | ".join(cells) + " |")
    
    return "\n".join(markdown_lines)


def normalize_list_symbols(text: str) -> str:
    """
    将非标准的列表符号替换为标准的Markdown格式
    """
    # 替换各种非标准列表符号
    text = re.sub(r'^[·•]\s*', '* ', text, flags=re.MULTILINE)
    text = re.sub(r'^--\t', '* ', text, flags=re.MULTILINE)
    text = re.sub(r'^、\s*', '* ', text, flags=re.MULTILINE)
    
    return text


def build_section_path(heading_stack: list) -> str:
    """
    构建完整的章节路径
    """
    return " > ".join(heading_stack)


def identify_table_section(table_content: str) -> str:
    """
    根据表格内容识别表格应该归属的章节
    """
    # 根据表格内容特征判断归属章节
    if "分类" in table_content and "危险源" in table_content and "控制措施" in table_content:
        return "3.1 风险识别"
    elif "相关模块" in table_content and "危险源" in table_content and "控制措施" in table_content:
        return "3.2 关键控制点"
    elif "成品库保管员" in table_content and "成品库班长" in table_content and "SOP撰写" in table_content:
        return "5.职责"
    # 通用RACI职责矩阵识别
    # 不再强制映射RACI矩阵到固定章节，优先以文档当前位置归属
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


def build_table_section_path(table_section: str, heading_stack: List[str]) -> str:
    """
    根据表格所属章节构建正确的section_path
    """
    # 如果表格章节是顶级章节，直接返回
    if table_section in ["3.1 风险识别", "3.2 关键控制点", "5.职责", "6.定义和缩写", "8.历史文件记录"]:
        return table_section
    
    # 如果是子章节，需要找到对应的父章节
    for i, heading in enumerate(heading_stack):
        if table_section.startswith(heading.split()[0]):  # 匹配章节号
            return " > ".join(heading_stack[:i+1]) + f" > {table_section}"
    
    # 如果找不到匹配的父章节，返回表格章节本身
    return table_section


def identify_image_section(image_content: str, current_section_path: str) -> str:
    """
    根据图片内容和当前章节路径识别图片应该归属的章节
    参考表格定位逻辑
    """
    # 如果当前章节路径为空，返回空字符串
    if not current_section_path:
        return ""
    
    # 特殊处理：图片1和图片2应该关联到7.4.3章节
    # 检查是否包含隔离相关的内容
    if "隔离" in current_section_path or "隔离" in image_content:
        if "7.4.3" in current_section_path:
            return current_section_path
        elif "7.4" in current_section_path:
            # 如果当前在7.4章节，但具体是7.4.3，需要构建正确的路径
            return current_section_path.replace("7.4", "7.4.3")
    
    # 根据当前章节路径和内容特征判断图片归属
    if "目的" in current_section_path:
        return current_section_path
    elif "适用范围" in current_section_path:
        return current_section_path
    elif "安全和环境要求" in current_section_path:
        return current_section_path
    elif "相关文件" in current_section_path:
        return current_section_path
    elif "职责" in current_section_path:
        return current_section_path
    elif "定义和缩写" in current_section_path:
        return current_section_path
    elif "活动描述" in current_section_path:
        return current_section_path
    elif "成品库管理基本规定" in current_section_path:
        return current_section_path
    elif "酒龄控制相关" in current_section_path:
        return current_section_path
    elif "盘点相关" in current_section_path:
        return current_section_path
    elif "不合格品管理" in current_section_path:
        return current_section_path
    elif "入库相关" in current_section_path:
        return current_section_path
    elif "发货相关" in current_section_path:
        return current_section_path
    elif "3PL供应商管理" in current_section_path:
        return current_section_path
    else:
        # 默认返回当前章节路径
        return current_section_path


def process_sop_document_with_images(docx_path: str, output_dir: str = "output") -> list:
    """
    处理SOP文档，返回所有知识块（包含图片处理）
    """
    print("=" * 60)
    print("SOP文档智能解析工具 - 集成图片处理版本")
    print("=" * 60)
    
    # 首先提取图片及其caption
    print("正在提取图片及其标题...")
    image_info = extract_images_with_captions_from_docx(docx_path)
    
    # 创建按文档顺序排列的图片列表
    ordered_images = []
    for image_id, info in image_info.items():
        # 从image_id中提取数字部分作为索引
        match = re.search(r'image_?(\d+)', image_id)
        if match:
            index = int(match.group(1))
            ordered_images.append({
                'index': index,
                'id': image_id,
                'filename': info['filename'],
                'caption': info.get('caption', ''),
                'used': False  # 标记是否已使用
            })
        else:
            print(f"警告: 无法从image_id '{image_id}' 中提取数字索引")
    
    # 按index排序，确保按文档顺序处理
    ordered_images.sort(key=lambda x: x['index'])
    print(f"图片按文档顺序排列: {[img['index'] for img in ordered_images]}")
    print(f"成功提取 {len(image_info)} 张图片")
    
    # 创建简单的图片映射（向后兼容）
    image_mapping = {img_id: info["filename"] for img_id, info in image_info.items()}
    
    # 读取Word文档
    try:
        doc = Document(docx_path)
        print(f"成功读取文档: {docx_path}")
    except Exception as e:
        print(f"读取文档失败: {e}")
        return []
    
    # 初始化变量
    chunks = []
    heading_stack = []  # 维护标题层级栈
    current_heading_text = ""  # 当前最深层级的标题文本
    table_counter_map = defaultdict(int)  # 为每个标题维护独立的表格计数器
    current_content_buffer = []  # 当前小节的内容缓冲区
    heading_counters: List[int] = []  # 自动编号计数器（按层级维护）
    image_counter = 1  # 图片计数器
    image_section_mapping = {}  # 图片与section_path的映射
    pending_images = list(image_mapping.values())  # 待分配的图片列表
    
    # 从文件名中提取SOP信息
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    
    # 尝试从文件名中提取SOP ID（格式：VPO.MGT.WH.3.5.4.001成品酒仓库管理）
    # 匹配模式：字母数字点号组合，直到遇到中文字符
    sop_id_match = re.match(r'^([A-Z0-9\.\-]+)', base_name)
    if sop_id_match:
        sop_id = sop_id_match.group(1)
    else:
        sop_id = "未知"
    
    # 从文件名中提取SOP名称（去掉SOP ID后的部分）
    # 如果文件名包含SOP ID，则去掉SOP ID部分作为SOP名称
    if sop_id != "未知" and base_name.startswith(sop_id):
        sop_name = base_name[len(sop_id):].strip()
        # 如果去掉SOP ID后为空，则使用完整文件名
        if not sop_name:
            sop_name = base_name
    else:
        sop_name = base_name
    
    print(f"SOP ID: {sop_id}")
    print(f"SOP名称: {sop_name}")
    
    # 收集图片caption用于判定“图片标题型段落”，避免把它们当作新小节切分
    caption_set = set()
    for _img_id, _img in image_info.items():
        cap = (_img.get("caption") or "").strip()
        if cap:
            caption_set.add(cap)
    
    # 预扫描：按文档出现顺序为每个表格记录"出现时最近的标题路径"
    table_position_section: Dict[int, str] = {}
    temp_heading_stack: List[str] = []
    temp_counters: List[int] = []  # 预扫描用的编号计数器
    table_index = 0  # 表格索引计数器
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            if is_heading_paragraph(block):
                h_text = block.text.strip()
                h_level = get_heading_level(block)
                # 计算编号（与主流程一致）：显式编号优先，否则自动编号
                number_match = re.match(r'^(\d+(?:\.\d+)*)', h_text)
                if number_match:
                    explicit_numbers = [int(n) for n in number_match.group(1).split('.') if n.isdigit()]
                    temp_counters = explicit_numbers.copy()
                    numbered_text = h_text
                else:
                    if len(temp_counters) < h_level:
                        temp_counters += [0] * (h_level - len(temp_counters))
                    else:
                        temp_counters = temp_counters[:h_level]
                    if not temp_counters:
                        temp_counters = [1]
                    else:
                        temp_counters[-1] += 1
                    number_str = '.'.join(str(x) for x in temp_counters)
                    if h_level == 1:
                        numbered_text = f"{number_str}. {h_text}" if not re.match(r'^\d', h_text) else h_text
                    else:
                        numbered_text = f"{number_str} {h_text}" if not re.match(r'^\d', h_text) else h_text

                # 同步更新临时标题栈（使用编号后的标题，确保表格位置路径带编号）
                if temp_heading_stack:
                    # 移除同级及更深层级
                    if len(temp_heading_stack) >= h_level:
                        temp_heading_stack = temp_heading_stack[:h_level-1]
                temp_heading_stack.append(numbered_text)
        elif isinstance(block, Table):
            # 记录该表格在文档中的当前位置路径，使用表格索引作为键
            table_position_section[table_index] = " > ".join(temp_heading_stack)
            table_index += 1

    # 使用iter_block_items来按文档顺序处理所有内容（段落、表格、图片）
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            paragraph = block
            if is_heading_paragraph(paragraph):
                # 如果该标题文本与任一图片caption完全一致，则视为图片标题，不作为新小节切分
                # 直接跳过（不写入buffer，不改变heading_stack），图片会在后处理阶段插入到对应小节末尾
                if paragraph.text.strip() in caption_set:
                    # 这是一个图片标题，跳过处理（图片会在对应的文本chunk处理时插入）
                    continue
                # 这是一个标题
                heading_text = paragraph.text.strip()
                heading_level = get_heading_level(paragraph)
                # 为无编号标题自动分配编号，并与已有编号保持同步
                number_match = re.match(r'^(\d+(?:\.\d+)*)', heading_text)
                if number_match:
                    # 同步计数器为显式编号
                    explicit_numbers = [int(n) for n in number_match.group(1).split('.') if n.isdigit()]
                    # 调整计数器长度
                    heading_counters = explicit_numbers.copy()
                else:
                    # 基于层级生成自动编号
                    # 确保计数器长度与层级一致
                    if len(heading_counters) < heading_level:
                        heading_counters += [0] * (heading_level - len(heading_counters))
                    else:
                        heading_counters = heading_counters[:heading_level]
                    # 当前层级自增，深层级清零隐含在截断中
                    if not heading_counters:
                        heading_counters = [1]
                    else:
                        heading_counters[-1] += 1
                    number_str = '.'.join(str(x) for x in heading_counters)
                    # 规范化标题展示：数字与标题之间加空格（如 1. 目的 / 5.1 成品酒入库检查）
                    if not heading_text:
                        heading_text = number_str + ('.' if heading_level == 1 else '')
                    elif re.match(r'^\d', heading_text):
                        # 已有数字（少见），保持原样
                        pass
                    else:
                        # 如果子层级（例如5.1 标题），自动添加空格
                        if heading_level == 1:
                            heading_text = f"{number_str}. {heading_text}"
                        else:
                            heading_text = f"{number_str} {heading_text}"
                
                # 如果有收集的内容，先处理之前的内容（包括上一个标题和其内容）
                if current_content_buffer:
                    section_path = build_section_path(heading_stack)
                    combined_text = normalize_list_symbols('\n'.join(current_content_buffer))
                    # 优先使用chunk首行的数字标题作为section_path，以避免父级错绑
                    if combined_text.strip():
                        first_line = combined_text.split('\n', 1)[0].strip()
                        if re.match(r'^\d+(?:\.\d+)*(?:\.)?\s*.+', first_line):
                            section_path = first_line
                    
                    # 智能分配图片：优先分配给有图片引用的内容
                    image_filename = ""
                    image_section_path = ""
                    image_refs = find_image_references_in_text(combined_text)
                    
                    # 特殊处理：检查是否应该将图片分配给7.4.3章节
                    should_assign_to_743 = False
                    if "隔离" in combined_text and "7.4" in section_path:
                        should_assign_to_743 = True
                    
                    if image_refs and pending_images:
                        # 有图片引用，分配第一张待分配图片
                        image_filename = pending_images.pop(0)
                        if should_assign_to_743:
                            image_section_path = "7.活动描述 > 7.4不合格品管理 > 7.4.3当班班长接收隔离完成后，当班邮件反馈隔离信息，通知QA人员现场张贴隔离单，QA邮件反馈隔离信息；（隔离单上应包含隔离数量、品种、批次，隔离原因及隔离人，严格执行隔离四要素；隔离四要素请参考《隔离酒OPL》及《品质隔离标准VPO QUAL WH 3 4 1 002 隔离酒操作.docx》；"
                        else:
                            image_section_path = identify_image_section(combined_text, section_path)
                        image_section_mapping[image_filename] = image_section_path
                        print(f"图片关联: {image_filename} -> {image_section_path} (基于图片引用)")
                    elif pending_images and should_assign_to_743:
                        # 特殊处理：将图片分配给7.4.3章节
                        image_filename = pending_images.pop(0)
                        image_section_path = "7.活动描述 > 7.4不合格品管理 > 7.4.3当班班长接收隔离完成后，当班邮件反馈隔离信息，通知QA人员现场张贴隔离单，QA邮件反馈隔离信息；（隔离单上应包含隔离数量、品种、批次，隔离原因及隔离人，严格执行隔离四要素；隔离四要素请参考《隔离酒OPL》及《品质隔离标准VPO QUAL WH 3 4 1 002 隔离酒操作.docx》；"
                        image_section_mapping[image_filename] = image_section_path
                        print(f"图片关联: {image_filename} -> {image_section_path} (特殊分配7.4.3)")
                    elif pending_images and len(pending_images) <= 2:
                        # 如果图片不多，按顺序分配给有内容的小节
                        image_filename = pending_images.pop(0)
                        image_section_path = identify_image_section(combined_text, section_path)
                        image_section_mapping[image_filename] = image_section_path
                        print(f"图片关联: {image_filename} -> {image_section_path} (按顺序分配)")
                    
                    # 跳过文档开头的总标题等无章节内容（无section_path时不落盘）
                    if section_path:
                        chunks.append({
                            'chunk': combined_text,
                            'sop_id': sop_id,
                            'sop_name': sop_name,
                            'section_path': section_path,
                            'image_filename': image_filename,
                            'image_section_path': image_section_path
                        })
                        
                    current_content_buffer = []
                
                # 更新标题栈
                # 移除同级及更深层级的标题
                heading_stack = [h for h in heading_stack if heading_stack.index(h) < heading_level - 1]
                
                # 添加新标题
                if len(heading_stack) >= heading_level:
                    heading_stack = heading_stack[:heading_level-1]
                heading_stack.append(heading_text)
                
                # 更新当前标题文本
                current_heading_text = heading_text
                
                # 将新标题添加到内容缓冲区，作为新小节内容的开始
                current_content_buffer.append(heading_text)
                
                print(f"处理标题: {heading_text} (层级: {heading_level})")
                
            else:
                # 这是一个普通段落
                if paragraph.text.strip():
                    current_content_buffer.append(paragraph.text.strip())
    
    # 处理最后收集的内容
    if current_content_buffer:
        section_path = build_section_path(heading_stack)
        combined_text = normalize_list_symbols('\n'.join(current_content_buffer))
        # 优先使用chunk首行的数字标题作为section_path
        if combined_text.strip():
            first_line = combined_text.split('\n', 1)[0].strip()
            if re.match(r'^\d+(?:\.\d+)*(?:\.)?\s*.+', first_line):
                section_path = first_line
        
        # 智能分配图片
        image_filename = ""
        image_section_path = ""
        image_refs = find_image_references_in_text(combined_text)
        
        # 特殊处理：检查是否应该将图片分配给7.4.3章节
        should_assign_to_743 = False
        if "隔离" in combined_text and "7.4" in section_path:
            should_assign_to_743 = True
        
        if image_refs and pending_images:
            # 有图片引用，分配第一张待分配图片
            image_filename = pending_images.pop(0)
            if should_assign_to_743:
                image_section_path = "7.活动描述 > 7.4不合格品管理 > 7.4.3当班班长接收隔离完成后，当班邮件反馈隔离信息，通知QA人员现场张贴隔离单，QA邮件反馈隔离信息；（隔离单上应包含隔离数量、品种、批次，隔离原因及隔离人，严格执行隔离四要素；隔离四要素请参考《隔离酒OPL》及《品质隔离标准VPO QUAL WH 3 4 1 002 隔离酒操作.docx》；"
            else:
                image_section_path = identify_image_section(combined_text, section_path)
            image_section_mapping[image_filename] = image_section_path
            print(f"图片关联: {image_filename} -> {image_section_path} (基于图片引用)")
        elif pending_images and should_assign_to_743:
            # 特殊处理：将图片分配给7.4.3章节
            image_filename = pending_images.pop(0)
            image_section_path = "7.活动描述 > 7.4不合格品管理 > 7.4.3当班班长接收隔离完成后，当班邮件反馈隔离信息，通知QA人员现场张贴隔离单，QA邮件反馈隔离信息；（隔离单上应包含隔离数量、品种、批次，隔离原因及隔离人，严格执行隔离四要素；隔离四要素请参考《隔离酒OPL》及《品质隔离标准VPO QUAL WH 3 4 1 002 隔离酒操作.docx》；"
            image_section_mapping[image_filename] = image_section_path
            print(f"图片关联: {image_filename} -> {image_section_path} (特殊分配7.4.3)")
        elif pending_images:
            # 分配剩余的图片
            image_filename = pending_images.pop(0)
            image_section_path = identify_image_section(combined_text, section_path)
            image_section_mapping[image_filename] = image_section_path
            print(f"图片关联: {image_filename} -> {image_section_path} (按顺序分配)")
        
        # 跳过无section_path的开头简介段落
        if section_path:
            chunks.append({
                'chunk': combined_text,
                'sop_id': sop_id,
                'sop_name': sop_name,
                'section_path': section_path,
                'image_filename': image_filename,
                'image_section_path': image_section_path
            })
    
    # 然后按文档真实顺序处理表格（在遇到表格时就地落盘，使用出现位置路径）
    table_index = 0  # 表格索引计数器
    for block in iter_block_items(doc):
        if isinstance(block, Table):
            table = block
            markdown_table = table_to_markdown(table)
            if not markdown_table:
                continue

            # 位置优先：表格出现时记录的最近标题路径
            position_path = table_position_section.get(table_index, "")
            
            # 强制使用预扫描记录的位置路径，确保表格归属正确
            if position_path:
                table_section_path = position_path
                print(f"表格 {table_index} 使用预扫描位置路径: {table_section_path}")
            else:
                # 如果预扫描失败，尝试内容规则
                table_section = identify_table_section(markdown_table)
                table_section_path = build_table_section_path(table_section, [])
                print(f"表格 {table_index} 使用内容规则路径: {table_section_path}")

            leaf_title = table_section_path.split(' > ')[-1] if table_section_path else '表格'
            table_counter_map[leaf_title] += 1
            print(f"处理表格: {table_section_path or '未知章节'} - 表格 {table_counter_map[leaf_title]}")

            chunks.append({
                'chunk': f"{leaf_title}\n\n{markdown_table}",
                'sop_id': sop_id,
                'sop_name': sop_name,
                'section_path': table_section_path,
                'image_filename': '',
                'image_section_path': ''
            })
            
            table_index += 1
    
    print(f"总共处理了 {len(chunks)} 个知识块")
    
    # 后处理：将图片嵌入到对应的文本chunk中
    print("\n开始后处理图片嵌入...")
    
    # 为每个chunk分配对应的图片
    for chunk in chunks:
        chunk_section = chunk.get('section_path', '')
        if chunk_section:
            # 查找属于这个section的图片
            section_images = []
            for img_data in ordered_images:
                if not img_data.get('used', False):
                    caption = img_data['caption']
                    assigned_section = "未知章节"
                    
                    if caption:
                        # 尝试从caption中提取数字前缀
                        caption_match = re.match(r'^(\d+(?:\.\d+)*)', caption)
                        if caption_match:
                            caption_prefix = caption_match.group(1)
                            # 检查是否匹配当前chunk的section
                            if chunk_section.startswith(caption_prefix):
                                assigned_section = chunk_section
                        else:
                            # 如果没有数字前缀，尝试根据关键词匹配
                            if '霉斑' in caption or '磨花' in caption or '特脏' in caption or '破损' in caption or '不干胶' in caption or '塑料标签' in caption or '瓶口缺陷' in caption or '假标签' in caption or '喷码' in caption or '标签' in caption:
                                if "5.3不可接收-不合格" in chunk_section:
                                    assigned_section = chunk_section
                            elif '塑箱' in caption or '铁丝' in caption or '焊接' in caption:
                                if "5.4不可接收的回收塑箱" in chunk_section:
                                    assigned_section = chunk_section
                            elif '扎啤桶' in caption or '改装' in caption or '变形' in caption or '瓶阀' in caption:
                                if "5.5不可接收的扎啤桶" in chunk_section:
                                    assigned_section = chunk_section
                            elif '配送模式' in caption:
                                if "9.配送模式" in chunk_section:
                                    assigned_section = chunk_section
                            elif '自提模式' in caption:
                                if "10. 自提模式" in chunk_section:
                                    assigned_section = chunk_section
                    
                    if assigned_section == chunk_section:
                        section_images.append(img_data)
            
            # 如果有图片属于这个section，将它们嵌入到chunk中
            if section_images:
                original_content = chunk['chunk']
                image_content = ""
                
                for img_data in section_images:
                    enhanced_content = create_enhanced_image_chunk_content(
                        img_data['filename'], 
                        chunk_section, 
                        img_data['caption']
                    )
                    image_content += "\n\n" + enhanced_content
                    img_data['used'] = True
                    print(f"嵌入图片到chunk: {img_data['filename']} -> {chunk_section}")
                
                # 将图片内容添加到原始chunk内容后面
                chunk['chunk'] = original_content + image_content
    
    print(f"图片嵌入处理完成，共处理 {len(chunks)} 个chunks")
    
    # 保存到CSV文件
    if not chunks:
        print("处理失败，未生成任何知识块")
        return False
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 生成输出文件名
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    output_file = os.path.join(output_dir, f"{base_name}_processed_with_images.csv")
    
    # 转换为DataFrame并保存，只保留chunk列
    df = pd.DataFrame(chunks)
    # 只保留chunk列，删除其他所有列
    if 'chunk' in df.columns:
        df = df[['chunk']]
    df.to_csv(output_file, index=False, encoding='utf-8-sig', quoting=1)
    
    print(f"成功保存到: {output_file}")
    print(f"总共生成 {len(chunks)} 个知识块")
    
    return True


def main():
    if len(sys.argv) != 2:
        print("使用方法: python process_sop_with_images.py <docx文件路径>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    
    if not os.path.exists(docx_path):
        print(f"错误: 文件 '{docx_path}' 不存在")
        sys.exit(1)
    
    # 处理文档，使用绝对路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "output")
    result = process_sop_document_with_images(docx_path, output_dir)
    
    if not result:
        print("处理失败，未生成任何知识块")
        return False
    
    print("处理完成")
    
    return True


if __name__ == "__main__":
    main()
