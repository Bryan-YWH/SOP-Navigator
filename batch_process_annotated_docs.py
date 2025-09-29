#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量处理已标注文档脚本
使用一体化处理流程处理"标注完的文档"文件夹中的所有Word文档
"""

import os
import sys
import glob
from pathlib import Path
from process_sop_with_images import process_sop_document_with_images

def batch_process_annotated_documents():
    """批量处理已标注的文档"""
    
    # 设置路径
    current_dir = Path(__file__).parent
    annotated_dir = current_dir / "标注完的文档"
    output_dir = current_dir / "output"
    images_dir = current_dir / "sop_images"
    
    # 确保输出目录存在
    output_dir.mkdir(exist_ok=True)
    images_dir.mkdir(exist_ok=True)
    
    # 查找所有Word文档
    docx_files = list(annotated_dir.glob("*.docx"))
    doc_files = list(annotated_dir.glob("*.doc"))
    
    all_files = docx_files + doc_files
    
    if not all_files:
        print("❌ 在'标注完的文档'文件夹中未找到任何Word文档")
        return
    
    print(f"📁 找到 {len(all_files)} 个文档待处理:")
    for i, file_path in enumerate(all_files, 1):
        print(f"  {i}. {file_path.name}")
    
    print(f"\n🚀 开始批量处理...")
    print("=" * 60)
    
    success_count = 0
    error_count = 0
    processed_files = []
    
    for i, docx_path in enumerate(all_files, 1):
        print(f"\n📄 处理文档 {i}/{len(all_files)}: {docx_path.name}")
        print("-" * 40)
        
        try:
            # 处理文档
            # 输出CSV到 output_dir，图片提取由处理函数内部固定到 sop_images
            result = process_sop_document_with_images(str(docx_path), str(output_dir))
            
            if result:
                success_count += 1
                processed_files.append(docx_path.name)
                print(f"✅ 成功处理: {docx_path.name}")
            else:
                error_count += 1
                print(f"❌ 处理失败: {docx_path.name}")
                
        except Exception as e:
            error_count += 1
            print(f"❌ 处理出错: {docx_path.name}")
            print(f"   错误信息: {str(e)}")
    
    # 输出处理结果统计
    print("\n" + "=" * 60)
    print("📊 批量处理完成统计:")
    print(f"✅ 成功处理: {success_count} 个文档")
    print(f"❌ 处理失败: {error_count} 个文档")
    print(f"📁 总文档数: {len(all_files)} 个")
    
    if processed_files:
        print(f"\n📋 成功处理的文档列表:")
        for i, filename in enumerate(processed_files, 1):
            print(f"  {i}. {filename}")
    
    # 显示输出文件
    print(f"\n📂 输出文件位置:")
    print(f"  CSV文件: {output_dir}")
    print(f"  图片文件: {images_dir}")
    
    # 列出生成的CSV文件
    csv_files = list(output_dir.glob("*_processed_with_images.csv"))
    if csv_files:
        print(f"\n📄 生成的CSV文件 ({len(csv_files)} 个):")
        for csv_file in sorted(csv_files):
            print(f"  - {csv_file.name}")
    
    # 列出生成的图片文件
    image_files = list(images_dir.glob("*.png")) + list(images_dir.glob("*.jpg"))
    if image_files:
        print(f"\n🖼️  提取的图片文件 ({len(image_files)} 个):")
        for image_file in sorted(image_files):
            print(f"  - {image_file.name}")

if __name__ == "__main__":
    batch_process_annotated_documents()
