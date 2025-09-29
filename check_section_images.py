#!/usr/bin/env python3
"""
检查每个小节对应的图片数量
"""

import re
import os

def check_section_images():
    csv_file = "output/China RTP-001 可回收包装物接收政策_processed_with_images.csv"
    
    if not os.path.exists(csv_file):
        print(f"❌ 文件不存在: {csv_file}")
        return
    
    print("=" * 80)
    print("China RTP-001 可回收包装物接收政策 - 章节图片统计")
    print("=" * 80)
    
    with open(csv_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 按行分割内容
    lines = content.split('\n')
    
    current_section = ""
    section_images = {}
    image_count = 0
    
    for line in lines:
        line = line.strip()
        if not line or line.startswith('"chunk"'):
            continue
            
        # 移除CSV的引号
        line = line.strip('"')
        
        # 检查是否是章节标题（以数字开头）
        section_match = re.match(r'^(\d+(?:\.\d+)*)\s*[^图]', line)
        if section_match:
            current_section = line
            if current_section not in section_images:
                section_images[current_section] = []
            continue
        
        # 检查是否包含图片
        if '[图：' in line:
            image_count += 1
            if current_section:
                section_images[current_section].append(line)
    
    # 输出统计结果
    print(f"📊 总图片数量: {image_count}")
    print(f"📋 章节数量: {len(section_images)}")
    print("\n" + "=" * 80)
    print("各章节图片统计:")
    print("=" * 80)
    
    for section, images in section_images.items():
        print(f"\n📁 {section}")
        print(f"   图片数量: {len(images)}")
        if images:
            for i, img in enumerate(images, 1):
                # 提取图片名称
                img_match = re.search(r'\[图：([^\]]+)\]', img)
                if img_match:
                    img_name = img_match.group(1)
                    print(f"   {i}. {img_name}")
    
    # 检查未知章节
    unknown_images = []
    for line in lines:
        line = line.strip().strip('"')
        if '未知章节' in line and '[图：' in line:
            unknown_images.append(line)
    
    if unknown_images:
        print(f"\n❓ 未知章节图片 ({len(unknown_images)}张):")
        for i, img in enumerate(unknown_images, 1):
            img_match = re.search(r'\[图：([^\]]+)\]', img)
            if img_match:
                img_name = img_match.group(1)
                print(f"   {i}. {img_name}")

if __name__ == "__main__":
    check_section_images()
