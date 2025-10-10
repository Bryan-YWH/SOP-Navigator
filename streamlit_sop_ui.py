#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit UI for SOP Document Processing
智能SOP文档处理系统的Web界面
"""

import streamlit as st
import os
import tempfile
import zipfile
from pathlib import Path
import pandas as pd
from typing import List, Tuple
import time
import shutil
from process_sop_with_images import process_sop_document_with_images

# 页面配置
st.set_page_config(
    page_title="SOP Navigator - 智能文档处理系统",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 2rem;
    }
    .upload-section {
        border: 2px dashed #4CAF50;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background-color: #f8f9fa;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def init_session_state():
    """初始化会话状态"""
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = []
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'output_dir' not in st.session_state:
        st.session_state.output_dir = None

def create_output_directories() -> Tuple[Path, Path]:
    """创建输出目录"""
    current_dir = Path(__file__).parent
    output_dir = current_dir / "streamlit_output"
    images_dir = current_dir / "sop_images"
    
    output_dir.mkdir(exist_ok=True)
    images_dir.mkdir(exist_ok=True)
    
    return output_dir, images_dir

def is_word_document(filename: str) -> bool:
    """检查文件是否为Word文档"""
    return filename.lower().endswith(('.docx', '.doc'))

def create_images_zip(images_dir: Path) -> bytes:
    """创建包含所有图片的ZIP文件"""
    import io
    
    zip_buffer = io.BytesIO()
    
    # 获取所有图片文件
    image_files = list(images_dir.glob("*.png")) + list(images_dir.glob("*.jpg")) + list(images_dir.glob("*.jpeg"))
    
    if not image_files:
        return b""
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for img_file in sorted(image_files):
            # 使用相对路径作为ZIP内的文件名
            zip_file.write(img_file, img_file.name)
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def save_uploaded_file(uploaded_file, temp_dir: Path) -> Path:
    """保存上传的文件到临时目录"""
    file_path = temp_dir / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

def extract_zip_files(zip_file, temp_dir: Path) -> List[Path]:
    """从ZIP文件中提取Word文档"""
    word_files = []
    
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
        
        # 递归查找所有Word文档
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                if is_word_document(file):
                    word_files.append(Path(root) / file)
    
    return word_files

def process_single_document(doc_path: Path, output_dir: Path) -> Tuple[bool, str]:
    """处理单个文档"""
    try:
        result = process_sop_document_with_images(str(doc_path), str(output_dir))
        if result:
            return True, f"✅ 成功处理: {doc_path.name}"
        else:
            return False, f"❌ 处理失败: {doc_path.name}"
    except Exception as e:
        return False, f"❌ 处理出错: {doc_path.name} - {str(e)}"

def display_results(processed_files: List[dict], output_dir: Path, images_dir: Path, context: str = "main"):
    """显示处理结果"""
    # 如果是历史记录且没有processed_files，显示所有可用文件
    show_all_files = (context == "history" and not processed_files)
    
    if not processed_files and not show_all_files:
        return
    
    success_count = sum(1 for f in processed_files if f['success']) if processed_files else 0
    error_count = len(processed_files) - success_count if processed_files else 0
    
    # 显示统计信息
    if not show_all_files:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("总文档数", len(processed_files))
        with col2:
            st.metric("成功处理", success_count, delta=success_count)
        with col3:
            st.metric("处理失败", error_count, delta=-error_count if error_count > 0 else 0)
        
        # 显示详细结果
        st.subheader("📋 处理详情")
        for file_info in processed_files:
            if file_info['success']:
                st.success(file_info['message'])
            else:
                st.error(file_info['message'])
    
    # 显示生成的文件
    if show_all_files:
        # 历史记录模式：显示所有文件
        csv_files = list(output_dir.glob("*_processed_with_images.csv"))
        image_files = list(images_dir.glob("*.png")) + list(images_dir.glob("*.jpg")) + list(images_dir.glob("*.jpeg"))
    else:
        # 当前会话模式：只显示当前处理会话相关的文件
        # 从processed_files获取成功处理的文件名（不含扩展名）
        processed_doc_names = []
        for file_info in processed_files:
            if file_info['success']:
                # 移除文件扩展名，获取基础名称
                base_name = file_info['filename'].replace('.docx', '').replace('.doc', '')
                processed_doc_names.append(base_name)
        
        # 只显示与当前处理会话相关的CSV文件
        all_csv_files = list(output_dir.glob("*_processed_with_images.csv"))
        csv_files = []
        for csv_file in all_csv_files:
            csv_base_name = csv_file.name.replace('_processed_with_images.csv', '')
            if csv_base_name in processed_doc_names:
                csv_files.append(csv_file)
        
        # 只显示与当前处理会话相关的图片文件
        all_image_files = list(images_dir.glob("*.png")) + list(images_dir.glob("*.jpg")) + list(images_dir.glob("*.jpeg"))
        image_files = []
        for img_file in all_image_files:
            # 检查图片文件名是否包含当前处理的文档名
            for doc_name in processed_doc_names:
                if doc_name in img_file.name:
                    image_files.append(img_file)
                    break
    
    if csv_files:
        st.subheader(f"📄 生成的CSV文件 ({len(csv_files)} 个)")
        
        # 显示CSV文件路径和统计信息
        col_csv1, col_csv2 = st.columns(2)
        with col_csv1:
            st.info(f"📁 CSV保存路径: `{output_dir}`")
        with col_csv2:
            total_size = sum(csv_file.stat().st_size for csv_file in csv_files)
            st.info(f"📊 总大小: {total_size / 1024:.1f} KB")
        
        # 显示每个CSV文件
        for i, csv_file in enumerate(sorted(csv_files)):
            col_file, col_btn = st.columns([2, 1])
            
            with col_file:
                file_size = csv_file.stat().st_size / 1024  # KB
                st.write(f"📄 {csv_file.name} ({file_size:.1f} KB)")
            
            with col_btn:
                # 提供下载按钮 - 使用上下文和索引创建唯一key
                with open(csv_file, 'rb') as f:
                    st.download_button(
                        label="⬇️ 下载",
                        data=f.read(),
                        file_name=csv_file.name,
                        mime="text/csv",
                        key=f"download_csv_{context}_{i}_{csv_file.stem}"
                    )
    
    if image_files:
        # 创建标题和下载按钮的列布局
        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader(f"🖼️ 提取的图片文件 ({len(image_files)} 个)")
        with col2:
            # 创建ZIP文件并提供下载
            zip_data = create_images_zip(images_dir)
            if zip_data:
                st.download_button(
                    label="📦 下载所有图片 (ZIP)",
                    data=zip_data,
                    file_name=f"sop_images_{time.strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    key=f"download_images_zip_{context}"
                )
        
        # 显示图片路径和统计信息
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            st.info(f"📁 图片保存路径: `{images_dir}`")
        with col_info2:
            png_count = len([f for f in image_files if f.suffix.lower() == '.png'])
            jpg_count = len([f for f in image_files if f.suffix.lower() in ['.jpg', '.jpeg']])
            st.info(f"📊 格式统计: PNG ({png_count}), JPG ({jpg_count})")
        
        # 图片预览（显示前几张）
        preview_count = min(6, len(image_files))
        cols = st.columns(3)
        for i, img_file in enumerate(sorted(image_files)[:preview_count]):
            with cols[i % 3]:
                try:
                    st.image(str(img_file), caption=img_file.name, width=200)
                except Exception as e:
                    st.write(f"🖼️ {img_file.name}")
        
        if len(image_files) > preview_count:
            st.write(f"... 还有 {len(image_files) - preview_count} 张图片")

def main():
    """主函数"""
    init_session_state()
    
    # 页面标题
    st.markdown("""
    <div class="main-header">
        <h1>📚 SOP Navigator</h1>
        <h3>智能SOP文档处理系统</h3>
        <p>将Word格式的SOP文档转换为结构化CSV数据，支持图片提取和表格处理</p>
    </div>
    """, unsafe_allow_html=True)
    
    # 侧边栏说明
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        ### 支持的文件格式：
        - Word文档 (.docx, .doc)
        - ZIP压缩包（包含Word文档）
        
        ### 处理功能：
        - ✅ 智能标题识别
        - ✅ 表格转换与归属
        - ✅ 图片自动提取
        - ✅ 内容结构化处理
        - ✅ 批量文档处理
        
        ### 输出结果：
        - 📄 结构化CSV文件 (streamlit_output/)
        - 🖼️ 提取的图片文件 (sop_images/)
        - 📦 图片批量下载 (ZIP压缩包)
        - 📊 处理统计报告
        
        ### 🚨 常见问题解决：
        - **403错误**: 文件过大或网络问题，尝试刷新页面
        - **上传失败**: 确保文件格式正确(.docx/.doc/.zip)
        - **处理超时**: 大文件需要更长时间，请耐心等待
        """)
        
        st.header("🔧 系统信息")
        st.info("基于Python-docx和Pandas构建")
        st.success("✅ 当前配置: 最大文件200MB, CORS已禁用")
        
        # 显示当前文件状态
        current_dir = Path(__file__).parent
        output_dir = current_dir / "streamlit_output"
        images_dir = current_dir / "sop_images"
        
        if output_dir.exists():
            existing_csvs = list(output_dir.glob("*_processed_with_images.csv"))
            if existing_csvs:
                total_csv_size = sum(csv.stat().st_size for csv in existing_csvs) / 1024
                st.success(f"📄 已有 {len(existing_csvs)} 个CSV文件 ({total_csv_size:.1f}KB)")
            else:
                st.info("📄 CSV目录为空，处理文档后将生成CSV")
        
        if images_dir.exists():
            existing_images = list(images_dir.glob("*.png")) + list(images_dir.glob("*.jpg")) + list(images_dir.glob("*.jpeg"))
            if existing_images:
                st.success(f"�️ 已有 {len(existing_images)} 张提取的图片")
            else:
                st.info("�️ 图片目录为空，处理文档后将显示图片")
    
    # 主要内容区域
    tab1, tab2 = st.tabs(["📁 文档上传处理", "📊 处理历史"])
    
    with tab1:
        st.markdown("""
        <div class="upload-section">
            <h3>📁 拖拽上传文档</h3>
            <p>支持单个文档或文件夹（ZIP格式）上传</p>
        </div>
        """, unsafe_allow_html=True)
        
        # 文件上传组件
        try:
            uploaded_files = st.file_uploader(
                "选择Word文档或ZIP文件",
                type=['docx', 'doc', 'zip'],
                accept_multiple_files=True,
                help="支持拖拽多个文件或ZIP压缩包（单个文件最大200MB）"
            )
        except Exception as e:
            st.error(f"文件上传出错: {str(e)}")
            st.info("💡 提示：如果遇到403错误，请尝试刷新页面或使用较小的文件")
            uploaded_files = None
        
        if uploaded_files:
            # 验证文件大小
            max_size = 200 * 1024 * 1024  # 200MB
            valid_files = []
            
            for file in uploaded_files:
                if file.size > max_size:
                    st.error(f"❌ 文件 {file.name} 太大 ({file.size / 1024 / 1024:.1f}MB)，请使用小于200MB的文件")
                else:
                    valid_files.append(file)
            
            if valid_files:
                st.success(f"已上传 {len(valid_files)} 个有效文件")
                
                # 显示上传的文件列表
                st.subheader("📋 上传文件列表")
                for file in valid_files:
                    file_type = "ZIP压缩包" if file.name.endswith('.zip') else "Word文档"
                    file_size_mb = file.size / 1024 / 1024
                    st.write(f"📄 {file.name} ({file_type}, {file_size_mb:.1f}MB)")
                
                uploaded_files = valid_files  # 只处理有效文件
            else:
                st.warning("没有有效的文件可以处理")
                uploaded_files = None
            
            # 处理按钮
            if st.button("🚀 开始处理文档", type="primary", use_container_width=True, key="process_documents_btn"):
                output_dir, images_dir = create_output_directories()
                
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_path = Path(temp_dir)
                    all_word_files = []
                    
                    # 处理上传的文件
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for i, uploaded_file in enumerate(uploaded_files):
                        status_text.text(f"正在处理文件 {i+1}/{len(uploaded_files)}: {uploaded_file.name}")
                        
                        if uploaded_file.name.endswith('.zip'):
                            # 处理ZIP文件
                            zip_path = save_uploaded_file(uploaded_file, temp_path)
                            word_files = extract_zip_files(zip_path, temp_path)
                            all_word_files.extend(word_files)
                        else:
                            # 处理Word文档
                            if is_word_document(uploaded_file.name):
                                word_path = save_uploaded_file(uploaded_file, temp_path)
                                all_word_files.append(word_path)
                        
                        progress_bar.progress((i + 1) / len(uploaded_files) / 2)
                    
                    if not all_word_files:
                        st.error("❌ 未找到可处理的Word文档！")
                        return
                    
                    st.info(f"📄 找到 {len(all_word_files)} 个Word文档，开始处理...")
                    
                    # 处理Word文档
                    processed_files = []
                    for i, doc_path in enumerate(all_word_files):
                        status_text.text(f"正在处理文档 {i+1}/{len(all_word_files)}: {doc_path.name}")
                        
                        success, message = process_single_document(doc_path, output_dir)
                        processed_files.append({
                            'filename': doc_path.name,
                            'success': success,
                            'message': message
                        })
                        
                        progress_bar.progress(0.5 + (i + 1) / len(all_word_files) / 2)
                    
                    status_text.text("处理完成！")
                    progress_bar.progress(1.0)
                    
                    # 保存处理结果到会话状态
                    st.session_state.processed_files = processed_files
                    st.session_state.processing_complete = True
                    st.session_state.output_dir = output_dir
                    
                    # 显示结果
                    st.success("🎉 文档处理完成！")
                    display_results(processed_files, output_dir, images_dir, "processing")
    
    with tab2:
        st.subheader("📊 处理历史")
        
        if st.session_state.processing_complete and st.session_state.processed_files:
            output_dir, images_dir = create_output_directories()
            display_results(st.session_state.processed_files, output_dir, images_dir, "history")
            
            # 清除历史按钮
            if st.button("🗑️ 清除处理历史", type="secondary", key="clear_history_btn"):
                st.session_state.processed_files = []
                st.session_state.processing_complete = False
                st.session_state.output_dir = None
                st.rerun()
        else:
            st.info("📝 暂无处理历史记录。请先在上传处理页面处理文档。")
    
    # 页脚
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🔧 SOP Navigator v1.0 | 智能SOP文档处理系统</p>
        <p>💡 支持Word文档批量处理、图片提取、表格转换等功能</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()