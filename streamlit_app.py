import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os
import tempfile
from datetime import datetime
from dashscope import Generation
import json
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import numpy as np
import io

# 设置页面配置
st.set_page_config(
    page_title="条款合规性对比工具",
    page_icon="📄",
    layout="wide"
)

# 页面标题
st.title("📄 条款合规性对比工具 (优化版)")
st.write("上传基准文件和待比较文件，系统将优先通过文本提取分析，必要时使用OCR识别图像内容。")

# 检查Tesseract是否安装
def check_tesseract_installation():
    try:
        # 尝试运行Tesseract版本检查
        pytesseract.get_tesseract_version()
        return True
    except pytesseract.TesseractNotFoundError:
        return False

# 检查Tesseract状态
tesseract_available = check_tesseract_installation()

# 侧边栏配置
with st.sidebar:
    st.subheader("配置选项")
    
    # Qwen API配置
    st.subheader("Qwen大模型配置")
    qwen_api_key = st.text_input("请输入阿里云DashScope API密钥", type="password")
    if qwen_api_key:
        os.environ["DASHSCOPE_API_KEY"] = qwen_api_key
    
    # Tesseract配置（如果可用）
    if not tesseract_available:
        st.warning("未检测到Tesseract OCR引擎，图片型PDF将无法处理")
        st.info("""
        安装Tesseract指南：
        1. 下载安装包：https://github.com/UB-Mannheim/tesseract/wiki
        2. 安装时勾选中文语言包
        3. 配置环境变量或在应用中指定路径
        """)
    else:
        tesseract_path = st.text_input(
            "Tesseract安装路径", 
            value=pytesseract.pytesseract.tesseract_cmd,
            help="默认路径: C:\\Program Files\\Tesseract-OCR\\tesseract.exe (Windows) 或 /usr/bin/tesseract (Linux)"
        )
        if tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path

# 辅助函数：从docx文件中提取文本
def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():  # 只添加非空段落
            full_text.append(para.text)
    return '\n'.join(full_text)

# 辅助函数：检查PDF页面是否包含可复制文本
def has_selectable_text(page):
    text = page.get_text("text")
    # 排除仅包含少量字符的情况（可能是页眉页脚）
    clean_text = text.strip()
    return len(clean_text) > 50  # 认为50个字符以上是有效文本

# 辅助函数：从PDF中提取文本（优先文本提取，必要时OCR）
def extract_text_from_pdf(file_path):
    doc = fitz.open(file_path)
    full_text = []
    
    for page_num, page in enumerate(doc):
        # 先尝试提取文本
        if has_selectable_text(page):
            page_text = page.get_text("text")
            if page_text.strip():
                full_text.append(f"[第{page_num+1}页 - 文本提取]\n{page_text}")
                continue
        
        # 如果文本提取失败且Tesseract可用，则使用OCR
        if tesseract_available:
            try:
                # 将页面转换为图片
                pix = page.get_pixmap(dpi=300)  # 高DPI提高识别精度
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                # 预处理图像提高识别率
                img = img.convert('L')  # 转为灰度图
                img_np = np.array(img)
                # 二值化处理（简单阈值）
                threshold = 150
                img_np = (img_np > threshold) * 255
                img = Image.fromarray(img_np.astype(np.uint8))
                
                # 进行OCR识别（中英文）
                ocr_text = pytesseract.image_to_string(
                    img, 
                    lang="chi_sim+eng",
                    config='--psm 6'  # 假设图片是单一均匀的文本块
                )
                
                if ocr_text.strip():
                    full_text.append(f"[第{page_num+1}页 - OCR识别]\n{ocr_text}")
                else:
                    full_text.append(f"[第{page_num+1}页 - 无法提取文本]")
                    
            except Exception as e:
                full_text.append(f"[第{page_num+1}页 - OCR处理失败: {str(e)}]")
        else:
            full_text.append(f"[第{page_num+1}页 - 包含图片内容，需要Tesseract OCR处理]")
    
    doc.close()
    return '\n\n'.join(full_text)

# 统一的文件提取函数
def extract_text_from_file(uploaded_file, file_type):
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_type}") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
        
        # 根据文件类型提取文本
        if file_type == "docx":
            text = extract_text_from_docx(tmp_path)
        elif file_type == "pdf":
            text = extract_text_from_pdf(tmp_path)
        else:
            text = ""
        
        # 清理临时文件
        os.unlink(tmp_path)
        return text
    
    except Exception as e:
        st.error(f"文件处理错误: {str(e)}")
        return ""

# 优化的中文条款拆分函数
def split_chinese_terms(text):
    # 中文条款常见的编号模式
    patterns = [
        r'(第[一二三四五六七八九十百]+条\s?)',  # 第一条 第二款
        r'([一二三四五六七八九十]+、\s?)',       # 一、 二、
        r'(\([一二三四五六七八九十]+\)\s?)',     # (一) (二)
        r'(\d+\.\s?)',                             # 1. 2.1.
        r'(\(\d+\)\s?)',                           # (1) (2)
        r'([A-Za-z]\.\s?)'                         # A. B.
    ]
    
    # 组合所有模式
    combined_pattern = '|'.join(patterns)
    
    # 拆分文本
    parts = re.split(combined_pattern, text)
    
    # 重组条款
    terms = []
    current_term = ""
    
    for part in parts:
        if not part.strip():
            continue
            
        # 检查当前部分是否是条款编号
        is_numbering = any(re.fullmatch(pattern.strip(), part.strip()) for pattern in patterns)
        
        if is_numbering:
            if current_term:  # 如果已有内容，保存当前条款
                terms.append(current_term.strip())
            current_term = part  # 开始新条款
        else:
            current_term += part  # 添加到当前条款
    
    # 添加最后一个条款
    if current_term.strip():
        terms.append(current_term.strip())
    
    return terms

# 使用Qwen大模型进行条款匹配和合规性分析
def analyze_terms_with_qwen(benchmark_term, compare_terms):
    if not qwen_api_key:
        st.error("请先配置Qwen API密钥")
        return []
    
    compare_text = "\n".join([f"条款{i+1}: {term}" for i, term in enumerate(compare_terms)])
    
    prompt = f"""你是一名条款合规性分析专家，请对比以下基准条款与待比较条款列表，找出最匹配的条款。
    基准条款: {benchmark_term}
    
    待比较条款列表:
    {compare_text}
    
    分析要求:
    1. 从待比较条款中找出与基准条款内容最相似的条款
    2. 评估匹配度（0-100分）
    3. 简要说明匹配点和差异点
    4. 如果匹配度低于70分，判定为不合规
    
    请以JSON格式返回结果，包含:
    - best_match_index: 最匹配条款的索引（从0开始，没有则为-1）
    - similarity_score: 匹配度分数
    - analysis: 分析说明
    - is_compliant: 是否合规（true/false）
    """
    
    try:
        response = Generation.call(
            model="qwen-plus",
            prompt=prompt,
            result_format="json"
        )
        
        if response.status_code == 200:
            result = json.loads(response.output.text)
            return result
        else:
            st.error(f"Qwen API调用失败: {response.message}")
            return {"best_match_index": -1, "similarity_score": 0, "analysis": "分析失败", "is_compliant": False}
    except Exception as e:
        st.error(f"条款分析出错: {str(e)}")
        return {"best_match_index": -1, "similarity_score": 0, "analysis": f"分析出错: {str(e)}", "is_compliant": False}

# 生成Word报告
def generate_word_report(benchmark_name, compare_results):
    doc = docx.Document()
    
    # 标题
    title = doc.add_heading("条款合规性对比报告", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 基本信息
    doc.add_paragraph(f"报告生成时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
    doc.add_paragraph(f"基准文件: {benchmark_name}")
    doc.add_paragraph(f"对比文件数量: {len(compare_results)}")
    doc.add_page_break()
    
    # 目录
    doc.add_heading("目录", 1)
    for i, (file_name, _) in enumerate(compare_results.items(), 1):
        p = doc.add_paragraph(f"{i}. {file_name}", style='List Number')
        p.paragraph_format.left_indent = Pt(20)
    
    doc.add_page_break()
    
    # 各文件分析结果
    for file_name, result in compare_results.items():
        doc.add_heading(file_name, 1)
        
        # 合规性概要
        compliant_count = sum(1 for item in result if item["analysis"]["is_compliant"])
        total_count = len(result)
        doc.add_heading("合规性概要", 2)
        doc.add_paragraph(f"总条款数: {total_count}")
        doc.add_paragraph(f"合规条款数: {compliant_count}")
        doc.add_paragraph(f"不合规条款数: {total_count - compliant_count}")
        doc.add_paragraph(f"合规率: {compliant_count/total_count*100:.2f}%")
        
        # 详细匹配结果
        doc.add_heading("条款匹配详情", 2)
        for i, item in enumerate(result, 1):
            doc.add_heading(f"基准条款 {i}: {item['benchmark_term'][:50]}...", 3)
            
            p = doc.add_paragraph("匹配结果: ")
            p.add_run(f"{'合规' if item['analysis']['is_compliant'] else '不合规'} ").bold = True
            p.add_run(f"(匹配度: {item['analysis']['similarity_score']}分)")
            
            if item['analysis']['best_match_index'] != -1:
                match_term = item['compare_terms'][item['analysis']['best_match_index']]
                doc.add_paragraph(f"最匹配条款: {match_term[:100]}...")
            
            doc.add_paragraph("分析说明:")
            doc.add_paragraph(item['analysis']['analysis'], style='List Bullet')
        
        doc.add_page_break()
    
    # 保存到 BytesIO
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

# 主函数
def main():
    # 文件上传
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("基准文件")
        benchmark_file = st.file_uploader("上传基准文件 (PDF或DOCX)", type=["pdf", "docx"], key="benchmark")
    
    with col2:
        st.subheader("对比文件")
        compare_files = st.file_uploader(
            "上传对比文件 (PDF或DOCX，可多选)", 
            type=["pdf", "docx"], 
            key="compare",
            accept_multiple_files=True
        )
    
    # 分析按钮
    if st.button("开始分析", disabled=not (benchmark_file and compare_files)):
        with st.spinner("正在处理文件和分析条款..."):
            # 处理基准文件
            bench_type = benchmark_file.name.split(".")[-1].lower()
            st.info(f"正在处理基准文件: {benchmark_file.name}")
            bench_text = extract_text_from_file(benchmark_file, bench_type)
            
            # 拆分基准条款
            st.info("正在拆分基准条款...")
            bench_terms = split_chinese_terms(bench_text)
            st.success(f"成功拆分出 {len(bench_terms)} 条基准条款")
            
            # 处理对比文件
            compare_results = {}
            
            for compare_file in compare_files:
                file_name = compare_file.name
                st.info(f"正在处理对比文件: {file_name}")
                
                # 提取文本
                compare_type = file_name.split(".")[-1].lower()
                compare_text = extract_text_from_file(compare_file, compare_type)
                
                # 拆分条款
                compare_terms = split_chinese_terms(compare_text)
                st.success(f"成功拆分出 {len(compare_terms)} 条对比条款")
                
                # 条款匹配分析
                file_results = []
                progress_bar = st.progress(0)
                
                for i, bench_term in enumerate(bench_terms):
                    analysis = analyze_terms_with_qwen(bench_term, compare_terms)
                    file_results.append({
                        "benchmark_term": bench_term,
                        "compare_terms": compare_terms,
                        "analysis": analysis
                    })
                    progress_bar.progress((i + 1) / len(bench_terms))
                
                compare_results[file_name] = file_results
                progress_bar.empty()
            
            # 显示结果
            st.success("分析完成！")
            
            # 创建结果标签页
            tabs = st.tabs(["汇总报告"] + list(compare_results.keys()))
            
            # 汇总报告
            with tabs[0]:
                st.subheader("合规性汇总")
                for i, (file_name, result) in enumerate(compare_results.items(), 1):
                    compliant_count = sum(1 for item in result if item["analysis"]["is_compliant"])
                    total_count = len(result)
                    st.write(f"**{file_name}**")
                    st.write(f"合规率: {compliant_count/total_count*100:.2f}% ({compliant_count}/{total_count})")
                    
                    non_compliant = [item for item in result if not item["analysis"]["is_compliant"]]
                    if non_compliant:
                        with st.expander(f"查看不合规条款 ({len(non_compliant)})"):
                            for item in non_compliant:
                                st.write(f"**基准条款:** {item['benchmark_term'][:100]}...")
                                st.write(f"**分析:** {item['analysis']['analysis']}")
                                st.write("---")
            
            # 各文件详细结果
            for i, (file_name, result) in enumerate(compare_results.items(), 1):
                with tabs[i]:
                    st.subheader(f"{file_name} 分析结果")
                    
                    # 合规性概览
                    compliant_count = sum(1 for item in result if item["analysis"]["is_compliant"])
                    total_count = len(result)
                    st.metric("合规率", f"{compliant_count/total_count*100:.2f}%", f"{compliant_count}/{total_count}")
                    
                    # 详细条款对比
                    for j, item in enumerate(result):
                        with st.expander(f"基准条款 {j+1}: {item['benchmark_term'][:80]}..."):
                            col_a, col_b = st.columns(2)
                            
                            with col_a:
                                st.write("**基准条款全文:**")
                                st.write(item['benchmark_term'])
                            
                            with col_b:
                                st.write("**分析结果:**")
                                st.write(f"匹配度: {item['analysis']['similarity_score']}分")
                                st.write(f"合规性: {'✅ 合规' if item['analysis']['is_compliant'] else '❌ 不合规'}")
                                
                                if item['analysis']['best_match_index'] != -1:
                                    match_term = item['compare_terms'][item['analysis']['best_match_index']]
                                    st.write("**最匹配条款:**")
                                    st.write(match_term)
                                
                                st.write("**分析说明:**")
                                st.write(item['analysis']['analysis'])
            
            # 生成并提供下载报告
            st.subheader("生成报告")
            report_buffer = generate_word_report(benchmark_file.name, compare_results)
            st.download_button(
                label="下载合规性报告 (Word)",
                data=report_buffer,
                file_name=f"条款合规性对比报告_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if __name__ == "__main__":
    main()
    
