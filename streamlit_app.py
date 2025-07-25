import streamlit as st
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os
import tempfile
from datetime import datetime
from dashscope import Generation
import json
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import numpy as np
import io

# 页面配置
st.set_page_config(
    page_title="中文条款合规性对比工具",
    page_icon="📄",
    layout="wide"
)

# 页面标题
st.title("📄 中文条款合规性对比工具")
st.write("优化中文条款解析，支持PDF（含图片PDF）和Word文件，精确拆分条款并进行合规性分析")

# 侧边栏配置
with st.sidebar:
    st.subheader("配置")
    
    # Qwen API配置
    st.text("Qwen大模型配置")
    qwen_api_key = st.text_input("阿里云DashScope API密钥", type="password")
    if qwen_api_key:
        os.environ["DASHSCOPE_API_KEY"] = qwen_api_key
    
    # OCR配置
    st.text("\nOCR配置")
    tesseract_path = st.text_input(
        "Tesseract OCR路径", 
        value=r"C:\Program Files\Tesseract-OCR\tesseract.exe" if os.name == 'nt' else "/usr/bin/tesseract"
    )
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
    
    st.info("提示：处理扫描件PDF需要安装Tesseract OCR及中文语言包")

# ------------------------------
# 中文文本处理优化函数
# ------------------------------

def is_chinese_char(c):
    """判断是否为中文字符"""
    return '\u4e00' <= c <= '\u9fff'

def clean_chinese_text(text):
    """清理中文文本，去除多余空行和空格"""
    # 处理中文标点符号前后的空格
    text = re.sub(r'(\s+)([，。；：！？,.;:!?])', r'\2', text)
    text = re.sub(r'([，。；：！？,.;:!?])(\s+)', r'\1', text)
    
    # 合并过多空行
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    return '\n'.join(lines)

# ------------------------------
# PDF解析优化（针对中文）
# ------------------------------

def extract_text_from_pdf(pdf_path):
    """从PDF提取文本，优化中文处理"""
    doc = fitz.open(pdf_path)
    full_text = []
    
    for page_num, page in enumerate(doc):
        # 尝试直接提取文本
        page_text = page.get_text("text")
        
        # 检查页面是否有足够的文本，判断是否为扫描页
        chinese_chars = sum(1 for c in page_text if is_chinese_char(c))
        if len(page_text.strip()) < 50 and chinese_chars < 10:
            # 扫描页，使用OCR
            st.warning(f"第{page_num+1}页可能为图片，将使用OCR提取文本")
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # 预处理图像以提高OCR精度（针对中文）
            img_np = np.array(img)
            # 转换为灰度图
            gray = np.mean(img_np, axis=2).astype(np.uint8)
            # 简单二值化处理
            threshold = 150
            binary = (gray > threshold) * 255
            img_processed = Image.fromarray(binary.astype(np.uint8))
            
            # 识别中文文本
            ocr_text = pytesseract.image_to_string(
                img_processed, 
                lang="chi_sim+eng",  # 中英文混合识别
                config='--psm 6'  # 假设为单一均匀文本块
            )
            full_text.append(ocr_text)
        else:
            # 正常文本页，清理后添加
            cleaned_text = clean_chinese_text(page_text)
            full_text.append(cleaned_text)
    
    return '\n'.join(full_text)

# ------------------------------
# 条款拆分优化（针对中文条款特点）
# ------------------------------

def split_chinese_terms(text):
    """优化中文条款拆分，处理各种常见的条款编号格式"""
    # 中文条款常见编号格式正则表达式
    # 匹配：1. 、(1)、一、1.1 、1.1.1、第一条、第一款等格式
    patterns = [
        r'^(第[一二三四五六七八九十百]+[条款项点])(\.?\s?)',  # 第一条、第一款
        r'^(\d+)\.\s',  # 1. 
        r'^(\d+\.\d+)\.\s',  # 1.1.
        r'^(\(\d+\))\s',  # (1)
        r'^([一二三四五六七八九十]+)\.\s',  # 一. 
        r'^(\d+\.\d+\.\d+)\s'  # 1.1.1 
    ]
    
    terms = []
    current_term = []
    current_number = None
    
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
            
        # 检查是否为新条款开头
        matched = False
        for pattern in patterns:
            match = re.match(pattern, line)
            if match:
                # 如果有当前条款，先保存
                if current_term:
                    terms.append({
                        'number': current_number,
                        'content': ' '.join(current_term).strip()
                    })
                
                # 开始新条款
                current_number = match.group(1)
                current_term = [line[len(match.group(0)):].strip()]
                matched = True
                break
        
        if not matched and current_term:
            # 不是新条款，添加到当前条款
            current_term.append(line)
    
    # 添加最后一个条款
    if current_term:
        terms.append({
            'number': current_number,
            'content': ' '.join(current_term).strip()
        })
    
    # 如果没有匹配到任何条款格式，使用Qwen大模型进行智能拆分
    if len(terms) <= 1:
        st.info("检测到条款格式复杂，将使用Qwen大模型进行智能拆分")
        return split_terms_with_qwen(text)
    
    return terms

def split_terms_with_qwen(text):
    """使用Qwen大模型智能拆分中文条款"""
    if not qwen_api_key:
        st.error("请先配置Qwen API密钥以处理复杂条款")
        return []
    
    prompt = f"""请帮我从以下中文文本中提取条款，每条条款作为一个独立项。
    条款通常以以下形式开头：
    - 数字编号：1. 、1.1 、(1) 等
    - 中文编号：第一条、第一款、一、一等
    
    请以JSON数组格式返回，每个元素是包含"number"（条款编号）和"content"（条款内容）的对象。
    确保准确拆分，保持条款的完整性和独立性。
    
    文本内容：
    {text[:3000]}
    """
    
    try:
        response = Generation.call(
            model="qwen-plus",
            prompt=prompt,
            result_format="json"
        )
        
        if response.status_code == 200:
            try:
                terms = json.loads(response.output.text)
                return terms if isinstance(terms, list) else []
            except json.JSONDecodeError:
                st.warning("Qwen返回结果格式不正确，将使用备用拆分方法")
                return split_chinese_terms(text)
        else:
            st.error(f"Qwen API调用失败: {response.message}")
            return split_chinese_terms(text)
    except Exception as e:
        st.error(f"条款拆分出错: {str(e)}")
        return split_chinese_terms(text)

# ------------------------------
# 通用文件处理函数
# ------------------------------

def extract_text_from_docx(file_path):
    """从Word文档提取文本"""
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text)
    return clean_chinese_text('\n'.join(full_text))

def extract_text_from_file(file, file_type):
    """根据文件类型提取文本"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_type}') as f:
        f.write(file.getbuffer())
        temp_path = f.name
    
    try:
        if file_type == 'pdf':
            text = extract_text_from_pdf(temp_path)
        elif file_type in ['docx', 'doc']:  # 简单支持doc格式
            text = extract_text_from_docx(temp_path)
        else:
            text = ""
            st.error(f"不支持的文件类型: {file_type}")
    finally:
        os.unlink(temp_path)
    
    return text

# ------------------------------
# 条款匹配与分析
# ------------------------------

def analyze_compliance(benchmark_terms, compare_terms):
    """分析条款合规性"""
    if not qwen_api_key:
        st.error("请配置Qwen API密钥以进行合规性分析")
        return []
    
    results = []
    
    # 显示进度条
    progress_bar = st.progress(0)
    total = len(benchmark_terms)
    
    for i, bench_term in enumerate(benchmark_terms):
        progress_bar.progress((i + 1) / total)
        
        prompt = f"""请分析以下对比条款与基准条款的合规性：
        
        基准条款[{bench_term.get('number', '')}]：
        {bench_term.get('content', '')}
        
        对比条款列表：
        {json.dumps(compare_terms, ensure_ascii=False, indent=2)}
        
        请找出最匹配的对比条款，并分析：
        1. 是否合规（相似度是否达到80%以上）
        2. 主要差异点（如果不合规）
        3. 匹配的条款编号
        
        请以JSON格式返回，包含：
        - benchmark_number: 基准条款编号
        - matched_number: 匹配的对比条款编号（如无则为null）
        - is_compliant: 是否合规（true/false）
        - similarity: 相似度（0-100）
        - differences: 差异描述（如不合规）
        """
        
        try:
            response = Generation.call(
                model="qwen-plus",
                prompt=prompt,
                result_format="json"
            )
            
            if response.status_code == 200:
                analysis = json.loads(response.output.text)
                results.append({
                    'benchmark': bench_term,
                    'analysis': analysis
                })
            else:
                st.warning(f"分析条款[{bench_term.get('number')}]失败: {response.message}")
        except Exception as e:
            st.warning(f"分析条款[{bench_term.get('number')}]出错: {str(e)}")
    
    progress_bar.empty()
    return results

# ------------------------------
# 报告生成
# ------------------------------

def generate_report(benchmark_name, compare_files_results):
    """生成Word报告"""
    doc = docx.Document()
    
    # 标题
    title = doc.add_heading("条款合规性对比报告", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 基本信息
    doc.add_paragraph(f"基准文件: {benchmark_name}")
    doc.add_paragraph(f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph("")
    
    # 目录
    doc.add_heading("目录", 1)
    for i, (file_name, results) in enumerate(compare_files_results.items(), 1):
        doc.add_paragraph(f"{i}. {file_name}", style='List Number')
    doc.add_paragraph("")
    
    # 每个对比文件的分析结果
    for file_name, results in compare_files_results.items():
        doc.add_heading(f"文件: {file_name}", 1)
        
        # 合规条款
        doc.add_heading("1. 合规条款", 2)
        compliant_terms = [r for r in results if r['analysis'].get('is_compliant', False)]
        
        if compliant_terms:
            for term in compliant_terms:
                p = doc.add_paragraph()
                p.add_run(f"基准条款[{term['benchmark']['number']}]: ").bold = True
                p.add_run(term['benchmark']['content'])
                
                p = doc.add_paragraph()
                p.add_run(f"匹配条款[{term['analysis']['matched_number']}]: ").bold = True
                p.add_run(f"(相似度: {term['analysis']['similarity']}%)")
                doc.add_paragraph("")
        else:
            doc.add_paragraph("无合规条款")
        
        # 不合规条款
        doc.add_heading("2. 不合规条款", 2)
        non_compliant_terms = [r for r in results if not r['analysis'].get('is_compliant', False)]
        
        if non_compliant_terms:
            for term in non_compliant_terms:
                p = doc.add_paragraph()
                p.add_run(f"基准条款[{term['benchmark']['number']}]: ").bold = True
                p.add_run(term['benchmark']['content'])
                
                p = doc.add_paragraph()
                p.add_run("差异分析: ").bold = True
                p.add_run(term['analysis'].get('differences', '无详细分析'))
                doc.add_paragraph("")
        else:
            doc.add_paragraph("无不合规条款")
        
        doc.add_page_break()
    
    # 保存到内存
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

# ------------------------------
# 主程序
# ------------------------------

def main():
    # 文件上传
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("基准文件")
        benchmark_file = st.file_uploader("上传基准条款文件", type=['pdf', 'docx', 'doc'], key='benchmark')
    
    with col2:
        st.subheader("对比文件")
        compare_files = st.file_uploader(
            "上传需要对比的文件（可多选）", 
            type=['pdf', 'docx', 'doc'], 
            key='compare',
            accept_multiple_files=True
        )
    
    # 分析按钮
    if st.button("开始分析", disabled=not (benchmark_file and compare_files)):
        with st.spinner("正在处理文件..."):
            # 处理基准文件
            st.info(f"正在解析基准文件: {benchmark_file.name}")
            bench_type = benchmark_file.name.split('.')[-1].lower()
            bench_text = extract_text_from_file(benchmark_file, bench_type)
            
            # 拆分基准条款
            st.info("正在拆分基准条款...")
            benchmark_terms = split_chinese_terms(bench_text)
            st.success(f"成功拆分基准条款: {len(benchmark_terms)}条")
            
            # 处理对比文件
            compare_files_results = {}
            
            for compare_file in compare_files:
                st.info(f"正在处理对比文件: {compare_file.name}")
                file_type = compare_file.name.split('.')[-1].lower()
                file_text = extract_text_from_file(compare_file, file_type)
                
                # 拆分对比条款
                st.info(f"正在拆分{compare_file.name}的条款...")
                compare_terms = split_chinese_terms(file_text)
                st.success(f"成功拆分{compare_file.name}的条款: {len(compare_terms)}条")
                
                # 分析合规性
                st.info(f"正在分析{compare_file.name}的合规性...")
                results = analyze_compliance(benchmark_terms, compare_terms)
                compare_files_results[compare_file.name] = results
            
            # 展示结果
            st.success("分析完成！")
            
            # 使用标签页展示每个文件的结果
            tabs = st.tabs(list(compare_files_results.keys()))
            
            for tab, (file_name, results) in zip(tabs, compare_files_results.items()):
                with tab:
                    # 合规条款
                    st.subheader("✅ 可匹配的条款")
                    compliant = [r for r in results if r['analysis'].get('is_compliant', False)]
                    
                    if compliant:
                        for item in compliant:
                            with st.expander(f"基准条款[{item['benchmark']['number']}] 与 对比条款[{item['analysis']['matched_number']}] (相似度: {item['analysis']['similarity']}%)"):
                                col_bench, col_compare = st.columns(2)
                                with col_bench:
                                    st.write("**基准条款内容：**")
                                    st.write(item['benchmark']['content'])
                                with col_compare:
                                    st.write("**对比条款内容：**")
                                    # 这里需要根据matched_number找到对应的条款内容
                                    st.write("(对比条款内容将在此显示)")
                    else:
                        st.info("没有找到可匹配的条款")
                    
                    # 不合规条款
                    st.subheader("❌ 不合规条款总结")
                    non_compliant = [r for r in results if not r['analysis'].get('is_compliant', False)]
                    
                    if non_compliant:
                        for item in non_compliant:
                            with st.expander(f"基准条款[{item['benchmark']['number']}]"):
                                st.write("**基准条款内容：**")
                                st.write(item['benchmark']['content'])
                                st.write("**差异分析：**")
                                st.write(item['analysis'].get('differences', '无详细分析'))
                    else:
                        st.info("所有条款均合规")
            
            # 生成报告
            st.subheader("📥 生成报告")
            report_buffer = generate_report(benchmark_file.name, compare_files_results)
            st.download_button(
                label="下载Word报告",
                data=report_buffer,
                file_name=f"条款合规性对比报告_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if __name__ == "__main__":
    main()
    
