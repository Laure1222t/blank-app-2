import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os
import tempfile
from datetime import datetime
from dashscope import Generation  # 阿里云Qwen大模型SDK
import json
import fitz  # PyMuPDF用于PDF文本提取
import pytesseract  # OCR识别库
from PIL import Image
import io

# 设置页面配置
st.set_page_config(
    page_title="条款合规性对比工具",
    page_icon="📄",
    layout="wide"
)

# 页面标题
st.title("📄 条款合规性对比工具 (Qwen增强版)")
st.write("上传基准文件和待比较文件（支持Word和PDF，包括图片PDF），系统将使用Qwen大模型进行智能条款匹配分析并生成合规性报告。")

# Qwen API密钥配置
with st.sidebar:
    st.subheader("Qwen大模型配置")
    qwen_api_key = st.text_input("请输入阿里云DashScope API密钥", type="password")
    if qwen_api_key:
        os.environ["DASHSCOPE_API_KEY"] = qwen_api_key
    st.info("需要阿里云账号和DashScope服务访问权限，获取API密钥: https://dashscope.console.aliyun.com/")
    
    st.subheader("OCR配置")
    st.warning("处理图片PDF需要Tesseract OCR支持，请确保已安装并配置好环境")
    tess_path = st.text_input("Tesseract OCR路径", value=r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if tess_path:
        pytesseract.pytesseract.tesseract_cmd = tess_path

# 辅助函数：从文件中提取文本（支持Word和PDF）
def extract_text_from_file(file_path, file_type):
    if file_type == "docx":
        return extract_text_from_docx(file_path)
    elif file_type == "pdf":
        return extract_text_from_pdf(file_path)
    else:
        st.error(f"不支持的文件类型: {file_type}")
        return ""

# 从Word文档提取文本
def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():  # 只添加非空段落
            full_text.append(para.text)
    return '\n'.join(full_text)

# 从PDF提取文本（支持图片PDF的OCR识别）
def extract_text_from_pdf(file_path):
    pdf_document = fitz.open(file_path)
    full_text = []
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text()
        
        # 如果页面文本为空，尝试OCR识别图片内容
        if not text.strip():
            st.info(f"PDF页面 {page_num + 1} 看起来是图片，正在使用OCR提取文本...")
            text = ocr_pdf_page(page)
        
        if text.strip():
            full_text.append(text)
    
    pdf_document.close()
    return '\n'.join(full_text)

# 对PDF页面进行OCR识别
def ocr_pdf_page(page):
    # 将PDF页面转换为图片
    pix = page.get_pixmap(dpi=300)  # 高DPI提高识别精度
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    
    # 使用Tesseract进行OCR识别
    try:
        text = pytesseract.image_to_string(img, lang="chi_sim+eng")
        return text
    except Exception as e:
        st.error(f"OCR识别失败: {str(e)}")
        st.warning("请确保已正确安装Tesseract OCR并配置了正确路径")
        return ""

# 辅助函数：使用Qwen大模型拆分条款
def split_terms_with_qwen(text):
    if not qwen_api_key:
        st.error("请先配置Qwen API密钥")
        return []
    
    prompt = f"""请帮我从以下文本中提取条款，每条条款作为一个独立项。条款通常以数字编号开头，如"1. "、"2.1 "等。
    请以JSON数组格式返回，每个元素是一个条款的完整内容（包含编号）。如果没有明确的条款结构，按逻辑段落拆分。
    
    文本内容：
    {text[:2000]}  # 限制输入长度，避免超出API限制
    """
    
    try:
        response = Generation.call(
            model="qwen-plus",
            prompt=prompt,
            result_format="json"
        )
        
        if response.status_code == 200:
            terms = json.loads(response.output.text)
            return terms if isinstance(terms, list) else []
        else:
            st.error(f"Qwen API调用失败: {response.message}")
            return []
    except Exception as e:
        st.error(f"条款拆分出错: {str(e)}")
        # 备用方案：使用正则表达式拆分
        return split_terms_with_regex(text)

# 备用函数：使用正则表达式拆分条款
def split_terms_with_regex(text):
    # 使用正则表达式匹配条款编号，如1. 1.1 2. 等
    pattern = r'(\d+\.\s|\d+\.\d+\s|\d+\s)'
    terms = re.split(pattern, text)
    
    # 重组条款，将编号和内容合并
    result = []
    for i in range(1, len(terms), 2):
        if i + 1 < len(terms):
            term_number = terms[i].strip()
            term_content = terms[i+1].strip()
            result.append(f"{term_number} {term_content}")
    
    # 如果没有匹配到条款编号格式，将整个文本作为一个条款
    if not result:
        result.append(text)
    
    return result

# 辅助函数：使用Qwen大模型进行条款匹配和合规性分析
def analyze_terms_with_qwen(benchmark_term, compare_terms):
    if not qwen_api_key:
        st.error("请先配置Qwen API密钥")
        return None
    
    compare_text = "\n".join([f"[{i+1}] {term}" for i, term in enumerate(compare_terms)])
    
    prompt = f"""作为法律条款合规性专家，请分析以下基准条款与待比较条款的匹配度和合规性：
    
    基准条款：
    {benchmark_term}
    
    待比较条款列表：
    {compare_text}
    
    请找出最匹配的条款，并分析其合规性。如果存在合规问题，请指出具体差异和不合规之处。
    请以JSON格式返回，包含以下字段：
    - best_match_index: 最匹配条款的索引（从0开始），如果没有匹配项则为-1
    - similarity_score: 匹配度评分（0-100）
    - compliance_analysis: 合规性分析说明
    - is_compliant: 是否合规（true/false）
    """
    
    try:
        response = Generation.call(
            model="qwen-plus",
            prompt=prompt,
            result_format="json"
        )
        
        if response.status_code == 200:
            return json.loads(response.output.text)
        else:
            st.error(f"Qwen API调用失败: {response.message}")
            return None
    except Exception as e:
        st.error(f"条款分析出错: {str(e)}")
        return None

# 生成合规性报告Word文档
def generate_report(benchmark_terms, compare_terms, analysis_results):
    doc = docx.Document()
    
    # 设置标题
    title = doc.add_heading("条款合规性对比报告", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 添加报告信息
    doc.add_paragraph(f"生成日期: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph(f"基准条款数量: {len(benchmark_terms)}")
    doc.add_paragraph(f"待比较条款数量: {len(compare_terms)}")
    doc.add_paragraph("---")
    
    # 添加匹配条款分析
    doc.add_heading("1. 可匹配条款分析", level=1)
    matched_count = 0
    
    for i, (benchmark_term, analysis) in enumerate(zip(benchmark_terms, analysis_results)):
        if analysis and analysis.get('best_match_index', -1) != -1:
            matched_count += 1
            match_idx = analysis['best_match_index']
            match_term = compare_terms[match_idx]
            
            doc.add_heading(f"1.{matched_count} 基准条款 {i+1}", level=2)
            p = doc.add_paragraph(benchmark_term)
            p.paragraph_format.space_after = Pt(12)
            
            doc.add_heading(f"1.{matched_count}.1 匹配条款 {match_idx+1}", level=3)
            p = doc.add_paragraph(match_term)
            p.paragraph_format.space_after = Pt(12)
            
            doc.add_heading(f"1.{matched_count}.2 合规性分析", level=3)
            p = doc.add_paragraph(analysis['compliance_analysis'])
            p.paragraph_format.space_after = Pt(12)
            
            # 添加合规性标识
            compliant_text = "合规" if analysis['is_compliant'] else "不合规"
            compliant_color = "green" if analysis['is_compliant'] else "red"
            p = doc.add_paragraph(f"合规性: {compliant_text} (匹配度: {analysis['similarity_score']}/100)")
            p.font.color.rgb = docx.shared.RGBColor.from_string(compliant_color)
            p.paragraph_format.space_after = Pt(24)
    
    # 添加不合规条款总结
    doc.add_heading("2. 不合规条款总结", level=1)
    
    non_compliant = [
        (i, term, analysis) 
        for i, (term, analysis) in enumerate(zip(benchmark_terms, analysis_results))
        if analysis and not analysis.get('is_compliant', False)
    ]
    
    if not non_compliant:
        doc.add_paragraph("未发现不合规条款")
    else:
        for i, (term_idx, term, analysis) in enumerate(non_compliant):
            doc.add_heading(f"2.{i+1} 基准条款 {term_idx+1}", level=2)
            p = doc.add_paragraph(term)
            p.paragraph_format.space_after = Pt(12)
            
            p = doc.add_paragraph(f"不合规原因: {analysis['compliance_analysis']}")
            p.paragraph_format.space_after = Pt(18)
    
    # 保存到临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        doc.save(tmp.name)
        return tmp.name

# 主程序
def main():
    # 文件上传
    col1, col2 = st.columns(2)
    
    with col1:
        benchmark_file = st.file_uploader("上传基准文件 (Word或PDF)", type=["docx", "pdf"])
    
    with col2:
        compare_file = st.file_uploader("上传待比较文件 (Word或PDF)", type=["docx", "pdf"])
    
    if benchmark_file and compare_file and st.button("开始分析"):
        with st.spinner("正在处理文件..."):
            # 保存上传的文件到临时目录
            with tempfile.TemporaryDirectory() as tmpdir:
                # 保存基准文件
                benchmark_path = os.path.join(tmpdir, benchmark_file.name)
                with open(benchmark_path, "wb") as f:
                    f.write(benchmark_file.getvalue())
                
                # 保存待比较文件
                compare_path = os.path.join(tmpdir, compare_file.name)
                with open(compare_path, "wb") as f:
                    f.write(compare_file.getvalue())
                
                # 提取文本
                benchmark_type = benchmark_file.name.split('.')[-1].lower()
                compare_type = compare_file.name.split('.')[-1].lower()
                
                benchmark_text = extract_text_from_file(benchmark_path, benchmark_type)
                compare_text = extract_text_from_file(compare_path, compare_type)
                
                # 拆分条款
                st.subheader("条款提取结果")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.info("正在使用Qwen大模型拆分基准条款...")
                    benchmark_terms = split_terms_with_qwen(benchmark_text)
                    st.success(f"成功提取基准条款 {len(benchmark_terms)} 条")
                    with st.expander("查看基准条款"):
                        for i, term in enumerate(benchmark_terms):
                            st.write(f"{i+1}. {term}")
                
                with col2:
                    st.info("正在使用Qwen大模型拆分待比较条款...")
                    compare_terms = split_terms_with_qwen(compare_text)
                    st.success(f"成功提取待比较条款 {len(compare_terms)} 条")
                    with st.expander("查看待比较条款"):
                        for i, term in enumerate(compare_terms):
                            st.write(f"{i+1}. {term}")
                
                # 条款匹配与分析
                if benchmark_terms and compare_terms:
                    st.subheader("条款匹配与合规性分析")
                    analysis_results = []
                    
                    progress_bar = st.progress(0)
                    for i, benchmark_term in enumerate(benchmark_terms):
                        st.info(f"正在分析基准条款 {i+1}/{len(benchmark_terms)}...")
                        analysis = analyze_terms_with_qwen(benchmark_term, compare_terms)
                        analysis_results.append(analysis)
                        progress_bar.progress((i+1)/len(benchmark_terms))
                    
                    # 显示分析结果
                    st.subheader("分析结果摘要")
                    compliant_count = sum(1 for res in analysis_results if res and res.get('is_compliant', False))
                    st.metric("合规条款数量", compliant_count)
                    st.metric("不合规条款数量", len(benchmark_terms) - compliant_count)
                    
                    # 生成并提供下载报告
                    st.subheader("生成合规性报告")
                    with st.spinner("正在生成Word报告..."):
                        report_path = generate_report(benchmark_terms, compare_terms, analysis_results)
                        
                        with open(report_path, "rb") as f:
                            st.download_button(
                                label="下载合规性报告",
                                data=f,
                                file_name=f"条款合规性对比报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

if __name__ == "__main__":
    main()
    
