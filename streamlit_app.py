import streamlit as st
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import re
import os
import tempfile
from datetime import datetime
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import numpy as np
import requests
import json
from io import BytesIO

# 页面配置
st.set_page_config(
    page_title="条款合规性对比工具",
    page_icon="📄",
    layout="wide"
)

# 初始化会话状态
if 'analysis_results' not in st.session_state:
    st.session_state.analysis_results = {}
if 'bench_terms' not in st.session_state:
    st.session_state.bench_terms = []
if 'comparison_terms' not in st.session_state:
    st.session_state.comparison_terms = {}
if 'debug_mode' not in st.session_state:
    st.session_state.debug_mode = False


### 1. 工具函数：文件解析与文本提取
def check_tesseract_installation():
    """检查Tesseract是否安装"""
    try:
        pytesseract.get_tesseract_version()
        return True
    except:
        return False

def has_selectable_text(page):
    """判断PDF页面是否为可选择文本（非图片）"""
    text = page.get_text("text").strip()
    return len(text) > 50

def ocr_image(image):
    """对图片进行OCR识别（中文优先）"""
    try:
        # 图像预处理：转为灰度图并二值化
        gray_image = image.convert('L')
        threshold = 150
        binary_image = gray_image.point(lambda p: p > threshold and 255)
        
        # 执行OCR（中英文混合）
        text = pytesseract.image_to_string(
            binary_image,
            lang='chi_sim+eng',
            config='--psm 6'  # 假设单一均匀文本块
        )
        return text.strip()
    except Exception as e:
        st.warning(f"OCR识别出错: {str(e)}")
        return ""

def extract_text_from_pdf(pdf_path):
    """从PDF提取文本（优先文本提取，必要时OCR）"""
    text = []
    try:
        doc = fitz.open(pdf_path)
        tesseract_available = check_tesseract_installation()
        
        for page_num, page in enumerate(doc):
            if has_selectable_text(page):
                page_text = page.get_text("text").strip()
                text.append(f"{page_text}")
            else:
                if tesseract_available:
                    pix = page.get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_result = ocr_image(img)
                    text.append(f"{ocr_result}")
                else:
                    text.append(f"[无法提取文本：未安装Tesseract]")
        
        doc.close()
        return "\n\n".join(text)
    except Exception as e:
        st.error(f"PDF解析失败: {str(e)}")
        return ""

def extract_text_from_docx(docx_path):
    """从DOCX提取文本"""
    try:
        doc = docx.Document(docx_path)
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())
        return "\n".join(full_text)
    except Exception as e:
        st.error(f"DOCX解析失败: {str(e)}")
        return ""

def extract_text_from_file(uploaded_file, file_type):
    """统一文件提取入口"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_type}") as temp_file:
            temp_file.write(uploaded_file.getvalue())
            temp_path = temp_file.name
        
        if file_type == "pdf":
            return extract_text_from_pdf(temp_path)
        elif file_type == "docx":
            return extract_text_from_docx(temp_path)
        else:
            return ""
    finally:
        if 'temp_path' in locals() and os.path.exists(temp_path):
            os.unlink(temp_path)


### 2. 增强版中文条款拆分函数
def split_chinese_terms(text):
    """
    增强版中文条款拆分，针对中文法律/合规文件特点优化
    支持多种条款编号格式，提高拆分准确性
    """
    # 输入验证
    if not text or not isinstance(text, str) or len(text.strip()) < 10:
        st.warning("输入文本为空或太短，无法拆分条款")
        return []
    
    # 调试模式：显示原始文本
    if st.session_state.debug_mode:
        with st.expander("查看原始文本（用于调试）"):
            st.text_area("", text, height=200)
    
    # 文本预处理
    processed_text = text
    # 替换全角标点为半角，便于统一处理
    processed_text = processed_text.replace('。', '.').replace('，', ',').replace('；', ';')
    # 去除多余空行
    processed_text = re.sub(r'\n+', '\n', processed_text.strip())
    # 处理可能的连在一起的编号（如"1.条款内容2.条款内容"）
    processed_text = re.sub(r'(\d+)\.([^\d])', r'\1.\n\2', processed_text)
    processed_text = re.sub(r'(\d+)\.(\d+)\.', r'\1.\2.\n', processed_text)
    
    # 中文条款常见编号格式（增强版正则模式）
    # 按优先级排序，更具体的模式排在前面
    patterns = [
        r'((?:第)?[一二三四五六七八九十百]+(?:条|款|项|点|节)\s*)',  # 第一条、第一款、第一项
        r'((?:第)?\d+(?:条|款|项|点|节)\s*)',                          # 第1条、第1款、第1项
        r'(\d+\.\d+\.\d+\s*)',                                          # 1.1.1 
        r'(\d+\.\d+\s*)',                                                # 1.1 
        r'(\d+\s*)',                                                     # 1 
        r'([一二三四五六七八九十]+\、\s*)',                             # 一、二、
        r'(\(\d+\)\s*)',                                                 # (1) (2)
        r'(\(\D+\)\s*)',                                                 # (一) (二)
        r'([A-Za-z]\.\s*)',                                              # A. B.
        r'(\d+\)\s*)'                                                    # 1) 2)
    ]
    
    # 组合所有模式，使用正向前瞻确保只匹配作为开头的编号
    combined_pattern = r'(?m)^(' + '|'.join(patterns).replace('(', '(?:') + ')'
    
    # 拆分文本并重组条款
    terms = []
    current_term = ""
    
    # 使用finditer找到所有匹配的编号位置
    matches = list(re.finditer(combined_pattern, processed_text, re.MULTILINE))
    
    if not matches:
        # 如果没有找到任何编号模式，尝试按空行拆分
        st.info("未检测到标准条款编号格式，尝试按空行拆分")
        raw_terms = re.split(r'\n\s*\n', processed_text)
        return [term.strip() for term in raw_terms if term.strip()]
    
    # 处理第一个条款之前的内容（如果有）
    first_match = matches[0]
    if first_match.start() > 0:
        prefix = processed_text[:first_match.start()].strip()
        if prefix:
            terms.append(prefix)
    
    # 处理所有匹配的条款
    for i, match in enumerate(matches):
        # 当前编号
        numbering = match.group(0)
        # 计算当前条款的结束位置
        if i < len(matches) - 1:
            end_pos = matches[i+1].start()
        else:
            end_pos = len(processed_text)
        
        # 提取条款内容
        content = processed_text[match.end():end_pos].strip()
        full_term = f"{numbering}{content}"
        terms.append(full_term)
    
    # 调试模式：显示拆分结果
    if st.session_state.debug_mode:
        with st.expander("查看条款拆分结果（用于调试）"):
            for i, term in enumerate(terms):
                st.write(f"条款{i+1}: {term[:100]}...")
    
    # 过滤过短的条款（可能是误拆分）
    filtered_terms = []
    for term in terms:
        if len(term) > 10:  # 过滤掉长度小于10的条款
            filtered_terms.append(term)
        elif st.session_state.debug_mode:
            st.warning(f"过滤短条款：{term}")
    
    # 拆分效果评估
    if len(filtered_terms) < 3 and len(processed_text) > 1000:
        st.info(f"检测到可能的条款拆分效果不佳（共{len(filtered_terms)}条），建议检查文件格式")
    
    return filtered_terms


### 3. Qwen大模型调用（兼容模式API）
def call_qwen_api(prompt, api_key):
    """调用阿里云DashScope兼容模式API"""
    if not api_key:
        return None, "未提供API密钥"
    
    url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    
    payload = {
        "model": "qwen-plus",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "choices" in result and len(result["choices"]) > 0:
            return result["choices"][0]["message"]["content"], None
        else:
            return None, f"API返回格式异常: {str(result)}"
    except Exception as e:
        return None, f"API调用失败: {str(e)}"

def analyze_terms_with_qwen(bench_term, compare_term, api_key):
    """用Qwen分析条款匹配度"""
    prompt = f"""请对比以下两个条款的合规性：
    【基准条款】：{bench_term[:500]}
    【待比条款】：{compare_term[:500]}
    
    请按以下格式回答：
    1. 匹配度（0-100分）：[分数]
    2. 相同点：[简要说明相同内容]
    3. 差异点：[简要说明差异内容]
    4. 合规性判断：[符合/部分符合/不符合]
    5. 判断依据：[说明判断理由]
    """
    
    result, error = call_qwen_api(prompt, api_key)
    if error:
        return None, error
    
    # 解析结果
    try:
        score_match = re.search(r'匹配度（0-100分）：(\d+)', result)
        score = int(score_match.group(1)) if score_match else 0
        
        compliance_match = re.search(r'合规性判断：([^\n]+)', result)
        compliance = compliance_match.group(1) if compliance_match else "无法判断"
        
        return {
            "score": score,
            "full_analysis": result,
            "compliance": compliance
        }, None
    except:
        return {
            "score": 0,
            "full_analysis": f"解析失败，原始结果：{result}",
            "compliance": "无法判断"
        }, None


### 4. 结果报告生成
def generate_word_report(bench_terms, comparison_results, bench_filename):
    """生成可下载的Word报告"""
    doc = docx.Document()
    
    # 设置中文字体
    style = doc.styles['Normal']
    style.font.name = 'SimSun'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    style.font.size = Pt(10.5)
    
    # 标题
    title = doc.add_heading("条款合规性对比报告", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 基本信息
    doc.add_paragraph(f"报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"基准文件：{bench_filename}")
    doc.add_paragraph(f"对比文件数量：{len(comparison_results)}")
    doc.add_page_break()
    
    # 按文件生成结果
    for file_name, result in comparison_results.items():
        doc.add_heading(f"对比文件：{file_name}", level=1)
        
        # 可匹配条款
        doc.add_heading("1. 可匹配条款", level=2)
        if result["matched_terms"]:
            for idx, item in enumerate(result["matched_terms"], 1):
                doc.add_heading(f"1.{idx} 基准条款：{item['bench_term'][:30]}...", level=3)
                doc.add_paragraph(f"对比条款：{item['compare_term'][:50]}...")
                doc.add_paragraph(f"匹配度：{item['analysis']['score']}分")
                doc.add_paragraph(f"合规性：{item['analysis']['compliance']}")
                doc.add_paragraph("分析详情：")
                doc.add_paragraph(item['analysis']['full_analysis'], style='Normal')
        else:
            doc.add_paragraph("无匹配条款")
        
        # 不合规条款总结
        doc.add_heading("2. 不合规条款总结", level=2)
        if result["non_compliant_terms"]:
            for term in result["non_compliant_terms"]:
                doc.add_paragraph(f"- {term[:100]}...")
        else:
            doc.add_paragraph("未发现不合规条款")
        
        doc.add_page_break()
    
    # 保存到内存
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


### 5. 主函数
def main():
    st.title("📄 条款合规性对比工具（增强版）")
    st.write("支持上传基准文件和多个对比文件（PDF/DOCX），优化了中文条款拆分效果")
    
    # 侧边栏配置
    with st.sidebar:
        st.subheader("配置")
        qwen_api_key = st.text_input("阿里云DashScope API密钥", type="password")
        st.session_state.debug_mode = st.checkbox("启用调试模式", value=False)
        st.info("获取密钥：https://dashscope.console.aliyun.com/")
        st.divider()
        st.subheader("使用说明")
        st.write("1. 上传1个基准文件和多个对比文件")
        st.write("2. 点击开始分析")
        st.write("3. 查看结果并下载报告")
    
    # 文件上传
    col1, col2 = st.columns(2)
    with col1:
        bench_file = st.file_uploader("上传基准文件（PDF/DOCX）", type=["pdf", "docx"], accept_multiple_files=False)
    with col2:
        compare_files = st.file_uploader("上传对比文件（PDF/DOCX）", type=["pdf", "docx"], accept_multiple_files=True)
    
    # 分析按钮
    if st.button("开始分析", disabled=not (bench_file and compare_files)):
        with st.spinner("正在处理文件..."):
            # 1. 提取基准文件文本并拆分条款
            bench_type = bench_file.name.split('.')[-1].lower()
            bench_text = extract_text_from_file(bench_file, bench_type)
            bench_terms = split_chinese_terms(bench_text)
            st.session_state.bench_terms = bench_terms
            st.success(f"基准文件解析完成，提取条款：{len(bench_terms)}条")
        
        # 2. 处理每个对比文件
        all_results = {}
        progress_bar = st.progress(0)
        
        for file_idx, compare_file in enumerate(compare_files):
            file_name = compare_file.name
            st.subheader(f"处理对比文件：{file_name}")
            
            # 提取文本并拆分条款
            compare_type = file_name.split('.')[-1].lower()
            compare_text = extract_text_from_file(compare_file, compare_type)
            compare_terms = split_chinese_terms(compare_text)
            st.session_state.comparison_terms[file_name] = compare_terms
            st.info(f"提取条款：{len(compare_terms)}条")
            
            # 条款对比分析
            matched_terms = []
            non_compliant_terms = []
            
            with st.spinner(f"正在分析 {file_name} 的条款..."):
                min_terms = min(len(bench_terms), len(compare_terms))
                for i in range(min_terms):
                    bench_term = bench_terms[i]
                    compare_term = compare_terms[i]
                    
                    # 调用Qwen分析
                    if qwen_api_key:
                        analysis, error = analyze_terms_with_qwen(bench_term, compare_term, qwen_api_key)
                        if error:
                            st.warning(f"条款{i+1}分析失败：{error}")
                            continue
                    else:
                        # 无API时的基础判断
                        analysis = {
                            "score": 50 if len(set(bench_term[:100]) & set(compare_term[:100])) > 10 else 20,
                            "full_analysis": "未使用Qwen API，无法提供详细分析",
                            "compliance": "未知（需API密钥）"
                        }
                    
                    # 分类
                    if analysis["score"] >= 70:
                        matched_terms.append({
                            "bench_term": bench_term,
                            "compare_term": compare_term,
                            "analysis": analysis
                        })
                    else:
                        non_compliant_terms.append(compare_term)
            
            # 保存结果
            all_results[file_name] = {
                "matched_terms": matched_terms,
                "non_compliant_terms": non_compliant_terms
            }
            
            # 更新进度
            progress_bar.progress((file_idx + 1) / len(compare_files))
        
        # 3. 展示结果
        st.session_state.analysis_results = all_results
        st.success("所有文件分析完成！")
        
        # 4. 生成报告
        if st.button("生成Word报告"):
            with st.spinner("正在生成报告..."):
                report_buffer = generate_word_report(
                    bench_terms, 
                    all_results, 
                    bench_file.name
                )
                st.download_button(
                    label="下载报告",
                    data=report_buffer,
                    file_name=f"条款合规性对比报告_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )


if __name__ == "__main__":
    main()
    
