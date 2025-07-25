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
            config='--psm 6'
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
                    text.append(f"[警告：无法提取文本（未安装Tesseract）]")
        
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


### 2. 中文条款拆分函数（修复正则错误版）
def split_chinese_terms(text):
    """拆分中文条款（修复正则表达式错误，增强稳定性）"""
    # 输入验证
    if not text or not isinstance(text, str):
        st.warning("输入文本为空或无效，无法拆分条款")
        return []
    
    # 预处理：清除多余空行和空格，统一标点
    processed_text = re.sub(r'\n+', '\n', text.strip())
    processed_text = re.sub(r'\s+', ' ', processed_text)
    # 替换全角标点为半角，便于统一处理
    processed_text = processed_text.replace('。', '.').replace('，', ',').replace('；', ';')
    processed_text = processed_text.replace('：', ':').replace('（', '(').replace('）', ')')
    
    # 中文条款常见编号格式（正则模式）
    patterns = [
        r'(\d+\.\d+\.\d+\s+)',        # 1.1.1 
        r'(\d+\.\d+\s+)',             # 1.1 
        r'(\d+\.\s+)',                # 1. 
        r'(\(\d+\)\.\s+)',            # (1). 
        r'(\(\d+\)\s+)',              # (1) 
        r'([一二三四五六七八九十]+、\s+)',  # 一、 
        r'(第[一二三四五六七八九十]+条\s+)', # 第一条
        r'(第[一二三四五六七八九十]+款\s+)', # 第一款
        r'(第[一二三四五六七八九十]+项\s+)', # 第一项
        r'(\d+\)\s+)',                # 1)
        r'([A-Za-z]\.\s+)',           # A. 
        r'([A-Za-z]\)\s+)',           # A)
    ]
    
    try:
        # 组合所有模式
        combined_pattern = '|'.join(patterns)
        
        # 查找所有匹配的条款编号位置
        matches = list(re.finditer(combined_pattern, processed_text, re.MULTILINE))
        
        if not matches:
            # 如果没有找到标准编号，尝试按空行拆分
            st.info("未检测到标准条款编号，尝试按空行拆分")
            terms = [t.strip() for t in re.split(r'\n\s*\n', processed_text) if t.strip()]
            return terms
        
        # 提取条款内容
        terms = []
        # 第一条条款从文本开始到第一个匹配
        first_match = matches[0]
        if first_match.start() > 0:
            pre_text = processed_text[:first_match.start()].strip()
            if pre_text:
                terms.append(pre_text)
        
        # 处理中间的条款
        for i in range(len(matches)):
            current_match = matches[i]
            start_idx = current_match.start()
            
            # 确定当前条款的结束位置
            if i < len(matches) - 1:
                end_idx = matches[i+1].start()
            else:
                end_idx = len(processed_text)
            
            # 提取条款内容（包含编号）
            term_content = processed_text[start_idx:end_idx].strip()
            if term_content:
                terms.append(term_content)
        
        # 过滤过短的条款（可能是误拆分）
        min_term_length = 10  # 最小条款长度
        terms = [term for term in terms if len(term) >= min_term_length]
        
        st.success(f"成功拆分条款：{len(terms)}条")
        return terms
        
    except re.error as e:
        st.error(f"正则表达式错误: {str(e)}")
        # 出错时的备选方案：按空行拆分
        st.info("使用备选方案拆分条款")
        terms = [t.strip() for t in re.split(r'\n\s*\n', processed_text) if t.strip()]
        return terms
    except Exception as e:
        st.error(f"条款拆分失败: {str(e)}")
        return [processed_text]  # 返回原始文本作为最后的备选


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
    prompt = f"""请对比以下两个条款的匹配程度：
    【基准条款】：{bench_term[:500]}
    【待比条款】：{compare_term[:500]}
    
    请按以下格式回答：
    1. 匹配度（0-100分）：[分数]
    2. 相同点：[简要说明相同内容]
    3. 匹配依据：[说明为何认为这两个条款相匹配]
    """
    
    result, error = call_qwen_api(prompt, api_key)
    if error:
        return None, error
    
    # 解析结果
    try:
        score_match = re.search(r'匹配度（0-100分）：(\d+)', result)
        score = int(score_match.group(1)) if score_match else 0
        
        return {
            "score": score,
            "full_analysis": result
        }, None
    except:
        return {
            "score": 0,
            "full_analysis": f"解析失败，原始结果：{result}"
        }, None


### 4. 结果报告生成（只包含匹配条款）
def generate_word_report(bench_terms, comparison_results, bench_filename):
    """生成只包含匹配条款的Word报告"""
    doc = docx.Document()
    
    # 设置中文字体
    style = doc.styles['Normal']
    style.font.name = 'SimSun'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    style.font.size = Pt(10.5)
    
    # 标题
    title = doc.add_heading("条款匹配分析报告", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 基本信息
    doc.add_paragraph(f"报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"基准文件：{bench_filename}")
    doc.add_paragraph(f"对比文件数量：{len(comparison_results)}")
    doc.add_page_break()
    
    # 按文件生成结果（只包含匹配条款）
    for file_name, result in comparison_results.items():
        doc.add_heading(f"对比文件：{file_name}", level=1)
        
        # 只展示匹配条款
        doc.add_heading(f"匹配条款（共{len(result['matched_terms'])}条）", level=2)
        if result["matched_terms"]:
            for idx, item in enumerate(result["matched_terms"], 1):
                doc.add_heading(f"{idx}. 基准条款：{item['bench_term'][:30]}...", level=3)
                doc.add_paragraph(f"对比条款：{item['compare_term'][:50]}...")
                doc.add_paragraph(f"匹配度：{item['analysis']['score']}分")
                doc.add_paragraph("匹配分析：")
                doc.add_paragraph(item['analysis']['full_analysis'], style='Normal')
        else:
            doc.add_paragraph("未发现匹配条款")
        
        doc.add_page_break()
    
    # 保存到内存
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


### 5. 主函数
def main():
    st.title("📄 条款匹配分析工具")
    st.write("只展示基准文件与对比文件中可匹配的条款，支持PDF/DOCX格式")
    
    # 侧边栏配置
    with st.sidebar:
        st.subheader("配置")
        qwen_api_key = st.text_input("阿里云DashScope API密钥", type="password")
        st.info("获取密钥：https://dashscope.console.aliyun.com/")
        match_threshold = st.slider("匹配度阈值（分）", 0, 100, 70)
        st.write("高于此分数的条款将被视为匹配")
    
    # 文件上传
    col1, col2 = st.columns(2)
    with col1:
        bench_file = st.file_uploader("上传基准文件（PDF/DOCX）", type=["pdf", "docx"], accept_multiple_files=False)
    with col2:
        compare_files = st.file_uploader("上传对比文件（PDF/DOCX）", type=["pdf", "docx"], accept_multiple_files=True)
    
    # 分析按钮
    if st.button("开始分析", disabled=not (bench_file and compare_files)):
        with st.spinner("正在处理基准文件..."):
            # 1. 提取基准文件文本并拆分条款
            bench_type = bench_file.name.split('.')[-1].lower()
            bench_text = extract_text_from_file(bench_file, bench_type)
            
            # 显示部分提取的文本供参考
            with st.expander("查看提取的基准文件文本（前500字符）"):
                st.text(bench_text[:500])
            
            bench_terms = split_chinese_terms(bench_text)
            st.session_state.bench_terms = bench_terms
        
        # 2. 处理每个对比文件
        all_results = {}
        progress_bar = st.progress(0)
        
        for file_idx, compare_file in enumerate(compare_files):
            file_name = compare_file.name
            st.subheader(f"处理对比文件：{file_name}")
            
            # 提取文本并拆分条款
            compare_type = file_name.split('.')[-1].lower()
            compare_text = extract_text_from_file(compare_file, compare_type)
            
            with st.expander(f"查看提取的{file_name}文本（前500字符）"):
                st.text(compare_text[:500])
            
            compare_terms = split_chinese_terms(compare_text)
            st.session_state.comparison_terms[file_name] = compare_terms
            
            # 条款对比分析（只保留匹配的条款）
            matched_terms = []
            
            with st.spinner(f"正在分析 {file_name} 的条款匹配度..."):
                # 为每个基准条款寻找最匹配的对比条款
                for bench_term in bench_terms[:30]:  # 限制处理数量，避免超时
                    best_match = None
                    highest_score = 0
                    
                    for compare_term in compare_terms[:30]:
                        # 调用Qwen分析
                        if qwen_api_key:
                            analysis, error = analyze_terms_with_qwen(bench_term, compare_term, qwen_api_key)
                            if error:
                                st.warning(f"条款分析失败：{error}")
                                continue
                        else:
                            # 无API时的基础判断
                            common_words = len(set(bench_term[:100]) & set(compare_term[:100]))
                            score = min(100, common_words * 5)  # 简单的关键词匹配计分
                            analysis = {
                                "score": score,
                                "full_analysis": "未使用Qwen API，基于关键词匹配"
                            }
                        
                        # 记录最高匹配度
                        if analysis["score"] > highest_score:
                            highest_score = analysis["score"]
                            best_match = {
                                "bench_term": bench_term,
                                "compare_term": compare_term,
                                "analysis": analysis
                            }
                    
                    # 只保留达到阈值的匹配条款
                    if best_match and highest_score >= match_threshold:
                        matched_terms.append(best_match)
            
            # 保存结果（只包含匹配条款）
            all_results[file_name] = {
                "matched_terms": matched_terms
            }
            
            # 显示当前文件的匹配结果
            st.success(f"{file_name} 分析完成，找到 {len(matched_terms)} 条匹配条款")
            if matched_terms:
                with st.expander(f"查看 {file_name} 的匹配条款"):
                    for i, match in enumerate(matched_terms, 1):
                        st.write(f"**匹配项 {i}（匹配度：{match['analysis']['score']}分）**")
                        st.text(f"基准条款：{match['bench_term'][:100]}...")
                        st.text(f"对比条款：{match['compare_term'][:100]}...")
                        with st.expander("查看详细分析"):
                            st.text(match['analysis']['full_analysis'])
                        st.divider()
            
            # 更新进度
            progress_bar.progress((file_idx + 1) / len(compare_files))
        
        # 保存所有结果
        st.session_state.analysis_results = all_results
        st.success("所有文件分析完成！")
        
        # 生成报告
        if st.button("生成匹配条款报告"):
            with st.spinner("正在生成报告..."):
                report_buffer = generate_word_report(
                    bench_terms, 
                    all_results, 
                    bench_file.name
                )
                st.download_button(
                    label="下载Word报告",
                    data=report_buffer,
                    file_name=f"条款匹配分析报告_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )


if __name__ == "__main__":
    main()
    
