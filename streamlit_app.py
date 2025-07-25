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
import pytesseract
from PIL import Image
import numpy as np
import io
import base64

# 设置页面配置
st.set_page_config(
    page_title="条款合规性对比工具",
    page_icon="📄",
    layout="wide"
)

# 页面标题
st.title("📄 条款合规性对比工具")
st.write("上传基准文件和多个对比文件，系统将分析条款匹配情况并生成合规性报告。")

# 检查Tesseract是否安装
def check_tesseract_installation():
    try:
        # 尝试运行tesseract命令
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False

# 配置Tesseract路径（针对本地运行）
def configure_tesseract():
    if not check_tesseract_installation():
        with st.sidebar:
            st.warning("未检测到Tesseract OCR引擎，图片型PDF将无法处理")
            st.info("""
            安装指南：
            1. 下载安装Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
            2. 安装时勾选中文语言包
            3. 在下方输入安装路径（如C:\\Program Files\\Tesseract-OCR\\tesseract.exe）
            """)
            tesseract_path = st.text_input("Tesseract安装路径")
            if tesseract_path:
                try:
                    pytesseract.pytesseract.tesseract_cmd = tesseract_path
                    if check_tesseract_installation():
                        st.success("Tesseract配置成功")
                except Exception as e:
                    st.error(f"配置失败: {str(e)}")
    return check_tesseract_installation()

# Qwen API密钥配置
with st.sidebar:
    st.subheader("Qwen大模型配置")
    qwen_api_key = st.text_input("请输入阿里云DashScope API密钥", type="password")
    if qwen_api_key:
        os.environ["DASHSCOPE_API_KEY"] = qwen_api_key
    st.info("需要阿里云账号和DashScope服务访问权限，获取API密钥: https://dashscope.console.aliyun.com/")
    
    # 配置Tesseract
    tesseract_available = configure_tesseract()

# 检查页面是否包含可选择的文本
def has_selectable_text(page):
    text = page.get_text().strip()
    # 如果文本长度大于50个字符，认为是可选择的文本
    return len(text) > 50

# 从PDF中提取文本（优先文本提取，必要时使用OCR）
def extract_text_from_pdf(file_path):
    doc = fitz.open(file_path)
    full_text = []
    tesseract_available = check_tesseract_installation()
    
    with st.spinner("正在提取PDF内容..."):
        progress_bar = st.progress(0)
        for i, page in enumerate(doc):
            # 检查是否有可选择的文本
            if has_selectable_text(page):
                text = page.get_text().strip()
                full_text.append(f"[文本提取] 第{i+1}页:\n{text}")
            else:
                # 没有可选择的文本，尝试OCR
                if tesseract_available:
                    with st.spinner(f"正在对第{i+1}页进行OCR识别..."):
                        # 将页面转换为图片
                        pix = page.get_pixmap(dpi=300)
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        
                        # 预处理：转为灰度并二值化
                        img_gray = img.convert('L')
                        img_np = np.array(img_gray)
                        thresh = 150  # 阈值调整
                        img_binary = (img_np > thresh) * 255
                        img_processed = Image.fromarray(img_binary.astype(np.uint8))
                        
                        # 进行OCR识别，支持中英文
                        ocr_text = pytesseract.image_to_string(
                            img_processed, 
                            lang="chi_sim+eng"
                        ).strip()
                        
                        full_text.append(f"[OCR识别] 第{i+1}页:\n{ocr_text}")
                else:
                    full_text.append(f"[无法识别] 第{i+1}页: 未安装Tesseract OCR，无法处理图片型PDF内容")
            
            progress_bar.progress((i + 1) / len(doc))
    
    return '\n\n'.join(full_text)

# 从docx文件中提取文本
def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():  # 只添加非空段落
            full_text.append(para.text)
    return '\n'.join(full_text)

# 统一的文件文本提取函数
def extract_text_from_file(uploaded_file, file_type):
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_type}') as temp_file:
            temp_file.write(uploaded_file.getvalue())
            temp_path = temp_file.name
        
        # 根据文件类型提取文本
        if file_type == 'pdf':
            text = extract_text_from_pdf(temp_path)
        elif file_type in ['docx', 'doc']:  # 简单处理，实际doc需要额外库
            text = extract_text_from_docx(temp_path)
        else:
            text = ""
        
        # 清理临时文件
        os.unlink(temp_path)
        return text
    except Exception as e:
        st.error(f"文件处理错误: {str(e)}")
        return ""

# 优化的中文条款拆分函数
def split_chinese_terms(text):
    """拆分中文条款，支持多种编号格式，增加空值和异常处理"""
    # 首先检查输入是否有效
    if not text or not isinstance(text, str):
        st.warning("输入文本为空或无效，无法进行条款拆分")
        return []
    
    # 清除多余空行和空格
    text = re.sub(r'\n+', '\n', text.strip())
    
    # 中文条款常见的编号格式正则表达式
    patterns = [
        r'(\d+\.\s+)',                # 1. 
        r'(\d+\.\d+\s+)',             # 1.1 
        r'(\(\d+\)\s+)',              # (1) 
        r'([一二三四五六七八九十]+\、\s+)',  # 一、 
        r'(第[一二三四五六七八九十]条\s+)',   # 第一条
        r'(第[一二三四五六七八九十]款\s+)',   # 第一款
        r'(\d+\)\s+)',                # 1)
        r'([A-Za-z]\.\s+)',           # A. 
    ]
    
    # 组合所有模式
    combined_pattern = '|'.join(patterns)
    
    # 拆分文本
    parts = re.split(combined_pattern, text)
    
    terms = []
    current_term = ""
    
    for part in parts:
        # 跳过空值或仅含空白字符的部分
        if not part or not part.strip():
            continue
            
        # 检查当前部分是否为条款编号
        is_numbering = any(re.fullmatch(pattern.strip(), part.strip()) for pattern in patterns)
        
        if is_numbering:
            # 如果已有内容，先保存当前条款
            if current_term.strip():
                terms.append(current_term.strip())
            # 开始新条款
            current_term = part
        else:
            # 累加条款内容
            current_term += part
    
    # 添加最后一个条款
    if current_term.strip():
        terms.append(current_term.strip())
    
    # 如果条款数量较少，可能是拆分效果不好，提示用户
    if len(terms) < 3 and len(text) > 500:
        st.info(f"检测到可能的条款拆分效果不佳（共{len(terms)}条），建议检查文件格式")
    
    return terms

# 使用Qwen大模型进行条款匹配和合规性分析
def analyze_terms_with_qwen(benchmark_term, compare_terms):
    if not qwen_api_key:
        st.error("请先配置Qwen API密钥")
        return None, "未配置API密钥"
    
    prompt = f"""你是一个条款合规性分析专家。请分析对比条款与基准条款的匹配程度和差异。
    基准条款: {benchmark_term}
    
    待比较条款列表:
    {chr(10).join([f"{i+1}. {term}" for i, term in enumerate(compare_terms)])}
    
    请先判断哪个待比较条款与基准条款最匹配，然后分析它们的差异。
    输出格式要求:
    1. 匹配条款编号: [数字，如1表示第一个待比较条款]
    2. 匹配度: [0-100的数字，表示匹配百分比]
    3. 相同点: [简要描述相同内容]
    4. 差异点: [简要描述不同内容]
    5. 合规性判断: [合规/部分合规/不合规]
    6. 理由: [说明判断依据]
    
    请用中文输出，确保结果简洁明了。
    """
    
    try:
        response = Generation.call(
            model="qwen-plus",
            prompt=prompt
        )
        
        if response.status_code == 200:
            result = response.output.text
            # 提取匹配度
            match_score = re.search(r'匹配度: (\d+)', result)
            score = int(match_score.group(1)) if match_score else 0
            return score, result
        else:
            st.error(f"Qwen API调用失败: {response.message}")
            return 0, f"API调用失败: {response.message}"
    except Exception as e:
        st.error(f"分析出错: {str(e)}")
        return 0, f"分析出错: {str(e)}"

# 生成Word报告
def generate_word_report(benchmark_name, compare_results):
    doc = docx.Document()
    
    # 添加标题
    title = doc.add_heading("条款合规性对比报告", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 添加报告信息
    doc.add_paragraph(f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph(f"基准文件: {benchmark_name}")
    doc.add_paragraph(f"对比文件数量: {len(compare_results)}")
    doc.add_page_break()
    
    # 添加目录
    doc.add_heading("目录", 1)
    for i, (file_name, _) in enumerate(compare_results.items(), 1):
        para = doc.add_paragraph(f"{i}. {file_name}", style='List Number')
        para.hyperlink = f"#{file_name}"  # 简单的目录链接标记
    
    doc.add_page_break()
    
    # 为每个对比文件添加分析结果
    for file_name, analysis in compare_results.items():
        # 文件标题
        heading = doc.add_heading(file_name, 1)
        heading.paragraph_format.keep_with_next = True
        
        # 匹配的条款
        doc.add_heading("可匹配条款", 2)
        matched_terms = [t for t in analysis if t['score'] >= 70]
        
        if matched_terms:
            for term in matched_terms:
                doc.add_heading(f"基准条款: {term['benchmark_term'][:30]}...", 3)
                doc.add_paragraph(f"匹配条款: {term['matched_term'][:50]}...")
                doc.add_paragraph(f"匹配度: {term['score']}%")
                doc.add_paragraph("分析:")
                doc.add_paragraph(term['analysis'], style='List Bullet')
                doc.add_paragraph("")
        else:
            doc.add_paragraph("未发现可匹配的条款")
        
        # 不合规的条款
        doc.add_heading("不合规条款总结", 2)
        non_compliant = [t for t in analysis if t['score'] < 70]
        
        if non_compliant:
            for term in non_compliant:
                para = doc.add_paragraph(f"基准条款: {term['benchmark_term'][:50]}...", style='List Number')
                para.runs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # 红色
                doc.add_paragraph(f"匹配度: {term['score']}%")
                doc.add_paragraph("分析:")
                doc.add_paragraph(term['analysis'], style='List Bullet')
                doc.add_paragraph("")
        else:
            doc.add_paragraph("未发现不合规的条款")
        
        doc.add_page_break()
    
    # 保存到字节流
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

# 主函数
def main():
    # 上传文件
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("基准文件")
        benchmark_file = st.file_uploader("上传基准文件 (PDF或DOCX)", type=['pdf', 'docx'], key='benchmark')
    
    with col2:
        st.subheader("对比文件")
        compare_files = st.file_uploader(
            "上传一个或多个对比文件 (PDF或DOCX)", 
            type=['pdf', 'docx'], 
            key='compare',
            accept_multiple_files=True
        )
    
    # 分析按钮
    if st.button("开始分析", disabled=not (benchmark_file and compare_files)):
        # 处理基准文件
        st.subheader("正在处理基准文件...")
        bench_type = benchmark_file.name.split('.')[-1].lower()
        bench_text = extract_text_from_file(benchmark_file, bench_type)
        
        if not bench_text:
            st.error("无法从基准文件中提取文本内容")
            return
        
        # 拆分基准条款
        st.info("正在拆分基准条款...")
        bench_terms = split_chinese_terms(bench_text)
        st.success(f"成功拆分出 {len(bench_terms)} 条基准条款")
        
        # 显示部分基准条款
        with st.expander("查看部分基准条款"):
            for i, term in enumerate(bench_terms[:5]):
                st.write(f"{i+1}. {term[:100]}...")
        
        # 处理每个对比文件
        compare_results = {}
        
        for compare_file in compare_files:
            st.subheader(f"正在处理对比文件: {compare_file.name}")
            file_type = compare_file.name.split('.')[-1].lower()
            compare_text = extract_text_from_file(compare_file, file_type)
            
            if not compare_text:
                st.warning(f"无法从 {compare_file.name} 中提取文本内容，跳过该文件")
                continue
            
            # 拆分对比条款
            st.info(f"正在拆分 {compare_file.name} 的条款...")
            compare_terms = split_chinese_terms(compare_text)
            st.success(f"成功拆分出 {len(compare_terms)} 条对比条款")
            
            # 分析条款匹配情况
            st.info(f"正在分析 {compare_file.name} 与基准文件的匹配情况...")
            progress_bar = st.progress(0)
            analysis_results = []
            
            for i, bench_term in enumerate(bench_terms):
                # 分析当前基准条款与所有对比条款的匹配度
                score, analysis = analyze_terms_with_qwen(bench_term, compare_terms)
                
                # 找到最匹配的条款（简化处理，实际应遍历所有对比条款）
                matched_term = compare_terms[0] if compare_terms else "无对应条款"
                
                analysis_results.append({
                    "benchmark_term": bench_term,
                    "matched_term": matched_term,
                    "score": score,
                    "analysis": analysis
                })
                
                progress_bar.progress((i + 1) / len(bench_terms))
            
            compare_results[compare_file.name] = analysis_results
        
        # 显示结果
        st.subheader("分析结果")
        tabs = st.tabs(list(compare_results.keys()))
        
        for tab, (file_name, results) in zip(tabs, compare_results.items()):
            with tab:
                # 显示匹配的条款
                st.subheader("可匹配条款")
                matched = [r for r in results if r['score'] >= 70]
                
                if matched:
                    for i, res in enumerate(matched[:10]):  # 只显示前10条
                        with st.expander(f"基准条款 {i+1} (匹配度: {res['score']}%)"):
                            st.write("**基准条款:**", res['benchmark_term'])
                            st.write("**匹配条款:**", res['matched_term'])
                            st.write("**分析:**", res['analysis'])
                    if len(matched) > 10:
                        st.info(f"共 {len(matched)} 条匹配条款，显示前10条")
                else:
                    st.info("未发现可匹配的条款")
                
                # 显示不合规条款
                st.subheader("不合规条款")
                non_compliant = [r for r in results if r['score'] < 70]
                
                if non_compliant:
                    for i, res in enumerate(non_compliant[:10]):  # 只显示前10条
                        with st.expander(f"基准条款 {i+1} (匹配度: {res['score']}%)"):
                            st.write("**基准条款:**", res['benchmark_term'])
                            st.write("**匹配条款:**", res['matched_term'])
                            st.write("**分析:**", res['analysis'])
                    if len(non_compliant) > 10:
                        st.info(f"共 {len(non_compliant)} 条不合规条款，显示前10条")
                else:
                    st.success("未发现不合规的条款")
        
        # 生成报告
        st.subheader("生成报告")
        if compare_results:
            report_buffer = generate_word_report(benchmark_file.name, compare_results)
            
            # 提供下载
            b64 = base64.b64encode(report_buffer.read()).decode()
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="条款合规性对比报告_{datetime.now().strftime("%Y%m%d")}.docx">下载Word报告</a>'
            st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
    
