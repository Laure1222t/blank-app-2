import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os
import tempfile
from datetime import datetime
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
st.write("上传基准文件和待比较文件，系统将进行条款匹配分析并生成合规性报告。")

# 初始化会话状态
if 'api_key_valid' not in st.session_state:
    st.session_state.api_key_valid = False

# Qwen API密钥配置
with st.sidebar:
    st.subheader("Qwen大模型配置")
    qwen_api_key = st.text_input("请输入阿里云DashScope API密钥", type="password")
    
    # 验证API密钥
    if qwen_api_key:
        os.environ["DASHSCOPE_API_KEY"] = qwen_api_key
        # 简单验证格式（实际有效性需调用API时才知道）
        if len(qwen_api_key) == 32 and qwen_api_key.startswith('sk-'):
            st.session_state.api_key_valid = True
            st.success("API密钥格式有效")
        else:
            st.session_state.api_key_valid = False
            st.warning("API密钥格式似乎不正确，应为以sk-开头的32位字符串")
    else:
        st.session_state.api_key_valid = False
        st.info("需要阿里云账号和DashScope服务访问权限，获取API密钥: https://dashscope.console.aliyun.com/")
        st.info("若无API密钥，将使用基础模式进行文本比对")

# 检查Tesseract是否安装
def check_tesseract_installation():
    try:
        # 尝试获取Tesseract版本信息
        pytesseract.get_tesseract_version()
        return True
    except pytesseract.TesseractNotFoundError:
        return False
    except Exception as e:
        st.error(f"Tesseract检查出错: {str(e)}")
        return False

# 检查Tesseract状态并提示
tesseract_available = check_tesseract_installation()
if not tesseract_available:
    with st.sidebar:
        st.warning("⚠️ 未检测到Tesseract OCR引擎，图片型PDF处理功能将受限")
        st.info("""
        安装Tesseract指南：
        1. 下载安装包：https://github.com/UB-Mannheim/tesseract/wiki
        2. 安装时选择中文语言包
        3. 配置环境变量或在设置中指定路径
        """)

# 辅助函数：判断PDF页面是否包含可选文本
def has_selectable_text(page):
    text = page.get_text("text")
    # 过滤空白字符后检查长度
    clean_text = re.sub(r'\s+', '', text)
    return len(clean_text) > 50  # 认为50个以上非空白字符为有效文本

# 辅助函数：从PDF中提取文本（优先文本提取，必要时OCR）
def extract_text_from_pdf(file_path):
    doc = fitz.open(file_path)
    full_text = []
    page_count = len(doc)
    
    with st.spinner(f"正在解析PDF文件（共{page_count}页）..."):
        progress_bar = st.progress(0)
        
        for i, page in enumerate(doc):
            # 更新进度
            progress_bar.progress((i + 1) / page_count)
            
            # 先尝试文本提取
            if has_selectable_text(page):
                text = page.get_text("text")
                full_text.append(f"[页面{i+1} - 文本提取]\n{text}")
            else:
                # 文本提取失败，尝试OCR
                if not tesseract_available:
                    full_text.append(f"[页面{i+1} - 无法处理]\n警告：未安装Tesseract OCR，无法提取图片中的文本内容。")
                    continue
                
                try:
                    # 将页面转换为图片
                    pix = page.get_pixmap(dpi=300)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    
                    # 预处理：转为灰度并二值化增强识别率
                    img_gray = img.convert('L')
                    threshold = 150  # 阈值可调整
                    img_binary = img_gray.point(lambda p: p > threshold and 255)
                    
                    # 进行OCR识别（中英文）
                    ocr_text = pytesseract.image_to_string(img_binary, lang="chi_sim+eng")
                    full_text.append(f"[页面{i+1} - OCR识别]\n{ocr_text}")
                except Exception as e:
                    full_text.append(f"[页面{i+1} - 处理失败]\n错误：{str(e)}")
        
        progress_bar.empty()
    
    return '\n\n'.join(full_text)

# 辅助函数：从docx文件中提取文本
def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():  # 只添加非空段落
            full_text.append(para.text)
    return '\n'.join(full_text)

# 统一的文件提取函数
def extract_text_from_file(uploaded_file, file_type):
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_type}') as temp_file:
            temp_file.write(uploaded_file.getvalue())
            temp_path = temp_file.name
        
        # 根据文件类型提取文本
        if file_type == 'pdf':
            text = extract_text_from_pdf(temp_path)
        elif file_type == 'docx':
            text = extract_text_from_docx(temp_path)
        else:
            text = ""
        
        # 清理临时文件
        os.unlink(temp_path)
        return text
    except Exception as e:
        st.error(f"文件处理出错: {str(e)}")
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
    
    # 条款拆分效果评估
    if len(terms) < 3 and len(text) > 500:
        st.info(f"检测到可能的条款拆分效果不佳（共{len(terms)}条），建议检查文件格式")
    
    return terms

# 基础模式的条款匹配（无API密钥时使用）
def basic_term_matching(benchmark_term, compare_terms):
    """简单的基于关键词的条款匹配"""
    best_match = None
    best_score = 0
    
    # 提取基准条款关键词
    bench_words = set(re.findall(r'[\u4e00-\u9fff]+', benchmark_term))  # 提取中文字符
    bench_words.update(re.findall(r'\b[a-zA-Z]+\b', benchmark_term))  # 提取英文字符
    bench_words = [w for w in bench_words if len(w) > 1]  # 过滤单字
    
    if not bench_words:
        return None, 0
    
    for term in compare_terms:
        # 提取对比条款关键词
        term_words = set(re.findall(r'[\u4e00-\u9fff]+', term))
        term_words.update(re.findall(r'\b[a-zA-Z]+\b', term))
        term_words = [w for w in term_words if len(w) > 1]
        
        if not term_words:
            continue
            
        # 计算相似度（交集/并集）
        common = len(bench_words & term_words)
        total = len(bench_words | term_words)
        score = common / total if total > 0 else 0
        
        if score > best_score:
            best_score = score
            best_match = term
    
    return best_match, best_score

# 生成Word报告
def generate_word_report(benchmark_name, compare_results):
    doc = docx.Document()
    
    # 标题
    title = doc.add_heading('条款合规性对比报告', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 基本信息
    doc.add_paragraph(f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph(f"基准文件: {benchmark_name}")
    doc.add_paragraph("")
    
    # 目录
    doc.add_heading('目录', level=1)
    for i, (file_name, _) in enumerate(compare_results.items()):
        doc.add_paragraph(f"{i+1}. {file_name}", style='List Number')
    doc.add_paragraph("")
    
    # 详细结果
    for file_name, result in compare_results.items():
        doc.add_heading(f"文件: {file_name}", level=1)
        
        # 可匹配条款
        doc.add_heading("可匹配条款", level=2)
        if result['matched']:
            for idx, item in enumerate(result['matched'], 1):
                doc.add_heading(f"匹配项 {idx} (相似度: {item['score']:.2f})", level=3)
                
                p = doc.add_paragraph("基准条款: ")
                p.add_run(item['benchmark']).bold = True
                
                p = doc.add_paragraph("对比条款: ")
                p.add_run(item['compare']).bold = True
                
                if 'analysis' in item:
                    doc.add_paragraph(f"分析: {item['analysis']}")
        else:
            doc.add_paragraph("未找到可匹配的条款")
        
        # 不合规条款
        doc.add_heading("不合规条款总结", level=2)
        if result['non_compliant']:
            for idx, term in enumerate(result['non_compliant'], 1):
                doc.add_paragraph(f"{idx}. {term}", style='List Number')
        else:
            doc.add_paragraph("未发现不合规条款")
        
        doc.add_page_break()
    
    # 保存到内存
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
        # 提取基准文件文本和条款
        with st.spinner("正在处理基准文件..."):
            bench_type = benchmark_file.name.split('.')[-1].lower()
            bench_text = extract_text_from_file(benchmark_file, bench_type)
            
            if not bench_text:
                st.error("无法从基准文件中提取文本内容")
                return
                
            st.success(f"基准文件处理完成，提取到文本长度: {len(bench_text)}字符")
            bench_terms = split_chinese_terms(bench_text)
            st.info(f"从基准文件中拆分出 {len(bench_terms)} 条条款")
        
        # 处理每个对比文件
        compare_results = {}
        use_advanced = st.session_state.api_key_valid
        
        if use_advanced:
            st.info("将使用Qwen大模型进行高级条款匹配分析")
        else:
            st.info("未检测到有效API密钥，将使用基础模式进行条款匹配")
        
        # 显示总体进度
        progress_bar = st.progress(0)
        total_files = len(compare_files)
        
        for file_idx, compare_file in enumerate(compare_files, 1):
            st.subheader(f"正在处理: {compare_file.name}")
            
            # 提取对比文件文本和条款
            with st.spinner(f"提取文本和拆分条款..."):
                comp_type = compare_file.name.split('.')[-1].lower()
                comp_text = extract_text_from_file(compare_file, comp_type)
                
                if not comp_text:
                    st.warning(f"无法从 {compare_file.name} 中提取文本内容，跳过该文件")
                    progress_bar.progress(file_idx / total_files)
                    continue
                    
                comp_terms = split_chinese_terms(comp_text)
                st.info(f"从 {compare_file.name} 中拆分出 {len(comp_terms)} 条条款")
            
            # 条款匹配分析
            matched_terms = []
            comp_terms_used = set()  # 跟踪已匹配的条款
            
            with st.spinner(f"正在进行条款匹配分析..."):
                for bench_idx, bench_term in enumerate(bench_terms[:20]):  # 限制前20条以提高效率
                    # 显示当前进度
                    if len(bench_terms) > 0:
                        sub_progress = (bench_idx + 1) / len(bench_terms)
                        st.progress(sub_progress, text=f"处理条款 {bench_idx + 1}/{len(bench_terms)}")
                    
                    # 查找最佳匹配
                    if use_advanced:
                        # 这里应该是调用Qwen大模型的代码
                        # 为了避免错误，当API不可用时使用基础模式
                        try:
                            from dashscope import Generation
                            
                            prompt = f"""
                            请对比以下两个条款的内容，并判断它们的匹配程度（0-100分）。
                            同时分析它们的相同点和不同点，并给出合规性判断。
                            
                            基准条款: {bench_term[:200]}
                            
                            请从以下对比条款中找到最匹配的一项:
                            {chr(10).join([f"{i+1}. {t[:100]}..." for i, t in enumerate(comp_terms)])}
                            
                            请以JSON格式返回:
                            {{
                                "best_match_index": 最匹配条款的索引(从0开始),
                                "similarity_score": 匹配度(0-100),
                                "analysis": "相同点和不同点分析，以及合规性判断"
                            }}
                            """
                            
                            response = Generation.call(
                                model="qwen-plus",
                                prompt=prompt,
                                result_format="json"
                            )
                            
                            if response.status_code == 200:
                                try:
                                    analysis_result = json.loads(response.output.text)
                                    match_idx = analysis_result.get("best_match_index", -1)
                                    score = analysis_result.get("similarity_score", 0) / 100  # 转换为0-1范围
                                    analysis = analysis_result.get("analysis", "")
                                    
                                    if 0 <= match_idx < len(comp_terms) and match_idx not in comp_terms_used:
                                        comp_terms_used.add(match_idx)
                                        matched_terms.append({
                                            "benchmark": bench_term,
                                            "compare": comp_terms[match_idx],
                                            "score": score,
                                            "analysis": analysis
                                        })
                                except:
                                    # 解析结果失败，使用基础模式
                                    best_match, score = basic_term_matching(bench_term, comp_terms)
                                    if best_match:
                                        matched_terms.append({
                                            "benchmark": bench_term,
                                            "compare": best_match,
                                            "score": score
                                        })
                            else:
                                # API调用失败，使用基础模式
                                best_match, score = basic_term_matching(bench_term, comp_terms)
                                if best_match:
                                    matched_terms.append({
                                        "benchmark": bench_term,
                                        "compare": best_match,
                                        "score": score
                                    })
                        except Exception as e:
                            st.warning(f"高级分析出错，使用基础模式: {str(e)}")
                            best_match, score = basic_term_matching(bench_term, comp_terms)
                            if best_match:
                                matched_terms.append({
                                    "benchmark": bench_term,
                                    "compare": best_match,
                                    "score": score
                                })
                    else:
                        # 使用基础模式
                        best_match, score = basic_term_matching(bench_term, comp_terms)
                        if best_match:
                            matched_terms.append({
                                "benchmark": bench_term,
                                "compare": best_match,
                                "score": score
                            })
            
            # 筛选出匹配度高的条款（>0.7）
            valid_matches = [m for m in matched_terms if m['score'] > 0.7]
            valid_matches.sort(key=lambda x: x['score'], reverse=True)
            
            # 找出未匹配的条款（不合规）
            non_compliant = [comp_terms[i] for i in range(len(comp_terms)) if i not in comp_terms_used]
            
            # 保存结果
            compare_results[compare_file.name] = {
                "matched": valid_matches,
                "non_compliant": non_compliant[:10]  # 限制显示前10条
            }
            
            # 更新总体进度
            progress_bar.progress(file_idx / total_files)
        
        progress_bar.empty()
        
        # 显示结果
        st.success("所有文件分析完成！")
        
        # 创建结果标签页
        tabs = st.tabs([f"📄 {name}" for name in compare_results.keys()])
        
        for tab, (file_name, result) in zip(tabs, compare_results.items()):
            with tab:
                st.header(f"文件: {file_name}")
                
                # 显示匹配条款
                st.subheader("可匹配条款")
                if result['matched']:
                    for i, item in enumerate(result['matched']):
                        with st.expander(f"匹配项 {i+1} (相似度: {item['score']:.2f})"):
                            col_a, col_b = st.columns(2)
                            with col_a:
                                st.markdown("**基准条款:**")
                                st.write(item['benchmark'])
                            with col_b:
                                st.markdown("**对比条款:**")
                                st.write(item['compare'])
                            if 'analysis' in item:
                                st.markdown("**分析:**")
                                st.write(item['analysis'])
                else:
                    st.info("未找到可匹配的条款")
                
                # 显示不合规条款
                st.subheader("不合规条款总结")
                if result['non_compliant']:
                    for i, term in enumerate(result['non_compliant']):
                        st.write(f"{i+1}. {term[:200]}...")  # 显示前200字符
                else:
                    st.success("未发现不合规条款")
        
        # 生成并提供下载报告
        st.subheader("生成报告")
        report_buffer = generate_word_report(benchmark_file.name, compare_results)
        
        # 提供下载
        b64 = base64.b64encode(report_buffer.getvalue()).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="条款合规性对比报告_{datetime.now().strftime("%Y%m%d")}.docx">下载Word报告</a>'
        st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
    
