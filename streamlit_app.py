import streamlit as st
import fitz  # PyMuPDF
import re
import time
import requests
import json
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 设置页面配置
st.set_page_config(
    page_title="多文件政策比对分析工具",
    page_icon="📜",
    layout="wide"
)

# 自定义CSS
st.markdown("""
<style>
    .stButton>button {
        margin-top: 1rem;
    }
    .analysis-box {
        border: 1px solid #e0e0e0;
        border-radius: 5px;
        padding: 1rem;
        margin-top: 1rem;
        background-color: #f9f9f9;
    }
    .file-tab {
        padding: 0.5rem 1rem;
        border-radius: 4px;
        margin: 0.25rem;
        cursor: pointer;
        display: inline-block;
    }
    .file-tab.active {
        background-color: #007bff;
        color: white;
    }
    .file-tab.inactive {
        background-color: #e9ecef;
        color: #495057;
    }
    .clause-item {
        padding: 0.5rem;
        margin: 0.25rem 0;
        border-radius: 3px;
        background-color: #f0f2f6;
    }
    .parse-status {
        font-size: 0.9rem;
        color: #6c757d;
    }
</style>
""", unsafe_allow_html=True)

# 初始化会话状态
if 'target_clauses' not in st.session_state:
    st.session_state.target_clauses = []
if 'compare_files' not in st.session_state:
    st.session_state.compare_files = {}  # {文件名: {条款: [], 分析结果: ""}}
if 'current_file' not in st.session_state:
    st.session_state.current_file = None
if 'api_key' not in st.session_state:
    st.session_state.api_key = os.getenv("QWEN_API_KEY", "")
if 'max_clauses' not in st.session_state:
    st.session_state.max_clauses = 30  # 默认最大条款数
if 'parse_method' not in st.session_state:
    st.session_state.parse_method = "智能识别"  # 解析方法

# 页面标题
st.title("📜 多文件政策比对分析工具")
st.markdown("上传目标政策文件和多个待比对文件，系统将逐一进行条款比对与合规性分析")
st.markdown("---")

# 条款提取设置
st.sidebar.subheader("条款提取设置")
st.session_state.max_clauses = st.sidebar.slider(
    "最大条款数量", 
    min_value=0, 
    max_value=50, 
    value=st.session_state.max_clauses,
    help="设置从文件中提取的最大条款数量，0表示无限制（最多50条）"
)

# 条款拆分精细度设置
clause_precision = st.sidebar.select_slider(
    "条款拆分精细度",
    options=["粗略", "中等", "精细"],
    value="中等",
    help="设置条款拆分的精细程度，精细模式会识别更多子条款"
)

# 解析方法选择
st.session_state.parse_method = st.sidebar.radio(
    "解析方法",
    ["智能识别", "按标题层级", "按段落拆分"],
    help="当智能识别效果不佳时，可尝试其他解析方法"
)

# API配置
with st.expander("🔑 API 配置", expanded=not st.session_state.api_key):
    st.session_state.api_key = st.text_input("请输入Qwen API密钥", value=st.session_state.api_key, type="password")
    model_option = st.selectbox(
        "选择Qwen模型",
        ["qwen-turbo", "qwen-plus", "qwen-max"],
        index=0  # 默认使用轻量版
    )
    st.caption("提示：可从阿里云DashScope平台获取API密钥，不同模型能力和成本不同")

# 优化的PDF解析函数 - 解决解析不完全问题
def parse_pdf(file, max_clauses=30, precision="中等", method="智能识别"):
    """解析PDF文件并提取结构化条款，优化解析完整性"""
    try:
        with st.spinner("正在解析文件..."):
            doc = fitz.open(stream=file.read(), filetype="pdf")
            total_pages = len(doc)
            text = ""
            page_texts = []  # 存储每页的文本，用于处理跨页条款
            
            # 逐页读取文本，保留页面分隔信息
            for page_num, page in enumerate(doc, 1):
                page_text = page.get_text()
                page_texts.append(f"[[PAGE {page_num}]]\n{page_text}")
                text += page_text + "\n\n"
            
            # 文本预处理 - 增强版
            text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)  # 移除控制字符
            text = re.sub(r'(\r\n|\r|\n)+', '\n', text)  # 统一换行符
            text = re.sub(r'[^\S\n]+', ' ', text)  # 替换非换行的空白字符为空格
            text = text.strip()
            
            # 根据选择的解析方法处理
            clauses = []
            
            if method == "智能识别":
                # 智能识别模式 - 尝试多种模式
                clauses = parse_with_patterns(text, precision)
                # 如果提取的条款太少，尝试其他模式补充
                if len(clauses) < 5:
                    st.markdown('<p class="parse-status">智能识别提取条款较少，尝试补充提取...</p>', unsafe_allow_html=True)
                    heading_clauses = parse_by_headings(text)
                    # 合并条款并去重
                    combined = list(clauses)
                    for clause in heading_clauses:
                        if clause not in combined:
                            combined.append(clause)
                    clauses = combined
            
            elif method == "按标题层级":
                # 按标题层级解析
                clauses = parse_by_headings(text)
            
            else:  # 按段落拆分
                # 按段落拆分模式
                clauses = parse_by_paragraphs(text)
            
            # 后处理：过滤过短条款和空白条款
            clauses = [clause.strip() for clause in clauses if clause.strip() and len(clause.strip()) > 30]
            
            # 处理跨页条款（简单合并可能被分页符分割的条款）
            if len(clauses) > 1:
                merged_clauses = []
                i = 0
                while i < len(clauses):
                    # 检查是否包含页码标记，且不是最后一条
                    if "[[PAGE" in clauses[i] and i < len(clauses) - 1:
                        # 合并当前条款和下一条款
                        merged = clauses[i] + " " + clauses[i+1]
                        merged = re.sub(r'\[\[PAGE \d+\]\]', '', merged)  # 移除页码标记
                        merged_clauses.append(merged)
                        i += 2  # 跳过下一条
                    else:
                        # 移除页码标记
                        clean_clause = re.sub(r'\[\[PAGE \d+\]\]', '', clauses[i])
                        merged_clauses.append(clean_clause)
                        i += 1
                clauses = merged_clauses
            
            # 应用最大条款数限制
            max_clauses = min(max_clauses, 50) if max_clauses > 0 else 50
            final_clauses = clauses[:max_clauses]
            
            # 显示解析状态
            st.markdown(f'<p class="parse-status">共解析 {total_pages} 页，提取 {len(final_clauses)} 条有效条款</p>', unsafe_allow_html=True)
            return final_clauses
            
    except Exception as e:
        st.error(f"文件解析错误: {str(e)}")
        return []

# 按模式识别条款
def parse_with_patterns(text, precision):
    # 根据精细度选择不同的条款提取模式
    patterns = []
    
    if precision == "精细":
        # 精细模式：识别更多类型的条款
        patterns = [
            # 数字编号条款（支持多级）
            re.compile(r'(\d+\.\d+\.\d+\.\d+\s+.*?)(?=\d+\.\d+\.\d+\.\d+\s+|$)', re.DOTALL),  # 四级
            re.compile(r'(\d+\.\d+\.\d+\s+.*?)(?=\d+\.\d+\.\d+\s+|$)', re.DOTALL),          # 三级
            re.compile(r'(\d+\.\d+\s+.*?)(?=\d+\.\d+\s+|$)', re.DOTALL),                  # 二级
            re.compile(r'(\d+\s+.*?)(?=\d+\s+|$)', re.DOTALL),                            # 一级
            
            # 中文编号条款
            re.compile(r'([一二三四五六七八九十百]+\.\s+.*?)(?=[一二三四五六七八九十百]+\.\s+|$)', re.DOTALL),  # 中文数字
            re.compile(r'(\([一二三四五六七八九十]\)\s+.*?)(?=\([一二三四五六七八九十]\)\s+|$)', re.DOTALL),  # 带括号中文
            re.compile(r'([甲乙丙丁戊己庚辛壬癸]+\.\s+.*?)(?=[甲乙丙丁戊己庚辛壬癸]+\.\s+|$)', re.DOTALL),  # 天干
            
            # 字母编号条款
            re.compile(r'([A-Z]\.\s+.*?)(?=[A-Z]\.\s+|$)', re.DOTALL),                    # 大写字母
            re.compile(r'([a-z]\.\s+.*?)(?=[a-z]\.\s+|$)', re.DOTALL),                    # 小写字母
            re.compile(r'(\([A-Za-z]\)\s+.*?)(?=\([A-Za-z]\)\s+|$)', re.DOTALL)           # 带括号字母
        ]
    elif precision == "中等":
        # 中等模式：识别主要层级条款
        patterns = [
            re.compile(r'(\d+\.\d+\.\d+\s+.*?)(?=\d+\.\d+\.\d+\s+|$)', re.DOTALL),          # 三级
            re.compile(r'(\d+\.\d+\s+.*?)(?=\d+\.\d+\s+|$)', re.DOTALL),                  # 二级
            re.compile(r'(\d+\s+.*?)(?=\d+\s+|$)', re.DOTALL),                            # 一级
            re.compile(r'([一二三四五六七八九十]+\.\s+.*?)(?=[一二三四五六七八九十]+\.\s+|$)', re.DOTALL),  # 中文数字
            re.compile(r'([A-Z]\.\s+.*?)(?=[A-Z]\.\s+|$)', re.DOTALL)                     # 大写字母
        ]
    else:  # 粗略
        # 粗略模式：只识别主要条款
        patterns = [
            re.compile(r'(\d+\.\d+\s+.*?)(?=\d+\.\d+\s+|$)', re.DOTALL),                  # 二级
            re.compile(r'(\d+\s+.*?)(?=\d+\s+|$)', re.DOTALL),                            # 一级
            re.compile(r'([一二三四五六七八九十]+\.\s+.*?)(?=[一二三四五六七八九十]+\.\s+|$)', re.DOTALL)   # 中文数字
        ]
    
    clauses = []
    for pattern in patterns:
        matches = pattern.findall(text)
        if matches:
            # 过滤过短的条款
            clauses = [match.strip() for match in matches if len(match.strip()) > 20]
            break
    
    return clauses

# 按标题层级解析
def parse_by_headings(text):
    # 匹配常见的标题格式
    heading_patterns = [
        re.compile(r'(第[一二三四五六七八九十百]+章\s+.*?)(?=第[一二三四五六七八九十百]+章\s+|$)', re.DOTALL),  # 章节
        re.compile(r'(第[一二三四五六七八九十百]+条\s+.*?)(?=第[一二三四五六七八九十百]+条\s+|$)', re.DOTALL),  # 条
        re.compile(r'([一二三四五六七八九十]+\、\s+.*?)(?=[一二三四五六七八九十]+\、\s+|$)', re.DOTALL),        # 中文序号加顿号
    ]
    
    for pattern in heading_patterns:
        matches = pattern.findall(text)
        if matches and len(matches) > 1:
            return [match.strip() for match in matches]
    
    # 如果没有识别到标题，使用通用模式
    return re.split(r'(?<=[。；！？])\s+', text)

# 按段落拆分
def parse_by_paragraphs(text):
    # 使用多种标点符号作为段落分隔符
    separators = r'。(?=\s+)|！(?=\s+)|？(?=\s+)|；(?=\s+)|[\n]{2,}'
    paragraphs = re.split(separators, text)
    # 过滤过短段落并补充结尾标点
    processed = []
    for para in paragraphs:
        para = para.strip()
        if len(para) > 50:
            if not para.endswith(('。', '！', '？', '；', '.')):
                para += '。'
            processed.append(para)
    return processed

# 调用Qwen API进行分析
def call_qwen_api(prompt, api_key, model="qwen-turbo"):
    """调用Qwen API进行合规性分析"""
    if not api_key:
        st.error("请先配置API密钥")
        return None
    
    try:
        with st.spinner("正在调用API进行分析..."):
            url = "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation"
            
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
            
            data = {
                "model": model,
                "input": {
                    "prompt": prompt
                },
                "parameters": {
                    "temperature": 0.6,
                    "top_p": 0.9,
                    "max_tokens": 1500
                }
            }
            
            response = requests.post(url, headers=headers, data=json.dumps(data))
            response_data = response.json()
            
            if response.status_code == 200 and "output" in response_data:
                return response_data["output"]["text"]
            else:
                st.error(f"API调用失败: {response_data.get('message', '未知错误')}")
                return None
                
    except Exception as e:
        st.error(f"API请求错误: {str(e)}")
        return None

# 合规性分析函数
def analyze_compliance(target_clauses, compare_clauses, api_key, model):
    """生成分析提示并调用API"""
    if not target_clauses or not compare_clauses:
        st.warning("缺少条款内容，无法进行分析")
        return None
    
    # 准备条款文本
    target_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(target_clauses)])
    compare_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(compare_clauses)])
    
    # 分析提示词
    prompt = """
    你是政策合规性分析专家，需要比对两份文件的条款并进行合规性分析。请严格按照以下要求执行：
    
    1. 全面覆盖提供的所有条款，不要遗漏重要内容
    2. 重点分析合规性：对于不同之处，判断是否存在冲突、不一致或不合规的情况
    3. 对于相同或一致的条款，简要说明即可
    4. 分析时请基于条款内容本身，不要添加外部知识
    5. 输出格式：
       - 先列出条款对应关系
       - 再分析差异点
       - 最后给出合规性判断及建议
    
    目标政策文件条款：
    {target_text}
    
    待比对文件条款：
    {compare_text}
    
    请用中文详细输出分析结果，确保逻辑清晰、结论明确。
    """.format(target_text=target_text, compare_text=compare_text)
    
    return call_qwen_api(prompt, api_key, model)

# 生成Word文档
def generate_word_document(analysis_result, target_filename, compare_filename):
    """生成Word格式分析报告"""
    try:
        doc = Document()
        
        # 标题
        title = doc.add_heading("政策文件合规性分析报告", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 基本信息
        doc.add_paragraph(f"目标政策文件: {target_filename}")
        doc.add_paragraph(f"待比对文件: {compare_filename}")
        doc.add_paragraph(f"分析日期: {time.strftime('%Y年%m月%d日')}")
        doc.add_paragraph("")
        
        # 分析结果
        doc.add_heading("一、分析结果", level=1)
        
        # 处理分析结果为段落
        paragraphs = re.split(r'\n+', analysis_result)
        for para in paragraphs:
            para = para.strip()
            if para:
                if para.startswith(('1.', '2.', '3.')) or para.endswith('：'):
                    p = doc.add_paragraph(para)
                    p.style = 'Heading 2'
                else:
                    p = doc.add_paragraph(para)
                    p.paragraph_format.space_after = Pt(6)
        
        # 保存到临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            doc.save(tmp.name)
            return tmp.name
            
    except Exception as e:
        st.error(f"生成Word文档失败: {str(e)}")
        return None

# 主界面布局
col1, col2 = st.columns([1, 2], gap="large")

with col1:
    st.subheader("目标政策文件")
    st.caption("作为基准的政策文件")
    target_file = st.file_uploader("上传目标政策文件 (PDF)", type="pdf", key="target")
    
    if target_file:
        # 使用当前设置解析目标文件
        st.session_state.target_clauses = parse_pdf(
            target_file, 
            max_clauses=st.session_state.max_clauses,
            precision=clause_precision,
            method=st.session_state.parse_method
        )
        st.success(f"✅ 解析完成，提取到 {len(st.session_state.target_clauses)} 条条款")
        
        with st.expander(f"查看提取的条款 (共 {len(st.session_state.target_clauses)} 条)"):
            for i, clause in enumerate(st.session_state.target_clauses):
                display_text = clause[:150] + "..." if len(clause) > 150 else clause
                st.markdown(f'<div class="clause-item"><strong>条款 {i+1}:</strong> {display_text}</div>', unsafe_allow_html=True)
    
    # 多文件上传区域
    st.subheader("待比对文件")
    st.caption("可上传多个文件，将逐一与目标文件比对")
    compare_files = st.file_uploader(
        "上传待比对文件 (PDF)", 
        type="pdf", 
        key="compare",
        accept_multiple_files=True
    )
    
    # 处理上传的多个文件
    if compare_files:
        for file in compare_files:
            if file.name not in st.session_state.compare_files:
                # 使用当前设置解析待比对文件
                clauses = parse_pdf(
                    file, 
                    max_clauses=st.session_state.max_clauses,
                    precision=clause_precision,
                    method=st.session_state.parse_method
                )
                st.session_state.compare_files[file.name] = {
                    "clauses": clauses,
                    "analysis": None
                }
                st.success(f"✅ 已添加 {file.name}，提取到 {len(clauses)} 条条款")
    
    # 显示已上传的待比对文件列表
    if st.session_state.compare_files:
        st.subheader("已上传文件")
        for filename in st.session_state.compare_files.keys():
            col_a, col_b = st.columns([3, 1])
            with col_a:
                st.markdown(f"- {filename} (条款数: {len(st.session_state.compare_files[filename]['clauses'])})")
            with col_b:
                if st.button("分析", key=f"analyze_{filename}") and st.session_state.target_clauses:
                    result = analyze_compliance(
                        st.session_state.target_clauses,
                        st.session_state.compare_files[filename]["clauses"],
                        st.session_state.api_key,
                        model_option
                    )
                    if result:
                        st.session_state.compare_files[filename]["analysis"] = result
                        st.session_state.current_file = filename
                        st.success(f"✅ {filename} 分析完成")

with col2:
    st.subheader("分析结果")
    
    # 显示文件选择标签
    if st.session_state.compare_files:
        st.markdown("**选择文件查看结果：**")
        # 计算每行显示的文件标签数量
        cols_per_row = 3
        files = list(st.session_state.compare_files.items())
        rows = (len(files) + cols_per_row - 1) // cols_per_row
        
        for row in range(rows):
            cols = st.columns(cols_per_row)
            for col_idx in range(cols_per_row):
                file_idx = row * cols_per_row + col_idx
                if file_idx < len(files):
                    filename, data = files[file_idx]
                    with cols[col_idx]:
                        status = " ✓" if data["analysis"] else ""
                        if st.button(f"{filename.split('.')[0]}{status}", key=f"tab_{filename}"):
                            st.session_state.current_file = filename
    
    # 显示当前选中文件的分析结果
    if st.session_state.current_file and st.session_state.compare_files[st.session_state.current_file]["analysis"]:
        filename = st.session_state.current_file
        analysis_result = st.session_state.compare_files[filename]["analysis"]
        
        st.markdown('<div class="analysis-box">', unsafe_allow_html=True)
        for para in re.split(r'\n+', analysis_result):
            if para.strip():
                st.markdown(f"{para.strip()}  \n")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 生成并下载Word文档
        if target_file:
            word_file = generate_word_document(
                analysis_result,
                target_file.name,
                filename
            )
            
            if word_file:
                with open(word_file, "rb") as f:
                    st.download_button(
                        label=f"💾 下载 {filename} 的分析报告",
                        data=f,
                        file_name=f"政策合规性分析报告_{filename}_{time.strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                os.unlink(word_file)
    elif st.session_state.compare_files:
        st.info("请选择一个文件进行分析，或点击文件旁的'分析'按钮")
    else:
        st.info("请上传待比对文件")

# 帮助信息
with st.expander("ℹ️ 使用帮助"):
    st.markdown("""
    ### 提高解析完整性的技巧
    1. **尝试不同解析方法**：
       - 智能识别：自动识别多种条款格式（默认）
       - 按标题层级：优先识别章节、条款等标题结构
       - 按段落拆分：简单按标点符号拆分文本
    
    2. **调整精细度**：
       - 复杂文件建议使用"精细"模式
       - 结构简单的文件可使用"粗略"模式提高效率
    
    3. **其他建议**：
       - 确保PDF文件可复制（非图片扫描件）
       - 若文件加密，请先解密再上传
       - 对于特别长的文件，可适当增加最大条款数量
    
    ### 基本使用流程
    1. 上传目标政策文件和待比对文件
    2. 配置API密钥（首次使用）
    3. 根据文件特点调整解析参数
    4. 点击"分析"按钮生成比对结果
    5. 查看结果并下载报告
    """)
