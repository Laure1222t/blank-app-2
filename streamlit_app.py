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
    }
    .file-tab.active {
        background-color: #007bff;
        color: white;
    }
    .file-tab.inactive {
        background-color: #e9ecef;
        color: #495057;
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

# 页面标题
st.title("📜 多文件政策比对分析工具")
st.markdown("上传目标政策文件和多个待比对文件，系统将逐一进行条款比对与合规性分析")
st.markdown("---")

# API配置
with st.expander("🔑 API 配置", expanded=not st.session_state.api_key):
    st.session_state.api_key = st.text_input("请输入Qwen API密钥", value=st.session_state.api_key, type="password")
    model_option = st.selectbox(
        "选择Qwen模型",
        ["qwen-turbo", "qwen-plus", "qwen-max"],
        index=0  # 默认使用轻量版
    )
    st.caption("提示：可从阿里云DashScope平台获取API密钥，不同模型能力和成本不同")

# PDF解析函数
def parse_pdf(file):
    """解析PDF文件并提取条款"""
    try:
        with st.spinner("正在解析文件..."):
            doc = fitz.open(stream=file.read(), filetype="pdf")
            text = ""
            for page in doc:
                text += page.get_text()
            
            # 清理文本
            text = re.sub(r'\s+', ' ', text).strip()
            
            # 条款提取
            clause_patterns = [
                re.compile(r'(\d+\.\s+.*?)(?=\d+\.\s+|$)', re.DOTALL),
                re.compile(r'(\d+\.\d+\s+.*?)(?=\d+\.\d+\s+|\d+\.\s+|$)', re.DOTALL),
                re.compile(r'(\d+\.\d+\.\d+\s+.*?)(?=\d+\.\d+\.\d+\s+|\d+\.\d+\s+|$)', re.DOTALL)
            ]
            
            clauses = []
            for pattern in clause_patterns:
                matches = pattern.findall(text)
                if matches:
                    clauses = [match.strip() for match in matches if len(match.strip()) > 20]
                    break
            
            if not clauses:
                paragraphs = [p.strip() for p in text.split('\n') if len(p.strip()) > 50]
                clauses = paragraphs
            
            return clauses[:30]
            
    except Exception as e:
        st.error(f"文件解析错误: {str(e)}")
        return []

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
    target_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(target_clauses[:15])])
    compare_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(compare_clauses[:15])])
    
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
        st.session_state.target_clauses = parse_pdf(target_file)
        st.success(f"✅ 解析完成，提取到 {len(st.session_state.target_clauses)} 条条款")
        
        with st.expander(f"查看提取的条款 (显示前10条)"):
            for i, clause in enumerate(st.session_state.target_clauses[:10]):
                st.markdown(f"**条款 {i+1}:** {clause[:150]}..." if len(clause) > 150 else f"**条款 {i+1}:** {clause}")
    
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
                clauses = parse_pdf(file)
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
                st.markdown(f"- {filename}")
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
        cols = st.columns(min(3, len(st.session_state.compare_files)))
        for i, (filename, data) in enumerate(st.session_state.compare_files.items()):
            with cols[i % min(3, len(st.session_state.compare_files))]:
                status = "✓" if data["analysis"] else ""
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
    1. 首先上传目标政策文件（左侧）
    2. 配置Qwen API密钥（首次使用需要）
    3. 上传一个或多个待比对文件（左侧）
    4. 对每个待比对文件点击"分析"按钮
    5. 在右侧查看不同文件的分析结果并下载报告
    
    注意：
    - API调用可能产生费用，请参考阿里云DashScope平台定价
    - 分析结果取决于文件质量和条款清晰度
    - 可同时上传多个文件，逐一进行分析和查看
    """)
