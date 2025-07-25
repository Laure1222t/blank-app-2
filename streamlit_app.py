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

# 设置页面配置
st.set_page_config(
    page_title="政策文件比对分析工具",
    page_icon="📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义CSS
st.markdown("""
<style>
    .stButton>button {
        width: 100%;
        margin-top: 1rem;
    }
    .analysis-box {
        border: 1px solid #e0e0e0;
        border-radius: 5px;
        padding: 1rem;
        margin-top: 1rem;
    }
    .api-key-warning {
        color: #e74c3c;
        padding: 10px;
        border-radius: 5px;
        background-color: #fdf2f2;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# 初始化会话状态
if 'target_clauses' not in st.session_state:
    st.session_state.target_clauses = []
if 'compare_clauses' not in st.session_state:
    st.session_state.compare_clauses = []
if 'analysis_result' not in st.session_state:
    st.session_state.analysis_result = None
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""

# 页面标题和说明
st.title("📜 中文政策文件比对分析工具")
st.markdown("上传目标政策文件和待比对文件，系统将自动解析并进行条款比对与合规性分析")
st.markdown("---")

# API设置
with st.expander("🔑 API 设置", expanded=False):
    st.session_state.api_key = st.text_input(
        "请输入你的Qwen API密钥", 
        value=st.session_state.api_key,
        type="password"
    )
    api_endpoint = st.text_input(
        "API 端点", 
        value="https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        help="Qwen API的访问端点，默认为阿里云DashScope"
    )
    model_version = st.selectbox(
        "选择模型版本",
        ["qwen-turbo", "qwen-plus", "qwen1.5-7b-chat"],
        index=0,
        help="qwen-turbo为轻量版，响应速度快且成本低"
    )

# 检查API密钥
if not st.session_state.api_key:
    st.markdown('<div class="api-key-warning">⚠️ 请先输入API密钥以使用分析功能</div>', unsafe_allow_html=True)

# PDF解析函数
def parse_pdf(file):
    """解析PDF文件并提取结构化条款"""
    try:
        with st.spinner("正在解析文件..."):
            doc = fitz.open(stream=file.read(), filetype="pdf")
            text = ""
            for page in doc:
                text += page.get_text()
            
            # 清理文本
            text = re.sub(r'\s+', ' ', text).strip()
            
            # 条款提取策略：优先识别多级编号条款
            clause_patterns = [
                re.compile(r'(\d+\.\s+.*?)(?=\d+\.\s+|$)', re.DOTALL),  # 一级条款 (1. ...)
                re.compile(r'(\d+\.\d+\s+.*?)(?=\d+\.\d+\s+|\d+\.\s+|$)', re.DOTALL),  # 二级条款 (1.1 ...)
                re.compile(r'(\d+\.\d+\.\d+\s+.*?)(?=\d+\.\d+\.\d+\s+|\d+\.\d+\s+|$)', re.DOTALL)  # 三级条款
            ]
            
            clauses = []
            for pattern in clause_patterns:
                matches = pattern.findall(text)
                if matches:
                    clauses = [match.strip() for match in matches if len(match.strip()) > 20]  # 过滤过短条目
                    break
            
            # 如果没有识别到条款格式，按段落分割
            if not clauses:
                paragraphs = [p.strip() for p in text.split('\n') if len(p.strip()) > 50]  # 过滤过短段落
                clauses = paragraphs
            
            return clauses[:30]  # 限制最大条款数量
            
    except Exception as e:
        st.error(f"文件解析错误: {str(e)}")
        return []

# 通过API调用Qwen模型进行合规性分析
def analyze_compliance_api(target_clauses, compare_clauses, api_key, endpoint, model):
    """使用API调用Qwen模型进行合规性分析"""
    if not api_key:
        return "请先设置API密钥"
    
    try:
        with st.spinner("正在进行条款比对和合规性分析..."):
            # 准备条款文本
            target_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(target_clauses[:15])])
            compare_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(compare_clauses[:15])])
            
            # 构建提示词
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
            
            # 构建API请求
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
            
            data = {
                "model": model,
                "messages": [
                    {"role": "system", "content": "你是专业的政策合规性分析专家，擅长比对政策文件条款并分析合规性。"},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.6,
                "max_tokens": 1200
            }
            
            # 发送请求
            response = requests.post(endpoint, headers=headers, data=json.dumps(data))
            response_data = response.json()
            
            # 处理响应
            if response.status_code == 200 and "choices" in response_data:
                return response_data["choices"][0]["message"]["content"]
            else:
                error_msg = response_data.get("error", {}).get("message", "API调用失败")
                return f"分析失败: {error_msg} (状态码: {response.status_code})"
                
    except Exception as e:
        st.error(f"分析过程出错: {str(e)}")
        return f"分析失败: {str(e)}"

# 生成Word文档函数
def generate_word_document(analysis_result, target_filename, compare_filename):
    """生成格式化的Word分析报告"""
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
col1, col2 = st.columns(2, gap="large")

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

with col2:
    st.subheader("待比对文件")
    st.caption("需要检查合规性的文件")
    compare_file = st.file_uploader("上传待比对文件 (PDF)", type="pdf", key="compare")
    
    if compare_file:
        st.session_state.compare_clauses = parse_pdf(compare_file)
        st.success(f"✅ 解析完成，提取到 {len(st.session_state.compare_clauses)} 条条款")
        
        with st.expander(f"查看提取的条款 (显示前10条)"):
            for i, clause in enumerate(st.session_state.compare_clauses[:10]):
                st.markdown(f"**条款 {i+1}:** {clause[:150]}..." if len(clause) > 150 else f"**条款 {i+1}:** {clause}")

# 分析控制
st.markdown("---")

# 分析按钮
if st.session_state.api_key and st.session_state.target_clauses and st.session_state.compare_clauses:
    if st.button("🔍 开始比对与合规性分析"):
        with st.spinner("正在进行深度分析，请稍候..."):
            st.session_state.analysis_result = analyze_compliance_api(
                st.session_state.target_clauses, 
                st.session_state.compare_clauses,
                st.session_state.api_key,
                api_endpoint,
                model_version
            )

# 显示分析结果
if st.session_state.analysis_result:
    st.markdown("### 📊 合规性分析结果")
    st.markdown('<div class="analysis-box">', unsafe_allow_html=True)
    for para in re.split(r'\n+', st.session_state.analysis_result):
        if para.strip():
            st.markdown(f"{para.strip()}  \n")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # 生成并下载Word文档
    if target_file and compare_file:
        word_file = generate_word_document(
            st.session_state.analysis_result,
            target_file.name,
            compare_file.name
        )
        
        if word_file:
            with open(word_file, "rb") as f:
                st.download_button(
                    label="💾 下载分析报告 (Word格式)",
                    data=f,
                    file_name=f"政策合规性分析报告_{time.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            os.unlink(word_file)

# 帮助信息
with st.expander("ℹ️ 使用帮助"):
    st.markdown("""
    1. 首先在API设置中输入你的Qwen API密钥
    2. 上传目标政策文件（左侧）和待比对文件（右侧）
    3. 点击"开始比对与合规性分析"按钮
    4. 分析完成后可以查看结果并下载Word报告
    
    API获取提示：
    - Qwen API密钥可从阿里云DashScope平台获取
    - 推荐使用qwen-turbo轻量模型，响应速度快且成本低
    - 请注意API调用可能产生费用，请参考相关平台的定价政策
    """)
    
