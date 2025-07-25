import streamlit as st
import fitz  # PyMuPDF
import re
import time
import torch
from transformers import AutoTokenizer, AutoModelForCausalLM, BitsAndBytesConfig
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os

# 设置页面配置 - 优化移动端显示
st.set_page_config(
    page_title="政策文件比对分析工具",
    page_icon="📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义CSS - 优化显示效果
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
    .section-header {
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
        font-weight: bold;
        color: #2c3e50;
    }
</style>
""", unsafe_allow_html=True)

# 初始化会话状态 - 避免重复加载
if 'target_clauses' not in st.session_state:
    st.session_state.target_clauses = []
if 'compare_clauses' not in st.session_state:
    st.session_state.compare_clauses = []
if 'analysis_result' not in st.session_state:
    st.session_state.analysis_result = None
if 'model_loaded' not in st.session_state:
    st.session_state.model_loaded = False
if 'tokenizer' not in st.session_state:
    st.session_state.tokenizer = None
if 'model' not in st.session_state:
    st.session_state.model = None

# 页面标题和说明
st.title("📜 中文政策文件比对分析工具")
st.markdown("上传目标政策文件和待比对文件，系统将自动解析并进行条款比对与合规性分析")
st.markdown("---")

# 检查PyTorch是否可用
try:
    import torch
    torch_available = True
except ImportError:
    torch_available = False
    st.error("⚠️ 未检测到PyTorch库，请检查依赖配置。")

# 优化的PDF解析函数 - 更准确的条款提取
def parse_pdf(file):
    """解析PDF文件并提取结构化条款"""
    try:
        # 读取PDF内容
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
            
            return clauses[:30]  # 限制最大条款数量，避免内存问题
            
    except Exception as e:
        st.error(f"文件解析错误: {str(e)}")
        return []

# 模型加载优化 - 使用量化技术减少内存占用
@st.cache_resource(show_spinner=False)
def load_optimized_model():
    """加载量化后的Qwen模型，适合云环境运行"""
    if not torch_available:
        return None, None
        
    try:
        # 4位量化配置 - 大幅减少内存使用
        bnb_config = BitsAndBytesConfig(
            load_in_4bit=True,
            bnb_4bit_use_double_quant=True,
            bnb_4bit_quant_type="nf4",
            bnb_4bit_compute_dtype=torch.float16
        )
        
        # 使用较小的模型版本，适合云环境
        model_name = "Qwen/Qwen-1.8B-Chat"
        
        with st.spinner(f"正在加载模型 {model_name}...\n这可能需要几分钟时间"):
            tokenizer = AutoTokenizer.from_pretrained(
                model_name, 
                trust_remote_code=True,
                cache_dir="./cache"
            )
            
            model = AutoModelForCausalLM.from_pretrained(
                model_name,
                quantization_config=bnb_config,
                device_map="auto",
                trust_remote_code=True,
                cache_dir="./cache"
            )
            model.eval()
            
            return tokenizer, model
            
    except Exception as e:
        st.error(f"模型加载失败: {str(e)}")
        st.info("提示：如果持续加载失败，可能是资源限制，请尝试使用更小的模型或本地部署。")
        return None, None

# 合规性分析函数优化 - 更明确的提示词和批处理
def analyze_compliance(target_clauses, compare_clauses):
    """使用优化的提示词进行合规性分析"""
    if not st.session_state.tokenizer or not st.session_state.model:
        return "模型未加载，无法进行分析。"
    
    try:
        with st.spinner("正在进行条款比对和合规性分析..."):
            # 准备条款文本 - 限制长度以适应模型
            target_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(target_clauses[:15])])
            compare_text = "\n".join([f"条款{i+1}: {clause[:200]}" for i, clause in enumerate(compare_clauses[:15])])
            
            # 优化的提示词 - 更明确的指令
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
            
            # 模型推理参数优化
            inputs = st.session_state.tokenizer(prompt, return_tensors="pt").to(st.session_state.model.device)
            
            with torch.no_grad():  # 禁用梯度计算，节省内存
                outputs = st.session_state.model.generate(
                    **inputs,
                    max_new_tokens=1200,  # 限制输出长度，避免超时
                    temperature=0.6,      # 降低随机性，提高稳定性
                    top_p=0.9,
                    repetition_penalty=1.1  # 减少重复内容
                )
            
            result = st.session_state.tokenizer.decode(outputs[0], skip_special_tokens=True)
            
            # 提取有效结果（去除提示词部分）
            result_start = result.find("目标政策文件条款：")
            if result_start != -1:
                result = result[result_start:]
                
            return result
            
    except Exception as e:
        st.error(f"分析过程出错: {str(e)}")
        return f"分析失败: {str(e)}"

# 生成Word文档函数优化
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
                # 识别标题行
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

# 模型加载和分析控制
st.markdown("---")

# 单独的模型加载按钮，避免重复加载
if not st.session_state.model_loaded and torch_available:
    if st.button("📦 加载分析模型 (首次使用需要几分钟)"):
        st.session_state.tokenizer, st.session_state.model = load_optimized_model()
        if st.session_state.tokenizer and st.session_state.model:
            st.session_state.model_loaded = True
            st.success("模型加载成功，可以开始分析了！")
        else:
            st.session_state.model_loaded = False

# 分析按钮
if st.session_state.model_loaded and st.session_state.target_clauses and st.session_state.compare_clauses:
    if st.button("🔍 开始比对与合规性分析"):
        with st.spinner("正在进行深度分析，请稍候..."):
            st.session_state.analysis_result = analyze_compliance(
                st.session_state.target_clauses, 
                st.session_state.compare_clauses
            )

# 显示分析结果
if st.session_state.analysis_result:
    st.markdown("### 📊 合规性分析结果")
    st.markdown('<div class="analysis-box">', unsafe_allow_html=True)
    # 将分析结果按段落显示，增强可读性
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
            # 清理临时文件
            os.unlink(word_file)

# 帮助信息
with st.expander("ℹ️ 使用帮助"):
    st.markdown("""
    1. 首先上传目标政策文件（左侧）和待比对文件（右侧）
    2. 点击"加载分析模型"按钮（首次使用需要几分钟）
    3. 模型加载完成后，点击"开始比对与合规性分析"
    4. 分析完成后可以查看结果并下载Word报告
    
    注意：
    - 模型加载需要一定时间和资源，请耐心等待
    - 为保证分析效果，建议上传清晰的PDF文件
    - 分析结果仅供参考，重要决策请咨询专业人士
    """)
