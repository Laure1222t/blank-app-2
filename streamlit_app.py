import streamlit as st
import fitz  # PyMuPDF
import re
import json
import torch
from transformers import AutoTokenizer, AutoModelForCausalLM
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import tempfile
import time

# 设置页面配置
st.set_page_config(
    page_title="政策文件比对分析工具",
    page_icon="📜",
    layout="wide"
)

# 页面标题
st.title("📜 中文政策文件比对分析工具")
st.markdown("上传目标政策文件和待比对文件，系统将自动解析并进行条款比对与合规性分析")

# 初始化会话状态
if 'target_doc' not in st.session_state:
    st.session_state.target_doc = None
if 'compare_doc' not in st.session_state:
    st.session_state.compare_doc = None
if 'analysis_result' not in st.session_state:
    st.session_state.analysis_result = None

# 加载Qwen模型和tokenizer
@st.cache_resource
def load_model():
    try:
        with st.spinner("正在加载Qwen大模型，请稍候..."):
            tokenizer = AutoTokenizer.from_pretrained("Qwen/Qwen-7B-Chat", trust_remote_code=True)
            model = AutoModelForCausalLM.from_pretrained(
                "Qwen/Qwen-7B-Chat", 
                device_map="auto", 
                trust_remote_code=True
            )
            model.eval()
            return tokenizer, model
    except Exception as e:
        st.error(f"模型加载失败: {str(e)}")
        return None, None

# 解析PDF文件
def parse_pdf(file):
    try:
        doc = fitz.open(stream=file.read(), filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()
        
        # 简单的条款提取，可根据实际需求优化
        clauses = []
        # 匹配以数字加点开头的条款（如 1. 2.1 等）
        pattern = re.compile(r'(\d+\.\s+|\d+\.\d+\s+).+?(?=\d+\.\s+|\d+\.\d+\s+|$)', re.DOTALL)
        matches = pattern.findall(text)
        
        if matches:
            for match in matches:
                clause_text = match[0] + match[1].strip()
                clauses.append(clause_text)
        else:
            # 如果没有明确的条款格式，按段落分割
            paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
            clauses = paragraphs[:20]  # 取前20个段落
        
        return clauses
    except Exception as e:
        st.error(f"PDF解析失败: {str(e)}")
        return []

# 使用Qwen模型进行合规性分析
def analyze_compliance(target_clauses, compare_clauses, tokenizer, model):
    if not tokenizer or not model:
        return "模型加载失败，无法进行分析"
    
    try:
        with st.spinner("正在进行条款比对和合规性分析，请稍候..."):
            # 构建分析提示
            prompt = """
            请比对以下两份政策文件的条款，进行合规性分析。
            分析要求：
            1. 尽量覆盖所有条款，确保分析全面
            2. 重点进行合规性分析，判断不同之处是否存在冲突，而不仅仅是指出差异
            3. 对于相同或一致的条款可以简要说明
            4. 对于存在差异的条款，详细分析是否存在合规性冲突，以及可能的影响
            
            目标政策文件条款：
            {}
            
            待比对文件条款：
            {}
            
            请以结构化的方式输出分析结果，包括条款对应关系、差异点和合规性分析。
            """.format("\n".join(target_clauses[:10]), "\n".join(compare_clauses[:10]))
            
            # 调用模型
            inputs = tokenizer(prompt, return_tensors="pt").to(model.device)
            outputs = model.generate(
                **inputs,
                max_new_tokens=1500,
                temperature=0.7,
                top_p=0.9
            )
            result = tokenizer.decode(outputs[0], skip_special_tokens=True)
            
            # 提取模型回答（去除提示部分）
            result_start = result.find("目标政策文件条款：")
            if result_start != -1:
                result = result[result_start:]
                
            return result
    except Exception as e:
        st.error(f"分析过程出错: {str(e)}")
        return f"分析失败: {str(e)}"

# 生成Word文档
def generate_word_document(analysis_result, target_filename, compare_filename):
    try:
        doc = docx.Document()
        
        # 添加标题
        title = doc.add_heading("政策文件合规性分析报告", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 添加文件信息
        doc.add_paragraph(f"目标政策文件: {target_filename}")
        doc.add_paragraph(f"待比对文件: {compare_filename}")
        doc.add_paragraph(f"分析日期: {time.strftime('%Y年%m月%d日')}")
        doc.add_paragraph("")
        
        # 添加分析结果
        doc.add_heading("分析结果", level=1)
        
        # 简单处理分析结果，按换行分割
        paragraphs = analysis_result.split('\n')
        for para in paragraphs:
            if para.strip():
                p = doc.add_paragraph(para.strip())
                p.paragraph_format.space_after = Pt(12)
        
        # 保存到临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            doc.save(tmp.name)
            return tmp.name
    except Exception as e:
        st.error(f"生成Word文档失败: {str(e)}")
        return None

# 主界面布局
col1, col2 = st.columns(2)

with col1:
    st.subheader("目标政策文件 (左侧)")
    target_file = st.file_uploader("上传目标政策文件 (PDF)", type="pdf", key="target")
    
    if target_file:
        st.session_state.target_doc = parse_pdf(target_file)
        st.success(f"文件解析成功，提取到 {len(st.session_state.target_doc)} 条条款")
        
        with st.expander("查看提取的条款"):
            for i, clause in enumerate(st.session_state.target_doc[:10]):  # 只显示前10条
                st.write(f"条款 {i+1}: {clause[:100]}...")

with col2:
    st.subheader("待比对文件 (右侧)")
    compare_file = st.file_uploader("上传待比对文件 (PDF)", type="pdf", key="compare")
    
    if compare_file:
        st.session_state.compare_doc = parse_pdf(compare_file)
        st.success(f"文件解析成功，提取到 {len(st.session_state.compare_doc)} 条条款")
        
        with st.expander("查看提取的条款"):
            for i, clause in enumerate(st.session_state.compare_doc[:10]):  # 只显示前10条
                st.write(f"条款 {i+1}: {clause[:100]}...")

# 分析按钮
if st.button("开始比对与合规性分析"):
    if not st.session_state.target_doc or not st.session_state.compare_doc:
        st.warning("请先上传并解析两份文件")
    else:
        # 加载模型
        tokenizer, model = load_model()
        if tokenizer and model:
            # 进行分析
            st.session_state.analysis_result = analyze_compliance(
                st.session_state.target_doc, 
                st.session_state.compare_doc,
                tokenizer,
                model
            )

# 显示分析结果
if st.session_state.analysis_result:
    st.subheader("📊 合规性分析结果")
    st.text_area("", st.session_state.analysis_result, height=400)
    
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
                    label="下载分析报告 (Word)",
                    data=f,
                    file_name=f"政策文件合规性分析报告_{time.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            # 清理临时文件
            os.unlink(word_file)

# 页脚信息
st.markdown("---")
st.markdown("工具说明：本工具用于政策文件的条款比对与合规性分析，结果仅供参考。")
