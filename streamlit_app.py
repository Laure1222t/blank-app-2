import streamlit as st
from PyPDF2 import PdfReader
from difflib import SequenceMatcher
import base64
import re
import requests
import jieba
import time
from io import StringIO
from typing import List, Tuple, Optional

# 页面设置
st.set_page_config(
    page_title="合规性分析工具",
    page_icon="📊",
    layout="wide"
)

# 自定义样式
st.markdown("""
<style>
    .stApp { max-width: 1200px; margin: 0 auto; }
    .status-box { padding: 10px; border-radius: 5px; margin: 10px 0; }
    .disabled-hint { color: #666; font-style: italic; }
</style>
""", unsafe_allow_html=True)

# API配置
QWEN_API_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"

# 会话状态初始化
if 'analysis_running' not in st.session_state:
    st.session_state.analysis_running = False
if 'button_disabled' not in st.session_state:
    st.session_state.button_disabled = True
if 'disabled_reason' not in st.session_state:
    st.session_state.disabled_reason = "请完成所有必要设置"

def update_button_state(base_file, target_files, api_key):
    """更新按钮状态和禁用原因"""
    if st.session_state.analysis_running:
        st.session_state.button_disabled = True
        st.session_state.disabled_reason = "分析正在进行中"
    elif not base_file:
        st.session_state.button_disabled = True
        st.session_state.disabled_reason = "请上传基准文件"
    elif not target_files:
        st.session_state.button_disabled = True
        st.session_state.disabled_reason = "请上传至少一个目标文件"
    elif not api_key or api_key.strip() == "":
        st.session_state.button_disabled = True
        st.session_state.disabled_reason = "请输入API密钥"
    else:
        st.session_state.button_disabled = False
        st.session_state.disabled_reason = ""

# 简化的核心函数（保持功能但精简代码）
def call_qwen_api(prompt: str, api_key: str) -> Optional[str]:
    try:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
        data = {
            "model": "qwen-plus",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
            "max_tokens": 1000
        }
        
        response = requests.post(QWEN_API_URL, headers=headers, json=data, timeout=30)
        if response.status_code == 200:
            response_json = response.json()
            if "choices" in response_json and len(response_json["choices"]) > 0:
                return response_json["choices"][0]["message"]["content"]
        return None
    except:
        return None

def extract_text_from_pdf(file) -> str:
    try:
        pdf_reader = PdfReader(file)
        text = ""
        for page in pdf_reader.pages[:20]:  # 限制页数
            page_text = page.extract_text() or ""
            text += page_text.replace("  ", "").replace("\n", "")
        return text[:50000]  # 限制文本长度
    except:
        return ""

def split_into_clauses(text: str, max_clauses: int = 20) -> List[str]:
    patterns = [
        r'(第[一二三四五六七八九十百]+条\s+.*?)(?=第[一二三四五六七八九十百]+条\s+|$)',
        r'(\d+\.\s+.*?)(?=\d+\.\s+|$)',
    ]
    for pattern in patterns:
        clauses = re.findall(pattern, text, re.DOTALL)
        if len(clauses) > 3:
            return [clause.strip() for clause in clauses if clause.strip()][:max_clauses]
    paragraphs = re.split(r'[。；！？]\s*', text)
    return [p.strip() for p in paragraphs if p.strip() and len(p) > 10][:max_clauses]

def match_clauses_with_base(base_clauses, target_clauses) -> List[Tuple[str, str, float]]:
    matched_pairs = []
    used_indices = set()
    for base_clause in base_clauses[:15]:
        best_match = None
        best_ratio = 0.35
        best_idx = -1
        for idx, target_clause in enumerate(target_clauses[:20]):
            if idx not in used_indices:
                ratio = SequenceMatcher(None, base_clause, target_clause).ratio()
                if ratio > best_ratio:
                    best_ratio = ratio
                    best_match = target_clause
                    best_idx = idx
        if best_match:
            matched_pairs.append((base_clause, best_match, best_ratio))
            used_indices.add(best_idx)
    return matched_pairs

def generate_target_report(matched_pairs, base_name, target_name, api_key) -> str:
    report = [f"合规性分析报告: {target_name} vs {base_name}\n{'-'*50}\n"]
    for i, (base_clause, target_clause, ratio) in enumerate(matched_pairs):
        report.append(f"条款对 {i+1} (相似度: {ratio:.2%})")
        report.append(f"基准条款: {base_clause[:150]}...")
        report.append(f"目标条款: {target_clause[:150]}...\n")
        
        prompt = f"分析基准条款: {base_clause[:300]} 与目标条款: {target_clause[:300]} 的合规性，简要说明符合程度、差异和建议。"
        analysis = call_qwen_api(prompt, api_key)
        report.append(f"分析: {analysis if analysis else '无法获取分析结果'}\n{'-'*50}\n")
    return "\n".join(report)

def get_download_link(text: str, filename: str) -> str:
    b64 = base64.b64encode(text.encode()).decode()
    return f'<a href="data:text/plain;base64,{b64}" download="{filename}" style="padding:8px 16px;background:#007bff;color:white;text-decoration:none;border-radius:4px;margin:5px 0;display:inline-block;">下载报告</a>'

def main():
    st.title("合规性分析工具")
    st.write("基准文件与多目标文件条款对比")
    
    # 侧边栏设置
    with st.sidebar:
        st.subheader("API设置")
        api_key = st.text_input("Qwen API密钥", type="password", key="api_key")
        max_clauses = st.slider("最大条款数/文件", 5, 30, 10)
    
    # 文件上传
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("基准文件")
        base_file = st.file_uploader("上传基准PDF", type="pdf", key="base_file")
    
    with col2:
        st.subheader("目标文件")
        target_files = st.file_uploader(
            "上传目标PDF（可多个）", 
            type="pdf", 
            key="target_files",
            accept_multiple_files=True
        )
    
    # 更新按钮状态
    update_button_state(base_file, target_files, api_key)
    
    # 显示按钮禁用原因（如果有）
    if st.session_state.button_disabled and st.session_state.disabled_reason:
        st.markdown(f'<p class="disabled-hint">🔒 开始分析按钮已禁用: {st.session_state.disabled_reason}</p>', unsafe_allow_html=True)
    
    # 分析按钮
    if st.button("开始分析", disabled=st.session_state.button_disabled):
        st.session_state.analysis_running = True
        
        try:
            # 处理基准文件
            with st.spinner("处理基准文件..."):
                base_text = extract_text_from_pdf(base_file)
                if not base_text:
                    st.error("无法从基准文件提取文本")
                    st.session_state.analysis_running = False
                    return
                base_clauses = split_into_clauses(base_text, max_clauses)
                st.success(f"基准文件处理完成，提取到 {len(base_clauses)} 条条款")
            
            # 处理目标文件
            for idx, target_file in enumerate(target_files, 1):
                st.subheader(f"分析目标文件 {idx}/{len(target_files)}: {target_file.name}")
                target_text = extract_text_from_pdf(target_file)
                if not target_text:
                    st.warning("无法提取文件内容，跳过")
                    continue
                
                target_clauses = split_into_clauses(target_text, max_clauses)
                matched_pairs = match_clauses_with_base(base_clauses, target_clauses)
                
                if not matched_pairs:
                    st.warning("未找到匹配条款，无法分析")
                    continue
                
                report = generate_target_report(matched_pairs, base_file.name, target_file.name, api_key)
                st.markdown(get_download_link(report, f"{target_file.name}_合规性报告.txt"), unsafe_allow_html=True)
                with st.expander("查看报告预览"):
                    st.text_area("报告内容", report, height=200)
            
            st.session_state.analysis_running = False
            st.success("所有文件分析完成！")
                
        except Exception as e:
            st.error(f"分析出错: {str(e)}")
            st.session_state.analysis_running = False

if __name__ == "__main__":
    main()
