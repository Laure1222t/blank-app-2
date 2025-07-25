import streamlit as st
from PyPDF2 import PdfReader
from difflib import SequenceMatcher
import base64
import re
import requests
import jieba
import time
import sys
from io import StringIO
from typing import List, Tuple, Optional, Dict

# 页面设置 - 尽可能早地设置
st.set_page_config(
    page_title="高效合规性分析工具",
    page_icon="📊",
    layout="wide"
)

# 强制刷新缓存
for key in list(st.session_state.keys()):
    if key not in ['api_key', 'analysis_running']:
        del st.session_state[key]

# 自定义样式
st.markdown("""
<style>
    .stApp { max-width: 1200px; margin: 0 auto; }
    .status-box { padding: 10px; border-radius: 5px; margin: 10px 0; }
    .loading-box { background-color: #e3f2fd; }
    .error-box { background-color: #ffebee; }
    .success-box { background-color: #e8f5e9; }
</style>
""", unsafe_allow_html=True)

# API配置
QWEN_API_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"

# 会话状态初始化
if 'analysis_running' not in st.session_state:
    st.session_state.analysis_running = False
if 'last_error' not in st.session_state:
    st.session_state.last_error = ""
if 'partial_results' not in st.session_state:
    st.session_state.partial_results = {}

def limited_print(text: str, max_length: int = 500):
    """限制打印长度，避免内存占用过多"""
    if len(text) > max_length:
        print(f"{text[:max_length]}...")
    else:
        print(text)

def call_qwen_api(prompt: str, api_key: str) -> Optional[str]:
    """优化的API调用函数，减少超时和内存问题"""
    try:
        # 限制提示词长度
        if len(prompt) > 3000:
            prompt = prompt[:3000] + "...（提示词已截断以提高处理效率）"
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
        data = {
            "model": "qwen-plus",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2,
            "max_tokens": 1000  # 减少返回内容长度
        }
        
        # 使用较短的超时时间
        response = requests.post(
            QWEN_API_URL,
            headers=headers,
            json=data,
            timeout=30  # 缩短超时时间
        )
        
        if response.status_code == 200:
            response_json = response.json()
            if "choices" in response_json and len(response_json["choices"]) > 0:
                return response_json["choices"][0]["message"]["content"]
            
        st.warning(f"API返回非预期状态: {response.status_code}")
        return None
                
    except requests.exceptions.Timeout:
        st.session_state.last_error = "API请求超时"
        return None
    except requests.exceptions.RequestException as e:
        st.session_state.last_error = f"网络错误: {str(e)}"
        return None
    except Exception as e:
        st.session_state.last_error = f"API处理错误: {str(e)}"
        return None

def extract_text_from_pdf(file) -> str:
    """轻量级PDF文本提取"""
    try:
        # 限制PDF大小
        if file.size > 10 * 1024 * 1024:  # 10MB
            st.error("文件过大，建议使用10MB以下的PDF")
            return ""
            
        pdf_reader = PdfReader(file)
        # 限制处理页数
        max_pages = 20
        pages_to_process = pdf_reader.pages[:max_pages]
        
        text = ""
        for page in pages_to_process:
            page_text = page.extract_text() or ""
            page_text = page_text.replace("  ", "").replace("\n", "").replace("\r", "")
            text += page_text
            # 限制总文本长度
            if len(text) > 50000:
                text = text[:50000]
                break
                
        return text
    except Exception as e:
        st.session_state.last_error = f"PDF提取错误: {str(e)}"
        return ""

def split_into_clauses(text: str, max_clauses: int = 20) -> List[str]:
    """更高效的条款分割"""
    if not text:
        return []
        
    # 简化的条款识别模式
    patterns = [
        r'(第[一二三四五六七八九十百]+条\s+.*?)(?=第[一二三四五六七八九十百]+条\s+|$)',
        r'(\d+\.\s+.*?)(?=\d+\.\s+|$)',
    ]
    
    for pattern in patterns:
        clauses = re.findall(pattern, text, re.DOTALL)
        if len(clauses) > 3:
            return [clause.strip() for clause in clauses if clause.strip()][:max_clauses]
    
    # 备用分割方式
    paragraphs = re.split(r'[。；！？]\s*', text)
    return [p.strip() for p in paragraphs if p.strip() and len(p) > 10][:max_clauses]

def chinese_text_similarity(text1: str, text2: str) -> float:
    """优化的相似度计算，减少计算量"""
    # 限制文本长度以提高速度
    text1 = text1[:500]
    text2 = text2[:500]
    
    words1 = list(jieba.cut(text1))
    words2 = list(jieba.cut(text2))
    
    # 限制分词数量
    if len(words1) > 100:
        words1 = words1[:100]
    if len(words2) > 100:
        words2 = words2[:100]
        
    return SequenceMatcher(None, words1, words2).ratio()

def match_clauses_with_base(base_clauses: List[str], target_clauses: List[str]) -> List[Tuple[str, str, float]]:
    """更高效的条款匹配"""
    matched_pairs = []
    used_indices = set()
    
    # 限制匹配数量
    max_matches = min(15, len(base_clauses), len(target_clauses))
    
    for i, base_clause in enumerate(base_clauses[:max_matches]):
        best_match = None
        best_ratio = 0.35
        best_idx = -1
        
        # 只检查前20个条款，提高速度
        for idx, target_clause in enumerate(target_clauses[:20]):
            if idx not in used_indices:
                ratio = chinese_text_similarity(base_clause, target_clause)
                if ratio > best_ratio:
                    best_ratio = ratio
                    best_match = target_clause
                    best_idx = idx
        
        if best_match:
            matched_pairs.append((base_clause, best_match, best_ratio))
            used_indices.add(best_idx)
            if len(matched_pairs) >= max_matches:
                break
    
    return matched_pairs

def analyze_compliance_with_base(base_clause: str, target_clause: str, 
                               base_name: str, target_name: str, 
                               api_key: str) -> Optional[str]:
    """简化的合规性分析"""
    # 限制条款长度
    base_clause = base_clause[:300]
    target_clause = target_clause[:300]
    
    prompt = f"""
    以{base_name}为基准，分析合规性：
    
    基准条款：{base_clause}
    目标条款：{target_clause}
    
    请简要回答：
    1. 符合程度（高/中/低）
    2. 主要差异（1-2点）
    3. 修改建议
    """
    
    return call_qwen_api(prompt, api_key)

def generate_target_report(matched_pairs: List[Tuple[str, str, float]],
                          base_name: str, target_name: str,
                          api_key: str) -> str:
    """生成目标文件报告（优化性能）"""
    report = [
        f"条款合规性分析报告: {target_name} vs {base_name}",
        f"生成时间: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "="*50 + "\n"
    ]
    
    report.append(f"分析概要: 共匹配 {len(matched_pairs)} 条条款\n")
    
    # 分析每对条款
    for i, (base_clause, target_clause, ratio) in enumerate(matched_pairs):
        report.append(f"条款对 {i+1} (相似度: {ratio:.2%})")
        report.append(f"基准条款: {base_clause[:150]}...")
        report.append(f"目标条款: {target_clause[:150]}...\n")
        
        # 合规性分析
        analysis = analyze_compliance_with_base(
            base_clause, target_clause, 
            base_name, target_name, 
            api_key
        )
        
        if analysis:
            report.append("合规性分析:")
            report.append(analysis)
        else:
            report.append("合规性分析: 无法获取分析结果")
        
        report.append("\n" + "-"*50 + "\n")
        
        # 保存部分结果
        st.session_state.partial_results[target_name] = "\n".join(report)
        
        # 检查是否需要中断
        if st.session_state.last_error:
            report.append(f"\n分析中断: {st.session_state.last_error}")
            return "\n".join(report)
        
        time.sleep(0.5)  # 短暂延迟，减少并发
        
    return "\n".join(report)

def get_download_link(text: str, filename: str) -> str:
    """生成下载链接"""
    b64 = base64.b64encode(text.encode()).decode()
    return f'<a href="data:text/plain;base64,{b64}" download="{filename}" style="display:inline-block;padding:8px 16px;background-color:#007bff;color:white;text-decoration:none;border-radius:4px;margin:5px 0;">下载报告</a>'

def main():
    st.title("高效合规性分析工具")
    st.write("基准文件与多目标文件条款对比（优化版）")
    
    # 显示最后错误（如果有）
    if st.session_state.last_error:
        st.markdown(f'<div class="status-box error-box">⚠️ 上次错误: {st.session_state.last_error}</div>', unsafe_allow_html=True)
    
    # 侧边栏设置
    with st.sidebar:
        st.subheader("设置")
        api_key = st.text_input("Qwen API密钥", type="password", key="api_key")
        max_clauses = st.slider("最大条款数/文件", 5, 30, 10)
        st.info("减少条款数可提高速度，降低加载失败概率")
    
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
    
    # 控制按钮
    col1, col2 = st.columns(2)
    with col1:
        start_btn = st.button("开始分析", 
                             disabled=not (base_file and target_files and api_key) or st.session_state.analysis_running)
    with col2:
        if st.button("重置分析", disabled=not st.session_state.analysis_running):
            st.session_state.analysis_running = False
            st.session_state.last_error = ""
            st.session_state.partial_results = {}
            st.experimental_rerun()
    
    if start_btn:
        st.session_state.analysis_running = True
        st.session_state.last_error = ""
        
        try:
            # 处理基准文件
            with st.spinner("处理基准文件..."):
                status_box = st.markdown('<div class="status-box loading-box">处理基准文件中...</div>', unsafe_allow_html=True)
                
                base_text = extract_text_from_pdf(base_file)
                if not base_text:
                    st.error("无法从基准文件提取文本")
                    st.session_state.analysis_running = False
                    return
                
                base_clauses = split_into_clauses(base_text, max_clauses)
                status_box.markdown(f'<div class="status-box success-box">基准文件处理完成，提取到 {len(base_clauses)} 条条款</div>', unsafe_allow_html=True)
            
            # 处理目标文件
            for idx, target_file in enumerate(target_files, 1):
                if not st.session_state.analysis_running:
                    break
                    
                st.subheader(f"分析目标文件 {idx}/{len(target_files)}: {target_file.name}")
                target_status = st.markdown('<div class="status-box loading-box">准备分析...</div>', unsafe_allow_html=True)
                
                # 提取目标文件内容
                target_text = extract_text_from_pdf(target_file)
                if not target_text:
                    target_status.markdown('<div class="status-box error-box">无法提取文件内容，跳过</div>', unsafe_allow_html=True)
                    continue
                
                target_clauses = split_into_clauses(target_text, max_clauses)
                target_status.markdown(f'<div class="status-box loading-box">提取到 {len(target_clauses)} 条条款，正在匹配...</div>', unsafe_allow_html=True)
                
                # 匹配条款
                matched_pairs = match_clauses_with_base(base_clauses, target_clauses)
                if not matched_pairs:
                    target_status.markdown('<div class="status-box error-box">未找到匹配条款，无法分析</div>', unsafe_allow_html=True)
                    continue
                
                target_status.markdown(f'<div class="status-box loading-box">找到 {len(matched_pairs)} 对匹配条款，正在分析...</div>', unsafe_allow_html=True)
                
                # 生成报告
                report = generate_target_report(
                    matched_pairs,
                    base_file.name,
                    target_file.name,
                    api_key
                )
                
                # 显示结果
                target_status.markdown('<div class="status-box success-box">分析完成！</div>', unsafe_allow_html=True)
                st.markdown(get_download_link(report, f"{target_file.name}_合规性报告.txt"), unsafe_allow_html=True)
                
                with st.expander("查看报告预览"):
                    st.text_area("报告内容", report, height=200)
            
            # 完成所有分析
            st.session_state.analysis_running = False
            st.balloons()
            st.success("所有文件分析完成！")
                
        except Exception as e:
            st.session_state.analysis_running = False
            st.session_state.last_error = str(e)
            st.markdown(f'<div class="status-box error-box">分析出错: {str(e)}</div>', unsafe_allow_html=True)
            
            # 显示部分结果
            if st.session_state.partial_results:
                st.subheader("部分分析结果")
                for name, report in st.session_state.partial_results.items():
                    st.markdown(get_download_link(report, f"部分_{name}_合规性报告.txt"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
    
