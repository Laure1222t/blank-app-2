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
    page_title="条款式政策比对分析工具",
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
    .matched-clause {
        border-left: 4px solid #28a745;
        padding: 0.75rem;
        margin: 1rem 0;
        background-color: #f8fff8;
    }
    .difference-section {
        border-left: 4px solid #ffc107;
        padding: 0.75rem;
        margin: 0.5rem 0;
        background-color: #fffcf2;
    }
    .summary-box {
        border: 1px solid #007bff;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
        background-color: #f0f7ff;
    }
    .parse-info {
        font-size: 0.9rem;
        color: #6c757d;
        margin-top: 0.5rem;
    }
    .clause-item {
        padding: 0.5rem;
        margin: 0.25rem 0;
        border-radius: 3px;
        background-color: #f0f2f6;
    }
</style>
""", unsafe_allow_html=True)

# 初始化会话状态
if 'target_clauses' not in st.session_state:
    st.session_state.target_clauses = {}  # {条款号: 内容}
if 'compare_files' not in st.session_state:
    st.session_state.compare_files = {}  # {文件名: {条款: {}, 分析结果: {匹配结果: {}, 总结: ""}}}
if 'current_file' not in st.session_state:
    st.session_state.current_file = None
if 'api_key' not in st.session_state:
    st.session_state.api_key = os.getenv("QWEN_API_KEY", "")
if 'max_clauses' not in st.session_state:
    st.session_state.max_clauses = 30  # 默认最大条款数
if 'parse_precision' not in st.session_state:
    st.session_state.parse_precision = "中等"  # 解析精度

# 页面标题
st.title("📜 条款式政策比对分析工具")
st.markdown("精确解析政策文件条款，按编号匹配并分析合规性与差异")
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

# 解析精度设置
st.session_state.parse_precision = st.sidebar.select_slider(
    "条款解析精度",
    options=["宽松", "中等", "严格"],
    value=st.session_state.parse_precision,
    help="宽松：提取更多可能的条款；严格：只提取明确符合格式的条款"
)

# API配置
with st.expander("🔑 API 配置", expanded=not st.session_state.api_key):
    st.session_state.api_key = st.text_input("请输入Qwen API密钥", value=st.session_state.api_key, type="password")
    model_option = st.selectbox(
        "选择Qwen模型",
        ["qwen-turbo", "qwen-plus", "qwen-max"],
        index=0  # 默认使用轻量版
    )
    st.caption("提示：可从阿里云DashScope平台获取API密钥")

# 优化的PDF解析函数 - 增强条款拆分能力
def parse_pdf_by_clauses(file, max_clauses=30, precision="中等"):
    """解析PDF文件并按条款号提取结构化条款，优化拆分逻辑"""
    try:
        with st.spinner("正在解析文件并拆分条款..."):
            doc = fitz.open(stream=file.read(), filetype="pdf")
            total_pages = len(doc)
            full_text = ""
            
            # 逐页读取文本，保留页面信息
            for page_num, page in enumerate(doc, 1):
                page_text = page.get_text()
                # 清理页面文本并添加页分隔符
                full_text += f"\n\n[[PAGE {page_num}]]\n{page_text}"
            
            # 文本预处理 - 增强条款识别
            full_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', full_text)  # 移除控制字符
            full_text = re.sub(r'(\r\n|\r|\n)+', '\n', full_text)  # 统一换行符
            full_text = re.sub(r'[^\S\n]+', ' ', full_text)  # 替换非换行空白字符为空格
            full_text = re.sub(r'(\d+)\.(\d+)', r'\1.\2', full_text)  # 修复数字间的点
            full_text = full_text.strip()
            
            # 根据精度选择不同的条款提取模式
            clauses = {}
            
            # 主要条款模式：第X条
            primary_pattern = re.compile(r'(第[零一二三四五六七八九十百千\d]+\s*条\s+.*?)(?=第[零一二三四五六七八九十百千\d]+\s*条\s+|$)', re.DOTALL)
            primary_matches = primary_pattern.findall(full_text)
            
            if primary_matches:
                # 从主要模式提取条款
                for match in primary_matches:
                    clause_num_match = re.search(r'第([零一二三四五六七八九十百千\d]+)\s*条', match)
                    if clause_num_match:
                        clause_num = clause_num_match.group(1)
                        # 清理条款内容
                        clause_content = re.sub(r'第[零一二三四五六七八九十百千\d]+\s*条\s*', '', match).strip()
                        # 移除页码标记
                        clause_content = re.sub(r'\[\[PAGE \d+\]\]', '', clause_content)
                        
                        # 根据精度过滤条款
                        if clause_content:
                            if precision == "严格" and len(clause_content) > 50:
                                clauses[clause_num] = clause_content
                            elif precision == "中等" and len(clause_content) > 30:
                                clauses[clause_num] = clause_content
                            elif precision == "宽松" and len(clause_content) > 20:
                                clauses[clause_num] = clause_content
                    
                    # 达到最大条款数则停止
                    if 0 < max_clauses <= len(clauses):
                        break
            
            # 如果主要模式提取不足，尝试辅助模式
            if not clauses or len(clauses) < 5:
                st.markdown('<p class="parse-info">尝试其他条款格式提取...</p>', unsafe_allow_html=True)
                
                # 辅助模式1：数字编号 (1., 1.1, 1.1.1等)
                alt_patterns = [
                    re.compile(r'(\d+\.\d+\.\d+\s+.*?)(?=\d+\.\d+\.\d+\s+|$)', re.DOTALL),  # 三级
                    re.compile(r'(\d+\.\d+\s+.*?)(?=\d+\.\d+\s+|$)', re.DOTALL),          # 二级
                    re.compile(r'(\d+\s+.*?)(?=\d+\s+|$)', re.DOTALL)                     # 一级
                ]
                
                for pattern in alt_patterns:
                    alt_matches = pattern.findall(full_text)
                    if alt_matches and len(alt_matches) > len(clauses):
                        for i, match in enumerate(alt_matches):
                            match = match.strip()
                            if match:
                                # 移除页码标记
                                clean_match = re.sub(r'\[\[PAGE \d+\]\]', '', match)
                                # 提取数字编号
                                num_match = re.search(r'^(\d+(\.\d+)*)', clean_match)
                                if num_match:
                                    clause_num = num_match.group(1)
                                    clause_content = re.sub(r'^\d+(\.\d+)*\s*', '', clean_match).strip()
                                else:
                                    clause_num = str(i+1)
                                    clause_content = clean_match
                                
                                if clause_content:
                                    clauses[clause_num] = clause_content
                                
                                # 达到最大条款数则停止
                                if 0 < max_clauses <= len(clauses):
                                    break
                        if clauses:
                            break
            
            # 最终过滤和整理
            final_clauses = {}
            for num, content in clauses.items():
                # 移除多余空白和清理内容
                cleaned = re.sub(r'\s+', ' ', content).strip()
                if len(cleaned) > 20:  # 确保条款有足够内容
                    final_clauses[num] = cleaned
            
            # 限制最大条款数
            if 0 < max_clauses < len(final_clauses):
                final_clauses = dict(list(final_clauses.items())[:max_clauses])
            
            st.success(f"共解析 {total_pages} 页，成功提取 {len(final_clauses)} 条条款")
            return final_clauses
            
    except Exception as e:
        st.error(f"文件解析错误: {str(e)}")
        return {}

# 调用Qwen API进行条款比对分析
def call_qwen_api(prompt, api_key, model="qwen-turbo"):
    """调用Qwen API进行条款比对分析"""
    if not api_key:
        st.error("请先配置API密钥")
        return None
    
    try:
        with st.spinner("正在分析条款..."):
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
                    "temperature": 0.5,
                    "top_p": 0.9,
                    "max_tokens": 800
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

# 合规性分析函数 - 按条款匹配
def analyze_clause_matches(target_clauses, compare_clauses, api_key, model):
    """按条款匹配进行合规性分析，只分析匹配的条款"""
    if not target_clauses or not compare_clauses:
        st.warning("缺少条款内容，无法进行分析")
        return None, None
    
    # 找到匹配的条款（条款号相同）
    matched_clause_nums = [num for num in target_clauses if num in compare_clauses]
    
    if not matched_clause_nums:
        st.info("未找到匹配的条款")
        return {}, "未找到匹配的条款，无法进行合规性分析。"
    
    # 分析每个匹配的条款
    matched_results = {}
    for clause_num in matched_clause_nums:
        target_content = target_clauses[clause_num]
        compare_content = compare_clauses[clause_num]
        
        # 生成条款比对提示
        prompt = f"""
        请比对以下两条政策条款的合规性和差异：
        
        目标条款（第{clause_num}条）：
        {target_content[:300]}
        
        待比对条款（第{clause_num}条）：
        {compare_content[:300]}
        
        分析要求：
        1. 判断待比对条款是否符合目标条款要求
        2. 指出两者的主要差异点（如无差异则说明一致）
        3. 分析差异可能带来的影响
        4. 用简洁的中文（不超过300字）输出分析结果
        """
        
        # 调用API分析
        result = call_qwen_api(prompt, api_key, model)
        if result:
            matched_results[clause_num] = {
                "target": target_content,
                "compare": compare_content,
                "analysis": result
            }
    
    # 生成总体总结
    summary_prompt = f"""
    以下是目标政策文件与待比对文件中匹配条款的分析结果：
    {json.dumps(matched_results, ensure_ascii=False, indent=2)}
    
    请基于以上分析，用简洁的中文（不超过300字）总结：
    1. 总体合规性情况
    2. 主要差异点汇总
    3. 简要的合规建议
    """
    
    summary = call_qwen_api(summary_prompt, api_key, model) or "无法生成总结，请检查API配置。"
    
    return matched_results, summary

# 生成Word文档
def generate_word_document(matched_results, summary, target_filename, compare_filename):
    """生成Word格式分析报告"""
    try:
        doc = Document()
        
        # 标题
        title = doc.add_heading("政策文件条款比对分析报告", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 基本信息
        doc.add_paragraph(f"目标政策文件: {target_filename}")
        doc.add_paragraph(f"待比对文件: {compare_filename}")
        doc.add_paragraph(f"分析日期: {time.strftime('%Y年%m月%d日')}")
        doc.add_paragraph("")
        
        # 总体总结
        doc.add_heading("一、总体总结", level=1)
        for para in re.split(r'\n+', summary):
            if para.strip():
                doc.add_paragraph(para.strip())
        
        # 匹配条款分析
        doc.add_heading("二、匹配条款详细分析", level=1)
        
        for clause_num, details in matched_results.items():
            doc.add_heading(f"第{clause_num}条", level=2)
            
            p = doc.add_paragraph("目标条款内容：")
            p.style = 'Heading 3'
            doc.add_paragraph(details["target"])
            
            p = doc.add_paragraph("待比对条款内容：")
            p.style = 'Heading 3'
            doc.add_paragraph(details["compare"])
            
            p = doc.add_paragraph("分析结果：")
            p.style = 'Heading 3'
            for para in re.split(r'\n+', details["analysis"]):
                if para.strip():
                    doc.add_paragraph(para.strip())
        
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
    st.caption("作为基准的政策文件，系统将自动识别并拆分条款")
    target_file = st.file_uploader("上传目标政策文件 (PDF)", type="pdf", key="target")
    
    if target_file:
        # 解析目标文件条款
        st.session_state.target_clauses = parse_pdf_by_clauses(
            target_file, 
            max_clauses=st.session_state.max_clauses,
            precision=st.session_state.parse_precision
        )
        
        with st.expander(f"查看提取的条款 (共 {len(st.session_state.target_clauses)} 条)"):
            for num, content in st.session_state.target_clauses.items():
                display_text = content[:150] + "..." if len(content) > 150 else content
                st.markdown(f'<div class="clause-item"><strong>第{num}条:</strong> {display_text}</div>', unsafe_allow_html=True)
    
    # 多文件上传区域
    st.subheader("待比对文件")
    st.caption("可上传多个文件，系统将自动拆分条款并按编号匹配")
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
                # 解析待比对文件条款
                clauses = parse_pdf_by_clauses(
                    file, 
                    max_clauses=st.session_state.max_clauses,
                    precision=st.session_state.parse_precision
                )
                # 确保新文件的字典结构完整
                st.session_state.compare_files[file.name] = {
                    "clauses": clauses,
                    "matched_results": None,
                    "summary": None
                }
                st.success(f"✅ 已添加 {file.name}，提取到 {len(clauses)} 条条款")
    
    # 显示已上传的待比对文件列表
    if st.session_state.compare_files:
        st.subheader("已上传文件")
        for filename in st.session_state.compare_files.keys():
            col_a, col_b = st.columns([3, 1])
            with col_a:
                clause_count = len(st.session_state.compare_files[filename]["clauses"])
                st.markdown(f"- {filename} (条款数: {clause_count})")
            with col_b:
                if st.button("分析", key=f"analyze_{filename}") and st.session_state.target_clauses:
                    # 进行条款匹配分析
                    matched_results, summary = analyze_clause_matches(
                        st.session_state.target_clauses,
                        st.session_state.compare_files[filename]["clauses"],
                        st.session_state.api_key,
                        model_option
                    )
                    if matched_results is not None:
                        st.session_state.compare_files[filename]["matched_results"] = matched_results
                        st.session_state.compare_files[filename]["summary"] = summary
                        st.session_state.current_file = filename
                        st.success(f"✅ {filename} 分析完成，找到 {len(matched_results)} 条匹配条款")

with col2:
    st.subheader("分析结果")
    
    # 显示文件选择标签
    if st.session_state.compare_files:
        st.markdown("**选择文件查看结果：**")
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
                        # 安全检查：确保matched_results存在且不为None
                        if "matched_results" in data and data["matched_results"]:
                            match_count = len(data["matched_results"])
                            status = f" ({match_count}条匹配)"
                        else:
                            status = ""
                        
                        if st.button(f"{filename.split('.')[0]}{status}", key=f"tab_{filename}"):
                            st.session_state.current_file = filename
    
    # 显示当前选中文件的分析结果
    if st.session_state.current_file:
        filename = st.session_state.current_file
        # 确保文件数据存在
        if filename in st.session_state.compare_files:
            file_data = st.session_state.compare_files[filename]
            # 安全获取匹配结果和总结
            matched_results = file_data.get("matched_results", None)
            summary = file_data.get("summary", "")
            
            if matched_results is not None:
                # 显示总体总结
                st.markdown("### 📊 总体分析总结")
                st.markdown('<div class="summary-box">', unsafe_allow_html=True)
                for para in re.split(r'\n+', summary):
                    if para.strip():
                        st.markdown(f"{para.strip()}  \n")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # 显示匹配条款的详细分析
                if matched_results:
                    st.markdown(f"### 🔍 匹配条款详情 ({len(matched_results)} 条)")
                    
                    for clause_num, details in matched_results.items():
                        st.markdown(f'#### 第{clause_num}条')
                        st.markdown('<div class="matched-clause">', unsafe_allow_html=True)
                        
                        st.markdown("**目标条款内容：**")
                        st.write(details["target"][:500] + "..." if len(details["target"]) > 500 else details["target"])
                        
                        st.markdown("**待比对条款内容：**")
                        st.write(details["compare"][:500] + "..." if len(details["compare"]) > 500 else details["compare"])
                        
                        st.markdown('<div class="difference-section">', unsafe_allow_html=True)
                        st.markdown("**分析结果：**")
                        for para in re.split(r'\n+', details["analysis"]):
                            if para.strip():
                                st.markdown(f"{para.strip()}  \n")
                        st.markdown('</div>', unsafe_allow_html=True)
                        
                        st.markdown('</div>', unsafe_allow_html=True)
                
                # 生成并下载Word文档
                if target_file and matched_results is not None:
                    word_file = generate_word_document(
                        matched_results,
                        summary,
                        target_file.name,
                        filename
                    )
                    
                    if word_file:
                        with open(word_file, "rb") as f:
                            st.download_button(
                                label=f"💾 下载 {filename} 的分析报告",
                                data=f,
                                file_name=f"政策条款比对报告_{filename}_{time.strftime('%Y%m%d')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        os.unlink(word_file)
            else:
                st.info("请点击文件旁的'分析'按钮生成分析结果")
        else:
            st.warning("所选文件不存在，请重新选择")
    else:
        st.info("请上传待比对文件并选择一个文件查看分析结果")

# 帮助信息
with st.expander("ℹ️ 使用帮助"):
    st.markdown("""
    ### 条款解析说明
    系统采用多层次条款识别机制：
    1. 优先识别"第X条"格式的标准条款
    2. 如识别不足，自动尝试数字编号格式(1., 1.1等)
    3. 可通过侧边栏调整解析精度：
       - 宽松：提取更多可能的条款，适合格式不规范的文件
       - 中等：平衡提取数量和准确性（默认）
       - 严格：只提取明确符合格式的条款，适合规范文件
    
    ### 使用流程
    1. 上传目标政策文件和待比对文件
    2. 可在侧边栏调整条款提取参数
    3. 点击"分析"按钮进行条款匹配与分析
    4. 查看总体总结和匹配条款的详细分析
    5. 下载Word格式报告（可选）
    
    ### 提高解析效果的建议
    - 确保PDF文件为可复制文本（非图片扫描件）
    - 条款编号清晰的文件解析效果更佳
    - 如提取条款不足，尝试调整为"宽松"解析精度
    - 过长文件可适当增加最大条款数量限制
    """)
    
