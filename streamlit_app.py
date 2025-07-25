import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import os
import tempfile
from datetime import datetime

# 设置页面配置
st.set_page_config(
    page_title="条款合规性对比工具",
    page_icon="📄",
    layout="wide"
)

# 页面标题
st.title("📄 条款合规性对比工具")
st.write("上传基准文件和待比较文件，系统将自动进行条款匹配分析并生成合规性报告。")

# 辅助函数：从docx文件中提取文本
def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():  # 只添加非空段落
            full_text.append(para.text)
    return '\n'.join(full_text)

# 辅助函数：拆分条款（假设条款以数字开头，如"1. "、"2.1 "等）
def split_terms(text):
    # 使用正则表达式匹配条款编号，如1. 1.1 2. 等
    pattern = r'(\d+\.\s|\d+\.\d+\s|\d+\s)'
    terms = re.split(pattern, text)
    
    # 重组条款，将编号和内容合并
    result = []
    for i in range(1, len(terms), 2):
        if i + 1 < len(terms):
            term_number = terms[i].strip()
            term_content = terms[i+1].strip()
            result.append(f"{term_number} {term_content}")
    
    # 如果没有匹配到条款编号格式，将整个文本作为一个条款
    if not result:
        result.append(text)
    
    return result

# 辅助函数：简单的条款匹配（基于关键词相似度）
def match_terms(benchmark_term, compare_terms, threshold=0.3):
    benchmark_words = set(re.findall(r'\w+', benchmark_term.lower()))
    best_match = None
    best_score = 0
    
    for term in compare_terms:
        compare_words = set(re.findall(r'\w+', term.lower()))
        # 计算词集交集比例
        common_words = benchmark_words.intersection(compare_words)
        score = len(common_words) / len(benchmark_words) if benchmark_words else 0
        
        if score > best_score and score >= threshold:
            best_score = score
            best_match = term
    
    return best_match, best_score

# 辅助函数：生成Word报告
def generate_word_report(benchmark_name, compare_name, matched_terms, unmatched_benchmark, unmatched_compare):
    doc = docx.Document()
    
    # 添加标题
    title = doc.add_heading('条款合规性对比报告', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 添加报告信息
    doc.add_paragraph(f"基准文件: {benchmark_name}")
    doc.add_paragraph(f"对比文件: {compare_name}")
    doc.add_paragraph(f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph("")  # 空行
    
    # 添加匹配的条款部分
    doc.add_heading('一、匹配的条款', level=1)
    if matched_terms:
        for i, (benchmark_term, compare_term, score) in enumerate(matched_terms, 1):
            doc.add_heading(f"匹配项 {i} (相似度: {score:.2f})", level=2)
            
            p = doc.add_paragraph("基准条款: ")
            p.add_run(benchmark_term).bold = True
            
            p = doc.add_paragraph("对比条款: ")
            p.add_run(compare_term).bold = True
            
            doc.add_paragraph("")  # 空行
    else:
        doc.add_paragraph("未找到匹配的条款")
    
    # 添加基准文件中未匹配的条款
    doc.add_heading('二、基准文件中未匹配的条款', level=1)
    if unmatched_benchmark:
        for term in unmatched_benchmark:
            p = doc.add_paragraph(term)
            p.italic = True
    else:
        doc.add_paragraph("所有基准条款均找到匹配项")
    
    # 添加对比文件中未匹配的条款（不合规总结）
    doc.add_heading('三、对比文件中未匹配的条款（不合规总结）', level=1)
    if unmatched_compare:
        for term in unmatched_compare:
            p = doc.add_paragraph(term)
            p.italic = True
    else:
        doc.add_paragraph("对比文件所有条款均与基准文件匹配")
    
    # 保存到临时文件
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(temp_file.name)
    temp_file.close()
    
    return temp_file.name

# 主函数
def main():
    # 文件上传区域
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("基准文件")
        benchmark_file = st.file_uploader("上传基准条款文件 (docx)", type=["docx"], key="benchmark")
    
    with col2:
        st.subheader("待比较文件")
        compare_file = st.file_uploader("上传待比较条款文件 (docx)", type=["docx"], key="compare")
    
    # 分析按钮
    if st.button("开始分析") and benchmark_file and compare_file:
        with st.spinner("正在分析文件..."):
            # 保存上传的文件到临时目录
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as f:
                f.write(benchmark_file.getbuffer())
                benchmark_path = f.name
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as f:
                f.write(compare_file.getbuffer())
                compare_path = f.name
            
            # 提取文本
            benchmark_text = extract_text_from_docx(benchmark_path)
            compare_text = extract_text_from_docx(compare_path)
            
            # 拆分条款
            benchmark_terms = split_terms(benchmark_text)
            compare_terms = split_terms(compare_text)
            
            # 显示条款数量
            st.info(f"基准文件提取到 {len(benchmark_terms)} 条条款")
            st.info(f"对比文件提取到 {len(compare_terms)} 条条款")
            
            # 进行条款匹配
            matched_terms = []
            matched_compare_indices = set()
            
            for benchmark_term in benchmark_terms:
                match, score = match_terms(benchmark_term, compare_terms)
                if match:
                    matched_terms.append((benchmark_term, match, score))
                    # 记录已匹配的对比条款索引
                    matched_compare_indices.add(compare_terms.index(match))
            
            # 找出未匹配的条款
            unmatched_benchmark = []
            matched_benchmark_terms = [term[0] for term in matched_terms]
            for term in benchmark_terms:
                if term not in matched_benchmark_terms:
                    unmatched_benchmark.append(term)
            
            unmatched_compare = []
            for i, term in enumerate(compare_terms):
                if i not in matched_compare_indices:
                    unmatched_compare.append(term)
            
            # 显示结果
            st.subheader("分析结果")
            
            # 显示匹配的条款
            with st.expander("查看匹配的条款", expanded=True):
                if matched_terms:
                    for i, (benchmark_term, compare_term, score) in enumerate(matched_terms, 1):
                        st.markdown(f"**匹配项 {i} (相似度: {score:.2f})**")
                        st.markdown(f"基准条款: {benchmark_term}")
                        st.markdown(f"对比条款: {compare_term}")
                        st.markdown("---")
                else:
                    st.warning("未找到匹配的条款")
            
            # 显示未匹配的基准条款
            with st.expander("查看基准文件中未匹配的条款"):
                if unmatched_benchmark:
                    for term in unmatched_benchmark:
                        st.markdown(f"- {term}")
                else:
                    st.success("所有基准条款均找到匹配项")
            
            # 显示未匹配的对比条款（不合规总结）
            with st.expander("查看对比文件中未匹配的条款（不合规总结）"):
                if unmatched_compare:
                    for term in unmatched_compare:
                        st.markdown(f"- {term}")
                else:
                    st.success("对比文件所有条款均与基准文件匹配")
            
            # 生成并提供下载Word报告
            st.subheader("生成报告")
            report_path = generate_word_report(
                benchmark_file.name, 
                compare_file.name, 
                matched_terms, 
                unmatched_benchmark, 
                unmatched_compare
            )
            
            # 提供下载
            with open(report_path, "rb") as file:
                st.download_button(
                    label="下载合规性报告 (Word)",
                    data=file,
                    file_name=f"合规性对比报告_{benchmark_file.name.split('.')[0]}_vs_{compare_file.name.split('.')[0]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            # 清理临时文件
            os.unlink(benchmark_path)
            os.unlink(compare_path)
            os.unlink(report_path)

if __name__ == "__main__":
    main()
