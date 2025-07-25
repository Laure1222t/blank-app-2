def split_chinese_terms(text):
    """拆分中文条款（修复正则表达式错误，增强稳定性）"""
    # 输入验证
    if not text or not isinstance(text, str):
        st.warning("输入文本为空或无效，无法拆分条款")
        return []
    
    # 预处理：清除多余空行和空格，统一标点
    processed_text = re.sub(r'\n+', '\n', text.strip())
    processed_text = re.sub(r'\s+', ' ', processed_text)
    # 替换全角标点为半角，便于统一处理
    processed_text = processed_text.replace('。', '.').replace('，', ',').replace('；', ';')
    processed_text = processed_text.replace('：', ':').replace('（', '(').replace('）', ')')
    
    # 中文条款常见编号格式（正则模式）
    # 修复了可能导致"nothing to repeat"错误的模式，确保每个模式都是有效的
    patterns = [
        r'(\d+\.\d+\.\d+\s+)',        # 1.1.1 
        r'(\d+\.\d+\s+)',             # 1.1 
        r'(\d+\.\s+)',                # 1. 
        r'(\(\d+\)\.\s+)',            # (1). 
        r'(\(\d+\)\s+)',              # (1) 
        r'([一二三四五六七八九十]+、\s+)',  # 一、 
        r'(第[一二三四五六七八九十]+条\s+)', # 第一条
        r'(第[一二三四五六七八九十]+款\s+)', # 第一款
        r'(第[一二三四五六七八九十]+项\s+)', # 第一项
        r'(\d+\)\s+)',                # 1)
        r'([A-Za-z]\.\s+)',           # A. 
        r'([A-Za-z]\)\s+)',           # A)
    ]
    
    try:
        # 组合所有模式，使用非捕获组包裹
        combined_pattern = '|'.join([f'({p})' for p in patterns])
        
        # 查找所有匹配的条款编号位置
        matches = list(re.finditer(combined_pattern, processed_text, re.MULTILINE))
        
        if not matches:
            # 如果没有找到标准编号，尝试按空行拆分
            st.info("未检测到标准条款编号，尝试按空行拆分")
            terms = [t.strip() for t in re.split(r'\n\s*\n', processed_text) if t.strip()]
            return terms
        
        # 提取条款内容
        terms = []
        # 第一条条款从文本开始到第一个匹配
        first_match = matches[0]
        if first_match.start() > 0:
            pre_text = processed_text[:first_match.start()].strip()
            if pre_text:
                terms.append(pre_text)
        
        # 处理中间的条款
        for i in range(len(matches)):
            current_match = matches[i]
            start_idx = current_match.start()
            
            # 确定当前条款的结束位置
            if i < len(matches) - 1:
                end_idx = matches[i+1].start()
            else:
                end_idx = len(processed_text)
            
            # 提取条款内容（包含编号）
            term_content = processed_text[start_idx:end_idx].strip()
            if term_content:
                terms.append(term_content)
        
        # 过滤过短的条款（可能是误拆分）
        min_term_length = 10  # 最小条款长度
        terms = [term for term in terms if len(term) >= min_term_length]
        
        # 拆分效果反馈
        st.success(f"成功拆分条款：{len(terms)}条")
        return terms
        
    except re.error as e:
        st.error(f"正则表达式错误: {str(e)}")
        # 出错时的备选方案：按空行拆分
        st.info("使用备选方案拆分条款")
        terms = [t.strip() for t in re.split(r'\n\s*\n', processed_text) if t.strip()]
        return terms
    except Exception as e:
        st.error(f"条款拆分失败: {str(e)}")
        return [processed_text]  # 返回原始文本作为最后的备选
    
