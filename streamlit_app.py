# 优化的中文条款拆分函数，增加空值和异常处理
def split_chinese_terms(text):
    """
    拆分中文条款，支持多种编号格式，增加空值和异常处理
    """
    # 首先检查输入是否有效
    if not text or not isinstance(text, str):
        st.warning("输入文本为空或无效，无法进行条款拆分")
        return []
    
    # 清除多余空行和空格
    text = re.sub(r'\n+', '\n', text.strip())
    
    # 中文条款常见的编号格式正则表达式
    # 支持: 1.  1.1  (1)  一、  第一条  第一款  1)  等格式
    patterns = [
        r'(\d+\.\s+)',                # 1. 
        r'(\d+\.\d+\s+)',             # 1.1 
        r'(\(\d+\)\s+)',              # (1) 
        r'([一二三四五六七八九十]+\、\s+)',  # 一、 
        r'(第[一二三四五六七八九十]条\s+)',   # 第一条
        r'(第[一二三四五六七八九十]款\s+)',   # 第一款
        r'(\d+\)\s+)',                # 1)
        r'([A-Za-z]\.\s+)',           # A. 
    ]
    
    # 组合所有模式
    combined_pattern = '|'.join(patterns)
    
    # 拆分文本
    parts = re.split(combined_pattern, text)
    
    terms = []
    current_term = ""
    
    for part in parts:
        # 跳过空值或仅含空白字符的部分
        if not part or not part.strip():
            continue
            
        # 检查当前部分是否为条款编号
        is_numbering = any(re.fullmatch(pattern.strip(), part.strip()) for pattern in patterns)
        
        if is_numbering:
            # 如果已有内容，先保存当前条款
            if current_term.strip():
                terms.append(current_term.strip())
            # 开始新条款
            current_term = part
        else:
            # 累加条款内容
            current_term += part
    
    # 添加最后一个条款
    if current_term.strip():
        terms.append(current_term.strip())
    
    # 如果条款数量较少，可能是拆分效果不好，提示用户
    if len(terms) < 3 and len(text) > 500:
        st.info(f"检测到可能的条款拆分效果不佳（共{len(terms)}条），建议检查文件格式")
    
    return terms
    
