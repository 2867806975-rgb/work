import re

def strong_clean_excel():
    # 输入文件名
    input_file = input("请输入你要清洗的文件名: ").strip()
    output_file = input_file.replace('.csv', '_强力清洗.csv')
    
    print(f"🧹 开始强力清洗: {input_file}")
    
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        lines = content.split('\n')
        cleaned_lines = []
        
        # 保留表头
        if lines:
            cleaned_lines.append(lines[0] + ',清洗状态')
        
        for i, line in enumerate(lines[1:], 1):  # 跳过表头
            if not line.strip():
                continue
            
            original_line = line
            cleaned_line = strong_clean_line(line)
            
            if cleaned_line and cleaned_line != original_line:
                cleaned_lines.append(cleaned_line + ',已清洗')
        
        # 保存结果
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(cleaned_lines))
        
        print(f"✅ 强力清洗完成！")
        print(f"原始行数: {len(lines)}")
        print(f"清洗后行数: {len(cleaned_lines)}")
        print(f"保存到: {output_file}")
        
        # 显示清洗效果
        show_cleaning_effect(lines, cleaned_lines)
        
    except FileNotFoundError:
        print(f"❌ 文件 {input_file} 不存在！")
    except Exception as e:
        print(f"❌ 错误: {e}")

def strong_clean_line(line):
    """强力清洗单行数据"""
    if not line.strip():
        return ""
    
    # 1. 移除多余空格和换行符
    line = re.sub(r'\s+', ' ', line)
    
    # 2. 移除各种干扰信息
    patterns_to_remove = [
        r'点赞\d+',           # 点赞数
        r'收藏\d+',           # 收藏数  
        r'分享\d+',           # 分享数
        r'回复\d+',           # 回复数
        r'\d+小时前',         # 时间信息
        r'\d+分钟前',
        r'\d+天前',
        r'\d+-\d+-\d+',      # 日期
        r'举报',              # 举报按钮
        r'作者赞过',          # 作者点赞标记
        r'展开',              # 展开按钮
        r'收起',              # 收起按钮
        r'查看更多回复',       # 查看更多
        r'http\S+',          # 网址链接
        r'@\S+',             # @用户
        r'#\S+#',            # 话题标签
        r'【.*?】',           # 方括号内容
        r'\[.*?\]',          # 英文方括号
        r'<.*?>',            # HTML标签
        r'[♡♥❤️💕💖]',       # 爱心表情
        r'[👍👎❤️🔥]',       # 其他表情
    ]
    
    for pattern in patterns_to_remove:
        line = re.sub(pattern, '', line)
    
    # 3. 移除特定位置的干扰文本
    noise_texts = [
        '点击查看', '查看图片', '图片评论', '语音评论',
        '视频评论', '位置:', '发布于', '编辑于',
        '已编辑', '删除', '置顶'
    ]
    
    for text in noise_texts:
        line = line.replace(text, '')
    
    # 4. 清理标点符号（保留中文标点）
    line = re.sub(r'[^\w\u4e00-\u9fff\s，。！？：；（）《》]', '', line)
    
    # 5. 移除纯数字或过短的行
    line = line.strip()
    if len(line) < 5 or line.isdigit():
        return ""
    
    # 6. 智能判断是否是有效评论
    if not is_valid_comment(line):
        return ""
    
    return line

def is_valid_comment(text):
    """判断是否是有效评论"""
    if len(text) < 5:
        return False
    
    # 无效内容特征
    invalid_patterns = [
        r'^回复$', r'^点赞$', r'^收藏$', r'^分享$',
        r'^作者$', r'^用户$', r'^评论$', r'^内容$',
        r'^\d+$', r'^\.+$'
    ]
    
    for pattern in invalid_patterns:
        if re.match(pattern, text):
            return False
    
    # 必须包含中文或实际内容
    if not re.search(r'[\u4e00-\u9fff]', text) and len(text) < 10:
        return False
    
    return True

def show_cleaning_effect(original_lines, cleaned_lines):
    """显示清洗效果对比"""
    print("\n" + "="*60)
    print("🧼 清洗效果对比")
    print("="*60)
    
    # 显示几个清洗前后的例子
    print("\n📝 清洗前后对比示例:")
    count = 0
    for i, orig_line in enumerate(original_lines[1:6], 1):  # 前5条数据
        if i < len(cleaned_lines):
            clean_line = cleaned_lines[i].split(',')[0]  # 取清洗后的内容部分
            
            print(f"\n示例 {i}:")
            print(f"  清洗前: {orig_line[:80]}..." if len(orig_line) > 80 else f"  清洗前: {orig_line}")
            print(f"  清洗后: {clean_line[:80]}..." if len(clean_line) > 80 else f"  清洗后: {clean_line}")
            count += 1
    
    # 统计信息
    print(f"\n📊 清洗统计:")
    print(f"  - 原始数据: {len(original_lines)} 行")
    print(f"  - 清洗后: {len(cleaned_lines)} 行") 
    print(f"  - 过滤掉: {len(original_lines) - len(cleaned_lines)} 行")
    print(f"  - 保留率: {len(cleaned_lines)/len(original_lines)*100:.1f}%")

# 运行清洗
if __name__ == "__main__":
    strong_clean_excel()
