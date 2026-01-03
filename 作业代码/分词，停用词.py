import pandas as pd
import jieba
import re
import os

# -------------------------- 1. 自动定位Excel文件（洛天依/赵丽颖） --------------------------
def find_idol_excel():
    """扫描当前目录，找到含洛天依/赵丽颖的Excel文件"""
    excel_map = {}
    current_dir = os.getcwd()
    for file in os.listdir(current_dir):
        if file.endswith(".xlsx"):
            if "洛天依" in file:
                excel_map["洛天依"] = os.path.join(current_dir, file)
            elif "赵丽颖" in file:
                excel_map["赵丽颖"] = os.path.join(current_dir, file)
    if not excel_map:
        raise FileNotFoundError("未找到洛天依/赵丽颖的Excel文件！请确认文件名含偶像名")
    print("✅ 找到目标Excel：")
    for idol, path in excel_map.items():
        print(f"- {idol}：{os.path.basename(path)}")
    return excel_map

# -------------------------- 2. 文本清洗+分词核心函数 --------------------------
def clean_comment(comment):
    """清洗单条中文评论：去空值、特殊字符、多余空格"""
    if pd.isna(comment) or str(comment).strip() == "":
        return None
    # 仅保留中文，去除所有符号/数字/英文
    cleaned = re.sub(r"[^\u4e00-\u9fa5\s]", "", str(comment))
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned if cleaned else None

def get_raw_tokens(cleaned_comment):
    """仅分词（保留停用词）→ 返回单词列表"""
    if not cleaned_comment:
        return []
    return jieba.lcut(cleaned_comment, cut_all=False)  # 精确分词

def get_filtered_tokens(cleaned_comment):
    """分词+去停用词 → 返回单词列表"""
    if not cleaned_comment:
        return []
    # 1. 先分词
    tokens = jieba.lcut(cleaned_comment, cut_all=False)
    # 2. 中文停用词库（核心无意义词）
    stop_words = set([
        "的", "了", "是", "我", "你", "他", "她", "它", "我们", "你们", "他们",
        "这", "那", "此", "彼", "和", "与", "及", "或", "但", "而", "却", "若",
        "因为", "所以", "虽然", "但是", "如果", "只要", "只有", "由于", "因此",
        "在", "于", "到", "从", "向", "对", "对于", "关于", "把", "被", "让",
        "能", "会", "可以", "可能", "应该", "必须", "需要", "要", "不要", "没",
        "不", "没", "无", "非", "否", "别", "很", "非常", "太", "更", "最", "比较",
        "还", "也", "又", "再", "才", "就", "都", "全", "总", "共", "所有",
        "一个", "一些", "一点", "一样", "一起", "一直", "一定", "一般",
        "啊", "呀", "呢", "吗", "吧", "啦", "哦", "哇", "唉"
    ])
    # 3. 过滤停用词+单字
    return [t for t in tokens if t not in stop_words and len(t)>=2]

# -------------------------- 3. 生成示例格式的TXT（每行一个词） --------------------------
def generate_txt_by_idol(idol_name, excel_path):
    """为单个偶像生成两个TXT：仅分词.txt + 去停用词.txt"""
    print(f"\n===== 处理【{idol_name}】=====")
    # 读取Excel
    df = pd.read_excel(excel_path, engine="openpyxl")
    print(f"1. 原始数据：{len(df)}条")
    
    # 选择评论列
    print(f"\n评论列选择（输入序号）：")
    for i, col in enumerate(df.columns, 1):
        print(f"   {i}. {col}")
    col_idx = int(input("   输入评论列序号：")) - 1
    comment_col = df.columns[col_idx]
    
    # 清洗评论
    df["清洗后"] = df[comment_col].apply(clean_comment)
    df_valid = df.dropna(subset=["清洗后"]).reset_index(drop=True)
    print(f"2. 有效评论：{len(df_valid)}条")
    
    # 收集所有分词结果（扁平化：所有评论的词合并成一个列表）
    all_raw_tokens = []    # 含停用词的所有词
    all_filtered_tokens = []  # 去停用词的所有词
    for comment in df_valid["清洗后"]:
        all_raw_tokens.extend(get_raw_tokens(comment))
        all_filtered_tokens.extend(get_filtered_tokens(comment))
    
    # 生成TXT（每行一个词，和参考格式完全一致）
    raw_txt = f"{idol_name}_仅分词结果.txt"
    filtered_txt = f"{idol_name}_去除停用词之后结果.txt"
    
    # 保存仅分词（含停用词）
    with open(raw_txt, "w", encoding="utf-8") as f:
        for token in all_raw_tokens:
            f.write(token + "\n")
    # 保存去停用词后
    with open(filtered_txt, "w", encoding="utf-8") as f:
        for token in all_filtered_tokens:
            f.write(token + "\n")
    
    print(f"3. 文件生成完成：")
    print(f"   - {raw_txt}（共{len(all_raw_tokens)}个词，含停用词）")
    print(f"   - {filtered_txt}（共{len(all_filtered_tokens)}个词，无停用词）")

# -------------------------- 4. 主流程 --------------------------
if __name__ == "__main__":
    try:
        excel_map = find_idol_excel()
        for idol, path in excel_map.items():
            generate_txt_by_idol(idol, path)
        print("\n🎉 全部处理完成！每个偶像生成2个TXT（每行一个词）：")
        print("   格式完全匹配参考示例，可直接打开查看/使用！")
    except Exception as e:
        print(f"❌ 错误：{e}")
