# 导入所需库
import pandas as pd
from snownlp import SnowNLP
import jieba
from collections import Counter
import re

# ===================== 1. 读取Excel格式数据（核心修改部分） =====================
# 请替换为你的Excel文件路径（和代码同文件夹直接写文件名，否则写完整路径，如：C:/data/赵丽颖评论_清洗后.xlsx）
zhao_file_path = "C:/桌面/赵丽颖评论爬取.xlsx"  # 赵丽颖Excel文件
luo_file_path = "C:/桌面/洛天依评论爬取(1).xlsx"    # 洛天依Excel文件

# 读取Excel（指定编码，解决中文乱码；usecols只读取需要的列，提升速度）
try:
    # 优先用utf-8编码，若报错改用gbk（中文Windows常见）
    zhao_df = pd.read_excel(zhao_file_path, usecols=["评论内容"], encoding="utf-8")
    luo_df = pd.read_excel(luo_file_path, usecols=["评论内容"], encoding="utf-8")
except:
    zhao_df = pd.read_excel(zhao_file_path, usecols=["评论内容"], encoding="gbk")
    luo_df = pd.read_excel(luo_file_path, usecols=["评论内容"], encoding="gbk")

# 清理空值评论（避免分析报错）
zhao_df = zhao_df.dropna(subset=["评论内容"])
luo_df = luo_df.dropna(subset=["评论内容"])
zhao_df["评论内容"] = zhao_df["评论内容"].astype(str)  # 统一转为字符串
luo_df["评论内容"] = luo_df["评论内容"].astype(str)

# ===================== 2. 情感极性&强度分析 =====================
# 定义情感分析函数（含反讽词过滤，提升准确性）
def emotion_analysis(text):
    if pd.isna(text) or text.strip() == "" or len(text) < 5:  # 过滤短评论
        return {"polarity": "中性", "score": 0.5}
    
    # 反讽词库（可根据数据补充）
    irony_words = ["呵呵", "真厉害", "绝了（反讽）", "就这", "不愧是你（反讽）"]
    for word in irony_words:
        if word in text:
            return {"polarity": "负向", "score": 0.2}
    
    s = SnowNLP(text)
    score = s.sentiments  # 情感得分（0-1，越接近1越正向）
    if score >= 0.6:
        polarity = "正向"
    elif score <= 0.4:
        polarity = "负向"
    else:
        polarity = "中性"
    return {"polarity": polarity, "score": score}

# 批量分析赵丽颖粉丝情感
zhao_df["情感极性"] = zhao_df["评论内容"].apply(lambda x: emotion_analysis(x)["polarity"])
zhao_df["情感得分"] = zhao_df["评论内容"].apply(lambda x: emotion_analysis(x)["score"])

# 批量分析洛天依粉丝情感
luo_df["情感极性"] = luo_df["评论内容"].apply(lambda x: emotion_analysis(x)["polarity"])
luo_df["情感得分"] = luo_df["评论内容"].apply(lambda x: emotion_analysis(x)["score"])

# 统计情感极性占比（核心对比数据）
zhao_polarity = zhao_df["情感极性"].value_counts(normalize=True) * 100
luo_polarity = luo_df["情感极性"].value_counts(normalize=True) * 100

# 统计情感强度均值
zhao_score_mean = zhao_df["情感得分"].mean()
luo_score_mean = luo_df["情感得分"].mean()

# ===================== 3. 情感强度细化分析（强情感词+感叹号） =====================
# 定义强情感词库
positive_strong = ["封神", "永远支持", "绝了", "最好", "心疼", "守护", "超棒", "惊艳"]
negative_strong = ["失望透顶", "拉胯", "尴尬", "糟糕", "崩溃", "差评", "难看"]

# 统计单条评论强情感词数量
def count_strong_words(text):
    if pd.isna(text):
        return 0
    count = 0
    for word in positive_strong + negative_strong:
        count += text.count(word)
    return count

# 统计感叹号数量（含中英文）
def count_exclamation(text):
    if pd.isna(text):
        return 0
    return text.count("!") + text.count("！")

# 批量计算
zhao_df["强情感词数"] = zhao_df["评论内容"].apply(count_strong_words)
luo_df["强情感词数"] = luo_df["评论内容"].apply(count_strong_words)

zhao_df["感叹号数"] = zhao_df["评论内容"].apply(count_exclamation)
luo_df["感叹号数"] = luo_df["评论内容"].apply(count_exclamation)

# 统计均值
zhao_strong_word_mean = zhao_df["强情感词数"].mean()
luo_strong_word_mean = luo_df["强情感词数"].mean()

zhao_excla_mean = zhao_df["感叹号数"].mean()
luo_excla_mean = luo_df["感叹号数"].mean()

# ===================== 4. 情感焦点关键词提取 =====================
# 加载停用词表（需提前下载哈工大停用词表，保存为stopwords.txt）
try:
    with open("stopwords.txt", "r", encoding="utf-8") as f:
        stopwords = set(f.read().splitlines())
except:
    # 若没有停用词表，用默认停用词
    stopwords = {"的", "了", "是", "我", "你", "他", "她", "它", "我们", "你们", "他们", "这", "那", "在", "有", "就", "都", "而", "及", "与"}

# 文本清洗+分词函数
def clean_and_cut(text):
    if pd.isna(text):
        return []
    # 去除特殊字符、数字、字母
    text = re.sub(r"[^\u4e00-\u9fa5]", "", text)
    # 分词
    words = jieba.lcut(text)
    # 过滤停用词、短词、无意义词
    words = [w for w in words if w not in stopwords and len(w) > 1 and w not in ["评论", "内容", "用户"]]
    return words

# 提取正向评论的关键词（聚焦核心情感）
zhao_positive = zhao_df[zhao_df["情感极性"] == "正向"]["评论内容"].apply(clean_and_cut)
luo_positive = luo_df[luo_df["情感极性"] == "正向"]["评论内容"].apply(clean_and_cut)

# 统计Top20关键词
zhao_word_count = Counter([word for sublist in zhao_positive for word in sublist])
luo_word_count = Counter([word for sublist in luo_positive for word in sublist])

# ===================== 5. 输出所有结果 =====================
print("="*60)
print("【赵丽颖粉丝情感分析结果】")
print("="*60)
print("1. 情感极性占比：")
print(zhao_polarity.round(1))  # 保留1位小数
print(f"2. 情感强度均值：{zhao_score_mean:.2f}")
print(f"3. 强情感词平均出现次数：{zhao_strong_word_mean:.2f}")
print(f"4. 感叹号平均出现次数：{zhao_excla_mean:.2f}")
print("5. 正向评论Top20关键词：")
for word, count in zhao_word_count.most_common(20):
    print(f"   {word}: {count}次")

print("\n" + "="*60)
print("【洛天依粉丝情感分析结果】")
print("="*60)
print("1. 情感极性占比：")
print(luo_polarity.round(1))
print(f"2. 情感强度均值：{luo_score_mean:.2f}")
print(f"3. 强情感词平均出现次数：{luo_strong_word_mean:.2f}")
print(f"4. 感叹号平均出现次数：{luo_excla_mean:.2f}")
print("5. 正向评论Top20关键词：")
for word, count in luo_word_count.most_common(20):
    print(f"   {word}: {count}次")

# ===================== 6. 导出分析结果到Excel（方便后续可视化） =====================
zhao_df.to_excel("赵丽颖情感分析结果.xlsx", index=False, encoding="utf-8")
luo_df.to_excel("洛天依情感分析结果.xlsx", index=False, encoding="utf-8")
print("\n✅ 分析结果已导出为Excel文件：")
print("   - 赵丽颖情感分析结果.xlsx")
print("   - 洛天依情感分析结果.xlsx")
