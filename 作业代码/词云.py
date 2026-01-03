# 终极版：双偶像过滤+JPG蒙版+高频情感对比
import pandas as pd
import jieba
import re
import os
import numpy as np
from PIL import Image
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# ---------------------- 全局配置（JPG蒙版+路径） ----------------------
ZLY_MASK = r"C:\Users\GHS\Desktop\爱心.png"  # JPG格式
LTY_MASK = r"C:\Users\GHS\Desktop\爱心.png"
ZLY_EXCEL = "赵丽颖评论爬取.xlsx"
LTY_EXCEL = "洛天依评论爬取.xlsx"
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")

# ---------------------- 1. 情感词+专属过滤词配置 ----------------------
# 核心情感词表
emotion_dict = {
    "正面情感": ["喜欢", "爱", "热爱", "开心", "快乐", "感动", "暖心", "惊艳", "优秀", "棒", "好", "完美", "赞", "支持", "认可", "欣赏", "可爱", "好听", "过瘾", "值得", "骄傲", "甜蜜", "上头", "本命", "入坑", "治愈", "温柔"],
    "负面情感": ["失望", "难过", "不满", "讨厌", "差", "不好", "遗憾", "吐槽", "无语", "生气", "伤心", "无奈"],
    "态度倾向": ["期待", "希望", "觉得", "认为", "感觉", "想要", "愿意", "应该"]
}
# 双偶像专属过滤词（剔除命名类无效词）
filter_dict = {
    "赵丽颖": {"赵", "颖", },
    "洛天依": {"洛天", "洛", "依",}
}
# 基础停用词
stopwords_basic = {"的", "了", "在", "是", "我", "你", "他", "这", "那", "和", "也", "都", "只", "又","真的", "这个", "就是", "一直", "觉得","一个", "这么","但是", "不是", "越来越", "可以", "天依来","回复","没有"}

# ---------------------- 2. 统一数据预处理（双偶像过滤） ----------------------
def process_data(excel_path, idol_name):
    """通用预处理：适配双偶像过滤+情感词提取"""
    df = pd.read_excel(excel_path)
    comments = df["评论内容"].dropna().astype(str)
    
    # 1. 清洗+分词+专属过滤
    all_words = []
    idol_filter = filter_dict[idol_name]
    for c in comments:
        c_clean = re.sub(r"[^\u4e00-\u9fa5]", "", c)
        words = jieba.lcut(c_clean)
        for w in words:
            # 过滤规则：非基础停用词+非偶像命名词+长度≥2
            if w not in stopwords_basic and w not in idol_filter and len(w)>=2:
                all_words.append(w)
    
    # 2. 统计情感词频次
    emotion_freq = {"正面情感": {}, "负面情感": {}, "态度倾向": {}}
    total_emotion_words = []
    for emo_type, emo_words in emotion_dict.items():
        for w in emo_words:
            cnt = all_words.count(w)
            if cnt > 0:
                emotion_freq[emo_type][w] = cnt
                total_emotion_words.extend([w]*cnt)
    
    # 3. 生成词云文本（过滤低频词）
    word_freq = {}
    for w in all_words:
        word_freq[w] = word_freq.get(w, 0) + 1
    final_words = [w for w in all_words if word_freq[w] >= 2]  # 高频核心词
    text = ' '.join(final_words)
    
    print(f"✅ {idol_name}数据处理完成：")
    print(f"   核心词汇数：{len(final_words)} | 情感词总数：{len(total_emotion_words)}")
    return text, emotion_freq

# 处理双偶像数据
zly_text, zly_emo_freq = process_data(ZLY_EXCEL, "赵丽颖")
lt_text, lt_emo_freq = process_data(LTY_EXCEL, "洛天依")

# ---------------------- 3. JPG蒙版适配（双偶像通用） ----------------------
def fix_jpg_mask(mask_path):
    """处理JPG蒙版，解决压缩杂色问题"""
    if not os.path.exists(mask_path):
        print(f"⚠️  未找到蒙版：{mask_path}")
        return None
    img = Image.open(mask_path).convert("RGB")
    mask = np.array(img)
    # JPG专用二值化阈值（适配压缩杂色）
    mask_gray = np.sum(mask, axis=2)
    mask_binary = np.where(mask_gray > 550, 255, 0)
    return mask_binary

zly_mask = fix_jpg_mask(ZLY_MASK)
lt_mask = fix_jpg_mask(LTY_MASK)

# ---------------------- 4. 高频情感词对比条形图 ----------------------
def plot_emotion_bar():
    """生成双偶像TOP10正面情感词对比条形图"""
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    # 提取TOP10高频正面情感词
    def get_top10(emo_freq):
        top = sorted(emo_freq["正面情感"].items(), key=lambda x: x[1], reverse=True)[:10]
        return [w[0] for w in top], [w[1] for w in top]
    
    zly_words, zly_counts = get_top10(zly_emo_freq)
    lt_words, lt_counts = get_top10(lt_emo_freq)
    
    # 绘制对比图
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
    # 赵丽颖
    ax1.barh(zly_words, zly_counts, color='#FF6B6B', alpha=0.8)
    ax1.set_title("赵丽颖粉丝TOP10正面情感词", fontsize=14, fontweight='bold')
    ax1.set_xlabel("出现频次", fontsize=12)
    ax1.grid(axis='x', alpha=0.3)
    # 洛天依
    ax2.barh(lt_words, lt_counts, color='#66CCFF', alpha=0.8)
    ax2.set_title("洛天依粉丝TOP10正面情感词", fontsize=14, fontweight='bold')
    ax2.set_xlabel("出现频次", fontsize=12)
    ax2.grid(axis='x', alpha=0.3)
    
    plt.suptitle("虚拟vs真实偶像粉丝高频正面情感词对比", fontsize=16, fontweight='bold', y=1.02)
    plt.tight_layout()
    bar_save = os.path.join(DESKTOP, "高频情感词对比条形图.png")
    plt.savefig(bar_save, bbox_inches='tight', dpi=300)
    plt.show()
    print(f"✅ 高频情感词对比条形图已保存：{bar_save}")

plot_emotion_bar()

# ---------------------- 5. 双偶像超饱满词云（无无效词） ----------------------
font_path = 'C:/Windows/Fonts/simhei.ttf'

# 赵丽颖词云（专属过滤+JPG蒙版+暖色调）
zly_wc = WordCloud(
    font_path=font_path,
    background_color='white',
    mask=zly_mask,
    max_words=300,          # 词汇量拉满
    random_state=42,
    contour_width=2,
    contour_color='#FF6B6B',
    prefer_horizontal=0.8,  # 80%水平词，填充更满
    relative_scaling=0.9,   # 词频关联度最大化
    font_step=1,            # 字号梯度最小
    collocations=False,     # 关闭词汇搭配
    scale=2,                # 分辨率翻倍
    color_func=lambda *args, **kwargs: '#FF6B6B'
).generate(zly_text)
zly_save = os.path.join(DESKTOP, "赵丽颖_情感词云_无无效词_JPG.png")
zly_wc.to_file(zly_save)
print(f"✅ 赵丽颖词云已保存：{zly_save}")

# 洛天依词云（专属过滤+JPG蒙版+冷色调）
lt_wc = WordCloud(
    font_path=font_path,
    background_color='white',
    mask=lt_mask,
    max_words=300,
    random_state=42,
    contour_width=2,
    contour_color='#66CCFF',
    prefer_horizontal=0.8,
    relative_scaling=0.9,
    font_step=1,
    collocations=False,
    scale=2,
    color_func=lambda *args, **kwargs: '#66CCFF'
).generate(lt_text)
lt_save = os.path.join(DESKTOP, "洛天依_情感词云_无空版_JPG.png")
lt_wc.to_file(lt_save)
print(f"✅ 洛天依词云已保存：{lt_save}")

# ---------------------- 6. 核心统计结果 ----------------------
print("\n" + "="*60)
print("📊 双偶像情感统计结果：")
def sum_emo(emo_freq):
    return {t: sum(freq.values()) for t, freq in emo_freq.items()}

zly_total = sum_emo(zly_emo_freq)
lt_total = sum_emo(lt_emo_freq)

print(f"赵丽颖：正面{zly_total['正面情感']} | 负面{zly_total['负面情感']} | 态度{zly_total['态度倾向']}")
print(f"洛天依：正面{lt_total['正面情感']} | 负面{lt_total['负面情感']} | 态度{lt_total['态度倾向']}")
print("="*60)
print("\n🎉 全部完成！桌面文件：")
print("   1. 高频情感词对比条形图.png（TOP10正面情感词对比）")
print("   2. 赵丽颖_情感词云_无无效词_JPG.png（无命名词+超饱满）")
print("   3. 洛天依_情感词云_无空版_JPG.png（无命名词+填充拉满）")
