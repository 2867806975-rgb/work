# 最终完美版：桌面保存+高密度+纯净情感词云
import pandas as pd
import jieba
import re
import os
import numpy as np
from PIL import Image
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# ---------------------- 1. 核心配置（直接保存到桌面） ----------------------
# 桌面路径（自动获取，无需手动改）
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
# 蒙版路径（替换为你的实际路径）
ZLY_MASK = r"C:\Users\GHS\Desktop\赵丽颖蒙版_处理后.jpg"
LTY_MASK = r"C:\Users\GHS\Desktop\洛天依蒙版_处理后.jpg"
# Excel文件路径（确保和代码同文件夹）
ZLY_EXCEL = "赵丽颖评论爬取.xlsx"
LTY_EXCEL = "洛天依评论爬取.xlsx"

# ---------------------- 2. 纯净情感词表（无任何无用词） ----------------------
core_emotion_words = {
    # 正面情感（补充更多，增加填充量）
    "喜欢", "爱", "热爱", "开心", "快乐", "感动", "暖心", "惊艳", "优秀", "棒", "好", "完美", 
    "赞", "支持", "认可", "欣赏", "可爱", "好听", "过瘾", "值得", "骄傲", "甜蜜", "上头", 
    "本命", "入坑", "治愈", "温柔", "心疼", "佩服", "满意", "幸福", "给力", "惊艳", "圈粉", "心动",
    # 负面情感
    "失望", "难过", "不满", "讨厌", "差", "不好", "遗憾", "吐槽", "无语", "生气", "伤心", "无奈",
    # 态度倾向
    "期待", "希望", "觉得", "认为", "感觉", "想要", "愿意", "应该", "盼望", "向往", "憧憬"
}
# 剔除命名类词
name_filter = {"赵丽颖", "颖宝", "赵", "颖", "洛天", "洛天依", "天依", "洛", "依"}

# ---------------------- 3. 数据处理（提升填充密度） ----------------------
def process_emotion_data(excel_path):
    df = pd.read_excel(excel_path)
    comments = df["评论内容"].dropna().astype(str)
    pure_words = []
    for c in comments:
        c_clean = re.sub(r"[^\u4e00-\u9fa5]", "", c)
        words = jieba.lcut(c_clean)
        # 只保留情感词，且放宽频次（至少1次），增加词汇量
        for w in words:
            if w in core_emotion_words and w not in name_filter:
                pure_words.append(w)
    # 重复词汇，提升填充密度（核心！解决空的问题）
    dense_words = pure_words * 3  # 词汇量翻3倍，填充更满
    return ' '.join(dense_words), pure_words

# 处理数据
zly_text, zly_origin = process_emotion_data(ZLY_EXCEL)
lt_text, lt_origin = process_emotion_data(LTY_EXCEL)
print(f"✅ 赵丽颖情感词总数（含重复）：{len(zly_text.split())}")
print(f"✅ 洛天依情感词总数（含重复）：{len(lt_text.split())}")

# ---------------------- 4. 蒙版处理（优化形状贴合） ----------------------
def fix_mask(mask_path):
    if not os.path.exists(mask_path):
        print(f"⚠️  蒙版未找到：{mask_path}，生成矩形高密度词云")
        return None
    img = Image.open(mask_path).convert("L")
    mask = np.array(img)
    mask_binary = np.where(mask > 200, 255, 0)
    return mask_binary

zly_mask = fix_mask(ZLY_MASK)
lt_mask = fix_mask(LTY_MASK)

# ---------------------- 5. 生成高密度纯净词云（保存到桌面） ----------------------
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False
font_path = 'C:/Windows/Fonts/msyh.ttc'  # 微软雅黑，更美观

# --- 赵丽颖词云（高密度+暖色调） ---
zly_wc = WordCloud(
    font_path=font_path,
    background_color='white',
    mask=zly_mask,
    max_words=500,          # 最大词汇量拉满
    random_state=42,
    contour_width=1,
    contour_color='#FF6B6B',
    prefer_horizontal=0.6,  # 60%水平+40%垂直，填充更密
    relative_scaling=0.5,   # 字号差异减小，填充更均匀
    font_step=1,            # 字号梯度最小
    collocations=False,
    color_func=lambda *args, **kwargs: np.random.choice(['#FF6B6B', '#FF8E8E', '#FFA8A8'])  # 渐变暖色
).generate(zly_text)

# 保存到桌面
zly_save = os.path.join(DESKTOP, "赵丽颖_高密度纯净情感词云.png")
zly_wc.to_file(zly_save)
# 强制显示
plt.figure(figsize=(10, 8))
plt.imshow(zly_wc, interpolation='bilinear')
plt.axis('off')
plt.title("赵丽颖粉丝高密度纯净情感词云", fontsize=14, fontweight='bold')
plt.show()
print(f"✅ 赵丽颖词云已保存到桌面：{zly_save}")

# --- 洛天依词云（高密度+冷色调） ---
lt_wc = WordCloud(
    font_path=font_path,
    background_color='white',
    mask=lt_mask,
    max_words=500,
    random_state=42,
    contour_width=1,
    contour_color='#66CCFF',
    prefer_horizontal=0.6,
    relative_scaling=0.5,
    font_step=1,
    collocations=False,
    color_func=lambda *args, **kwargs: np.random.choice(['#66CCFF', '#87CEEB', '#B0E0E6'])  # 渐变冷色
).generate(lt_text)

# 保存到桌面
lt_save = os.path.join(DESKTOP, "洛天依_高密度纯净情感词云.png")
lt_wc.to_file(lt_save)
# 强制显示
plt.figure(figsize=(10, 8))
plt.imshow(lt_wc, interpolation='bilinear')
plt.axis('off')
plt.title("洛天依粉丝高密度纯净情感词云", fontsize=14, fontweight='bold')
plt.show()
print(f"✅ 洛天依词云已保存到桌面：{lt_save}")

# ---------------------- 6. 生成高频情感词对比条形图（桌面保存） ----------------------
def plot_bar():
    # 统计原始词频（去重）
    def get_freq(words):
        freq = {}
        for w in words:
            freq[w] = freq.get(w, 0) + 1
        return sorted(freq.items(), key=lambda x: x[1], reverse=True)[:8]
    
    zly_top = get_freq(zly_origin)
    lt_top = get_freq(lt_origin)
    zly_words = [w[0] for w in zly_top]
    zly_counts = [w[1] for w in zly_top]
    lt_words = [w[0] for w in lt_top]
    lt_counts = [w[1] for w in lt_top]
    
    # 绘制高颜值条形图
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    # 赵丽颖
    ax1.barh(zly_words, zly_counts, color=['#FF6B6B', '#FF8E8E', '#FFA8A8', '#FFC0CB', '#FFD1DC', '#FFE4E1', '#FFF0F5', '#F8C8DC'])
    ax1.set_title("赵丽颖粉丝核心情感词TOP8", fontsize=13, fontweight='bold')
    ax1.set_xlabel("出现频次")
    ax1.grid(axis='x', alpha=0.2)
    # 洛天依
    ax2.barh(lt_words, lt_counts, color=['#66CCFF', '#87CEEB', '#B0E0E6', '#E0FFFF', '#F0F8FF', '#87CEFA', '#ADD8E6', '#B0C4DE'])
    ax2.set_title("洛天依粉丝核心情感词TOP8", fontsize=13, fontweight='bold')
    ax2.set_xlabel("出现频次")
    ax2.grid(axis='x', alpha=0.2)
    
    plt.suptitle("虚拟vs真实偶像粉丝核心情感词对比", fontsize=15, fontweight='bold', y=1.02)
    plt.tight_layout()
    # 保存到桌面
    bar_save = os.path.join(DESKTOP, "高频情感词对比条形图_高颜值.png")
    plt.savefig(bar_save, dpi=300, bbox_inches='tight')
    plt.show()
    print(f"✅ 条形图已保存到桌面：{bar_save}")

plot_bar()

# ---------------------- 最终提示 ----------------------
print("\n🎉 全部完成！所有文件已保存到你的电脑桌面：")
print("1. 赵丽颖_高密度纯净情感词云.png（无无用词+填充饱满）")
print("2. 洛天依_高密度纯净情感词云.png（无无用词+填充饱满）")
print("3. 高频情感词对比条形图_高颜值.png（核心情感词对比）")
