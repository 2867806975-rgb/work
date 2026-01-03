import pandas as pd
import jieba
import json
import re
import os

# ---------------------- 第一步：自动定位Excel文件（桌面/当前文件夹） ----------------------
def find_excel_file(file_name: str) -> str:
    """
    自动查找Excel文件（先查桌面，再查当前文件夹）
    :param file_name: 文件名（如"赵丽颖评论爬取.xlsx"）
    :return: 完整路径/空字符串（未找到）
    """
    # 桌面路径
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    desktop_path = os.path.join(desktop, file_name)
    if os.path.exists(desktop_path):
        print(f"✅ 找到文件：{desktop_path}")
        return desktop_path
    
    # 当前文件夹路径
    current_path = os.path.join(os.getcwd(), file_name)
    if os.path.exists(current_path):
        print(f"✅ 找到文件：{current_path}")
        return current_path
    
    # 未找到，提示用户
    print(f"❌ 未找到文件：{file_name}")
    print("📌 请确认文件存在于以下位置之一：")
    print(f"   1. 桌面：{desktop}")
    print(f"   2. 代码所在文件夹：{os.getcwd()}")
    return ""

# 自动查找你的Excel文件（无需手动改路径）
zly_excel = find_excel_file("赵丽颖评论爬取.xlsx")
lty_excel = find_excel_file("洛天依评论爬取.xlsx")

# ---------------------- 第二步：数据预处理（容错） ----------------------
def load_and_clean_comments(excel_path: str, idol_name: str) -> list:
    if not excel_path:
        return []
    try:
        df = pd.read_excel(excel_path)
        # 兼容不同列名（评论/评论内容/粉丝评论）
        comment_cols = [col for col in df.columns if "评论" in col]
        if not comment_cols:
            print(f"❌ {idol_name}Excel中无“评论”相关列")
            return []
        comments = df[comment_cols[0]].dropna().astype(str).tolist()
    except Exception as e:
        print(f"❌ 读取{idol_name}Excel失败：{str(e)[:50]}")
        return []
    
    # 清洗规则
    stop_names = {idol_name, "颖宝", "天依", "殿下", "赵姐", "洛殿"}
    clean_comments = []
    for c in comments:
        # 剔除非中文+去空格
        c_clean = re.sub(r"[^\u4e00-\u9fa5]", "", c).strip()
        # 剔除偶像命名
        for name in stop_names:
            c_clean = c_clean.replace(name, "")
        # 过滤短文本
        if len(c_clean) >= 3:
            clean_comments.append(c_clean)
    print(f"✅ {idol_name}有效评论数：{len(clean_comments)}")
    return clean_comments

# 加载并清洗数据
zly_comments = load_and_clean_comments(zly_excel, "赵丽颖")
lty_comments = load_and_clean_comments(lty_excel, "洛天依")
all_comments = zly_comments + lty_comments

if not all_comments:
    print("❌ 无有效评论数据，程序退出")
    exit()

# ---------------------- 第三步：离线实体/关系抽取（无外网依赖） ----------------------
# 预设核心实体库（贴合你的场景）
PRESET_ENTITIES = {
    "赵丽颖": {"type": "人物", "sub_type": "真实偶像"},
    "洛天依": {"type": "人物", "sub_type": "虚拟偶像"},
    "花千骨": {"type": "作品"},
    "风吹半夏": {"type": "作品"},
    "演唱会": {"type": "行为"},
    "歌曲": {"type": "作品"},
    "演技": {"type": "特质"},
    "声音": {"type": "特质"}
}

# 预设情感词库
EMOTION_WORDS = {
    "喜欢", "爱", "热爱", "开心", "快乐", "感动", "暖心", "惊艳", "优秀", "棒", "好", "完美",
    "赞", "支持", "认可", "欣赏", "可爱", "好听", "过瘾", "值得", "骄傲", "甜蜜", "治愈",
    "失望", "难过", "不满", "讨厌", "差", "不好", "遗憾", "吐槽", "无语", "生气"
}

# 抽取实体和三元组
entities = []  # 格式：[(实体名, 实体类型), ...]
triples = []   # 格式：[(头实体, 关系, 尾实体), ...]

for comment in all_comments:
    # 1. 匹配偶像实体
    for idol, info in PRESET_ENTITIES.items():
        if idol in comment and idol in ["赵丽颖", "洛天依"]:
            # 添加偶像实体
            entities.append((idol, info["type"]))
            entities.append((info["sub_type"], "偶像类型"))
            
            # 2. 匹配情感词，生成情感关系
            for emo in EMOTION_WORDS:
                if emo in comment:
                    triples.append(("粉丝", emo, idol))
                    entities.append((emo, "情感词"))
            
            # 3. 匹配作品/特质，生成关联关系
            for entity, e_info in PRESET_ENTITIES.items():
                if entity in comment and entity not in ["赵丽颖", "洛天依"]:
                    entities.append((entity, e_info["type"]))
                    # 生成关系（根据实体类型）
                    if e_info["type"] == "作品":
                        triples.append((idol, "关联", entity))
                    elif e_info["type"] == "特质":
                        triples.append((idol, "拥有", entity))

# 去重（避免重复实体/关系）
entities = list(set(entities))
triples = list(set([tuple(t) for t in triples]))

# ---------------------- 第四步：保存结果到桌面（可视化用） ----------------------
# 桌面路径
desktop = os.path.join(os.path.expanduser("~"), "Desktop")

# 保存实体
entity_path = os.path.join(desktop, "知识图谱_实体.json")
with open(entity_path, "w", encoding="utf-8") as f:
    json.dump(entities, f, ensure_ascii=False, indent=2)

# 保存三元组
triple_path = os.path.join(desktop, "知识图谱_三元组.json")
with open(triple_path, "w", encoding="utf-8") as f:
    json.dump(triples, f, ensure_ascii=False, indent=2)

# ---------------------- 结果提示 ----------------------
print("\n🎉 离线抽取完成！文件已保存到桌面：")
print(f"1. 实体文件：{entity_path}")
print(f"2. 三元组文件：{triple_path}")
print("\n📌 实体示例（前5个）：")
for e in entities[:5]:
    print(f"   - {e[0]}（{e[1]}）")
print("\n📌 三元组示例（前5个）：")
for t in triples[:5]:
    print(f"   - {t[0]} → {t[1]} → {t[1]}")
