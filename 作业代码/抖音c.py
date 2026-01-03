import time
import csv
import re

print("=== 抖音全部评论采集（仅提取评论内容）===")
print("📝 操作步骤：")
print("1. 打开抖音视频评论区，下拉加载所有需要采集的评论")
print("2. Ctrl+A 全选评论区内容 → Ctrl+C 复制")
print("3. 回到本程序窗口，Ctrl+V 粘贴所有内容")
print("4. 新起一行输入「完成」，按回车开始处理")
print()
print("请粘贴评论内容（粘贴后输入「完成」结束）：")

# 1. 接收用户粘贴的内容
all_lines = []
while True:
    try:
        line = input()
        if line.strip() == "完成":
            break
        if line.strip():  # 跳过空行
            all_lines.append(line.strip())
    except (EOFError, KeyboardInterrupt):
        break

print(f"📊 接收到 {len(all_lines)} 行文本，正在提取所有评论内容...")

# 2. 核心逻辑：提取所有纯评论内容（包括一级和二级）
# 定义需要过滤的无关信息正则
# 匹配规则：
# - 以「@用户名」开头的（回复标识，保留后面的评论内容）
# - 包含「赞」「万」的数字（点赞数，如「123赞」「4.5万赞」）
# - 纯数字/数字+符号（可能是抖音号或时间）
# - 表情符号
irrelevant_pattern = re.compile(
    r"^\d+(\.\d+)?万?赞$|^\d+$|[\U00010000-\U0010ffff]",
    re.UNICODE
)

# 合并所有行，按标点符号分割（解决评论内容换行的问题）
merged_text = " ".join(all_lines)
# 按常见标点分割（。！？；：，、）），保留分割符
potential_comments = re.split(r"([。！？；：，、）])", merged_text)
# 重组分割后的内容（将分割符还原到评论末尾）
comments_with_punct = []
for i in range(0, len(potential_comments), 2):
    comment_part = potential_comments[i]
    punct_part = potential_comments[i+1] if i+1 < len(potential_comments) else ""
    if comment_part.strip():
        comments_with_punct.append(comment_part.strip() + punct_part)

# 过滤无效内容，保留纯评论
pure_comments = []
seen = set()  # 去重

for item in comments_with_punct:
    # 过滤条件：
    # - 不匹配无关信息正则（排除点赞数、纯数字、表情等）
    # - 长度≥5（排除过短的无效内容）
    # - 不是空字符串
    if (
        not irrelevant_pattern.match(item) 
        and len(item.strip()) >= 5
        and item.strip() != ""
    ):
        # 去除开头的@用户名（如果有），保留后面的评论内容
        cleaned_comment = re.sub(r"^@\w+\s*", "", item.strip())
        if cleaned_comment and cleaned_comment not in seen:
            seen.add(cleaned_comment)
            pure_comments.append(cleaned_comment)

# 3. 输出结果
print(f"✅ 成功提取 {len(pure_comments)} 条有效评论（含一级和二级）")

if pure_comments:
    # 生成CSV文件（仅包含序号和评论内容）
    filename = f"抖音全部评论_纯内容_{time.strftime('%Y%m%d_%H%M%S')}.csv"
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["序号", "评论内容"])  # 仅保留两个核心字段
        for idx, comment in enumerate(pure_comments, 1):
            writer.writerow([idx, comment])

    print(f"💾 文件已保存至：{filename}")
    print("\n📋 前10条评论预览：")
    print("-" * 60)
    for i, comment in enumerate(pure_comments[:10], 1):
        print(f"{i:2d}. {comment}")
    if len(pure_comments) > 10:
        print(f"... 共 {len(pure_comments)} 条评论，其余内容已保存至文件")
    print("-" * 60)
else:
    print("❌ 未识别到有效评论，请检查粘贴内容是否包含完整的评论区信息")

input("按回车键退出...")
