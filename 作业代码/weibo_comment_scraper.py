import time
import os
from openpyxl import Workbook

# 获取桌面路径（适配Windows系统）
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
print("=== 微博评论批量整理（保存到桌面Excel版） ===")
print("操作步骤：")
print("1. 在微博评论区（PC网页版），Ctrl+A全选评论，Ctrl+C复制")
print("2. 在此处Ctrl+V粘贴所有内容")
print("3. 新起一行输入“完成”，按回车生成Excel文件")
print()
print("请粘贴评论内容：")

# 读取粘贴内容，直到输入“完成”
all_lines = []
while True:
    try:
        line = input()
        if line.strip() == "完成":
            break
        if line.strip():
            all_lines.append(line.strip())
    except:
        break

print(f"接收到 {len(all_lines)} 行文本，正在处理...")

# 识别有效微博评论（适配微博格式：用户名+时间+内容）
comments = []
current_comment = ""
for line in all_lines:
    line = line.strip()
    if not line:
        continue
    
    # 判定新评论：微博评论特征（用户名+“·”+时间，如“张三·1小时前”）
    is_new_comment = (
        # 匹配“用户名·时间”格式（如“李四·昨天 12:30”）
        ("·" in line and any(key in line for key in ["前", "天", "月", "年", ":"]))
        # 补充匹配点赞数特征（如“100赞”）
        or ("赞" in line and line[:3].isdigit())
    )
    
    if is_new_comment:
        if current_comment and len(current_comment) > 5:
            comments.append(current_comment)
        current_comment = line
    else:
        current_comment = current_comment + " " + line if current_comment else line

# 保存最后一条评论
if current_comment and len(current_comment) > 5:
    comments.append(current_comment)

print(f"成功识别 {len(comments)} 条评论")

# 生成Excel文件并保存到桌面
if comments:
    # 生成带时间戳的文件名
    filename = f"微博评论_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_path = os.path.join(desktop_path, filename)
    
    # 创建Excel工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "微博评论"
    
    # 写入表头
    ws.append(['序号', '评论内容', '评论级别', '采集时间'])
    
    # 写入评论数据（区分一级/二级评论）
    for i, com in enumerate(comments, 1):
        level = "二级评论" if ("回复" in com or "@" in com) else "一级评论"
        ws.append([i, com, level, time.strftime("%Y-%m-%d %H:%M:%S")])
    
    # 保存文件
    wb.save(excel_path)
    print(f"\nExcel文件已保存到桌面：{filename}")
    print(f"文件路径：{excel_path}")
    
    # 预览前10条评论
    print("\n前10条评论预览：")
    print("-" * 50)
    for i, com in enumerate(comments[:10], 1):
        preview = com[:50] + "..." if len(com) > 50 else com
        print(f"{i}. {preview}")
    if len(comments) > 10:
        print(f"... 还有 {len(comments)-10} 条评论")
else:
    print("\n未识别到有效评论，请重新复制评论区内容（避开广告/按钮文字）")

input("\n按回车键退出...")
