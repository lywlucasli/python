import pyperclip
import re

# 1. 从剪贴板读取原始文本
raw = pyperclip.paste()

# 2. 提取所有连续数字串（工单编号）
#    \d+ 匹配一个或多个数字，按出现顺序返回
ticket_numbers = re.findall(r'\d+', raw)

# 3. 拼接成每行一个编号的字符串
result = '\n'.join(ticket_numbers)

# 4. 复制到剪贴板
pyperclip.copy(result)

# 5. 提示并预览
print("✅ 解析完成，所有工单编号已复制到剪贴板！")
print(f"共提取 {len(ticket_numbers)} 个编号")
print("预览（前200字符）：")
print(result[:200] + ("..." if len(result) > 200 else ""))