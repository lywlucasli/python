#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import sys
import pyperclip  # 需要 pip install pyperclip

def expand_boxes(box_str: str):
    """解析箱号字符串，返回数字列表（支持 - 连续区间）"""
    # 先去除首尾空格，再按常见分隔符拆分
    parts = re.split('[、,;|/]+', box_str.strip())
    boxes = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        nums = re.findall(r'\d+', part)
        if not nums:
            continue
        if '-' in part and len(nums) >= 2:
            start = int(nums[0])
            end = int(nums[1])
            if start > end:
                start, end = end, start
            boxes.extend(range(start, end + 1))
        else:
            boxes.append(int(nums[0]))
    return boxes

def process_line(line: str):
    """处理一行，返回该行对应的所有完整箱号列表（只含箱号字符串）"""
    line = line.strip()
    if not line:
        return []
    # 尝试制表符分割，否则按第一个空格分割
    if '\t' in line:
        parts = line.split('\t', 1)
    else:
        parts = line.split(maxsplit=1)
    if len(parts) < 2:
        return []
    
    so = parts[0].strip()          # 去除首尾空格
    box_str = parts[1].strip()     # 去除首尾空格
    
    if not so or not box_str:
        return []
    
    box_numbers = expand_boxes(box_str)
    so_fmt = f"A{so.zfill(12)}"    # 工单号补足12位（若so本身超过12位则不补）
    return [so_fmt + f"C{str(num).zfill(4)}" for num in box_numbers]

def main():
    try:
        clipboard_text = pyperclip.paste()
    except Exception as e:
        print(f"读取剪贴板失败: {e}", file=sys.stderr)
        sys.exit(1)
    
    if not clipboard_text.strip():
        print("剪贴板为空，请复制两列数据后再运行。", file=sys.stderr)
        sys.exit(1)

    lines = clipboard_text.splitlines()
    all_boxes = []
    for line in lines:
        all_boxes.extend(process_line(line))

    if all_boxes:
        output = "\n".join(all_boxes)
        pyperclip.copy(output)          # 复制到剪贴板
        # 屏幕打印也只输出箱号，不附加任何说明
        for box in all_boxes:
            print(box)
    else:
        # 无有效数据时，剪贴板保持不变，屏幕输出仅提示（不影响实际复制）
        print("未生成任何箱号，请检查数据格式。", file=sys.stderr)

if __name__ == "__main__":
    main()