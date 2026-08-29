# -*- coding: utf-8 -*-
"""
仓库物料数据导出工具（提取标签编号 + 统计频次/数量 + Excel导出）
- 每页10条（接口原生支持）
- 失败自动重试3次
- 返回为空才停止，不依赖总数判断
- 自动从label_code中提取第二个#号之间的编号（RT）
- 增加两列：物料编号出现次数、RT对应数量总和
- 输出为.xlsx文件
"""
import requests
import time
from openpyxl import Workbook

# ========== 配置区 ==========
API_URL = "http://192.168.9.100:8081/mate/searchByInfo"
PAGE_SIZE = 10000           # 接口每页最大条数（实际可能限制为10，但保持原意）
OUTPUT_FILE = "mate_export.xlsx"
USER_NAME = "NAF684"
USER_GROUP_POWER = "1"
USER_LANGUAGE = "zh"
MAX_RETRY = 3
RETRY_DELAY = 2

EXPORT_FIELDS = [
    "in_time",
    "update_time",
    "shelf_id",
    "warehouse_name",
    "warehouse_category_name",
    "part_num",
    "quantity",
    "label_code",
    "position_info",
]

# 最终表头（原始列 + RT列 + 新增两列）
EXCEL_HEAD = [
    "入库时间", "更新时间", "货架编号", "仓库名称", "仓库类别",
    "物料编号", "数量", "标签编码", "库位信息",
    "RT",                           # 提取的编号
    "物料编号出现次数",              # 新增
    "RT对应数量总和"                 # 新增
]
# ============================

def build_payload(start, rows):
    return {
        "part_num": "",
        "warehouse_category_id": 0,
        "warehouse_name_id": 0,
        "save_id": "",
        "lot_code": "",
        "mfg_date_start": "",
        "supplier_name": "",
        "mfg_date_end": "",
        "start": start,
        "rows": rows,
        "user_name": USER_NAME,
        "user_group_power": USER_GROUP_POWER,
        "user_language": USER_LANGUAGE,
        "shelf_id": "",
        "position_info": "",
        "start_date": "",
        "end_date": "",
    }

def fetch_page(start):
    """请求单页，带重试"""
    payload = build_payload(start, PAGE_SIZE)
    for attempt in range(1, MAX_RETRY + 1):
        try:
            resp = requests.post(API_URL, json=payload, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            if data.get("result") == 1:
                return data
            else:
                print(f"  [WARN] 第{attempt}次 result!=1: {data}")
        except Exception as e:
            print(f"  [WARN] 第{attempt}次请求异常: {e}")
        if attempt < MAX_RETRY:
            time.sleep(RETRY_DELAY)
    return None

def extract_code_from_label(label):
    """从标签编码中提取第二个#号之间的编号，例如 '0014497575'"""
    if not label:
        return ""
    try:
        parts = label.split('#')
        if len(parts) >= 2:
            return parts[1].strip()
        else:
            return ""
    except Exception:
        return ""

def fetch_all():
    all_items = []    # 每行数据（列表），包含原始字段 + RT
    all_raw = []      # 原始记录（字典），用于统计
    start = 0
    total = None
    consecutive_empty = 0

    while True:
        data = fetch_page(start)
        if data is None:
            print(f"[ERROR] start={start} 重试{MAX_RETRY}次仍失败，停止")
            break

        mate_list = data.get("mate_list", [])

        if total is None:
            total = data.get("rows", 0)
            print(f"接口声明总记录数: {total}")

        if not mate_list:
            consecutive_empty += 1
            print(f"  [INFO] start={start} 返回空页 ({consecutive_empty}/3)")
            if consecutive_empty >= 3:
                print("连续3次空页，确认数据已拉完")
                break
            start += PAGE_SIZE
            time.sleep(0.5)
            continue

        consecutive_empty = 0

        for item in mate_list:
            all_raw.append(item)   # 保存原始字典
            # 构建行：基础字段
            row = [item.get(f, "") for f in EXPORT_FIELDS]
            # 提取RT
            label = item.get("label_code", "")
            rt = extract_code_from_label(label)
            row.append(rt)         # RT放在最后一列（索引9，因为原来有9个字段）
            all_items.append(row)

        print(f"已获取: {len(all_items)} 条 (本次返回 {len(mate_list)} 条, start={start})")

        if total > 0 and len(all_items) > total * 1.5:
            print(f"[WARN] 已获取 {len(all_items)} 条，超过声明总数 {total} 的1.5倍，强制停止")
            break

        start += PAGE_SIZE
        time.sleep(0.1)

    return all_items, all_raw, total

def export_excel(items, raw_records, headers):
    """统计并导出为.xlsx"""
    # ----- 统计频次和数量总和 -----
    part_num_count = {}
    rt_quantity_sum = {}

    for item in raw_records:
        part = item.get("part_num", "")
        qty_str = item.get("quantity", "0")
        try:
            qty = float(qty_str) if qty_str else 0.0
        except ValueError:
            qty = 0.0

        part_num_count[part] = part_num_count.get(part, 0) + 1

        rt = extract_code_from_label(item.get("label_code", ""))
        rt_quantity_sum[rt] = rt_quantity_sum.get(rt, 0.0) + qty

    # ----- 写入Excel -----
    wb = Workbook()
    ws = wb.active
    ws.title = "物料数据"
    ws.append(headers)   # 写入表头

    for row, item in zip(items, raw_records):
        # row目前包含原始9个字段 + RT（共10个元素）
        part = item.get("part_num", "")
        rt = row[-1]  # RT已在最后
        # 追加两列
        count = part_num_count.get(part, 0)
        sum_qty = rt_quantity_sum.get(rt, 0.0)
        row.append(count)
        row.append(sum_qty)
        ws.append(row)

    wb.save(OUTPUT_FILE)
    print(f"\n导出完成: {OUTPUT_FILE}  共 {len(items)} 条")

if __name__ == "__main__":
    data_rows, raw_data, total = fetch_all()
    export_excel(data_rows, raw_data, EXCEL_HEAD)
    if total > 0:
        print(f"对比: 接口声明 {total} 条, 实际导出 {len(data_rows)} 条")
        if len(data_rows) < total:
            print(f"[注意] 缺失 {total - len(data_rows)} 条，可能是中途请求失败，请检查上方日志")