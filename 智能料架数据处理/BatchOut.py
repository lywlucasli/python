import requests
import time
import sys
import re
# ==================== 配置区 ====================
BASE_URL = "http://192.168.9.100:8081"
SEARCH_URL = f"{BASE_URL}/mate/searchByInfo"
OUT_URL = f"{BASE_URL}/mate/outByPosition"
# 需要处理的仓库类型ID列表（可自由增删）
WAREHOUSE_NAME_IDS = [43, 44]   # 根据实际需求修改 看服务器。
# 搜索基础参数（part_num 或 save_id 会在 main 中动态替换，其他参数固定）
SEARCH_BASE = {
    "part_num": "",                     # 占位，由用户输入
    "warehouse_category_id": 1,
    "save_id": "",                      # 占位，由用户输入
    "lot_code": "",
    "mfg_date_start": "",
    "supplier_name": "",
    "mfg_date_end": "",
    "start": 0,
    "rows": 1000,                       # 每页 1000 条
    "user_name": "NAF684",
    "user_group_power": "1",
    "user_language": "zh",
    "shelf_id": "",
    "position_info": "",
    "start_date": "",
    "end_date": ""
}
# 出库固定参数（color=3 表示出库）
OUT_TEMPLATE = {
    "color": 3,
    "user_name": "NAF684",
    "user_group_power": "1",
    "user_language": "zh"
}
# 请求间隔（秒），避免过快
REQUEST_INTERVAL = 0.2
# ==================== 核心函数 ====================
def get_materials_by_save_id(save_id_value, warehouse_ids, search_base):
    """
    根据 save_id 中的编号获取所有匹配的物料
    :param save_id_value: save_id 中的编号（如 0011555445）
    :param warehouse_ids: 仓库类型ID列表
    :param search_base: 搜索基础参数
    :return: 物料列表
    """
    materials = []
    
    # 复制一份基础参数
    payload = search_base.copy()
    payload["save_id"] = save_id_value
    payload["part_num"] = ""  # 清空 part_num
    
    for warehouse_name_id in warehouse_ids:
        payload["warehouse_name_id"] = warehouse_name_id
        start = 0
        rows = payload.get("rows", 1000)
        
        print(f"  Searching warehouse_name_id={warehouse_name_id}, save_id={save_id_value}")
        
        while True:
            payload["start"] = start
            try:
                resp = requests.post(SEARCH_URL, json=payload, timeout=10)
            except Exception as e:
                print(f"    Request exception: {e}")
                break
            if resp.status_code != 200:
                print(f"    HTTP {resp.status_code}, stop fetching")
                break
            data = resp.json()
            if data.get("result") != 1:
                print(f"    API error response: {data}, stop fetching")
                break
            mate_list = data.get("mate_list", [])
            if not mate_list:
                break
            materials.extend(mate_list)
            total = data.get("rows", 0)
            start += rows
            if start >= total:
                break
            time.sleep(0.1)  # 分页间隔
    
    return materials

def get_materials_by_part_num(part_num, warehouse_ids, search_base):
    """
    根据 part_num 获取所有匹配的物料
    :param part_num: part_num 值
    :param warehouse_ids: 仓库类型ID列表
    :param search_base: 搜索基础参数
    :return: 物料列表
    """
    materials = []
    
    # 复制一份基础参数
    payload = search_base.copy()
    payload["part_num"] = part_num
    payload["save_id"] = ""  # 清空 save_id
    
    for warehouse_name_id in warehouse_ids:
        payload["warehouse_name_id"] = warehouse_name_id
        start = 0
        rows = payload.get("rows", 1000)
        
        print(f"  Searching warehouse_name_id={warehouse_name_id}, part_num={part_num}")
        
        while True:
            payload["start"] = start
            try:
                resp = requests.post(SEARCH_URL, json=payload, timeout=10)
            except Exception as e:
                print(f"    Request exception: {e}")
                break
            if resp.status_code != 200:
                print(f"    HTTP {resp.status_code}, stop fetching")
                break
            data = resp.json()
            if data.get("result") != 1:
                print(f"    API error response: {data}, stop fetching")
                break
            mate_list = data.get("mate_list", [])
            if not mate_list:
                break
            materials.extend(mate_list)
            total = data.get("rows", 0)
            start += rows
            if start >= total:
                break
            time.sleep(0.1)  # 分页间隔
    
    return materials

def out_material(shelf_id, position, out_template):
    """
    调用出库接口
    :return: (是否成功, 消息)
    """
    payload = {
        "shelf_id": shelf_id,
        "position": position,
        "color": out_template["color"],
        "user_name": out_template["user_name"],
        "user_group_power": out_template["user_group_power"],
        "user_language": out_template["user_language"]
    }
    try:
        resp = requests.post(OUT_URL, json=payload, timeout=10)
    except Exception as e:
        return False, f"Request exception: {e}"
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"
    data = resp.json()
    if data.get("result") == 1:
        return True, "Success"
    else:
        return False, f"API result != 1: {data}"

def batch_out(materials, out_template, operation_name=""):
    """
    批量出库
    """
    if not materials:
        print(f"  {operation_name}: No materials to outbound")
        return 0, 0, []
    success = 0
    fail = 0
    fail_details = []
    total = len(materials)
    print(f"\n  Start batch outbound [{operation_name}], total {total} materials")
    for idx, mat in enumerate(materials, start=1):
        shelf_id = mat.get("shelf_id")
        position = mat.get("position")
        if shelf_id is None or position is None:
            fail += 1
            fail_details.append(f"Index{idx}: Missing shelf_id or position")
            continue
        ok, msg = out_material(shelf_id, position, out_template)
        if ok:
            success += 1
            print(f"  [{idx}/{total}] ✓ Outbound OK: shelf_id={shelf_id} position={position}")
        else:
            fail += 1
            fail_details.append(f"[{idx}] shelf_id={shelf_id} position={position} Failed: {msg}")
            print(f"  [{idx}/{total}] ✗ Outbound FAILED: shelf_id={shelf_id} position={position} -> {msg}")
        time.sleep(REQUEST_INTERVAL)
    return success, fail, fail_details

def main():
    print("===== Batch Outbound Tool (support part_num / multiple save_id) =====")
    # ---------- 选择搜索模式 ----------
    print("\nPlease select search mode:")
    print("1. Search by single part_num")
    print("2. Search by multiple save_id (one per line)")
    
    mode_choice = input("Please input option (1 or 2): ").strip()
    
    all_materials = []
    operation_names = []
    
    if mode_choice == "1":
        # 模式1：按 part_num
        search_mode = "part_num"
        if len(sys.argv) > 1:
            search_value = sys.argv[1]
            print(f"Read part_num from argument = {search_value}")
        else:
            search_value = input("Please input part_num (example: 23-132499-030-R007): ").strip()
            if not search_value:
                print("Error: part_num is empty, exit.")
                sys.exit(1)
        
        print(f"\nSearch mode: part_num = {search_value}")
        print(f"Target warehouse_name_id list: {WAREHOUSE_NAME_IDS}")
        
        # 获取物料
        materials = get_materials_by_part_num(search_value, WAREHOUSE_NAME_IDS, SEARCH_BASE)
        if materials:
            all_materials.extend(materials)
            operation_names.append(f"part_num={search_value}")
        
    elif mode_choice == "2":
        # 模式2：按多个 save_id 编号
        search_mode = "save_id"
        print("\nPlease input save_id list (example: 0011555445), one per line")
        print("End input: press Ctrl+D (Linux/Mac) or Ctrl+Z (Windows) then Enter")
        print("Or input blank line and press Enter to finish")
        
        save_id_list = []
        while True:
            try:
                line = input().strip()
                if not line:
                    break
                save_id_list.append(line)
            except EOFError:
                break
        
        if not save_id_list:
            print("Error: No save_id input, exit.")
            sys.exit(1)
        
        print(f"\nTotal {len(save_id_list)} save_id loaded:")
        for sid in save_id_list:
            print(f"  - {sid}")
        
        print(f"\nTarget warehouse_name_id list: {WAREHOUSE_NAME_IDS}")
        
        # 逐个处理每个 save_id
        for idx, save_id_value in enumerate(save_id_list, 1):
            print(f"\n--- Processing {idx}/{len(save_id_list)} save_id: {save_id_value} ---")
            materials = get_materials_by_save_id(save_id_value, WAREHOUSE_NAME_IDS, SEARCH_BASE)
            if materials:
                all_materials.extend(materials)
                operation_names.append(f"save_id={save_id_value}")
                print(f"  Found {len(materials)} materials")
            else:
                print(f"  No matched materials")
        
    else:
        print("Error: Invalid option, please input 1 or 2")
        sys.exit(1)
    # 检查是否获取到物料
    if not all_materials:
        print("\nNo materials found, please check search parameters or warehouse ID")
        sys.exit(1)
    print(f"\nTotal loaded materials: {len(all_materials)}")
    print(f"Operations: {', '.join(operation_names)}")
    # 显示部分物料信息供确认
    if all_materials:
        print("\nFirst 10 materials for confirmation:")
        for i, mat in enumerate(all_materials[:10], 1):
            print(f"  {i}. shelf_id={mat.get('shelf_id')}, position={mat.get('position')}, "
                  f"part_num={mat.get('part_num')}, save_id={mat.get('save_id')}")
    # 2. 执行批量出库
    confirm = input("\nConfirm to execute outbound? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Operation cancelled by user")
        sys.exit(0)
    # 执行出库
    success, fail, details = batch_out(all_materials, OUT_TEMPLATE, "All materials")
    # 3. 输出汇总
    print("\n===== Outbound Summary =====")
    print(f"Success: {success}")
    print(f"Failed : {fail}")
    if fail > 0:
        print("Failure details:")
        for d in details:
            print(f"  {d}")

if __name__ == "__main__":
    main()
