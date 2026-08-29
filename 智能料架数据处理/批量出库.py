import requests
import time
import sys

# ==================== 配置区 ====================
BASE_URL = "http://192.168.9.100:8081"
SEARCH_URL = f"{BASE_URL}/mate/searchByInfo"
OUT_URL = f"{BASE_URL}/mate/outByPosition"

# 需要处理的仓库类型ID列表（可自由增删）
WAREHOUSE_NAME_IDS = [43, 44]   # 根据实际需求修改 看服务器。

# 搜索基础参数（part_num 会在 main 中动态替换，其他参数固定）
SEARCH_BASE = {
    "part_num": "",                     # 占位，由用户输入
    "warehouse_category_id": 1,
    "save_id": "",
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

def get_materials_for_warehouse(warehouse_name_id, search_base):
    """
    针对单个 warehouse_name_id 分页获取所有物料
    :param warehouse_name_id: 仓库类型ID
    :param search_base: 搜索基础参数字典（会被复制并修改）
    :return: 该仓库下的物料列表
    """
    materials = []
    start = 0
    rows = search_base.get("rows", 1000)

    # 复制一份基础参数，避免影响其他仓库的循环
    payload = search_base.copy()
    payload["warehouse_name_id"] = warehouse_name_id

    print(f"开始获取 warehouse_name_id={warehouse_name_id} 的物料...")

    while True:
        payload["start"] = start
        try:
            resp = requests.post(SEARCH_URL, json=payload, timeout=10)
        except Exception as e:
            print(f" 请求异常: {e}")
            break

        if resp.status_code != 200:
            print(f" HTTP {resp.status_code}，停止获取")
            break

        data = resp.json()
        if data.get("result") != 1:
            print(f" 返回错误: {data}，停止获取")
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

    print(f" 共获取 {len(materials)} 条物料")
    return materials


def get_all_materials(warehouse_ids, search_base):
    """
    遍历所有仓库类型，合并物料列表
    :param warehouse_ids: 仓库类型ID列表
    :param search_base: 搜索基础参数
    :return: 全部物料列表
    """
    all_materials = []
    for wid in warehouse_ids:
        part = get_materials_for_warehouse(wid, search_base)
        all_materials.extend(part)
    return all_materials


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
        return False, f"请求异常: {e}"

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"

    data = resp.json()
    if data.get("result") == 1:
        return True, "成功"
    else:
        return False, f"返回结果非1: {data}"


def batch_out(materials, out_template):
    """
    批量出库
    """
    success = 0
    fail = 0
    fail_details = []
    total = len(materials)

    print(f"\n开始批量出库，共 {total} 条物料")

    for idx, mat in enumerate(materials, start=1):
        shelf_id = mat.get("shelf_id")
        position = mat.get("position")
        if shelf_id is None or position is None:
            fail += 1
            fail_details.append(f"索引{idx}: 缺少 shelf_id 或 position")
            continue

        ok, msg = out_material(shelf_id, position, out_template)
        if ok:
            success += 1
            print(f"[{idx}/{total}] ✓ 出库成功: {shelf_id} 位置 {position}")
        else:
            fail += 1
            fail_details.append(f"[{idx}] {shelf_id} 位置 {position} 失败: {msg}")
            print(f"[{idx}/{total}] ✗ 出库失败: {shelf_id} 位置 {position} -> {msg}")

        time.sleep(REQUEST_INTERVAL)

    return success, fail, fail_details


def main():
    print("===== 批量出库脚本（多仓库类型，可自定义 part_num） =====")

    # ---------- 获取 part_num ----------
    if len(sys.argv) > 1:
        part_num = sys.argv[1]
        print(f"从命令行参数读取 part_num = {part_num}")
    else:
        part_num = input("请输入要出库的 part_num（例如 23-132499-030-R007）: ").strip()
        if not part_num:
            print("错误：未输入 part_num，退出。")
            sys.exit(1)

    # 更新搜索参数
    SEARCH_BASE["part_num"] = part_num
    print(f"待处理仓库类型ID: {WAREHOUSE_NAME_IDS}")

    # 1. 获取所有物料
    materials = get_all_materials(WAREHOUSE_NAME_IDS, SEARCH_BASE)
    if not materials:
        print("未获取到任何物料，请检查搜索参数或仓库ID")
        sys.exit(1)

    print(f"总计获取物料 {len(materials)} 条")

    # 2. 执行批量出库
    success, fail, details = batch_out(materials, OUT_TEMPLATE)

    # 3. 输出汇总
    print("\n===== 出库汇总 =====")
    print(f"成功: {success}")
    print(f"失败: {fail}")
    if fail > 0:
        print("失败详情:")
        for d in details:
            print(f"  {d}")


if __name__ == "__main__":
    main()