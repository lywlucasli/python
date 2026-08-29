import sys
import subprocess
import importlib.util
import os
import math
from datetime import datetime

# ===================== 自动安装缺失库 =====================
def auto_install(package_name: str):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

if not importlib.util.find_spec("requests"):
    print("[AUTO-INSTALL] requests not found, installing...")
    auto_install("requests")
if not importlib.util.find_spec("openpyxl"):
    print("[AUTO-INSTALL] openpyxl not found, installing...")
    auto_install("openpyxl")
if not importlib.util.find_spec("pyperclip"):
    print("[AUTO-INSTALL] pyperclip not found, installing...")
    auto_install("pyperclip")

import hashlib
import requests
import pyperclip
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from typing import List, Dict, Tuple

# ===================== 全局配置 =====================
GLOBAL_COMPANY_ID = "6801"
USER_NAME = "N59596"
PASSWORD = "000000"

# ===================== MES登录 =====================
class MES:
    def encrypt_password(self, password: str) -> str:
        md5 = hashlib.md5()
        md5.update(password.encode('utf-8'))
        return md5.hexdigest()

    def login(self, account: str, password: str) -> str:
        print("[MES] 登录中...")
        encrypted_pwd = self.encrypt_password(password)
        url = "http://10.97.245.205:86/login"
        headers = {"Content-Type": "application/json;charset=UTF-8"}
        payload = {"account": account, "password": encrypted_pwd}
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=15)
            resp.raise_for_status()
            res_json = resp.json()
            if res_json.get("code") == 200 and "data" in res_json and "token" in res_json["data"]:
                token = res_json["data"]["token"]
                print("[MES] 登录成功")
                return token
            else:
                print(f"[MES] 登录失败: {res_json.get('info', '未知错误')}")
                sys.exit(1)
        except Exception as e:
            print(f"[MES] 登录异常: {e}")
            sys.exit(1)

# ===================== BOM下载（扩展：返回物料名称、规格） =====================
def fetch_work_order_bom(token: str, work_order: str) -> List[Dict]:
    """
    下载指定工单的BOM明细，返回列表，每个元素包含：
        partno_id, qty, uom, partno_name, specification
    """
    url = (
        f"http://10.97.245.205:92/prod_baseapi/mes/wobomdetail/page"
        f"?rows=10000&page=1"
        f"&sidx=partno_id+asc,WORK_ORDER+DESC,rspos+asc"
        f"&sord="
        f"&company_id={GLOBAL_COMPANY_ID}"
        f"&work_order={work_order}"
        f"&lead_work_order="
        f"&partno_id="
    )
    headers = {
        "token": token,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    out_rows = []
    try:
        resp = requests.get(url, headers=headers, timeout=20)
        resp.raise_for_status()
        js = resp.json()
        if js.get("code") != 200:
            print(f"[BOM-API] 工单 {work_order} 错误: {js.get('info')}")
            return []
        rows = js.get("data", {}).get("rows", [])
        for row in rows:
            partno_id = row.get("partno_id", "").strip()
            req_qty = float(row.get("request_qty", 0.0))
            sobkz = row.get("sobkz", "")
            dumps = row.get("dumps", "")
            dbskz = row.get("dbskz", "")
            # 原有过滤逻辑
            if req_qty <= 0:
                continue
            if sobkz == "Q" and dumps == "X":
                continue
            if dbskz == "E":
                continue
            uom = row.get("uom", "EA")
            # 尝试获取物料名称和规格（字段名可能不同，此处列举常见）
            partno_name = row.get("partno_name") or row.get("description") or row.get("part_desc") or ""
            specification = row.get("specification") or row.get("spec") or row.get("model") or row.get("part_spec") or ""
            out_rows.append({
                "partno_id": partno_id,
                "qty": req_qty,
                "work_order": work_order,
                "uom": uom,
                "partno_name": partno_name,
                "specification": specification
            })
        print(f"[BOM-API] 工单 {work_order} 有效行数: {len(out_rows)}")
        return out_rows
    except Exception as e:
        print(f"[BOM-API] 工单 {work_order} 异常: {e}")
        return []

# ===================== 解析剪贴板（工单号 数量） =====================
def parse_wo_qty_clip(raw_text: str) -> List[Tuple[str, float]]:
    """解析剪贴板中的工单号和数量，返回 [(wo, qty), ...]"""
    pairs = []
    for line in raw_text.strip().splitlines():
        parts = line.strip().split()
        if len(parts) >= 2:
            wo = parts[0].strip()
            try:
                qty = float(parts[1].strip())
                pairs.append((wo, qty))
            except ValueError:
                continue
    return pairs

# ===================== 汇总BOM，计算单位用量 =====================
def aggregate_bom(bom_rows: List[Dict], wo_qty: float) -> List[Dict]:
    """
    按料号汇总总需求，计算单位用量（总需求 / 工单数量，向上取整）
    返回列表，每个元素包含 partno, name, spec, unit_qty
    """
    agg = {}
    for row in bom_rows:
        pn = row["partno_id"]
        if pn not in agg:
            agg[pn] = {
                "qty": 0.0,
                "name": row.get("partno_name", ""),
                "spec": row.get("specification", "")
            }
        agg[pn]["qty"] += row["qty"]
    result = []
    for pn, data in agg.items():
        unit_qty = math.ceil(data["qty"] / wo_qty)   # 向上取整
        result.append({
            "partno": pn,
            "name": data["name"],
            "spec": data["spec"],
            "unit_qty": unit_qty
        })
    # 按料号排序（可选）
    result.sort(key=lambda x: x["partno"])
    return result

# ===================== 生成Excel（简洁样式） =====================
def generate_single_bom_excel(wo: str, wo_qty: float, agg_rows: List[Dict], filepath: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"

    # 表头
    headers = ["序号", "物料编号", "物料名称", "规格型号", "单位用量", "位置号"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # 数据行
    for i, row in enumerate(agg_rows, 1):
        ws.cell(row=i+1, column=1, value=i)
        ws.cell(row=i+1, column=2, value=row["partno"])
        ws.cell(row=i+1, column=3, value=row["name"])
        ws.cell(row=i+1, column=4, value=row["spec"])
        ws.cell(row=i+1, column=5, value=row["unit_qty"])
        ws.cell(row=i+1, column=6, value="")   # 位置号留空

    # 自动调整列宽
    for col_idx in range(1, 7):
        col_letter = chr(64 + col_idx)
        ws.column_dimensions[col_letter].width = 20

    # 设置边框和对齐
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wb.save(filepath)
    print(f"[EXCEL] 生成 {filepath} (共 {len(agg_rows)} 行)")

# ===================== 主入口 =====================
if __name__ == "__main__":
    print("===== 工单BOM直接下载工具 =====")
    raw_clip = pyperclip.paste().strip()
    if not raw_clip:
        print("[错误] 剪贴板为空，请复制工单号和数量（每行：工单号 数量）")
        sys.exit(1)

    # 解析剪贴板
    wo_qty_pairs = parse_wo_qty_clip(raw_clip)
    if not wo_qty_pairs:
        print("[错误] 未能解析出有效的工单号和数量，请检查格式（空格或Tab分隔）")
        sys.exit(1)
    print(f"[解析] 共识别 {len(wo_qty_pairs)} 个工单")

    # 登录MES
    mes = MES()
    token = mes.login(USER_NAME, PASSWORD)

    # 逐个处理
    success_count = 0
    for wo, qty in wo_qty_pairs:
        print(f"\n处理工单: {wo}  数量: {qty}")
        bom_raw = fetch_work_order_bom(token, wo)
        if not bom_raw:
            print(f"[跳过] 工单 {wo} 无BOM数据")
            continue
        agg = aggregate_bom(bom_raw, qty)
        if not agg:
            print(f"[跳过] 工单 {wo} 汇总后无有效数据")
            continue
        # 生成文件
        qty_int = int(qty) if qty.is_integer() else qty
        filename = f"{wo}_{qty_int}.xlsx"
        generate_single_bom_excel(wo, qty, agg, filename)
        success_count += 1

    print(f"\n===== 完成 =====")
    print(f"成功生成 {success_count} 个Excel文件，文件名格式：工单号_数量.xlsx")
    input("\n按 Enter 键退出...")