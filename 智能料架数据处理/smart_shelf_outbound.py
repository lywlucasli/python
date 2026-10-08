# -*- coding: utf-8 -*-
"""
Batch Outbound Tool (tkinter version v6)

Features:
1. Query stock: 只按配置的 warehouse_name_id 分别分页拉取并合并
   - 仓库 ID 可在界面编辑，默认 43,44，自动保存/读取
2. Extract RT + PN from shelf data
3. By PN: fuzzy search + outbound by selected PN
   - 出库前会重新向服务器查询该 PN（只查配置的仓库 ID）
4. By RT: paste RTs -> validate -> show matched rows ->
          user can select rows (single / multi / select-all) then outbound
   - 出库前校验编辑框 ID 与内存一致，不一致则拦截
5. WiFi: manual switch
6. Export Stock: export current shelf data to Excel (English headers)
7. UI style: white + blue
"""

import os
import re
import sys
import json
import time
import threading
import subprocess
from collections import Counter, defaultdict
from typing import List, Optional, Tuple

import requests
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

# ==================== Config ====================
BASE_URL = "http://192.168.9.100:8081"
SEARCH_URL = f"{BASE_URL}/mate/searchByInfo"
OUT_URL = f"{BASE_URL}/mate/outByPosition"

# 默认目标仓库类型 ID（可被配置文件覆盖）
DEFAULT_WAREHOUSE_NAME_IDS = [43, 44]

PAGE_SIZE = 10000
USER_NAME = "NAF684"
USER_GROUP_POWER = "1"
USER_LANGUAGE = "zh"
REQUEST_INTERVAL = 0.2
MAX_RETRY = 3
RETRY_DELAY = 2

# ---- 本地配置 ----
CONFIG_FILENAME = "batch_outbound_config.json"

SEARCH_BASE = {
    "part_num": "",
    "warehouse_category_id": 1,
    "warehouse_name_id": "",   # 运行时按仓库填入
    "save_id": "",
    "lot_code": "",
    "mfg_date_start": "",
    "supplier_name": "",
    "mfg_date_end": "",
    "start": 0,
    "rows": PAGE_SIZE,
    "user_name": USER_NAME,
    "user_group_power": USER_GROUP_POWER,
    "user_language": USER_LANGUAGE,
    "shelf_id": "",
    "position_info": "",
    "start_date": "",
    "end_date": "",
}

OUT_TEMPLATE = {
    "color": 3,
    "user_name": USER_NAME,
    "user_group_power": USER_GROUP_POWER,
    "user_language": USER_LANGUAGE,
}

# ---- Excel export ----
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

EXCEL_HEAD = [
    "In Time",
    "Update Time",
    "Shelf ID",
    "Warehouse Name",
    "Warehouse Category",
    "Part Number",
    "Quantity",
    "Label Code",
    "Position Info",
    "RT",
    "Part Num Occurrences",
    "RT Total Qty",
]

# ---- WiFi ----
WIFI_PROFILE_NAME = "AOI"
WIFI_PASSWORD     = "123456789"
WIFI_PROFILE_XML  = f"""<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{WIFI_PROFILE_NAME}</name>
    <SSIDConfig>
        <SSID>
            <name>{WIFI_PROFILE_NAME}</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>{WIFI_PASSWORD}</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"""

# ---- 配色 ----
COLOR_BG        = "#FFFFFF"
COLOR_BG_ALT    = "#F5F9FF"
COLOR_PRIMARY   = "#1976D2"
COLOR_PRIMARY_D = "#0D47A1"
COLOR_ACCENT    = "#E3F2FD"
COLOR_BORDER    = "#BBDEFB"
COLOR_TEXT      = "#1A1A1A"
COLOR_MUTED     = "#607D8B"
COLOR_SUCCESS   = "#2E7D32"
COLOR_WARN      = "#EF6C00"
COLOR_ERROR     = "#C62828"

FONT_BASE  = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_SUB   = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)


# ==================== Config File ====================
def config_path() -> str:
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base = os.getcwd()
    return os.path.join(base, CONFIG_FILENAME)


def load_config() -> dict:
    default = {"warehouse_name_ids": list(DEFAULT_WAREHOUSE_NAME_IDS)}
    path = config_path()
    if not os.path.isfile(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        ids = data.get("warehouse_name_ids")
        clean = []
        if isinstance(ids, list):
            for x in ids:
                try:
                    clean.append(int(x))
                except (TypeError, ValueError):
                    pass
        data["warehouse_name_ids"] = clean or list(DEFAULT_WAREHOUSE_NAME_IDS)
        return data
    except Exception:
        return default


def save_config(cfg: dict) -> bool:
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def parse_warehouse_ids(text: str) -> List[int]:
    """
    解析用户输入：支持 '43,44' / '43 44' / '43;44' / '43，44' 等。
    返回去重后的 int 列表（保持输入顺序）。空输入返回 []。
    """
    if not text:
        return []
    tokens = re.split(r"[\s,;，、]+", text.strip())
    result = []
    seen = set()
    for t in tokens:
        if not t:
            continue
        try:
            v = int(t)
        except ValueError:
            raise ValueError(f"Invalid warehouse ID: {t}")
        if v not in seen:
            seen.add(v)
            result.append(v)
    return result


# ==================== WiFi Manager ====================
class WifiManager:
    @staticmethod
    def _run(args, timeout=15) -> str:
        try:
            out = subprocess.run(args, capture_output=True, timeout=timeout)
            return (out.stdout or b"").decode("gbk", errors="ignore")
        except Exception:
            return ""

    @staticmethod
    def get_current_ssid() -> str:
        out = WifiManager._run(["netsh", "wlan", "show", "interfaces"])
        for line in out.splitlines():
            if "SSID" in line and "BSSID" not in line:
                parts = line.split(":", 1)
                if len(parts) == 2:
                    return parts[1].strip()
        return ""

    @staticmethod
    def _get_state_and_ssid() -> Tuple[str, str]:
        out = WifiManager._run(["netsh", "wlan", "show", "interfaces"])
        state = ""
        ssid = ""
        for line in out.splitlines():
            if "State" in line or "状态" in line:
                if ":" in line:
                    state = line.split(":", 1)[1].strip().lower()
            if "SSID" in line and "BSSID" not in line:
                if ":" in line:
                    ssid = line.split(":", 1)[1].strip()
        return state, ssid

    @staticmethod
    def list_profiles() -> List[str]:
        out = WifiManager._run(["netsh", "wlan", "show", "profiles"])
        profiles = []
        for line in out.splitlines():
            if ":" in line and ("All User Profile" in line
                                or "所有用户配置文件" in line
                                or "User Profile" in line
                                or "用户配置文件" in line):
                profiles.append(line.split(":", 1)[1].strip())
        return profiles

    @staticmethod
    def ensure_profile() -> bool:
        profiles = WifiManager.list_profiles()
        if WIFI_PROFILE_NAME in profiles:
            return True
        try:
            from pathlib import Path
            xml_path = Path(__file__).resolve().parent / "_wifi_profile.xml"
            xml_path.write_text(WIFI_PROFILE_XML, encoding="utf-8")
            WifiManager._run(
                ["netsh", "wlan", "add", "profile",
                 f"filename={xml_path}", "user=all"],
                timeout=15
            )
            try: xml_path.unlink()
            except Exception: pass
            return WIFI_PROFILE_NAME in WifiManager.list_profiles()
        except Exception:
            return False

    @staticmethod
    def disconnect():
        WifiManager._run(["netsh", "wlan", "disconnect"], timeout=10)
        time.sleep(1.5)

    @staticmethod
    def connect(profile_name: str, timeout: int = 30) -> bool:
        state, ssid = WifiManager._get_state_and_ssid()
        if ssid == profile_name and state == "connected":
            return True

        WifiManager.disconnect()

        try:
            subprocess.run(
                ["netsh", "wlan", "connect", f"name={profile_name}"],
                check=False, timeout=15
            )
        except Exception:
            return False

        t0 = time.time()
        last_retry = t0
        while time.time() - t0 < timeout:
            state, ssid = WifiManager._get_state_and_ssid()
            if ssid == profile_name and state == "connected":
                time.sleep(1.5)
                state2, ssid2 = WifiManager._get_state_and_ssid()
                if ssid2 == profile_name and state2 == "connected":
                    return True
            if time.time() - last_retry > 15:
                last_retry = time.time()
                try:
                    subprocess.run(
                        ["netsh", "wlan", "connect", f"name={profile_name}"],
                        check=False, timeout=15
                    )
                except Exception:
                    pass
            time.sleep(0.5)
        return False


# ==================== Helpers ====================
def extract_rt_from_label(label: str) -> str:
    if not label:
        return ""
    try:
        parts = label.split("#")
        return parts[1].strip() if len(parts) >= 2 else ""
    except Exception:
        return ""


def build_search_payload(start, rows=PAGE_SIZE, warehouse_name_id=None):
    payload = SEARCH_BASE.copy()
    payload["start"] = start
    payload["rows"] = rows
    if warehouse_name_id is not None:
        payload["warehouse_name_id"] = warehouse_name_id
    return payload


# ==================== Network ====================
def fetch_page(start, warehouse_name_id=None, logger=None):
    payload = build_search_payload(start, warehouse_name_id=warehouse_name_id)
    for attempt in range(1, MAX_RETRY + 1):
        try:
            resp = requests.post(SEARCH_URL, json=payload, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            if data.get("result") == 1:
                return data
            if logger:
                logger(f"  [WARN] attempt {attempt} result != 1: {data}")
        except Exception as e:
            if logger:
                logger(f"  [WARN] attempt {attempt} request error: {e}")
        if attempt < MAX_RETRY:
            time.sleep(RETRY_DELAY)
    return None


def fetch_all_materials(warehouse_ids: List[int],
                        log_cb=None, progress_cb=None):
    """
    按 warehouse_name_id 分别分页拉取并合并。
    接口不支持数组，因此逐个 ID 请求。
    返回 (materials, any_ok)
    """
    if not warehouse_ids:
        if log_cb:
            log_cb("[ERROR] warehouse_ids is empty")
        return [], False

    all_materials = []
    any_ok = False

    for wh_id in warehouse_ids:
        if log_cb:
            log_cb(f"---- Fetching warehouse_name_id={wh_id} ----")

        start = 0
        total = None
        consecutive_empty = 0
        wh_ok = True
        wh_count_before = len(all_materials)

        while True:
            data = fetch_page(start, warehouse_name_id=wh_id, logger=log_cb)
            if data is None:
                if log_cb:
                    log_cb(f"[ERROR] wh={wh_id} start={start} failed after "
                           f"{MAX_RETRY} retries, stop this warehouse")
                wh_ok = False
                break

            mate_list = data.get("mate_list", [])

            if total is None:
                total = data.get("rows", 0)
                if log_cb:
                    log_cb(f"  Server reported total records (wh={wh_id}): {total}")

            if not mate_list:
                consecutive_empty += 1
                if log_cb:
                    log_cb(f"  [INFO] wh={wh_id} start={start} empty page "
                           f"({consecutive_empty}/3)")
                if consecutive_empty >= 3:
                    if log_cb:
                        log_cb(f"  wh={wh_id}: 3 consecutive empty pages, done")
                    break
                start += PAGE_SIZE
                time.sleep(0.5)
                continue

            consecutive_empty = 0
            all_materials.extend(mate_list)
            if log_cb:
                log_cb(f"  wh={wh_id} loaded {len(all_materials) - wh_count_before} "
                       f"records (this page {len(mate_list)}, start={start})")

            if total > 0 and (len(all_materials) - wh_count_before) > total * 1.5:
                if log_cb:
                    log_cb(f"  [WARN] wh={wh_id} loaded > 1.5x declared total, "
                           f"force stop")
                break

            start += PAGE_SIZE
            time.sleep(0.1)

        if wh_ok:
            any_ok = True
        if log_cb:
            log_cb(f"---- warehouse_name_id={wh_id} done, "
                   f"{len(all_materials) - wh_count_before} records "
                   f"(ok={wh_ok}) ----")

    if log_cb:
        log_cb(f"All target warehouses combined: {len(all_materials)} records")

    if progress_cb:
        progress_cb(len(all_materials), len(all_materials))

    return all_materials, any_ok


def search_materials_by_part_num(part_num, warehouse_ids: List[int], logger=None):
    """
    按 part_num 精确搜索目标仓库的物料（用于出库）。
    只查传入的 warehouse_ids。
    """
    materials = []
    for warehouse_name_id in warehouse_ids:
        payload = build_search_payload(0, warehouse_name_id=warehouse_name_id)
        payload["part_num"] = part_num
        payload["save_id"] = ""
        start = 0
        rows = payload["rows"]

        while True:
            payload["start"] = start
            try:
                resp = requests.post(SEARCH_URL, json=payload, timeout=10)
            except Exception as e:
                if logger:
                    logger(f"    Request exception: {e}")
                break
            if resp.status_code != 200:
                break
            data = resp.json()
            if data.get("result") != 1:
                break
            mate_list = data.get("mate_list", [])
            if not mate_list:
                break
            materials.extend(mate_list)
            total = data.get("rows", 0)
            start += rows
            if start >= total:
                break
            time.sleep(0.1)
    return materials


def out_material(shelf_id, position):
    payload = {
        "shelf_id": shelf_id,
        "position": position,
        "color": OUT_TEMPLATE["color"],
        "user_name": OUT_TEMPLATE["user_name"],
        "user_group_power": OUT_TEMPLATE["user_group_power"],
        "user_language": OUT_TEMPLATE["user_language"],
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


# ==================== Excel Export ====================
def export_to_excel(raw_records: List[dict], output_path: str,
                    log_cb=None) -> Tuple[bool, str]:
    if not HAS_OPENPYXL:
        return False, "openpyxl not installed, cannot export"
    if not raw_records:
        return False, "No data to export"

    part_num_count = {}
    rt_quantity_sum = {}

    for item in raw_records:
        part = item.get("part_num", "")
        qty_str = item.get("quantity", "0")
        try:
            qty = float(qty_str) if qty_str not in (None, "") else 0.0
        except (TypeError, ValueError):
            qty = 0.0

        part_num_count[part] = part_num_count.get(part, 0) + 1

        rt = extract_rt_from_label(item.get("label_code", ""))
        rt_quantity_sum[rt] = rt_quantity_sum.get(rt, 0.0) + qty

    wb = Workbook()
    ws = wb.active
    ws.title = "Stock"
    ws.append(EXCEL_HEAD)

    for item in raw_records:
        base_row = [item.get(f, "") for f in EXPORT_FIELDS]
        rt = extract_rt_from_label(item.get("label_code", ""))
        count = part_num_count.get(item.get("part_num", ""), 0)
        sum_qty = rt_quantity_sum.get(rt, 0.0)
        base_row.append(rt)
        base_row.append(count)
        base_row.append(sum_qty)
        ws.append(base_row)

    wb.save(output_path)
    if log_cb:
        log_cb(f"Exported {len(raw_records)} rows -> {output_path}")
    return True, output_path


# ==================== Main Window ====================
class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Batch Outbound Tool 批量出库工具")
        self.root.configure(bg=COLOR_BG)
        self._setup_styles()
        self._center_window(1180, 820)
        self.root.minsize(1040, 720)

        # 读取本地配置
        self.config_data = load_config()
        self.warehouse_ids: List[int] = list(self.config_data.get(
            "warehouse_name_ids", DEFAULT_WAREHOUSE_NAME_IDS))

        self.all_materials = []
        self.part_stats = []
        self.rt_map = defaultdict(list)
        self._busy = False
        self._original_wifi = ""

        self._build_ui()
        self._refresh_wifi_label()

    # ---------------- Styles ----------------
    def _setup_styles(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except Exception:
            pass

        s.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=FONT_BASE)
        s.configure("TFrame", background=COLOR_BG)
        s.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=FONT_BASE)
        s.configure("Muted.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=FONT_SUB)
        s.configure("Status.TLabel", background=COLOR_BG_ALT,
                    foreground=COLOR_PRIMARY_D, font=FONT_BOLD)

        s.configure("Card.TLabelframe", background=COLOR_BG_ALT,
                    bordercolor=COLOR_BORDER, relief="solid", borderwidth=1)
        s.configure("Card.TLabelframe.Label", background=COLOR_BG_ALT,
                    foreground=COLOR_PRIMARY_D, font=FONT_BOLD)

        s.configure("TEntry", fieldbackground=COLOR_BG, foreground=COLOR_TEXT,
                    bordercolor=COLOR_BORDER, insertcolor=COLOR_PRIMARY)

        s.configure("Primary.TButton",
                    background=COLOR_PRIMARY, foreground="white",
                    font=FONT_BOLD, padding=(12, 8),
                    borderwidth=0, focusthickness=0)
        s.map("Primary.TButton",
              background=[("active", COLOR_PRIMARY_D), ("pressed", COLOR_PRIMARY_D)],
              foreground=[("active", "white")])

        s.configure("Secondary.TButton",
                    background=COLOR_ACCENT, foreground=COLOR_PRIMARY_D,
                    font=FONT_BASE, padding=(10, 7),
                    borderwidth=0, focusthickness=0)
        s.map("Secondary.TButton",
              background=[("active", COLOR_BORDER), ("pressed", COLOR_BORDER)],
              foreground=[("active", COLOR_PRIMARY_D)])

        s.configure("Danger.TButton",
                    background="#FFEBEE", foreground=COLOR_ERROR,
                    font=FONT_BASE, padding=(10, 7),
                    borderwidth=0, focusthickness=0)
        s.map("Danger.TButton",
              background=[("active", "#FFCDD2"), ("pressed", "#FFCDD2")],
              foreground=[("active", COLOR_ERROR)])

        s.configure("TNotebook", background=COLOR_BG, borderwidth=0)
        s.configure("TNotebook.Tab", background=COLOR_ACCENT, foreground=COLOR_PRIMARY_D,
                    font=FONT_BOLD, padding=(16, 8))
        s.map("TNotebook.Tab",
              background=[("selected", COLOR_PRIMARY), ("active", COLOR_BORDER)],
              foreground=[("selected", "white"), ("active", COLOR_PRIMARY_D)])

        s.configure("Treeview", background=COLOR_BG, fieldbackground=COLOR_BG,
                    foreground=COLOR_TEXT, rowheight=26, font=FONT_BASE, borderwidth=0)
        s.configure("Treeview.Heading",
                    background=COLOR_ACCENT, foreground=COLOR_PRIMARY_D,
                    font=FONT_BOLD, relief="flat", padding=(6, 8))
        s.map("Treeview.Heading", background=[("active", COLOR_BORDER)])
        s.map("Treeview",
              background=[("selected", COLOR_PRIMARY)],
              foreground=[("selected", "white")])

        s.configure("Blue.Horizontal.TProgressbar",
                    troughcolor=COLOR_ACCENT, background=COLOR_PRIMARY,
                    bordercolor=COLOR_ACCENT, lightcolor=COLOR_PRIMARY,
                    darkcolor=COLOR_PRIMARY, thickness=14)

    def _center_window(self, w, h):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2 - 20)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ---------------- UI ----------------
    def _build_ui(self):
        header = tk.Frame(self.root, bg=COLOR_PRIMARY, height=60)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="Batch Outbound Tool  批量出库工具",
                 bg=COLOR_PRIMARY, fg="white", font=("Segoe UI", 15, "bold")
                 ).pack(side=tk.LEFT, padx=20)

        body = ttk.Frame(self.root, style="TFrame", padding=(16, 12, 16, 12))
        body.pack(fill=tk.BOTH, expand=True)

        # ---- 配置卡片 ----
        cfg = ttk.LabelFrame(body, text=" Configuration 配置 ",
                             style="Card.TLabelframe", padding=(14, 10))
        cfg.pack(fill=tk.X, pady=(0, 10))

        # ---- WiFi 行 ----
        ttk.Label(cfg, text="WiFi 当前网络:", background=COLOR_BG_ALT).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 6))
        self.wifi_var = tk.StringVar(value="(checking...)")
        ttk.Label(cfg, textvariable=self.wifi_var, background=COLOR_BG_ALT,
                  foreground=COLOR_PRIMARY_D, font=FONT_BOLD).grid(
            row=0, column=1, sticky=tk.W)

        ttk.Button(cfg, text="Refresh 刷新", style="Secondary.TButton",
                   command=self._refresh_wifi_label).grid(row=0, column=2, padx=(8, 6))
        ttk.Button(cfg, text="To AOI 切到AOI", style="Secondary.TButton",
                   command=self._switch_to_aoi).grid(row=0, column=3, padx=(0, 6))
        ttk.Button(cfg, text="Restore 切回原网", style="Secondary.TButton",
                   command=self._switch_to_original).grid(row=0, column=4, padx=(0, 6))

        # ---- 仓库 ID 行 ----
        ttk.Label(cfg, text="Warehouse IDs 仓库ID:", background=COLOR_BG_ALT
                  ).grid(row=1, column=0, sticky=tk.W, padx=(0, 6), pady=(10, 0))

        self.wh_var = tk.StringVar(value=",".join(str(x) for x in self.warehouse_ids))
        wh_entry = ttk.Entry(cfg, textvariable=self.wh_var, width=30)
        wh_entry.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))

        ttk.Label(cfg, text="(逗号或空格分隔，如 43,44)",
                  style="Muted.TLabel").grid(row=1, column=2, columnspan=2,
                                             sticky=tk.W, padx=(8, 0), pady=(10, 0))

        ttk.Button(cfg, text="Apply & Save 应用并保存",
                   style="Secondary.TButton",
                   command=self.on_apply_warehouse_ids
                   ).grid(row=1, column=4, sticky=tk.W, pady=(10, 0))

        # ---- 操作行 ----
        self.btn_download = ttk.Button(cfg, text="(1) Query Stock 查询库存",
                                       style="Primary.TButton",
                                       command=self.on_query_stock)
        self.btn_download.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))

        self.btn_export = ttk.Button(cfg, text="(2) Export Stock 导出库存",
                                     style="Primary.TButton",
                                     command=self.on_export_stock,
                                     state=tk.DISABLED)
        self.btn_export.grid(row=2, column=2, columnspan=2, sticky=tk.W, pady=(10, 0))

        self.btn_stats = ttk.Button(cfg, text="Stats 统计",
                                    style="Secondary.TButton",
                                    command=self.show_stats)
        self.btn_stats.grid(row=2, column=4, sticky=tk.W, pady=(10, 0))

        self.btn_clear_log = ttk.Button(cfg, text="Clear Log 清空日志",
                                        style="Secondary.TButton",
                                        command=self.clear_log)
        self.btn_clear_log.grid(row=3, column=4, sticky=tk.W, pady=(6, 0))

        # ---- 状态栏 ----
        status_wrap = tk.Frame(body, bg=COLOR_BG_ALT,
                               highlightbackground=COLOR_BORDER, highlightthickness=1)
        status_wrap.pack(fill=tk.X, pady=(0, 10))
        self.status_var = tk.StringVar(value="Ready 就绪")
        ttk.Label(status_wrap, textvariable=self.status_var, style="Status.TLabel",
                  padding=(10, 6)).pack(anchor=tk.W)

        # ---- 进度条 ----
        prog_wrap = ttk.Frame(body, style="TFrame")
        prog_wrap.pack(fill=tk.X, pady=(0, 10))
        self.progress = ttk.Progressbar(prog_wrap, mode="determinate",
                                        style="Blue.Horizontal.TProgressbar",
                                        maximum=100, value=0)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_text = tk.StringVar(value="")
        ttk.Label(prog_wrap, textvariable=self.progress_text,
                  background=COLOR_BG, foreground=COLOR_PRIMARY_D,
                  font=FONT_BOLD, width=26).pack(side=tk.LEFT, padx=(10, 0))

        # ---- Notebook ----
        nb = ttk.Notebook(body)
        nb.pack(fill=tk.BOTH, expand=True)

        # ===== PN Tab =====
        pn_tab = ttk.Frame(nb, style="TFrame", padding=(8, 8))
        nb.add(pn_tab, text=" By PN 按料号 ")

        pn_top = ttk.Frame(pn_tab, style="TFrame")
        pn_top.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(pn_top, text="Search PN 模糊搜索:").pack(side=tk.LEFT, padx=(0, 6))
        self.pn_filter_var = tk.StringVar()
        pn_entry = ttk.Entry(pn_top, textvariable=self.pn_filter_var, width=40)
        pn_entry.pack(side=tk.LEFT, padx=(0, 8))
        pn_entry.bind("<KeyRelease>", lambda e: self._apply_pn_filter())
        ttk.Button(pn_top, text="Clear 清除", style="Secondary.TButton",
                   command=lambda: (self.pn_filter_var.set(""),
                                    self._apply_pn_filter())).pack(side=tk.LEFT)

        self.btn_out_pn = ttk.Button(pn_top, text="Outbound by PN 按PN出库",
                                     style="Primary.TButton",
                                     command=self.on_outbound_by_pn,
                                     state=tk.DISABLED)
        self.btn_out_pn.pack(side=tk.LEFT, padx=(12, 0))

        ttk.Label(pn_top, text="(select a row, then click Outbound by PN)",
                  style="Muted.TLabel").pack(side=tk.LEFT, padx=12)

        pn_tree_frame = ttk.Frame(pn_tab, style="TFrame")
        pn_tree_frame.pack(fill=tk.BOTH, expand=True)

        self.pn_tree = ttk.Treeview(pn_tree_frame,
                                    columns=("part_num", "count"),
                                    show="headings", selectmode="browse")
        self.pn_tree.heading("part_num", text="Part Number 料号")
        self.pn_tree.heading("count", text="Occurrences 出现次数")
        self.pn_tree.column("part_num", anchor=tk.W, stretch=True, width=600)
        self.pn_tree.column("count", anchor=tk.CENTER, stretch=False, width=140)

        pn_vsb = ttk.Scrollbar(pn_tree_frame, orient=tk.VERTICAL,
                               command=self.pn_tree.yview)
        self.pn_tree.configure(yscrollcommand=pn_vsb.set)
        self.pn_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pn_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # ===== RT Tab =====
        rt_tab = ttk.Frame(nb, style="TFrame", padding=(8, 8))
        nb.add(rt_tab, text=" By RT 按RT ")

        rt_top = ttk.Frame(rt_tab, style="TFrame")
        rt_top.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(rt_top, text="Paste RTs (one per line) 粘贴 RT（一行一个）:",
                  font=FONT_BOLD).pack(anchor=tk.W)

        rt_input_wrap = ttk.Frame(rt_tab, style="TFrame")
        rt_input_wrap.pack(fill=tk.X, pady=(0, 6))

        self.rt_input = scrolledtext.ScrolledText(
            rt_input_wrap, height=6, font=FONT_MONO, wrap=tk.NONE,
            bg=COLOR_BG, fg=COLOR_TEXT, relief="solid", borderwidth=1,
            highlightthickness=0
        )
        self.rt_input.pack(fill=tk.X, expand=False)

        rt_btn_row = ttk.Frame(rt_tab, style="TFrame")
        rt_btn_row.pack(fill=tk.X, pady=(4, 6))

        ttk.Button(rt_btn_row, text="Clear 清空", style="Secondary.TButton",
                   command=lambda: self.rt_input.delete("1.0", tk.END)).pack(side=tk.LEFT)
        ttk.Button(rt_btn_row, text="Validate 校验 RT", style="Secondary.TButton",
                   command=self.on_validate_rt).pack(side=tk.LEFT, padx=6)

        ttk.Button(rt_btn_row, text="Select All 全选", style="Secondary.TButton",
                   command=self.on_select_all_rt).pack(side=tk.LEFT, padx=(12, 6))
        ttk.Button(rt_btn_row, text="Deselect All 取消全选", style="Secondary.TButton",
                   command=self.on_deselect_all_rt).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(rt_btn_row, text="Invert 反选", style="Secondary.TButton",
                   command=self.on_invert_rt).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(rt_btn_row, text="Outbound Selected 出库选中",
                   style="Primary.TButton",
                   command=self.on_outbound_rt_selected).pack(side=tk.LEFT, padx=(12, 6))
        ttk.Button(rt_btn_row, text="Outbound All Matched 出库全部匹配",
                   style="Danger.TButton",
                   command=self.on_outbound_rt_all).pack(side=tk.LEFT, padx=(0, 6))

        rt_tree_label = ttk.Frame(rt_tab, style="TFrame")
        rt_tree_label.pack(fill=tk.X, pady=(4, 2))
        ttk.Label(rt_tree_label, text="Matched materials 匹配到的物料:",
                  font=FONT_BOLD).pack(side=tk.LEFT)
        self.rt_count_var = tk.StringVar(value="")
        ttk.Label(rt_tree_label, textvariable=self.rt_count_var,
                  style="Muted.TLabel").pack(side=tk.LEFT, padx=8)

        rt_tree_frame = ttk.Frame(rt_tab, style="TFrame")
        rt_tree_frame.pack(fill=tk.BOTH, expand=True)

        self.rt_tree = ttk.Treeview(
            rt_tree_frame,
            columns=("rt", "part_num", "shelf_id", "position", "label_code"),
            show="headings", selectmode="extended")
        self.rt_tree.heading("rt", text="RT")
        self.rt_tree.heading("part_num", text="Part Number 料号")
        self.rt_tree.heading("shelf_id", text="Shelf ID")
        self.rt_tree.heading("position", text="Position")
        self.rt_tree.heading("label_code", text="Label Code")
        self.rt_tree.column("rt", width=140, anchor=tk.W)
        self.rt_tree.column("part_num", width=260, anchor=tk.W)
        self.rt_tree.column("shelf_id", width=100, anchor=tk.CENTER)
        self.rt_tree.column("position", width=100, anchor=tk.CENTER)
        self.rt_tree.column("label_code", width=360, anchor=tk.W)

        rt_vsb = ttk.Scrollbar(rt_tree_frame, orient=tk.VERTICAL,
                               command=self.rt_tree.yview)
        self.rt_tree.configure(yscrollcommand=rt_vsb.set)
        self.rt_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rt_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # ===== Log Tab =====
        log_tab = ttk.Frame(nb, style="TFrame", padding=(8, 8))
        nb.add(log_tab, text=" Log 日志 ")
        self.log = scrolledtext.ScrolledText(
            log_tab, height=14, wrap=tk.WORD, font=FONT_MONO,
            bg=COLOR_BG, fg=COLOR_TEXT, relief="flat",
            highlightthickness=0, insertbackground=COLOR_PRIMARY
        )
        self.log.pack(fill=tk.BOTH, expand=True)
        self.log.tag_config(COLOR_TEXT, foreground=COLOR_TEXT)
        self.log.tag_config(COLOR_WARN, foreground=COLOR_WARN)
        self.log.tag_config(COLOR_ERROR, foreground=COLOR_ERROR)
        self.log.tag_config(COLOR_SUCCESS, foreground=COLOR_SUCCESS)
        self.log.configure(state=tk.DISABLED)

    # ---------------- Log ----------------
    def append_log(self, msg, level="INFO"):
        def _do():
            ts = time.strftime("%H:%M:%S")
            color = {"INFO": COLOR_TEXT, "WARN": COLOR_WARN,
                     "ERROR": COLOR_ERROR, "OK": COLOR_SUCCESS}.get(level, COLOR_TEXT)
            self.log.configure(state=tk.NORMAL)
            self.log.insert(tk.END, f"[{ts}] {level}: {msg}\n", (color,))
            self.log.see(tk.END)
            self.log.configure(state=tk.DISABLED)
        self.root.after(0, _do)

    def clear_log(self):
        self.log.configure(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.configure(state=tk.DISABLED)

    def _set_progress(self, cur, total, text=""):
        if total <= 0:
            self.root.after(0, lambda: self.progress.configure(value=0))
            self.root.after(0, lambda: self.progress_text.set(""))
            return
        pct = max(0.0, min(100.0, cur * 100.0 / total))
        self.root.after(0, lambda: self.progress.configure(value=pct))
        label = f"{cur}/{total} ({pct:.0f}%)"
        if text:
            label += f"  {text}"
        self.root.after(0, lambda: self.progress_text.set(label))

    def _reset_progress(self):
        self.root.after(0, lambda: self.progress.configure(value=0))
        self.root.after(0, lambda: self.progress_text.set(""))

    # ---------------- Warehouse IDs ----------------
    def _sync_warehouse_ids_from_ui(self, silent=False) -> bool:
        """
        从编辑框读取仓库 ID，更新 self.warehouse_ids 并保存配置。
        返回 True 表示成功（并已应用），False 表示校验失败。
        """
        try:
            ids = parse_warehouse_ids(self.wh_var.get())
        except ValueError as e:
            if not silent:
                messagebox.showerror("Invalid 无效", str(e))
            return False
        if not ids:
            if not silent:
                messagebox.showwarning("Warning", "Warehouse ID list is empty.")
            return False

        changed = (ids != self.warehouse_ids)
        self.warehouse_ids = ids
        self.config_data["warehouse_name_ids"] = ids
        save_config(self.config_data)

        normalized = ",".join(str(x) for x in ids)
        self.wh_var.set(normalized)
        if changed and not silent:
            self.append_log(f"Warehouse IDs applied & saved: {normalized}", "OK")
        return True

    def on_apply_warehouse_ids(self):
        if self._sync_warehouse_ids_from_ui():
            self.status_var.set(
                f"Warehouse IDs: {','.join(str(x) for x in self.warehouse_ids)} (saved)")

    # ---------------- WiFi ----------------
    def _refresh_wifi_label(self):
        ssid = WifiManager.get_current_ssid() or "(not connected 未连接)"
        self.wifi_var.set(ssid)

    def _switch_to_aoi(self):
        self.append_log(f"Switching WiFi to {WIFI_PROFILE_NAME}... 切换中...")
        if not WifiManager.ensure_profile():
            self.append_log(f"Profile {WIFI_PROFILE_NAME} missing 无法创建", "ERROR")
            messagebox.showerror("WiFi", f"Cannot ensure profile {WIFI_PROFILE_NAME}")
            return

        cur = WifiManager.get_current_ssid()
        if cur and cur != WIFI_PROFILE_NAME:
            self._original_wifi = cur

        def worker():
            ok = WifiManager.connect(WIFI_PROFILE_NAME)
            self.append_log("WiFi OK 已连接 AOI" if ok else "WiFi 连接失败",
                            "OK" if ok else "ERROR")
            self.root.after(0, self._refresh_wifi_label)
        threading.Thread(target=worker, daemon=True).start()

    def _switch_to_original(self):
        if not self._original_wifi:
            messagebox.showinfo("WiFi", "No original WiFi recorded 未记录原网络")
            return
        target = self._original_wifi
        self.append_log(f"Switching WiFi back to {target}... 切回原网...")

        def worker():
            ok = WifiManager.connect(target)
            self.append_log(f"WiFi switched back to {target}" if ok
                            else f"Failed to switch back to {target}",
                            "OK" if ok else "ERROR")
            self.root.after(0, self._refresh_wifi_label)
        threading.Thread(target=worker, daemon=True).start()

    # ---------------- Query Stock ----------------
    def on_query_stock(self):
        if self._busy:
            return

        # 先把编辑框内容同步为当前配置（会自动保存）
        if not self._sync_warehouse_ids_from_ui(silent=False):
            return

        self._busy = True
        self.btn_download.configure(state=tk.DISABLED)
        self.btn_export.configure(state=tk.DISABLED)
        self.btn_out_pn.configure(state=tk.DISABLED)

        for item in self.pn_tree.get_children():
            self.pn_tree.delete(item)
        for item in self.rt_tree.get_children():
            self.rt_tree.delete(item)
        self.rt_count_var.set("")
        self.all_materials = []
        self.part_stats = []
        self.rt_map = defaultdict(list)

        self.status_var.set("Querying stock... 正在查询库存...")
        self._reset_progress()
        self.append_log("=" * 60)
        self.append_log(f"Start querying stock, warehouse_ids={self.warehouse_ids}")

        threading.Thread(target=self._query_thread,
                         args=(list(self.warehouse_ids),),
                         daemon=True).start()

    def _query_thread(self, warehouse_ids):
        def progress_cb(cur, total):
            self._set_progress(cur, total, "shelf 货架")

        materials, ok = fetch_all_materials(
            warehouse_ids,
            log_cb=lambda m: self.append_log(m),
            progress_cb=progress_cb)
        self.root.after(0, self._on_query_finished, materials, ok)

    def _on_query_finished(self, materials, ok):
        self._reset_progress()
        self._busy = False
        self.btn_download.configure(state=tk.NORMAL)

        if not ok:
            self.status_var.set("Query failed 查询失败")
            self.append_log("Query failed, please check network / server", "ERROR")
            messagebox.showerror("Error", "Failed to query stock. Please check the log.")
            return

        self.all_materials = materials
        self.append_log(f"Query complete, total {len(materials)} records", "OK")

        counter = Counter()
        for item in materials:
            part = item.get("part_num", "")
            if part:
                counter[part] += 1
        self.part_stats = counter.most_common()

        for item in materials:
            label = item.get("label_code", "") or ""
            rt = extract_rt_from_label(label)
            if not rt:
                rt = item.get("rt", "") or ""
            if rt:
                self.rt_map[rt].append(item)

        self.append_log(f"{len(self.part_stats)} distinct PNs, "
                        f"{len(self.rt_map)} distinct RTs", "OK")

        self._apply_pn_filter()

        if self.part_stats:
            self.btn_out_pn.configure(state=tk.NORMAL)
            self.btn_export.configure(state=tk.NORMAL)
            self.status_var.set(
                f"Loaded {len(materials)} records, "
                f"{len(self.part_stats)} PNs, {len(self.rt_map)} RTs "
                f"(warehouse_ids={self.warehouse_ids})")
        else:
            self.status_var.set("No data 无数据")
            messagebox.showinfo("Info", "No material data retrieved.")

    # ---------------- Export Stock ----------------
    def on_export_stock(self):
        if self._busy:
            return
        if not self.all_materials:
            messagebox.showwarning(
                "Warning",
                "No stock data in memory. Please click Query Stock first.")
            return
        if not HAS_OPENPYXL:
            messagebox.showerror(
                "Missing openpyxl",
                "openpyxl is not installed. Run: pip install openpyxl")
            return

        ts = time.strftime("%Y%m%d_%H%M%S")
        default_name = f"mate_export_{ts}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Export Stock 导出库存",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx")])
        if not path:
            self.append_log("Export cancelled by user")
            return

        self.append_log(f"Exporting stock to {path} ...")
        self.status_var.set("Exporting stock... 正在导出库存...")
        self._busy = True
        self.btn_download.configure(state=tk.DISABLED)
        self.btn_export.configure(state=tk.DISABLED)
        self.btn_out_pn.configure(state=tk.DISABLED)

        def worker():
            try:
                ok, msg = export_to_excel(
                    self.all_materials, path,
                    log_cb=lambda m: self.append_log(m))
                self.root.after(0, self._on_export_finished, ok, msg, path)
            except Exception as e:
                self.root.after(0, self._on_export_finished, False, str(e), path)

        threading.Thread(target=worker, daemon=True).start()

    def _on_export_finished(self, ok, msg, path):
        self._busy = False
        self.btn_download.configure(state=tk.NORMAL)
        if self.part_stats:
            self.btn_export.configure(state=tk.NORMAL)
            self.btn_out_pn.configure(state=tk.NORMAL)

        if ok:
            self.append_log(f"Export complete: {path}", "OK")
            self.status_var.set(f"Export done 导出完成: {path}")
            messagebox.showinfo("Export 导出", f"Saved 已保存:\n{path}")
        else:
            self.append_log(f"Export failed: {msg}", "ERROR")
            self.status_var.set("Export failed 导出失败")
            messagebox.showerror("Export 导出失败", msg)

    # ---------------- PN filter / table ----------------
    def _apply_pn_filter(self):
        keyword = self.pn_filter_var.get().strip().lower()
        for item in self.pn_tree.get_children():
            self.pn_tree.delete(item)

        max_count = self.part_stats[0][1] if self.part_stats else 1
        for part, cnt in self.part_stats:
            if keyword and keyword not in part.lower():
                continue
            tags = ("top",) if cnt == max_count else ()
            self.pn_tree.insert("", tk.END, values=(part, cnt), tags=tags)
        self.pn_tree.tag_configure("top", background="#FFF3B0")

    # ---------------- Outbound by PN ----------------
    def on_outbound_by_pn(self):
        if self._busy:
            return

        # 出库前先把编辑框内容同步（含保存），确保用的就是配置 ID
        if not self._sync_warehouse_ids_from_ui(silent=False):
            return

        sel = self.pn_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a PN first.")
            return
        values = self.pn_tree.item(sel[0], "values")
        part_num = values[0]
        count = values[1]

        reply = messagebox.askyesno(
            "Confirm 确认",
            f"Are you sure you want to outbound ALL materials of this PN?\n\n"
            f"Part Number: {part_num}\n"
            f"Occurrences in queried data: {count}\n"
            f"Target warehouse IDs: {self.warehouse_ids}\n\n"
            f"确认按此料号出库？（将重新向服务器查询该 PN 的物料并全部出库，"
            f"仅限上述仓库 ID）",
            icon="warning", default="no")
        if not reply:
            self.append_log("User cancelled PN outbound")
            return

        self._busy = True
        self.btn_download.configure(state=tk.DISABLED)
        self.btn_export.configure(state=tk.DISABLED)
        self.btn_out_pn.configure(state=tk.DISABLED)
        self._reset_progress()
        self.append_log("=" * 60)
        self.append_log(f"Start PN outbound: part_num={part_num}, "
                        f"warehouse_ids={self.warehouse_ids}")

        threading.Thread(target=self._pn_out_thread,
                         args=(part_num, list(self.warehouse_ids)),
                         daemon=True).start()

    def _pn_out_thread(self, part_num, warehouse_ids):
        self.append_log(f"Searching materials for part_num={part_num} "
                        f"(warehouse_ids={warehouse_ids}) ...")
        materials = search_materials_by_part_num(
            part_num, warehouse_ids, self.append_log)
        if not materials:
            self.append_log("No materials found for outbound", "WARN")
            self.root.after(0, self._on_out_finished, 0, 0, [], part_num)
            return

        total = len(materials)
        self.append_log(f"Found {total} materials, starting outbound...", "OK")
        success, fail, details = self._do_batch_out(materials, total)
        self.root.after(0, self._on_out_finished, success, fail, details, part_num)

    # ---------------- RT input helpers ----------------
    def _parse_rt_input(self):
        raw = self.rt_input.get("1.0", tk.END)
        tokens = re.split(r"[\s,;]+", raw)
        return [t.strip() for t in tokens if t.strip()]

    def on_validate_rt(self):
        if not self.rt_map:
            messagebox.showwarning("Warning", "Please query stock first (Step 1).")
            return
        rts = self._parse_rt_input()
        if not rts:
            messagebox.showwarning("Warning", "Please paste at least one RT.")
            return

        found, not_found = [], []
        for rt in rts:
            if rt in self.rt_map:
                found.append(rt)
            else:
                not_found.append(rt)

        for item in self.rt_tree.get_children():
            self.rt_tree.delete(item)
        for rt in found:
            for mat in self.rt_map[rt]:
                self.rt_tree.insert("", tk.END, values=(
                    rt,
                    mat.get("part_num", ""),
                    mat.get("shelf_id", ""),
                    mat.get("position", ""),
                    mat.get("label_code", ""),
                ))

        self.rt_count_var.set(
            f"Input {len(rts)} | Matched {len(found)} RTs / "
            f"{len(self.rt_tree.get_children())} rows | Not found {len(not_found)}")

        msg = f"Input RTs: {len(rts)}\nMatched: {len(found)}\nNot found: {len(not_found)}"
        if not_found:
            preview = "\n".join(not_found[:10])
            msg += f"\n\nNot found (first 10):\n{preview}"
        self.append_log(f"RT validate: {msg}", "OK" if not not_found else "WARN")
        messagebox.showinfo("Validate RT", msg)

    # ---------------- RT selection helpers ----------------
    def on_select_all_rt(self):
        items = self.rt_tree.get_children()
        if not items:
            return
        self.rt_tree.selection_set(items)

    def on_deselect_all_rt(self):
        self.rt_tree.selection_remove(self.rt_tree.selection())

    def on_invert_rt(self):
        items = self.rt_tree.get_children()
        current = set(self.rt_tree.selection())
        new_sel = [i for i in items if i not in current]
        self.rt_tree.selection_set(new_sel)

    # ---------------- RT outbound ----------------
    def on_outbound_rt_selected(self):
        self._do_rt_outbound(mode="selected")

    def on_outbound_rt_all(self):
        self._do_rt_outbound(mode="all")

    def _do_rt_outbound(self, mode: str):
        if self._busy:
            return

        # 校验：编辑框里的 ID 必须和当前内存数据一致
        try:
            current_ids = parse_warehouse_ids(self.wh_var.get())
        except ValueError as e:
            messagebox.showerror("Invalid 无效", str(e))
            return
        if not current_ids:
            messagebox.showwarning("Warning", "Warehouse ID list is empty.")
            return
        if current_ids != self.warehouse_ids:
            messagebox.showwarning(
                "Warehouse IDs changed 仓库ID已修改",
                f"Memory 数据对应的仓库ID: {self.warehouse_ids}\n"
                f"当前输入框仓库ID: {current_ids}\n\n"
                f"RT 出库使用的是已查询到的内存数据。\n"
                f"请先点击 'Query Stock 查询库存' 重新拉取后再出库。")
            return

        if not self.rt_map:
            messagebox.showwarning("Warning", "Please query stock first (Step 1).")
            return

        items = self.rt_tree.get_children()
        if not items:
            messagebox.showwarning(
                "Warning",
                "No matched rows in the table.\n"
                "Please paste RTs and click Validate first.")
            return

        if mode == "selected":
            sel = self.rt_tree.selection()
            if not sel:
                messagebox.showwarning(
                    "Warning",
                    "No rows selected. Please select rows, "
                    "or use 'Outbound All Matched'.")
                return
            target_items = list(sel)
        else:
            target_items = list(items)

        materials = []
        rts_set = set()
        for iid in target_items:
            vals = self.rt_tree.item(iid, "values")
            rt, part_num, shelf_id, position, label_code = vals
            rts_set.add(rt)
            materials.append({
                "rt": rt,
                "part_num": part_num,
                "shelf_id": shelf_id,
                "position": position,
                "label_code": label_code,
            })

        mode_label = "Selected 选中" if mode == "selected" else "All Matched 全部匹配"
        preview_rts = "\n".join(list(rts_set)[:10])
        reply = messagebox.askyesno(
            "Confirm 确认",
            f"Are you sure you want to outbound by RT?\n\n"
            f"Mode: {mode_label}\n"
            f"Rows to outbound: {len(materials)}\n"
            f"Distinct RTs: {len(rts_set)}\n"
            f"Warehouse IDs (from memory): {self.warehouse_ids}\n\n"
            f"First 10 RTs:\n{preview_rts}\n\n"
            f"确认出库？",
            icon="warning", default="no")
        if not reply:
            self.append_log("User cancelled RT outbound")
            return

        self._busy = True
        self.btn_download.configure(state=tk.DISABLED)
        self.btn_export.configure(state=tk.DISABLED)
        self.btn_out_pn.configure(state=tk.DISABLED)
        self._reset_progress()
        self.append_log("=" * 60)
        self.append_log(f"Start RT outbound [{mode_label}]: "
                        f"rows={len(materials)}, RTs={len(rts_set)}, "
                        f"warehouse_ids={self.warehouse_ids}")

        threading.Thread(target=self._rt_out_thread,
                         args=(materials, mode_label),
                         daemon=True).start()

    def _rt_out_thread(self, materials, mode_label):
        total = len(materials)
        success, fail, details = self._do_batch_out(materials, total)
        self.root.after(0, self._on_out_finished, success, fail, details,
                        f"RT-{mode_label}")

    # ---------------- Common outbound loop ----------------
    def _do_batch_out(self, materials, total):
        success = 0
        fail = 0
        details = []
        for idx, mat in enumerate(materials, start=1):
            shelf_id = mat.get("shelf_id")
            position = mat.get("position")
            if shelf_id is None or position is None:
                fail += 1
                details.append(f"[{idx}] missing shelf_id or position")
                self._set_progress(idx, total, "outbound 出库")
                continue

            ok, msg = out_material(shelf_id, position)
            if ok:
                success += 1
                self.append_log(f"[{idx}/{total}] OK  shelf_id={shelf_id} "
                                f"position={position}", "OK")
            else:
                fail += 1
                details.append(f"[{idx}] shelf_id={shelf_id} "
                               f"position={position} FAILED: {msg}")
                self.append_log(f"[{idx}/{total}] FAIL shelf_id={shelf_id} "
                                f"position={position} -> {msg}", "ERROR")
            self._set_progress(idx, total, "outbound 出库")
            time.sleep(REQUEST_INTERVAL)
        return success, fail, details

    def _on_out_finished(self, success, fail, details, tag):
        self._reset_progress()
        self._busy = False
        self.btn_download.configure(state=tk.NORMAL)
        if self.part_stats:
            self.btn_out_pn.configure(state=tk.NORMAL)
            self.btn_export.configure(state=tk.NORMAL)

        self.append_log(f"Outbound finished [{tag}]: success {success}, failed {fail}",
                        "OK" if fail == 0 else "WARN")
        if details:
            self.append_log("Failure details:")
            for d in details:
                self.append_log("  " + d, "ERROR")

        self.status_var.set(f"Outbound done [{tag}]: success {success} / failed {fail}")
        messagebox.showinfo(
            "Outbound Complete 出库完成",
            f"Outbound finished [{tag}]\nSuccess: {success}\nFailed: {fail}"
        )

    # ---------------- Stats ----------------
    def show_stats(self):
        total = len(self.all_materials)
        pns = len(self.part_stats)
        rts = len(self.rt_map)
        total_qty = 0.0
        for m in self.all_materials:
            try:
                total_qty += float(m.get("quantity") or 0)
            except (TypeError, ValueError):
                pass
        messagebox.showinfo(
            "Stats 统计",
            f"Total records 总条数 : {total}\n"
            f"Distinct PNs 料号数  : {pns}\n"
            f"Distinct RTs RT数    : {rts}\n"
            f"Total qty 总数量     : {total_qty:.0f}\n"
            f"Warehouse IDs 仓库ID : {self.warehouse_ids}"
        )


# ==================== Entry ====================
def main():
    root = tk.Tk()
    MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
