# -*- coding: utf-8 -*-
"""
MES Auto Putaway Tool - GUI Version (v6)

- RT 合并上架 + FIFO 消耗 label
- label_code 唯一码防重复; putaway_rt 记录 RT 累计已上架数量
- Step1 前切 AOI WiFi, Step1 后切回原 WiFi (修复首次切换失败)
- 只在用户点击"Export 导出"时导出 Excel; 平时不落地任何文件
- 界面白蓝配色, 居中显示; 英文在前, 中文在后
- 进度条改为 determinate 模式, 显示具体进度
"""

import hashlib
import requests
import time
import json
import sqlite3
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from typing import Optional, Dict, List, Tuple, Callable
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager
from openpyxl import Workbook

# ======================== Configuration ========================
MES_LOGIN_URL = "http://10.97.245.205:86/login"
MES_ACCOUNT = "NAF684"
MES_PASSWORD = "666888"

WAITING_REPORT_URL = "http://10.97.245.205:92/prod_baseapi/wms/instorereport/getPendingRep/page"
COMPANY_ID = "6801"

SMART_SHELF_URL = "http://192.168.9.100:8081/mate/searchByInfo"
SMART_USER_NAME = "NAF684"
SMART_USER_GROUP_POWER = "1"
SMART_USER_LANGUAGE = "zh"

GET_RT_INFO_URL = "http://10.97.245.205:92/prod_baseapi/wms/sendproductionbill/getRTInfo"
PUTAWAY_URL = "http://10.97.245.205:92/prod_baseapi/CommonAPI/CallSp"

MAX_RETRY = 3
RETRY_DELAY = 2

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

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "putaway_state.db"

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

FONT_BASE   = ("Segoe UI", 10)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_TITLE  = ("Segoe UI", 15, "bold")
FONT_SUB    = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)


# ======================== WiFi Helper ========================
class WifiManager:
    """
    修复要点：
    1. 切换前先 disconnect，避免旧网卡自动重连
    2. 轮询时同时检查 State == connected 且 SSID 匹配
    3. 每 500ms 轮询一次，超时默认 30s
    4. 15s 内没成功自动再发一次 connect
    """

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
        """返回 (state, ssid)，state 如 connected / disconnected / authenticating"""
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
            xml_path = APP_DIR / "_wifi_profile.xml"
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
        """
        连接并阻塞等待到真正 connected。
        修复首次失败：先 disconnect，15s 内没成功再发一次 connect。
        """
        # 已是目标 SSID 且 state = connected 就直接返回
        state, ssid = WifiManager._get_state_and_ssid()
        if ssid == profile_name and state == "connected":
            return True

        # 先断开，避免系统自动连回旧的强信号网络
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
                # 稳定一下再确认（避免刚连上就掉）
                time.sleep(1.5)
                state2, ssid2 = WifiManager._get_state_and_ssid()
                if ssid2 == profile_name and state2 == "connected":
                    return True
            # 15 秒还没成功，再发一次 connect
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


# ======================== Local Database ========================
class LocalDB:
    _lock = threading.Lock()

    def __init__(self, path: Path = DB_PATH):
        self.path = str(path)
        self._init_schema()

    @contextmanager
    def conn(self):
        c = sqlite3.connect(self.path, timeout=30, check_same_thread=False)
        c.row_factory = sqlite3.Row
        try:
            yield c
            c.commit()
        finally:
            c.close()

    def _init_schema(self):
        with self.conn() as c:
            c.executescript("""
                CREATE TABLE IF NOT EXISTS shelf_snapshot (
                    label_code    TEXT PRIMARY KEY,
                    snapshot_time TEXT NOT NULL,
                    rt            TEXT,
                    part_num      TEXT,
                    qty           REAL,
                    in_time       TEXT,
                    position_info TEXT,
                    shelf_id      TEXT
                );
                CREATE TABLE IF NOT EXISTS putaway_rt (
                    rt            TEXT PRIMARY KEY,
                    putaway_qty   REAL NOT NULL DEFAULT 0,
                    last_time     TEXT
                );
                CREATE TABLE IF NOT EXISTS putaway_label (
                    label_code   TEXT PRIMARY KEY,
                    rt           TEXT,
                    part_num     TEXT,
                    qty          REAL,
                    location     TEXT,
                    putaway_time TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS putaway_log (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    rt         TEXT,
                    part_num   TEXT,
                    qty        REAL,
                    location   TEXT,
                    result     TEXT NOT NULL,
                    message    TEXT,
                    log_time   TEXT NOT NULL
                );
            """)
            cols = [r["name"] for r in c.execute("PRAGMA table_info(shelf_snapshot)").fetchall()]
            for name, typ in [("in_time", "TEXT"), ("rt", "TEXT"), ("part_num", "TEXT"),
                              ("qty", "REAL"), ("position_info", "TEXT"), ("shelf_id", "TEXT")]:
                if name not in cols:
                    c.execute(f"ALTER TABLE shelf_snapshot ADD COLUMN {name} {typ}")
            c.executescript("""
                CREATE INDEX IF NOT EXISTS idx_snap_rt   ON shelf_snapshot(rt);
                CREATE INDEX IF NOT EXISTS idx_snap_time ON shelf_snapshot(snapshot_time);
                CREATE INDEX IF NOT EXISTS idx_snap_in   ON shelf_snapshot(in_time);
                CREATE INDEX IF NOT EXISTS idx_pl_rt     ON putaway_label(rt);
                CREATE INDEX IF NOT EXISTS idx_log_rt    ON putaway_log(rt);
                CREATE INDEX IF NOT EXISTS idx_log_time  ON putaway_log(log_time);
            """)

    def save_shelf_snapshot(self, items: List[dict], snapshot_time: str):
        with self._lock, self.conn() as c:
            rows = []
            for it in items:
                label = it.get("label_code") or ""
                if not label: continue
                rt = SmartShelfDownloader.extract_rt_from_label(label)
                try: qty = float(it.get("quantity") or 0)
                except (TypeError, ValueError): qty = 0.0
                rows.append((label, snapshot_time, rt, it.get("part_num", ""), qty,
                             it.get("in_time", ""), it.get("position_info", ""),
                             it.get("shelf_id", "")))
            c.executemany("""INSERT OR REPLACE INTO shelf_snapshot
                (label_code, snapshot_time, rt, part_num, qty, in_time, position_info, shelf_id)
                VALUES (?,?,?,?,?,?,?,?)""", rows)

    def load_latest_snapshot(self) -> List[dict]:
        with self.conn() as c:
            rows = c.execute("""SELECT label_code, rt, part_num, qty, in_time,
                                       position_info, shelf_id FROM shelf_snapshot""").fetchall()
            return [dict(r) for r in rows]

    def get_putaway_qty_map(self, rts: List[str]) -> Dict[str, float]:
        if not rts: return {}
        result = {}
        with self.conn() as c:
            BATCH = 900
            for i in range(0, len(rts), BATCH):
                chunk = rts[i:i + BATCH]
                ph = ",".join("?" * len(chunk))
                for row in c.execute(
                    f"SELECT rt, putaway_qty FROM putaway_rt WHERE rt IN ({ph})", chunk
                ):
                    result[row["rt"]] = row["putaway_qty"]
        return result

    def add_putaway_rt(self, rt: str, qty: float):
        now = datetime.now().isoformat(timespec="seconds")
        with self._lock, self.conn() as c:
            c.execute("""INSERT INTO putaway_rt(rt, putaway_qty, last_time)
                         VALUES(?, ?, ?)
                         ON CONFLICT(rt) DO UPDATE SET
                           putaway_qty = putaway_qty + excluded.putaway_qty,
                           last_time   = excluded.last_time""", (rt, qty, now))

    def mark_label_done(self, label_code, rt, part_num, qty, location):
        now = datetime.now().isoformat(timespec="seconds")
        with self._lock, self.conn() as c:
            c.execute("""INSERT OR IGNORE INTO putaway_label
                (label_code, rt, part_num, qty, location, putaway_time)
                VALUES (?,?,?,?,?,?)""",
                (label_code, rt, part_num, qty, location, now))

    def log_putaway(self, rt, part_num, qty, location, result, message):
        now = datetime.now().isoformat(timespec="seconds")
        with self._lock, self.conn() as c:
            c.execute("""INSERT INTO putaway_log
                (rt, part_num, qty, location, result, message, log_time)
                VALUES (?,?,?,?,?,?,?)""",
                (rt, part_num, qty, location, result, message, now))

    def stats(self) -> dict:
        with self.conn() as c:
            rts  = c.execute("SELECT COUNT(*) FROM putaway_rt").fetchone()[0]
            qty  = c.execute("SELECT COALESCE(SUM(putaway_qty),0) FROM putaway_rt").fetchone()[0]
            lbls = c.execute("SELECT COUNT(*) FROM putaway_label").fetchone()[0]
            snap = c.execute("SELECT COUNT(*) FROM shelf_snapshot").fetchone()[0]
            logs = c.execute("SELECT COUNT(*) FROM putaway_log").fetchone()[0]
        return {"rt_count": rts, "rt_qty": qty, "label_count": lbls,
                "snapshot_labels": snap, "log_rows": logs}

    def query_success_rt(self, rt_filter: str = "", part_filter: str = "",
                         limit: int = 2000) -> List[dict]:
        sql = """SELECT rt, putaway_qty, last_time FROM putaway_rt WHERE 1=1"""
        params = []
        if rt_filter:
            sql += " AND rt LIKE ?"; params.append(f"%{rt_filter}%")
        sql += " ORDER BY last_time DESC LIMIT ?"
        params.append(limit)
        with self.conn() as c:
            return [dict(r) for r in c.execute(sql, params).fetchall()]

    def query_success_labels(self, rt_filter: str = "", part_filter: str = "",
                             limit: int = 2000) -> List[dict]:
        sql = """SELECT label_code, rt, part_num, qty, location, putaway_time
                 FROM putaway_label WHERE 1=1"""
        params = []
        if rt_filter:
            sql += " AND rt LIKE ?"; params.append(f"%{rt_filter}%")
        if part_filter:
            sql += " AND part_num LIKE ?"; params.append(f"%{part_filter}%")
        sql += " ORDER BY putaway_time DESC LIMIT ?"
        params.append(limit)
        with self.conn() as c:
            return [dict(r) for r in c.execute(sql, params).fetchall()]

    def query_log(self, rt_filter: str = "", result_filter: str = "",
                  limit: int = 2000) -> List[dict]:
        sql = "SELECT rt, part_num, qty, location, result, message, log_time FROM putaway_log WHERE 1=1"
        params = []
        if rt_filter:
            sql += " AND rt LIKE ?"; params.append(f"%{rt_filter}%")
        if result_filter:
            sql += " AND result = ?"; params.append(result_filter)
        sql += " ORDER BY log_time DESC LIMIT ?"
        params.append(limit)
        with self.conn() as c:
            return [dict(r) for r in c.execute(sql, params).fetchall()]

    def reset_rt(self, rt: str):
        with self._lock, self.conn() as c:
            c.execute("DELETE FROM putaway_rt WHERE rt=?", (rt,))
            c.execute("DELETE FROM putaway_label WHERE rt=?", (rt,))

    def clear_log(self):
        with self._lock, self.conn() as c:
            c.execute("DELETE FROM putaway_log")


# ======================== Logger ================================
class GuiLogger:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def info(self, msg):      self._log("INFO", msg)
    def warning(self, msg):   self._log("WARN", msg)
    def error(self, msg):     self._log("ERROR", msg)
    def exception(self, msg): self._log("ERROR", msg)

    def _log(self, level, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        color_map = {"INFO": COLOR_TEXT, "WARN": COLOR_WARN, "ERROR": COLOR_ERROR}
        color = color_map.get(level, COLOR_TEXT)
        self.text_widget.insert(tk.END, f"[{ts}] {level}: {msg}\n", (color,))
        self.text_widget.see(tk.END)
        self.text_widget.update_idletasks()


# ======================== MES Auth ================================
class MESAuth:
    @staticmethod
    def encrypt_password(password: str) -> str:
        md5 = hashlib.md5(); md5.update(password.encode('utf-8'))
        return md5.hexdigest()

    @staticmethod
    def login(account: str, password: str, logger) -> Optional[str]:
        logger.info("Logging in to MES... 登录 MES...")
        encrypted_pwd = MESAuth.encrypt_password(password)
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Connection": "keep-alive",
            "Content-Type": "application/json;charset=UTF-8",
            "Host": "10.97.245.205:86",
            "Origin": "http://10.97.245.205:92",
            "Referer": "http://10.97.245.205:92/",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        data = {"account": account, "password": encrypted_pwd}
        try:
            resp = requests.post(MES_LOGIN_URL, headers=headers, json=data, timeout=10)
            resp.raise_for_status()
            result = resp.json()
            if result.get("code") == 200 and "data" in result and "token" in result["data"]:
                token = result["data"]["token"]
                logger.info("Login successful 登录成功")
                return token
            logger.error(f"Login failed 登录失败: {result.get('info', 'Unknown')}")
            return None
        except Exception as e:
            logger.exception(f"Login exception 登录异常: {e}")
            return None


# ======================== Smart Shelf ================================
class SmartShelfDownloader:
    @staticmethod
    def build_payload(start: int, rows: int) -> dict:
        return {
            "part_num": "", "warehouse_category_id": 0, "warehouse_name_id": 0,
            "save_id": "", "lot_code": "", "mfg_date_start": "", "supplier_name": "",
            "mfg_date_end": "", "start": start, "rows": rows,
            "user_name": SMART_USER_NAME,
            "user_group_power": SMART_USER_GROUP_POWER,
            "user_language": SMART_USER_LANGUAGE,
            "shelf_id": "", "position_info": "", "start_date": "", "end_date": "",
        }

    @staticmethod
    def fetch_page(start: int, rows: int = 10000, logger=None) -> Optional[dict]:
        payload = SmartShelfDownloader.build_payload(start, rows)
        for attempt in range(1, MAX_RETRY + 1):
            try:
                resp = requests.post(SMART_SHELF_URL, json=payload, timeout=30)
                resp.raise_for_status()
                data = resp.json()
                if data.get("result") == 1:
                    return data
                if logger:
                    logger.warning(f"Shelf attempt {attempt} result!=1: {data}")
            except Exception as e:
                if logger:
                    logger.warning(f"Shelf attempt {attempt} exception: {e}")
            if attempt < MAX_RETRY:
                time.sleep(RETRY_DELAY)
        return None

    @staticmethod
    def extract_rt_from_label(label: str) -> str:
        if not label: return ""
        try:
            parts = label.split('#')
            return parts[1].strip() if len(parts) >= 2 else ""
        except Exception:
            return ""

    @staticmethod
    def download_all_raw(logger, progress_cb: Optional[Callable[[int, int], None]] = None) -> List[dict]:
        all_items: List[dict] = []
        start = 0
        consecutive_empty = 0
        total = None
        logger.info("Downloading shelf data... 下载货架数据...")
        while True:
            data = SmartShelfDownloader.fetch_page(start, logger=logger)
            if data is None:
                logger.error(f"start={start} retry failed, stopping")
                break
            mate_list = data.get("mate_list", [])
            if total is None:
                total = data.get("rows", 0)
                logger.info(f"Shelf total records 货架总条数: {total}")
            if not mate_list:
                consecutive_empty += 1
                logger.info(f"  start={start} empty page ({consecutive_empty}/3)")
                if consecutive_empty >= 3:
                    logger.info("3 empty pages, complete")
                    break
                start += 10000
                time.sleep(0.5)
                continue
            consecutive_empty = 0
            all_items.extend(mate_list)
            logger.info(f"Fetched {len(mate_list)} items (acc {len(all_items)})")
            if progress_cb and total:
                progress_cb(len(all_items), total)
            if total and len(mate_list) < 10000:
                break
            start += 10000
            time.sleep(0.1)
        logger.info(f"Shelf download complete 下载完成: {len(all_items)} items")
        if progress_cb and total:
            progress_cb(len(all_items), total)
        return all_items


# ======================== Waiting Report ================================
class WaitingReportDownloader:
    @staticmethod
    def fetch_page(token: str, page: int, rows: int = 100000, logger=None) -> Optional[dict]:
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Connection": "keep-alive",
            "Host": "10.97.245.205:92",
            "Referer": "http://10.97.245.205:92/wms/report/PendingRep/index",
            "token": token,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        params = {
            "rows": rows, "page": page, "sidx": "vtech_lot", "sord": "asc",
            "start_date": "", "end_date": "", "company_id": COMPANY_ID,
            "startTime": "00:00:00", "endTime": "00:00:00",
            "partNo": "", "batch": ""
        }
        for attempt in range(1, MAX_RETRY + 1):
            try:
                resp = requests.get(WAITING_REPORT_URL, headers=headers,
                                    params=params, timeout=30)
                resp.raise_for_status()
                data = resp.json()
                if data.get("code") == 200:
                    return data.get("data", {})
                if logger:
                    logger.warning(f"Attempt {attempt} code!=200: {data}")
            except Exception as e:
                if logger:
                    logger.warning(f"Attempt {attempt} exception: {e}")
            if attempt < MAX_RETRY:
                time.sleep(RETRY_DELAY)
        return None

    @staticmethod
    def download_all(token: str, logger) -> List[dict]:
        all_rows = []
        page = 1
        consecutive_empty = 0
        while True:
            logger.info(f"Fetching waiting report page {page}...")
            data = WaitingReportDownloader.fetch_page(token, page, logger=logger)
            if data is None:
                consecutive_empty += 1
                if consecutive_empty >= 3:
                    logger.warning("3 consecutive failures, stopping")
                    break
                page += 1
                continue
            rows = data.get("rows", [])
            records = data.get("records", 0)
            logger.info(f"  Returned {len(rows)} rows, total {records}")
            if not rows:
                consecutive_empty += 1
                if consecutive_empty >= 3:
                    break
            else:
                consecutive_empty = 0
                all_rows.extend(rows)
            if len(rows) < 100000:
                break
            if records > 0 and len(all_rows) >= records:
                break
            page += 1
            time.sleep(0.3)
        return all_rows

    @staticmethod
    def extract_wait_map(rows: List[dict]) -> Dict[str, dict]:
        result: Dict[str, dict] = {}
        for r in rows:
            rt = r.get("vtech_lot", "")
            if not rt: continue
            try: q = float(r.get("unit_qty") or 0)
            except (TypeError, ValueError): q = 0.0
            if q <= 0: continue
            if rt not in result:
                result[rt] = {"qty": 0.0, "part_num": r.get("part_no", "")}
            result[rt]["qty"] += q
        return result


# ======================== Putaway ================================
class PutawayProcessor:
    @staticmethod
    def get_rt_info(token: str, vtech_lot: str, logger) -> Optional[dict]:
        url = f"{GET_RT_INFO_URL}/{vtech_lot}/ON"
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Connection": "keep-alive",
            "Host": "10.97.245.205:92",
            "Referer": "http://10.97.245.205:92/wms/sendproduction/onshelf/index",
            "token": token,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        for attempt in range(1, MAX_RETRY + 1):
            try:
                resp = requests.get(url, headers=headers, timeout=10)
                resp.raise_for_status()
                data = resp.json()
                if isinstance(data, list) and data:
                    return data[0]
                return None
            except Exception as e:
                if logger:
                    logger.warning(f"getRTInfo attempt {attempt} exception: {e}")
            if attempt < MAX_RETRY:
                time.sleep(RETRY_DELAY)
        return None

    @staticmethod
    def do_putaway(token: str, batch: str, bincode: str, qty: float,
                   account: str, logger) -> Tuple[bool, str]:
        json_param = {
            "p_bincode": bincode, "p_batch": batch,
            "p_totalqty": qty, "p_inqty": 0, "p_qty": qty,
            "account": account
        }
        json_str = json.dumps(json_param, ensure_ascii=False)
        url = f"{PUTAWAY_URL}?spname=do_wms_batch_instore_sp&json={json_str}"
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Connection": "keep-alive",
            "Host": "10.97.245.205:92",
            "Origin": "http://10.97.245.205:92",
            "Referer": "http://10.97.245.205:92/wms/sendproduction/onshelf/index",
            "token": token,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        for attempt in range(1, MAX_RETRY + 1):
            try:
                resp = requests.put(url, headers=headers, timeout=30)
                resp.raise_for_status()
                result = resp.json()
                if "Table" in result and result["Table"]:
                    row = result["Table"][0]
                    if row.get("o_res") == "OK":
                        return True, "Success"
                    return False, f"Error: {row}"
                return False, f"Unexpected response: {result}"
            except Exception as e:
                if logger:
                    logger.warning(f"Putaway attempt {attempt} exception: {e}")
            if attempt < MAX_RETRY:
                time.sleep(RETRY_DELAY)
        return False, "Network error after retries"

    @staticmethod
    def process_putaway(shelf_items, waiting_rows, account, forced_location, db,
                        logger, progress_cb: Optional[Callable[[int, int, str], None]] = None):
        token = MESAuth.login(MES_ACCOUNT, MES_PASSWORD, logger)
        if not token:
            logger.error("Login failed, cannot putaway")
            return []

        shelf_rt: Dict[str, dict] = {}
        for it in shelf_items:
            rt = it.get("rt") or SmartShelfDownloader.extract_rt_from_label(it.get("label_code", ""))
            if not rt: continue
            if rt not in shelf_rt:
                shelf_rt[rt] = {"rt": rt, "part_num": it.get("part_num", ""),
                                "labels": [], "total_qty": 0.0}
            try: q = float(it.get("qty") or it.get("quantity") or 0)
            except (TypeError, ValueError): q = 0.0
            shelf_rt[rt]["labels"].append({
                "label_code": it.get("label_code"),
                "qty": q,
                "in_time": it.get("in_time", ""),
            })
            shelf_rt[rt]["total_qty"] += q

        logger.info(f"Shelf unique RTs: {len(shelf_rt)}")

        wait_map = WaitingReportDownloader.extract_wait_map(waiting_rows)
        logger.info(f"Waiting RTs (MES): {len(wait_map)}")

        already_map = db.get_putaway_qty_map(list(shelf_rt.keys()))

        # 预筛要处理的 RT，用于进度条总数
        todo_rts = [rt for rt in shelf_rt.keys()
                    if rt in wait_map and (shelf_rt[rt]["total_qty"] - already_map.get(rt, 0)) > 0]
        total_todo = len(todo_rts)
        if progress_cb:
            progress_cb(0, max(total_todo, 1), "")
        logger.info(f"RTs to process 待处理 RT: {total_todo}")

        results: List[dict] = []
        skipped_in_this_run = set()
        done_count = 0

        for rt in todo_rts:
            if rt in skipped_in_this_run:
                done_count += 1
                if progress_cb:
                    progress_cb(done_count, total_todo, rt)
                continue
            info = shelf_rt[rt]
            wait_qty  = wait_map[rt]["qty"]
            shelf_qty = info["total_qty"]
            already   = already_map.get(rt, 0.0)
            need      = shelf_qty - already
            put_qty   = min(need, wait_qty)

            logger.info(f"\nRT {rt}: shelf={shelf_qty}, already={already}, "
                        f"need={need}, wait={wait_qty}, put={put_qty}")

            if put_qty <= 0:
                msg = f"nothing to put (need={need}, wait={wait_qty})"
                logger.info(f"  Skip 跳过: {msg}")
                db.log_putaway(rt, info["part_num"], 0, forced_location, "Skipped", msg)
                results.append({"rt": rt, "part_num": info["part_num"],
                                "qty": 0, "location": forced_location,
                                "status": "Skipped", "message": msg})
                done_count += 1
                if progress_cb:
                    progress_cb(done_count, total_todo, rt)
                continue

            success, msg = PutawayProcessor.do_putaway(
                token, rt, forced_location, put_qty, account, logger)

            if not success and ("o_res" in msg and "None" in msg):
                logger.warning("  o_res=None, re-fetch RT info...")
                rt_info = PutawayProcessor.get_rt_info(token, rt, logger)
                if rt_info is not None:
                    try: remain = float(rt_info.get("unit_qty") or 0)
                    except (TypeError, ValueError): remain = 0.0
                    retry_qty = min(put_qty, remain) if remain > 0 else 0
                    if retry_qty > 0 and retry_qty != put_qty:
                        logger.info(f"  Retry with qty={retry_qty}")
                        success, msg = PutawayProcessor.do_putaway(
                            token, rt, forced_location, retry_qty, account, logger)
                        if success: put_qty = retry_qty

            if success:
                logger.info(f"  OK 成功: RT={rt}, qty={put_qty}")
                db.add_putaway_rt(rt, put_qty)
                labels_sorted = sorted(info["labels"],
                    key=lambda x: (x.get("in_time") or "", x.get("label_code") or ""))
                remain = put_qty
                for lb in labels_sorted:
                    if remain <= 0: break
                    db.mark_label_done(lb["label_code"], rt, info["part_num"],
                                       min(lb["qty"], remain), forced_location)
                    remain -= lb["qty"]
                db.log_putaway(rt, info["part_num"], put_qty, forced_location, "Success", msg)
                results.append({"rt": rt, "part_num": info["part_num"],
                                "qty": put_qty, "location": forced_location,
                                "status": "Success", "message": msg})
            else:
                logger.error(f"  FAIL 失败: {msg}")
                skipped_in_this_run.add(rt)
                db.log_putaway(rt, info["part_num"], put_qty, forced_location, "Failed", msg)
                results.append({"rt": rt, "part_num": info["part_num"],
                                "qty": put_qty, "location": forced_location,
                                "status": "Failed", "message": msg})

            done_count += 1
            if progress_cb:
                progress_cb(done_count, total_todo, rt)

        s = sum(1 for r in results if r["status"] == "Success")
        f = sum(1 for r in results if r["status"] == "Failed")
        k = sum(1 for r in results if r["status"] == "Skipped")
        logger.info(f"\nPutaway complete 上架完成: Success {s}, Failed {f}, Skipped {k}")
        return results


# ======================== GUI - 主窗口 ================================
class PutawayApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MES Auto Putaway Tool 自动上架工具")
        self.root.configure(bg=COLOR_BG)
        self._setup_styles()
        self._center_window(1080, 720)
        self.root.minsize(960, 640)

        self.db = LocalDB()
        self.logger = None
        self.is_running = False
        self._shelf_items: List[dict] = []
        self._original_wifi: str = ""

        self._build_ui()
        self._refresh_stats()
        self._refresh_wifi_label()

    def _setup_styles(self):
        s = ttk.Style()
        try: s.theme_use("clam")
        except Exception: pass

        s.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=FONT_BASE)
        s.configure("TFrame", background=COLOR_BG)
        s.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=FONT_BASE)
        s.configure("Muted.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=FONT_SUB)
        s.configure("Status.TLabel", background=COLOR_BG_ALT, foreground=COLOR_PRIMARY_D, font=FONT_BOLD)

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

        # 进度条：用蓝色底 + 蓝色填充
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

    def _build_ui(self):
        # 顶部标题栏
        header = tk.Frame(self.root, bg=COLOR_PRIMARY, height=64)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="MES Auto Putaway Tool  自动上架工具",
                 bg=COLOR_PRIMARY, fg="white", font=("Segoe UI", 16, "bold")
                 ).pack(side=tk.LEFT, padx=20)

        body = ttk.Frame(self.root, style="TFrame", padding=(16, 12, 16, 12))
        body.pack(fill=tk.BOTH, expand=True)

        # ---- 配置卡片 ----
        cfg = ttk.LabelFrame(body, text=" Configuration 配置 ",
                             style="Card.TLabelframe", padding=(14, 10))
        cfg.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(cfg, text="Location 上架库位:", background=COLOR_BG_ALT).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 6))
        self.location_var = tk.StringVar(value="SA00")
        ttk.Entry(cfg, textvariable=self.location_var, width=14).grid(
            row=0, column=1, sticky=tk.W)
        ttk.Label(cfg, text="(e.g. SA00)", style="Muted.TLabel",
                  background=COLOR_BG_ALT).grid(row=0, column=2, sticky=tk.W, padx=(6, 20))

        ttk.Label(cfg, text="WiFi 当前网络:", background=COLOR_BG_ALT).grid(
            row=0, column=3, sticky=tk.W, padx=(0, 6))
        self.wifi_var = tk.StringVar(value="(checking...)")
        ttk.Label(cfg, textvariable=self.wifi_var, background=COLOR_BG_ALT,
                  foreground=COLOR_PRIMARY_D, font=FONT_BOLD).grid(
            row=0, column=4, sticky=tk.W)
        ttk.Button(cfg, text="Refresh 刷新", style="Secondary.TButton",
                   command=self._refresh_wifi_label).grid(row=0, column=5, padx=(8, 6))
        ttk.Button(cfg, text="To AOI 切到AOI", style="Secondary.TButton",
                   command=self._switch_to_aoi).grid(row=0, column=6)

        # ---- 操作栏 ----
        actions = ttk.Frame(body, style="TFrame")
        actions.pack(fill=tk.X, pady=(0, 10))

        self.btn_step1 = ttk.Button(actions, text="Step 1  Download Shelf  下载货架",
                                    style="Primary.TButton", command=self.run_step1)
        self.btn_step1.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_step2 = ttk.Button(actions, text="Step 2  Putaway  上架",
                                    style="Primary.TButton", command=self.run_step2)
        self.btn_step2.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_dbview = ttk.Button(actions, text="Ledger 成功台账",
                                     style="Secondary.TButton", command=self.open_db_viewer)
        self.btn_dbview.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_stats = ttk.Button(actions, text="Stats 统计",
                                    style="Secondary.TButton", command=self.show_stats)
        self.btn_stats.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_reset = ttk.Button(actions, text="Reset RT 重置",
                                    style="Danger.TButton", command=self.reset_rt_dialog)
        self.btn_reset.pack(side=tk.LEFT)

        ttk.Button(actions, text="Clear Log 清空日志", style="Secondary.TButton",
                   command=self.clear_log).pack(side=tk.RIGHT)

        # ---- 状态栏 ----
        status_wrap = tk.Frame(body, bg=COLOR_BG_ALT, highlightbackground=COLOR_BORDER,
                               highlightthickness=1)
        status_wrap.pack(fill=tk.X, pady=(0, 10))
        self.status_var = tk.StringVar(value="Ready 就绪")
        ttk.Label(status_wrap, textvariable=self.status_var, style="Status.TLabel",
                  padding=(10, 6)).pack(anchor=tk.W)

        # ---- 进度条（determinate，带右侧百分比） ----
        prog_wrap = ttk.Frame(body, style="TFrame")
        prog_wrap.pack(fill=tk.X, pady=(0, 10))

        self.progress = ttk.Progressbar(prog_wrap, mode='determinate',
                                        style="Blue.Horizontal.TProgressbar",
                                        maximum=100, value=0)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress_text = tk.StringVar(value="")
        ttk.Label(prog_wrap, textvariable=self.progress_text,
                  background=COLOR_BG, foreground=COLOR_PRIMARY_D,
                  font=FONT_BOLD, width=22).pack(side=tk.LEFT, padx=(10, 0))

        # ---- 日志卡片 ----
        log_frame = ttk.LabelFrame(body, text=" Log 运行日志 ",
                                   style="Card.TLabelframe", padding=(8, 6))
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=16, font=FONT_MONO, wrap=tk.WORD,
            bg=COLOR_BG, fg=COLOR_TEXT, relief="flat",
            highlightthickness=0, insertbackground=COLOR_PRIMARY
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config(COLOR_TEXT, foreground=COLOR_TEXT)
        self.log_text.tag_config(COLOR_WARN, foreground=COLOR_WARN)
        self.log_text.tag_config(COLOR_ERROR, foreground=COLOR_ERROR)
        self.log_text.tag_config(COLOR_SUCCESS, foreground=COLOR_SUCCESS)

        self.logger = GuiLogger(self.log_text)

        foot = ttk.Label(body, text=f"DB: {DB_PATH}", style="Muted.TLabel")
        foot.pack(anchor=tk.W, pady=(6, 0))

    # ---------- 进度条辅助 ----------
    def _set_progress(self, cur: int, total: int, text: str = ""):
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

    # ---------- WiFi ----------
    def _refresh_wifi_label(self):
        ssid = WifiManager.get_current_ssid() or "(not connected 未连接)"
        self.wifi_var.set(ssid)

    def _switch_to_aoi(self):
        self.logger.info(f"Switching WiFi to {WIFI_PROFILE_NAME}... 切换中...")
        if not WifiManager.ensure_profile():
            self.logger.error(f"Profile {WIFI_PROFILE_NAME} missing 无法创建")
            messagebox.showerror("WiFi", f"Cannot ensure profile {WIFI_PROFILE_NAME}")
            return
        # 放到后台线程，避免 UI 卡住
        def worker():
            ok = WifiManager.connect(WIFI_PROFILE_NAME)
            self.logger.info("OK 已连接" if ok else "Failed 连接失败")
            self.root.after(0, self._refresh_wifi_label)
        threading.Thread(target=worker, daemon=True).start()

    def _switch_wifi(self, target: str, logger=None) -> bool:
        lg = logger or self.logger
        if not target: return True
        if WifiManager.get_current_ssid() == target:
            lg.info(f"Already on WiFi: {target}")
            return True
        lg.info(f"Switching WiFi to {target}...")
        ok = WifiManager.connect(target)
        lg.info(f"WiFi switched to {target}" if ok else f"Failed to switch to {target}")
        self.root.after(0, self._refresh_wifi_label)
        return ok

    # ---------- 通用 ----------
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def set_buttons_enabled(self, enabled):
        state = tk.NORMAL if enabled else tk.DISABLED
        for b in (self.btn_step1, self.btn_step2, self.btn_dbview,
                  self.btn_stats, self.btn_reset):
            b.config(state=state)

    def _refresh_stats(self):
        s = self.db.stats()
        self.status_var.set(
            f"Ledger 台账: {s['rt_count']} RT / {s['rt_qty']:.0f} qty | "
            f"labels {s['label_count']} | snapshot {s['snapshot_labels']} | "
            f"logs {s['log_rows']}"
        )

    def show_stats(self):
        s = self.db.stats()
        messagebox.showinfo(
            "Local DB Stats 本地数据库统计",
            f"DB: {DB_PATH}\n\n"
            f"RTs putaway 已上架 RT  : {s['rt_count']}\n"
            f"Total qty 总数量       : {s['rt_qty']:.0f}\n"
            f"Labels 已上架 label    : {s['label_count']}\n"
            f"Snapshot labels 快照   : {s['snapshot_labels']}\n"
            f"Log rows 日志行        : {s['log_rows']}"
        )

    def reset_rt_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Reset RT 重置 RT")
        dlg.configure(bg=COLOR_BG)
        dlg.transient(self.root); dlg.grab_set()
        w, h = 480, 180
        dlg.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        ttk.Label(dlg, text="RT to reset (removes ledger for this RT):\n"
                            "要重置的 RT（将删除该 RT 的台账）:",
                  background=COLOR_BG).pack(anchor=tk.W, padx=14, pady=(12, 4))
        var = tk.StringVar()
        e = ttk.Entry(dlg, textvariable=var, width=56)
        e.pack(padx=14, fill=tk.X); e.focus_set()

        def do():
            rt = var.get().strip()
            if not rt: return
            if messagebox.askyesno("Confirm 确认", f"Reset RT {rt} ?"):
                self.db.reset_rt(rt)
                self.logger.warning(f"[MANUAL RESET] RT={rt}")
                dlg.destroy(); self._refresh_stats()

        btns = ttk.Frame(dlg, style="TFrame")
        btns.pack(pady=14)
        ttk.Button(btns, text="Reset 重置", style="Danger.TButton",
                   command=do).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Cancel 取消", style="Secondary.TButton",
                   command=dlg.destroy).pack(side=tk.LEFT, padx=6)

    def open_db_viewer(self):
        LedgerViewer(self.root, self.db)

    # ---------- Step 1 ----------
    def run_step1(self):
        if self.is_running: return
        self.is_running = True
        self.set_buttons_enabled(False)
        self._reset_progress()
        self.status_var.set("Step 1: switching WiFi + downloading shelf... 切换 WiFi + 下载货架...")
        threading.Thread(target=self._run_step1, daemon=True).start()

    def _run_step1(self):
        try:
            self._original_wifi = WifiManager.get_current_ssid()
            self.logger.info(f"WiFi before: {self._original_wifi or '(none)'}")

            if not WifiManager.ensure_profile():
                self.logger.error(f"WiFi profile {WIFI_PROFILE_NAME} missing")
                self.status_var.set("Step 1 failed: WiFi profile missing")
                return
            if not self._switch_wifi(WIFI_PROFILE_NAME):
                self.logger.error("Cannot switch to AOI, abort 无法切换到 AOI")
                self.status_var.set("Step 1 failed: WiFi switch failed")
                return

            self.logger.info("=" * 46)
            self.logger.info("STEP 1: Download Smart Shelf Data 下载货架数据")
            self.logger.info("=" * 46)

            def progress_cb(cur, total):
                self._set_progress(cur, total, "shelf 货架")

            items = SmartShelfDownloader.download_all_raw(self.logger, progress_cb)
            self._shelf_items = items

            snapshot_time = datetime.now().isoformat(timespec="seconds")
            self.db.save_shelf_snapshot(items, snapshot_time)
            self.logger.info(f"Snapshot saved 快照已写入: {len(items)} items")

            total_qty = 0.0; rts = set()
            for it in items:
                try: total_qty += float(it.get("quantity") or 0)
                except (TypeError, ValueError): pass
                rt = SmartShelfDownloader.extract_rt_from_label(it.get("label_code", ""))
                if rt: rts.add(rt)
            self.logger.info(f"总库存条数 items: {len(items)}  "
                             f"总数量 qty: {total_qty:.0f}  唯一 RT: {len(rts)}")
            self.status_var.set(
                f"Step 1 done 完成: {len(items)} items, qty {total_qty:.0f}, {len(rts)} RTs")
        except Exception as e:
            self.logger.exception(f"Step 1 failed: {e}")
            self.status_var.set("Step 1 failed 失败")
        finally:
            if self._original_wifi:
                self._switch_wifi(self._original_wifi)
            self._reset_progress()
            self.is_running = False
            self.set_buttons_enabled(True)
            self._refresh_stats()

    # ---------- Step 2 ----------
    def run_step2(self):
        if self.is_running: return
        if not self._shelf_items:
            rows = self.db.load_latest_snapshot()
            if rows:
                self._shelf_items = rows
                self.logger.info(f"Loaded snapshot from DB 从数据库读取快照: {len(rows)}")
            else:
                if messagebox.askyesno("Missing Data 缺少数据",
                                       "No shelf data. Run Step 1 first?\n"
                                       "没有货架数据，先跑 Step 1?"):
                    self.run_step1()
                return
        self.is_running = True
        self.set_buttons_enabled(False)
        self._reset_progress()
        self.status_var.set("Step 2: downloading report + putaway... 下载待上架 + 上架...")
        threading.Thread(target=self._run_step2, daemon=True).start()

    def _run_step2(self):
        try:
            self.logger.info("=" * 46)
            self.logger.info("STEP 2: Report & Putaway 待上架 + 上架")
            self.logger.info("=" * 46)

            forced_location = self.location_var.get().strip() or "SA00"
            self.location_var.set(forced_location)
            self.logger.info(f"Location 库位: {forced_location}")

            token = MESAuth.login(MES_ACCOUNT, MES_PASSWORD, self.logger)
            if not token:
                self.logger.error("Login failed 登录失败"); return

            rows = WaitingReportDownloader.download_all(token, self.logger)
            if not rows:
                self.logger.warning("No waiting report data 无待上架数据"); return

            def progress_cb(cur, total, rt):
                self._set_progress(cur, total, f"RT {rt}" if rt else "")

            results = PutawayProcessor.process_putaway(
                self._shelf_items, rows, MES_ACCOUNT, forced_location,
                self.db, self.logger, progress_cb)

            if results:
                s = sum(1 for r in results if r["status"] == "Success")
                self.status_var.set(f"Step 2 done 完成: {s} success / {len(results)} total")
            else:
                self.status_var.set("Step 2 done 完成: nothing to do 无任务")
        except Exception as e:
            self.logger.exception(f"Step 2 failed: {e}")
            self.status_var.set("Step 2 failed 失败")
        finally:
            self._reset_progress()
            self.is_running = False
            self.set_buttons_enabled(True)
            self._refresh_stats()


# ======================== Ledger Viewer ================================
class LedgerViewer:
    def __init__(self, parent, db: LocalDB):
        self.db = db
        self.win = tk.Toplevel(parent)
        self.win.title("Ledger 成功台账")
        self.win.configure(bg=COLOR_BG)
        self.win.transient(parent)

        w, h = 1060, 660
        self.win.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - w) // 2
        y = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.win.geometry(f"{w}x{h}+{max(0,x)}+{max(0,y)}")

        top = tk.Frame(self.win, bg=COLOR_BG_ALT,
                       highlightbackground=COLOR_BORDER, highlightthickness=1)
        top.pack(fill=tk.X, padx=10, pady=(10, 8))

        tk.Label(top, text="RT:", bg=COLOR_BG_ALT, fg=COLOR_TEXT, font=FONT_BOLD
                 ).pack(side=tk.LEFT, padx=(10, 4), pady=8)
        self.rt_filter = tk.StringVar()
        ttk.Entry(top, textvariable=self.rt_filter, width=18).pack(side=tk.LEFT, padx=(0, 12))

        tk.Label(top, text="Part:", bg=COLOR_BG_ALT, fg=COLOR_TEXT, font=FONT_BOLD
                 ).pack(side=tk.LEFT, padx=(0, 4))
        self.part_filter = tk.StringVar()
        ttk.Entry(top, textvariable=self.part_filter, width=18).pack(side=tk.LEFT, padx=(0, 12))

        tk.Label(top, text="Result:", bg=COLOR_BG_ALT, fg=COLOR_TEXT, font=FONT_BOLD
                 ).pack(side=tk.LEFT, padx=(0, 4))
        self.result_filter = tk.StringVar(value="All")
        ttk.Combobox(top, textvariable=self.result_filter, width=10, state="readonly",
                     values=["All", "Success", "Failed", "Skipped"]).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Button(top, text="Query 查询", style="Primary.TButton",
                   command=self.query).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Refresh 刷新", style="Secondary.TButton",
                   command=self.query).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Export Excel 导出", style="Secondary.TButton",
                   command=self.export).pack(side=tk.LEFT)

        nb = ttk.Notebook(self.win)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        rt_tab = ttk.Frame(nb, style="TFrame")
        nb.add(rt_tab, text="Success RT 成功 RT")
        self.rt_tree = self._make_tree(
            rt_tab,
            cols=[("rt", "RT", 240), ("qty", "Qty 数量", 120),
                  ("last_time", "Last Time 最后时间", 200)])

        lb_tab = ttk.Frame(nb, style="TFrame")
        nb.add(lb_tab, text="Success Label 成功 Label")
        self.lb_tree = self._make_tree(
            lb_tab,
            cols=[("label_code", "Label Code 唯一码", 380),
                  ("rt", "RT", 110), ("part_num", "Part No 料号", 150),
                  ("qty", "Qty 数量", 80), ("location", "Location 库位", 80),
                  ("putaway_time", "Time 时间", 170)])

        lg_tab = ttk.Frame(nb, style="TFrame")
        nb.add(lg_tab, text="Logs 日志")
        self.lg_tree = self._make_tree(
            lg_tab,
            cols=[("log_time", "Time 时间", 160),
                  ("rt", "RT", 130), ("part_num", "Part No 料号", 150),
                  ("qty", "Qty 数量", 80), ("location", "Loc 库位", 70),
                  ("result", "Result 结果", 90),
                  ("message", "Message 消息", 320)])

        self.query()

    def _make_tree(self, parent, cols):
        wrap = tk.Frame(parent, bg=COLOR_BG)
        wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        col_ids = [c[0] for c in cols]
        tree = ttk.Treeview(wrap, columns=col_ids, show="headings")
        for cid, text, w in cols:
            tree.heading(cid, text=text)
            tree.column(cid, width=w, anchor=tk.W, stretch=(cid in ("label_code", "message")))
        sb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        return tree

    def _clear(self, tree):
        for i in tree.get_children():
            tree.delete(i)

    def query(self):
        rt_f = self.rt_filter.get().strip()
        part_f = self.part_filter.get().strip()
        result_f = self.result_filter.get().strip()

        self._clear(self.rt_tree)
        for r in self.db.query_success_rt(rt_f, part_f, limit=5000):
            self.rt_tree.insert("", tk.END, values=(
                r["rt"], f"{r['putaway_qty']:.0f}", r["last_time"] or ""))

        self._clear(self.lb_tree)
        for r in self.db.query_success_labels(rt_f, part_f, limit=5000):
            self.lb_tree.insert("", tk.END, values=(
                r["label_code"], r["rt"], r["part_num"], f"{r['qty']:.0f}",
                r["location"], r["putaway_time"]))

        self._clear(self.lg_tree)
        rfilter = "" if result_f in ("", "All") else result_f
        for r in self.db.query_log(rt_f, rfilter, limit=5000):
            self.lg_tree.insert("", tk.END, values=(
                r["log_time"], r["rt"], r["part_num"], f"{r['qty']:.0f}",
                r["location"], r["result"], r["message"] or ""))

    def export(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            title="Export Ledger 导出台账", defaultextension=".xlsx",
            initialfile=f"ledger_{ts}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not path: return

        rt_f = self.rt_filter.get().strip()
        part_f = self.part_filter.get().strip()
        result_f = self.result_filter.get().strip()

        wb = Workbook()
        ws1 = wb.active; ws1.title = "Success RT"
        ws1.append(["RT", "Qty", "Last Time"])
        for r in self.db.query_success_rt(rt_f, part_f, limit=100000):
            ws1.append([r["rt"], r["putaway_qty"], r["last_time"]])

        ws2 = wb.create_sheet("Success Label")
        ws2.append(["Label Code", "RT", "Part No", "Qty", "Location", "Putaway Time"])
        for r in self.db.query_success_labels(rt_f, part_f, limit=100000):
            ws2.append([r["label_code"], r["rt"], r["part_num"],
                        r["qty"], r["location"], r["putaway_time"]])

        ws3 = wb.create_sheet("Logs")
        ws3.append(["Time", "RT", "Part No", "Qty", "Location", "Result", "Message"])
        rfilter = "" if result_f in ("", "All") else result_f
        for r in self.db.query_log(rt_f, rfilter, limit=100000):
            ws3.append([r["log_time"], r["rt"], r["part_num"], r["qty"],
                        r["location"], r["result"], r["message"]])

        wb.save(path)
        messagebox.showinfo("Export 导出", f"Saved 已保存:\n{path}")


# ======================== Main ================================
def main():
    root = tk.Tk()
    PutawayApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()