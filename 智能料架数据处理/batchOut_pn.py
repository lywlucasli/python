# -*- coding: utf-8 -*-
"""
Batch Outbound Tool (tkinter version)
Features:
1. Query stock: download all material data from server into memory
2. Count occurrences of "part_num", show in descending order
3. Select a row, click "Batch Outbound" to outbound that part_num
4. Second confirmation required before outbound
"""
import sys
import time
import threading
from collections import Counter

import requests
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# ==================== Config ====================
BASE_URL = "http://192.168.9.100:8081"
SEARCH_URL = f"{BASE_URL}/mate/searchByInfo"
OUT_URL = f"{BASE_URL}/mate/outByPosition"

# Warehouse type IDs used for outbound
WAREHOUSE_NAME_IDS = [43, 44]

PAGE_SIZE = 10000
USER_NAME = "NAF684"
USER_GROUP_POWER = "1"
USER_LANGUAGE = "zh"
REQUEST_INTERVAL = 0.2
MAX_RETRY = 3
RETRY_DELAY = 2

# Base search payload
SEARCH_BASE = {
    "part_num": "",
    "warehouse_category_id": 1,
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

# Fixed outbound params
OUT_TEMPLATE = {
    "color": 3,
    "user_name": USER_NAME,
    "user_group_power": USER_GROUP_POWER,
    "user_language": USER_LANGUAGE,
}


# ==================== Network ====================
def build_search_payload(start, rows=PAGE_SIZE):
    payload = SEARCH_BASE.copy()
    payload["start"] = start
    payload["rows"] = rows
    return payload


def fetch_page(start):
    """Fetch a single page with retry"""
    payload = build_search_payload(start)
    for attempt in range(1, MAX_RETRY + 1):
        try:
            resp = requests.post(SEARCH_URL, json=payload, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            if data.get("result") == 1:
                return data
            else:
                print(f"  [WARN] attempt {attempt} result != 1: {data}")
        except Exception as e:
            print(f"  [WARN] attempt {attempt} request error: {e}")
        if attempt < MAX_RETRY:
            time.sleep(RETRY_DELAY)
    return None


def fetch_all_materials(log_cb=None):
    """Download all material data"""
    all_materials = []
    start = 0
    total = None
    consecutive_empty = 0

    while True:
        data = fetch_page(start)
        if data is None:
            if log_cb:
                log_cb(f"[ERROR] start={start} failed after {MAX_RETRY} retries, stop")
            return all_materials, False

        mate_list = data.get("mate_list", [])

        if total is None:
            total = data.get("rows", 0)
            if log_cb:
                log_cb(f"Server reported total records: {total}")

        if not mate_list:
            consecutive_empty += 1
            if log_cb:
                log_cb(f"  [INFO] start={start} empty page ({consecutive_empty}/3)")
            if consecutive_empty >= 3:
                if log_cb:
                    log_cb("3 consecutive empty pages, data fully loaded")
                break
            start += PAGE_SIZE
            time.sleep(0.5)
            continue

        consecutive_empty = 0
        all_materials.extend(mate_list)
        if log_cb:
            log_cb(f"Loaded {len(all_materials)} records "
                   f"(this page {len(mate_list)}, start={start})")

        if total > 0 and len(all_materials) > total * 1.5:
            if log_cb:
                log_cb("[WARN] Loaded > 1.5x declared total, force stop")
            break

        start += PAGE_SIZE
        time.sleep(0.1)

    return all_materials, True


# ==================== Outbound ====================
def search_materials_by_part_num(part_num):
    """Search materials by part_num in target warehouses (for outbound)"""
    materials = []
    for warehouse_name_id in WAREHOUSE_NAME_IDS:
        payload = build_search_payload(0)
        payload["part_num"] = part_num
        payload["save_id"] = ""
        payload["warehouse_name_id"] = warehouse_name_id
        start = 0
        rows = payload["rows"]

        while True:
            payload["start"] = start
            try:
                resp = requests.post(SEARCH_URL, json=payload, timeout=10)
            except Exception as e:
                print(f"    Request exception: {e}")
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
    """Call outbound API"""
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


# ==================== Main Window ====================
class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Batch Outbound Tool")
        self.root.geometry("1000x700")

        self.all_materials = []
        self.part_stats = []          # [(part_num, count), ...] descending
        self._busy = False

        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        # Top buttons
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill=tk.X)

        self.btn_download = ttk.Button(top, text="(1) Query Stock",
                                       command=self.on_query_stock, width=24)
        self.btn_download.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_out = ttk.Button(top, text="(2) Batch Outbound (selected)",
                                  command=self.on_batch_out, width=28,
                                  state=tk.DISABLED)
        self.btn_out.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_status = ttk.Label(top, text="No data loaded", foreground="#888")
        self.lbl_status.pack(side=tk.LEFT, padx=8)

        # Progress bar
        self.progress = ttk.Progressbar(self.root, mode="determinate", maximum=100)
        self.progress.pack(fill=tk.X, padx=8, pady=(0, 4))

        # Middle: table + log
        paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # --- Table ---
        table_frame = ttk.Frame(paned)
        paned.add(table_frame, weight=3)

        columns = ("part_num", "count")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                 selectmode="browse")
        self.tree.heading("part_num", text="Part Number (part_num)")
        self.tree.heading("count", text="Occurrences")
        self.tree.column("part_num", anchor=tk.W, stretch=True, width=600)
        self.tree.column("count", anchor=tk.CENTER, stretch=False, width=120)

        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.tag_configure("top", background="#FFF3B0")

        # --- Log ---
        log_frame = ttk.Frame(paned)
        paned.add(log_frame, weight=1)

        self.log = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD,
                                             font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)
        self.log.configure(state=tk.DISABLED)

    # ---------------- Log ----------------
    def append_log(self, msg):
        def _do():
            self.log.configure(state=tk.NORMAL)
            self.log.insert(tk.END, msg + "\n")
            self.log.see(tk.END)
            self.log.configure(state=tk.DISABLED)
        self.root.after(0, _do)

    # ---------------- Query Stock ----------------
    def on_query_stock(self):
        if self._busy:
            return
        self._busy = True
        self.btn_download.configure(state=tk.DISABLED)
        self.btn_out.configure(state=tk.DISABLED)
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.all_materials = []
        self.part_stats = []
        self.lbl_status.configure(text="Querying stock...")
        self.progress.configure(mode="indeterminate")
        self.progress.start(15)
        self.append_log("=" * 60)
        self.append_log("Start querying stock ...")

        threading.Thread(target=self._query_thread, daemon=True).start()

    def _query_thread(self):
        materials, ok = fetch_all_materials(log_cb=self.append_log)
        self.root.after(0, self._on_query_finished, materials, ok)

    def _on_query_finished(self, materials, ok):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=100)
        self._busy = False
        self.btn_download.configure(state=tk.NORMAL)

        if not ok:
            self.lbl_status.configure(text="Query failed")
            self.append_log("Query failed, please check network / server")
            messagebox.showerror("Error", "Failed to query stock. Please check the log.")
            return

        self.all_materials = materials
        self.append_log(f"Query complete, total {len(materials)} records")

        # Count part_num occurrences
        counter = Counter()
        for item in materials:
            part = item.get("part_num", "")
            if part:
                counter[part] += 1

        self.part_stats = counter.most_common()   # already descending
        self.append_log(f"{len(self.part_stats)} distinct part numbers, "
                        f"sorted by occurrences descending")

        self._populate_table()

        if self.part_stats:
            self.btn_out.configure(state=tk.NORMAL)
            self.lbl_status.configure(
                text=f"Loaded {len(materials)} records, "
                     f"{len(self.part_stats)} part numbers")
        else:
            self.lbl_status.configure(text="No data")
            messagebox.showinfo("Info", "No material data retrieved.")

    def _populate_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        max_count = self.part_stats[0][1] if self.part_stats else 1
        for part, cnt in self.part_stats:
            tags = ("top",) if cnt == max_count else ()
            self.tree.insert("", tk.END, values=(part, cnt), tags=tags)

    # ---------------- Batch Outbound ----------------
    def on_batch_out(self):
        if self._busy:
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a part number first.")
            return

        values = self.tree.item(sel[0], "values")
        part_num = values[0]
        count = values[1]

        # Second confirmation
        reply = messagebox.askyesno(
            "Confirm",
            f"Are you sure you want to batch outbound this part number?\n\n"
            f"Part number: {part_num}\n"
            f"Occurrences in queried data: {count}\n\n"
            f"Target warehouse IDs: {WAREHOUSE_NAME_IDS}\n"
            f"This will call the outbound API. Please confirm!",
            icon="warning",
            default="no",
        )
        if not reply:
            self.append_log("User cancelled outbound operation")
            return

        self._busy = True
        self.btn_download.configure(state=tk.DISABLED)
        self.btn_out.configure(state=tk.DISABLED)
        self.progress.configure(mode="determinate", value=0)
        self.append_log("=" * 60)
        self.append_log(f"Start batch outbound: part_num={part_num}")

        threading.Thread(target=self._out_thread, args=(part_num,), daemon=True).start()

    def _out_thread(self, part_num):
        self.append_log(f"Searching materials for part_num={part_num} ...")
        materials = search_materials_by_part_num(part_num)
        if not materials:
            self.append_log("No materials found for outbound")
            self.root.after(0, self._on_out_finished, 0, 0, [])
            return

        total = len(materials)
        self.append_log(f"Found {total} materials, starting batch outbound...")

        success = 0
        fail = 0
        fail_details = []
        for idx, mat in enumerate(materials, start=1):
            shelf_id = mat.get("shelf_id")
            position = mat.get("position")
            if shelf_id is None or position is None:
                fail += 1
                fail_details.append(f"[{idx}] missing shelf_id or position")
                self._update_progress(idx, total)
                continue

            ok, msg = out_material(shelf_id, position)
            if ok:
                success += 1
                self.append_log(f"[{idx}/{total}] OK  shelf_id={shelf_id} position={position}")
            else:
                fail += 1
                fail_details.append(
                    f"[{idx}] shelf_id={shelf_id} position={position} FAILED: {msg}")
                self.append_log(
                    f"[{idx}/{total}] FAIL shelf_id={shelf_id} position={position} -> {msg}")

            self._update_progress(idx, total)
            time.sleep(REQUEST_INTERVAL)

        self.root.after(0, self._on_out_finished, success, fail, fail_details)

    def _update_progress(self, current, total):
        if total > 0:
            val = int(current * 100 / total)
            self.root.after(0, lambda v=val: self.progress.configure(value=v))

    def _on_out_finished(self, success, fail, details):
        self.progress.configure(value=100)
        self._busy = False
        self.btn_download.configure(state=tk.NORMAL)
        self.btn_out.configure(state=tk.NORMAL)
        self.append_log(f"Outbound finished: success {success}, failed {fail}")
        if details:
            self.append_log("Failure details:")
            for d in details:
                self.append_log("  " + d)

        messagebox.showinfo(
            "Outbound Complete",
            f"Outbound finished\nSuccess: {success}\nFailed: {fail}"
        )


# ==================== Entry ====================
def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()