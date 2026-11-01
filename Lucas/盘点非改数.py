#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
库存偏差分配工具（GUI）

偏差输入支持两种方式：
1. 直接粘贴到文本框（从 Excel 复制两列：PN、QTY，制表符分隔）；
2. 选择偏差 Excel 文件。

文本框有内容时优先使用文本框。

逻辑：
1. 偏差按 PN 汇总；
2. 库存按 PN 分组，组内按 Batch 排序；
3. 负偏差从第一行扣，扣到 0 继续下一行；正偏差全部加到第一行；
4. 输出 Excel，并写警告文件。

导出选项：
- 全部行
- 仅修改过的行（Book quantity 有变化），并带原值对照列

依赖：
    pip install openpyxl pandas
"""

import os
import threading
from decimal import Decimal, getcontext

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

getcontext().prec = 50


# ---------- 工具函数 ----------

def to_decimal(v) -> Decimal:
    if v is None:
        return Decimal("0")
    if isinstance(v, float) and pd.isna(v):
        return Decimal("0")
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return Decimal("0")
    s = s.replace(",", "")
    return Decimal(s)


def fmt_decimal(d: Decimal):
    if d == d.to_integral_value():
        return int(d)
    return float(d)


def parse_pasted_offsets(text: str, log):
    """
    解析粘贴的偏差文本。
    支持：
      PN<TAB>QTY
      PN,QTY
      PN 空格 QTY
    忽略表头行（PN/QTY/Material/Quantity 等）。
    """
    offsets = {}
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return offsets

    for ln in lines:
        if "\t" in ln:
            parts = ln.split("\t")
        elif "," in ln:
            parts = ln.split(",")
        else:
            parts = ln.split()

        parts = [p.strip() for p in parts if p.strip() != ""]
        if len(parts) < 2:
            continue

        pn = parts[0]
        qty_str = parts[1]

        if pn.lower() in ("pn", "material", "料号", "item"):
            continue
        if qty_str.lower() in ("qty", "quantity", "数量", "偏差"):
            continue

        try:
            qty = to_decimal(qty_str)
        except Exception:
            log(f"跳过无法解析的行：{ln}")
            continue

        offsets[pn] = offsets.get(pn, Decimal("0")) + qty

    return offsets


def read_offsets_from_excel(path, log):
    off_df = pd.read_excel(path)
    pn_col = None
    qty_col = None
    for c in off_df.columns:
        name = str(c).strip().lower()
        if name in ("pn", "material", "料号", "item"):
            pn_col = c
        if name in ("qty", "quantity", "偏差", "数量"):
            qty_col = c
    if pn_col is None or qty_col is None:
        raise ValueError(f"偏差文件必须包含 PN 和 QTY 列，实际列：{list(off_df.columns)}")

    offsets = {}
    for _, row in off_df.iterrows():
        pn = str(row[pn_col]).strip()
        if not pn or pn.lower() == "nan":
            continue
        qty = to_decimal(row[qty_col])
        offsets[pn] = offsets.get(pn, Decimal("0")) + qty
    return offsets


# ---------- 核心处理 ----------

def process(inventory_path, offsets, output_path, group_mode, sort_mode,
            export_mode, log):
    log(f"偏差表共 {len(offsets)} 个 PN")

    log("读取库存文件...")
    inv_df = pd.read_excel(inventory_path)

    required = ["Phys. Inventory Doc.", "Material", "Batch", "Storage Location", "Book quantity"]
    for col in required:
        if col not in inv_df.columns:
            raise ValueError(f"库存文件缺少列：{col}，实际列：{list(inv_df.columns)}")

    inv_df = inv_df.reset_index().rename(columns={"index": "_orig_idx"})
    log(f"库存表共 {len(inv_df)} 行")

    if group_mode == "material":
        key_cols = ["Material"]
    else:
        key_cols = ["Phys. Inventory Doc.", "Material", "Storage Location"]

    if sort_mode == "batch_asc":
        sort_cols = key_cols + ["Batch"]
        ascending = [True] * len(key_cols) + [True]
    elif sort_mode == "batch_desc":
        sort_cols = key_cols + ["Batch"]
        ascending = [True] * len(key_cols) + [False]
    elif sort_mode == "book_asc":
        sort_cols = key_cols + ["Book quantity"]
        ascending = [True] * len(key_cols) + [True]
    elif sort_mode == "book_desc":
        sort_cols = key_cols + ["Book quantity"]
        ascending = [True] * len(key_cols) + [False]
    else:
        sort_cols = key_cols + ["Batch"]
        ascending = [True] * len(key_cols) + [True]

    inv_df = inv_df.sort_values(by=sort_cols, ascending=ascending, kind="mergesort").reset_index(drop=True)
    inv_df["_book_in"] = inv_df["Book quantity"].apply(to_decimal)
    inv_df["_book_out"] = inv_df["_book_in"]

    warnings = []
    grouped = inv_df.groupby(key_cols, sort=False)

    for key, grp in grouped:
        if not isinstance(key, tuple):
            key = (key,)
        pn = str(key[0]).strip() if group_mode == "material" else str(key[1]).strip()

        total_offset = offsets.get(pn, Decimal("0"))
        if total_offset == 0:
            continue

        idxs = grp.index.tolist()

        if total_offset > 0:
            first = idxs[0]
            inv_df.at[first, "_book_out"] = inv_df.at[first, "_book_out"] + total_offset
        else:
            remaining = -total_offset
            for i in idxs:
                if remaining <= 0:
                    break
                book = inv_df.at[i, "_book_out"]
                if book <= 0:
                    continue
                deduct = min(book, remaining)
                inv_df.at[i, "_book_out"] = book - deduct
                remaining -= deduct

            if remaining > 0:
                warnings.append((key, pn, total_offset, remaining))

    # 标记修改行
    inv_df["_changed"] = inv_df["_book_in"] != inv_df["_book_out"]

    # 写回 Book quantity
    inv_df["Book quantity"] = inv_df["_book_out"].apply(fmt_decimal)

    # 恢复原始行顺序
    inv_df = inv_df.sort_values("_orig_idx").reset_index(drop=True)

    # 根据导出模式选择行
    if export_mode == "changed":
        out_df = inv_df[inv_df["_changed"]].copy()
        log(f"仅导出修改行：{len(out_df)} 行")
    else:
        out_df = inv_df.copy()
        log(f"导出全部行：{len(out_df)} 行")

    # 为导出添加原值对照列
    out_df["原Book quantity"] = out_df["_book_in"].apply(fmt_decimal)

    # 清理内部列
    drop_cols = ["_orig_idx", "_book_in", "_book_out", "_changed"]
    out_df = out_df.drop(columns=[c for c in drop_cols if c in out_df.columns])

    log("写出结果...")
    out_df.to_excel(output_path, index=False)

    report_path = os.path.splitext(output_path)[0] + "_warnings.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        if not warnings:
            f.write("没有未分配完的负偏差。\n")
        else:
            f.write("以下组负偏差大于该组 Book quantity 总和，剩余未分配：\n")
            for key, pn, total_offset, remaining in warnings:
                f.write(f"key={key}, PN={pn}, total_offset={total_offset}, remaining={remaining}\n")

    log(f"完成：{output_path}")
    log(f"警告文件：{report_path}")

    if warnings:
        log(f"⚠ 有 {len(warnings)} 组负偏差未扣完，见警告文件")
    else:
        log("✅ 所有偏差已分配完毕")


# ---------- GUI ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("库存偏差分配工具")
        self.geometry("1000x860")

        self.inventory_path = tk.StringVar()
        self.offset_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.group_mode = tk.StringVar(value="material")
        self.sort_mode = tk.StringVar(value="batch_asc")
        self.export_mode = tk.StringVar(value="all")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # 文件
        frm_files = ttk.LabelFrame(self, text="文件")
        frm_files.pack(fill="x", **pad)
        self._file_row(frm_files, 0, "库存文件（Excel）", self.inventory_path, self.pick_inventory)
        self._file_row(frm_files, 1, "偏差文件（可选）", self.offset_path, self.pick_offset)
        self._file_row(frm_files, 2, "输出文件（Excel）", self.output_path, self.pick_output)

        # 偏差粘贴
        frm_paste = ttk.LabelFrame(self, text="偏差输入（直接粘贴，优先于偏差文件）—— 两列：PN 和 QTY，从 Excel 复制即可")
        frm_paste.pack(fill="both", expand=False, **pad)
        self.offset_text = scrolledtext.ScrolledText(frm_paste, height=8)
        self.offset_text.pack(fill="both", expand=True, padx=4, pady=4)

        # 选项
        frm_opts = ttk.LabelFrame(self, text="选项")
        frm_opts.pack(fill="x", **pad)

        ttk.Label(frm_opts, text="分组键：").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Radiobutton(frm_opts, text="按 Material（推荐，全局汇总偏差）",
                        variable=self.group_mode, value="material").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(frm_opts, text="按 盘点单+Material+库位",
                        variable=self.group_mode, value="doc_material_storage").grid(row=0, column=2, sticky="w")

        ttk.Label(frm_opts, text="排序：").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Radiobutton(frm_opts, text="Batch 升序", variable=self.sort_mode, value="batch_asc").grid(row=1, column=1, sticky="w")
        ttk.Radiobutton(frm_opts, text="Batch 降序", variable=self.sort_mode, value="batch_desc").grid(row=1, column=2, sticky="w")
        ttk.Radiobutton(frm_opts, text="Book quantity 升序", variable=self.sort_mode, value="book_asc").grid(row=1, column=3, sticky="w")
        ttk.Radiobutton(frm_opts, text="Book quantity 降序", variable=self.sort_mode, value="book_desc").grid(row=1, column=4, sticky="w")

        ttk.Label(frm_opts, text="导出：").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        ttk.Radiobutton(frm_opts, text="全部行", variable=self.export_mode, value="all").grid(row=2, column=1, sticky="w")
        ttk.Radiobutton(frm_opts, text="仅修改过的行（带原值对照列）", variable=self.export_mode, value="changed").grid(row=2, column=2, sticky="w", columnspan=2)

        # 按钮
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", **pad)
        self.btn_run = ttk.Button(frm_btn, text="开始处理", command=self.run)
        self.btn_run.pack(side="left", padx=8)
        ttk.Button(frm_btn, text="清空日志", command=self.clear_log).pack(side="left")
        ttk.Button(frm_btn, text="清空偏差输入", command=lambda: self.offset_text.delete("1.0", "end")).pack(side="left", padx=8)

        # 日志
        frm_log = ttk.LabelFrame(self, text="日志")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_box = scrolledtext.ScrolledText(frm_log, height=15)
        self.log_box.pack(fill="both", expand=True, padx=4, pady=4)

    def _file_row(self, parent, row, label, var, cmd):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(parent, textvariable=var, width=80).grid(row=row, column=1, padx=4, pady=4)
        ttk.Button(parent, text="浏览...", command=cmd).grid(row=row, column=2, padx=8, pady=4)

    def pick_inventory(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.inventory_path.set(p)

    def pick_offset(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.offset_path.set(p)

    def pick_output(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.output_path.set(p)

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def run(self):
        inv = self.inventory_path.get().strip()
        out = self.output_path.get().strip()
        paste_text = self.offset_text.get("1.0", "end").strip()

        if not inv or not os.path.exists(inv):
            messagebox.showerror("错误", "请选择有效的库存文件")
            return
        if not out:
            messagebox.showerror("错误", "请选择输出文件")
            return

        self.btn_run.config(state="disabled")
        self.clear_log()

        def worker():
            try:
                if paste_text:
                    self.log("使用粘贴的偏差数据...")
                    offsets = parse_pasted_offsets(paste_text, self.log)
                    if not offsets:
                        raise ValueError("粘贴的偏差数据解析为空，请检查格式（PN<TAB>QTY）")
                else:
                    off = self.offset_path.get().strip()
                    if not off or not os.path.exists(off):
                        raise ValueError("请粘贴偏差数据，或选择偏差文件")
                    self.log("使用偏差文件...")
                    offsets = read_offsets_from_excel(off, self.log)

                process(
                    inv, offsets, out,
                    self.group_mode.get(),
                    self.sort_mode.get(),
                    self.export_mode.get(),
                    self.log,
                )
                messagebox.showinfo("完成", f"处理完成：\n{out}")
            except Exception as e:
                self.log(f"❌ 错误：{e}")
                messagebox.showerror("错误", str(e))
            finally:
                self.btn_run.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
