import sys
import subprocess
import os

# ============ 自动安装依赖 ============
def install_dependencies():
    try:
        import ttkbootstrap
    except ImportError:
        print("正在首次自动安装界面美化组件 (ttkbootstrap)，请稍候...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "ttkbootstrap"])
        print("组件安装完成，正在启动程序...")

install_dependencies()

# ============ 以下为正式业务代码 ============
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import json
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

DATA_FILE = "wip_save_data.json"

# ============ 核心解析逻辑 ============
def parse_order_groups(text):
    """① 解析工单组数据：只要有工单号，统统提取出来加入列表"""
    lines = [l for l in text.splitlines() if l.strip()]
    orders = []

    for l in lines:
        if 'Order' in l and 'WIP QTY' in l:
            continue

        if '\t' in l:
            parts = l.split('\t')
        else:
            parts = re.split(r'\s{2,}', l.strip())

        parts = [p.strip() for p in parts]

        order_num = ""
        for p in parts[:3]:
            match = re.search(r'\b\d{9,12}\b', p)
            if match:
                order_num = match.group(0)
                break

        if order_num and order_num not in orders:
            orders.append(order_num)

    return orders

def parse_wip(text):
    """② 解析实际WIP数据：严格按 Tab 切分，按表头列名定位 Diff Qty"""
    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        return []

    # ---------- 1. 找表头，按 Tab 切分，定位列索引 ----------
    header_idx = None
    header_cols = []
    for i, l in enumerate(lines):
        if 'Diff Qty' in l and 'Order' in l:
            header_idx = i
            header_cols = [c.strip() for c in l.split('\t')]
            break

    if header_idx is None:
        return []

    def find_col(name):
        target = name.strip().lower().replace(' ', '')
        for idx, c in enumerate(header_cols):
            if c.strip().lower().replace(' ', '') == target:
                return idx
        return -1

    idx_order     = find_col('Order')
    idx_material  = find_col('Material')
    idx_component = find_col('Component')
    idx_diff      = find_col('Diff Qty')

    # 表头没找到时退回 SAP 固定列位（0基）
    if idx_order     < 0: idx_order     = 0
    if idx_material  < 0: idx_material  = 3
    if idx_component < 0: idx_component = 4
    if idx_diff      < 0: idx_diff      = 5

    # ---------- 2. 逐行严格按 Tab 切分取值 ----------
    result = []
    for l in lines[header_idx + 1:]:
        parts = [p.strip() for p in l.split('\t')]
        need = max(idx_order, idx_material, idx_component, idx_diff)
        if len(parts) <= need:
            continue

        order_val = parts[idx_order]
        if not re.fullmatch(r'\d{9,12}', order_val):
            continue

        material_val  = parts[idx_material]
        component_val = parts[idx_component]
        diff_raw      = parts[idx_diff]

        diff_val = diff_raw.replace(',', '').strip()
        if not re.fullmatch(r'-?\d+(\.\d+)?', diff_val):
            diff_val = "0"

        result.append({
            'Order': order_val,
            'Material': material_val,
            'Component': component_val,
            'Diff Qty': diff_val,
            '_used': False
        })

    return result

def sanitize_qty(v):
    try:
        val = Decimal(str(v).replace(',', '').strip())
        return Decimal('0') if val < 0 else val
    except (InvalidOperation, ValueError):
        return Decimal('0')

def fmt_num(v):
    if isinstance(v, (str, int, float)):
        v = Decimal(str(v))
    rounded = v.quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)
    if rounded == rounded.to_integral():
        return str(int(rounded))
    s = f"{rounded:f}"
    return s.rstrip('0').rstrip('.') if '.' in s else s

DEFAULT_GROUP_DATA = """WBS Element\tOrder\tMaterial\tMaterial description\tSO# QTY\tBasic fin.\tLeadgOrder\tOrder Type\tPlant\tDel. qty\tConf. qty\tWIP QTY\tUnit\tRelease
  PASCAL_VTX\t681000004932\t80-011340-019-RX00\tPCBA, AUDIO AMP, L-PRO2S MKII NHS\t3\t05/05/2026\t681000004932\tPPI1\t6801\t0\t0\t3\tEA\t
  PASCAL_VTX\t681000004931\t50-011340-019-PX00\tINDIRECT ASSY,L-PRO2S MKII BOARD\t3\t05/04/2026\t681000004932\tPPI1\t6801\t0\t0\t3\tEA\t05/06/2026
  PASCAL_VTX\t681000004930\t85-011680-000-PX07\tINDIRECT MATERIAL ASSY,PX3401,VTX\t3\t05/04/2026\t681000004932\tPPI1\t6801\t0\t0\t3\tEA\t
  PASCAL_VTX\t681000004929\t50-011340-019-RX00\tPCB ASSY, AUDIO AMP, L-PRO2S MKII VTX\t3\t05/04/2026\t681000004932\tPPI1\t6801\t0\t0\t3\tEA\t05/06/2026
  PASCAL_VTX\t681000004928\t52-011340-019-PX00\tINDIRECT ASSY,L-PRO2S MKII BOARD\t3\t05/04/2026\t681000004932\tPPI1\t6801\t0\t0\t3\tEA\t05/06/2026"""

class WipApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WIP 入库与匹配管理工具")
        self.root.geometry("1100x800")

        self.order_list = []
        self.wip_rows = []
        self.inventory_list = []
        self.toast_timer = None

        self.setup_ui()
        self.load_saved_data()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=X, padx=10, pady=5)

        group_box = ttk.LabelFrame(top_frame, text="① 工单组数据 (只负责提取工单号)", padding=10)
        group_box.pack(fill=X, pady=5)
        self.txt_group = ttk.Text(group_box, height=4)
        self.txt_group.pack(fill=X, pady=5)

        btn_bar = ttk.Frame(group_box)
        btn_bar.pack(fill=X)
        ttk.Button(btn_bar, text="解析工单组", bootstyle=INFO, command=self.btn_parse_groups).pack(side=LEFT)
        ttk.Button(btn_bar, text="一键全解析并计算 ⚡", bootstyle=SUCCESS, command=self.auto_generate_all).pack(side=LEFT, padx=10)

        wip_box = ttk.LabelFrame(top_frame, text="② 实际WIP数据 (取 Diff Qty 累加)", padding=10)
        wip_box.pack(fill=X, pady=5)
        self.txt_wip = ttk.Text(wip_box, height=4)
        self.txt_wip.pack(fill=X, pady=5)

        bar_wip = ttk.Frame(wip_box)
        bar_wip.pack(fill=X)
        ttk.Button(bar_wip, text="解析实际WIP数据", bootstyle=INFO, command=self.btn_parse_wip).pack(side=LEFT)
        self.lbl_status = ttk.Label(bar_wip, text="等待操作", bootstyle=SECONDARY)
        self.lbl_status.pack(side=LEFT, padx=10)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=YES, padx=10, pady=5)

        self.tab1 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab1, text="📦 功能1：入库清单面板")
        self.build_tab1()

        self.tab2 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab2, text="🔍 功能2：明细匹配 (消耗锁定)")
        self.build_tab2()

    def build_tab1(self):
        f_top = ttk.Frame(self.tab1)
        f_top.pack(fill=X, pady=5)
        ttk.Label(f_top, text="手动输入工单号:", font=("微软雅黑", 10, "bold")).pack(side=LEFT)
        self.ent_order = ttk.Entry(f_top, width=30)
        self.ent_order.pack(side=LEFT, padx=10)
        self.ent_order.bind('<Return>', lambda e: self.calc_order(self.ent_order.get().strip()))

        ttk.Button(f_top, text="生成并添加 (Enter)", bootstyle=SUCCESS, command=lambda: self.calc_order(self.ent_order.get().strip())).pack(side=LEFT)
        ttk.Button(f_top, text="清空列表", bootstyle=DANGER, command=self.reset_inventory_list).pack(side=RIGHT)

        cols = ("order", "total", "status")
        self.tree = ttk.Treeview(self.tab1, columns=cols, show="headings", selectmode="browse", height=15)
        self.tree.heading("order", text="工单号")
        self.tree.heading("total", text="Diff Qty 总数")
        self.tree.heading("status", text="入库状态 (单击切换)")

        self.tree.column("order", width=250, anchor=CENTER)
        self.tree.column("total", width=250, anchor=CENTER)
        self.tree.column("status", width=250, anchor=CENTER)

        self.tree.tag_configure('locked', foreground='#198754', font=("微软雅黑", 10, "bold"))
        self.tree.tag_configure('unlocked', foreground='#212529')

        self.tree.pack(fill=BOTH, expand=YES, pady=10)
        self.tree.bind('<ButtonRelease-1>', self.on_tree_click)

        ttk.Label(self.tab1, text="操作提示：单击列表中的行，可【锁定/解锁】。锁定时自动复制“工单号+换行+数量”。", bootstyle=SECONDARY).pack(anchor=W)

    def auto_generate_all(self):
        self.btn_parse_groups(silent=True)
        self.btn_parse_wip(silent=True)

        if not self.order_list:
            self.show_toast("工单组数据为空，无法解析", "danger")
            return
        if not self.wip_rows:
            self.show_toast("实际WIP数据为空，无法提取 Diff Qty", "danger")
            return

        self.inventory_list = []
        for order in self.order_list:
            self.calc_order(order, silent_copy=True, auto_append=True)

        self.refresh_tree()
        self.show_toast(f"一键提取完成！已生成 {len(self.inventory_list)} 个工单。", "success")

    def calc_order(self, order_num, silent_copy=False, auto_append=False):
        if not order_num:
            return False
        if not self.wip_rows:
            if not auto_append:
                self.show_toast("请确保已加载WIP数据", "danger")
            return False

        total = Decimal('0')
        for r in self.wip_rows:
            if r.get('Order','').strip() == order_num:
                total += sanitize_qty(r.get('Diff Qty', '0'))

        total_str = fmt_num(total)

        exists = False
        for item in self.inventory_list:
            if item['order'] == order_num:
                item['total'] = total_str
                exists = True
                break

        if not exists:
            self.inventory_list.append({
                "order": order_num,
                "total": total_str,
                "locked": False
            })

        if auto_append: return True

        self.refresh_tree()

        if not silent_copy:
            self.copy_to_clipboard(f"{order_num}\n{total_str}", f"{order_num} 已添加并自动复制")
            self.ent_order.delete(0, END)

            for child in self.tree.get_children():
                if self.tree.item(child, 'values')[0] == order_num:
                    self.tree.selection_set(child)
                    self.tree.see(child)
                    break

        return True

    def on_tree_click(self, event):
        selection = self.tree.selection()
        if not selection: return

        item_values = self.tree.item(selection[0], 'values')
        target_order = item_values[0]

        item = next((x for x in self.inventory_list if x['order'] == target_order), None)
        if not item: return

        item['locked'] = not item['locked']

        if item['locked']:
            self.copy_to_clipboard(f"{item['order']}\n{item['total']}", f"已锁定 {item['order']} 并复制")
        else:
            self.show_toast(f"已解锁 {item['order']}", "warning")

        self.refresh_tree()

        for child in self.tree.get_children():
            if self.tree.item(child, 'values')[0] == target_order:
                self.tree.selection_set(child)
                break

    def refresh_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        self.inventory_list.sort(key=lambda x: x['order'])

        for item in self.inventory_list:
            if item['locked']:
                status_text = "✅ 已入库 (锁定)"
                tag = 'locked'
            else:
                status_text = "⏳ 待操作 (未锁定)"
                tag = 'unlocked'

            self.tree.insert("", END, values=(item['order'], item['total'], status_text), tags=(tag,))

    def reset_inventory_list(self):
        if messagebox.askyesno("确认", "确定要清空当前入库列表吗？"):
            self.inventory_list = []
            self.refresh_tree()

    def build_tab2(self):
        f = ttk.Frame(self.tab2)
        f.pack(fill=X, pady=5)
        ttk.Label(f, text="粘贴需要匹配的明细行（Tab 分隔：Order / WBS / Material / Component / Plant / [Diff Qty]）:").pack(anchor=W)

        self.txt_in2 = ttk.Text(self.tab2, height=6)
        self.txt_in2.pack(fill=X, pady=5)
        self.txt_in2.bind('<Control-Return>', lambda e: (self.match_rows(), 'break'))

        bar = ttk.Frame(self.tab2)
        bar.pack(fill=X, pady=5)
        ttk.Button(bar, text="匹配 Diff Qty (Ctrl+Enter)", bootstyle=SUCCESS, command=self.match_rows).pack(side=LEFT)
        ttk.Button(bar, text="重置消耗标记", bootstyle=WARNING, command=self.reset_used_flags).pack(side=LEFT, padx=10)

        self.out2 = ttk.Text(self.tab2, height=8)
        self.out2.pack(fill=BOTH, expand=YES, pady=5)

    def _split_match_line(self, l):
        """把匹配行按 Tab 切成列；返回 (order, material, component)"""
        if '\t' in l:
            parts = [p.strip() for p in l.split('\t')]
        else:
            parts = [p.strip() for p in re.split(r'\s{2,}', l.strip())]

        order = ""
        material = ""
        component = ""

        # 优先按列位取：0=Order, 2=Material, 3=Component（你这份 6 列格式）
        if len(parts) >= 4:
            m = re.fullmatch(r'\d{9,12}', parts[0])
            if m:
                order = parts[0]
                material = parts[2]
                component = parts[3]
                return order, material, component

        # 兜底：整行逐个找 9-12 位数字作为 Order；再找两个“物料型”字符串
        m = re.search(r'\b\d{9,12}\b', l)
        if m:
            order = m.group(0)
        mats = re.findall(r'[0-9A-Z]{2,}-[0-9A-Z]{4,}-[0-9A-Z]{2,}-[0-9A-Z]{3,}', l)
        if len(mats) >= 1:
            material = mats[0]
        if len(mats) >= 2:
            component = mats[1]
        return order, material, component

    def match_rows(self):
        if not self.wip_rows:
            self.show_toast("请先加载 WIP 数据", "danger")
            return
        lines = [l for l in self.txt_in2.get("1.0", "end").splitlines() if l.strip()]
        if not lines:
            return

        out_nums = []
        miss = 0

        for r in self.wip_rows:
            r['_used'] = False

        for l in lines:
            order, material, component = self._split_match_line(l)

            matched = False
            if order:
                for r in self.wip_rows:
                    if r['_used']:
                        continue
                    if r.get('Order','').strip() != order:
                        continue

                    r_mat = r.get('Material', '').strip()
                    r_comp = r.get('Component', '').strip()

                    # 三者都要求完全相等（Material/Component 有值时）
                    if (not material or r_mat == material) and (not component or r_comp == component):
                        out_nums.append(fmt_num(sanitize_qty(r.get('Diff Qty','0'))))
                        r['_used'] = True
                        matched = True
                        break

            if not matched:
                out_nums.append("0")   # 没有就当 0
                miss += 1

        result_str = "\n".join(out_nums)
        self.out2.delete("1.0", "end")
        self.out2.insert("1.0", result_str)

        msg = "匹配完成！已自动复制"
        if miss:
            msg += f" (有 {miss} 行未匹配，已按 0 处理)"
        self.copy_to_clipboard(result_str, msg)

        self.txt_in2.delete("1.0", "end")
        self.notebook.select(self.tab1)

    def reset_used_flags(self):
        for r in self.wip_rows: r['_used'] = False
        self.show_toast("消耗标记已重置", "info")

    def on_close(self):
        data = {
            "group_text": self.txt_group.get("1.0", "end").strip(),
            "wip_text": self.txt_wip.get("1.0", "end").strip(),
            "inventory": self.inventory_list
        }
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
        self.root.destroy()

    def load_saved_data(self):
        if not os.path.exists(DATA_FILE):
            self.txt_group.insert("1.0", DEFAULT_GROUP_DATA)
            return
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("group_text"):
                self.txt_group.insert("1.0", data["group_text"])
            if data.get("wip_text"):
                self.txt_wip.insert("1.0", data["wip_text"])
            if data.get("inventory"):
                self.inventory_list = data["inventory"]
                self.refresh_tree()

            if data.get("group_text"): self.btn_parse_groups(silent=True)
            if data.get("wip_text"): self.btn_parse_wip(silent=True)

        except Exception:
            pass

    def btn_parse_groups(self, silent=False):
        text = self.txt_group.get("1.0", "end")
        self.order_list = parse_order_groups(text)
        if not silent:
            if self.order_list:
                self.show_toast(f"成功解析 {len(self.order_list)} 个独立工单", "success")
            else:
                self.show_toast("解析失败，请检查格式", "danger")

    def btn_parse_wip(self, silent=False):
        text = self.txt_wip.get("1.0", "end")
        self.wip_rows = parse_wip(text)
        if not silent:
            if self.wip_rows:
                self.show_toast(f"成功加载 {len(self.wip_rows)} 条WIP明细", "success")
            else:
                self.show_toast("解析失败，请检查格式", "danger")

    def copy_to_clipboard(self, content, success_msg):
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.show_toast(success_msg, "success")

    def show_toast(self, msg, style="success"):
        if self.toast_timer: self.root.after_cancel(self.toast_timer)
        self.lbl_status.config(text=msg, bootstyle=style)
        self.toast_timer = self.root.after(3000, lambda self=self: self.lbl_status.config(text="就绪", bootstyle="secondary"))

if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetDpiAwareness(1) if hasattr(windll.shcore, "SetDpiAwareness") else windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = ttk.Window(themename="cosmo")
    app = WipApp(root)
    root.mainloop()
