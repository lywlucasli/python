import os
import sys
import subprocess
import threading
import queue
import sqlite3
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
import warnings
import re

warnings.filterwarnings("ignore")

# ==========================================
# 1. 自动安装依赖
# ==========================================
def install_packages():
    required = {'pandas': 'pandas', 'openpyxl': 'openpyxl', 'xlrd': 'xlrd', 'tabulate': 'tabulate'}
    missing = []
    for mod, pkg in required.items():
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    
    if missing:
        print(f"正在安装依赖: {missing}...")
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing, '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple'])
        except:
            messagebox.showerror("错误", "依赖安装失败")
            sys.exit(1)

install_packages()
import pandas as pd

# ==========================================
# 2. 主程序
# ==========================================
class ModernExcelSearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 秒搜工具 Pro (终极版)")
        self.root.geometry("1280x900")
        self.root.configure(bg="#F3F4F6")
        
        # --- 窗口居中 ---
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 1280) // 2
        y = (screen_height - 900) // 2
        self.root.geometry(f"1280x900+{x}+{y}")
        
        # --- 现代化颜色配置 ---
        self.colors = {
            "bg": "#F3F4F6", "card_bg": "#FFFFFF", "primary": "#2563EB", 
            "primary_hover": "#1D4ED8", "text_main": "#1F2937", "text_sub": "#6B7280",
            "border": "#E5E7EB", "success": "#10B981", "danger": "#EF4444",
            "preview_bg": "#F8FAFC", "highlight_row": "#DBEAFE"
        }
        self.fonts = {
            "title": ("Microsoft YaHei UI", 16, "bold"),
            "body": ("Microsoft YaHei UI", 10),
            "body_bold": ("Microsoft YaHei UI", 10, "bold"),
            "small": ("Microsoft YaHei UI", 9),
            "mono": ("Consolas", 10)
        }

        self.db_path = "search_index.db"
        self.init_db()
        
        self.selected_files = []
        self.ui_queue = queue.Queue()
        self.is_processing = False
        self.current_search_keyword = ""
        self.current_selected_file_path = ""
        
        self.click_timer = None
        
        self.setup_styles()
        self.create_widgets()
        self.check_queue()

    def init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('CREATE TABLE IF NOT EXISTS files (path TEXT PRIMARY KEY, mtime REAL)')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT, file_name TEXT, sheet_name TEXT, row_idx INTEGER, content TEXT
                )
            ''')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_content ON data(content)')
            conn.commit()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("Primary.TButton", font=self.fonts["body_bold"], foreground="white", 
                        background=self.colors["primary"], borderwidth=0, padding=(15, 8))
        style.map("Primary.TButton", background=[("active", self.colors["primary_hover"]), ("disabled", "#A0AEC0")])
        
        style.configure("Secondary.TButton", font=self.fonts["body"], foreground=self.colors["text_main"], 
                        background="#E5E7EB", borderwidth=0, padding=(15, 8))
        style.map("Secondary.TButton", background=[("active", "#D1D5DB"), ("disabled", "#F3F4F6")])

        style.configure("Modern.Treeview", font=self.fonts["body"], background="white", 
                        fieldbackground="white", rowheight=34, borderwidth=0)
        style.configure("Modern.Treeview.Heading", font=self.fonts["body_bold"], 
                        background="#F9FAFB", relief="flat", padding=(10, 6))
        style.map("Modern.Treeview", background=[("selected", "#DBEAFE")], foreground=[("selected", self.colors["primary"])])
        
        style.configure("Modern.Vertical.TScrollbar", 
                        gripcount=0, background="#D1D5DB", darkcolor="#F3F4F6", lightcolor="#F3F4F6", 
                        troughcolor="#F3F4F6", bordercolor="#F3F4F6", arrowcolor="#6B7280")
        style.map("Modern.Vertical.TScrollbar", background=[('active', '#9CA3AF')])

        style.configure("Modern.Horizontal.TProgressbar", troughcolor="#E5E7EB", 
                        background=self.colors["primary"], borderwidth=0, thickness=8)

    def create_widgets(self):
        main = tk.Frame(self.root, bg=self.colors["bg"])
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # ================= 顶部区域 =================
        top_frame = tk.Frame(main, bg=self.colors["bg"])
        top_frame.pack(fill="x", pady=(0, 15))

        file_card = tk.Frame(top_frame, bg=self.colors["card_bg"], padx=20, pady=15, highlightthickness=1, highlightbackground=self.colors["border"])
        file_card.pack(side="left", fill="both", expand=True, padx=(0, 10))

        tk.Label(file_card, text="📁 数据源", bg=self.colors["card_bg"], fg=self.colors["text_main"], font=self.fonts["body_bold"]).pack(anchor="w", pady=(0, 10))
        btn_box = tk.Frame(file_card, bg=self.colors["card_bg"])
        btn_box.pack(fill="x")
        
        ttk.Button(btn_box, text="选择文件", style="Secondary.TButton", command=self.select_files).pack(side="left", padx=(0, 10))
        ttk.Button(btn_box, text="扫描文件夹", style="Secondary.TButton", command=self.scan_folder).pack(side="left", padx=(0, 10))
        self.update_index_btn = ttk.Button(btn_box, text="⚡ 更新索引", style="Primary.TButton", command=self.build_index_thread, state="disabled")
        self.update_index_btn.pack(side="left")

        self.file_label = tk.Label(file_card, text="尚未选择文件", bg=self.colors["card_bg"], fg=self.colors["text_sub"], font=self.fonts["small"])
        self.file_label.pack(anchor="w", pady=(10, 0))

        search_card = tk.Frame(top_frame, bg=self.colors["card_bg"], padx=20, pady=15, highlightthickness=1, highlightbackground=self.colors["border"])
        search_card.pack(side="right", fill="both", expand=True, padx=(10, 0))

        tk.Label(search_card, text="🔍 快速搜索", bg=self.colors["card_bg"], fg=self.colors["text_main"], font=self.fonts["body_bold"]).pack(anchor="w", pady=(0, 10))
        s_box = tk.Frame(search_card, bg=self.colors["card_bg"])
        s_box.pack(fill="x")
        
        entry_frame = tk.Frame(s_box, bg="white", highlightthickness=1, highlightbackground=self.colors["border"])
        entry_frame.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=4)
        tk.Label(entry_frame, text="🔎", bg="white", fg=self.colors["text_sub"]).pack(side="left", padx=(10, 0))
        
        self.search_entry = tk.Entry(entry_frame, font=self.fonts["body"], bg="white", bd=0, highlightthickness=0)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=10)
        self.search_entry.bind('<Return>', lambda e: self.perform_search())
        
        self.search_entry.insert(0, "支持 * 和 ? 通配符，也支持多行粘贴")
        self.search_entry.config(fg=self.colors["text_sub"])
        self.search_entry.bind("<FocusIn>", self.on_entry_focus_in)
        self.search_entry.bind("<FocusOut>", self.on_entry_focus_out)
        
        self.search_btn = ttk.Button(s_box, text="搜索", style="Primary.TButton", command=self.perform_search)
        self.search_btn.pack(side="right")

        # ================= 中间：结果表格 =================
        result_card = tk.Frame(main, bg=self.colors["card_bg"], highlightthickness=1, highlightbackground=self.colors["border"])
        result_card.pack(fill="both", expand=True, pady=(0, 15))

        tree_container = tk.Frame(result_card, bg=self.colors["card_bg"], padx=2, pady=2)
        tree_container.pack(fill="both", expand=True)
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        cols = ("file", "sheet", "row", "content")
        self.tree = ttk.Treeview(tree_container, show="headings", columns=cols, style="Modern.Treeview")
        self.tree.heading("file", text="文件名", anchor="w")
        self.tree.heading("sheet", text="工作表", anchor="center")
        self.tree.heading("row", text="行号", anchor="center")
        self.tree.heading("content", text="匹配内容预览", anchor="w")
        
        self.tree.column("file", width=220, stretch=False)
        self.tree.column("sheet", width=100, anchor="center", stretch=False)
        self.tree.column("row", width=80, anchor="center", stretch=False)
        self.tree.column("content", width=650, anchor="w")
        
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview, style="Modern.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        
        self.tree.tag_configure('odd', background='#FFFFFF')
        self.tree.tag_configure('even', background='#F9FAFB')
        self.tree.bind("<<TreeviewSelect>>", self.on_result_select)
        self.tree.bind("<Control-c>", self.copy_main_row)

        # ================= 底部：详细预览 =================
        preview_card = tk.Frame(main, bg=self.colors["card_bg"], highlightthickness=1, highlightbackground=self.colors["border"])
        preview_card.pack(fill="x", pady=(0, 15))

        preview_head = tk.Frame(preview_card, bg=self.colors["card_bg"])
        preview_head.pack(fill="x", padx=15, pady=(10, 5))
        
        tk.Label(preview_head, text="📄 详细数据预览 (支持鼠标拖拽复制)", bg=self.colors["card_bg"], fg=self.colors["text_main"], font=self.fonts["body_bold"]).pack(side="left")
        
        self.open_file_btn = ttk.Button(preview_head, text="📂 打开该文件", style="Secondary.TButton", command=self.open_selected_file, state="disabled")
        self.open_file_btn.pack(side="right")
        
        preview_container = tk.Frame(preview_card, bg=self.colors["preview_bg"], padx=10, pady=10)
        preview_container.pack(fill="both", expand=True)
        preview_container.grid_rowconfigure(0, weight=1)
        preview_container.grid_columnconfigure(0, weight=1)

        self.preview_text = tk.Text(preview_container, height=7, wrap="none", font=self.fonts["mono"], 
                                    bg=self.colors["preview_bg"], bd=0, highlightthickness=0, padx=5, pady=5)
        self.preview_text.grid(row=0, column=0, sticky="nsew")
        
        p_vsb = ttk.Scrollbar(preview_container, orient="vertical", command=self.preview_text.yview, style="Modern.Vertical.TScrollbar")
        p_vsb.grid(row=0, column=1, sticky="ns")
        self.preview_text.configure(yscrollcommand=p_vsb.set)
        
        self.preview_text.tag_configure("highlight", background="#FEF08A", foreground="#DC2626")
        self.preview_text.tag_configure("current_row", background=self.colors["highlight_row"], font=(self.fonts["mono"][0], 10, "bold"))
        
        self.text_menu = tk.Menu(self.root, tearoff=0, font=self.fonts["small"])
        self.text_menu.add_command(label="📋 复制选中文本", command=lambda: self.preview_text.event_generate("<<Copy>>"))
        self.preview_text.bind("<Button-3>", self.show_text_menu)

        # ================= 底部状态区 =================
        status_frame = tk.Frame(main, bg=self.colors["bg"])
        status_frame.pack(fill="x", side="bottom")
        
        info_box = tk.Frame(status_frame, bg=self.colors["bg"])
        info_box.pack(fill="x", pady=(0, 5))
        
        self.status_label = tk.Label(info_box, text="就绪", bg=self.colors["bg"], fg=self.colors["text_sub"], font=self.fonts["body"], anchor="w")
        self.status_label.pack(side="left")
        
        self.progress_bar = ttk.Progressbar(status_frame, style="Modern.Horizontal.TProgressbar", mode='determinate')
        self.progress_bar.pack(fill="x", pady=(0, 5))

        self.tree.bind("<MouseWheel>", lambda e: self.tree.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.preview_text.bind("<MouseWheel>", lambda e: self.preview_text.yview_scroll(int(-1*(e.delta/120)), "units"))

    def show_text_menu(self, event):
        try:
            self.text_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.text_menu.grab_release()

    def copy_main_row(self, event=None):
        selected = self.tree.selection()
        if not selected: return
        item = self.tree.item(selected[0])
        values = item['values']
        if not values: return
        
        text = "\t".join([str(v) for v in values])
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        
        self.status_label.config(text="✅ 已复制主表格行内容到剪贴板", fg=self.colors["success"])
        self.root.after(2000, lambda: self.status_label.config(text="就绪", fg=self.colors["text_sub"]))

    def on_entry_focus_in(self, event):
        if self.search_entry.get() == "支持 * 和 ? 通配符，也支持多行粘贴":
            self.search_entry.delete(0, tk.END)
            self.search_entry.config(fg=self.colors["text_main"])

    def on_entry_focus_out(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, "支持 * 和 ? 通配符，也支持多行粘贴")
            self.search_entry.config(fg=self.colors["text_sub"])

    def toggle_buttons(self, state):
        self.update_index_btn.config(state=state)
        self.search_btn.config(state=state)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel/CSV", "*.xlsx;*.xls;*.csv")])
        if files:
            self.selected_files = list(files)
            self.file_label.config(text=f"已选择 {len(files)} 个文件", fg=self.colors["success"])
            self.update_index_btn.config(state="normal")

    def scan_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            found = []
            for root, _, fs in os.walk(folder):
                for f in fs:
                    if f.lower().endswith(('.xlsx', '.xls', '.csv')) and not f.startswith('~$'):
                        found.append(os.path.join(root, f))
            
            self.selected_files = found
            if found:
                self.file_label.config(text=f"已扫描 {len(found)} 个文件", fg=self.colors["success"])
                self.update_index_btn.config(state="normal")
            else:
                self.file_label.config(text="未找到支持的表格文件", fg=self.colors["danger"])

    def build_index_thread(self):
        if not self.selected_files or self.is_processing: return
        
        self.is_processing = True
        self.toggle_buttons("disabled")
        for item in self.tree.get_children(): self.tree.delete(item)
        self.preview_text.delete('1.0', tk.END)
        self.open_file_btn.config(state="disabled")
            
        threading.Thread(target=self._build_index_worker, args=(self.selected_files,), daemon=True).start()

    def _build_index_worker(self, files):
        total = len(files)
        updated_count = 0
        skipped_count = 0
        error_list = []
        current_valid_paths = set(files)
        
        existing_files = {}
        try:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.cursor()
                cur.execute("SELECT path, mtime FROM files")
                existing_files = {row[0]: row[1] for row in cur.fetchall()}
        except: pass

        for idx, file_path in enumerate(files, 1):
            try:
                mtime = os.path.getmtime(file_path)
                
                if file_path in existing_files and existing_files[file_path] == mtime:
                    skipped_count += 1
                    self.ui_queue.put(('progress', (idx, total, f"跳过: {os.path.basename(file_path)}")))
                    continue
                
                self.ui_queue.put(('progress', (idx, total, f"读取: {os.path.basename(file_path)}")))
                
                ext = Path(file_path).suffix.lower()
                data = None
                if ext == '.csv':
                    for enc in ['utf-8', 'gbk', 'gb18030', 'latin1']:
                        try:
                            data = {'CSV': pd.read_csv(file_path, encoding=enc, on_bad_lines='skip', dtype=str)}
                            break
                        except: continue
                elif ext in ['.xls', '.xlsx']:
                    try:
                        xls = pd.ExcelFile(file_path)
                        data = {sheet: pd.read_excel(file_path, sheet_name=sheet, dtype=str) for sheet in xls.sheet_names}
                    except Exception:
                        if ext == '.xls':
                            try: data = {'TSV': pd.read_csv(file_path, sep='\t', encoding='gbk', dtype=str)}
                            except: pass
                
                if data:
                    with sqlite3.connect(self.db_path) as conn:
                        cur = conn.cursor()
                        cur.execute("DELETE FROM data WHERE file_path = ?", (file_path,))
                        cur.execute("DELETE FROM files WHERE path = ?", (file_path,))
                        
                        batch_data = []
                        for sheet_name, df in data.items():
                            for row_idx, row in df.iterrows():
                                row_content = " | ".join([str(x) for x in row.values if str(x) != 'nan'])
                                batch_data.append((file_path, os.path.basename(file_path), sheet_name, row_idx+1, row_content))
                                
                                if len(batch_data) >= 1000:
                                    cur.executemany("INSERT INTO data (file_path, file_name, sheet_name, row_idx, content) VALUES (?, ?, ?, ?, ?)", batch_data)
                                    batch_data = []
                        
                        if batch_data:
                            cur.executemany("INSERT INTO data (file_path, file_name, sheet_name, row_idx, content) VALUES (?, ?, ?, ?, ?)", batch_data)
                            
                        cur.execute("INSERT INTO files (path, mtime) VALUES (?, ?)", (file_path, mtime))
                        conn.commit()
                    updated_count += 1
                else:
                    error_list.append(os.path.basename(file_path))
            except Exception:
                error_list.append(os.path.basename(file_path))
            
            self.root.update_idletasks()

        try:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.cursor()
                cur.execute("SELECT path FROM files")
                db_files = set(row[0] for row in cur.fetchall())
                deleted_paths = db_files - current_valid_paths
                if deleted_paths:
                    for dp in deleted_paths:
                        cur.execute("DELETE FROM data WHERE file_path = ?", (dp,))
                        cur.execute("DELETE FROM files WHERE path = ?", (dp,))
                    conn.commit()
        except: pass

        self.ui_queue.put(('index_done', (updated_count, skipped_count, error_list)))

    # =================== 核心修复：多通配符智能检索 ===================
    def perform_search(self):
        keyword = self.search_entry.get().strip()
        if not keyword or keyword == "支持 * 和 ? 通配符，也支持多行粘贴":
            messagebox.showwarning("提示", "请输入搜索关键字")
            return
        if self.is_processing: return

        self.current_search_keyword = keyword
        self.is_processing = True
        self.toggle_buttons("disabled")
        for item in self.tree.get_children(): self.tree.delete(item)
        self.preview_text.delete('1.0', tk.END)
        self.open_file_btn.config(state="disabled")
        self.current_selected_file_path = ""
        
        self.status_label.config(text="正在从索引库检索...", fg=self.colors["primary"])
        self.progress_bar["mode"] = "indeterminate"
        self.progress_bar.start(10)

        # 1. 处理多行粘贴输入 (支持你一次性粘贴多个料号)
        raw_keywords = [k.strip() for k in keyword.split('\n') if k.strip()]
        
        # 2. 将用户的 * 语法转换为 SQL LIKE 语法
        sql_patterns = []
        for kw in raw_keywords:
            # 连续多个星号(如 ******)会被合并为一个 %
            cleaned_kw = re.sub(r'\*+', '*', kw)
            
            if '*' not in cleaned_kw and '?' not in cleaned_kw:
                sql_pat = f"%{cleaned_kw}%"
            else:
                sql_pat = cleaned_kw.replace('*', '%').replace('?', '_')
            
            # 如果用户用了星号但没全包含，自动补两头，确保搜索范围广
            if not sql_pat.startswith('%'): sql_pat = "%" + sql_pat
            if not sql_pat.endswith('%'): sql_pat = sql_pat + "%"
            
            sql_patterns.append(sql_pat)

        # 3. 拼装多条件 SQL
        # 如果输入了多个料号，使用 OR 连接
        conditions = " OR ".join(["content LIKE ?"] * len(sql_patterns))
        final_sql = f"SELECT file_name, sheet_name, row_idx, content, file_path FROM data WHERE {conditions} LIMIT 3000"

        threading.Thread(target=self._search_db_worker, args=(final_sql, sql_patterns), daemon=True).start()

    def _search_db_worker(self, final_sql, params):
        try:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.cursor()
                cur.execute(final_sql, tuple(params))
                rows = cur.fetchall()
                
                batch_size = 50
                for i in range(0, len(rows), batch_size):
                    self.ui_queue.put(('search_batch', rows[i:i+batch_size]))
            
            self.ui_queue.put(('search_done', len(rows)))
        except Exception as e:
            self.ui_queue.put(('error', str(e)))

    def on_result_select(self, event):
        if self.click_timer:
            self.root.after_cancel(self.click_timer)
        self.click_timer = self.root.after(100, self._process_preview)

    def _process_preview(self):
        selected = self.tree.selection()
        if not selected: return
        
        item = self.tree.item(selected[0])
        values = item['values']
        if not values: return
        
        file_name, sheet_name, row_str = values[0], values[1], values[2]
        row_idx = int(row_str)
        
        self.preview_text.delete('1.0', tk.END)

        try:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT row_idx, content FROM data 
                    WHERE file_name = ? AND sheet_name = ? AND row_idx >= ? AND row_idx <= ?
                    ORDER BY row_idx ASC
                """, (file_name, sheet_name, max(1, row_idx - 2), row_idx + 2))
                context_rows = cur.fetchall()
                
                cur.execute("SELECT file_path FROM data WHERE file_name = ? LIMIT 1", (file_name,))
                path_res = cur.fetchone()
                if path_res:
                    self.current_selected_file_path = path_res[0]
                    self.open_file_btn.config(state="normal")
                else:
                    self.current_selected_file_path = ""
                    self.open_file_btn.config(state="disabled")
                
                if context_rows:
                    for r_idx, r_content in context_rows:
                        is_current = (r_idx == row_idx)
                        prefix = ">> " if is_current else "   "
                        line_text = f"{prefix}第 {r_idx} 行: {r_content}\n"
                        
                        start_index = self.preview_text.index(tk.END)
                        self.preview_text.insert(tk.END, line_text)
                        end_index = self.preview_text.index(tk.END)
                        
                        if is_current:
                            self.preview_text.tag_add("current_row", start_index, end_index)
                    
                    self.highlight_preview_keyword()
                    
        except Exception as e:
            pass

    def highlight_preview_keyword(self):
        kw = self.current_search_keyword
        if not kw: return
        
        # 针对多行粘贴的情况，把每一行的去星号结果都高亮
        clean_kws = [k.strip().replace('*', '').replace('?', '') for k in kw.split('\n') if k.strip()]
        
        for clean_kw in clean_kws:
            if not clean_kw: continue
            idx = '1.0'
            while True:
                idx = self.preview_text.search(clean_kw, idx, nocase=True, stopindex=tk.END)
                if not idx: break
                last_idx = f"{idx}+{len(clean_kw)}c"
                self.preview_text.tag_add("highlight", idx, last_idx)
                idx = last_idx

    def open_selected_file(self):
        if not self.current_selected_file_path: return
        
        if not os.path.exists(self.current_selected_file_path):
            messagebox.showerror("错误", "文件不存在，可能已被移动或删除。")
            return
            
        try:
            if sys.platform == 'win32':
                os.startfile(self.current_selected_file_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', self.current_selected_file_path])
            else:
                subprocess.call(['xdg-open', self.current_selected_file_path])
        except Exception as e:
            messagebox.showerror("打开失败", f"无法调用系统程序打开文件: {e}")

    def check_queue(self):
        try:
            for _ in range(50):
                msg_type, data = self.ui_queue.get_nowait()
                
                if msg_type == 'progress':
                    current, total, text = data
                    self.progress_bar["mode"] = "determinate"
                    self.progress_bar["maximum"] = total
                    self.progress_bar["value"] = current
                    self.status_label.config(text=f"[{current}/{total}] {text}", fg=self.colors["text_sub"])
                
                elif msg_type == 'index_done':
                    updated, skipped, errors = data
                    self.is_processing = False
                    self.toggle_buttons("normal")
                    self.progress_bar["value"] = 0
                    
                    msg = f"索引更新完成 (更新/新增:{updated}, 跳过无变化:{skipped})"
                    self.status_label.config(text=msg, fg=self.colors["success"])
                    if errors:
                        messagebox.showwarning("注意", f"完成，但 {len(errors)} 个文件读取失败被跳过。")

                elif msg_type == 'search_batch':
                    for row in data:
                        file_name, sheet, row_idx, content, file_path = row
                        if len(content) > 200: content = content[:200] + "..."
                        tag = 'even' if len(self.tree.get_children()) % 2 == 0 else 'odd'
                        self.tree.insert("", "end", values=(file_name, sheet, row_idx, content), tags=(tag,))
                
                elif msg_type == 'search_done':
                    count = data
                    self.is_processing = False
                    self.progress_bar.stop()
                    self.progress_bar["mode"] = "determinate"
                    self.progress_bar["value"] = 0
                    self.status_label.config(text=f"检索完成！共找到 {count} 条匹配记录", fg=self.colors["success"])
                    self.toggle_buttons("normal")
                    
                elif msg_type == 'error':
                    messagebox.showerror("错误", data)
                    self.is_processing = False
                    self.progress_bar.stop()
                    self.toggle_buttons("normal")

        except queue.Empty:
            pass
        
        self.root.after(50, self.check_queue)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ModernExcelSearchTool(root)
        root.mainloop()
    except Exception as e:
        print(f"启动失败: {e}")