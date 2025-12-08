import os
import sys
import queue
import threading
import tempfile
import shutil
import contextlib
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
import keyring

# 作成したモジュールをインポート
import config
import notion_api
import email_service
import pdf_generator
import settings_gui

class QueueIO:
    def __init__(self, q):
        self.q = q
    def write(self, text):
        self.q.put(("log", text))
    def flush(self):
        pass

class TopPane(ttk.Frame):
    """上部のアクションとフィルター領域のUI"""
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app

        # --- フィルターとアカウントを並べて表示するコンテナ ---
        sub_pane = ttk.Frame(self)
        sub_pane.pack(fill=tk.X, expand=True)
        sub_pane.columnconfigure(0, weight=3)
        sub_pane.columnconfigure(1, weight=0)

        # --- 1b. 部署名フィルター ---
        department_container = ttk.Frame(sub_pane, style="Highlight.TFrame", relief="solid", borderwidth=1, padding=(0, 5))
        department_container.grid(row=0, column=0, sticky="ew", pady=(5, 0))

        title_label = ttk.Label(department_container, text="部署名フィルター", font=("Yu Gothic UI", 12, "bold"), foreground=self.app.PRIMARY_COLOR, background="#EAF2F8")
        title_label.pack(anchor="w", padx=10, pady=(0, 2))

        departments = config.load_departments()
        self.app.department_vars = {name: tk.BooleanVar() for name in departments}
        
        checkbox_container = ttk.Frame(department_container, style="Highlight.TFrame")
        checkbox_container.pack(fill=tk.X, padx=5)

        for i, name in enumerate(departments):
            cb = ttk.Checkbutton(checkbox_container, text=name, variable=self.app.department_vars[name], style="Highlight.TCheckbutton", command=self.app.on_department_selection_change)
            cb.grid(row=0, column=i, padx=(5, 15), pady=(0,5), sticky='w')

        # --- 右側のアカウント選択 ---
        account_frame = ttk.Frame(sub_pane, style="Highlight.TFrame", relief="solid", borderwidth=1, padding=(0, 5))
        account_frame.grid(row=0, column=1, sticky="n", padx=(20, 0), pady=(5, 0))
        ttk.Label(account_frame, text="送信者アカウント", font=("Yu Gothic UI", 12, "bold"),
                  foreground=self.app.PRIMARY_COLOR, background="#E0E9FF").pack(anchor="w", padx=10, pady=(0, 2))

        account_contents = ttk.Frame(account_frame, style="Highlight.TFrame")
        account_contents.pack(fill=tk.X, padx=10, pady=(2, 10))

        account_display_names = sorted(list(self.app.display_name_to_key_map.keys()))
        self.account_selector = ttk.Combobox(account_contents, textvariable=self.app.selected_account_display_name, values=account_display_names, state="readonly", width=25, font=("Yu Gothic UI", 10))
        self.account_selector.pack(side=tk.LEFT, padx=(0, 10), pady=4)

        sender_label_frame = ttk.Frame(account_contents, style="Highlight.TFrame")
        sender_label_frame.pack(side=tk.LEFT, fill=tk.X, padx=(0, 10), pady=4)
        ttk.Label(sender_label_frame, text="送信元:", font=("Yu Gothic UI", 10)).pack(side=tk.LEFT)
        ttk.Label(sender_label_frame, textvariable=self.app.sender_email_var, font=("Yu Gothic UI", 10, "bold")).pack(side=tk.LEFT, padx=5)

        # --- 1a. データ取得ボタン ---
        action_frame = ttk.Frame(self)
        action_frame.pack(fill=tk.X, pady=(0, 10), anchor='w')

        self.get_data_button = ttk.Button(action_frame, text="Notionからデータを取得", command=self.app.start_data_retrieval, style="Primary.TButton")
        self.get_data_button.pack(side=tk.LEFT, ipady=5)

        self.spinner_label = ttk.Label(action_frame, textvariable=self.app.spinner_var, font=("Yu Gothic UI", 16), foreground=self.app.ACCENT_COLOR)

    def toggle_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        self.get_data_button.config(state=state)

    def start_spinner(self):
        self.spinner_label.pack(side=tk.LEFT, padx=(10, 0), anchor='center')

    def stop_spinner(self):
        self.spinner_label.pack_forget()

class MiddlePane(ttk.PanedWindow):
    """中央のデータ表示領域 (仕入先リストと注文データ) のUI"""
    def __init__(self, master, app):
        super().__init__(master, orient=tk.HORIZONTAL)
        self.app = app

        supplier_pane = ttk.LabelFrame(self, text="仕入先を選択")
        self.add(supplier_pane, weight=1)
        
        table_pane = ttk.LabelFrame(self, text="発注対象データ")
        self.add(table_pane, weight=3)

        # --- 仕入先リスト ---
        self.supplier_listbox = ttk.Treeview(supplier_pane, columns=("supplier_name",), show="headings", selectmode="browse")
        self.supplier_listbox.heading("supplier_name", text="仕入先")
        self.supplier_listbox.column("supplier_name", width=200, anchor=tk.W)
        self.supplier_listbox.tag_configure('sent', foreground='gray', background='#F0F0F0')
        self.supplier_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(1,0), pady=1)
        self.supplier_listbox.bind("<ButtonRelease-1>", self.app.on_supplier_select)
        
        supplier_vsb = ttk.Scrollbar(supplier_pane, orient="vertical", command=self.supplier_listbox.yview, style="Vertical.TScrollbar")
        self.supplier_listbox.config(yscrollcommand=supplier_vsb.set)
        supplier_vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=1)

        # --- 注文データテーブル ---
        columns = ("maker", "part_num", "qty")
        self.tree = ttk.Treeview(table_pane, columns=columns, show="headings")
        self.tree.heading("maker", text="メーカー"); self.tree.heading("part_num", text="品番"); self.tree.heading("qty", text="数量")
        self.tree.column("maker", width=150); self.tree.column("part_num", width=250); self.tree.column("qty", width=60, anchor=tk.E)
        vsb = ttk.Scrollbar(table_pane, orient="vertical", command=self.tree.yview, style="Vertical.TScrollbar")
        hsb = ttk.Scrollbar(table_pane, orient="horizontal", command=self.tree.xview, style="Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=1); hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=1); self.tree.pack(fill=tk.BOTH, expand=True, padx=1, pady=(1,0))

    def update_supplier_list(self, data):
        self.supplier_listbox.delete(*self.supplier_listbox.get_children())
        suppliers = sorted(list(set(i["supplier_name"] for i in data)))
        for s in suppliers: self.supplier_listbox.insert('', tk.END, values=(s,))

    def update_table_for_supplier(self, supplier_name):
        self.tree.delete(*self.tree.get_children())
        items_to_display = self.app.orders_by_supplier.get(supplier_name, [])
        for item in items_to_display:
            self.tree.insert("", tk.END, values=(item.get("maker_name", ""), item.get("db_part_number", ""), item.get("quantity", 0)))

    def mark_supplier_as_sent(self, supplier):
        for iid in self.supplier_listbox.get_children():
            if self.supplier_listbox.item(iid, 'values')[0] == supplier:
                self.supplier_listbox.item(iid, tags=('sent',))
                self.supplier_listbox.selection_remove(iid)
                break

    def clear_displays(self):
        self.tree.delete(*self.tree.get_children())
        self.supplier_listbox.delete(*self.supplier_listbox.get_children())

class BottomPane(ttk.PanedWindow):
    """下部のプレビューとログ領域のUI"""
    def __init__(self, master, app):
        super().__init__(master, orient=tk.HORIZONTAL)
        self.app = app

        # --- プレビューエリア ---
        preview_frame = ttk.LabelFrame(self, text="内容を確認して送信")
        self.add(preview_frame, weight=1)

        self.to_var, self.cc_var, self.contact_var, self.pdf_var = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()
        preview_grid = ttk.Frame(preview_frame, padding=10, style="Light.TFrame")
        preview_grid.pack(fill="both", expand=True)
        preview_labels = {"宛先(To):": self.to_var, "宛先(Cc):": self.cc_var, "担当者:": self.contact_var}
        for i, (text, var) in enumerate(preview_labels.items()):
            ttk.Label(preview_grid, text=text, font=("Yu Gothic UI", 9), style="Light.TLabel").grid(row=i, column=0, sticky=tk.W, padx=5, pady=3)
            ttk.Label(preview_grid, textvariable=var, font=("Yu Gothic UI", 9, "bold"), style="Light.TLabel").grid(row=i, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(preview_grid, text="添付PDF:", font=("Yu Gothic UI", 9), style="Light.TLabel").grid(row=3, column=0, sticky=tk.W, padx=5, pady=3)
        self.pdf_label = ttk.Label(preview_grid, textvariable=self.pdf_var, cursor="hand2", style="PdfLink.TLabel")
        self.pdf_label.grid(row=3, column=1, sticky=tk.W, padx=5)
        self.pdf_label.bind("<Button-1>", self.app.open_current_pdf)
        preview_grid.columnconfigure(1, weight=1)

        # --- ログエリア ---
        log_frame = ttk.LabelFrame(self, text="通知メッセージ")
        self.add(log_frame, weight=1)

        self.log_display = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD, bg=self.app.LIGHT_BG, fg=self.app.TEXT_COLOR, font=("Consolas", 11), relief="flat", borderwidth=0, highlightthickness=0)
        self.log_display.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self.log_display.configure(state='disabled')
        self.log_display.tag_configure("emphasis", foreground=self.app.EMPHASIS_COLOR, font=("Yu Gothic UI", 12, "bold"))
        self.log_display.tag_configure("error", foreground="red", font=("Consolas", 11, "bold"))

    def update_preview(self, info, pdf_path):
        self.app.current_pdf_path = pdf_path
        self.to_var.set(info.get("email", ""))
        self.cc_var.set(info.get("email_cc", ""))
        self.contact_var.set(info.get("sales_contact", ""))
        self.pdf_var.set(os.path.basename(pdf_path) if pdf_path else "作成失敗")

    def clear_preview(self):
        self.to_var.set(""); self.cc_var.set(""); self.contact_var.set(""); self.pdf_var.set("")
        self.app.current_pdf_path = None

    def log(self, message, tag=None):
        self.log_display.config(state="normal")
        if tag:
            self.log_display.insert(tk.END, message + "\n", tag)
        else:
            self.log_display.insert(tk.END, message + "\n")
        self.log_display.see(tk.END)
        self.log_display.config(state="disabled")

    def clear_log(self):
        self.log_display.config(state="normal")
        self.log_display.delete(1.0, tk.END)
        self.log_display.config(state="disabled")

class Application(ttk.Frame):
    """アプリケーションのメインコントローラー"""
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.configure_styles()
        self.create_menu()
        self.q = queue.Queue()
        self.queue_io = QueueIO(self.q)
        
        # --- 一時フォルダと事前生成PDFの管理 ---
        self.temp_dir = tempfile.TemporaryDirectory()
        self.pregenerated_pdfs = {}

        # --- 状態管理 ---
        self.processing = False
        self.order_data = []
        self.orders_by_supplier = {}
        self.current_pdf_path = None
        self.sent_suppliers = set()
        self.selected_departments = []

        # --- Tkinter変数 ---
        self.department_vars = {}
        self.selected_account_display_name = tk.StringVar()
        self.sender_email_var = tk.StringVar()
        self.spinner_var = tk.StringVar()

        # --- スピナー関連 ---
        self.spinner_running = False
        self.spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.spinner_index = 0
        self._after_id = None

        # --- 設定とマッピング ---
        self.accounts = config.load_email_accounts()
        self.department_defaults = config.load_department_defaults()
        self.display_name_to_key_map = {v.get('display_name', k): k for k, v in self.accounts.items()}
        
        self.selected_account_display_name.trace_add("write", self.update_sender_label)

        self.configure_styles()
        self.create_widgets()
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.initialize_app_state()

    def configure_styles(self):
        style = ttk.Style(self.master)
        style.theme_use("clam")

        self.BG_COLOR = "#E8F0FF"
        self.TEXT_COLOR = "#1F2A44"
        self.PRIMARY_COLOR = "#1E40AF"
        self.ACCENT_COLOR = "#2563EB"
        self.LIGHT_BG = "#FFFFFF"
        self.HEADER_FG = "#FFFFFF"
        self.SELECT_BG = "#D1E2FF"
        self.SELECT_FG = "#0F172A"
        self.EMPHASIS_COLOR = self.ACCENT_COLOR
        self.BORDER_COLOR = "#C7D5EF"

        style.configure(".", background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=("Yu Gothic UI", 10))
        style.configure("TFrame", background=self.BG_COLOR)
        style.configure("TLabel", background=self.BG_COLOR, foreground=self.TEXT_COLOR)
        style.configure("Light.TLabel", background=self.LIGHT_BG)
        style.configure("PdfLink.TLabel", background=self.LIGHT_BG, foreground=self.ACCENT_COLOR, font=("Yu Gothic UI", 9, "underline"))
        style.configure("TLabelFrame", background=self.BG_COLOR, bordercolor=self.BORDER_COLOR, relief="solid", borderwidth=1)
        style.configure("TLabelFrame.Label", background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=("Yu Gothic UI", 12, "bold"))
        style.configure("TButton", font=("Yu Gothic UI", 10, "bold"), borderwidth=0, padding=(15, 10))
        style.configure("Primary.TButton", background=self.PRIMARY_COLOR, foreground=self.HEADER_FG)
        style.map("Primary.TButton", background=[("active", "#1B3AA8"), ("disabled", "#A9CCE3")])
        style.configure("Secondary.TButton", background="#D9E8FF", foreground=self.TEXT_COLOR)
        style.map("Secondary.TButton", background=[("active", "#BAC8EF"), ("disabled", "#E0E0E0")])
        style.configure("Treeview", background=self.LIGHT_BG, fieldbackground=self.LIGHT_BG, foreground=self.TEXT_COLOR, rowheight=28, font=("Yu Gothic UI", 10))
        style.map("Treeview", background=[("selected", self.SELECT_BG)], foreground=[("selected", self.SELECT_FG)])
        style.configure("Treeview.Heading", background="#6C757D", foreground=self.HEADER_FG, font=("Yu Gothic UI", 10, "bold"), padding=8)
        style.map("Treeview.Heading", background=[("active", "#495057")])
        style.configure("TCombobox", font=("Yu Gothic UI", 10), fieldbackground=self.LIGHT_BG, padding=5)
        style.configure("Vertical.TScrollbar", background=self.PRIMARY_COLOR, troughcolor=self.BG_COLOR, arrowcolor=self.HEADER_FG)
        style.configure("Horizontal.TScrollbar", background=self.PRIMARY_COLOR, troughcolor=self.BG_COLOR, arrowcolor=self.HEADER_FG)
        style.configure("TPanedWindow", background=self.BORDER_COLOR)
        style.configure("Highlight.TFrame", background="#E0E9FF", bordercolor=self.PRIMARY_COLOR)
        style.configure("Light.TFrame", background=self.LIGHT_BG)
        style.configure("Highlight.TCheckbutton", background="#E0E9FF", font=("Yu Gothic UI", 10))
        style.map("Highlight.TCheckbutton", background=[('active', '#E0E9FF')], indicatorbackground=[('active', '#E0E9FF')], foreground=[('selected', '#1E40AF')], font=[('selected', ("Yu Gothic UI", 10, "bold"))])

    def create_menu(self):
        """アプリケーションメニューを生成する"""
        menubar = tk.Menu(self.master)
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="設定を開く", command=self.open_settings_window)
        settings_menu.add_command(label="設定リロード", command=self.reload_ui_after_settings_change)
        menubar.add_cascade(label="　⚙ 設定", menu=settings_menu)
        self.master.config(menu=menubar)

    def create_widgets(self):
        self.pack(fill=tk.BOTH, expand=True)
        main_container = ttk.Frame(self, padding=15)
        main_container.pack(fill=tk.BOTH, expand=True)
        main_container.rowconfigure(1, weight=1)
        main_container.rowconfigure(2, weight=1)
        main_container.columnconfigure(0, weight=1)

        self.top_pane = TopPane(main_container, self)
        self.top_pane.grid(row=0, column=0, sticky="ew", pady=(0, 15))

        self.middle_pane = MiddlePane(main_container, self)
        self.middle_pane.grid(row=1, column=0, sticky="nsew", pady=0)

        self.bottom_pane = BottomPane(main_container, self)
        self.bottom_pane.grid(row=2, column=0, sticky="nsew", pady=(15, 0))

        send_button_area = ttk.Frame(main_container)
        send_button_area.grid(row=3, column=0, sticky="ew", pady=(15, 0))
        send_button_container = ttk.Frame(send_button_area)
        send_button_container.pack(expand=True)
        self.send_mail_button = ttk.Button(send_button_container, text="メール送信", command=self.send_single_mail, state="disabled", style="Primary.TButton")
        self.send_mail_button.pack(ipadx=40, ipady=15)

    def initialize_app_state(self):
        if self.accounts:
            self.set_default_sender_account()
        self.update_sender_label()
        self.log("----------------------------------------")
        self.log("STEP 1: アプリケーション開始")
        self.log("----------------------------------------")
        self.log("1. 部署名を選択してください。", "emphasis")
        self.log("2. 送信者アカウントを確認してください。", "emphasis")
        self.log("3. 「Notionからデータを取得」ボタンをクリックしてください。", "emphasis")

    # --- スピナー管理 ---
    def start_spinner(self):
        if self.spinner_running: return
        self.spinner_running = True
        self.top_pane.start_spinner()
        self.animate_spinner()

    def stop_spinner(self):
        if not self.spinner_running: return
        if self._after_id: self.master.after_cancel(self._after_id)
        self._after_id = None
        self.spinner_var.set("")
        self.top_pane.stop_spinner()
        self.spinner_running = False

    def animate_spinner(self):
        if not self.spinner_running: return
        self.spinner_var.set(f"Loading {self.spinner_frames[self.spinner_index]}")
        self.spinner_index = (self.spinner_index + 1) % len(self.spinner_frames)
        self._after_id = self.master.after(config.AppConstants.SPINNER_ANIMATION_DELAY, self.animate_spinner)

    # --- UIイベントハンドラ ---
    def on_department_selection_change(self):
        self.set_default_sender_account()

    def update_sender_label(self, *args):
        selected_display_name = self.selected_account_display_name.get()
        account_key = self.display_name_to_key_map.get(selected_display_name)
        self.sender_email_var.set(self.accounts[account_key]["sender"] if account_key and account_key in self.accounts else "")

    def start_data_retrieval(self):
        if self.processing: return
        self.selected_departments = [name for name, var in self.department_vars.items() if var.get()]
        self.reset_temp_storage()
        self.processing = True
        self.toggle_buttons(False)
        self.clear_displays()
        self.start_spinner()
        threading.Thread(target=self.run_thread, args=(self.get_data_task,)).start()
        self.master.after(config.AppConstants.QUEUE_CHECK_INTERVAL, self.check_queue)

    def on_supplier_select(self, event):
        selected_iid = self.middle_pane.supplier_listbox.identify_row(event.y)
        if not selected_iid or self.processing: return
        self.middle_pane.supplier_listbox.selection_set(selected_iid)
        selected_supplier = self.middle_pane.supplier_listbox.item(selected_iid, 'values')[0]
        
        # 発注対象データテーブルを更新
        self.middle_pane.update_table_for_supplier(selected_supplier)

        if selected_supplier in self.sent_suppliers:
            self.clear_preview()
            return

        # 事前生成されたPDFのパスを取得
        pdf_path = self.pregenerated_pdfs.get(selected_supplier)
        items = self.orders_by_supplier.get(selected_supplier, [])

        if pdf_path and items:
            # プレビューを即時更新
            self.update_preview_ui((items[0], pdf_path))
            self.log(f"\n「{selected_supplier}」のプレビューを表示しました。", "emphasis")
        elif not items:
            self.log(f"エラー: 「{selected_supplier}」の注文アイテムが見つかりません。", "error")
            self.clear_preview()
        else:
            self.log("PDFはまだ準備中です。少し待ってからもう一度お試しください。", "emphasis")
            self.clear_preview()

    def send_single_mail(self):
        if self.processing or not self.current_pdf_path: return
        
        selected_iids = self.middle_pane.supplier_listbox.selection()
        if not selected_iids: return
        selected_supplier = self.middle_pane.supplier_listbox.item(selected_iids[0], 'values')[0]
        items = self.orders_by_supplier.get(selected_supplier, [])
        if not items:
            messagebox.showerror("データなし", f"「{selected_supplier}」の注文データが見つかりません。")
            return
        supplier_departments = []
        for item in items:
            for dept in (item.get("departments") or []):
                dept = dept.strip()
                if dept and dept not in supplier_departments:
                    supplier_departments.append(dept)
        department_for_pdf = next((dept for dept in self.selected_departments if dept in supplier_departments), None)
        if not department_for_pdf and supplier_departments:
            department_for_pdf = supplier_departments[0]

        # PDFを本保存先にコピー
        try:
            # --- 保存先パスの計算ロジックを追加 ---
            base_dest_dir = config.PDF_SAVE_DIR
            final_dest_dir = base_dest_dir
            
            # 部署フォルダを補完
            if department_for_pdf:
                department_dir = os.path.join(base_dest_dir, department_for_pdf)
                os.makedirs(department_dir, exist_ok=True) # フォルダがなければ作成
                final_dest_dir = department_dir

            # 最終的なパスを構築
            final_dest_path = os.path.join(final_dest_dir, os.path.basename(self.current_pdf_path))
            # --- ここまで ---

            shutil.copy2(self.current_pdf_path, final_dest_path)
            self.log(f"注文書を正式な保存先にコピーしました: {final_dest_path}")
        except Exception as e:
            messagebox.showerror("ファイルコピーエラー", f"PDFを保存フォルダにコピーできませんでした。\n{e}")
            return

        if not messagebox.askyesno("メール送信確認", f"{self.bottom_pane.to_var.get()} 宛にメールを送信します。よろしいですか？"): return
        
        self.processing = True
        self.toggle_buttons(False)
        self.start_spinner()
        threading.Thread(target=self.run_thread, args=(self.send_mail_task,)).start()
        self.master.after(config.AppConstants.QUEUE_CHECK_INTERVAL, self.check_queue)

    def open_settings_window(self):
        settings_win = settings_gui.SettingsWindow(self.master)
        self.master.wait_window(settings_win)
        result = getattr(settings_win, "save_result", None)
        if result and result.get("saved"):
            import importlib
            importlib.reload(config)
            self.reload_ui_after_settings_change(message=result.get("message"))
        else:
            self.log(result.get("message", "設定変更をキャンセルしました。"))

    def open_current_pdf(self, event=None):
        if self.current_pdf_path and os.path.exists(self.current_pdf_path):
            try:
                os.startfile(self.current_pdf_path)
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルを開けませんでした。\n{e}")
        else:
            messagebox.showwarning("ファイルなし", "PDFファイルが見つかりません。")

    def on_closing(self):
        if self.processing: return messagebox.showwarning("処理中", "処理が実行中です。終了できません。")
        self.cleanup()
        self.master.destroy()

    # --- スレッドタスク ---
    def run_thread(self, task_func, *args):
        try:
            with contextlib.redirect_stdout(self.queue_io):
                task_func(*args)
        except Exception as e:
            self.q.put(("log", f"\nスレッド処理中にエラーが発生しました: {e}", "error"))
            self.q.put(("task_complete", None))

    def get_data_task(self):
        self.log("----------------------------------------")
        self.log("STEP 2: Notionからデータを取得")
        self.log("----------------------------------------")
        self.log(f"部署名「{', '.join(self.selected_departments)}」でフィルタリング中...\nNotionからデータ取得中..." if self.selected_departments else "部署名フィルターは未選択です。\nNotionからデータ取得中...")
        
        # 専門関数を呼び出すだけに変更
        processed_data = notion_api.fetch_and_process_orders(department_names=self.selected_departments)
        
        order_count = len(processed_data.get("all_orders", []))
        self.log(f"✅ 完了 ({order_count}件の要発注データが見つかりました)")
        self.q.put(("update_data_ui", processed_data))



    def send_mail_task(self):
        # --- 必要な情報をUIスレッドから取得 ---
        account_key = self.display_name_to_key_map.get(self.selected_account_display_name.get())
        if not account_key:
            self.log("エラー: 送信者アカウントが選択されていません。", "error")
            return self.q.put(("task_complete", None))

        sender_creds = self.accounts[account_key]
        selected_iids = self.middle_pane.supplier_listbox.selection()
        if not selected_iids:
            self.log("エラー: 送信する仕入先が選択されていません。", "error")
            return self.q.put(("task_complete", None))

        selected_supplier = self.middle_pane.supplier_listbox.item(selected_iids[0], 'values')[0]
        items = self.orders_by_supplier.get(selected_supplier, [])
        supplier_departments = []
        for item in items:
            for dept in (item.get("departments") or []):
                dept = dept.strip()
                if dept and dept not in supplier_departments:
                    supplier_departments.append(dept)
        department_for_mail = next((dept for dept in self.selected_departments if dept in supplier_departments), None)
        if not department_for_mail and supplier_departments:
            department_for_mail = supplier_departments[0]

        self.log(f"「{selected_supplier}」宛にメールを送信中 (From: {sender_creds['sender']})...")

        success, error_message = email_service.prepare_and_send_order_email(
            account_key,
            sender_creds,
            items,
            self.current_pdf_path,
            department_for_mail
        )

        if success:
            self.q.put(("ask_and_update_notion", (selected_supplier, [item['page_id'] for item in items])))
        else:
            user_message = error_message or "メール送信に失敗しました。詳細はログを確認してください。"
            self.log(f"✗ {user_message}", "error")
            self.q.put(("email_error", user_message))
            self.q.put(("task_complete", None))

    def update_notion_task(self, page_ids):
        notion_api.update_notion_pages(page_ids)
        supplier = next((item["supplier_name"] for item in self.order_data if item["page_id"] == page_ids[0]), None)
        if supplier: self.q.put(("mark_as_sent_after_update", supplier))

    def pregenerate_pdfs_task(self):
        """全ての仕入先のPDFをバックグラウンドで事前生成する"""
        self.log("\n----------------------------------------")
        self.log("STEP 3a: 注文書をバックグラウンドで準備中...")
        self.log("----------------------------------------")
        
        account_key = self.display_name_to_key_map.get(self.selected_account_display_name.get())
        if not account_key:
            self.log("エラー: 送信者アカウントが不明なため、PDFの事前生成を中止しました。", "error")
            self.q.put(("task_complete", None))
            return

        sender_creds = self.accounts[account_key]
        department_guidance_numbers = config.load_department_guidance_numbers()

        def resolve_departments(items):
            departments = []
            for item in items:
                for dept in (item.get("departments") or []):
                    dept = dept.strip()
                    if dept and dept not in departments:
                        departments.append(dept)
            return departments

        def render_pdf(supplier, items):
            departments = resolve_departments(items)
            department_for_pdf = next((dept for dept in self.selected_departments if dept in departments), None)
            if not department_for_pdf and departments:
                department_for_pdf = departments[0]
            raw_guidance = department_guidance_numbers.get(department_for_pdf, "")
            guidance_number = "".join(filter(str.isdigit, raw_guidance))
            sender_info = {
                "name": sender_creds.get("display_name", account_key),
                "email": sender_creds["sender"],
                "guidance_number": guidance_number
            }
            pdf_path, _, error_message = pdf_generator.generate_order_pdf_flow(
                supplier,
                items,
                sender_info,
                selected_department=department_for_pdf,
                save_dir=self.temp_dir.name
            )
            return supplier, pdf_path, error_message

        futures = []
        total_suppliers = len(self.orders_by_supplier)
        max_workers = max(1, min(4, total_suppliers))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            for supplier, items in self.orders_by_supplier.items():
                if not items:
                    continue
                futures.append(executor.submit(render_pdf, supplier, items))
            for future in as_completed(futures):
                try:
                    supplier, pdf_path, error_message = future.result()
                except Exception as exc:
                    self.log(f"    -> PDF準備中に例外が発生しました: {exc}", "error")
                    continue
                if pdf_path:
                    self.pregenerated_pdfs[supplier] = pdf_path
                else:
                    self.log(f"    -> 準備中にエラーが発生: {error_message}", "error")
        self.log("✅ 全ての注文書の準備が完了しました。", "emphasis")
        self.q.put(("task_complete", None))


    # --- キュー処理 ---
    def check_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                command, message = item[0], item[1]
                if command == "log": self.log(message.strip(), item[2] if len(item) > 2 else None)
                elif command == "update_data_ui": self.update_data_ui(message)
                elif command == "ask_and_update_notion": self.ask_and_update_notion(message[0], message[1])
                elif command == "mark_as_sent_after_update": self.mark_as_sent(message)
                elif command == "update_preview_ui": self.update_preview_ui(message)
                elif command == "email_error": self.show_email_send_error(message)
                elif command == "task_complete":
                    self.processing = False
                    self.toggle_buttons(True)
                    self.stop_spinner()
        except queue.Empty: 
            self.master.after(config.AppConstants.QUEUE_CHECK_INTERVAL, self.check_queue)
        except Exception as e:
            self.log(f"UI更新中に致命的なエラーが発生しました: {e}", "error")
            self.processing = False
            self.toggle_buttons(True)
            self.stop_spinner()
            self.master.after(config.AppConstants.QUEUE_CHECK_INTERVAL, self.check_queue)

    # --- UI更新メソッド ---
    def log(self, message, tag=None):
        self.bottom_pane.log(message, tag)

    def show_email_send_error(self, message: str):
        suggestion = "不明なエラーが発生しました。ログを確認してください。"
        if "PDFファイルが見つかりません" in message:
            suggestion = "一度プレビューを閉じてPDFを再生成してから再試行してください。"
        elif "宛先メールアドレス" in message:
            suggestion = "Notionの仕入先情報に宛先メールアドレスが設定されているか確認してください。"
        elif "資格情報ストア" in message or "パスワード" in message:
            suggestion = "[設定]画面で対象アカウントを開き、パスワードを再保存してください。"
        elif "SMTP認証" in message:
            suggestion = "メールアドレスとパスワードが正しいか確認した上で再試行してください。"
        elif "接続できません" in message:
            suggestion = "ネットワーク環境やSMTPサーバー設定を確認してください。"
        messagebox.showerror("メール送信エラー", f"{message}\n\n対処ヒント: {suggestion}")

    def update_data_ui(self, processed_data):
        # 事前処理済みのデータを展開
        self.orders_by_supplier = processed_data.get("orders_by_supplier", {})
        all_orders = processed_data.get("all_orders", [])
        unlinked_count = processed_data.get("unlinked_count", 0)
        self.order_data = all_orders

        # UIの更新
        self.sent_suppliers.clear()
        self.middle_pane.tree.delete(*self.middle_pane.tree.get_children())
        self.middle_pane.update_supplier_list(all_orders)
        
        if unlinked_count > 0:
            self.log("", None)
            self.log(f"⚠️ 警告: 仕入先が未設定の「要発注」データが{unlinked_count}件見つかりました。\n       これらのデータはリストに表示されていません。\n       Notionで「DB_仕入先リスト」を設定してください。", "error")
        
        if not all_orders:
            self.log("-> 発注対象のデータは見つかりませんでした。")
            self.q.put(("task_complete", None))
        else:
            self.log("\n----------------------------------------")
            self.log("STEP 3: 仕入先の選択")
            self.log("----------------------------------------")
            self.log("▼リストから仕入先を選択してください。", "emphasis")
            # PDFの事前生成をバックグラウンドで開始
            self.start_spinner()
            threading.Thread(target=self.run_thread, args=(self.pregenerate_pdfs_task,)).start()

    def update_preview_ui(self, data):
        info, pdf_path = data
        self.bottom_pane.update_preview(info, pdf_path)
        if pdf_path:
            self.log("✅ 完了")
            self.log(f"  -> {os.path.basename(pdf_path)}")
        else:
            self.log("❌ PDF作成に失敗しました。")
        self.log("\n----------------------------------------")
        self.log("STEP 5: 内容の確認とメール送信")
        self.log("----------------------------------------")
        self.log("▼生成されたPDFプレビューで、以下の内容が正しいかご確認ください。", "emphasis")
        self.log("  (ファイル名をクリックするとPDFが開きます)", "emphasis")
        self.log("  - 宛先", "emphasis")
        self.log("  - 担当者", "emphasis")
        self.log("  - 注文内容", "emphasis")
        self.log("\n-> 問題がなければ「メール送信」ボタンをクリックしてください。", "emphasis")
        if pdf_path: self.send_mail_button.config(state="normal")
        self.q.put(("task_complete", None))

    def ask_and_update_notion(self, supplier, page_ids):
        if messagebox.askyesno("Notion更新確認", f"メール送信が完了しました。\n\n「{supplier}」のNotionページの「発注日」を更新しますか？"):
            self.processing = True
            self.toggle_buttons(False)
            self.log(f"「{supplier}」のNotionページを更新中...")
            self.start_spinner()
            threading.Thread(target=self.run_thread, args=(self.update_notion_task, page_ids)).start()
        else:
            self.mark_as_sent(supplier, updated=False)

    def mark_as_sent(self, supplier, updated=True):
        self.sent_suppliers.add(supplier)
        self.middle_pane.mark_supplier_as_sent(supplier)
        self.clear_preview()
        self.log(f"-> 「{supplier}」は送信済みとしてマークされました。({'更新済み' if updated else '更新スキップ'})")
        self.q.put(("task_complete", None))

    def clear_displays(self): 
        self.middle_pane.clear_displays()
        self.sent_suppliers.clear()
        self.clear_preview()
        self.bottom_pane.clear_log()

    def clear_preview(self):
        self.bottom_pane.clear_preview()
        self.send_mail_button.config(state="disabled")

    def reset_temp_storage(self):
        """Reset temporary PDF storage directory."""
        try:
            if self.temp_dir:
                self.temp_dir.cleanup()
        except Exception:
            pass
        self.temp_dir = tempfile.TemporaryDirectory()
        self.pregenerated_pdfs = {}

    def cleanup(self):
        """アプリケーション終了時にリソースをクリーンアップする"""
        try:
            self.temp_dir.cleanup()
        except Exception:
            pass

    def toggle_buttons(self, enabled):
        self.top_pane.toggle_buttons(enabled)

    def reload_ui_after_settings_change(self, message=None):
        if message:
            self.log(message)
        else:
            self.log("設定をリロードしました。")
        self.accounts = config.load_email_accounts()
        self.department_defaults = config.load_department_defaults()
        self.display_name_to_key_map = {v.get('display_name', k): k for k, v in self.accounts.items()}
        self.top_pane.account_selector['values'] = sorted(list(self.display_name_to_key_map.keys()))
        self.set_default_sender_account()

    def set_default_sender_account(self):
        if not self.accounts: return
        default_account_key = None
        selected_departments = [name for name, var in self.department_vars.items() if var.get()]
        if selected_departments:
            for dep_name in selected_departments:
                if dep_name in self.department_defaults:
                    default_account_key = self.department_defaults[dep_name]
                    break
        if not default_account_key:
            if self.department_defaults:
                first_dep_name = next(iter(self.department_defaults.keys()), None)
                if first_dep_name:
                    default_account_key = self.department_defaults.get(first_dep_name)
        if not default_account_key or default_account_key not in self.accounts:
            default_account_key = next(iter(self.accounts), None)
        if default_account_key:
            default_display_name = self.accounts[default_account_key].get("display_name", default_account_key)
            self.selected_account_display_name.set(default_display_name)
