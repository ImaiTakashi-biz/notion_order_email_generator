import os
import sys
import queue
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk

# 作成したモジュールをインポート
import config
import notion_api
import email_service
import pdf_generator

class QueueIO:
    def __init__(self, q):
        self.q = q
    def write(self, text):
        self.q.put(("log", text))
    def flush(self):
        pass

class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.q = queue.Queue()
        self.queue_io = QueueIO(self.q)
        self.processing = False
        self.order_data = []
        self.current_pdf_path = None
        self.sent_suppliers = set()
        self.department_vars = {}
        
        self.selected_account_display_name = tk.StringVar()
        self.sender_email_var = tk.StringVar()

        # スピナー関連
        self.spinner_var = tk.StringVar()
        self.spinner_running = False
        self.spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.spinner_index = 0
        self._after_id = None

        # アカウント情報をロードし、表示名と内部キーの対応表を作成
        self.accounts = config.load_email_accounts()
        self.display_name_to_key_map = {v.get('display_name', k): k for k, v in self.accounts.items()}
        
        self.selected_account_display_name.trace_add("write", self.update_sender_label)

        self.configure_styles()
        self.create_widgets()
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 初期値を設定
        if self.accounts:
            # JSONファイルで最初に記述されているアカウントをデフォルトにする
            first_account_key = next(iter(self.accounts))
            default_display_name = self.accounts[first_account_key].get("display_name", first_account_key)
            self.selected_account_display_name.set(default_display_name)
        
        self.update_sender_label() # 初期表示

    def configure_styles(self):
        style = ttk.Style(self.master)
        style.theme_use("clam")

        # ---モダンなライトテーマのカラーパレット (アクセント追加) ---
        self.BG_COLOR = "#F8F9FA"
        self.TEXT_COLOR = "#212529"
        self.PRIMARY_COLOR = "#4A90E2"      # 落ち着いた青
        self.ACCENT_COLOR = "#8A2BE2"       # アクセント: ブルーバイオレット
        self.LIGHT_BG = "#FFFFFF"
        self.HEADER_FG = "#FFFFFF"
        self.SELECT_BG = "#E8DAF5"          # 選択行の背景色 (紫系)
        self.SELECT_FG = "#000000"          # 選択行のテキスト色
        self.EMPHASIS_COLOR = self.ACCENT_COLOR # 強調色
        self.BORDER_COLOR = "#DEE2E6"

        style.configure(".", background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=("Yu Gothic UI", 10))
        style.configure("TFrame", background=self.BG_COLOR)
        style.configure("TLabel", background=self.BG_COLOR, foreground=self.TEXT_COLOR)
        style.configure("Light.TLabel", background=self.LIGHT_BG) # 白背景のLabelスタイル

        # PDFリンク用のスタイル
        style.configure("PdfLink.TLabel", background=self.LIGHT_BG, foreground=self.ACCENT_COLOR, font=("Yu Gothic UI", 9, "underline"))
        style.configure("TLabelFrame", background=self.BG_COLOR, bordercolor=self.BORDER_COLOR, relief="solid", borderwidth=1)
        style.configure("TLabelFrame.Label", background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=("Yu Gothic UI", 11, "bold")) # フォントサイズを11に戻す

        style.configure("TButton", font=("Yu Gothic UI", 10, "bold"), borderwidth=0, padding=(15, 10))
        style.configure("Primary.TButton", background=self.PRIMARY_COLOR, foreground=self.HEADER_FG)
        style.map("Primary.TButton", background=[("active", "#357ABD"), ("disabled", "#A9CCE3")])
        
        style.configure("Treeview", background=self.LIGHT_BG, fieldbackground=self.LIGHT_BG, foreground=self.TEXT_COLOR, rowheight=28, font=("Yu Gothic UI", 10))
        style.map("Treeview", background=[("selected", self.SELECT_BG)], foreground=[("selected", self.SELECT_FG)])
        style.configure("Treeview.Heading", background=self.PRIMARY_COLOR, foreground=self.HEADER_FG, font=("Yu Gothic UI", 10, "bold"), padding=8)
        style.map("Treeview.Heading", background=[("active", "#357ABD")])

        style.configure("TCombobox", font=("Yu Gothic UI", 10), fieldbackground=self.LIGHT_BG, padding=5)
        style.configure("Vertical.TScrollbar", background=self.PRIMARY_COLOR, troughcolor=self.BG_COLOR, arrowcolor=self.HEADER_FG)
        style.configure("Horizontal.TScrollbar", background=self.PRIMARY_COLOR, troughcolor=self.BG_COLOR, arrowcolor=self.HEADER_FG)
        
        style.configure("TPanedWindow", background=self.BORDER_COLOR)

        # --- 強調用スタイル ---
        style.configure("Highlight.TFrame", background="#EAF2F8", bordercolor=self.PRIMARY_COLOR) # 枠線の色も変更
        style.configure("Light.TFrame", background=self.LIGHT_BG) # 白背景のFrameスタイル
        
        # チェックボタンのスタイル
        style.configure("Highlight.TCheckbutton", background="#EAF2F8", font=("Yu Gothic UI", 10))
        style.map("Highlight.TCheckbutton",
                  background=[('active', '#EAF2F8')],
                  indicatorbackground=[('active', '#EAF2F8')],
                  # 選択されたら文字色を濃い青に、フォントを太字にする
                  foreground=[('selected', '#357ABD')], # より濃い青に変更
                  font=[('selected', ("Yu Gothic UI", 10, "bold"))])

    def create_widgets(self):
        self.pack(fill=tk.BOTH, expand=True)

        # --- メインコンテナ ---
        main_container = ttk.Frame(self, padding=15)
        main_container.pack(fill=tk.BOTH, expand=True)
        main_container.rowconfigure(1, weight=1) # 中段のペインが伸縮
        main_container.rowconfigure(2, weight=1) # 下段(プレビュー&ログ)が伸縮
        main_container.rowconfigure(3, weight=0) # ボタンエリアは固定
        main_container.columnconfigure(0, weight=1)

        # --- 1. 上段: アクションとアカウント ---
        top_pane = ttk.Frame(main_container)
        top_pane.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        # --- 左側のアクションエリア ---
        left_action_area = ttk.Frame(top_pane)
        left_action_area.pack(side=tk.LEFT, anchor='n', fill=tk.X, expand=True)

        # --- 1a. データ取得ボタン ---
        action_frame = ttk.Frame(left_action_area)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.get_data_button = ttk.Button(action_frame, text="1. Notionからデータを取得", command=self.start_data_retrieval, style="Primary.TButton")
        self.get_data_button.pack(side=tk.LEFT, ipady=5)

        self.spinner_label = ttk.Label(action_frame, textvariable=self.spinner_var, font=("Yu Gothic UI", 16), foreground=self.ACCENT_COLOR)

        # --- 1b. 部署名フィルター (強調表示) ---
        # LabelFrameの代わりにFrameとLabelで模倣する
        department_container = ttk.Frame(left_action_area, style="Highlight.TFrame", relief="solid", borderwidth=1, padding=(0, 5))
        department_container.pack(fill=tk.X, pady=(5,0))

        # タイトルラベル
        title_label = ttk.Label(department_container, text="部署名フィルター", font=("Yu Gothic UI", 12, "bold"), foreground=self.PRIMARY_COLOR, background="#EAF2F8") # フォントサイズを12に戻す
        title_label.pack(anchor="w", padx=10, pady=(0, 2))

        departments = ["生産部", "品質保証部", "営業部", "総務部"]
        self.department_vars = {name: tk.BooleanVar() for name in departments}
        
        # チェックボックスを配置する内部フレーム
        checkbox_container = ttk.Frame(department_container, style="Highlight.TFrame")
        checkbox_container.pack(fill=tk.X, padx=5)

        for i, name in enumerate(departments):
            # チェックボタンにもスタイルを適用して背景色を合わせる
            cb = ttk.Checkbutton(checkbox_container, text=name, variable=self.department_vars[name], style="Highlight.TCheckbutton")
            cb.grid(row=0, column=i, padx=(5, 15), pady=(0,5), sticky='w')

        # --- 右側のアカウント選択 ---
        account_frame = ttk.LabelFrame(top_pane, text="送信者アカウント")
        account_frame.pack(side=tk.RIGHT, anchor='n', padx=(20, 0))
        
        account_display_names = sorted(list(self.display_name_to_key_map.keys()))
        self.account_selector = ttk.Combobox(account_frame, textvariable=self.selected_account_display_name, values=account_display_names, state="readonly", width=25, font=("Yu Gothic UI", 10))
        self.account_selector.pack(side=tk.LEFT, padx=10, pady=10)
        
        sender_label_frame = ttk.Frame(account_frame)
        sender_label_frame.pack(side=tk.LEFT, fill=tk.X, padx=(0,10), pady=10)
        ttk.Label(sender_label_frame, text="送信元:").pack(side=tk.LEFT)
        ttk.Label(sender_label_frame, textvariable=self.sender_email_var, font=("Yu Gothic UI", 10, "bold")).pack(side=tk.LEFT, padx=5)

        # --- 2. 中段: データ選択 (仕入先リストとデータテーブル) ---
        middle_pane = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        middle_pane.grid(row=1, column=0, sticky="nsew", pady=0)

        supplier_pane = ttk.LabelFrame(middle_pane, text="2. 仕入先を選択")
        middle_pane.add(supplier_pane, weight=1)
        
        table_pane = ttk.LabelFrame(middle_pane, text="発注対象データ")
        middle_pane.add(table_pane, weight=3)

        self.supplier_listbox = tk.Listbox(supplier_pane, exportselection=False, bg=self.LIGHT_BG, fg=self.TEXT_COLOR, selectbackground=self.SELECT_BG, selectforeground=self.SELECT_FG, font=("Yu Gothic UI", 11), relief="flat", borderwidth=0, highlightthickness=0)
        self.supplier_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(1,0), pady=1)
        self.supplier_listbox.bind("<<ListboxSelect>>", self.on_supplier_select)
        supplier_vsb = ttk.Scrollbar(supplier_pane, orient="vertical", command=self.supplier_listbox.yview, style="Vertical.TScrollbar")
        self.supplier_listbox.config(yscrollcommand=supplier_vsb.set)
        supplier_vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=1)

        columns = ("maker", "part_num", "qty")
        self.tree = ttk.Treeview(table_pane, columns=columns, show="headings")
        self.tree.heading("maker", text="メーカー"); self.tree.heading("part_num", text="品番"); self.tree.heading("qty", text="数量")
        self.tree.column("maker", width=150); self.tree.column("part_num", width=250); self.tree.column("qty", width=60, anchor=tk.E)
        vsb = ttk.Scrollbar(table_pane, orient="vertical", command=self.tree.yview, style="Vertical.TScrollbar")
        hsb = ttk.Scrollbar(table_pane, orient="horizontal", command=self.tree.xview, style="Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=1); hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=1); self.tree.pack(fill=tk.BOTH, expand=True, padx=1, pady=(1,0))

        # --- 3. 下段: プレビューとログ (左右分割) ---
        bottom_and_log_pane = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        bottom_and_log_pane.grid(row=2, column=0, sticky="nsew", pady=(15, 0))

        # --- 3a. プレビューエリア (左) ---
        # ttk.LabelFrameに戻し、スタイルで色を統一する
        preview_frame = ttk.LabelFrame(bottom_and_log_pane, text="3. 内容を確認して送信")
        bottom_and_log_pane.add(preview_frame, weight=1)

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
        self.pdf_label.bind("<Button-1>", self.open_current_pdf)
        preview_grid.columnconfigure(1, weight=1)

        # --- 3b. ログエリア (右) ---
        log_frame = ttk.LabelFrame(bottom_and_log_pane, text="ログ")
        bottom_and_log_pane.add(log_frame, weight=1) # weight=1で均等分割

        self.log_display = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD, bg=self.LIGHT_BG, fg=self.TEXT_COLOR, font=("Consolas", 11), relief="flat", borderwidth=0, highlightthickness=0)
        self.log_display.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self.log_display.configure(state='disabled')
        self.log_display.tag_configure("emphasis", foreground=self.EMPHASIS_COLOR, font=("Yu Gothic UI", 12, "bold"))

        # --- 4. 最下段: メール送信ボタン ---
        send_button_area = ttk.Frame(main_container)
        send_button_area.grid(row=3, column=0, sticky="ew", pady=(15, 0))
        
        # 中央に配置するためのコンテナ
        send_button_container = ttk.Frame(send_button_area)
        send_button_container.pack(expand=True)

        self.send_mail_button = ttk.Button(send_button_container, text="メール送信", command=self.send_single_mail, state="disabled", style="Primary.TButton")
        self.send_mail_button.pack(ipadx=40, ipady=15) # サイズ調整は維持

    def start_spinner(self):
        if self.spinner_running: return
        self.spinner_running = True
        self.spinner_label.pack(side=tk.LEFT, padx=(10, 0), anchor='center')
        self.animate_spinner()

    def stop_spinner(self):
        if not self.spinner_running: return
        if self._after_id: self.master.after_cancel(self._after_id)
        self._after_id = None
        self.spinner_var.set("")
        self.spinner_label.pack_forget()
        self.spinner_running = False

    def animate_spinner(self):
        if not self.spinner_running: return
        self.spinner_var.set(f"Loading {self.spinner_frames[self.spinner_index]}")
        self.spinner_index = (self.spinner_index + 1) % len(self.spinner_frames)
        self._after_id = self.master.after(80, self.animate_spinner)

    def update_sender_label(self, *args):
        selected_display_name = self.selected_account_display_name.get()
        account_key = self.display_name_to_key_map.get(selected_display_name)
        if account_key and account_key in self.accounts:
            self.sender_email_var.set(self.accounts[account_key]["sender"])
        else:
            self.sender_email_var.set("")

    def start_data_retrieval(self):
        if self.processing: return
        
        env_missing = []
        if not config.NOTION_API_TOKEN: env_missing.append("・Notion APIトークン (NOTION_API_TOKEN)")
        if not config.PAGE_ID_CONTAINING_DB: env_missing.append("・Notion データベースID (NOTION_DATABASE_ID)")
        if not self.accounts: env_missing.append("・Emailアカウント (email_accounts.json)")
        if not config.EXCEL_TEMPLATE_PATH: env_missing.append("・Excelテンプレートのパス (EXCEL_TEMPLATE_PATH)")
        if not config.PDF_SAVE_DIR: env_missing.append("・PDF保存先フォルダ (PDF_SAVE_DIR)")
        if env_missing:
            messagebox.showerror("設定エラー", "以下の設定を確認してください。\n\n" + "\n".join(env_missing))
            return

        if not os.path.exists(config.EXCEL_TEMPLATE_PATH):
            messagebox.showerror("設定エラー (Excel)", f"""Excelテンプレートが見つかりません。
パスを確認してください。

現在のパス: {config.EXCEL_TEMPLATE_PATH}"""
            )
            return

        self.selected_departments = [name for name, var in self.department_vars.items() if var.get()]

        self.processing = True; self.toggle_buttons(False); self.clear_displays()
        self.start_spinner()
        threading.Thread(target=self.run_thread, args=(self.get_data_task,)).start()
        self.master.after(100, self.check_queue)

    def on_supplier_select(self, event=None):
        if self.processing or not self.supplier_listbox.curselection(): return
        selected_supplier = self.supplier_listbox.get(self.supplier_listbox.curselection())
        if selected_supplier in self.sent_suppliers: self.clear_preview(); return
        self.processing = True; self.toggle_buttons(False)
        self.start_spinner()
        self.update_table_for_supplier(selected_supplier)
        threading.Thread(target=self.run_thread, args=(self.pdf_creation_flow_task,)).start()
        self.master.after(100, self.check_queue)

    def send_single_mail(self):
        if self.processing or not self.current_pdf_path: return
        if not messagebox.askyesno("メール送信確認", f"{self.to_var.get()} 宛にメールを送信します。よろしいですか？"): return
        self.processing = True; self.toggle_buttons(False)
        self.start_spinner()
        threading.Thread(target=self.run_thread, args=(self.send_mail_task,)).start()
        self.master.after(100, self.check_queue)

    def run_thread(self, task_func, *args):
        original_stdout = sys.stdout; sys.stdout = self.queue_io
        try:
            task_func(*args)
        except Exception as e:
            print(f"\nスレッド処理中にエラーが発生しました: {e}")
            self.q.put(("task_complete", None))
        finally:
            sys.stdout = original_stdout

    def get_data_task(self):
        if self.selected_departments:
            print(f"部署名「{', '.join(self.selected_departments)}」でフィルタリング中...")
        else:
            print("部署名フィルターは未選択です。")
        print("Notionからデータ取得中...")
        data = notion_api.get_order_data_from_notion(department_names=self.selected_departments)
        self.q.put(("update_data_ui", data))

    def pdf_creation_flow_task(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            selected_display_name = self.selected_account_display_name.get()
            account_key = self.display_name_to_key_map.get(selected_display_name)
            if not account_key: return self.q.put(("task_complete", None))
            
            sender_creds = self.accounts[account_key]
            sender_info = {"name": sender_creds.get("display_name", account_key), "email": sender_creds["sender"]}
            selected_supplier = self.supplier_listbox.get(self.supplier_listbox.curselection())
            
            print(f"「{selected_supplier}」の注文書PDF作成準備中...")
            items = [item for item in self.order_data if item["supplier_name"] == selected_supplier]
            if not items: return self.q.put(("task_complete", None))

            sales_contact = items[0]["sales_contact"]
            pdf_path = pdf_generator.create_order_pdf(selected_supplier, items, sales_contact, sender_info)

            if pdf_path:
                self.q.put(("update_preview_ui", (items[0], pdf_path)))
            else:
                self.q.put(("task_complete", None))
        finally:
            pythoncom.CoUninitialize()

    def send_mail_task(self):
        selected_display_name = self.selected_account_display_name.get()
        account_key = self.display_name_to_key_map.get(selected_display_name)
        if not account_key: return self.q.put(("task_complete", None))

        sender_creds = self.accounts[account_key]
        display_name = sender_creds.get("display_name", account_key)
        selected_supplier = self.supplier_listbox.get(self.supplier_listbox.curselection())
        print(f"「{selected_supplier}」宛にメールを送信中 (From: {sender_creds['sender']})...")
        items = [item for item in self.order_data if item["supplier_name"] == selected_supplier]
        success = email_service.send_smtp_mail(items[0], self.current_pdf_path, sender_creds, display_name)
        if success:
            page_ids_to_update = [item['page_id'] for item in items]
            self.q.put(("ask_and_update_notion", (selected_supplier, page_ids_to_update)))
        else:
            self.q.put(("task_complete", None))

    def check_queue(self):
        try:
            while True:
                command, data = self.q.get_nowait()
                if command == "log": self.log(data.strip())
                elif command == "update_data_ui": self.update_data_ui(data)
                elif command == "ask_and_update_notion": self.ask_and_update_notion(data[0], data[1])
                elif command == "mark_as_sent_after_update": self.mark_as_sent(data)
                elif command == "update_preview_ui": self.update_preview_ui(data)
                elif command == "task_complete":
                    self.processing = False
                    self.toggle_buttons(True)
                    self.stop_spinner()
        except queue.Empty: 
            self.master.after(100, self.check_queue)
        except Exception as e:
            error_message = f"UI更新中に致命的なエラーが発生しました: {e}"
            self.log(error_message, "emphasis")
            print(error_message)
            import traceback
            traceback.print_exc()
            self.processing = False
            self.toggle_buttons(True)
            self.master.after(100, self.check_queue)

    def log(self, message, tag=None):
        self.log_display.config(state="normal")
        if tag:
            self.log_display.insert(tk.END, message + "\n", tag)
        else:
            self.log_display.insert(tk.END, message + "\n")
        self.log_display.see(tk.END)
        self.log_display.config(state="disabled")

    def update_data_ui(self, data):
        self.order_data = data; self.sent_suppliers.clear()
        self.tree.delete(*self.tree.get_children())
        self.update_supplier_list(data)
        if not data:
            self.log("-> 発注対象のデータは見つかりませんでした。")
        else:
            self.log("-> データ取得完了！", "emphasis")
            self.log("   「仕入先を選択」リストから、メールを送信したい仕入先をクリックしてください。", "emphasis")
            self.log("   選択すると、自動的に注文書PDFが作成され、プレビューが表示されます。")
            self.log("")
        self.q.put(("task_complete", None))

    def update_preview_ui(self, data):
        info, pdf_path = data
        self.current_pdf_path = pdf_path
        self.to_var.set(info.get("email", "")); self.cc_var.set(info.get("email_cc", "")); self.contact_var.set(info.get("sales_contact", ""))
        self.pdf_var.set(os.path.basename(pdf_path) if pdf_path else "作成失敗")
        log_message = """-> プレビューの準備ができました。

   宛先、担当者、PDFの内容を必ず確認してください。
   (PDFファイル名をクリックすると内容を開けます)

-> 内容を確認し、問題がなければ「メール送信」ボタンをクリックしてください！"""
        self.log(log_message, "emphasis")
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

    def update_notion_task(self, page_ids):
        notion_api.update_notion_pages(page_ids)
        supplier = next((item["supplier_name"] for item in self.order_data if item["page_id"] == page_ids[0]), None)
        if supplier: self.q.put(("mark_as_sent_after_update", supplier))

    def mark_as_sent(self, supplier, updated=True):
        self.sent_suppliers.add(supplier)
        try:
            idx = self.supplier_listbox.get(0, "end").index(supplier)
            self.supplier_listbox.itemconfig(idx, {'fg': 'gray'}); self.supplier_listbox.selection_clear(idx)
        except ValueError: pass
        self.clear_preview()
        status = "更新済み" if updated else "更新スキップ"
        self.log(f"-> 「{supplier}」は送信済みとしてマークされました。({status})")
        self.q.put(("task_complete", None))

    def update_table_for_supplier(self, supplier_name):
        self.tree.delete(*self.tree.get_children())
        items_to_display = [item for item in self.order_data if item["supplier_name"] == supplier_name]
        for item in items_to_display:
            self.tree.insert("", tk.END, values=(item.get("maker_name", ""), item.get("db_part_number", ""), item.get("quantity", 0)))

    def update_supplier_list(self, data):
        self.supplier_listbox.delete(0, tk.END)
        suppliers = sorted(list(set(i["supplier_name"] for i in data)))
        for s in suppliers: self.supplier_listbox.insert(tk.END, s)

    def clear_displays(self, event=None): 
        self.tree.delete(*self.tree.get_children()); self.supplier_listbox.delete(0, tk.END); self.sent_suppliers.clear(); self.clear_preview()
        self.log_display.config(state="normal"); self.log_display.delete(1.0, tk.END); self.log_display.config(state="disabled")

    def clear_preview(self):
        self.to_var.set(""); self.cc_var.set(""); self.contact_var.set(""); self.pdf_var.set(""); self.current_pdf_path = None; self.send_mail_button.config(state="disabled")

    def toggle_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        self.get_data_button.config(state=state)

    def open_current_pdf(self, event=None):
        if self.current_pdf_path and os.path.exists(self.current_pdf_path):
            try:
                os.startfile(self.current_pdf_path)
            except Exception as e:
                messagebox.showerror("エラー", f"""ファイルを開けませんでした。
{e}"""
                )
        else:
            messagebox.showwarning("ファイルなし", "PDFファイルが見つかりません。")

    def on_closing(self):
        if self.processing: return messagebox.showwarning("処理中", "処理が実行中です。終了できません。")
        self.master.destroy()
