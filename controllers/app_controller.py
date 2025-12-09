"""アプリケーションのメインコントローラー"""
import os
import queue
import threading
import tempfile
import shutil
import contextlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List, Optional, Tuple, Callable
import tkinter as tk
from tkinter import messagebox, ttk

# 作成したモジュールをインポート
import config
import notion_api
import email_service
import pdf_generator
import settings_gui
import logger_config

# UIコンポーネントをインポート
from ui.queue_io import QueueIO
from ui.top_pane import TopPane
from ui.middle_pane import MiddlePane
from ui.bottom_pane import BottomPane

# ロガーの取得
logger = logger_config.get_logger(__name__)


class Application(ttk.Frame):
    """アプリケーションのメインコントローラー"""
    def __init__(self, master: Optional[tk.Tk] = None) -> None:
        super().__init__(master)
        self.master = master
        self.configure_styles()
        self.create_menu()
        self.q = queue.Queue()
        self.queue_io = QueueIO(self.q)
        
        # --- 一時フォルダと事前生成PDFの管理 ---
        self.temp_dir = tempfile.TemporaryDirectory()
        self.pregenerated_pdfs: Dict[str, str] = {}
        
        # --- 状態管理 ---
        self.processing = False
        self.order_data: List[Dict[str, Any]] = []
        self.orders_by_supplier: Dict[str, List[Dict[str, Any]]] = {}
        self.current_pdf_path: Optional[str] = None
        self.sent_suppliers: set = set()
        self.selected_departments: List[str] = []
        
        # --- Tkinter変数 ---
        self.department_vars: Dict[str, tk.BooleanVar] = {}
        self.selected_account_display_name = tk.StringVar()
        self.sender_email_var = tk.StringVar()
        self.spinner_var = tk.StringVar()
        
        # --- スピナー関連 ---
        self.spinner_running = False
        self.spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.spinner_index = 0
        self._after_id: Optional[str] = None
        
        # --- 設定とマッピング ---
        self.accounts = config.load_email_accounts()
        self.department_defaults = config.load_department_defaults()
        self.display_name_to_key_map = {v.get('display_name', k): k for k, v in self.accounts.items()}
        
        self.selected_account_display_name.trace_add("write", self.update_sender_label)
        
        self.configure_styles()
        self.create_widgets()
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.initialize_app_state()
    
    def configure_styles(self) -> None:
        """UIスタイルを設定する"""
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
        style.configure("TCombobox", font=("Yu Gothic UI", 10), fieldbackground=self.LIGHT_BG, background=self.LIGHT_BG, foreground=self.TEXT_COLOR, padding=5, selectbackground=self.LIGHT_BG, selectforeground=self.TEXT_COLOR)
        style.map("TCombobox", 
                  fieldbackground=[("readonly", self.LIGHT_BG)], 
                  background=[("readonly", self.LIGHT_BG)], 
                  foreground=[("readonly", self.TEXT_COLOR)],
                  selectbackground=[("readonly", self.LIGHT_BG)],
                  selectforeground=[("readonly", self.TEXT_COLOR)])
        # Comboboxのエントリ部分のスタイルを設定
        style.configure("TCombobox.field", selectbackground=self.LIGHT_BG, selectforeground=self.TEXT_COLOR, background=self.LIGHT_BG, foreground=self.TEXT_COLOR)
        style.map("TCombobox.field", 
                  selectbackground=[("readonly", self.LIGHT_BG)],
                  selectforeground=[("readonly", self.TEXT_COLOR)],
                  background=[("readonly", self.LIGHT_BG)],
                  foreground=[("readonly", self.TEXT_COLOR)])
        style.configure("Vertical.TScrollbar", background=self.PRIMARY_COLOR, troughcolor=self.BG_COLOR, arrowcolor=self.HEADER_FG)
        style.configure("Horizontal.TScrollbar", background=self.PRIMARY_COLOR, troughcolor=self.BG_COLOR, arrowcolor=self.HEADER_FG)
        style.configure("TPanedWindow", background=self.BORDER_COLOR)
        style.configure("Highlight.TFrame", background="#E0E9FF", bordercolor=self.PRIMARY_COLOR)
        style.configure("Light.TFrame", background=self.LIGHT_BG)
        style.configure("Highlight.TCheckbutton", background="#E0E9FF", font=("Yu Gothic UI", 10))
        style.map("Highlight.TCheckbutton", background=[('active', '#E0E9FF')], indicatorbackground=[('active', '#E0E9FF')], foreground=[('selected', '#1E40AF')], font=[('selected', ("Yu Gothic UI", 10, "bold"))])
    
    def create_menu(self) -> None:
        """アプリケーションメニューを生成する"""
        menubar = tk.Menu(self.master)
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="設定を開く", command=self.open_settings_window)
        settings_menu.add_command(label="設定リロード", command=self.reload_ui_after_settings_change)
        menubar.add_cascade(label="　⚙ 設定", menu=settings_menu)
        self.master.config(menu=menubar)
    
    def create_widgets(self) -> None:
        """ウィジェットを作成する"""
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
    
    def initialize_app_state(self) -> None:
        """アプリケーションの初期状態を設定する"""
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
    def start_spinner(self) -> None:
        """スピナーを開始する"""
        if self.spinner_running: return
        self.spinner_running = True
        self.top_pane.start_spinner()
        self.animate_spinner()
    
    def stop_spinner(self) -> None:
        """スピナーを停止する"""
        if not self.spinner_running: return
        if self._after_id: self.master.after_cancel(self._after_id)
        self._after_id = None
        self.spinner_var.set("")
        self.top_pane.stop_spinner()
        self.spinner_running = False
    
    def animate_spinner(self) -> None:
        """スピナーアニメーションを更新する"""
        if not self.spinner_running: return
        self.spinner_var.set(f"Loading {self.spinner_frames[self.spinner_index]}")
        self.spinner_index = (self.spinner_index + 1) % len(self.spinner_frames)
        self._after_id = self.master.after(config.AppConstants.SPINNER_ANIMATION_DELAY, self.animate_spinner)
    
    # --- UIイベントハンドラ ---
    def on_department_selection_change(self) -> None:
        """部署選択変更時のハンドラ"""
        self.set_default_sender_account()
    
    def update_sender_label(self, *args: Any) -> None:
        """送信者ラベルを更新する"""
        selected_display_name = self.selected_account_display_name.get()
        account_key = self.display_name_to_key_map.get(selected_display_name)
        self.sender_email_var.set(self.accounts[account_key]["sender"] if account_key and account_key in self.accounts else "")
    
    def start_data_retrieval(self) -> None:
        """データ取得を開始する"""
        if self.processing: return
        self.selected_departments = [name for name, var in self.department_vars.items() if var.get()]
        self.reset_temp_storage()
        self.processing = True
        self.toggle_buttons(False)
        self.clear_displays()
        self.start_spinner()
        threading.Thread(target=self.run_thread, args=(self.get_data_task,)).start()
        self.master.after(config.AppConstants.QUEUE_CHECK_INTERVAL, self.check_queue)
    
    def on_supplier_select(self, event: tk.Event) -> None:
        """仕入先選択時のハンドラ"""
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
    
    def send_single_mail(self) -> None:
        """単一メールを送信する"""
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
    
    def open_settings_window(self) -> None:
        """設定ウィンドウを開く"""
        settings_win = settings_gui.SettingsWindow(self.master)
        self.master.wait_window(settings_win)
        result = getattr(settings_win, "save_result", None)
        if result and result.get("saved"):
            import importlib
            importlib.reload(config)
            self.reload_ui_after_settings_change(message=result.get("message"))
        else:
            self.log(result.get("message", "設定変更をキャンセルしました。"))
    
    def open_current_pdf(self, event: Optional[tk.Event] = None) -> None:
        """現在のPDFを開く"""
        if self.current_pdf_path and os.path.exists(self.current_pdf_path):
            try:
                os.startfile(self.current_pdf_path)
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルを開けませんでした。\n{e}")
        else:
            messagebox.showwarning("ファイルなし", "PDFファイルが見つかりません。")
    
    def on_closing(self) -> None:
        """アプリケーション終了時のハンドラ"""
        if self.processing: return messagebox.showwarning("処理中", "処理が実行中です。終了できません。")
        self.cleanup()
        self.master.destroy()
    
    # --- スレッドタスク ---
    def run_thread(self, task_func: Callable[..., None], *args: Any) -> None:
        """バックグラウンドスレッドでタスクを実行する"""
        try:
            with contextlib.redirect_stdout(self.queue_io):
                task_func(*args)
        except Exception as e:
            logger.error(f"スレッド処理中にエラーが発生しました: {e}", exc_info=True)
            self.q.put(("log", f"\nスレッド処理中にエラーが発生しました: {e}", "error"))
            self.q.put(("task_complete", None))
    
    def get_data_task(self) -> None:
        """Notionからデータを取得するタスク"""
        self.log("----------------------------------------")
        self.log("STEP 2: Notionからデータを取得")
        self.log("----------------------------------------")
        self.log(f"部署名「{', '.join(self.selected_departments)}」でフィルタリング中...\nNotionからデータ取得中..." if self.selected_departments else "部署名フィルターは未選択です。\nNotionからデータ取得中...")
        
        # 専門関数を呼び出すだけに変更
        processed_data = notion_api.fetch_and_process_orders(department_names=self.selected_departments)
        
        order_count = len(processed_data.get("all_orders", []))
        self.log(f"✅ 完了 ({order_count}件の要発注データが見つかりました)")
        self.q.put(("update_data_ui", processed_data))
    
    def send_mail_task(self) -> None:
        """メール送信タスク"""
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
    
    def update_notion_task(self, page_ids: List[str]) -> None:
        """Notionページ更新タスク"""
        notion_api.update_notion_pages(page_ids)
        supplier = next((item["supplier_name"] for item in self.order_data if item["page_id"] == page_ids[0]), None)
        if supplier: self.q.put(("mark_as_sent_after_update", supplier))
    
    def pregenerate_pdfs_task(self) -> None:
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
        
        def resolve_departments(items: List[Dict[str, Any]]) -> List[str]:
            """アイテムから部署リストを解決する"""
            departments = []
            for item in items:
                for dept in (item.get("departments") or []):
                    dept = dept.strip()
                    if dept and dept not in departments:
                        departments.append(dept)
            return departments
        
        def render_pdf(supplier: str, items: List[Dict[str, Any]]) -> Tuple[str, Optional[str], Optional[str]]:
            """PDFをレンダリングする"""
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
    def check_queue(self) -> None:
        """キューをチェックしてUIを更新する"""
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
            logger.error(f"UI更新中に致命的なエラーが発生しました: {e}", exc_info=True)
            self.log(f"UI更新中に致命的なエラーが発生しました: {e}", "error")
            self.processing = False
            self.toggle_buttons(True)
            self.stop_spinner()
            self.master.after(config.AppConstants.QUEUE_CHECK_INTERVAL, self.check_queue)
    
    # --- UI更新メソッド ---
    def log(self, message: str, tag: Optional[str] = None) -> None:
        """ログメッセージを表示する"""
        self.bottom_pane.log(message, tag)
    
    def show_email_send_error(self, message: str) -> None:
        """メール送信エラーを表示する"""
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
    
    def update_data_ui(self, processed_data: Dict[str, Any]) -> None:
        """データUIを更新する"""
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
    
    def update_preview_ui(self, data: Tuple[Dict[str, Any], Optional[str]]) -> None:
        """プレビューUIを更新する"""
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
    
    def ask_and_update_notion(self, supplier: str, page_ids: List[str]) -> None:
        """Notion更新を確認して実行する"""
        if messagebox.askyesno("Notion更新確認", f"メール送信が完了しました。\n\n「{supplier}」のNotionページの「発注日」を更新しますか？"):
            self.processing = True
            self.toggle_buttons(False)
            self.log(f"「{supplier}」のNotionページを更新中...")
            self.start_spinner()
            threading.Thread(target=self.run_thread, args=(self.update_notion_task, page_ids)).start()
        else:
            self.mark_as_sent(supplier, updated=False)
    
    def mark_as_sent(self, supplier: str, updated: bool = True) -> None:
        """仕入先を送信済みとしてマークする"""
        self.sent_suppliers.add(supplier)
        self.middle_pane.mark_supplier_as_sent(supplier)
        self.clear_preview()
        self.log(f"-> 「{supplier}」は送信済みとしてマークされました。({'更新済み' if updated else '更新スキップ'})")
        self.q.put(("task_complete", None))
    
    def clear_displays(self) -> None:
        """全ての表示をクリアする"""
        self.middle_pane.clear_displays()
        self.sent_suppliers.clear()
        self.clear_preview()
        self.bottom_pane.clear_log()
    
    def clear_preview(self) -> None:
        """プレビューをクリアする"""
        self.bottom_pane.clear_preview()
        self.send_mail_button.config(state="disabled")
    
    def reset_temp_storage(self) -> None:
        """一時PDFストレージディレクトリをリセットする"""
        try:
            if self.temp_dir:
                self.temp_dir.cleanup()
        except Exception:
            pass
        self.temp_dir = tempfile.TemporaryDirectory()
        self.pregenerated_pdfs = {}
    
    def cleanup(self) -> None:
        """アプリケーション終了時にリソースをクリーンアップする"""
        try:
            self.temp_dir.cleanup()
        except Exception:
            pass
    
    def toggle_buttons(self, enabled: bool) -> None:
        """ボタンの有効/無効を切り替える"""
        self.top_pane.toggle_buttons(enabled)
    
    def reload_ui_after_settings_change(self, message: Optional[str] = None) -> None:
        """設定変更後にUIをリロードする"""
        if message:
            self.log(message)
        else:
            self.log("設定をリロードしました。")
        self.accounts = config.load_email_accounts()
        self.department_defaults = config.load_department_defaults()
        self.display_name_to_key_map = {v.get('display_name', k): k for k, v in self.accounts.items()}
        self.top_pane.account_selector['values'] = sorted(list(self.display_name_to_key_map.keys()))
        self.set_default_sender_account()
    
    def set_default_sender_account(self) -> None:
        """デフォルトの送信者アカウントを設定する"""
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

