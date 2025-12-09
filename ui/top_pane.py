"""上部のアクションとフィルター領域のUI"""
import tkinter as tk
from tkinter import ttk
from typing import TYPE_CHECKING

import config

if TYPE_CHECKING:
    from controllers.app_controller import Application


class TopPane(ttk.Frame):
    """上部のアクションとフィルター領域のUI"""
    def __init__(self, master: ttk.Frame, app: 'Application') -> None:
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

        title_label = ttk.Label(department_container, text="部署名フィルター", font=("Yu Gothic UI", 12, "bold"), foreground=self.app.PRIMARY_COLOR, background="#E0E9FF")
        title_label.pack(anchor="w", padx=10, pady=(0, 2))

        departments = config.load_departments()
        self.app.department_vars = {name: tk.BooleanVar() for name in departments}
        
        checkbox_container = ttk.Frame(department_container, style="Highlight.TFrame")
        checkbox_container.pack(fill=tk.X, padx=5)

        for i, name in enumerate(departments):
            cb = ttk.Checkbutton(checkbox_container, text=name, variable=self.app.department_vars[name], style="Highlight.TCheckbutton", command=self.app.on_department_selection_change)
            cb.grid(row=0, column=i, padx=(5, 15), pady=(0, 5), sticky='w')

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
        def on_combobox_selected(event: tk.Event) -> None:
            # フォーカスを外して選択状態を解除
            account_contents.focus_set()
            # エントリ部分の選択範囲を解除
            try:
                # ttk.Comboboxのエントリ部分にアクセス
                entry_widget = self.account_selector.children.get('!entry')
                if entry_widget:
                    entry_widget.selection_clear()
            except (AttributeError, KeyError):
                pass
        self.account_selector.bind("<<ComboboxSelected>>", on_combobox_selected)

        sender_label_frame = ttk.Frame(account_contents, style="Highlight.TFrame")
        sender_label_frame.pack(side=tk.LEFT, fill=tk.X, padx=(0, 10), pady=4)
        ttk.Label(sender_label_frame, text="送信元:", font=("Yu Gothic UI", 10), background="#E0E9FF").pack(side=tk.LEFT)
        ttk.Label(sender_label_frame, textvariable=self.app.sender_email_var, font=("Yu Gothic UI", 10, "bold"), background="#E0E9FF").pack(side=tk.LEFT, padx=5)

        # --- 1a. データ取得ボタン ---
        action_frame = ttk.Frame(self)
        action_frame.pack(fill=tk.X, pady=(0, 10), anchor='w')

        self.get_data_button = ttk.Button(action_frame, text="Notionからデータを取得", command=self.app.start_data_retrieval, style="Primary.TButton")
        self.get_data_button.pack(side=tk.LEFT, ipady=5)

        self.spinner_label = ttk.Label(action_frame, textvariable=self.app.spinner_var, font=("Yu Gothic UI", 16), foreground=self.app.ACCENT_COLOR)

    def toggle_buttons(self, enabled: bool) -> None:
        """ボタンの有効/無効を切り替える"""
        state = "normal" if enabled else "disabled"
        self.get_data_button.config(state=state)

    def start_spinner(self) -> None:
        """スピナーを表示する"""
        self.spinner_label.pack(side=tk.LEFT, padx=(10, 0), anchor='center')

    def stop_spinner(self) -> None:
        """スピナーを非表示にする"""
        self.spinner_label.pack_forget()

