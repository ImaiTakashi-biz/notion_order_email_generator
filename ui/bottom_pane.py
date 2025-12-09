"""下部のプレビューとログ領域のUI"""
import os
import tkinter as tk
from tkinter import scrolledtext, ttk
from typing import TYPE_CHECKING, Dict, Any, Optional

if TYPE_CHECKING:
    from controllers.app_controller import Application


class BottomPane(ttk.PanedWindow):
    """下部のプレビューとログ領域のUI"""
    def __init__(self, master: ttk.Frame, app: 'Application') -> None:
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

    def update_preview(self, info: Dict[str, Any], pdf_path: Optional[str]) -> None:
        """プレビュー情報を更新する"""
        self.app.current_pdf_path = pdf_path
        self.to_var.set(info.get("email", ""))
        self.cc_var.set(info.get("email_cc", ""))
        self.contact_var.set(info.get("sales_contact", ""))
        self.pdf_var.set(os.path.basename(pdf_path) if pdf_path else "作成失敗")

    def clear_preview(self) -> None:
        """プレビューをクリアする"""
        self.to_var.set(""); self.cc_var.set(""); self.contact_var.set(""); self.pdf_var.set("")
        self.app.current_pdf_path = None

    def log(self, message: str, tag: Optional[str] = None) -> None:
        """ログメッセージを表示する"""
        self.log_display.config(state="normal")
        if tag:
            self.log_display.insert(tk.END, message + "\n", tag)
        else:
            self.log_display.insert(tk.END, message + "\n")
        self.log_display.see(tk.END)
        self.log_display.config(state="disabled")

    def clear_log(self) -> None:
        """ログをクリアする"""
        self.log_display.config(state="normal")
        self.log_display.delete(1.0, tk.END)
        self.log_display.config(state="disabled")

