import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.simpledialog import askstring
import config
import keyring
from typing import Any, List

# アプリケーション識別のためのサービス名
SERVICE_NAME = "NotionOrderApp"

class AccountDialog(tk.Toplevel):
    """アカウントの追加・編集を行うダイアログ"""
    def __init__(self, master, account_info=None, account_key=None, is_new=True):
        super().__init__(master)
        self.transient(master)
        self.grab_set()

        self.result = None
        self.is_new = is_new
        self.original_sender = account_info.get("sender", "") if account_info else ""

        self.title("アカウントの追加" if is_new else "アカウントの編集")

        # --- 変数 ---
        self.key_var = tk.StringVar(value=account_key if account_key else "")
        self.display_name_var = tk.StringVar()
        self.sender_var = tk.StringVar()
        self.password_var = tk.StringVar()

        if account_info:
            self.display_name_var.set(account_info.get("display_name", ""))
            sender_email = account_info.get("sender", "")
            self.sender_var.set(sender_email)
            # keyringからパスワードを取得
            if sender_email:
                password = keyring.get_password(SERVICE_NAME, sender_email)
                if password:
                    self.password_var.set(password)

        # --- UI ---
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="一意のキー:").grid(row=0, column=0, sticky=tk.W, pady=5)
        key_entry = ttk.Entry(frame, textvariable=self.key_var, width=40)
        key_entry.grid(row=0, column=1, sticky="ew", pady=5)
        if not is_new:
            key_entry.config(state="readonly")

        ttk.Label(frame, text="表示名:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.display_name_var).grid(row=1, column=1, sticky="ew", pady=5)

        ttk.Label(frame, text="メールアドレス:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.sender_var).grid(row=2, column=1, sticky="ew", pady=5)

        ttk.Label(frame, text="パスワード:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.password_var, show="*").grid(row=3, column=1, sticky="ew", pady=5)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, sticky="e", pady=(15, 0))
        ttk.Button(button_frame, text="OK", command=self.ok_pressed).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.destroy).pack(side=tk.LEFT)

    def ok_pressed(self):
        key = self.key_var.get().strip()
        sender_email = self.sender_var.get().strip()
        password = self.password_var.get()

        if not key or not sender_email:
            messagebox.showerror("入力エラー", "一意のキーとメールアドレスは必須です。", parent=self)
            return

        # 編集時かつメールアドレスが変更された場合、古いキーリングエントリを削除
        if not self.is_new and self.original_sender and self.original_sender != sender_email:
            try:
                keyring.delete_password(SERVICE_NAME, self.original_sender)
            except keyring.errors.PasswordDeleteError:
                # エントリが存在しない場合は何もしない
                pass

        # keyringにパスワードを保存
        keyring.set_password(SERVICE_NAME, sender_email, password)

        self.result = {
            "key": key,
            "details": {
                "display_name": self.display_name_var.get(),
                "sender": sender_email,
                # "password" は結果に含めない
            }
        }
        self.destroy()

class DepartmentDialog(tk.Toplevel):
    """部署の追加・編集を行うダイアログ"""
    def __init__(self, master, department_name=None, guidance_number=None, is_new=True):
        super().__init__(master)
        self.transient(master)
        self.grab_set()

        self.result = None

        self.title("部署の追加" if is_new else "部署の編集")

        # --- 変数 ---
        self.department_var = tk.StringVar(value=department_name if department_name else "")
        self.guidance_var = tk.StringVar(value=guidance_number if guidance_number else "")

        # --- UI ---
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="部署名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.department_var, width=40).grid(row=0, column=1, sticky="ew", pady=5)

        ttk.Label(frame, text="ガイダンス番号:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.guidance_var, width=40).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(frame, text="例: 内線101", font=("", 8), foreground="gray").grid(row=2, column=1, sticky=tk.W, pady=(0, 10))

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=2, sticky="e", pady=(15, 0))
        ttk.Button(button_frame, text="OK", command=self.ok_pressed).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.destroy).pack(side=tk.LEFT)

    def ok_pressed(self):
        department = self.department_var.get().strip()
        if not department:
            messagebox.showerror("入力エラー", "部署名は必須です。", parent=self)
            return
        
        self.result = {
            "department": department,
            "guidance_number": self.guidance_var.get().strip()
        }
        self.destroy()

class SettingsWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("設定")
        self.geometry("800x500")
        self.resizable(False, False)

        self.grab_set()
        self.focus_set()
        self.transient(master)

        self.accounts_data = {}
        self.department_defaults_vars = {}
        self.department_guidance_vars = {}
        self.save_result = {"saved": False, "message": "設定変更をキャンセルしました。"}
        self.original_settings = config._load_settings_from_json()

        main_frame = ttk.Frame(self, padding=(10, 10, 10, 0))
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        self.tab_accounts = ttk.Frame(notebook, padding=10)
        self.tab_departments = ttk.Frame(notebook, padding=10)

        notebook.add(self.tab_accounts, text="メールアカウント")
        notebook.add(self.tab_departments, text="部署設定")

        self.setup_accounts_tab()
        self.setup_departments_tab()

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, sticky="se", pady=(10, 0))
        ttk.Button(button_frame, text="保存して閉じる", command=self.save_and_close).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.close_without_save).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self.close_without_save)

        self.load_settings()

    def setup_accounts_tab(self):
        frame = self.tab_accounts
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        columns = ("display_name", "sender")
        self.accounts_tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        self.accounts_tree.grid(row=0, column=0, sticky="nsew")
        self.accounts_tree.heading("display_name", text="表示名")
        self.accounts_tree.heading("sender", text="メールアドレス")
        self.accounts_tree.column("display_name", width=150)
        self.accounts_tree.column("sender", width=300)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.accounts_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.accounts_tree.configure(yscrollcommand=vsb.set)

        button_container = ttk.Frame(frame)
        button_container.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(button_container, text="追加...", command=self.add_account).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="編集...", command=self.edit_account).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="削除", command=self.delete_account).pack(side=tk.LEFT)

    def setup_departments_tab(self):
        frame = self.tab_departments
        frame.columnconfigure(1, weight=1)

        # 部署リストフレーム
        list_frame = ttk.LabelFrame(frame, text="部署名リスト", padding=10)
        list_frame.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        list_frame.rowconfigure(0, weight=1)

        self.departments_listbox = tk.Listbox(list_frame, height=8)
        self.departments_listbox.grid(row=0, column=0, columnspan=2, sticky="nsew")

        ttk.Button(list_frame, text="追加", command=self.add_department).grid(row=1, column=0, pady=(5,0), sticky="ew")
        ttk.Button(list_frame, text="編集", command=self.edit_department).grid(row=2, column=0, pady=(2,0), sticky="ew")
        ttk.Button(list_frame, text="削除", command=self.delete_department).grid(row=1, column=1, pady=(5,0), sticky="ew")

        # 右側の設定フレーム
        right_frame = ttk.Frame(frame)
        right_frame.grid(row=0, column=1, sticky="nsew")
        right_frame.rowconfigure(1, weight=1)

        # デフォルト送信者設定
        defaults_frame = ttk.LabelFrame(right_frame, text="部署ごとのデフォルト送信者", padding=10)
        defaults_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.defaults_frame_content = ttk.Frame(defaults_frame)
        self.defaults_frame_content.pack(fill="both", expand=True)

        # ガイダンス番号設定
        guidance_frame = ttk.LabelFrame(right_frame, text="部署ごとのガイダンス番号", padding=10)
        guidance_frame.grid(row=1, column=0, sticky="nsew")
        self.guidance_frame_content = ttk.Frame(guidance_frame)
        self.guidance_frame_content.pack(fill="both", expand=True)

    def add_account(self):
        dialog = AccountDialog(self, is_new=True)
        self.wait_window(dialog)
        if dialog.result:
            key = dialog.result["key"]
            if key in self.accounts_data:
                messagebox.showerror("エラー", f"キー '{key}' は既に使用されています。", parent=self)
            else:
                self.accounts_data[key] = dialog.result["details"]
                self.refresh_accounts_tree()
                self.refresh_defaults_ui()
                self.refresh_guidance_ui()

    def edit_account(self):
        selected_key = self.accounts_tree.focus()
        if not selected_key:
            messagebox.showwarning("選択エラー", "編集するアカウントを選択してください。", parent=self)
            return

        account_info = self.accounts_data.get(selected_key)
        dialog = AccountDialog(self, account_info=account_info, account_key=selected_key, is_new=False)
        self.wait_window(dialog)

        if dialog.result:
            self.accounts_data[selected_key] = dialog.result["details"]
            self.refresh_accounts_tree()
            self.refresh_defaults_ui()
            self.refresh_guidance_ui()

    def delete_account(self):
        selected_key = self.accounts_tree.focus()
        if not selected_key:
            messagebox.showwarning("選択エラー", "削除するアカウントを選択してください。", parent=self)
            return
        
        account_info = self.accounts_data.get(selected_key, {})
        display_name = account_info.get("display_name", selected_key)
        sender_email = account_info.get("sender")

        if messagebox.askyesno("削除確認", f"アカウント「{display_name}」を削除しますか？\nこのアカウントを使用している部署のデフォルト設定は解除されます。", parent=self):
            # keyringからパスワードを削除
            if sender_email:
                try:
                    keyring.delete_password(SERVICE_NAME, sender_email)
                except keyring.errors.PasswordDeleteError:
                    # エントリが存在しない場合は何もしない
                    pass
            
            # データからアカウントを削除
            if selected_key in self.accounts_data:
                del self.accounts_data[selected_key]
            
            self.refresh_accounts_tree()
            self.refresh_defaults_ui()
            self.refresh_guidance_ui()

    def refresh_accounts_tree(self):
        self.accounts_tree.delete(*self.accounts_tree.get_children())
        for key, details in self.accounts_data.items():
            self.accounts_tree.insert("", tk.END, iid=key, values=(details.get("display_name", ""), details.get("sender", "")))

    def add_department(self):
        dialog = DepartmentDialog(self, is_new=True)
        self.wait_window(dialog)
        if dialog.result:
            department = dialog.result["department"]
            current_depts = self.departments_listbox.get(0, tk.END)
            if department in current_depts:
                messagebox.showwarning("重複", "その部署名は既に使用されています。", parent=self)
            else:
                self.departments_listbox.insert(tk.END, department)
                self.refresh_defaults_ui()
                self.refresh_guidance_ui()

    def edit_department(self):
        selected_indices = self.departments_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("選択エラー", "編集する部署を選択してください。", parent=self)
            return
        
        dept_name = self.departments_listbox.get(selected_indices[0])
        current_guidance = self.department_guidance_vars.get(dept_name, tk.StringVar()).get()
        
        dialog = DepartmentDialog(self, department_name=dept_name, guidance_number=current_guidance, is_new=False)
        self.wait_window(dialog)
        
        if dialog.result:
            new_dept = dialog.result["department"]
            new_guidance = dialog.result["guidance_number"]
            
            # 部署名が変更された場合
            if new_dept != dept_name:
                current_depts = self.departments_listbox.get(0, tk.END)
                if new_dept in current_depts:
                    messagebox.showwarning("重複", "その部署名は既に使用されています。", parent=self)
                    return
                # リストボックスの項目を更新
                self.departments_listbox.delete(selected_indices[0])
                self.departments_listbox.insert(selected_indices[0], new_dept)
            
            # ガイダンス番号を更新
            if dept_name in self.department_guidance_vars:
                self.department_guidance_vars[dept_name].set(new_guidance)
            
            self.refresh_defaults_ui()
            self.refresh_guidance_ui()

    def delete_department(self):
        selected_indices = self.departments_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("選択エラー", "削除する部署を選択してください。", parent=self)
            return
        
        dept_name = self.departments_listbox.get(selected_indices[0])
        if messagebox.askyesno("削除確認", f"部署「{dept_name}」を削除しますか？", parent=self):
            self.departments_listbox.delete(selected_indices[0])
            self.refresh_defaults_ui()
            self.refresh_guidance_ui()

    def refresh_defaults_ui(self):
        for widget in self.defaults_frame_content.winfo_children():
            widget.destroy()

        self.department_defaults_vars = {}
        account_names = [details.get("display_name") for details in self.accounts_data.values()]

        departments = self.departments_listbox.get(0, tk.END)
        current_defaults = config._load_settings_from_json().get("department_defaults", {})
        display_name_map = {v.get("display_name"): k for k, v in self.accounts_data.items()}

        for i, dept in enumerate(departments):
            ttk.Label(self.defaults_frame_content, text=f"{dept}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=5)
            combo = ttk.Combobox(self.defaults_frame_content, values=account_names, state="readonly")
            combo.grid(row=i, column=1, sticky="ew", padx=5)
            self.department_defaults_vars[dept] = combo
            
            default_key = current_defaults.get(dept)
            if default_key and default_key in self.accounts_data:
                display_name = self.accounts_data[default_key].get("display_name")
                combo.set(display_name)

    def refresh_guidance_ui(self):
        for widget in self.guidance_frame_content.winfo_children():
            widget.destroy()

        self.department_guidance_vars = {}
        departments = self.departments_listbox.get(0, tk.END)
        current_guidance = config._load_settings_from_json().get("department_guidance_numbers", {})

        for i, dept in enumerate(departments):
            ttk.Label(self.guidance_frame_content, text=f"{dept}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=5)
            entry = ttk.Entry(self.guidance_frame_content, width=20)
            entry.grid(row=i, column=1, sticky="ew", padx=5)
            self.department_guidance_vars[dept] = entry
            
            guidance_number = current_guidance.get(dept, "")
            entry.insert(0, guidance_number)

    def load_settings(self):
        self.accounts_data = config.load_email_accounts()
        self.refresh_accounts_tree()

        self.departments_listbox.delete(0, tk.END)
        departments = config.load_departments()
        if departments:
            for dept in departments:
                self.departments_listbox.insert(tk.END, dept)
        self.refresh_defaults_ui()
        self.refresh_guidance_ui()

    def save_and_close(self):
        current_json_data = config._load_settings_from_json()
        current_json_data["accounts"] = self.accounts_data
        current_json_data["departments"] = list(self.departments_listbox.get(0, tk.END))

        # デフォルト送信者設定を保存
        new_defaults = {}
        key_map = {v.get("display_name"): k for k, v in self.accounts_data.items()}
        for dept, combo in self.department_defaults_vars.items():
            display_name = combo.get()
            if display_name in key_map:
                new_defaults[dept] = key_map[display_name]
        current_json_data["department_defaults"] = new_defaults

        # ガイダンス番号設定を保存
        new_guidance = {}
        for dept, entry in self.department_guidance_vars.items():
            guidance_number = entry.get().strip()
            if guidance_number:
                new_guidance[dept] = guidance_number
        current_json_data["department_guidance_numbers"] = new_guidance

        success, message = config.save_settings(current_json_data)
        if success:
            summary = self._build_change_summary(self.original_settings, current_json_data)
            if summary:
                full_message = f"{message} 変更内容: {summary}"
            else:
                full_message = message
            self.save_result = {"saved": True, "message": full_message}
            messagebox.showinfo("成功", full_message, parent=self)
            self.destroy()
        else:
            messagebox.showerror("エラー", message, parent=self)

    def close_without_save(self):
        self.save_result = {"saved": False, "message": "設定変更をキャンセルしました。"}
        self.destroy()

    def _build_change_summary(self, before: dict, after: dict) -> str:
        parts: List[str] = []
        self._append_diff(parts, "メールアカウント", before.get("accounts", {}), after.get("accounts", {}))
        self._append_diff(parts, "部署", before.get("departments", []), after.get("departments", []))
        self._append_diff(parts, "デフォルト送信者", before.get("department_defaults", {}), after.get("department_defaults", {}))
        self._append_diff(parts, "ガイダンス番号", before.get("department_guidance_numbers", {}), after.get("department_guidance_numbers", {}))
        return "、".join(parts) if parts else ""

    def _append_diff(self, target: List[str], label: str, before: Any, after: Any):
        diff = self._extract_diff(before, after)
        if diff:
            target.append(f"{label}: {len(diff)}項目変更")

    def _extract_diff(self, before: Any, after: Any) -> List[Any]:
        if isinstance(before, dict) and isinstance(after, dict):
            keys = set(before.keys()) | set(after.keys())
            return [key for key in keys if before.get(key) != after.get(key)]
        if isinstance(before, (list, tuple)) and isinstance(after, (list, tuple)):
            if set(before) == set(after):
                return []
            return list(set(after).symmetric_difference(set(before)))
        if before != after:
            return [after]
        return []

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    app = SettingsWindow(root)
    app.protocol("WM_DELETE_WINDOW", lambda: root.destroy())
    root.mainloop()
