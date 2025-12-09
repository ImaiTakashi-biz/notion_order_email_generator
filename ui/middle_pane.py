"""中央のデータ表示領域 (仕入先リストと注文データ) のUI"""
import tkinter as tk
from tkinter import ttk
from typing import TYPE_CHECKING, Dict, Any, List

if TYPE_CHECKING:
    from controllers.app_controller import Application


class MiddlePane(ttk.PanedWindow):
    """中央のデータ表示領域 (仕入先リストと注文データ) のUI"""
    def __init__(self, master: ttk.Frame, app: 'Application') -> None:
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

    def update_supplier_list(self, data: List[Dict[str, Any]]) -> None:
        """仕入先リストを更新する"""
        self.supplier_listbox.delete(*self.supplier_listbox.get_children())
        suppliers = sorted(list(set(i["supplier_name"] for i in data)))
        for s in suppliers: self.supplier_listbox.insert('', tk.END, values=(s,))

    def update_table_for_supplier(self, supplier_name: str) -> None:
        """指定された仕入先の注文データをテーブルに表示する"""
        self.tree.delete(*self.tree.get_children())
        items_to_display = self.app.orders_by_supplier.get(supplier_name, [])
        for item in items_to_display:
            self.tree.insert("", tk.END, values=(item.get("maker_name", ""), item.get("db_part_number", ""), item.get("quantity", 0)))

    def mark_supplier_as_sent(self, supplier: str) -> None:
        """仕入先を送信済みとしてマークする"""
        for iid in self.supplier_listbox.get_children():
            if self.supplier_listbox.item(iid, 'values')[0] == supplier:
                self.supplier_listbox.item(iid, tags=('sent',))
                self.supplier_listbox.selection_remove(iid)
                break

    def clear_displays(self) -> None:
        """表示をクリアする"""
        self.tree.delete(*self.tree.get_children())
        self.supplier_listbox.delete(*self.supplier_listbox.get_children())

