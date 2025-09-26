import tkinter as tk
from tkinter import messagebox
import sys
import config
from app_gui import Application

def run_pre_flight_checks():
    """アプリケーション起動前の設定検証"""
    is_valid, errors = config.validate_config()
    if not is_valid:
        # GUIを表示する前にエラーを表示するため、一時的なTkルートを作成
        root = tk.Tk()
        root.withdraw() # メインウィンドウは表示しない
        error_message = "アプリケーションを開始できませんでした。\n以下の設定を確認してください:\n\n" + "\n".join(f"・{e}" for e in errors)
        messagebox.showerror("設定エラー", error_message)
        root.destroy()
        return False
    return True

if __name__ == "__main__":
    # 起動前チェックを実行
    if not run_pre_flight_checks():
        sys.exit(1)

    # アプリケーションのメインウィンドウを作成
    root = tk.Tk()
    
    # アプリケーションのインスタンスを作成
    app = Application(master=root)
    
    # ウィンドウのタイトルを設定
    app.master.title("Notion注文書メール作成アプリ")
    
    # ウィンドウを最大化して表示
    root.state('zoomed')
    
    # メインループを開始
    app.mainloop()
