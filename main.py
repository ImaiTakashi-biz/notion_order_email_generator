import tkinter as tk
from tkinter import messagebox
import config
from app_gui import Application

def main():
    """
    アプリケーションのメイン関数
    """
    # 起動前に設定を検証
    is_valid, errors = config.validate_config()

    if not is_valid:
        # エラーメッセージを整形
        error_title = "設定エラー"
        error_details = "\n".join(errors)
        full_message = f"アプリケーションを起動できませんでした。\n以下の項目を確認してください:\n\n{error_details}"
        
        # GUIを使わずにエラーダイアログを表示
        root = tk.Tk()
        root.withdraw() # メインウィンドウを非表示にする
        messagebox.showerror(error_title, full_message)
        root.destroy()
        return # アプリケーションを終了

    # 検証が成功した場合のみGUIを起動
    root = tk.Tk()
    root.title("Notion 注文書メール自動作成アプリ")
    root.state('zoomed')
    app = Application(master=root)
    app.mainloop()

if __name__ == '__main__':
    main()