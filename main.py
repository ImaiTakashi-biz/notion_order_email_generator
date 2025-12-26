import tkinter as tk
from tkinter import messagebox
import os
import sys
import config
from controllers.app_controller import Application
from version import APP_NAME, APP_VERSION

def _get_resource_path(relative_path: str) -> str:
    """
    PyInstallerでビルドされた場合と通常実行の場合の両方に対応してリソースパスを取得する
    
    Args:
        relative_path: リソースファイルの相対パス
        
    Returns:
        リソースファイルの絶対パス
    """
    # PyInstallerでビルドされた場合、一時ディレクトリのパスがsys._MEIPASSに設定される
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        # 通常実行時はスクリプトのディレクトリを基準にする
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

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
    root.title(f"{APP_NAME}  {APP_VERSION}")
    
    # アイコンの設定（存在する場合）
    icon_path = _get_resource_path("app_icon.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception:
            # アイコンの読み込みに失敗してもアプリは継続
            pass
    
    root.state('zoomed')
    app = Application(master=root)
    app.mainloop()

if __name__ == '__main__':
    main()
