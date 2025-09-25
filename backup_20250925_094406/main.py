import tkinter as tk
from app_gui import Application

if __name__ == "__main__":
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
