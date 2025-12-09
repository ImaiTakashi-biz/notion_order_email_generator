"""
PNG画像をICO形式に変換するスクリプト
"""
from PIL import Image
import os

def convert_png_to_ico(png_path: str, ico_path: str) -> None:
    """
    PNG画像をICO形式に変換する
    
    Args:
        png_path: 入力PNGファイルのパス
        ico_path: 出力ICOファイルのパス
    """
    if not os.path.exists(png_path):
        raise FileNotFoundError(f"PNGファイルが見つかりません: {png_path}")
    
    # PNG画像を読み込む
    img = Image.open(png_path)
    
    # 複数のサイズを含むICOファイルとして保存
    # Windowsアイコンは通常、16x16, 32x32, 48x48, 256x256を含む
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(ico_path, format='ICO', sizes=sizes)
    print(f"ICOファイルを生成しました: {ico_path}")

if __name__ == '__main__':
    convert_png_to_ico('mail_app_image.png', 'app_icon.ico')

