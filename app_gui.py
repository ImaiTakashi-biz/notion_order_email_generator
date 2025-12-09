"""
アプリケーションメインファイル（後方互換性のため）
分割されたモジュールからApplicationクラスをインポート
"""
# 後方互換性のため、Applicationクラスをエクスポート
from controllers.app_controller import Application

__all__ = ['Application']
