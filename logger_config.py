"""
ロギング設定モジュール
アプリケーション全体で使用するロガーを設定する
"""
import logging

# ログフォーマット
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# ロガーのキャッシュ
_loggers: dict[str, logging.Logger] = {}


def get_logger(name: str, level: int = logging.INFO) -> logging.Logger:
    """
    指定された名前のロガーを取得または作成する
    
    Args:
        name: ロガー名（通常は__name__）
        level: ログレベル（デフォルト: INFO）
    
    Returns:
        設定済みのロガーインスタンス
    """
    if name in _loggers:
        return _loggers[name]
    
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # 既にハンドラーが設定されている場合はスキップ
    if logger.handlers:
        _loggers[name] = logger
        return logger
    
    # ハンドラーなし（ログはUIの「通知メッセージ」エリアにのみ表示）
    # ターミナルへの出力は行わない
    
    _loggers[name] = logger
    return logger

