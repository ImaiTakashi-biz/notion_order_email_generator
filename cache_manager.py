"""
キャッシュ管理モジュール
Notionデータ取得の結果をキャッシュして、同じ条件での再取得を高速化する
"""
import hashlib
import json
import time
from typing import Any, Dict, List, Optional
import logger_config

logger = logger_config.get_logger(__name__)

# キャッシュの有効期限（秒）
CACHE_TTL = 300  # 5分

# メモリキャッシュ
_cache: Dict[str, Dict[str, Any]] = {}


def _generate_cache_key(department_names: Optional[List[str]]) -> str:
    """
    部署名リストからキャッシュキーを生成する
    
    Args:
        department_names: 部署名のリスト
    
    Returns:
        キャッシュキー（ハッシュ値）
    """
    if department_names is None:
        department_names = []
    
    # ソートして一意性を保証
    sorted_departments = sorted(department_names)
    key_data = json.dumps(sorted_departments, ensure_ascii=False)
    return hashlib.md5(key_data.encode('utf-8')).hexdigest()


def get_cached_data(department_names: Optional[List[str]]) -> Optional[Dict[str, Any]]:
    """
    キャッシュからデータを取得する
    
    Args:
        department_names: 部署名のリスト
    
    Returns:
        キャッシュされたデータ、またはNone（キャッシュなし/期限切れ）
    """
    cache_key = _generate_cache_key(department_names)
    
    if cache_key not in _cache:
        return None
    
    cached_item = _cache[cache_key]
    current_time = time.time()
    
    # キャッシュの有効期限チェック
    if current_time - cached_item['timestamp'] > CACHE_TTL:
        logger.debug(f"キャッシュが期限切れです: {cache_key}")
        del _cache[cache_key]
        return None
    
    logger.debug(f"キャッシュからデータを取得: {cache_key}")
    return cached_item['data']


def set_cached_data(department_names: Optional[List[str]], data: Dict[str, Any]) -> None:
    """
    データをキャッシュに保存する
    
    Args:
        department_names: 部署名のリスト
        data: キャッシュするデータ
    """
    cache_key = _generate_cache_key(department_names)
    _cache[cache_key] = {
        'data': data,
        'timestamp': time.time()
    }
    logger.debug(f"データをキャッシュに保存: {cache_key}")


def clear_cache() -> None:
    """
    キャッシュをクリアする
    """
    _cache.clear()
    logger.info("キャッシュをクリアしました")


def get_cache_stats() -> Dict[str, Any]:
    """
    キャッシュの統計情報を取得する
    
    Returns:
        キャッシュの統計情報
    """
    current_time = time.time()
    valid_count = 0
    expired_count = 0
    
    for item in _cache.values():
        if current_time - item['timestamp'] <= CACHE_TTL:
            valid_count += 1
        else:
            expired_count += 1
    
    return {
        'total': len(_cache),
        'valid': valid_count,
        'expired': expired_count
    }

