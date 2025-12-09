"""標準出力をキューにリダイレクトするためのIOラッパー"""
import queue
from typing import Any


class QueueIO:
    """標準出力をキューにリダイレクトするためのIOラッパー"""
    def __init__(self, q: queue.Queue) -> None:
        self.q = q
    
    def write(self, text: str) -> None:
        self.q.put(("log", text))
    
    def flush(self) -> None:
        pass

