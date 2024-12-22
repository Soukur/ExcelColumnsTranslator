"""
ユーティリティ関数モジュール
"""
import time
from typing import Optional, Tuple

def validate_column_range(columns, max_col):
    """列範囲の妥当性を検証"""
    for col in columns:
        if col < 1 or col > max_col:
            raise ValueError(f"無効な列番号です: {col}")

def format_progress_bar(progress: float, 
                       width: int = 50, 
                       current_cell: Optional[str] = None,
                       start_time: Optional[float] = None,
                       total_cells: Optional[int] = None,
                       processed_cells: Optional[int] = None) -> str:
    """
    詳細なプログレスバーの生成

    Args:
        progress (float): 進捗率（0-100）
        width (int): プログレスバーの幅
        current_cell (Optional[str]): 現在処理中のセル位置
        start_time (Optional[float]): 処理開始時刻
        total_cells (Optional[int]): 総処理セル数
        processed_cells (Optional[int]): 処理済みセル数

    Returns:
        str: フォーマットされたプログレスバー文字列
    """
    filled = int(width * progress // 100)
    bar = '=' * filled + '>' + '-' * (width - filled - 1)
    status = f"[{bar}] {progress:.1f}%"

    if current_cell:
        status += f" | 現在の位置: {current_cell}"

    if all(x is not None for x in [start_time, total_cells, processed_cells]):
        elapsed = time.time() - start_time
        cells_per_sec = processed_cells / elapsed if elapsed > 0 else 0
        remaining_cells = total_cells - processed_cells
        eta = remaining_cells / cells_per_sec if cells_per_sec > 0 else 0

        status += f" | 速度: {cells_per_sec:.1f}セル/秒"
        status += f" | 残り時間: {format_time(eta)}"

    return status

def format_time(seconds: float) -> str:
    """
    秒数を時間形式に変換

    Args:
        seconds (float): 秒数

    Returns:
        str: フォーマットされた時間文字列
    """
    if seconds < 60:
        return f"{seconds:.0f}秒"
    elif seconds < 3600:
        minutes = seconds / 60
        return f"{minutes:.0f}分"
    else:
        hours = seconds / 3600
        return f"{hours:.1f}時間"

def excel_column_to_number(column_letter):
    """Excel列文字を数値に変換"""
    result = 0
    for char in column_letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result

def number_to_excel_column(column_number):
    """数値をExcel列文字に変換"""
    result = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result