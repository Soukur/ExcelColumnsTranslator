"""
翻訳履歴を管理するモジュール
"""
import json
import os
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Optional

class TranslationHistory:
    def __init__(self, history_file: str = "translation_history.json"):
        """
        翻訳履歴管理クラスの初期化

        Args:
            history_file (str): 履歴を保存するJSONファイルのパス
        """
        self.history_file = history_file
        self.history: List[Dict] = self._load_history()

    def _load_history(self) -> List[Dict]:
        """既存の履歴を読み込む"""
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                print(f"警告: 履歴ファイルの読み込みに失敗しました。新しい履歴を作成します。")
                return []
        return []

    def add_entry(self, 
                  source_text: str, 
                  translated_text: str, 
                  excel_file: str,
                  sheet_name: str,
                  source_cell: str,
                  target_cell: str) -> None:
        """
        翻訳エントリーを追加

        Args:
            source_text (str): 原文
            translated_text (str): 翻訳文
            excel_file (str): Excelファイル名
            sheet_name (str): シート名
            source_cell (str): 翻訳元セル
            target_cell (str): 翻訳先セル
        """
        # 日本のタイムゾーン（UTC+9）を設定
        JST = timezone(timedelta(hours=+9))

        entry = {
            'timestamp': datetime.now(JST).isoformat(),
            'excel_file': excel_file,
            'sheet_name': sheet_name,
            'source_cell': source_cell,
            'target_cell': target_cell,
            'source_text': source_text,
            'translated_text': translated_text
        }
        self.history.append(entry)
        self._save_history()

    def _save_history(self) -> None:
        """履歴をファイルに保存"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"警告: 履歴の保存に失敗しました: {str(e)}")

    def get_recent_entries(self, limit: Optional[int] = None) -> List[Dict]:
        """
        最近の翻訳履歴を取得

        Args:
            limit (Optional[int]): 取得するエントリー数の上限

        Returns:
            List[Dict]: 翻訳履歴のリスト
        """
        if limit is None:
            return self.history
        return self.history[-limit:]

    def clear_history(self) -> None:
        """履歴を全て削除"""
        self.history = []
        if os.path.exists(self.history_file):
            try:
                os.remove(self.history_file)
            except Exception as e:
                print(f"警告: 履歴ファイルの削除に失敗しました: {str(e)}")