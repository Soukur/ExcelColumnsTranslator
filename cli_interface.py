"""
対話型CLIインターフェースモジュール
"""
import os
import argparse
from typing import Dict, List, Tuple, Union, TypedDict, Optional

class TranslationParams(TypedDict):
    batch_mode: bool
    input_path: str
    output_path: str
    source_cols: List[int]
    target_cols: List[int]
    row_range: Tuple[int, int]
    api_key: Optional[str]  # Added API key parameter

class CLIInterface:
    def __init__(self):
        pass

    def _get_input_path(self) -> str:
        """入力ファイルパスの取得"""
        while True:
            path = input("翻訳するExcelファイルのパスを入力してください: ").strip().strip('"')
            if os.path.exists(path):
                return path
            print("ファイルが見つかりません。正しいパスを入力してください。")

    def _get_output_path(self, input_path: str) -> str:
        """出力ファイルパスの取得"""
        default_path = self._generate_default_output_path(input_path)
        path = input(f"出力ファイルのパスを入力してください（デフォルト: {default_path}）: ").strip().strip('"')
        return path if path else default_path

    def _generate_default_output_path(self, input_path: str) -> str:
        """デフォルトの出力パスを生成"""
        dir_name = os.path.dirname(input_path)
        base_name = os.path.basename(input_path)
        name, ext = os.path.splitext(base_name)
        return os.path.join(dir_name, f"{name}_translated{ext}")

    def _parse_column_input(self, prompt: Optional[str] = None) -> List[int]:
        """列指定の解析（アルファベット形式）"""
        while True:
            try:
                if prompt:
                    col_input = input(prompt).strip().upper()
                else:
                    return []

                cols = []
                for part in col_input.split(','):
                    if '-' in part:
                        start, end = part.split('-')
                        start_num = excel_column_to_number(start.strip())
                        end_num = excel_column_to_number(end.strip())
                        cols.extend(range(start_num, end_num + 1))
                    else:
                        cols.append(excel_column_to_number(part.strip()))
                return cols
            except ValueError:
                print("無効な入力です。列名または範囲（例: A,B,C-E）で入力してください。")

    def _get_row_range(self) -> Tuple[int, int]:
        """行範囲の取得"""
        while True:
            try:
                start = int(input("開始行を入力してください: ").strip())
                end = int(input("終了行を入力してください: ").strip())
                if start <= end:
                    return (start, end)
                print("開始行は終了行以下である必要があります。")
            except ValueError:
                print("無効な入力です。数値で入力してください。")

    def parse_args(self) -> TranslationParams:
        """コマンドライン引数の解析"""
        parser = argparse.ArgumentParser(description='Excel翻訳ツール')
        parser.add_argument('--batch', action='store_true', help='バッチモードで実行')
        parser.add_argument('--input', help='入力Excelファイルのパス')
        parser.add_argument('--output', help='出力Excelファイルのパス')
        parser.add_argument('--source-cols', help='翻訳元の列（例: A,B,C-E）')
        parser.add_argument('--target-cols', help='翻訳先の列（例: F,G,H-J）')
        parser.add_argument('--row-start', type=int, help='開始行番号')
        parser.add_argument('--row-end', type=int, help='終了行番号')
        parser.add_argument('--api-key', help='DeepL APIキー（指定しない場合は環境変数DEEPL_API_KEYを使用）')

        args = parser.parse_args()

        if args.batch:
            if not all([args.input, args.source_cols, args.target_cols, 
                       args.row_start is not None, args.row_end is not None]):
                parser.error("バッチモードでは全てのパラメータが必要です")

            if not os.path.exists(args.input):
                parser.error(f"入力ファイルが見つかりません: {args.input}")

            # 列の解析
            try:
                source_cols = []
                for part in args.source_cols.split(','):
                    if '-' in part:
                        start, end = part.split('-')
                        start_num = excel_column_to_number(start.strip())
                        end_num = excel_column_to_number(end.strip())
                        source_cols.extend(range(start_num, end_num + 1))
                    else:
                        source_cols.append(excel_column_to_number(part.strip()))

                target_cols = []
                for part in args.target_cols.split(','):
                    if '-' in part:
                        start, end = part.split('-')
                        start_num = excel_column_to_number(start.strip())
                        end_num = excel_column_to_number(end.strip())
                        target_cols.extend(range(start_num, end_num + 1))
                    else:
                        target_cols.append(excel_column_to_number(part.strip()))
            except ValueError as e:
                parser.error(f"列の指定が無効です: {str(e)}")

            if len(source_cols) != len(target_cols):
                parser.error("翻訳元と翻訳先の列数が一致しません")

            if args.row_start > args.row_end:
                parser.error("開始行は終了行以下である必要があります")

            return {
                'batch_mode': True,
                'input_path': args.input,
                'output_path': args.output or self._generate_default_output_path(args.input),
                'source_cols': source_cols,
                'target_cols': target_cols,
                'row_range': (args.row_start, args.row_end),
                'api_key': args.api_key
            }
        return {
            'batch_mode': False,
            'input_path': "",
            'output_path': "",
            'source_cols': [],
            'target_cols': [],
            'row_range': (0, 0),
            'api_key': args.api_key
        }

    def get_parameters(self) -> TranslationParams:
        """全パラメータの取得"""
        # バッチモードのチェック
        args = self.parse_args()
        if args.get('batch_mode'):
            return args

        # 対話モード
        input_path = self._get_input_path()
        source_cols = self._parse_column_input(
            "翻訳元の列を入力してください（カンマ区切り、範囲指定可能 例: A,B,C-E）: "
        )
        target_cols = self._parse_column_input(
            "翻訳先の列を入力してください（カンマ区切り、範囲指定可能 例: F,G,H-J）: "
        )

        # 翻訳元と翻訳先の列数チェック
        if len(source_cols) != len(target_cols):
            raise ValueError("翻訳元と翻訳先の列数が一致しません。")

        row_range = self._get_row_range()
        return {
            'batch_mode': False,
            'input_path': input_path,
            'output_path': self._get_output_path(input_path),
            'source_cols': source_cols,
            'target_cols': target_cols,
            'row_range': row_range,
            'api_key': args.get('api_key')
        }

from utils import excel_column_to_number, number_to_excel_column