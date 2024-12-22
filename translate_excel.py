#!/usr/bin/env python3
"""
Excel翻訳ツールのメインスクリプト
"""
import sys
import os
import time
from cli_interface import CLIInterface
from excel_handler import ExcelHandler
from deepl_client import DeepLTranslator
from translation_history import TranslationHistory
from utils import format_progress_bar

def main() -> None:
    try:
        # CLIインターフェースの初期化
        cli = CLIInterface()
        params = cli.get_parameters()

        # バッチモードの場合は進捗表示を簡略化
        is_batch_mode = params['batch_mode']

        # Excelハンドラーの初期化
        excel_handler = ExcelHandler(
            input_path=str(params['input_path']),
            output_path=str(params['output_path'])
        )

        # DeepL翻訳クライアントの初期化（APIキーをパラメータから取得）
        translator = DeepLTranslator(api_key=params.get('api_key'))

        # 翻訳履歴ハンドラーの初期化
        history_handler = TranslationHistory()

        # 翻訳処理の実行
        row_start, row_end = params['row_range']
        total_rows = row_end - row_start + 1
        total_cols = len(params['source_cols'])
        processed = 0
        total_cells = total_rows * total_cols
        start_time = time.time()

        for src_col, dest_col in zip(params['source_cols'], params['target_cols']):
            for row in range(row_start, row_end + 1):
                # 進捗表示の更新
                processed += 1
                progress = (processed / total_cells) * 100
                current_cell = excel_handler.get_cell_address(row, src_col)

                if is_batch_mode:
                    # バッチモードでは簡略化された進捗表示
                    print(f"\r進捗: {progress:.1f}% | セル: {current_cell}", end='', file=sys.stderr)
                else:
                    # 対話モードでは詳細な進捗表示
                    progress_bar = format_progress_bar(
                        progress=progress,
                        current_cell=current_cell,
                        start_time=start_time,
                        total_cells=total_cells,
                        processed_cells=processed
                    )
                    print(f"\r{progress_bar}", end='', file=sys.stderr)

                # セルの翻訳
                source_text = excel_handler.get_cell_value(row, src_col)
                if source_text:
                    translated = translator.translate(source_text)
                    if translated is not None:
                        # 翻訳結果をセルに設定
                        excel_handler.set_cell_value(row, dest_col, translated)

                        # 翻訳履歴に追加
                        history_handler.add_entry(
                            source_text=source_text,
                            translated_text=translated,
                            excel_file=os.path.basename(params['input_path']),
                            sheet_name=excel_handler.get_sheet_name(),
                            source_cell=excel_handler.get_cell_address(row, src_col),
                            target_cell=excel_handler.get_cell_address(row, dest_col)
                        )

        # 保存
        excel_handler.save()
        print("\n翻訳が完了しました！")

        # 実行時間の表示
        total_time = time.time() - start_time
        print(f"処理時間: {total_time:.1f}秒")

        if is_batch_mode:
            # バッチモードでは出力パスを表示
            print(f"出力ファイル: {params['output_path']}")

        # 翻訳履歴の保存先を表示
        print(f"翻訳履歴は {history_handler.history_file} に保存されました。")

    except KeyboardInterrupt:
        print("\n処理が中断されました。")
        sys.exit(1)
    except Exception as e:
        print(f"\nエラーが発生しました: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()