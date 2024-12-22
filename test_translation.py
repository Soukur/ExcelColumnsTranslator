#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse
from deepl_client import DeepLTranslator

def main():
    # コマンドライン引数の解析
    parser = argparse.ArgumentParser(description='DeepL API接続テスト')
    parser.add_argument('--api-key', help='DeepL APIキー（指定しない場合は環境変数DEEPL_API_KEYを使用）')
    args = parser.parse_args()

    try:
        translator = DeepLTranslator(api_key=args.api_key)

        print("DeepL API接続テスト中...")
        if translator.test_connection():
            print("API接続テスト成功！")

            # 実際の翻訳テスト
            test_text = "Hello, this is a test message."
            print(f"\n翻訳テスト:")
            print(f"原文: {test_text}")
            translated = translator.translate(test_text)
            print(f"訳文: {translated}")
        else:
            print("API接続テストに失敗しました。")

    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")

if __name__ == "__main__":
    main()