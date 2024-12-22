"""
DeepL APIクライアントモジュール
"""
import os
import requests
from time import sleep
from typing import Optional

class DeepLTranslator:
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv('DEEPL_API_KEY')
        if not self.api_key:
            raise ValueError("DeepL APIキーが指定されていません。コマンドライン引数 --api-key または環境変数 DEEPL_API_KEY で指定してください。")

        self.base_url = "https://api-free.deepl.com/v2/translate"
        self.headers = {
            "Authorization": f"DeepL-Auth-Key {self.api_key}",
            "Content-Type": "application/json"
        }

    def test_connection(self) -> bool:
        """API接続のテスト"""
        try:
            # 短いテキストで翻訳をテスト
            result = self.translate("Hello")
            return bool(result)
        except Exception as e:
            print(f"API接続テストに失敗しました: {str(e)}")
            return False

    def translate(self, text: str, target_lang: str = "JA", max_retries: int = 3) -> Optional[str]:
        """テキストを翻訳"""
        if not text:
            return ""

        for attempt in range(max_retries):
            try:
                response = requests.post(
                    self.base_url,
                    headers=self.headers,
                    json={
                        "text": [text],
                        "target_lang": target_lang
                    }
                )

                if response.status_code == 429:  # Rate limit
                    wait_time = 2 ** attempt
                    print(f"API制限に達しました。{wait_time}秒待機します...")
                    sleep(wait_time)
                    continue
                elif response.status_code == 403:
                    raise Exception("APIキーが無効です")
                elif response.status_code == 456:
                    raise Exception("文字制限を超えています")

                response.raise_for_status()
                return response.json()["translations"][0]["text"]

            except requests.exceptions.RequestException as e:
                if attempt == max_retries - 1:
                    raise Exception(f"翻訳APIでエラーが発生しました: {str(e)}")
                wait_time = 2 ** attempt
                print(f"エラーが発生しました。{wait_time}秒後にリトライします...")
                sleep(wait_time)

        return None