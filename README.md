# Excel翻訳ツール

DeepL APIを使用したExcelファイル翻訳CLIツール。対話型インターフェースまたはバッチモードで操作可能で、列指定（A,B,C形式）、行範囲指定、出力パス指定が可能です。

## 特徴
- DeepL API連携による自動翻訳
- 対話型CLI操作とバッチ処理モード
- Excel列・行範囲の柔軟な指定
- カスタム出力パス設定
- APIエラーハンドリング
- 翻訳履歴のJSON形式での保存

## インストール方法

```bash
https://github.com/Soukur/ExcelColumnsTranslator.git
cd excel-translator
pip install -r requirements.txt
```

## APIキーの設定

以下の2つの方法でDeepL APIキーを設定できます：

1. コマンドライン引数での指定（推奨）
```bash
python translate_excel.py --api-key 'YOUR-API-KEY'
```

2. 環境変数での指定

### Windows環境での環境変数設定

#### システム環境変数の設定（永続的）
1. Windowsキー + Rを押して「実行」を開く
2. `sysdm.cpl` と入力してEnter
3. 「詳細設定」タブをクリック
4. 「環境変数」ボタンをクリック
5. 「システム環境変数」セクションで「新規」をクリック
6. 変数名: `DEEPL_API_KEY`
7. 変数値: あなたのDeepL APIキー
8. OKをクリックして全ての画面を閉じる
9. コマンドプロンプトを再起動

#### コマンドプロンプトでの一時的な設定
```cmd
set DEEPL_API_KEY=YOUR-API-KEY
python translate_excel.py
```

#### PowerShellでの設定
```powershell
$env:DEEPL_API_KEY = 'YOUR-API-KEY'
python translate_excel.py
```

### Unix/Linux/macOS環境での設定
```bash
export DEEPL_API_KEY='YOUR-API-KEY'
python translate_excel.py
```

## 実行モード

### 1. 対話モード
対話形式で必要な情報を入力しながら翻訳を実行します。
```bash
python translate_excel.py
```

### 2. バッチモード
コマンドライン引数で全てのパラメータを指定して実行します。
```bash
python translate_excel.py --batch --input example.xlsx --source-cols A,B --target-cols C,D --row-start 1 --row-end 10
```

## オプション一覧

| オプション | 説明 | 必須 | デフォルト値 |
|------------|------|------|--------------|
| --batch | バッチモードで実行 | × | False |
| --input | 入力Excelファイルのパス | バッチモード時○ | - |
| --output | 出力Excelファイルのパス | × | 入力ファイル名_translated |
| --source-cols | 翻訳元の列（例: A,B,C-E） | バッチモード時○ | - |
| --target-cols | 翻訳先の列（例: F,G,H-J） | バッチモード時○ | - |
| --row-start | 開始行番号 | バッチモード時○ | - |
| --row-end | 終了行番号 | バッチモード時○ | - |
| --api-key | DeepL APIキー | × | 環境変数から取得 |

## 注意点

1. APIキーについて
   - APIキーは必ず指定してください（コマンドライン引数または環境変数）
   - 無効なAPIキーを指定した場合はエラーメッセージが表示されます

2. 列指定について
   - 翻訳元と翻訳先の列数は一致させてください
   - 列指定は「A」や「B」などのExcel形式で指定します
   - 範囲指定も可能です（例：A-C）

3. 行範囲について
   - 開始行は終了行以下である必要があります
   - 行番号は1から始まります

4. ファイル処理について
   - 出力ファイルが既に存在する場合は上書きされます
   - 大きなファイルの場合、処理に時間がかかる場合があります

5. Windows環境特有の注意点
   - 環境変数を設定した後は、コマンドプロンプトを再起動してください
   - パスを指定する際は、バックスラッシュ（\）またはスラッシュ（/）が使用可能です
   - スペースを含むパスは引用符（""）で囲んでください

## 使用例

### 基本的な使用例（バッチモード）
```bash
# A列の内容をB列に翻訳（1行目から3行目まで）
python translate_excel.py --batch --input test.xlsx --source-cols A --target-cols B --row-start 1 --row-end 3 --api-key 'YOUR-API-KEY'
```

### 複数列の翻訳例
```bash
# A,B列の内容をC,D列に翻訳
python translate_excel.py --batch --input test.xlsx --source-cols A,B --target-cols C,D --row-start 1 --row-end 10 --api-key 'YOUR-API-KEY'
```

### 列範囲指定の例
```bash
# A-C列の内容をD-F列に翻訳
python translate_excel.py --batch --input test.xlsx --source-cols A-C --target-cols D-F --row-start 1 --row-end 5 --api-key 'YOUR-API-KEY'
```

## 翻訳履歴

翻訳履歴は `translation_history.json` に保存され、以下の情報が記録されます：
- 翻訳日時（JST）
- 原文と訳文
- Excelファイル名
- シート名
- 翻訳元・翻訳先セル

## エラー発生時の対応

1. APIエラー
   - レート制限：自動的にリトライします
   - 認証エラー：APIキーを確認してください
   - ネットワークエラー：接続を確認してリトライしてください

2. ファイルエラー
   - ファイルが見つからない：パスを確認してください
   - アクセス権限エラー：ファイルのアクセス権限を確認してください
   - 保存エラー：出力先のディスクの空き容量を確認してください

3. Windows環境でのエラー
   - 環境変数が認識されない：コマンドプロンプトを再起動してください
   - パス関連のエラー：パスの区切り文字とスペースの扱いを確認してください
   - 権限エラー：管理者権限で実行するか、アクセス権限を確認してください
