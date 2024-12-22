from openpyxl import Workbook

def create_test_file():
    wb = Workbook()
    ws = wb.active
    
    # テストデータの設定
    test_data = [
        "Hello, this is a test message.",
        "The weather is nice today.",
        "I love programming."
    ]
    
    # A列にテストデータを設定
    for i, text in enumerate(test_data, 1):
        ws.cell(row=i, column=1, value=text)
    
    # ファイルの保存
    wb.save('test.xlsx')
    print("テスト用Excelファイルを作成しました: test.xlsx")

if __name__ == "__main__":
    create_test_file()
