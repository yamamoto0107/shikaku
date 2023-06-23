"""
要件定義
・前回作成した答え合わせソフトで、画面設計。
・openpyxlとファイル選択機能。
・フォルダでの一括添削機能。
・入力欄１：正解エクセルファイルの選択。
・入力欄２：回答エクセルファイルの選択。
・入力欄３：正誤を〇×形式か〇答え形式選択。
・入力欄４：添削ファイルの保存場所設定。(初期値は回答エクセルのアドレス)
・入力欄５：学生の回答問題数と回答開始行と列番号。
(初期値は、20,1,2)

・機能１：エクセルファイルの正誤入力と保存。
"""
import openpyxl as xl

def checker2(x,y):
    for i in range(len(x.value)):
        for j in range(len(y.value)):
            if x.value[i] == y.value[j]

def checker(ch,ans,save,row,start,col):
    try:
        wb1 = xl.load_workbook(ch)
        wb2 = xl.load_workbook(ans)
        ws1 = wb1.worksheets[0]
        ws2 = wb2.worksheets[0]
        for i in range(row):
            cell1 = ws1.cell(row+start,col)
            cell2 = ws2.cell(row+start,col)
            checker2(cell1,cell2)


        return 1
    except Exception as e:

        return e

ch = "C:\Users\yamamoto\Downloads\答え合わせソフト\FE_0623.xlsx"
ans = "C:\Users\yamamoto\Downloads\答え合わせソフト\FE_0623(1).xlsx"
test=checker(ch,ans,ans,13,1,2)
#正解、回答、出力先、問題数、開始行、回答列