import sys
import openpyxl as xl
import glob
from tkinter import messagebox
import pandas as pd

PSpath = sys.argv[1] #コマンドライン引数 [0]はファイル名のため[1]になる

xlfiles = glob.glob(PSpath + "/*.xls*") #xls,xlsx,xlsmのファイルをリスト化
csvfile = glob.glob(PSpath + "/*.csv") #csvファイルをリスト化

df = pd.read_csv(csvfile[0],encoding='utf8') #1行目はヘッダーとみなす
count = len(df) #len(df)でデータフレームの行数を取得（ヘッダー行除く）

for file in xlfiles:
    wb = xl.load_workbook(filename=file, keep_vba=True) #keep_vba=Trueでxlsmを編集可能
    sheet = wb.active

    for i in range(count): #range(5)の場合、0,1,2,3,4。dfの行は0から始まるため補正なしでOK
        r = df.loc[i , 'row'] #i行row列の値を取り出す
        c = df.loc[i , 'column'] #i行column列の値を取り出す
        v = df.loc[i , 'value'] #i行value列の値を取り出す
        
        sheet.cell(row=r, column=c).value = v
    
    wb.save(file)
    
messagebox.showinfo('Python', 'Finished.')