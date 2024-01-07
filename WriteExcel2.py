import sys
import glob
import pandas as pd
import openpyxl as xl
from tkinter import messagebox

def main():
    ps_path = sys.argv[1]
    xls_files = glob.glob(ps_path + "/*.xls*")
    csv_file = glob.glob(ps_path + "/*.csv")[0]
    sheet_name = sys.argv[2]

    df = pd.read_csv(csv_file, encoding='utf8')

    for xls_file in xls_files:
        process_excel_file(xls_file, sheet_name, df)

    messagebox.showinfo('Python', 'Finished.')

def process_excel_file(file, sheet_name, df):
    wb = xl.load_workbook(filename=file, keep_vba=True)

    for sheet in wb.worksheets:
        if sheet.title == sheet_name:
            ws = wb[sheet_name]

            for _, row in df.iterrows():
                r, c, v = row['row'], row['column'], row['value']
                ws.cell(row=r, column=c).value = v

            wb.save(file)

if __name__ == "__main__":
    main()