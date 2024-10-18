import glob
import pandas as pd
from datetime import datetime

# 年、月、日を取得
now = datetime.now()
current_year = now.year
current_month = now.month
current_day = now.day


# 指定したフォルダのパス
read_excel_path = 'C:/Users/81808/source/hinai/read_excel' # 希望シフトが書いてあるファイルが入ってるフォルダのパス 
write_excel_path = 'C:/Users/81808/source/hinai/write_excel' # シフト表のテンプレが入ってるフォルダのパス

# フォルダ内のExcelファイル（拡張子が .xlsx のもの）を検索
read_excel_file = glob.glob(f'{read_excel_path}/*.xlsx')
write_excel_file = glob.glob(f'{write_excel_path}/*.xlsx')

def find_file(excelfile):
    if excelfile:
        file_path = excelfile[0]
        df = pd.read_excel(file_path)
        print(f'読み込んだファイル: {file_path}')
    else:
        print('Excelファイルが見つかりませんでした。')
    return df

df_read = find_file(read_excel_file)
df_write = find_file(write_excel_file)

# 各行をリストとして格納
rows_as_list = []
name_list = []
for index, row in df_read.iterrows():
    rows_as_list.append(row.tolist())  # 行をリストに変換して追加

# name_listに名前を格納
for row in rows_as_list:
    name_list.append(row[2])
# print(name_list)

# writeファイルの1列目に一致する名前があるか検索
for name in name_list:
    count = 0
    find = False
    for write_sel_1 in df_write.iloc[:, 0]:  # 1列目を取得
        count += 1
        if name == write_sel_1:  # 一致する名前がname_listにあるか確認
            print(f"{name}はwriteファイルの{count+1}行目にありました")
            find = True
            break
    if not (find):
        print(f"{name}を見つけられませんでした")

# for i in range(4, 4 + 15):
