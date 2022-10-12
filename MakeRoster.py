#ライブラリのインポート
import openpyxl #Excel操作系
import os #ファイル存在確認

#設定ファイルの格納(固定のため埋め込み) 参考URL:https://python.keicode.com/lang/file-exists.php
StrSettingFilePass = 'Setting\\UserSettingConfig.xlsx'

#編集するExcelファイル、シートの読み込み
excel_file = openpyxl.load_workbook('ファイル名')
excel_sheet = excel_file["シート名"]

#セル値の取得
excel_sheet['セル番地'].value

#セル値の代入・変更
excel_sheet['セル番地'].value = '代入・変更したい値'


#Excelファイルの上書き保存
excel_file.save('ファイル名')