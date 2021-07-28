#フォルダ内を比較して不一致ファイルを一旦htmlにしてexcelで開いてPDFでエクスポートする。


import subprocess
import csv
import os
import shutil
from enum import Enum
from tkinter import messagebox
import pythoncom

#コマンドプロント上で以下を実行しインストールする。
#python -m pip install pywin32
#バージョンアップも必要かも
#python -m pip install pywin32 -U
import win32com.client 

#exe化はpyinstallerを使用
#インストールは以下
#pip install pyinstaller
#実行はコマンドプロンプトで
#pyinstaller GUI.py --onefile  --noconsole

path_Null = 'null.txt'
path_Winmerge = r'C:/Program Files/WinMerge/WinMergeU.exe'

#比較結果判定用文字列
tag_File_Same = 'テキスト ファイルは同一です'
tag_File_OnlyBefore = '左側のみ'
tag_File_OnlyAfter = '右側のみ'


def TestMain():
    path_Before = 'target/Before/'
    path_After  = 'target/After/'
    path_Output = 'target/Output/'
    abspath_Before = os.path.abspath(path_Before)
    abspath_After  = os.path.abspath(path_After)
    abspath_Output = os.path.abspath(path_Output)
    CodeToPdf(abspath_Before,abspath_After,abspath_Output)
    
def getState():
    return StringState


def CodeToPdf(abspath_Before,abspath_After,abspath_Output):
    global StringState
    StringState = ''

    #絶対パスに変更
    abspath_Null= os.path.abspath(path_Null)
    abspath_OutputTmp = os.path.abspath(abspath_Output+'/tmp')
    abspath_DiffList  = os.path.abspath(abspath_Output+'/DiffList.csv')

    print('以下フォルダの差分を取得します。')
    print('Before='+abspath_Before)
    print('After ='+abspath_After)
    print('Output='+abspath_Output)

    StringState = '初期化'

    #出力先フォルダの削除
    try:
        shutil.rmtree(abspath_Output)
        os.makedirs(abspath_OutputTmp)
    except :
        raise Exception('Error001:Outputフォルダを削除できませんでした。')

    #フォルダ比較結果をCSVで出力
    subprocess.run( [\
        path_Winmerge, \
        abspath_Before, \
        abspath_After, \
        '-minimize', \
        '-noninteractive', \
        '-noprefs', \
        '-cfg', 'Settings/DirViewExpandSubdirs=1', \
        '-cfg', 'ReportFiles/ReportType=0', \
        '-cfg', 'ReportFiles/IncludeFileCmpReport=1', \
        '-r', \
        '-u', \
        '-or', abspath_DiffList \
        ])

    #比較結果読み込み
    with open(abspath_DiffList) as f:
        reader = csv.reader(f)
        DiffList = [row for row in reader]

    for i in range(len(DiffList)):
        #テスト：出力先htmlの文字列生成
        subfolder = ''
        path_File_Before     =  ''
        path_File_After      =  ''
        path_File_Report     =  ''
        path_File_PDF        =  ''
        if (i > 3) & (len(DiffList[i]) > 1):
            if (DiffList[i][5] != "") & (DiffList[i][2] != tag_File_Same):
                #サブフォルダのパスとフォルダ出力用のサブフォルダ名生成
                if(DiffList[i][1] != ''):
                    subfolder     = '\\'+DiffList[i][1]
                    subfoldername = '\\'+DiffList[i][1].replace('\\','＞')+'＞'
                else:
                    subfolder     = ''
                    subfoldername = ''

                #比較ファイルのパス生成
                #片方しかない場合はnull.txtと比較する。
                if(DiffList[i][2].startswith(tag_File_OnlyAfter)):
                    path_File_After  = abspath_After  + subfolder+'\\' + DiffList[i][0]
                    path_File_Before = abspath_Null
                elif(DiffList[i][2].startswith(tag_File_OnlyBefore)):
                    path_File_After  = abspath_Null
                    path_File_Before = abspath_Before + subfolder+'\\' + DiffList[i][0]
                else:
                    path_File_After  = abspath_After  + subfolder+'\\' + DiffList[i][0]
                    path_File_Before = abspath_Before + subfolder+'\\' + DiffList[i][0]

                #出力ファイルパス生成
                path_File_Report     = abspath_OutputTmp + subfoldername + DiffList[i][0] + '.html'
                path_File_PDF        = abspath_Output    + subfoldername + DiffList[i][0] + '.pdf'

                #比較レポート生成(csvファイル生成)
                
                StringState ='PDF出力：' + subfoldername + DiffList[i][0]
                MakeDiff_ReportFile( \
                    path_File_Before, \
                    path_File_After, \
                    path_File_Report)

                #PDFへ変換
                HtmlToPDF_with_Excel(path_File_Report,path_File_PDF)

    StringState = '完了'

#ファイル比較レポート出力(HTML形式)
def MakeDiff_ReportFile(Before,After,Output):
    subprocess.run( [\
        path_Winmerge, \
        Before, \
        After, \
        '/minimize', \
        '/noninteractive', \
        '/u', \
        '/or', \
        Output])

#HTML→PDF
def HtmlToPDF_with_Excel(HtmlFile,PDFFile):

    #htmlをエクセルで開く
    pythoncom.CoInitialize()  # Excelを起動する前にこれを呼び出す
    excel = win32com.client.Dispatch("Excel.Application")
    excel.visible = False
    file = excel.Workbooks.Open(HtmlFile, UpdateLinks=0, ReadOnly=True)
    file.WorkSheets(1).Select()

    #横幅調整
    file.WorkSheets(1).Columns(1).ColumnWidth = 5
    file.WorkSheets(1).Columns(2).ColumnWidth = 100
    file.WorkSheets(1).Columns(3).ColumnWidth = 5
    file.WorkSheets(1).Columns(4).ColumnWidth = 100

    
    file.WorkSheets(1).Pagesetup.PrintTitleRows  = "$1:$1"
    file.WorkSheets(1).Cells.Font.Size = 8
    file.WorkSheets(1).Cells.Font.Name = "Consolas"

    #印刷設定
    file.WorkSheets(1).Pagesetup.Zoom  = False
    file.WorkSheets(1).Pagesetup.Orientation = 2
    file.WorkSheets(1).Pagesetup.FitToPagesWide = 1
    file.WorkSheets(1).Pagesetup.FitToPagesTall = False
    file.WorkSheets(1).Pagesetup.CenterHorizontally = True

    file.WorkSheets(1).Pagesetup.Papersize = 8
    file.WorkSheets(1).Pagesetup.RightMargin = 1
    file.WorkSheets(1).Pagesetup.LeftMargin = 1
    file.WorkSheets(1).Pagesetup.TopMargin = 1
    file.WorkSheets(1).Pagesetup.BottomMargin = 1
    file.WorkSheets(1).Pagesetup.HeaderMargin = 0
    file.WorkSheets(1).Pagesetup.FooterMargin = 0

    try:
        file.ActiveSheet.ExportAsFixedFormat(0,PDFFile)
    except :
        raise Exception('Error002:PDF出力できませんでした。')
    file.Close(SaveChanges=False)
