import subprocess
import csv
import os
#コマンドプロント上で以下を実行しインストールする。
#python -m pip install pywin32
#バージョンアップも必要かも
import win32com.client 

path_Null = 'null.txt'
path_Before = 'target/Before/'
path_After  = 'target/After/'
path_Output = 'target/Output/'
path_DiffList = 'target/Output/DiffList.csv'
path_Winmerge = r'C:/Program Files/WinMerge/WinMergeU.exe'
tag_directory_Diff = 'フォルダーは異なっています'
tag_directory_Same = '同一'
tag_File_Same = 'テキスト ファイルは同一です'
tag_File_OnlyBefore = '左側のみ'
tag_File_OnlyAfter = '右側のみ'

def main():
    #絶対パスに変更
    abspath_Null= os.path.abspath(path_Null)
    abspath_Before = os.path.abspath(path_Before)
    abspath_After  = os.path.abspath(path_After)
    abspath_Output = os.path.abspath(path_Output)
    abspath_DiffList  = os.path.abspath(path_DiffList)

    #フォルダ比較結果をCSVで出力
    print('以下フォルダの差分を取得します。')
    print('Before='+abspath_Before)
    print('After ='+abspath_After)

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


    print(len(DiffList))
    for i in range(len(DiffList)):
        #テスト：出力先htmlの文字列生成
        subfolder = ''
        path_File_Before     =  ''
        path_File_After      =  ''
        path_File_Report     =  ''
        path_File_PDF        =  ''
        if (i > 3) & (len(DiffList[i]) > 1):
            if (DiffList[i][5] != "") & (DiffList[i][2] != tag_File_Same):
                subfolder = '/'+DiffList[i][1]
                path_File_After      = abspath_After +subfolder+'/' +DiffList[i][0]
                path_File_Before     = abspath_Before+subfolder+'/' +DiffList[i][0]
                if(DiffList[i][2][0:4] == tag_File_OnlyAfter):
                    path_File_Before     = abspath_Null
                if(DiffList[i][2][0:4] == tag_File_OnlyBefore):
                    path_File_After      = abspath_Null
                path_File_Report     = abspath_Output+"/tmp"+subfolder+'＞'+DiffList[i][0]+'.html'
                path_File_PDF        = abspath_Output+subfolder+'＞'+DiffList[i][0]+'.pdf'

                print(i.__str__()+":"+DiffList[i][2][0:4])
                print(" :"+path_File_Before)
                print(" :"+path_File_After)

                MakeDiff_ReportFile( \
                    path_File_Before, \
                    path_File_After, \
                    path_File_Report)

                HtmlToPDF(path_File_Report,path_File_PDF)

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
def HtmlToPDF(HtmlFile,PDFFile):

    #テスト:PDF出力
    excel = win32com.client.Dispatch("Excel.Application")
    file = excel.Workbooks.Open(HtmlFile, UpdateLinks=0, ReadOnly=True)
    file.WorkSheets(1).Select()
    file.WorkSheets(1).Pagesetup.Zoom  = False
    file.WorkSheets(1).Pagesetup.Orientation = 2
    file.WorkSheets(1).Pagesetup.FitToPagesWide = 1
    file.WorkSheets(1).Pagesetup.CenterHorizontally = True

    file.WorkSheets(1).Pagesetup.RightMargin = 1
    file.WorkSheets(1).Pagesetup.LeftMargin = 1
    file.WorkSheets(1).Pagesetup.TopMargin = 1
    file.WorkSheets(1).Pagesetup.BottomMargin = 1
    file.WorkSheets(1).Pagesetup.HeaderMargin = 0
    file.WorkSheets(1).Pagesetup.FooterMargin = 0

    try:
        file.ActiveSheet.ExportAsFixedFormat(0,PDFFile)
    except ZeroDivisionError:
        print('Error')
    file.Close(SaveChanges=False)

#実行
main()