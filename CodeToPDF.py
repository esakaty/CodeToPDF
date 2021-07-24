import subprocess
import csv

#コマンドプロント上で以下を実行しインストールする。
#python -m pip install pywin32
#バージョンアップも必要かも
import win32com.client 

path_Before = 'target/Before'
path_After  = 'target/After'
path_Output = 'target/Output/'
path_DiffList = 'target/Output/out.csv'

def main():
    #フォルダ比較結果をCSVで出力
    subprocess.run([r'C:/Program Files/WinMerge/WinMergeU.exe', \
        path_Before, \
        path_After, \
        '-minimize', \
        '-noninteractive', \
        '-noprefs', \
        '-cfg', 'Settings/DirViewExpandSubdirs=1', \
        '-cfg', 'ReportFiles/ReportType=0', \
        '-cfg', 'ReportFiles/IncludeFileCmpReport=1', \
        '-r', \
        '-u', \
        '-or', path_DiffList])

    #比較結果読み込み
    with open(path_DiffList) as f:
        reader = csv.reader(f)
        DiffList = [row for row in reader]

    #テスト：出力先htmlの文字列生成
    path_ReportFile = path_Output+DiffList[4][0]+'.html'

    MakeDifPDF( \
        path_Before+'/'+DiffList[4][1]+'/'+DiffList[4][0], \
        path_After+'/'+DiffList[4][1]+'/'+DiffList[4][0], \
        path_ReportFile)

def MakeDifPDF(Before,After,Output):
    #テスト：ファイル比較
    subprocess.run( \
        [r'C:/Program Files/WinMerge/WinMergeU.exe', \
        Before, \
        After, \
        '/minimize', \
        '/noninteractive', \
        '/u', \
        '/or', \
        Output])

    #テスト:PDF出力
    excel = win32com.client.Dispatch("Excel.Application")
    file = excel.Workbooks.Open(r'M:/WorkSpace/PrjCodeToPDF/'+Output, UpdateLinks=0, ReadOnly=True)
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
        file.ActiveSheet.ExportAsFixedFormat(0,r'M:/WorkSpace/PrjCodeToPDF/bb')
    except ZeroDivisionError:
        print('Error')
    file.Close(SaveChanges=False)

#実行
main()