
import os,sys
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import DiffAndPDF
import time
import threading

def_path_Before = 'target/Before/'
def_path_After  = 'target/After/'
def_path_Output = 'target/Output/'

# フォルダ指定の関数
def dirdialog_clicked():
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    StringVar_path_Before.set(iDirPath)
    
def dirdialog_clicked2():
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    StringVar_path_After.set(iDirPath)
    
def dirdialog_clicked3():
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    StringVar_path_Output.set(iDirPath)
    
# 実行ボタン押下時の実行関数
#Todo:ボタンの無効化処理が仮実装
RunFlg = True
def StateUpdate():
    button1['state'] = "disable" 
    RunFlg = True
    while(DiffAndPDF.StringState != '完了'):
        StringVar_State.set(DiffAndPDF.StringState)
        time.sleep(1)
    button1['state'] = "enable"

#変換処理のコール(スレッド用)
def RunOperation(abspath_Before,abspath_After,abspath_Output):
    try:
        DiffAndPDF.CodeToPdf(abspath_Before,abspath_After,abspath_Output)
        RunFlg = False
    except Exception as e:
        messagebox.showerror("error", e)
        RunFlg = False


def conductMain():

    path_Before = StringVar_path_Before.get()
    path_After  = StringVar_path_After.get()
    path_Output = StringVar_path_Output.get()
    
    if (path_Before!='')&(path_After!='')&(path_Output!=''):
        abspath_Before = os.path.abspath(path_Before)
        abspath_After  = os.path.abspath(path_After)
        abspath_Output = os.path.abspath(path_Output)
            
        th1 = threading.Thread( \
            target=RunOperation, \
            args=(abspath_Before,abspath_After,abspath_Output))
            
        th2 = threading.Thread( \
            target=StateUpdate)
        th1.start()
        th2.start()
                
    else:
        messagebox.showerror("error", "パスの指定がありません。")

     

def main():
    global StringVar_path_Before
    global StringVar_path_After
    global StringVar_path_Output
    global StringVar_State
    global button1

    # rootの作成
    root = Tk()
    root.title("Diff and PDF")

    # Frame1の作成
    frame1 = ttk.Frame(root, padding=10)
    frame1.grid(row=0, column=1, sticky=E)

    # 「フォルダ参照」ラベルの作成
    IDirLabe_Before = ttk.Label(frame1, text="フォルダ参照：変更前", padding=(5, 2))
    IDirLabe_Before.pack(side=LEFT)

    # 「フォルダ参照」エントリーの作成
    StringVar_path_Before = StringVar()
    IDirEntry = ttk.Entry(frame1, textvariable=StringVar_path_Before, width=60)
    IDirEntry.insert(0,os.path.abspath(def_path_Before))
    IDirEntry.pack(side=LEFT)

    # 「フォルダ参照」ボタンの作成
    IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    IDirButton.pack(side=LEFT)

    # Frame2の作成
    frame2 = ttk.Frame(root, padding=10)
    frame2.grid(row=1, column=1, sticky=E)

    # 「ファイル参照」ラベルの作成
    IDirLabe_After = ttk.Label(frame2, text="フォルダ参照：変更後", padding=(5, 2))
    IDirLabe_After.pack(side=LEFT)

    # 「ファイル参照」エントリーの作成
    StringVar_path_After = StringVar()
    IFileEntry = ttk.Entry(frame2, textvariable=StringVar_path_After, width=60)
    IFileEntry.insert(0,os.path.abspath(def_path_After))
    IFileEntry.pack(side=LEFT)

    # 「ファイル参照」ボタンの作成
    IFileButton = ttk.Button(frame2, text="参照", command=dirdialog_clicked2)
    IFileButton.pack(side=LEFT)

    # Frame3の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid(row=2,column=1,sticky=W)

    # 「ファイル参照」ラベルの作成
    IDirLabe_Output = ttk.Label(frame3, text="フォルダ参照：出力　", padding=(5, 2))
    IDirLabe_Output.pack(side=LEFT)

    # 「ファイル参照」エントリーの作成
    StringVar_path_Output = StringVar()
    IFileEntry = ttk.Entry(frame3, textvariable=StringVar_path_Output, width=60)
    IFileEntry.insert(0,os.path.abspath(def_path_Output))
    IFileEntry.pack(side=LEFT)

    # 「ファイル参照」ボタンの作成
    IFileButton = ttk.Button(frame3, text="参照", command=dirdialog_clicked3)
    IFileButton.pack(side=LEFT)

    # Frame4の作成
    frame4 = ttk.Frame(root, padding=10)
    frame4.grid(row=4, column=1, sticky=E)


    # 実行ボタンの設置
    button1 = ttk.Button(frame4, text="PDF出力", command=conductMain)
    button1.pack(fill = "x", padx=30, side = "left")

    IDirLabe_State = ttk.Label(frame4, text="出力状態", padding=(5, 2))
    IDirLabe_State.pack(side=LEFT)

    # 「ファイル参照」ラベルの作成
    StringVar_State = StringVar()
    StateStra = ttk.Entry(frame4, textvariable=StringVar_State, width=60)
    StateStra.insert(0,' ')
    StateStra.pack(side=LEFT)


    root.mainloop()

if __name__ == "__main__":
    main()