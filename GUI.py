
import os,sys
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import DiffAndPDF

def_path_Before = 'target/Before/'
def_path_After  = 'target/After/'
def_path_Output = 'target/Output/'

# フォルダ指定の関数
def dirdialog_clicked():
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    entry1.set(iDirPath)
    
def dirdialog_clicked2():
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    entry2.set(iDirPath)
    
def dirdialog_clicked3():
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    entry3.set(iDirPath)
    
# 実行ボタン押下時の実行関数
def conductMain():
    text = ""

    path_Before = entry1.get()
    path_After  = entry2.get()
    path_Output = entry3.get()
    
    if (path_Before!='')&(path_After!='')&(path_Output!=''):
        abspath_Before = os.path.abspath(path_Before)
        abspath_After  = os.path.abspath(path_After)
        abspath_Output = os.path.abspath(path_Output)
        try:
            DiffAndPDF.CodeToPdf(abspath_Before,abspath_After,abspath_Output)
            messagebox.showinfo("完了", "出力完了しました。")
        except Exception as e:
            messagebox.showerror("error", e)
    else:
        messagebox.showerror("error", "パスの指定がありません。")

if __name__ == "__main__":

    # rootの作成
    root = Tk()
    root.title("Diff and PDF")

    # Frame1の作成
    frame1 = ttk.Frame(root, padding=10)
    frame1.grid(row=0, column=1, sticky=E)

    # 「フォルダ参照」ラベルの作成
    IDirLabel = ttk.Label(frame1, text="フォルダ参照：変更前", padding=(5, 2))
    IDirLabel.pack(side=LEFT)

    # 「フォルダ参照」エントリーの作成
    entry1 = StringVar()
    IDirEntry = ttk.Entry(frame1, textvariable=entry1, width=60)
    IDirEntry.insert(0,os.path.abspath(def_path_Before))
    IDirEntry.pack(side=LEFT)

    # 「フォルダ参照」ボタンの作成
    IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    IDirButton.pack(side=LEFT)

    # Frame2の作成
    frame2 = ttk.Frame(root, padding=10)
    frame2.grid(row=1, column=1, sticky=E)

    # 「ファイル参照」ラベルの作成
    IFileLabel = ttk.Label(frame2, text="フォルダ参照：変更後", padding=(5, 2))
    IFileLabel.pack(side=LEFT)

    # 「ファイル参照」エントリーの作成
    entry2 = StringVar()
    IFileEntry = ttk.Entry(frame2, textvariable=entry2, width=60)
    IFileEntry.insert(0,os.path.abspath(def_path_After))
    IFileEntry.pack(side=LEFT)

    # 「ファイル参照」ボタンの作成
    IFileButton = ttk.Button(frame2, text="参照", command=dirdialog_clicked2)
    IFileButton.pack(side=LEFT)

    # Frame3の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid(row=2,column=1,sticky=W)

    # 「ファイル参照」ラベルの作成
    IFileLabel = ttk.Label(frame3, text="フォルダ参照：出力　", padding=(5, 2))
    IFileLabel.pack(side=LEFT)

    # 「ファイル参照」エントリーの作成
    entry3 = StringVar()
    IFileEntry = ttk.Entry(frame3, textvariable=entry3, width=60)
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

    # Frame5の作成
    frame5 = ttk.Frame(root, padding=10)
    frame5.grid(row=5, column=1, sticky=E)


    root.mainloop()
