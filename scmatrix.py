# coding: utf-8

import tkinter
from tkinter import filedialog
import pandas as pd
import seaborn as sns
from PIL import Image
import xlwings as xw
import os

# 散布図行列 を作る本体関数
def makeScatMtrx():
    df = pd.read_excel(filename, index_col=0)
    wb = xw.Book(filename)
    sheet = wb.sheets.add('scat_mtrx')

    df = pd.read_excel(filename, index_col=0)
    fig = sns.pairplot(df, kind='reg')
    # get_figure() が機能しないので一旦save
    tempPath = '\\'.join(file.split('/')[:-1] + ['temp_fig.jpg'])
    fig.savefig(tempPath)
    # fig をもう一回拾ってﾘｻｲｽﾞ後ﾜｰｸｼｰﾄに貼り付け
    img = Image.open(tempPath)
    img = img.resize((512, 512))
    img.save(tempPath)
    sheet.pictures.add(tempPath)
    # 一時ﾌｧｲﾙを削除
    os.unlink(tempPath)
    
    # 決定係数の表
    corr = df.corr()
    coList = []
    for row in range(len(corr.index)):
        rList = [corr.index[row]]
        for col in range(len(corr.index)):
            rList.append(corr.iloc[row, col] ** 2)
        coList.append(rList)

    sheet.range("J3").value = coList
    sheet.range("K2").value = list(corr.index)
    sheet.range("K1").value = "重決定係数R2"

tk = tkinter.Tk()
tk.withdraw()

fTyp = [("Files","*.*")]
iDir = r'C:\Users'
titleText = "Select file"
file = filedialog.askopenfilename(filetypes = fTyp, title = titleText, initialdir = iDir)
filename = '\\'.join(file.split('/'))

makeScatMtrx()
