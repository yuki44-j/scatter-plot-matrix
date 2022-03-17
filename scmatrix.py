# coding: utf-8

import tkinter
from tkinter import filedialog
import pandas as pd
import seaborn as sns

tk = tkinter.Tk()
tk.withdraw()

fTyp = [("Files","*.*")]
iDir = r'C:\Users'
titleText = "Select file"
file = filedialog.askopenfilename(filetypes = fTyp, title = titleText, initialdir = iDir)
filename = '\\'.join(file.split('/'))

multiDf = pd.read_excel(filename, index_col=0)
multiFig = sns.pairplot(multiDf)

fTyp = [("JPEG", ".jpg"), ("PNG", ".png"), ("PDF", ".pdf") ]
iDir = r'C:\Users'
titleText = "Select file"
file = filedialog.asksaveasfilename(filetypes = fTyp, title = titleText, initialdir = iDir, defaultextension = "jpg")
filename = '\\'.join(file.split('/'))

multiFig.savefig(filename)
