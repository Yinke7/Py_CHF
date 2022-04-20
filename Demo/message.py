import tkinter as tk
import tkinter.messagebox

def ShowMsg(info):
    top = tkinter.Tk()
    top.withdraw()
    top.update()
    tkinter.messagebox.showwarning('提示', info)
    top.destroy()
    return
