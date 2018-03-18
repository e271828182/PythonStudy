#!/usr/bin/python
# -*- coding: utf-8 -*-

from Tkinter import Tk
from time import sleep, ctime
from tkMessageBox import showwarning
from urllib import urlopen
import win32com.client as win32

# warn = lambda app: showwarning(app, "Exit?")
RANGE = range(3, 8)
# TICKS = ("YHOO", "GOOG", "EBAY", "AMZN")
# COLS = ("TICKER", "PRICE", "CHG", "%AGE")
# URL = 'http://quote.yahoo.com/d/quotes.csv?s=%s&f=sl2c1p2'

def excel():
    app = "Excel"
    x1 = win32.gencache.EnsureDispatch('%s.Application' % app)
    ss = x1.Workbooks.Add()
    sh = ss.ActiveSheet
    x1.Visible = True
    sleep(1)

    sh.Cells(1, 1).Value = 'Python-to-%s Demo' % app
    sleep(1)
    for i in RANGE:
        sh.Cells(i, 1).Value = 'Line %d' % i
        # sleep(1)
    sh.Cells(i + 2, 1).Value = "Th-th-th-that's all folks!"

    # warn(app)
    ss.Close(SaveChanges=True, Filename='E:\\study\\Python2\\PythonStudy\\aa.xlsx')
    print(ss.Close.__defaults__)
    print(ss.Close.__doc__)
    print(ss.Close.__code__)
    x1.Application.Quit()


if __name__ == '__main__':
    Tk().wm_withdraw()
    excel()
