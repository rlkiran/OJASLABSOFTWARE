import win32api
import win32print
import os

pwd = os.getcwd()
filename = ''


def printFile(filename):
    win32api.ShellExecute(
        0,
        "printto",
        filename,
        '"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )
    os.chdir("..")
