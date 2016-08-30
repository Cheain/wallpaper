import os
import urllib.request
import win32api
import win32gui
import zipfile

import win32con
from PIL import ImageFile
from retrying import retry


@retry
def down():
    URL = 'http://himawari8-dl.nict.go.jp/himawari8/img/D531106/1d/550/2016/08/23/112000_0_0.png'
    print(URL)
    req = urllib.request.Request(URL, headers={
        'Connection': 'Keep-Alive',
        'Accept': 'text/html, application/xhtml+xml, */*',
        'Accept-Language': 'en-US,en;q=0.8,zh-Hans-CN;q=0.5,zh-Hans;q=0.3',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko'
    })
    image = urllib.request.urlopen(req, timeout=1).read()
    im = ImageFile.Parser()
    im.show()
    print('OK')


def t():
    fp = open("D:\\2.png", "rb")
    fp = urllib.request.urlopen('http://himawari8-dl.nict.go.jp/himawari8/'
                                'img/D531106/1d/550/2016/08/23/112000_0_0.png').read()
    print(1)
    # fp=fp.read()
    print(2)
    p = ImageFile.Parser()
    p.feed(fp)
    im = p.close()
    im.show()
    while 1:
        s = fp.read(1024)
        if not s:
            break
        p.feed(s)
    im = p.close()
    im.show()


def downTTF():
    ziti = urllib.request.urlopen('http://jsdx.sc.chinaz.com/Files/DownLoad/font4/201511/bb5300.rar').read()
    with open(os.getcwd() + 'ziti.rar', 'rw') as f:
        f.write(ziti)
        f.close()
    f = zipfile.ZipFile(os.getcwd() + 'ziti.rar', 'r')
    for file in f.namelist():
        f.extract(file, os.getcwd())
    f.close()


import win32com.client


def pro():
    wmi = win32com.client.GetObject('winmgmts:')
    # for p in wmi.InstancesOf('win32_process'):
    #     print(p.Name, p.Properties_('ProcessId'),
    #           int(p.Properties_('UserModeTime').Value) + int(p.Properties_('KernelModeTime').Value))
    #     children = wmi.ExecQuery('Select * from win32_process where ParentProcessId=%s' % p.Properties_('ProcessId'))
    #     for child in children:
    #         print('\t', child.Name, child.Properties_('ProcessId'),
    #               int(child.Properties_('UserModeTime').Value) + int(child.Properties_('KernelModeTime').Value))
    for p in wmi.InstancesOf('win32_process'):
        if p.name == 'H8Img.exe':
            print(p.Name, p.Properties_('ProcessId'),
                  int(p.Properties_('UserModeTime').Value) + int(p.Properties_('KernelModeTime').Value))
            break


def SetWallpaper(filename):
    key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, 'Control Panel\Desktop', 0, win32con.KEY_ALL_ACCESS)
    print(key)
    current = win32api.RegQueryValueEx(key, 'Wallpaper')[0]
    print(win32api.RegQueryValueEx(key, 'Wallpaper'), '\n', current)
    if current != filename:
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, filename, 1 + 2)
    win32api.RegCloseKey(key)


def setStartup():
    key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                              0, win32con.KEY_ALL_ACCESS)
    win32api.RegSetValueEx(key, 'H8Img', 0, win32con.REG_SZ, os.path.abspath(__file__))
    win32api.RegCloseKey(key)

from tkinter import *
import tkinter.messagebox as messagebox

class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.nameInput = Entry(self)
        self.nameInput.pack()
        self.alertButton = Button(self, text='Hello', command=self.hello)
        self.alertButton.pack()

    def hello(self):
        name = self.nameInput.get() or 'world'
        messagebox.showinfo('Message', 'Hello, %s' % name)

def App():
    app = Application()
    # 设置窗口标题:
    app.master.title('Hello World')
    # 主消息循环:
    app.mainloop()

class MainWindow:
    def __init__(self):
        self.frame = Tk()

        self.label_name = Label(self.frame,text = "name:")
        self.label_age = Label(self.frame,text = "age:")
        self.label_sex = Label(self.frame,text = "sex:")

        self.text_name = Text(self.frame,height = "1",width = 30)
        self.text_age = Text(self.frame,height = "1",width = 30)
        self.text_sex = Text(self.frame,height = "1",width = 30)

        self.label_name.grid(row = 0,column = 0)
        self.label_age.grid(row = 1,column = 0)
        self.label_sex.grid(row = 2,column = 0)

        self.button_ok = Button(self.frame,text = "ok",width = 10)
        self.button_cancel = Button(self.frame,text = "cancel",width = 10)

        self.text_name.grid(row = 0,column = 1)
        self.text_age.grid(row = 1,column = 1)
        self.text_sex.grid(row = 2,column = 1)

        self.button_ok.grid(row = 3,column = 0)
        self.button_cancel.grid(row = 3,column = 1)

        self.frame.mainloop()

frame = MainWindow()

import tkinter.ttk