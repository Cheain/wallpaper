import os
import urllib.request
import zipfile

import win32api

import win32con
import win32gui
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
    print(win32api.RegQueryValueEx(key, 'Wallpaper'),'\n',current)
    if current != filename:
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, filename, 1 + 2)
    win32api.RegCloseKey(key)

SetWallpaper('E:\himawari8\wallpaper\2016年08月25日21时10分00秒.jpg')