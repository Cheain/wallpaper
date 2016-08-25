import datetime
import gc
import os
import re
import sys
import time
import urllib.request
import win32api
import win32gui

import win32com.client
import win32con
from PIL import Image, ImageDraw, ImageFont, ImageFile
from retrying import retry


def getDate():
    localTime = datetime.datetime.now() - datetime.timedelta(minutes=30)
    zeroTime = localTime - datetime.timedelta(hours=8)
    localDate = "{}年{}月{}日\n  {}时{}分00秒" \
        .format(localTime.year, '0' * (2 - len(str(localTime.month))) + str(localTime.month),
                '0' * (2 - len(str(localTime.day))) + str(localTime.day),
                '0' * (2 - len(str(localTime.hour))) + str(localTime.hour),
                str(localTime.minute // 10) + '0' * (2 - len(str(localTime.minute // 10))))
    zeroDate = {'year': str(zeroTime.year), 'month': '0' * (2 - len(str(zeroTime.month))) + str(zeroTime.month),
                'day': '0' * (2 - len(str(zeroTime.day))) + str(zeroTime.day),
                'hour': '0' * (2 - len(str(zeroTime.hour))) + str(zeroTime.hour),
                'minute': str(zeroTime.minute // 10) + '0' * (2 - len(str(zeroTime.minute // 10)))}
    return localDate, zeroDate


def getURL(multiple, zeroDate, i, j):
    rawURL = 'http://himawari8-dl.nict.go.jp/himawari8/img/D531106/{}d/550/{}/{}/{}/{}{}00_{}_{}.png'
    earthURL = rawURL.format(multiple, zeroDate['year'], zeroDate['month'], zeroDate['day'], zeroDate['hour'],
                             zeroDate['minute'], i, j)
    return earthURL


def downloadImg(multiple, zeroDate):
    earth = Image.new('RGB', (multiple * 550, multiple * 550))
    for i in range(multiple):
        for j in range(multiple):
            URL = getURL(multiple, zeroDate, i, j)
            print('{}/{} picture started to download\n'
                  '    URL->{}'.format(i * multiple + j + 1, pow(multiple, 2), URL))
            image = down(URL)
            f = ImageFile.Parser()
            f.feed(image)
            tempImg = f.close()
            earth.paste(tempImg, (i * 550, j * 550, (i + 1) * 550, (j + 1) * 550))
            print('{}/{} picture has been download'.format(i * multiple + j + 1, pow(multiple, 2)))
    return earth


# @retry(wait_random_min=0, wait_random_max=3000)
@retry
def down(URL):
    print('    Try')
    headers = {'Connection': 'Keep-Alive',
               'Accept': 'text/html, application/xhtml+xml, */*',
               'Accept-Language': 'en-US,en;q=0.8,zh-Hans-CN;q=0.5,zh-Hans;q=0.3',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko'
               }
    req = urllib.request.Request(URL, headers=headers)
    image = urllib.request.urlopen(req, timeout=1).read()
    return image


def getFontName():
    g = lambda i: re.search('\.ttf$', i, re.I) or re.search('\.ttc$', i, re.I)
    for path in (os.getcwd(), 'C:\\Windows\\Fonts'):
        for i in os.listdir(path):
            if g(i):
                fontName = os.path.join(path, i)
                return fontName
    win32api.MessageBox(0, "找不到字体文件，请下载喜欢的字体文件到{}目录下，并重启程序".format(os.getcwd()),
                        "错误", win32con.MB_OK)
    sys.exit()


def drawDate(localDate, earth, imgPath):
    height, width = earth.size
    font = ImageFont.truetype(getFontName(), height // 30)
    date = localDate.translate(str.maketrans('年月时分', '//::', '日秒'))
    textHeight, textWidth = font.getsize(date[0:11])
    draw = ImageDraw.Draw(earth)
    draw.text((height - textHeight, textWidth // 2), date, (255, 255, 255), font=font)
    imgName = os.path.join(imgPath, '{}.jpg'.format(localDate.replace('\n  ', '')))
    # for i in os.listdir(wallPath):
    #     os.remove(os.path.join(wallPath, i))
    earth.save(os.path.join(imgPath, imgName))
    print('Picture has been saved at {}'.format(imgName))
    return imgName


def getPath():
    for disk in ('E:', 'D:', 'C:'):
        if os.path.isdir(disk):
            imgPath = os.path.join(disk, os.path.sep, 'himawari8')
            if not os.path.isdir(imgPath):
                os.makedirs(imgPath)
            return imgPath


def setWallpaper(filename):
    key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, 'Control Panel\Desktop', 0, win32con.KEY_ALL_ACCESS)
    current = win32api.RegQueryValueEx(key, 'Wallpaper')[0]
    if current != filename:
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, filename, 1 + 2)
        print('Picture {} has been setted as wallpaper'.format(filename))
    win32api.RegCloseKey(key)


def checkProcess():
    wmi = win32com.client.GetObject('winmgmts:')
    # for p in wmi.InstancesOf('win32_process'):
    #     print(p.Name, p.Properties_('ProcessId'),
    #           int(p.Properties_('UserModeTime').Value) + int(p.Properties_('KernelModeTime').Value))
    #     children = wmi.ExecQuery('Select * from win32_process where ParentProcessId=%s' % p.Properties_('ProcessId'))
    #     for child in children:
    #         print('\t', child.Name, child.Properties_('ProcessId'),
    #               int(child.Properties_('UserModeTime').Value) + int(child.Properties_('KernelModeTime').Value))
    count = 0
    for p in wmi.InstancesOf('win32_process'):
        if p.name == 'H8Img.exe':
            print(p.Name, p.Properties_('ProcessId'),
                  int(p.Properties_('UserModeTime').Value) + int(p.Properties_('KernelModeTime').Value))
            count += 1
            if count > 2:
                win32api.MessageBox(0, "程序已运行，请勿重复开启", "提示", win32con.MB_OK)
                sys.exit(0)


def setStartup():
    key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                              0, win32con.KEY_ALL_ACCESS)
    name = re.sub('\.py$', '.exe', os.path.abspath(__file__))
    win32api.RegSetValueEx(key, 'H8Img', 0, win32con.REG_SZ, name)
    win32api.RegCloseKey(key)


def main():
    multiple = 8
    localDate, zeroDate = getDate()
    imgPath = getPath()
    earth = downloadImg(multiple, zeroDate)
    ImgName = drawDate(localDate, earth, imgPath)
    setWallpaper(ImgName)


if __name__ == '__main__':
    checkProcess()
    setStartup()
    while True:
        try:
            a = time.time()
            gc.disable()
            main()
            gc.enable()
            b = time.time()
            print('{}s'.format(b - a))
            gc.collect()
            c = time.time()
            print('{}s'.format(c - b))
        except Exception as e:
            print(e)
        print('Start waiting')
        time.sleep(300)
