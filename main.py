#!/usr/bin/env python3

from datetime import datetime, timedelta
from os.path import dirname, join
from time import strptime, strftime, sleep
from json import loads
from urllib.request import urlopen
from win32com.shell import shell, shellcon
import pythoncom

from PIL import Image

LEVEL = 4
WIDTH = 550
DL_TIMEOUT = 60
BASE_DIR = "D:\\pictures\\"

LAST_TIME = ''
LAST_X = 0
LAST_Y = 0
REPEAT_TIMES = 10
SLEEP_TIME = 120

g_desk = ''
WINDOWS_WPSTYLE = shellcon.WPSTYLE_MAX

def getDeskComObject():
    global g_desk
    if not g_desk:
        g_desk = pythoncom.CoCreateInstance(shell.CLSID_ActiveDesktop, \
                                             None, pythoncom.CLSCTX_INPROC_SERVER, \
                                             shell.IID_IActiveDesktop)
    return g_desk

def setWallPaper(paper):
    desktop = getDeskComObject()
    if desktop:
        desktop.SetWallpaper(paper, 0)
        desktop.SetWallpaperOptions(WINDOWS_WPSTYLE)
        desktop.ApplyChanges(shellcon.AD_APPLY_ALL)

def download_chunk(x, y, latest):
    url_format = "http://himawari8.nict.go.jp/img/D531106/{}d/{}/{}_{}_{}.png"
    tTime = strftime("%Y/%m/%d/%H%M%S", latest)
    url = url_format.format(LEVEL, WIDTH, tTime, x, y)

    with urlopen(url , timeout=DL_TIMEOUT) as tile_w:
        tiledata = tile_w.read()
    print('Download:'+url)
    downloadPicture(url, str(x)+str(y)+".png")
    return x, y

def getLatestPictureTime():
    with urlopen("http://himawari8-dl.nict.go.jp/himawari8/img/D531106/latest.json") as latest_json:
        return loads(latest_json.read().decode("utf-8"))["date"]

def getBeijingTime(nowtime):
    delta = timedelta(hours=8)
    flag = datetime.strptime(nowtime, "%Y-%m-%d %H:%M:%S")
    return flag + delta

def downloadPicture(url, name):
    with urlopen(url , timeout=DL_TIMEOUT) as conn:
        if conn:
            with open(BASE_DIR+"\\"+name,'wb') as f:
                f.write(conn.read())
                print('Pic' + name + 'Saved!')

def mosaicPicture(timeStr):
    mw = 550
    ms = 550
    msize = mw * ms
    toImage = Image.new('RGB', (mw*4, ms*4))

    for x in  range(LEVEL):
        for y in range(LEVEL):
            fname = BASE_DIR+str(x)+str(y)+".png"
            fromImage = Image.open(fname)
            toImage.paste(fromImage,( x*mw, y*mw))
    picturePath = BASE_DIR+timeStr+".jpg"
    toImage.save(picturePath)
    return picturePath

def main(latest):
    global LAST_X, LAST_Y
    nowx = LAST_X
    nowy = LAST_Y
    requested_time = strptime(latest, "%Y-%m-%d %H:%M:%S")
    for x in range(nowx, LEVEL):
        for y in range(LEVEL):
            if x <= nowx and y < nowy:
                pass
            else:
                LAST_X = x
                LAST_Y = y
                download_chunk(x, y, requested_time)
    return requested_time

def checkTime():
    global LAST_TIME
    try:
        latest = getLatestPictureTime()
    except Exception:
        return False
    if LAST_TIME == latest:
        return False
    else:
        LAST_TIME = latest
        print(latest)
        return latest

while True:
    try:
        latest = checkTime()
        if latest:
            try_times = 1
            while try_times <= REPEAT_TIMES:
                try:
                    timeStruct = main(latest)
                except Exception:
                    print("\nTimeout while download "+str(LAST_X)+"."+str(LAST_Y)+"."+str(latest)+" AND try times "+str(try_times))
                    try_times = try_times + 1
                    continue
                if timeStruct:
                    timeStr = strftime("%Y%m%d-%H%M%S", timeStruct)
                    picturePath = mosaicPicture(timeStr)
                    setWallPaper(picturePath)
                    break
            LAST_X = 0
            LAST_Y = 0
        SLEEP_TIME(60)
    except Exception:
        pass
