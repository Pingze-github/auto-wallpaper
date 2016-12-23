# usr/bin/env python
# -*-coding=utf-8-*-

from PyQt4 import QtGui,QtCore
import sys
import os
import shutil
import urllib
from PIL import Image
import win32api,win32con,win32gui
import threading
import random
import re
import time
import traceback
import json

#默认参数声明
mode_default = 1
userPath1_default = r'D:\Users\win10\Pictures\wallpaper\bing'
userPath2_default = r'D:\Users\win10\Pictures\wallpaper\collected'
userPath3_default = r'D:\Users\win10\Pictures\wallpaper\windows lockscreen'
#设置参数声明
mode=1
userPath1 = r'D:\Users\win10\Pictures\wallpaper\bing'
userPath2 = r'D:\Users\win10\Pictures\wallpaper\collected'
userPath3 = r'D:\Users\win10\Pictures\wallpaper\windows lockscreen'

#全局声明
lockscreenPath=r'C:\Users\71969\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets'
url=r'http://www.bing.com/?mkt=zh-CN'
iconPath = 'AW.ico'
procName = u'AutoWallpaper'
ISOTIMEFORMAT = '%Y%m%d'

def stopCheck():
    '''
    中止循环线程
    '''
    global timeri, timerc
    timeri.cancel()
    timerc.cancel()

def mkdir(path):
    '''
    带有异常检查的目录创建
    '''
    if(os.path.exists(path)==False or os.path.isdir(path)==False):
        os.mkdir(path)
        print('--- Create directory '+path+' ---')

def mkfile(path):
    '''
    带有异常检查的文件创建
    '''
    if(os.path.exists(path)==False or os.path.isfile(path)==False):
        os.mknod(path)
        print('--- Create file '+path+' ---')

def getCheckedIndex(checkGroup):
    '''
    获取一个QButtonGroup中被选中的QCheckBox的索引数
    '''
    i = 1
    for check in checkGroup.buttons():
        if check.isChecked() == 1:
            return i
        i += 1

def setCheckedIndex(checkGroup, index):
    '''
    获取一个QButtonGroup中被选中的QCheckBox的索引数
    '''
    i = 1
    for check in checkGroup.buttons():
        if i == index:
            check.setChecked(True)
        else:
            check.setChecked(False)
        i += 1

def getProcDir():
    '''
    获取程序所在的目录
    '''
    path = sys.path[0]
    if os.path.isdir(path): #脚本文件返回目录
        return path
    elif os.path.isfile(path): #编译后文件返回文件路径
        return os.path.dirname(path)

def saveConfig():
    '''
    向程序目录下的config.json文件写入设置
    '''
    config = {'mode':mode,
            'userPath1':str(userPath1),
            'userPath2':str(userPath2),
            'userPath3':str(userPath3)
            }
    jsonConfig = json.dumps(config, encoding='utf-8', sort_keys=True, indent=2)
    print('save as %s' % jsonConfig)
    with open(getProcDir()+r'/config.json', 'w') as f:
        f.write(jsonConfig)

def loadConfig():
    '''
    从程序目录下的config.json文件获取设置
    '''
    global mode, userPath1, userPath2, userPath3
    try:
        filePath = getProcDir()+r'\config.json'
        if os.path.exists(filePath) and os.path.isfile(filePath):
            with open(filePath, 'r') as f:
                jsonConfig = json.load(f)
            mode = jsonConfig['mode']
            userPath1 = jsonConfig['userPath1']
            userPath2 = jsonConfig['userPath2']
            userPath3 = jsonConfig['userPath3']
        else:
            mode = mode_default
            userPath1 = userPath1_default
            userPath2 = userPath2_default
            userPath3 = userPath3_default
            saveConfig()
    except:
        traceback.print_exc()
        sys.exit(-1)

def setWallpaper(imagePath):
    '''
    根据图像文件路径设置桌面壁纸
    '''
    k = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,"Control Panel\\Desktop",0,win32con.KEY_SET_VALUE)
    win32api.RegSetValueEx(k, "WallpaperStyle", 0, win32con.REG_SZ, "2")
    win32api.RegSetValueEx(k, "TileWallpaper", 0, win32con.REG_SZ, "0")
    win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER,imagePath, 1+2)
    print(imagePath)
    print('--- Wallpaper has been set '+'\"'+imagePath+'\" ---')

def fromBing():
    '''
    模式1下，从必应主页下载每日主页到指定目录并返回图像路径
    '''
    html=urllib.urlopen(url).read()
    if html=='':
        print('html invalid!')
        sys.exit(-1)
    html=html.decode('utf-8')
    reg=re.compile('g_img={url: (.*?),id',re.S) #Cn
    text=re.findall(reg,html)
    text = text[0][1:-1]
    print(text)
    if text.find('http')<0:
        text='http://www.bing.com'+text
    else:
        text = 'http://www.bing.com'+text[text.index('net')+3:]
    urlImage = text
    nameImage=str(urlImage)[str(urlImage).rfind('/')+1:]
    pathImage=userPath1+'\\'+nameImage
    urllib.urlretrieve(urlImage,pathImage)
    print('--- '+nameImage+' was successfully saved ! ---')
    return pathImage

def fromDir():
    '''
    模式2下，从用户指定目录随机选取图像返回其路径
    '''
    for root,dirs,files in os.walk(userPath2):
        if len(files)>0:
            i=random.randint(0,len(files)-1)
            pathImage=os.path.join(userPath2,files[i])
            return pathImage
        else:
            return 'null'        

def fromLockscreen():
    '''
    模式3下，从系统锁屏拷贝图片到目录并返回随机图像地址
    '''
    fileList=[]
    ctimeList=[]
    for root,dirs,files in os.walk(lockscreenPath):
        for fn in files:
            filePath=os.path.join(lockscreenPath, fn)
            ctime=os.path.getctime(filePath)
            size=os.path.getsize(filePath)
            filePathUser=os.path.join(userPath3, fn)
            filePathUser+='.jpg'
            if size>200000:
                shutil.copy(filePath,filePathUser)
                image=Image.open(filePathUser)
                image.close()
                if str(image.size)!='(1920, 1080)':
                    os.remove(filePathUser)
                else:
                    ctimeList.append(ctime)
                    fileList.append(filePathUser)
    index=0
    ctime=0.0
    '''         
    for i in range(0,len(fileList)-1): #set wallpaper by create time
        if float(ctimeList[i])>ctime:
            ctime=float(ctimeList[i])
            index=i
    '''
    index=random.randint(0,len(fileList)-1) #set wallpaper randomly
    pathImage=fileList[index]
    return pathImage

def setFromMode():
    '''
    根据模式返回图像路径
    '''
    if mode==1:
        mkdir(userPath1)
        return fromBing()
    elif mode==2:
        mkdir(userPath2)
        return fromDir()
    elif mode==3:
        mkdir(userPath3)
        return fromLockscreen()   
    else:
        print ('--- Error: mode wrong! ---')
        return 'null'

def initWallpaper():
    '''
    每次启动程序或更改模式执行，更改壁纸
    '''
    global timeri
    print('--- AutoWallpaper initing ---')
    try:
        setWallpaper(setFromMode())
        global aw
        aw.showMessage(u'壁纸已更换', 500, 'Info')
        print('--- init wallpaper successfully ---')
        global lastDate
        lastDate=int(time.strftime(ISOTIMEFORMAT, time.localtime()))
        if timeri.isAlive():
            timeri.cancel() #停止initWallpaper线程级循环
        tc = threading.Thread(target=checkWallpaper)
        tc.start() #启动checkWallpaper
    except:
        print('--- Error: init wallpaper unsuccessfully ---')
        traceback.print_exc()
        print('--- init wallpaper again ---')
        timeri = threading.Timer(5, initWallpaper)
        timeri.start() #启动initWallpaper线程级循环

def checkWallpaper():
    '''
    循环时间检查，隔天时更换壁纸
    '''
    print('--- AutoWallpaper checking ---')
    try:
        thisDate=int(time.strftime(ISOTIMEFORMAT, time.localtime()))
        print(time.ctime())
        global lastDate
        if thisDate!=lastDate:
            print 'not equal'
            thisImage=setFromMode()
            lastImage=win32gui.SystemParametersInfo(win32con.SPI_GETDESKWALLPAPER)
            if thisImage!=lastImage:
                setWallpaper(thisImage)
                lastDate = thisDate
                global aw
                aw.showMessage(u'壁纸已更换', 500, 'Info')
                print('--- wallpaper changed ---')
    except:
        print('--- Error: checkWallpaper unsuccessfully ---')
        traceback.print_exc()
    finally:
        global timerc
        timerc = threading.Timer(10, checkWallpaper)
        timerc.start() #启动checkWallpaper线程级循环

class AW(QtGui.QDialog):
    def __init__(self):
        '''
        构造方法：初始化通知区域图标和菜单，以及设置界面，并初始化设置参数
        '''
        super(AW, self).__init__() # super(classname, self).func() 调用父类方法
        self.procName = 'AutoWallpaper'
        #通知区域图标和菜单
        showItem = QtGui.QAction(self)
        showItem.setText(u'设置')
        showItem.triggered.connect(self.showNormal)
        quitItem = QtGui.QAction(self)
        quitItem.setText(u'退出')
        quitItem.triggered.connect(QtGui.qApp.quit)
        quitItem.triggered.connect(stopCheck)
        self.menu = QtGui.QMenu(self)
        self.menu.addAction(showItem)
        self.menu.addAction(quitItem)
        self.trayIcon = QtGui.QSystemTrayIcon(self)
        self.trayIcon.setContextMenu(self.menu)
        self.icon = QtGui.QIcon(iconPath)       
        self.trayIcon.setIcon(QtGui.QIcon('AW.ico'))
        self.trayIcon.show()
        #设置界面内容
        self.setWindowTitle(self.procName)
        self.resize(500,250)
        self.setWindowIcon(QtGui.QIcon('AW.ico'))
        saveButton = QtGui.QPushButton(u'保存', self)
        saveButton.clicked.connect(self.save)
        quitButton = QtGui.QPushButton(u'取消', self)
        quitButton.clicked.connect(self.cancel)
        self.check_1 = QtGui.QCheckBox(u'从必应主页获取每日壁纸', self)
        self.check_2 = QtGui.QCheckBox(u'从用户文件夹随机获取壁纸', self)
        self.check_3 = QtGui.QCheckBox(u'从系统锁屏中随机获取壁纸', self)
        self.checkGroup = QtGui.QButtonGroup(self)
        self.checkGroup.setExclusive(True)
        self.checkGroup.addButton(self.check_1)
        self.checkGroup.addButton(self.check_2)
        self.checkGroup.addButton(self.check_3)
        label_1 = QtGui.QLabel(u'必应壁纸保存路径')
        label_2 = QtGui.QLabel(u'用户壁纸保存路径')
        label_3 = QtGui.QLabel(u'系统锁屏保存路径')
        self.lineEdit_1 = QtGui.QLineEdit('')
        self.lineEdit_2 = QtGui.QLineEdit('')
        self.lineEdit_3 = QtGui.QLineEdit('')
        search_1 = QtGui.QPushButton(u'浏览')
        search_1.clicked.connect(self.search_1_Dialog)
        search_2 = QtGui.QPushButton(u'浏览')
        search_2.clicked.connect(self.search_2_Dialog)
        search_3 = QtGui.QPushButton(u'浏览')
        search_3.clicked.connect(self.search_3_Dialog)
        hBox_1 = QtGui.QHBoxLayout()
        hBox_1.addWidget(self.check_1)
        hBox_1.addWidget(self.check_2)
        hBox_1.addWidget(self.check_3)
        hBox_2 = QtGui.QHBoxLayout()
        hBox_2.addWidget(label_1)
        hBox_2.addWidget(self.lineEdit_1)
        hBox_2.addWidget(search_1)
        hBox_3 = QtGui.QHBoxLayout()
        hBox_3.addWidget(label_2)
        hBox_3.addWidget(self.lineEdit_2)
        hBox_3.addWidget(search_2)
        hBox_4 = QtGui.QHBoxLayout()
        hBox_4.addWidget(label_3)
        hBox_4.addWidget(self.lineEdit_3)
        hBox_4.addWidget(search_3)
        hBox_5 = QtGui.QHBoxLayout()
        hBox_5.addWidget(saveButton)
        hBox_5.addWidget(quitButton)
        vBox = QtGui.QVBoxLayout()
        vBox.addLayout(hBox_1)
        vBox.addLayout(hBox_2)
        vBox.addLayout(hBox_3)
        vBox.addLayout(hBox_4)
        vBox.addLayout(hBox_5)
        self.setLayout(vBox)
        #设置初始化
        t = threading.Thread(target=loadConfig)
        t.start()
        t.join()
        setCheckedIndex(self.checkGroup, mode)
        self.lineEdit_1.setText(userPath1)
        self.lineEdit_2.setText(userPath2)
        self.lineEdit_3.setText(userPath3)

    def search_1_Dialog(self):
        '''
        文件夹选择对话框；文本框有路径时默认该路径；文本框无路径默认userPath；不选择时保留文本框信息
        '''
        dirPath = self.lineEdit_1.text()
        if not( os.path.exists(dirPath) and os.path.isdir(dirPath) ):
            dirPath = userPath1_default
        dirPath = QtGui.QFileDialog.getExistingDirectory(self, u'文件浏览', dirPath)
        if len(dirPath) > 0:
            self.lineEdit_1.setText(dirPath)

    def search_2_Dialog(self):
        dirPath = self.lineEdit_2.text()
        if os.path.exists(dirPath) and os.path.isdir(dirPath):
            pass
        else:
            dirPath = userPath2_default
        dirPath = QtGui.QFileDialog.getExistingDirectory(self, u'文件浏览', dirPath)
        if len(dirPath) > 0:
            self.lineEdit_2.setText(dirPath)

    def search_3_Dialog(self):
        dirPath = self.lineEdit_3.text()
        if os.path.exists(dirPath) and os.path.isdir(dirPath):
            pass
        else:
            dirPath = userPath3_default
        dirPath = QtGui.QFileDialog.getExistingDirectory(self, u'文件浏览', dirPath)
        if len(dirPath) > 0:
            self.lineEdit_3.setText(dirPath)

    def cancel(self):
        '''
        按下‘取消’时关闭窗口
        '''
        self.close()

    def save(self):
        '''
        按下‘保存’时更新内存中设置变量并写入设置文件；若模式改变，立刻重新初始化
        '''
        global mode, userPath1, userPath2, userPath3
        modeNew = getCheckedIndex(self.checkGroup)
        userPath1 = str(self.lineEdit_1.text()) #text()返回QString，必须换成String
        userPath2 = str(self.lineEdit_2.text())
        userPath3 = str(self.lineEdit_3.text())
        print modeNew, mode
        if modeNew != mode:
            mode = modeNew
            global timeri, timerc
            timeri.cancel()
            timerc.cancel()
            t = threading.Thread(target = initWallpaper)
            t.start()
        t = threading.Thread(target = saveConfig)
        t.start()
        t.join()
        self.close()

    def showMessage(self, message, time=500, mtype='Proc'):
        '''
        显示气泡通知
        '''
        if mtype == 'Proc':
            icon = QtGui.QSystemTrayIcon.MessageIcon()
        elif mtype == 'Error':
            icon = QtGui.QSystemTrayIcon.Critical
        elif mtype == 'Warn':
            icon = QtGui.QSystemTrayIcon.Warning
        elif mtype == 'Info':
            icon = QtGui.QSystemTrayIcon.Information      
        self.trayIcon.showMessage(
            self.procName, message, icon, time)


def main():
    global timeri, timerc
    timeri = threading.Timer(0, initWallpaper)
    timerc = threading.Timer(0, checkWallpaper)
    print('--- AutoWallpaper start ---')
    app = QtGui.QApplication(sys.argv)
    QtGui.QApplication.setQuitOnLastWindowClosed(False)
    global aw
    aw = AW()
    ti = threading.Thread(target = initWallpaper)
    ti.start()  #启动initWallpaper
    sys.exit(app.exec_())

if __name__ == '__main__': 
    main()
