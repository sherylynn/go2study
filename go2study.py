time_count=0
time_space=90
time_pdf=180
block_title="微信firefox"
allow_title="通关宝典考试软件选择命令提示符"
allow_pdf_title="福昕"
test_title="Sublime"
alert_message_list=["还在玩,快学习,快起床,快想想你要开玛莎拉蒂,",
"我们有时候选择去做一件事情,不是因为它有多简单,而是因为它很难,我们突破自己,挑战自己,变得更强",
"可能刚刚你走神了,可能刚刚你在忙其他事,不要紧,现在这一分钟就很重要,来做题吧,现在开始就不会迟",
"想一想,还有人是仰望着你,朝着你走过的路前进,你当做好表率,继续当好榜样,引领别人前行",
"手机很好玩,微博很好刷,与别人聊天也很有趣,但是你是审计从业者,如果你对注册会计师的证,都觉得无趣,那你这样做,应该吗",
"有时候需要放松,有时候需要调整,你休息好了吗,现在开始,继续前进前进"]
#pip3 install pynput==1.6.8

from pynput import mouse
from datetime import datetime
from threading import Timer

#pip3 install pywin32
import win32api
import win32con
import win32gui as w
from datetime import date,timedelta

import random

import pyttsx3 
# 模块初始化

#pip3 install pycaw
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))


import os

def sound_max():
    #WM_APPCOMMAND = 0x319
    #APPCOMMAND_VOLUME_MAX = 0x0a
    #APPCOMMAND_VOLUME_MIN = 0x09
    #win32api.SendMessage(-1, WM_APPCOMMAND, 0x30292, APPCOMMAND_VOLUME_MAX * 0x10000)
    
    if volume.GetMute() == 1 :
        """
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            process_volume = session.SimpleAudioVolume
            if session.Process and (session.Process.name() == "go2study.exe" or session.Process.name() == "python.exe" ):
                process_volume.SetMute(0, None)
        """
        volume.SetMute(0,None)
    volume.SetMasterVolumeLevel(0.0, None)

def alert_message():
    return alert_message_list[random.randint(0,len(alert_message_list)-1)]


def killexe(exe):
    os.system('start /B taskkill /F /IM '+exe)
    #加入 start /B 去除黑色窗口 失效 ,还是去掉这个功能吧

def remain_daytime():
    d1=date(2021, 8, 27)
    d2=date.today()
    return (d1-d2).days


def alert():
    sound_max()
    #win32api.MessageBox(None,"快去看书","要命啦!!!!!!!!!!!!!!!!!!!!!",win32con.MB_OK)
    engine = pyttsx3.init()
    # 设置要播报的Unicode字符串
    engine.say(alert_message()+"离考试倒数天数只有%d天了"%remain_daytime()) 
    # 等待语音播报完毕 
    engine.runAndWait()


# 打印时间函数
def countTime(inc):
    global time_count
    global time_space
    global time_pdf
    time_count+=1
    
    #killexe("按键精灵2014.exe")
    print(time_count)

    title = w.GetWindowText (w.GetForegroundWindow())

    time_deadline=time_space
    if allow_pdf_title in title or test_title in title:
        time_deadline=time_pdf
        
    if time_count >= time_deadline:
        alert()
    #print("开始下一个计时")
    t = Timer(inc, countTime, (inc,))
    t.start()

#循环1s计时
countTime(1)

def clear_time_count():
    global block_title
    global time_count
    title = w.GetWindowText (w.GetForegroundWindow())
    #如果不是被禁止的窗口,点击鼠标会清空
    #if title not in block_title:
    #    time_count=0
    print(title)
    if title !="" and (title in allow_title or allow_pdf_title in title or test_title in title):
        time_count=0

def on_move(x, y):
    clear_time_count()

def on_click(x, y, button, pressed):
    clear_time_count()

def on_scroll(x, y, dx, dy):
    clear_time_count()

# Collect events until released
with mouse.Listener(
        on_move=on_move,
        on_click=on_click,
        on_scroll=on_scroll) as listener:
    listener.join()

# ...or, in a non-blocking fashion:
listener = mouse.Listener(
    on_move=on_move,
    on_click=on_click,
    on_scroll=on_scroll)
listener.start()