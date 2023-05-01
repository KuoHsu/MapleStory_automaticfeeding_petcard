import time
import pyautogui
import win32gui
import win32api
import win32con
import keyboard

hwn = win32gui.FindWindow("MapleStoryClassTW","MapleStory")
win32gui.SetForegroundWindow(hwn)

time.sleep(1)

rect = win32gui.GetWindowRect(hwn)
x,y = rect[0],rect[1]
sx,sy = rect[2]-rect[0],rect[3]-rect[1]
print(x)
print(y)
print(sx)
print(sy)

#win32gui.MoveWindow(hwn,x,y,sx,sy,True)



pyautogui.screenshot(r'./img/screenshot.png',region=(x+750,y+430,50,30))

