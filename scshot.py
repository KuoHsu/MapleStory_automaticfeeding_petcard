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
xx = x + 725
yy = y + 325
dx = 60
dy = 75



pyautogui.screenshot(r'./img/screenshot.png',region=(x+700,y+280,40,20))

row = 4

while(1):
    pyautogui.moveTo(xx+200,yy)
    pyautogui.scroll(-400)
    print("scroll!")
    a = pyautogui.locateOnScreen(r'./img/scroll_reg.png',region=(x+690,y+270,60,40))
    if(a != None):
        print("new row")
        row += 1
    else:
        print("not new row.")
        break
    time.sleep(1)

print("row count: " + str(row))
pyautogui.scroll(8000)


'''
pyautogui.moveTo(xx+200,yy)



pyautogui.scroll(-400)
time.sleep(2)




pyautogui.scroll(-400)
time.sleep(2)
pyautogui.scroll(-400)

time.sleep(3)
pyautogui.scroll(8000)
#pyautogui.scroll(-200)
#pyautogui.scroll(-200)
'''
'''
for r in range(0,4):
    for c in range(0,4):
        px = xx + c * dx - 15
        py = yy + r * dy - 25

        pyautogui.screenshot(r'./img/screenshot' + str(r) + str(c) + '.png',region=(px,py,40,40))
        a = pyautogui.locateOnScreen(r'./img/card_null_reg.png',region=(px,py,40,40))
        if(a != None):
            print(str(r) + str(c) + 'isnull')
'''




while(0):
    a = pyautogui.position()
    print(a.x - x, a.y-y)
    time.sleep(1)

