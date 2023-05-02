import time
import pyautogui
import win32gui
import win32api
import win32con
import keyboard

import clipboard

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


'''
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
clipboard.copy("保留對象")

for r in range(0,4):
    for c in range(0,4):
        px = xx + c * dx 
        py = yy + r * dy 
        
        nx = x + 565
        ny = y + 420
        pyautogui.moveTo(px,py)
        pyautogui.rightClick()
        time.sleep(0.1)
        pyautogui.moveTo(nx,ny)
        pyautogui.click(clicks = 1)
        time.sleep(0.1)
        keyboard.press("ctrl+v")
        time.sleep(0.1)
        pyautogui.moveTo(x+725,y+445)
        pyautogui.click(clicks = 1)
        time.sleep(0.1)
        
        pyautogui.moveTo(px,py)
        time.sleep(0.01)
        

for r in range(0,4):
    for c in range(0,4):
        px = xx + c * dx 
        py = yy + r * dy 
        pyautogui.moveTo(px,py)
        time.sleep(0.1)
        pyautogui.screenshot(r'./img/screenshot' + str(r) + str(c) + '.png',region=(px,py,250,300))
        a = pyautogui.locateOnScreen(r'./img/keep_reg.png',region=(px,py,250,300))
        if(a != None):
            print(str(r+1) + str(c+1) + '為保留對象')





count = 0
while(0):
    count+=1
    #a = pyautogui.position()
    #r,g,b = pyautogui.pixel(a.x -5,a.y - 5)
    #print(a.x - x, a.y-y, r,g,b)
    #pyautogui.screenshot(r'./img/sc' + str(count) +'.png',region = (a.x, a.y,30,30))
    print("\rloading" + "." * count,end="")
    time.sleep(1)

