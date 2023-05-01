import time
import pyautogui
import win32gui
import win32api
import win32con
import keyboard

hwn = win32gui.FindWindow("MapleStoryClassTW","MapleStory")




class Script:
    def __init__(self,hwn):
        self.exit_flag = False
        self.eat_flag = False
        self.maplestroy_hwn = hwn
        self.package_0_0_disx = 25
        self.package_0_0_disy = 95
        self.package_blank = 43
        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        keyboard.add_hotkey("f1",lambda: self.__eat_status__())
        keyboard.add_hotkey("f2",lambda: self.__exit__())


    def __eat_status__(self):
        self.eat_flag = not self.eat_flag
        return



    def __exit__(self):
        self.exit_flag = True
        return

    def __eat__(self):
        rect = win32gui.GetWindowRect(self.maplestroy_hwn)
        x,y = rect[0],rect[1]


        #檢查有無開背包
        ca = pyautogui.locateOnScreen(r'./img/consume_active.png',region=(x+35,y+50,40,30), grayscale=True, confidence=0.75)
        cua = pyautogui.locateOnScreen(r'./img/consume_unactive.png',region=(x+35,y+50,40,30), grayscale=True, confidence=0.75)
        if(ca == None and cua == None):
            print("請先打開背包!")
            self.eat_flag = False
            return



        #開啟消耗欄
        print("開始吃背包中的萌獸。")

        pyautogui.moveTo(x+45,y+65)
        pyautogui.click(clicks = 1)

        package_0_0_posx = x + self.package_0_0_disx
        package_0_0_posy = y + self.package_0_0_disy
        package_blank = self.package_blank
        row = 7
        col = 15

        check = 0
        pet = 0

        while(col >= 0 and self.eat_flag == True and self.exit_flag != True):
            while(row >= 0 and self.eat_flag == True and self.exit_flag != True):
                

                px = package_0_0_posx + col * package_blank
                py = package_0_0_posy + row * package_blank
                pyautogui.moveTo(px,py)
                reg = pyautogui.locateOnScreen(r'./img/reg_cute_pet_1.png',region=(px+15,py+130,40,20), grayscale=True, confidence=0.75)
                is_cute_pet = True if reg != None else False

                if(is_cute_pet):
                    pyautogui.click(clicks = 2)

                    pyautogui.screenshot(r'./img/screenshot.png',region=(x+750,y+430,50,30))
                    fullreg = pyautogui.locateOnScreen(r'./img/full_msg_ok.png',region=(x+750,y+430,50,30), grayscale=True, confidence=0.75)
                    if(fullreg != None):
                        print("萌獸本已滿。")
                        pyautogui.moveTo(x+760,y+435)
                        pyautogui.click(clicks = 1)
                        self.eat_flag = False
                        break
                    else:
                        pet += 1
                        print("背包第" + str(row+1) + "列第" + str(col+1) + "欄為萌獸，已吃掉。")
                    
                row -= 1
                time.sleep(0.1)
                check += 1
            row = 7
            col -= 1
        
        print("已吃掉" + str(pet) + "隻萌獸。")
        if(check != 128):
            print("吃萌獸程序已中斷。")
        else:
            print("背包中的萌獸已經吃完了。")
        self.eat_flag = False
    



    def run(self):
        print_msg = False
        while(self.exit_flag != True):
            if(not print_msg):
                print("----- F1: 吃背包中的萌獸  F2: 離開程式 -----")
                print_msg = True

            if(self.eat_flag == True):
                self.__eat__()
                print_msg = False
            time.sleep(0.5)


        


script_run = Script(hwn)  
script_run.run()


