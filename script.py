import time
import pyautogui
import win32gui
import win32api
import win32con
import keyboard
import queue
import clipboard

hwn = win32gui.FindWindow("MapleStoryClassTW","MapleStory")

if(hwn == 0 ):
    print("找不到楓之谷遊戲視窗，請確認是否有開啟楓之谷。")
    exit()


class Script:
    def __init__(self,hwn):

        self.exit_flag = False
        self.eat_flag = False
        self.upgrade_flag = False
        self.keep_flag = False
        self.maplestroy_hwn = hwn

        
        #吃背包萌獸座標設定(1366*768)
        self.package_0_0_disx = 25 #背包第一格在遊戲視窗的x座標
        self.package_0_0_disy = 95 #背包第一格在遊戲視窗的y座標
        self.package_blank = 43 #背包每格位移距離 列欄位移值相同 (x=y)



        #吃收藏本萌獸座標設定(1366*768)
        self.book_class_y = 270 #潛能等級y座標
        self.book_class_0_x = 765 #普通等級x座標
        self.book_class_blank_x = 40 #潛能等級x座標位移植 (同列，無y座標位移植)
        self.book_0_0_disx = 725 #第一隻萌獸在遊戲視窗的x座標
        self.book_0_0_disy = 325 #第一隻萌獸在遊戲視窗的y座標
        self.book_blank_row = 75 #萌獸每列(y)位移值
        self.book_blank_col = 60 #萌獸每欄(x)位移值
        self.book_eat_confirm_x = 555 #萌獸獻祭確認鍵x座標
        self.book_eat_confirm_y = 555 #萌獸獻祭確認鍵y座標

        self.book_empty_screen_disx = -15 #空位截圖起始座標x位移值
        self.book_empty_screen_disy = -25 #空位截圖起始座標y位移值

        self.book_lock_screen_disx = 0 #上鎖截圖起始座標x位移值
        self.book_lock_screen_disy = -50 #上鎖截圖起始座標y位移值

        
        keyboard.add_hotkey("f1",lambda: self.__eat_status__())
        keyboard.add_hotkey("f2",lambda: self.__upgrade_status__())
        keyboard.add_hotkey("f3",lambda: self.__keep_status__())
        keyboard.add_hotkey("f4",lambda: self.__exit__())


    def __eat_status__(self):
        self.eat_flag = not self.eat_flag
        if(self.eat_flag):
            self.upgrade_flag = False
            self.keep_flag = False
        return

    def __upgrade_status__(self):
        self.upgrade_flag = not self.upgrade_flag
        if(self.upgrade_flag):
            self.eat_flag = False
            self.keep_flag = False
        return

    def __keep_status__(self):
        self.keep_flag = not self.keep_flag
        if(self.keep_flag):
            self.eat_flag = False
            self.upgrade_flag = False

        return

    def __exit__(self):
        self.exit_flag = True
        return

    def __add__(self):
        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        rect = win32gui.GetWindowRect(self.maplestroy_hwn)
        x,y = rect[0],rect[1]


        #檢查有無開背包
        ca = pyautogui.locateOnScreen(r'./img/consume_active_reg.png',region=(x+35,y+50,40,30), grayscale=True, confidence=0.75)
        cua = pyautogui.locateOnScreen(r'./img/consume_unactive_reg.png',region=(x+35,y+50,40,30), grayscale=True, confidence=0.75)
        if(ca == None and cua == None):
            self.__print_msg__("請先打開背包並將背包置於遊戲畫面左上角！")
            self.eat_flag = False
            return



        #開啟消耗欄
        pyautogui.moveTo(x+45,y+65)
        pyautogui.click(clicks = 1)



        self.__print_msg__("開始將背包中的萌獸新增至收藏本。")
        package_0_0_x = x + self.package_0_0_disx
        package_0_0_y = y + self.package_0_0_disy
        package_blank = self.package_blank

        #從背包最後面開始新增
        index = 127

        row = 7
        col = 15

        add_count = 0

        while(index >= 0 and self.eat_flag and not self.exit_flag):
            region = index // 32
            col_in_region = index % 4
            row = (index- region * 32) // 4
            col = region * 4 + col_in_region
            px = package_0_0_x + col * package_blank
            py = package_0_0_y + row * package_blank
            pyautogui.moveTo(px,py)

            #判斷是否為萌獸
            card_reg = pyautogui.locateOnScreen(r'./img/card_reg_1366768.png',region=(px+15,py+130,40,20), grayscale=True, confidence=0.75)
            is_card = True if card_reg != None else False

            if(is_card):
                pyautogui.click(clicks = 2)

                #確認是否滿了(跳確認視窗)
                fullreg = pyautogui.locateOnScreen(r'./img/msg_ok_reg.png',region=(x+750,y+430,50,30), grayscale=True, confidence=0.75)
                if(fullreg != None):
                    self.__print_msg__("萌獸收藏本已經滿了。")
                    pyautogui.moveTo(x+760,y+435)
                    pyautogui.click(clicks = 1)
                    self.eat_flag = False
                    break
                else:
                    add_count += 1
                    self.__print_msg__("背包第" + str(row+1) + "列第" + str(col+1) + "欄為萌獸，已新增至收藏本。")
            index -= 1

        
        self.__print_msg__("已新增" + str(add_count) + "隻萌獸。")
        if(index != -1):
            self.__print_msg__("新增萌獸程序已中止。")
        else:
            self.__print_msg__("背包中的萌獸已新增完畢。")
        self.eat_flag = False
    

    def __upgrade__(self):
        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        rect = win32gui.GetWindowRect(self.maplestroy_hwn)
        x,y = rect[0],rect[1]

        time.sleep(0.05)

        #確認有開萌獸收藏本
        bookreg = pyautogui.locateOnScreen(r'./img/petbook_reg.png',region=(x+915,y+220,50,30), grayscale=True, confidence=0.75)
        if(bookreg == None):
            self.__print_msg__("請先打開萌獸收藏本並放上要強化的萌獸! 萌獸收藏本須放在預設位置，若有移動到請重新打開")
            self.upgrade_flag = False
            return

        #確認有放要強化的萌獸
        setreg = pyautogui.locateOnScreen(r'./img/pet_is_set_reg.png',region=(x+440,y+435,20,20), grayscale=True, confidence=0.75)
        if(setreg == None):
            self.__print_msg__("請先放上要強化的萌獸!")
            self.upgrade_flag = False
            return

        
        

        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        rect = win32gui.GetWindowRect(self.maplestroy_hwn)
        x,y = rect[0],rect[1]

        eat_count = 0

        #keep_cards = self.__upgrade_keep__()


        class_0_x = x + self.book_class_0_x
        class_y = y + self.book_class_y
        class_col = 0
        class_name = ['普通','特殊','稀有']

        #只吃普通、特殊跟稀有的卡片
        while(class_col < 3 and self.upgrade_flag and not self.exit_flag):
            self.__print_msg__('切換至 ' + class_name[class_col] + ' 等級')

            px = class_0_x + class_col * self.book_class_blank_x
            py = class_y
            pyautogui.moveTo(px,py)
            pyautogui.click(clicks = 1)
            time.sleep(0.05)

            #吃完當前等級的所有萌獸
            eat_count += self.__upgrade_eat__(x,y)

            class_col += 1

        self.__print_msg__("所有能吃的萌獸都吃完了，總共吃了" + str(eat_count) + "隻萌獸。")
        pyautogui.moveTo(x + self.book_0_0_disx - 200,y + self.book_0_0_disy + 50)

        self.upgrade_flag = False



    def __upgrade_eat__(self,x,y):

        eat_count = 0
        
        confirm_x = x + self.book_eat_confirm_x
        confirm_y = y + self.book_eat_confirm_y

      


        end_flag = False

        while(not end_flag and self.upgrade_flag and not self.exit_flag):
            pyautogui.moveTo(x + self.book_0_0_disx + 200,y + self.book_0_0_disy)
            pyautogui.scroll(8000)
            time.sleep(0.05)
            index = 0
            count = 0
            scroll_count = 0


            #吃之前先判斷要升級的萌獸是否滿等(沒有顯示LV)
            maxlv_reg = pyautogui.locateOnScreen(r'./img/pet_is_set_reg.png',region=(x+440,y+435,20,20), grayscale=True, confidence=0.75)
            if(maxlv_reg == None):
                self.__print_msg__("萌獸已強化至最高等。")
                self.upgrade_flag = False
                break


            #將4個素材放進去
            while(count < 4 and self.upgrade_flag and not self.exit_flag):
                pyautogui.moveTo(x + self.book_0_0_disx + 200,y + self.book_0_0_disy)
                time.sleep(0.05)

                row = index // 4
                col = index % 4

                if(row > 3 and col == 1):
                    pyautogui.moveTo(x + self.book_0_0_disx + 200,y + self.book_0_0_disy)
                    pyautogui.scroll(-400)
                    time.sleep(0.05)
                    new_row_reg = pyautogui.locateOnScreen(r'./img/scroll_reg.png',region=(x+690,y+270,60,40))
                    #剛好16張會觸發
                    if(new_row_reg == None):
                        break
                    else:
                        scroll_count += 1

                px = x + self.book_0_0_disx + col * self.book_blank_col
                py = y + self.book_0_0_disy + (row - scroll_count) * self.book_blank_row


                #判斷當前位置是否有卡片，沒有代表已經尋訪完所有卡片
                empty_x = px + self.book_empty_screen_disx #(px-15)
                empty_y = py + self.book_empty_screen_disy #(py-25)
                #pyautogui.screenshot(r'./img/shot.png',region=(empty_x,empty_y,40,40))
                null_reg = pyautogui.locateOnScreen(r'./img/pet_empty_reg.png',region=(empty_x,empty_y,40,40))
                if(null_reg != None):
                    end_flag = True
                    break

                #判斷當前位置的卡片是否為上鎖，上鎖就找下一個位置，沒上鎖代表可以吃，將順序加入到序列
                lock_x = px + self.book_lock_screen_disx #(px+0)
                lock_y = py + self.book_lock_screen_disy #(py-50)
                lock_reg = pyautogui.locateOnScreen(r'./img/lock_reg.png',region=(lock_x,lock_y,30,30), grayscale=True, confidence=0.75)
                if(lock_reg != None):
                    self.__print_msg__(str(row+1) + '列'+ str(col+1) + '欄為上鎖萌獸')
                    index += 1
                    continue
                 

                #判斷是否為保留對象
                pyautogui.moveTo(px,py)
                time.sleep(0.05)
                keep_reg = pyautogui.locateOnScreen(r'./img/keep_reg.png',region=(px,py,250,300), grayscale=True, confidence=0.75)
                pyautogui.screenshot(r'./img/screenshot' + str(row) + str(col) + '.png',region=(px,py,250,300))
                if(keep_reg != None):
                    self.__print_msg__(str(row+1) + '列'+ str(col+1) + '欄為要保留的萌獸')
                    index += 1
                    continue
                
         

                pyautogui.click(clicks = 2)
                time.sleep(0.05)
                index += 1
                count += 1

            #點獻祭台下面的確認
            pyautogui.moveTo(confirm_x,confirm_y)
            pyautogui.click(clicks = 1)
            time.sleep(0.05)

            #點跳出來的確認視窗
            pyautogui.moveTo(x+725,y+445)
            pyautogui.click(clicks = 1)
            time.sleep(0.05)
            


            fullreg = pyautogui.locateOnScreen(r'./img/msg_ok_reg.png',region=(x+750,y+430,50,30), grayscale=True, confidence=0.75)
            if(fullreg != None):
                self.__print_msg__("沒楓幣了。")
                pyautogui.moveTo(x+760,y+435)
                pyautogui.click(clicks = 1)
                self.upgrade_flag = False
                break
            else:
                eat_count += count
                #等動畫跑完
                an_count = 0
                while(1):
                    r,g,b = pyautogui.pixel(x+450,y+245)
                    if(r > 40 and g > 40 and b > 40):
                        break
                    self.__print_msg__("跑動畫中" + "."  * an_count)
                    an_count+=1
                    time.sleep(0.2)
                
            time.sleep(0.2)


        return eat_count


    
    def __set_keep__(self):
        keep_card = []
        class_name = ['普通','特殊','稀有']


        for i in range(0,3):
            if(not self.keep_flag):
                self.__print_msg__("設定保留萌獸程序中止。")
                return


            classn = class_name[i]
            keep_in_classn = []
            keep_count = -1
            win32gui.SetForegroundWindow(self.maplestroy_hwn)
            rect = win32gui.GetWindowRect(self.maplestroy_hwn)
            x,y = rect[0],rect[1]

            cx = x + self.book_class_0_x + i * self.book_class_blank_x 
            cy = y + self.book_class_y

            pyautogui.moveTo(cx,cy)
            pyautogui.click()
            time.sleep(0.05)

            while(keep_count == -1):
                keep_count = int(input("請輸入要保留的 " + classn + " 等級萌獸數量："))
                if(keep_count == -1):
                    self.__print_msg__("請輸入 0 或 0 以上的數字。")


            for c in range(1,keep_count + 1):
                if(not self.keep_flag):
                    self.__print_msg__("設定保留萌獸程序中斷。")
                    break


                keep_c_row = -1
                keep_c_col = -1
                while(keep_c_row == -1):
                    keep_c_row = int(input("第 " + str(c) + "隻需保留的萌獸於 " + class_name[i] + " 等級頁面中所處的列數 (>=1)："))
                    keep_c_row = -1 if(keep_c_row< 1) else keep_c_row
                    if(keep_c_row == -1):
                        self.__print_msg__("輸入錯誤! 請輸入大於等於 1 的數字。\n")
                while(keep_c_col == -1):
                    keep_c_col = int(input("第 " + str(c) + "隻需保留的萌獸於 " + class_name[i] + " 等級頁面中所處的欄數 (1~4)："))
                    keep_c_col = -1 if(keep_c_col< 1 or keep_c_col > 4) else keep_c_col
                    if(keep_c_col == -1):
                        self.__print_msg__("輸入錯誤! 請輸入介於 1~4 之間的數字。\n")
                keep_in_classn.append((keep_c_row,keep_c_col))
            keep_card.append(keep_in_classn)

        if(not self.keep_flag):
            self.__print_msg__("設定保留萌獸程序中止。")
            return

        #確認有開萌獸收藏本
        while(1):
            win32gui.SetForegroundWindow(self.maplestroy_hwn)
            bookreg = pyautogui.locateOnScreen(r'./img/petbook_reg.png',region=(x+915,y+220,50,30), grayscale=True, confidence=0.75)
            if(bookreg == None):
                self.__print_msg__("請先打開萌獸收藏本！打開收藏本後程序將繼續執行\r")
                self.upgrade_flag = False
                time.sleep(1)
            else:
                break
        
        count = 0


        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        rect = win32gui.GetWindowRect(self.maplestroy_hwn)
        x,y = rect[0],rect[1]


        clipboard.copy("保留")

        nx = x + 565
        ny = y + 420

        okx = x + 725
        oky = y + 445

        print(keep_card)

        for i in range(0,3):
            if(not self.keep_flag):
                break
            classn = class_name[i]
            keep_in_classn = keep_card[i]
            #print(keep_in_classn)
            if(len(keep_in_classn) == 0):
                continue
            
            cx = x + self.book_class_0_x + i * self.book_class_blank_x
            cy = y + self.book_class_y

            pyautogui.moveTo(cx,cy)
            pyautogui.click()
            time.sleep(0.05)


            self.__print_msg__("開始設定 " + classn +" 等級要保留的萌獸。")
            keep_in_classn.sort()
            scroll_count = 0
            for rc in keep_in_classn:
                if(not self.keep_flag):
                    break
                row = rc[0] - 1
                col = rc[1] - 1
                
                while((row - scroll_count) > 3):
                    pyautogui.moveTo(x + self.book_0_0_disx + 200,y + self.book_0_0_disy)
                    pyautogui.scroll(-400)
                    time.sleep(0.05)
                px = x + self.book_0_0_disx + col * self.book_blank_col
                py = y + self.book_0_0_disy + row * self.book_blank_row
                pyautogui.moveTo(px,py)
                time.sleep(0.05)
                pyautogui.rightClick()
                time.sleep(0.05)
                pyautogui.moveTo(nx,ny)
                pyautogui.click(clicks = 1)
                time.sleep(0.05)
                keyboard.press("ctrl+v")
                time.sleep(0.05)
                pyautogui.moveTo(okx,oky)
                pyautogui.click(clicks = 1)
                count += 1
                
        
        if(not self.keep_flag):
            self.__print_msg__("設定保留萌獸程序中止。") 
        self.__print_msg__("保留對象設定完成，共設定了 " + str(count) + " 個要保留的萌獸。")
        self.keep_flag = False
        return

    def __print_info_msg__(self):
        print("╔══════════════════════════════════════════════════════════════════════════════════════════════════")
        print("║  　　　　　　　　　　！！重要！！　　　　　　　　　　　           ")
        print("║ 強化前請先將「要強化的」與「要保留的」萌獸名稱設定為「保留」。")
        print("║ 「已上鎖」的萌獸不必設定為保留。")
        print("║ 可使用程式中的功能設定要保留的萌獸。　　　　　　　              ")
        print("╟──────────────────────────────────────────────────────────────────────────────────────────────────")
        print("║ 按下對應按鍵可執行對應程序，按下按鍵前先確認楓之谷內對應按鍵上沒有放重要的東西 (按下會一同觸發)。")
        print("║ 程序執行中可再按一次對應按鍵中止程序，程序執行中按 F4 可直接關閉程式。")
        print("║ 例：按 F2 執行升級萌獸程序，執行過程中可再按下一次 F2 中止該程序，按下 F4 則會中止程序並關閉程式。")
        print("╚══════════════════════════════════════════════════════════════════════════════════════════════════")


    def __print_msg__(self,msg):
        print("\r                                   ",end="")
        print("\r" + msg, end= "")


    def run(self):
        print_msg = False
        self.__print_info_msg__()
        while(not self.exit_flag):
            if(not print_msg):
                print("\n-----\nF1: 吃背包中的萌獸\nF2: 升級萌獸\nF3: 設定要保留的萌獸\nF4: 離開程式\n-----\n")
                print_msg = True
            if(self.eat_flag):
                self.__add__()
                print_msg = False
            elif(self.upgrade_flag):
                self.__upgrade__()
                print_msg = False
            elif(self.keep_flag):
                self.__set_keep__()
                print_msg = False
            time.sleep(0.5)


        


script_run = Script(hwn)  
script_run.run()


