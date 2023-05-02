import time
import pyautogui
import win32gui
import win32api
import win32con
import keyboard
import queue

hwn = win32gui.FindWindow("MapleStoryClassTW","MapleStory")

if(hwn == 0 ):
    print("找不到楓之谷遊戲視窗，請確認是否有開啟楓之谷。")
    exit()


class Script:
    def __init__(self,hwn):
        self.exit_flag = False
        self.eat_flag = False
        self.upgrade_flag = False
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

        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        keyboard.add_hotkey("f1",lambda: self.__eat_status__())
        keyboard.add_hotkey("f2",lambda: self.__upgrade_status__())
        keyboard.add_hotkey("f3",lambda: self.__exit__())


    def __eat_status__(self):
        self.eat_flag = not self.eat_flag
        return

    def __upgrade_status__(self):
        self.upgrade_flag = not self.upgrade_flag
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
            print("請先打開背包!")
            self.eat_flag = False
            return



        #開啟消耗欄
        pyautogui.moveTo(x+45,y+65)
        pyautogui.click(clicks = 1)



        print("開始將背包中的萌獸新增至收藏本。")
        package_0_0_x = x + self.package_0_0_disx
        package_0_0_y = y + self.package_0_0_disy
        package_blank = self.package_blank

        #從背包最後面開始新增
        row = 7
        col = 15

        check = 0
        add_count = 0

        while(col >= 0 and self.eat_flag and not self.exit_flag ):
            while(row >= 0 and self.eat_flag and not self.exit_flag):
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
                        print("萌獸收藏本已經滿了。")
                        pyautogui.moveTo(x+760,y+435)
                        pyautogui.click(clicks = 1)
                        self.eat_flag = False
                        break
                    else:
                        add_count += 1
                        print("背包第" + str(row+1) + "列第" + str(col+1) + "欄為萌獸，已新增至收藏本。")
                row -= 1
                time.sleep(0.05)
                check += 1
            row = 7
            col -= 1
        
        print("已新增" + str(add_count) + "隻萌獸。")
        if(check != 128):
            print("新增萌獸程序已中斷。")
        else:
            print("背包中的萌獸已新增完畢。")
        self.eat_flag = False
    

    def __upgrade__(self):
       





        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        rect = win32gui.GetWindowRect(self.maplestroy_hwn)
        x,y = rect[0],rect[1]


        #確認有開萌獸收藏本
        bookreg = pyautogui.locateOnScreen(r'./img/petbook_reg.png',region=(x+915,y+220,50,30), grayscale=True, confidence=0.75)
        if(bookreg == None):
            print("請先打開萌獸收藏本並放上要強化的萌獸! 萌獸收藏本須放在預設位置，若有移動到請重新打開")
            self.upgrade_flag = False
            return

        #確認有放要強化的萌獸
        setreg = pyautogui.locateOnScreen(r'./img/pet_is_set_reg.png',region=(x+440,y+435,20,20), grayscale=True, confidence=0.75)
        if(setreg == None):
            print("請先放上要強化的萌獸!")
            self.upgrade_flag = False
            return

        
        print("該程序會將沒有上鎖的普通、特殊與稀有萌獸\"全部\"當作素材吃掉，")
        print("若有需要排除在外的萌獸請先上鎖。")
        print("(不建議使用該程序餵普通、特殊與稀有等級的萌獸，已經試過了，強化對象會被吃掉。)")
        confirm = str(input("確定執行自動吃萌獸程序，請輸入「Yes」(大小寫需相符): "))
        if (confirm != "Yes")
            return

        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        rect = win32gui.GetWindowRect(self.maplestroy_hwn)
        x,y = rect[0],rect[1]

        eat_count = 0

        except_cards = self.__upgrade_except__()


        class_0_x = x + self.book_class_0_x
        class_y = y + self.book_class_y
        class_col = 0
        class_name = ['普通','特殊','稀有']

        #只吃普通、特殊跟稀有的卡片
        while(class_col < 3 and self.upgrade_flag and not self.exit_flag):
            print('切換至 ' + class_name[class_col] + ' 等級')

            px = class_0_x + class_col * self.book_class_blank_x
            py = class_y
            pyautogui.moveTo(px,py)
            pyautogui.click(clicks = 1)
            time.sleep(0.05)

            #except_card_inclass = except_cards[class_col]
            ary = []

            #吃完當前等級的所有萌獸
            eat_count += self.__upgrade_eat__(x,y,ary)

            class_col += 1

        print("所有能吃的萌獸都吃完了，總共吃了" + str(eat_count) + "隻萌獸。")
        pyautogui.moveTo(x + self.book_0_0_disx - 200,y + self.book_0_0_disy + 50)

        self.upgrade_flag = False


    def __upgrade_except__(self):

        except_card = []
        class_name = ['普通','特殊','稀有']


        print("請先輸入排除萌獸的潛能等級、列數與欄數資料。")
        print("排除萌獸包含強化對象、進化合成素材，已鎖定的萌獸不必輸入。")
        print("列數為由上往下遞增(1~4列)，欄數為由左往右遞增(1~4欄)")
        print("例如：要強化的萌獸潛能等級是特殊等級，該卡片在特殊等級頁面中所處的列數與欄數分別為第1列與第3欄，則於列數輸入1、欄數輸入3。")
        

        for i in range(0,3):
            classn = class_name[i]
            except_in_classn = []
            except_count = -1
            while(except_count == -1):
                except_count = int(input("請輸入要排除的 " + classn + " 等級萌獸數量："))
                if(except_count == -1):
                    print("請輸入0或0以上的數字。")
            for c in range(1,except_count + 1):
                except_c_row = -1
                except_c_col = -1
                while(except_c_row == -1):
                    except_c_row = int(input("請輸入第 " + str(c) + "隻萌獸於 " + class_name[i] + " 等級頁面中所處的列數(>=1)："))
                    except_c_row = -1 if(except_c_row< 1) else except_c_row
                    if(except_c_row == -1):
                        print("輸入錯誤! 請輸入大於等於1的數字。\n")
                while(except_c_col == -1):
                    except_c_col = int(input("請輸入第 " + str(c) + "隻萌獸於 " + class_name[i] + " 等級頁面中所處的欄數(1~4)："))
                    except_c_col = -1 if(except_c_col< 1 or except_c_col > 4) else except_c_col
                    if(except_c_col == -1):
                        print("輸入錯誤! 請輸入介於1~4之間的數字。\n")
                except_in_classn.append((except_c_row,except_c_col))
            except_card.append(except_in_classn)

        win32gui.SetForegroundWindow(self.maplestroy_hwn)
        
        return except_card






    def __upgrade_eat__(self,x,y,except_cards):

        eat_count = 0
        
        confirm_x = x + self.book_eat_confirm_x
        confirm_y = y + self.book_eat_confirm_y

        row = 0
        col = 0

        eat_index_queue =  queue.Queue()
    
        count = 0
        index = 0
        c_index = 0
        scroll_count = 0

        pyautogui.moveTo(x + self.book_0_0_disx + 200,y + self.book_0_0_disy)
        pyautogui.scroll(8000)
        time.sleep(0.2)

        #將待吃萌獸加入到序列中，加入的索引值為要吃的當下在書中的位置
        while(1):
            row = index // 4
            col = index % 4

            
            #要判斷是否有新的一列，要捲下去
            if(row > 3 and col == 1):
                pyautogui.moveTo(x + self.book_0_0_disx + 200,y + self.book_0_0_disy)
                pyautogui.scroll(-400)
                time.sleep(0.2)
                new_row_reg = pyautogui.locateOnScreen(r'./img/scroll_reg.png',region=(x+690,y+270,60,40))
                #剛好16張會觸發
                if(new_row_reg == None):
                    break
                else:
                    scroll_count += 1

            px = x + self.book_0_0_disx + col * self.book_blank_col
            py = y + self.book_0_0_disy + (row - scroll_count) * self.book_blank_row

            #判斷當前位置是否有卡片，沒有代表當前等級能吃的都加入到序列了
            empty_x = px + self.book_empty_screen_disx #(px-15)
            empty_y = py + self.book_empty_screen_disy #(py-25)
            #pyautogui.screenshot(r'./img/shot.png',region=(empty_x,empty_y,40,40))
            null_reg = pyautogui.locateOnScreen(r'./img/pet_empty_reg.png',region=(empty_x,empty_y,40,40))
            if(null_reg != None):
                break

            #先判斷當前位置是否是要被強化的卡片
            if((row+1,col+1) in except_cards):
                print(str(row+1) + '列'+ str(col+1) + '欄為排除對象，換下一個位置')
                time.sleep(0.05)
            else:
                #判斷當前位置的卡片是否為上鎖，上鎖就找下一個位置，沒上鎖代表可以吃，將順序加入到序列
                lock_x = px + self.book_lock_screen_disx #(px+0)
                lock_y = py + self.book_lock_screen_disy #(py-50)
                lock_reg = pyautogui.locateOnScreen(r'./img/lock_reg.png',region=(lock_x,lock_y,30,30), grayscale=True, confidence=0.75)
                if(lock_reg != None):
                    print(str(row+1) + '列'+ str(col+1) + '欄為上鎖卡片，換下一個位置')
                    time.sleep(0.05)
                else:
                    eat_index_queue.put(c_index)
                    count += 1

                if(count == 4):
                    count = 0
                    c_index -= 4
            index += 1
            c_index += 1
    


        pyautogui.moveTo(x + self.book_0_0_disx + 200,y + self.book_0_0_disy)
        pyautogui.scroll(8000)
        time.sleep(0.2)


        scroll_count = 0
        

        check_str = ''

        while(not eat_index_queue.empty()):

            #吃之前先判斷要升級的萌獸是否滿等(沒有顯示LV)
            setreg = pyautogui.locateOnScreen(r'./img/pet_is_set_reg.png',region=(x+440,y+435,20,20), grayscale=True, confidence=0.75)
            if(setreg == None):
                print("萌獸已強化至最高等。")
                self.upgrade_flag = False
                return eat_count


            #取序列中前4個
            count = 0
            while(count < 4 and not eat_index_queue.empty()):


                c_index = eat_index_queue.get()
                check_str += ' ' + str(c_index)
                row = c_index // 4
                col = c_index % 4
                while((row - scroll_count) >= 4):
                    pyautogui.moveTo(self.book_0_0_disx+200,self.book_0_0_disy)
                    pyautogui.scroll(-400)
                    scroll_count += 1


                px = x + self.book_0_0_disx + col * self.book_blank_col
                py = y + self.book_0_0_disy + (row - scroll_count) * self.book_blank_row

                pyautogui.moveTo(px,py)
                pyautogui.click(clicks = 2)
                count += 1
                time.sleep(0.2)

            
            #點獻祭台下面的確認
            pyautogui.moveTo(confirm_x,confirm_y)
            pyautogui.click(clicks = 1)
            time.sleep(0.05)

            #點跳出來的確認視窗
            pyautogui.moveTo(x+725,y+440)
            pyautogui.click(clicks = 1)
            time.sleep(0.05)
            


            fullreg = pyautogui.locateOnScreen(r'./img/msg_ok_reg.png',region=(x+750,y+430,50,30), grayscale=True, confidence=0.75)
            if(fullreg != None):
                print("沒楓幣了。")
                pyautogui.moveTo(x+760,y+435)
                pyautogui.click(clicks = 1)
                self.upgrade_flag = False
                break
            else:
                eat_count += count

            time.sleep(0.5)

        print(check_str)

        time.sleep(3)
        return eat_count


    



       



    def run(self):
        print_msg = False
        while(not self.exit_flag):
            if(not print_msg):
                print("----- F1: 吃背包中的萌獸  F2: 升級萌獸  F3: 離開程式 -----")
                print_msg = True

            if(self.eat_flag):
                
                self.__add__()
                print_msg = False
            elif(self.upgrade_flag):
                self.__upgrade__()
                print_msg = False
            time.sleep(0.5)


        


script_run = Script(hwn)  
script_run.run()


