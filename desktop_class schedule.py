from pystray import Icon, MenuItem, Menu
from threading import Thread
import schedule
import PySimpleGUI as sg
import csv
import ctypes
from PIL import Image,ImageDraw,ImageFont
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from subprocess import CREATE_NO_WINDOW
from time import sleep
import os
import datetime
import sys

# サンプルのURL
TABLE_URL = 'http://furi0web.html.xdomain.jp/' 

# データの読み出し関数
def load_data():
    csv_data = {}
    with open("data.csv") as f:
        for i in csv.reader(f):
            if len(i)==2:
                csv_data[i[0]] = i[1]
    return csv_data

# データの書き込み関数
def write_data(csv_data):
    with open('data.csv', 'w') as f:  
        writer = csv.writer(f, lineterminator='\n')
        for k, v in csv_data.items():
            writer.writerow([k, v])

# スクレイピングで1～4時限の科目名を読み取る
def table_read(chromedriver_path, target_url):
    week_num = datetime.date.today().weekday()
    # 曜日の確認
    if week_num <= 4:
        try:
            options = Options()
            options.add_argument('--headless')
            service = Service(chromedriver_path)
            service.creationflags = CREATE_NO_WINDOW
            driver = webdriver.Chrome(service=service, chrome_options=options)
            error_flg = False
            driver.get(target_url)
            sleep(3)
        except Exception:
            # スクレイピングの設定失敗
            return ['Error','Miss scraping set','','']
        
        try:
            table = driver.find_element(by=By.XPATH, value='/html/body/table')
            print(table)
            table_val = []
            for i in table.find_elements(by=By.TAG_NAME, value="tr"):
                tr_val = [] 
                for j in i.find_elements(by=By.TAG_NAME, value="td"):
                    tr_val.append(j.text)
                table_val.append(tr_val)
            
        except Exception:
            # 表の取得失敗
            return ['Error','Not found table','','']
        
        # 文字数に制限をかける
        lesson = []
        for i in range(1,5):
            lesson_text = table_val[i][week_num+1]
            # 10文字で切る
            if len(lesson_text) > 10:
                lesson_text = lesson_text[:10]
            lesson.append(lesson_text)
        return lesson
    elif week_num in (5,6):
        # 土日は無いため空欄とする
        return ['','','','']
    else:
        # 曜日が取得できないとき
        return ['Error','Day get failed','','']



class taskTray:
    # 初期化
    def __init__(self, image):
        self.status = False
        self.flag = False
        image = Image.open(image)
        menu = Menu(
                    MenuItem('Update', self.update),
                    MenuItem('GetTable', self.update_table),  
                    MenuItem('Setting', self.setting),
                    MenuItem('Exit', self.stopProgram),
                )

        self.icon = Icon(name='Schedule', title='Schedule', icon=image, menu=menu)
        self.dir_path = os.path.dirname(os.path.abspath(__file__))+'\\'
        self.wallpaper = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.SystemParametersInfoW(115,len(self.wallpaper),ctypes.byref(self.wallpaper),0)
        self.img_path = self.wallpaper.value.replace('\\','\\\\')
        self.get_table()
        self.update()
    
    # 位置やメモの設定
    def setting(self):
        self.flag = True
        self.picker()
        self.update()

    # 時刻表を更新して表示
    def update_table(self):
        self.flag = True
        self.get_table()
        self.update()

    # 表の取得を取得して保存
    def get_table(self):
        csv_data = load_data()
        cd_path = self.dir_path+'chromedriver' if csv_data['chromedriver'] == '' else csv_data['chromedriver']
        lesson_list = table_read(cd_path,TABLE_URL)
        csv_data['lesson1'] = lesson_list[0]
        csv_data['lesson2'] = lesson_list[1]
        csv_data['lesson3'] = lesson_list[2]
        csv_data['lesson4'] = lesson_list[3]
        csv_data['day'] = datetime.datetime.now().strftime('%Y:%m:%d')
        write_data(csv_data)

    # 時刻表の表示
    def update(self):
        # 表示位置の取得
        csv_data = load_data()
        if csv_data['day'] != datetime.datetime.now().strftime('%Y:%m:%d'):
            self.get_table()
            csv_data = load_data()
        try:
            pos_x = int(csv_data['x'])
        except Exception:
            pos_x = 0
        try:
            pos_y = int(csv_data['y'])
        except Exception:
            pos_y = 0
        
        ctypes.windll.user32.SystemParametersInfoW(115,len(self.wallpaper),ctypes.byref(self.wallpaper),0)
        img_path2 = self.wallpaper.value.replace('\\','\\\\')

        if img_path2 == self.dir_path+'img2.png':
            img1 = Image.open(self.img_path).convert('RGBA')
        else:
            img1 = Image.open(img_path2).convert('RGBA')

        img1 = Image.open(self.img_path).convert('RGBA')
        img2 = Image.open("back1.png").convert('RGBA')
        img_a = img1.copy()
        img_a.putalpha(alpha=0)
        img_a.paste(img2, (pos_x, pos_y))
        draw = ImageDraw.Draw(img1)
        font = ImageFont.truetype("HGRME.TTC", 30)
        draw.text((pos_x+210,pos_y+159), csv_data['lesson1'] ,(255,255,255),font=font)
        draw.text((pos_x+210,pos_y+204), csv_data['lesson2'] ,(255,255,255),font=font)
        draw.text((pos_x+210,pos_y+249), csv_data['lesson3'] ,(255,255,255),font=font)
        draw.text((pos_x+210,pos_y+294), csv_data['lesson4'] ,(255,255,255),font=font)
        memo1_text = csv_data['memo'] if len(csv_data['memo'])<=14 else csv_data['memo'][:14] 
        memo2_text = csv_data['memo2'] if len(csv_data['memo2'])<=14 else csv_data['memo2'][:14]
        draw.text((pos_x+80,pos_y+450), memo1_text+'\n'+memo2_text ,(255,255,255),font=font)
        img3 = Image.alpha_composite(img1, img_a)
        img3.save(self.dir_path+'img2.png')
        csv_data['dir_path'] = self.dir_path+'img2.png'
        ctypes.windll.user32.SystemParametersInfoW(20, 0, self.dir_path+'img2.png', 0)

    # 終了処理
    def stopProgram(self, icon):
        self.status = False
        self.icon.stop()
        ctypes.windll.user32.SystemParametersInfoW(20, 0, None, 0)
    
    # 設定画面の描画処理
    def picker(self):
        csv_data = load_data()

        layout2 = [[sg.Text("Possition X",size=(10,1)),sg.Input(key='<Pos_X>',default_text=csv_data['x'])],
                    [sg.Text("Possition Y",size=(10,1)),sg.Input(key='<Pos_Y>',default_text=csv_data['y'])],
                    [sg.Text("Memo",size=(10,1)),sg.Input(key='<Memo>',default_text=csv_data['memo'])],
                    [sg.Text("Memo2",size=(10,1)),sg.Input(key='<Memo2>',default_text=csv_data['memo2'])],
                    [sg.Button('登録')]]
        setting_window = sg.Window('設定', layout2,finalize=True,modal=True)

        event, values = setting_window.read()
        
        if event == '登録':
            csv_data['x'] = values['<Pos_X>']
            csv_data['y'] = values['<Pos_Y>']
            csv_data['memo'] = values['<Memo>']
            csv_data['memo2'] = values['<Memo2>']
            write_data(csv_data)
        setting_window.close()

    # スケジュール処理
    def runSchedule(self):
        schedule.every().day.at("00:01").do(self.update)
        while self.status:
            schedule.run_pending()
            sleep(5)
    
    # 実行処理
    def runProgram(self):
        self.status = True
        task_thread = Thread(target=self.runSchedule)
        task_thread.start()
        self.icon.run()


if __name__ == '__main__':
    system_tray = taskTray(image="icon.png")
    system_tray.runProgram()
    sys.exit()

