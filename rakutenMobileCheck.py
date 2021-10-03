import json
from time import sleep
import datetime
import win32com.client
import threading

import eel 
from common.driver import Driver
from common.logger import set_logger


DATE_FORMAT = '%Y-%m-%d-%H-%M-%S'

class Search:
    """
    楽天モバイルで在庫を確認する処理
    """


    def __init__(self, rakuten_url, setting_json):
        self.rakuten_url = rakuten_url
        self.setting_json = setting_json

        self.stock = []
        self.logger = set_logger("__name__")

    # main処理
    def search_rakuten(self):

        #setting.jsonから設定値を取得
        json_parameter = self.get_josn()
        
        #終了時間または、STOPを押すまで無限ループ
        while self.end_judge():

            #処理ステータスを表示
            eel.view_status("在庫チェック中")

            # driverを起動
            driver = Driver.set_driver(False)

            # Webサイトを開く
            driver.get(self.rakuten_url)

            self.logger.info("楽天モバイルページへ遷移")
            sleep(5)

            #在庫チェック時の情報を出力
            dt = datetime.datetime
            self.check_time =dt.now().strftime('%Y-%m-%d %H:%M:%S')

            #機種名を取得
            self.model_name = driver.find_element_by_css_selector(
            "div.rktn-equipment__title-container-position > h1.rktn-equipment__title-container-name").text


            eel.view_log_js(f"{self.check_time}：「{self.model_name}」の在庫チェック開始")


            #各色のブロックごとにリストに格納
            color_table = driver.find_elements_by_css_selector(
            "div.rktn-product-colors__content > div.rktn-product-colors__item")
                
            # #各色のブロックごろに処理を繰り返す
            for i, color in enumerate(color_table):

                # #2回目以降エラーになるので再度ドライバを指定
                # color = driver.find_elements_by_css_selector("div.rktn-product-colors__content")

                #色を格納       
                color_name = color.find_element_by_css_selector(
                    "div.rktn-product-colors__item-content-text").text
                
                #「在庫なし」というキーワードがない場合、HTMLがないとエラーになる
                try:    
                    #在庫の確認
                    stock_check = color.find_element_by_css_selector(
                        "div.rktn-product-colors__item-status").text
                except:
                    stock_check = "在庫あり"
                    #在庫ありリストの追加
                    self.stock.append(color_name)

                #ログ出力
                self.logger.info(f"「{color_name}」：{stock_check}")
                #デスクトップテキストエリアへの出力
                eel.view_log_js(f"「{color_name}」：{stock_check}")

                sleep(3)

            #メモリのブロックごとにリストに格納
            memory_table = driver.find_elements_by_css_selector(
                "div.rktn-equipment-memory__memories > div.rktn-equipment-memory__memories-content")

            #各メモリのブロックごろに処理を繰り返す
            for i, memory in enumerate(memory_table):

                # #2回目以降エラーになるので再度ドライバを指定
                # color = driver.find_elements_by_css_selector("div.rktn-product-colors__content")

                #メモリを格納       
                memory_name = memory.find_element_by_css_selector(
                    "div.rktn-equipment-memory__memories-content-item-info > div.rktn-equipment-memory__memories-content-item-info-value").text
                
                #在庫の確認
                #「在庫なし」というキーワードがない場合、HTMLがないとエラーになる
                try:    
                    #在庫の確認
                    stock_check = memory.find_element_by_css_selector(
                        "div.rktn-equipment-memory__memories-content-item-status").text
                except:
                    stock_check = "在庫あり"
                    #在庫ありリストの追加
                    self.stock.append(color_name)

                #ログ出力
                self.logger.info(f"「{memory_name}GB」：{stock_check}")
                #デスクトップテキストエリアへの出力
                eel.view_log_js(f"「{memory_name}GB」：{stock_check}")

                sleep(1)

            #ブラウザを閉じる
            driver.quit()

            #在庫があったらメールを送る
            if len(self.stock) > 0:

                #メール送信クラスを呼び出し
                self.logger.info("メール通知処理")
                sendemail = SendEmail(json_parameter,self.stock,self.check_time,self.model_name)
                sendemail.send_email()
            
            else:
                #ログ出力
                self.logger.info("全種類在庫なし")

            #終了時刻の場合、インターバルを無視
            if self.end_judge():

                #ログ出力
                self.logger.info("インターバル中")
                 #処理ステータスを表示
                eel.view_status("インターバル中")
                
                #setting.jsonで指定した時間まで待機
                sleep(int(json_parameter["environment_preference"]["interval"]))

                #在庫ありリストリセット
                self.stock = []

        #ログ出力
        self.logger.info("終了時刻となったので、処理を終了します")
        eel.view_status("処理停止中")

    #setting.jsonから設定値を取得
    def get_josn(self):
        #setting.jsonを開く
        json_open = open(self.setting_json, 'r')
        #setting.jsonを読み込む
        json_load = json.load(json_open)

        return json_load

    #終了時間の判定を行う
    def end_judge(self):

        #今の時刻
        dt = datetime.datetime
        dt_now =dt.now()

        #setting.jsonから設定値を取得
        json_parameter = self.get_josn()
        
        now_time = datetime.datetime(dt_now.year,dt_now.month,dt_now.day,dt_now.hour,dt_now.minute)
        
        end_time = json_parameter["environment_preference"]["endTime"]
        end_time = end_time.split(",")

        end_time = datetime.datetime(
            int(end_time[0]),int(end_time[1]),int(end_time[2]),int(end_time[3]),int(end_time[4]))

        #時間比較
        return now_time < end_time

    #STOP判定用ファイルを作成します
    def stop_flag_assign(self):
        self.logger.info("停止判定ファイル生成")
        with open("./stop_flg.dat", 'w', encoding='UTF-8') as f:
            f.write("ダミーファイルです")

class SendEmail():

    def __init__(self, json_parameter, stock_list, check_time, model_name):
        self.json_parameter = json_parameter
        self.stock_list = stock_list
        self.check_time = check_time
        self.model_name = model_name

    def send_email(self):

        #Outlook設定
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.to = self.json_parameter["environment_preference"]["emailAddress"]
        mail.subject = f"★{self.model_name}の在庫がありました：チェック時間[{self.check_time}]"
        mail.bodyFormat = 1

        #在庫ありのリストから本文を作成
        for stock in self.stock_list:
            message = f"「{stock}」:在庫あり" + "\n"
            mail.body += message + "\n"
        
        mail.Send()
        
