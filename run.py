import os
import json
import sys
import threading

import eel
import common.desktop as desktop
from rakutenMobileCheck import Search


RAKUTENMOBILEURL = "https://portal.mobile.rakuten.co.jp/equipment-details?id=50150604"
#RAKUTENMOBILEURL = "https://portal.mobile.rakuten.co.jp/equipment-details?id=50150294&selectAvailable=true"
SETTING_JSON = "./setting.json"

app_name="web"
end_point="index.html"
size=(700,1000)

search_start = Search(RAKUTENMOBILEURL, SETTING_JSON)

@ eel.expose
def loop_start():
    try:
        t = threading.Thread(target=search_start.search_rakuten())
        t.start()
    except Exception as e:
        eel.view_status("※エラーが発生しました。setting.jsonやブラウザをご確認ください")

@ eel.expose
def loop_stop():
    t = threading.Thread(target=search_start.stop_flag_assign)
    t.start()

if __name__ == "__main__":
    desktop.start(app_name,end_point,size)
