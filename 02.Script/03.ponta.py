# ------ ライブラリのインポート ------ 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import sys
import json
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from PIL import Image
import io


# ------ 画像を縦方向または横方向に結合する関数 ------ 
def get_concat_image(img1, img2, direction='v'):
    if direction == 'v':
        new_img = Image.new('RGB', (max(img1.width, img2.width), (img1.height + img2.height)))
        new_img.paste(img1, (0, 0))
        new_img.paste(img2, ((img1.width - img2.width) //  2, img1.height))
    elif direction == 'h':
        new_img = Image.new('RGB', ((img1.width + img2.width), max(img1.height, img2.height)))
        new_img.paste(img1, (0, 0))
        new_img.paste(img2, (img1.width, (img1.height - img2.height) // 2))
    else:
        return None
    return new_img



# ------ スクリーンショットを取得し、スクロールバーを削除する関数 ------ 
def get_screenshot_image(driver):
    c_width, c_height, i_width, i_height = driver.execute_script(
        'return [document.documentElement.clientWidth, document.documentElement.clientHeight, window.innerWidth, window.innerHeight]')
    image = Image.open(io.BytesIO(driver.get_screenshot_as_png()))
    w = image.size[0] * c_width  // i_width  
    h = image.size[1] * c_height // i_height
    return image.crop((0,0,w, h))



# ------ 画面をスクロールしながら全画面を取得する関数 ------ 
def get_full_screenshot_image(driver, wait_time=3):
    s_height, c_height = driver.execute_script(
        'return [document.body.scrollHeight, document.documentElement.clientHeight]')
    driver.execute_script('window.scrollTo(0, arguments[0]);', s_height)
    time.sleep(wait_time)
    ## 1 ページ目を入れる
    driver.execute_script('window.scrollTo(0, arguments[0]);', 0)
    full_page = get_screenshot_image(driver)
    ## 少しずつスクロールしながらスクリーンショットを取得する
    for y_coord in range(c_height, s_height, c_height):
        driver.execute_script('window.scrollTo(0, arguments[0]);', y_coord)
        single_page = get_screenshot_image(driver)
        ## 最後のスクロールは適度にクロップする
        if s_height - y_coord < c_height:
            h = single_page.size[1] * (c_height - (s_height - y_coord)) // c_height
            single_page = single_page.crop((0, h, single_page.size[0], single_page.size[1]))
        ## スクリーンショットを結合する
        full_page = get_concat_image(full_page, single_page)
    ## ページトップヘ戻る
    driver.execute_script('window.scrollTo(0, arguments[0]);', 0)
    return full_page



# ------ JSONファイル読み込み（出力フォルダパス） ------
temp_json = open('01.Config\\00.temp.json', 'r',encoding="utf-8")
temp_data = json.load(temp_json)
FilePath = temp_data["FileDir"]
downloadsFilePath = FilePath



# ------ ChromeDriver のオプション指定 ------
chrome_options = webdriver.ChromeOptions()
prefs = {
    "plugins.always_open_pdf_externally": True,                                              # PDFをブラウザのビューワーで開かせない
    "credentials_enable_service": False,                                                     # 認証情報(ログイン情報やパスワード)の自動保存を無効
    "profile.password_manager_enabled" : False,                                              # ブラウザが自動的に保存するパスワードマネージャーを無効
    "profile.default_content_settings.popups": 1,                                            # ポップアップウィンドウを許可(ブロックされずに表示されるようにする)
    "download.default_directory": os.path.abspath(downloadsFilePath) + r"\\",                # ダウンロードディレクトリのパスを設定
    "directory_upgrade": True                                                                # 新しい設定を有効にする
}
chrome_options.add_argument('--ignore-certificate-errors')                                          # 証明書エラーを無視する
chrome_options.add_argument('--ignore-ssl-errors')                                                  # SSLエラーを無視する
chrome_options.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])  # 自動化やログの有効化を除外する
chrome_options.add_experimental_option("prefs", prefs)                                                 # 定義したプリファレンスを設定オプションに追加



# ------ ChromeDriver の起動 ------
driver_path = ChromeDriverManager().install()
service =Service(driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get('https://www.amazon.co.jp/mypoints/')
time.sleep(3)
wait = WebDriverWait(driver, 10)
element = wait.until(EC.visibility_of_all_elements_located)



###################################################
###############   メイン処理開始   #################
###################################################
# ------ 手動ログオン ------
sub_win = tk.Tk()
sub_win.withdraw()
sub_win.attributes("-topmost",True)
messagebox.showwarning("手動でログイン処理をお願いします!","ログイン画面上の「ユーザID」と「パスワード」を入力し、手動でログインしてください。\nログインが終わり画面が遷移したら、このポップアップの「OK」をクリックしてください。", parent=sub_win)
sub_win.attributes("-topmost",False)
driver.maximize_window()



# ------ 期間の絞り込み ------
dropdown = driver.find_element(By.CLASS_NAME, "select")
select = Select(dropdown)
month =  temp_data["month"] 
if int(month) > 9:
    acqDate = temp_data["year"]+"年"+temp_data["month"]+"月"
else:
    acqDate = temp_data["year"]+"年"+"0"+temp_data["month"]+"月"
select.select_by_visible_text(acqDate)
time.sleep(3)
# 取得するデータがない場合は、処理をスキップ
try:
    history = driver.find_element(By.CLASS_NAME, "history-list")
except:
    file = open(FilePath+"※取得できるデータがありません", 'w')
    file.close()
    driver.close()
    sys.exit(0)



# ------ スクリーンショット取得 ------
screenshot = get_full_screenshot_image(driver)
screenshot.save(FilePath+acqDate+'pontaポイント.png')



# ------ テーブルから情報を抽出しリストへ格納 ------ 
historys = driver.find_elements(By.TAG_NAME, 'p')
dates = []
contents = []
for history in historys[4::1]:
    content = history.text
    flg = acqDate in content
    if len(dates) == 0:
        dates.append(content)    
    elif flg == True:
        dates.append(content)
        del dates[len(dates)-2]
    else:
        contents.append(content)
        dates.append(dates[len(dates)-1])
del dates[len(dates)-1]
start = contents.index("")
end = len(contents)
del contents[start:end]
del dates[start:end]
# 履歴からポイント抽出
history = driver.find_element(By.CLASS_NAME, "history-list")
historys = history.find_elements(By.TAG_NAME, 'dd')
points = []
for history in historys:
    point = history.text
    point = point.replace('P','')                 # 文字列「P」を削除
    point = point.replace(",",'')                 # 文字列「,」を削除
    point = int(point)                            # 整数型に変換
    points.append(int(point))



# ------ データフレーム ------
# リストをデータフレームに格納
df = pd.DataFrame()
df["日付"] = dates
df["内容"] = contents
df["ポイント"] = points[4::1]
kinds = ['獲得ポイント合計(ponta)','獲得ポイント合計(aupay限定)','利用ポイント合計(ponta)','利用ポイント合計(aupay限定)']
del points[4:len(points)]   # 先頭4行のみ取り出し（合計ポイント）
df_Total = pd.DataFrame()
df_Total["種類"] = kinds
df_Total["ポイント"] = points
df_Change = df[df['内容'].str.contains('ポイント交換に伴う利用', na=False)]     # 獲得予定ポイントを排除
df_Change = df_Change.drop(['日付','内容'], axis=1)
df_Change = df_Change.rename(columns={'ポイント':'交換'})
df_Change.set_axis([2], axis="index", inplace=True)
df_Change.loc[0] = [0]
df_Change.loc[1] = [0]
df_Change.loc[3] = [0]
df_Change = df_Change.sort_index() 
df_Total = pd.concat([df_Total,df_Change], axis = 1) # 合計ポイント
df_a = df.loc[df['ポイント'] >= 0] # 獲得ポイント
df_u = df.loc[df['ポイント'] < 0] # 利用ポイント
# Excel形式でファイル保存
writer = pd.ExcelWriter(FilePath+acqDate+"Pontaポイント.xlsx", engine='openpyxl')
df_Total.to_excel(writer, sheet_name = "合計ポイント")
df_a.to_excel(writer, sheet_name = "獲得ポイント履歴")
df_u.to_excel(writer, sheet_name = "利用ポイント履歴")
writer.save() 
driver.quit()
