# ライブラリのインポート
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import sys
import json
import time
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



# ------ 整数変換する関数 ------ 
def listCreate(lists):
    lists2 = []
    for list in lists:
        if not list in "-":
            list = list.replace(",",'')                 # 文字列「,」を削除
            list = int(list)                            # 整数型に変換    
        lists2.append(list)    
    return lists2



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
driver.get('https://www.biccamera.com/bc/member/CSfLogin.jsp?autoLogin')
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
driver.find_element(By.PARTIAL_LINK_TEXT, "ポイント履歴").click()



# ------ 期間の絞り込み（カレンダー） ------
dropdown = driver.find_element(By.NAME, "POINT_SEARCH_DM")
select = Select(dropdown)
select.select_by_visible_text('1年以内')
year = temp_data["year"]+"年"
month = temp_data["month"]+"月"



# ------ 取得するデータがない場合は、処理をスキップ ------
try:
    elem_rireki = driver.find_element(By.TAG_NAME, "tbody")
    rirekis = []
    elem_rirekis = elem_rireki.find_elements(By.TAG_NAME, "td")
    for elem_rireki in elem_rirekis:
        rireki = elem_rireki.text
        rirekis.append(rireki) 
    Dates = rirekis[0::6]
    Contents = rirekis[1::6]
    ItemIDs = rirekis[2::6]
    acqPoints = rirekis[3::6]
    usePoints = rirekis[4::6]
    Details = rirekis[5::6]
    acqPoints = listCreate(acqPoints)
    usePoints = listCreate(usePoints)
except:
    file = open(FilePath+"※取得できるデータがありません", 'w')
    file.close()
    screenshot = get_full_screenshot_image(driver)
    screenshot.save(FilePath+month+'_ビックポイント.png')
    driver.close()
    sys.exit(0)



# ------ データフレーム ------
# リストをデータフレームに格納
df = pd.DataFrame()
df['日付'] = Dates
df['内容'] = Contents
df['注文番号'] = ItemIDs
df['獲得ポイント'] = acqPoints
df['利用ポイント'] = usePoints
df['詳細'] = Details
df2 = df[df['日付'].str.contains(year+month)]
# スクショ
screenshot = get_full_screenshot_image(driver)
screenshot.save(FilePath+month+'_ビックポイント.png')
# Excel形式でファイル保存
writer = pd.ExcelWriter(FilePath+month+"_ビックポイント.xlsx", engine='openpyxl')
df2.to_excel(writer, sheet_name = month + "_獲得・利用ポイント")
df.to_excel(writer, sheet_name = "獲得・利用ポイント")
writer.save() 
driver.quit()