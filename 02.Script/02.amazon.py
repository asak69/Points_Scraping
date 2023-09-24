# ------ ライブラリのインポート ------
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os
import sys
import json
import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import tkinter.simpledialog as simpledialog
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



# ------ テーブルから情報を抽出しリストとして返す関数 ------ 
def listCreate(driver, lists, NumCase, elemClass):
    num=1
    while NumCase >= 0:
        flg = "a-text-right" in elemClass
        if flg:
            screenshot = get_full_screenshot_image(driver)
            screenshot.save(FilePath+month+'_Amazonポイント{}.png'.format(num))
        # テーブルからリストへ格納
        elem = driver.find_element(By.CLASS_NAME, 'a-normal')
        elem_tables = elem.find_elements(By.CLASS_NAME, elemClass)
        for elem_table in elem_tables:
            list = elem_table.text.replace("\u3000","_")
            list = list.replace("+","")
            lists.append(list)
        try:
            driver.find_element(By.PARTIAL_LINK_TEXT, "次のページ").click() # 次のページに移動
            num += 1
            time.sleep(2)  # ページ遷移待機
        except:
            time.sleep(2)  # ページ遷移待機     
        NumCase -= 10
    return lists



def page_remove(driver, NumCase):
    while NumCase >= 0:
        try:
            driver.find_element(By.PARTIAL_LINK_TEXT, "前のページ").click() # 前のページに移動
            time.sleep(2)  # ページ遷移待機
        except:
            time.sleep(2)  # ページ遷移待機     
        NumCase -= 10



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


# ------ 期間の絞り込み（カレンダー） ------
elem_cal = driver.find_element(By.CLASS_NAME, "a-input-text-addon")
elem_cal.click()
# カレンダー表の取得
elem_cal = driver.find_element(By.CLASS_NAME, "a-cal-month-container")
year = temp_data["year"]
dt_now = datetime.now()
now_year = dt_now.year
if int(year) > int(now_year):
    # カレンダー表の年度　一つ戻る
    elem_cal.find_element(By.CLASS_NAME, "a-icon").click()
month = temp_data["month"]+"月"
driver.find_element(By.PARTIAL_LINK_TEXT, month).click()
time.sleep(3)



# 取得するデータがない場合は、処理をスキップ
try:
    driver.find_element(By.CLASS_NAME, 'a-normal')
except:
    file = open(FilePath+"※取得できるデータがありません", 'w')
    file.close()
    driver.close()
    sys.exit(0)
#  件数の取得
elem_his = driver.find_element(By.ID, "points-transaction-tab-container")
elem_NumCases = elem_his.find_elements(By.CLASS_NAME, "a-section")
NumCase = elem_NumCases[2].text
start = NumCase.find("(")
end = NumCase.find("件")
NumCase_rep = int(NumCase[start+1:end])
# 空のリスト作成
points = []
points2 = []
ptypes = []
L_rirekis = []
L_Rirekis = listCreate(driver, L_rirekis, NumCase_rep, "a-text-left")
page_remove(driver, NumCase_rep)
Dates = L_Rirekis[0::3]
Items = L_Rirekis[1::3]
Links = L_Rirekis[2::3]
pTypes = listCreate(driver, ptypes, NumCase_rep, "a-text-center")
page_remove(driver, NumCase_rep)
Points = listCreate(driver, points, NumCase_rep, "a-text-right")
page_remove(driver, NumCase_rep)
# ポイント型変換（int）
for point in Points:
    point = point.replace(",", "")
    flg = '獲得予定' in point or '期間限定' in point or 'ポイント' in point
    if not flg:
        point = int(point)
        points2.append(point)
    else:
        point = point.replace("\n", " ")
        points2.append(point)



# ------ データフレーム ------
# リストをデータフレームに格納
df = pd.DataFrame()
df["日付"] = Dates[1::1]
df["項目"] = Items[1::1]
df["リンク"] = Links[1::1]
df["種類"] = pTypes[1::1]
df["ポイント"] = points2[1::1]
# Excel形式でファイル保存
writer = pd.ExcelWriter(FilePath+month+"_Amazonポイント.xlsx", engine='openpyxl')
try:
    points_df = df[~df['ポイント'].str.contains('獲得予定', na=False)]     # 獲得予定の文字を含むdfをを除外
    points_df = points_df[~points_df['ポイント'].str.contains('期間限定', na=False)]     # 期間限定の文字を含むdfを除外
    acq_points_df = points_df.loc[points_df['ポイント'] >= 0] # 獲得ポイント
    used_points_df = points_df.loc[points_df['ポイント'] < 0] # 利用ポイント
    acq_schedule_df = df[df['ポイント'].str.contains('獲得予定', na=False)]         # 獲得予定ポイント
    limited_time_df = df[df['ポイント'].str.contains('期間限定', na=False)]         # 期間限定予定ポイント
    acq_schedule_df.to_excel(writer, sheet_name = "獲得予定ポイント")
    limited_time_df.to_excel(writer, sheet_name = "期間限定ポイント") 
except:
    # エラーハンドリング
    acq_points_df = df.loc[df['ポイント'] >= 0] # 獲得ポイント
    used_points_df = df.loc[df['ポイント'] < 0] # 利用ポイント
# Excel形式でファイル保存
acq_points_df.to_excel(writer, sheet_name = "獲得ポイント")
used_points_df.to_excel(writer, sheet_name = "利用ポイント")
writer.save() 
driver.quit()
