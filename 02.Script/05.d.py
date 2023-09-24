# ------ ライブラリのインポート ------
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import sys
import json
import time
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



# ------ 文字列が整数であるか判定する関数 ------ 
def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


     
# ------ テーブルから情報を抽出しリストとして返す関数 ------ 
def extract_data(driver, max_page, screenshotFilePath):
    num =1
    reflection_dates, use_dates, contents, details, types, points, expiration_dates = [], [], [], [], [], [], []
    while num <= max_page:
        try:
            elem = driver.find_element(By.CLASS_NAME, "data01")
            table_rows = elem.find_elements(By.TAG_NAME, "tr")
            for row in table_rows:
                row_data = row.text.replace("★\n", "").replace("\u3000", "").replace("\n", " ")
                row_parts = row_data.split(" ", maxsplit=6)
                reflection_dates.append(row_parts[0])
                use_dates.append(row_parts[1])
                contents.append(row_parts[2])
                details.append(row_parts[3])
                types.append(row_parts[4])
                points.append(row_parts[5].replace(",", ""))
                expiration_dates.append(row_parts[6])
            screenshot = get_full_screenshot_image(driver)
            screenshot.save(screenshotFilePath + '{}.png'.format(num))
            driver.find_element(By.NAME, "root_GKKCGW001RirekiFTZnextPage").click()
        except:
            reflection_dates, use_dates, contents, details, types, points, expiration_dates = ["-", "出力情報なし"] * 7
        num += 1
    return reflection_dates, use_dates, contents, details, types, points, expiration_dates



# ------ ページ数を取得する関数 ------ 
def get_max_page(driver, default_max_page):
    # ページネーション要素を取得
    elem_links = driver.find_elements(By.CLASS_NAME, "t-char")
    # 最大ページ数を取得
    links = elem_links[0].text
    start = links.find("/")
    end = len(links)
    max_page = int(links[start+1:end])
    return max_page



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
driver.get('https://point.smt.docomo.ne.jp/guide/value/gkkcp001.srv?dcmancr=0ce21c5b356bcf3cc3c136eed388f588.1671036534669.8106.1671036545114_1278837160.1671036529_415&_ga=2.136407088.1883638374.1671036529-1278837160.1671036529')
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
dropdown = driver.find_element(By.NAME, "root_GKKCGW001_RIYONENGETPULDOWNN")
select = Select(dropdown)
if len(temp_data["month"]) == 1:
    temp_data["month"] = "0" + temp_data["month"]
tgt_date = temp_data["year"]+"年"+temp_data["month"]+"月分"
select.select_by_visible_text(tgt_date)



# ------ 取得するデータがない場合は、処理をスキップ ------
try:
    elem = driver.find_element(By.CLASS_NAME, "data01")
except:
    file = open(FilePath+"※取得できるデータがありません", 'w')
    file.close()
    driver.close()
    sys.exit(0)



# ------ 獲得ポイント取得 ------
dropdown = driver.find_element(By.NAME, "root_GKKCGW001_HYOUJIOPTIONPULDOWNN1")
select = Select(dropdown)
select.select_by_visible_text("獲得・利用別 　　　　　　　　　　　")
driver.find_element(By.NAME, "root_GKKCGW001SubmitHyoji").click()
time.sleep(5)
dropdown = driver.find_element(By.NAME, "root_GKKCGW001_HYOUJIOPTIONPULDOWNN2")
select = Select(dropdown)
select.select_by_visible_text("獲得 　　　　　　　　　　　　　　　")
driver.find_element(By.NAME, "root_GKKCGW001SubmitHyoji").click()
time.sleep(5)
max_page = get_max_page(driver,0) # 関数（get_max_page）呼び出し
screenshotFilePath = FilePath+tgt_date + 'dポイント獲得'
ACQs = extract_data(driver, max_page, screenshotFilePath) # 関数（extract_data）呼び出し
point= ACQs[5][1]
Check = is_int(point)
ACQPoint = []
if Check is True:
    for point in ACQs[5][1::1]:
        point = int(point)
        ACQPoint.append(int(point))
else:
    for point in ACQs[5][1::1]:
        point = point
        ACQPoint.append(point)



# ------ 利用ポイント取得 ------
dropdown = driver.find_element(By.NAME, "root_GKKCGW001_HYOUJIOPTIONPULDOWNN2")
select = Select(dropdown)
select.select_by_visible_text("利用 　　　　　　　　　　　　　　　")
driver.find_element(By.NAME, "root_GKKCGW001SubmitHyoji").click()
time.sleep(5)
max_page = get_max_page(driver,0) # 関数（get_max_page）呼び出し
screenshotFilePath = FilePath+tgt_date + 'dポイント利用'
USEs = extract_data(driver, max_page, screenshotFilePath) # 関数（extract_data）呼び出し
point= USEs[5][1]
Check = is_int(point)
USEPoint = []
if Check is True:
    for point in USEs[5][1::1]:
        point = int(point)
        USEPoint.append(int(point))
else:
    for point in USEs[5][1::1]:
        point = point
        USEPoint.append(point)



# リストをデータフレームに格納
import pandas as pd
df = pd.DataFrame()
df_u = pd.DataFrame()
df["反映日"] = ACQs[0][1::1]
df["利用日"] = ACQs[1][1::1]
df["内容"] = ACQs[2][1::1]
df["ご利用内容詳細"] = ACQs[3][1::1]
df["種別"] = ACQs[4][1::1]
df["ポイント数"] = ACQPoint
df_u["反映日"] = USEs[0][1::1]
df_u["利用日"] = USEs[1][1::1]
df_u["内容"] = USEs[2][1::1]
df_u["内容"] = USEs[2][1::1]
df_u["ご利用内容詳細"] = USEs[3][1::1]
df_u["種別"] = USEs[4][1::1]
df_u["ポイント数"] = USEPoint
writer = pd.ExcelWriter(FilePath+tgt_date+"_dポイント.xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name = "獲得ポイント履歴")
df_u.to_excel(writer, sheet_name = "利用ポイント履歴")
writer.save()
driver.quit()