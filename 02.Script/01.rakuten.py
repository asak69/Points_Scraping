# ------ ライブラリのインポート ------ 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import json
from PIL import Image
import io
import time
import pandas as pd
from tkinter import messagebox



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
def create_list(driver, lists, elem_class):
    lists = [] # 結果のリストを初期化
    try:
        elem_table = driver.find_element(By.CLASS_NAME, "history-table")
    except:
        lists = ["出力情報なし"] # テーブルが存在しない場合は、"出力情報なし" のリストを代入して終了
    if elem_class == "date":
        # 日付の場合
        elem_dates = elem_table.find_elements(By.CLASS_NAME, elem_class)
        for elem_date in elem_dates:
            date = elem_date.text.replace("\n", "/")  # 改行をスラッシュに置換
            lists.append(date)
    elif elem_class == "point":
        # ポイントの場合
        elem_points = elem_table.find_elements(By.CLASS_NAME, elem_class)
        for elem_point in elem_points:
            point = elem_point.text.replace('\n',' ')  # 改行をスペースに置換
            point = point.replace(',','')  # カンマを削除
            if "(" in point:
                p_start = point.find("(")
                p_end = point.find(")")
                point = point[p_start+1:p_end]  # 括弧内のポイントのみ抽出
                point = int(point)  # 整数に変換
            else:
                point = int(point)  # 整数に変換
            lists.append(point)
    else:
        # その他の場合
        elem_others = elem_table.find_elements(By.CLASS_NAME, elem_class)
        for elem_other in elem_others:
            other = elem_other.text.replace('\n',' ')  # 改行をスペースに置換
            lists.append(other)
    return lists



# ------ ポイント履歴から各要素をリストとして返す関数 ------ 
def scrape_point_history(driver,max_page,screenshot_dir, wait_time=3):
    # 各情報を保持するリスト
    dates = []
    services = []
    details = []
    actions = []
    points = []
    notes = []
    # ページを繰り返し処理
    for page in range(1, max_page):
        url = 'https://point.rakuten.co.jp/history/?page={}'.format(page)
        driver.get(url)
        time.sleep(wait_time)
        screenshot = get_full_screenshot_image(driver)
        screenshot.save(screenshot_dir + '{}.png'.format(page))
        # 各要素を取得してリストに追加
        dates = create_list(driver, dates, "date")
        services = create_list(driver, services, "service")
        details = create_list(driver, details, "detail")
        actions = create_list(driver, actions, "action")
        points = create_list(driver, points, "point")
        notes = create_list(driver, notes, "note")
    # 奇数番目の要素を取得
    dates = dates[0::2]
    return dates, services, details, actions, points, notes



# ------ 画面スクロール関数 ------ 
def scroll(driver):
    driver.execute_script("window.scrollBy(0, 1000);")
    elem_btn = driver.find_element(By.CLASS_NAME, "btn-toggle_trigger")
    elem_btn.click()



# ------ ページ数を取得する関数 ------ 
def get_max_page(driver, default_max_page):
    # ページネーション要素を取得
    elem_pagination = driver.find_element(By.CLASS_NAME, "pagination")
    elem_links = elem_pagination.find_elements(By.TAG_NAME, "a")
    # 最大ページ数を取得
    max_page = len(elem_links)
    # 最大ページ数が1ページの場合、2ページに修正
    if max_page == 1:
        max_page += 1
    return max_page



# ------ ポイント履歴ページの期間指定要素を取得する関数 ------ 
def get_period_elements(driver):
    # 履歴検索条件要素を取得
    elem_history = driver.find_element(By.ID, "history-search-condition")   
    # 獲得履歴と失効履歴のチェックボックス要素を取得
    elem_acquired = elem_history.find_elements(By.CLASS_NAME, "check-items-child")[0].find_elements(By.CLASS_NAME, "checkbox")
    elem_expired = elem_history.find_elements(By.CLASS_NAME, "check-items-child")[1].find_elements(By.CLASS_NAME, "checkbox")
    # 期間指定の要素を取得
    elem_period = elem_history.find_elements(By.CLASS_NAME, "wrap_fromate")
    elem_from = elem_period[0].find_element(By.NAME, "fromDate")
    elem_to = elem_period[1].find_element(By.NAME, "toDate")
    return elem_acquired, elem_expired, elem_from, elem_to



# ------ 検索ボタン実行関数 ------ 
def perform_search(driver):
    elem_history = driver.find_element(By.ID, "history-search-condition")
    elem_search_button = elem_history.find_element(By.NAME, "submit-search-btn")
    elem_search_button.click()



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
chrome_options.add_experimental_option("prefs", prefs)                                              # 定義したプリファレンスを設定オプションに追加



# ------ ChromeDriver の起動 ------
driver_path = ChromeDriverManager().install()
service =Service(driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get('https://point.rakuten.co.jp/history/?l-id=point_top_history_pc')
wait = WebDriverWait(driver, 10)
element = wait.until(EC.visibility_of_all_elements_located)



###################################################
###############   メイン処理開始   #################
###################################################
sub_win = tk.Tk()
sub_win.withdraw()
sub_win.attributes("-topmost",True)
messagebox.showwarning("手動でログイン処理をお願いします!","ログイン画面上の「ユーザID」と「パスワード」を入力し、手動でログインしてください。\nログインが終わり画面が遷移したら、このポップアップの「OK」をクリックしてください。", parent=sub_win)
sub_win.attributes("-topmost",False)
driver.maximize_window()



# ------ 期間の絞り込みFrom ~ To ------
element = wait.until(EC.visibility_of_all_elements_located)
scroll(driver) #　関数（scroll）呼び出し
PeriodList = get_period_elements(driver) #　関数（get_period_elements）呼び出し
select = Select(PeriodList[2]) 
fdate = temp_data["year"]+"年"+temp_data["month"]+"月"
select.select_by_visible_text(fdate)
select = Select(PeriodList[3]) 
tdate = fdate
select.select_by_visible_text(tdate)



# ------ 項目の絞り込み(獲得ポイント) 通常ポイント / 期間限定ポイント ------
PeriodList[0][0].click()
PeriodList[0][2].click()
perform_search(driver) #　関数（perform_search）呼び出し
print('>\n***********************************************')
print('***********  獲得ポイント抽出中...  ***********')
print('***********************************************')
max_page = get_max_page(driver,0) #　関数（get_max_page）呼び出し
FName = FilePath+fdate+"_楽天ポイント獲得"
ACQs = scrape_point_history(driver,max_page,FName) #　関数（scrape_point_history）呼び出し



# ------ 項目の絞り込み(利用ポイント) 通常利用 ------
element = wait.until(EC.visibility_of_all_elements_located)
scroll(driver) #　関数（scroll）呼び出し
PeriodList = get_period_elements(driver) #　関数（get_period_elements）呼び出し
select = Select(PeriodList[2])
select.select_by_visible_text(fdate)
select = Select(PeriodList[3])
select.select_by_visible_text(tdate)
PeriodList[0][0].click() #　項目の絞り込み(獲得ポイント) 初期化
PeriodList[0][2].click() #　項目の絞り込み(獲得ポイント) 初期化
PeriodList[1][0].click() #　項目の絞り込み(利用ポイント) 通常利用
perform_search(driver) #　関数（perform_search）呼び出し
print('>\n***********************************************')
print('***********  利用ポイント抽出中...  ***********')
print('***********************************************')
max_page = get_max_page(driver,0) #　関数（get_max_page）呼び出し
FName = FilePath+fdate+"_楽天ポイント利用"
USEs = scrape_point_history(driver,max_page,FName) #　関数（scrape_point_history）呼び出し



# ------ データフレーム ------
#　獲得ポイントデータフレーム
acq_df = pd.DataFrame()
acq_df["日付"] = ACQs[0]
acq_df["サービス"] = ACQs[1]
acq_df["内容"] = ACQs[2]
acq_df["アクション"] = ACQs[3]
acq_df["ポイント"] = ACQs[4]
acq_df["備考"] = ACQs[5]
#　利用ポイントデータフレーム
use_df = pd.DataFrame()
use_df["日付"] = USEs[0]
use_df["サービス"] = USEs[1]
use_df["内容"] = USEs[2]
use_df["アクション"] = USEs[3]
use_df["ポイント"] = USEs[4]
use_df["備考"] = USEs[5]
#　データフレーム分割
acq_points_df = acq_df[~acq_df['アクション'].str.contains('失効')]                    # 獲得ポイント
exp_points_df = acq_df[acq_df['アクション'].str.contains('失効')]                     # 失効ポイント
used_points_df = use_df[~use_df['内容'].str.contains('楽天キャッシュ')]               # 利用ポイント（内容：楽天キャッシュを除外）
used_points_df = used_points_df[~used_points_df['備考'].str.contains('キャッシュ')]   # 利用ポイント（備考：キャッシュを除外）
cash_points1_df = use_df[use_df['内容'].str.contains('楽天キャッシュ')]               # キャッシュポイント（内容：楽天キャッシュ）
cash_points2_df = use_df[use_df['備考'].str.contains('キャッシュ')]                   # キャッシュポイント（備考：キャッシュ）
concat_df = pd.concat([cash_points1_df, cash_points2_df])                           # キャッシュポイント（結合）
#　Excelにデータフレームを保存
writer = pd.ExcelWriter(FilePath+fdate+"_楽天ポイント.xlsx", engine='xlsxwriter')
acq_points_df.to_excel(writer, sheet_name = "獲得ポイント履歴")
exp_points_df.to_excel(writer, sheet_name = "失効ポイント履歴")
used_points_df.to_excel(writer, sheet_name = "利用ポイント履歴")
concat_df.to_excel(writer, sheet_name = "利用ポイント(キャッシュ)履歴")
writer.save()
driver.quit()
