import tkinter
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import json
import codecs
import time
from datetime import datetime
import subprocess

# 参照ボタンの動作
def ask_folder():
    path = filedialog.askdirectory()
    folder_path.set(path)


# 実行ボタンの動作
def app():
    res = messagebox.askquestion("ポイント取得開始","ポイント取得を開始しますか？")
    if res == "yes":
        point=point_comb.get()
        month=month_comb.get()
        year=year_comb.get()
        now_month = dt_now.strftime("%m")
        old_month = int(now_month) - 5

        output_dir=folder_path.get()
        output_dir = output_dir.replace("/","\\")+"\\"

        # pyファイルを実行するバラメータ「temp.json」を作成
        str = {
        "FileDir" : output_dir,
        
        "year" : year,
        "month" : month
        }
        jsonfile = codecs.open('01.Config\\00.temp.json', 'w', 'utf-8')
        json.dump(str, jsonfile, indent=2, ensure_ascii=False)
        jsonfile.close()
        msg = point+"の取得を開始しました!"
        var.set(msg) 

        time.sleep(3)

        # 日付チェック
        flg=1
        while flg == 1:
            if int(now_year) == int(year) and int(now_month) < int(month):
                    messagebox.showwarning("取得対象年月を選択しなおしてください","未来の日付を選択してるため、取得できません!!")
                    break
            elif int(now_year) > int(year) and int(now_month) > int(month):
                    messagebox.showwarning("取得対象年月を選択しなおしてください","過去１年より前の日付を選択してるため、取得できません!!")
                    break
            elif point == "Pontaポイント":   # Pontaは半年間の履歴までしか出力できない為、判定
                if old_month > 0 and old_month > int(month):                                         # 半年前が年をまたがない場合
                    messagebox.showwarning("取得対象年月を選択しなおしてください","Pontaポイントは過去半年より前の日付は取得できません!!")
                    break
                elif old_month <= 0 and int(now_year) > int(year) and old_month+12 > int(month):     # 半年前が年をまたぐ場合
                    messagebox.showwarning("取得対象年月を選択しなおしてください","Pontaポイントは過去半年より前の日付は取得できません!!")
                    break
                # 出力するポイントを判定
                else:
                    main_win.update()
                    subprocess.run(['python', '02.Script\\03.ponta.py'])
                    messagebox.showinfo("ポイント取得完了", point+"の取得が完了しました!!")
                    subprocess.Popen(["explorer", output_dir])          
                    flg=0
            # 出力するポイントを判定
            else:
                main_win.update()
                if point == '楽天ポイント':
                    subprocess.run(['python', '02.Script\\01.rakuten.py'])
                elif point == 'Amazonポイント':
                    subprocess.run(['python', '02.Script\\02.amazon.py'])
                elif point == 'ビックポイント':
                    subprocess.run(['python', '02.Script\\04.bic.py'])
                else:
                    subprocess.run(['python', '02.Script\\05.d.py'])
                messagebox.showinfo("ポイント取得完了", point+"の取得が完了しました!!")
                subprocess.Popen(["explorer", output_dir])
                flg=0
        var.set("※注1 : 過去１年までの履歴を出力できます\n※注2 : Pontaポイントは半年までしか出力できません") 
        main_win.update()
    else:
        messagebox.showinfo('キャンセル','ポイント取得をキャンセルしました！')



###################################################
###############   ウインドウ作成   ###############
###################################################
# メインウィンドウ
main_win = tkinter.Tk()
photo = tkinter.PhotoImage(file = "01.Config\\image\\robot.gif")
main_win.iconphoto(False, photo)
main_win.title("ポイント取得ツール")
# ウィンドウのサイズを設定
window_width = 500
window_height = 250
main_win.geometry(f"{window_width}x{window_height}")
# 画面のサイズを取得
screen_width = main_win.winfo_screenwidth()
screen_height = main_win.winfo_screenheight()
# ウィンドウの幅の4分の1の値を計算
quarter_width = window_width // 4
# ウィンドウの高さの2分の1の値を計算
half_height = window_height // 2
# ウィンドウを右下にずらす座標を計算
x_coordinate = screen_width - window_width - quarter_width
y_coordinate = screen_height - window_height - half_height
# ウィンドウを計算した座標に配置
main_win.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
main_win.resizable(0,0)
main_win.attributes("-topmost",True)
# メインフレーム
main_frm = ttk.Frame(main_win)
main_frm.grid(column=0, row=0, sticky=tkinter.NSEW, padx=15, pady=10)
# パラメータ
folder_path = tkinter.StringVar()
var = tkinter.StringVar(value="※注1 : 過去１年までの履歴を出力できます\n※注2 : Pontaポイントは半年までしか出力できません")
point_List=['楽天ポイント','Amazonポイント','Pontaポイント','dポイント']
dt_now = datetime.now()
now_year = dt_now.year



###################################################
###############   ウィジェット作成   ###############
###################################################
# 取得対象ポイント
get_point = ttk.Label(main_frm, text="取得対象ポイント")
point_comb = ttk.Combobox(main_frm, values=point_List, width=15, state="readonly")
point_comb.current(0)
# 取得対象年度
year_List=[now_year-1,now_year]
get_year = ttk.Label(main_frm, text="取得対象年度")
year_comb = ttk.Combobox(main_frm, values=year_List, width=10, state="readonly")
year_comb.current(0)
# 取得対象月
month_List = tuple(range(1,13))
get_month = ttk.Label(main_frm, text="取得対象月")
month_comb = ttk.Combobox(main_frm, values=month_List, width=10, state="readonly")
month_comb.current(0)
# フォルダパス
folder_label = ttk.Label(main_frm, text="出力フォルダ")
folder_box = ttk.Entry(main_frm, textvariable=folder_path)
folder_btn = ttk.Button(main_frm, text="参照", command=ask_folder)
# メッセージボックス
message = ttk.Label(main_frm, textvariable=var)
# 実行ボタン
app_btn = ttk.Button(main_frm, text="実行", command=app)



###################################################
###############   ウィジェット配置   ###############
###################################################
# 取得対象ポイント
get_point.grid(column=0, row=0, pady=10)
point_comb.grid(columnspan=2, column=1, row=0, sticky=tkinter.W, padx=10)
# 取得対象年度
get_year.grid(column=0, row=1, pady=10)
year_comb.grid(column=1, row=1, sticky=tkinter.W, padx=10)
# 取得対象月
get_month.grid(column=2, row=1, pady=10)
month_comb.grid(column=3, row=1,  padx=10)
# フォルダパス
folder_label.grid(column=0, row=2, pady=10)
folder_box.grid(columnspan=3, column=1, row=2, sticky=tkinter.EW, padx=10)
folder_btn.grid(column=4, row=2, padx=10)
# 出力内容
message.grid(columnspan=5, column=0, row=3, pady=15)
# 実行ボタン
app_btn.grid(column=2, row=4, pady=10)
# 配置設定
main_win.columnconfigure(0, weight=1)
main_win.rowconfigure(0, weight=1)
main_frm.columnconfigure(1, weight=1) # 真ん中の列（column=1）が延びる
main_win.mainloop()