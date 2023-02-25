import tkinter as tk
from tkinter import messagebox
import openpyxl
import datetime

# エクセル開く
book = openpyxl.load_workbook('./sub-job.xlsx')

# シート開く
sheet = book['Sheet1']

# 既に書いてあるデータを消す
for data in sheet["A2:E2"]:
  for cell in data:
    cell.value = ""
  book.save('./sub-job.xlsx')

# 画面作成
frame = tk.Tk()
frame.title("副業勤怠")
frame.geometry('400x300')

# 勤務ボタン押下時
def change_goout_button():
  
  # 終了時
  if start_btn['state'] == "disabled":
    end_btn['state'] = "disabled"

    dt = datetime.datetime.now()
    sheet['C2'] = dt.strftime('%Y/%m/%d %H:%M')
    book.save('./sub-job.xlsx')
    
    frame.destroy()

  # 開始時
  else:
    start_btn['state'] = "disabled"
    end_btn['state'] = "normal"
    start_stay_btn['state'] = "normal"
    
    dt = datetime.datetime.now()
    sheet['A2'] = dt.strftime('%m月%d日')
    sheet['B2'] = dt.strftime('%Y/%m/%d %H:%M')
    book.save('./sub-job.xlsx')
    
# 休憩ボタン押下時
def change_stay_button():
  
  # 休憩終了時
  if start_stay_btn['state'] == "disabled":
    end_stay_btn['state'] = "disabled"

    dt = datetime.datetime.now()
    sheet['E2'] = dt.strftime('%Y/%m/%d %H:%M')
    book.save('./sub-job.xlsx')

  # 休憩開始時
  else:
    start_stay_btn['state'] = "disabled"
    end_stay_btn['state'] = "normal"
    
    dt = datetime.datetime.now()
    sheet['D2'] = dt.strftime('%Y/%m/%d %H:%M')
    book.save('./sub-job.xlsx')
    

# ボタン設置
start_btn = tk.Button(frame, text="勤務開始！", font=(20), bg='#b0c4de',command=change_goout_button)
end_btn = tk.Button(frame, text="勤務終了！", font=(20), command=change_goout_button, state="disabled")
start_stay_btn = tk.Button(frame, text="休憩開始！", font=(20), bg='#fff0f5', command=change_stay_button, state="disabled")
end_stay_btn = tk.Button(frame, text="休憩終了！", font=(20), command=change_stay_button, state="disabled")

start_btn.pack(pady=14)
end_btn.pack(pady=14)
start_stay_btn.pack(pady=14)
end_stay_btn.pack(pady=14)

# ×ボタン押下時のアラート
def close_alert():
  if messagebox.askokcancel("最終確認", "閉じてOK?\n閉じると勤怠データはなくなるよ"):
    frame.destroy()

frame.protocol("WM_DELETE_WINDOW", close_alert)

# 画面表示
frame.mainloop()

