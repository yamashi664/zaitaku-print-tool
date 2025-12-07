# gui_input.py
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

def input_date_gui() -> datetime:
    """
    YYYY/MM/DDをGUIで入力させて文字列で返す。
    キャンセルされたらNoneを返す。
    """    

    def on_ok():
        s = entry.get().strip()
        try:
            dt = datetime.strptime(s, "%Y/%m/%d")
        except ValueError:
            messagebox.showerror("形式エラー", "日付は YYYY/MM/DD 形式で入力してください")
            return
        result["target_date"] = dt
        root.destroy()
        
    def on_cancel():
        result["target_date"] = None
        root.destroy()

    #今日の日付を取得
    today_datetime = datetime.now().strftime("%Y/%m/%d")

    root = tk.Tk()
    root.title("日付入力")
    root.geometry("320x140")

    result = {"target_date": None}
    
    #指示文
    lbl1 = tk.Label(
        root,
        text="印刷対象とする報告書ファイルの更新日付を"
    )
    lbl1.pack()

    lbl2 = tk.Label(
        root,
        text="YYYY/MM/DD形式で入力してください"
    )
    lbl2.pack()
    
    #入力窓
    entry = tk.Entry(
        root,width=20,
        justify="center",
        font=("Segoe UI", 12)
    )
    entry.pack(pady=12)
    entry.insert(0, today_datetime)
    
    #ボタン設定
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=5)

    OK_btn = tk.Button(
        btn_frame,
        text="実行",
        width=10,
        command=on_ok
    )
    OK_btn.pack(side="left",padx=6)

    cancel_btn = tk.Button(
        btn_frame,
        text="キャンセル",
        width=10,
        command=on_cancel
    )
    cancel_btn.pack(side="left",padx=6)
    
    root.mainloop()
    return result["target_date"]
    
    
    

    
