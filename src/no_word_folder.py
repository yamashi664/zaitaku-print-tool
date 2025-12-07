# no_word_folder.py
import tkinter as tk
from tkinter import ttk
from typing import List

def no_word(no_word_folder: List[str]):
    """
    wordファイルのないフォルダ名を列挙する
    """
    
    root = tk.Tk()
    root.title("wordファイルのないフォルダ")
    root.geometry("600x300")

    # --- スクロール可能な領域を作る ---
    container = ttk.Frame(root)
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # --- 対象フォルダ名の羅列 ---
    li = []
    for folder_name in no_word_folder:
        text = f"・{folder_name}"
        lbl = tk.Label(scroll_frame, text=text)
        lbl.pack()
        li.append(lbl)

    # --- 下部ボタン群 ---
    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", pady=1)    

    flag = False
    def cont():
        nonlocal flag
        flag = True
        root.destroy()

    def end():
        root.destroy()

    end_btn = ttk.Button(btn_frame, text="終了", command=end)
    end_btn.pack(side="right",padx=5)
    continue_btn = ttk.Button(btn_frame, text="続行", command=cont)
    continue_btn.pack(side="right",padx=5)
    
    root.mainloop()
    return flag

