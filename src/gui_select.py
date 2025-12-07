# gui_select.py
from typing import List, Tuple
from pathlib import Path

def select_targets_gui(targets: List[Tuple[str, Path, str]]) -> List[Tuple[str, Path, str]]:
    """
    targets をチェックボックス付きで表示し、選ばれたものだけ返す。
    tkinter標準のみ。
    """
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("印刷するファイルを選択")
    root.geometry("600x600")

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

    # --- チェックボックス行を生成 ---
    vars_ = []  # (BooleanVar, (kind, path))
    for kind, path, fname in targets:
        v = tk.BooleanVar(value=True)  # デフォルト全選択
        text = f"[{kind.upper():4}]  {fname}"
        cb = ttk.Checkbutton(scroll_frame, text=text, variable=v)
        cb.pack(anchor="w", padx=8, pady=2)
        vars_.append((v, (kind, path,fname)))

    # --- 下部ボタン群 ---
    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", pady=6)

    def select_all():
        for v, _ in vars_:
            v.set(True)

    def clear_all():
        for v, _ in vars_:
            v.set(False)

    selected: List[Tuple[str, Path, str]] = []

    def done():
        nonlocal selected
        selected = [item for v, item in vars_ if v.get()]
        root.destroy()

    ttk.Button(btn_frame, text="全選択", command=select_all).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="全解除", command=clear_all).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="選択したものを印刷", command=done).pack(side="right", padx=5)

    root.mainloop()
    return selected

