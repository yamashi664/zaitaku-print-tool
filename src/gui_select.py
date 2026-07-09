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

    # --- マウスホイールでのスクロールを有効化する設定 ---
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # --- ttkチェックボックスの見た目を維持したまま文字サイズを変更 ---
    style = ttk.Style()
    style.configure("Custom.TCheckbutton", font=("MS Gothic", 11))

    # --- チェックボックス行を生成 ---
    vars_ = []  # (BooleanVar, (kind, path))
    for kind, path, fname in targets:
        v = tk.BooleanVar(value=True)  # デフォルト全選択
        
        # =================================================================
        # ★★★【修正】PDFの場合、前後を削りつつ「CCC」と「DDD」を入れ替える ★★★
        # =================================================================
        display_name = fname
        if kind == "pdf":
            # 拡張子 (.pdf) を取り除いてアンダースコアで分割
            pure_name = Path(fname).stem
            parts = pure_name.split("_")
            
            # 必要な要素数（AAA, BBB, CCC, DDD, GGG の最低5つ）があるかチェック
            if len(parts) >= 5:
                # parts[2]がCCC、parts[3]がDDD なので、順序を入れ替えて残りの要素(parts[4:-1])と結合
                swapped_parts = [parts[3], parts[2]] + parts[4:-1]
                display_name = "_".join(swapped_parts)
            elif len(parts) > 3:
                # 万が一要素が足りない場合は、前後を落とすだけの安全処理
                display_name = "_".join(parts[2:-1])
        
        text = f"[{kind.upper():4}]  {display_name}"
        
        # 元の ttk.Checkbutton のまま、カスタムスタイルを適用
        cb = ttk.Checkbutton(
            scroll_frame, 
            text=text, 
            variable=v, 
            style="Custom.TCheckbutton"
        )
        cb.pack(anchor="w", padx=8, pady=2)
        
        # 内部データとしては、加工前の正しい path や fname をそのまま保持します
        vars_.append((v, (kind, path, fname)))

    # --- 下部ボタン群 ---
    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", pady=8)

    # ボタン用共通フォント
    btn_font = ("MS Gothic", 10, "bold")

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

    # 各ボタンの配置
    tk.Button(btn_frame, text="全選択", command=select_all, font=btn_font, padx=5).pack(side="left", padx=5)
    tk.Button(btn_frame, text="全解除", command=clear_all, font=btn_font, padx=5).pack(side="left", padx=5)
    tk.Button(btn_frame, text="選択したものを印刷", command=done, font=btn_font, bg="#0078d7", fg="white", padx=10, pady=2).pack(side="right", padx=5)

    root.mainloop()
    return selected