# hokokusyo_print.py
from pathlib import Path
import sys
import shutil
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import module1 as m
import gui_select as gs
import time

# ===== ZIPファイルと薬局を順番に選択する画面（D&D対応版） =====
def select_pharmacy_and_zip(pharmacies: list) -> tuple:
    selected_data = {"pharmacy": None, "zip_path": None}
    
    root = TkinterDnD.Tk()
    root.title("報告書印刷ツール (ZIP・送付状対応版)")
    root.geometry("450x370")  # ボタン位置調整のため少し高さを最適化
    root.resizable(False, False)
    
    # 画面を中央に配置
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (450 // 2)
    y = (root.winfo_screenheight() // 2) - (370 // 2)
    root.geometry(f"+{x}+{y}")

    # --- レイアウト ---
    
    # ステップ1: タイトル案内
    lbl1 = tk.Label(root, text="1. 処理するZIPファイルをドロップするか選択してください", font=("MS Gothic", 10, "bold"))
    lbl1.pack(pady=(15, 5))
    
    # 視覚的なドロップエリア（白い四角のボックス）
    drop_frame = tk.Frame(root, width=350, height=100, bg="white", bd=2, relief="groove")
    drop_frame.pack(pady=(5, 2)) # ボタンとの間隔を少し狭くしました
    drop_frame.pack_propagate(False)

    # 四角の真ん中に配置する「＋」と案内メッセージ
    lbl_plus = tk.Label(drop_frame, text="＋", bg="white", fg="#0078d7", font=("MS Gothic", 24, "bold"))
    lbl_plus.pack(pady=(15, 0))
    
    lbl_zip_name = tk.Label(drop_frame, text="ここにZIPファイルをドロップ", bg="white", fg="gray", font=("MS Gothic", 9))
    lbl_zip_name.pack(pady=(0, 10))

    # --- ボタン・D&Dの動作設定 ---
    
    def activate_next_step(file_path_str: str):
        clean_path = file_path_str.strip('{}')
        path_obj = Path(clean_path)
        
        if path_obj.suffix.lower() != '.zip':
            messagebox.showwarning("警告", "ZIPファイル（.zip）をドロップしてください。")
            return
            
        selected_data["zip_path"] = path_obj
        
        # ファイルが選ばれたら、白い四角の中身を「緑色の成功表示」に変える
        drop_frame.config(bg="#e6f4ea")
        lbl_plus.config(text="✓", bg="#e6f4ea", fg="green")
        lbl_zip_name.config(text=f"選択中: {path_obj.name}", bg="#e6f4ea", fg="green", font=("MS Gothic", 9, "bold"))
        
        # 薬局選択と確定ボタンを有効化する
        lbl2.config(fg="black")
        combo.config(state="readonly")
        btn_submit.config(state="normal", bg="#0078d7")

    def on_select_zip():
        file_path = filedialog.askopenfilename(
            title="報告書ZIPファイルを選択",
            filetypes=[("ZIPファイル", "*.zip"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            activate_next_step(file_path)

    def on_drop_zip(event):
        if event.data:
            activate_next_step(event.data)

    def on_submit():
        idx = combo.current()
        if idx == -1:
            messagebox.showwarning("警告", "薬局を選択してください。")
            return
        
        selected_data["pharmacy"] = pharmacies[idx]
        root.destroy()

    # ★ 「またはファイルを選択...」のボタンをドロップボックスのすぐ下に配置
    btn_zip = tk.Button(root, text="またはファイルを選択...", command=on_select_zip, font=("MS Gothic", 9))
    btn_zip.pack(pady=(2, 5))

    # ステップ2: 薬局選択
    lbl2 = tk.Label(root, text="2. 送付状に記載する薬局を選択してください", font=("MS Gothic", 10, "bold"), fg="gray")
    lbl2.pack(pady=(15, 5))

    pharmacy_names = [p.get("pharmacy_name", "名称未設定") for p in pharmacies]
    combo = ttk.Combobox(root, values=pharmacy_names, state="disabled", width=40, font=("MS Gothic", 10))
    combo.pack(pady=5)
    if pharmacy_names:
        combo.current(0)

    # 画面全体および「白い四角（各ラベル含む）」のどこに落としても反応するように紐付け
    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop_zip)
    drop_frame.drop_target_register(DND_FILES)
    drop_frame.dnd_bind('<<Drop>>', on_drop_zip)
    lbl_plus.drop_target_register(DND_FILES)
    lbl_plus.dnd_bind('<<Drop>>', on_drop_zip)
    lbl_zip_name.drop_target_register(DND_FILES)
    lbl_zip_name.dnd_bind('<<Drop>>', on_drop_zip)

    btn_submit = tk.Button(root, text="この薬局で印刷処理を開始", command=on_submit, state="disabled", bg="gray", fg="white", font=("MS Gothic", 10, "bold"), padx=10, pady=5)
    btn_submit.pack(pady=(15, 10))

    def on_closing():
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
    
    return selected_data["pharmacy"], selected_data["zip_path"]

def main():
    # 1) 設定読み込み（config.json が無い/壊れている時はGUIで通知）
    try:
        cfg = m.load_config()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "設定ファイルエラー",
            f"config.json を読み込めませんでした。\n\n"
            f"exe と同じフォルダに config.json があるか確認してください。\n\n"
            f"詳細: {e}"
        )
        root.destroy()
        return
    
    printer_name = cfg["printer_name"]
    soffice_path = Path(cfg["soffice_path"])
    pdftoprinter_path = Path(cfg["pdftoprinter_path"])
    queue_limit = int(cfg.get("queue_limit", 6))
    queue_wait_interval_sec = float(cfg.get("queue_wait_interval_sec", 1))
    
    # config.jsonから薬局リストを取得（なければ空リスト）
    pharmacies = cfg.get("pharmacies", [])
    if not pharmacies:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("設定エラー", "config.json 内に薬局情報（pharmacies）が設定されていません。")
        root.destroy()
        return

    # 2) 薬局の選択 ＆ ZIPファイルの選択
    selected_pharmacy, zip_path = select_pharmacy_and_zip(pharmacies)
    if selected_pharmacy is None or zip_path is None:
        print("キャンセルのため終了します。")
        return

    # ★★★【追加】薬局選択の直後、印刷リスト表示の前に送付状テンプレートの有無をチェック ★★★
    template_path = m.base_dir() / "fax_template.docx"
    if not template_path.exists():
        # 目立つ警告ダイアログを表示
        root = tk.Tk()
        root.withdraw()
        proceed = messagebox.askyesno(
            "【警告】送付状テンプレートが見つかりません",
            f"送付状の型紙ファイルが見つかりません。\n"
            f"場所: {template_path}\n"
            f"送付状のファイル名は'fax_template.docx'にしてください。\n\n"
            f"※このまま続行すると、送付状は作成されず【PDFのみ】の印刷リストになります。\n"
            f"このまま処理を続行しますか？"
        )
        root.destroy()
        
        if not proceed:
            print("ユーザーにより処理が中断されました。")
            return  # 「いいえ」を押した場合は安全に終了する
        else:
            print("警告を了承の上、送付状なしで続行します。")

    print(f"選択された薬局: {selected_pharmacy.get('pharmacy_name')}")
    print(f"対象ZIPファイル: {zip_path.name}")

    temp_dir = None
    try:
        # 3) ZIPファイルの解凍 ＆ 各フォルダごとの送付状自動生成
        temp_dir, print_list = m.process_zip_and_generate_fax(zip_path, cfg, selected_pharmacy)
        print(f"展開・生成された総印刷対象件数: {len(print_list)}")

        if not print_list:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("情報", "ZIPファイル内にPDFを含む有効なフォルダが見つかりませんでした。")
            root.destroy()
            return

        # 4) GUIで最終確認・選択（印刷順はすでに PDF ➡️ 送付状 になっています）
        selected = gs.select_targets_gui(print_list)
        print(f"ユーザーが選択した印刷件数: {len(selected)}")
        if not selected:
            print("何も選択されなかったので終了します。")
            return
        
        # 5) 印刷実行（進捗GUIつき）
        from print_progress_gui import run_print_with_gui

        def _print_pdf(path):
            m.wait_if_queue_full(printer_name, queue_limit, queue_wait_interval_sec)
            m.print_pdf_with_pdftoprinter(pdftoprinter_path, printer_name, path)
            time.sleep(2.0)

        def _print_word(path):
            m.wait_if_queue_full(printer_name, queue_limit, queue_wait_interval_sec)
            m.print_word_with_soffice(soffice_path, printer_name, path)
            time.sleep(4.0)

        # GUI付きで印刷を走らせる
        ok = run_print_with_gui(
            selected,
            print_pdf_func=_print_pdf,
            print_word_func=_print_word,
            printer_name=printer_name
        )

        if ok:
            print("\n=== 完了 ===")
            print(f"成功: {len(selected)} / 失敗: 0")
        else:
            print("\n=== 終了 ===")
            print("中止または失敗がありました。")

    finally:
        # 6) 【超重要】使い終わった一時フォルダを綺麗さっぱり削除してパソコンを汚さない
        if temp_dir and temp_dir.exists():
            print(f"一時フォルダを削除しています...: {temp_dir}")
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"致命的エラー: {e}")
        sys.exit(1)