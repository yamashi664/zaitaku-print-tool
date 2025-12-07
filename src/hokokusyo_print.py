# hokokusyo_print.py
from pathlib import Path
from datetime import datetime
import sys

import module1 as m
import gui_select as gs
import gui_input as gi
import no_word_folder as nw


def main():
    # 1) 設定読み込み
    cfg = m.load_config()
    parent_folder = Path(cfg["parent_folder"])
    printer_name = cfg["printer_name"]
    soffice_path = Path(cfg["soffice_path"])
    pdftoprinter_path = Path(cfg["pdftoprinter_path"])
    queue_limit = int(cfg.get("queue_limit", 6))
    queue_wait_interval_sec = float(cfg.get("queue_wait_interval_sec", 1))

    # 2) 日付入力
    target_date = gi.input_date_gui()
    if target_date is None:
        print("キャンセルのため終了")
        return
    else:
        print(f"\n対象：{target_date.strftime('%Y/%m/%d')} に更新されたPDF\n")        

    # 3) 対象収集
    targets, no_word_folder = m.collect_targets(parent_folder, target_date)
    print(f"印刷対象件数: {len(targets)}")

    # 4) wordファイルの無いフォルダの表示
    if no_word_folder:
        if not nw.no_word(no_word_folder):
            return

    # 5) GUIで選択
    selected = gs.select_targets_gui(targets)
    print(f"選択された印刷件数: {len(selected)}")
    if not selected:
        print("何も選択されなかったので終了します。")
        return
    
    # 6) 印刷実行（進捗GUIつき）
    from print_progress_gui import run_print_with_gui

    # --- 既存の module1 の印刷関数をGUI用にラップ ---
    def _print_pdf(path):
        # キュー上限付き投入（元コードと同じ）
        m.wait_if_queue_full(printer_name, queue_limit, queue_wait_interval_sec)
        m.print_pdf_with_pdftoprinter(pdftoprinter_path, printer_name, path)

    def _print_word(path):
        m.wait_if_queue_full(printer_name, queue_limit, queue_wait_interval_sec)
        m.print_word_with_soffice(soffice_path, printer_name, path)

    # --- GUI付きで印刷を走らせる ---
    ok = run_print_with_gui(
        selected,
        print_pdf_func=_print_pdf,
        print_word_func=_print_word,
        printer_name=printer_name
    )

    # --- 結果表示（GUIは「キュー空」まで待ってから完了になる） ---
    if ok:
        print("\n=== 完了 ===")
        print(f"成功: {len(selected)} / 失敗: 0")
    else:
        # 中止 or 一部失敗
        print("\n=== 終了 ===")
        print("中止または失敗がありました。")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"致命的エラー: {e}")
        sys.exit(1)

