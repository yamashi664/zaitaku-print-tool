# hokokusyo_print.py
from pathlib import Path
from datetime import datetime
import sys

import module1 as m
import gui_select as g


def main():
    # 1) 設定読み込み
    cfg = m.load_config()
    parent_folder = Path(cfg["parent_folder"])
    printer_name = cfg["printer_name"]
    soffice_path = Path(cfg["soffice_path"])
    pdftoprinter_path = Path(cfg["pdftoprinter_path"])
    queue_limit = int(cfg.get("queue_limit", 6))
    sleep_full = float(cfg.get("sleep_when_queue_full_sec", 1))

    # 2) 日付入力
    while True:
        s = input("印刷対象とするPDFの更新日付を入力してください (YYYY/MM/DD): ")
        try:
            target_date = m.parse_input_date(s)
            break
        except ValueError as e:
            print(e)

    print(f"\n対象：{target_date.strftime('%Y/%m/%d')} に更新されたPDF\n")

    # 3) 対象収集
    targets = m.collect_targets(parent_folder, target_date)
    print(f"印刷対象件数: {len(targets)}")

    # 4) GUIで選択
    targets = g.select_targets_gui(targets)
    print(f"選択された印刷件数: {len(targets)}")
    if not targets:
        print("何も選択されなかったので終了します。")
        return
    
    # 5) 印刷実行
    ok_count = 0
    ng_count = 0

    for kind, path, fname in targets:
        try:
            # キュー上限付き投入（未実装なら即スルー）
            m.wait_if_queue_full(printer_name, queue_limit, sleep_full)

            if kind == "pdf":
                print(f"[PDF] {path}")
                m.print_pdf_with_pdftoprinter(pdftoprinter_path, printer_name, path)
            else:
                print(f"[WORD] {path}")
                m.print_word_with_soffice(soffice_path, printer_name, path)

            ok_count += 1
            
        except Exception as e:
            ng_count += 1
            print(f"  -> 失敗: {e}")


    print("\n=== 完了 ===")
    print(f"成功: {ok_count} / 失敗: {ng_count}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"致命的エラー: {e}")
        sys.exit(1)

