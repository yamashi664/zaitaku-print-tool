# module1.py
import sys
import json
import time
import subprocess
from pathlib import Path
from datetime import datetime
from typing import List, Tuple, Optional

# ===== パス基準（exeの隣を見るための定番） =====
def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# ===== 設定読み込み =====
def load_config() -> dict:
    cfg_path = base_dir() / "config.json"
    if not cfg_path.exists():
        raise FileNotFoundError(f"config.json が見つかりません: {cfg_path}")
    return json.loads(cfg_path.read_text(encoding="utf-8"))


# ===== ファイル収集 =====
def collect_targets(parent_folder: Path, target_date: datetime) -> Tuple[List[Tuple[str, Path, str]], List[str]]:
    """
    バッチ仕様：
      - parent 配下の各サブフォルダを走査
      - target_dateに更新されたPDFを印刷対象
      - そのサブフォルダでPDFが1つでも対象になったら、同フォルダのwordファイルも全部対象
    戻り値: [("pdf", pdf_path, pdf_name), ("word", docm_path, pdf_name), ...]
      - サブフォルダに対象PDFがあるのにwordファイルがない場合、そのサブフォルダ名も返す
    """
    targets: List[Tuple[str, Path, str]] = []
    no_word_folder: List[str] = []

    for sub in sorted(parent_folder.iterdir()):
        if not sub.is_dir():
            continue

        # PDFの中から更新日がtarget_dateのものだけ拾う
        pdfs = sorted(sub.glob("*.pdf"))
        recent_pdfs = []
        for p in pdfs:
            mtime = datetime.fromtimestamp(p.stat().st_mtime) #pdfファイルの更新時刻を取得、datetime型に変換
            if mtime.date() == target_date.date():
                recent_pdfs.append(p)

        if recent_pdfs:
            # PDFを対象に追加
            for p in recent_pdfs:
                targets.append(("pdf", p, p.name))
            
            # PDFがあったフォルダだけwordファイルを対象に追加
            doc_exts = {".doc", ".docx", ".docm"}
            docms = [
                p for p in sorted(sub.glob("*.doc*")) 
                if p.suffix.lower() in doc_exts and not p.name.startswith("~$")
                ]
            if docms:
                for w in docms:
                   targets.append(("word", w, w.name))
            else:
                no_word_folder.append(sub.name)
            
    return targets, no_word_folder


# ===== 印刷：PDFtoPrinter =====
def print_pdf_with_pdftoprinter(pdftoprinter_path: Path, printer_name: str, pdf_path: Path):
    if not pdftoprinter_path.exists():
        raise FileNotFoundError(f"PDFtoPrinter.exe が見つかりません: {pdftoprinter_path}")
    # PDFtoPrinter.exe "file.pdf" "Printer Name"
    subprocess.run(
        [str(pdftoprinter_path),str(pdf_path), printer_name],
        check=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )

# ===== 印刷：LibreOffice headless =====
def print_word_with_soffice(soffice_path: Path, printer_name: str, word_path: Path):
    if not soffice_path.exists():
        raise FileNotFoundError(f"soffice.com が見つかりません: {soffice_path}")
    # soffice --headless --pt "Printer Name" "file.docm"
    subprocess.run(
        [str(soffice_path), "--headless", "--pt", printer_name, str(word_path)],
        check=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )


# ===== キュー制御（骨組み） =====
def get_print_queue_size(printer_name: str) -> int:
    """
    PowerShellで印刷キュー内ジョブ数を取得（Windows標準）。
    失敗時は None ではなく大きめの値を返して安全側に倒す。
    """
    
    try:
        # Get-PrintJob は Windows 8/Server2012 以降で使える標準コマンド
        # 返り値はジョブ数のみを数値で出す
        safe_name = printer_name.replace("'", "''")
        cmd = [
            "powershell",
            "-NoProfile",
            "-Command",
            f"(Get-PrintJob -PrinterName '{safe_name}' | Measure-Object).Count"
        ]
        out = subprocess.check_output(
            cmd,
            stderr=subprocess.STDOUT,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        ).strip()
        if out == "":
            return 0
        return int(out)

    except Exception:
        # 判定不能なら「多い」扱いにして待ち側へ寄せる
        return 9999


def wait_if_queue_full(printer_name: str, queue_limit: int, queue_wait_interval_sec: float):
    """
    キュー上限付き高速投入の骨組み。
    queue_limit 以上たまっていたら空くまで sleep_sec ごとに待つ。
    """
    size = get_print_queue_size(printer_name)
    if size is None:
        return  # 未実装なら待たない
    while size >= queue_limit:
        time.sleep(queue_wait_interval_sec)
        size = get_print_queue_size(printer_name)
        if size is None:
            return  # 待機中に取得不能になったら諦めて抜ける
