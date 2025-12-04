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


# ===== 日付入力パース（YYYY/MM/DD） =====
def parse_input_date(s: str) -> datetime:
    # バッチのチェック相当（形式が違えば例外）
    try:
        return datetime.strptime(s.strip(), "%Y/%m/%d")
    except ValueError:
        raise ValueError("日付形式が違います。YYYY/MM/DD で入力してください。")


# ===== ファイル収集 =====
def collect_targets(parent_folder: Path, target_date: datetime) -> List[Tuple[str, Path, str]]:
    """
    バッチ仕様：
      - parent 配下の各サブフォルダを走査
      - target_dateに更新されたPDFを印刷対象
      - そのサブフォルダでPDFが1つでも対象になったら、同フォルダのdocmも全部対象
    戻り値: [("pdf", pdf_path), ("word", docm_path), ...]
    """
    targets: List[Tuple[str, Path, str]] = []

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
            for w in docms:
                targets.append(("word", w, w.name))
            
    return targets


# ===== 印刷：PDFtoPrinter =====
def print_pdf_with_pdftoprinter(pdftoprinter_path: Path, printer_name: str, pdf_path: Path):
    if not pdftoprinter_path.exists():
        raise FileNotFoundError(f"PDFtoPrinter.exe が見つかりません: {pdftoprinter_path}")
    # PDFtoPrinter.exe "file.pdf" "Printer Name"
    subprocess.run([str(pdftoprinter_path), str(pdf_path), printer_name], check=True)


# ===== 印刷：LibreOffice headless =====
def print_word_with_soffice(soffice_path: Path, printer_name: str, word_path: Path):
    if not soffice_path.exists():
        raise FileNotFoundError(f"soffice.com が見つかりません: {soffice_path}")
    # soffice --headless --pt "Printer Name" "file.docm"
    subprocess.run(
        [str(soffice_path), "--headless", "--pt", printer_name, str(word_path)],
        check=True
    )


# ===== キュー制御（骨組み） =====
def get_print_queue_size(printer_name: str) -> Optional[int]:
    """
    ここは後で実装：
      - pywin32 を入れて win32print.EnumJobs を使う想定
    ひとまず None を返して、キュー制御なしで動く骨組みにしておく
    """
    return None


def wait_if_queue_full(printer_name: str, queue_limit: int, sleep_sec: float):
    """
    キュー上限付き高速投入の骨組み。
    get_print_queue_size が実装されたら効く。
    """
    size = get_print_queue_size(printer_name)
    if size is None:
        return  # 未実装なら待たない
    while size >= queue_limit:
        time.sleep(sleep_sec)
        size = get_print_queue_size(printer_name)

