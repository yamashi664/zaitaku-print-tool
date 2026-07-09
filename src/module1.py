# module1.py
import sys
import json
import time
import subprocess
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime
from typing import List, Tuple, Optional
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

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


# ======zipファイル対応版追加要素==========
def process_zip_and_generate_fax(zip_path: Path, config: dict, selected_pharmacy: dict) -> Tuple[Path, List[Tuple[str, Path, str]]]:
    """
    V2用ロジック (アンダースコア区切りフォルダ対応版):
      1) zipファイルを一時フォルダに解凍
      2) 内部の「個人名_施設名_役職」という形式のフォルダを検知し、宛先情報を抽出
      3) fax_template.docx の各タグを、抽出した情報および config 情報で置換
      4) 印刷対象リストを返却
    """
    temp_dir = Path(tempfile.mkdtemp(prefix="zaitaku_print_"))
    
    # zipファイルを解凍
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for member in zip_ref.infolist():
            try:
                filename = member.filename.encode('cp437').decode('cp932')
            except Exception:
                filename = member.filename
            
            target_path = temp_dir / filename
            if member.is_dir():
                target_path.mkdir(parents=True, exist_ok=True)
            else:
                target_path.parent.mkdir(parents=True, exist_ok=True)
                with zip_ref.open(member) as source, open(target_path, "wb") as target:
                    target.write(source.read())

    # 印刷対象を詰め込む最終的なリスト
    print_list = []

    # 1. 一時フォルダの直下にある「宛先フォルダ」を1つずつループで巡回する
    #    (例: temp_dir / "ケアマネA_ほにゃらら介護相談室_ケアマネジャー" など)
    for target_folder in temp_dir.iterdir():
        if not target_folder.is_dir():
            continue  # ファイル（もしあれば）はスキップして、フォルダだけを処理

        # 2. そのフォルダ名から「個人名」「施設名」を抽出する
        folder_name = target_folder.name
        parts = folder_name.split("_")
        
        if len(parts) >= 2:
            personal_name = parts[0]
            facility_name = parts[1]
        else:
            personal_name = "関係者"
            facility_name = folder_name

        # 3. そのフォルダ「の内部だけ」からPDFファイルを収集する
        pdf_files = list(target_folder.glob("*.pdf"))  # **/* ではなく *.pdf でそのフォルダ直下のみ
        report_count = len(pdf_files)

        if report_count == 0:
            continue  # もしPDFが1枚もないフォルダなら送付状を作る必要がないのでスキップ

        # 4. このフォルダ（宛先）専用の送付状をWordテンプレートから自動生成する
        # プログラム/exeと同じ場所にあるテンプレートを見に行く
        template_path = base_dir() / "fax_template.docx"
        generated_word_path = None

        if template_path.exists():
            doc = Document(str(template_path))
            today_str = datetime.now().strftime("%Y年%m月%d日")
            
            # 選択された薬局情報
            ph_info = selected_pharmacy if selected_pharmacy else {}
            
            replacements = {
                "{{post-code}}": ph_info.get("post_code", "〒000-0000"),
                "{{address}}": ph_info.get("address", "薬局の住所が未設定です"),
                "{{pharmacy_name}}": ph_info.get("pharmacy_name", "〇〇薬局"), # キー名をpharmacy_nameに統一
                "{{tel_num}}": ph_info.get("tel_num", "000-000-0000"),
                "{{fax_num}}": ph_info.get("fax_num", "000-000-0000"),
                "{{facility_name}}": facility_name,
                "{{personal_name}}": personal_name,
                "{{date}}": today_str,
                "{{report_count}}": str(report_count)
            }
            
            def safe_replace(paragraph, replacements):
                for key, val in replacements.items():
                    if key in paragraph.text:
                        # 1. まず段落全体で文字列を置換（これでRunの分断を無視して中身を入れ替える）
                        paragraph.text = paragraph.text.replace(key, val)
            
                        # 2. ただしこれだと段落全体の書式がリセットされる場合があるため、
                        #    ここで「薬局名」だけ特別にフォントサイズを戻す処理を行う
                        if key == "{{pharmacy_name}}":
                            for run in paragraph.runs:
                                if val in run.text:
                                    run.font.size = Pt(14) # ここで強制的に指定サイズにする

            # 本文の置換処理
            for p in doc.paragraphs:
                safe_replace(p, replacements)

            # テーブル内の置換処理
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            safe_replace(p, replacements)
            
            # ファイル名に宛名を組み込んでこのフォルダ内に保存
            fax_filename = f"【送付状】{personal_name} 様_{facility_name}.docx"
            generated_word_path = target_folder / fax_filename
            doc.save(str(generated_word_path))

        # 5. このフォルダの「印刷セット」を順序通りに全体のリストに追加する
        # (まず、そのフォルダ内のPDFを名前順（01, 02...）に配置)
        for pdf_path in pdf_files:
            print_list.append(("pdf", pdf_path, pdf_path.name))
            
        # (最後に、出来上がった送付状を一番後ろに配置)
        if generated_word_path:
            print_list.append(("word", generated_word_path, generated_word_path.name))

    # すべてのフォルダの処理が終わったら、一時フォルダのパスと、完成した全印刷リストを返す
    return temp_dir, print_list