# print_progress_gui.py
# Windows 11 用：印刷進捗GUI + 印刷キュー監視（確定版）
#
# 仕様:
# - GUI表示: 「印刷中」, 進捗（プログレスバー＋件数）, 中止ボタン
# - 完了条件:
#     1) 印刷対象リストが空（= 印刷投入が終わり "sent_all" を受信）
#     2) OS印刷キューが空（空判定が連続N回続いたら確定）
#   => メッセージボックス「印刷完了しました」 + 「終了」ボタン有効化
# - 中止ボタン押下時:
#   cancel_event を立て、次の投入前で停止。
#   完了時の表示は「印刷を中止しました」にする。
#
# 依存:
# - pywin32 があれば win32print で確認（高速で確実）
# - 無ければ PowerShell(Get-PrintJob)でフォールバック
#
# 使い方:
#   from print_progress_gui import run_print_with_gui
#   ok = run_print_with_gui(selected, _print_pdf, _print_word, printer_name)

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import queue
import time
import subprocess

try:
    import win32print
except ImportError:
    win32print = None


def is_printer_queue_empty(printer_name: str) -> bool:
    """
    指定プリンタのスプーラジョブが空なら True。

    優先順:
      1) pywin32 (win32print.EnumJobs)
      2) PowerShell Get-PrintJob
    失敗/判定不能なら False（完了誤判定を避ける）
    """
    if not printer_name:
        return False

    # 1) pywin32 があればそれで判定
    if win32print is not None:
        h = None
        try:
            h = win32print.OpenPrinter(printer_name)
            jobs = win32print.EnumJobs(h, 0, -1, 1)
            return len(jobs) == 0
        except Exception:
            # pywin32 側が失敗したら PS へフォールバック
            pass
        finally:
            if h:
                win32print.ClosePrinter(h)

    # 2) PowerShell フォールバック
    try:
        safe_name = printer_name.replace("'", "''")  # PS シングルクォートエスケープ
        cmd = [
            "powershell",
            "-NoProfile",
            "-Command",
            f"(Get-PrintJob -PrinterName '{safe_name}' | Measure-Object).Count"
        ]
        out = subprocess.check_output(
            cmd,
            stderr=subprocess.STDOUT,
            text=True
            #,creationflags=subprocess.CREATE_NO_WINDOW
        ).strip()
        if out == "":
            return True
        return int(out) == 0
    except Exception:
        return False


class PrintProgressWindow(tk.Tk):
    """
    印刷進捗GUI（Windowsスプーラ監視つき / プリンタ名は外部指定）

    印刷側から event_queue にイベントを put する。
      ("init", total)
      ("start_item", idx, name)
      ("done_item", idx, name)
      ("error_item", idx, name, msg)
      ("log", text)        # 任意
      ("sent_all", )       # 印刷対象リストを全てスプーラに送信し終わった合図
    """

    POLL_MS = 150
    CHECK_SPOOL_MS = 700
    EMPTY_STREAK_REQUIRED = 3   # 空判定が連続N回続いたら完了確定

    def __init__(self, event_queue: queue.Queue, cancel_event: threading.Event,
                 printer_name: str):
        super().__init__()
        self.title("印刷進捗")
        self.geometry("560x260")
        self.resizable(False, False)

        self.q = event_queue
        self.cancel_event = cancel_event
        self.printer_name = printer_name

        self.total = 0
        self.done = 0
        self.error = 0
        self.sent_all = False
        self.empty_streak = 0

        self._build_ui()

        self.after(self.POLL_MS, self._poll_queue)
        self.after(self.CHECK_SPOOL_MS, self._check_completion_condition)

        # 完了/中止確定までは×で閉じさせない
        self.protocol("WM_DELETE_WINDOW", self._block_close)

    def _build_ui(self):
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        # 上：印刷中表示
        self.lbl_title = ttk.Label(main, text="印刷中", font=("", 14, "bold"))
        self.lbl_title.pack(anchor="w")

        # 進捗バー
        self.progress = ttk.Progressbar(
            main, orient="horizontal", mode="determinate", length=520
        )
        self.progress.pack(pady=(8, 0))

        self.lbl_counts = ttk.Label(main, text="0 / 0")
        self.lbl_counts.pack(anchor="w", pady=(6, 0))

        # 現在の対象
        self.lbl_current = ttk.Label(main, text="現在の印刷対象: (なし)")
        self.lbl_current.pack(anchor="w", pady=(8, 0))

        # キュー/プリンタ状態
        self.lbl_spool = ttk.Label(
            main,
            text=f"プリンタ: {self.printer_name} / キュー状態: (確認中)"
        )
        self.lbl_spool.pack(anchor="w", pady=(4, 0))

        # ボタン行
        btn_row = ttk.Frame(main)
        btn_row.pack(fill="x", pady=(14, 0))

        self.btn_cancel = ttk.Button(btn_row, text="中止", command=self._on_cancel)
        self.btn_cancel.pack(side="left")

        self.btn_exit = ttk.Button(btn_row, text="終了", command=self._on_exit)
        self.btn_exit.state(["disabled"])
        self.btn_exit.pack(side="right")

    def _update_counts(self):
        self.lbl_counts.configure(
            text=f"{self.done + self.error} / {self.total} (失敗:{self.error})"
        )
        self.progress["maximum"] = max(self.total, 1)
        self.progress["value"] = self.done + self.error

    def _poll_queue(self):
        try:
            while True:
                ev = self.q.get_nowait()
                self._handle_event(ev)
        except queue.Empty:
            pass

        self.after(self.POLL_MS, self._poll_queue)

    def _handle_event(self, ev):
        if not ev:
            return
        etype = ev[0]

        if etype == "init":
            self.total = int(ev[1])
            self.done = 0
            self.error = 0
            self.sent_all = False
            self.empty_streak = 0
            self._update_counts()

        elif etype == "start_item":
            _, idx, name = ev
            self.lbl_current.configure(text=f"現在の印刷対象: {name}")

        elif etype == "done_item":
            _, idx, name = ev
            self.done += 1
            self._update_counts()

        elif etype == "error_item":
            _, idx, name, msg = ev
            self.error += 1
            self._update_counts()

        elif etype == "log":
            # いまは画面に出してないが、必要なら拡張可
            pass

        elif etype == "sent_all":
            # 「印刷対象リストが空（=送信完了）」の合図
            self.sent_all = True
            self.empty_streak = 0

    def _check_completion_condition(self):
        """
        完了条件:
          1) 印刷対象リストが空（= sent_all 済み）
          2) OS印刷キューが空（連続N回）
        """
        if self.sent_all:
            spool_empty = is_printer_queue_empty(self.printer_name)
            status = "空" if spool_empty else "残りあり"
            self.lbl_spool.configure(
                text=f"プリンタ: {self.printer_name} / キュー状態: {status}"
            )

            if spool_empty:
                self.empty_streak += 1
            else:
                self.empty_streak = 0

            if self.empty_streak >= self.EMPTY_STREAK_REQUIRED:
                self._on_all_done()
                return

        self.after(self.CHECK_SPOOL_MS, self._check_completion_condition)

    def _on_all_done(self):
        # 完了/中止 表示を切り替え
        self.lbl_current.configure(text="現在の印刷対象: (なし)")
        self.btn_cancel.state(["disabled"])
        self.btn_exit.state(["!disabled"])

        if self.cancel_event.is_set():
            self.lbl_title.configure(text="印刷を中止しました")
            messagebox.showinfo("中止", "印刷を中止しました")
        else:
            self.lbl_title.configure(text="印刷完了")
            messagebox.showinfo("完了", "印刷完了しました")

    def _on_cancel(self):
        # 印刷スレッドへ中止シグナル
        self.cancel_event.set()
        self.btn_cancel.state(["disabled"])
        self.lbl_title.configure(text="中止処理中…")

    def _on_exit(self):
        self.destroy()

    def _block_close(self):
        # 完了/中止確定までは×で閉じない
        if "disabled" in self.btn_exit.state():
            messagebox.showwarning("印刷中", "完了または中止確定まで閉じられません。")
            return
        self._on_exit()


def run_print_with_gui(selected, print_pdf_func, print_word_func, printer_name: str):
    """
    既存印刷処理をGUI付きで走らせるためのラッパ。

    selected: [(kind, path, fname), ...]
    print_pdf_func(path): 既存PDF印刷関数
    print_word_func(path): 既存Word印刷関数
    printer_name: config.json から渡す監視対象プリンタ名

    返り値:
      True  = 正常完了（失敗なし＆中止なし）
      False = 中止 or 失敗あり
    """
    q = queue.Queue()
    cancel_event = threading.Event()

    gui = PrintProgressWindow(q, cancel_event, printer_name=printer_name)

    def worker():
        q.put(("init", len(selected)))

        for i, (kind, path, fname) in enumerate(selected):
            if cancel_event.is_set():
                break

            q.put(("start_item", i, fname))
            try:
                if kind == "pdf":
                    print_pdf_func(path)
                else:
                    print_word_func(path)

                q.put(("done_item", i, fname))
            except Exception as e:
                q.put(("error_item", i, fname, str(e)))

        # 印刷対象リストが空になった合図（送信完了）
        q.put(("sent_all",))

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    gui.mainloop()

    if cancel_event.is_set():
        return False
    if gui.error > 0:
        return False
    return True
