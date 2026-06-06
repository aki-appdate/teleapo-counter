"""
架電カウンター
- 縦長・細幅レイアウト（常に最前面）
- Google Sheets にリアルタイム書き込み
"""

import tkinter as tk
from tkinter import messagebox
from google.oauth2.service_account import Credentials
import gspread
import json
import os
import sys
import platform
import threading
from datetime import datetime, date

# ─── OS別フォント設定 ─────────────────────────────────────────────────────────
_OS = platform.system()  # "Darwin"=Mac / "Windows" / "Linux"

def F(size: int, bold: bool = False) -> tuple:
    """OSに合わせた日本語対応フォントを返す"""
    if _OS == "Darwin":
        name = "Hiragino Sans"       # Mac標準の日本語フォント
    elif _OS == "Windows":
        name = "Meiryo"              # Windows標準の日本語フォント
    else:
        name = "TkDefaultFont"
    weight = "bold" if bold else "normal"
    return (name, size, weight)

# ─── パス解決 ─────────────────────────────────────────────────────────────────
def base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE = base_dir()
WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]

def jp_date(d: date) -> str:
    return f"{d.month}/{d.day}({WEEKDAYS_JP[d.weekday()]})"

# ─── 設定 ─────────────────────────────────────────────────────────────────────
def load_config() -> dict:
    with open(os.path.join(BASE, "config.json"), encoding="utf-8") as f:
        return json.load(f)

def save_config(cfg: dict):
    with open(os.path.join(BASE, "config.json"), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ─── Google Sheets ────────────────────────────────────────────────────────────
def build_client():
    creds_path = os.path.join(BASE, "credentials.json")
    creds = Credentials.from_service_account_file(
        creds_path, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds)

def col_to_index(letter: str) -> int:
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - 64)
    return result

def find_row(ws, date_str: str, start_row: int):
    col_a = ws.col_values(1)
    for i, val in enumerate(col_a[start_row - 1:], start=start_row):
        if val.strip() == date_str.strip():
            return i
    return None

# ─── アプリ ───────────────────────────────────────────────────────────────────
class App(tk.Tk):
    W = 200  # ウィンドウ幅

    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.gc = None
        self._lock = threading.Lock()
        self._row_cache = {}
        self._timers = {}  # デバウンス用タイマー

        self.counters = {s: tk.IntVar(value=0)
                         for s in self.cfg.get("column_mapping", {})}

        self.title("架電カウンター")
        self.attributes("-topmost", True)

        self._build()

        # コンテンツに合わせて幅・高さ両方自動調整
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        self.geometry(f"{w}x{h}")
        self.resizable(False, False)

        self._place_window()
        self._connect()

    # ── UI構築 ────────────────────────────────────────────────────────────────
    def _build(self):
        for w in self.winfo_children():
            w.destroy()

        # ── ヘッダー ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg="#1e3a5f")
        hdr.pack(fill="x")

        tk.Label(hdr, text="架電カウンター", bg="#1e3a5f", fg="white",
                 font=F(12, bold=True),
                 anchor="center").pack(fill="x", pady=(8, 4))

        # 名前
        r1 = tk.Frame(hdr, bg="#1e3a5f")
        r1.pack(fill="x", padx=8, pady=2)
        tk.Label(r1, text="名前", bg="#1e3a5f", fg="white",
                 font=F(11), width=4, anchor="w").pack(side="left")
        self.var_member = tk.StringVar()
        members = self.cfg.get("members", [])
        month   = self.cfg.get("month", "")
        # ドロップダウンは名前のみ表示、シート名は内部で「名前+月」に結合
        opt = tk.OptionMenu(r1, self.var_member, *members)
        opt.config(bg="white", fg="black", font=F(11),
                   relief="flat", highlightthickness=0,
                   activebackground="#ddecff", activeforeground="black",
                   width=8)
        opt["menu"].config(bg="white", fg="black", font=F(11))
        if members:
            self.var_member.set(members[0])
        opt.pack(side="left")
        # 現在の月をヘッダーに表示
        tk.Label(r1, text=f"({month})", bg="#1e3a5f", fg="#aaccff",
                 font=F(9)).pack(side="left", padx=(4,0))
        self.var_member.trace_add("write", self._ctx_changed)

        # 日付（テキスト入力＋前後ボタン）
        r2 = tk.Frame(hdr, bg="#1e3a5f")
        r2.pack(fill="x", padx=8, pady=2)
        tk.Label(r2, text="日付", bg="#1e3a5f", fg="white",
                 font=F(11), width=4, anchor="w").pack(side="left")

        btn_prev = tk.Label(r2, text="◀", bg="#2a4a6f", fg="white",
                            font=F(11), cursor="hand2", padx=4)
        btn_prev.pack(side="left")
        btn_prev.bind("<Button-1>", lambda e: self._shift_date(-1))

        self.var_date = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        ent = tk.Entry(r2, textvariable=self.var_date, width=10,
                       font=F(11), bg="white", fg="black",
                       insertbackground="black", relief="flat")
        ent.pack(side="left", padx=2)
        ent.bind("<Return>",   self._ctx_changed)
        ent.bind("<FocusOut>", self._ctx_changed)

        btn_next = tk.Label(r2, text="▶", bg="#2a4a6f", fg="white",
                            font=F(11), cursor="hand2", padx=4)
        btn_next.pack(side="left")
        btn_next.bind("<Button-1>", lambda e: self._shift_date(+1))

        # 接続ステータス
        self.lbl_st = tk.Label(hdr, text="● 未接続", bg="#1e3a5f",
                                fg="#888888", font=F(9))
        self.lbl_st.pack(anchor="w", padx=8, pady=(4, 6))

        # ── 区切り線 ──────────────────────────────────────────────────────────
        tk.Frame(self, bg="#3a5a8f", height=2).pack(fill="x")

        # ── カウンター一覧 ────────────────────────────────────────────────────
        body = tk.Frame(self, bg="#f5f6f8")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)  # カウント＋ボタン群

        statuses = list(self.cfg.get("column_mapping", {}).keys())
        for i, status in enumerate(statuses):
            bg = "#ffffff" if i % 2 == 0 else "#eef0f4"

            # ステータス名
            tk.Label(body, text=status, bg=bg, fg="#2c3e50",
                     font=F(11), anchor="w", padx=4
                     ).grid(row=i, column=0, sticky="nsew", pady=1)

            # 右側フレーム（－ カウント ＋ を完全均等に並べる）
            right = tk.Frame(body, bg=bg)
            right.grid(row=i, column=1, sticky="nsew", padx=(2,6), pady=1)
            right.columnconfigure(0, weight=1, uniform="r")
            right.columnconfigure(1, weight=1, uniform="r")
            right.columnconfigure(2, weight=1, uniform="r")

            btn_m = tk.Label(right, text="－", bg="#e74c3c", fg="white",
                             font=F(12, bold=True), cursor="hand2",
                             relief="raised", bd=2, anchor="center")
            btn_m.grid(row=0, column=0, sticky="nsew", padx=(0,1), ipady=3)
            btn_m.bind("<Button-1>", lambda e, s=status: self._change(s, -1))
            btn_m.bind("<Enter>",    lambda e, b=btn_m: b.config(bg="#c0392b"))
            btn_m.bind("<Leave>",    lambda e, b=btn_m: b.config(bg="#e74c3c"))

            # カウント数（クリックで直接入力に切り替え）
            lbl_count = tk.Label(right, textvariable=self.counters[status],
                                 bg="white", fg="#2c3e50", font=F(12, bold=True),
                                 relief="solid", anchor="center", bd=1, cursor="xterm")
            lbl_count.grid(row=0, column=1, sticky="nsew", padx=1, ipady=3)
            lbl_count.bind("<Button-1>", lambda e, s=status, r=right: self._start_edit(s, r))

            btn_p = tk.Label(right, text="＋", bg="#27ae60", fg="white",
                             font=F(12, bold=True), cursor="hand2",
                             relief="raised", bd=2, anchor="center")
            btn_p.grid(row=0, column=2, sticky="nsew", padx=(1,0), ipady=3)
            btn_p.bind("<Button-1>", lambda e, s=status: self._change(s, +1))
            btn_p.bind("<Enter>",    lambda e, b=btn_p: b.config(bg="#1e8449"))
            btn_p.bind("<Leave>",    lambda e, b=btn_p: b.config(bg="#27ae60"))

        # ── フッター（Labelで代用→非アクティブでも色が消えない） ────────────
        tk.Frame(self, bg="#3a5a8f", height=2).pack(fill="x")
        ftr = tk.Frame(self, bg="#dde2ea")
        ftr.pack(fill="x")

        btn_load = tk.Label(ftr, text="シートから読込", bg="#2980b9", fg="white",
                            font=F(9), relief="raised", bd=2, cursor="hand2", padx=4)
        btn_load.pack(side="left", padx=6, pady=4)
        btn_load.bind("<Button-1>", lambda e: self._load())
        btn_load.bind("<Enter>", lambda e: btn_load.config(bg="#3498db"))
        btn_load.bind("<Leave>", lambda e: btn_load.config(bg="#2980b9"))

        btn_rst = tk.Label(ftr, text="全リセット", bg="#7f8c8d", fg="white",
                           font=F(9), relief="raised", bd=2, cursor="hand2", padx=4)
        btn_rst.pack(side="right", padx=6, pady=4)
        btn_rst.bind("<Button-1>", lambda e: self._reset())
        btn_rst.bind("<Enter>", lambda e: btn_rst.config(bg="#95a5a6"))
        btn_rst.bind("<Leave>", lambda e: btn_rst.config(bg="#7f8c8d"))

        self.update_idletasks()

    # ── ウィンドウ位置（右端・縦中央） ────────────────────────────────────────
    def _place_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        ww = self.winfo_reqwidth()
        wh = self.winfo_reqheight()
        x = sw - ww - 4
        y = max(0, (sh - wh) // 2)
        self.geometry(f"+{x}+{y}")

    # ── Sheets接続 ────────────────────────────────────────────────────────────
    def _connect(self):
        self._set_st("● 接続中…", "#e6ac00")
        threading.Thread(target=self._do_connect, daemon=True).start()

    def _do_connect(self):
        try:
            gc = build_client()
            with self._lock:
                self.gc = gc
            self.after(0, lambda: self._set_st("● 接続済", "#27ae60"))
        except FileNotFoundError:
            self.after(0, lambda: self._set_st("● credentials.json なし", "#e74c3c"))
        except Exception as e:
            self.after(0, lambda: self._set_st(f"● 接続エラー", "#e74c3c"))

    def _set_st(self, text, color="#888888"):
        if self.lbl_st.winfo_exists():
            self.lbl_st.config(text=text, fg=color)

    # ── 日付を前後にずらす ────────────────────────────────────────────────────
    def _shift_date(self, delta: int):
        from datetime import timedelta
        try:
            d = datetime.strptime(self.var_date.get().strip(), "%Y-%m-%d").date()
            self.var_date.set((d + timedelta(days=delta)).strftime("%Y-%m-%d"))
            self._ctx_changed()
        except ValueError:
            pass

    # ── コンテキスト変更 ──────────────────────────────────────────────────────
    def _ctx_changed(self, *_):
        self._row_cache.clear()
        for v in self.counters.values():
            v.set(0)

    # ── カウント直接入力 ──────────────────────────────────────────────────────
    def _start_edit(self, status, right_frame):
        """カウント数ラベルをEntryに差し替えて直接入力モードにする"""
        current = self.counters[status].get()

        # column=1のラベルを非表示にしてEntryを置く
        ent = tk.Entry(right_frame, font=F(12, bold=True),
                       bg="#fffde7", fg="#2c3e50",
                       justify="center", relief="solid", bd=1, width=3)
        ent.grid(row=0, column=1, sticky="nsew", padx=1, ipady=3)
        ent.insert(0, str(current))
        ent.select_range(0, tk.END)
        ent.focus_set()

        def confirm(e=None):
            try:
                val = max(0, int(ent.get()))
            except ValueError:
                val = current
            self.counters[status].set(val)
            ent.destroy()
            # スプシに書き込み
            if status in self._timers:
                self.after_cancel(self._timers[status])
            self._timers[status] = self.after(
                100, lambda s=status, v=val: threading.Thread(
                    target=self._do_write, args=(s, v), daemon=True).start()
            )

        def cancel(e=None):
            ent.destroy()

        ent.bind("<Return>",  confirm)
        ent.bind("<Escape>",  cancel)
        ent.bind("<FocusOut>", confirm)

    # ── カウント変更（デバウンス：0.8秒後にまとめて書き込み） ────────────────
    def _change(self, status, delta):
        new = max(0, self.counters[status].get() + delta)
        self.counters[status].set(new)

        # 既存の書き込み予約をキャンセルして再予約
        if status in self._timers:
            self.after_cancel(self._timers[status])
        self._timers[status] = self.after(
            800, lambda s=status, v=new: threading.Thread(
                target=self._do_write, args=(s, v), daemon=True).start()
        )

    def _do_write(self, status, value):
        self.after(0, lambda: self._set_st("● 書込中…", "#e6ac00"))
        try:
            gc, sid, sheet, row = self._resolve()
            col = self.cfg["column_mapping"][status]
            gc.open_by_key(sid).worksheet(sheet).update_cell(
                row, col_to_index(col), value)
            self.after(0, lambda: self._set_st("● 接続済", "#27ae60"))
        except _RowNotFound as e:
            msg = str(e)
            self.after(0, lambda: self._set_st(f"● {msg}", "#e67e22"))
        except Exception as e:
            msg = str(e)
            print(f"[書込エラー] {msg}")
            self.after(0, lambda: self._set_st(f"● エラー: {msg[:20]}", "#e74c3c"))

    # ── シートから読込 ────────────────────────────────────────────────────────
    def _load(self):
        self._set_st("● 読込中…", "#e6ac00")
        threading.Thread(target=self._do_load, daemon=True).start()

    def _do_load(self):
        try:
            gc, sid, sheet, row = self._resolve()
            ws = gc.open_by_key(sid).worksheet(sheet)
            for status, col in self.cfg["column_mapping"].items():
                val = ws.cell(row, col_to_index(col)).value
                num = int(val) if val and str(val).isdigit() else 0
                self.after(0, lambda s=status, v=num: self.counters[s].set(v))
            self.after(0, lambda: self._set_st("● 接続済", "#27ae60"))
        except _RowNotFound as e:
            msg = str(e)
            self.after(0, lambda: self._set_st(f"● {msg}", "#e67e22"))
        except Exception:
            self.after(0, lambda: self._set_st("● エラー", "#e74c3c"))

    # ── コンテキスト解決 ──────────────────────────────────────────────────────
    def _resolve(self):
        with self._lock:
            gc = self.gc
        if not gc:
            raise RuntimeError("未接続")
        name = self.var_member.get()
        if not name:
            raise RuntimeError("名前未選択")
        # 「名前」+「月」を結合してシート名にする（例：秋山 + 6月度 = 秋山6月度）
        month = self.cfg.get("month", "")
        sheet = f"{name}{month}"
        try:
            d = datetime.strptime(self.var_date.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            raise RuntimeError("日付形式エラー")
        date_str = jp_date(d)
        sid = self.cfg["spreadsheet_id"]
        start = self.cfg.get("data_start_row", 3)
        key = (sheet, date_str)
        if key not in self._row_cache:
            ws = gc.open_by_key(sid).worksheet(sheet)
            row = find_row(ws, date_str, start)
            if row is None:
                raise _RowNotFound(f"{date_str} の行なし")
            self._row_cache[key] = row
        return gc, sid, sheet, self._row_cache[key]

    # ── リセット ──────────────────────────────────────────────────────────────
    def _reset(self):
        if messagebox.askyesno("確認", "全カウントを 0 にリセットしますか？"):
            self._row_cache.clear()
            for v in self.counters.values():
                v.set(0)


class _RowNotFound(Exception):
    pass


if __name__ == "__main__":
    app = App()
    app.mainloop()
