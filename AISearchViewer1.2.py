import tkinter as tk
import ctypes
import sys
import os  # ← ここで os を確実にインポートする

# 1. まずアプリIDを設定（タスクバー分離防止）
def set_appusermodel_id():
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Ame.AISearchViewer")
    except:
        pass

set_appusermodel_id()

# 2. リソースパス取得関数を定義
def resource_path(relative_path):
    """ PyInstaller環境と通常実行環境の両方に対応するパス取得 """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 3. その他のインポート
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
import urllib.parse
import webbrowser
import re
# import os  ← 下の方にあったら削除またはコメントアウト
import string
import configparser
from pathlib import Path
import logging
from typing import Optional, List
# =====================
# ログ設定
# =====================
logging.basicConfig(
    filename='ai_search_viewer.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s'
)

# =====================
# Utility Functions
# =====================
def safe_text(v):
    return "" if pd.isna(v) else str(v)

def ai_url(text):
    text = safe_text(text)
    if text:
        return f'=HYPERLINK("https://www.perplexity.ai/search?q={urllib.parse.quote(text)}","AI検索")'
    return ""

def google_url(text):
    text = safe_text(text)
    if text:
        return f'=HYPERLINK("https://www.google.com/search?q={urllib.parse.quote(text)}","Google検索")'
    return ""

def extract_url(v):
    if isinstance(v, str):
        m = re.search(r'HYPERLINK\("(.+?)"', v)
        if m:
            return m.group(1)
    return None

def display_text(v):
    if pd.isna(v):
        return ""
    if isinstance(v, str) and v.startswith("=HYPERLINK"):
        m = re.search(r',"(.+?)"\)', v)
        if m:
            return m.group(1)
    return v

def get_excel_header(col_index):
    result = ""
    while col_index > 0:
        col_index, rem = divmod(col_index - 1, 26)
        result = string.ascii_uppercase[rem] + result
    return result


def normalize_row_values(row, ncols: int):
    """Convert a row-like to list[str] length ncols."""
    vals = []
    for i in range(ncols):
        try:
            v = row[i]
        except Exception:
            v = ""
        vals.append("" if pd.isna(v) else str(v))
    if len(vals) < ncols:
        vals += [""] * (ncols - len(vals))
    return vals[:ncols]


# =====================
# ToolTip Class
# =====================
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        self.id = None
        self.widget.bind("<Enter>", self.schedule)
        self.widget.bind("<Leave>", self.hide)

    def schedule(self, event=None):
        self.id = self.widget.after(500, self.show)

    def show(self):
        if self.tip:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width()
        y = self.widget.winfo_rooty() + self.widget.winfo_height()
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw, text=self.text, background="#ffffe0",
            relief="solid", borderwidth=1, font=("Segoe UI", 9), justify="left"
        ).pack(ipadx=4, ipady=2)

    def hide(self, event=None):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None
        if self.tip:
            self.tip.destroy()
            self.tip = None

# =====================
# Main Application Class
# =====================
class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI検索ビューア")

        self.root.geometry("1400x800")
        self.root.minsize(1100, 650)

        self.current_df: Optional[pd.DataFrame] = None
        self.raw_df: Optional[pd.DataFrame] = None
        self.header_row_current: int = int(getattr(self, "header_row_default", 1) or 1)
        self.excel_path: Optional[str] = None
        self.base_col_name: Optional[str] = None
        self.base_col_names: List[str] = []  # 複数検索語句列
        self.base_joiner: str = " "  # 複数列を結合する区切り
        self.sort_state = {}
        self.sorted_col: Optional[str] = None

        # セル編集
        self.edit_entry = None
        self.edit_row = None
        self.edit_col = None
        self._edit_is_pre = False
        self._edit_data_index = -1
        self._edit_raw_col = None

        # Undo / Redo / 状態
        self.undo_stack: List[pd.DataFrame] = []
        self.redo_stack: List[pd.DataFrame] = []
        self.unsaved_changes = False

        # ヘッダークリックの遅延ソート制御（ダブルクリックでキャンセルする）
        self._header_click_job = None
        self._header_click_col = None

        # フロー改善
        self._onboard_shown = False
        self.op_history: List[str] = []  # 操作履歴（Undo単位）

        # config
        self.config_path = os.path.join(os.path.expanduser("~"), ".ai_search_viewer.ini")
        self.config = configparser.ConfigParser()
        self.confirm_rebuild = True  # 既定：確認あり
        self.load_config()

        # --- 表示色の既定値 ---
        if "Colors" not in self.config:
            self.config["Colors"] = {}
        self.config["Colors"].setdefault("headerrow", "#ffb6c1")   # 見出し(強調)行
        self.config["Colors"].setdefault("preheader", "#f5f5f5")   # 見出しより上(グレー)
        self.config["Colors"].setdefault("even", "#ffffff")        # 偶数行
        self.config["Colors"].setdefault("odd", "#f9f9f9")         # 奇数行
        self.colors = dict(self.config["Colors"])

        
        
        self.root.geometry("1200x650")

        self.setup_theme_style()
        self.setup_treeview_style()
        self.setup_menu()
        self.setup_ui()
        self.root.after(100, self.load_once)

    # ---------------------
    # Style
    # ---------------------
    def setup_theme_style(self):
        """ttkのテーマ/ボタン等の基本スタイルを設定します。"""
        try:
            style = ttk.Style(self.root)
            # テーマによってボタン背景色が効かないため、比較的反映されやすい clam を優先
            if "clam" in style.theme_names():
                style.theme_use("clam")

            # 危険操作（上書き保存など）用の赤ボタン
            style.configure("Red.TButton", foreground="white")
            style.map(
                "Red.TButton",
                foreground=[("disabled", "gray"), ("!disabled", "white")],
                background=[("active", "#cc4444"), ("!disabled", "#cc3333")],
            )
        except Exception:
            # 失敗しても致命的ではない（既定テーマで続行）
            return

    # ---------------------
    # Config
    # ---------------------
    def load_config(self):
        # 設定はユーザーのホームに保存（exe化しても消えにくい）
        self.header_row_default = 1
        self.base_col_index_default = 1
        self.preview_rows_default = 20
        self.base_col_indices_default: List[int] = []
        # 起動時の動作
        self.startup_open_last = False
        self.startup_show_file_dialog = True
        self.startup_always_show_load_settings = True
        self.last_file = ""
        # 検索リンク生成
        self.generate_ai = True
        self.generate_google = True
        # AI検索サービス（URLテンプレは {q} を検索語句に置換）
        self.ai_service = "Perplexity"
        self.ai_url_template = "https://www.perplexity.ai/search?q={q}"
        self.link_insert_mode = "fixed2"  # fixed2 / after_base / rightmost
        # Undo 最大数
        self.undo_limit = 20
        if os.path.exists(self.config_path):
            try:
                self.config.read(self.config_path, encoding="utf-8")
                # 互換：過去版の設定（列名指定）
                self.base_col_name = self.config.get("Settings", "base_column", fallback=None) or None
                # v1.2.3：見出し行 / 検索語句列（列番号）の記憶
                self.header_row_default = self.config.getint("Settings", "header_row", fallback=1)
                self.base_col_index_default = self.config.getint("Settings", "base_column_index", fallback=1)
                self.confirm_rebuild = self.config.getboolean("Settings", "confirm_rebuild", fallback=True)
                self.preview_rows_default = self.config.getint("Settings", "preview_rows", fallback=20)

                # v1.2.5+: 複数検索語句列
                base_cols = self.config.get("Settings", "base_columns", fallback="").strip()
                if base_cols:
                    self.base_col_names = [c.strip() for c in base_cols.split(",") if c.strip()]
                else:
                    self.base_col_names = []
                self.base_joiner = self.config.get("Settings", "base_joiner", fallback=" ").replace("\\t", "\t")
                # 互換：列番号のリスト（あれば次回の初期選択に利用）
                idxs = self.config.get("Settings", "base_column_indices", fallback="").strip()
                if idxs:
                    try:
                        self.base_col_indices_default = [int(x) for x in idxs.split(",") if x.strip().isdigit()]
                    except Exception:
                        self.base_col_indices_default = []
                else:
                    self.base_col_indices_default = []

                # 起動時
                self.startup_open_last = self.config.getboolean("Settings", "startup_open_last", fallback=False)
                self.startup_show_file_dialog = self.config.getboolean("Settings", "startup_show_file_dialog", fallback=True)
                self.startup_always_show_load_settings = self.config.getboolean("Settings", "startup_always_show_load_settings", fallback=True)
                self.last_file = self.config.get("Settings", "last_file", fallback="")
                # 検索リンク生成
                self.generate_ai = self.config.getboolean("Settings", "generate_ai", fallback=True)
                self.generate_google = self.config.getboolean("Settings", "generate_google", fallback=True)
                self.link_insert_mode = self.config.get("Settings", "link_insert_mode", fallback="fixed2")
                # AI検索サービス
                self.ai_service = self.config.get("Settings", "ai_service", fallback="Perplexity")
                self.ai_url_template = self.config.get("Settings", "ai_url_template", fallback="https://www.perplexity.ai/search?q={q}")
                # Undo
                self.undo_limit = self.config.getint("Settings", "undo_limit", fallback=20)
            except Exception as e:
                logging.error(f"Config error: {e}")
    def save_config(self):
        """設定ファイル保存（壊れない版）
        - 既存の self.config の内容をそのまま保存
        - 主要なUI設定は、存在する属性だけ Settings に反映
        """
        try:
            if "Settings" not in self.config:
                self.config["Settings"] = {}

            s = self.config["Settings"]

            # 主要設定（属性がある場合のみ）
            if hasattr(self, "header_row_default"):
                s["header_row"] = str(getattr(self, "header_row_default", 1) or 1)
            if hasattr(self, "confirm_rebuild"):
                s["confirm_rebuild"] = "1" if bool(getattr(self, "confirm_rebuild", False)) else "0"
            if hasattr(self, "preview_rows_default"):
                s["preview_rows"] = str(getattr(self, "preview_rows_default", 20) or 20)

            # 起動時
            if hasattr(self, "startup_open_last"):
                s["startup_open_last"] = "1" if bool(getattr(self, "startup_open_last", False)) else "0"
            if hasattr(self, "startup_show_file_dialog"):
                s["startup_show_file_dialog"] = "1" if bool(getattr(self, "startup_show_file_dialog", True)) else "0"
            if hasattr(self, "startup_always_show_load_settings"):
                s["startup_always_show_load_settings"] = "1" if bool(getattr(self, "startup_always_show_load_settings", True)) else "0"
            if hasattr(self, "last_file"):
                s["last_file"] = str(getattr(self, "last_file", "") or "")

            # 検索語句列（複数）
            if hasattr(self, "base_col_names"):
                s["base_columns"] = ",".join(getattr(self, "base_col_names", []) or [])

            # 生成設定
            if hasattr(self, "gen_mode"):
                s["gen_mode"] = str(getattr(self, "gen_mode", "both") or "both")
            if hasattr(self, "ai_service"):
                s["ai_service"] = str(getattr(self, "ai_service", "Perplexity") or "Perplexity")
            if hasattr(self, "insert_position"):
                s["insert_position"] = str(getattr(self, "insert_position", "right") or "right")

            # 色設定は show_color_settings で self.config["Colors"] を更新済み

            with open(self.config_path, "w", encoding="utf-8") as f:
                self.config.write(f)

        except Exception as e:
            try:
                logging.warning(f"Config save failed: {e}")
            except Exception:
                pass

    def setup_menu(self):
        menubar = tk.Menu(self.root)

        view = tk.Menu(menubar, tearoff=0)
        view.add_command(label="使い方", command=self.show_help_window)
        view.add_command(label="操作履歴", command=self.show_history_window)
        menubar.add_cascade(label="表示", menu=view)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="環境設定…", command=self.open_settings_dialog)
        menubar.add_cascade(label="設定", menu=settings_menu)

        self.root.config(menu=menubar)

    def open_settings_dialog(self):
        """環境設定ダイアログ（OKで保存、キャンセルで破棄）"""
        dlg = tk.Toplevel(self.root)
        dlg.title("環境設定")
        dlg.transient(self.root)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        # variables（作業用コピー）
        var_header = tk.IntVar(value=int(getattr(self, "header_row_default", 1)))
        var_basecol = tk.IntVar(value=int(getattr(self, "base_col_index_default", 1)))
        var_preview = tk.IntVar(value=int(getattr(self, "preview_rows_default", 20)))
        var_confirm = tk.BooleanVar(value=bool(getattr(self, "confirm_rebuild", True)))

        # 起動時の動作
        var_startup_open_last = tk.BooleanVar(value=bool(getattr(self, 'startup_open_last', False)))
        var_startup_file_dialog = tk.BooleanVar(value=bool(getattr(self, 'startup_show_file_dialog', True)))
        var_startup_always_settings = tk.BooleanVar(value=bool(getattr(self, 'startup_always_show_load_settings', True)))

        # 検索リンク生成
        var_gen_ai = tk.BooleanVar(value=bool(getattr(self, 'generate_ai', True)))
        var_gen_google = tk.BooleanVar(value=bool(getattr(self, 'generate_google', True)))
        var_insert_mode = tk.StringVar(value=str(getattr(self, 'link_insert_mode', 'fixed2')))

        # Undo
        var_undo_limit = tk.IntVar(value=int(getattr(self, 'undo_limit', 20) or 20))

        ttk.Label(frm, text="見出し行（初期値）").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Spinbox(frm, from_=1, to=1000, width=8, textvariable=var_header).grid(row=0, column=1, sticky="w", pady=(0, 6), padx=(8, 0))

        ttk.Label(frm, text="検索語句列（初期値）").grid(row=1, column=0, sticky="w", pady=(0, 6))
        ttk.Spinbox(frm, from_=1, to=1000, width=8, textvariable=var_basecol).grid(row=1, column=1, sticky="w", pady=(0, 6), padx=(8, 0))

        ttk.Label(frm, text="プレビュー行数（初期値）").grid(row=2, column=0, sticky="w", pady=(0, 10))
        ttk.Spinbox(frm, from_=5, to=200, width=8, textvariable=var_preview).grid(row=2, column=1, sticky="w", pady=(0, 10), padx=(8, 0))

        ttk.Checkbutton(frm, text="検索語句更新の前に確認する", variable=var_confirm).grid(row=3, column=0, columnspan=2, sticky="w")

        ttk.Separator(frm).grid(row=4, column=0, columnspan=2, sticky='ew', pady=(10, 8))
        ttk.Label(frm, text='起動時の動作').grid(row=5, column=0, columnspan=2, sticky='w')
        ttk.Checkbutton(frm, text='最後に開いたファイルを自動で開く', variable=var_startup_open_last).grid(row=6, column=0, columnspan=2, sticky='w')
        ttk.Checkbutton(frm, text='起動時にファイル選択ダイアログを出す', variable=var_startup_file_dialog).grid(row=7, column=0, columnspan=2, sticky='w')
        ttk.Checkbutton(frm, text='起動時に読み込み設定ダイアログを必ず出す', variable=var_startup_always_settings).grid(row=8, column=0, columnspan=2, sticky='w')

        ttk.Separator(frm).grid(row=9, column=0, columnspan=2, sticky='ew', pady=(10, 8))
        ttk.Label(frm, text='検索リンク生成').grid(row=10, column=0, columnspan=2, sticky='w')
        ttk.Checkbutton(frm, text='AI検索 列を生成', variable=var_gen_ai).grid(row=11, column=0, columnspan=2, sticky='w')
        ttk.Checkbutton(frm, text='Google検索 列を生成', variable=var_gen_google).grid(row=12, column=0, columnspan=2, sticky='w')
        ttk.Label(frm, text='挿入位置').grid(row=13, column=0, sticky='w', pady=(6, 0))
        cmb_insert = ttk.Combobox(frm, state='readonly', width=22, values=['2列目固定', '検索語句列の右', '一番右'])
        _map = {'fixed2': '2列目固定', 'after_base': '検索語句列の右', 'rightmost': '一番右'}
        cmb_insert.set(_map.get(var_insert_mode.get(), '2列目固定'))
        cmb_insert.grid(row=13, column=1, sticky='w', padx=(8, 0), pady=(6, 0))

        ttk.Separator(frm).grid(row=14, column=0, columnspan=2, sticky='ew', pady=(10, 8))
        ttk.Label(frm, text='Undo 最大数').grid(row=15, column=0, sticky='w')
        ttk.Spinbox(frm, from_=0, to=200, width=8, textvariable=var_undo_limit).grid(row=15, column=1, sticky='w', padx=(8, 0))

        btns = ttk.Frame(frm)
        btns.grid(row=16, column=0, columnspan=2, sticky="e", pady=(12, 0))

        def _ok():
            try:
                h = int(var_header.get())
                b = int(var_basecol.get())
                p = int(var_preview.get())
                if h < 1 or b < 1 or p < 1:
                    raise ValueError
            except Exception:
                messagebox.showerror("入力エラー", "見出し行・検索語句列・プレビュー行数は 1 以上の整数で指定してください。")
                return

            self.header_row_default = h
            self.base_col_index_default = b
            self.preview_rows_default = p
            self.confirm_rebuild = bool(var_confirm.get())
            # 起動時
            self.startup_open_last = bool(var_startup_open_last.get())
            self.startup_show_file_dialog = bool(var_startup_file_dialog.get())
            self.startup_always_show_load_settings = bool(var_startup_always_settings.get())
            # 検索リンク生成
            self.generate_ai = bool(var_gen_ai.get())
            self.generate_google = bool(var_gen_google.get())
            disp = (cmb_insert.get() or '').strip()
            if disp == '検索語句列の右':
                self.link_insert_mode = 'after_base'
            elif disp == '一番右':
                self.link_insert_mode = 'rightmost'
            else:
                self.link_insert_mode = 'fixed2'
            # Undo
            try:
                self.undo_limit = int(var_undo_limit.get())
            except Exception:
                self.undo_limit = 20
            self.save_config()
            dlg.destroy()

        def _cancel():
            dlg.destroy()

        ttk.Button(btns, text="キャンセル", command=_cancel).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="表示色設定...", command=self.show_color_settings).pack(side="left")
        ttk.Button(btns, text="OK", command=_ok).pack(side="right")

        dlg.bind("<Return>", lambda e: _ok())
        dlg.bind("<Escape>", lambda e: _cancel())

        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (dlg.winfo_width() // 2)
        y = self.root.winfo_rooty() + (self.root.winfo_height() // 2) - (dlg.winfo_height() // 2)
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def show_help_window(self):
        win = tk.Toplevel(self.root)
        win.title("使い方（かんたん）")
        win.geometry("560x320")
        win.grab_set()

        msg = (
            "【基本の流れ】\n"
            "1) Excelファイルを開く\n"
            "2) 見出し行・検索語句列を指定してOK\n"
            "3) 『検索語句更新』でリンク列を作成\n"
            "4) リンク列をダブルクリックでブラウザ検索\n\n"
            "【ヒント】\n"
            "・再読み込みでも同じ指定画面が出ます\n"
            "・設定 → 環境設定… で初期値を変更できます\n"
        )

        txt = tk.Text(win, wrap="word", height=14)
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", msg)
        txt.configure(state="disabled")

        ttk.Button(win, text="閉じる", command=win.destroy).pack(pady=6)

    def show_history_window(self):
        win = tk.Toplevel(self.root)
        win.title("操作履歴（Undo単位）")
        win.geometry("520x360")
        win.grab_set()

        tk.Label(win, text="直近の操作（新しい順）").pack(anchor="w", padx=10, pady=6)

        lb = tk.Listbox(win)
        lb.pack(fill="both", expand=True, padx=10, pady=6)

        history = getattr(self, "op_history", [])
        for s in reversed(history[-200:]):
            lb.insert("end", s)

        ttk.Button(win, text="閉じる", command=win.destroy).pack(pady=6)

    def _log_action(self, text: str):
        self.op_history.append(text)
        if len(self.op_history) > 500:
            self.op_history = self.op_history[-500:]

    # ---------------------
    # Undo/Redo helpers（A案）
    # ---------------------
    def _df_changed(self, before: pd.DataFrame, after: pd.DataFrame) -> bool:
        try:
            return not before.equals(after)
        except Exception:
            return True

    def commit_df(self, before: pd.DataFrame, after: pd.DataFrame, action: str, *, refresh_view=True) -> bool:
        """変更があった時だけ Undo積む/Redoクリア/未保存ON。"""
        if not self._df_changed(before, after):
            if refresh_view:
                self.current_df = after
                self.show_dataframe(self.current_df)
            self.update_undo_redo_buttons()
            return False

        self.undo_stack.append(before.copy())
        if len(self.undo_stack) > int(getattr(self, 'undo_limit', 20) or 20):
            self.undo_stack.pop(0)

        self.redo_stack.clear()
        self.current_df = after
        if refresh_view:
            self.show_dataframe(self.current_df)

        self.set_unsaved(True)
        self.update_undo_redo_buttons()
        self._log_action(action)
        return True

    def update_undo_redo_buttons(self):
        self.btn_undo.config(state="normal" if len(self.undo_stack) > 0 else "disabled")
        self.btn_redo.config(state="normal" if len(self.redo_stack) > 0 else "disabled")

    def undo(self):
        if not self.undo_stack or self.current_df is None:
            return
        self.redo_stack.append(self.current_df.copy())
        self.current_df = self.undo_stack.pop()
        self.show_dataframe(self.current_df)
        self.set_unsaved(True)
        self.update_undo_redo_buttons()
        self._log_action("Undo")

    def redo(self):
        if not self.redo_stack or self.current_df is None:
            return
        self.undo_stack.append(self.current_df.copy())
        self.current_df = self.redo_stack.pop()
        self.show_dataframe(self.current_df)
        self.set_unsaved(True)
        self.update_undo_redo_buttons()
        self._log_action("Redo")

    # ---------------------
    # UI
    # ---------------------
    def setup_treeview_style(self):
        style = ttk.Style(self.root)
        # 既定テーマの上で Treeview の見た目を整えます
        try:
            style.theme_use(style.theme_use())
        except Exception:
            pass


        # Treeview 全体（罫線風）
        style.configure(
            "Treeview",
            rowheight=24,
            borderwidth=1,
            relief="solid",
            background="white",
            fieldbackground="white"
        )
        style.map("Treeview", background=[("selected", "#cce5ff")])

        # ヘッダー（境界線を強調）
        style.configure(
            "Treeview.Heading",
            background="#ccffff",
            foreground="black",
            borderwidth=1,
            relief="solid"
        )

        style.configure("Vertical.TScrollbar", troughcolor="#f0f0f0", background="#e0e0e0")
        style.configure("Red.TButton", background="#ffaaaa", foreground="black")
        style.map("Red.TButton",
                  background=[('active', '#ff6666'), ('pressed', '#cc0000')],
                  foreground=[('active', 'black'), ('pressed', 'white')])

    def setup_ui(self):
        top = tk.Frame(self.root)
        top.pack(fill="x", pady=5)

        btn_add_col = ttk.Button(top, text="空白列追加", command=self.add_empty_column)
        btn_add_col.pack(side="left", padx=5)
        ToolTip(btn_add_col, "右側に新しい列を追加します。")

        btn_add_row = ttk.Button(top, text="空白行追加", command=self.add_empty_row)
        btn_add_row.pack(side="left", padx=5)
        ToolTip(btn_add_row, "一番下に新しい行を追加します。")

        btn_sel_base = ttk.Button(top, text="検索語句列の変更", command=self.select_base_columns)
        btn_sel_base.pack(side="left", padx=5)
        ToolTip(btn_sel_base, "検索の元になる列（例：商品名）を選びます。")

        btn_rebuild = ttk.Button(top, text="検索語句更新", command=self.rebuild_search_columns)
        btn_rebuild.pack(side="left", padx=5)
        ToolTip(btn_rebuild, "選んだ列からAI/Google検索リンクを作り直します。")

        btn_fit = ttk.Button(top, text="列幅自動調整", command=self.auto_adjust_columns)
        btn_fit.pack(side="left", padx=5)
        ToolTip(btn_fit, "見やすい幅に自動調整します。")

        self.btn_undo = ttk.Button(top, text="Undo", command=self.undo, state="disabled")
        self.btn_undo.pack(side="left", padx=5)
        ToolTip(self.btn_undo, "直前の変更を戻します（変更がある時だけ有効）。")

        self.btn_redo = ttk.Button(top, text="Redo", command=self.redo, state="disabled")
        self.btn_redo.pack(side="left", padx=5)
        ToolTip(self.btn_redo, "戻した変更をやり直します（Undoの後に有効）。")

        self.unsaved_label = tk.Label(top, text="", fg="red")
        self.unsaved_label.pack(side="left", padx=5)
        # 保存操作（ロータリーボックス + 保存ボタン）
        self.save_mode_var = tk.StringVar(value="別名で保存")
        self.save_mode_combo = ttk.Combobox(
            top,
            textvariable=self.save_mode_var,
            state="readonly",
            width=18,
            values=("元ファイルをコピー", "別名で保存", "CSV")
        )
        # 右側に配置（上書き保存ボタンは一番右）
        self.save_mode_combo.pack(side="right", padx=5)
        ToolTip(self.save_mode_combo, "保存方法を選択します。")

        def _do_save_selected():
            mode = self.save_mode_var.get()
            if mode == "元ファイルをコピー":
                self.copy_current_file()
            elif mode == "別名で保存":
                self.save_as_new()
            elif mode == "CSV":
                self.save_as_csv()

        self.btn_save = ttk.Button(top, text="保存", command=_do_save_selected)
        self.btn_save.pack(side="right", padx=5)
        ToolTip(self.btn_save, "選択した方法で保存します。保存後にExcelで開く確認が出ます。")

        # 上書き保存（赤）※一番右
        self.btn_overwrite = ttk.Button(top, text="上書き保存", command=self.save_current_file, style="Red.TButton")
        self.btn_overwrite.pack(side="right", padx=5)
        ToolTip(self.btn_overwrite, "元ファイルへ上書き保存します。保存後にExcelで開く確認が出ます。")

        btn_open = ttk.Button(top, text="別のファイルを開く", command=self.open_new_file)
        btn_open.pack(side="right", padx=5)
        ToolTip(btn_open, "別のExcelファイルを開きます（未保存がある場合は確認します）。")

        self.btn_reload = ttk.Button(top, text="再読み込み", command=self.reload_original)
        self.btn_reload.pack(side="right", padx=5)
        ToolTip(self.btn_reload, "元ファイルを読み込み直します（変更は破棄されます）。")

        
        # Treeview
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill="both", expand=True)

        self.vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        self.vsb.pack(side="right", fill="y")

        self.tree = ttk.Treeview(tree_frame, yscrollcommand=self.vsb.set)
        self.tree.pack(fill="both", expand=True)
        self.vsb.config(command=self.tree.yview)

        self.hsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        self.hsb.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=self.hsb.set)

        # Status bar（下部）
        status = tk.Frame(self.root, bd=1, relief="sunken")
        status.pack(side="bottom", fill="x")

        self.status_left = tk.Label(status, text="", anchor="w")
        self.status_left.pack(side="left", padx=6)

        self.status_mid = tk.Label(status, text="", anchor="w")
        self.status_mid.pack(side="left", padx=12)

        self.status_right = tk.Label(status, text="", anchor="e")
        self.status_right.pack(side="right", padx=6)

        self.status_msg = tk.Label(status, text="", fg="gray", anchor="w")
        self.status_msg.pack(side="right", padx=12)

        # Treeviewイベント
        self.tree.bind("<Button-1>", self.on_header_click)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.on_header_right_click)

    # ---------------------
    # Status / Toast
    # ---------------------
    def update_status_bar(self):
        if self.current_df is None:
            self.status_left.config(text="")
            self.status_mid.config(text="")
            self.status_right.config(text="")
            return
        rows = len(self.current_df)
        cols = len(self.current_df.columns)
        self.status_left.config(text=f"{rows} rows | {cols} cols")
        cols = getattr(self, "base_col_names", [])
        if not cols and self.base_col_name:
            cols = [self.base_col_name]
        disp = " + ".join(cols) if cols else "-"
        self.status_mid.config(text=f"検索語句列: {disp}")
        if self.sorted_col:
            arrow = "▲" if self.sort_state.get(self.sorted_col, True) else "▼"
            self.status_right.config(text=f"並び替え: {self.sorted_col} {arrow}")
        else:
            self.status_right.config(text="")

    def toast(self, msg: str, ms: int = 2800):
        self.status_msg.config(text=msg)
        def clear():
            self.status_msg.config(text="")
        self.root.after(ms, clear)

    # ---------------------
    # Excel / Treeview
    # ---------------------
    def load_excel(self, path):
        try:
            self.raw_df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl")
            self.header_row_current = int(getattr(self, "header_row_default", 1) or 1)
            self._build_current_df_from_raw()
            logging.info(f"Loaded: {path}")
        except Exception as e:
            messagebox.showerror("エラー", f"読み込み失敗: {e}")
            self.raw_df = None
            self.current_df = None

    def _reset_for_new_file(self):
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.set_unsaved(False)
        self.sorted_col = None
        self._onboard_shown = False
        self.op_history.clear()
        self.update_undo_redo_buttons()

    
    def _build_current_df_from_raw(self):
        """raw_df(全行)と header_row_current から current_df(ヘッダ下の表)を作る。"""
        if self.raw_df is None:
            self.current_df = None
            return
        hr = int(getattr(self, "header_row_current", 1) or 1)
        hr = max(1, hr)
        hdr_r = hr - 1
        if hdr_r >= len(self.raw_df):
            hdr_r = max(0, len(self.raw_df) - 1)
        ncols = int(self.raw_df.shape[1] or 0)
        header_vals = []
        if ncols > 0:
            header_vals = ["" if pd.isna(v) else str(v) for v in self.raw_df.iloc[hdr_r].tolist()]
            if len(header_vals) < ncols:
                header_vals += [""] * (ncols - len(header_vals))
            header_vals = header_vals[:ncols]
        # 空ヘッダは Excel列名で補う
        cols = []
        for i, v in enumerate(header_vals):
            name = (v or "").strip()
            if not name:
                name = f"列{get_excel_header(i+1)}"
            cols.append(name)
        # 重複ヘッダはサフィックスで回避
        seen = {}
        uniq = []
        for c in cols:
            if c not in seen:
                seen[c] = 1
                uniq.append(c)
            else:
                seen[c] += 1
                uniq.append(f"{c}_{seen[c]}")
        data = self.raw_df.iloc[hdr_r+1:].copy()
        data.columns = uniq[:data.shape[1]]
        data = data.fillna("").astype(str)
        self.current_df = data.reset_index(drop=True)

        # 現在のヘッダ名（表示・選択用に保存）
        self._header_vals_raw = header_vals
        self._current_columns = list(self.current_df.columns)

    def _compose_output_raw(self) -> pd.DataFrame:
        """raw_df(上部+ヘッダ行) + current_df(データ) から保存用の DataFrame を作る（header=Noneで書く）。"""
        if self.raw_df is None:
            return pd.DataFrame()
        hr = int(getattr(self, "header_row_current", 1) or 1)
        hr = max(1, hr)
        hdr_r = hr - 1
        hdr_r = min(hdr_r, max(0, len(self.raw_df)-1))
        pre = self.raw_df.iloc[:hdr_r].copy()
        header_row = self.raw_df.iloc[[hdr_r]].copy()

        # current_df の列数に合わせて列を拡張（リンク列追加のため）
        base_n = int(self.raw_df.shape[1] or 0)
        cur = self.current_df.copy() if self.current_df is not None else pd.DataFrame()
        # 保存は「表示の列順」を優先（current_df列順）
        cur_cols = list(cur.columns)
        # ヘッダ行（Excel上）は current_df の列名を反映
        # rawのヘッダ行を current_df 列数に合わせて再構築
        out_cols_n = max(base_n, len(cur_cols))
        # pre/headerの列数も合わせる
        def _pad_df(df, n):
            if df.shape[1] < n:
                for k in range(df.shape[1], n):
                    df[k] = ""
            return df.iloc[:, :n]
        pre = _pad_df(pre.reset_index(drop=True), out_cols_n)
        header_row = _pad_df(header_row.reset_index(drop=True), out_cols_n)

        # ヘッダ行を更新（リンク列もここに入れる）
        for i in range(out_cols_n):
            header_row.iat[0, i] = cur_cols[i] if i < len(cur_cols) else ("" if i >= base_n else header_row.iat[0, i])

        # データ行を current_df から作る（列数 out_cols_n に揃える）
        data_rows = []
        if len(cur) > 0:
            for _, r in cur.iterrows():
                row = [r.get(c, "") for c in cur_cols]
                if len(row) < out_cols_n:
                    row += [""] * (out_cols_n - len(row))
                data_rows.append(row[:out_cols_n])
        data_df = pd.DataFrame(data_rows)

        out = pd.concat([pre, header_row, data_df], ignore_index=True)
        out = out.fillna("").astype(str)
        return out

    def load_once(self):
        # 起動直後の動作（環境設定で切替）
        path = None

        # 1) 最後に開いたファイルを開く
        if getattr(self, "startup_open_last", False):
            cand = getattr(self, "last_file", "") or ""
            if cand and os.path.exists(cand):
                path = cand

        # 2) ファイル選択ダイアログ
        if path is None and getattr(self, "startup_show_file_dialog", True):
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

        # 3) 何も無ければ終了
        if not path:
            self.root.destroy()
            return

        # 起動時に読み込み設定ダイアログを必ず出すか？
        if getattr(self, "startup_always_show_load_settings", True):
            self._load_excel_with_dialog(path, first_time=True, force_select_base=False)
        else:
            ok = self._load_excel_no_dialog(path, first_time=True)
            if not ok:
                # 失敗したら従来方式で救済
                self._load_excel_with_dialog(path, first_time=True, force_select_base=False)

    def _load_excel_no_dialog(self, path: str, *, first_time: bool = False) -> bool:
        """設定の初期値だけで読み込む（ダイアログを出さない）。
        失敗したら False を返す。
        """
        try:
            header_row = int(getattr(self, "header_row_default", 1) or 1)
            base_col_index = int(getattr(self, "base_col_index_default", 1) or 1)
            self.raw_df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl")
            self.header_row_current = int(header_row)
            self._build_current_df_from_raw()
        except Exception as e:
            messagebox.showerror("エラー", f"読み込み失敗: {e}")
            self.current_df = None
            return False

        self.excel_path = path
        self.last_file = path
        self.save_config()
        self._reset_for_new_file()

        if self.current_df is None or len(self.current_df.columns) == 0:
            return False

        # 検索語句列（列番号→列名）
        try:
            idx = max(1, min(int(base_col_index), len(self.current_df.columns))) - 1
        except Exception:
            idx = 0
        self.base_col_name = str(self.current_df.columns[idx])
        # 既定：単一列（ここから後で複数選択に変えられます）
        if not getattr(self, "base_col_names", None):
            self.base_col_names = [self.base_col_name]

        # 検索リンク列が無い場合は、検索語句列の選択ウィンドウを出す
        missing_links = ("AI検索" not in self.current_df.columns) or ("Google検索" not in self.current_df.columns)
        if missing_links:
            self.select_base_columns()  # 適用時に rebuild_search_columns() まで実行
        else:
            self.show_dataframe(self.current_df)
            self.update_status_bar()

        return True

    def _load_excel_with_dialog(self, path: str, first_time: bool = False, force_select_base: bool = False):
        """ファイル読み込み時に、見出し行＆検索語句列を指定するダイアログを出してから読み込む。"""
        settings = self._show_load_settings_dialog(path)
        if settings is None:
            if first_time:
                self.root.destroy()
            return

        header_row, base_col_index = settings

        # 設定を記憶（次回起動時の初期値）
        self.header_row_default = header_row
        self.base_col_index_default = base_col_index
        self.save_config()

        # 実読み込み（見出し行をヘッダーとして扱う）
        try:
            self.raw_df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl")
            self.header_row_current = int(header_row)
            self._build_current_df_from_raw()
            logging.info(f"Loaded: {path} (header_row={header_row}, base_col_index={base_col_index})")
        except Exception as e:
            messagebox.showerror("エラー", f"読み込み失敗: {e}")
            self.current_df = None
            return

        self.excel_path = path
        self.last_file = path
        self.save_config()
        self._reset_for_new_file()

        if self.current_df is None or len(self.current_df.columns) == 0:
            return

        # 検索語句列（列番号→列名）
        try:
            idx = max(1, min(int(base_col_index), len(self.current_df.columns))) - 1
        except Exception:
            idx = 0
        self.base_col_name = str(self.current_df.columns[idx])
        # 検索リンク列（AI検索/Google検索）が無い場合は、検索語句列の選択ウィンドウを出す
        missing_links = ("AI検索" not in self.current_df.columns) or ("Google検索" not in self.current_df.columns)

        if force_select_base or missing_links:
            # ここで選択ダイアログを出して、選ばれた列でリンク列を生成
            self.select_base_columns()  # 適用時に rebuild_search_columns() まで実行
        else:
            # 既にリンク列がある場合はそのまま表示（必要なら手動で「検索語句更新」）
            self.show_dataframe(self.current_df)
            self.update_status_bar()

    def _show_load_settings_dialog(self, path: str):
        """見出し行 / 検索語句列を指定するダイアログ（プレビュー付き）。
        戻り値: (header_row:int, base_col_index:int) / None(キャンセル)
        """
        init_header = int(getattr(self, "header_row_default", 1) or 1)
        init_base = int(getattr(self, "base_col_index_default", 1) or 1)

        win = tk.Toplevel(self.root)
        win.title("読み込み設定（見出し行 / 検索語句列）")
        win.geometry("900x520")
        win.grab_set()

        info = tk.Label(
            win,
            text="見出し行（ヘッダー）と、検索語句列（列番号）を指定してください。\n"
                 "プレビューで見出し行をハイライトします。\n"
                 "※ 見出し行より上（タイトル/注記）は、読み込み後の一覧には表示しません。",
            justify="left",
        )
        info.pack(anchor="w", padx=10, pady=(10, 6))

        top = ttk.Frame(win)
        top.pack(fill="x", padx=10)

        ttk.Label(top, text="見出し行（1〜）:").pack(side="left")
        header_var = tk.IntVar(value=max(1, init_header))
        sp_header = ttk.Spinbox(top, from_=1, to=500, width=6, textvariable=header_var)
        sp_header.pack(side="left", padx=(6, 14))

        ttk.Label(top, text="検索語句列（1〜）:").pack(side="left")
        base_var = tk.IntVar(value=max(1, init_base))
        sp_base = ttk.Spinbox(top, from_=1, to=500, width=6, textvariable=base_var)
        sp_base.pack(side="left", padx=(6, 14))

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=10, pady=(6, 6))

        preview_frame = ttk.Frame(win)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # プレビュー読み込み（header=Noneで“生”の行を表示）
        try:
            preview_n = max(25, header_var.get() + 10)
            preview_df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl", nrows=preview_n)
        except Exception as e:
            messagebox.showerror("エラー", f"プレビュー読み込み失敗: {e}")
            win.destroy()
            return None

        ncols = int(preview_df.shape[1] or 1)
        sp_base.config(to=max(1, ncols))

        cols = [f"C{i+1}" for i in range(ncols)]
        tree = ttk.Treeview(preview_frame, columns=cols, show="headings", height=16)
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=120, anchor="w")

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        tree.tag_configure("headerline", background="#ffe6cc")

        result = {"value": None}

        def ensure_preview_rows(hr: int):
            nonlocal preview_df, ncols
            need_n = max(25, hr + 10)
            if len(preview_df) < need_n:
                try:
                    preview_df = pd.read_excel(path, header=None, dtype=str, engine="openpyxl", nrows=need_n)
                except Exception:
                    return
            ncols = int(preview_df.shape[1] or 1)

        def render():
            hr = max(1, int(header_var.get() or 1))
            ensure_preview_rows(hr)
            sp_base.config(to=max(1, ncols))

            tree.delete(*tree.get_children())
            for r in range(len(preview_df)):
                row_vals = []
                for c in range(ncols):
                    v = preview_df.iat[r, c] if c < preview_df.shape[1] else ""
                    row_vals.append("" if pd.isna(v) else str(v))
                tags = ("headerline",) if (r + 1) == hr else ()
                tree.insert("", "end", values=row_vals, tags=tags)

        def ok():
            hr = max(1, int(header_var.get() or 1))
            bc = max(1, int(base_var.get() or 1))
            result["value"] = (hr, bc)
            win.destroy()

        def cancel():
            result["value"] = None
            win.destroy()

        def on_change(*_):
            render()

        header_var.trace_add("write", on_change)
        base_var.trace_add("write", on_change)

        ttk.Button(btns, text="OK（読み込み）", command=ok).pack(side="right", padx=6)
        ttk.Button(btns, text="キャンセル", command=cancel).pack(side="right")

        render()
        self.root.wait_window(win)
        return result["value"]
    # データ操作
    # ---------------------
    def add_empty_column(self):
        if self.current_df is None:
            return
        before = self.current_df.copy()
        after = self.current_df.copy()
        new_col_name = f"新規列_{len(after.columns) + 1}"
        after[new_col_name] = ""
        self.commit_df(before, after, "空白列追加", refresh_view=True)
        self.update_status_bar()

    def add_empty_row(self):
        if self.current_df is None:
            return
        before = self.current_df.copy()
        after = self.current_df.copy()
        after.loc[len(after)] = [""] * len(after.columns)
        self.commit_df(before, after, "空白行追加", refresh_view=True)
        self.update_status_bar()

    # ---------------------
    # 検索語句列の選択
    # ---------------------

    def select_base_columns(self):
        """検索語句に使う列を複数選択できます（例：メーカー + 商品名）。"""
        if self.current_df is None or len(self.current_df.columns) == 0:
            return

        win = tk.Toplevel(self.root)
        win.title("検索語句の列を選択（複数可）")
        win.geometry("520x430")
        win.grab_set()

        tk.Label(win, text="AI/Google検索のキーワードに使う列を複数選択できます（Ctrl/Shiftで複数選択）。").pack(padx=16, pady=(12, 6), anchor="w")

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=16, pady=6)

        lb = tk.Listbox(frm, selectmode="extended", height=14)
        sb = ttk.Scrollbar(frm, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)

        lb.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        frm.grid_rowconfigure(0, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        cols = list(self.current_df.columns)
        for c in cols:
            lb.insert("end", c)

        # 既存選択を復元（base_col_names 優先）
        selected = []
        if getattr(self, "base_col_names", None):
            selected = [c for c in self.base_col_names if c in cols]
        elif self.base_col_name and self.base_col_name in cols:
            selected = [self.base_col_name]

        for i, c in enumerate(cols):
            if c in selected:
                lb.selection_set(i)

        opt = ttk.Frame(win)
        opt.pack(fill="x", padx=16, pady=(4, 0))

        ttk.Label(opt, text="結合の区切り:").pack(side="left")
        joiner_var = tk.StringVar(value=getattr(self, "base_joiner", " "))
        ent_join = ttk.Entry(opt, textvariable=joiner_var, width=12)
        ent_join.pack(side="left", padx=(6, 10))
        ttk.Label(opt, text='例：" "（半角スペース） / " / " / "\\t"').pack(side="left")

        preview = tk.Label(win, text="", fg="gray")
        preview.pack(fill="x", padx=16, pady=(6, 0), anchor="w")

        def render_preview():
            sel = [cols[i] for i in lb.curselection()]
            j = joiner_var.get().replace("\\t", "\t")
            if not sel:
                preview.config(text="選択なし（最低1列は選んでください）")
            else:
                preview.config(text="検索語句: " + j.join(sel))

        lb.bind("<<ListboxSelect>>", lambda e: render_preview())
        joiner_var.trace_add("write", lambda *args: render_preview())
        render_preview()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=16, pady=12)

        def decide():
            sel_idx = list(lb.curselection())
            if not sel_idx:
                messagebox.showerror("入力エラー", "検索語句列を1つ以上選択してください。")
                return
            sel_cols = [cols[i] for i in sel_idx]
            self.base_col_names = sel_cols
            self.base_col_name = sel_cols[0]  # 互換（旧処理・表示用）
            self.base_joiner = joiner_var.get().replace("\\t", "\t").replace("\\t", "\t")

            # 次回初期選択用に列番号を保存（1始まり）
            self.base_col_indices_default = [i + 1 for i in sel_idx]

            self.save_config()
            win.destroy()
            self.rebuild_search_columns()

        ttk.Button(btns, text="キャンセル", command=win.destroy).pack(side="right")
        ttk.Button(btns, text="適用", command=decide).pack(side="right", padx=(0, 8))


    def _build_keyword_series(self) -> pd.Series:
        """選択された複数列から検索語句を合成した Series を返します。"""
        if self.current_df is None:
            return pd.Series([], dtype=str)

        cols = []
        # base_col_names を優先し、なければ base_col_name を使う
        if getattr(self, "base_col_names", None):
            cols = [c for c in self.base_col_names if c in self.current_df.columns]
        elif self.base_col_name and self.base_col_name in self.current_df.columns:
            cols = [self.base_col_name]

        if not cols:
            # 最低限：先頭列
            cols = [str(self.current_df.columns[0])]

        joiner = getattr(self, "base_joiner", " ")
        joiner = "\t" if joiner == "\\t" else joiner

        def make_row_keyword(row):
            parts = []
            for c in cols:
                v = safe_text(row.get(c, ""))
                v = v.strip()
                if v:
                    parts.append(v)
            return joiner.join(parts).strip()

        return self.current_df.apply(make_row_keyword, axis=1)

    def _make_hyperlink_formula(self, text: str, template: str, label: str) -> str:
        """Excelの=HYPERLINK式を作る（template内の{q}をURLエンコードした検索語句に置換）"""
        text = safe_text(text)
        if not text:
            return ""
        tpl = (template or "").strip()
        if not tpl:
            return ""
        if "{q}" not in tpl:
            join = "&" if "?" in tpl else "?"
            tpl = tpl + join + "q={q}"
        url = tpl.replace("{q}", urllib.parse.quote(text))
        return f'=HYPERLINK("{url}","{label}")'


    # ---------------------
    # 確認ダイアログ（チェック付き）
    # ---------------------
    def confirm_rebuild_dialog(self) -> bool:
        win = tk.Toplevel(self.root)
        win.title("確認")
        win.geometry("460x190")
        win.grab_set()

        tk.Label(
            win,
            text="検索リンクを更新しますか？\n"
                 "※元データは消えません（Undoで戻せます）",
            justify="left"
        ).pack(padx=14, pady=10, anchor="w")

        var_noask = tk.BooleanVar(value=False)
        chk = ttk.Checkbutton(win, text="今後は確認しない", variable=var_noask)
        chk.pack(padx=14, pady=6, anchor="w")

        result = {"ok": False}

        def yes():
            result["ok"] = True
            if var_noask.get():
                self.confirm_rebuild = False
                self.save_config()
            win.destroy()

        def no():
            result["ok"] = False
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=10)
        ttk.Button(btns, text="はい", command=yes).pack(side="left", padx=8)
        ttk.Button(btns, text="いいえ", command=no).pack(side="left", padx=8)

        win.wait_window()
        return result["ok"]

    def rebuild_search_columns(self):
        if self.current_df is None:
            return
        valid_cols = []
        if getattr(self, "base_col_names", None):
            valid_cols = [c for c in self.base_col_names if c in self.current_df.columns]
        elif self.base_col_name and self.base_col_name in self.current_df.columns:
            valid_cols = [self.base_col_name]
        if not valid_cols:
            self.select_base_columns()
            return

        if self.confirm_rebuild:
            if not self.confirm_rebuild_dialog():
                return

        self.finish_edit(None)
        before = self.current_df.copy()

        df = self.current_df.copy()
        # 生成対象
        gen_ai = bool(getattr(self, 'generate_ai', True))
        gen_google = bool(getattr(self, 'generate_google', True))
        if (not gen_ai) and (not gen_google):
            gen_ai = True  # どちらもOFFは事故るので救済

        # 既存のリンク列は作り直す前提でいったん除外
        base_cols = [c for c in df.columns if c not in ['AI検索', 'Google検索']]
        df = df[base_cols].copy()

        keywords = self._build_keyword_series()

        link_cols = []
        if gen_ai:
            df['AI検索'] = keywords.apply(lambda t: self._make_hyperlink_formula(t, getattr(self, 'ai_url_template', 'https://www.perplexity.ai/search?q={q}'), 'AI検索'))
            link_cols.append('AI検索')
        if gen_google:
            df['Google検索'] = keywords.apply(lambda t: self._make_hyperlink_formula(t, 'https://www.google.com/search?q={q}', 'Google検索'))
            link_cols.append('Google検索')

        # 挿入位置
        mode = getattr(self, 'link_insert_mode', 'fixed2')
        cols = [c for c in df.columns if c not in link_cols]

        def insert_at(pos: int):
            nonlocal cols
            pos = max(0, min(pos, len(cols)))
            for i, lc in enumerate(link_cols):
                cols.insert(pos + i, lc)

        if mode == 'after_base':
            try:
                base_pos = cols.index(self.base_col_name) + 1
            except Exception:
                base_pos = 1
            insert_at(base_pos)
        elif mode == 'rightmost':
            insert_at(len(cols))
        else:
            # fixed2: 2列目固定
            insert_at(1)

        df = df[cols]

        changed = self.commit_df(before, df, "検索リンク更新", refresh_view=True)
        self.update_status_bar()

        if changed and (not self._onboard_shown):
            self._onboard_shown = True
            self.toast(
                "使い方：1) 普通の列を編集 2)「検索語句更新」 3) AI/Google列をダブルクリックで検索を開く（編集不可）",
                5200
            )

    # ---------------------
    # Treeview表示
    # ---------------------
    def apply_row_colors(self):
        """Treeviewのタグ色を設定値から反映"""
        if not getattr(self, "tree", None):
            return
        colors = getattr(self, "colors", None) or {}
        header = colors.get("headerrow", "#ffb6c1")
        pre = colors.get("preheader", "#f5f5f5")
        even = colors.get("even", "#ffffff")
        odd = colors.get("odd", "#f9f9f9")
        try:
            self.tree.tag_configure("headerrow", background=header, foreground="black")
            self.tree.tag_configure("preheader", background=pre, foreground="#666666")
            self.tree.tag_configure("even", background=even)
            self.tree.tag_configure("odd", background=odd)
        except Exception:
            pass

    def show_dataframe(self, df):
        """表示：Excelで見える行はすべて表示（固定なし）
        - raw_df の全行を表示
        - 指定の見出し行は強調表示し、列ヘッダ(heading)の表示文字にも使う
        """
        self.tree.delete(*self.tree.get_children())

        if self.raw_df is None:
            return

        hr = int(getattr(self, "header_row_current", getattr(self, "header_row_default", 1)) or 1)
        hr = max(1, hr)
        hdr_r = hr - 1
        hdr_r = min(hdr_r, max(0, len(self.raw_df)-1))

        base_n = int(self.raw_df.shape[1] or 0)
        cur_n = int(getattr(self.current_df, "shape", (0,0))[1] or 0) if self.current_df is not None else 0
        ncols = max(base_n, cur_n)

        cols = [f"__c{i}__" for i in range(ncols)]
        self.tree["columns"] = cols
        self.tree["show"] = "headings"

        self.apply_row_colors()

        # ヘッダ行の値をheading表示に反映
        header_vals = []
        try:
            header_vals = ["" if pd.isna(v) else str(v) for v in self.raw_df.iloc[hdr_r].tolist()]
        except Exception:
            header_vals = []

        # current_df の列順（検索列を含む）をそのまま尊重
        cur_cols = list(self.current_df.columns) if self.current_df is not None else []

        # 表示列(index) -> raw列(index) の対応
        # raw_df の列は header=None 読み込みなので通常 int（0,1,2...）。
        # 検索列（AI検索/Google検索）や「行」など文字列列は raw に対応させない（None）。
        
        # 表示列(index) -> raw列(index) の対応
        # raw_df から作った元の列名（self._current_columns）に一致する列だけ raw に対応させる。
        # 検索列（AI検索/Google検索）など追加列は raw に存在しないため None 扱い。
        base_col_names = list(getattr(self, "_current_columns", []) or [])
        raw_name_to_idx = {name: i for i, name in enumerate(base_col_names)}

        view_to_raw = []
        for name in cur_cols:
            if name in raw_name_to_idx:
                view_to_raw.append(raw_name_to_idx[name])
            else:
                view_to_raw.append(None)

        # current_df が無い場合は raw をそのまま
        if not cur_cols:
            view_to_raw = [i for i in range(ncols)]

        # 念のため長さを合わせる
        if len(view_to_raw) < ncols:
            view_to_raw.extend([None] * (ncols - len(view_to_raw)))

        self._view_to_raw_index = view_to_raw

        
        for i, col in enumerate(cols):
            label = ""
            # raw列に対応する場合はヘッダ行の値を優先
            raw_idx = None
            try:
                raw_idx = view_to_raw[i] if i < len(view_to_raw) else None
            except Exception:
                raw_idx = None

            if raw_idx is not None and raw_idx < len(header_vals):
                label = header_vals[raw_idx]
            elif i < len(cur_cols):
                # 検索列など（文字列）
                label = str(cur_cols[i])
            elif i < len(header_vals):
                label = header_vals[i]

            text = f"{get_excel_header(i+1)} {label}".strip()
            self.tree.heading(col, text=text)
            self.tree.column(col, width=150, anchor="w")

        # tag styles

        # raw部分（0..hdr_r）をそのまま表示（リンク列は空欄）
        view_to_raw = getattr(self, "_view_to_raw_index", [i for i in range(ncols)])
        for r in range(0, hdr_r+1):
            row_vals = [""] * ncols
            for i in range(ncols):
                raw_idx = view_to_raw[i] if i < len(view_to_raw) else None
                if raw_idx is None:
                    row_vals[i] = ""
                else:
                    try:
                        v = self.raw_df.iat[r, raw_idx]
                    except Exception:
                        v = ""
                    row_vals[i] = "" if pd.isna(v) else str(v)
            tag = ("headerrow",) if r == hdr_r else ("preheader",)
            self.tree.insert("", "end", values=row_vals, tags=tag)

        # data部分（hdr_r+1..）は current_df を表示（rawが短い場合は補完）
        if self.current_df is not None:
            for i, row in enumerate(self.current_df.itertuples(index=False)):
                vals = [display_text(v) for v in row]
                if len(vals) < ncols:
                    vals += [""] * (ncols - len(vals))
                tag = ("even",) if (i % 2 == 0) else ("odd",)
                self.tree.insert("", "end", values=vals[:ncols], tags=tag)

        self.update_status_bar()


    # ---------------------
    # 列幅自動調整
    # ---------------------
    def auto_adjust_columns(self):
        if self.current_df is None:
            return
        CHAR_WIDTH = 7
        MIN_WIDTH = 60
        MAX_WIDTH = 400

        for col in self.tree["columns"]:
            header_text = self.tree.heading(col, "text")
            max_len = len(str(header_text)) if header_text else 0

            for item in self.tree.get_children():
                val = self.tree.set(item, col)
                if val is not None:
                    max_len = max(max_len, len(str(val)))

            width = max_len * CHAR_WIDTH + 16
            width = max(MIN_WIDTH, min(width, MAX_WIDTH))
            self.tree.column(col, width=width)

        self.toast("列幅を自動調整しました。", 1800)

    # ---------------------
    # ヘッダー: シングルクリック = ソート（遅延）
    # ---------------------
    def on_header_click(self, event):
        if self.current_df is None:
            return
        region = self.tree.identify_region(event.x, event.y)

        # リンク列をクリックしたときの説明（初心者向け）
        if region == "heading":
            col = self.tree.identify_column(event.x)
            if col:
                idx = int(col.replace("#", "")) - 1
                if 0 <= idx < len(self.current_df.columns):
                    col_name = self.current_df.columns[idx]
                    if col_name in ("AI検索", "Google検索"):
                        self.toast("リンク専用列です（編集不可）。ダブルクリックで開きます。", 2400)
                        return

        if region != "heading":
            return

        col = self.tree.identify_column(event.x)
        if not col:
            return
        c = int(col.replace("#", "")) - 1
        if c < 0 or c >= len(self.current_df.columns):
            return
        col_name = self.current_df.columns[c]

        if col_name in ("AI検索", "Google検索"):
            return

        if self._header_click_job:
            try:
                self.root.after_cancel(self._header_click_job)
            except Exception:
                pass

        self._header_click_col = col_name
        self._header_click_job = self.root.after(220, self._do_sort_reserved)

    def _do_sort_reserved(self):
        self._header_click_job = None
        col_name = self._header_click_col
        self._header_click_col = None
        if not col_name or self.current_df is None:
            return
        self.sort_by_column(col_name)

    def sort_by_column(self, col_name):
        if self.current_df is None:
            return
        asc = self.sort_state.get(col_name, True)
        try:
            self.current_df[col_name] = pd.to_numeric(self.current_df[col_name], errors='ignore')
        except Exception:
            pass
        self.current_df = self.current_df.sort_values(by=col_name, ascending=asc, kind="mergesort").reset_index(drop=True)
        self.sort_state[col_name] = not asc
        self.sorted_col = col_name
        self.show_dataframe(self.current_df)

    # ---------------------
    # ダブルクリック
    # ---------------------
    def on_double_click(self, event):
        if self.current_df is None:
            return

        if self._header_click_job:
            try:
                self.root.after_cancel(self._header_click_job)
            except Exception:
                pass
            self._header_click_job = None
            self._header_click_col = None

        region = self.tree.identify_region(event.x, event.y)

        if region == "heading":
            col = self.tree.identify_column(event.x)
            if not col:
                return
            c = int(col.replace("#", "")) - 1
            if c < 0 or c >= len(self.current_df.columns):
                return
            old_name = self.current_df.columns[c]

            if old_name in ("AI検索", "Google検索"):
                return

            self.open_header_name_editor(old_name)
            return

        if region == "cell":
            self.start_edit(event)
            return

    # ---------------------
    # ヘッダー名編集
    # ---------------------
    def open_header_name_editor(self, old_name):
        if self.current_df is None:
            return

        win = tk.Toplevel(self.root)
        win.title("列名の変更")
        win.geometry("420x160")
        win.grab_set()

        tk.Label(win, text="新しい列名を入力してください：").pack(pady=8)
        ent = tk.Entry(win)
        ent.pack(fill="x", padx=12)
        ent.insert(0, old_name)
        ent.focus_set()

        tk.Label(win, text="※ヘッダー：シングル=並び替え / ダブル=列名変更 / 右クリック=列一括編集", fg="gray").pack(pady=4)

        def apply():
            if self.current_df is None:
                return
            new_name = ent.get().strip()
            if not new_name:
                return

            if new_name in ("AI検索", "Google検索"):
                messagebox.showerror("エラー", "その列名は予約されています。")
                return

            if new_name in self.current_df.columns and new_name != old_name:
                messagebox.showerror("エラー", "同名の列が既にあります。")
                return

            before = self.current_df.copy()
            after = self.current_df.rename(columns={old_name: new_name})

            changed = self.commit_df(before, after, f"列名変更: {old_name} → {new_name}", refresh_view=True)
            if changed:
                if self.base_col_name == old_name:
                    self.base_col_name = new_name
                    self.save_config()
                win.destroy()

        ttk.Button(win, text="適用", command=apply).pack(pady=10)

    # ---------------------
    # ヘッダー右クリック（列一括編集）
    # ---------------------
    def on_header_right_click(self, event):
        if self.current_df is None:
            return
        if self.tree.identify_region(event.x, event.y) != "heading":
            return

        col = self.tree.identify_column(event.x)
        if not col:
            return
        col_index = int(col.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.current_df.columns):
            return
        col_name = self.current_df.columns[col_index]

        if col_name in ("AI検索", "Google検索"):
            self.toast("リンク専用列です（編集不可）。ダブルクリックで開きます。", 2400)
            return

        self.open_formula_editor(col_name)

    # ---------------------
    # セル編集
    # ---------------------
    def start_edit(self, event):
        if self.current_df is None:
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        r = self.tree.index(row_id)
        c = int(col_id.replace("#", "")) - 1

        # view上の行 r は raw行(0..hdr_r) + data行(0..)
        hr = int(getattr(self, "header_row_current", getattr(self, "header_row_default", 1)) or 1)
        hdr_r = max(0, hr - 1)

        # ヘッダ行は編集不可
        if r == hdr_r:
            self.toast("見出し行は編集できません。", 2000)
            return

        is_pre = (r < hdr_r)
        data_index = r - (hdr_r + 1)

        if is_pre:
            # raw側編集（リンク列は編集不可）
            view_to_raw = getattr(self, "_view_to_raw_index", None)
            raw_c = None
            if view_to_raw and c < len(view_to_raw):
                raw_c = view_to_raw[c]
            if raw_c is None:
                return
            col_name = f"RAW{raw_c+1}"
            self._edit_raw_col = int(raw_c)
        else:
            if self.current_df is None:
                return
            if data_index < 0 or data_index >= len(self.current_df):
                return
            if c < 0 or c >= len(self.current_df.columns):
                return
            col_name = self.current_df.columns[c]

        if col_name in ("AI検索", "Google検索"):
            self.toast("リンク列です。ダブルクリックで検索を開きます（編集不可）。", 2400)
            url = extract_url(self.current_df.iat[r, c])
            if url:
                webbrowser.open(url)
            return

        x, y, w, h = self.tree.bbox(row_id, col_id)
        self.edit_entry = tk.Entry(self.tree)
        self.edit_entry.place(x=x, y=y, width=w, height=h)
        if is_pre:
            raw_c = int(getattr(self, "_edit_raw_col", c))
            old_v = self.raw_df.iat[r, raw_c] if self.raw_df is not None else ""
        else:
            old_v = self.current_df.iat[data_index, c]
        self.edit_entry.insert(0, display_text(old_v))
        self.edit_entry.focus()

        self.edit_row, self.edit_col = r, c
        self._edit_is_pre = bool(is_pre)
        self._edit_data_index = int(data_index)
        self.edit_entry.bind("<Return>", self.finish_edit)
        self.edit_entry.bind("<Escape>", lambda e: self._cancel_edit())

    def _cancel_edit(self):
        if self.edit_entry:
            try:
                self.edit_entry.destroy()
            except Exception:
                pass
        self.edit_entry = None
        self.edit_row = None
        self.edit_col = None

    def finish_edit(self, event):
        if self.current_df is None or self.edit_entry is None:
            return
        val = self.edit_entry.get()
        r_view = int(self.edit_row)
        c = int(self.edit_col)
        is_pre = bool(getattr(self, "_edit_is_pre", False))
        data_index = int(getattr(self, "_edit_data_index", -1))
        if is_pre:
            old = display_text(self.raw_df.iat[r_view, c]) if self.raw_df is not None else ""
        else:
            old = display_text(self.current_df.iat[data_index, c])

        if val != old:
            if is_pre:
                # raw側（上部）編集：undo対象外（シンプル運用）
                raw_c = int(getattr(self, "_edit_raw_col", c))
                try:
                    if self.raw_df is not None:
                        self.raw_df.iat[r_view, raw_c] = val
                except Exception:
                    pass
                # 表を再構築して表示更新
                self._build_current_df_from_raw()
                self.show_dataframe(self.current_df)
                self.set_unsaved(True)
                self._log_action(f"上部行編集: R{r_view+1}C{c+1}")
            else:
                before = self.current_df.copy()
                after = self.current_df.copy()
                after.iat[data_index, c] = val
                changed = self.commit_df(before, after, f"セル編集: R{data_index+hr+1+1}C{c+1}", refresh_view=True)
                if changed:
                    # raw_dfにも反映（データ領域）
                    try:
                        if self.raw_df is not None:
                            hdr_r = max(0, hr-1)
                            raw_r = (hdr_r + 1) + data_index
                            # 表示列 -> raw列
                            view_to_raw = getattr(self, "_view_to_raw_index", None)
                            raw_c = None
                            if view_to_raw and c < len(view_to_raw):
                                raw_c = view_to_raw[c]
                            if raw_c is None:
                                # リンク列などはrawに反映しない
                                raise Exception("no raw column")
                            raw_c = int(raw_c)
                            # rawの列数が足りなければ拡張
                            if raw_c >= self.raw_df.shape[1]:
                                for k in range(self.raw_df.shape[1], raw_c+1):
                                    self.raw_df[k] = ""
                            if raw_r < len(self.raw_df):
                                self.raw_df.iat[raw_r, raw_c] = val
                            else:
                                # rawが短い場合は行追加
                                while len(self.raw_df) <= raw_r:
                                    self.raw_df.loc[len(self.raw_df)] = [""] * self.raw_df.shape[1]
                                self.raw_df.iat[raw_r, raw_c] = val
                    except Exception:
                        pass

        self._cancel_edit()

    # ---------------------
    # 列全体編集（プレビュー付き）
    # ---------------------
    def open_formula_editor(self, col_name):
        if self.current_df is None:
            return

        win = tk.Toplevel(self.root)
        win.title(f"{col_name} の列一括編集")
        win.geometry("640x390")
        win.grab_set()

        tk.Label(
            win,
            text="ここに入力した内容を、この列の全行にまとめて入れます。\n\n"
                 "Excelで使う「式」や「文字」を、そのまま書けます。\n"
                 "{ROW} は「行番号」に自動で置き換わります。\n\n"
                 "■ よく使う例\n"
                 "① 他の列の値をつなげる  =A{ROW}&\"_\"&B{ROW}\n"
                 "② 金額に「円」を付ける   =A{ROW}&\"円\"\n"
                 "③ 割り算（単価・比率など）=F{ROW}/E{ROW}\n"
                 "④ 固定の文字を入れる     メモ\n"
                 "⑤ 日付や関数を使う       =TODAY()\n\n"
                 "行番号の例：\n"
                 "  入力：=F{ROW}/E{ROW}\n"
                 "  2行目 → =F2/E2\n"
                 "  3行目 → =F3/E3\n\n"
                 "※ この画面では計算されません。Excelで開いたときに計算されます。\n"
                 "※ 下に先頭3行のプレビューが出ます。",
            justify="left"
        ).pack(pady=8, anchor="w", padx=10)

        ent = tk.Entry(win)
        ent.pack(fill="x", padx=10, pady=6)
        ent.insert(0, '例：=A{ROW}&\"_\"&B{ROW}')

        tk.Label(win, text="プレビュー（先頭3行）", anchor="w").pack(fill="x", padx=10)

        pv = tk.Text(win, height=6, wrap="none")
        pv.pack(fill="x", padx=10, pady=6)
        pv.configure(state="disabled")

        def compute_val(formula: str, row_no: int) -> str:
            if "{ROW}" in formula:
                return formula.replace("{ROW}", str(row_no))
            return formula.replace("2", str(row_no))  # 互換（非推奨）

        def render_preview():
            if self.current_df is None:
                return
            formula = ent.get()
            lines = []
            n = min(3, len(self.current_df))
            for i in range(n):
                row_no = i + 2
                val = compute_val(formula, row_no)
                lines.append(f"行{row_no}: {val}")
            if n == 0:
                lines = ["(データがありません)"]
            pv.configure(state="normal")
            pv.delete("1.0", "end")
            pv.insert("1.0", "\n".join(lines))
            pv.configure(state="disabled")

        ent.bind("<KeyRelease>", lambda e: render_preview())
        render_preview()

        def apply_changes():
            if self.current_df is None:
                return
            formula = ent.get()
            before = self.current_df.copy()

            new_col = []
            for i in range(len(self.current_df)):
                row_no = i + 2
                new_col.append(compute_val(formula, row_no))

            after = self.current_df.copy()
            after[col_name] = new_col
            self.commit_df(before, after, f"列一括編集: {col_name}", refresh_view=True)
            win.destroy()

        ttk.Button(win, text="適用", command=apply_changes).pack(pady=10)

    # ---------------------
    # ファイル操作
    # ---------------------
    def set_unsaved(self, flag: bool):
        self.unsaved_changes = flag
        self.unsaved_label.config(text="● 未保存" if flag else "")
    def open_new_file(self):
        if self.unsaved_changes and not messagebox.askyesno("確認", "変更を破棄して新しいファイルを開きますか？"):
            return
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self._load_excel_with_dialog(path, first_time=False, force_select_base=False)
    def reload_original(self):
        if not self.excel_path:
            return
        msg = "元ファイルを再読み込みします。\\n（変更は破棄されます）\\n続行しますか？"
        if self.unsaved_changes and not messagebox.askyesno("再読み込み", msg):
            return
        # 再読み込み時も、見出し行/検索語句列の指定ダイアログを表示
        self._load_excel_with_dialog(self.excel_path, first_time=False, force_select_base=True)
    def prompt_open_in_excel(self, path: str):
        """保存後にExcelで開くか確認し、はいなら開く。"""
        try:
            ans = messagebox.askyesno("保存", "保存したファイルをexcelで開きますか？")
        except Exception:
            ans = False
        if ans:
            try:
                import os
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("Excel起動", f"Excelで開けませんでした: {e}")
    def show_color_settings(self):
        """表示色設定（A案：見出し/グレー/偶奇）"""
        win = tk.Toplevel(self.root)
        win.title("表示色設定")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        items = [
            ("headerrow", "見出し(強調)行", self.colors.get("headerrow", "#ffb6c1")),
            ("preheader", "見出しより上(グレー)", self.colors.get("preheader", "#f5f5f5")),
            ("even", "通常行(偶数)", self.colors.get("even", "#ffffff")),
            ("odd", "通常行(奇数)", self.colors.get("odd", "#f9f9f9")),
        ]

        vars_ = {}
        for key, label, default in items:
            row = ttk.Frame(frm)
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=label, width=20).pack(side="left")
            v = tk.StringVar(value=default)
            vars_[key] = v
            ttk.Entry(row, textvariable=v, width=12).pack(side="left", padx=6)

            def _pick(k=key):
                cur = vars_[k].get()
                rgb, hx = colorchooser.askcolor(color=cur, parent=win)
                if hx:
                    vars_[k].set(hx)

            ttk.Button(row, text="選択", command=_pick).pack(side="left")

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))

        def _ok():
            if "Colors" not in self.config:
                self.config["Colors"] = {}
            for k, v in vars_.items():
                self.config["Colors"][k] = v.get()
                self.colors[k] = v.get()
            try:
                self.save_config()
            except Exception:
                pass
            self.apply_row_colors()
            try:
                self.show_dataframe(self.current_df)
            except Exception:
                pass
            win.destroy()

        ttk.Button(btns, text="OK", command=_ok).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="キャンセル", command=lambda: win.destroy()).pack(side="right")


    def show_save_dialog(self):
        """保存ダイアログ：[コピー][別名で保存][上書き保存][CSV]"""
        if not getattr(self, "excel_path", None):
            messagebox.showwarning("保存", "先にExcelファイルを開いてください。")
            return

        win = tk.Toplevel(self.root)
        win.title("保存")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="保存方法を選んでください").pack(anchor="w")

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))

        def _close():
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()

        def _copy():
            self.copy_current_file()
            _close()

        def _save_as():
            self.save_as_new()
            _close()

        def _overwrite():
            self.save_current_file()
            _close()

        def _csv():
            self.save_as_csv()
            _close()

        ttk.Button(btns, text="コピー", command=_copy).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="別名で保存", command=_save_as).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="上書き保存", command=_overwrite).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="CSV", command=_csv).pack(side="left")

        ttk.Button(frm, text="閉じる", command=_close).pack(anchor="e", pady=(10, 0))

    def copy_current_file(self):
        """元ファイルと同じフォルダーにコピーを作成する"""
        if not getattr(self, "excel_path", None):
            messagebox.showwarning("コピー", "元ファイルがありません。")
            return
        src_path = Path(self.excel_path)
        if not src_path.exists():
            messagebox.showerror("コピー", "元ファイルが見つかりません。")
            return

        # 保存内容を反映したコピーを作る（現在の表示/編集内容を書き出す）
        out_df = self._compose_output_raw()

        folder = src_path.parent
        stem = src_path.stem
        suffix = src_path.suffix or ".xlsx"

        # 重複しないファイル名を作る
        cand = folder / f"{stem}_copy{suffix}"
        i = 2
        while cand.exists():
            cand = folder / f"{stem}_copy{i}{suffix}"
            i += 1

        try:
            out_df.to_excel(cand, index=False, header=False)
            self.prompt_open_in_excel(str(cand))
            self.set_unsaved(False)
            self.toast(f"コピー作成: {cand.name}", 2500)
            logging.info(f"Copied to: {cand}")
        except Exception as e:
            messagebox.showerror("コピー", f"コピー作成に失敗しました: {e}")

    def save_as_csv(self):
        """CSVを書き出し（headerなしでExcel見た目を維持）"""
        if not getattr(self, "excel_path", None):
            messagebox.showwarning("CSV", "先にExcelファイルを開いてください。")
            return
        try:
            initial = str(Path(self.excel_path).with_suffix(".csv"))
        except Exception:
            initial = "output.csv"

        csv_path = filedialog.asksaveasfilename(
            title="CSVとして保存",
            defaultextension=".csv",
            initialfile=os.path.basename(initial),
            filetypes=[("CSV", "*.csv")],
        )
        if not csv_path:
            return

        try:
            out_df = self._compose_output_raw()
            out_df.to_csv(csv_path, index=False, header=False, encoding="utf-8-sig")
            self.prompt_open_in_excel(csv_path)
            self.toast("CSV保存しました", 2000)
            logging.info(f"Saved CSV: {csv_path}")
        except Exception as e:
            messagebox.showerror("CSV", f"CSV保存に失敗しました: {e}")


    def save_current_file(self):
        if self.current_df is None or not self.excel_path:
            return False
        try:
            out_df = self._compose_output_raw()
            out_df.to_excel(self.excel_path, index=False, header=False)
            self.prompt_open_in_excel(self.excel_path)
            self.set_unsaved(False)
            self.toast("保存しました。", 1600)
            return True
        except Exception as e:
            messagebox.showerror("エラー", f"保存失敗: {e}")
            return False

    def save_and_open_choice(self):
        if not self.excel_path:
            return
        if not self.save_current_file():
            return
        choice = messagebox.askquestion("開く方法", "Excelで開きますか？\n（いいえ：CSVを保存してフォルダを開きます）")
        if choice == "yes":
            try:
                os.startfile(self.excel_path)
            except Exception as e:
                messagebox.showerror("エラー", f"開けません: {e}")
        else:
            csv_path = os.path.abspath(re.sub(r"\.xls[xm]?$", ".csv", self.excel_path, flags=re.IGNORECASE))
            try:
                out_df = self._compose_output_raw()
                out_df.to_csv(csv_path, index=False, header=False, encoding="utf-8-sig")
                folder = os.path.dirname(csv_path)
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("エラー", f"保存/表示に失敗: {e}")

    def save_as_new(self):
        if self.current_df is None:
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            try:
                out_df = self._compose_output_raw()
                out_df.to_excel(path, index=False, header=False)
                self.prompt_open_in_excel(path)
                messagebox.showinfo("保存", "保存しました。")
                self.set_unsaved(False)
            except Exception as e:
                messagebox.showerror("エラー", f"保存失敗: {e}")

# =====================
# Main
# =====================
def main():
    root = tk.Tk()
    ExcelViewerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
