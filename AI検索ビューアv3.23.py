import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import urllib.parse
import webbrowser
import re
import os
import string
import configparser
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
        self.current_df: Optional[pd.DataFrame] = None
        self.excel_path: Optional[str] = None
        self.base_col_name: Optional[str] = None
        self.sort_state = {}
        self.sorted_col: Optional[str] = None

        # セル編集
        self.edit_entry = None
        self.edit_row = None
        self.edit_col = None

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

        self.version = "v3.23"
        self.root.title(f"AI検索ビューア {self.version}")
        self.root.geometry("1200x650")

        self.setup_style()
        self.setup_menu()
        self.setup_ui()
        self.root.after(100, self.load_once)

    # ---------------------
    # Config
    # ---------------------
    def load_config(self):
        if os.path.exists(self.config_path):
            try:
                self.config.read(self.config_path, encoding="utf-8")
                self.base_col_name = self.config.get("Settings", "base_column", fallback=None)
                self.confirm_rebuild = self.config.getboolean("Settings", "confirm_rebuild", fallback=True)
            except Exception as e:
                logging.error(f"Config error: {e}")

    def save_config(self):
        if "Settings" not in self.config:
            self.config["Settings"] = {}
        self.config["Settings"]["base_column"] = self.base_col_name if self.base_col_name else ""
        self.config["Settings"]["confirm_rebuild"] = "1" if self.confirm_rebuild else "0"
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.config.write(f)
        except Exception as e:
            logging.error(f"Config save error: {e}")

    # ---------------------
    # Menu（使い方 / 操作履歴）
    # ---------------------
    def setup_menu(self):
        menubar = tk.Menu(self.root)

        view = tk.Menu(menubar, tearoff=0)
        view.add_command(label="使い方", command=self.show_help_window)
        view.add_command(label="操作履歴", command=self.show_history_window)
        menubar.add_cascade(label="表示", menu=view)

        self.root.config(menu=menubar)

    def show_help_window(self):
        win = tk.Toplevel(self.root)
        win.title("使い方（かんたん）")
        win.geometry("560x320")
        win.grab_set()

        msg = (
            "使い方（基本）\n"
            "1. 編集したいセルをダブルクリックして入力\n"
            "2. 検索元の列を選ぶ（最初だけ）\n"
            "3. 「検索語句更新」でリンク列（AI/Google）を作成\n"
            "4. AI/Google列をダブルクリックで検索を開く\n"
            "5. 保存する（別名保存がおすすめ）\n\n"
            "ポイント\n"
            "・編集は通常の列で行います（AI/Google列はリンク専用で編集不可）\n"
            "・ヘッダー：シングル=並び替え / ダブル=列名変更 / 右クリック=列一括編集\n"
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

        for s in reversed(self.op_history[-200:]):
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
        if len(self.undo_stack) > 20:
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
    def setup_style(self):
        style = ttk.Style()
        style.theme_use("default")

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

        btn_sel_base = ttk.Button(top, text="検索語句列の変更", command=self.select_base_column)
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

        btn_save_as = ttk.Button(top, text="名前を付けて保存", command=self.save_as_new)
        btn_save_as.pack(side="right", padx=5)
        ToolTip(btn_save_as, "別名で保存します（元ファイルは残ります）。")

        self.btn_save_open = ttk.Button(top, text="上書き保存して開く", command=self.save_and_open_choice, style="Red.TButton")
        self.btn_save_open.pack(side="right", padx=5)
        ToolTip(self.btn_save_open,
                "元ファイルに上書き保存します。\n"
                "はい：Excelで開きます。\n"
                "いいえ：CSVを保存し、フォルダを開きます。")

        btn_open = ttk.Button(top, text="別のファイルを開く", command=self.open_new_file)
        btn_open.pack(side="right", padx=5)
        ToolTip(btn_open, "別のExcelファイルを開きます（未保存がある場合は確認します）。")

        self.btn_reload = ttk.Button(top, text="再読み込み", command=self.reload_original)
        self.btn_reload.pack(side="right", padx=5)
        ToolTip(self.btn_reload, "元ファイルを読み込み直します（変更は破棄されます）。")

        self.version_label = tk.Label(top, text=self.version, fg="gray")
        self.version_label.pack(side="right", padx=5)

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
        self.status_mid.config(text=f"検索語句列: {self.base_col_name or '-'}")
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
            self.current_df = pd.read_excel(path, dtype=str, engine="openpyxl")
            logging.info(f"Loaded: {path}")
        except Exception as e:
            messagebox.showerror("エラー", f"読み込み失敗: {e}")
            self.current_df = None

    def _reset_for_new_file(self):
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.set_unsaved(False)
        self.sorted_col = None
        self._onboard_shown = False
        self.op_history.clear()
        self.update_undo_redo_buttons()

    def load_once(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not self.excel_path:
            self.root.destroy()
            return
        self.load_excel(self.excel_path)
        self._reset_for_new_file()
        if self.current_df is not None:
            if self.base_col_name not in self.current_df.columns:
                self.select_base_column()
            else:
                self.rebuild_search_columns()
        self.update_status_bar()

    # ---------------------
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
    def select_base_column(self):
        if self.current_df is None or len(self.current_df.columns) == 0:
            return

        win = tk.Toplevel(self.root)
        win.title("検索語句の列を選択")
        win.grab_set()

        tk.Label(win, text="AI/Google検索のキーワードとする列を選択してください:").pack(padx=20, pady=10)
        box = ttk.Combobox(win, values=list(self.current_df.columns), state="readonly")

        if self.base_col_name in self.current_df.columns:
            box.set(self.base_col_name)
        else:
            box.set(self.current_df.columns[0])

        box.pack(padx=20, pady=10)

        def decide():
            col = box.get()
            if not col:
                return
            self.base_col_name = col
            self.save_config()
            win.destroy()
            self.rebuild_search_columns()

        ttk.Button(win, text="適用", command=decide).pack(pady=10)

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
        if (not self.base_col_name) or (self.base_col_name not in self.current_df.columns):
            self.select_base_column()
            return

        if self.confirm_rebuild:
            if not self.confirm_rebuild_dialog():
                return

        self.finish_edit(None)
        before = self.current_df.copy()

        df = self.current_df.copy()
        df["AI検索"] = df[self.base_col_name].apply(ai_url)
        df["Google検索"] = df[self.base_col_name].apply(google_url)

        cols = [c for c in df.columns if c not in ["AI検索", "Google検索"]]
        cols.insert(1, "AI検索")
        cols.insert(2, "Google検索")
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
    def show_dataframe(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for i, c in enumerate(df.columns):
            hdr = f"{get_excel_header(i + 1)} {c}"
            if self.sorted_col == c:
                hdr = ("▲ " if self.sort_state.get(c, True) else "▼ ") + hdr
            self.tree.heading(c, text=hdr)
            self.tree.column(c, width=150, anchor="w")

        for i, row in enumerate(df.itertuples(index=False)):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=[display_text(v) for v in row], tags=(tag,))

        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd", background="#f9f9f9")

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
        if r < 0 or r >= len(self.current_df):
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
        self.edit_entry.insert(0, display_text(self.current_df.iat[r, c]))
        self.edit_entry.focus()

        self.edit_row, self.edit_col = r, c
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
        old = display_text(self.current_df.iat[self.edit_row, self.edit_col])

        if val != old:
            before = self.current_df.copy()
            after = self.current_df.copy()
            after.iat[self.edit_row, self.edit_col] = val
            self.commit_df(before, after, f"セル編集: R{self.edit_row+2}C{self.edit_col+1}", refresh_view=True)

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
        ent.insert(0, '例：=A{ROW}&"_"&B{ROW}')

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
            self.excel_path = path
            self.load_excel(path)
            self._reset_for_new_file()
            if self.current_df is None:
                return
            if self.base_col_name not in self.current_df.columns:
                self.select_base_column()
            else:
                self.rebuild_search_columns()
            self.update_status_bar()

    def reload_original(self):
        if self.excel_path and messagebox.askyesno("再読み込み", "元ファイルを再読み込みします。\n（変更は破棄されます）"):
            self.load_excel(self.excel_path)
            self._reset_for_new_file()
            if self.current_df is None:
                return
            self.rebuild_search_columns()
            self.set_unsaved(False)
            self.update_status_bar()

    def save_current_file(self):
        if self.current_df is None or not self.excel_path:
            return False
        try:
            self.current_df.to_excel(self.excel_path, index=False)
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
                self.current_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
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
                self.current_df.to_excel(path, index=False)
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
