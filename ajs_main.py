#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AST (AJS Support Tool) - Main GUI
v6.0 (2026-04-29) - UI刷新: クラス化, ダークテーマ, 2カテゴリ×サブタブ

機能:
  [AJS操作]
    - 定義退避・変換 (Print & Convert)
    - 定義回復 (Recover)
  [ユニット解析]
    - 入出力解析 (In/Out Analysis)
    - 逆引き解析 (Dependency Trace)
"""

import json
import threading
import pathlib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import paramiko
import platform

from ajs_constants import *
from ajs_print_logic import print_start_job
from ajs_define_logic import define_start_job
from ajs_inout_logic import inout_start_job
from ajs_rel_logic import pre_start_job
from ajs_depend_logic import open_t5_job_runner
from ajs_exception_editor import open_editor_window
from ajs_bridge_editor import open_bridge_editor


# =========================================================================
# カラーテーマ (ダーク)
# =========================================================================
# Draculaベースの配色。純黒(#1e1e2e)よりやや明るい#282a36を背景にすることで
# 長時間使っても目が疲れにくく、テキストとのコントラストも十分に確保できる。
_COLORS = {
    "bg":           "#282a36",       # 全体背景 (Dracula Background)
    "fg":           "#f0f0f5",       # 通常テキスト (ほぼ白、高コントラスト)
    "fg_dim":       "#6272a4",       # 薄い補助テキスト (Dracula Comment)
    "accent":       "#7aa2f7",       # アクセント (Tokyo Night Blue)
    "accent_hover": "#89b4fa",       # アクセントホバー
    "button_bg":    "#44475a",       # ボタン背景 (Dracula Selection)
    "button_fg":    "#f0f0f5",       # ボタンテキスト
    "entry_bg":     "#21222c",       # 入力欄背景 (bgより少し暗い)
    "entry_fg":     "#f0f0f5",       # 入力欄テキスト
    "entry_border": "#6272a4",       # 入力欄の枠線 (背景に溶けないよう明るめ)
    "border":       "#44475a",       # ボタン等の枠線
    "success":      "#a6e3a1",       # 成功 (緑)
    "error":        "#ff6e6e",       # エラー (赤)
    "warning":      "#f9e2af",       # 警告 (黄)
    "progress_bg":  "#44475a",       # プログレスバーのトラフ
}

PAD = 8            # 標準パディング
LABEL_W = 14       # ラベル統一幅(文字数)


# =========================================================================
# Tooltip
# =========================================================================
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tooltip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify='left', background="#FFFDE7",
                 foreground="#282a36", relief='solid', borderwidth=1,
                 wraplength=400, padx=8, pady=4).pack()

    def hide(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None


# =========================================================================
# FileListEditor (目標ユニット入力用のリストエディタ)
# =========================================================================
class FileListEditor(ttk.Frame):
    """行を動的に追加・削除できるリストエディタ。
    逆引き解析で目標ユニットのパスを複数入力するために使う。
    """
    def __init__(self, parent):
        super().__init__(parent)

        self.canvas_frm = ttk.Frame(self, style="Main.TFrame")
        self.canvas_frm.pack(fill="both", expand=True, pady=(0, 5))

        self.canvas = tk.Canvas(self.canvas_frm, borderwidth=0,
                                background=_COLORS["entry_bg"], height=60,
                                highlightthickness=0)
        self.scroll_frame = ttk.Frame(self.canvas, style="Main.TFrame")
        self.vsb = ttk.Scrollbar(self.canvas_frm, orient="vertical",
                                 command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas_window_id = self.canvas.create_window(
            (0, 0), window=self.scroll_frame, anchor="nw")

        self.scroll_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.rows = []

        btn_frame = ttk.Frame(self, style="Main.TFrame")
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="+ 行を追加", font=("", 9),
                  bg=_COLORS["button_bg"], fg=_COLORS["button_fg"],
                  activebackground=_COLORS["border"],
                  activeforeground=_COLORS["button_fg"],
                  relief="flat", padx=8, pady=2, cursor="hand2",
                  command=lambda: self.add_row("")).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="全クリア", font=("", 9),
                  bg=_COLORS["button_bg"], fg=_COLORS["button_fg"],
                  activebackground=_COLORS["border"],
                  activeforeground=_COLORS["button_fg"],
                  relief="flat", padx=8, pady=2, cursor="hand2",
                  command=self.clear_all).pack(side="right")

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window_id, width=event.width)

    def add_row(self, text_value):
        row_f = ttk.Frame(self.scroll_frame, style="Main.TFrame")
        row_f.pack(fill="x", pady=2, padx=2)

        entry = tk.Entry(row_f, bg=_COLORS["entry_bg"], fg=_COLORS["entry_fg"],
                         insertbackground=_COLORS["fg"], relief="flat", bd=0,
                         highlightthickness=1, highlightcolor=_COLORS["accent"],
                         highlightbackground=_COLORS["entry_border"])
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        entry.insert(0, text_value)

        del_btn = tk.Button(row_f, text="x", width=3, font=("", 8),
                            bg=_COLORS["button_bg"], fg=_COLORS["error"],
                            activebackground=_COLORS["border"],
                            relief="flat", cursor="hand2",
                            command=lambda f=row_f: self.remove_row(f))
        del_btn.pack(side="left")

        self.rows.append((row_f, entry))

    def remove_row(self, frame):
        for i, (f, e) in enumerate(self.rows):
            if f == frame:
                f.destroy()
                self.rows.pop(i)
                break

    def clear_all(self):
        for f, e in self.rows:
            f.destroy()
        self.rows = []

    def get_values(self):
        return [e.get().strip() for f, e in self.rows if e.get().strip()]

    def set_values(self, values_list):
        self.clear_all()
        for v in values_list:
            if v:
                self.add_row(v)


# =========================================================================
# メインアプリケーション
# =========================================================================
class ASTApp(tk.Tk):
    """AST (AJS Support Tool) のメインウィンドウ。"""

    def __init__(self):
        super().__init__()
        self.title("AST v6.0")
        self.geometry("860x800")
        self.minsize(780, 600)
        self.configure(bg=_COLORS["bg"])

        self._setup_styles()
        self._init_variables()
        self._build_ui()

    # -----------------------------------------------------------------------
    # ttkスタイル設定
    # -----------------------------------------------------------------------
    def _setup_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass

        # 共通フォントサイズ: デフォルトの9ptから10ptに上げて視認性向上
        _font = ("", 10)
        _font_bold = ("", 10, "bold")
        _font_sm = ("", 9)

        style.configure("Main.TFrame", background=_COLORS["bg"])

        style.configure("Main.TLabel", background=_COLORS["bg"],
                        foreground=_COLORS["fg"], font=_font)
        style.configure("Dim.TLabel", background=_COLORS["bg"],
                        foreground=_COLORS["fg_dim"], font=_font_sm)

        style.configure("Main.TCheckbutton", background=_COLORS["bg"],
                        foreground=_COLORS["fg"], font=_font)
        style.configure("Main.TRadiobutton", background=_COLORS["bg"],
                        foreground=_COLORS["fg"], font=_font)

        # セクション (LabelFrame)
        style.configure("Section.TLabelframe", background=_COLORS["bg"],
                        foreground=_COLORS["accent"])
        style.configure("Section.TLabelframe.Label", background=_COLORS["bg"],
                        foreground=_COLORS["accent"], font=_font_bold)

        # プログレスバー
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=_COLORS["progress_bg"],
                        background=_COLORS["accent"])

        # Notebook (大カテゴリタブ)
        style.configure("TNotebook", background=_COLORS["bg"])
        style.configure("TNotebook.Tab", background=_COLORS["button_bg"],
                        foreground=_COLORS["fg"], padding=[14, 6],
                        font=_font_bold)
        style.map("TNotebook.Tab",
                  background=[("selected", _COLORS["accent"])],
                  foreground=[("selected", "#282a36")])

        # サブタブ用Notebook
        # 選択中タブにアクセント色の下線が見えるよう、背景を明るめにして差をつける
        style.configure("Sub.TNotebook", background=_COLORS["bg"])
        style.configure("Sub.TNotebook.Tab", background=_COLORS["button_bg"],
                        foreground=_COLORS["fg_dim"], padding=[10, 4],
                        font=_font_bold)
        style.map("Sub.TNotebook.Tab",
                  background=[("selected", _COLORS["accent"])],
                  foreground=[("selected", "#282a36")])

        # Combobox (ダークテーマ対応)
        style.configure("TCombobox", font=_font)
        style.map("TCombobox",
                  fieldbackground=[("readonly", _COLORS["entry_bg"])],
                  foreground=[("readonly", _COLORS["fg"])])

    # -----------------------------------------------------------------------
    # GUI変数の初期化
    # -----------------------------------------------------------------------
    def _init_variables(self):
        # 共通接続情報
        self.v_ip = tk.StringVar(self)
        self.v_user = tk.StringVar(self)
        self.v_pass = tk.StringVar(self)
        self.v_srv_c = tk.StringVar(self, 'SJIS')

        # Tab1: 定義退避・変換
        self.v_print_ajs_path = tk.StringVar(self)
        self.v_print_kind = tk.StringVar(self, "verify")
        self.v_print_conv_flg = tk.StringVar(self, "no")
        self.v_print_bank = tk.StringVar(self)
        self.v_print_detail = tk.StringVar(self)
        self.v_print_out_c = tk.StringVar(self, 'SJIS(CP932)')
        self.v_print_out_n = tk.StringVar(self, 'CRLF(Windows)')
        self.print_custom_pairs = []

        # Tab2: 定義回復
        self.v_recover_file = tk.StringVar(self)
        self.v_recover_unit = tk.StringVar(self, "")

        # Tab3: 入出力解析
        self.v_inout_ajs = tk.StringVar(self)
        self.v_inout_res = tk.StringVar(self)
        self.v_inout_bank = tk.StringVar(self)
        self.v_inout_format = tk.StringVar(self, "Excel")
        self.inout_custom_vars = []

        # 逆引き解析 (旧Tab5, Tab4 GUIは廃止)
        self.v_dep_ajs = tk.StringVar(self)
        self.v_dep_res = tk.StringVar(self)
        self.v_dep_tgt_files = tk.StringVar(self)
        self.v_dep_bank = tk.StringVar(self)
        self.v_dep_out_c = tk.StringVar(self, 'SJIS(CP932)')
        self.v_dep_out_n = tk.StringVar(self, 'CRLF(Windows)')
        self.dep_custom_vars = []

        # コマンド・環境変数
        self.v_ajs_print_path = tk.StringVar(self, AJS_PRINT_PATH)
        self.v_ajs_define_path = tk.StringVar(self, AJS_DEFINE_PATH)
        self.v_jp1_hostname = tk.StringVar(self, DEFAULT_JP1_HOSTNAME)
        self.v_jp1_username = tk.StringVar(self, DEFAULT_JP1_USERNAME)

        # ステータス
        self.status_var = tk.StringVar(self, '準備完了')
        self.progress = tk.DoubleVar(self, 0.0)

        # 履歴
        self.hist = self._load_hist()

        # 実行ボタン参照用
        self.run_buttons = {}

    # -----------------------------------------------------------------------
    # UI構築
    # -----------------------------------------------------------------------
    def _build_ui(self):
        main_frm = ttk.Frame(self, style="Main.TFrame", padding=PAD)
        main_frm.pack(fill="both", expand=True)

        # --- 底から先にpackして確実にスペースを確保する ---
        # packは「先に呼んだものが先に場所を取る」ルール。
        # side="bottom"で底からステータスバー→ログの順に積み上げてから、
        # 残ったスペースをNotebookに渡す。こうす��とウィンドウが小さくても
        # ログエリアが潰れない。
        self._build_status_bar(main_frm)
        self._build_log_area(main_frm)

        # --- 上から ---
        self._build_connection_section(main_frm)

        # --- 残りのスペースをNotebookが使う ---
        self.notebook = ttk.Notebook(main_frm)
        self.notebook.pack(fill="both", expand=True, pady=(PAD, 0))

        # [AJS操作] カテゴリ
        tab_ops = ttk.Frame(self.notebook, style="Main.TFrame")
        self.notebook.add(tab_ops, text="  AJS操作  ")
        self.sub_nb_ops = ttk.Notebook(tab_ops, style="Sub.TNotebook")
        self.sub_nb_ops.pack(fill="both", expand=True, padx=4, pady=4)

        tab_print = ttk.Frame(self.sub_nb_ops, style="Main.TFrame", padding=PAD)
        self.sub_nb_ops.add(tab_print, text=" 定義退避・変換 ")
        self._build_tab_print(tab_print)

        tab_recover = ttk.Frame(self.sub_nb_ops, style="Main.TFrame", padding=PAD)
        self.sub_nb_ops.add(tab_recover, text=" 定義回復 ")
        self._build_tab_recover(tab_recover)

        # [ユニット解析] カテゴリ
        tab_analysis = ttk.Frame(self.notebook, style="Main.TFrame")
        self.notebook.add(tab_analysis, text="  ユニット解析  ")
        self.sub_nb_analysis = ttk.Notebook(tab_analysis, style="Sub.TNotebook")
        self.sub_nb_analysis.pack(fill="both", expand=True, padx=4, pady=4)

        tab_inout = ttk.Frame(self.sub_nb_analysis, style="Main.TFrame", padding=PAD)
        self.sub_nb_analysis.add(tab_inout, text=" 入出力解析 ")
        self._build_tab_inout(tab_inout)

        tab_dep = ttk.Frame(self.sub_nb_analysis, style="Main.TFrame", padding=PAD)
        self.sub_nb_analysis.add(tab_dep, text=" 逆引き解析 ")
        self._build_tab_dep(tab_dep)

    # -----------------------------------------------------------------------
    # 共通接続情報セクション
    # -----------------------------------------------------------------------
    def _build_connection_section(self, parent):
        sec = ttk.LabelFrame(parent, text=" 共通接続情報 ",
                             style="Section.TLabelframe", padding=PAD)
        sec.pack(fill="x", pady=(0, 4))

        row = ttk.Frame(sec, style="Main.TFrame")
        row.pack(fill="x")

        self._label(row, "IP:", width=4).pack(side="left")
        self._combobox(row, self.v_ip,
                       self.hist.get('ip', []), width=16).pack(
            side="left", padx=(0, 8))

        self._label(row, "User:", width=5).pack(side="left")
        self._combobox(row, self.v_user,
                       self.hist.get('user', []), width=12).pack(
            side="left", padx=(0, 8))

        self._label(row, "Pass:", width=5).pack(side="left")
        self._entry(row, self.v_pass, show="*", width=12).pack(
            side="left", padx=(0, 8))

        self._label(row, "Enc:", width=4).pack(side="left")
        ttk.Combobox(row, textvariable=self.v_srv_c,
                     values=['SJIS', 'UTF-8'], width=6,
                     state='readonly').pack(side="left", padx=(0, 8))

        self._btn(row, "詳細...", self._open_advanced_settings).pack(side="left")

    # -----------------------------------------------------------------------
    # 詳細設定ウィンドウ
    # -----------------------------------------------------------------------
    def _open_advanced_settings(self):
        adv_win = tk.Toplevel(self)
        adv_win.title("詳細設定")
        adv_win.geometry("500x220")
        adv_win.transient(self)
        adv_win.grab_set()
        adv_win.configure(bg=_COLORS["bg"])

        frm = ttk.LabelFrame(adv_win, text=" コマンド・環境変数 ",
                             style="Section.TLabelframe", padding=12)
        frm.pack(fill="both", expand=True, padx=12, pady=8)
        frm.columnconfigure(1, weight=1)

        for i, (lbl, var) in enumerate([
            ("ajsprint パス:", self.v_ajs_print_path),
            ("ajsdefine パス:", self.v_ajs_define_path),
        ]):
            self._label(frm, lbl).grid(row=i, column=0, sticky="e", padx=5, pady=5)
            self._entry(frm, var).grid(row=i, column=1, sticky="ew", padx=5, pady=5)

        ttk.Separator(frm, orient="horizontal").grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=8)

        for i, (lbl, var) in enumerate([
            ("JP1_HOSTNAME:", self.v_jp1_hostname),
            ("JP1_USERNAME:", self.v_jp1_username),
        ], start=3):
            self._label(frm, lbl).grid(row=i, column=0, sticky="e", padx=5, pady=5)
            self._entry(frm, var).grid(row=i, column=1, sticky="ew", padx=5, pady=5)

        btn_frm = ttk.Frame(frm, style="Main.TFrame")
        btn_frm.grid(row=5, column=0, columnspan=2, pady=8)

        def restore():
            self.v_ajs_print_path.set(AJS_PRINT_PATH)
            self.v_ajs_define_path.set(AJS_DEFINE_PATH)
            self.v_jp1_hostname.set(DEFAULT_JP1_HOSTNAME)
            self.v_jp1_username.set(DEFAULT_JP1_USERNAME)

        self._btn(btn_frm, "デフォルトに戻す", restore).pack(side="left", padx=8)
        self._btn(btn_frm, "閉じる", adv_win.destroy).pack(side="left", padx=8)

    # -----------------------------------------------------------------------
    # Tab: 定義退避・変換
    # -----------------------------------------------------------------------
    def _build_tab_print(self, parent):
        # --- ボタンを先にpack(side=bottom)して底に固定 ---
        self.run_buttons['print'] = self._accent_btn(
            parent, "  定義退避・変換 実行  ", self._run_print)
        self.run_buttons['print'].pack(side="bottom", pady=PAD)

        # 出力定義設定も底側に
        self._build_output_selector(parent, self.v_print_out_c,
                                    self.v_print_out_n).pack(
            side="bottom", fill="x", pady=5)

        # --- 設定セクション (上から) ---
        sec = ttk.LabelFrame(parent, text=" 実行設定 ",
                             style="Section.TLabelframe", padding=PAD)
        sec.pack(fill="x", pady=(0, PAD))
        sec.columnconfigure(1, weight=1)

        # AJSパス
        self._label(sec, "AJS パス:").grid(row=0, column=0, sticky="e", pady=6)
        self._combobox(sec, self.v_print_ajs_path,
                       self.hist.get('print_ajs_path', [])).grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=5)

        # 取得する定義
        self._label(sec, "取得する定義:").grid(row=1, column=0, sticky="e", pady=6)
        f_def = ttk.Frame(sec, style="Main.TFrame")
        f_def.grid(row=1, column=1, columnspan=2, sticky="w")
        for t, v in [("回復用", "recover"), ("確認用", "verify"), ("両方", "both")]:
            ttk.Radiobutton(f_def, text=t, variable=self.v_print_kind, value=v,
                            style="Main.TRadiobutton").pack(side="left", padx=5)

        # 変換を行うか
        self._label(sec, "変換を行うか:").grid(row=2, column=0, sticky="e", pady=6)
        f_conv = ttk.Frame(sec, style="Main.TFrame")
        f_conv.grid(row=2, column=1, columnspan=2, sticky="w")
        for t, v in [("はい", "yes"), ("いいえ", "no")]:
            ttk.Radiobutton(f_conv, text=t, variable=self.v_print_conv_flg, value=v,
                            style="Main.TRadiobutton",
                            command=self._toggle_conv_section).pack(side="left", padx=5)

        # === 変換設定セクション (「はい」のときだけ表示) ===
        self._print_sec = sec
        self.conv_section = ttk.LabelFrame(parent, text=" 変換設定 ",
                                           style="Section.TLabelframe", padding=PAD)
        self.conv_section.columnconfigure(1, weight=1)

        # 銀行名
        self._label(self.conv_section, "銀行名:").grid(
            row=0, column=0, sticky="e", pady=3)
        banks_for_conv = [b for b in BANKS if b != "その他"] + ["その他"]
        self.v_print_bank.set(banks_for_conv[0])
        self.cb_print_bank = ttk.Combobox(
            self.conv_section, textvariable=self.v_print_bank,
            values=banks_for_conv, state="readonly", width=20)
        self.cb_print_bank.grid(row=0, column=1, sticky="w", padx=5)
        self.cb_print_bank.bind("<<ComboboxSelected>>", self._on_bank_select)

        # 変換詳細
        self._label(self.conv_section, "変換詳細:").grid(
            row=1, column=0, sticky="ne", pady=3)
        det_frame = ttk.Frame(self.conv_section, style="Main.TFrame")
        det_frame.grid(row=1, column=1, columnspan=2, sticky="w")
        self.det_rbs = []
        self.DETAILS = ["本番⇒ミラー", "本番⇒開発", "ミラー⇒本番",
                        "ミラー⇒開発", "開発⇒本番", "開発⇒ミラー", "カスタム"]
        self.v_print_detail.set(self.DETAILS[0])
        for i, d in enumerate(self.DETAILS):
            rb = ttk.Radiobutton(det_frame, text=d, variable=self.v_print_detail,
                                 value=d, style="Main.TRadiobutton",
                                 command=self._on_detail_select)
            rb.grid(row=i // 3, column=i % 3, sticky="w", padx=4)
            self.det_rbs.append(rb)

        self.custom_detail_btn = self._btn(
            det_frame, "詳細...",
            lambda: self._open_key_value_window(
                "カスタム変換設定", self.print_custom_pairs,
                "変換元文言", "変換先文言"))
        self.custom_detail_btn.config(state="disabled")
        self.custom_detail_btn.grid(
            row=len(self.DETAILS) // 3, column=len(self.DETAILS) % 3,
            sticky="w", padx=10, pady=5)

        self._toggle_conv_section()

    # -----------------------------------------------------------------------
    # Tab: 定義回復
    # -----------------------------------------------------------------------
    def _build_tab_recover(self, parent):
        # --- ボタンを底に固定 ---
        self.run_buttons['recover'] = self._accent_btn(
            parent, "  定義回復 実行  ", self._run_recover)
        self.run_buttons['recover'].pack(side="bottom", pady=PAD)

        # --- 設定セクション (項目が少ないのでpyを大きめに取って余裕を持たせる) ---
        sec = ttk.LabelFrame(parent, text=" 実行設定 ",
                             style="Section.TLabelframe", padding=(PAD, PAD * 2))
        sec.pack(fill="x")
        sec.columnconfigure(1, weight=1)

        self._label(sec, "回復用AJS定義:").grid(row=0, column=0, sticky="e", pady=8)
        self._entry(sec, self.v_recover_file).grid(
            row=0, column=1, sticky="ew", padx=5)
        self._btn(sec, "参照...",
                  lambda: self.v_recover_file.set(
                      filedialog.askopenfilename())).grid(row=0, column=2)

        self._label(sec, "回復先AJSパス:").grid(row=1, column=0, sticky="e", pady=8)
        self._combobox(sec, self.v_recover_unit,
                       self.hist.get('recover_unit_name', [])).grid(
            row=1, column=1, columnspan=2, sticky="ew", padx=5)

    # -----------------------------------------------------------------------
    # 銀行選択ウィジェット生成 (入出力解析・逆引き解析で共通)
    # -----------------------------------------------------------------------
    def _build_bank_selector(self, parent, bank_var, custom_vars, row_num):
        """銀行名のComboboxと「その他」用の初期変数設定ボタンを配置する。"""
        self._label(parent, "銀行名:").grid(
            row=row_num, column=0, sticky="e", pady=3)
        bank_frame = ttk.Frame(parent, style="Main.TFrame")
        bank_frame.grid(row=row_num, column=1, columnspan=2, sticky="w")

        banks_list = [b for b in BANKS if b != "その他"] + ["その他"]
        bank_var.set(banks_list[0])
        cb = ttk.Combobox(bank_frame, textvariable=bank_var,
                          values=banks_list, state="readonly", width=20)
        cb.pack(side="left", padx=(5, 8))

        custom_btn = self._btn(
            bank_frame, "初期変数設定",
            lambda: self._open_key_value_window(
                "「その他」用 初期変数設定", custom_vars,
                "変数名 (例: BSDIR)", "値 (例: /HN)"))
        custom_btn.pack(side="left")
        custom_btn.pack_forget()  # 初期は非表示

        # 「その他」選択時のみボタンを表示
        def _on_bank_change(event=None):
            if bank_var.get() == "その他":
                custom_btn.pack(side="left")
            else:
                custom_btn.pack_forget()

        cb.bind("<<ComboboxSelected>>", _on_bank_change)
        return cb, custom_btn

    # -----------------------------------------------------------------------
    # Tab: 入出力解析
    # -----------------------------------------------------------------------
    def _build_tab_inout(self, parent):
        # --- ボタン行を底に固定 (逆引き解析と同じ構成) ---
        btn_frame = ttk.Frame(parent, style="Main.TFrame")
        btn_frame.pack(side="bottom", pady=PAD)
        self.run_buttons['inout'] = self._accent_btn(
            btn_frame, "  入出力解析 実行  ", self._run_inout)
        self.run_buttons['inout'].pack(side="left", padx=5)
        self._btn(btn_frame, "I/Oルール編集",
                  lambda: open_editor_window(
                      self, [b for b in BANKS if b != "その他"] + ["*"]
                  )).pack(side="left", padx=5)

        # --- 解析結果 (ボタンの上、expand=Trueで余白を吸収) ---
        res_sec = ttk.LabelFrame(parent, text=" 解析結果 ",
                                 style="Section.TLabelframe", padding=PAD)
        res_sec.pack(side="bottom", fill="both", expand=True, pady=(PAD, 0))
        self.t3_text_box = self._create_text_widget(res_sec, height=6)

        # --- 設定セクション (上から) ---
        sec = ttk.LabelFrame(parent, text=" 実行設定 ",
                             style="Section.TLabelframe", padding=PAD)
        sec.pack(fill="x")
        sec.columnconfigure(1, weight=1)

        # AJSパス
        self._label(sec, "AJSパス:").grid(row=0, column=0, sticky="e", pady=3)
        self._combobox(sec, self.v_inout_ajs,
                       self.hist.get('inout_ajs_path', [])).grid(
            row=0, column=1, sticky="ew", columnspan=2, padx=5)

        # リソースパス
        self._label(sec, "リソースパス:").grid(row=1, column=0, sticky="e", pady=3)
        self._combobox(sec, self.v_inout_res,
                       self.hist.get('inout_res_path', [])).grid(
            row=1, column=1, sticky="ew", padx=5)
        self._btn(sec, "参照...",
                  lambda: self.v_inout_res.set(
                      filedialog.askdirectory())).grid(row=1, column=2)

        # 銀行名
        self._build_bank_selector(sec, self.v_inout_bank,
                                  self.inout_custom_vars, row_num=2)

        # 出力形式
        self._label(sec, "出力形式:").grid(row=3, column=0, sticky="e", pady=3)
        t3_out_frame = ttk.Frame(sec, style="Main.TFrame")
        t3_out_frame.grid(row=3, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(t3_out_frame, text="Excel", variable=self.v_inout_format,
                        value="Excel", style="Main.TRadiobutton").pack(
            side="left", padx=5)
        ttk.Radiobutton(t3_out_frame, text="CSV", variable=self.v_inout_format,
                        value="CSV", style="Main.TRadiobutton").pack(
            side="left", padx=5)

    # -----------------------------------------------------------------------
    # Tab: 逆引き解析
    # -----------------------------------------------------------------------
    def _build_tab_dep(self, parent):
        # --- 底から順にpack: ボタン→出力設定 ---
        btn_frame = ttk.Frame(parent, style="Main.TFrame")
        btn_frame.pack(side="bottom", pady=PAD)
        self.run_buttons['dep'] = self._accent_btn(
            btn_frame, "  逆引き解析 実行  ", self._run_dep)
        self.run_buttons['dep'].pack(side="left", padx=5)
        self._btn(btn_frame, "DBブリッジ編集",
                  lambda: open_bridge_editor(
                      self, [b for b in BANKS if b != "その他"] + ["*"]
                  )).pack(side="left", padx=5)

        self._build_output_selector(parent, self.v_dep_out_c,
                                    self.v_dep_out_n).pack(
            side="bottom", fill="x", pady=5)

        # --- 上から順にpack: 設定→目標ユニット ---
        sec = ttk.LabelFrame(parent, text=" 実行設定 ",
                             style="Section.TLabelframe", padding=PAD)
        sec.pack(fill="x")
        sec.columnconfigure(1, weight=1)

        self._label(sec, "AJSパス:").grid(row=0, column=0, sticky="e", pady=3)
        self._combobox(sec, self.v_dep_ajs,
                       self.hist.get('dep_ajs_path', [])).grid(
            row=0, column=1, sticky="ew", padx=5)

        self._label(sec, "リソースパス:").grid(row=1, column=0, sticky="e", pady=3)
        self._combobox(sec, self.v_dep_res,
                       self.hist.get('dep_res_path', [])).grid(
            row=1, column=1, sticky="ew", padx=5)
        self._btn(sec, "参照...",
                  lambda: self.v_dep_res.set(
                      filedialog.askdirectory())).grid(row=1, column=2)

        self._build_bank_selector(sec, self.v_dep_bank,
                                  self.dep_custom_vars, row_num=2)

        # 目標ユニット (secの外に独立セクションとして配置)
        tgt_sec = ttk.LabelFrame(parent, text=" 目標ユニット ",
                                 style="Section.TLabelframe", padding=PAD)
        tgt_sec.pack(fill="x", pady=(PAD, 0))
        self.t5_list_editor = FileListEditor(tgt_sec)
        self.t5_list_editor.pack(fill="x")

        # --- 残りのスペースを解析結果が使う (最後にpack) ---
        res_sec = ttk.LabelFrame(
            parent, text=" 解析結果 ",
            style="Section.TLabelframe", padding=PAD)
        res_sec.pack(fill="both", expand=True, pady=(PAD, 0))
        self.t5_text_box = self._create_text_widget(res_sec, height=6)

    # -----------------------------------------------------------------------
    # ログエリア
    # -----------------------------------------------------------------------
    def _build_log_area(self, parent):
        log_frame = ttk.Frame(parent, style="Main.TFrame")
        log_frame.pack(side="bottom", fill="x", pady=(PAD, 0))

        self.log_text = tk.Text(
            log_frame, height=3,
            bg=_COLORS["entry_bg"], fg=_COLORS["fg"],
            insertbackground=_COLORS["fg"],
            relief='flat', font=("Consolas", 9),
            state='disabled', padx=6, pady=4,
            highlightthickness=1,
            highlightcolor=_COLORS["accent"],
            highlightbackground=_COLORS["entry_border"],
        )
        log_scroll = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side='right', fill='y')
        self.log_text.pack(side='left', fill='both', expand=True)

    # -----------------------------------------------------------------------
    # ステータスバー
    # -----------------------------------------------------------------------
    def _build_status_bar(self, parent):
        status_frm = ttk.Frame(parent, style="Main.TFrame")
        status_frm.pack(side="bottom", fill="x", pady=(4, 0))

        ttk.Progressbar(
            status_frm, variable=self.progress, maximum=100,
            length=300, mode='determinate',
            style="Custom.Horizontal.TProgressbar"
        ).pack(fill="x", side="left", expand=True, padx=(0, 8))

        ttk.Label(status_frm, textvariable=self.status_var,
                  style="Dim.TLabel", anchor="e").pack(side="right")

    # -----------------------------------------------------------------------
    # 出力定義設定 (定義退避・変換, 逆引き解析で共通)
    # -----------------------------------------------------------------------
    def _build_output_selector(self, parent, outc_var, outn_var):
        frm = ttk.LabelFrame(parent, text=" 出力定義設定 ",
                             style="Section.TLabelframe", padding=5)
        row = ttk.Frame(frm, style="Main.TFrame")
        row.pack(fill="x")
        self._label(row, "出力文字コード:", width=14).pack(side="left")
        ttk.Combobox(row, textvariable=outc_var,
                     values=['SJIS(CP932)', 'UTF-8'], width=15,
                     state='readonly').pack(side="left", padx=(0, 16))
        self._label(row, "出力改行コード:", width=14).pack(side="left")
        ttk.Combobox(row, textvariable=outn_var,
                     values=['CRLF(Windows)', 'LF(Unix)'], width=15,
                     state='readonly').pack(side="left")
        return frm

    # -----------------------------------------------------------------------
    # UIヘルパーメソッド
    # -----------------------------------------------------------------------
    def _label(self, parent, text, **kw):
        """統一幅のラベルを生成する"""
        return ttk.Label(parent, text=text, style='Main.TLabel',
                         width=kw.get('width', LABEL_W), anchor='w')

    def _btn(self, parent, text, cmd):
        """通常ボタン"""
        return tk.Button(parent, text=text, font=("", 9),
                         bg=_COLORS["button_bg"], fg=_COLORS["button_fg"],
                         activebackground=_COLORS["border"],
                         activeforeground=_COLORS["button_fg"],
                         relief='flat', padx=10, pady=3, cursor='hand2',
                         command=cmd)

    def _accent_btn(self, parent, text, cmd):
        """アクセントカラーの実行ボタン"""
        return tk.Button(
            parent, text=text, font=("", 11, "bold"),
            bg=_COLORS["accent"], fg="#282a36",
            activebackground=_COLORS["accent_hover"],
            activeforeground="#282a36",
            relief="flat", padx=20, pady=6, cursor="hand2",
            command=cmd)

    def _entry(self, parent, textvariable, **kw):
        """ダークテーマ用Entry"""
        return tk.Entry(parent, textvariable=textvariable,
                        bg=_COLORS["entry_bg"], fg=_COLORS["entry_fg"],
                        insertbackground=_COLORS["fg"],
                        relief='flat', bd=0,
                        highlightthickness=1,
                        highlightcolor=_COLORS["accent"],
                        highlightbackground=_COLORS["entry_border"],
                        **kw)

    def _combobox(self, parent, textvariable, values, **kw):
        """履歴付きCombobox"""
        return ttk.Combobox(parent, textvariable=textvariable,
                            values=values, **kw)

    def _create_text_widget(self, parent, height=5):
        """横スクロールバー付き、折り返しなしのテキストボックス"""
        frame = ttk.Frame(parent, style="Main.TFrame")
        frame.pack(fill="both", expand=True)

        v_scroll = tk.Scrollbar(frame)
        h_scroll = tk.Scrollbar(frame, orient="horizontal")

        text_widget = tk.Text(
            frame, height=height, wrap="none", undo=False,
            bg=_COLORS["entry_bg"], fg=_COLORS["fg"],
            insertbackground=_COLORS["fg"],
            selectbackground=_COLORS["accent"],
            selectforeground="#282a36",
            relief='flat',
            highlightthickness=1,
            highlightcolor=_COLORS["accent"],
            highlightbackground=_COLORS["entry_border"])

        text_widget.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        text_widget.config(yscrollcommand=v_scroll.set,
                           xscrollcommand=h_scroll.set)
        v_scroll.config(command=text_widget.yview)
        h_scroll.config(command=text_widget.xview)

        return text_widget

    # -----------------------------------------------------------------------
    # ログ出力
    # -----------------------------------------------------------------------
    def _log(self, msg):
        """ログエリアにメッセージを追加する (スレッドセーフ)"""
        def _append():
            self.log_text.configure(state='normal')
            self.log_text.insert('end', f"> {msg}\n")
            self.log_text.see('end')
            self.log_text.configure(state='disabled')
        self.after(0, _append)

    # -----------------------------------------------------------------------
    # ステータス・通知
    # -----------------------------------------------------------------------
    def update_status(self, msg, p_val=None):
        self.after(0, lambda: self.status_var.set(msg))
        if p_val is not None:
            self.after(0, lambda: self.progress.set(p_val))
        self._log(msg)

    def show_error(self, msg):
        self.after(0, lambda: messagebox.showerror('エラー', msg))
        self.update_status('エラー')

    def show_info(self, msg):
        self.after(0, lambda: messagebox.showinfo('完了', msg))
        self.update_status('完了')

    # -----------------------------------------------------------------------
    # スレッド実行
    # -----------------------------------------------------------------------
    def _check_thread(self, thread, btn_key):
        if thread.is_alive():
            self.after(100, lambda: self._check_thread(thread, btn_key))
        else:
            btn = self.run_buttons.get(btn_key)
            if btn:
                btn.config(state="normal")

    def run_in_thread(self, target_func, btn_key=None):
        """ロジック関数をバックグラウンドスレッドで実行する。"""
        def wrapper(*args, **kwargs):
            btn = self.run_buttons.get(btn_key)
            if btn:
                btn.config(state="disabled")

            thread = threading.Thread(target=target_func, args=args,
                                      kwargs=kwargs, daemon=True)
            thread.start()

            if btn:
                self._check_thread(thread, btn_key)
        return wrapper

    # -----------------------------------------------------------------------
    # SSH接続
    # -----------------------------------------------------------------------
    def get_ssh_client(self):
        ip, user, pw = self.v_ip.get(), self.v_user.get(), self.v_pass.get()
        if not all([ip, user, pw]):
            raise ValueError("接続情報 (IP, ユーザー, パスワード) を入力してください。")

        ssh = paramiko.SSHClient()
        known_hosts = pathlib.Path.home() / ".ssh" / "known_hosts"
        if known_hosts.exists():
            ssh.load_host_keys(str(known_hosts))
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ip, username=user, password=pw, timeout=10)
        ssh.get_host_keys().save(str(known_hosts))
        return ssh

    # -----------------------------------------------------------------------
    # 履歴
    # -----------------------------------------------------------------------
    def _load_hist(self):
        if not HIST_FILE.exists():
            return {}
        try:
            return json.loads(HIST_FILE.read_text(encoding='utf-8'))
        except Exception:
            return {}

    def save_hist(self):
        hist_data = self._load_hist()
        hist_items = {
            'ip': self.v_ip.get(),
            'user': self.v_user.get(),
            'print_ajs_path': self.v_print_ajs_path.get(),
            'recover_unit_name': self.v_recover_unit.get(),
            'inout_ajs_path': self.v_inout_ajs.get(),
            'inout_res_path': self.v_inout_res.get(),
            'dep_ajs_path': self.v_dep_ajs.get(),
            'dep_res_path': self.v_dep_res.get(),
            'dep_tgt_files': self.v_dep_tgt_files.get(),
        }

        for key, value in hist_items.items():
            if not value:
                continue
            new_list = [value] + [x for x in hist_data.get(key, []) if x != value]
            hist_data[key] = new_list[:MAX_HIST]

        HIST_FILE.write_text(
            json.dumps(hist_data, indent=2, ensure_ascii=False), encoding='utf-8')

    # -----------------------------------------------------------------------
    # ロジック呼び出し用の辞書 (既存ロジックファイルとのインターフェース)
    # -----------------------------------------------------------------------
    def _gui_funcs_common(self):
        return {
            'update_status': self.update_status,
            'get_ssh_client': self.get_ssh_client,
            'save_hist': self.save_hist,
            'show_info': self.show_info,
            'show_error': self.show_error,
            'run_in_thread': lambda func: self.run_in_thread(func),
        }

    # -----------------------------------------------------------------------
    # Tab1 変換系コールバック
    # -----------------------------------------------------------------------
    def _on_detail_select(self, *_):
        state = "disabled"
        if self.v_print_detail.get() == "カスタム":
            state = "normal"
        self.custom_detail_btn.config(state=state)

    def _on_bank_select(self, *_):
        bank = self.v_print_bank.get()
        if bank == "その他":
            # 「その他」では変換パターンが不明なのでカスタム固定
            self.v_print_detail.set("カスタム")
            for rb in self.det_rbs:
                st = "disabled" if rb.cget("value") != "カスタム" else "normal"
                rb.config(state=st)
        else:
            for rb in self.det_rbs:
                rb.config(state="normal")
        self._on_detail_select()

    def _toggle_conv_section(self, *_):
        """「変換を行うか」の切り替えで変換設定セクション全体を表示/非表示する。

        従来はdisable(灰色)で対応していたが、ダークテーマでは灰色が見分けにくい。
        関係ない設定は隠す方がUIとして明快。
        """
        if self.v_print_conv_flg.get() == "yes":
            # 実行設定セクション(sec)の直後に変換設定を挿入
            self.conv_section.pack(fill="x", pady=(PAD, 0),
                                   after=self._print_sec)
            self._on_bank_select()
        else:
            self.conv_section.pack_forget()

    # -----------------------------------------------------------------------
    # キー・バリュー編集ウィンドウ
    # -----------------------------------------------------------------------
    def _open_key_value_window(self, title, data_list, key_label, value_label):
        current_data = list(data_list)
        cust_win = tk.Toplevel(self)
        cust_win.title(title)
        cust_win.geometry("500x400")
        cust_win.transient(self)
        cust_win.grab_set()
        cust_win.configure(bg=_COLORS["bg"])

        hdr_frm = ttk.Frame(cust_win, style="Main.TFrame", padding=(10, 10, 10, 0))
        hdr_frm.pack(fill="x")
        ttk.Label(hdr_frm, text=key_label, width=30,
                  style="Main.TLabel",
                  font=("", 10, "bold")).pack(side="left", padx=5)
        ttk.Label(hdr_frm, text=value_label, width=30,
                  style="Main.TLabel",
                  font=("", 10, "bold")).pack(side="left", padx=5)

        canvas_frm = ttk.Frame(cust_win, style="Main.TFrame", padding=5)
        canvas_frm.pack(fill="both", expand=True)
        canvas = tk.Canvas(canvas_frm, borderwidth=0,
                           background=_COLORS["entry_bg"], highlightthickness=0)
        scroll_frame = ttk.Frame(canvas, style="Main.TFrame", padding=(10, 0))
        vsb = tk.Scrollbar(canvas_frm, command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def _conf(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        scroll_frame.bind("<Configure>", _conf)

        def _wheel(e):
            if e.delta:
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _wheel)
        cust_win.bind("<Destroy>",
                      lambda e: canvas.unbind_all("<MouseWheel>"))

        all_rows = []

        def add_pair_row(key_text="", val_text=""):
            row_f = ttk.Frame(scroll_frame, style="Main.TFrame")
            row_f.pack(fill="x", pady=2)
            key_e = self._entry(row_f, tk.StringVar(value=key_text), width=25)
            key_e.pack(side="left", padx=5, expand=True, fill="x")
            val_e = self._entry(row_f, tk.StringVar(value=val_text), width=25)
            val_e.pack(side="left", padx=5, expand=True, fill="x")
            del_btn = self._btn(row_f, "削除",
                                lambda f=row_f: remove_pair_row(f))
            del_btn.pack(side="left", padx=5)
            all_rows.append((row_f, key_e, val_e))

        def remove_pair_row(frame_to_remove):
            for i, (f, k, v) in enumerate(all_rows):
                if f == frame_to_remove:
                    f.destroy()
                    all_rows.pop(i)
                    break

        if not current_data:
            add_pair_row()
        else:
            for k, v in current_data:
                add_pair_row(k, v)

        f_btn = ttk.Frame(cust_win, style="Main.TFrame", padding=10)
        f_btn.pack(fill="x")

        def save_and_close():
            data_list.clear()
            for f, key_e, val_e in all_rows:
                k, v = key_e.get(), val_e.get()
                if k:
                    data_list.append((k, v))
            cust_win.destroy()

        self._btn(f_btn, "保存して閉じる", save_and_close).pack(
            side="right", padx=10)
        self._btn(f_btn, "行を追加", lambda: add_pair_row()).pack(side="right")

    # -----------------------------------------------------------------------
    # 実行関数
    # -----------------------------------------------------------------------
    def _run_print(self):
        gui_vars_map = {
            'v_print_ajs_path': self.v_print_ajs_path,
            'v_print_kind': self.v_print_kind,
            'v_srv_c': self.v_srv_c,
            'v_print_conv_flg': self.v_print_conv_flg,
            'v_print_bank': self.v_print_bank,
            'v_print_detail': self.v_print_detail,
            'v_print_out_c': self.v_print_out_c,
            'v_print_out_n': self.v_print_out_n,
            'v_ajs_print_path': self.v_ajs_print_path,
            'v_print_custom_pairs': self.print_custom_pairs,
            'v_jp1_hostname': self.v_jp1_hostname,
            'v_jp1_username': self.v_jp1_username,
        }
        self.run_in_thread(print_start_job, btn_key='print')(
            gui_vars_map, self._gui_funcs_common())

    def _run_recover(self):
        gui_vars_map = {
            'v_recover_file': self.v_recover_file,
            'v_recover_unit': self.v_recover_unit,
            'v_srv_c': self.v_srv_c,
            'v_ajs_define_path': self.v_ajs_define_path,
            'v_jp1_hostname': self.v_jp1_hostname,
            'v_jp1_username': self.v_jp1_username,
        }
        self.run_in_thread(define_start_job, btn_key='recover')(
            gui_vars_map, self._gui_funcs_common())

    def _run_inout(self):
        gui_vars_map = {
            'v_inout_ajs': self.v_inout_ajs,
            'v_inout_res': self.v_inout_res,
            'v_ajs_print_path': self.v_ajs_print_path,
            'v_inout_bank': self.v_inout_bank,
            'v_inout_format': self.v_inout_format,
            'v_inout_custom_vars': self.inout_custom_vars,
            'v_jp1_hostname': self.v_jp1_hostname,
            'v_jp1_username': self.v_jp1_username,
            'inout_text_box': self.t3_text_box,
        }
        self.run_in_thread(inout_start_job, btn_key='inout')(
            gui_vars_map, self._gui_funcs_common())

    def _run_dep(self):
        file_list = self.t5_list_editor.get_values()
        self.v_dep_tgt_files.set("\n".join(file_list))

        gui_vars_map = {
            'v_dep_ajs': self.v_dep_ajs,
            'v_dep_res': self.v_dep_res,
            'v_dep_tgt_files': self.v_dep_tgt_files,
            'v_dep_bank': self.v_dep_bank,
            'v_dep_custom_vars': self.dep_custom_vars,
            'v_t5_out_c': self.v_dep_out_c,
            'v_t5_out_n': self.v_dep_out_n,
            'v_ajs_print_path': self.v_ajs_print_path,
            'v_jp1_hostname': self.v_jp1_hostname,
            'v_jp1_username': self.v_jp1_username,
        }
        open_t5_job_runner(gui_vars_map, self._gui_funcs_common(),
                           self.t5_text_box)


# =========================================================================
# エントリーポイント
# =========================================================================
if __name__ == '__main__':
    app = ASTApp()
    app.mainloop()
