#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool (Main GUI)
v4.6 (2025-11-23) - バグ修正(create_result_textbox定義漏れ), 解析結果エリアの横スクロール対応

機能:
1. 定義取得・変換 (Print & Convert)
2. 定義回復 (Recover)
3. 入出力解析 (In/Out Analysis)
4. 先行関係解析 (Predecessor Analysis)
5. 依存関係解析 (Dependency Analysis)
"""

import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import paramiko
import platform

# 外部ファイルをインポート
from ajs_constants import *
from ajs_print_logic import print_start_job
from ajs_define_logic import define_start_job
from ajs_inout_logic import inout_start_job
from ajs_rel_logic import pre_start_job
from ajs_depend_logic import open_t5_job_runner 
from ajs_exception_editor import open_editor_window

# ───────────── ★ デザイン定数 ─────────────
PAD_X = 5
PAD_Y = 3
BTN_PAD_Y = 10

# ───────────── ★ 共通ヘルパー ─────────────
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True) 
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tooltip, 
            text=self.text, 
            background="#ffffe0", 
            relief="solid", 
            borderwidth=1,
            font=("", 9, "normal"), 
            padx=4, 
            pady=4
        ) 
        label.pack()

    def hide_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()
        self.tooltip = None

# ───────────── ★ スクロール制御付きウィジェット作成関数 (ここを追加) ─────────────

def setup_scroll_handling(widget, main_scroll_func):
    """
    ウィジェット上のマウスホイール操作を、そのウィジェット専用にする。
    (親画面のスクロールを止める return "break" を使用)
    """
    def _on_local_wheel(event):
        if platform.system() == "Windows":
            widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            widget.yview_scroll(-1, "units")
        elif event.num == 5:
            widget.yview_scroll(1, "units")
        return "break" # 親への伝播を阻止

    def _bind_local(event):
        if platform.system() == "Windows":
            widget.bind_all("<MouseWheel>", _on_local_wheel)
        else:
            widget.bind_all("<Button-4>", _on_local_wheel)
            widget.bind_all("<Button-5>", _on_local_wheel)

    def _unbind_local(event):
        # メイン画面のスクロールに戻す
        if platform.system() == "Windows":
            widget.bind_all("<MouseWheel>", main_scroll_func)
        else:
            widget.bind_all("<Button-4>", main_scroll_func)
            widget.bind_all("<Button-5>", main_scroll_func)

    widget.bind("<Enter>", _bind_local)
    widget.bind("<Leave>", _unbind_local)


def create_result_textbox(parent, main_scroll_func, height=10):
    """
    横スクロールバー付き、折り返しなしのテキストボックスを作成する
    """
    frame = ttk.Frame(parent)
    frame.pack(fill="both", expand=True)

    # スクロールバー
    v_scroll = ttk.Scrollbar(frame, orient="vertical")
    h_scroll = ttk.Scrollbar(frame, orient="horizontal")

    # テキストウィジェット (wrap="none" で折り返しなし)
    text_widget = tk.Text(frame, height=height, wrap="none", undo=False, borderwidth=1, relief="solid")
    
    # 配置 (Gridを使用)
    text_widget.grid(row=0, column=0, sticky="nsew")
    v_scroll.grid(row=0, column=1, sticky="ns")
    h_scroll.grid(row=1, column=0, sticky="ew")

    # リサイズ対応
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    # 連動設定
    text_widget.config(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
    v_scroll.config(command=text_widget.yview)
    h_scroll.config(command=text_widget.xview)

    # スクロール干渉防止
    setup_scroll_handling(text_widget, main_scroll_func)

    return text_widget


# ───────────── ★ リストエディタ (スクロール制御付き) ─────────────
class FileListEditor(ttk.Frame):
    def __init__(self, parent, main_scroll_func):
        super().__init__(parent)
        self.main_scroll_func = main_scroll_func 
        
        self.canvas_frm = ttk.Frame(self, relief="sunken", borderwidth=1)
        self.canvas_frm.pack(fill="both", expand=True, pady=(0, 5))
        
        self.canvas = tk.Canvas(self.canvas_frm, borderwidth=0, background="#ffffff", height=100)
        self.scroll_frame = ttk.Frame(self.canvas) 
        self.vsb = ttk.Scrollbar(self.canvas_frm, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas_window_id = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        self.scroll_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # 干渉防止
        setup_scroll_handling(self.canvas, self.main_scroll_func)

        self.rows = [] 
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="＋ 行を追加", command=lambda: self.add_row("")).pack(side="left", padx=PAD_X)
        ttk.Button(btn_frame, text="全クリア", command=self.clear_all).pack(side="right", padx=PAD_X)

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window_id, width=event.width)

    def add_row(self, text_value):
        row_f = ttk.Frame(self.scroll_frame)
        row_f.pack(fill="x", pady=2, padx=2)
        
        entry = ttk.Entry(row_f)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        entry.insert(0, text_value)
        
        del_btn = ttk.Button(row_f, text="×", width=3, command=lambda f=row_f: self.remove_row(f))
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


# ───────────── メインウィンドウ構築 ─────────────
root = tk.Tk()
root.title("AJS Helper Tool v4.6")
root.geometry("800x650") 

# --- 全体スクロール用 Canvas ---
main_canvas = tk.Canvas(root)
main_scrollbar = ttk.Scrollbar(root, orient="vertical", command=main_canvas.yview)
scrollable_frame = ttk.Frame(main_canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
)

main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
main_canvas.configure(yscrollcommand=main_scrollbar.set)

main_scrollbar.pack(side="right", fill="y")
main_canvas.pack(side="left", fill="both", expand=True)

def on_main_mousewheel(event):
    if platform.system() == "Windows":
        main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    elif event.num == 4:
        main_canvas.yview_scroll(-1, "units")
    elif event.num == 5:
        main_canvas.yview_scroll(1, "units")

if platform.system() == "Windows":
    root.bind_all("<MouseWheel>", on_main_mousewheel)
else:
    root.bind_all("<Button-4>", on_main_mousewheel)
    root.bind_all("<Button-5>", on_main_mousewheel)


# ==========================================
# コンテンツ配置 (scrollable_frame内)
# ==========================================
main_frm = ttk.Frame(scrollable_frame, padding=15)
main_frm.pack(fill="both", expand=True)

def load_hist():
    if not HIST_FILE.exists():
        return {}
    try:
        return json.loads(HIST_FILE.read_text(encoding='utf-8'))
    except:
        return {}

hist = load_hist()

# --- 共通接続情報 ---
conn_frm = ttk.LabelFrame(main_frm, text="共通接続情報", padding=10)
conn_frm.pack(fill="x", expand=True, pady=(0, 10))
conn_frm.columnconfigure(1, weight=1)
conn_frm.columnconfigure(3, weight=1)
conn_frm.columnconfigure(5, weight=1)

ttk.Label(conn_frm, text="IP:").grid(row=0, column=0, sticky="e", padx=PAD_X)
v_ip = tk.StringVar(root)
ttk.Combobox(conn_frm, textvariable=v_ip, values=hist.get('ip', [])).grid(row=0, column=1, sticky="ew", padx=PAD_X)

ttk.Label(conn_frm, text="User:").grid(row=0, column=2, sticky="e", padx=PAD_X)
v_user = tk.StringVar(root)
ttk.Combobox(conn_frm, textvariable=v_user, values=hist.get('user', [])).grid(row=0, column=3, sticky="ew", padx=PAD_X)

ttk.Label(conn_frm, text="Pass:").grid(row=0, column=4, sticky="e", padx=PAD_X)
v_pass = tk.StringVar(root)
ttk.Entry(conn_frm, textvariable=v_pass, show="*").grid(row=0, column=5, sticky="ew", padx=PAD_X)

sub_frame = ttk.Frame(conn_frm)
sub_frame.grid(row=0, column=6, columnspan=2, sticky="e", padx=PAD_X)
v_srv_c = tk.StringVar(root, 'SJIS')
ttk.Label(sub_frame, text="Enc:").pack(side="left")
ttk.Combobox(sub_frame, textvariable=v_srv_c, values=['SJIS', 'UTF-8'], width=6, state='readonly').pack(side="left", padx=5)

def open_advanced_settings():
    adv_win = tk.Toplevel(root)
    adv_win.title("詳細設定")
    adv_win.geometry("500x200") 
    adv_win.transient(root)
    adv_win.grab_set() 
    
    frm = ttk.Frame(adv_win, padding=15)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)
    
    ttk.Label(frm, text="ajsprint パス:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    ttk.Entry(frm, textvariable=v_ajs_print_path).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
    
    ttk.Label(frm, text="ajsdefine パス:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    ttk.Entry(frm, textvariable=v_ajs_define_path).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
    
    ttk.Separator(frm, orient="horizontal").grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)
    
    ttk.Label(frm, text="JP1_HOSTNAME:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    e_host = ttk.Entry(frm, textvariable=v_jp1_hostname)
    e_host.grid(row=3, column=1, sticky="ew", padx=5, pady=5)

    ttk.Label(frm, text="JP1_USERNAME:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
    e_user = ttk.Entry(frm, textvariable=v_jp1_username)
    e_user.grid(row=4, column=1, sticky="ew", padx=5, pady=5)

    btn_frm = ttk.Frame(frm)
    btn_frm.grid(row=5, column=0, columnspan=2, pady=10)
    
    def restore():
        v_ajs_print_path.set(AJS_PRINT_PATH)
        v_ajs_define_path.set(AJS_DEFINE_PATH)
        v_jp1_hostname.set(DEFAULT_JP1_HOSTNAME)
        v_jp1_username.set(DEFAULT_JP1_USERNAME)

    ttk.Button(btn_frm, text="デフォルトに戻す", command=restore).pack(side="left", padx=10)
    ttk.Button(btn_frm, text="閉じる", command=adv_win.destroy).pack(side="left", padx=10)

ttk.Button(sub_frame, text="詳細...", width=6, command=open_advanced_settings).pack(side="left", padx=5)


# --- ノートブック ---
notebook = ttk.Notebook(main_frm)
notebook.pack(fill="both", expand=True, pady=5)

# --- 共通変数定義 ---
# Tab 1
v_print_ajs_path = tk.StringVar(root)
v_print_kind = tk.StringVar(root, "verify")
v_print_conv_flg = tk.StringVar(root, "no")
v_print_bank = tk.StringVar(root)
v_print_detail = tk.StringVar(root)
v_print_out_c = tk.StringVar(root, 'SJIS(CP932)')
v_print_out_n = tk.StringVar(root, 'CRLF(Windows)')
print_custom_pairs = [] 

# Tab 2
v_recover_file = tk.StringVar(root)
v_recover_unit = tk.StringVar(root, "") 

# Tab 3
v_inout_ajs = tk.StringVar(root)
v_inout_res = tk.StringVar(root)
v_inout_bank = tk.StringVar(root) 
v_inout_format = tk.StringVar(root, "Excel") 
inout_custom_vars = [] 

# Tab 4
v_pre_root = tk.StringVar(root)
v_pre_tgt = tk.StringVar(root)
v_pre_out_c = tk.StringVar(root, 'SJIS(CP932)')
v_pre_out_n = tk.StringVar(root, 'CRLF(Windows)')

# Tab 5
v_dep_ajs = tk.StringVar(root)
v_dep_res = tk.StringVar(root)
v_dep_tgt_files = tk.StringVar(root) 
v_dep_bank = tk.StringVar(root) 
v_dep_out_c = tk.StringVar(root, 'SJIS(CP932)')
v_dep_out_n = tk.StringVar(root, 'CRLF(Windows)')
dep_custom_vars = [] 

# Commands & Env
v_ajs_print_path = tk.StringVar(root, AJS_PRINT_PATH)
v_ajs_define_path = tk.StringVar(root, AJS_DEFINE_PATH)
v_jp1_hostname = tk.StringVar(root, DEFAULT_JP1_HOSTNAME)
v_jp1_username = tk.StringVar(root, DEFAULT_JP1_USERNAME)

status_var = tk.StringVar(root, '待機中')
progress = tk.DoubleVar(root, 0.0)
BANKS = ["香川", "徳島大正", "トマト", "高知", "大光", "大東", "栃木", "静岡中央", "三十三", "その他"]


# --- ヘルパー関数 ---
def update_status(msg, p_val=None):
    root.after(0, lambda: status_var.set(msg))
    if p_val is not None:
        root.after(0, lambda: progress.set(p_val))

def show_error(msg):
    root.after(0, lambda: messagebox.showerror('エラー', msg))
    update_status('エラー')

def show_info(msg):
    root.after(0, lambda: messagebox.showinfo('完了', msg))
    update_status('完了')

def check_thread(thread, btn):
    if thread.is_alive():
        root.after(100, lambda: check_thread(thread, btn))
    else:
        btn.config(state="normal")

def run_in_thread(target_func):
    def wrapper(*args, **kwargs):
        try:
            current_tab = notebook.index(notebook.select())
            btn_map = {0: run_btn_print, 1: run_btn_rec, 2: run_btn_inout, 3: run_btn_pre, 4: run_btn_dep}
            btn = btn_map.get(current_tab)
        except Exception:
            btn = None 

        if btn:
            btn.config(state="disabled")
        
        thread = threading.Thread(target=target_func, args=args, kwargs=kwargs, daemon=True)
        thread.start()
        
        if btn:
            check_thread(thread, btn)
    return wrapper

def get_ssh_client():
    ip, user, pw = v_ip.get(), v_user.get(), v_pass.get()
    if not all([ip, user, pw]):
        raise ValueError("接続情報 (IP, ユーザー, パスワード) を入力してください。")
    
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(ip, username=user, password=pw, timeout=10)
    return ssh

def save_hist():
    hist_data = load_hist()
    hist_items = {
        'ip': v_ip.get(), 'user': v_user.get(), 
        'print_ajs_path': v_print_ajs_path.get(),
        'recover_unit_name': v_recover_unit.get(), 
        'inout_ajs_path': v_inout_ajs.get(), 'inout_res_path': v_inout_res.get(), 
        'pre_root': v_pre_root.get(), 'pre_tgt': v_pre_tgt.get(),
        'dep_ajs_path': v_dep_ajs.get(), 'dep_res_path': v_dep_res.get(),
        'dep_tgt_files': v_dep_tgt_files.get(),
    }
    
    for key, value in hist_items.items():
        if not value: continue
        new_list = [value] + [x for x in hist_data.get(key, []) if x != value]
        hist_data[key] = new_list[:MAX_HIST]
        
    HIST_FILE.write_text(json.dumps(hist_data, indent=2, ensure_ascii=False), encoding='utf-8')

gui_funcs_common = {
    'update_status': update_status, 
    'get_ssh_client': get_ssh_client, 
    'save_hist': save_hist, 
    'show_info': show_info, 
    'show_error': show_error, 
    'run_in_thread': run_in_thread
}

def create_output_selector(parent, outc_var, outn_var):
    frm = ttk.LabelFrame(parent, text="出力定義設定", padding=5)
    ttk.Label(frm, text="出力文字コード").pack(side="left", padx=(5, 5))
    ttk.Combobox(frm, textvariable=outc_var, values=['SJIS(CP932)', 'UTF-8'], width=15, state='readonly').pack(side="left")
    ttk.Label(frm, text="出力改行コード").pack(side="left", padx=(20, 5))
    ttk.Combobox(frm, textvariable=outn_var, values=['CRLF(Windows)', 'LF(Unix)'], width=15, state='readonly').pack(side="left")
    return frm

def open_key_value_window(title, data_list, key_label, value_label):
    current_data = list(data_list)
    cust_win = tk.Toplevel(root)
    cust_win.title(title)
    cust_win.geometry("500x400")
    cust_win.transient(root)
    cust_win.grab_set()

    hdr_frm = ttk.Frame(cust_win, padding=(10, 10, 10, 0))
    hdr_frm.pack(fill="x")
    ttk.Label(hdr_frm, text=key_label, width=30, font=("", 10, "bold")).pack(side="left", padx=5)
    ttk.Label(hdr_frm, text=value_label, width=30, font=("", 10, "bold")).pack(side="left", padx=5)

    canvas_frm = ttk.Frame(cust_win, padding=5)
    canvas_frm.pack(fill="both", expand=True)
    canvas = tk.Canvas(canvas_frm, borderwidth=0, background="#ffffff")
    scroll_frame = ttk.Frame(canvas, padding=(10, 0)) 
    vsb = ttk.Scrollbar(canvas_frm, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

    def _conf(e): canvas.configure(scrollregion=canvas.bbox("all"))
    scroll_frame.bind("<Configure>", _conf)
    
    # キーバリューウィンドウはモーダルなので単純なbindでOK
    def _wheel(e): 
        if e.delta: canvas.yview_scroll(int(-1*(e.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", _wheel)
    cust_win.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

    all_rows = [] 
    def add_pair_row(key_text="", val_text=""):
        row_f = ttk.Frame(scroll_frame)
        row_f.pack(fill="x", pady=2)
        key_e = ttk.Entry(row_f, width=30)
        key_e.pack(side="left", padx=5, expand=True, fill="x")
        key_e.insert(0, key_text)
        val_e = ttk.Entry(row_f, width=30)
        val_e.pack(side="left", padx=5, expand=True, fill="x")
        val_e.insert(0, val_text)
        
        del_btn = ttk.Button(row_f, text="削除", width=5, command=lambda f=row_f: remove_pair_row(f))
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

    f_btn = ttk.Frame(cust_win, padding=10)
    f_btn.pack(fill="x")
    
    def save_and_close():
        data_list.clear()
        for f, key_e, val_e in all_rows:
            k, v = key_e.get(), val_e.get()
            if k:
                data_list.append((k, v))
        cust_win.destroy()

    ttk.Button(f_btn, text="保存して閉じる", command=save_and_close).pack(side="right", padx=10)
    ttk.Button(f_btn, text="行を追加", command=lambda: add_pair_row()).pack(side="right")


# --- Tab 1: 定義取得・変換 ---
tab1 = ttk.Frame(notebook, padding=10)
notebook.add(tab1, text="定義取得・変換")

t1_frm = ttk.LabelFrame(tab1, text="実行設定", padding=10)
t1_frm.pack(fill="x")
t1_frm.columnconfigure(1, weight=1)

ttk.Label(t1_frm, text="AJS パス").grid(row=0, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t1_frm, textvariable=v_print_ajs_path, values=hist.get('print_ajs_path', [])).grid(row=0, column=1, columnspan=2, sticky="ew", padx=PAD_X)

f_def = ttk.Frame(t1_frm)
f_def.grid(row=2, column=1, columnspan=2, sticky="w")
ttk.Label(t1_frm, text="取得する定義").grid(row=2, column=0, sticky="e", pady=PAD_Y)
for t, v in [("回復用", "recover"), ("確認用", "verify"), ("両方", "both")]:
    ttk.Radiobutton(f_def, text=t, variable=v_print_kind, value=v).pack(side="left", padx=5)

def on_detail_select(*_):
    state = "disabled"
    if v_print_conv_flg.get() == "yes" and v_print_detail.get() == "カスタム":
        state = "normal"
    custom_detail_btn.config(state=state)

def on_bank_select(*_):
    bank = v_print_bank.get()
    if bank == "その他":
        v_print_detail.set("カスタム")
        for rb in det_rbs:
            rb.config(state="disabled" if rb.cget("value") != "カスタム" else "normal")
    else:
        for rb in det_rbs:
            rb.config(state="normal")
    on_detail_select() 

def toggle_conv_widgets(*_):
    state = "normal" if v_print_conv_flg.get() == "yes" else "disabled"
    for w in bank_rbs + det_rbs: w.config(state=state)
    if state == "normal": on_bank_select()
    else: custom_detail_btn.config(state="disabled")

ttk.Label(t1_frm, text="変換を行うか").grid(row=3, column=0, sticky="e", pady=PAD_Y)
f_conv = ttk.Frame(t1_frm)
f_conv.grid(row=3, column=1, columnspan=2, sticky="w")
for t, v in [("はい", "yes"), ("いいえ", "no")]:
    ttk.Radiobutton(f_conv, text=t, variable=v_print_conv_flg, value=v, command=toggle_conv_widgets).pack(side="left", padx=5)

ttk.Label(t1_frm, text="銀行名").grid(row=4, column=0, sticky="ne", pady=PAD_Y)
bank_frame = ttk.Frame(t1_frm)
bank_frame.grid(row=4, column=1, columnspan=2, sticky="w")
bank_rbs = []
v_print_bank.set(BANKS[0])
for i, b in enumerate(BANKS):
    rb = ttk.Radiobutton(bank_frame, text=b, variable=v_print_bank, value=b, command=on_bank_select)
    rb.grid(row=i // 5, column=i % 5, sticky="w", padx=4)
    bank_rbs.append(rb)

ttk.Label(t1_frm, text="変換詳細").grid(row=5, column=0, sticky="ne", pady=PAD_Y)
det_frame = ttk.Frame(t1_frm)
det_frame.grid(row=5, column=1, columnspan=2, sticky="w")
det_rbs, DETAILS = [], ["本番⇒ミラー", "本番⇒開発", "ミラー⇒本番", "ミラー⇒開発", "開発⇒本番", "開発⇒ミラー", "カスタム"]
v_print_detail.set(DETAILS[0])
for i, d in enumerate(DETAILS):
    rb = ttk.Radiobutton(det_frame, text=d, variable=v_print_detail, value=d, command=on_detail_select)
    rb.grid(row=i // 3, column=i % 3, sticky="w", padx=4)
    det_rbs.append(rb)

custom_detail_btn = ttk.Button(det_frame, text="詳細...", state="disabled", command=lambda: open_key_value_window(
    "カスタム変換設定", print_custom_pairs, "変換元文言", "変換先文言"))
custom_detail_btn.grid(row=len(DETAILS) // 3, column=len(DETAILS) % 3, sticky="w", padx=10, pady=5)

toggle_conv_widgets() 
create_output_selector(tab1, v_print_out_c, v_print_out_n).pack(fill="x", pady=5, expand=True)

def create_print_runner():
    global print_custom_pairs
    gui_vars_map = {
        'v_print_ajs_path': v_print_ajs_path,
        'v_print_kind': v_print_kind,
        'v_srv_c': v_srv_c,
        'v_print_conv_flg': v_print_conv_flg,
        'v_print_bank': v_print_bank,
        'v_print_detail': v_print_detail,
        'v_print_out_c': v_print_out_c,
        'v_print_out_n': v_print_out_n,
        'v_ajs_print_path': v_ajs_print_path,
        'v_print_custom_pairs': print_custom_pairs,
        'v_jp1_hostname': v_jp1_hostname,
        'v_jp1_username': v_jp1_username,
    }
    run_in_thread(print_start_job)(gui_vars_map, gui_funcs_common)

run_btn_print = ttk.Button(tab1, text="取得＆変換 実行", command=create_print_runner)
run_btn_print.pack(pady=BTN_PAD_Y)


# --- Tab 2: 定義回復 ---
tab2 = ttk.Frame(notebook, padding=10)
notebook.add(tab2, text="定義回復")
t2_frm = ttk.LabelFrame(tab2, text="実行設定", padding=10)
t2_frm.pack(fill="x")
t2_frm.columnconfigure(1, weight=1)

ttk.Label(t2_frm, text="回復用AJS定義").grid(row=0, column=0, sticky="e", pady=PAD_Y)
ttk.Entry(t2_frm, textvariable=v_recover_file).grid(row=0, column=1, sticky="ew", padx=PAD_X)
ttk.Button(t2_frm, text="ファイル選択...", command=lambda: v_recover_file.set(filedialog.askopenfilename())).grid(row=0, column=2)

ttk.Label(t2_frm, text="回復先AJSパス").grid(row=1, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t2_frm, textvariable=v_recover_unit, values=hist.get('recover_unit_name', [])).grid(row=1, column=1, columnspan=2, sticky="ew", padx=PAD_X)

def create_recover_runner():
    gui_vars_map = {
        'v_recover_file': v_recover_file,
        'v_recover_unit': v_recover_unit,
        'v_srv_c': v_srv_c, 
        'v_ajs_define_path': v_ajs_define_path,
        'v_jp1_hostname': v_jp1_hostname,
        'v_jp1_username': v_jp1_username,
    }
    run_in_thread(define_start_job)(gui_vars_map, gui_funcs_common)

run_btn_rec = ttk.Button(tab2, text="回復実行", command=create_recover_runner)
run_btn_rec.pack(pady=BTN_PAD_Y)


# --- Tab 3: 入出力解析 ---
tab3 = ttk.Frame(notebook, padding=10)
notebook.add(tab3, text="入出力解析")

t3_frm = ttk.LabelFrame(tab3, text="実行設定", padding=10)
t3_frm.pack(fill="x")
t3_frm.columnconfigure(1, weight=1)

ttk.Label(t3_frm, text="AJSパス").grid(row=0, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t3_frm, textvariable=v_inout_ajs, values=hist.get('inout_ajs_path', [])).grid(row=0, column=1, sticky="ew", columnspan=2, padx=PAD_X)

ttk.Label(t3_frm, text="リソースパス").grid(row=1, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t3_frm, textvariable=v_inout_res, values=hist.get('inout_res_path', [])).grid(row=1, column=1, sticky="ew", padx=PAD_X)
ttk.Button(t3_frm, text="参照...", command=lambda: v_inout_res.set(filedialog.askdirectory())).grid(row=1, column=2)

ttk.Label(t3_frm, text="銀行名").grid(row=2, column=0, sticky="ne", pady=PAD_Y)
t3_bank_frame = ttk.Frame(t3_frm)
t3_bank_frame.grid(row=2, column=1, columnspan=2, sticky="w")
v_inout_bank.set(BANKS[0])

t3_custom_vars_btn = ttk.Button(t3_bank_frame, text="初期変数設定", state="disabled", command=lambda: open_key_value_window(
    "「その他」用 初期変数設定", inout_custom_vars, "変数名 (例: BSDIR)", "値 (例: /HN)"
))

def on_t3_bank_select(*_):
    if v_inout_bank.get() == "その他": t3_custom_vars_btn.config(state="normal")
    else: t3_custom_vars_btn.config(state="disabled")

for i, b in enumerate(BANKS):
    row = i // 5
    col = i % 5
    rb = ttk.Radiobutton(t3_bank_frame, text=b, variable=v_inout_bank, value=b, command=on_t3_bank_select)
    rb.grid(row=row, column=col, sticky="w", padx=4)
    if b == "その他": t3_custom_vars_btn.grid(row=row, column=col + 1, sticky="w", padx=(0, 10), pady=5)

ttk.Label(t3_frm, text="出力形式").grid(row=3, column=0, sticky="e", pady=PAD_Y)
t3_out_frame = ttk.Frame(t3_frm)
t3_out_frame.grid(row=3, column=1, columnspan=2, sticky="w")
ttk.Radiobutton(t3_out_frame, text="Excel", variable=v_inout_format, value="Excel").pack(side="left", padx=5)
ttk.Radiobutton(t3_out_frame, text="CSV", variable=v_inout_format, value="CSV").pack(side="left", padx=5)

# ルール編集ボタンをフレーム内に配置
ttk.Button(t3_frm, text="I/Oルール編集", command=lambda: open_editor_window(root, [b for b in BANKS if b != "その他"] + ["*"])).grid(row=3, column=2, sticky="e", padx=PAD_X)

t3_res_frm = ttk.LabelFrame(tab3, text="解析結果 (問題のあったユニット一覧)", padding=10)
t3_res_frm.pack(fill="both", expand=True, pady=(10, 0))

# ★修正: create_result_textbox を使用
t3_text_box = create_result_textbox(t3_res_frm, on_main_mousewheel, height=6)

def create_inout_runner():
    global inout_custom_vars
    gui_vars_map = {
        'v_inout_ajs': v_inout_ajs,
        'v_inout_res': v_inout_res,
        'v_ajs_print_path': v_ajs_print_path,
        'v_inout_bank': v_inout_bank,
        'v_inout_format': v_inout_format,
        'v_inout_custom_vars': inout_custom_vars,
        'v_jp1_hostname': v_jp1_hostname,
        'v_jp1_username': v_jp1_username,
        'inout_text_box': t3_text_box,
    }
    run_in_thread(inout_start_job)(gui_vars_map, gui_funcs_common)

run_btn_inout = ttk.Button(tab3, text="解析実行", command=create_inout_runner)
run_btn_inout.pack(pady=BTN_PAD_Y) 


# --- Tab 4: 先行関係解析 ---
tab4 = ttk.Frame(notebook, padding=10)
notebook.add(tab4, text="先行関係解析")

t4_frm = ttk.LabelFrame(tab4, text="実行設定", padding=10)
t4_frm.pack(fill="x")
t4_frm.columnconfigure(1, weight=1)

ttk.Label(t4_frm, text="AJSパス").grid(row=0, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t4_frm, textvariable=v_pre_root, values=hist.get('pre_root', [])).grid(row=0, column=1, sticky="ew", padx=PAD_X)

ttk.Label(t4_frm, text="解析対象ユニットパス").grid(row=1, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t4_frm, textvariable=v_pre_tgt, values=hist.get('pre_tgt', [])).grid(row=1, column=1, sticky="ew", padx=PAD_X)

create_output_selector(tab4, v_pre_out_c, v_pre_out_n).pack(fill="x", pady=5, expand=True)

t4_res_frm = ttk.LabelFrame(tab4, text="解析結果 (関連ユニット一覧)", padding=10)
t4_res_frm.pack(fill="both", expand=True)

# ★修正: create_result_textbox を使用
t4_text_box = create_result_textbox(t4_res_frm, on_main_mousewheel, height=10)

def create_pre_runner():
    gui_vars_map = {
        'v_pre_root': v_pre_root,
        'v_pre_tgt': v_pre_tgt,
        'v_srv_c': v_srv_c,
        'v_pre_out_c': v_pre_out_c,
        'v_pre_out_n': v_pre_out_n,
        'v_ajs_print_path': v_ajs_print_path,
        'v_jp1_hostname': v_jp1_hostname,
        'v_jp1_username': v_jp1_username,
    }
    run_in_thread(pre_start_job)(gui_vars_map, gui_funcs_common, t4_text_box)

run_btn_pre = ttk.Button(tab4, text="解析実行", command=create_pre_runner)
run_btn_pre.pack(pady=BTN_PAD_Y)


# --- Tab 5: 依存関係解析 ---
tab5 = ttk.Frame(notebook, padding=10)
notebook.add(tab5, text="依存関係解析")

t5_frm = ttk.LabelFrame(tab5, text="実行設定", padding=10)
t5_frm.pack(fill="x")
t5_frm.columnconfigure(1, weight=1)

ttk.Label(t5_frm, text="AJSパス").grid(row=0, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t5_frm, textvariable=v_dep_ajs, values=hist.get('dep_ajs_path', [])).grid(row=0, column=1, sticky="ew", padx=PAD_X)

ttk.Label(t5_frm, text="リソースパス").grid(row=1, column=0, sticky="e", pady=PAD_Y)
ttk.Combobox(t5_frm, textvariable=v_dep_res, values=hist.get('dep_res_path', [])).grid(row=1, column=1, sticky="ew", padx=PAD_X)
ttk.Button(t5_frm, text="参照...", command=lambda: v_dep_res.set(filedialog.askdirectory())).grid(row=1, column=2)

ttk.Label(t5_frm, text="銀行名").grid(row=2, column=0, sticky="ne", pady=PAD_Y)
t5_bank_frame = ttk.Frame(t5_frm)
t5_bank_frame.grid(row=2, column=1, columnspan=2, sticky="w")
v_dep_bank.set(BANKS[0])

t5_custom_vars_btn = ttk.Button(t5_bank_frame, text="初期変数設定", state="disabled", command=lambda: open_key_value_window(
    "「その他」用 初期変数設定", dep_custom_vars, "変数名 (例: BSDIR)", "値 (例: /HN)"
))

def on_t5_bank_select(*_):
    if v_dep_bank.get() == "その他": t5_custom_vars_btn.config(state="normal")
    else: t5_custom_vars_btn.config(state="disabled")

for i, b in enumerate(BANKS):
    row = i // 5
    col = i % 5
    rb = ttk.Radiobutton(t5_bank_frame, text=b, variable=v_dep_bank, value=b, command=on_t5_bank_select)
    rb.grid(row=row, column=col, sticky="w", padx=4)
    if b == "その他": t5_custom_vars_btn.grid(row=row, column=col + 1, sticky="w", padx=(0, 10))

lbl_files = ttk.Label(t5_frm, text="目標ファイル")
lbl_files.grid(row=3, column=0, sticky="ne", pady=PAD_Y)

# リストエディタ
t5_list_editor = FileListEditor(t5_frm, on_main_mousewheel)
t5_list_editor.grid(row=3, column=1, columnspan=2, sticky="ew", padx=5, pady=3)

# Tab5初期値は常に空
t5_list_editor.set_values([])

create_output_selector(tab5, v_dep_out_c, v_dep_out_n).pack(fill="x", pady=5, expand=True)

t5_res_frm = ttk.LabelFrame(tab5, text="解析結果 (抽出ユニットと外部入力ファイル一覧)", padding=10)
t5_res_frm.pack(fill="both", expand=True)

# ★修正: create_result_textbox を使用
t5_text_box = create_result_textbox(t5_res_frm, on_main_mousewheel, height=10)

def create_dep_runner():
    global dep_custom_vars
    file_list = t5_list_editor.get_values()
    v_dep_tgt_files.set("\n".join(file_list))
    
    gui_vars_map = {
        'v_dep_ajs': v_dep_ajs,
        'v_dep_res': v_dep_res,
        'v_dep_tgt_files': v_dep_tgt_files,
        'v_dep_bank': v_dep_bank,
        'v_dep_custom_vars': dep_custom_vars,
        'v_t5_out_c': v_dep_out_c,
        'v_t5_out_n': v_dep_out_n, 
        'v_ajs_print_path': v_ajs_print_path,
        'v_jp1_hostname': v_jp1_hostname,
        'v_jp1_username': v_jp1_username,
    }
    open_t5_job_runner(gui_vars_map, gui_funcs_common, t5_text_box)

run_btn_dep = ttk.Button(tab5, text="解析＆定義作成 実行", command=create_dep_runner)
run_btn_dep.pack(pady=BTN_PAD_Y)


# --- 共通ステータスバー ---
status_frm = ttk.Frame(main_frm)
status_frm.pack(fill="x", side="bottom", pady=(10, 0))
ttk.Progressbar(status_frm, variable=progress, maximum=100, length=260, mode='determinate').pack(fill="x")
ttk.Label(status_frm, textvariable=status_var).pack(fill="x", pady=(2, 0))

if __name__ == '__main__':
    root.mainloop()
