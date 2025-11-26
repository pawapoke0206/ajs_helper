#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - I/O 手動ルール エディタ
v1.5 (2025-11-12) - ソートを大文字/小文字 区別しないように修正
"""

import json
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import copy

# 定数ファイルをインポート
from ajs_constants import IO_EXCEPTION_FILE

# -----------------------------------------------------------------------------
# ★ 汎用「リスト編集」ウィジェット (変更なし)
# -----------------------------------------------------------------------------
class ListEditor(ttk.Frame):
    """
    「行追加」「行削除」機能を持つリスト入力ウィジェット
    """
    def __init__(self, parent, title, initial_list=None):
        super().__init__(parent)
        
        self.label = ttk.Label(self, text=title)
        self.label.pack(fill="x", padx=5)
        
        self.canvas_frm = ttk.Frame(self, relief="solid", borderwidth=1)
        self.canvas_frm.pack(fill="both", expand=True, padx=5, pady=(2, 5))
        
        self.canvas = tk.Canvas(self.canvas_frm, borderwidth=0, background="#ffffff", height=100)
        self.scroll_frame = ttk.Frame(self.canvas, padding=(5, 5)) 
        self.vsb = ttk.Scrollbar(self.canvas_frm, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas_window_id = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        self.scroll_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.all_rows = [] 
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=5)
        ttk.Button(btn_frame, text="行を追加", command=self.add_row).pack(side="left")

        if initial_list:
            for item in initial_list:
                self.add_row(item)
        else:
            self.add_row("") 

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window_id, width=event.width)

    def _on_mouse_wheel(self, event):
        if hasattr(event, 'delta'):
             self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 5:
             self.canvas.yview_scroll(1, "units")
        elif event.num == 4:
             self.canvas.yview_scroll(-1, "units")

    def add_row(self, value=""):
        row_f = ttk.Frame(self.scroll_frame)
        row_f.pack(fill="x", pady=1)
        entry = ttk.Entry(row_f)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        entry.insert(0, value)
        del_btn = ttk.Button(row_f, text="削除", width=5, command=lambda f=row_f: self.remove_row(f))
        del_btn.pack(side="left")
        self.all_rows.append((row_f, entry))

    def remove_row(self, frame_to_remove):
        for i, (f, entry) in enumerate(self.all_rows):
            if f == frame_to_remove:
                f.destroy()
                self.all_rows.pop(i)
                break

    def get_values(self):
        return [entry.get() for f, entry in self.all_rows if entry.get().strip()]

    def bind_mousewheel(self):
        self.canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)

    def unbind_mousewheel(self):
        self.canvas.unbind_all("<MouseWheel>")

# -----------------------------------------------------------------------------
# ★ ルール編集ポップアップ (変更なし)
# -----------------------------------------------------------------------------
class RuleEditDialog:
    """
    1件のルールを追加・編集するためのモーダルウィンドウ
    """
    def __init__(self, parent, bank, rule_data=None):
        self.parent = parent
        self.bank = bank
        self.rule_data = rule_data or {} 
        self.result_rule = None 
        
        self.win = tk.Toplevel(parent)
        self.win.title(f"ルール編集 (銀行: {bank})")
        self.win.geometry("700x500") 
        self.win.transient(parent)
        self.win.grab_set()

        form_frame = ttk.Frame(self.win, padding=10)
        form_frame.pack(fill="both", expand=True)

        # --- フォーム ---
        ttk.Label(form_frame, text="シェル名 (shell):").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        default_shell = "" if not self.rule_data else self.rule_data.get("shell", "")
        self.shell_var = tk.StringVar(form_frame, value=default_shell)
        ttk.Entry(form_frame, textvariable=self.shell_var, width=40).grid(row=0, column=1, sticky="w", padx=5)

        ttk.Label(form_frame, text="ユニット名 (unit):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.unit_var = tk.StringVar(form_frame, value=self.rule_data.get("unit", "*"))
        ttk.Entry(form_frame, textvariable=self.unit_var, width=40).grid(row=1, column=1, sticky="w", padx=5)

        ttk.Label(form_frame, text="備考タグ (source_tag):").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        default_tag = "手動ルール" if not self.rule_data else self.rule_data.get("source_tag", "手動ルール")
        self.tag_var = tk.StringVar(form_frame, value=default_tag)
        ttk.Entry(form_frame, textvariable=self.tag_var, width=40).grid(row=2, column=1, sticky="w", padx=5)

        # --- 入出力 (リストエディタ) ---
        io_frame = ttk.Frame(form_frame)
        io_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=10)
        form_frame.rowconfigure(3, weight=1)
        form_frame.columnconfigure(0, weight=1)
        form_frame.columnconfigure(1, weight=1)
        io_frame.columnconfigure(0, weight=1)
        io_frame.columnconfigure(1, weight=1)

        self.inputs_editor = ListEditor(io_frame, "入力 (inputs)", self.rule_data.get("inputs", []))
        self.inputs_editor.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        self.outputs_editor = ListEditor(io_frame, "出力 (outputs)", self.rule_data.get("outputs", []))
        self.outputs_editor.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        io_frame.rowconfigure(0, weight=1) 

        # --- ボタン ---
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=5)
        
        ttk.Button(btn_frame, text="OK", command=self.on_ok).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="キャンセル", command=self.on_cancel).pack(side="left", padx=10)
        
        self.win.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.win.bind("<Return>", lambda e: self.on_ok())
        self.win.bind("<Escape>", lambda e: self.on_cancel())
        
        self.inputs_editor.bind_mousewheel()
        self.outputs_editor.bind_mousewheel()

    def on_ok(self):
        shell = self.shell_var.get()
        if not shell:
            messagebox.showwarning("入力エラー", "シェル名は必須です。", parent=self.win)
            return

        self.result_rule = {
            "bank": self.bank,
            "shell": shell,
            "unit": self.unit_var.get(),
            "inputs": self.inputs_editor.get_values(),
            "outputs": self.outputs_editor.get_values(),
            "source_tag": self.tag_var.get()
        }
        
        self.inputs_editor.unbind_mousewheel()
        self.outputs_editor.unbind_mousewheel()
        self.win.destroy()

    def on_cancel(self):
        self.inputs_editor.unbind_mousewheel()
        self.outputs_editor.unbind_mousewheel()
        self.win.destroy()

    def show(self):
        self.win.wait_window()
        return self.result_rule

# -----------------------------------------------------------------------------
# ★ メインエディタウィンドウ (★ 修正)
# -----------------------------------------------------------------------------
class ExceptionEditor:
    """
    io_exceptions.json を編集するための専用ウィンドウクラス
    """
    def __init__(self, parent, banks):
        self.parent = parent
        self.banks = banks
        self.all_rules = self.load_rules() 
        self.is_dirty = False 
        
        self.win = tk.Toplevel(parent)
        self.win.title("I/O 手動ルール エディタ (io_exceptions.json)")
        self.win.geometry("800x600")
        self.win.transient(parent)
        self.win.grab_set()
        
        main_frame = ttk.Frame(self.win, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True, pady=5)
        
        self.bank_tabs = {} 
        for bank in self.banks:
            self.create_bank_tab(bank)
            
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)
        
        close_btn_frame = ttk.Frame(main_frame)
        close_btn_frame.pack(fill="x", pady=(10, 0))
        
        ttk.Label(close_btn_frame, text="* 変更は「保存して閉じる」までファイルに反映されません").pack(side="left")
        ttk.Button(close_btn_frame, text="保存して閉じる", command=self.on_save_and_close).pack(side="right", padx=5)
        
        self.win.protocol("WM_DELETE_WINDOW", self.on_close_window) 

    def load_rules(self):
        if not IO_EXCEPTION_FILE.exists():
            return []
        try:
            with open(IO_EXCEPTION_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get("rules", [])
        except Exception as e:
            messagebox.showerror("読込エラー", f"{IO_EXCEPTION_FILE.name} の読み込みに失敗しました。\n{e}", parent=self.win)
            return []

    def save_rules(self):
        """
        現在の all_rules をソートしてJSONファイルに書き戻す
        """
        # ★★★ 修正点 1 (Sort) ★★★
        # .lower() を使って大文字/小文字を区別せずにソート
        self.all_rules.sort(key=lambda r: (r.get("bank", "").lower(), r.get("shell", "").lower()))
        
        try:
            with open(IO_EXCEPTION_FILE, 'w', encoding='utf-8') as f:
                json.dump({"rules": self.all_rules}, f, indent=2, ensure_ascii=False)
            self.is_dirty = False 
            return True
        except Exception as e:
            messagebox.showerror("保存エラー", f"{IO_EXCEPTION_FILE.name} への保存に失敗しました。\n{e}", parent=self.win)
            return False

    def create_bank_tab(self, bank):
        """銀行ごとのタブUIを作成する"""
        tab_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_frame, text=bank)
        
        # 1. ルール一覧 (Treeview)
        tree_frame = ttk.Frame(tab_frame)
        tree_frame.pack(fill="both", expand=True)
        
        cols = ("shell", "unit", "inputs", "outputs", "tag")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15, selectmode="extended")
        tree.pack(side="left", fill="both", expand=True)
        
        tree.heading("shell", text="シェル名 (shell)")
        tree.heading("unit", text="ユニット名 (unit)")
        tree.heading("inputs", text="入力 (概要)")
        tree.heading("outputs", text="出力 (概要)")
        tree.heading("tag", text="備考タグ (source_tag)")
        
        tree.column("shell", width=150, anchor="w")
        tree.column("unit", width=100, anchor="w")
        tree.column("inputs", width=150, anchor="w")
        tree.column("outputs", width=150, anchor="w")
        tree.column("tag", width=100, anchor="w")
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=vsb.set)
        
        # 2. 操作ボタン (Treeviewの下)
        btn_frame = ttk.Frame(tab_frame)
        btn_frame.pack(fill="x", pady=10)
        
        btn_new = ttk.Button(btn_frame, text="ルールの追加", command=lambda: self.add_rule(bank))
        btn_new.pack(side="left", padx=5)
        
        btn_update = ttk.Button(btn_frame, text="選択したルールを変更", command=lambda: self.modify_rule(bank))
        btn_update.pack(side="left", padx=5)
        
        btn_del = ttk.Button(btn_frame, text="選択したルールを削除", command=lambda: self.delete_rule(bank))
        btn_del.pack(side="left", padx=5)
        
        self.bank_tabs[bank] = {
            "tree": tree,
            "item_map": {} 
        }
        
        tree.bind("<Double-1>", lambda e, b=bank: self.modify_rule(b))
        
        self.load_rules_for_bank(bank)

    def load_rules_for_bank(self, bank):
        """Treeview に指定された銀行のルールを読み込む"""
        widgets = self.bank_tabs[bank]
        tree = widgets["tree"]
        
        tree.delete(*tree.get_children())
        widgets["item_map"].clear()
            
        bank_rules = [r for r in self.all_rules if r.get("bank") == bank]
        # ★★★ 修正点 1 (Sort) ★★★
        # .lower() を使って大文字/小文字を区別せずにソート
        bank_rules.sort(key=lambda r: r.get("shell", "").lower())
        
        for rule in bank_rules:
            inputs_str = ", ".join(rule.get("inputs", []))
            outputs_str = ", ".join(rule.get("outputs", []))
            values = (
                rule.get("shell", ""), rule.get("unit", ""),
                inputs_str, outputs_str, rule.get("source_tag", "")
            )
            item_id = tree.insert("", "end", values=values)
            widgets["item_map"][item_id] = rule

    def add_rule(self, bank):
        """「ルールの追加...」ボタン"""
        dialog = RuleEditDialog(self.win, bank, None)
        new_rule = dialog.show()
        
        if new_rule:
            self.all_rules.append(new_rule)
            self.is_dirty = True
            # Treeview を再読み込み (ソートを反映させるため)
            self.load_rules_for_bank(bank)

    def modify_rule(self, bank):
        """「選択したルールを変更...」ボタン"""
        widgets = self.bank_tabs[bank]
        tree = widgets["tree"]
        
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("変更エラー", "変更するルールを一覧から選択してください。", parent=self.win)
            return
        
        item_id = selected_item[0]
        
        old_rule = widgets["item_map"].get(item_id)
        if not old_rule:
            messagebox.showerror("内部エラー", "選択されたルールの参照が見つかりません。", parent=self.win)
            return

        dialog = RuleEditDialog(self.win, bank, copy.deepcopy(old_rule))
        updated_rule = dialog.show()
        
        if updated_rule:
            for i, rule in enumerate(self.all_rules):
                if rule == old_rule:
                    self.all_rules[i] = updated_rule
                    break
            
            self.is_dirty = True
            # Treeview を再読み込み (ソートを反映させるため)
            self.load_rules_for_bank(bank) 


    def delete_rule(self, bank):
        """「選択したルールを削除」ボタン"""
        widgets = self.bank_tabs[bank]
        tree = widgets["tree"]
        
        selected_items = tree.selection() 
        if not selected_items:
            messagebox.showwarning("削除エラー", "削除するルールを一覧から選択してください。", parent=self.win)
            return

        # 確認メッセージ
        msg = f"{len(selected_items)} 件のルールを削除しますか？\n(保存するまでファイルには反映されません)"
        if len(selected_items) == 1:
            rule_to_delete = widgets["item_map"].get(selected_items[0])
            msg = f"以下のルールを削除しますか？\nシェル: {rule_to_delete.get('shell')}\nユニット: {rule_to_delete.get('unit')}"

        if messagebox.askyesno("確認", msg, parent=self.win):
            
            for item_id in selected_items:
                rule_to_delete = widgets["item_map"].get(item_id)
                if not rule_to_delete:
                    continue 

                if rule_to_delete in self.all_rules:
                    self.all_rules.remove(rule_to_delete)
                
                tree.delete(item_id)
                del widgets["item_map"][item_id] 
            
            self.is_dirty = True

    def check_dirty_and_save(self):
        """
        変更フラグをチェックし、必要なら保存を促す
        戻り値: (True=処理継続OK, False=処理キャンセル)
        """
        if not self.is_dirty:
            return True 
        
        answer = messagebox.askyesnocancel(
            "未保存の変更", 
            "ルールに未保存の変更があります。保存しますか？",
            parent=self.win
        )
        
        if answer is None: return False
        elif answer is True: return self.save_rules()
        else:
            self.is_dirty = False
            return True

    def on_tab_change(self, event):
        pass

    def on_save_and_close(self):
        """「保存して閉じる」ボタン"""
        if self.check_dirty_and_save():
            self.win.destroy()

    def on_close_window(self):
        """「X」ボタン"""
        if self.check_dirty_and_save():
            self.win.destroy()


# --- ajs_main.py から呼び出されるエントリーポイント ---
def open_editor_window(parent, banks):
    editor = ExceptionEditor(parent, banks)
