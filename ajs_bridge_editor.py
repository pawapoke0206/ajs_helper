#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - DBブリッジ エディタ
v1.0 (2026-04-25) - ファイルベースでは追えない依存関係（DB経由等）の手動定義を編集する

ブリッジ定義は「fromのユニットが必要なら、toのユニットも必要」という関係を定義する。
BFS中にfromユニットが見つかった時点でtoユニットも追加され、toの入力も再帰的に追跡される。
"""

import json
import tkinter as tk
from tkinter import ttk, messagebox
import copy

from ajs_constants import DB_BRIDGE_FILE


class BridgeEditDialog:
    """1件のブリッジを追加・編集するモーダルウィンドウ"""

    def __init__(self, parent, banks, bridge_data=None):
        self.parent = parent
        self.banks = banks
        self.bridge_data = bridge_data or {}
        self.result = None

        self.win = tk.Toplevel(parent)
        self.win.title("ブリッジ編集")
        self.win.geometry("600x400")
        self.win.transient(parent)
        self.win.grab_set()

        form = ttk.Frame(self.win, padding=10)
        form.pack(fill="both", expand=True)

        # 銀行
        ttk.Label(form, text="銀行:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.bank_var = tk.StringVar(value=self.bridge_data.get("bank", banks[0] if banks else "*"))
        bank_cb = ttk.Combobox(form, textvariable=self.bank_var, values=["*"] + list(banks), width=20)
        bank_cb.grid(row=0, column=1, sticky="w", padx=5)

        # From（このユニットが見つかったら）
        ttk.Label(form, text="このユニットが必要なら:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.from_var = tk.StringVar(value=self.bridge_data.get("from", ""))
        ttk.Entry(form, textvariable=self.from_var, width=60).grid(row=1, column=1, sticky="w", padx=5)

        # To（このユニットも追加する）
        ttk.Label(form, text="このユニットも追加:").grid(row=2, column=0, sticky="ne", padx=5, pady=5)
        to_frame = ttk.Frame(form)
        to_frame.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
        form.rowconfigure(2, weight=1)

        self.to_text = tk.Text(to_frame, height=8, width=60, wrap="none")
        self.to_text.pack(fill="both", expand=True)
        # 既存のtoをテキストに展開（1行1ユニット）
        for t in self.bridge_data.get("to", []):
            self.to_text.insert("end", t + "\n")

        ttk.Label(form, text="※1行に1ユニット（正規化パス推奨: /日次処理/...）").grid(
            row=3, column=1, sticky="w", padx=5)

        # 備考
        ttk.Label(form, text="備考:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.note_var = tk.StringVar(value=self.bridge_data.get("note", ""))
        ttk.Entry(form, textvariable=self.note_var, width=60).grid(row=4, column=1, sticky="w", padx=5)

        # ボタン
        btn_frame = ttk.Frame(form)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="OK", command=self.on_ok).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="キャンセル", command=self.on_cancel).pack(side="left", padx=10)

        self.win.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.win.bind("<Escape>", lambda e: self.on_cancel())

    def on_ok(self):
        from_val = self.from_var.get().strip()
        if not from_val:
            messagebox.showwarning("入力エラー", "Fromは必須です。", parent=self.win)
            return

        # Toテキストを行ごとに分割（空行除去）
        to_lines = [line.strip() for line in self.to_text.get("1.0", "end").splitlines()
                     if line.strip()]
        if not to_lines:
            messagebox.showwarning("入力エラー", "Toに1つ以上のユニットを入力してください。", parent=self.win)
            return

        self.result = {
            "bank": self.bank_var.get(),
            "from": from_val,
            "to": to_lines,
            "note": self.note_var.get().strip()
        }
        self.win.destroy()

    def on_cancel(self):
        self.win.destroy()

    def show(self):
        self.win.wait_window()
        return self.result


class BridgeEditor:
    """db_bridges.json を編集するウィンドウ"""

    def __init__(self, parent, banks):
        self.parent = parent
        self.banks = banks
        self.all_bridges = self._load()
        self.is_dirty = False

        self.win = tk.Toplevel(parent)
        self.win.title("DBブリッジ エディタ (db_bridges.json)")
        self.win.geometry("900x500")
        self.win.transient(parent)
        self.win.grab_set()

        main = ttk.Frame(self.win, padding=10)
        main.pack(fill="both", expand=True)

        # 説明
        desc = ttk.Label(main, text=(
            "ファイルベースでは追えない依存関係（DB経由等）を定義します。\n"
            "Fromのユニットが必要と判定された時、Toのユニットも自動的に追加されます。"
        ), foreground="#555555")
        desc.pack(fill="x", pady=(0, 10))

        # 一覧（Treeview）
        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill="both", expand=True)

        cols = ("bank", "from", "to", "note")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.heading("bank", text="銀行")
        self.tree.heading("from", text="このユニットが必要なら")
        self.tree.heading("to", text="このユニットも追加")
        self.tree.heading("note", text="備考")

        self.tree.column("bank", width=80, anchor="w")
        self.tree.column("from", width=200, anchor="w")
        self.tree.column("to", width=400, anchor="w")
        self.tree.column("note", width=150, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.bind("<Double-1>", lambda e: self.modify_bridge())

        # ボタン
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=10)
        ttk.Button(btn_frame, text="追加", command=self.add_bridge).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="編集", command=self.modify_bridge).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="削除", command=self.delete_bridge).pack(side="left", padx=5)

        # 保存・閉じる
        bottom = ttk.Frame(main)
        bottom.pack(fill="x")
        ttk.Label(bottom, text="* 変更は「保存して閉じる」までファイルに反映されません").pack(side="left")
        ttk.Button(bottom, text="保存して閉じる", command=self.on_save_and_close).pack(side="right", padx=5)

        self.win.protocol("WM_DELETE_WINDOW", self.on_close)

        # item_id → bridge のマップ
        self.item_map = {}
        self._refresh_tree()

    def _load(self):
        if not DB_BRIDGE_FILE.exists():
            return []
        try:
            with open(DB_BRIDGE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("bridges", [])
        except Exception as e:
            messagebox.showerror("読込エラー", f"db_bridges.json の読み込みに失敗しました。\n{e}",
                                 parent=self.win if hasattr(self, 'win') else self.parent)
            return []

    def _save(self):
        self.all_bridges.sort(key=lambda b: (b.get("bank", "").lower(), b.get("from", "").lower()))
        try:
            with open(DB_BRIDGE_FILE, 'w', encoding='utf-8') as f:
                json.dump({"bridges": self.all_bridges}, f, indent=2, ensure_ascii=False)
            self.is_dirty = False
            return True
        except Exception as e:
            messagebox.showerror("保存エラー", f"db_bridges.json への保存に失敗しました。\n{e}",
                                 parent=self.win)
            return False

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.item_map.clear()
        for bridge in sorted(self.all_bridges,
                             key=lambda b: (b.get("bank", "").lower(), b.get("from", "").lower())):
            to_str = ", ".join(bridge.get("to", []))
            values = (bridge.get("bank", "*"), bridge.get("from", ""),
                      to_str, bridge.get("note", ""))
            item_id = self.tree.insert("", "end", values=values)
            self.item_map[item_id] = bridge

    def add_bridge(self):
        dialog = BridgeEditDialog(self.win, self.banks)
        result = dialog.show()
        if result:
            self.all_bridges.append(result)
            self.is_dirty = True
            self._refresh_tree()

    def modify_bridge(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("選択エラー", "編集するブリッジを選択してください。", parent=self.win)
            return
        old = self.item_map.get(sel[0])
        if not old:
            return
        dialog = BridgeEditDialog(self.win, self.banks, copy.deepcopy(old))
        result = dialog.show()
        if result:
            for i, b in enumerate(self.all_bridges):
                if b is old:
                    self.all_bridges[i] = result
                    break
            self.is_dirty = True
            self._refresh_tree()

    def delete_bridge(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("選択エラー", "削除するブリッジを選択してください。", parent=self.win)
            return
        if messagebox.askyesno("確認", f"{len(sel)}件のブリッジを削除しますか？", parent=self.win):
            for item_id in sel:
                bridge = self.item_map.get(item_id)
                if bridge and bridge in self.all_bridges:
                    self.all_bridges.remove(bridge)
            self.is_dirty = True
            self._refresh_tree()

    def _check_dirty(self):
        if not self.is_dirty:
            return True
        answer = messagebox.askyesnocancel("未保存の変更",
                                           "未保存の変更があります。保存しますか？",
                                           parent=self.win)
        if answer is None:
            return False
        elif answer:
            return self._save()
        else:
            self.is_dirty = False
            return True

    def on_save_and_close(self):
        if self._check_dirty():
            self.win.destroy()

    def on_close(self):
        if self._check_dirty():
            self.win.destroy()


def open_bridge_editor(parent, banks):
    """ajs_main.pyから呼び出すエントリーポイント"""
    BridgeEditor(parent, banks)
