#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 共通定数定義
v3.6 (2025-11-23) - 出力フォルダ名変更 ("AJS定義取得" -> "ジョブ定義取得")
"""

import json
import logging
import pathlib
import sys

# ───────────── EXE化対応 基準パスの取得 ─────────────
def get_base_path():
    if getattr(sys, 'frozen', False):
        return pathlib.Path(sys.executable).resolve().parent
    else:
        return pathlib.Path(__file__).resolve().parent

BASE_DIR = get_base_path()

# ───────────── ログ・フォルダ設定 ─────────────
# ログフォルダを先に定義・作成
LOG_DIR = BASE_DIR / "log"
LOG_DIR.mkdir(exist_ok=True) 

# ログ設定 (logフォルダ配下に出力)
logging.basicConfig(filename=LOG_DIR / 'debug.log', level=logging.DEBUG, filemode='w', format='%(message)s')
log = logging.debug

# その他のファイルパス
HIST_FILE = BASE_DIR / 'history.json'
PARAM_FILE = BASE_DIR / 'AJS_trans.prm'
IO_EXCEPTION_FILE = BASE_DIR / 'io_exceptions.json'
DB_EXCEPTION_FILE = BASE_DIR / 'db_exceptions.json'
DB_BRIDGE_FILE = BASE_DIR / 'db_bridges.json'
CONFIG_FILE = BASE_DIR / 'config.json'

MAX_HIST = 10

# ───────────── 出力フォルダ名定義 ─────────────
DIR_NAME_PRINT = "ジョブ定義取得" # ★修正
DIR_NAME_INOUT = "入出力解析"
DIR_NAME_PRE   = "先行関係解析"
DIR_NAME_DEP   = "依存関係解析"

# ───────────── AJSコマンド設定 ─────────────
AJS_PRINT_PATH = "/opt/jp1ajs2/bin/ajsprint"
AJS_DEFINE_PATH = "/opt/jp1ajs2/bin/ajsdefine"

# --- JP1環境変数 デフォルト値 ---
DEFAULT_JP1_HOSTNAME = ""
DEFAULT_JP1_USERNAME = "jp1admin"

# ───────────── 文字コード・改行コード設定 ─────────────
ENC = {'SJIS': 'cp932', 'UTF-8': 'utf-8', 'SJIS(CP932)': 'cp932'}
NL = {'CRLF(Windows)': '\r\n', 'LF(Unix)': '\n'}

# ───────────── 銀行リスト (config.jsonから読み込み) ─────────────
def load_banks():
    # comenv_initial_vars_by_bank のキー一覧を銀行名リストとして使う
    # 銀行名はconfig.json(.gitignore対象)にのみ存在し、ソースコードには含めない
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            banks = list(json.load(f).get("comenv_initial_vars_by_bank", {}).keys())
            if banks:
                return banks
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    return ["その他"]

BANKS = load_banks()
