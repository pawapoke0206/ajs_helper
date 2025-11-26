#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 共通ユーティリティ
複数のロジックファイルで使い回す関数を集約する。
"""

import os
import re
import json
import codecs
import datetime


def make_logger(log_file):
    """指定ログファイルに書き込む _log 関数を生成して返す。
    使い方: _log = make_logger(LOG_DIR / "inout_run.log")
    """
    def _log(msg):
        try:
            timestamp = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {msg}\n")
        except Exception as e:
            print(f"[Warning] Log write failed ({log_file}): {e}", file=__import__('sys').stderr)
    return _log


def pre_normalize(path: str, base: str) -> str:
    """AJSパスを正規化する。
    1. サービス名プレフィックス（AJSROOT1: 等）を除去
    2. base部分を除去
    """
    p = re.sub(r'^[a-zA-Z0-9_]+:', '', path.strip())
    if base and base != "/" and p.startswith(base):
        p = p.replace(base, '', 1)
    return p


def write_detail_log(data_dict, log_file, logger=None):
    """詳細ログ(JSON)出力。set/dict/listを適切に変換してJSON保存する。
    logger: _log関数。指定されればログメッセージも出力する。
    """
    try:
        serializable = {}
        for k, v in data_dict.items():
            if isinstance(v, set):
                serializable[k] = sorted(list(v))
            elif isinstance(v, list):
                serializable[k] = v
            elif isinstance(v, dict):
                new_dict = {}
                for sub_k, sub_v in v.items():
                    if isinstance(sub_v, set):
                        new_dict[sub_k] = sorted(list(sub_v))
                    else:
                        new_dict[sub_k] = str(sub_v)
                serializable[k] = new_dict
            else:
                serializable[k] = str(v)

        with open(log_file, "w", encoding="utf-8") as f:
            json.dump(serializable, f, indent=2, ensure_ascii=False)
        if logger:
            logger(f"[Info] Detail log saved: {log_file}")
    except Exception as e:
        if logger:
            logger(f"[Error] Failed to write detail log: {e}")


# =============================================================================
# シェル特性解析（DB操作検出）
# =============================================================================
# テーブル名の判定: DMR/DMRC/DMG/UNY/IKO プレフィクスで始まる5文字以上
_TABLE_NAME_PREFIXES = ("DMR", "DMRC", "DMG", "UNY", "IKO")

# DB関連変数の代入パターン
_DB_VAR_RE = re.compile(
    r'(?:TBL_ID|TBLID|DB_ID|DB_NAME|DBS_NAME|TBL_NAME|DAY_TBL|TBLNAME|tbl_name)'
    r"""=["\']?([A-Za-z]\w+)["\']?"""
)

# サブシェル呼び出しパターン（管理テーブル操作）
# 更新系: MRKM900000, MRKH900000, MRKD900000
# 参照系: MRKM900005, MRKH900005
_SUB_SHELL_RE = re.compile(
    r'\$\{?(?:SHLDIR|BSDIR)\}?(?:/SHL)?/(\w+)\.sh\s+(\w+)'
)
_SUB_SHELL_WRITE = {"MRKM900000", "MRKH900000", "MRKD900000"}
_SUB_SHELL_READ = {"MRKM900005", "MRKH900005"}


def _is_table_name(name):
    """テーブル名として妥当かどうか判定"""
    upper = name.upper()
    return any(upper.startswith(p) for p in _TABLE_NAME_PREFIXES) and len(name) > 4


def _read_shell_content(shell_path):
    """シェルファイルを読む（cp932 → utf-8 フォールバック）"""
    for enc in ("cp932", "utf-8", "latin-1"):
        try:
            with open(shell_path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, OSError):
            continue
    return ""


def analyze_shell_db_ops(shell_path, db_exceptions=None):
    """シェルのDB操作情報を返す。

    db_exceptions.jsonの精査済みデータを優先的に使用する。
    db_exceptionsが渡されていない、またはシェルが登録されていない場合は
    空の結果を返す（自動検出は現在不採用）。

    Args:
        shell_path: シェルファイルのパス
        db_exceptions: load_db_exceptions()で読み込んだ辞書。
                       {シェル名(拡張子なし): {"db_tables": {...}}}

    戻り値: dict
      {
        "db_tables": {"DMR040D00": "RW", "DMR010D00": "R", ...},
        "mgmt_tables": []
      }
      DB操作がない場合は {"db_tables": {}, "mgmt_tables": []}

    操作の分類:
      - "W": 書き込みのみ（pdload -d 洗い替え等）
      - "R": 参照のみ（pdrorg -k unld, pdsql SELECT等）
      - "RW": 参照+更新（SELECT+UPDATE等）
    """
    empty = {"db_tables": {}, "mgmt_tables": []}

    if not db_exceptions:
        return empty

    # シェル名（拡張子なし）でルックアップ
    shell_name = os.path.splitext(os.path.basename(shell_path))[0]
    entry = db_exceptions.get(shell_name)
    if not entry:
        return empty

    return {"db_tables": dict(entry.get("db_tables", {})), "mgmt_tables": []}


# db_exceptions.jsonのキャッシュ（モジュールレベル）
_db_exceptions_cache = None
_db_exceptions_mtime = 0


def load_db_exceptions(base_path):
    """db_exceptions.jsonを読み込み、シェル名→DB操作の辞書を返す。

    ファイルの更新時刻が変わっていなければキャッシュを返す。
    ファイルが存在しない場合は空辞書を返す。

    Args:
        base_path: 基準パス（db_exceptions.jsonの親ディレクトリ）

    戻り値: {シェル名(拡張子なし): {"db_tables": {...}}}
    """
    global _db_exceptions_cache, _db_exceptions_mtime

    json_path = os.path.join(base_path, "db_exceptions.json")
    if not os.path.exists(json_path):
        return {}

    mtime = os.path.getmtime(json_path)
    if _db_exceptions_cache is not None and mtime == _db_exceptions_mtime:
        return _db_exceptions_cache

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # shells配列をシェル名辞書に変換
        result = {}
        for entry in data.get("shells", []):
            shell_name = entry.get("shell", "")
            if shell_name:
                result[shell_name] = {"db_tables": entry.get("db_tables", {})}
        _db_exceptions_cache = result
        _db_exceptions_mtime = mtime
        return result
    except Exception as e:
        print(f"[Warning] db_exceptions.json read failed: {e}", file=__import__('sys').stderr)
        return {}


def _get_mgmt_table(sub_shell_name):
    """サブシェル名から対応する管理テーブル名を返す"""
    mapping = {
        "MRKM900000": "DMR233M00",
        "MRKM900005": "DMR233M00",
        "MRKH900000": "DMR234H00",
        "MRKH900005": "DMR234H00",
        "MRKD900000": "DMR235D00",
    }
    return mapping.get(sub_shell_name, "UNKNOWN")
