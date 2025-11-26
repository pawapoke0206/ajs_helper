#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
入出力解析 - シェル解析エンジン (Shell Parser)

シェルスクリプトとcomenv(共通環境変数ファイル)を解析し、
各ジョブの入出力ファイルパスを特定する。

主要クラス:
  - ComenvParser: comenvのcase文を全パターン展開して変数辞書を生成
  - ShellParser: シェルスクリプトを手続きリスト（ASSIGN/IO_ASSIGN/RM/CP等）に変換
  - ShellExecutor: 手続きリストを仮想実行し、入出力パスを確定

ヘルパー:
  - _build_file_index(): リソースフォルダの全ファイルを辞書化
  - _create_replacer(): 変数展開用の置換関数を生成
  - inout_parse_ini_resource(): iniファイルからIO定義を正規表現で抽出
"""

import re
import os
import codecs
import copy
import itertools

from ajs_constants import LOG_DIR
from ajs_utils import make_logger

LOG_FILE_RUN = LOG_DIR / "inout_run.log"
_log = make_logger(LOG_FILE_RUN)

# -----------------------------------------------------------------------------
# 正規表現定義（シェル解析で使用）
# -----------------------------------------------------------------------------
RES_PAT = re.compile(r"^(FILE([IO])(\d{2})|CBL_SYS([IO\d])(\d{2})|([IO])(\d{2})FILE|IN_FILE|OUT_FILE)=(.+)$", re.VERBOSE)
ALL_VAR_PAT = re.compile(r"\$(\{([^}]+)\}|([a-zA-Z_][a-zA-Z0-9_]*)|(\d+))")
CASE_VAR_PAT = re.compile(r'case\s*["\']?\$(\{([^}]+)\}|([a-zA-Z_][a-zA-Z0-9_]*))["\']?')
VAR_ASSIGN_PAT = re.compile(r'^\s*(?:export\s+)?([^=\s]+)\s*=\s*(' r'[\'"](.*?)[\'"]|' r'(?:[^ \t;#]+)' r')')
RM_PAT = re.compile(r'^\s*rm\s+(?:-f\s+)?(.*)')
# cpコマンド: cp [-p] src dst
CP_PAT = re.compile(r'^\s*cp\s+(?:-[a-zA-Z]+\s+)?(\S+)\s+(\S+)')
# リダイレクト: 何か > ${変数} or >> ${変数}（IO変数への書き込み検出用）
REDIRECT_PAT = re.compile(r'>\s*(\$\{[^}]+\}|\$[a-zA-Z_]\w*)\s*$')


# -----------------------------------------------------------------------------
# ヘルパー関数
# -----------------------------------------------------------------------------

def _build_file_index(res_root):
    """res_root配下の全ファイルを1回スキャンして「ファイル名→フルパス」の辞書を返す。
    同名ファイルがある場合は最初に見つかったものが優先される。"""
    index = {}
    if not res_root or not os.path.isdir(res_root):
        return index
    for dirpath, _dirnames, filenames in os.walk(res_root):
        for fname in filenames:
            if fname not in index:
                index[fname] = os.path.join(dirpath, fname)
    return index


def _create_replacer(var_dict):
    def replacer(match):
        key = match.group(2) or match.group(3) or match.group(4)
        return var_dict.get(key, match.group(0)) if key else match.group(0)
    return replacer


# -----------------------------------------------------------------------------
# ComenvParser: 共通環境変数ファイルの解析
# -----------------------------------------------------------------------------

class ComenvParser:
    def __init__(self, comenv_path, initial_vars, log_data):
        self.comenv_path = comenv_path
        self.initial_vars = initial_vars
        self.log_data = log_data
        self.lines = self._read_comenv()
        self.case_patterns = {}
        self.master_var_dict = {}
        self.case_vars_order = []

    def _read_comenv(self):
        if not self.comenv_path: return []
        try:
            with codecs.open(self.comenv_path, "r", "cp932", errors="ignore") as f:
                return f.readlines()
        except Exception as e:
            _log(f"[Warning] comenv read error: {e}")
            return []

    def _resolve_value(self, value, current_vars):
        if "$" not in value: return value
        replacer_func = _create_replacer(current_vars)
        for _ in range(10):
            new_value = ALL_VAR_PAT.sub(replacer_func, value)
            if new_value == value: return new_value
            value = new_value
        return value

    def _evaluate_if_statement(self, line, current_vars):
        match = re.search(r'if\s*\[\s*"\$\{([^}]+)\}"\s*=\s*"([^"]+)"\s*\]', line)
        if match:
            return current_vars.get(match.group(1), "") == match.group(2)
        return True

    def parse_all_patterns(self):
        if not self.lines: return

        # case文の変数を抽出
        current_case_var = None
        for line in self.lines:
            line_stripped = line.strip()
            if line_stripped.startswith("case"):
                match = CASE_VAR_PAT.search(line_stripped)
                if match:
                    current_case_var = match.group(2) or match.group(3)
                    if current_case_var:
                        if current_case_var not in self.case_vars_order:
                            self.case_vars_order.append(current_case_var)
                        self.case_patterns.setdefault(current_case_var, set())
            elif current_case_var:
                match = re.match(r'^\s*([^)\s]+)\s*\)', line_stripped)
                if match:
                    pattern = match.group(1).strip()
                    if pattern != "*": self.case_patterns[current_case_var].add(pattern)
            elif line_stripped.startswith("esac"):
                current_case_var = None

        # パターン組み合わせ生成
        pattern_lists = []
        for var_name in self.case_vars_order:
            patterns = sorted(list(self.case_patterns.get(var_name, set())))
            patterns.append("*")
            pattern_lists.append(patterns)

        if not pattern_lists:
            pattern_lists = [["*"]]
            self.case_vars_order = ["_default_"]

        self.log_data["comenv_case_patterns"] = self.case_patterns
        all_combinations = list(itertools.product(*pattern_lists))

        # --- 各パターン組み合わせでcomenvを仮想実行 ---
        # comenvのcase文は「銀行コード」等で変数値を分岐させる。
        # 全パターンの組み合わせで1回ずつ実行し、結果の変数辞書をmaster_var_dictに格納する。
        #
        # active_block: 現在の行が「有効」かどうか。if/case分岐で切り替わる。
        #   - case文では最初にマッチしたパターンだけを有効にする（case_match_found）
        #   - naive実装（全パターンを有効にする）だと、case文の後段で上書きされて
        #     意図しない変数値になる
        # case_active: case〜esacの中にいるかどうか（パターン行の検出用）
        for combination in all_combinations:
            current_pattern_map = dict(zip(self.case_vars_order, combination))
            current_vars = self.initial_vars.copy()
            active_block = True
            case_active = False
            current_case_var = None
            case_match_found = False

            for line in self.lines:
                line_stripped = line.strip()
                if not line_stripped or line_stripped.startswith("#"): continue
                if line_stripped.startswith("if"):
                    active_block = self._evaluate_if_statement(line_stripped, current_vars)
                    continue
                elif line_stripped.startswith("else"):
                    active_block = not active_block
                    continue
                elif line_stripped.startswith("fi"):
                    active_block = True
                    continue

                if line_stripped.startswith("case"):
                    match = CASE_VAR_PAT.search(line_stripped)
                    if match:
                        current_case_var = match.group(2) or match.group(3)
                        case_match_found = False
                        case_active = True
                    continue
                elif line_stripped.startswith("esac"):
                    current_case_var = None
                    case_match_found = False
                    case_active = False
                    active_block = True
                    continue

                is_pattern_line = False
                if case_active:
                    match = re.match(r'^\s*([^)\s]+)\s*\)', line_stripped)
                    if match:
                        is_pattern_line = True
                        pattern = match.group(1).strip()
                        target_pattern = current_pattern_map.get(current_case_var, "*")
                        if not case_match_found and (pattern == target_pattern or pattern == "*"):
                            active_block = True
                            case_match_found = True
                        else:
                            active_block = False
                    elif line_stripped.startswith(";;"):
                        active_block = False
                        continue

                if active_block:
                    line_to_parse = line.replace("export ", "").strip()
                    if is_pattern_line: line_to_parse = line_to_parse.split(")", 1)[-1].strip()
                    if line_to_parse.endswith(";;"): line_to_parse = line_to_parse[:-2].strip()
                    line_to_parse = line_to_parse.split('#', 1)[0].strip()
                    var_match = VAR_ASSIGN_PAT.match(line_to_parse)
                    if var_match:
                        var_name = var_match.group(1)
                        value = var_match.group(3) if var_match.group(3) is not None else var_match.group(2)
                        value = value.strip("'\"")
                        current_vars[var_name] = self._resolve_value(value, current_vars)
            self.master_var_dict[combination] = current_vars

    def get_var_dict_for_env(self, ajs_env_str):
        if not self.master_var_dict: return self.initial_vars
        env_vars = dict(item.split('=', 1) for item in ajs_env_str.split(';') if '=' in item)
        key_tuple = []
        for var_name in self.case_vars_order:
            val = env_vars.get(var_name, "*")
            if var_name in self.case_patterns and val not in self.case_patterns[var_name]: val = "*"
            key_tuple.append(val)
        return self.master_var_dict.get(tuple(key_tuple), self.initial_vars)


# -----------------------------------------------------------------------------
# ShellParser: シェルスクリプトの静的解析
# -----------------------------------------------------------------------------

class ShellParser:
    def __init__(self, shell_path):
        self.shell_path = shell_path
        self.procedures = []
        self._parse_shell_to_procedures()

    def _parse_shell_to_procedures(self):
        try:
            with codecs.open(self.shell_path, "r", "cp932", errors="ignore") as f:
                for line in f:
                    line = line.strip().split('#', 1)[0].strip()
                    if not line: continue
                    rm_match = RM_PAT.match(line)
                    if rm_match:
                        self.procedures.append(("RM", rm_match.group(1).strip()))
                        continue
                    # cpコマンド検出: cp src dst -> ("CP", src, dst)
                    cp_match = CP_PAT.match(line)
                    if cp_match:
                        self.procedures.append(("CP", cp_match.group(1), cp_match.group(2)))
                    # リダイレクト検出: ... > ${VAR} -> ("REDIRECT_OUT", 変数参照文字列)
                    redir_match = REDIRECT_PAT.search(line)
                    if redir_match:
                        self.procedures.append(("REDIRECT_OUT", redir_match.group(1)))
                    a_match = VAR_ASSIGN_PAT.match(line)
                    if a_match:
                        val = a_match.group(3) if a_match.group(3) is not None else a_match.group(2)
                        self.procedures.append(("ASSIGN", a_match.group(1), val.strip("'\"")))
                    b_match = RES_PAT.match(line)
                    if b_match:
                        var = b_match.group(0).split('=',1)[0]
                        val = b_match.group(8)
                        io_groups = (b_match.group(1), b_match.group(2), b_match.group(4), b_match.group(6))
                        self.procedures.append(("IO_ASSIGN", var, val.strip("'\""), io_groups))
        except Exception as e:
            _log(f"[Warning] Shell parse error ({self.shell_path}): {e}")

    def get_procedures(self): return self.procedures


# -----------------------------------------------------------------------------
# ShellExecutor: 手続きリストの仮想実行
# -----------------------------------------------------------------------------

class ShellExecutor:
    def __init__(self, procedures, comenv_dict, ajs_record):
        self.procedures = procedures
        self.comenv_dict = comenv_dict
        self.ajs_record = ajs_record
        self.shell_context = {}
        self._init_context()

    def _init_context(self):
        self.shell_context = copy.deepcopy(self.comenv_dict)
        try:
            env_vars = dict(item.split('=', 1) for item in self.ajs_record['env'].split(';') if '=' in item)
            self.shell_context.update(env_vars)
        except Exception as e:
            _log(f"[Warning] env parse failed ({self.ajs_record.get('unit', '?')}): {e}")
        try:
            params = self.ajs_record['param'].split()
            for i, param in enumerate(params): self.shell_context[f"{i+1}"] = param
        except Exception as e:
            _log(f"[Warning] param parse failed ({self.ajs_record.get('unit', '?')}): {e}")

    def _resolve_value(self, value_template):
        if not value_template or "$" not in value_template: return value_template
        replacer_func = _create_replacer(self.shell_context)
        return ALL_VAR_PAT.sub(replacer_func, value_template)

    def execute(self):
        inputs, outputs = [], []
        unresolved_io_vars = set()
        # IO方向補正用: cp/リダイレクトで実際に入力・出力として使われたIO変数パスを記録
        actual_outputs = set()  # 書き込み先として使われたパス
        actual_inputs = set()   # 読み取り元として使われたパス
        # IO変数として定義されたパスの一覧（補正対象を限定するため）
        io_defined_paths = set()

        for proc in self.procedures:
            op = proc[0]
            if op == "RM":
                resolved = self._resolve_value(proc[1])
                outputs = [f for f in outputs if f != resolved]
                inputs = [f for f in inputs if f != resolved]
                continue
            # cpコマンド: srcが入力、dstが出力
            if op == "CP":
                src_resolved = self._resolve_value(proc[1])
                dst_resolved = self._resolve_value(proc[2])
                if '/dev/null' not in src_resolved:
                    actual_inputs.add(src_resolved)
                actual_outputs.add(dst_resolved)
                continue
            # リダイレクト: 書き込み先なので出力
            if op == "REDIRECT_OUT":
                redir_resolved = self._resolve_value(proc[1])
                actual_outputs.add(redir_resolved)
                continue
            name, val_tmpl = proc[1], proc[2]
            resolved_val = self._resolve_value(val_tmpl)
            if op == "ASSIGN":
                self.shell_context[name] = resolved_val
            elif op == "IO_ASSIGN":
                self.shell_context[name] = resolved_val
                io_defined_paths.add(resolved_val)
                is_unresolved = False
                if ALL_VAR_PAT.findall(resolved_val):
                    for mt in ALL_VAR_PAT.findall(resolved_val):
                        key = mt[1] or mt[2] or mt[3]
                        if key not in self.shell_context: is_unresolved = True
                if '`' in resolved_val or '$(' in resolved_val: is_unresolved = True
                if is_unresolved: unresolved_io_vars.add(val_tmpl)
                tag, io2, sys01, io3 = proc[3]
                is_input = False
                if "IN_FILE" in tag: is_input = True
                elif "OUT_FILE" in tag: is_input = False
                elif io2: is_input = (io2 == "I")
                elif sys01: is_input = (sys01 in ("I", "0"))
                elif io3: is_input = (io3 == "I")
                if is_input: inputs.append(resolved_val)
                else: outputs.append(resolved_val)

        # --- IO方向補正 ---
        # なぜ必要か: IO変数名(FILEI01等)の命名規則だけでは入出力の方向を誤判定する
        # ケースがある。例えばFILEI01（命名上は入力）に cp で書き込んでいる場合、
        # 実際には出力として扱うべき。cp/リダイレクトの構文解析結果(actual_inputs/
        # actual_outputs)と照合して方向を補正する。
        # io_defined_pathsに限定する理由: cp先がたまたまIO変数と同じパスだった場合の
        # 誤補正を防ぐため、IO_ASSIGNで明示的に定義されたパスのみ対象にしている。
        for path in list(inputs):
            if path in io_defined_paths and path in actual_outputs and path not in actual_inputs:
                inputs.remove(path)
                if path not in outputs:
                    outputs.append(path)
        for path in list(outputs):
            if path in io_defined_paths and path in actual_inputs and path not in actual_outputs:
                outputs.remove(path)
                if path not in inputs:
                    inputs.append(path)

        return inputs, outputs, list(unresolved_io_vars)


# -----------------------------------------------------------------------------
# iniファイルからのIO定義抽出
# -----------------------------------------------------------------------------

def inout_parse_ini_resource(path: str):
    inputs, outputs = [], []
    try:
        with codecs.open(path, "r", "cp932", errors="ignore") as f:
            for line in f:
                m = RES_PAT.match(line.strip())
                if not m: continue
                val = m.group(8).strip()
                tag, io2, sys01, io3 = m.group(1), m.group(2), m.group(4), m.group(6)
                is_input = False
                if "IN_FILE" in tag: is_input = True
                elif "OUT_FILE" in tag: is_input = False
                elif io2: is_input = (io2 == "I")
                elif sys01: is_input = (sys01 in ("I", "0"))
                elif io3: is_input = (io3 == "I")
                if is_input: inputs.append(val)
                else: outputs.append(val)
    except Exception as e:
        _log(f"[Warning] ini parse failed ({path}): {e}")
    return sorted(list(set(inputs))), sorted(list(set(outputs)))
