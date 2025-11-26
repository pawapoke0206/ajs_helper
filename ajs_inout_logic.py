#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 入出力解析ロジック (In/Out Logic)

入出力解析のオーケストレーション（全体の流れの制御）を担当する。
シェル解析エンジンは ajs_shell_parser.py、Excel/CSV出力は ajs_inout_report.py に分離。

このファイルの責務:
  - analyze_ajs_jobs(): IO解析のコアロジック（キャッシュ対応）
  - inout_start_job(): Tab3のエントリーポイント
  - inout_parse_exceptions_json(): 例外JSONルールの適用
  - inout_parse_ajsprint_output(): ajsprint出力のパース
"""

import re
import os
import sys
import time
import codecs
import shlex
import json
import pathlib
import copy
import traceback
import datetime
from fnmatch import fnmatch

# 定数ファイルをインポート
from ajs_constants import LOG_DIR, IO_EXCEPTION_FILE, DB_EXCEPTION_FILE, CONFIG_FILE, DIR_NAME_INOUT
from ajs_utils import make_logger, analyze_shell_db_ops, load_db_exceptions

# 分割モジュール
from ajs_shell_parser import (
    ComenvParser, ShellParser, ShellExecutor,
    _build_file_index, _create_replacer, inout_parse_ini_resource,
    ALL_VAR_PAT, VAR_ASSIGN_PAT, RES_PAT,
)
from ajs_inout_report import inout_write_excel, inout_write_csv

# --- グローバル変数: 解析結果のキャッシュ ---
_ANALYSIS_CACHE = {}
_LAST_CACHE_KEY = None

# ログファイルパス
LOG_FILE_RUN = LOG_DIR / "inout_run.log"
LOG_FILE_DETAIL = LOG_DIR / "inout_details.json"

_log = make_logger(LOG_FILE_RUN)


# =====================================================================
# ユーティリティ
# =====================================================================

def write_detail_log(log_data):
    """詳細ログ(JSON)出力"""
    try:
        ld = log_data.copy()
        if "comenv_master_dictionary" in ld and isinstance(ld["comenv_master_dictionary"], dict):
            ld["comenv_master_dictionary"] = {",".join(k): v for k, v in ld["comenv_master_dictionary"].items()}

        with open(LOG_FILE_DETAIL, "w", encoding="utf-8") as f:
            json.dump(ld, f, indent=2, ensure_ascii=False)
        _log(f"[Info] Detail log saved: {LOG_FILE_DETAIL}")
    except Exception as e:
        _log(f"[Error] Failed to write detail log: {e}")

def inout_parse_ajsprint_output(local_tmp_path):
    ajs_mapping_list = []
    headers = ["unit_full", "unit", "resource", "type", "env", "param"]
    with codecs.open(local_tmp_path, "r", "cp932", errors="ignore") as f:
        for line in f:
            parts = line.rstrip("\n").split("\t")
            if len(parts) < len(headers): parts.extend([""] * (len(headers) - len(parts)))
            ajs_mapping_list.append(dict(zip(headers, parts)))
    return ajs_mapping_list

def inout_resolve_path_variables(paths, var_dict):
    resolved_paths, unresolved_vars = [], set()
    replacer_func = _create_replacer(var_dict)
    for path in paths:
        if "$" not in path:
            resolved_paths.append(path)
            continue
        resolved_path = ALL_VAR_PAT.sub(replacer_func, path)
        resolved_paths.append(resolved_path)
        if ALL_VAR_PAT.findall(resolved_path):
            for mt in ALL_VAR_PAT.findall(resolved_path):
                key = mt[1] or mt[2] or mt[3]
                if key not in var_dict: unresolved_vars.add(key)
    return resolved_paths, list(unresolved_vars)


# =====================================================================
# 例外JSONルールの適用
# =====================================================================

def inout_parse_exceptions_json(ajs_record, rules, bank, var_dict, res_root="", file_index=None):
    def resolve_path(path_template, context_vars):
        path = path_template
        if "$" not in path: return path
        replacer_func = _create_replacer(context_vars)
        for _ in range(5):
            last_path = path
            path = ALL_VAR_PAT.sub(replacer_func, path)
            path = re.sub(r"\$\{PM\[(\d+)\]\}", lambda m: ajs_record['param'].split()[int(m.group(1))] if len(ajs_record['param'].split()) > int(m.group(1)) else m.group(0), path)
            # EN[変数名]をcontext_vars（comenv+AJS環境変数+パラメータ）から解決する
            path = re.sub(r"\$\{EN\[([^}]+)\]\}", lambda m: context_vars.get(m.group(1), m.group(0)), path)
            # TBL[n]をcontext_varsから解決する（tbl_lookupで登録済み）
            path = re.sub(r"\$\{TBL\[(\d+)\]\}", lambda m: context_vars.get(f"TBL[{m.group(1)}]", m.group(0)), path)
            if path == last_path: return path
        return path

    def _tbl_lookup(rule, context_vars):
        """tbl_lookup定義に基づきTBLファイルを検索し、結果をcontext_varsに登録する"""
        tbl_def = rule.get("tbl_lookup")
        if not tbl_def or not res_root:
            return
        tbl_filename = tbl_def.get("file", "")
        key_template = tbl_def.get("key", "")
        sep = tbl_def.get("sep", ":")
        match_mode = tbl_def.get("match", "field0")
        key = resolve_path(key_template, context_vars)
        if not key or "$" in key:
            return
        tbl_path = file_index.get(tbl_filename) if file_index else None
        if not tbl_path:
            return
        try:
            with open(tbl_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    fields = line.split(sep)
                    hit = False
                    if match_mode == "contains":
                        hit = key in line
                    else:
                        hit = fields[0] == key
                    if hit:
                        for i, val in enumerate(fields):
                            context_vars[f"TBL[{i}]"] = val
                        return
        except Exception:
            pass

    def _tbl_expand(expand_def, context_vars):
        """TBLファイルの多重展開。グループ番号でTBLを検索し、
        全ヒット行からファイルパスのリストを生成する。"""
        tbl_filename = expand_def.get("file", "")
        sep = expand_def.get("sep", ",")
        group_field = expand_def.get("group_field", 0)

        pm_clean = context_vars.get('PM_CLEAN', '')
        if not pm_clean:
            return None
        group_nums = pm_clean.split()

        tbl_path = file_index.get(tbl_filename) if file_index else None
        if not tbl_path:
            return None

        try:
            with open(tbl_path, "r", encoding="utf-8", errors="ignore") as f:
                tbl_lines = [l.strip() for l in f if l.strip()]
        except Exception:
            return None

        def _apply_template(template, fields):
            result = re.sub(
                r'\$\{F\[(\d+)\]\}',
                lambda m: fields[int(m.group(1))] if int(m.group(1)) < len(fields) else m.group(0),
                template)
            return resolve_path(result, context_vars)

        # --- テンプレートモード ---
        input_tmpl_raw = expand_def.get("input_template")
        output_tmpl_raw = expand_def.get("output_template")
        input_tmpls = [input_tmpl_raw] if isinstance(input_tmpl_raw, str) else (input_tmpl_raw or [])
        output_tmpls = [output_tmpl_raw] if isinstance(output_tmpl_raw, str) else (output_tmpl_raw or [])
        if input_tmpls or output_tmpls:
            inputs_result = []
            outputs_result = []
            for gnum in group_nums:
                for line in tbl_lines:
                    fields = line.split(sep)
                    if len(fields) <= group_field:
                        continue
                    if fields[group_field] == gnum:
                        for tmpl in input_tmpls:
                            inputs_result.append(_apply_template(tmpl, fields))
                        for tmpl in output_tmpls:
                            outputs_result.append(_apply_template(tmpl, fields))
            if not inputs_result and not outputs_result:
                return None
            return {"inputs": inputs_result, "outputs": outputs_result}

        # --- レガシーモード（従来のHULFT互換） ---
        expand_field = expand_def.get("expand_field", expand_def.get("outfile_field", 4))
        cycle_field = expand_def.get("cycle_field", 2)
        path_prefix = expand_def.get("path_prefix", "")
        if path_prefix:
            path_prefix = resolve_path(path_prefix, context_vars)

        cycle_dir_map = {
            'D': context_vars.get('JD1DIR', '${JD1DIR}'),
            'W': context_vars.get('JW1DIR', '${JW1DIR}'),
            'M': context_vars.get('JM1DIR', '${JM1DIR}'),
            'H': context_vars.get('JH1DIR', '${JH1DIR}'),
            'Y': context_vars.get('JY1DIR', '${JY1DIR}'),
            'O': context_vars.get('JO1DIR', '${JO1DIR}'),
        }

        result = []
        for gnum in group_nums:
            for line in tbl_lines:
                fields = line.split(sep)
                if len(fields) <= max(group_field, expand_field):
                    continue
                if fields[group_field] == gnum:
                    filename = fields[expand_field]
                    if path_prefix:
                        result.append(f"{path_prefix}/{filename}")
                    else:
                        cycle = fields[cycle_field] if len(fields) > cycle_field else 'D'
                        out_dir = cycle_dir_map.get(cycle, '${JD1DIR}')
                        result.append(f"{out_dir}/{filename}")

        return result if result else None

    for rule in rules:
        if not fnmatch(bank, rule.get("bank", "*")): continue
        if not fnmatch(os.path.basename(ajs_record['resource']), rule.get("shell", "*")): continue
        if not fnmatch(ajs_record['unit'], rule.get("unit", "*")): continue

        context_vars = copy.deepcopy(var_dict)
        try: context_vars.update(dict(item.split('=', 1) for item in ajs_record['env'].split(';') if '=' in item))
        except Exception as e:
            _log(f"[Warning] exception JSON env parse failed ({ajs_record.get('unit', '?')}): {e}")
        try:
            for i, p in enumerate(ajs_record['param'].split()): context_vars[f"{i+1}"] = p
        except Exception as e:
            _log(f"[Warning] exception JSON param parse failed ({ajs_record.get('unit', '?')}): {e}")
        # PM_CLEAN: パラメタからリダイレクト(2>&1等)以降を除去した値
        try:
            raw_params = ajs_record.get('param', '')
            clean = re.split(r'\s*\d*>&\d*', raw_params)[0].strip()
            context_vars['PM_CLEAN'] = clean if clean else raw_params.strip()
        except Exception:
            context_vars['PM_CLEAN'] = ajs_record.get('param', '').strip()
        # DLYCNT_PLUS1: DLYCNT+1の計算済み値（ディレード番号+1、exprの代替）
        try:
            dlycnt = context_vars.get('DLYCNT', '')
            if dlycnt.isdigit():
                context_vars['DLYCNT_PLUS1'] = str(int(dlycnt) + 1)
        except Exception as e:
            _log(f"[Warning] DLYCNT_PLUS1 calc failed ({ajs_record.get('unit', '?')}): {e}")

        _tbl_lookup(rule, context_vars)

        expand_def = rule.get("tbl_expand") or rule.get("hulft_tbl")
        if expand_def and res_root:
            expand_result = _tbl_expand(expand_def, context_vars)
            if expand_result is not None:
                source_tag = rule.get("source_tag", "例外JSON")
                if isinstance(expand_result, dict):
                    return expand_result.get("inputs", []), expand_result.get("outputs", []), source_tag
                expand_dir = expand_def.get("expand", "outputs")
                if expand_dir == "inputs":
                    outputs = [resolve_path(p, context_vars) for p in rule.get("outputs", [])]
                    return expand_result, outputs, source_tag
                else:
                    inputs = [resolve_path(p, context_vars) for p in rule.get("inputs", [])]
                    return inputs, expand_result, source_tag

        inputs = [resolve_path(p, context_vars) for p in rule.get("inputs", [])]
        outputs = [resolve_path(p, context_vars) for p in rule.get("outputs", [])]
        source_tag = rule.get("source_tag", "例外JSON")

        _, u_in = inout_resolve_path_variables(inputs, context_vars)
        _, u_out = inout_resolve_path_variables(outputs, context_vars)
        if u_in or u_out: source_tag = f"解析失敗: 未解決変数 (JSON) {{{', '.join(set(u_in + u_out))}}}"
        return inputs, outputs, source_tag
    return None, None, None


# =====================================================================
# 解析コアロジック (キャッシュ対応)
# =====================================================================

def analyze_ajs_jobs(gui_vars, gui_funcs, out_dir=None, use_cache=True):
    """
    AJS定義取得〜I/O解析までを行う再利用可能な関数
    """
    global _ANALYSIS_CACHE, _LAST_CACHE_KEY

    update_status = gui_funcs['update_status']
    get_ssh_client = gui_funcs['get_ssh_client']
    show_error = gui_funcs['show_error']

    # Tab 3 (inout) or Tab 5 (dep)
    var_ajs = gui_vars.get('v_inout_ajs') or gui_vars.get('v_dep_ajs')
    ajs_path = var_ajs.get().strip() if var_ajs else ""

    var_res = gui_vars.get('v_inout_res') or gui_vars.get('v_dep_res')
    res_root = var_res.get().strip() if var_res else ""

    var_bank = gui_vars.get('v_inout_bank') or gui_vars.get('v_dep_bank')
    bank = var_bank.get() if var_bank else ""

    ajs_print_path = gui_vars['v_ajs_print_path'].get()
    jp1_hostname = gui_vars['v_jp1_hostname'].get()
    jp1_username = gui_vars['v_jp1_username'].get()

    # キャッシュキーに例外JSONの更新日時を含める（編集時に自動で再解析される）
    ex_mtime = IO_EXCEPTION_FILE.stat().st_mtime if IO_EXCEPTION_FILE.exists() else 0
    db_ex_mtime = DB_EXCEPTION_FILE.stat().st_mtime if DB_EXCEPTION_FILE.exists() else 0
    current_key = (ajs_path, res_root, bank, jp1_hostname, jp1_username, ex_mtime, db_ex_mtime)

    if use_cache and current_key in _ANALYSIS_CACHE:
        update_status("解析結果をキャッシュから復元中...", 10)
        _log(f"[Cache] Hit! Using cached analysis data.")
        time.sleep(0.5)
        return _ANALYSIS_CACHE[current_key]

    if not all([ajs_path, res_root, bank]):
        raise ValueError("AJSパス、リソースパス、銀行名が不足しています。")

    if out_dir is None:
        ts = time.strftime("%Y%m%d%H%M%S")
        out_dir = pathlib.Path(sys.argv[0]).resolve().parent / "temp_analysis" / ts
        out_dir.mkdir(parents=True, exist_ok=True)

    log_data = {
        "config_initial_vars": {}, "comenv_path": "未検出", "comenv_case_patterns": {},
        "comenv_master_dictionary": {}, "shell_cache_build_log": {}, "ajs_mapping": [],
        "json_exceptions_log": {}, "shell_execution_log": {}, "ini_regex_log": {}, "final_records": []
    }

    _log(f"[Analyze] Start Analysis. Bank={bank}, Path={ajs_path}")

    initial_vars = {}
    if bank == "その他":
        custom_vars = gui_vars.get('v_inout_custom_vars', gui_vars.get('v_dep_custom_vars', []))
        initial_vars = dict(custom_vars)
    elif CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                initial_vars = json.load(f).get("comenv_initial_vars_by_bank", {}).get(bank, {})
        except Exception as e:
            _log(f"[Warning] config.json read failed: {e}")
    log_data["config_initial_vars"] = initial_vars

    # res_root配下のファイルを1回だけスキャンしてインデックス化（以降のファイル検索は辞書引き）
    file_index = _build_file_index(res_root)
    _log(f"[Analyze] File index built: {len(file_index)} files")

    update_status("comenv 解析中...", 15)
    comenv_path = file_index.get("comenv")
    log_data["comenv_path"] = comenv_path
    _log(f"[Analyze] comenv found: {comenv_path}")

    comenv_parser = ComenvParser(comenv_path, initial_vars, log_data)
    comenv_parser.parse_all_patterns()

    update_status("AJS定義取得中...", 30)
    export_list = []
    export_list.append(f'export JP1_HOSTNAME={shlex.quote(jp1_hostname)}')
    export_list.append(f'export JP1_USERNAME={shlex.quote(jp1_username)}')
    env_str = ' && '.join(export_list)

    tmp_dir = out_dir / "tmp"
    tmp_dir.mkdir(exist_ok=True)

    remote_tmp = f"/tmp/ajs_out_anl_{time.time()}.txt"
    local_tmp = tmp_dir / "ajs_out_raw.txt"

    cmd = f'{env_str} && {ajs_print_path} -F AJSROOT1 -f "%JN%t%jn%t%sc%t%TY%t%En%t%pm" -R {shlex.quote(ajs_path)} > {remote_tmp}'

    _log(f"[Analyze] Executing remote command: {cmd}")

    with get_ssh_client() as ssh:
        ch = ssh.get_transport().open_session()
        ch.exec_command(cmd.encode("cp932"))
        if ch.recv_exit_status() != 0:
            err = ch.makefile_stderr().read().decode('cp932','ignore')
            _log(f"[Error] AJS command failed: {err}")
            raise RuntimeError(f"AJSコマンドエラー: {err}")
        try:
            sftp = ssh.open_sftp()
            sftp.get(str(remote_tmp), str(local_tmp))
            sftp.close()
        finally:
            # 異常終了時もリモート一時ファイルを確実に削除
            try: ssh.exec_command(f"rm -f {remote_tmp}")
            except: pass

    ajs_mapping_list = inout_parse_ajsprint_output(local_tmp)
    _log(f"[Analyze] Retrieved {len(ajs_mapping_list)} units.")
    log_data["ajs_mapping"] = ajs_mapping_list

    ex_rules = []
    if IO_EXCEPTION_FILE.exists():
        try:
            with open(IO_EXCEPTION_FILE, "r", encoding="utf-8") as f: ex_rules = json.load(f).get("rules", [])
        except Exception as e:
            _log(f"[Warning] io_exceptions.json read failed: {e}")

    update_status("シェル解析キャッシュ構築...", 60)
    shell_cache = {}
    db_ops_cache = {}  # {resource_name: {"db_tables": {...}, "mgmt_tables": []}}
    # db_exceptions.json読み込み（精査済みDB操作データ）
    db_exceptions = load_db_exceptions(str(DB_EXCEPTION_FILE.parent))
    unique_shells = set(r['resource'] for r in ajs_mapping_list if r['resource'] and not r['resource'].endswith('.ini'))
    for sname in unique_shells:
        shl_path = file_index.get(os.path.basename(sname))
        if shl_path:
            procedures = ShellParser(shl_path).get_procedures()
            shell_cache[sname] = procedures
            log_data["shell_cache_build_log"][sname] = procedures
            # DB操作: db_exceptions.jsonの精査済みデータを使用
            db_ops_cache[sname] = analyze_shell_db_ops(shl_path, db_exceptions)

    update_status("I/O変数解決実行中...", 70)
    final_records = []
    for record in ajs_mapping_list:
        r_copy = record.copy()
        inputs, outputs, tag = [], [], None
        var_dict = comenv_parser.get_var_dict_for_env(r_copy['env'])

        if not r_copy['resource']: tag = "リソース指定なし"

        if tag is None and ex_rules:
            i_j, o_j, t_j = inout_parse_exceptions_json(r_copy, ex_rules, bank, var_dict, res_root, file_index)
            if t_j:
                inputs, outputs, tag = i_j, o_j, t_j
                log_data["json_exceptions_log"][r_copy['unit_full']] = {'in': inputs, 'out': outputs, 'tag': tag}

        if tag is None and r_copy['resource'] in shell_cache:
            executor = ShellExecutor(shell_cache[r_copy['resource']], var_dict, r_copy)
            inputs, outputs, unres = executor.execute()
            tag = f"解析失敗: 未解決 {unres}" if unres else ("シェル解析 (IO定義無)" if not inputs and not outputs else "シェル解析 (変数解決)")
            log_data["shell_execution_log"][r_copy['unit_full']] = {'in': inputs, 'out': outputs, 'unresolved': unres, 'tag': tag}

        if tag is None and r_copy['resource'] and r_copy['resource'].endswith('.ini'):
            ini_path = file_index.get(os.path.basename(r_copy['resource']))
            if ini_path:
                raw_in, raw_out = inout_parse_ini_resource(ini_path)
                inputs, u_in = inout_resolve_path_variables(raw_in, var_dict)
                outputs, u_out = inout_resolve_path_variables(raw_out, var_dict)
                tag = f"解析失敗: 未解決 {{{', '.join(set(u_in + u_out))}}}" if u_in or u_out else ("正規表現 (変数解決)" if inputs or outputs else "正規表現 (IO定義無)")
                log_data["ini_regex_log"][r_copy['unit_full']] = {'in': inputs, 'out': outputs, 'tag': tag}
            else: tag = "不明 (リソース無)"

        if tag is None: tag = "不明 (非解析対象)"

        r_copy['inputs'], r_copy['outputs'], r_copy['source_tag'] = inputs, outputs, tag
        # DB操作情報の付与（キャッシュから取得）
        db_info = db_ops_cache.get(r_copy.get('resource'), {})
        r_copy['db_tables'] = db_info.get('db_tables', {})
        r_copy['mgmt_tables'] = db_info.get('mgmt_tables', [])
        final_records.append(r_copy)

    log_data["final_records"] = final_records

    _ANALYSIS_CACHE[current_key] = (final_records, log_data)
    _LAST_CACHE_KEY = current_key

    return final_records, log_data


# =====================================================================
# Tab3 エントリーポイント
# =====================================================================

def inout_start_job(gui_vars, gui_funcs):
    update_status = gui_funcs['update_status']
    save_hist = gui_funcs['save_hist']
    show_info = gui_funcs['show_info']
    show_error = gui_funcs['show_error']
    text_box = gui_vars.get('inout_text_box')

    if text_box: text_box.delete('1.0', 'end')

    with open(LOG_FILE_RUN, "w", encoding="utf-8") as f:
        f.write(f"=== InOut Execution Start: {datetime.datetime.now()} ===\n")

    try:
        ts = time.strftime("%Y%m%d%H%M%S")
        base_dir = pathlib.Path(sys.argv[0]).resolve().parent
        out_dir = base_dir / DIR_NAME_INOUT / ts
        out_dir.mkdir(parents=True, exist_ok=True)
        _log(f"[Info] Output Dir: {out_dir}")

        final_records, log_data = analyze_ajs_jobs(gui_vars, gui_funcs, out_dir, use_cache=False)

        out_format = gui_vars['v_inout_format'].get()
        headers = ["ユニット完全名称","ユニット名称","リソース名称","入力ファイル","出力ファイル","備考 (取得方法)"]
        update_status(f"{out_format}出力中...", 90)

        if out_format == "Excel":
            out_path = out_dir / "unit_io_mapping.xlsx"
            inout_write_excel(out_path, final_records, headers, gui_vars=gui_vars)
        else:
            out_path = out_dir / "unit_io_mapping.csv"
            inout_write_csv(out_path, final_records, headers)

        _log(f"[Info] Saved result to: {out_path}")

        if text_box:
            problems = [f"・{r['unit_full']} ({r['source_tag']})" for r in final_records if "解析失敗" in r['source_tag']]
            if problems: text_box.insert('end', f"--- 問題検出 ({len(problems)}件) ---\n" + "\n".join(problems))
            else: text_box.insert('end', "--- 問題なし ---")

        write_detail_log(log_data)
        update_status("完了", 100)
        save_hist()
        _log("[Success] Completed.")
        show_info(f"解析が完了しました。\n結果は以下のファイルに出力されました:\n{out_path}")

    except Exception as e:
        tb = traceback.format_exc()
        _log(f"[Exception] {str(e)}\n{tb}")
        show_error(str(e))
    finally:
        update_status("待機中", 0)
