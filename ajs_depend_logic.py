#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 依存関係解析ロジック (Dependency Logic)
v3.7 (2025-11-23) - 中間ファイルをtmpフォルダに格納するように修正
"""

import os
import sys
import time
import shlex
import pathlib
import re
import collections
import traceback
import datetime
import json

# 定数・既存ロジック
from ajs_constants import ENC, NL, LOG_DIR, DIR_NAME_DEP
from ajs_rel_logic import pre_filter_definition, pre_parse_graph 
from ajs_inout_logic import analyze_ajs_jobs   

# ログファイルパス
LOG_FILE_RUN = LOG_DIR / "tab5_dep_run.log"
LOG_FILE_DETAIL = LOG_DIR / "tab5_dep_details.json"

def _log(msg):
    """実行ログへの書き込みヘルパー"""
    try:
        timestamp = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        with open(LOG_FILE_RUN, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {msg}\n")
    except: pass

def write_detail_log(data_dict):
    """詳細ログ(JSON)出力"""
    try:
        serializable = {}
        for k, v in data_dict.items():
            if isinstance(v, set):
                serializable[k] = sorted(list(v))
            elif isinstance(v, list):
                serializable[k] = v
            else:
                # ProducerMapのような {path: {unit, unit}} の構造対応
                if isinstance(v, dict):
                    new_dict = {}
                    for sub_k, sub_v in v.items():
                        if isinstance(sub_v, set):
                            new_dict[sub_k] = sorted(list(sub_v))
                        else:
                            new_dict[sub_k] = str(sub_v)
                    serializable[k] = new_dict
                else:
                    serializable[k] = str(v)
                
        with open(LOG_FILE_DETAIL, "w", encoding="utf-8") as f:
            json.dump(serializable, f, indent=2, ensure_ascii=False)
        _log(f"[Info] Detail log saved: {LOG_FILE_DETAIL}")
    except Exception as e:
        _log(f"[Error] Failed to write detail log: {e}")

def dep_start_job(gui_vars, gui_funcs, text_box):
    """依存関係解析のメイン処理"""
    
    update_status = gui_funcs['update_status']
    show_info = gui_funcs['show_info']
    show_error = gui_funcs['show_error']
    save_hist = gui_funcs['save_hist']
    get_ssh_client = gui_funcs['get_ssh_client']
    
    # ログ開始
    with open(LOG_FILE_RUN, "w", encoding="utf-8") as f:
        f.write(f"=== Tab 5 (Dependency) Execution Start: {datetime.datetime.now()} ===\n")

    # パラメータ取得
    target_files_str = gui_vars['v_dep_tgt_files'].get().strip()
    ajs_path_input = gui_vars['v_dep_ajs'].get().strip()
    res_root = gui_vars['v_dep_res'].get().strip()
    bank = gui_vars['v_dep_bank'].get()
    
    ajs_print_path = gui_vars['v_ajs_print_path'].get()
    jp1_hostname = gui_vars['v_jp1_hostname'].get()
    jp1_username = gui_vars['v_jp1_username'].get()
    
    _log(f"[Params] AJS Path: {ajs_path_input}")
    _log(f"[Params] Bank: {bank}")

    # パス正規化ロジック
    def normalize_unit_path(absolute_path, base_to_remove):
        path = re.sub(r'^[a-zA-Z0-9_]+:', '', absolute_path)
        if base_to_remove and base_to_remove != "/" and path.startswith(base_to_remove):
            path = path.replace(base_to_remove, '', 1)
        return path

    try:
        if not target_files_str:
            raise ValueError("作成したいファイルパスを入力してください。")
        
        target_files = [f.strip() for f in target_files_str.split('\n') if f.strip()]
        if not target_files:
            raise ValueError("ファイルパスが有効ではありません。")
        
        _log(f"[Params] Target Files ({len(target_files)}): {target_files}")

        # --- 0. 初期準備 ---
        ts = time.strftime("%Y%m%d%H%M%S")
        base_dir = pathlib.Path(sys.argv[0]).resolve().parent
        out_dir = base_dir / DIR_NAME_DEP / ts 
        out_dir.mkdir(parents=True, exist_ok=True)
        
        # ★追加: tmpフォルダ作成
        tmp_dir = out_dir / "tmp"
        tmp_dir.mkdir(exist_ok=True)
        
        _log(f"[Info] Output Directory: {out_dir}")

        # --- 1. 解析実行 ---
        update_status("I/O解析実行中...", 10)
        # out_dirを渡すと、ajs_inout_logic側で自動的に tmp/ajs_out_raw.txt を作成・使用してくれます
        final_records, _ = analyze_ajs_jobs(gui_vars, gui_funcs, out_dir, use_cache=True)
        
        producer_map = collections.defaultdict(set)
        for record in final_records:
            for output_path in record.get('outputs', []):
                producer_map[output_path].add(record['unit_full'])
        
        _log(f"[Info] Producer Map built. Total files tracked: {len(producer_map)}")

        # --- 2. 逆引きトレース (BFS) ---
        update_status("逆引き探索 (BFS) 実行中...", 40)
        queue = collections.deque(target_files)
        visited_files = set(target_files) 
        needed_units_full = set() # ここには %JN (絶対パス) が入る
        external_inputs = set()
        
        trace_log = [] 

        while queue:
            current_file = queue.popleft()
            producers = producer_map.get(current_file)
            
            if not producers:
                external_inputs.add(current_file)
                trace_log.append(f"File: {current_file} -> [External]")
                continue
            
            for unit_full_path in producers:
                if unit_full_path in needed_units_full: continue
                needed_units_full.add(unit_full_path)
                trace_log.append(f"File: {current_file} -> Job: {unit_full_path}")
                
                record = next((r for r in final_records if r['unit_full'] == unit_full_path), None)
                if record:
                    for input_file in record.get('inputs', []):
                        if input_file and input_file not in visited_files:
                            visited_files.add(input_file)
                            queue.append(input_file)
        
        _log(f"[Trace] Found {len(needed_units_full)} units, {len(external_inputs)} external inputs.")

        # --- 3. 定義フィルタリングの準備 ---
        update_status("定義フィルタリング準備...", 60)
        
        path_part = ajs_path_input.split(':', 1)[1] if ':' in ajs_path_input else ajs_path_input
        base_dir_to_remove = os.path.dirname(path_part) 
        
        normalized_need_set = set()
        
        def add_parents(path_set, n_path):
            anc = n_path
            while '/' in anc.strip('/'):
                anc = '/' + '/'.join(anc.strip('/').split('/')[:-1])
                if not anc.strip('/'): break
                path_set.add(anc)
        
        for unit_full in needed_units_full:
            norm_path = normalize_unit_path(unit_full, base_dir_to_remove)
            if norm_path.strip('/'):
                normalized_need_set.add(norm_path)
                add_parents(normalized_need_set, norm_path)

        # --- 4. 定義・関連情報の取得 ---
        update_status("AJS定義・関連情報を取得中...", 70)
        
        export_list = []
        export_list.append(f'export JP1_HOSTNAME={shlex.quote(jp1_hostname)}')
        export_list.append(f'export JP1_USERNAME={shlex.quote(jp1_username)}')
        env_str = ' && '.join(export_list)
        
        remote_def = f"/tmp/ajs_def_dep_{ts}.txt"
        # ★修正: 保存先を tmp 配下へ
        local_def = tmp_dir / "ajs_definition_original.txt"
        cmd_def = f'{env_str} && {ajs_print_path} -s yes -a {shlex.quote(ajs_path_input)} > {remote_def}'
        
        remote_dep = f"/tmp/ajs_dep_dep_{ts}.txt"
        # ★修正: 保存先を tmp 配下へ
        local_dep = tmp_dir / "ajs_graph.txt"
        cmd_dep = f'{env_str} && {ajs_print_path} -F AJSROOT1 -f %TY%t%JN%t%ar -R {shlex.quote(ajs_path_input)} > {remote_dep}'
        
        read_enc = ENC[gui_vars['v_t5_out_c'].get()] 
        if read_enc == 'utf-8': read_enc = 'cp932'

        _log(f"[Command-Def] {cmd_def}")
        _log(f"[Command-Rel] {cmd_dep}")

        with get_ssh_client() as ssh:
            ch = ssh.get_transport().open_session()
            ch.exec_command(cmd_def.encode(read_enc))
            if ch.recv_exit_status() != 0:
                err = ch.makefile_stderr().read().decode(read_enc,'ignore')
                _log(f"[Error] Def command: {err}")
                raise RuntimeError(f"AJS定義取得エラー: {err}")
            
            ch2 = ssh.get_transport().open_session()
            ch2.exec_command(cmd_dep.encode(read_enc))
            if ch2.recv_exit_status() != 0:
                err = ch2.makefile_stderr().read().decode(read_enc,'ignore')
                _log(f"[Error] Rel command: {err}")
                raise RuntimeError(f"AJS関連取得エラー: {err}")

            sftp = ssh.open_sftp()
            sftp.get(str(remote_def), str(local_def))
            sftp.get(str(remote_dep), str(local_dep))
            sftp.close()
            ssh.exec_command(f"rm -f {remote_def} {remote_dep}")
            
        ajs_def_txt = local_def.read_text(encoding=read_enc, errors='ignore')
        ajs_dep_txt = local_dep.read_text(encoding=read_enc, errors='ignore')
        
        # --- 5. グラフ構築とフィルタリング ---
        update_status("グラフ構築と再結線処理...", 90)
        
        G = pre_parse_graph(ajs_dep_txt, base_dir_to_remove)
        rec_txt = pre_filter_definition(ajs_def_txt, normalized_need_set, G)

        # --- 6. 出力 (成果物はルートへ) ---
        out_file_def = out_dir / 'recovery_definition.txt'
        out_file_def.write_text(rec_txt, encoding=ENC[gui_vars['v_t5_out_c'].get()], newline=NL[gui_vars['v_t5_out_n'].get()])

        out_file_ext = out_dir / 'missing_files.txt'
        ext_list = sorted(list(external_inputs))
        if ext_list:
            out_file_ext.write_text("\n".join(ext_list), encoding='utf-8')
        
        # GUI更新
        text_box.delete('1.0', 'end')
        if needed_units_full:
            text_box.insert('end', f"--- 抽出ジョブ ({len(needed_units_full)}件) ---\n" + "\n".join(sorted(list(needed_units_full))))
            text_box.insert('end', f"\n\n--- 外部入力ファイル(欠落ファイル) ({len(ext_list)}件) ---\n" + ("\n".join(ext_list) if ext_list else "なし"))
        else:
            text_box.insert('end', "探索結果: 該当ジョブなし")

        write_detail_log({
            "target_files": target_files,
            "producer_map": producer_map,
            "needed_units_full": needed_units_full,
            "normalized_need_set": normalized_need_set,
            "trace_log": trace_log
        })

        update_status("完了", 100)
        save_hist()
        _log("[Success] Completed.")
        show_info(f"解析完了: {out_dir}")
        
    except Exception as e:
        tb = traceback.format_exc()
        _log(f"[Exception] {str(e)}\n{tb}")
        show_error(str(e))
    finally:
        update_status("待機中", 0)

# --- エントリーポイント ---
def open_t5_job_runner(gui_vars, gui_funcs, text_box):
    gui_funcs['run_in_thread'](dep_start_job)(gui_vars, gui_funcs, text_box)
