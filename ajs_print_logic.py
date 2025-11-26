#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 定義取得＆変換ロジック (Print Logic)
v3.5 (2025-11-23) - フォルダ構成変更 (中間ファイルはtmpへ、成果物はルートへ)
"""

import os
import sys
import time
import shlex
import csv
import codecs
import pathlib
import traceback
import datetime

from ajs_constants import ENC, NL, PARAM_FILE, DIR_NAME_PRINT, LOG_DIR

LOG_FILE = LOG_DIR / "tab1_print_debug.log"

def _log(msg):
    try:
        timestamp = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f: f.write(f"[{timestamp}] {msg}\n")
    except: pass

def print_load_prm(path):
    if not os.path.exists(path): raise FileNotFoundError(f"パラメータファイルが見つかりません: {path}")
    mp = {}
    with codecs.open(path, "r", "cp932") as f:
        rdr = csv.reader(f)
        next(rdr, None) 
        for row in rdr:
            if len(row) < 4: continue
            bank, prod, mir, dev = [c.strip() for c in row[:4]]
            mp.setdefault(bank, []).append({'prod': prod, 'mir': mir, 'dev': dev})
    return mp

def print_build_table(mapping, bank, detail):
    tbl = []
    details_map = {
        "本番⇒ミラー": ("prod", "mir"), "本番⇒開発": ("prod", "dev"),
        "ミラー⇒本番": ("mir", "prod"), "ミラー⇒開発": ("mir", "dev"),
        "開発⇒本番": ("dev", "prod"), "開発⇒ミラー": ("dev", "mir")
    }
    src_key, dst_key = details_map.get(detail, (None, None))
    if not src_key: return [] 
    for parts in mapping.get(bank, []):
        src_val, dst_val = parts.get(src_key), parts.get(dst_key)
        if src_val and dst_val: tbl.append((src_val, dst_val))
    return tbl

def print_start_job(gui_vars, gui_funcs):
    update_status = gui_funcs['update_status']
    get_ssh_client = gui_funcs['get_ssh_client']
    save_hist = gui_funcs['save_hist']
    show_info = gui_funcs['show_info']
    show_error = gui_funcs['show_error']
    
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"=== Tab 1 (Print) Execution Start: {datetime.datetime.now()} ===\n")

    try:
        v_ajs_path = gui_vars['v_print_ajs_path']
        v_def_kind = gui_vars['v_print_kind']
        v_srv_c = gui_vars['v_srv_c']
        v_conv_flg = gui_vars['v_print_conv_flg']
        v_bank_sel = gui_vars['v_print_bank']
        v_detail_sel = gui_vars['v_print_detail']
        v_out_c = gui_vars['v_print_out_c']
        v_out_n = gui_vars['v_print_out_n']
        
        ajs_print_path = gui_vars['v_ajs_print_path'].get()
        custom_pairs = gui_vars.get('v_print_custom_pairs', [])

        jp1_hostname = gui_vars['v_jp1_hostname'].get()
        jp1_username = gui_vars['v_jp1_username'].get()
        
        _log(f"[Params] AJS Path: {v_ajs_path.get()}")
        _log(f"[Params] Kind: {v_def_kind.get()}, Convert: {v_conv_flg.get()}")

        export_list = []
        export_list.append(f'export JP1_HOSTNAME={shlex.quote(jp1_hostname)}')
        export_list.append(f'export JP1_USERNAME={shlex.quote(jp1_username)}')
        EXPORT_ENV = ' && '.join(export_list)

        update_status("処理開始...", 0)
        ajs_path = v_ajs_path.get().strip()
        if not ajs_path: raise ValueError("AJS パスを入力してください。")
        
        ts = time.strftime("%Y%m%d%H%M%S")
        base_dir = pathlib.Path(sys.argv[0]).resolve().parent
        out_dir = base_dir / DIR_NAME_PRINT / ts
        out_dir.mkdir(parents=True, exist_ok=True)
        
        # ★追加: tmpフォルダ作成
        tmp_dir = out_dir / "tmp"
        tmp_dir.mkdir(exist_ok=True)
        
        _log(f"[Info] Output Directory: {out_dir}")

        cmd_list = []
        if v_def_kind.get() in ("recover", "both"):
            remote_path = f"/tmp/AJS_recover_{ts}.txt"
            cmd_str = f'{EXPORT_ENV} && {ajs_print_path} -s yes -a {shlex.quote(ajs_path)} > {remote_path}'
            cmd_list.append(("AJS_recover.txt", remote_path, cmd_str))
            _log(f"[Command-Recover] {cmd_str}")
        
        if v_def_kind.get() in ("verify", "both"):
            remote_path = f"/tmp/AJS_verify_{ts}.txt"
            fmt = "-F AJSROOT1 -f %TY%t%JN%t%jn%t%sc%t%Te%t%pm%t%cm%t%En%t%ev%t%wk%t%rh%t%un%t%wt%t%Th%t%eU%t%FF%t%FC%t%FI%t%FO%t%ud%t%Ed%t%rg%t%pr%t%si%t%so%t%oa%t%se%t%ea%t%de%t%ha%t%Sd%t%St%t%ed%t%cy%t%sh%t%hd%t%sy%t%ey%t%jc%t%ms%t%mp%t%ow%t%gr%t%Rh%t%EI%t%eI"
            cmd_str = f'{EXPORT_ENV} && {ajs_print_path} {fmt} -R {shlex.quote(ajs_path)} > {remote_path}'
            cmd_list.append(("AJS_verify.txt", remote_path, cmd_str))
            _log(f"[Command-Verify] {cmd_str}")
        
        update_status("SSH 接続...", 10)
        srv_enc = ENC[v_srv_c.get()]
        
        with get_ssh_client() as ssh:
            sftp = ssh.open_sftp()
            local_files = []
            for i, (fname, rtmp, cmd) in enumerate(cmd_list):
                progress_val = 20 + i * 30
                update_status(f"{fname} 取得中...", progress_val)
                ch = ssh.get_transport().open_session()
                ch.exec_command(cmd.encode(srv_enc))
                if ch.recv_exit_status() != 0:
                    err_msg = ch.makefile_stderr().read().decode(srv_enc,'ignore')
                    _log(f"[Error] Command failed: {err_msg}")
                    raise RuntimeError(f"コマンド実行エラー:\n{err_msg}")
                
                # ★修正: 変換ありならtmpへ、なしならrootへ保存
                if v_conv_flg.get() == "yes":
                    lpath = tmp_dir / fname
                else:
                    lpath = out_dir / fname
                    
                update_status(f"{fname} ダウンロード中...", progress_val + 15)
                sftp.get(rtmp, str(lpath))
                local_files.append(lpath)
                try: ssh.exec_command(f"rm -f {rtmp}")
                except: pass
            sftp.close()

            out_c, out_n = ENC[v_out_c.get()], NL[v_out_n.get()]
            prm_path = base_dir / PARAM_FILE
            
            if v_conv_flg.get() == "yes":
                update_status("変換処理中...", 80)
                _log("[Info] Starting conversion...")
                table = []
                detail_selection = v_detail_sel.get()
                if detail_selection == "カスタム":
                    if not custom_pairs: raise ValueError("カスタム変換が選択されましたが、詳細が1件も登録されていません。")
                    table = custom_pairs
                else:
                    mapping = print_load_prm(prm_path)
                    table = print_build_table(mapping, v_bank_sel.get(), detail_selection)

                if not table:
                    update_status("変換テーブルが空のためスキップ", 85)
                    _log("[Info] Conversion table is empty. Skipping.")
                else:
                    for path in local_files:
                        txt = path.read_text(encoding=srv_enc, errors='ignore')
                        for src, dst in table:
                            if src: txt = txt.replace(src, dst)
                        
                        # ★修正: 変換後のファイルは常にルート(out_dir)に保存
                        conv_fname = f"{path.stem}_converted.txt"
                        conv_path = out_dir / conv_fname
                        
                        conv_path.write_text(txt, encoding=out_c, newline=out_n, errors='ignore')
                        _log(f"[Info] Converted file saved: {conv_path.name}")
        
        update_status("完了", 100)
        save_hist()
        _log("[Success] Job completed successfully.")
        show_info(f"結果を {out_dir} に出力しました。")
        
    except Exception as e:
        err_detail = str(e)
        tb = traceback.format_exc()
        _log(f"[Exception] {err_detail}\n{tb}")
        show_error(err_detail)
    finally:
        update_status("待機中", 0)
