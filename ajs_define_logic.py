#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 定義回復ロジック (Define/Recover Logic)
v3.2 (2025-11-23) - SVNAME削除, OS改行設定削除(LF固定), ログ出力強化
"""

import time
import shlex
import traceback
import datetime
from ajs_constants import ENC, LOG_DIR

# ログファイルパス
LOG_FILE = LOG_DIR / "tab2_define_debug.log"

def _log(msg):
    """ログファイルへの書き込みヘルパー"""
    try:
        timestamp = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {msg}\n")
    except: pass

def define_convert_newlines(local_path, target_enc):
    """
    ローカルの定義ファイルを読み込み、改行コードをLF(\n)に統一する。
    ※AJSの定義ファイルはUNIX/Windows問わずLFが標準的
    """
    content = None
    try:
        with open(local_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        try:
            with open(local_path, 'r', encoding='cp932') as f:
                content = f.read()
        except Exception as e:
            raise IOError(f"ファイルの読み込みに失敗しました (UTF-8, SJIS)。\n{e}")
    except Exception as e:
        raise IOError(f"ファイルを開けません。\n{e}")
    
    # 改行コードを正規化 (\r\n -> \n, \r -> \n)
    content = content.replace('\r\n', '\n').replace('\r', '\n')
    return content 

def define_start_job(gui_vars, gui_funcs):
    """定義回復のメイン処理"""
    
    update_status = gui_funcs['update_status']
    get_ssh_client = gui_funcs['get_ssh_client']
    save_hist = gui_funcs['save_hist']
    show_info = gui_funcs['show_info']
    show_error = gui_funcs['show_error']
    
    # ログ開始
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"=== Tab 2 (Define) Execution Start: {datetime.datetime.now()} ===\n")

    try:
        # --- パラメータ取得 ---
        v_file_path = gui_vars['v_recover_file']
        v_unit_name = gui_vars['v_recover_unit']
        v_srv_c = gui_vars['v_srv_c'] # 共通設定
        ajs_define_path = gui_vars['v_ajs_define_path'].get()

        jp1_hostname = gui_vars['v_jp1_hostname'].get()
        jp1_username = gui_vars['v_jp1_username'].get()
        
        # ログ出力
        _log(f"[Params] Local File: {v_file_path.get()}")
        _log(f"[Params] Target Unit: {v_unit_name.get()}")
        _log(f"[Params] Server Enc: {v_srv_c.get()}")

        # 環境変数 (SVNAME削除)
        export_list = []
        export_list.append(f'export JP1_HOSTNAME={shlex.quote(jp1_hostname)}')
        export_list.append(f'export JP1_USERNAME={shlex.quote(jp1_username)}')
        EXPORT_ENV = ' && '.join(export_list)

        update_status("処理開始...", 0)
        local_path = v_file_path.get()
        full_ajs_path = v_unit_name.get().strip()
        
        if not local_path or not full_ajs_path:
            raise ValueError("回復用AJS定義と回復先AJSパスを指定してください。")
        if ':' not in full_ajs_path:
            raise ValueError("回復先AJSパスにはAJSルートパス（例: AJSROOT1:/...）も含めて指定してください。")
        
        ajs_root, job_path = full_ajs_path.split(':', 1)
        target_enc = ENC[v_srv_c.get()]
        
        update_status("ファイル変換中...", 20)
        # 改行コードはLF固定
        converted_content = define_convert_newlines(local_path, target_enc)
        _log("[Info] File read and converted to LF.")
        
        update_status("SSH 接続...", 30)
        with get_ssh_client() as ssh:
            sftp = ssh.open_sftp()
            remote_path = f"/tmp/define_{time.strftime('%Y%m%d%H%M%S')}.txt"
            
            update_status("ファイルアップロード中...", 50)
            _log(f"[Info] Uploading to {remote_path}...")
            with sftp.open(remote_path, 'wb') as f:
                f.write(converted_content.encode(target_enc))
            
            update_status("ajsdefine 実行中...", 70)
            # コマンド構築
            cmd = (f'{EXPORT_ENV} && {ajs_define_path} -f -p -F {shlex.quote(ajs_root)} '
                   f'-d {shlex.quote(job_path)} {shlex.quote(remote_path)}')
            
            _log(f"[Command] {cmd}")
            
            ch = ssh.get_transport().open_session()
            ch.exec_command(cmd.encode(target_enc))
            if ch.recv_exit_status() != 0:
                err_msg = ch.makefile_stderr().read().decode(target_enc, 'ignore')
                _log(f"[Error] Command failed: {err_msg}")
                raise RuntimeError(f"ajsdefine 実行エラー:\n{err_msg}")
            
            sftp.remove(remote_path)
            sftp.close()
            
        update_status("完了", 100)
        save_hist()
        _log("[Success] Job completed successfully.")
        show_info("ジョブ定義の回復が正常に完了しました。")
        
    except Exception as e:
        err_detail = str(e)
        tb = traceback.format_exc()
        _log(f"[Exception] {err_detail}\n{tb}")
        show_error(err_detail)
    finally:
        update_status("待機中", 0)
