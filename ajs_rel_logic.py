#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 先行関係解析ロジック (Predecessor Logic)
v5.6 (2025-11-23) - 中間ファイルをtmpフォルダに格納するように修正
"""

import re
import os
import sys
import time
import shlex
import pathlib
import networkx as nx
import json 
import collections
import traceback
import datetime

# 定数ファイルをインポート
from ajs_constants import ENC, NL, LOG_DIR, DIR_NAME_PRE

# ログファイルパス
LOG_FILE_RUN = LOG_DIR / "tab4_pre_run.log"
LOG_FILE_DETAIL = LOG_DIR / "tab4_pre_details.json"

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
        # Set型をListに変換してJSON化
        serializable = {}
        for k, v in data_dict.items():
            if isinstance(v, set):
                serializable[k] = sorted(list(v))
            elif isinstance(v, list):
                serializable[k] = v
            else:
                serializable[k] = str(v)
                
        with open(LOG_FILE_DETAIL, "w", encoding="utf-8") as f:
            json.dump(serializable, f, indent=2, ensure_ascii=False)
        _log(f"[Info] Detail log saved: {LOG_FILE_DETAIL}")
    except Exception as e:
        _log(f"[Error] Failed to write detail log: {e}")

def pre_normalize(path: str, base: str) -> str:
    """AJSパスを正規化する"""
    p = re.sub(r'^[a-zA-Z0-9_]+:', '', path.strip())
    if base and base != "/" and p.startswith(base):
        p = p.replace(base, '', 1)
    return p

def pre_parse_graph(txt: str, base: str):
    """3つ組 (From, To, Type) 対応のグラフ構築"""
    G = nx.DiGraph()
    unit_types = {x.strip() for x in "mgroup,group,mnet,condn,net,rnet,rmnet,rrnet,job,rjob,pjob,rpjob,qjob,rqjob,jdjob,rjdjob,orjob,rorjob,fxjob,rfxjob,netcn".split(',')}
    
    current_parent = None
    parse_stats = {"nodes": 0, "edges": 0}

    for ln in txt.splitlines():
        cols = ln.split('\t')
        if not cols: continue
        
        if cols[0] in unit_types and len(cols) >= 2:
            current_parent = pre_normalize(cols[1], base)
            if current_parent.strip('/'):
                G.add_node(current_parent)
                parse_stats["nodes"] += 1
            rels_str = cols[2] if len(cols) >= 3 else ""
        else:
            rels_str = cols[0]
        
        if not current_parent or not rels_str:
            continue

        parts = [s.strip() for s in rels_str.split(',') if s.strip()]
        for i in range(0, len(parts) - 1, 3):
            from_name = parts[i]
            to_name = parts[i+1]
            
            if from_name == '-' or to_name == '-': continue

            def join_path(parent, child):
                return f"{parent.rstrip('/')}/{child}".replace("//", "/")

            from_path = join_path(current_parent, from_name)
            to_path = join_path(current_parent, to_name)

            if from_path.strip('/') and to_path.strip('/'):
                G.add_node(from_path)
                G.add_node(to_path)
                G.add_edge(from_path, to_path)
                parse_stats["edges"] += 1

    _log(f"[Graph] Parsed {parse_stats['nodes']} nodes, {parse_stats['edges']} edges.")
    return G

def pre_descendants(G: nx.DiGraph, n):
    """指定ノード配下の子孫ユニットをすべて列挙する"""
    prefix = n.rstrip('/') + '/'
    yield from (x for x in G.nodes if x.startswith(prefix) and x != n)

def pre_compute_need(G: nx.DiGraph, target: str):
    """指定ユニットの実行に必要な先行ユニット・親・子孫をすべて計算する"""
    need = set()
    q = [(target, 'seed')]
    
    while q:
        n, org = q.pop()
        
        if not n.strip('/') or n in need: continue
        
        if n not in G: G.add_node(n) 
        need.add(n)
        
        # 1. 先行ユニット
        if n in G:
            q.extend((p, 'pre') for p in G.predecessors(n))
        
        # 2. 親ユニット
        anc = n
        while '/' in anc.strip('/'):
            anc = '/' + '/'.join(anc.strip('/').split('/')[:-1])
            if not anc.strip('/'): break
            q.append((anc, 'par'))
            
        # 3. 子孫ユニット
        if org in ('seed', 'pre'):
            q.extend((d, 'desc') for d in pre_descendants(G, n))
            
    return need

def find_bridged_successors(G: nx.DiGraph, start_node: str, valid_siblings: set):
    """グラフ探索: 直近の有効な兄弟を探す"""
    targets = set()
    visited = set()
    queue = [start_node]
    visited.add(start_node)
    
    while queue:
        curr = queue.pop(0)
        if curr not in G: continue
        
        for succ in G.successors(curr):
            if succ in visited: continue
            visited.add(succ)
            
            if succ in valid_siblings:
                targets.add(succ) 
            else:
                queue.append(succ)
    return targets

def generate_ar_lines(G: nx.DiGraph, siblings: list, parent_path: str, indent_level: int):
    """兄弟間の先行関係を計算して ar行を生成"""
    lines = []
    valid_siblings_full = set()
    for s in siblings:
        full = f"{parent_path}/{s}".replace("//", "/")
        valid_siblings_full.add(full)
        
    indent_str = "\t" * indent_level
        
    for s in siblings:
        start_node = f"{parent_path}/{s}".replace("//", "/")
        if start_node not in G: continue

        next_nodes_full = find_bridged_successors(G, start_node, valid_siblings_full)
        
        for next_node in next_nodes_full:
            next_name = os.path.basename(next_node)
            lines.append(f"{indent_str}ar=(f={s},t={next_name},seq);")
            
    return lines

def build_hierarchy_map(need_set):
    h_map = collections.defaultdict(list)
    for path in need_set:
        if path == "/": continue
        parent_dir = os.path.dirname(path)
        basename = os.path.basename(path)
        if parent_dir == "/": parent_dir = ""
        h_map[parent_dir].append(basename)
    return h_map

def pre_filter_definition(txt: str, need: set, G: nx.DiGraph = None):
    """AJS定義から必要なユニットを抽出し、ar行のみ再構築する"""
    out = []
    stack = [] 
    skip = False
    sk_ind = None
    depth = 0
    
    normalized_need = {p.rstrip('/') for p in need}
    hierarchy_map = build_hierarchy_map(normalized_need)
    ar_written_map = {}

    for ln in txt.splitlines():
        st = ln.lstrip()
        curr_indent = ln.count('\t')

        if skip:
            depth += ln.count('{') - ln.count('}')
            if st.startswith('}') and curr_indent == sk_ind and depth == 0:
                skip = False
                while stack and stack[-1][0] >= sk_ind:
                    stack.pop()
            continue
        
        if st.startswith('unit='):
            ind = curr_indent
            while stack and stack[-1][0] >= ind:
                stack.pop()
            
            name = st.split('=', 1)[1].split(',', 1)[0].strip()
            stack.append((ind, name))
            path = '/' + '/'.join(n for _, n in stack)
            
            if path not in normalized_need:
                skip = True
                sk_ind = ind
                depth = 0
            else:
                out.append(ln)
            continue
        
        if st.startswith('ar='):
            if G is None:
                out.append(ln)
            else:
                if stack:
                    parent_path = '/' + '/'.join(n for _, n in stack)
                    if not ar_written_map.get(parent_path):
                        children = hierarchy_map[parent_path]
                        if children:
                            new_ars = generate_ar_lines(G, children, parent_path, curr_indent)
                            out.extend(new_ars)
                        ar_written_map[parent_path] = True
                continue

        out.append(ln)
            
    return '\n'.join(out)

def pre_start_job(gui_vars, gui_funcs, text_box):
    """先行関係解析のメイン処理"""
    update_status = gui_funcs['update_status']
    get_ssh_client = gui_funcs['get_ssh_client']
    save_hist = gui_funcs['save_hist']
    show_info = gui_funcs['show_info']
    show_error = gui_funcs['show_error']
    
    # ログ開始
    with open(LOG_FILE_RUN, "w", encoding="utf-8") as f:
        f.write(f"=== Tab 4 (Predecessor) Execution Start: {datetime.datetime.now()} ===\n")

    v_root = gui_vars['v_pre_root']
    v_tgt = gui_vars['v_pre_tgt']
    v_srv_c = gui_vars['v_srv_c']
    v_out_c = gui_vars['v_pre_out_c']
    v_out_n = gui_vars['v_pre_out_n']
    
    ajs_print_path = gui_vars['v_ajs_print_path'].get()
    jp1_hostname = gui_vars['v_jp1_hostname'].get()
    jp1_username = gui_vars['v_jp1_username'].get()
    
    # パラメータログ
    _log(f"[Params] Root: {v_root.get()}")
    _log(f"[Params] Target: {v_tgt.get()}")

    export_list = []
    export_list.append(f'export JP1_HOSTNAME={shlex.quote(jp1_hostname)}')
    export_list.append(f'export JP1_USERNAME={shlex.quote(jp1_username)}')
    env_str = ' && '.join(export_list)

    try:
        ajs_root, tgt_in = v_root.get().strip(), v_tgt.get().strip()
        if not all([ajs_root, tgt_in]):
            raise ValueError("AJSパスと解析対象ユニットパスを入力してください。")
        
        base_path_for_norm = os.path.dirname(ajs_root.split(':', 1)[-1])
        _log(f"[Info] Base Path for normalization: {base_path_for_norm}")
        
        ts = time.strftime('%Y%m%d%H%M%S')
        out_dir = pathlib.Path(sys.argv[0]).resolve().parent / DIR_NAME_PRE / ts 
        out_dir.mkdir(parents=True, exist_ok=True)
        
        # ★追加: tmpフォルダ作成
        tmp_dir = out_dir / "tmp"
        tmp_dir.mkdir(exist_ok=True)
        
        _log(f"[Info] Output Directory: {out_dir}")
        
        srv_enc = ENC[v_srv_c.get()]
        
        with get_ssh_client() as ssh:
            sftp = ssh.open_sftp()
            update_status("SSH 接続...", 10)
            tmp_def, tmp_dep = f'/tmp/def_{ts}.txt', f'/tmp/dep_{ts}.txt'
            
            update_status("ajsprint (1/2) 実行...", 25)
            cmd1 = f'{env_str} && {ajs_print_path} -s yes -a {shlex.quote(ajs_root)} > {tmp_def}'
            _log(f"[Command-Def] {cmd1}")
            
            ch1 = ssh.get_transport().open_session()
            ch1.exec_command(cmd1.encode(srv_enc))
            if ch1.recv_exit_status() != 0:
                err = ch1.makefile_stderr().read().decode(srv_enc, 'ignore')
                _log(f"[Error] Def command failed: {err}")
                raise RuntimeError(f"定義取得エラー:\n{err}")
            
            update_status("ajsprint (2/2) 実行...", 50)
            cmd2 = f'{env_str} && {ajs_print_path} -F AJSROOT1 -f %TY%t%JN%t%ar -R {shlex.quote(ajs_root)} > {tmp_dep}'
            _log(f"[Command-Rel] {cmd2}")
            
            ch2 = ssh.get_transport().open_session()
            ch2.exec_command(cmd2.encode(srv_enc))
            if ch2.recv_exit_status() != 0:
                err = ch2.makefile_stderr().read().decode(srv_enc, 'ignore')
                _log(f"[Error] Rel command failed: {err}")
                raise RuntimeError(f"関連取得エラー:\n{err}")
            
            update_status("ファイル取得...", 70)
            # ★修正: tmpフォルダに保存
            def_file = tmp_dir / 'ajs_definition_original.txt'
            dep_file = tmp_dir / 'ajs_graph.txt'
            
            sftp.get(tmp_def, str(def_file))
            sftp.get(tmp_dep, str(dep_file))
            ssh.exec_command(f"rm -f {tmp_def} {tmp_dep}")
            sftp.close()
            
        update_status("解析...", 80)
        def_txt = def_file.read_text(encoding=srv_enc, errors='ignore')
        dep_txt = dep_file.read_text(encoding=srv_enc, errors='ignore')
        
        # グラフ構築
        G = pre_parse_graph(dep_txt, base_path_for_norm)
        
        # 必要ユニット計算 (子孫含む)
        need = pre_compute_need(G, pre_normalize(tgt_in, base_path_for_norm))
        _log(f"[Info] Need set count: {len(need)}")
        
        # 詳細ログ出力
        write_detail_log({
            "need_set": need,
            "graph_nodes_count": len(G.nodes),
            "graph_edges_count": len(G.edges)
        })
        
        update_status("回復定義生成(再結線)...", 90)
        rec_txt = pre_filter_definition(def_txt, need, G)
        
        # 成果物はルートへ
        out_file = out_dir / 'recovery_definition.txt'
        out_file.write_text(rec_txt, encoding=ENC[v_out_c.get()], newline=NL[v_out_n.get()])
        
        text_box.delete('1.0', 'end')
        text_box.insert('end', '\n'.join(sorted(need)))
        
        update_status("完了", 100)
        save_hist()
        _log("[Success] Completed.")
        show_info(f"回復用定義を生成しました。\n出力フォルダ:\n{out_dir}")
        
    except Exception as e:
        tb = traceback.format_exc()
        _log(f"[Exception] {str(e)}\n{tb}")
        show_error(str(e))
    finally:
        update_status("待機中", 0)
