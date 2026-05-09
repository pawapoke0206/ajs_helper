#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJS Helper Tool - 依存関係解析ロジック (Dependency Logic)

依存関係解析のオーケストレーション（全体の流れの制御）を担当する。
探索ロジックは ajs_dep_tracer.py、Excel出力は ajs_dep_report.py に分離。

このファイルの責務:
  - dep_start_job(): GUI入力の受け取り → 各ステップの呼び出し → 結果出力
  - GUIダイアログ（ユーザー選択）
  - SSH経由のデータ取得
  - ブリッジ定義の読み込み・解決
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
from ajs_constants import ENC, NL, LOG_DIR, DIR_NAME_DEP, DB_BRIDGE_FILE
from ajs_utils import make_logger, pre_normalize, write_detail_log as _write_detail_log
from ajs_rel_logic import pre_filter_definition, pre_parse_graph, pre_compute_need
from ajs_inout_logic import analyze_ajs_jobs

# 分割モジュール
from ajs_dep_tracer import (
    _get_ancestors, discover_parent_jobnets, find_parent_jobnet,
    filter_dep_text_by_jobnet, find_producers_across_jobnets,
    filter_producers_by_graph, _bfs_trace, _supplement_check,
)
from ajs_dep_report import write_dependency_report

# ログファイルパス
LOG_FILE_RUN = LOG_DIR / "depend_run.log"
LOG_FILE_DETAIL = LOG_DIR / "depend_details.json"

_log = make_logger(LOG_FILE_RUN)


# =====================================================================
# GUIダイアログ: ユーザー選択
# =====================================================================

def _ask_user_producer_choice(items, title, description, root_window=None, warning=None):
    """ユニット/親JNの選択をユーザーに求める汎用ダイアログ。

    3つのケースで共通利用する:
      1. DFS中の外部ファイル: 複数親JNまたは同一親JN内の複数producer
      2. 充足チェック: 複数producerから選択
      3. BFS内: 並列producer（ジョブ設計の問題警告付き）

    Args:
        items: 選択対象のリスト。各要素は:
            {
              'file': ファイルパス,
              'consumer': 使用元ユニットのフルパス（表示用、省略可）,
              'candidates': [producer_unit_full, ...],  # 選択肢
            }
        title: ダイアログのタイトル
        description: 説明テキスト
        root_window: 親ウィンドウ（Noneなら自動選択）
        warning: 警告テキスト（ケース3用、Noneなら非表示）

    Returns:
        dict: {ファイルパス: 選択されたproducer_unit_full}
    """
    import tkinter as tk

    result = {}
    if not items:
        return result

    if root_window is None:
        for item in items:
            result[item['file']] = sorted(item['candidates'])[0]
        return result

    dialog = tk.Toplevel(root_window)
    dialog.title(title)
    # 画面幅の80%を使い、最低900px確保（長いパスが切れないように）
    try:
        screen_w = root_window.winfo_screenwidth()
        dlg_w = max(900, int(screen_w * 0.8))
    except Exception:
        dlg_w = 900
    dialog.geometry(f"{dlg_w}x500")
    dialog.transient(root_window)
    dialog.grab_set()
    dialog.configure(bg="#282a36")

    # 説明
    tk.Label(dialog, text=description,
             bg="#282a36", fg="#f0f0f5", font=("", 10),
             wraplength=dlg_w - 40, justify="left", pady=8).pack(fill="x", padx=16)

    # 警告（ケース3: ジョブ設計問題）
    if warning:
        tk.Label(dialog, text=warning,
                 bg="#282a36", fg="#ff6e6e", font=("", 9, "bold"),
                 wraplength=dlg_w - 40, justify="left", pady=4).pack(fill="x", padx=16)

    # スクロール可能なフレーム（水平スクロールも追加）
    scroll_outer = tk.Frame(dialog, bg="#282a36")
    scroll_outer.pack(fill="both", expand=True, padx=16, pady=(0, 8))

    canvas = tk.Canvas(scroll_outer, bg="#282a36", highlightthickness=0)
    v_scrollbar = tk.Scrollbar(scroll_outer, orient="vertical", command=canvas.yview)
    h_scrollbar = tk.Scrollbar(scroll_outer, orient="horizontal", command=canvas.xview)
    scroll_frame = tk.Frame(canvas, bg="#282a36")

    scroll_frame.bind("<Configure>",
                      lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

    # canvasの幅をダイアログに追従させる
    def _on_canvas_configure(event):
        canvas.itemconfig(canvas_window, width=event.width)
    canvas.bind("<Configure>", _on_canvas_configure)

    # マウスホイール対応
    def _on_mousewheel(event):
        if event.delta:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    dialog.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

    v_scrollbar.pack(side="right", fill="y")
    h_scrollbar.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)

    # 同じファイルが複数consumerで出る場合があるので、インデックスで管理
    selection_vars = {}  # {index: StringVar}

    for i, item in enumerate(items):
        if i > 0:
            tk.Frame(scroll_frame, bg="#44475a", height=1).pack(fill="x", pady=8)

        file_path = item['file']
        consumer = item.get('consumer', '')
        candidates = item['candidates']

        tk.Label(scroll_frame, text=f"ファイル: {os.path.basename(file_path)}",
                 bg="#282a36", fg="#7aa2f7", font=("", 10, "bold"),
                 anchor="w").pack(fill="x", padx=4)
        tk.Label(scroll_frame, text=f"  {file_path}",
                 bg="#282a36", fg="#6272a4", font=("", 9),
                 anchor="w").pack(fill="x", padx=4)

        if consumer:
            tk.Label(scroll_frame,
                     text=f"  使用元: {os.path.basename(consumer)}",
                     bg="#282a36", fg="#a6e3a1", font=("", 9),
                     anchor="w").pack(fill="x", padx=4)
            tk.Label(scroll_frame,
                     text=f"    {consumer}",
                     bg="#282a36", fg="#6272a4", font=("", 9),
                     anchor="w").pack(fill="x", padx=4)

        tk.Label(scroll_frame, text="  候補:",
                 bg="#282a36", fg="#f0f0f5", font=("", 9),
                 anchor="w").pack(fill="x", padx=4, pady=(4, 0))

        var = tk.StringVar(value=sorted(candidates)[0])
        selection_vars[i] = var

        for cand in sorted(candidates):
            cand_name = os.path.basename(cand)
            tk.Radiobutton(scroll_frame, text=cand_name,
                           variable=var, value=cand,
                           bg="#282a36", fg="#f0f0f5", selectcolor="#44475a",
                           activebackground="#282a36", activeforeground="#f0f0f5",
                           font=("", 9), anchor="w",
                           ).pack(fill="x", padx=24)
            tk.Label(scroll_frame, text=f"    {cand}",
                     bg="#282a36", fg="#6272a4", font=("", 8),
                     anchor="w").pack(fill="x", padx=24)

    # 決定ボタン
    def on_ok():
        for i, item in enumerate(items):
            var = selection_vars.get(i)
            if var:
                result[item['file']] = var.get()
        dialog.destroy()

    btn_frame = tk.Frame(dialog, bg="#282a36")
    btn_frame.pack(fill="x", padx=16, pady=8)
    tk.Button(btn_frame, text="  決定  ", font=("", 11, "bold"),
              bg="#7aa2f7", fg="#282a36", activebackground="#89b4fa",
              relief="flat", padx=20, pady=6, cursor="hand2",
              command=on_ok).pack()

    dialog.protocol("WM_DELETE_CLOSE", on_ok)
    root_window.wait_window(dialog)

    # フォールバック
    for item in items:
        if item['file'] not in result:
            result[item['file']] = sorted(item['candidates'])[0]

    return result


def _ask_user_jobnet_choice(choices, root_window=None):
    """後方互換用ラッパー: 親JN選択 -> 汎用ダイアログに変換して呼ぶ"""
    items = []
    for file_path, pj_candidates in choices:
        # 親JN候補をフラットなproducerリストに変換
        all_producers = []
        for pj_path, producers in pj_candidates.items():
            all_producers.extend(sorted(producers))
        items.append({
            'file': file_path,
            'candidates': all_producers,
        })
    return _ask_user_producer_choice(
        items,
        title="クロスジョブネット: 出力元の選択",
        description="以下のファイルについて、出力元のユニットを選択してください。",
        root_window=root_window)


# =====================================================================
# ブリッジ定義: DB経由の暗黙的依存関係
# =====================================================================

def _load_bridges(bank):
    """DBブリッジ定義を読み込む。
    ブリッジ定義は「このユニットが必要なら、このユニットも必要」という
    ファイルベースでは追えない依存関係（DB経由等）を手動で定義するもの。
    """
    if not DB_BRIDGE_FILE.exists():
        return []
    try:
        with open(DB_BRIDGE_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # bankが一致するもの、または"*"のものを返す
        return [b for b in data.get('bridges', [])
                if b.get('bank', '*') in ('*', bank)]
    except Exception as e:
        _log(f"[Warning] Failed to load db_bridges.json: {e}")
        return []


def _resolve_bridge_units(bridges, full_set, unit_name_to_fulls, base_dir_to_remove):
    """ブリッジ定義のfrom/toをunit_fullに解決する。
    末尾マッチ方式:
      - "JOBNAME01" -> unit_fullの末尾が /JOBNAME01 に一致するものを探す
      - "/グループ名/JOBNAME01" -> より深い階層で絞り込む
      - 一意なら採用、複数一致なら警告（上位階層の追加を促す）
    戻り値: ({from_unit_full: [to_unit_full, ...]}, [warnings])
    """
    resolved = {}
    warnings = []

    def resolve_one(name):
        """1つのユニット名/パスをunit_fullに解決する（末尾マッチ）"""
        if not name:
            return []
        # unit_full完全一致（フルパスをそのまま書いた場合）
        if name in full_set:
            return [name]
        # サービス名プレフィックス除去して完全一致（AJSROOT1:/...形式）
        stripped = re.sub(r'^[a-zA-Z0-9_]+:', '', name)
        if stripped in full_set:
            return [stripped]
        # 末尾マッチ: 指定文字列でunit_fullの末尾を照合
        suffix = name if name.startswith('/') else '/' + name
        candidates = [uf for uf in full_set if uf.endswith(suffix)]
        if len(candidates) == 1:
            return candidates
        elif len(candidates) > 1:
            warnings.append(f"Bridge: '{name}' が{len(candidates)}件一致。上位階層を追加してください")
            return []
        else:
            warnings.append(f"Bridge: '{name}' が見つかりません")
            return []

    for bridge in bridges:
        from_resolved = resolve_one(bridge.get('from', ''))
        to_names = bridge.get('to', [])
        to_resolved = []
        for t in to_names:
            to_resolved.extend(resolve_one(t))

        for f in from_resolved:
            if f not in resolved:
                resolved[f] = []
            resolved[f].extend(to_resolved)

    for w in warnings:
        _log(f"[Warning] {w}")

    return resolved, warnings


# =====================================================================
# SSH経由のデータ取得
# =====================================================================

def _fetch_graph_and_def(get_ssh_client, ajs_print_path, ajs_path_input,
                         jp1_hostname, jp1_username, read_enc,
                         base_dir_to_remove, tmp_dir, ts):
    """SSH経由でajsprintを実行し、先行関係グラフとAJS定義テキストを取得する。

    ajsprintを2回実行する:
      1回目: -s yes -a で定義テキスト取得（recovery_definition生成用）
      2回目: -f %TY%t%JN%t%ar -R でtype・先行関係取得（グラフ構築用）
    取得したファイルをSFTPでローカルにコピーし、リモート側は削除する。

    Returns:
        (G, ajs_def_txt, ajs_dep_txt):
          G: NetworkXの有向グラフ（先行関係）
          ajs_def_txt: AJS定義テキスト（文字列）
          ajs_dep_txt: type+先行関係テキスト（文字列、親ジョブネット判定にも使える）
    """
    export_list = []
    export_list.append(f'export JP1_HOSTNAME={shlex.quote(jp1_hostname)}')
    export_list.append(f'export JP1_USERNAME={shlex.quote(jp1_username)}')
    env_str = ' && '.join(export_list)

    remote_dep = f"/tmp/ajs_dep_dep_{ts}.txt"
    local_dep = tmp_dir / "ajs_graph.txt"
    cmd_dep = f'{env_str} && {ajs_print_path} -F AJSROOT1 -f %TY%t%JN%t%ar -R {shlex.quote(ajs_path_input)} > {remote_dep}'

    remote_def = f"/tmp/ajs_def_dep_{ts}.txt"
    local_def = tmp_dir / "ajs_definition_original.txt"
    cmd_def = f'{env_str} && {ajs_print_path} -s yes -a {shlex.quote(ajs_path_input)} > {remote_def}'

    _log(f"[Command-Def] {cmd_def}")
    _log(f"[Command-Rel] {cmd_dep}")

    with get_ssh_client() as ssh:
        ch = ssh.get_transport().open_session()
        ch.exec_command(cmd_def.encode(read_enc))
        if ch.recv_exit_status() != 0:
            err = ch.makefile_stderr().read().decode(read_enc, 'ignore')
            _log(f"[Error] Def command: {err}")
            raise RuntimeError(f"AJS定義取得エラー: {err}")

        ch2 = ssh.get_transport().open_session()
        ch2.exec_command(cmd_dep.encode(read_enc))
        if ch2.recv_exit_status() != 0:
            err = ch2.makefile_stderr().read().decode(read_enc, 'ignore')
            _log(f"[Error] Rel command: {err}")
            raise RuntimeError(f"AJS関連取得エラー: {err}")

        try:
            sftp = ssh.open_sftp()
            sftp.get(str(remote_def), str(local_def))
            sftp.get(str(remote_dep), str(local_dep))
            sftp.close()
        finally:
            # 異常終了時もリモート一時ファイルを確実に削除
            try: ssh.exec_command(f"rm -f {remote_def} {remote_dep}")
            except: pass

    ajs_def_txt = local_def.read_text(encoding=read_enc, errors='ignore')
    ajs_dep_txt = local_dep.read_text(encoding=read_enc, errors='ignore')

    G = pre_parse_graph(ajs_dep_txt, base_dir_to_remove)
    _log(f"[Graph] Nodes: {G.number_of_nodes()}, Edges: {G.number_of_edges()}")

    return G, ajs_def_txt, ajs_dep_txt


# =====================================================================
# 目標ユニット解決
# =====================================================================

def _resolve_targets(target_units, final_records):
    """目標ユニット名をunit_full(フルパス)に変換する。

    入力形式を3段階でフォールバックし、柔軟に受け付ける:
      1. unit_full完全一致: "/グループ名/.../JOBNAME01"
      2. サービス名除去後一致: "AJSROOT1:/グループ名/..." -> "/グループ名/..."
      3. ユニット名末尾一致: "JOBNAME01"（一意の場合のみ）

    Returns:
        (target_unit_fulls, not_found, ambiguous, full_set, unit_name_to_fulls):
          target_unit_fulls: set of unit_full
          not_found: list of unresolved names
          ambiguous: list of (name, [candidates])
          full_set: set of all unit_full
          unit_name_to_fulls: dict of {unit_name: [unit_full, ...]}
    """
    full_set = set()
    unit_name_to_fulls = collections.defaultdict(list)
    for record in final_records:
        full_set.add(record['unit_full'])
        unit_name_to_fulls[record['unit']].append(record['unit_full'])

    def strip_service_prefix(path):
        return re.sub(r'^[a-zA-Z0-9_]+:', '', path)

    target_unit_fulls = set()
    not_found = []
    ambiguous = []

    for tgt in target_units:
        # 1. unit_fullそのままでマッチ
        if tgt in full_set:
            target_unit_fulls.add(tgt)
            continue

        # 2. サービス名プレフィックス除去してマッチ
        stripped = strip_service_prefix(tgt)
        if stripped in full_set:
            target_unit_fulls.add(stripped)
            continue

        # 3. ユニット名（末尾部分）でマッチ
        basename = tgt.rsplit('/', 1)[-1] if '/' in tgt else tgt
        candidates = unit_name_to_fulls.get(basename, [])
        if len(candidates) == 1:
            target_unit_fulls.add(candidates[0])
        elif len(candidates) > 1:
            ambiguous.append((tgt, candidates))
        else:
            not_found.append(tgt)

    if ambiguous:
        for tgt, cands in ambiguous:
            _log(f"[Warning] Ambiguous unit '{tgt}': {cands}")
    if not_found:
        _log(f"[Warning] Units not found: {not_found}")

    return target_unit_fulls, not_found, ambiguous, full_set, unit_name_to_fulls


# =====================================================================
# メイン処理: dep_start_job
# =====================================================================

def dep_start_job(gui_vars, gui_funcs, text_box):
    """依存関係解析のメイン処理。各段階の関数を順に呼び出す。"""

    update_status = gui_funcs['update_status']
    show_info = gui_funcs['show_info']
    show_error = gui_funcs['show_error']
    save_hist = gui_funcs['save_hist']
    get_ssh_client = gui_funcs['get_ssh_client']

    with open(LOG_FILE_RUN, "w", encoding="utf-8") as f:
        f.write(f"=== Dependency Execution Start: {datetime.datetime.now()} ===\n")

    target_units_str = gui_vars['v_dep_tgt_files'].get().strip()
    ajs_path_input = gui_vars['v_dep_ajs'].get().strip()
    res_root = gui_vars['v_dep_res'].get().strip()
    bank = gui_vars['v_dep_bank'].get()

    ajs_print_path = gui_vars['v_ajs_print_path'].get()
    jp1_hostname = gui_vars['v_jp1_hostname'].get()
    jp1_username = gui_vars['v_jp1_username'].get()

    _log(f"[Params] AJS Path: {ajs_path_input}")
    _log(f"[Params] Bank: {bank}")

    try:
        if not target_units_str:
            raise ValueError("目標ユニット名を入力してください。")

        target_units = [u.strip() for u in target_units_str.split('\n') if u.strip()]
        if not target_units:
            raise ValueError("ユニット名が有効ではありません。")

        _log(f"[Params] Target Units ({len(target_units)}): {target_units}")

        # --- 0. 初期準備 ---
        ts = time.strftime("%Y%m%d%H%M%S")
        base_dir = pathlib.Path(sys.argv[0]).resolve().parent
        out_dir = base_dir / DIR_NAME_DEP / ts
        out_dir.mkdir(parents=True, exist_ok=True)

        tmp_dir = out_dir / "tmp"
        tmp_dir.mkdir(exist_ok=True)

        _log(f"[Info] Output Directory: {out_dir}")

        path_part = ajs_path_input.split(':', 1)[1] if ':' in ajs_path_input else ajs_path_input
        base_dir_to_remove = os.path.dirname(path_part)

        read_enc = ENC[gui_vars['v_t5_out_c'].get()]
        if read_enc == 'utf-8': read_enc = 'cp932'

        # --- 1. I/O解析実行 ---
        update_status("I/O解析実行中...", 10)
        final_records, _ = analyze_ajs_jobs(gui_vars, gui_funcs, out_dir, use_cache=True)

        producer_map = collections.defaultdict(set)
        db_producer_map = collections.defaultdict(set)  # {テーブル名 -> W元unit_full}
        for record in final_records:
            for output_path in record.get('outputs', []):
                producer_map[output_path].add(record['unit_full'])
            # DB W操作ユニットをdb_producer_mapに登録
            for table_name, op in record.get('db_tables', {}).items():
                if op in ('W', 'RW'):
                    db_producer_map[table_name].add(record['unit_full'])

        _log(f"[Info] Producer Map built. Files: {len(producer_map)}, DB tables: {len(db_producer_map)}")

        # --- 2. 先行関係グラフ + AJS定義取得（範囲全体を一括取得） ---
        update_status("先行関係グラフ取得中...", 30)
        G_full, ajs_def_txt, ajs_dep_txt = _fetch_graph_and_def(
            get_ssh_client, ajs_print_path, ajs_path_input,
            jp1_hostname, jp1_username, read_enc,
            base_dir_to_remove, tmp_dir, ts)

        # --- 2b. 親ジョブネット一覧の判定 ---
        dep_lines = ajs_dep_txt.splitlines()
        type_lines = []
        for line in dep_lines:
            parts = line.split('\t')
            if len(parts) >= 2:
                type_lines.append(f"{parts[0]}\t{parts[1]}")

        parent_jobnets = discover_parent_jobnets(type_lines, path_part)
        _log(f"[CrossJobnet] Parent jobnets ({len(parent_jobnets)}): {parent_jobnets}")

        # --- 3. 目標ユニット特定 ---
        update_status("目標ユニット特定中...", 40)
        target_unit_fulls, not_found, ambiguous, full_set, unit_name_to_fulls = \
            _resolve_targets(target_units, final_records)

        if not target_unit_fulls:
            msgs = []
            if not_found: msgs.append(f"見つからない: {not_found}")
            if ambiguous: msgs.append(f"同名が複数: {[t for t,_ in ambiguous]}")
            raise ValueError(f"指定されたユニットを特定できません。{' / '.join(msgs)}")

        # --- 2c. 目標ユニットを親ジョブネットごとにグループ分け ---
        targets_by_pj = collections.defaultdict(set)
        targets_no_pj = set()
        for tuf in target_unit_fulls:
            pj = find_parent_jobnet(tuf, parent_jobnets)
            if pj:
                targets_by_pj[pj].add(tuf)
            else:
                targets_no_pj.add(tuf)
                _log(f"[Warning] Target unit not in any parent jobnet: {tuf}")

        if targets_by_pj:
            pj_list = sorted(targets_by_pj.keys())
            _log(f"[CrossJobnet] Target parent jobnets: {pj_list}")
        if targets_no_pj:
            _log(f"[CrossJobnet] Targets without parent jobnet: {targets_no_pj}")

        # --- 4. DFSルートごとに独立して探索 ---
        update_status("逆引き探索実行中...", 50)

        record_by_full = {}
        for record in final_records:
            record_by_full[record['unit_full']] = record

        bridges_raw = _load_bridges(bank)
        bridge_map, bridge_warnings = _resolve_bridge_units(
            bridges_raw, full_set, unit_name_to_fulls, base_dir_to_remove)
        if bridge_map:
            _log(f"[Bridge] Loaded {len(bridge_map)} bridge rules.")

        # ユーザー選択のキャッシュ（ルート間共有、Excelレポートにも渡す）
        choice_cache = {}

        # GUIのrootウィンドウ（ユーザー選択ダイアログ用）
        try:
            import tkinter as tk
            root_window = tk._default_root
        except Exception:
            root_window = None

        # 全ルートの結果を収集（最後にマージ）
        all_route_results = []

        # --- DFSルート実行関数 ---
        def _run_dfs_route(route_targets, route_start_pj):
            """1つのDFSルートを実行する。
            route_targets: この起点の目標ユニット群
            route_start_pj: 起点の親JNパス（Noneなら親JN不明）
            """
            route_result = {}

            def _route_needed():
                """ルート内で蓄積された全必要ユニットを返す"""
                s = set()
                for pj_data in route_result.values():
                    s.update(pj_data['needed'])
                return s

            # --- 初回BFS ---
            pj_label = os.path.basename(route_start_pj) if route_start_pj else "全体"
            _log(f"[DFS Route] Start: {pj_label}, targets={len(route_targets)}")

            # ask_choice_fn: BFS内で並列producerが見つかった時にユーザーに選択を求める関数
            bfs_result = _bfs_trace(G_full, route_targets, producer_map,
                                    record_by_full, base_dir_to_remove, bridge_map,
                                    root_window=root_window,
                                    db_producer_map=db_producer_map,
                                    ask_choice_fn=_ask_user_producer_choice)

            # 初回BFS結果をルート結果に記録
            pj_key = route_start_pj or "__no_pj__"
            route_result[pj_key] = {
                'needed': set(bfs_result['needed_units_full']),
                'true_externals': dict(bfs_result['true_externals']),
                'disconnected_externals': dict(bfs_result['disconnected_externals']),
                'trace_data': [("SECTION", "", "", "", f"DFS: {pj_label}")] + list(bfs_result['trace_data']),
                'trace_lines': list(bfs_result['trace_lines']),
                'bridged': set(bfs_result['bridged_units']),
            }

            # --- DFSラウンド（disconnected_externalsを追跡） ---
            if not parent_jobnets or len(parent_jobnets) <= 1:
                return route_result

            current_round_files = {f: route_start_pj
                                   for f in bfs_result['disconnected_externals'].keys()}
            max_rounds = 10

            for dfs_round in range(max_rounds):
                if not current_round_files:
                    break

                _log(f"[DFS Route {pj_label}] Round {dfs_round + 1}: "
                     f"{len(current_round_files)} files")

                # 一括で親JN横断検索
                auto_resolved = {}
                needs_user_choice = []

                file_consumers = {}
                for ext_file in current_round_files.keys():
                    all_consumers = bfs_result['disconnected_externals'].get(ext_file, set())
                    for prev_pj_data in route_result.values():
                        all_consumers = all_consumers | prev_pj_data.get(
                            'disconnected_externals', {}).get(ext_file, set())
                    source_pj = current_round_files.get(ext_file)
                    filter_pj = source_pj or route_start_pj
                    if filter_pj:
                        pj_prefix = filter_pj.rstrip('/') + '/'
                        filtered = {c for c in all_consumers
                                    if c.startswith(pj_prefix) or c == filter_pj}
                        file_consumers[ext_file] = filtered if filtered else all_consumers
                    else:
                        file_consumers[ext_file] = all_consumers

                for ext_file, source_pj in current_round_files.items():
                    consumers = file_consumers.get(ext_file, set())

                    exclude = {source_pj} if source_pj else None
                    if ext_file.startswith("[DB]"):
                        db_table = ext_file[4:]
                        pj_producers = find_producers_across_jobnets(
                            db_table, db_producer_map, parent_jobnets,
                            exclude_jobnets=exclude)
                    else:
                        pj_producers = find_producers_across_jobnets(
                            ext_file, producer_map, parent_jobnets,
                            exclude_jobnets=exclude)

                    if not pj_producers:
                        continue

                    all_candidates = []
                    for pj_path, producers in pj_producers.items():
                        all_candidates.extend(producers)

                    for consumer in sorted(consumers):
                        cache_key = (ext_file, consumer)

                        if cache_key in choice_cache:
                            chosen = choice_cache[cache_key]
                            pj = find_parent_jobnet(chosen, parent_jobnets)
                            if pj:
                                auto_resolved.setdefault(ext_file, {})[consumer] = (pj, chosen)
                            continue

                        if len(all_candidates) == 1:
                            chosen = all_candidates[0]
                            pj = find_parent_jobnet(chosen, parent_jobnets)
                            auto_resolved.setdefault(ext_file, {})[consumer] = (pj, chosen)
                            choice_cache[cache_key] = chosen
                        else:
                            needs_user_choice.append({
                                'file': ext_file,
                                'consumer': consumer,
                                'candidates': all_candidates,
                            })

                # ユーザー選択
                if needs_user_choice:
                    _log(f"[DFS Route {pj_label}] User choice: {len(needs_user_choice)} files")
                    user_choices = _ask_user_producer_choice(
                        needs_user_choice,
                        title="クロスジョブネット: 出力元の選択",
                        description="以下のファイルについて、出力元のユニットを選択してください。",
                        root_window=root_window)
                    for item in needs_user_choice:
                        ext_file = item['file']
                        consumer = item['consumer']
                        chosen = user_choices.get(ext_file)
                        if not chosen:
                            chosen = sorted(item['candidates'])[0]
                        pj = find_parent_jobnet(chosen, parent_jobnets)
                        auto_resolved.setdefault(ext_file, {})[consumer] = (pj, chosen)
                        choice_cache[(ext_file, consumer)] = chosen

                # 親JNごとにBFS実行
                route_known = _route_needed()
                bfs_by_pj = collections.defaultdict(set)
                file_to_pj = {}
                for ext_file, consumer_choices in auto_resolved.items():
                    for consumer, (pj_path, chosen) in consumer_choices.items():
                        if chosen not in route_known:
                            bfs_by_pj[pj_path].add(chosen)
                        file_to_pj[ext_file] = pj_path

                next_round_files = {}

                for pj_path, producer_units in bfs_by_pj.items():
                    if not producer_units:
                        continue
                    _log(f"[DFS Route {pj_label}] BFS in "
                         f"{os.path.basename(pj_path)}: {len(producer_units)} producers")

                    cross_result = _bfs_trace(
                        G_full, producer_units, producer_map,
                        record_by_full, base_dir_to_remove, bridge_map,
                        root_window=root_window,
                        pre_known=route_known,
                        db_producer_map=db_producer_map,
                        ask_choice_fn=_ask_user_producer_choice)

                    trigger_files = [os.path.basename(f) for f, pj in file_to_pj.items()
                                     if pj == pj_path]
                    trigger_info = ", ".join(trigger_files[:3])
                    if len(trigger_files) > 3:
                        trigger_info += "..."

                    if pj_path not in route_result:
                        route_result[pj_path] = {
                            'needed': set(), 'true_externals': {},
                            'disconnected_externals': {}, 'trace_data': [],
                            'trace_lines': [], 'bridged': set(),
                        }
                    pj_data = route_result[pj_path]
                    pj_data['needed'].update(cross_result['needed_units_full'])
                    pj_data['bridged'].update(cross_result['bridged_units'])
                    for f, c in cross_result['true_externals'].items():
                        pj_data['true_externals'].setdefault(f, set()).update(c)
                    for f, c in cross_result['disconnected_externals'].items():
                        pj_data['disconnected_externals'].setdefault(f, set()).update(c)

                    pj_data['trace_data'].append(
                        ("SECTION", "", "", "",
                         f"DFS R{dfs_round + 1}: {os.path.basename(pj_path)} "
                         f"(起因: {trigger_info})"))
                    pj_data['trace_data'].extend(cross_result['trace_data'])
                    pj_data['trace_lines'].extend(cross_result['trace_lines'])

                    for new_ext in cross_result['disconnected_externals']:
                        if new_ext not in next_round_files and new_ext not in current_round_files:
                            next_round_files[new_ext] = pj_path

                current_round_files = next_round_files

            return route_result

        # --- 各DFSルートを実行 ---
        if targets_by_pj:
            for pj_path, pj_targets in sorted(targets_by_pj.items()):
                route_result = _run_dfs_route(pj_targets, pj_path)
                all_route_results.append(route_result)

        if targets_no_pj:
            route_result = _run_dfs_route(targets_no_pj, None)
            all_route_results.append(route_result)

        # --- 全ルート結果を親JNごとにマージ ---
        _log(f"[Merge] Merging {len(all_route_results)} route results")
        needed_units_full = set()
        true_externals = {}
        disconnected_externals = {}
        trace_data = []
        trace_lines = []
        bridged_units = set()

        for route in all_route_results:
            for pj_key, pj_data in route.items():
                needed_units_full.update(pj_data['needed'])
                bridged_units.update(pj_data['bridged'])
                for f, c in pj_data['true_externals'].items():
                    true_externals.setdefault(f, set()).update(c)
                for f, c in pj_data['disconnected_externals'].items():
                    disconnected_externals.setdefault(f, set()).update(c)
                trace_data.extend(pj_data['trace_data'])
                trace_lines.extend(pj_data['trace_lines'])

        _log(f"[Merge] Total needed: {len(needed_units_full)}, "
             f"true_ext: {len(true_externals)}, disc_ext: {len(disconnected_externals)}")

        # --- 4c. 充足チェック（マージ後にG_fullで1回） ---
        update_status("充足チェック中...", 60)

        predecessor_cache = {}
        supplemented_units = set()
        # ask_choice_fn: 充足チェック内で複数候補がある時にユーザーに選択を求める関数
        supplemented_externals = _supplement_check(
            needed_units_full, record_by_full, producer_map,
            G_full, base_dir_to_remove, predecessor_cache,
            true_externals, disconnected_externals, supplemented_units,
            root_window=root_window,
            ask_choice_fn=_ask_user_producer_choice)

        # 充足チェック結果をトレースログに追記
        trace_lines.append("")
        trace_lines.append("=" * 60)
        trace_lines.append("===== 充足チェック =====")
        trace_lines.append("=" * 60)
        if supplemented_units:
            trace_lines.append(f"")
            trace_lines.append(f"  BFS完了後、必要ジョブの入力が結果セット内で")
            trace_lines.append(f"  作成されていないケースを検出し、producerを追加。")
            trace_lines.append(f"")
            for s_unit in sorted(supplemented_units):
                s_rec = record_by_full.get(s_unit)
                s_name = s_rec['unit'] if s_rec else os.path.basename(s_unit)
                trace_lines.append(f"  + {s_name}  ({s_unit})")
        else:
            trace_lines.append(f"")
            trace_lines.append(f"  追加なし（全入力が充足済み）")

        for f, consumers in supplemented_externals.items():
            true_externals.setdefault(f, set()).update(consumers)

        # DBブリッジ結果をトレースログに追記
        if bridged_units or bridge_warnings:
            trace_lines.append("")
            trace_lines.append("=" * 60)
            trace_lines.append("===== DBブリッジ =====")
            trace_lines.append("=" * 60)
            if bridged_units:
                trace_lines.append(f"")
                trace_lines.append(f"  ブリッジ定義により追加されたユニット:")
                trace_lines.append(f"")
                for b_unit in sorted(bridged_units):
                    b_rec = record_by_full.get(b_unit)
                    b_name = b_rec['unit'] if b_rec else os.path.basename(b_unit)
                    trace_lines.append(f"  + {b_name}  ({b_unit})")
            if bridge_warnings:
                trace_lines.append(f"")
                for w in bridge_warnings:
                    trace_lines.append(f"  [警告] {w}")
            if not bridged_units and not bridge_warnings:
                trace_lines.append(f"")
                trace_lines.append(f"  追加なし")

        # トレースログのサマリ + ファイル出力
        trace_lines.append("")
        trace_lines.append("=" * 60)
        trace_lines.append(f"===== 結果サマリ =====")
        trace_lines.append("=" * 60)
        trace_lines.append(f"目標ユニット:     {len(target_unit_fulls)}件")
        trace_lines.append(f"必要ジョブ:       {len(needed_units_full)}件")
        if supplemented_units:
            trace_lines.append(f"  (うち充足チェック追加: {len(supplemented_units)}件)")
        if bridged_units:
            trace_lines.append(f"  (うちDBブリッジ追加: {len(bridged_units)}件)")
        all_external_files = set(true_externals.keys()) | set(disconnected_externals.keys())
        trace_lines.append(f"外部入力ファイル: {len(all_external_files)}件")
        if true_externals:
            trace_lines.append(f"  (外部ファイル: {len(true_externals)}件)")
        if disconnected_externals:
            trace_lines.append(f"  (マスタ系: {len(disconnected_externals)}件)")

        trace_log_path = tmp_dir / 'trace_log.txt'
        trace_log_path.write_text("\n".join(trace_lines), encoding='utf-8')

        # --- 5. 定義フィルタリングの準備 ---
        update_status("定義フィルタリング準備...", 70)

        normalized_need_set = set()

        def add_parents(path_set, n_path):
            anc = n_path
            while '/' in anc.strip('/'):
                anc = '/' + '/'.join(anc.strip('/').split('/')[:-1])
                if not anc.strip('/'): break
                path_set.add(anc)

        for unit_full in needed_units_full:
            norm_path = pre_normalize(unit_full, base_dir_to_remove)
            if norm_path.strip('/'):
                normalized_need_set.add(norm_path)
                add_parents(normalized_need_set, norm_path)

        # --- 6. 定義フィルタリング＋ar行再結線 ---
        update_status("グラフ構築と再結線処理...", 90)
        rec_txt = pre_filter_definition(ajs_def_txt, normalized_need_set, G_full)

        # --- 7. 出力 ---
        out_file_def = out_dir / 'recovery_definition.txt'
        out_file_def.write_text(rec_txt, encoding=ENC[gui_vars['v_t5_out_c'].get()], newline=NL[gui_vars['v_t5_out_n'].get()])

        out_file_ext = tmp_dir / 'missing_files.txt'
        true_ext_sorted = sorted(true_externals.keys())
        disc_ext_sorted = sorted(disconnected_externals.keys())
        sup_ext_sorted = sorted(supplemented_externals.keys())
        ext_lines = []
        if true_ext_sorted:
            ext_lines.append("--- 外部ファイル（TBL・事前準備等） ---")
            ext_lines.extend(true_ext_sorted)
        if disc_ext_sorted:
            ext_lines.append("")
            ext_lines.append("--- マスタ系ファイル ---")
            ext_lines.extend(disc_ext_sorted)
        if sup_ext_sorted:
            ext_lines.append("")
            ext_lines.append("--- 充足チェック追加ユニットの外部入力 ---")
            ext_lines.extend(sup_ext_sorted)
        ext_list = true_ext_sorted + disc_ext_sorted + sup_ext_sorted
        if ext_lines:
            out_file_ext.write_text("\n".join(ext_lines), encoding='utf-8')

        # Excelレポート出力
        try:
            write_dependency_report(
                out_path=out_dir / 'dependency_report.xlsx',
                true_externals=true_externals,
                disconnected_externals=disconnected_externals,
                producer_map=producer_map,
                record_by_full=record_by_full,
                trace_data=trace_data,
                needed_units_full=needed_units_full,
                target_unit_fulls=target_unit_fulls,
                supplemented_units=supplemented_units,
                bridged_units=bridged_units,
                not_found=not_found,
                ambiguous=ambiguous,
                gui_vars=gui_vars,
                G=G_full,
                base_dir_to_remove=base_dir_to_remove,
                parent_jobnets=parent_jobnets,
                choice_cache=choice_cache,
                db_producer_map=db_producer_map,
            )
        except Exception as e:
            _log(f"[Warning] Excel report failed: {e}")

        # GUI更新
        text_box.delete('1.0', 'end')
        if needed_units_full:
            bfs_units = sorted(needed_units_full - supplemented_units - bridged_units)
            sup_units = sorted(supplemented_units)
            brg_units = sorted(bridged_units)

            text_box.insert('end', f"--- 目標ユニット ({len(target_unit_fulls)}件) ---\n" + "\n".join(sorted(target_unit_fulls)))
            text_box.insert('end', f"\n\n--- 必要ジョブ ({len(bfs_units)}件) ---\n" + "\n".join(bfs_units))
            if sup_units:
                text_box.insert('end', f"\n\n--- 充足チェック追加 ({len(sup_units)}件) ---\n" + "\n".join(sup_units))
            if brg_units:
                text_box.insert('end', f"\n\n--- DBブリッジ追加 ({len(brg_units)}件) ---\n" + "\n".join(brg_units))
            text_box.insert('end', f"\n\n--- 外部入力ファイル ({len(ext_list)}件) ---\n" + ("\n".join(ext_list) if ext_list else "なし"))
            if not_found:
                text_box.insert('end', f"\n\n--- 見つからなかったユニット ({len(not_found)}件) ---\n" + "\n".join(not_found))
        else:
            text_box.insert('end', "探索結果: 該当ジョブなし")

        _write_detail_log({
            "target_units": target_units,
            "target_unit_fulls": target_unit_fulls,
            "producer_map": producer_map,
            "needed_units_full": needed_units_full,
            "normalized_need_set": normalized_need_set,
            "true_externals": true_externals,
            "disconnected_externals": disconnected_externals,
            "not_found_units": not_found,
        }, LOG_FILE_DETAIL, _log)

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
