#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
依存関係解析 - 探索エンジン (Dependency Tracer)

BFS逆引きトレース・producer絞り込み・充足チェックなど、
依存関係解析の「探索ロジック」を担当する。
GUIやSSHには依存しない純粋なロジックモジュール。

主要関数:
  - _bfs_trace(): BFS逆引きトレース（依存関係の核）
  - filter_producers_by_graph(): グラフ上のproducer絞り込み
  - _supplement_check(): BFS結果の充足チェック
  - discover_parent_jobnets(): 親ジョブネット抽出
  - find_parent_jobnet(): ユニット→親ジョブネット特定
"""

import os
import collections

from ajs_constants import LOG_DIR, CONFIG_FILE
from ajs_utils import make_logger, pre_normalize
from ajs_rel_logic import pre_compute_need

LOG_FILE_RUN = LOG_DIR / "depend_run.log"
_log = make_logger(LOG_FILE_RUN)


# =====================================================================
# ユーティリティ: パス・ジョブネット操作
# =====================================================================

def _get_ancestors(norm_path):
    """正規化パスの親階層を全て返す（自分→親→祖父→...の順）"""
    ancestors = [norm_path]
    p = norm_path
    while '/' in p.strip('/'):
        p = '/' + '/'.join(p.strip('/').split('/')[:-1])
        if not p.strip('/'): break
        ancestors.append(p)
    return ancestors


def discover_parent_jobnets(type_lines, range_path):
    """ajsprintの%TY出力と探索範囲から、対象の親ジョブネット一覧を抽出する。

    親ジョブネット = range_path配下で最も浅い階層にあるジョブネット型ユニット。
    あるnetの配下にある別のnet（サブネット）は含めない。
    range_path自体がジョブネット型の場合、それ自身のみを返す。

    Args:
        type_lines: "%TY\\t%JN" 形式の行リスト（list）または改行区切り文字列
        range_path: 探索範囲のAJSパス（例: "/グループ名"）

    Returns:
        list: 親ジョブネットのパス一覧
    """
    # 文字列が渡された場合はリストに変換（後方互換）
    if isinstance(type_lines, str):
        type_lines = type_lines.strip().split('\n')

    range_norm = range_path.rstrip('/')
    # ジョブネット型の判定用（短縮形・フルスペル両対応）
    net_types = {'n', 'ne', 'nm', 'net'}

    # まずrange配下の全netを収集
    all_nets = []
    for line in type_lines:
        if isinstance(line, str):
            line = line.strip()
        if not line:
            continue
        parts = line.split('\t')
        if len(parts) < 2:
            continue
        ty, jn = parts[0], parts[1]
        if ty not in net_types:
            continue
        jn_norm = jn.rstrip('/')

        # range自身 or range配下
        if jn_norm == range_norm or jn_norm.startswith(range_norm + '/'):
            all_nets.append(jn_norm)

    # range自身がnetの場合 → それだけを返す（配下のサブネットは含めない）
    if range_norm in all_nets:
        return [range_norm]

    # 「上位に別のnetがないnet」= 親ジョブネット
    # あるnetが別のnetの配下にある場合、そのnetはサブネットなので除外
    parent_jobnets = []
    for jn in all_nets:
        is_subnet = False
        for other_jn in all_nets:
            if other_jn == jn:
                continue
            # otherがjnの祖先（jnがother配下にある）なら、jnはサブネット
            if jn.startswith(other_jn + '/'):
                is_subnet = True
                break
        if not is_subnet:
            parent_jobnets.append(jn)

    return parent_jobnets


def find_parent_jobnet(unit_path, parent_jobnets):
    """ユニットパスが属する親ジョブネットを特定する。

    Args:
        unit_path: ユニットのフルパス
        parent_jobnets: discover_parent_jobnetsで取得した親ジョブネット一覧

    Returns:
        str or None: 一致する親ジョブネットパス
    """
    if not parent_jobnets:
        return None
    unit_norm = unit_path.rstrip('/')
    for pj in parent_jobnets:
        pj_norm = pj.rstrip('/')
        if unit_norm == pj_norm or unit_norm.startswith(pj_norm + '/'):
            return pj_norm
    return None


def filter_dep_text_by_jobnet(dep_text, parent_jobnet_path):
    """先行関係テキストを指定親ジョブネット配下のみに絞り込む。

    AR継続行（タブを含まない行）は、直前のヘッダ行が対象なら含める。

    Args:
        dep_text: ajsprintの先行関係出力テキスト（%TY\\t%JN\\t%ar形式）
        parent_jobnet_path: 絞り込み対象の親ジョブネットパス

    Returns:
        str: 絞り込み後のテキスト
    """
    pj_norm = parent_jobnet_path.rstrip('/')
    filtered_lines = []
    in_target = False  # 直前のヘッダ行が対象かどうか
    for line in dep_text.strip().split('\n'):
        parts = line.strip().split('\t')
        if len(parts) >= 2:
            # ヘッダ行（type\tpath\t...）
            jn = parts[1].rstrip('/')
            in_target = (jn == pj_norm or jn.startswith(pj_norm + '/'))
        # else: AR継続行（タブなし）→ in_targetを引き継ぐ
        if in_target:
            filtered_lines.append(line)
    return '\n'.join(filtered_lines)


def find_producers_across_jobnets(file_path, producer_map, parent_jobnets,
                                  exclude_jobnets=None):
    """外部入力ファイルの出力元を、全親ジョブネットから検索する。

    あるファイルのproducer（出力元ユニット）がどの親ジョブネットに属するかを調べ、
    親ジョブネット単位でグルーピングして返す。

    Args:
        file_path: 検索対象のファイルパス
        producer_map: {ファイルパス -> 出力元unit_fullの集合}
        parent_jobnets: 親ジョブネット一覧
        exclude_jobnets: 除外する親ジョブネットのset（ループ防止用）

    Returns:
        dict: {親ジョブネットパス -> set of producer unit_full}
              producerが見つからなければ空dict
    """
    if exclude_jobnets is None:
        exclude_jobnets = set()

    producers = producer_map.get(file_path)
    if not producers:
        return {}

    result = {}
    for producer_uf in producers:
        pj = find_parent_jobnet(producer_uf, parent_jobnets)
        if pj is None:
            continue
        if pj in exclude_jobnets:
            continue
        result.setdefault(pj, set()).add(producer_uf)

    return result


# =====================================================================
# producer絞り込み: グラフベースのフィルタリング
# =====================================================================

def filter_producers_by_graph(G, producers, consumer_unit_full, base_dir_to_remove, predecessor_cache, skip_order_filter=False):
    """
    producer候補から、先行関係グラフ上でconsumerの先行にあたるものを返す。

    Args:
        skip_order_filter: Trueの場合、Step 2（順序による絞り込み）をスキップし、
                           接続された全producerを返す。DB依存の場合に使用。
                           ファイルは「最後の上書きが最終結果」だが、
                           DBは「全てのW操作が必要」なため。

    戻り値: (採用set, Step1除外set, Step2除外set)
      - 採用: 最終的に必要と判定されたproducer
      - Step1除外: pre_compute_needの先行集合に含まれない（未接続）
      - Step2除外: is_beforeで先行と判定された（より後に実行されるproducerがいる）
    """
    consumer_norm = pre_normalize(consumer_unit_full, base_dir_to_remove)

    # consumerの先行ユニット集合を計算（キャッシュあれば再利用）
    if consumer_norm not in predecessor_cache:
        predecessor_cache[consumer_norm] = pre_compute_need(G, consumer_norm)

    need_set = predecessor_cache[consumer_norm]

    # --- Step 1: 先行関係で接続されたproducerに絞る ---
    # なぜ2段階か: 同じファイルを出力するジョブが複数あるとき、
    # 「consumerより先に実行されるもの」だけが有効なproducer。
    # Step 1で先行集合にいないもの（別ジョブネット等）を除外し、
    # Step 2で残った複数候補から「最後に実行されるもの」を選ぶ。
    # Step 2が必要な理由: ファイルは後から上書きされるので、最後のwriterが
    # 最終結果を持つ。ただしDBは全W操作が必要なのでStep 2をスキップする。
    connected = set()
    step1_rejected = set()
    for p in producers:
        p_norm = pre_normalize(p, base_dir_to_remove)
        if p_norm in need_set:
            connected.add(p)
        else:
            step1_rejected.add(p)

    if len(connected) <= 1:
        return connected, step1_rejected, set()

    # DB依存の場合は全W元が必要なので、順序による絞り込みをスキップ
    if skip_order_filter:
        return connected, step1_rejected, set()

    # Step 2: 接続されたproducerが複数 -> 最後に実行されるものを選ぶ
    #
    # 考え方: producer A の親ジョブネットが producer B の親ジョブネットの
    #         先行にあるなら、Aの方が先に実行される -> Aは不要

    import networkx as nx

    producer_ancestor_list = {}  # 順序付きリスト（自分→親→祖父）
    for p in connected:
        p_norm = pre_normalize(p, base_dir_to_remove)
        producer_ancestor_list[p] = _get_ancestors(p_norm)

    # producer AがBより先に実行されるか判定する
    # A（またはAの祖先）-> B（またはBの祖先）へのパスがグラフ上に存在するかを確認
    def is_before(a, b):
        """aがbより先に実行されるか（a->bの先行関係があるか）"""
        a_ancs = producer_ancestor_list[a]
        b_ancs = producer_ancestor_list[b]
        for a_anc in a_ancs:
            if a_anc not in G: continue
            for b_anc in b_ancs:
                if b_anc not in G or a_anc == b_anc: continue
                if nx.has_path(G, a_anc, b_anc):
                    return True  # aの祖先 -> bの祖先 へパスあり -> aが先
        return False

    # 「自分より後に実行されるproducerが存在する」-> 自分は除外
    latest = set(connected)
    for p in connected:
        for other_p in connected:
            if other_p == p: continue
            if is_before(p, other_p):
                # pはother_pより先に実行される -> pを除外
                latest.discard(p)
                break

    result = latest if latest else connected
    step2_rejected = connected - result
    return result, step1_rejected, step2_rejected


# =====================================================================
# BFS逆引きトレース: 依存関係解析の核
# =====================================================================

def _bfs_trace(G, target_unit_fulls, producer_map, record_by_full,
               base_dir_to_remove, bridge_map, root_window=None,
               pre_known=None, db_producer_map=None, ask_choice_fn=None):
    """BFS逆引きトレースを実行し、必要ユニット群と外部入力を特定する。

    目標ユニットの入力ファイルから出発して:
      「このファイルを作ってるジョブは？」-> そのジョブの入力は？-> ...
    を繰り返し、全ての依存関係を解決する。
    加えて、DB依存（R/RWテーブルのW元探索）も同時に行う。

    Args:
        G: 先行関係グラフ（NetworkX DiGraph）
        target_unit_fulls: 目標ユニットのunit_fullの集合
        producer_map: {ファイルパス -> 出力元unit_fullの集合}
        record_by_full: {unit_full -> I/O解析レコード}
        base_dir_to_remove: パス正規化用の基準パス
        bridge_map: {unit_full -> [bridged_unit_full, ...]} DBブリッジ定義
        root_window: GUIルートウィンドウ（並列producer選択ダイアログ用）
        pre_known: set or None -- 同一DFSルート内で既に発見済みのユニット。
                   BFS内でこれらのユニットに到達したらスキップし、再探索を避ける。
        db_producer_map: {テーブル名 -> 出力元unit_fullの集合} DB W操作ユニット
        ask_choice_fn: ユーザー選択ダイアログ関数（依存性の注入）。
                       Noneの場合、並列producer警告をスキップする。

    Returns:
        dict: 以下のキーを持つ結果辞書
          needed_units_full: set   - 必要と判定されたユニットの集合
          true_externals: dict     - {ファイル/テーブル: {consumer_unit_full, ...}}
          disconnected_externals: dict - 同上（producer存在するが未接続）
          trace_data: list         - [(layer, name, full, file, info)]
          trace_lines: list        - 人が読めるトレースログ行
          bridged_units: set       - DBブリッジで追加されたユニット
    """
    predecessor_cache = {}
    queue = collections.deque()
    # pre_knownがあれば初期値に設定。既知ユニットはBFS内でスキップされる。
    needed_units_full = set(pre_known) if pre_known else set()
    true_externals = {}
    disconnected_externals = {}
    trace_lines = []
    trace_data = []
    bridged_units = set()
    parallel_warnings = []  # 並列producer警告 [{file, consumer, candidates}]
    _db_searched = set()    # DB検索済みユニット（二重検索防止）

    # 入力なしユニットの記録用（次レイヤーのtraceに出力）
    _no_input_units = []

    # --- 入力ファイルのproducer探索（Layer 0とBFSループで共用） ---
    def _search_inputs(consumer_unit_full, input_files, current_layer, layer_discovered):
        """consumerの入力ファイル群についてproducerを探索し、結果を記録する。"""
        consumer_rec = record_by_full.get(consumer_unit_full)
        consumer_name = consumer_rec['unit'] if consumer_rec else os.path.basename(consumer_unit_full)

        for current_file in input_files:
            producers = producer_map.get(current_file)
            if producers:
                producers = producers - {consumer_unit_full}

            if not producers:
                true_externals.setdefault(current_file, set()).add(consumer_unit_full)
                if consumer_unit_full not in layer_discovered:
                    layer_discovered[consumer_unit_full] = []
                layer_discovered[consumer_unit_full].append(
                    (current_file, "外部入力（作成元ジョブなし）", None))
                continue

            connected, step1_rej, step2_rej = filter_producers_by_graph(
                G, producers, consumer_unit_full, base_dir_to_remove, predecessor_cache)

            if not connected:
                disconnected_externals.setdefault(current_file, set()).add(consumer_unit_full)
                producer_names = [record_by_full[p]['unit'] if p in record_by_full else os.path.basename(p) for p in sorted(producers)]
                if consumer_unit_full not in layer_discovered:
                    layer_discovered[consumer_unit_full] = []
                layer_discovered[consumer_unit_full].append(
                    (current_file, f"候補: {', '.join(producer_names)} → 全て未接続 → 外部入力", None))
                continue

            def _uname(p):
                return record_by_full[p]['unit'] if p in record_by_full else os.path.basename(p)

            connected_names = [_uname(p) for p in sorted(connected)]

            if len(producers) > 1:
                info = f"候補: {', '.join([_uname(p) for p in sorted(producers)])}"
                if step1_rej:
                    info += f" → 除外(未接続): {', '.join([_uname(p) for p in sorted(step1_rej)])}"
                if step2_rej:
                    info += f" → 除外(順序): {', '.join([_uname(p) for p in sorted(step2_rej)])}"
                info += f" → 採用: {', '.join(connected_names)}"
            else:
                info = f"→ {connected_names[0]} ✓"

            # 並列producer警告: connectedが複数 = 順序不定のproducerが残った
            if len(connected) > 1:
                parallel_warnings.append({
                    'file': current_file,
                    'consumer': consumer_unit_full,
                    'candidates': sorted(connected),
                })

            for unit_full_path in connected:
                if unit_full_path in needed_units_full:
                    if consumer_unit_full not in layer_discovered:
                        layer_discovered[consumer_unit_full] = []
                    layer_discovered[consumer_unit_full].append(
                        (current_file, info + " (既出)", None))
                    continue

                needed_units_full.add(unit_full_path)

                if consumer_unit_full not in layer_discovered:
                    layer_discovered[consumer_unit_full] = []
                layer_discovered[consumer_unit_full].append(
                    (current_file, info, unit_full_path))

                rec = record_by_full.get(unit_full_path)
                new_inputs = [inp for inp in rec.get('inputs', []) if inp] if rec else []
                if new_inputs:
                    for inp in new_inputs:
                        queue.append((inp, unit_full_path, current_layer + 1))
                else:
                    _no_input_units.append((current_layer + 1, unit_full_path))

                # 新規発見ユニットのDB依存も探索
                if unit_full_path not in _db_searched:
                    _db_searched.add(unit_full_path)
                    _search_db_inputs(unit_full_path, current_layer, layer_discovered)

                # DBブリッジ
                if unit_full_path in bridge_map:
                    bridged_names_for_trace = []
                    for bridged in bridge_map[unit_full_path]:
                        if bridged in needed_units_full:
                            continue
                        needed_units_full.add(bridged)
                        bridged_units.add(bridged)
                        b_rec = record_by_full.get(bridged)
                        b_name = b_rec['unit'] if b_rec else os.path.basename(bridged)
                        bridged_names_for_trace.append(b_name)
                        if b_rec:
                            for inp in b_rec.get('inputs', []):
                                if inp:
                                    queue.append((inp, bridged, current_layer + 1))
                    if bridged_names_for_trace:
                        p_name = record_by_full[unit_full_path]['unit'] if unit_full_path in record_by_full else os.path.basename(unit_full_path)
                        if consumer_unit_full not in layer_discovered:
                            layer_discovered[consumer_unit_full] = []
                        layer_discovered[consumer_unit_full].append(
                            (f"[DBブリッジ: {p_name}]",
                             f"DBブリッジ → {', '.join(bridged_names_for_trace)}",
                             None))

    # --- DB依存の除外テーブル ---
    # なぜ除外するか: これらは全JNから参照/更新される共用管理テーブルで、
    # 追跡するとほぼ全ユニットが必要ユニットに入ってしまい結果が意味をなさない。
    # 顧客固有のテーブル名をソースコードに含めないためconfig.jsonから読み込む。
    try:
        import json as _json
        with open(CONFIG_FILE, 'r', encoding='utf-8') as _f:
            _DB_EXCLUDE_TABLES = set(_json.load(_f).get('db_exclude_tables', []))
    except (FileNotFoundError, _json.JSONDecodeError):
        _DB_EXCLUDE_TABLES = set()

    # --- DB依存のproducer探索 ---
    def _search_db_inputs(consumer_unit_full, current_layer, layer_discovered):
        """consumerのDB R/RWテーブルについてW元（db_producer_map）を探索する。"""
        if not db_producer_map:
            return

        consumer_rec = record_by_full.get(consumer_unit_full)
        if not consumer_rec:
            return

        db_tables = consumer_rec.get('db_tables', {})
        if not db_tables:
            return

        for table_name, op in db_tables.items():
            # R または RW のテーブルのみ追跡（Wは読んでないので依存なし）
            if op == 'W':
                continue
            # 除外テーブル
            if table_name in _DB_EXCLUDE_TABLES:
                continue
            # master.* は除外（テーブル名にmaster含む）
            if 'master' in table_name.lower():
                continue

            db_key = f"[DB]{table_name}"
            producers = db_producer_map.get(table_name)
            if producers:
                producers = producers - {consumer_unit_full}

            if not producers:
                # W元がない -> 外部DB
                true_externals.setdefault(db_key, set()).add(consumer_unit_full)
                if consumer_unit_full not in layer_discovered:
                    layer_discovered[consumer_unit_full] = []
                layer_discovered[consumer_unit_full].append(
                    (db_key, f"外部DB（書込元ジョブなし）[{op}]", None))
                continue

            # グラフ上で接続されたproducerを絞り込み
            # DB依存: 全W元が必要なのでStep 2（順序絞り込み）をスキップ
            connected, step1_rej, step2_rej = filter_producers_by_graph(
                G, producers, consumer_unit_full, base_dir_to_remove, predecessor_cache,
                skip_order_filter=True)

            if not connected:
                disconnected_externals.setdefault(db_key, set()).add(consumer_unit_full)
                producer_names = [record_by_full[p]['unit'] if p in record_by_full else os.path.basename(p) for p in sorted(producers)]
                if consumer_unit_full not in layer_discovered:
                    layer_discovered[consumer_unit_full] = []
                layer_discovered[consumer_unit_full].append(
                    (db_key, f"DB候補: {', '.join(producer_names)} → 全て未接続 → 外部DB", None))
                continue

            def _uname(p):
                return record_by_full[p]['unit'] if p in record_by_full else os.path.basename(p)

            connected_names = [_uname(p) for p in sorted(connected)]

            if len(producers) > 1:
                info = f"DB候補: {', '.join([_uname(p) for p in sorted(producers)])}"
                if step1_rej:
                    info += f" → 除外(未接続): {', '.join([_uname(p) for p in sorted(step1_rej)])}"
                if step2_rej:
                    info += f" → 除外(順序): {', '.join([_uname(p) for p in sorted(step2_rej)])}"
                info += f" → 採用: {', '.join(connected_names)}"
            else:
                info = f"→ {connected_names[0]} ✓ [DB:{table_name}]"

            for unit_full_path in connected:
                if unit_full_path in needed_units_full:
                    if consumer_unit_full not in layer_discovered:
                        layer_discovered[consumer_unit_full] = []
                    layer_discovered[consumer_unit_full].append(
                        (db_key, info + " (既出)", None))
                    continue

                needed_units_full.add(unit_full_path)

                if consumer_unit_full not in layer_discovered:
                    layer_discovered[consumer_unit_full] = []
                layer_discovered[consumer_unit_full].append(
                    (db_key, info, unit_full_path))

                rec = record_by_full.get(unit_full_path)
                new_inputs = [inp for inp in rec.get('inputs', []) if inp] if rec else []
                if new_inputs:
                    for inp in new_inputs:
                        queue.append((inp, unit_full_path, current_layer + 1))
                else:
                    _no_input_units.append((current_layer + 1, unit_full_path))

                # DB経由で見つかったユニットのDB依存も再帰的に探索
                if unit_full_path not in _db_searched:
                    _db_searched.add(unit_full_path)
                    _search_db_inputs(unit_full_path, current_layer, layer_discovered)

    # --- trace記録の出力（レイヤー単位） ---
    def _flush_layer(lyr, discovered, label=""):
        no_inputs_this_layer = [(l, uf) for l, uf in _no_input_units if l == lyr]
        if not discovered and not no_inputs_this_layer:
            return
        trace_lines.append("")
        trace_lines.append("=" * 60)
        hdr = f"===== Layer {lyr}"
        if label:
            hdr += f" ({label})"
        hdr += " ====="
        trace_lines.append(hdr)
        trace_lines.append("=" * 60)
        for consumer_uf, file_infos in discovered.items():
            c_rec = record_by_full.get(consumer_uf)
            c_name = c_rec['unit'] if c_rec else os.path.basename(consumer_uf)
            trace_lines.append(f"")
            trace_lines.append(f"  {c_name}  ({consumer_uf})")
            for fpath, info_str, discovered_uf in file_infos:
                trace_lines.append(f"    入力: {fpath}")
                trace_lines.append(f"      {info_str}")
                trace_data.append((lyr, c_name, consumer_uf, fpath, info_str))
        for _, uf in no_inputs_this_layer:
            ni_rec = record_by_full.get(uf)
            ni_name = ni_rec['unit'] if ni_rec else os.path.basename(uf)
            trace_lines.append(f"")
            trace_lines.append(f"  {ni_name}  ({uf})")
            trace_lines.append(f"    → 入力なし（探索終了）")
            trace_data.append((lyr, ni_name, uf, "", "入力なし（探索終了）"))
        for item in no_inputs_this_layer:
            _no_input_units.remove(item)

    # --- Layer 0: 目標ユニットの入力探索 ---
    layer = 0
    layer_discovered = collections.OrderedDict()

    for unit_full in target_unit_fulls:
        needed_units_full.add(unit_full)
        _db_searched.add(unit_full)
        rec = record_by_full.get(unit_full)
        unit_name = rec['unit'] if rec else os.path.basename(unit_full)

        input_files = [f for f in rec.get('inputs', []) if f] if rec else []

        # DBブリッジ: 目標ユニットに紐づくユニットがあれば追加
        if unit_full in bridge_map:
            for bridged in bridge_map[unit_full]:
                if bridged in needed_units_full:
                    continue
                needed_units_full.add(bridged)
                bridged_units.add(bridged)
                b_rec = record_by_full.get(bridged)
                if b_rec:
                    for inp in b_rec.get('inputs', []):
                        if inp:
                            queue.append((inp, bridged, layer + 1))

        if not input_files:
            bridged_targets = bridge_map.get(unit_full, [])
            bridged_in_this = [b for b in bridged_targets if b in bridged_units]
            if bridged_in_this:
                bridge_names = ", ".join(
                    record_by_full[b]['unit'] if b in record_by_full else os.path.basename(b)
                    for b in bridged_in_this)
                trace_data.append((0, unit_name, unit_full, "",
                                   f"目標ユニット（入力なし）/ DBブリッジ → {bridge_names}"))
            else:
                trace_data.append((0, unit_name, unit_full, "", "目標ユニット（入力なし）"))
            continue

        _search_inputs(unit_full, input_files, layer, layer_discovered)
        # DB依存探索
        _search_db_inputs(unit_full, layer, layer_discovered)

    _flush_layer(0, layer_discovered, "目標ユニット")

    # --- Layer 1, 2, 3, ... ---
    layer_discovered = collections.OrderedDict()
    while queue:
        next_layer = queue[0][2]
        if next_layer > layer:
            _flush_layer(layer, layer_discovered)
            layer = next_layer
            layer_discovered = collections.OrderedDict()

        current_file, consumer_unit_full, file_layer = queue.popleft()
        _search_inputs(consumer_unit_full, [current_file], layer, layer_discovered)

        # 新しく発見されたユニットのDB依存も探索
        if consumer_unit_full not in _db_searched:
            _db_searched.add(consumer_unit_full)
            _search_db_inputs(consumer_unit_full, layer, layer_discovered)

        if not queue or queue[0][2] != layer:
            _flush_layer(layer, layer_discovered)
            layer_discovered = collections.OrderedDict()

    if _no_input_units:
        last_layer = max(l for l, _ in _no_input_units)
        _flush_layer(last_layer, collections.OrderedDict())

    _log(f"[Trace] Found {len(needed_units_full)} units, "
         f"{len(true_externals)} true externals, "
         f"{len(disconnected_externals)} disconnected externals.")

    # 並列producer警告がある場合、ユーザーに選択させて不要なproducerを除去
    # ask_choice_fn が渡されていない場合はスキップ（テスト時やCLI実行時）
    if parallel_warnings and ask_choice_fn:
        _log(f"[Warning] {len(parallel_warnings)} parallel producer cases detected")
        user_choices = ask_choice_fn(
            parallel_warnings,
            title="警告: 並列producerの選択",
            description="以下のファイルを出力するユニットが同一ジョブネット内に複数存在し、"
                        "実行順序が定義されていません。使用する出力元を選択してください。",
            root_window=root_window,
            warning="ジョブ定義の先行関係(ar)に問題がある可能性があります。確認を推奨します。")
        # 選択されなかったproducerをneeded_units_fullから除去
        for item in parallel_warnings:
            chosen = user_choices.get(item['file'])
            if chosen:
                for cand in item['candidates']:
                    if cand != chosen and cand not in target_unit_fulls:
                        needed_units_full.discard(cand)

    return {
        'needed_units_full': needed_units_full,
        'true_externals': true_externals,
        'disconnected_externals': disconnected_externals,
        'trace_data': trace_data,
        'trace_lines': trace_lines,
        'bridged_units': bridged_units,
    }


# =====================================================================
# 充足チェック: BFS結果の漏れを補完
# =====================================================================

def _supplement_check(needed_units_full, record_by_full, producer_map,
                      G, base_dir_to_remove, predecessor_cache,
                      true_externals, disconnected_externals, supplemented_units,
                      root_window=None, ask_choice_fn=None):
    """充足チェック: BFS結果セット内の全入力が結果セット内で作成されているか確認する。

    BFSで見つかったジョブの入力ファイルについて、そのproducerが
    needed_units内にいるか確認する。いなければ追加する。
    複数のproducer候補がある場合はユーザーに選択させる。

    Returns:
        supplemented_externals: dict - 追加ユニットの外部入力
    """
    supplemented_externals_local = {}
    check_targets = set(needed_units_full)
    checked = set()

    while check_targets:
        new_additions = set()
        # 複数候補をバッチで収集してまとめてユーザーに聞く
        multi_choice_items = []  # [{file, consumer, candidates}, ...]

        for unit_full in check_targets:
            if unit_full in checked:
                continue
            checked.add(unit_full)

            rec = record_by_full.get(unit_full)
            if not rec:
                continue

            for input_file in rec.get('inputs', []):
                if not input_file:
                    continue

                producers = producer_map.get(input_file)
                if not producers:
                    if unit_full in supplemented_units:
                        supplemented_externals_local.setdefault(input_file, set()).add(unit_full)
                    else:
                        true_externals.setdefault(input_file, set()).add(unit_full)
                    continue

                has_producer_in_set = any(p in needed_units_full for p in producers)
                if has_producer_in_set:
                    continue

                connected, _, _ = filter_producers_by_graph(
                    G, producers, unit_full, base_dir_to_remove, predecessor_cache)

                if not connected:
                    if unit_full in supplemented_units:
                        supplemented_externals_local.setdefault(input_file, set()).add(unit_full)
                    else:
                        disconnected_externals.setdefault(input_file, set()).add(unit_full)
                    continue

                if len(connected) == 1:
                    p = list(connected)[0]
                    if p not in needed_units_full:
                        needed_units_full.add(p)
                        supplemented_units.add(p)
                        new_additions.add(p)
                else:
                    # 複数候補 -> ユーザー選択に回す
                    multi_choice_items.append({
                        'file': input_file,
                        'consumer': unit_full,
                        'candidates': sorted(connected),
                    })

        # 複数候補をまとめてユーザーに聞く
        if multi_choice_items and ask_choice_fn:
            user_choices = ask_choice_fn(
                multi_choice_items,
                title="充足チェック: 作成元の選択",
                description="以下の入力ファイルについて、作成元のユニットを選択してください。",
                root_window=root_window)
            for item in multi_choice_items:
                chosen = user_choices.get(item['file'])
                if not chosen:
                    chosen = item['candidates'][0]
                if chosen not in needed_units_full:
                    needed_units_full.add(chosen)
                    supplemented_units.add(chosen)
                    new_additions.add(chosen)
        elif multi_choice_items:
            # ask_choice_fnがない場合は先頭候補を自動選択
            for item in multi_choice_items:
                chosen = item['candidates'][0]
                if chosen not in needed_units_full:
                    needed_units_full.add(chosen)
                    supplemented_units.add(chosen)
                    new_additions.add(chosen)

        check_targets = new_additions

    _log(f"[Supplement] Added {len(supplemented_units)} units by sufficiency check.")
    return supplemented_externals_local
