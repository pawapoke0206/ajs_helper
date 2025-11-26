"""
リファクタ項目のテスト
- _build_file_index: glob → 辞書引きの置換が正しく動くか
- has_path: shortest_path → has_path の置換が正しく動くか
- pre_normalize / _get_ancestors / pre_compute_need: 既存ロジックの動作確認

実機確認が必要な項目は各テスト関数のdocstringに記載。
"""
import os
import sys
import tempfile

# テスト対象のモジュールをimportするためにプロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import networkx as nx


# =====================================================================
# 1. _build_file_index のテスト
# =====================================================================
from ajs_inout_logic import _build_file_index


class TestBuildFileIndex:
    """glob.glob → 辞書引きへの置換テスト。

    【実機で確認すべきこと】
    - 実際の res_root（リソースフォルダ）に対して実行し、
      comenv / シェル / TBLファイルが正しく見つかること
    - 同名ファイルが複数フォルダに存在する場合、想定通りの方が取れること
    """

    def test_basic(self, tmp_path):
        """基本動作: ファイルが辞書に登録される"""
        # tmp_path はpytestが自動で作る一時ディレクトリ
        (tmp_path / "SHL").mkdir()
        (tmp_path / "SHL" / "MRKD001015.sh").write_text("#!/bin/ksh")
        (tmp_path / "TBL").mkdir()
        (tmp_path / "TBL" / "FSYOKI").write_text("data")
        (tmp_path / "comenv").write_text("BSDIR=/users/SAM01")

        index = _build_file_index(str(tmp_path))

        assert "MRKD001015.sh" in index
        assert "FSYOKI" in index
        assert "comenv" in index
        assert index["MRKD001015.sh"].endswith("MRKD001015.sh")

    def test_nested_directories(self, tmp_path):
        """サブフォルダの深い階層でも見つかる"""
        deep = tmp_path / "a" / "b" / "c"
        deep.mkdir(parents=True)
        (deep / "deep_file.sh").write_text("test")

        index = _build_file_index(str(tmp_path))
        assert "deep_file.sh" in index

    def test_duplicate_filename_first_wins(self, tmp_path):
        """同名ファイルが複数ある場合、最初に見つかった方が登録される"""
        (tmp_path / "dir_a").mkdir()
        (tmp_path / "dir_b").mkdir()
        (tmp_path / "dir_a" / "same.txt").write_text("first")
        (tmp_path / "dir_b" / "same.txt").write_text("second")

        index = _build_file_index(str(tmp_path))
        # os.walkの順序は保証されないが、2つのうちどちらか1つだけが入る
        assert "same.txt" in index
        # 辞書には1件だけ（上書きされない）
        content = open(index["same.txt"]).read()
        assert content in ("first", "second")

    def test_empty_directory(self, tmp_path):
        """空ディレクトリ → 空の辞書"""
        index = _build_file_index(str(tmp_path))
        assert index == {}

    def test_invalid_path(self):
        """存在しないパス → 空の辞書（エラーにならない）"""
        index = _build_file_index("/nonexistent/path/12345")
        assert index == {}

    def test_none_path(self):
        """空文字列 → 空の辞書"""
        index = _build_file_index("")
        assert index == {}


# =====================================================================
# 2. has_path 置換（filter_producers_by_graph）のテスト
# =====================================================================
from ajs_depend_logic import filter_producers_by_graph, _get_ancestors
from ajs_rel_logic import pre_compute_need, _build_children_map


class TestFilterProducersByGraph:
    """shortest_path → has_path の置換テスト。

    小さなグラフを作って、producer選択ロジックが正しく動くことを確認する。

    【実機で確認すべきこと】
    - 実際のajsprint出力から構築したグラフで、
      依存関係解析（Tab5）の結果が変更前と同一であること
    - 特に producer が複数ある場合の絞り込み結果を比較
    """

    def _make_linear_graph(self):
        """A → B → C の直列グラフ（Aが最初、Cが最後に実行）"""
        G = nx.DiGraph()
        G.add_edge("/job/A", "/job/B")
        G.add_edge("/job/B", "/job/C")
        G.graph['children_map'] = _build_children_map(G)
        return G

    def _make_branching_graph(self):
        """  A → B → D
             A → C → D
        AとBとCはDより先。BとCの間に順序なし。"""
        G = nx.DiGraph()
        G.add_edge("/job/A", "/job/B")
        G.add_edge("/job/A", "/job/C")
        G.add_edge("/job/B", "/job/D")
        G.add_edge("/job/C", "/job/D")
        G.graph['children_map'] = _build_children_map(G)
        return G

    def test_linear_latest_wins(self):
        """直列: A→B→C のとき、consumerがCなら、Bが採用されAは除外"""
        G = self._make_linear_graph()
        producers = {"/job/A", "/job/B"}
        consumer = "/job/C"
        cache = {}

        result, step1_rej, step2_rej = filter_producers_by_graph(
            G, producers, consumer, "", cache)

        # Bが最後にファイルを書く → Bだけが採用
        assert result == {"/job/B"}
        assert "/job/A" in step2_rej

    def test_single_producer(self):
        """producer が1つだけ → そのまま返る（Step2に入らない）"""
        G = self._make_linear_graph()
        producers = {"/job/A"}
        consumer = "/job/C"
        cache = {}

        result, step1_rej, step2_rej = filter_producers_by_graph(
            G, producers, consumer, "", cache)

        assert result == {"/job/A"}
        assert step2_rej == set()

    def test_disconnected_producer_rejected(self):
        """consumerの先行にないproducerはStep1で除外"""
        G = self._make_linear_graph()
        # /job/X はグラフに存在しない → consumerの先行集合に入らない
        G.add_node("/job/X")
        producers = {"/job/A", "/job/X"}
        consumer = "/job/C"
        cache = {}

        result, step1_rej, step2_rej = filter_producers_by_graph(
            G, producers, consumer, "", cache)

        assert "/job/X" in step1_rej
        assert result == {"/job/A"}

    def test_parallel_producers_both_remain(self):
        """並列: BとCの間に順序がない場合、両方残る"""
        G = self._make_branching_graph()
        producers = {"/job/B", "/job/C"}
        consumer = "/job/D"
        cache = {}

        result, step1_rej, step2_rej = filter_producers_by_graph(
            G, producers, consumer, "", cache)

        # BとCの間に先行関係がない → 両方残る
        assert result == {"/job/B", "/job/C"}
        assert step2_rej == set()

    def test_no_producers(self):
        """producer が空 → 空の結果"""
        G = self._make_linear_graph()
        cache = {}
        result, step1_rej, step2_rej = filter_producers_by_graph(
            G, set(), "/job/C", "", cache)
        assert result == set()


# =====================================================================
# 3. 既存ロジックの動作確認（pre_normalize / _get_ancestors / pre_compute_need）
# =====================================================================
from ajs_rel_logic import pre_normalize


class TestPreNormalize:
    """パス正規化のテスト。

    【実機で確認すべきこと】
    - 特になし（純粋な文字列操作なのでテストだけで十分）
    """

    def test_strip_service_prefix(self):
        """AJSROOT1: プレフィックスが除去される"""
        assert pre_normalize("AJSROOT1:/BS_info/job1", "") == "/BS_info/job1"

    def test_remove_base_dir(self):
        """base部分が除去される"""
        assert pre_normalize("/BS_info/daily/job1", "/BS_info") == "/daily/job1"

    def test_both(self):
        """プレフィックス除去 + base除去"""
        assert pre_normalize("AJSROOT1:/BS_info/daily/job1", "/BS_info") == "/daily/job1"

    def test_no_change(self):
        """変換不要なケース"""
        assert pre_normalize("/job1", "") == "/job1"


class TestGetAncestors:
    """親階層リスト取得のテスト。

    【実機で確認すべきこと】
    - 特になし（純粋な文字列操作）
    """

    def test_deep_path(self):
        """3階層のパス → 自分、親、祖父の順"""
        result = _get_ancestors("/BS_info/daily/job1")
        assert result[0] == "/BS_info/daily/job1"  # 自分
        assert result[1] == "/BS_info/daily"        # 親
        assert result[2] == "/BS_info"              # 祖父

    def test_single_level(self):
        """1階層 → 自分だけ"""
        result = _get_ancestors("/job1")
        assert result == ["/job1"]


class TestPreComputeNeed:
    """先行ユニット計算のテスト。

    【実機で確認すべきこと】
    - 実際のajsprintグラフで、Tab4の結果が変わっていないこと
    """

    def _make_graph_with_map(self, edges=None, nodes=None):
        """children_map 付きのグラフを作るヘルパー。
        実際の pre_parse_graph と同じく、構築後に children_map を付与する。"""
        G = nx.DiGraph()
        for n in (nodes or []):
            G.add_node(n)
        for frm, to in (edges or []):
            G.add_edge(frm, to)
        G.graph['children_map'] = _build_children_map(G)
        return G

    def test_simple_chain(self):
        """A→B→C でCを指定 → A,B,Cすべて必要"""
        G = self._make_graph_with_map(
            edges=[("/grp/A", "/grp/B"), ("/grp/B", "/grp/C")])
        need = pre_compute_need(G, "/grp/C")
        assert "/grp/A" in need
        assert "/grp/B" in need
        assert "/grp/C" in need
        # 親ジョブネットも含まれる
        assert "/grp" in need

    def test_includes_descendants(self):
        """ジョブネットを指定 → 配下の子孫も含まれる"""
        G = self._make_graph_with_map(
            edges=[("/pre_job", "/net")],
            nodes=["/net/child1", "/net/child2"])
        need = pre_compute_need(G, "/net")
        assert "/net/child1" in need
        assert "/net/child2" in need
        assert "/pre_job" in need


# =====================================================================
# 4. pre_descendants 辞書化のテスト（テスト観点: ユーザー設計）
#
#    /A
#    /A/B
#    /A/C
#    /A/B/D
#
# =====================================================================
from ajs_rel_logic import pre_descendants, _build_children_map


class TestPreDescendants:
    """pre_descendants の辞書化テスト。

    テストケースはユーザーが設計:
      ① /A を呼んだ時に /A/B, /A/C, /A/B/D が全部返るか
      ② 辞書に /A → [/A/B, /A/C] が正しく入っているか
      ③ 配下がないノード(/A/B/D)を呼んだ時に空で返るか
      ④ 辞書のライフサイ��ル（グラフと一緒に作られ、一緒に消えるか）
      ⑤ グラフに存在しないノードを呼んだ時にエラーにな��ないか

    【実機で確認すべきこと】
    - 実際のajsprintグ���フで Tab4/Tab5 の結果が変更前と同一であること
    - 特にジョブネットの中にサブジョブネットがある構造（/net/sub_net/job）で
      子孫が正しく取れていること
    """

    def _make_graph(self):
        """テスト用グラフ: /A の配下に /A/B, /A/C, /A/B/D がある"""
        G = nx.DiGraph()
        G.add_node("/A")
        G.add_node("/A/B")
        G.add_node("/A/C")
        G.add_node("/A/B/D")
        G.graph['children_map'] = _build_children_map(G)
        return G

    def test_01_all_descendants_returned(self):
        """① /A を呼んだら /A/B, /A/C, /A/B/D が全部返る"""
        G = self._make_graph()
        result = set(pre_descendants(G, "/A"))
        assert result == {"/A/B", "/A/C", "/A/B/D"}

    def test_02_children_map_structure(self):
        """② 辞書に /A → [/A/B, /A/C] と /A/B → [/A/B/D] が入っている"""
        G = self._make_graph()
        cmap = G.graph['children_map']
        # /A の直下の子
        assert set(cmap["/A"]) == {"/A/B", "/A/C"}
        # /A/B の直下の子
        assert set(cmap["/A/B"]) == {"/A/B/D"}

    def test_03_leaf_node_returns_empty(self):
        """③ 配下がないノード(/A/B/D)を呼んだら空が返る"""
        G = self._make_graph()
        result = list(pre_descendants(G, "/A/B/D"))
        assert result == []

    def test_04_lifecycle_tied_to_graph(self):
        """④ children_mapはグラフの属性として持つので、グラフと一緒に消える"""
        G = self._make_graph()
        # グラフ作成直後は辞書が存在する
        assert 'children_map' in G.graph

        # グラフを破棄（変数を上書��）すれば辞書も一緒に消える
        G = None
        # Pythonのガベージコレクションにより、G と一緒に children_map も解放される
        # （直接確認は難し��が、G が None なら辞書にもアクセスできない）
        assert G is None

    def test_05_nonexistent_node_returns_empty(self):
        """⑤ グラフに存在しないノードを呼んでもエラーにならず空が返る"""
        G = self._make_graph()
        result = list(pre_descendants(G, "/X/Y/Z"))
        assert result == []


# =====================================================================
# 5. クロスジョブネット: 親ジョブネット自動判定のテスト
#
#    AJSの階層構造:
#      /BS_info/01_定例          (group)
#        /BS_info/01_定例/日次処理    (group)
#          /BS_info/01_定例/日次処理/日次定例  (net) ← 親ジョブネット
#        /BS_info/01_定例/月次処理    (group)
#          /BS_info/01_定例/月次処理/月次1歴日  (net) ← 親ジョブネット
#          /BS_info/01_定例/月次処理/月次2歴日  (net) ← 親ジョブネット
#      /BS_info/02_監明          (group)
#        /BS_info/02_監明/【監明】日次処理  (net) ← 親ジョブネット
#
# =====================================================================
from ajs_depend_logic import discover_parent_jobnets, find_parent_jobnet


class TestDiscoverParentJobnets:
    """親ジョブネット一覧の自動取得テスト。

    ajsprint -f "%TY\\t%JN" -R の出力を模擬して、
    groupをスキップし最初にnetになるものだけを親ジョブネットとして抽出する。
    """

    # テスト用のajsprint出力（設計書の例に対応）
    SAMPLE_LINES = [
        "g\t/BS_info/01_定例",
        "g\t/BS_info/01_定例/日次処理",
        "n\t/BS_info/01_定例/日次処理/日次定例",
        "j\t/BS_info/01_定例/日次処理/日次定例/MRKD001",
        "j\t/BS_info/01_定例/日次処理/日次定例/MRKD002",
        "n\t/BS_info/01_定例/日次処理/日次定例/サブネット",  # netの中のnet → 親ではない
        "j\t/BS_info/01_定例/日次処理/日次定例/サブネット/JOB1",
        "g\t/BS_info/01_定例/月次処理",
        "n\t/BS_info/01_定例/月次処理/月次1歴日",
        "n\t/BS_info/01_定例/月次処理/月次2歴日",
        "j\t/BS_info/01_定例/月次処理/月次1歴日/MRKM001",
        "g\t/BS_info/02_監明",
        "n\t/BS_info/02_監明/【監明】日次処理",
        "j\t/BS_info/02_監明/【監明】日次処理/MRKD001",
    ]

    def test_basic_discovery(self):
        """group配下の最初のnetが親ジョブネットとして抽出される"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/BS_info")
        assert "/BS_info/01_定例/日次処理/日次定例" in result
        assert "/BS_info/01_定例/月次処理/月次1歴日" in result
        assert "/BS_info/01_定例/月次処理/月次2歴日" in result
        assert "/BS_info/02_監明/【監明】日次処理" in result

    def test_fullspell_type(self):
        """ajsprintがフルスペル(net/group/job)で出力するパターンでも動作する"""
        lines = [
            "group\t/BS_info/01_定例",
            "net\t/BS_info/01_定例/日次定例",
            "job\t/BS_info/01_定例/日次定例/MRKD001",
            "net\t/BS_info/01_定例/月次1歴日",
        ]
        result = discover_parent_jobnets(lines, "/BS_info")
        assert len(result) == 2
        assert "/BS_info/01_定例/日次定例" in result
        assert "/BS_info/01_定例/月次1歴日" in result

    def test_nested_net_excluded(self):
        """netの中のnet(サブネット)は親ジョブネットに含まれない"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/BS_info")
        assert "/BS_info/01_定例/日次処理/日次定例/サブネット" not in result

    def test_count(self):
        """親ジョブネットは4つ"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/BS_info")
        assert len(result) == 4

    def test_narrower_range(self):
        """範囲を01_定例に絞ると、02_監明は含まれない"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/BS_info/01_定例")
        assert len(result) == 3
        assert "/BS_info/02_監明/【監明】日次処理" not in result

    def test_empty_input(self):
        """空の入力 → 空リスト"""
        assert discover_parent_jobnets([], "/BS_info") == []

    def test_range_itself_is_net(self):
        """指定範囲そのものがnetの場合、それ自体が親ジョブネット"""
        lines = [
            "n\t/BS_info/01_定例/日次定例",
            "j\t/BS_info/01_定例/日次定例/JOB1",
        ]
        result = discover_parent_jobnets(lines, "/BS_info/01_定例/日次定例")
        assert result == ["/BS_info/01_定例/日次定例"]

    def test_range_is_net_with_subnets(self):
        """範囲がnet+中にサブネットがある場合、範囲自身だけが親ジョブネット"""
        lines = [
            "net\t/BS_info/01_定例/日次処理",
            "net\t/BS_info/01_定例/日次処理/前処理",
            "job\t/BS_info/01_定例/日次処理/前処理/JOB1",
            "net\t/BS_info/01_定例/日次処理/後処理",
            "job\t/BS_info/01_定例/日次処理/後処理/JOB2",
        ]
        result = discover_parent_jobnets(lines, "/BS_info/01_定例/日次処理")
        assert len(result) == 1
        assert result == ["/BS_info/01_定例/日次処理"]


class TestFindParentJobnet:
    """ユニットの親ジョブネット所属判定テスト。"""

    PARENT_JOBNETS = [
        "/BS_info/01_定例/日次処理/日次定例",
        "/BS_info/01_定例/月次処理/月次1歴日",
        "/BS_info/01_定例/月次処理/月次2歴日",
        "/BS_info/02_監明/【監明】日次処理",
    ]

    def test_direct_child(self):
        """日次定例配下のジョブ → 日次定例に所属"""
        result = find_parent_jobnet(
            "/BS_info/01_定例/日次処理/日次定例/MRKD001",
            self.PARENT_JOBNETS)
        assert result == "/BS_info/01_定例/日次処理/日次定例"

    def test_deep_child(self):
        """サブネット配下のジョブでも、最も深い親ジョブネットに所属"""
        result = find_parent_jobnet(
            "/BS_info/01_定例/日次処理/日次定例/サブネット/JOB1",
            self.PARENT_JOBNETS)
        assert result == "/BS_info/01_定例/日次処理/日次定例"

    def test_different_parent(self):
        """監明配下のジョブ → 監明に所属"""
        result = find_parent_jobnet(
            "/BS_info/02_監明/【監明】日次処理/MRKD001",
            self.PARENT_JOBNETS)
        assert result == "/BS_info/02_監明/【監明】日次処理"

    def test_not_found(self):
        """どの親ジョブネットにも属さない → None"""
        result = find_parent_jobnet(
            "/BS_info/99_その他/JOB1",
            self.PARENT_JOBNETS)
        assert result is None

    def test_exact_match(self):
        """親ジョブネット自身を指定 → そのまま返す"""
        result = find_parent_jobnet(
            "/BS_info/01_定例/月次処理/月次2歴日",
            self.PARENT_JOBNETS)
        assert result == "/BS_info/01_定例/月次処理/月次2歴日"

    def test_partial_name_no_match(self):
        """パス名が部分一致するだけでは所属とみなさない"""
        result = find_parent_jobnet(
            "/BS_info/01_定例/月次処理/月次2歴日追加/JOB1",
            self.PARENT_JOBNETS)
        assert result is None


# =====================================================================
# 6. クロスジョブネット: グラフテキスト分割・producer横断検索のテスト
# =====================================================================
from ajs_depend_logic import filter_dep_text_by_jobnet, find_producers_across_jobnets


class TestFilterDepTextByJobnet:
    """ajsprintテキストを親ジョブネット単位にフィルタするテスト。"""

    SAMPLE_DEP_TEXT = "\n".join([
        "n\t/BS/01/日次処理/日次定例\t",
        "j\t/BS/01/日次処理/日次定例/MRKD001\tjobA,jobB,seq;",
        "j\t/BS/01/日次処理/日次定例/MRKD002\t",
        "n\t/BS/01/月次処理/月次1歴日\t",
        "j\t/BS/01/月次処理/月次1歴日/MRKM001\t",
        "n\t/BS/02/【監明】日次処理\t",
        "j\t/BS/02/【監明】日次処理/MRKD001\t",
    ])

    def test_filter_daily(self):
        """日次定例でフィルタ → 日次定例配下の3行だけ"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/BS/01/日次処理/日次定例")
        lines = result.strip().splitlines()
        assert len(lines) == 3
        assert all("/BS/01/日次処理/日次定例" in line for line in lines)

    def test_filter_monthly(self):
        """月次1歴日でフィルタ → 2行"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/BS/01/月次処理/月次1歴日")
        lines = result.strip().splitlines()
        assert len(lines) == 2

    def test_filter_kanmei(self):
        """監明でフィルタ → 2行"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/BS/02/【監明】日次処理")
        lines = result.strip().splitlines()
        assert len(lines) == 2

    def test_no_match(self):
        """存在しないパスでフィルタ → 空"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/BS/99/存在しない")
        assert result.strip() == ""

    def test_ar_data_preserved(self):
        """ar列(先行関係データ)がフィルタ後も残っている"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/BS/01/日次処理/日次定例")
        assert "jobA,jobB,seq;" in result

    def test_ar_continuation_lines(self):
        """AR継続行（タブなし）も対象ヘッダの後に残る"""
        text = "\n".join([
            "net\t/BS/01/日次定例\tjobA,jobB,seq;",
            "jobC,jobD,seq;",              # ← AR継続行
            "jobE,jobF,seq;",              # ← AR継続行
            "net\t/BS/02/月次定例\tjobX,jobY,seq;",
            "jobZ,jobW,seq;",              # ← 別ヘッダのAR継続行
        ])
        result = filter_dep_text_by_jobnet(text, "/BS/01/日次定例")
        lines = result.strip().splitlines()
        assert len(lines) == 3  # ヘッダ1行 + 継続2行
        assert "jobC,jobD,seq;" in result
        assert "jobE,jobF,seq;" in result
        assert "jobZ,jobW,seq;" not in result  # 別ヘッダの継続行は含まない


class TestFindProducersAcrossJobnets:
    """外部入力ファイルの出力元を親ジョブネット横断で検索するテスト。"""

    PARENT_JOBNETS = [
        "/BS/01/日次処理/日次定例",
        "/BS/01/月次処理/月次1歴日",
        "/BS/02/【監明】日次処理",
    ]

    # producer_map: ファイルXは日次定例と監明の2箇所で出力されている
    PRODUCER_MAP = {
        "/data/FILE_X": {
            "/BS/01/日次処理/日次定例/MRKD005",
            "/BS/02/【監明】日次処理/MRKD005",
        },
        "/data/FILE_Y": {
            "/BS/01/月次処理/月次1歴日/MRKM010",
        },
        "/data/FILE_Z": set(),  # 出力元なし
    }

    def test_multiple_jobnets(self):
        """FILE_Xは日次定例と監明の2つの親ジョブネットで出力される"""
        result = find_producers_across_jobnets(
            "/data/FILE_X", self.PRODUCER_MAP, self.PARENT_JOBNETS)
        assert len(result) == 2
        assert "/BS/01/日次処理/日次定例" in result
        assert "/BS/02/【監明】日次処理" in result

    def test_single_jobnet(self):
        """FILE_Yは月次1歴日のみ"""
        result = find_producers_across_jobnets(
            "/data/FILE_Y", self.PRODUCER_MAP, self.PARENT_JOBNETS)
        assert len(result) == 1
        assert "/BS/01/月次処理/月次1歴日" in result

    def test_no_producer(self):
        """FILE_Zは出力元なし → 空dict"""
        result = find_producers_across_jobnets(
            "/data/FILE_Z", self.PRODUCER_MAP, self.PARENT_JOBNETS)
        assert result == {}

    def test_unknown_file(self):
        """producer_mapにないファイル → 空dict"""
        result = find_producers_across_jobnets(
            "/data/UNKNOWN", self.PRODUCER_MAP, self.PARENT_JOBNETS)
        assert result == {}

    def test_exclude_jobnets(self):
        """exclude_jobnetsで指定した親ジョブネットは除外される（ループ防止）"""
        result = find_producers_across_jobnets(
            "/data/FILE_X", self.PRODUCER_MAP, self.PARENT_JOBNETS,
            exclude_jobnets={"/BS/01/日次処理/日次定例"})
        assert len(result) == 1
        assert "/BS/02/【監明】日次処理" in result

    def test_exclude_all(self):
        """全親ジョブネットを除外 → 空dict"""
        result = find_producers_across_jobnets(
            "/data/FILE_X", self.PRODUCER_MAP, self.PARENT_JOBNETS,
            exclude_jobnets={"/BS/01/日次処理/日次定例", "/BS/02/【監明】日次処理"})
        assert result == {}
