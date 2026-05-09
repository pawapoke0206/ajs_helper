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
        (tmp_path / "SHL" / "JOB001.sh").write_text("#!/bin/ksh")
        (tmp_path / "TBL").mkdir()
        (tmp_path / "TBL" / "DATA_TBL").write_text("data")
        (tmp_path / "comenv").write_text("APPDIR=/users/APP01")

        index = _build_file_index(str(tmp_path))

        assert "JOB001.sh" in index
        assert "DATA_TBL" in index
        assert "comenv" in index
        assert index["JOB001.sh"].endswith("JOB001.sh")

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
      依存関係解析の結果が変更前と同一であること
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
        """SVC1: プレフィックスが除去される"""
        assert pre_normalize("SVC1:/PROD/job1", "") == "/PROD/job1"

    def test_remove_base_dir(self):
        """base部分が除去される"""
        assert pre_normalize("/PROD/daily/job1", "/PROD") == "/daily/job1"

    def test_both(self):
        """プレフィックス除去 + base除去"""
        assert pre_normalize("SVC1:/PROD/daily/job1", "/PROD") == "/daily/job1"

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
        result = _get_ancestors("/PROD/daily/job1")
        assert result[0] == "/PROD/daily/job1"  # 自分
        assert result[1] == "/PROD/daily"        # 親
        assert result[2] == "/PROD"              # 祖父

    def test_single_level(self):
        """1階層 → 自分だけ"""
        result = _get_ancestors("/job1")
        assert result == ["/job1"]


class TestPreComputeNeed:
    """先行ユニット計算のテスト。

    【実機で確認すべきこと】
    - 実際のajsprintグラフで、結果が変わっていないこと
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
      (1) /A を呼んだ時に /A/B, /A/C, /A/B/D が全部返るか
      (2) 辞書に /A → [/A/B, /A/C] が正しく入っているか
      (3) 配下がないノード(/A/B/D)を呼んだ時に空で返るか
      (4) 辞書のライフサイクル（グラフと一緒に作られ、一緒に消えるか）
      (5) グラフに存在しないノードを呼んだ時にエラーにならないか

    【実機で確認すべきこと】
    - 実際のajsprintグラフで結果が変更前と同一であること
    - 特にジョブネットの中にサブジョブネットがある構造で
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
        """(1) /A を呼んだら /A/B, /A/C, /A/B/D が全部返る"""
        G = self._make_graph()
        result = set(pre_descendants(G, "/A"))
        assert result == {"/A/B", "/A/C", "/A/B/D"}

    def test_02_children_map_structure(self):
        """(2) 辞書に /A → [/A/B, /A/C] と /A/B → [/A/B/D] が入っている"""
        G = self._make_graph()
        cmap = G.graph['children_map']
        # /A の直下の子
        assert set(cmap["/A"]) == {"/A/B", "/A/C"}
        # /A/B の直下の子
        assert set(cmap["/A/B"]) == {"/A/B/D"}

    def test_03_leaf_node_returns_empty(self):
        """(3) 配下がないノード(/A/B/D)を呼んだら空が返る"""
        G = self._make_graph()
        result = list(pre_descendants(G, "/A/B/D"))
        assert result == []

    def test_04_lifecycle_tied_to_graph(self):
        """(4) children_mapはグラフの属性として持つので、グラフと一緒に消える"""
        G = self._make_graph()
        # グラフ作成直後は辞書が存在する
        assert 'children_map' in G.graph

        # グラフを破棄（変数を上書き）すれば辞書も一緒に消える
        G = None
        # Pythonのガベージコレクションにより、G と一緒に children_map も解放される
        assert G is None

    def test_05_nonexistent_node_returns_empty(self):
        """(5) グラフに存在しないノードを呼んでもエラーにならず空が返る"""
        G = self._make_graph()
        result = list(pre_descendants(G, "/X/Y/Z"))
        assert result == []


# =====================================================================
# 5. クロスジョブネット: 親ジョブネット自動判定のテスト
#
#    AJSの階層構造（テスト用ダミー）:
#      /PROD/01_regular            (group)
#        /PROD/01_regular/daily_proc    (group)
#          /PROD/01_regular/daily_proc/daily_batch  (net) ← 親JN
#        /PROD/01_regular/monthly_proc  (group)
#          /PROD/01_regular/monthly_proc/monthly_1st  (net) ← 親JN
#          /PROD/01_regular/monthly_proc/monthly_2nd  (net) ← 親JN
#      /PROD/02_audit              (group)
#        /PROD/02_audit/audit_daily  (net) ← 親JN
#
# =====================================================================
from ajs_depend_logic import discover_parent_jobnets, find_parent_jobnet


class TestDiscoverParentJobnets:
    """親ジョブネット一覧の自動取得テスト。

    ajsprint -f "%TY\\t%JN" -R の出力を模擬して、
    groupをスキップし最初にnetになるものだけを親ジョブネットとして抽出する。
    """

    # テスト用のajsprint出力
    SAMPLE_LINES = [
        "g\t/PROD/01_regular",
        "g\t/PROD/01_regular/daily_proc",
        "n\t/PROD/01_regular/daily_proc/daily_batch",
        "j\t/PROD/01_regular/daily_proc/daily_batch/JOB001",
        "j\t/PROD/01_regular/daily_proc/daily_batch/JOB002",
        "n\t/PROD/01_regular/daily_proc/daily_batch/subnet",  # netの中のnet → 親ではない
        "j\t/PROD/01_regular/daily_proc/daily_batch/subnet/JOB1",
        "g\t/PROD/01_regular/monthly_proc",
        "n\t/PROD/01_regular/monthly_proc/monthly_1st",
        "n\t/PROD/01_regular/monthly_proc/monthly_2nd",
        "j\t/PROD/01_regular/monthly_proc/monthly_1st/JOB010",
        "g\t/PROD/02_audit",
        "n\t/PROD/02_audit/audit_daily",
        "j\t/PROD/02_audit/audit_daily/JOB001",
    ]

    def test_basic_discovery(self):
        """group配下の最初のnetが親ジョブネットとして抽出される"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/PROD")
        assert "/PROD/01_regular/daily_proc/daily_batch" in result
        assert "/PROD/01_regular/monthly_proc/monthly_1st" in result
        assert "/PROD/01_regular/monthly_proc/monthly_2nd" in result
        assert "/PROD/02_audit/audit_daily" in result

    def test_fullspell_type(self):
        """ajsprintがフルスペル(net/group/job)で出力するパターンでも動作する"""
        lines = [
            "group\t/PROD/01_regular",
            "net\t/PROD/01_regular/daily_batch",
            "job\t/PROD/01_regular/daily_batch/JOB001",
            "net\t/PROD/01_regular/monthly_1st",
        ]
        result = discover_parent_jobnets(lines, "/PROD")
        assert len(result) == 2
        assert "/PROD/01_regular/daily_batch" in result
        assert "/PROD/01_regular/monthly_1st" in result

    def test_nested_net_excluded(self):
        """netの中のnet(サブネット)は親ジョブネットに含まれない"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/PROD")
        assert "/PROD/01_regular/daily_proc/daily_batch/subnet" not in result

    def test_count(self):
        """親ジョブネットは4つ"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/PROD")
        assert len(result) == 4

    def test_narrower_range(self):
        """範囲を01_regularに絞ると、02_auditは含まれない"""
        result = discover_parent_jobnets(self.SAMPLE_LINES, "/PROD/01_regular")
        assert len(result) == 3
        assert "/PROD/02_audit/audit_daily" not in result

    def test_empty_input(self):
        """空の入力 → 空リスト"""
        assert discover_parent_jobnets([], "/PROD") == []

    def test_range_itself_is_net(self):
        """指定範囲そのものがnetの場合、それ自体が親ジョブネット"""
        lines = [
            "n\t/PROD/01_regular/daily_batch",
            "j\t/PROD/01_regular/daily_batch/JOB1",
        ]
        result = discover_parent_jobnets(lines, "/PROD/01_regular/daily_batch")
        assert result == ["/PROD/01_regular/daily_batch"]

    def test_range_is_net_with_subnets(self):
        """範囲がnet+中にサブネットがある場合、範囲自身だけが親ジョブネット"""
        lines = [
            "net\t/PROD/01_regular/daily_proc",
            "net\t/PROD/01_regular/daily_proc/pre_proc",
            "job\t/PROD/01_regular/daily_proc/pre_proc/JOB1",
            "net\t/PROD/01_regular/daily_proc/post_proc",
            "job\t/PROD/01_regular/daily_proc/post_proc/JOB2",
        ]
        result = discover_parent_jobnets(lines, "/PROD/01_regular/daily_proc")
        assert len(result) == 1
        assert result == ["/PROD/01_regular/daily_proc"]


class TestFindParentJobnet:
    """ユニットの親ジョブネット所属判定テスト。"""

    PARENT_JOBNETS = [
        "/PROD/01_regular/daily_proc/daily_batch",
        "/PROD/01_regular/monthly_proc/monthly_1st",
        "/PROD/01_regular/monthly_proc/monthly_2nd",
        "/PROD/02_audit/audit_daily",
    ]

    def test_direct_child(self):
        """daily_batch配下のジョブ → daily_batchに所属"""
        result = find_parent_jobnet(
            "/PROD/01_regular/daily_proc/daily_batch/JOB001",
            self.PARENT_JOBNETS)
        assert result == "/PROD/01_regular/daily_proc/daily_batch"

    def test_deep_child(self):
        """サブネット配下のジョブでも、最も深い親ジョブネットに所属"""
        result = find_parent_jobnet(
            "/PROD/01_regular/daily_proc/daily_batch/subnet/JOB1",
            self.PARENT_JOBNETS)
        assert result == "/PROD/01_regular/daily_proc/daily_batch"

    def test_different_parent(self):
        """audit配下のジョブ → audit_dailyに所属"""
        result = find_parent_jobnet(
            "/PROD/02_audit/audit_daily/JOB001",
            self.PARENT_JOBNETS)
        assert result == "/PROD/02_audit/audit_daily"

    def test_not_found(self):
        """どの親ジョブネットにも属さない → None"""
        result = find_parent_jobnet(
            "/PROD/99_other/JOB1",
            self.PARENT_JOBNETS)
        assert result is None

    def test_exact_match(self):
        """親ジョブネット自身を指定 → そのまま返す"""
        result = find_parent_jobnet(
            "/PROD/01_regular/monthly_proc/monthly_2nd",
            self.PARENT_JOBNETS)
        assert result == "/PROD/01_regular/monthly_proc/monthly_2nd"

    def test_partial_name_no_match(self):
        """パス名が部分一致するだけでは所属とみなさない"""
        result = find_parent_jobnet(
            "/PROD/01_regular/monthly_proc/monthly_2nd_extra/JOB1",
            self.PARENT_JOBNETS)
        assert result is None


# =====================================================================
# 6. クロスジョブネット: グラフテキスト分割・producer横断検索のテスト
# =====================================================================
from ajs_depend_logic import filter_dep_text_by_jobnet, find_producers_across_jobnets


class TestFilterDepTextByJobnet:
    """ajsprintテキストを親ジョブネット単位にフィルタするテスト。"""

    SAMPLE_DEP_TEXT = "\n".join([
        "n\t/NS/01/daily_proc/daily_batch\t",
        "j\t/NS/01/daily_proc/daily_batch/JOB001\tjobA,jobB,seq;",
        "j\t/NS/01/daily_proc/daily_batch/JOB002\t",
        "n\t/NS/01/monthly_proc/monthly_1st\t",
        "j\t/NS/01/monthly_proc/monthly_1st/JOB010\t",
        "n\t/NS/02/audit_daily\t",
        "j\t/NS/02/audit_daily/JOB001\t",
    ])

    def test_filter_daily(self):
        """daily_batchでフィルタ → daily_batch配下の3行だけ"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/NS/01/daily_proc/daily_batch")
        lines = result.strip().splitlines()
        assert len(lines) == 3
        assert all("/NS/01/daily_proc/daily_batch" in line for line in lines)

    def test_filter_monthly(self):
        """monthly_1stでフィルタ → 2行"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/NS/01/monthly_proc/monthly_1st")
        lines = result.strip().splitlines()
        assert len(lines) == 2

    def test_filter_audit(self):
        """audit_dailyでフィルタ → 2行"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/NS/02/audit_daily")
        lines = result.strip().splitlines()
        assert len(lines) == 2

    def test_no_match(self):
        """存在しないパスでフィルタ → 空"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/NS/99/nonexistent")
        assert result.strip() == ""

    def test_ar_data_preserved(self):
        """ar列(先行関係データ)がフィルタ後も残っている"""
        result = filter_dep_text_by_jobnet(
            self.SAMPLE_DEP_TEXT, "/NS/01/daily_proc/daily_batch")
        assert "jobA,jobB,seq;" in result

    def test_ar_continuation_lines(self):
        """AR継続行（タブなし）も対象ヘッダの後に残る"""
        text = "\n".join([
            "net\t/NS/01/daily_batch\tjobA,jobB,seq;",
            "jobC,jobD,seq;",              # ← AR継続行
            "jobE,jobF,seq;",              # ← AR継続行
            "net\t/NS/02/monthly_batch\tjobX,jobY,seq;",
            "jobZ,jobW,seq;",              # ← 別ヘッダのAR継続行
        ])
        result = filter_dep_text_by_jobnet(text, "/NS/01/daily_batch")
        lines = result.strip().splitlines()
        assert len(lines) == 3  # ヘッダ1行 + 継続2行
        assert "jobC,jobD,seq;" in result
        assert "jobE,jobF,seq;" in result
        assert "jobZ,jobW,seq;" not in result  # 別ヘッダの継続行は含まない


class TestFindProducersAcrossJobnets:
    """外部入力ファイルの出力元を親ジョブネット横断で検索するテスト。"""

    PARENT_JOBNETS = [
        "/NS/01/daily_proc/daily_batch",
        "/NS/01/monthly_proc/monthly_1st",
        "/NS/02/audit_daily",
    ]

    # producer_map: ファイルXはdaily_batchとaudit_dailyの2箇所で出力されている
    PRODUCER_MAP = {
        "/data/FILE_X": {
            "/NS/01/daily_proc/daily_batch/JOB005",
            "/NS/02/audit_daily/JOB005",
        },
        "/data/FILE_Y": {
            "/NS/01/monthly_proc/monthly_1st/JOB010",
        },
        "/data/FILE_Z": set(),  # 出力元なし
    }

    def test_multiple_jobnets(self):
        """FILE_Xはdaily_batchとaudit_dailyの2つの親JNで出力される"""
        result = find_producers_across_jobnets(
            "/data/FILE_X", self.PRODUCER_MAP, self.PARENT_JOBNETS)
        assert len(result) == 2
        assert "/NS/01/daily_proc/daily_batch" in result
        assert "/NS/02/audit_daily" in result

    def test_single_jobnet(self):
        """FILE_Yはmonthly_1stのみ"""
        result = find_producers_across_jobnets(
            "/data/FILE_Y", self.PRODUCER_MAP, self.PARENT_JOBNETS)
        assert len(result) == 1
        assert "/NS/01/monthly_proc/monthly_1st" in result

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
            exclude_jobnets={"/NS/01/daily_proc/daily_batch"})
        assert len(result) == 1
        assert "/NS/02/audit_daily" in result

    def test_exclude_all(self):
        """全親ジョブネットを除外 → 空dict"""
        result = find_producers_across_jobnets(
            "/data/FILE_X", self.PRODUCER_MAP, self.PARENT_JOBNETS,
            exclude_jobnets={"/NS/01/daily_proc/daily_batch", "/NS/02/audit_daily"})
        assert result == {}
