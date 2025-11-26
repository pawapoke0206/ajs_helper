# AST - ビルド手順 & Defender誤検知対策

## 通常のビルド手順

`build.bat` をダブルクリックするだけです。
成果物は `dist\AST\AST.exe` に出力されます。


## exeがDefenderに消される場合

### 原因

PyInstallerのプリコンパイル済みbootloader（exe内に埋め込まれる起動用バイナリ）が
世界中で共有されているため、同じbootloaderを使うマルウェアとバイナリのハッシュが
一致してDefenderに引っかかります。ツール名やコードの内容は無関係です。

### 対策（3層）

build.bat と build.spec には以下の3層の対策が組み込まれています。

| 層 | 対策 | 効果 |
|----|------|------|
| 1 | UPX圧縮を無効化 | UPX圧縮がマルウェアのパッキングと同じ技術なので、無効にするだけで誤検知が大幅に減る |
| 2 | bootloaderを自前ビルド | ビルド環境固有のバイナリになるため、既知シグネチャに一致しなくなる |
| 3 | onedirモード | onefileの一時展開挙動がマルウェア的で怪しまれるが、onedirにはその問題がない |


### 層2（bootloader自前ビルド）が失敗する場合

C/C++コンパイラが必要です。以下の手順でインストールしてください。

1. https://visualstudio.microsoft.com/ja/visual-cpp-build-tools/ からダウンロード
2. インストーラーで「C++ によるデスクトップ開発」にチェックを入れてインストール
3. PCを再起動
4. 再度 `build.bat` を実行

正常にbootloaderがビルドされると、build.batの出力に
`Building bootloader for ...` のような表示が出ます。


### それでも誤検知される場合（根本解決: コード署名）

デジタル証明書でexeに署名すると、Defenderが信頼済みとして扱います。

```
signtool sign /f 証明書.pfx /p パスワード /tr http://timestamp.digicert.com /td sha256 /fd sha256 dist\AST\AST.exe
```

正規の証明書は年間数万円かかるため、個人ツールでは上記3層の対策で十分なケースが多いです。
