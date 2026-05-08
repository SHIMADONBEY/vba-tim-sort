# vba-tim-sort

VBA で Tim Sort アルゴリズムを実装したライブラリです。
Excelの並べ替え機能を使わずに、マクロの中でライブラリに依存せず高速な並べ替えを行うことができます。
また、オブジェクトを内包したコレクションや配列について効果的に並べ替えすることができます。

[English](README.md) | 日本語

## 導入方法

1. ファイルのインポート
   - VBA エディタを開き、メニューの「ファイル」→「ファイルのインポート」で次のファイルをプロジェクトに追加します。  
     - [vba-files/VbaTimSort.bas](vba-files/VbaTimSort.bas#L1)  
     - [vba-files/IComparator.cls](vba-files/IComparator.cls#L1)  
   - `IComparator` の実装サンプルは `vba-files/test/comparators/` にあります。GitHub リポジトリ上では参照できますが、`vba-files/test/**` は `export-ignore` のため `git archive` ベースのリリース ZIP には含まれません。必要な場合はリポジトリから取得してプロジェクトに追加してください。

2. 使い方 （配列）

``` vb
Dim arr As Variant
arr = Array(5, 2, 9, 1)

' 自然順（数値・文字列・日付）でソートする場合は comparator に Nothing を渡す
Dim sorted As Variant
sorted = SortArrayInPlace(arr, Nothing)             ' 昇順

' 降順にする場合
Dim sortedDesc As Variant
sortedDesc = SortArrayInPlace(arr, Nothing, True)   ' 降順
```

3. 使い方 （Collection）

``` vb
Dim coll As New Collection
coll.Add "b"
coll.Add "a"
Dim sortedColl As Collection
Set sortedColl = SortCollection(coll, Nothing)
```

4. IComparator を実装して使う例

``` vb
' MyComparator.cls
Implements IComparator

Private Function IComparator_Compare(ByVal a As Variant, ByVal b As Variant) As Integer
    ' カスタム比較ロジック（例：数値比較）
    If a < b Then
        IComparator_Compare = -1
    ElseIf a > b Then
        IComparator_Compare = 1
    Else
        IComparator_Compare = 0
    End If
End Function
```

``` vb
' example.bas
Dim comp As IComparator
Set comp = New MyComparator
Dim sortedWithComp As Variant
sortedWithComp = SortArrayInPlace(arr, comp)
```

5. 補足

- `SortArrayInPlace()` / `SortCollection()` は元の並びを変えず、新しい配列／Collection を返します。
- オブジェクトを比較する場合は、`IComparator` を必ず提供してください。（渡さないとエラーになります。）

## OSS プロジェクトガイドライン

このプロジェクトはオープンソースソフトウェア (OSS) として運用されています。
以下は、利用者とコントリビューターに向けた一般的な方針です。

### ライセンス

本プロジェクトのライセンス条件は [LICENSE](LICENSE) に記載されています。
利用、改変、再配布を行う場合は、同ライセンスに従ってください。

### 開発ソースについて

#### ブランチ構成

本プロジェクトは以下のブランチ構成をとっています。

|ブランチ名|内容|
|---|---|
|`main`|リリース済みの安定ブランチ。ここに直接コミットしないこと。リリースは `main` へマージ後にタグ付けします。  |
|`develop`|次回リリースに向けた開発ブランチ。日常の統合先。|
|`release/vX.Y.Z`|リリース作業ブランチ。`develop`から作成し、Pull-Requestを通して、`main`へマージ（Squash and Merge Commit）し、マージ完了後はブランチ削除と`develop`との同期を行います。|
|`archive`|保管用ブランチ（古い履歴など）。|

#### テストランナー (`test-runner.xlsb`) の取り扱い

- 本リポジトリではテスト実行用バイナリ `test-runner.xlsb` を Git LFS で管理しています。ローカル作業前に `git lfs install` で LFS を有効化し、`git lfs pull` で LFS オブジェクトをダウンロードしてください。`.gitattributes` へのコミットが必要なのは、追跡ルールを変更する場合のみです。未設定の場合は `CONTRIBUTING.ja.md` を参照してください。
- PR 作成前に必ず最新の `develop` を取り込み、ローカルで検証を実行してください。
- PR には以下を添付してください：
  - ローカル実行ログ（テキスト） — 実行した手順と結果（成功/失敗・エラー出力）  
  - 使用した `test-runner.xlsb` の SHA256 ハッシュ（下記コマンド参照）
- `test-runner.xlsb` の直接更新はリポジトリ所有者([@Shimadonbey](https://github.com/SHIMADONBEY))の承認が必要です。バイナリ更新が必要な場合は PR でその理由を明記してください。
- 詳しいローカル検証手順や LFS の使い方は [CONTRIBUTING](CONTRIBUTING.ja.md) を参照してください。

ハッシュ確認コマンド（例）
- PowerShell (Windows):
```powershell
Get-FileHash test-runner.xlsb -Algorithm SHA256 | Select-Object -ExpandProperty Hash
```

- Linux / macOS:
``` bash
sha256sum test-runner.xlsb | cut -d' ' -f1
```

#### 運用ルールなどについて

##### ブランチ保護

`main` と `develop` は force-push 禁止とします。

##### Pull-Request 等の対応について

このリポジトリはCI実行は難しいため、ローカル検証必須とする運用としています。
詳しい手順は[CONTRIBUTING](CONTRIBUTING.ja.md) をご参照ください。

`test-runner.xlsb` の直接更新はリポジトリ所有者の承認が必要です。
バイナリの更新が必要な場合は、Pull-Request に更新理由とローカル検証ログおよび、評価で用いた`test-runner.xlsb`のSHA256ハッシュを添えて、承認を得てください。公式配布はGitHub Releases にも配置します。

添付する SHA は「実際に検証で使用した `test-runner.xlsb` から算出した値」を必ず記載してください。
添付ログに機密情報が含まれていないことを確認してから公開してください。

##### マージ方針

Issue 対応等によるブランチマージは原則、Squash マージとします。
`archive` ブランチから`main`、`develop` ブランチのマージは禁止としています。
（リポジトリもできないように設定しています。）

##### バージョンについて

SemVer 推奨（`v1.2.3` タグ）とします。

##### ブランチの取り扱いについて

長期間のブランチは避け、早めに統合してコンフリクトを最小化します。

#### 配布・リリースガイド

このプロジェクトはリリースアーティファクトがプロダクションコードと必要なライセンス/通知ファイルのみを含むよう、厳密なリリース手順を推奨しています。
開発専用フレームワークコード（`archive` ブランチ内）はリリースに含めないでください。

リリースに含めるもの
- `VbaTimSort.bas`
- `IComparator.cls`
- README
- LICENSE
- `THIRD_PARTY_NOTICES.md`

含めてはいけないもの
- `vba-files/test/` 配下のテストコード — テスト専用モジュールは開発・検証用であり、エンドユーザー向けの配布物には含めてはいけません。
- `xvba_modules/` や `archive` ブランチにのみ存在する開発用コード（例：`XDebug` 等）


##### リリース手順

1. 最新の`develop` ブランチをPull します。

``` bash
git fetch origin
git checkout develop
git pull origin develop
# HEADコミットを確認
git rev-parse --short HEAD
```

2. リリース用ブランチを作成します。

``` bash
git checkout -b release/v1.0.0
```

3. 必要に応じてCHANGELOG等の情報を更新し、Commitします。

4. 更新確認後、Push します。

``` bash
git push -u origin release/v1.0.0
```

5. ローカルでアーカイブを作成します。 (`.gitattributes` のexport-ignore を利用します。)

``` bash
git archive --format=zip --worktree-attributes --output=release-test.zip HEAD
```

6. 内容物を確認します。

``` bash
unzip -l release-test.zip

# 内容物にxvba_module が含まれていないか確認してください。

unzip -l release-test.zip | grep -E '^ *[0-9]+' | awk '{print $4}' | grep -E '^xvba_modules/' || echo "OK: xvba_modules not found"
```

7. リリースアーカイブとチェックサムを作成します。

``` bash
git archive --format=zip --worktree-attributes --output=vba-tim-sort-v1.0.0.zip HEAD
sha256sum vba-tim-sort-v1.0.0.zip > vba-tim-sort-v1.0.0.zip.sha256
```

8. リリースアーカイブを公開します。`release/v1.0.0` ブランチから、`main`ブランチへマージするPull-Request を発行します。レビュー完了後、バージョン番号のタグをつけ、Squashマージします。

9. マージ後、`main` ブランチの内容を`develop` ブランチにマージします。

``` bash
git checkout develop
git pull origin develop
git merge --no-ff origin/main
git push origin develop
```

#### 補足事項

このライブラリではXVBAという開発フレームワークを利用して開発を行っています。
XVBAのフレームワークコードは`archive` ブランチで保管しています。
XVBAのフレームワークコードはリリース物には含まれませんのでご承知おきください。

### コントリビュート

コントリビューションを歓迎します。
一般的な流れは以下のとおりです。

1. リポジトリをフォークする。
2. 機能追加または修正用のブランチを作成する。
3. 振る舞いを変更した場合は、テストを追加または更新する。
4. 変更内容を明確にしたプルリクエストを作成する。

変更は小さく焦点を絞り、レビューしやすい形を意識してください。

### Issue と要望

不具合報告や機能要望は Issue で受け付けます。
可能であれば、次の情報を含めてください。

- 期待される動作
- 実際の動作
- 再現手順
- VBA / Excel のバージョン情報

### セキュリティ

[セキュリティポリシー（日本語）](SECURITY.ja.md) をご参照ください。

### サポート範囲

本プロジェクトはベストエフォートでメンテナンスされています。
返信速度やリリース時期は保証されません。

### 謝辞

Issue 報告、ドキュメント改善、コード提供に協力してくださるすべての方に感謝します。

## CHANGELOG

|Date|Version|Description|
|--:|---|---|
|2026/05/06|`v1.0.0`|初版リリース.|
