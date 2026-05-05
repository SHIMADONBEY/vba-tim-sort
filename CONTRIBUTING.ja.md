# ソースの更新について

[English](CONTRIBUTING.md) | 日本語

## 目的

ソースに修正手順並びに、ローカルでの評価方法とPush前に行ってほしいことを明記する。

## 必ずやってほしいこと。

- [ ] Excel の設定: Excel のトラストセンターの設定から、 *VBAマクロを有効にし* 、 *VBAオブジェクトモデルへのアクセスを信頼する* ようにしてください。
- [ ] *全ての Excel インスタンスを閉じて*から実行してください。
- [ ] ワーキングディレクトリ: リポジトリルートで作業してください。

## コードの修正について

### VBAの修正について

`vba-files` 内でソースを変更した場合は、必ずローカルでのテストを実行してください。

出力ファイルの一時名は衝突を避けるため乱数＋Timer 等のユニーク化を維持してください。

### Powershellスクリプトの修正

PS は「VBA が出力した JSON を信頼する」方針です。成功時は PS 側で上書きしないでください。

`Write-ResultAndExit` の役割は「フォールバックで結果を書く」ことに限定します（参照: .github/scripts/run-tests-core.ps1）。

### Pull-Request 作成前に対処してほしいこと

最新の `develop` ブランチをフェッチ・マージ／リベースし、最新コードを取り込んだうえで動作検証を行ってください。
[ローカルテスト](#ローカルテストの方法)で必ず合格した状態にしてください。

Pull-Request に以下を添付してください。

- テスト結果: `test-result.{timestamp}.{runId}.json`
- テストログ: `test-result.log` または、`test-result.write-debug.log`
- テスト評価で使った `test-runner.xlsb` のSHA256ハッシュ（下記コマンド参照）

NOTE: テスト結果、テストログ、テストコード類には*機密情報を載せない*でください。

#### `test-runner.xlsb` について

このリポジトリでは `test-runner.xlsb` をGit LFS で管理します。
ローカルで作業する前に `git lfs install` を実行し、LFS を有効にしてください。LFS 設定を行ったら `.gitattributes` の変更をコミットしてください。

ファイルの直接更新はリポジトリ所有者（@Shimadonbey）の承認が必要です。
更新が必要な場合は、PRに更新理由とローカル検証ログを添えてください。

公式ビルドはGitHub Releases にも配布します。

##### ハッシュ確認例

- Powershell
``` powershell
Get-FileHash .\test-runner.xlsb -Algorithm SHA256 | Select-Object -ExpandProperty Hash
```

- Linux / macOS
``` bash
sha256sum test-runner.xlsb | cut -d' ' -f1
```

##### Git LFS の導入

``` bash
git lfs install
git lfs track "test-runner.xlsb"
git add .gitattributes
git add test-runner.xlsb
git commit -m "Track test-runner.xlsb with Git LFS"
git push
```

#### Pull-Request 事前チェックリスト

- [ ] 目的・関連Issues に記載済みか？
- [ ] 最新の `develop` ブランチの内容を反映した状態にし、競合が起きないようにしているか？
- [ ] ローカルテスト結果: 実行ログ・JSON を添付しているか。
- [ ] バイナリ更新: 必要ならメンテナ承認済みか。

### テスト・レビュー手順

[ローカルテストの方法](#ローカルテストの方法) に従ってローカルテストを実行してください。
テスト結果 ( `test-result.{timestamp}.{runId}.json` ) の `"passed"` が `true` となっていればテスト成功です。

`test-runner.xlsb` のマクロ、`TestTimSort.RunAll` から手動検証を行うこともできます。

### コミットについて

コミットについては、修正した範囲が明確になるような粒度でコミットしてください。
コミットコメントの先頭には、対応したIssues の番号を付記してください。

## ローカルテストの方法

### コードの取り込み

以下のコマンドを実行してください。

``` powershell
.\.github\scripts\run-tests.ps1 -Export
```

テスト実行後、カレントディレクトリに `test-result.{timestamp}.{runId}.json` が生成されます。ログは `test-result.write-debug.log` と `test-result.log` を確認してください。

### テスト結果について

`test-result.{timestamp}.{runId}.json` に`test-runner.xlsb` のテスト結果が出力されます。

エラーが発生した場合は、`test-result.{timestamp}.{runId}.json.writerror.json` が出力されることもあります。

#### テスト結果コード
|終了コード|内容|
|-----|---|
|`0`  |テスト成功|
|`2`  |テストに成功したが、テスト不合格|
|`3`  |`test-runner.xlsb` マクロの実行エラー|
|`4`  |`test-runner.xlsb` 取り込みエラー|
|`5`  |`test-runner.xlsb` ワークブックなし|
|`7`  |`test-result.{timestamp}.{runid}.json` が未生成、または JSON 不正・切り詰め等により結果を正しく読み取れない|

