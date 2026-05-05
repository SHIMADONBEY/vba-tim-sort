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

エラーが発生した場合は、`test-result.{timestamp}.{runId}.json.writeerror.json` が出力されることもあります。

#### テスト結果コード
|終了コード|内容|
|-----|---|
|`0`  |テスト成功|
|`2`  |テストに成功したが、テスト不合格|
|`3`  |`test-runner.xlsb` マクロの実行エラー|
|`4`  |`test-runner.xlsb` 取り込みエラー|
|`5`  |`test-runner.xlsb` ワークブックなし|
|`7`  |`test-result.{timestamp}.{runid}.json` が未生成、または JSON 不正・切り詰め等により結果を正しく読み取れない|

