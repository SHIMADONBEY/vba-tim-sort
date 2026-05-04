## Summary
(Short description of what this PR changes)

## Changes
- Bullet points of changes

## Impact
- Note any build/runtime/compatibility impacts (if any)

## Checklist (required before merge)
- [ ] Automatic CI (lint / unit tests) passed
- [ ] Performed local Excel verification on Windows
- [ ] Added the `local-verified` label to this PR
- [ ] If applicable, attached `test-result.json` or added test results in a comment

## Local verification (quick guide)
1. Checkout repository on a Windows machine.
2. Install dependencies: npm ci
3. Export `vba-files/` into `test-runner.xlsb` (example): `npx xvba export --src vba-files --target test-runner.xlsb` (Confirm actual xvba-cli subcommand with `npx xvba --help`.)
4. Run the test script (example): `powershell -ExecutionPolicy Bypass -File .github/scripts/run-tests.ps1`
- `test-result.json` is expected as output.
- Exit code `0` indicates pass.

## Labeling (how-to)
- After successful local verification, the PR author or reviewer should add the `local-verified` label.
- Example with `gh` CLI: `gh pr edit <PR_NUMBER> --add-label local-verified`


## Notes
- PRs without the `local-verified` label will fail the repository's required status check and cannot be merged.
- Reviewers must verify the presence of the label before merging.

---

## 概要（日本語）
（このPRで何を変更するか、短く記載）

## 変更点
- 追加/変更点を箇条書き

## 影響範囲
- ビルドや実行時、互換性への影響（あれば記載）

## マージ前チェック（必須）
- [ ] 自動CI（lint / unit tests）が通っていること
- [ ] Windows 上でローカルの Excel 検証を実施したこと
- [ ] `local-verified` ラベルを付与済みであること
- [ ] 必要なら `test-result.json` を PR に添付、または実行結果をコメントに記載

## ローカル検証手順（概要）
1. Windows 環境でリポジトリをチェックアウト
2. 依存をインストール: `npm ci`
3. vba-files を test-runner.xlsb に反映
（例）: `npx xvba export --src vba-files --target test-runner.xlsb`（`xvba-cli` のコマンドは環境に合わせて確認してください）
4. テストを実行
（例）: `powershell -ExecutionPolicy Bypass -File .github/scripts/run-tests.ps1`
- `test-result.json` を出力する想定
- Exit code `0` が合格

## ラベル付与方法
- ローカル検証が合格したら PR 作成者またはレビュワーが `local-verified` ラベルを付与してください。
- `gh` CLI の例: `gh pr edit <PR番号> --add-label local-verified`

## 備考
- `local-verified` ラベルがない PR は必須ステータスチェックによりマージ不可です。レビュワーは必ずラベルを確認してください。
