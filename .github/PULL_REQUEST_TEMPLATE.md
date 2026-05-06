## Summary
(Short description of what this PR changes)

## Changes
- Bullet points of changes

## Impact
- Note any build/runtime/compatibility impacts (if any)

## Checklist (required before merge)

### Automated CI checks (handled by GitHub Actions)
- [ ] **vba-lint** passed — `Option Explicit` present, no prohibited constructs (`On Error Resume Next` / `Stop` / `Debug.Print` / Windows API calls via `Declare`/`Lib`), no trailing whitespace, files end with newline
  - Applies strict rules to `vba-files/` (production) and relaxed rules to `vba-files/test/` (test)
- [ ] **Require Local Verification** check passed (requires the `local-verified` label)

### Local Windows/Excel verification (manual — cannot be automated in CI)
- [ ] Performed local Excel verification on Windows
- [ ] Added the `local-verified` label to this PR
- [ ] If applicable, attached `test-result.json` or added test results in a comment

## What the CI checks cover

| Check | Scope | Detail |
|---|---|---|
| `Option Explicit` | Production only | Every `.bas`/`.cls`/`.frm` in `vba-files/` (not `test/`) must declare `Option Explicit`. |
| No `On Error Resume Next` | Production only | Use structured error handling instead. |
| No `Stop` | Production only | Remove all `Stop` statements before merging. |
| No `Debug.Print` | Production only | Remove all debug output before merging. |
| Trailing whitespace | Production + Test | Lines must not end with spaces or tabs. |
| Final newline | Production + Test | Each file must end with a newline character. |

> **Note:** Runtime correctness and Excel API compatibility cannot be verified by CI. Local Windows/Excel verification is still required and enforced by the `local-verified` label.

## Local verification (quick guide)
1. Checkout repository on a Windows machine.
2. Export `vba-files/` into `test-runner.xlsb` (example): `npx @localsmart/xvba-cli export --src vba-files --target test-runner.xlsb` (Confirm actual xvba-cli subcommand with `npx xvba --help`.)
3. Run the test script (example): `powershell -ExecutionPolicy Bypass -File .github/scripts/run-tests.ps1`
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

### 自動CIチェック（GitHub Actionsが実行）
- [ ] **vba-lint** が通っていること — `Option Explicit` の存在、禁止構文（`On Error Resume Next` / `Stop` / `Debug.Print` / `Declare`/`Lib` によるWindows APIの使用）なし、行末スペースなし、ファイル末尾改行あり
  - `vba-files/`（製品コード）には厳格ルール、`vba-files/test/`（テストコード）には緩和ルールを適用
- [ ] **Require Local Verification** チェックが通っていること（`local-verified` ラベルが必要）

### ローカルの Windows/Excel 検証（手動 — CIでは自動化不可）
- [ ] Windows 上でローカルの Excel 検証を実施したこと
- [ ] `local-verified` ラベルを付与済みであること
- [ ] 必要なら `test-result.json` を PR に添付、または実行結果をコメントに記載

## CIチェックの対象

| チェック内容 | 対象 | 詳細 |
|---|---|---|
| `Option Explicit` | 製品コードのみ | `vba-files/`（`test/` 除く）の `.bas`/`.cls`/`.frm` に必須。 |
| `On Error Resume Next` 禁止 | 製品コードのみ | 構造化エラーハンドリングを使用すること。 |
| `Stop` 禁止 | 製品コードのみ | マージ前に削除すること。 |
| `Debug.Print` 禁止 | 製品コードのみ | マージ前に削除すること。 |
| 行末スペース禁止 | 製品コード＋テスト | 行末にスペースやタブを残さないこと。 |
| ファイル末尾の改行 | 製品コード＋テスト | 各ファイルは改行で終わること。 |

> **注意:** 実行時の正確性と Excel API との互換性は CI では検証できません。ローカルの Windows/Excel 検証は引き続き必須であり、`local-verified` ラベルで管理されます。

## ローカル検証手順（概要）
1. Windows 環境でリポジトリをチェックアウト
2. vba-files を test-runner.xlsb に反映
（例）: `npx @localsmart/xvba-cli export --src vba-files --target test-runner.xlsb`（`xvba-cli` のコマンドは環境に合わせて確認してください）
3. テストを実行
（例）: `powershell -ExecutionPolicy Bypass -File .github/scripts/run-tests.ps1`
    - `test-result.json` を出力する想定
    - Exit code `0` が合格

## ラベル付与方法
- ローカル検証が合格したら PR 作成者またはレビュワーが `local-verified` ラベルを付与してください。
- `gh` CLI の例: `gh pr edit <PR番号> --add-label local-verified`

## 備考
- `local-verified` ラベルがない PR は必須ステータスチェックによりマージ不可です。レビュワーは必ずラベルを確認してください。
