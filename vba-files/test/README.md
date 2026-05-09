This folder stores test-only VBA modules for validating VbaTimSort.

Important:
- Keep these modules out of release workbooks and package outputs intended for end users.
- Use this folder only in verification or benchmark workbooks.

Recommended usage:
- Production modules: vba-files root
- Test modules: vba-files/test
- Before release, verify that test modules are not included in the distributed workbook

---

このフォルダには、VbaTimSortの検証用にのみ使用されるVBAモジュールが格納されています。

重要：
- エンドユーザー向けのリリース用ワークブックやパッケージ出力には、これらのモジュールを含めないでください。
- このフォルダーは、検証用またはベンチマーク用のワークブックでのみ使用してください。

推奨される使用方法：
- 本番用モジュール：vba-files ルート
- テスト用モジュール：vba-files/test
- リリース前に、配布するワークブックにテスト用モジュールが含まれていないことを確認してください
