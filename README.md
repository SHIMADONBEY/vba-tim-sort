# vba-tim-sort

A library that implements Tim Sort's sorting algorithm in VBA.

English | [日本語](README.ja.md)

## How to Install

1. Importing Files
   - Open the VBA Editor, go to “File” → “Import File” in the menu, and add the following files to your project:  
     - [vba-files/VbaTimSort.bas](vba-files/VbaTimSort.bas#L1)  
     - [vba-files/IComparator.cls](vba-files/IComparator.cls#L1)  
   - Sample implementations of `IComparator` are located in `vba-files/test/comparators/`. Add them to your project as needed.

2. Usage (Arrays)

``` vb
Dim arr As Variant
arr = Array(5, 2, 9, 1)

' To sort in natural order (numbers, strings, dates), pass `Nothing` to the comparator.
Dim sorted As Variant
sorted = SortArrayInPlace(arr, Nothing)             ' Ascending

' To sort in descending order, pass True as the optional third argument.
Dim sortedDesc As Variant
sortedDesc = SortArrayInPlace(arr, Nothing, True)   ' Descending
```

3. Usage （Collection）

``` vb
Dim coll As New Collection
coll.Add "b"
coll.Add "a"
Dim sortedColl As Collection
Set sortedColl = SortCollection(coll, Nothing)
```

4. An example of implementing IComparator

``` vb
' MyComparator.cls
Implements IComparator

Private Function IComparator_Compare(ByVal a As Variant, ByVal b As Variant) As Integer
    ' Customized comparison
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

5. Notes

- `SortArrayInPlace()` / `SortCollection()` does not modify the original array or collection; instead, it returns a new array or collection.
- When comparing objects, you must provide an `IComparator`. (Failure to do so will result in an error.)

## Open Source Project Guidelines

This project is maintained as an open source software (OSS) library.
The following notes describe common expectations for users and contributors.

### License

This project is distributed under the terms described in [LICENSE](LICENSE).
By using, modifying, or redistributing this project, you agree to follow that license.

### Development sources

#### Branch structure

This project uses the following branches:

|Name|Description|
|---|---|
|`main`|Stable release branch. Do not commit directly. Releases are tagged after merging into `main`.|
|`develop`|Development integration branch for the next release.|
|`release/vX.Y.Z`|Release branch. Create it from `develop`, merge it into `main` via a pull request (using a “squash and merge” commit), and after the merge is complete, delete the branch and sync it with `develop`.|
|`archive`|Storage branch for legacy code and framework code.|

#### Test runner (`test-runner.xlsb`) handling

- The repository distributes a test runner binary `test-runner.xlsb`. To avoid repository bloat, this file is managed via Git LFS. Before working locally, install/configure Git LFS with `git lfs install`, and if the file is not present locally, fetch LFS-managed content with `git lfs pull`. Commit `.gitattributes` only when intentionally changing LFS tracking patterns. See [CONTRIBUTING](CONTRIBUTING.md) for details.
- Before opening a PR, always update from the latest `develop` and run local verification.
- Attach the following to your PR:
  - Local run logs (text) — the steps you ran and the results (success/failure and any error output)
  - The SHA256 hash of the `test-runner.xlsb` used for verification (see commands below)
- Direct updates to `test-runner.xlsb` require approval from the repository owner ([@Shimadonbey](https://github.com/SHIMADONBEY)). If an update is necessary, state the reason in the PR and attach local verification logs.
- Official builds may also be published as GitHub Releases.

Hash check examples
- PowerShell (Windows):
```powershell
Get-FileHash test-runner.xlsb -Algorithm SHA256 | Select-Object -ExpandProperty Hash
```
#### Operational rules

##### Branch protection

main and develop are protected from force-push.

##### Pull requests

Because CI execution is limited for this project, local verification is mandatory.
See [CONTRIBUTING](CONTRIBUTING.md) for full instructions.

When updating the test runner binary, include the update reason, local verification logs, and the SHA256 of the `test-runner.xlsb` used for verification in the PR, then obtain approval from the repository owner.
Always ensure the attached SHA is computed from the exact file you used to verify. Confirm that attached logs do not contain sensitive information before publishing.

##### Merge policy

Feature/issue branches should be merged by squash merge by default.
Merges from `archive` into `main` or `develop` are disallowed.

##### Versioning

Semantic Versioning (SemVer) is recommended; use tags like `v1.2.3`.

##### Branch lifespan

Avoid long-lived branches; merge early and often to minimize conflicts.

#### Distribution and Release Guidelines

This project follows a controlled release process to ensure release artifacts contain only production code and the required license/notice files, and to prevent inclusion of development-only framework sources (archived in the `archive` branch).

What to include in release artifacts
- vba-files/VbaTimSort.bas
- vba-files/IComparator.cls
- README (short usage/instructions)
- LICENSE
- THIRD_PARTY_NOTICES.md
- examples/ or sample files (if present)
Do NOT include:
- xvba_modules/ or any code that lives only on the `archive` branch (for example: XDebug or other development-only tools).
- test-only modules under vba-files/test

Why xvba_modules must not be included
- xvba_modules and related tools are development-environment specific and intentionally archived in the `archive` branch. They must not be redistributed or mixed into `develop`/`main` release artifacts.

Preparing a release (recommended flow)
1. Start from a verified `develop`:
   - git fetch origin
   - git checkout develop
   - git pull origin develop
   - Confirm HEAD commit: `git rev-parse --short HEAD`

2. Create a release branch based on develop (example):
   - git checkout -b release/v0.1.0
   - Update CHANGELOG or version notes if needed
   - Commit changes and push:
     - git push -u origin release/v0.1.0

3. Create the release archive locally to verify content (git archive respects .gitattributes export-ignore):
   - # create test archive from the current branch
     git archive --format=zip --worktree-attributes --output=release-test.zip HEAD
   - # inspect archive content and ensure no archive-only files are present
     unzip -l release-test.zip
     # verify xvba_modules not included:
     unzip -l release-test.zip | grep -E '^ *[0-9]+' | awk '{print $4}' | grep -E '^xvba_modules/' || echo "OK: xvba_modules not found"

4. Create final archive and compute checksum:
   - git archive --format=zip --worktree-attributes --output=vba-tim-sort-vX.Y.Z.zip HEAD
   - sha256sum vba-tim-sort-vX.Y.Z.zip > vba-tim-sort-vX.Y.Z.zip.sha256

5. Publish the release
   - From `release/v1.0.0`, open a Pull Request into `main`. After review, merge the PR according to team policy: either “Squash and merge” (recommended) or “Create a merge commit”. After merging, create a release tag (for example `v1.0.0`) and publish the release artifacts (the repository’s release workflow will publish artifacts on main pushes if configured).

Checklist before releasing
- [ ] The archive contains only intended production files (no xvba_modules or files from `archive` branch).
- [ ] LICENSE and THIRD_PARTY_NOTICES.md are included in the artifact.
- [ ] README/usage instructions are up-to-date.
- [ ] CHANGELOG/version information is updated.
- [ ] SHA256 checksum for the artifact is computed and attached to the release or PR.
- [ ] Local verification (if required) completed and `local-verified` label added to the release PR.

Automation notes
- The repository contains a release workflow (.github/workflows/release.yml) that runs on pushes to `main` and uses `git archive --worktree-attributes`. Ensure `.gitattributes` correctly marks archive-only files (e.g. `xvba_modules/** export-ignore`) so `git archive` excludes them.
- Consider adding a pre-merge CI check to detect accidental inclusion of `xvba_modules/` in PRs targeting `develop` or `main`. (This is optional and can be added later as an automated safeguard.)

Policies recap
- `archive` branch: holds development-only framework code (xvba_modules, XDebug, etc.). Do not merge or include into `develop` or `main`.
- Merge policy: use squash merges for feature/release branches; merges from `archive` into `main` or `develop` are disallowed.

#### Notes

This library is developed using the XVBA framework.
XVBA framework code is stored on the archive branch and is not included in release artifacts. 

### Contributing

Contributions are welcome.
Typical contribution flow:

1. Fork the repository.
2. Create a feature or fix branch.
3. Add or update tests when behavior changes.
4. Open a pull request with a clear summary.

Please keep changes focused, well-scoped, and easy to review.

### Issues and Requests

If you find a bug or want to request a feature, please open an issue.
When possible, include:

- Expected behavior
- Actual behavior
- Reproducible steps
- VBA / Excel version details

### Security

Please see [SECURITY](SECURITY.md)

### Support Scope

This project is maintained on a best-effort basis.
Response times and release schedules are not guaranteed.

### Acknowledgements

Thanks to everyone who reports issues, improves documentation, and contributes code.
