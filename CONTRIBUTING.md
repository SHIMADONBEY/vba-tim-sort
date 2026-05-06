# Contributing

English | [Japanese](CONTRIBUTING.ja.md)

This document explains how to run and verify tests locally, how to make source changes, and what to check before opening a pull request.

## Purpose
Provide clear instructions for making changes, running local verification, and required checks before pushing or creating a PR.

## Required preconditions
- Enable VBA macros and enable “Trust access to the VBA project object model” in Excel Trust Center.
- Close all running Excel instances before running tests.
- Run commands from the repository root (the working directory affects test output paths).

## VBA coding rules

The following rules apply to all VBA source files under `vba-files/`.
Production files (everything except `vba-files/test/`) are subject to stricter checks enforced by the **vba-lint** CI job.

### Production code (`vba-files/` root — strict rules)

| Rule | Detail |
|---|---|
| `Option Explicit` required | Every module must declare `Option Explicit` at the top. |
| No `On Error Resume Next` | Use structured error handling (`On Error GoTo label`) instead. |
| No `Stop` | Remove all `Stop` statements before merging. |
| No `Debug.Print` | Remove all debug output before merging. |
| No trailing whitespace | Lines must not end with spaces or tabs. |
| File ends with a newline | Each file must end with a single newline character (LF). |

### Test code (`vba-files/test/` — relaxed rules)

Test modules are only checked for basic text hygiene (trailing whitespace, final newline).
`Debug.Print`, `Stop`, and `On Error Resume Next` are permitted in test code.

### What the CI checks automatically

The **vba-lint** workflow (`.github/workflows/vba-lint.yml`) runs on every pull request that touches `vba-files/`.
It runs `.github/scripts/lint-vba.sh` on an Ubuntu runner and produces file-and-line annotations directly in the PR.

The following items are **not** checked by CI and require local Windows/Excel verification (see the `local-verified` label requirement):
- Runtime correctness and sorting behaviour
- Excel API compatibility
- Test suite results (`test-result.json`)

## About making changes

### Changing VBA
- If you edit code under `vba-files`, always run local tests before pushing.
- When writing test output files, keep temporary filenames unique (for example, combine random number + Timer) to avoid collisions and partial-read races.

### Changing PowerShell scripts
- PowerShell trusts the JSON produced by VBA as authoritative on success. Do not overwrite VBA-authored output for successful runs.
- `Write-ResultAndExit` is intended only as a fallback to write results when VBA did not produce a stable result (see `.github/scripts/run-tests-core.ps1`).

### Before creating a Pull Request
- Fetch and merge/rebase the latest develop and verify behaviour locally. Ensure the run passes using the Local testing instructions.
- Attach the following to the Pull Request:
  - Test result: `test-result.{timestamp}.{runId}.json`
  - Test logs: `test-result.log or test-result.write-debug.log`
  - SHA256 hash of the `test-runner.xlsb` used for verification (see command examples below)

NOTE: Do not include confidential information in test results, logs, or test code.

### About `test-runner.xlsb`

- This repository manages `test-runner.xlsb` with Git LFS. 
- Before working locally, run `git lfs install` to enable LFS (one-time per environment), then run `git lfs pull` to download LFS objects. Only commit changes to `.gitattributes` if you are updating LFS tracking patterns.
```bash
git lfs install        # one-time per environment
git fetch --all        # fetch LFS pointers
git lfs pull           # download LFS objects for the current checkout
```

- Direct updates to the binary require approval from the repository owner (@Shimadonbey). If an update is necessary, include the update reason and local verification logs in the PR.
- Official builds are also distributed via GitHub Releases.

#### Hash check examples

- PowerShell:
``` powershell
Get-FileHash .\test-runner.xlsb -Algorithm SHA256 | Select-Object -ExpandProperty Hash
```

- Linux / macOS:
``` bash
sha256sum test-runner.xlsb | cut -d' ' -f1
```

#### Git LFS setup example

``` bash
git lfs install
git lfs track "test-runner.xlsb"
git add .gitattributes
git add test-runner.xlsb
git commit -m "Track test-runner.xlsb with Git LFS"
git push
```

### Pull Request pre-checklist

- [ ] Have you documented the purpose and related Issues?
- [ ] Have you incorporated the latest develop changes and resolved conflicts?
- [ ] Have you attached local test results (logs and JSON)?
- [ ] If you updated the binary, do you have maintainer approval?

### Test & review flow
- Follow "Local testing" below to run tests locally.
- A test is considered successful if the generated `test-result.{timestamp}.{runId}.json` contains `"passed": true`.
- You can also validate manually by running the `TestTimSort.RunAll` macro in `test-runner.xlsb` via the VBA editor.

### Commits
- Make small, focused commits so the scope of changes is clear.
- Prefix commit messages with the related issue number when applicable.

## Local testing

### Running tests
- Recommended (wrapper): run the wrapper script which spawns a child process and isolates COM:

``` powershell
.\.github\scripts\run-tests.ps1 -Export
```

After execution, a per-run file `test-result.{timestamp}.{runId}.json` will be created in the current directory. Check logs: `test-result.write-debug.log` and `test-result.log`.

### Test outputs
- The test-runner (`test-runner.xlsb`) should produce `test-result.{timestamp}.{runId}.json`. That JSON is the authoritative result on success.
- If a write error occurs, an error dump named `test-result.{timestamp}.{runId}.json.writerror.json` may be produced.

### Exit codes
| Exit code | Meaning |
|---:|---|
| `0` | Tests passed (passed = true) |
| `2` | Tests ran but failed (passed = false) |
| `3` | Macro execution error or fatal write error |
| `4` | Importer-related error (missing importer / import failure) |
| `5` | Workbook not found (`test-runner.xlsb`) |
| `7` | Result file missing or unstable (could not obtain stable output) |

## Notes and troubleshooting
- If you see truncated or invalid JSON, ensure the VBA writer uses a unique temporary filename and atomically moves the file into place.
- If PowerShell appears to overwrite VBA output, confirm you are using the latest `.github/scripts/run-tests-core.ps1`, which treats VBA output as authoritative on success.
- If importer issues occur, check `.github/scripts/import-vba-into-workbook.ps1` and run the wrapper with `-Export` to capture importer logs.

## Pull request checklist
- I ran `.\.github\scripts\run-tests.ps1 -Export` locally and verified results.
- I tested any `vba-files` changes and included reproduction steps.
- I updated documentation (CONTRIBUTING / README) as needed.
- I included instructions to reproduce the test run in the PR description.
