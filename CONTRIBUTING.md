# Contributing

English | [Japanese](CONTRIBUTING.ja.md)

This document explains how to run and verify tests locally, how to make source changes, and what to check before opening a pull request.

## Purpose
Provide clear instructions for making changes, running local verification, and required checks before pushing or creating a PR.

## Required preconditions
- Enable VBA macros and enable “Trust access to the VBA project object model” in Excel Trust Center.
- Close all running Excel instances before running tests.
- Run commands from the repository root (the working directory affects test output paths).

## About making changes

### Changing VBA
- If you edit code under `vba-files`, always run local tests before pushing.
- When writing test output files, keep temporary filenames unique (for example, combine random number + Timer) to avoid collisions and partial-read races.

### Changing PowerShell scripts
- PowerShell trusts the JSON produced by VBA as authoritative on success. Do not overwrite VBA-authored output for successful runs.
- `Write-ResultAndExit` is intended only as a fallback to write results when VBA did not produce a stable result (see `.github/scripts/run-tests-core.ps1`).

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
