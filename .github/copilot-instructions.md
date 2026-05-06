# Copilot instructions (vba-tim-sort)

These instructions apply to Copilot code suggestions and Copilot code review for this repository.

## Primary goal
Keep this VBA library **portable and embeddable**, including **Excel for Mac** compatibility.

## Review focus (highest priority)
1. Enforce the repository VBA rules (see below).
2. Flag portability risks (Windows-only features) and propose pure-VBA alternatives.
3. When behavior changes are possible, request evidence (tests / local run notes).

## Repository VBA rules

### Scope distinction
- **Production code**: `vba-files/` excluding `vba-files/test/` (strict)
- **Test code**: `vba-files/test/` (relaxed)

### Production code rules (strict)
- `Option Explicit` is required at the top of every module.
- Do **not** use `On Error Resume Next`. Use structured error handling (`On Error GoTo ...`) and do not silently ignore errors.
- Do **not** use `Stop` or `Debug.Print`.
- No trailing whitespace. Files must end with a newline (LF).

### Test code rules (relaxed)
- Keep basic text hygiene: no trailing whitespace; file ends with a newline (LF).
- `Debug.Print`, `Stop`, and `On Error Resume Next` may be acceptable in test modules.
  Do not report them as production violations when changes are limited to `vba-files/test/`.

## Portability rule (Excel for Mac compatibility)
- **Windows API calls are prohibited.**
  - Prohibited patterns include: `Declare` / `Declare PtrSafe`, any `Lib "..."` declarations
    (e.g. `kernel32`, `user32`, `advapi32`), and Windows-only API wrappers.
  - Prefer pure VBA and Excel built-in features instead.

## Output style for reviews
- Be specific: mention the **file path** and the **exact construct** (e.g. `On Error Resume Next`, `Declare PtrSafe ... Lib "kernel32"`).
- Prefer small, safe fixes. Avoid large rewrites unless necessary.
- If sorting correctness/stability could be affected (edge cases: empty input, duplicates, already-sorted, reverse-sorted, very large inputs),
  request additional test coverage or local run notes.
