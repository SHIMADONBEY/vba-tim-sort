#!/usr/bin/env bash
# .github/scripts/lint-vba.sh
#
# Lint VBA source files under vba-files/.
#
# Production files (vba-files/**/*.bas, *.cls, *.frm — excluding vba-files/test/**):
#   - Option Explicit must be present
#   - On Error Resume Next is disallowed
#   - Stop is disallowed
#   - Debug.Print is disallowed
#   - Windows API calls (Declare/Lib) are disallowed (for Excel for Mac compatibility)
#   - No trailing whitespace
#   - File must end with a newline
#
# Test files (vba-files/test/**/*.bas, *.cls, *.frm):
#   - No trailing whitespace
#   - File must end with a newline
#
# Exit code: 0 = all checks passed, 1 = one or more violations found.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
VBA_DIR="${REPO_ROOT}/vba-files"

errors=0

# ── helpers ────────────────────────────────────────────────────────────────────

err() {
  echo "::error file=${1},line=${2}::${3}" >&2
  errors=$((errors + 1))
}

# Return 0 (true) if the line is a VBA comment-only line.
# Excludes:
#   - lines starting with "'" (apostrophe comment)
#   - lines starting with "Rem " / "REM " etc.
is_comment_line() {
  local line="$1"
  # Normalize CRLF -> LF and then test
  line="${line%$'\r'}"

  if [[ "$line" =~ ^[[:space:]]*\' ]]; then
    return 0
  fi
  # Case-insensitive "Rem"
  shopt -s nocasematch
  if [[ "$line" =~ ^[[:space:]]*rem([[:space:]]+|$) ]]; then
    shopt -u nocasematch
    return 0
  fi
  shopt -u nocasematch
  return 1
}

# Check that a file ends with a newline character.
check_final_newline() {
  local file="$1"
  if [[ -s "$file" ]] && [[ "$(tail -c1 "$file" | wc -c)" -gt 0 ]]; then
    local last_byte
    last_byte=$(tail -c1 "$file" | od -An -tx1 | tr -d ' ')
    if [[ "$last_byte" != "0a" ]]; then
      err "$file" "EOF" "File does not end with a newline."
    fi
  fi
}

# Check for trailing whitespace (spaces or tabs before line ending).
check_trailing_whitespace() {
  local file="$1"
  local lineno=0
  while IFS= read -r line; do
    lineno=$((lineno + 1))
    # Strip CR if present (CRLF files on Linux)
    line="${line%$'\r'}"

    if [[ "$line" =~ [[:blank:]]$ ]]; then
      err "$file" "$lineno" "Trailing whitespace detected."
    fi
  done < "$file"
}

# ── production checks ──────────────────────────────────────────────────────────

check_production_file() {
  local file="$1"

  # 1. Option Explicit must appear in the file
  if ! grep -qiE '^\s*Option\s+Explicit\s*$' "$file"; then
    err "$file" "1" "'Option Explicit' is missing. All production modules must declare 'Option Explicit'."
  fi

  # 2. Prohibited constructs
  local lineno=0
  while IFS= read -r line; do
    lineno=$((lineno + 1))
    local stripped="${line%$'\r'}"

    # Skip comment-only lines
    if is_comment_line "$stripped"; then
      continue
    fi

    # Case-insensitive matching for VBA keywords
    shopt -s nocasematch

    if [[ "$stripped" =~ ^[[:space:]]*On[[:space:]]+Error[[:space:]]+Resume[[:space:]]+Next ]]; then
      shopt -u nocasematch
      err "$file" "$lineno" "Prohibited construct: 'On Error Resume Next' is not allowed in production code."
      continue
    fi

    if [[ "$stripped" =~ ^[[:space:]]*Stop[[:space:]]*$ ]]; then
      shopt -u nocasematch
      err "$file" "$lineno" "Prohibited construct: 'Stop' is not allowed in production code."
      continue
    fi

    if [[ "$stripped" =~ ^[[:space:]]*Debug\.Print([^[:alnum:]_]|$) ]]; then
      shopt -u nocasematch
      err "$file" "$lineno" "Prohibited construct: 'Debug.Print' is not allowed in production code."
      continue
    fi

    # Windows API calls (Declare/Lib) are prohibited for Excel for Mac compatibility
    if [[ "$stripped" =~ ^[[:space:]]*(Public[[:space:]]+|Private[[:space:]]+)?Declare([[:space:]]+PtrSafe)?([^[:alnum:]_]|$) ]]; then
      shopt -u nocasematch
      err "$file" "$lineno" "Prohibited construct: Windows API declaration ('Declare' / 'Declare PtrSafe') is not allowed in production code (Excel for Mac compatibility)."
      continue
    fi

    if [[ "$stripped" =~ (^|[^[:alnum:]_])Lib[[:space:]]+\"[^\"]+\" ]]; then
      shopt -u nocasematch
      err "$file" "$lineno" "Prohibited construct: external library binding via 'Lib \"...\"' is not allowed in production code (Excel for Mac compatibility)."
      continue
    fi

    shopt -u nocasematch

    # Trailing whitespace
    if [[ "$stripped" =~ [[:blank:]]$ ]]; then
      err "$file" "$lineno" "Trailing whitespace detected."
    fi
  done < "$file"

  check_final_newline "$file"
}

# ── test checks ────────────────────────────────────────────────────────────────

check_test_file() {
  local file="$1"
  check_trailing_whitespace "$file"
  check_final_newline "$file"
}

# ── file discovery ─────────────────────────────────────────────────────────────

found_any=0

while IFS= read -r -d '' file; do
  found_any=1
  # Normalise path separator for display
  rel="${file#"${REPO_ROOT}/"}"

  # Determine if the file lives under vba-files/test/
  if [[ "$file" == "${VBA_DIR}/test/"* ]]; then
    echo "  [test]       $rel"
    check_test_file "$file"
  else
    echo "  [production] $rel"
    check_production_file "$file"
  fi
done < <(find "${VBA_DIR}" -type f \( -iname '*.bas' -o -iname '*.cls' -o -iname '*.frm' \) -print0 | sort -z)

if [[ "$found_any" -eq 0 ]]; then
  echo "Warning: no .bas/.cls/.frm files found under ${VBA_DIR}" >&2
fi

# ── summary ────────────────────────────────────────────────────────────────────

echo ""
if [[ "$errors" -gt 0 ]]; then
  echo "::error::VBA lint failed with ${errors} error(s). See annotations above for details."
  exit 1
else
  echo "VBA lint passed — no issues found."
fi
