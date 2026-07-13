#!/bin/bash
# Verify that a published macOS OfficeCLI Mach-O only references Apple system
# libraries. This is a read-only release gate: it never rewrites install names.
# Exit codes: 2 invalid input, 3 inspection unavailable/failed, 4 non-portable.

set -euo pipefail

if [ "$#" -ne 1 ]; then
    echo "Usage: $0 <mach-o-binary>" >&2
    exit 2
fi

BINARY="$1"
FILE_TOOL="/usr/bin/file"
OTOOL="/usr/bin/otool"

if [ ! -f "$BINARY" ]; then
    echo "ERROR: macOS portability input does not exist: $BINARY" >&2
    exit 2
fi
if [ ! -x "$FILE_TOOL" ] || [ ! -x "$OTOOL" ]; then
    echo "ERROR: macOS portability inspection requires /usr/bin/file and /usr/bin/otool" >&2
    exit 3
fi

FILE_KIND="$($FILE_TOOL -b "$BINARY")"
if [[ "$FILE_KIND" != *"Mach-O"* ]]; then
    echo "ERROR: macOS portability input is not Mach-O: $BINARY ($FILE_KIND)" >&2
    exit 2
fi

if ! DEPENDENCY_OUTPUT="$($OTOOL -L "$BINARY" 2>&1)"; then
    echo "ERROR: otool -L failed for $BINARY: $DEPENDENCY_OUTPUT" >&2
    exit 3
fi
if ! LOAD_COMMAND_OUTPUT="$($OTOOL -l "$BINARY" 2>&1)"; then
    echo "ERROR: otool -l failed for $BINARY: $LOAD_COMMAND_OUTPUT" >&2
    exit 3
fi

is_system_path() {
    case "$1" in
        /System/Library|/System/Library/*|/usr/lib|/usr/lib/*) return 0 ;;
        *) return 1 ;;
    esac
}

VIOLATIONS=()
DEPENDENCIES="$(
    printf '%s\n' "$DEPENDENCY_OUTPUT" |
        /usr/bin/sed -E '1d; s/^[[:space:]]+//; s/[[:space:]]+\(compatibility version.*$//; /^$/d'
)"
while IFS= read -r DEPENDENCY; do
    [ -z "$DEPENDENCY" ] && continue
    if ! is_system_path "$DEPENDENCY"; then
        VIOLATIONS+=("dependency: $DEPENDENCY")
    fi
done <<< "$DEPENDENCIES"

RPATHS="$(
    printf '%s\n' "$LOAD_COMMAND_OUTPUT" |
        /usr/bin/awk '
            $1 == "cmd" && $2 == "LC_RPATH" { expecting_path = 1; next }
            expecting_path && $1 == "path" {
                line = $0
                sub(/^[[:space:]]*path[[:space:]]+/, "", line)
                sub(/[[:space:]]+\(offset[[:space:]][0-9]+\)[[:space:]]*$/, "", line)
                print line
                expecting_path = 0
            }
        '
)"
while IFS= read -r RPATH; do
    [ -z "$RPATH" ] && continue
    if ! is_system_path "$RPATH"; then
        VIOLATIONS+=("LC_RPATH: $RPATH")
    fi
done <<< "$RPATHS"

if [ "${#VIOLATIONS[@]}" -ne 0 ]; then
    echo "ERROR: non-portable macOS release artifact: $BINARY" >&2
    printf '  - %s\n' "${VIOLATIONS[@]}" >&2
    echo "Only /System/Library/** and /usr/lib/** are allowed; publish with the official .NET SDK/CI." >&2
    exit 4
fi

echo "macOS portability check passed: $BINARY"
