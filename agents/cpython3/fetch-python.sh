#!/bin/bash
# Downloads the official Windows embeddable CPython that the cpython3 engine ships (agents/cpython3/python-embed/, git-ignored).
# Pinned by version and SHA-256 so builds are reproducible; run automatically by Build.sh when the folder is missing.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
# PY_VERSION and PY_SHA256 live in python-embed.txt (shared with fetch-python.ps1)
# shellcheck source=python-embed.txt
. "$HERE/python-embed.txt"
DEST="$HERE/python-embed"
[ -f "$DEST/python312.dll" ] && exit 0
URL="https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT
echo "fetching $URL"
curl -fsSL -o "$TMP/py.zip" "$URL"
echo "$PY_SHA256  $TMP/py.zip" | sha256sum -c - >/dev/null || { echo "SHA-256 mismatch for $URL" >&2; exit 1; }
mkdir -p "$DEST" && unzip -q -o "$TMP/py.zip" -d "$DEST"
# not needed for embedding (and the executables only invite antivirus attention)
rm -f "$DEST/python.exe" "$DEST/pythonw.exe"
echo "python-embed ready: $(ls "$DEST" | wc -l) files"
