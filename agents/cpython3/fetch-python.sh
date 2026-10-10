#!/bin/bash
# Downloads the official Windows embeddable CPython that the cpython3 engine ships (agents/cpython3/python-embed/, git-ignored).
# Pinned by version and SHA-256 so builds are reproducible; run automatically by Build.sh when the folder is missing.
set -euo pipefail
PY_VERSION=3.12.8
PY_SHA256=8d3f33be9eb810f23c102f08475af2854e50484b8e4e06275e937be61ce3d2fb
HERE="$(cd "$(dirname "$0")" && pwd)"
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
