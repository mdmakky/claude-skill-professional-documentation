#!/usr/bin/env bash
# Offline smoke test: builds every template, then checks numbering, Unicode and workbook formulas.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PY="$HOME/.local/share/professional-documentation-venv/bin/python3"
exec "$PY" -I "$HERE/tests/test_smoke.py" "$HERE"
