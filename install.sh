#!/usr/bin/env bash
# Installs the Python environment for the professional-documentation skill.
# Usage: ./install.sh [--google]   (--google adds the Google Docs/Sheets API client)
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$HOME/.local/share/professional-documentation-venv"
REQ="requirements.txt"; [ "${1:-}" = "--google" ] && REQ="requirements-google.txt"
python3 -m venv "$VENV"
"$VENV/bin/pip" install --quiet -r "$HERE/$REQ"
mkdir -p "$HOME/.config/professional-documentation"; chmod 700 "$HOME/.config/professional-documentation"
echo "Python env ready: $VENV"
command -v soffice >/dev/null || echo "note: LibreOffice (soffice) not found; check_xlsx.py and render checks need it."
[ "$HERE" = "$HOME/.claude/skills/professional-documentation" ] || echo "note: Claude Code loads this skill only from ~/.claude/skills/professional-documentation (clone it there)."
echo "Next: ./tests/run.sh   (and see README.md 'Google Docs and Sheets' for --google setup)"
