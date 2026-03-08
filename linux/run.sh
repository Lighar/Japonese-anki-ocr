#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

source "$SCRIPT_DIR/venv/bin/activate"
pip install -q -r "$SCRIPT_DIR/requirements.txt"
python jap.py
