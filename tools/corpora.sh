#!/bin/sh
# Pinned public Office fixtures, kept outside the repository. Resumes interrupted fetches.
# Usage: sh tools/corpora.sh [DIR] [libreoffice|poi|openxml-sdk ...]
# Supplemental slicer/timeline inputs: sh tools/corpora.sh DIR ledger-features
set -eu
exec python3 "$(dirname "$0")/ooxml/corpora.py" "$@"
