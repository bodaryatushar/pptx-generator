#!/usr/bin/env bash
# Render a .pptx to PNGs with LibreOffice for quick visual checks.
# Usage: scripts/preview.sh out/deck.pptx [outdir]
set -euo pipefail
in="${1:?pptx path}"
out="${2:-$(dirname "$in")/preview}"
mkdir -p "$out"
soffice --headless --convert-to pdf --outdir "$out" "$in" >/dev/null 2>&1
pdf="$out/$(basename "${in%.*}").pdf"
pdftoppm -png -r 60 "$pdf" "$out/slide"
echo "$out"
