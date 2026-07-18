#!/usr/bin/env bash
set -euo pipefail

ROOT="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"

if ! command -v pandoc >/dev/null 2>&1; then
  echo "pandoc est requis." >&2
  exit 1
fi

if command -v xelatex >/dev/null 2>&1; then
  ENGINE=xelatex
elif command -v lualatex >/dev/null 2>&1; then
  ENGINE=lualatex
else
  echo "xelatex ou lualatex est requis." >&2
  exit 1
fi

build_one() {
  local lang="$1"
  local source="$ROOT/documentation-$lang.md"
  local output="$ROOT/documentation-$lang.pdf"
  local tmp
  tmp="$(mktemp -d)"
  trap 'rm -rf "$tmp"' RETURN

  pandoc "$source" \
    --from=gfm \
    --standalone \
    --toc \
    --number-sections \
    --pdf-engine="$ENGINE" \
    --include-in-header="$ROOT/preamble.tex" \
    --metadata-file="$ROOT/metadata-$lang.yaml" \
    -V papersize=a4 \
    -V geometry:margin=2cm \
    -V fontsize=10pt \
    -V mainfont="DejaVu Sans" \
    -V monofont="DejaVu Sans Mono" \
    -V colorlinks=true \
    -V linkcolor=blue \
    -V urlcolor=blue \
    --resource-path="$ROOT" \
    --output="$tmp/output.pdf"

  mv "$tmp/output.pdf" "$output"
  rm -rf "$tmp"
  trap - RETURN
  echo "Créé : $output"
}

build_one fr
build_one en
