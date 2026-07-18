#!/usr/bin/env bash
set -euo pipefail

ROOT="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT INT TERM HUP

OPENPYXL_SPEC="${OPENPYXL_SPEC:-openpyxl>=3.1,<4}"

resolve_executable() {
  local candidate="$1"
  if [[ "$candidate" == */* ]]; then
    [[ -x "$candidate" ]] || return 1
    (CDPATH= cd -- "$(dirname -- "$candidate")" && printf '%s/%s\n' "$PWD" "$(basename -- "$candidate")")
  else
    command -v "$candidate" 2>/dev/null
  fi
}

find_babet() {
  local local_babet="$ROOT/bin/babet"

  if [[ -n "${BABET_BIN:-}" ]]; then
    resolve_executable "$BABET_BIN"
    return
  fi

  if [[ -e "$local_babet" ]]; then
    if [[ ! -x "$local_babet" ]]; then
      echo "Babet trouvé dans $local_babet, mais il n'est pas exécutable." >&2
      echo "Exécute : chmod +x '$local_babet'" >&2
      return 2
    fi
    resolve_executable "$local_babet"
    return
  fi

  resolve_executable babet
}

find_lua() {
  local candidate
  if [[ -n "${LUA_BIN:-}" ]]; then
    resolve_executable "$LUA_BIN"
    return
  fi

  for candidate in lua5.5 lua5.4 lua5.3 lua texlua; do
    if resolve_executable "$candidate" >/dev/null 2>&1; then
      resolve_executable "$candidate"
      return
    fi
  done
  return 1
}

run_babet_script() {
  local source="$1"
  local dir="$TMP/babet-$(basename "$source" .lua)-$RANDOM"
  mkdir -p "$dir"
  cp "$ROOT/xlsx.lua" "$ROOT/dataframe.lua" "$dir/"
  cp "$source" "$dir/main.lua"
  "$BABET_CMD" "$dir"
}

create_python_venv() {
  local venv="$TMP/openpyxl-venv"

  if ! command -v python3 >/dev/null 2>&1; then
    echo "python3 est requis pour le test d'interopérabilité openpyxl." >&2
    return 1
  fi

  echo "=== Environnement Python temporaire ==="
  if ! python3 -m venv "$venv"; then
    echo "Impossible de créer le venv Python." >&2
    echo "Sous Debian/Ubuntu, installe généralement le paquet python3-venv." >&2
    return 1
  fi

  local python="$venv/bin/python"
  if [[ ! -x "$python" ]]; then
    echo "Le venv a été créé sans interpréteur Python utilisable." >&2
    return 1
  fi

  echo "Installation temporaire de $OPENPYXL_SPEC"
  mkdir -p "$TMP/pip-cache" "$TMP/pip-tmp"
  PIP_DISABLE_PIP_VERSION_CHECK=1 \
  PIP_CACHE_DIR="$TMP/pip-cache" \
  TMPDIR="$TMP/pip-tmp" \
    "$python" -m pip install --no-input --no-cache-dir "$OPENPYXL_SPEC"

  "$python" -c 'import openpyxl; print("openpyxl : " + openpyxl.__version__)'
  PYTHON_CMD="$python"
}

BABET_CMD=""
BABET_STATUS=0
if BABET_CMD="$(find_babet)"; then
  :
else
  BABET_STATUS=$?
  BABET_CMD=""
  if [[ "$BABET_STATUS" -eq 2 ]]; then
    exit 1
  fi
fi

LUA_CMD=""
if LUA_CMD="$(find_lua)"; then
  :
else
  LUA_CMD=""
fi

ran=0
if [[ -n "$BABET_CMD" ]]; then
  echo "=== Babet : $BABET_CMD ==="
  run_babet_script "$ROOT/selftest.lua"
  ran=1
else
  echo "=== Babet : SKIP (bin/babet et PATH absents) ==="
fi

if [[ -n "$LUA_CMD" ]]; then
  echo "=== Lua standard : $LUA_CMD ==="
  (cd "$ROOT" && "$LUA_CMD" selftest.lua)
  ran=1
else
  echo "=== Lua standard : SKIP (aucun interpréteur trouvé) ==="
fi

if [[ "$ran" -eq 0 ]]; then
  echo "Aucun runtime trouvé. Place Babet dans bin/babet ou définis BABET_BIN/LUA_BIN." >&2
  exit 1
fi

PYTHON_CMD=""
create_python_venv
INPUT="$TMP/openpyxl-input.xlsx"
"$PYTHON_CMD" "$ROOT/tests/interop.py" prepare "$INPUT"

interop_runs=0
if [[ -n "$BABET_CMD" ]]; then
  echo "=== Interopérabilité openpyxl + Babet ==="
  OUTPUT_BABET="$TMP/lua-output-babet.xlsx"
  LUA_XLSX_INPUT="$INPUT" LUA_XLSX_OUTPUT="$OUTPUT_BABET" \
    run_babet_script "$ROOT/tests/interop.lua"
  "$PYTHON_CMD" "$ROOT/tests/interop.py" check "$OUTPUT_BABET"
  interop_runs=$((interop_runs + 1))
fi

if [[ -n "$LUA_CMD" ]]; then
  echo "=== Interopérabilité openpyxl + Lua standard ==="
  OUTPUT_LUA="$TMP/lua-output-standard.xlsx"
  LUA_XLSX_INPUT="$INPUT" LUA_XLSX_OUTPUT="$OUTPUT_LUA" \
    "$LUA_CMD" "$ROOT/tests/interop.lua"
  "$PYTHON_CMD" "$ROOT/tests/interop.py" check "$OUTPUT_LUA"
  interop_runs=$((interop_runs + 1))
fi

if [[ "$interop_runs" -eq 0 ]]; then
  echo "Aucun runtime n'a pu exécuter le test d'interopérabilité." >&2
  exit 1
fi

echo "=== TOUS LES TESTS SONT OK ==="
echo "Le venv Python et tous les fichiers temporaires seront supprimés automatiquement."
