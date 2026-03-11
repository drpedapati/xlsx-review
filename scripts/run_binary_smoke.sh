#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
BINARY_PATH=""
WORK_DIR=""

usage() {
  cat <<'EOF'
Usage: ./scripts/run_binary_smoke.sh --binary <path> [--work-dir <path>]

Options:
  --binary <path>    Binary to smoke-test
  --work-dir <path>  Scratch directory for smoke outputs (default: mktemp -d)
  -h, --help         Show help
EOF
}

while (($# > 0)); do
  case "$1" in
    --binary)
      BINARY_PATH="$2"
      shift 2
      ;;
    --work-dir)
      WORK_DIR="$2"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ -z "${BINARY_PATH}" ]]; then
  echo "--binary is required" >&2
  usage >&2
  exit 1
fi

if [[ ! -x "${BINARY_PATH}" ]]; then
  echo "Binary not found or not executable: ${BINARY_PATH}" >&2
  exit 1
fi

cleanup() {
  if [[ -n "${TEMP_WORK_DIR:-}" && -d "${TEMP_WORK_DIR}" ]]; then
    rm -rf "${TEMP_WORK_DIR}"
  fi
}
trap cleanup EXIT

if [[ -z "${WORK_DIR}" ]]; then
  TEMP_WORK_DIR="$(mktemp -d)"
  WORK_DIR="${TEMP_WORK_DIR}"
else
  mkdir -p "${WORK_DIR}"
fi

cd "${ROOT_DIR}"

echo "Smoke-testing binary: ${BINARY_PATH}"
"${BINARY_PATH}" --version > "${WORK_DIR}/version.txt"
grep -q '^xlsx-review ' "${WORK_DIR}/version.txt"

"${BINARY_PATH}" examples/test_old.xlsx --read --json > "${WORK_DIR}/read.json"
ruby -rjson -e 'j = JSON.parse(File.read(ARGV[0])); abort("sheet count") unless j.dig("workbook", "sheet_count") == 2; abort("sheet name") unless j.dig("sheets", 0, "name") == "Data"' "${WORK_DIR}/read.json"

"${BINARY_PATH}" --diff examples/test_old.xlsx examples/test_new.xlsx --json > "${WORK_DIR}/diff.json"
ruby -rjson -e 'j = JSON.parse(File.read(ARGV[0])); abort("diff changed") unless j.dig("summary", "identical") == false' "${WORK_DIR}/diff.json"

"${BINARY_PATH}" examples/test_old.xlsx examples/sample-edits.json -o "${WORK_DIR}/edited.xlsx" --json > "${WORK_DIR}/edit.json"
grep -q '"success": true' "${WORK_DIR}/edit.json"

"${BINARY_PATH}" --create -o "${WORK_DIR}/created.xlsx" examples/sample-create.json --json > "${WORK_DIR}/create.json"
grep -q '"success": true' "${WORK_DIR}/create.json"

"${BINARY_PATH}" "${WORK_DIR}/created.xlsx" --read --json > "${WORK_DIR}/created-read.json"
ruby -rjson -e 'j = JSON.parse(File.read(ARGV[0])); abort("created workbook") unless j.dig("workbook", "sheet_count") == 2' "${WORK_DIR}/created-read.json"

echo "✅ Binary smoke passed"
