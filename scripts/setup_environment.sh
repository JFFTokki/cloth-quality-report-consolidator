#!/bin/zsh
set -euo pipefail

PROJECT_ROOT="${0:A:h:h}"
VENV_PATH="$PROJECT_ROOT/.venv"
ARTIFACT_VENDOR="$PROJECT_ROOT/quality-report-consolidator/vendor/js/artifact-tool"
ARTIFACT_NODE_PATH="$PROJECT_ROOT/quality-report-consolidator/node_modules/@oai/artifact-tool"

if ! command -v pdftoppm >/dev/null 2>&1 || ! command -v pdfinfo >/dev/null 2>&1; then
  echo "缺少 Poppler。请先执行：brew install poppler" >&2
  exit 1
fi

python3 -m venv "$VENV_PATH"
"$VENV_PATH/bin/python" -m pip install --upgrade pip
"$VENV_PATH/bin/python" -m pip install -r "$PROJECT_ROOT/requirements.txt"

if [[ ! -f "$ARTIFACT_VENDOR/index.mjs" ]]; then
  echo "项目包缺少内置Node依赖：$ARTIFACT_VENDOR" >&2
  exit 1
fi
if [[ ! -e "$ARTIFACT_NODE_PATH" ]]; then
  mkdir -p "${ARTIFACT_NODE_PATH:h}"
  ln -s "../../vendor/js/artifact-tool" "$ARTIFACT_NODE_PATH"
fi

"$VENV_PATH/bin/python" "$PROJECT_ROOT/pipeline/qc_all/check_environment.py" --smoke-ocr
