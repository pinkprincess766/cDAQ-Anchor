#!/bin/bash
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "Создаю виртуальное окружение..."
  python3 -m venv .venv
  .venv/bin/python -m pip install -r requirements.txt
fi

exec .venv/bin/python main.py "$@"
