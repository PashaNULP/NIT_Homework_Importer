#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
cd "$SCRIPT_DIR"

if [[ ! -x ".venv/bin/python" ]]; then
  echo "Не знайдено .venv. Спочатку виконайте інструкцію встановлення з README_NIT_Importer.md"
  echo "Натисніть будь-яку клавішу, щоб закрити вікно."
  read -k 1
  exit 1
fi

exec .venv/bin/python nit_web.py
