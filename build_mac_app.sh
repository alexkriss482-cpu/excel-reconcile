#!/bin/bash
# Сборка Mac-приложения (.app) для передачи коллеге.
set -euo pipefail

cd "$(dirname "$0")"

echo "Установка зависимостей для сборки..."
python3 -m pip install --user -q -r requirements.txt pyinstaller

echo "Сборка приложения (может занять 1–3 минуты)..."
python3 -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "SverkaOFD" \
  --hidden-import=openpyxl.cell._writer \
  --hidden-import=openpyxl.formatting.rule \
  --osx-bundle-identifier=com.kristall.sverkaofd \
  main.py

DIST_DIR="dist/SverkaOFD.app"
if [[ ! -d "$DIST_DIR" ]]; then
  echo "Ошибка: приложение не создано."
  exit 1
fi

echo ""
echo "Готово: $DIST_DIR"
echo "Для передачи коллеге можно заархивировать:"
echo "  ditto -c -k --sequesterRsrc --keepParent dist/SverkaOFD.app dist/SverkaOFD-mac.zip"
