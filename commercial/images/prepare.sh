#!/usr/bin/env bash
# Унификация скриншотов для коммерческого предложения Automacon WCS.
#
# Что делает скрипт:
#   1. Для каждого файла screen-*.png рядом со скриптом
#      создаёт оптимизированную копию в ./processed/
#   2. Приводит ширину к 1600 px (если оригинал шире), сохраняет пропорции.
#   3. Лёгкая постобработка: чуть приподнимает контраст и резкость,
#      ограничивает уровень белого, чтобы скриншоты смотрелись
#      аккуратно при печати в коммерческом предложении.
#   4. Опционально добавляет тонкую серую рамку 1 px вокруг кадра.
#
# Скрипт ничего не удаляет — оригиналы остаются на месте.
#
# Требуется ImageMagick (convert или magick).

set -euo pipefail
cd "$(dirname "$0")"

MAGICK="${MAGICK:-}"
if [[ -z "$MAGICK" ]]; then
  if command -v magick >/dev/null 2>&1; then
    MAGICK="magick"
  elif command -v convert >/dev/null 2>&1; then
    MAGICK="convert"
  else
    echo "ImageMagick не найден. Установите 'imagemagick' и повторите."
    exit 1
  fi
fi

mkdir -p processed

shopt -s nullglob
files=( screen-*.png )
if [[ ${#files[@]} -eq 0 ]]; then
  echo "Не найдено файлов screen-*.png в $(pwd)."
  echo "Положите оригиналы скриншотов сюда и запустите скрипт ещё раз."
  exit 0
fi

for src in "${files[@]}"; do
  dst="processed/$src"
  echo "→ обработка: $src"
  "$MAGICK" "$src" \
    -strip \
    -resize '1600x>' \
    -modulate 100,103,100 \
    -level 0%,98%,1.0 \
    -unsharp 0x0.6+0.4+0.005 \
    -bordercolor '#d6dde3' -border 1 \
    -define png:compression-level=9 \
    "$dst"
done

echo
echo "Готово. Оптимизированные копии лежат в ./processed/"
echo "Просмотрите их и при необходимости перенесите в эту папку поверх оригиналов:"
echo "  mv processed/screen-*.png ."
