#!/usr/bin/env bash
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
cd "$root"

mapfile -t files < <(rg --files -g '*.bas' -g '*.cls')

if [ "${#files[@]}" -eq 0 ]; then
  echo "Nenhum arquivo VBA (.bas/.cls) encontrado."
  exit 1
fi

missing_option_explicit=0
placeholder_only=0

for f in "${files[@]}"; do
  if ! rg -qi '^\s*Option\s+Explicit\b' "$f"; then
    echo "[WARN] Sem Option Explicit: $f"
    missing_option_explicit=$((missing_option_explicit + 1))
  fi

  code_lines=$(awk '
    BEGIN { count = 0 }
    {
      line = $0
      gsub(/^[ \t]+|[ \t]+$/, "", line)
      if (line == "") next
      if (line ~ /^\x27/) next
      count++
    }
    END { print count }
  ' "$f")

  if [ "$code_lines" -eq 0 ]; then
    echo "[WARN] Possível placeholder sem implementação: $f"
    placeholder_only=$((placeholder_only + 1))
  fi
done

echo ""
echo "Resumo:"
echo "- Arquivos sem Option Explicit: $missing_option_explicit"
echo "- Arquivos sem implementação detectável: $placeholder_only"

if [ "$placeholder_only" -gt 0 ]; then
  exit 2
fi
