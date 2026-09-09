#!/bin/bash
# Baixa a versao mais recente da planilha publicada pelo GitHub Pages
# e guarda nesta pasta. Basta dar dois cliques neste arquivo.
cd "$(dirname "$0")" || exit 1
URL="https://pcwaguilera.github.io/curtailment-ons/curtailment_solar.xlsx"
DEST="curtailment_solar.xlsx"

echo "Baixando de $URL ..."
if curl -fL --retry 3 --progress-bar -o "$DEST.parcial" "$URL"; then
  mv "$DEST.parcial" "$DEST"
  TAM=$(du -h "$DEST" | cut -f1)
  echo ""
  echo "Pronto: $(pwd)/$DEST  ($TAM)"
  echo "Atualizado em $(date '+%d/%m/%Y %H:%M')"
  open -R "$DEST"
else
  rm -f "$DEST.parcial"
  echo ""
  echo "Nao consegui baixar."
  echo "Se o site ainda nao existe, rode o workflow uma vez ate o fim:"
  echo "  https://github.com/pcwaguilera/curtailment-ons/actions"
fi
echo ""
read -n 1 -s -r -p "Pressione qualquer tecla para fechar."
