#!/bin/bash
# ─────────────────────────────────────────────────────────────
#  CAPFI — Instalador Mac
#  Execute uma vez antes de rodar o processar_trf1.py
# ─────────────────────────────────────────────────────────────

echo ""
echo "╔══════════════════════════════════════════════╗"
echo "║   CAPFI — Instalação de dependências (Mac)   ║"
echo "╚══════════════════════════════════════════════╝"
echo ""

# Verifica Python
if ! command -v python3 &>/dev/null; then
    echo "❌ Python3 não encontrado."
    echo "   Instale em: https://www.python.org/downloads/"
    exit 1
fi
echo "✅ Python3: $(python3 --version)"

# Instala pip se necessário
python3 -m ensurepip --upgrade 2>/dev/null

# Instala dependências
echo ""
echo "📦 Instalando openpyxl e playwright..."
pip3 install --upgrade openpyxl playwright

# Instala navegador Chromium
echo ""
echo "🌐 Instalando navegador Chromium (uma vez só, ~150MB)..."
python3 -m playwright install chromium

echo ""
echo "╔══════════════════════════════════════════════╗"
echo "║   ✅ Instalação concluída!                   ║"
echo "║                                              ║"
echo "║   Próximo passo:                             ║"
echo "║   1. Coloque a planilha .xlsx na mesma pasta ║"
echo "║   2. Execute: python3 processar_trf1.py      ║"
echo "╚══════════════════════════════════════════════╝"
echo ""
