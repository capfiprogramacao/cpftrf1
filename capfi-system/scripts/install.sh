#!/bin/bash
# ─── CAPFI System — Script de instalação VPS Hostinger ────────────────────────
# Execute este script no terminal do VPS após conectar
# Comando: bash install.sh

set -e  # Para em caso de erro

echo ""
echo "╔══════════════════════════════════════════════════════════╗"
echo "║          CAPFI System — Instalação VPS Hostinger         ║"
echo "╚══════════════════════════════════════════════════════════╝"
echo ""

# ── PASSO 1: Atualizar sistema ─────────────────────────────────────────────────
echo "📦 [1/7] Atualizando sistema..."
apt update && apt upgrade -y
apt install -y git curl wget unzip

# ── PASSO 2: Instalar Docker ───────────────────────────────────────────────────
echo ""
echo "🐳 [2/7] Instalando Docker..."
if ! command -v docker &> /dev/null; then
    curl -fsSL https://get.docker.com -o get-docker.sh
    sh get-docker.sh
    rm get-docker.sh
    echo "✅ Docker instalado!"
else
    echo "✅ Docker já instalado"
fi

# ── PASSO 3: Instalar Docker Compose ──────────────────────────────────────────
echo ""
echo "🔧 [3/7] Instalando Docker Compose..."
if ! command -v docker-compose &> /dev/null; then
    curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" \
         -o /usr/local/bin/docker-compose
    chmod +x /usr/local/bin/docker-compose
    echo "✅ Docker Compose instalado!"
else
    echo "✅ Docker Compose já instalado"
fi

# ── PASSO 4: Criar pasta do sistema ───────────────────────────────────────────
echo ""
echo "📁 [4/7] Criando estrutura de pastas..."
mkdir -p /opt/capfi-system
cd /opt/capfi-system
mkdir -p nginx/ssl postgres n8n-workflows scripts app

# ── PASSO 5: Criar arquivo .env ────────────────────────────────────────────────
echo ""
echo "⚙️  [5/7] Configurando variáveis de ambiente..."

if [ ! -f .env ]; then
    # Gerar senha aleatória para o banco
    DB_PASS=$(openssl rand -base64 16 | tr -d "=+/")
    SECRET=$(openssl rand -base64 24 | tr -d "=+/")

    cat > .env << EOF
# CAPFI System — Configurações
# Edite este arquivo com suas chaves reais

DB_PASSWORD=${DB_PASS}
SECRET_KEY=${SECRET}

# Preencha as chaves abaixo:
CLAUDE_API_KEY=SUA_CHAVE_CLAUDE_AQUI
NECTAR_API_TOKEN=SEU_TOKEN_NECTAR_AQUI
N8N_USER=admin
N8N_PASSWORD=admin123
VPS_IP=$(curl -s ifconfig.me 2>/dev/null || echo "SEU_IP_AQUI")
EOF
    echo "✅ Arquivo .env criado com senhas aleatórias"
    echo ""
    echo "⚠️  IMPORTANTE: Edite o arquivo .env com suas chaves de API:"
    echo "    nano /opt/capfi-system/.env"
    echo ""
else
    echo "✅ Arquivo .env já existe"
fi

# ── PASSO 6: Baixar modelo Llama ───────────────────────────────────────────────
echo ""
echo "🤖 [6/7] Preparando Ollama/Llama..."
echo "   O modelo Llama será baixado na primeira inicialização"
echo "   Tamanho: ~2GB — pode demorar alguns minutos"

# ── PASSO 7: Configurar inicialização automática ───────────────────────────────
echo ""
echo "🔄 [7/7] Configurando inicialização automática..."
cat > /etc/systemd/system/capfi.service << 'EOF'
[Unit]
Description=CAPFI System
After=docker.service
Requires=docker.service

[Service]
Type=oneshot
RemainAfterExit=yes
WorkingDirectory=/opt/capfi-system
ExecStart=/usr/local/bin/docker-compose up -d
ExecStop=/usr/local/bin/docker-compose down
TimeoutStartSec=300

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable capfi.service
echo "✅ Sistema configurado para iniciar automaticamente"

# ── Resumo final ───────────────────────────────────────────────────────────────
VPS_IP=$(curl -s ifconfig.me 2>/dev/null || echo "SEU_IP")

echo ""
echo "╔══════════════════════════════════════════════════════════╗"
echo "║              ✅ Instalação concluída!                    ║"
echo "╚══════════════════════════════════════════════════════════╝"
echo ""
echo "📋 PRÓXIMOS PASSOS:"
echo ""
echo "1. Edite as chaves de API:"
echo "   nano /opt/capfi-system/.env"
echo ""
echo "2. Envie os arquivos do projeto (via GitHub ou SCP)"
echo ""
echo "3. Inicie o sistema:"
echo "   cd /opt/capfi-system"
echo "   docker-compose up -d"
echo ""
echo "4. Baixe o modelo Llama (aguarde ~5 minutos):"
echo "   docker exec capfi_ollama ollama pull llama3.2"
echo ""
echo "5. Acesse o sistema:"
echo "   App principal: http://${VPS_IP}"
echo "   n8n (automações): http://${VPS_IP}/n8n"
echo ""
