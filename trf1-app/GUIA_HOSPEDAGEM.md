# Guia de Hospedagem — TRF1 Consulta de Precatórios

## Estrutura do projeto

```
trf1-app/
├── main.py              ← Backend FastAPI (API + WebSocket)
├── scraper.py           ← Automação do navegador (Playwright)
├── requirements.txt     ← Dependências Python
├── Dockerfile           ← Configuração do container
├── docker-compose.yml   ← Para rodar localmente com Docker
├── .dockerignore
└── static/
    └── index.html       ← Frontend (interface do usuário)
```

---

## Pré-requisito: publicar o código no GitHub

Todas as plataformas de hospedagem leem o código do GitHub.

1. Crie uma conta em https://github.com (gratuito)
2. Crie um repositório novo (ex: `trf1-precatorios`)
3. Faça upload de todos os arquivos da pasta `trf1-app/`
4. Anote a URL do repositório (ex: `https://github.com/seu-usuario/trf1-precatorios`)

---

## Opção 1 — Railway (Recomendado para começar)

**Por quê Railway?**
- Interface mais simples do mercado
- Detecta o Dockerfile automaticamente
- Deploy em menos de 5 minutos
- Plano gratuito: 500 horas/mês + 512 MB RAM

**Limitação:** O Playwright + Chromium precisa de ~600 MB. No plano gratuito
pode haver lentidão. Para uso regular, o plano pago ($5/mês) é suficiente.

### Passo a passo Railway

1. Acesse https://railway.app e clique em **"Start a New Project"**
2. Clique em **"Deploy from GitHub repo"**
3. Autorize o Railway a acessar seu GitHub
4. Selecione o repositório `trf1-precatorios`
5. Railway detecta o `Dockerfile` automaticamente — clique **Deploy**
6. Aguarde o build (3–5 minutos na primeira vez)
7. Após concluir, clique em **"Generate Domain"** para obter a URL pública
8. Acesse a URL — o app estará funcionando!

### Custo Railway
| Plano | Preço | RAM | Indicado para |
|---|---|---|---|
| Hobby (gratuito) | $0 | 512 MB | Testes |
| Pro | $5/mês | 1 GB+ | Uso regular |

---

## Opção 2 — Render

**Por quê Render?**
- Plano gratuito generoso
- Deploy via Dockerfile simples
- Boa documentação

**Limitação:** No plano gratuito, o serviço "dorme" após 15 minutos sem uso
(primeira requisição demora ~30s para acordar). Para uso contínuo, o plano
pago ($7/mês) mantém o serviço sempre ativo.

### Passo a passo Render

1. Acesse https://render.com e crie uma conta
2. Clique em **"New +"** → **"Web Service"**
3. Conecte seu GitHub e selecione o repositório `trf1-precatorios`
4. Configure:
   - **Name:** `trf1-precatorios`
   - **Environment:** `Docker`
   - **Region:** `Ohio (US East)` ou `Frankfurt (EU)` (mais próximo do Brasil)
   - **Instance Type:** Free (ou Starter $7/mês para produção)
5. Clique em **"Create Web Service"**
6. Aguarde o build (5–8 minutos)
7. Render fornece uma URL do tipo `https://trf1-precatorios.onrender.com`
8. Acesse a URL — app funcionando!

### Custo Render
| Plano | Preço | RAM | Observação |
|---|---|---|---|
| Free | $0 | 512 MB | Dorme após 15min sem uso |
| Starter | $7/mês | 512 MB | Sempre ativo |
| Standard | $25/mês | 2 GB | Melhor para Playwright |

---

## Opção 3 — VPS (Recomendado para produção)

Para uso com 2.000+ CPFs de forma estável, uma VPS é a melhor opção.
Você tem controle total e mais memória disponível.

**Provedores recomendados:**
- **DigitalOcean** — Droplet básico a partir de $6/mês (1 GB RAM)
- **Vultr** — A partir de $6/mês (1 GB RAM)
- **Contabo** — A partir de €4,99/mês (4 GB RAM — melhor custo-benefício)

### Passo a passo VPS (DigitalOcean como exemplo)

1. Crie conta em https://digitalocean.com
2. Crie um **Droplet** com:
   - Sistema: Ubuntu 22.04
   - Plano: Basic — 1 GB RAM / 1 CPU ($6/mês)
3. Conecte via SSH:
   ```bash
   ssh root@IP_DO_SERVIDOR
   ```
4. Instale Docker:
   ```bash
   curl -fsSL https://get.docker.com | sh
   ```
5. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/trf1-precatorios.git
   cd trf1-precatorios
   ```
6. Inicie o app:
   ```bash
   docker compose up -d
   ```
7. O app estará rodando em `http://IP_DO_SERVIDOR:8000`

### Configurar domínio + HTTPS (opcional)
```bash
# Instalar Nginx
apt install nginx

# Configurar proxy reverso para a porta 8000
# Instalar SSL gratuito com Let's Encrypt
apt install certbot python3-certbot-nginx
certbot --nginx -d seu-dominio.com
```

---

## Resumo comparativo

| Critério | Railway | Render | VPS |
|---|---|---|---|
| Facilidade | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐ |
| Custo mensal | $0–5 | $0–7 | $6–10 |
| RAM disponível | 512MB–1GB | 512MB–2GB | 1GB–ilimitado |
| Performance Playwright | Média | Média | Alta |
| Indicado para | Testes / inicial | Testes / inicial | Produção (2000+ CPFs) |

---

## Testar localmente antes de hospedar

Se tiver Docker instalado no computador:

```bash
cd trf1-app
docker compose up --build
```

Acesse http://localhost:8000 — o app funcionará igual à versão hospedada.
