# ─── Imagem base com Playwright + Chromium já instalados ─────────────────────
FROM mcr.microsoft.com/playwright/python:v1.44.0-jammy

# Diretório de trabalho
WORKDIR /app

# Instalar dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar o browser Chromium via Playwright
RUN playwright install chromium

# Copiar código da aplicação
COPY . .

# Criar pastas necessárias
RUN mkdir -p uploads results

# Expor porta
EXPOSE 8000

# Iniciar servidor
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
