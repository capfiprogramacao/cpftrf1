FROM mcr.microsoft.com/playwright/python:v1.44.0-jammy

WORKDIR /app

# Instalar dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Playwright Chromium
RUN playwright install chromium

# Copiar código
COPY . .

# Criar pastas necessárias
RUN mkdir -p uploads results checkpoints

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
