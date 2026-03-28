"""
CAPFI — Consulta TRF1 Precatórios
Script para APOSENTADOS — execução local (Mac)

Como usar:
  1. Coloque este arquivo na mesma pasta que "Aposentados Sindireceita.xlsx"
  2. Abra um novo Terminal (separado do Dependentes)
  3. Execute: caffeinate -i python3 processar_aposentados_mac.py
"""

import asyncio
import re
import sys
from datetime import datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from openpyxl import load_workbook

# ─── CONFIGURAÇÃO ─────────────────────────────────────────────
EXCEL_PATH      = "Aposentados Sindireceita.xlsx"
URL             = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"
NUM_WORKERS     = 5   # navegadores simultâneos
# ──────────────────────────────────────────────────────────────

KEYWORDS_PAGO = [
    "depósito","deposito","saque","pagamento","pago","levantamento",
    "alvará","alvar","requisição de pequeno valor","rpv levantado",
    "quitação","quitado","crédito levantado","transferência","ordem de pagamento expedida"
]
STATUS_SKIP = {
    "Sem CPF","CPF inválido","Encontrado","Pago/Excluir","Sem PRECAT",
    "Não encontrado","Encontrado - Sem proposta orçamentária","Encontrado - Sem aba Movimentação"
}


def clean_cpf(cpf): return re.sub(r'\D', '', str(cpf).strip())

def format_cpf(cpf):
    d = clean_cpf(cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d) == 11 else cpf

def preparar_planilha():
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    if ws.cell(1, 3).value is None:
        for i, h in enumerate(["Processo PRECAT", "Processo Originário", "Precatório Expedido",
                                "Ano Orçamentário", "Data Proposta", "Status"], start=3):
            ws.cell(row=1, column=i).value = h
        for row in range(2, ws.max_row + 1):
            cpf = ws.cell(row, 2).value
            if not cpf: continue
            d = clean_cpf(str(cpf))
            if len(d) != 11:
                ws.cell(row, 8).value = "CPF inválido" if cpf else "Sem CPF"
        wb.save(EXCEL_PATH)
        print("✅ Colunas criadas na planilha")
    return wb, ws

async def consultar_cpf(page, cpf_fmt):
    await page.goto(URL, timeout=30000)
    await page.fill('input[name="cbf"]', cpf_fmt)
    await page.press('input[name="cbf"]', 'Enter')
    await page.wait_for_load_state('networkidle', timeout=30000)

    content = await page.content()
    texto   = await page.inner_text('body')
    texto_low = texto.lower()

    if 'nenhum processo encontrado' in texto_low or 'não foram encontrados' in texto_low:
        return {"status": "Não encontrado"}

    proc_match = re.search(r'(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})', content)
    if not proc_match:
        return {"status": "Não encontrado"}

    proc_precat = proc_match.group(1)
    proc_orig   = ""
    ano_orc     = ""
    data_prop   = ""
    status      = "Encontrado"
    prec_exp    = "Não"

    # Tenta acessar aba Movimentação
    try:
        await page.click('text=Movimentação', timeout=5000)
        await page.wait_for_load_state('networkidle', timeout=10000)
        mov_texto = (await page.inner_text('body')).lower()
        if any(k in mov_texto for k in KEYWORDS_PAGO):
            status = "Pago/Excluir"
        prec_exp = "Sim"
    except Exception:
        status = "Encontrado - Sem aba Movimentação"

    return {
        "proc_precat": proc_precat,
        "proc_orig":   proc_orig,
        "prec_exp":    prec_exp,
        "ano_orc":     ano_orc,
        "data_prop":   data_prop,
        "status":      status,
    }

async def worker(worker_id, fila, wb, ws, lock, playwright, contagem):
    browser = await playwright.chromium.launch(headless=True)
    page    = await browser.new_page()
    try:
        while True:
            try:
                row = fila.get_nowait()
            except Exception:
                break

            cpf_raw = ws.cell(row, 2).value
            cpf_fmt = format_cpf(clean_cpf(str(cpf_raw)))

            try:
                resultado = await consultar_cpf(page, cpf_fmt)
            except Exception as e:
                resultado = {"status": "Erro"}

            async with lock:
                ws.cell(row, 3).value = resultado.get("proc_precat", "")
                ws.cell(row, 4).value = resultado.get("proc_orig",   "")
                ws.cell(row, 5).value = resultado.get("prec_exp",    "")
                ws.cell(row, 6).value = resultado.get("ano_orc",     "")
                ws.cell(row, 7).value = resultado.get("data_prop",   "")
                ws.cell(row, 8).value = resultado.get("status",      "Erro")
                wb.save(EXCEL_PATH)
                contagem[0] += 1
                print(f"[W{worker_id}] [{contagem[0]}] Linha {row} | {cpf_fmt} → {resultado.get('status')}", flush=True)

            fila.task_done()
    finally:
        await browser.close()

async def main():
    excel = Path(EXCEL_PATH)
    if not excel.exists():
        print(f"❌ Arquivo não encontrado: {EXCEL_PATH}")
        print(f"   Certifique-se de estar na pasta correta: {Path.cwd()}")
        sys.exit(1)

    print(f"📂 Arquivo: {excel.resolve()}")
    wb, ws = preparar_planilha()

    # Coleta linhas pendentes
    pendentes = []
    for row in range(2, ws.max_row + 1):
        cpf_raw = ws.cell(row, 2).value
        if not cpf_raw: continue
        d = clean_cpf(str(cpf_raw))
        if len(d) != 11: continue
        status = ws.cell(row, 8).value
        if status in STATUS_SKIP: continue
        pendentes.append(row)

    total = ws.max_row - 1
    feitos = total - len(pendentes)
    print(f"📊 Total: {total} | Já feitos: {feitos} | Pendentes: {len(pendentes)}")
    print(f"🚀 Iniciando {NUM_WORKERS} workers em paralelo...\n")

    if not pendentes:
        print("✅ Todos os CPFs já foram processados!")
        return

    fila = asyncio.Queue()
    for r in pendentes:
        fila.put_nowait(r)

    lock     = asyncio.Lock()
    contagem = [0]
    inicio   = datetime.now()

    async with async_playwright() as p:
        await asyncio.gather(*[
            worker(i + 1, fila, wb, ws, lock, p, contagem)
            for i in range(NUM_WORKERS)
        ])

    elapsed = (datetime.now() - inicio).seconds
    print(f"\n🏁 Concluído! {contagem[0]} CPFs em {elapsed // 60}m{elapsed % 60}s")
    print(f"📁 Salvo em: {excel.resolve()}")

if __name__ == "__main__":
    asyncio.run(main())
