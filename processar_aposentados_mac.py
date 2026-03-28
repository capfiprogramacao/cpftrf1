"""
CAPFI — Consulta TRF1 Precatórios
Script para APOSENTADOS — execução local (Mac)

Como usar:
  1. Coloque este arquivo na mesma pasta que "Aposentados Sindireceita.xlsx"
  2. Abra um Terminal na pasta ~/Desktop/cpftrf1
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
EXCEL_PATH  = "Aposentados Sindireceita.xlsx"
URL         = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"
NUM_WORKERS = 5
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

excel_lock = asyncio.Lock()

async def save_result(wb, row, resultado):
    async with excel_lock:
        wb2 = load_workbook(EXCEL_PATH)
        ws2 = wb2.active
        ws2.cell(row, 3).value = resultado.get("proc_precat", "")
        ws2.cell(row, 4).value = resultado.get("proc_orig", "")
        ws2.cell(row, 5).value = resultado.get("prec_exp", "")
        ws2.cell(row, 6).value = resultado.get("ano_orc", "")
        ws2.cell(row, 7).value = resultado.get("data_prop", "")
        ws2.cell(row, 8).value = resultado.get("status", "Erro")
        wb2.save(EXCEL_PATH)

async def consultar_cpf(page, cpf_fmt, primeira_vez):
    res = {"proc_precat": "", "proc_orig": "", "prec_exp": "Não",
           "ano_orc": "", "data_prop": "", "status": "Não encontrado"}
    try:
        if primeira_vez:
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(3)

        # Clica no campo, limpa e preenche o CPF
        inp = page.locator('input[name="cpf_cnpj"]')
        await inp.click()
        await inp.evaluate('el => el.value = ""')
        await inp.fill(clean_cpf(cpf_fmt))

        # Clica no botão de pesquisa
        await page.locator('input#enviar').click()
        await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        body = await page.evaluate('document.body.innerText')
        body_low = body.lower()

        if 'partes encontradas' not in body_low and 'nome da parte' not in body_low:
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return res

        # Clica no primeiro processo da lista
        links = page.locator('table a')
        if await links.count() == 0:
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return res

        await links.first.click()
        await page.wait_for_load_state('networkidle', timeout=20000)
        await asyncio.sleep(2)

        # Extrai dados da tabela de processos
        rows_data = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll("table tr")).map(row => {
                const c = Array.from(row.querySelectorAll("td"));
                const l = row.querySelector("a");
                return { col1: c[0] ? c[0].innerText.trim() : "",
                         col2: c[1] ? c[1].innerText.trim() : "",
                         href: l ? l.href : null };
            }).filter(r => r.href && r.col1.length > 0);
        }''')

        precat_rows = [r for r in rows_data if '(PRC)' in r['col1'] or '(PREC)' in r['col1']]
        if not precat_rows:
            res["status"] = "Sem PRECAT"
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return res

        pr = precat_rows[-1]
        res["proc_precat"] = pr['col1'].strip()
        res["proc_orig"] = pr['col2'].split('/')[0].strip() if '/' in pr['col2'] else pr['col2'].strip()
        res["prec_exp"] = "Sim"

        # Abre o processo PRECAT
        todos = page.locator('table a', has_text='PRC')
        n = await todos.count()
        await (todos.nth(n - 1) if n > 0 else todos.first).click()
        await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        # Verifica aba Movimentação
        mov = page.locator('a', has_text='Movimentação')
        if await mov.count() == 0:
            res["status"] = "Encontrado - Sem aba Movimentação"
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return res

        await mov.click()
        await page.wait_for_load_state('networkidle', timeout=20000)
        await asyncio.sleep(2)

        ml = (await page.evaluate('document.body.innerText')).lower()
        mr = await page.evaluate('''() =>
            Array.from(document.querySelectorAll("table tr"))
            .map(row => Array.from(row.querySelectorAll("td")).map(c => c.innerText.trim()))
            .filter(r => r.length >= 3)''')

        prop = False
        for r in mr:
            joined = ' '.join(r).lower()
            if 'proposta orçamentária' in joined and 'cjf' in joined:
                prop = True
                anos = re.findall(r'\b(20\d{2})\b', r[-1] if r else '')
                if anos: res["ano_orc"] = anos[0]
                m = re.search(r'(\d{2}/\d{2}/\d{4})', r[0] if r else '')
                if m: res["data_prop"] = m.group(1)
                break

        if any(kw in ml for kw in KEYWORDS_PAGO):
            res["status"] = "Pago/Excluir"
        elif prop:
            res["status"] = "Encontrado"
        else:
            res["status"] = "Encontrado - Sem proposta orçamentária"

        await page.goto(URL, timeout=30000, wait_until='networkidle')
        await asyncio.sleep(2)

    except PlaywrightTimeout:
        res["status"] = "Timeout"
        print(f"    ⚠️  Timeout em {cpf_fmt}", flush=True)
        try:
            await page.goto(URL, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except: pass

    except Exception as e:
        res["status"] = "Erro"
        print(f"    ❌ Erro em {cpf_fmt}: {e}", flush=True)
        try:
            await page.goto(URL, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except: pass

    return res

async def worker(worker_id, fila, wb, lock, playwright, contagem):
    # Lança browser com proteções anti-detecção
    browser = await playwright.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled"
        ]
    )
    ctx = await browser.new_context(
        user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 800},
        locale="pt-BR"
    )
    page = await ctx.new_page()
    await page.add_init_script('Object.defineProperty(navigator, "webdriver", {get: () => undefined})')

    primeira_vez = True
    try:
        while True:
            try:
                row, nome, cpf_raw = fila.get_nowait()
            except Exception:
                break

            cpf_fmt = format_cpf(clean_cpf(str(cpf_raw)))
            resultado = await consultar_cpf(page, cpf_fmt, primeira_vez)
            primeira_vez = False

            await save_result(wb, row, resultado)

            async with lock:
                contagem[0] += 1
                emoji = "✅" if resultado["status"] == "Encontrado" else \
                        "🔴" if "Pago" in resultado["status"] else \
                        "⚪" if resultado["status"] in ("Sem PRECAT", "Não encontrado") else "❌"
                print(f"[W{worker_id}][{contagem[0]:04d}] {nome[:35]:<35} {cpf_fmt} → {emoji} {resultado['status']}", flush=True)

            fila.task_done()
            await asyncio.sleep(1)
    finally:
        await browser.close()

async def main():
    excel = Path(EXCEL_PATH)
    if not excel.exists():
        print(f"❌ Arquivo não encontrado: {EXCEL_PATH}")
        print(f"   Pasta atual: {Path.cwd()}")
        sys.exit(1)

    print(f"📂 Arquivo: {excel.resolve()}")
    wb, ws = preparar_planilha()

    # Coleta pendentes
    pendentes = []
    for row in range(2, ws.max_row + 1):
        nome   = ws.cell(row, 1).value
        cpf    = ws.cell(row, 2).value
        status = ws.cell(row, 8).value
        if not nome or not cpf: continue
        if len(clean_cpf(str(cpf))) != 11: continue
        if str(status).strip() in STATUS_SKIP: continue
        pendentes.append((row, str(nome).strip(), str(cpf).strip()))

    total  = ws.max_row - 1
    feitos = total - len(pendentes)
    print(f"\n{'='*60}")
    print(f"  CAPFI — TRF1 | Aposentados Sindireceita")
    print(f"  Workers  : {NUM_WORKERS} navegadores em paralelo")
    print(f"  Total    : {total} | Feitos: {feitos} | Pendentes: {len(pendentes)}")
    print(f"  Início   : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"{'='*60}\n")

    if not pendentes:
        print("✅ Todos os CPFs já foram processados!")
        return

    fila = asyncio.Queue()
    for item in pendentes:
        fila.put_nowait(item)

    lock     = asyncio.Lock()
    contagem = [0]

    async with async_playwright() as p:
        await asyncio.gather(*[
            worker(i + 1, fila, wb, lock, p, contagem)
            for i in range(NUM_WORKERS)
        ])

    elapsed = (datetime.now() - datetime.now()).seconds
    print(f"\n🏁 Concluído! {contagem[0]} CPFs processados.")
    print(f"📁 Salvo em: {excel.resolve()}")

if __name__ == "__main__":
    asyncio.run(main())
