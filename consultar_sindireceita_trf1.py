"""
CAPFI — Consulta TRF1 por CPF
Planilha: SINDIRECEITA_COMPLETO.xlsx
Atualiza PROCESSO PRECAT (col 4) e ORÇAMENTO (col 5)

91 CPFs únicos → 445 linhas no total.
O script agrupa as linhas por CPF, consulta o TRF1 uma vez por CPF
e atualiza todas as linhas correspondentes.

Como usar:
  1. Coloque este script na mesma pasta que SINDIRECEITA_COMPLETO.xlsx
  2. Execute: caffeinate -i python3 consultar_sindireceita_trf1.py
"""

import asyncio
import re
import sys
from datetime import datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from openpyxl import load_workbook

# ─── CONFIGURAÇÃO ──────────────────────────────────────────────
EXCEL_PATH  = "SINDIRECEITA_COMPLETO.xlsx"
URL         = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"
NUM_WORKERS = 5
# ───────────────────────────────────────────────────────────────

STATUS_SKIP = {"OK", "Sem PRECAT", "Não encontrado", "Erro permanente"}

COL_PRECAT  = 4
COL_ORC     = 5
COL_STATUS  = 12   # coluna extra de controle (fora dos dados originais)

def clean_cpf(cpf):
    return re.sub(r'\D', '', str(cpf).strip())

def format_cpf(cpf):
    d = clean_cpf(cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d) == 11 else cpf

excel_lock = asyncio.Lock()

async def salvar_resultado(rows, precat, orcamento, status):
    """Atualiza todas as linhas do CPF de uma vez."""
    async with excel_lock:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
        for row in rows:
            if precat:
                ws.cell(row, COL_PRECAT).value = precat
            if orcamento:
                ws.cell(row, COL_ORC).value = orcamento
            ws.cell(row, COL_STATUS).value = status
        wb.save(EXCEL_PATH)

async def consultar_cpf(page, cpf_raw, primeira_vez):
    resultado = {"precat": None, "orcamento": None, "status": "Não encontrado"}
    cpf_fmt = format_cpf(cpf_raw)

    try:
        if primeira_vez:
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(3)

        # Preenche o CPF
        inp = page.locator('input[name="cpf_cnpj"]')
        await inp.click()
        await inp.evaluate('el => el.value = ""')
        await inp.fill(clean_cpf(cpf_raw))

        # Clica em Pesquisar
        await page.locator('input#enviar').click()
        await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        body = await page.evaluate('document.body.innerText')
        body_low = body.lower()

        # Sem resultados
        if 'partes encontradas' not in body_low and 'nome da parte' not in body_low:
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        # Clica no primeiro link (nome da pessoa)
        links = page.locator('table a')
        if await links.count() == 0:
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        await links.first.click()
        await page.wait_for_load_state('networkidle', timeout=20000)
        await asyncio.sleep(2)

        # Extrai processos da tabela
        rows_data = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll("table tr")).map(row => {
                const cols = Array.from(row.querySelectorAll("td"));
                const link = row.querySelector("a");
                return {
                    col1: cols[0] ? cols[0].innerText.trim() : "",
                    col2: cols[1] ? cols[1].innerText.trim() : "",
                    href: link ? link.href : null
                };
            }).filter(r => r.href && r.col1.length > 0);
        }''')

        # Procura processos PRECAT (PRC ou PREC)
        precat_rows = [r for r in rows_data if '(PRC)' in r['col1'] or '(PREC)' in r['col1']]

        if not precat_rows:
            resultado["status"] = "Sem PRECAT"
            await page.goto(URL, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        # Usa o último PRC encontrado
        pr = precat_rows[-1]
        resultado["precat"] = pr['col1'].strip()

        # Abre o processo PRC
        todos = page.locator('table a', has_text='PRC')
        n = await todos.count()
        await (todos.nth(n - 1) if n > 0 else todos.first).click()
        await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        # Clica na aba Movimentação
        mov = page.locator('a', has_text='Movimentação')
        if await mov.count() > 0:
            await mov.click()
            await page.wait_for_load_state('networkidle', timeout=20000)
            await asyncio.sleep(2)

            mr = await page.evaluate('''() =>
                Array.from(document.querySelectorAll("table tr"))
                .map(row => Array.from(row.querySelectorAll("td"))
                .map(c => c.innerText.trim()))
                .filter(r => r.length >= 3)''')

            for r in mr:
                joined = ' '.join(r).lower()
                if 'proposta orçamentária' in joined and 'cjf' in joined:
                    anos = re.findall(r'\b(20\d{2})\b', r[-1] if r else '')
                    if anos:
                        resultado["orcamento"] = anos[0]
                    break

        resultado["status"] = "OK"
        await page.goto(URL, timeout=30000, wait_until='networkidle')
        await asyncio.sleep(2)

    except PlaywrightTimeout:
        resultado["status"] = "Timeout"
        print(f"    ⚠️  Timeout: {cpf_fmt}", flush=True)
        try:
            await page.goto(URL, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except: pass

    except Exception as e:
        resultado["status"] = "Erro"
        print(f"    ❌ Erro em {cpf_fmt}: {e}", flush=True)
        try:
            await page.goto(URL, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except: pass

    return resultado

async def worker(worker_id, fila, lock, playwright, contagem):
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
                cpf_raw, nome, rows = fila.get_nowait()
            except Exception:
                break

            resultado = await consultar_cpf(page, cpf_raw, primeira_vez)
            primeira_vez = False

            await salvar_resultado(rows, resultado["precat"], resultado["orcamento"], resultado["status"])

            async with lock:
                contagem[0] += 1
                emoji = "✅" if resultado["status"] == "OK" else \
                        "⚪" if resultado["status"] in ("Sem PRECAT", "Não encontrado") else "❌"
                precat_str = resultado["precat"] or "-"
                orc_str    = resultado["orcamento"] or "-"
                print(f"[W{worker_id}][{contagem[0]:03d}] {nome[:30]:<30} | PRECAT: {precat_str[:35]} | Orç: {orc_str} {emoji}", flush=True)

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

    # Garante coluna de controle
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    if ws.cell(1, COL_STATUS).value is None:
        ws.cell(1, COL_STATUS).value = "STATUS_CONSULTA"
        wb.save(EXCEL_PATH)

    # Carrega e agrupa por CPF
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

    cpf_grupos = {}   # cpf_raw -> {"nome": ..., "rows": [...], "status": ...}
    for row in range(2, ws.max_row + 1):
        nome    = ws.cell(row, 1).value
        cpf_raw = ws.cell(row, 2).value
        status  = str(ws.cell(row, COL_STATUS).value or "").strip()
        if not cpf_raw:
            continue
        cpf_clean = clean_cpf(str(cpf_raw))
        if len(cpf_clean) != 11:
            continue
        if cpf_clean not in cpf_grupos:
            cpf_grupos[cpf_clean] = {"nome": str(nome or "").strip(), "rows": [], "status": status}
        cpf_grupos[cpf_clean]["rows"].append(row)
        # Usa o status mais recente (pode diferir entre linhas, usa o da última)
        cpf_grupos[cpf_clean]["status"] = status

    # Filtra pendentes
    pendentes = [
        (cpf, data["nome"], data["rows"])
        for cpf, data in cpf_grupos.items()
        if data["status"] not in STATUS_SKIP
    ]

    total  = len(cpf_grupos)
    feitos = total - len(pendentes)

    print(f"\n{'='*65}")
    print(f"  CAPFI — TRF1 | SINDIRECEITA COMPLETO")
    print(f"  Workers  : {NUM_WORKERS} navegadores em paralelo")
    print(f"  CPFs     : {total} únicos | Feitos: {feitos} | Pendentes: {len(pendentes)}")
    print(f"  Linhas   : {ws.max_row - 1} total")
    print(f"  Início   : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"{'='*65}\n")

    if not pendentes:
        print("✅ Todos os CPFs já foram consultados!")
        return

    fila = asyncio.Queue()
    for item in pendentes:
        fila.put_nowait(item)

    lock     = asyncio.Lock()
    contagem = [0]

    async with async_playwright() as p:
        await asyncio.gather(*[
            worker(i + 1, fila, lock, p, contagem)
            for i in range(min(NUM_WORKERS, len(pendentes)))
        ])

    print(f"\n🏁 Concluído! {contagem[0]} CPFs consultados → {ws.max_row - 1} linhas atualizadas.")
    print(f"📁 Salvo em: {excel.resolve()}")

if __name__ == "__main__":
    asyncio.run(main())
