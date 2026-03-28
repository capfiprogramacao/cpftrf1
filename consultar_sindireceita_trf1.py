"""
CAPFI — Consulta TRF1 por CPF
Planilha: SINDIRECEITA_COMPLETO.xlsx
Atualiza PROCESSO PRECAT (col 4) e ORÇAMENTO/LOA (col 5)

91 CPFs únicos → 445 linhas.
Para cada CPF, acessa a lista de processos no TRF1, encontra o PRC cujo
PROCESSO ORIGINÁRIO (col 2 da tabela) corresponde ao ORIG da planilha,
navega nesse PRC, extrai o ano LOA da aba Movimentação e salva.

Como usar (Mac):
  cd ~/Desktop/cpftrf1
  caffeinate -i python3 consultar_sindireceita_trf1.py
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
URL_BUSCA   = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"
NUM_WORKERS = 5
# ───────────────────────────────────────────────────────────────

STATUS_SKIP = {"OK", "Sem PRECAT", "Não encontrado", "Erro permanente"}

COL_NOME    = 1
COL_CPF     = 2
COL_ORIG    = 3
COL_PRECAT  = 4
COL_ORC     = 5
COL_STATUS  = 12   # coluna de controle (fora dos dados originais)

def clean_cpf(cpf):
    return re.sub(r'\D', '', str(cpf).strip())

def format_cpf(cpf):
    d = clean_cpf(cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d) == 11 else str(cpf)

def normaliza_orig(orig):
    """Remove sufixo /UF e espaços para comparação."""
    return re.sub(r'/\w+\s*$', '', str(orig or '').strip()).strip()

excel_lock = asyncio.Lock()

async def salvar_resultado(rows, precat, orcamento, status):
    async with excel_lock:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
        for row in rows:
            if precat:
                ws.cell(row, COL_PRECAT).value = precat
            if orcamento:
                ws.cell(row, COL_ORC).value = int(orcamento)
            ws.cell(row, COL_STATUS).value = status
        wb.save(EXCEL_PATH)

async def extrair_processos(page):
    """Extrai todos os processos da lista atual (col1=número, col2=orig_ref, href=link)."""
    return await page.evaluate('''() => {
        const linhas = [];
        document.querySelectorAll("table tr").forEach(row => {
            const cols = Array.from(row.querySelectorAll("td"));
            if (cols.length < 2) return;
            const c1 = cols[0].innerText.trim();
            const c2 = cols[1].innerText.trim();
            const link = row.querySelector("a");
            if (!c1 || !c2) return;
            linhas.push({ col1: c1, col2: c2, href: link ? link.href : null });
        });
        return linhas;
    }''')

async def consultar_cpf(page, cpf_raw, orig_proc, primeira_vez):
    resultado = {"precat": None, "orcamento": None, "status": "Não encontrado"}
    cpf_fmt  = format_cpf(cpf_raw)
    orig_norm = normaliza_orig(orig_proc)

    try:
        if primeira_vez:
            await page.goto(URL_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(3)

        # ── Preenche CPF e envia ──────────────────────────────────
        inp = page.locator('input[name="cpf_cnpj"]')
        await inp.click()
        await inp.evaluate('el => el.value = ""')
        await inp.fill(clean_cpf(cpf_raw))
        await page.locator('input#enviar').click()
        await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        body = await page.evaluate('document.body.innerText')
        if 'partes encontradas' not in body.lower() and 'nome da parte' not in body.lower():
            await page.goto(URL_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        # ── Clica no nome da pessoa (primeiro link da tabela) ─────
        await page.locator('table a').first.click()
        await page.wait_for_load_state('networkidle', timeout=25000)
        await asyncio.sleep(2)

        # ── Extrai lista de processos ─────────────────────────────
        processos = await extrair_processos(page)

        # ── Filtra PRCs que correspondem ao ORIG da planilha ──────
        prc_matches = [
            p for p in processos
            if ('(PRC)' in p['col1'] or '(PREC)' in p['col1'])
            and normaliza_orig(p['col2']) == orig_norm
        ]

        # Fallback: qualquer PRC (se ORIG não encontrar match exato)
        if not prc_matches:
            prc_matches = [p for p in processos if '(PRC)' in p['col1'] or '(PREC)' in p['col1']]

        if not prc_matches:
            resultado["status"] = "Sem PRECAT"
            await page.goto(URL_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        # Usa o PRC mais recente com match correto
        prc = prc_matches[-1]
        resultado["precat"] = prc['col1'].strip()

        # ── Abre o processo PRC ───────────────────────────────────
        if prc['href']:
            await page.goto(prc['href'], timeout=30000, wait_until='networkidle')
        else:
            # Clica no link pelo texto
            prc_texto = prc['col1'][:20]
            await page.locator(f'a:has-text("{prc_texto}")').first.click()
            await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        # ── Clica na aba Movimentação ─────────────────────────────
        mov_link = page.locator('a', has_text='Movimentação')
        if await mov_link.count() > 0:
            await mov_link.first.click()
            await page.wait_for_load_state('networkidle', timeout=25000)
            await asyncio.sleep(2)

            # Extrai linhas da aba de movimentação
            linhas_mov = await page.evaluate('''() =>
                Array.from(document.querySelectorAll("table tr"))
                .map(row => Array.from(row.querySelectorAll("td"))
                    .map(c => c.innerText.trim()))
                .filter(r => r.length >= 2)
            ''')

            for ln in linhas_mov:
                texto_full = ' '.join(ln)
                if 'proposta orçamentária' in texto_full.lower() and 'cjf' in texto_full.lower():
                    # Extrai ano especificamente após "DO EXERCÍCIO DE"
                    m = re.search(r'exercício\s+de\s+(\d{4})', texto_full, re.IGNORECASE)
                    if m:
                        resultado["orcamento"] = m.group(1)
                    break

        resultado["status"] = "OK"
        await page.goto(URL_BUSCA, timeout=30000, wait_until='networkidle')
        await asyncio.sleep(2)

    except PlaywrightTimeout:
        resultado["status"] = "Timeout"
        print(f"    ⚠️  Timeout: {cpf_fmt}", flush=True)
        try:
            await page.goto(URL_BUSCA, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except: pass

    except Exception as e:
        resultado["status"] = "Erro"
        print(f"    ❌ Erro em {cpf_fmt}: {e}", flush=True)
        try:
            await page.goto(URL_BUSCA, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except: pass

    return resultado


async def worker(worker_id, fila, lock, playwright, contagem):
    browser = await playwright.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
        ]
    )
    ctx = await browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        viewport={"width": 1280, "height": 800},
        locale="pt-BR",
    )
    page = await ctx.new_page()
    await page.add_init_script(
        'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
    )

    primeira_vez = True
    try:
        while True:
            try:
                cpf_raw, orig_proc, nome, rows = fila.get_nowait()
            except Exception:
                break

            resultado = await consultar_cpf(page, cpf_raw, orig_proc, primeira_vez)
            primeira_vez = False

            await salvar_resultado(rows, resultado["precat"], resultado["orcamento"], resultado["status"])

            async with lock:
                contagem[0] += 1
                emoji = "✅" if resultado["status"] == "OK" else \
                        "⚪" if resultado["status"] in ("Sem PRECAT", "Não encontrado") else "❌"
                precat_str = resultado["precat"] or "-"
                orc_str    = str(resultado["orcamento"]) if resultado["orcamento"] else "-"
                print(
                    f"[W{worker_id}][{contagem[0]:03d}] {nome[:28]:<28} | "
                    f"PRECAT: {precat_str[:35]} | LOA: {orc_str} {emoji}",
                    flush=True
                )

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

    # Garante cabeçalho na coluna de controle
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    if ws.cell(1, COL_STATUS).value is None:
        ws.cell(1, COL_STATUS).value = "STATUS_CONSULTA"
        wb.save(EXCEL_PATH)

    # Agrupa linhas por CPF
    wb  = load_workbook(EXCEL_PATH)
    ws  = wb.active
    cpf_grupos = {}
    for row in range(2, ws.max_row + 1):
        nome    = str(ws.cell(row, COL_NOME).value or '').strip()
        cpf_raw = ws.cell(row, COL_CPF).value
        orig    = str(ws.cell(row, COL_ORIG).value or '').strip()
        status  = str(ws.cell(row, COL_STATUS).value or '').strip()
        if not cpf_raw:
            continue
        cpf_c = clean_cpf(str(cpf_raw))
        if len(cpf_c) != 11:
            continue
        if cpf_c not in cpf_grupos:
            cpf_grupos[cpf_c] = {"nome": nome, "orig": orig, "rows": [], "status": status}
        cpf_grupos[cpf_c]["rows"].append(row)
        cpf_grupos[cpf_c]["status"] = status   # usa o último

    pendentes = [
        (cpf, data["orig"], data["nome"], data["rows"])
        for cpf, data in cpf_grupos.items()
        if data["status"] not in STATUS_SKIP
    ]

    total  = len(cpf_grupos)
    feitos = total - len(pendentes)

    print(f"\n{'='*68}")
    print(f"  CAPFI — TRF1 | SINDIRECEITA COMPLETO")
    print(f"  Workers  : {NUM_WORKERS}")
    print(f"  CPFs     : {total} únicos | Feitos: {feitos} | Pendentes: {len(pendentes)}")
    print(f"  Linhas   : {ws.max_row - 1} total")
    print(f"  Início   : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"{'='*68}\n")

    if not pendentes:
        print("✅ Todos os CPFs já foram processados!")
        return

    fila     = asyncio.Queue()
    for item in pendentes:
        fila.put_nowait(item)

    lock     = asyncio.Lock()
    contagem = [0]

    async with async_playwright() as p:
        await asyncio.gather(*[
            worker(i + 1, fila, lock, p, contagem)
            for i in range(min(NUM_WORKERS, len(pendentes)))
        ])

    print(f"\n🏁 Concluído! {contagem[0]} CPFs consultados.")
    print(f"📁 Arquivo salvo: {excel.resolve()}")


if __name__ == "__main__":
    asyncio.run(main())
