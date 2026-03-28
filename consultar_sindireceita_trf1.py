"""
CAPFI — Consulta TRF1 | SINDIRECEITA_COMPLETO.xlsx
Fluxo por LINHA (cada linha faz sua própria consulta completa):
  1. Busca CPF no TRF1
  2. Localiza o PRC vinculado ao PROCESSO ORIGINÁRIO da linha (col 3)
  3. Acessa o PRC, abre Movimentação, extrai o ANO LOA
  4. Salva PRECAT (col 4) e ORÇAMENTO (col 5)
  5. Repete para a próxima linha

Como usar (Mac):
  cd ~/Desktop/cpftrf1
  python3 reset_status_sindireceita.py
  caffeinate -i python3 consultar_sindireceita_trf1.py
"""

import asyncio
import re
import sys
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from openpyxl import load_workbook

# ─── CONFIGURAÇÃO ──────────────────────────────────────────────
EXCEL_PATH  = "SINDIRECEITA_COMPLETO.xlsx"
URL_BUSCA   = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"
NUM_WORKERS = 3
# ───────────────────────────────────────────────────────────────

STATUS_SKIP = {"OK", "Sem PRECAT", "Não encontrado", "Erro permanente"}

COL_NOME   = 1
COL_CPF    = 2
COL_ORIG   = 3
COL_PRECAT = 4
COL_ORC    = 5
COL_STATUS = 12

def clean_cpf(cpf):
    return re.sub(r'\D', '', str(cpf).strip())

def format_cpf(cpf):
    d = clean_cpf(cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d) == 11 else str(cpf)

def normaliza_orig(orig):
    return re.sub(r'/\w+\s*$', '', str(orig or '').strip()).strip()

excel_lock = asyncio.Lock()

async def salvar_linha(row_num, precat, orcamento, status):
    async with excel_lock:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
        if precat    is not None: ws.cell(row_num, COL_PRECAT).value = precat
        if orcamento is not None: ws.cell(row_num, COL_ORC).value   = int(orcamento)
        ws.cell(row_num, COL_STATUS).value = status
        wb.save(EXCEL_PATH)

def encontrar_prc(processos, orig_planilha):
    orig_norm = normaliza_orig(orig_planilha)
    matches = [p for p in processos
               if ('(PRC)' in p['col1'] or '(PREC)' in p['col1'])
               and normaliza_orig(p['col2']) == orig_norm]
    if matches:
        return matches[-1]
    fallback = [p for p in processos if '(PRC)' in p['col1'] or '(PREC)' in p['col1']]
    return fallback[-1] if fallback else None

async def processar_linha(page, row_num, cpf_raw, orig_proc, nome, primeira_vez):
    """Faz a consulta completa para uma única linha e retorna (precat, loa, status)."""

    # ── 1. Vai à página de busca ─────────────────────────────────
    if primeira_vez:
        await page.goto(URL_BUSCA, timeout=30000, wait_until='networkidle')
        await asyncio.sleep(3)
    else:
        await page.goto(URL_BUSCA, timeout=30000, wait_until='networkidle')
        await asyncio.sleep(1)

    # ── 2. Preenche CPF e envia ──────────────────────────────────
    inp = page.locator('input[name="cpf_cnpj"]')
    await inp.click()
    await inp.evaluate('el => el.value = ""')
    await inp.fill(clean_cpf(cpf_raw))
    await page.locator('input#enviar').click()
    await page.wait_for_load_state('networkidle', timeout=30000)
    await asyncio.sleep(2)

    body = await page.evaluate('document.body.innerText')
    if 'partes encontradas' not in body.lower() and 'nome da parte' not in body.lower():
        return None, None, "Não encontrado"

    # ── 3. Clica no nome da pessoa ───────────────────────────────
    await page.locator('table a').first.click()
    await page.wait_for_load_state('networkidle', timeout=25000)
    await asyncio.sleep(2)

    # ── 4. Extrai lista de processos ─────────────────────────────
    processos = await page.evaluate('''() => {
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

    # ── 5. Encontra o PRC pelo ORIG desta linha ──────────────────
    prc = encontrar_prc(processos, orig_proc)
    if prc is None:
        return None, None, "Sem PRECAT"

    precat = prc['col1'].strip()

    # ── 6. Navega ao PRC pelo href direto ────────────────────────
    prc_href = prc.get('href')
    if prc_href and prc_href.startswith('http'):
        await page.goto(prc_href, timeout=30000, wait_until='networkidle')
    else:
        prc_texto = prc['col1'][:20]
        await page.locator(f'a:has-text("{prc_texto}")').first.click()
        await page.wait_for_load_state('networkidle', timeout=30000)
    await asyncio.sleep(3)

    # ── 7. Clica na aba Movimentação ─────────────────────────────
    mov_link = page.locator('a', has_text='Movimentação')
    if await mov_link.count() == 0:
        return precat, None, "OK"

    await mov_link.first.click()
    # Aguarda mais tempo — conteúdo carrega via AJAX
    await page.wait_for_load_state('networkidle', timeout=30000)
    await asyncio.sleep(4)

    # ── 8. Extrai o ANO LOA ──────────────────────────────────────
    linhas_mov = await page.evaluate('''() =>
        Array.from(document.querySelectorAll("table tr"))
        .map(row => Array.from(row.querySelectorAll("td")).map(c => c.innerText.trim()))
        .filter(r => r.length >= 2)
    ''')

    loa = None
    for ln in linhas_mov:
        texto = ' '.join(ln)
        if 'cjf' in texto.lower() and 'exerc' in texto.lower():
            # Ano LOA está na última célula: ex. "2027,data 13/02/2026"
            for cell in reversed(ln):
                anos = re.findall(r'\b(20\d{2})\b', cell)
                for ano in anos:
                    if int(ano) >= 2024:
                        loa = ano
                        break
                if loa:
                    break
            if not loa:
                # Fallback: procura no texto completo
                m = re.search(r'exerc\S+\s+de\s+(20\d{2})', texto, re.IGNORECASE)
                if m:
                    loa = m.group(1)
            break

    return precat, loa, "OK"


async def worker(worker_id, fila, lock, playwright, contagem, total):
    browser = await playwright.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage",
              "--disable-blink-features=AutomationControlled"]
    )
    ctx = await browser.new_context(
        user_agent=("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"),
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
                row_num, cpf_raw, orig, nome = fila.get_nowait()
            except Exception:
                break

            try:
                precat, loa, status = await processar_linha(
                    page, row_num, cpf_raw, orig, nome, primeira_vez
                )
                primeira_vez = False

                await salvar_linha(row_num, precat, loa, status)

                async with lock:
                    contagem[0] += 1
                    emoji = "✅" if status == "OK" else "⚪"
                    print(
                        f"[W{worker_id}][{contagem[0]:03d}/{total}] "
                        f"{nome[:24]:<24} | "
                        f"ORIG: {orig[:28]:<28} | "
                        f"PRECAT: {(precat or '-')[:28]} | "
                        f"LOA: {loa or '-'} {emoji}",
                        flush=True
                    )

            except PlaywrightTimeout:
                primeira_vez = False
                await salvar_linha(row_num, None, None, "Timeout")
                async with lock:
                    contagem[0] += 1
                    print(f"[W{worker_id}][{contagem[0]:03d}/{total}] "
                          f"⚠️  Timeout: {nome[:30]}", flush=True)
                try:
                    await page.goto(URL_BUSCA, timeout=15000, wait_until='domcontentloaded')
                    await asyncio.sleep(2)
                except: pass

            except Exception as e:
                primeira_vez = False
                await salvar_linha(row_num, None, None, "Erro")
                async with lock:
                    contagem[0] += 1
                    print(f"[W{worker_id}][{contagem[0]:03d}/{total}] "
                          f"❌ Erro {nome[:20]}: {e}", flush=True)
                try:
                    await page.goto(URL_BUSCA, timeout=15000, wait_until='domcontentloaded')
                    await asyncio.sleep(2)
                except: pass

            fila.task_done()
            await asyncio.sleep(0.5)
    finally:
        await browser.close()


async def main():
    excel = Path(EXCEL_PATH)
    if not excel.exists():
        print(f"❌ Arquivo não encontrado: {EXCEL_PATH}")
        print(f"   Pasta atual: {Path.cwd()}")
        sys.exit(1)

    print(f"📂 Arquivo: {excel.resolve()}")

    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    if ws.cell(1, COL_STATUS).value is None:
        ws.cell(1, COL_STATUS).value = "STATUS_CONSULTA"
        wb.save(EXCEL_PATH)

    wb  = load_workbook(EXCEL_PATH)
    ws  = wb.active
    pendentes = []
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
        if status not in STATUS_SKIP:
            pendentes.append((row, str(cpf_raw), orig, nome))

    total   = ws.max_row - 1
    feitas  = total - len(pendentes)

    print(f"\n{'='*70}")
    print(f"  CAPFI — TRF1 | SINDIRECEITA COMPLETO")
    print(f"  Workers  : {NUM_WORKERS}")
    print(f"  Linhas   : {total} total | Feitas: {feitas} | Pendentes: {len(pendentes)}")
    print(f"  Início   : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"{'='*70}\n")

    if not pendentes:
        print("✅ Todas as linhas já foram processadas!")
        return

    fila     = asyncio.Queue()
    for item in pendentes:
        fila.put_nowait(item)

    lock     = asyncio.Lock()
    contagem = [0]

    async with async_playwright() as p:
        await asyncio.gather(*[
            worker(i + 1, fila, lock, p, contagem, len(pendentes))
            for i in range(min(NUM_WORKERS, len(pendentes)))
        ])

    print(f"\n🏁 Concluído! {contagem[0]}/{total} linhas processadas.")
    print(f"📁 Arquivo salvo: {excel.resolve()}")


if __name__ == "__main__":
    asyncio.run(main())
