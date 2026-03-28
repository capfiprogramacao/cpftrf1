"""
CAPFI — Consulta TRF1 | SINDIRECEITA_COMPLETO.xlsx
Fluxo por linha:
  1. Acessa TRF1 pelo CPF da linha
  2. Lê o PROCESSO ORIGINÁRIO do Excel (col 3) e encontra o PRC correspondente no TRF1
  3. Extrai PROCESSO PRECAT e ANO LOA
  4. Atualiza a linha
  5. Repete para todas as linhas do mesmo CPF (mesma consulta TRF1, novo ORIG se diferente)
  6. Segue para o próximo CPF

Como usar (Mac):
  cd ~/Desktop/cpftrf1
  python3 reset_status_sindireceita.py     ← limpa dados antigos
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
NUM_WORKERS = 3     # 3 workers para estabilidade
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
        if precat   is not None: ws.cell(row_num, COL_PRECAT).value = precat
        if orcamento is not None: ws.cell(row_num, COL_ORC).value  = int(orcamento)
        ws.cell(row_num, COL_STATUS).value = status
        wb.save(EXCEL_PATH)

# ──────────────────────────────────────────────────────────────
#  Extrai lista de processos da página TRF1 atual
# ──────────────────────────────────────────────────────────────
async def extrair_lista_processos(page):
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

# ──────────────────────────────────────────────────────────────
#  Busca CPF no TRF1 e retorna a lista de processos da pessoa
# ──────────────────────────────────────────────────────────────
async def buscar_cpf_trf1(page, cpf_raw, primeira_vez):
    if primeira_vez:
        await page.goto(URL_BUSCA, timeout=30000, wait_until='networkidle')
        await asyncio.sleep(3)

    inp = page.locator('input[name="cpf_cnpj"]')
    await inp.click()
    await inp.evaluate('el => el.value = ""')
    await inp.fill(clean_cpf(cpf_raw))
    await page.locator('input#enviar').click()
    await page.wait_for_load_state('networkidle', timeout=30000)
    await asyncio.sleep(2)

    body = await page.evaluate('document.body.innerText')
    if 'partes encontradas' not in body.lower() and 'nome da parte' not in body.lower():
        return None   # CPF não encontrado

    await page.locator('table a').first.click()
    await page.wait_for_load_state('networkidle', timeout=25000)
    await asyncio.sleep(2)

    return await extrair_lista_processos(page)

# ──────────────────────────────────────────────────────────────
#  Navega ao PRC e extrai o ano LOA da aba Movimentação
# ──────────────────────────────────────────────────────────────
async def extrair_loa_do_prc(page, prc):
    if prc.get('href'):
        await page.goto(prc['href'], timeout=30000, wait_until='networkidle')
    else:
        prc_texto = prc['col1'][:20]
        await page.locator(f'a:has-text("{prc_texto}")').first.click()
        await page.wait_for_load_state('networkidle', timeout=30000)
    await asyncio.sleep(2)

    mov_link = page.locator('a', has_text='Movimentação')
    if await mov_link.count() == 0:
        return None

    await mov_link.first.click()
    await page.wait_for_load_state('networkidle', timeout=25000)
    await asyncio.sleep(2)

    linhas_mov = await page.evaluate('''() =>
        Array.from(document.querySelectorAll("table tr"))
        .map(row => Array.from(row.querySelectorAll("td")).map(c => c.innerText.trim()))
        .filter(r => r.length >= 2)
    ''')

    for ln in linhas_mov:
        texto = ' '.join(ln)
        if 'proposta orçamentária' in texto.lower() and 'cjf' in texto.lower():
            m = re.search(r'exercício\s+de\s+(\d{4})', texto, re.IGNORECASE)
            if m:
                return m.group(1)
    return None

# ──────────────────────────────────────────────────────────────
#  Encontra o PRC na lista que corresponde ao ORIG da linha
# ──────────────────────────────────────────────────────────────
def encontrar_prc(processos, orig_planilha):
    orig_norm = normaliza_orig(orig_planilha)
    # Match exato por ORIG
    matches = [p for p in processos
               if ('(PRC)' in p['col1'] or '(PREC)' in p['col1'])
               and normaliza_orig(p['col2']) == orig_norm]
    if matches:
        return matches[-1]
    # Fallback: qualquer PRC
    fallback = [p for p in processos if '(PRC)' in p['col1'] or '(PREC)' in p['col1']]
    return fallback[-1] if fallback else None

# ──────────────────────────────────────────────────────────────
#  Worker: processa um grupo de CPF de cada vez
#  Para cada CPF: busca TRF1, depois percorre cada linha do grupo
# ──────────────────────────────────────────────────────────────
async def worker(worker_id, fila_cpfs, lock, playwright, contagem, total_linhas):
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
            # Pega próximo grupo (um CPF com todas suas linhas)
            try:
                cpf_raw, linhas_do_cpf = fila_cpfs.get_nowait()
            except Exception:
                break

            nome_ref = linhas_do_cpf[0]['nome']
            cpf_fmt  = format_cpf(cpf_raw)

            try:
                # ── Etapa 1: Busca CPF no TRF1 ───────────────────────────
                processos = await buscar_cpf_trf1(page, cpf_raw, primeira_vez)
                primeira_vez = False

                if processos is None:
                    # CPF não encontrado — marca todas as linhas
                    for linha in linhas_do_cpf:
                        await salvar_linha(linha['row'], None, None, "Não encontrado")
                        async with lock:
                            contagem[0] += 1
                            print(f"[W{worker_id}][{contagem[0]:03d}/{total_linhas}] "
                                  f"{linha['nome'][:26]:<26} | CPF não encontrado ⚪",
                                  flush=True)
                    await page.goto(URL_BUSCA, timeout=20000, wait_until='networkidle')
                    await asyncio.sleep(2)
                    fila_cpfs.task_done()
                    continue

                # Salva URL da lista de processos para voltar depois
                lista_url = page.url

                # ── Etapas 2-4: Para cada linha deste CPF ────────────────
                # Cache de (orig_norm → (precat, loa)) para este CPF
                orig_cache = {}

                for linha in linhas_do_cpf:
                    row_num = linha['row']
                    orig    = linha['orig']
                    orig_n  = normaliza_orig(orig)

                    if orig_n in orig_cache:
                        precat, loa = orig_cache[orig_n]
                        status = "OK" if precat else "Sem PRECAT"
                        cache_tag = " [=]"
                    else:
                        # Etapa 2: encontra o PRC pelo ORIG da linha
                        prc = encontrar_prc(processos, orig)

                        if prc is None:
                            orig_cache[orig_n] = (None, None)
                            precat, loa, status = None, None, "Sem PRECAT"
                            cache_tag = ""
                        else:
                            precat = prc['col1'].strip()
                            # Etapa 3: acessa PRC e extrai LOA
                            loa = await extrair_loa_do_prc(page, prc)
                            status = "OK"
                            orig_cache[orig_n] = (precat, loa)
                            cache_tag = ""
                            # Volta para a lista de processos (sem go_back)
                            await page.goto(lista_url, timeout=25000, wait_until='networkidle')
                            await asyncio.sleep(1)

                    # Etapa 4 (implícita): salva esta linha
                    await salvar_linha(row_num, precat, loa, status)

                    async with lock:
                        contagem[0] += 1
                        emoji = "✅" if status == "OK" else "⚪"
                        precat_str = precat or "-"
                        loa_str    = str(loa) if loa else "-"
                        print(
                            f"[W{worker_id}][{contagem[0]:03d}/{total_linhas}] "
                            f"{linha['nome'][:24]:<24} | "
                            f"ORIG: {orig[:30]:<30} | "
                            f"PRECAT: {precat_str[:30]} | LOA: {loa_str}{cache_tag} {emoji}",
                            flush=True
                        )

            except PlaywrightTimeout:
                print(f"[W{worker_id}] ⚠️  Timeout no CPF {cpf_fmt}", flush=True)
                for linha in linhas_do_cpf:
                    current_status = linha.get('status', '')
                    if current_status not in STATUS_SKIP:
                        await salvar_linha(linha['row'], None, None, "Timeout")
                        async with lock:
                            contagem[0] += 1
                try:
                    await page.goto(URL_BUSCA, timeout=15000, wait_until='domcontentloaded')
                    await asyncio.sleep(2)
                except: pass

            except Exception as e:
                print(f"[W{worker_id}] ❌ Erro no CPF {cpf_fmt}: {e}", flush=True)
                for linha in linhas_do_cpf:
                    current_status = linha.get('status', '')
                    if current_status not in STATUS_SKIP:
                        await salvar_linha(linha['row'], None, None, "Erro")
                        async with lock:
                            contagem[0] += 1
                try:
                    await page.goto(URL_BUSCA, timeout=15000, wait_until='domcontentloaded')
                    await asyncio.sleep(2)
                except: pass

            # ── Etapa 5: próximo CPF ──────────────────────────────────
            await page.goto(URL_BUSCA, timeout=20000, wait_until='networkidle')
            await asyncio.sleep(1)
            fila_cpfs.task_done()
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

    # Lê todas as linhas e agrupa por CPF (mantendo ordem de aparição)
    wb  = load_workbook(EXCEL_PATH)
    ws  = wb.active
    cpf_grupos = OrderedDict()

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
            cpf_grupos[cpf_c] = {'cpf_raw': str(cpf_raw), 'linhas': []}
        cpf_grupos[cpf_c]['linhas'].append({
            'row': row, 'nome': nome, 'orig': orig, 'status': status
        })

    # Filtra grupos com pelo menos uma linha pendente
    grupos_pendentes = []
    total_linhas_pendentes = 0
    for cpf_c, data in cpf_grupos.items():
        linhas_pend = [l for l in data['linhas'] if l['status'] not in STATUS_SKIP]
        if linhas_pend:
            grupos_pendentes.append((data['cpf_raw'], data['linhas']))
            total_linhas_pendentes += len(data['linhas'])

    total_cpfs   = len(cpf_grupos)
    total_linhas = sum(len(d['linhas']) for d in cpf_grupos.values())
    feitas = total_linhas - total_linhas_pendentes

    print(f"\n{'='*75}")
    print(f"  CAPFI — TRF1 | SINDIRECEITA COMPLETO")
    print(f"  Workers  : {NUM_WORKERS}")
    print(f"  CPFs     : {total_cpfs} únicos")
    print(f"  Linhas   : {total_linhas} total | Feitas: {feitas} | Pendentes: {total_linhas_pendentes}")
    print(f"  Início   : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"{'='*75}\n")

    if not grupos_pendentes:
        print("✅ Todas as linhas já foram processadas!")
        return

    fila_cpfs = asyncio.Queue()
    for item in grupos_pendentes:
        fila_cpfs.put_nowait(item)

    lock     = asyncio.Lock()
    contagem = [0]

    async with async_playwright() as p:
        await asyncio.gather(*[
            worker(i + 1, fila_cpfs, lock, p, contagem, total_linhas_pendentes)
            for i in range(min(NUM_WORKERS, len(grupos_pendentes)))
        ])

    print(f"\n🏁 Concluído! {contagem[0]}/{total_linhas} linhas processadas.")
    print(f"📁 Arquivo salvo: {excel.resolve()}")


if __name__ == "__main__":
    asyncio.run(main())
