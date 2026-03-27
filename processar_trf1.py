"""
CAPFI — Consulta TRF1 Precatórios
Versão para execução local (Mac/Windows/Linux)

Como usar:
  1. Coloque este arquivo na mesma pasta da planilha Excel
  2. Execute: python3 processar_trf1.py

Checkpoint automático: se interromper, retoma de onde parou.
"""

import asyncio
import re
import sys
import json
from datetime import datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from openpyxl import load_workbook

# ─── CONFIGURAÇÃO ─────────────────────────────────────────────
# Coloque o nome correto da sua planilha aqui:
EXCEL_PATH     = "Dependentes Sindireceita.xlsx"
CHECKPOINT_PATH = "checkpoint_trf1.json"
URL_CPF_BUSCA  = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"

# ──────────────────────────────────────────────────────────────

KEYWORDS_PAGO = [
    "depósito", "deposito", "saque", "pagamento", "pago",
    "levantamento", "alvará", "alvar", "requisição de pequeno valor",
    "rpv levantado", "quitação", "quitado", "crédito levantado",
    "transferência", "ordem de pagamento expedida"
]

STATUS_SKIP = {
    "Sem CPF", "CPF inválido", "Encontrado", "Pago/Excluir",
    "Sem PRECAT", "Não encontrado",
    "Encontrado - Sem proposta orçamentária",
    "Encontrado - Sem aba Movimentação"
}


def clean_cpf(cpf):
    return re.sub(r'\D', '', str(cpf).strip())


def format_cpf(cpf):
    d = clean_cpf(cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d) == 11 else cpf


def preparar_planilha():
    """Adiciona cabeçalhos e pré-preenche CPFs inválidos se ainda não feito."""
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

    # Verifica se cabeçalhos já existem
    if ws.cell(1, 3).value is None:
        headers = [
            "Processo PRECAT", "Processo Originário", "Precatório Expedido",
            "Ano Orçamentário", "Data Proposta", "Status"
        ]
        for i, h in enumerate(headers, start=3):
            ws.cell(row=1, column=i).value = h

        # Pré-preenche inválidos
        for row in range(2, ws.max_row + 1):
            cpf = ws.cell(row, 2).value
            if not cpf or str(cpf).strip() == '':
                ws.cell(row, 8).value = "Sem CPF"
            else:
                digits = re.sub(r'\D', '', str(cpf).strip())
                if len(digits) != 11:
                    ws.cell(row, 8).value = "CPF inválido"

        wb.save(EXCEL_PATH)
        print("✅ Planilha preparada com colunas de resultado.\n")


def load_pendentes():
    """Carrega lista de CPFs pendentes (pula já processados)."""
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    pendentes = []
    for row in range(2, ws.max_row + 1):
        nome   = ws.cell(row, 1).value
        cpf    = ws.cell(row, 2).value
        status = ws.cell(row, 8).value
        if not nome or not cpf:
            continue
        digits = re.sub(r'\D', '', str(cpf))
        if len(digits) != 11:
            continue
        st = str(status).strip() if status else ""
        if st in STATUS_SKIP:
            continue
        pendentes.append((row, str(nome).strip(), str(cpf).strip()))
    return pendentes


def update_row(row_idx, dados):
    """Salva resultado de um CPF na planilha imediatamente."""
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    ws.cell(row=row_idx, column=3).value = dados.get("processo_precat", "")
    ws.cell(row=row_idx, column=4).value = dados.get("processo_originario", "")
    ws.cell(row=row_idx, column=5).value = dados.get("precatorio_expedido", "")
    ws.cell(row=row_idx, column=6).value = dados.get("ano_orcamentario", "")
    ws.cell(row=row_idx, column=7).value = dados.get("data_proposta", "")
    ws.cell(row=row_idx, column=8).value = dados.get("status", "")
    wb.save(EXCEL_PATH)


def salvar_checkpoint(processados, total):
    """Salva progresso para retomada em caso de interrupção."""
    with open(CHECKPOINT_PATH, "w") as f:
        json.dump({"processados": processados, "total": total,
                   "atualizado": datetime.now().isoformat()}, f)


def clean_processo_originario(texto):
    if texto and '/' in texto:
        return texto.split('/')[0].strip()
    return texto.strip() if texto else ""


async def consultar_cpf(page, cpf, primeira_vez):
    resultado = {
        "processo_precat": "", "processo_originario": "",
        "precatorio_expedido": "Não", "ano_orcamentario": "",
        "data_proposta": "", "status": "Não encontrado"
    }
    try:
        cpf_digits = clean_cpf(cpf)

        if primeira_vez:
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(3)

        # Checkbox "mostrar baixados"
        checkbox = page.locator('input[name="mostrarBaixados"]')
        if await checkbox.count() > 0 and not await checkbox.is_checked():
            await checkbox.check()

        # Preenche CPF e busca
        cpf_input = page.locator('input[name="cpf_cnpj"]')
        await cpf_input.click()
        await cpf_input.evaluate('el => el.value = ""')
        await cpf_input.fill(cpf_digits)
        await page.locator('input#enviar').click()
        await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        body_text = await page.evaluate('document.body.innerText')
        if 'partes encontradas' not in body_text.lower() and 'nome da parte' not in body_text.lower():
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        parte_links = page.locator('table a')
        if await parte_links.count() == 0:
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        await parte_links.first.click()
        await page.wait_for_load_state('networkidle', timeout=20000)
        await asyncio.sleep(2)

        rows_data = await page.evaluate('''
            () => {
                const rows = Array.from(document.querySelectorAll("table tr"));
                return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    const link = row.querySelector("a");
                    return {
                        col1: cells[0] ? cells[0].innerText.trim() : "",
                        col2: cells[1] ? cells[1].innerText.trim() : "",
                        href: link ? link.href : null
                    };
                }).filter(r => r.href && r.col1.length > 0);
            }
        ''')

        precat_rows = [r for r in rows_data if '(PRC)' in r['col1'] or '(PREC)' in r['col1']]
        precat_row  = precat_rows[-1] if precat_rows else None

        if not precat_row:
            resultado["status"] = "Sem PRECAT"
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        resultado["processo_precat"]    = precat_row['col1'].strip()
        resultado["processo_originario"] = clean_processo_originario(precat_row['col2'])
        resultado["precatorio_expedido"] = "Sim"

        todos_prc = page.locator('table a', has_text='PRC')
        total_prc = await todos_prc.count()
        prc_link  = todos_prc.nth(total_prc - 1) if total_prc > 0 else todos_prc.first
        await prc_link.click()
        await page.wait_for_load_state('networkidle', timeout=30000)
        await asyncio.sleep(2)

        mov_link = page.locator('a', has_text='Movimentação')
        if await mov_link.count() == 0:
            resultado["status"] = "Encontrado - Sem aba Movimentação"
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until='networkidle')
            await asyncio.sleep(2)
            return resultado

        await mov_link.click()
        await page.wait_for_load_state('networkidle', timeout=20000)
        await asyncio.sleep(2)

        mov_text  = await page.evaluate('document.body.innerText')
        mov_lower = mov_text.lower()
        mov_rows  = await page.evaluate('''
            () => Array.from(document.querySelectorAll("table tr"))
                .map(row => Array.from(row.querySelectorAll("td")).map(c => c.innerText.trim()))
                .filter(r => r.length >= 3)
        ''')

        proposta_encontrada = False
        for row in mov_rows:
            row_text = ' '.join(row).lower()
            if 'proposta orçamentária' in row_text and 'cjf' in row_text:
                proposta_encontrada = True
                complemento = row[-1] if row else ''
                anos = re.findall(r'\b(20\d{2})\b', complemento)
                if anos:
                    resultado["ano_orcamentario"] = anos[0]
                m2 = re.search(r'(\d{2}/\d{2}/\d{4})', row[0] if row else '')
                if m2:
                    resultado["data_proposta"] = m2.group(1)
                break

        if any(kw in mov_lower for kw in KEYWORDS_PAGO):
            resultado["status"] = "Pago/Excluir"
        elif proposta_encontrada:
            resultado["status"] = "Encontrado"
        else:
            resultado["status"] = "Encontrado - Sem proposta orçamentária"

        await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until='networkidle')
        await asyncio.sleep(2)

    except PlaywrightTimeout:
        resultado["status"] = "Timeout"
        try:
            await page.goto(URL_CPF_BUSCA, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except Exception:
            pass
    except Exception as e:
        resultado["status"] = "Erro"
        print(f"\n  ❌ Erro inesperado: {e}", flush=True)
        try:
            await page.goto(URL_CPF_BUSCA, timeout=15000, wait_until='domcontentloaded')
            await asyncio.sleep(2)
        except Exception:
            pass

    return resultado


async def main():
    # Verifica se o arquivo Excel existe
    if not Path(EXCEL_PATH).exists():
        print(f"\n❌ Arquivo não encontrado: {EXCEL_PATH}")
        print("   Certifique-se de que a planilha está na mesma pasta deste script.\n")
        sys.exit(1)

    preparar_planilha()
    pendentes = load_pendentes()
    total = len(pendentes)

    if total == 0:
        print("✅ Nenhum CPF pendente. Planilha já está completa!")
        sys.exit(0)

    print(f"\n{'='*65}")
    print(f"  CAPFI — Consulta TRF1 Precatórios")
    print(f"  Arquivo    : {EXCEL_PATH}")
    print(f"  Pendentes  : {total} CPFs")
    print(f"  Início     : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"  Checkpoint : {CHECKPOINT_PATH} (retoma automaticamente se interromper)")
    print(f"{'='*65}\n")
    print("  Pressione Ctrl+C a qualquer momento para pausar.")
    print("  Execute novamente para retomar de onde parou.\n")
    sys.stdout.flush()

    stats = {"encontrados": 0, "pagos": 0, "sem_precat": 0,
             "nao_encontrados": 0, "erros": 0}
    processados = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled"
            ]
        )
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 800},
            locale="pt-BR"
        )
        page = await context.new_page()
        await page.add_init_script(
            'Object.defineProperty(navigator,"webdriver",{get:()=>undefined})'
        )

        primeira_vez = True
        try:
            for idx, (row_idx, nome, cpf) in enumerate(pendentes):
                num = idx + 1
                print(f"[{num:04d}/{total}] {nome[:40]:<40} | {format_cpf(cpf)}", end=" | ", flush=True)

                dados = await consultar_cpf(page, cpf, primeira_vez)
                primeira_vez = False
                update_row(row_idx, dados)
                processados += 1
                salvar_checkpoint(processados, total)

                st    = dados["status"]
                emoji = (
                    "✅" if st == "Encontrado" else
                    "🔴" if "Pago" in st else
                    "⚪" if st in ("Sem PRECAT", "Não encontrado") else
                    "❌"
                )
                print(f"{emoji} {st}", flush=True)

                if st == "Encontrado":         stats["encontrados"] += 1
                elif "Pago" in st:             stats["pagos"] += 1
                elif st == "Sem PRECAT":       stats["sem_precat"] += 1
                elif st == "Não encontrado":   stats["nao_encontrados"] += 1
                else:                          stats["erros"] += 1

                await asyncio.sleep(1.5)

        except KeyboardInterrupt:
            print("\n\n⏸  Pausado pelo usuário. Execute novamente para continuar.")

        finally:
            await browser.close()

    # Remove checkpoint se tudo foi processado
    pendentes_restantes = load_pendentes()
    if len(pendentes_restantes) == 0 and Path(CHECKPOINT_PATH).exists():
        Path(CHECKPOINT_PATH).unlink()

    print(f"\n{'='*65}")
    print(f"  Resultado parcial/final — {datetime.now().strftime('%H:%M:%S')}")
    print(f"  ✅ Encontrado c/ proposta  : {stats['encontrados']}")
    print(f"  🔴 Pago/Excluir            : {stats['pagos']}")
    print(f"  ⚪ Sem PRECAT              : {stats['sem_precat']}")
    print(f"  ⚪ Não encontrado          : {stats['nao_encontrados']}")
    print(f"  ❌ Erros/Timeout           : {stats['erros']}")
    print(f"  📋 Pendentes restantes     : {len(pendentes_restantes)}")
    print(f"{'='*65}\n")


if __name__ == "__main__":
    asyncio.run(main())
