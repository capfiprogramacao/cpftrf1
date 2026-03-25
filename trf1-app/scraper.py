"""
scraper.py — Lógica de automação do TRF1
Consultada por CPF/CNPJ e extração de dados de precatórios
"""

import asyncio
import re
from datetime import datetime
from typing import Callable, Optional

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

# ─── Constantes ───────────────────────────────────────────────────────────────
URL_CPF_BUSCA = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"

KEYWORDS_PAGO = [
    "depósito", "deposito", "saque", "pagamento", "pago",
    "levantamento", "alvará", "alvar", "requisição de pequeno valor",
    "rpv levantado", "quitação", "quitado", "crédito levantado",
    "transferência", "ordem de pagamento expedida"
]

# ─── Utilitários ──────────────────────────────────────────────────────────────

def clean_cpf(cpf: str) -> str:
    return re.sub(r"\D", "", str(cpf).strip())

def format_cpf(cpf: str) -> str:
    d = clean_cpf(cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d) == 11 else cpf

def clean_processo_originario(texto: str) -> str:
    if texto and "/" in texto:
        return texto.split("/")[0].strip()
    return texto.strip() if texto else ""

# ─── Consulta individual ──────────────────────────────────────────────────────

async def consultar_cpf(page, cpf: str, primeira_vez: bool) -> dict:
    resultado = {
        "processo_precat": "",
        "processo_originario": "",
        "precatorio_expedido": "Não",
        "ano_orcamentario": "",
        "data_proposta": "",
        "status": "Não encontrado",
    }

    try:
        cpf_digits = clean_cpf(cpf)

        # ETAPA 1 — Navegar para busca (apenas na primeira vez)
        if primeira_vez:
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until="networkidle")
            await asyncio.sleep(3)

        # ETAPA 2 — Preencher formulário
        checkbox = page.locator('input[name="mostrarBaixados"]')
        if await checkbox.count() > 0 and not await checkbox.is_checked():
            await checkbox.check()

        cpf_input = page.locator('input[name="cpf_cnpj"]')
        await cpf_input.click()
        await cpf_input.evaluate("el => el.value = ''")
        await cpf_input.fill(cpf_digits)

        # ETAPA 3 — Pesquisar
        await page.locator("input#enviar").click()
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(2)

        body_text = await page.evaluate("document.body.innerText")
        if "partes encontradas" not in body_text.lower() and "nome da parte" not in body_text.lower():
            resultado["status"] = "Não encontrado"
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until="networkidle")
            await asyncio.sleep(2)
            return resultado

        # ETAPA 4 — Expandir lista de processos
        parte_links = page.locator("table a")
        if await parte_links.count() == 0:
            resultado["status"] = "Não encontrado"
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until="networkidle")
            await asyncio.sleep(2)
            return resultado

        await parte_links.first.click()
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(2)

        # ETAPA 5 — Identificar processo PRC (último = mais recente)
        rows_data = await page.evaluate("""
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
        """)

        precat_rows = [r for r in rows_data if "(PRC)" in r["col1"] or "(PREC)" in r["col1"]]
        precat_row = precat_rows[-1] if precat_rows else None

        if not precat_row:
            resultado["status"] = "Sem PRECAT"
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until="networkidle")
            await asyncio.sleep(2)
            return resultado

        resultado["processo_precat"] = precat_row["col1"].strip()
        resultado["processo_originario"] = clean_processo_originario(precat_row["col2"])
        resultado["precatorio_expedido"] = "Sim"

        # ETAPA 6 — Abrir processo PRC (último da lista)
        todos_prc = page.locator("table a", has_text="PRC")
        total_prc = await todos_prc.count()
        prc_link = todos_prc.nth(total_prc - 1) if total_prc > 0 else todos_prc.first
        await prc_link.click()
        await page.wait_for_load_state("networkidle", timeout=30000)
        await asyncio.sleep(2)

        # ETAPA 7 — Aba Movimentação
        mov_link = page.locator("a", has_text="Movimentação")
        if await mov_link.count() == 0:
            resultado["status"] = "Encontrado - Sem aba Movimentação"
            await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until="networkidle")
            await asyncio.sleep(2)
            return resultado

        await mov_link.click()
        await page.wait_for_load_state("networkidle", timeout=20000)
        await asyncio.sleep(2)

        # ETAPA 8 — Extrair dados da movimentação
        mov_text = await page.evaluate("document.body.innerText")
        mov_lower = mov_text.lower()

        mov_rows = await page.evaluate("""
            () => Array.from(document.querySelectorAll("table tr"))
                .map(row => Array.from(row.querySelectorAll("td")).map(c => c.innerText.trim()))
                .filter(r => r.length >= 3)
        """)

        proposta_encontrada = False
        for row in mov_rows:
            row_text = " ".join(row).lower()
            if "proposta orçamentária" in row_text and "cjf" in row_text:
                proposta_encontrada = True
                complemento = row[-1] if row else ""

                match_ano = re.match(r"(\d{4})", complemento.strip())
                if match_ano:
                    resultado["ano_orcamentario"] = match_ano.group(1)
                else:
                    anos = re.findall(r"\b(20\d{2})\b", complemento)
                    if anos:
                        resultado["ano_orcamentario"] = anos[0]

                match_data = re.search(r"data\s+(\d{2}/\d{2}/\d{4})", complemento, re.I)
                if match_data:
                    resultado["data_proposta"] = match_data.group(1)
                else:
                    data_col = row[0] if row else ""
                    m2 = re.search(r"(\d{2}/\d{2}/\d{4})", data_col)
                    if m2:
                        resultado["data_proposta"] = m2.group(1)
                break

        status_pago = any(kw in mov_lower for kw in KEYWORDS_PAGO)
        if status_pago:
            resultado["status"] = "Pago/Excluir"
        elif proposta_encontrada:
            resultado["status"] = "Encontrado"
        else:
            resultado["status"] = "Encontrado - Sem proposta orçamentária"

        await page.goto(URL_CPF_BUSCA, timeout=30000, wait_until="networkidle")
        await asyncio.sleep(2)

    except PlaywrightTimeout:
        resultado["status"] = "Timeout"
        try:
            await page.goto(URL_CPF_BUSCA, timeout=15000, wait_until="domcontentloaded")
            await asyncio.sleep(2)
        except Exception:
            pass

    except Exception as e:
        resultado["status"] = f"Erro: {str(e)[:60]}"
        try:
            await page.goto(URL_CPF_BUSCA, timeout=15000, wait_until="domcontentloaded")
            await asyncio.sleep(2)
        except Exception:
            pass

    return resultado


# ─── Processamento em lote (com checkpoint/retomada) ─────────────────────────

STATUS_REPROCESSAR = {"Erro", "Timeout", ""}

async def processar_lista(
    registros: list[dict],
    on_progresso: Optional[Callable] = None,
    checkpoint_path: Optional[str] = None,
) -> list[dict]:
    """
    Processa uma lista de {'nome': str, 'cpf': str}.
    - on_progresso(item): callback chamado após cada CPF processado
    - checkpoint_path: caminho de um arquivo JSON para salvar/retomar progresso
      Se o arquivo existir, CPFs já processados com sucesso serão ignorados.
    Retorna lista com os resultados completos (incluindo os já salvos no checkpoint).
    """
    import json as _json
    from pathlib import Path as _Path

    # ── Carregar checkpoint existente ─────────────────────────────────────────
    resultados_salvos: dict[str, dict] = {}  # cpf_limpo -> resultado
    if checkpoint_path and _Path(checkpoint_path).exists():
        try:
            with open(checkpoint_path, "r", encoding="utf-8") as f:
                salvos = _json.load(f)
            for item in salvos:
                cpf_key = re.sub(r"\D", "", str(item.get("cpf", "")))
                status  = item.get("status", "")
                if cpf_key and status and status not in STATUS_REPROCESSAR:
                    resultados_salvos[cpf_key] = item
        except Exception:
            pass  # Checkpoint corrompido: reprocessar tudo

    # ── Filtrar pendentes ─────────────────────────────────────────────────────
    pendentes = []
    for reg in registros:
        cpf_key = re.sub(r"\D", "", str(reg.get("cpf", "")))
        if cpf_key not in resultados_salvos:
            pendentes.append(reg)

    ja_feitos = len(resultados_salvos)
    total = len(registros)

    # ── Processar pendentes ───────────────────────────────────────────────────
    novos_resultados: list[dict] = []

    if pendentes:
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage",
                      "--disable-blink-features=AutomationControlled"],
            )
            context = await browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 800},
                locale="pt-BR",
            )
            page = await context.new_page()
            await page.add_init_script(
                'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
            )

            primeira_vez = True
            for reg in pendentes:
                nome = reg["nome"]
                cpf  = reg["cpf"]
                dados = await consultar_cpf(page, cpf, primeira_vez)
                primeira_vez = False

                item = {
                    "idx":                 ja_feitos + len(novos_resultados) + 1,
                    "nome":                nome,
                    "cpf":                 format_cpf(cpf),
                    "processo_precat":     dados["processo_precat"],
                    "processo_originario": dados["processo_originario"],
                    "precatorio_expedido": dados["precatorio_expedido"],
                    "ano_orcamentario":    dados["ano_orcamentario"],
                    "data_proposta":       dados["data_proposta"],
                    "status":              dados["status"],
                    "hora":                datetime.now().strftime("%H:%M:%S"),
                }
                novos_resultados.append(item)

                # Salvar checkpoint imediatamente após cada CPF
                if checkpoint_path:
                    todos_ate_agora = list(resultados_salvos.values()) + novos_resultados
                    try:
                        with open(checkpoint_path, "w", encoding="utf-8") as f:
                            _json.dump(todos_ate_agora, f, ensure_ascii=False, indent=2)
                    except Exception:
                        pass

                if on_progresso:
                    await on_progresso(item)

                await asyncio.sleep(1.5)

            await browser.close()

    # ── Retornar todos (salvos + novos) ───────────────────────────────────────
    return list(resultados_salvos.values()) + novos_resultados
