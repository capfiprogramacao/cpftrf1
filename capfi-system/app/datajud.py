"""
datajud.py — Módulo de consulta à API pública DataJud (CNJ)
Complementa o scraping do TRF1 com dados estruturados via API REST
"""

import asyncio
import re
from datetime import datetime
from typing import Optional
import httpx

# ─── Configuração ──────────────────────────────────────────────────────────────
DATAJUD_BASE_URL = "https://api-publica.datajud.cnj.jus.br"
DATAJUD_API_KEY  = "cDZHYzlZa0JadVREZDJCendQbXY6SkJlTzNjLV9TRENyQk1RdnFKZGRQdw=="

ENDPOINTS = {
    "TRF1": f"{DATAJUD_BASE_URL}/api_publica_trf1/_search",
    "TRF2": f"{DATAJUD_BASE_URL}/api_publica_trf2/_search",
    "TRF3": f"{DATAJUD_BASE_URL}/api_publica_trf3/_search",
    "STJ":  f"{DATAJUD_BASE_URL}/api_publica_stj/_search",
    "STF":  f"{DATAJUD_BASE_URL}/api_publica_stf/_search",
}

HEADERS = {
    "Authorization": f"APIKey {DATAJUD_API_KEY}",
    "Content-Type": "application/json",
}

# ─── Consultar processo por número ────────────────────────────────────────────

async def consultar_por_numero(
    numero_processo: str,
    tribunal: str = "TRF1"
) -> Optional[dict]:
    """
    Busca um processo pelo número no DataJud.
    Retorna os dados estruturados ou None se não encontrado.
    """
    # Limpar número do processo (remover pontos, traços)
    numero_limpo = re.sub(r"\D", "", numero_processo)

    endpoint = ENDPOINTS.get(tribunal.upper(), ENDPOINTS["TRF1"])

    query = {
        "query": {
            "term": {"numeroProcesso": numero_limpo}
        },
        "size": 1,
        "_source": [
            "numeroProcesso", "classe", "tribunal", "grau",
            "dataAjuizamento", "dataHoraUltimaAtualizacao",
            "orgaoJulgador", "assuntos", "movimentos"
        ]
    }

    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(endpoint, json=query, headers=HEADERS)
            resp.raise_for_status()
            data = resp.json()

            hits = data.get("hits", {}).get("hits", [])
            if not hits:
                return None

            source = hits[0]["_source"]
            return _processar_resultado(source)

    except Exception as e:
        print(f"  ⚠️  DataJud erro [{numero_processo}]: {e}")
        return None


def _processar_resultado(source: dict) -> dict:
    """Normaliza o resultado do DataJud para o formato padrão do sistema."""

    # Extrair movimentações mais recentes (últimas 10)
    movimentos = source.get("movimentos", [])
    movimentos_resumo = []
    for mov in movimentos[-10:]:
        movimentos_resumo.append({
            "codigo": mov.get("codigo"),
            "nome":   mov.get("nome"),
            "data":   mov.get("dataHora", "")[:10],
            "orgao":  mov.get("orgaoJulgador", {}).get("nome", ""),
        })

    # Data de ajuizamento formatada
    data_raw = source.get("dataAjuizamento", "")
    data_fmt = ""
    if data_raw and len(data_raw) >= 8:
        try:
            dt = datetime.strptime(data_raw[:8], "%Y%m%d")
            data_fmt = dt.strftime("%d/%m/%Y")
        except Exception:
            data_fmt = data_raw[:10]

    return {
        "numero_processo":    source.get("numeroProcesso", ""),
        "tribunal":           source.get("tribunal", ""),
        "grau":               source.get("grau", ""),
        "classe_codigo":      source.get("classe", {}).get("codigo"),
        "classe_nome":        source.get("classe", {}).get("nome", ""),
        "data_ajuizamento":   data_fmt,
        "ultima_atualizacao": source.get("dataHoraUltimaAtualizacao", "")[:19],
        "orgao_julgador":     source.get("orgaoJulgador", {}).get("nome", ""),
        "assuntos":           [a.get("nome") for a in source.get("assuntos", [])],
        "movimentos":         movimentos_resumo,
        "total_movimentos":   len(movimentos),
    }


# ─── Consultar múltiplos processos ────────────────────────────────────────────

async def consultar_lista_processos(
    numeros: list[str],
    tribunal: str = "TRF1",
    delay: float = 0.5
) -> list[dict]:
    """
    Consulta uma lista de processos no DataJud com delay entre requisições
    para evitar sobrecarga da API pública.
    """
    resultados = []
    for numero in numeros:
        resultado = await consultar_por_numero(numero, tribunal)
        if resultado:
            resultados.append(resultado)
        await asyncio.sleep(delay)
    return resultados


# ─── Enriquecer resultado do TRF1 com dados do DataJud ────────────────────────

async def enriquecer_com_datajud(resultado_trf1: dict) -> dict:
    """
    Recebe um resultado do scraping TRF1 e complementa com dados do DataJud.
    Consulta tanto o processo PRECAT quanto o processo originário.
    """
    enriquecido = resultado_trf1.copy()
    enriquecido["datajud"] = {}

    # Consultar processo originário (mais rico em movimentações)
    processo_orig = resultado_trf1.get("processo_originario", "")
    if processo_orig:
        dados_dj = await consultar_por_numero(processo_orig)
        if dados_dj:
            enriquecido["datajud"]["processo_originario"] = dados_dj

    # Consultar processo PRECAT
    processo_precat = resultado_trf1.get("processo_precat", "")
    if processo_precat:
        # Extrair só os dígitos do processo PRECAT (remover classe e seção)
        nums = re.findall(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}", processo_precat)
        if nums:
            dados_prc = await consultar_por_numero(nums[0])
            if dados_prc:
                enriquecido["datajud"]["processo_precat"] = dados_prc

    return enriquecido
