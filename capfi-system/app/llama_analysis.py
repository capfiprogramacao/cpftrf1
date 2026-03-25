"""
llama_analysis.py — Análise de processos com Llama (local) e Claude (complexos)
Estratégia em camadas: Llama para triagem e análise geral, Claude para casos críticos
"""

import os
import json
from datetime import datetime
from typing import Optional
import httpx

# ─── Configuração ──────────────────────────────────────────────────────────────
OLLAMA_URL    = os.getenv("OLLAMA_URL", "http://ollama:11434")
CLAUDE_KEY    = os.getenv("CLAUDE_API_KEY", "")
MODELO_LLAMA  = "llama3.2"          # Roda local — custo zero
MODELO_CLAUDE = "claude-haiku-4-5-20251001"  # Usado apenas para casos complexos


# ─── Camada 1: Triagem rápida (Llama) ─────────────────────────────────────────

async def triagem_processo(dados_processo: dict) -> dict:
    """
    Triagem rápida para decidir se vale analisar o processo com profundidade.
    Usa Llama local — custo zero, muito rápido.
    Retorna: { vale_pena: bool, motivo: str, classificacao: str }
    """
    resumo = _montar_resumo_processo(dados_processo)

    prompt = f"""Você é um assistente jurídico especializado em precatórios.
Analise brevemente o processo abaixo e responda em JSON:

PROCESSO:
{resumo}

Responda APENAS com JSON válido, sem explicações:
{{
  "vale_pena": true ou false,
  "classificacao": "Alto valor" | "Médio valor" | "Baixo valor" | "Descartado",
  "motivo": "explicação em 1 frase"
}}

Critérios:
- vale_pena = true se: tem precatório expedido E não está pago/encerrado
- Alto valor: processo com sinais de valor significativo
- Descartado: pago, levantado, encerrado ou sem precatório"""

    try:
        resultado = await _chamar_llama(prompt, max_tokens=200)
        # Extrair JSON da resposta
        inicio = resultado.find("{")
        fim = resultado.rfind("}") + 1
        if inicio >= 0 and fim > inicio:
            return json.loads(resultado[inicio:fim])
    except Exception as e:
        print(f"  ⚠️  Triagem falhou: {e}")

    # Fallback baseado no status do TRF1
    status = dados_processo.get("status", "")
    vale = "Encontrado" in status and "Pago" not in status
    return {
        "vale_pena": vale,
        "classificacao": "Médio valor" if vale else "Descartado",
        "motivo": f"Baseado no status: {status}"
    }


# ─── Camada 2: Análise completa (Llama) ───────────────────────────────────────

async def analisar_processo_llama(dados_processo: dict) -> dict:
    """
    Análise completa do processo usando Llama local.
    Extrai informações relevantes e gera resumo estruturado.
    Custo: zero (roda no servidor VPS)
    """
    resumo = _montar_resumo_processo(dados_processo)

    prompt = f"""Você é um assistente jurídico especializado em precatórios federais.
Analise o processo abaixo e extraia as informações relevantes.

PROCESSO:
{resumo}

Responda APENAS com JSON válido:
{{
  "resumo": "Resumo do caso em 2-3 frases",
  "pontos_favoraveis": ["ponto 1", "ponto 2"],
  "pontos_desfavoraveis": ["ponto 1", "ponto 2"],
  "valor_estimado_descricao": "Descrição do valor potencial",
  "prazo_estimado": "Estimativa de prazo para recebimento",
  "recomendacao": "Ação recomendada em 1 frase",
  "urgencia": "Alta" | "Média" | "Baixa"
}}"""

    try:
        resultado = await _chamar_llama(prompt, max_tokens=600)
        inicio = resultado.find("{")
        fim = resultado.rfind("}") + 1
        if inicio >= 0 and fim > inicio:
            analise = json.loads(resultado[inicio:fim])
            analise["modelo_usado"] = MODELO_LLAMA
            analise["custo_usd"] = 0.0
            analise["analisado_em"] = datetime.now().isoformat()
            return analise
    except Exception as e:
        print(f"  ⚠️  Análise Llama falhou: {e}")

    return _analise_fallback(dados_processo, MODELO_LLAMA)


# ─── Camada 3: Análise profunda (Claude) — apenas para casos críticos ─────────

async def analisar_processo_claude(dados_processo: dict) -> dict:
    """
    Análise aprofundada usando Claude API.
    Usar apenas para casos de alto valor ou complexidade.
    Custo estimado: R$ 0.05-0.20 por análise.
    """
    if not CLAUDE_KEY:
        print("  ⚠️  CLAUDE_API_KEY não configurada — usando Llama")
        return await analisar_processo_llama(dados_processo)

    resumo = _montar_resumo_processo(dados_processo)

    prompt = f"""Você é um especialista em direito tributário e precatórios federais no Brasil.
Faça uma análise completa e estratégica do processo abaixo.

PROCESSO:
{resumo}

Forneça uma análise detalhada em JSON:
{{
  "resumo": "Resumo executivo do caso",
  "situacao_atual": "Situação atual do precatório",
  "pontos_favoraveis": ["lista detalhada"],
  "pontos_desfavoraveis": ["lista detalhada"],
  "valor_estimado_descricao": "Análise do potencial de valor",
  "prazo_estimado": "Estimativa fundamentada de prazo",
  "riscos": ["principais riscos identificados"],
  "recomendacao": "Recomendação estratégica detalhada",
  "urgencia": "Alta" | "Média" | "Baixa",
  "proximo_passo": "Ação imediata recomendada"
}}"""

    try:
        async with httpx.AsyncClient(timeout=60) as client:
            resp = await client.post(
                "https://api.anthropic.com/v1/messages",
                headers={
                    "x-api-key": CLAUDE_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json",
                },
                json={
                    "model": MODELO_CLAUDE,
                    "max_tokens": 1024,
                    "messages": [{"role": "user", "content": prompt}]
                }
            )
            resp.raise_for_status()
            data = resp.json()

            texto = data["content"][0]["text"]
            tokens = data.get("usage", {})
            custo = (tokens.get("input_tokens", 0) * 0.00000025 +
                     tokens.get("output_tokens", 0) * 0.00000125)

            inicio = texto.find("{")
            fim = texto.rfind("}") + 1
            if inicio >= 0 and fim > inicio:
                analise = json.loads(texto[inicio:fim])
                analise["modelo_usado"] = MODELO_CLAUDE
                analise["tokens_usados"] = tokens.get("input_tokens", 0) + tokens.get("output_tokens", 0)
                analise["custo_usd"] = round(custo, 6)
                analise["analisado_em"] = datetime.now().isoformat()
                return analise

    except Exception as e:
        print(f"  ⚠️  Claude falhou, usando Llama: {e}")

    return await analisar_processo_llama(dados_processo)


# ─── Pipeline completo: triagem + análise automática ──────────────────────────

async def pipeline_analise(dados_processo: dict, forcar_claude: bool = False) -> dict:
    """
    Pipeline inteligente de análise:
    1. Triagem com Llama (sempre)
    2. Se vale_pena=True → Análise completa com Llama
    3. Se forcar_claude=True ou classificacao='Alto valor' → Claude também
    """
    # Etapa 1: Triagem
    triagem = await triagem_processo(dados_processo)

    if not triagem.get("vale_pena", False):
        return {
            "triagem": triagem,
            "analise": None,
            "modelo_usado": MODELO_LLAMA,
            "custo_usd": 0.0,
            "descartado": True,
            "motivo_descarte": triagem.get("motivo")
        }

    # Etapa 2: Análise com Llama
    analise = await analisar_processo_llama(dados_processo)

    # Etapa 3: Análise com Claude se necessário
    if forcar_claude or triagem.get("classificacao") == "Alto valor":
        analise_claude = await analisar_processo_claude(dados_processo)
        analise["analise_claude"] = analise_claude

    return {
        "triagem": triagem,
        "analise": analise,
        "modelo_usado": analise.get("modelo_usado", MODELO_LLAMA),
        "custo_usd": analise.get("custo_usd", 0.0),
        "descartado": False,
    }


# ─── Funções auxiliares ────────────────────────────────────────────────────────

def _montar_resumo_processo(dados: dict) -> str:
    """Monta texto resumido do processo para enviar à IA."""
    linhas = []
    if dados.get("nome"):
        linhas.append(f"Parte: {dados['nome']}")
    if dados.get("processo_precat"):
        linhas.append(f"Processo PRECAT: {dados['processo_precat']}")
    if dados.get("processo_originario"):
        linhas.append(f"Processo Originário: {dados['processo_originario']}")
    if dados.get("precatorio_expedido"):
        linhas.append(f"Precatório Expedido: {dados['precatorio_expedido']}")
    if dados.get("ano_orcamentario"):
        linhas.append(f"Ano Orçamentário: {dados['ano_orcamentario']}")
    if dados.get("data_proposta"):
        linhas.append(f"Data da Proposta: {dados['data_proposta']}")
    if dados.get("status"):
        linhas.append(f"Status TRF1: {dados['status']}")

    # Dados do DataJud se disponíveis
    dj = dados.get("datajud", {})
    if dj.get("processo_originario"):
        orig = dj["processo_originario"]
        linhas.append(f"Classe: {orig.get('classe_nome', '')}")
        linhas.append(f"Ajuizado em: {orig.get('data_ajuizamento', '')}")
        linhas.append(f"Órgão Julgador: {orig.get('orgao_julgador', '')}")
        movs = orig.get("movimentos", [])
        if movs:
            linhas.append("Últimas movimentações:")
            for m in movs[-5:]:
                linhas.append(f"  - {m.get('data', '')} | {m.get('nome', '')}")

    return "\n".join(linhas)


async def _chamar_llama(prompt: str, max_tokens: int = 500) -> str:
    """Chama o Ollama local com o modelo Llama."""
    async with httpx.AsyncClient(timeout=120) as client:
        resp = await client.post(
            f"{OLLAMA_URL}/api/generate",
            json={
                "model": MODELO_LLAMA,
                "prompt": prompt,
                "stream": False,
                "options": {"num_predict": max_tokens, "temperature": 0.1}
            }
        )
        resp.raise_for_status()
        return resp.json().get("response", "")


def _analise_fallback(dados: dict, modelo: str) -> dict:
    """Análise básica quando a IA falha."""
    status = dados.get("status", "")
    return {
        "resumo": f"Processo com status: {status}",
        "pontos_favoraveis": ["Precatório identificado"] if "Encontrado" in status else [],
        "pontos_desfavoraveis": ["Análise IA indisponível"],
        "recomendacao": "Verificar manualmente",
        "urgencia": "Média",
        "modelo_usado": modelo,
        "custo_usd": 0.0,
        "analisado_em": datetime.now().isoformat()
    }
