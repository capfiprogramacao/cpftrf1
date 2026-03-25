-- ─── CAPFI System — Schema PostgreSQL ────────────────────────────────────────
-- Banco de dados central para todos os módulos do sistema

-- ── Extensões ─────────────────────────────────────────────────────────────────
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- ── Banco do n8n (criado separado para isolamento) ────────────────────────────
CREATE DATABASE n8n_db;

-- ── Tabela: pessoas ───────────────────────────────────────────────────────────
-- Armazena todos os leads/clientes consultados
CREATE TABLE IF NOT EXISTS pessoas (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    nome            VARCHAR(255) NOT NULL,
    cpf             VARCHAR(14) UNIQUE NOT NULL,
    email           VARCHAR(255),
    telefone        VARCHAR(20),
    nectar_id       VARCHAR(50),          -- ID do contato no Néctar CRM
    criado_em       TIMESTAMP DEFAULT NOW(),
    atualizado_em   TIMESTAMP DEFAULT NOW()
);

-- ── Tabela: lotes ─────────────────────────────────────────────────────────────
-- Representa cada upload de planilha / lote de processamento
CREATE TABLE IF NOT EXISTS lotes (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    nome_arquivo    VARCHAR(255),
    total           INTEGER DEFAULT 0,
    processados     INTEGER DEFAULT 0,
    encontrados     INTEGER DEFAULT 0,
    pagos           INTEGER DEFAULT 0,
    nao_encontrados INTEGER DEFAULT 0,
    erros           INTEGER DEFAULT 0,
    status          VARCHAR(50) DEFAULT 'aguardando',  -- aguardando, processando, concluido, erro
    criado_em       TIMESTAMP DEFAULT NOW(),
    concluido_em    TIMESTAMP
);

-- ── Tabela: consultas_trf1 ────────────────────────────────────────────────────
-- Resultados das consultas de precatórios no TRF1
CREATE TABLE IF NOT EXISTS consultas_trf1 (
    id                   UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    pessoa_id            UUID REFERENCES pessoas(id) ON DELETE CASCADE,
    lote_id              UUID REFERENCES lotes(id) ON DELETE SET NULL,
    processo_precat      VARCHAR(100),
    processo_originario  VARCHAR(100),
    precatorio_expedido  VARCHAR(10) DEFAULT 'Não',
    ano_orcamentario     VARCHAR(10),
    data_proposta        VARCHAR(20),
    status               VARCHAR(50) NOT NULL,    -- Encontrado, Pago/Excluir, Sem PRECAT, Não encontrado, Erro
    consultado_em        TIMESTAMP DEFAULT NOW(),
    UNIQUE(pessoa_id, processo_precat)             -- Evita duplicatas
);

-- ── Tabela: consultas_datajud ─────────────────────────────────────────────────
-- Resultados das consultas via API DataJud (CNJ)
CREATE TABLE IF NOT EXISTS consultas_datajud (
    id                   UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    pessoa_id            UUID REFERENCES pessoas(id) ON DELETE CASCADE,
    numero_processo      VARCHAR(30),
    classe_nome          VARCHAR(100),
    classe_codigo        INTEGER,
    tribunal             VARCHAR(20) DEFAULT 'TRF1',
    grau                 VARCHAR(10),
    data_ajuizamento     VARCHAR(30),
    ultima_atualizacao   TIMESTAMP,
    orgao_julgador       VARCHAR(200),
    movimentos_json      JSONB,                   -- Movimentações completas em JSON
    consultado_em        TIMESTAMP DEFAULT NOW()
);

-- ── Tabela: analises_ia ───────────────────────────────────────────────────────
-- Análises geradas pelo Llama/Claude sobre os processos
CREATE TABLE IF NOT EXISTS analises_ia (
    id                   UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    pessoa_id            UUID REFERENCES pessoas(id) ON DELETE CASCADE,
    numero_processo      VARCHAR(30),
    modelo_usado         VARCHAR(50),             -- llama3.2, claude-sonnet-4-6, etc.
    resumo               TEXT,
    classificacao        VARCHAR(50),             -- Alto valor, Médio, Baixo, Descartado
    pontos_favoraveis    TEXT,
    pontos_desfavoraveis TEXT,
    valor_estimado       DECIMAL(15,2),
    recomendacao         TEXT,
    tokens_usados        INTEGER,
    custo_usd            DECIMAL(10,6),
    analisado_em         TIMESTAMP DEFAULT NOW()
);

-- ── Tabela: integracoes_nectar ────────────────────────────────────────────────
-- Controle das sincronizações com o Néctar CRM
CREATE TABLE IF NOT EXISTS integracoes_nectar (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    pessoa_id       UUID REFERENCES pessoas(id) ON DELETE CASCADE,
    nectar_id       VARCHAR(50),
    tipo_acao       VARCHAR(50),                  -- criou_contato, criou_oportunidade, atualizou
    payload_json    JSONB,
    resposta_json   JSONB,
    sucesso         BOOLEAN DEFAULT FALSE,
    executado_em    TIMESTAMP DEFAULT NOW()
);

-- ── Índices para performance ───────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_pessoas_cpf ON pessoas(cpf);
CREATE INDEX IF NOT EXISTS idx_consultas_trf1_pessoa ON consultas_trf1(pessoa_id);
CREATE INDEX IF NOT EXISTS idx_consultas_trf1_status ON consultas_trf1(status);
CREATE INDEX IF NOT EXISTS idx_consultas_trf1_lote ON consultas_trf1(lote_id);
CREATE INDEX IF NOT EXISTS idx_analises_pessoa ON analises_ia(pessoa_id);
CREATE INDEX IF NOT EXISTS idx_lotes_status ON lotes(status);

-- ── View: dashboard resumo ────────────────────────────────────────────────────
CREATE OR REPLACE VIEW dashboard_resumo AS
SELECT
    (SELECT COUNT(*) FROM pessoas) AS total_pessoas,
    (SELECT COUNT(*) FROM consultas_trf1 WHERE status = 'Encontrado') AS encontrados,
    (SELECT COUNT(*) FROM consultas_trf1 WHERE status = 'Pago/Excluir') AS pagos,
    (SELECT COUNT(*) FROM consultas_trf1 WHERE status = 'Sem PRECAT') AS sem_precat,
    (SELECT COUNT(*) FROM consultas_trf1 WHERE status = 'Não encontrado') AS nao_encontrados,
    (SELECT COUNT(*) FROM analises_ia) AS total_analisados,
    (SELECT COUNT(*) FROM integracoes_nectar WHERE sucesso = TRUE) AS enviados_nectar;
