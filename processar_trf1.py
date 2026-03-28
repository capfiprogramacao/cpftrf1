"""
CAPFI — Consulta TRF1 Precatórios
Versão para execução local (Mac) — 3 workers paralelos

Como usar:
  1. Coloque este arquivo na mesma pasta da planilha Excel
  2. Execute: caffeinate -i python3 processar_trf1.py
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
EXCEL_PATH      = "Aposentados Sindireceita.xlsx"
CHECKPOINT_PATH = "checkpoint_trf1.json"
URL             = "https://processual.trf1.jus.br/consultaProcessual/cpfCnpjParte.php?secao=TRF1"
NUM_WORKERS     = 5   # navegadores simultâneos
CPFS_POR_WORKER = 999999  # sem limite — processa tudo de uma vez
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


def clean_cpf(cpf): return re.sub(r'\D','',str(cpf).strip())

def format_cpf(cpf):
    d = clean_cpf(cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d)==11 else cpf

def preparar_planilha():
    wb = load_workbook(EXCEL_PATH); ws = wb.active
    if ws.cell(1,3).value is None:
        for i, h in enumerate(["Processo PRECAT","Processo Originário","Precatório Expedido","Ano Orçamentário","Data Proposta","Status"], start=3):
            ws.cell(row=1, column=i).value = h
        for row in range(2, ws.max_row+1):
            cpf = ws.cell(row,2).value
            if not cpf or str(cpf).strip()=='':
                ws.cell(row,8).value = "Sem CPF"
            else:
                if len(re.sub(r'\D','',str(cpf)))!=11:
                    ws.cell(row,8).value = "CPF inválido"
        wb.save(EXCEL_PATH)
        print("✅ Planilha preparada.\n")

def load_pendentes():
    wb = load_workbook(EXCEL_PATH); ws = wb.active
    p = []
    for row in range(2, ws.max_row+1):
        nome=ws.cell(row,1).value; cpf=ws.cell(row,2).value
        st=str(ws.cell(row,8).value or "").strip()
        if not nome or not cpf: continue
        if len(re.sub(r'\D','',str(cpf)))!=11: continue
        if st not in STATUS_SKIP:
            p.append((row, str(nome).strip(), str(cpf).strip()))
    return p

excel_lock = asyncio.Lock()

async def save_result(row_idx, dados):
    async with excel_lock:
        wb = load_workbook(EXCEL_PATH); ws = wb.active
        ws.cell(row_idx,3).value = dados.get("processo_precat","")
        ws.cell(row_idx,4).value = dados.get("processo_originario","")
        ws.cell(row_idx,5).value = dados.get("precatorio_expedido","")
        ws.cell(row_idx,6).value = dados.get("ano_orcamentario","")
        ws.cell(row_idx,7).value = dados.get("data_proposta","")
        ws.cell(row_idx,8).value = dados.get("status","")
        wb.save(EXCEL_PATH)

async def consultar_cpf(page, cpf, pv):
    res={"processo_precat":"","processo_originario":"","precatorio_expedido":"Não",
         "ano_orcamentario":"","data_proposta":"","status":"Não encontrado"}
    try:
        if pv:
            await page.goto(URL,timeout=30000,wait_until='networkidle'); await asyncio.sleep(3)
        cb=page.locator('input[name="mostrarBaixados"]')
        if await cb.count()>0 and not await cb.is_checked(): await cb.check()
        inp=page.locator('input[name="cpf_cnpj"]')
        await inp.click(); await inp.evaluate('el=>el.value=""'); await inp.fill(clean_cpf(cpf))
        await page.locator('input#enviar').click()
        await page.wait_for_load_state('networkidle',timeout=30000); await asyncio.sleep(2)
        body=await page.evaluate('document.body.innerText')
        if 'partes encontradas' not in body.lower() and 'nome da parte' not in body.lower():
            await page.goto(URL,timeout=30000,wait_until='networkidle'); await asyncio.sleep(2); return res
        links=page.locator('table a')
        if await links.count()==0:
            await page.goto(URL,timeout=30000,wait_until='networkidle'); await asyncio.sleep(2); return res
        await links.first.click()
        await page.wait_for_load_state('networkidle',timeout=20000); await asyncio.sleep(2)
        rows_data=await page.evaluate('''()=>{return Array.from(document.querySelectorAll("table tr")).map(row=>{
            const c=Array.from(row.querySelectorAll("td"));const l=row.querySelector("a");
            return{col1:c[0]?c[0].innerText.trim():"",col2:c[1]?c[1].innerText.trim():"",href:l?l.href:null};
        }).filter(r=>r.href&&r.col1.length>0);}''')
        precat_rows=[r for r in rows_data if '(PRC)' in r['col1'] or '(PREC)' in r['col1']]
        if not precat_rows:
            res["status"]="Sem PRECAT"; await page.goto(URL,timeout=30000,wait_until='networkidle'); await asyncio.sleep(2); return res
        pr=precat_rows[-1]
        res["processo_precat"]=pr['col1'].strip()
        res["processo_originario"]=pr['col2'].split('/')[0].strip() if '/' in pr['col2'] else pr['col2'].strip()
        res["precatorio_expedido"]="Sim"
        todos=page.locator('table a',has_text='PRC'); n=await todos.count()
        await (todos.nth(n-1) if n>0 else todos.first).click()
        await page.wait_for_load_state('networkidle',timeout=30000); await asyncio.sleep(2)
        mov=page.locator('a',has_text='Movimentação')
        if await mov.count()==0:
            res["status"]="Encontrado - Sem aba Movimentação"; await page.goto(URL,timeout=30000,wait_until='networkidle'); await asyncio.sleep(2); return res
        await mov.click()
        await page.wait_for_load_state('networkidle',timeout=20000); await asyncio.sleep(2)
        ml=(await page.evaluate('document.body.innerText')).lower()
        mr=await page.evaluate('''()=>Array.from(document.querySelectorAll("table tr"))
            .map(row=>Array.from(row.querySelectorAll("td")).map(c=>c.innerText.trim())).filter(r=>r.length>=3)''')
        prop=False
        for row in mr:
            if 'proposta orçamentária' in ' '.join(row).lower() and 'cjf' in ' '.join(row).lower():
                prop=True; comp=row[-1] if row else ''
                anos=re.findall(r'\b(20\d{2})\b',comp)
                if anos: res["ano_orcamentario"]=anos[0]
                m=re.search(r'(\d{2}/\d{2}/\d{4})',row[0] if row else '')
                if m: res["data_proposta"]=m.group(1)
                break
        if any(kw in ml for kw in KEYWORDS_PAGO): res["status"]="Pago/Excluir"
        elif prop: res["status"]="Encontrado"
        else: res["status"]="Encontrado - Sem proposta orçamentária"
        await page.goto(URL,timeout=30000,wait_until='networkidle'); await asyncio.sleep(2)
    except PlaywrightTimeout:
        res["status"]="Timeout"
        try: await page.goto(URL,timeout=15000,wait_until='domcontentloaded'); await asyncio.sleep(2)
        except: pass
    except Exception as e:
        res["status"]="Erro"; print(f"\n  ❌ {e}",flush=True)
        try: await page.goto(URL,timeout=15000,wait_until='domcontentloaded'); await asyncio.sleep(2)
        except: pass
    return res

async def worker(pw, wid, fatia, fila):
    browser=await pw.chromium.launch(headless=True,
        args=["--no-sandbox","--disable-dev-shm-usage","--disable-blink-features=AutomationControlled"])
    ctx=await browser.new_context(
        user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        viewport={"width":1280,"height":800}, locale="pt-BR")
    page=await ctx.new_page()
    await page.add_init_script('Object.defineProperty(navigator,"webdriver",{get:()=>undefined})')
    pv=True
    for row_idx,nome,cpf in fatia:
        dados=await consultar_cpf(page,cpf,pv); pv=False
        await save_result(row_idx,dados)
        await fila.put((wid,nome,cpf,dados["status"]))
        await asyncio.sleep(1.5)
    await browser.close()

async def coletor(fila, total):
    stats={"Encontrado":0,"Pago/Excluir":0,"Sem PRECAT":0,"Não encontrado":0,"erros":0}
    for i in range(total):
        wid,nome,cpf,st=await asyncio.wait_for(fila.get(),timeout=300)
        emoji="✅" if st=="Encontrado" else "🔴" if "Pago" in st else "⚪" if st in("Sem PRECAT","Não encontrado") else "❌"
        print(f"[W{wid}][{i+1:04d}/{total}] {nome[:38]:<38} {format_cpf(cpf)} → {emoji} {st}",flush=True)
        if st=="Encontrado": stats["Encontrado"]+=1
        elif "Pago" in st: stats["Pago/Excluir"]+=1
        elif st=="Sem PRECAT": stats["Sem PRECAT"]+=1
        elif st=="Não encontrado": stats["Não encontrado"]+=1
        else: stats["erros"]+=1
    return stats

async def main():
    if not Path(EXCEL_PATH).exists():
        print(f"\n❌ Arquivo não encontrado: {EXCEL_PATH}")
        print("   Coloque a planilha na mesma pasta deste script.\n")
        sys.exit(1)

    preparar_planilha()
    pendentes=load_pendentes()
    total=len(pendentes)

    if total==0:
        print("✅ Nenhum CPF pendente. Planilha já completa!"); sys.exit(0)

    # Divide em fatias para cada worker
    fatias=[pendentes[i::NUM_WORKERS] for i in range(NUM_WORKERS)]

    print(f"\n{'='*65}")
    print(f"  CAPFI — Consulta TRF1 | {EXCEL_PATH}")
    print(f"  Workers  : {NUM_WORKERS} navegadores em paralelo")
    print(f"  Pendentes: {total} CPFs")
    print(f"  Início   : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"  Pressione Ctrl+C para pausar (retoma de onde parou)")
    print(f"{'='*65}\n")
    sys.stdout.flush()

    fila=asyncio.Queue()
    try:
        async with async_playwright() as pw:
            tasks=[asyncio.create_task(worker(pw,i+1,fatias[i],fila)) for i in range(NUM_WORKERS)]
            tasks.append(asyncio.create_task(coletor(fila,total)))
            resultados=await asyncio.gather(*tasks)
    except KeyboardInterrupt:
        print("\n\n⏸  Pausado. Execute novamente para continuar.")
        return

    stats=resultados[-1]
    restantes=len(load_pendentes())

    print(f"\n{'='*65}")
    print(f"  ✅ Concluído! {datetime.now().strftime('%H:%M:%S')}")
    print(f"  Encontrados : {stats['Encontrado']}")
    print(f"  Pago/Excluir: {stats['Pago/Excluir']}")
    print(f"  Sem PRECAT  : {stats['Sem PRECAT']}")
    print(f"  Não enc.    : {stats['Não encontrado']}")
    print(f"  Erros       : {stats['erros']}")
    print(f"  Pendentes   : {restantes}")
    if restantes==0: print(f"\n  🎉 PROCESSAMENTO 100% COMPLETO!")
    print(f"{'='*65}\n")

if __name__=="__main__":
    asyncio.run(main())
