"""
Reseta a coluna STATUS_CONSULTA do SINDIRECEITA_COMPLETO.xlsx
para que o script consultar_sindireceita_trf1.py reprocesse tudo.
"""
from openpyxl import load_workbook
from pathlib import Path

EXCEL_PATH = "SINDIRECEITA_COMPLETO.xlsx"
COL_STATUS = 12

excel = Path(EXCEL_PATH)
if not excel.exists():
    print(f"❌ Arquivo não encontrado: {EXCEL_PATH}")
    print(f"   Pasta atual: {Path.cwd()}")
    exit(1)

wb = load_workbook(EXCEL_PATH)
ws = wb.active

count = 0
for row in range(2, ws.max_row + 1):
    if ws.cell(row, COL_STATUS).value is not None:
        ws.cell(row, COL_STATUS).value = None
        count += 1

wb.save(EXCEL_PATH)
print(f"✅ Status resetado em {count} linhas. Pode rodar o script principal agora.")
