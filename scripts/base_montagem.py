"""Organiza a base de materiais baixados dos painéis em montagem (uma aba, ordenada por OS)."""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from familias import familia

COLS = ["OS", "Documento", "Nº Solicitação", "Cod. Material", "Material", "Quantidade", "Quantidade Baixada",
        "Necessidade", "Status Necessidade", "Data", "Cód. Aplicação", "Cód. Localização", "Situação"]
LARG = [8, 11, 12, 12, 70, 11, 12, 11, 28, 17, 11, 11, 11]

def gera(entradas, saida):
    d = pd.concat([pd.read_excel(p) for p in entradas], ignore_index=True)
    falta = [c for c in COLS if c not in d.columns]
    if falta: raise SystemExit(f"Colunas ausentes: {falta}")
    d = d[COLS].drop_duplicates().sort_values(["OS", "Documento", "Material"]).reset_index(drop=True)
    wb = Workbook(); ws = wb.active; ws.title = "Base"
    F = "Arial"
    for j, (h, w) in enumerate(zip(COLS, LARG), 1):
        x = ws.cell(1, j, h); x.font = Font(name=F, bold=True, color="FFFFFF")
        x.fill = PatternFill("solid", fgColor="1C5CAB"); x.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[L(j)].width = w
    ws.row_dimensions[1].height = 30; ws.freeze_panes = "A2"
    for i, row in enumerate(d.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if pd.isna(v): continue
            if j == 10: v = v.to_pydatetime()
            elif isinstance(v, float) and v.is_integer(): v = int(v)
            ws.cell(i, j, v).font = Font(name=F, size=10)
        ws.cell(i, 10).number_format = "dd/mm/yyyy"
        for j in (6, 7): ws.cell(i, j).number_format = "#,##0.##"
    t = Table(displayName="BaseMontagem", ref=f"A1:{L(len(COLS))}{len(d)+1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True); ws.add_table(t)
    materiais(wb, d)
    wb.save(saida)
    return d

def materiais(wb, d):
    """Aba com os materiais únicos (código MEGA) para conferir o cadastro no SAP."""
    m = d.groupby("Cod. Material").agg(Material=("Material", "first"), n_os=("OS", "nunique"),
            qtd=("Quantidade Baixada", "sum"), lista=("OS", lambda s: ", ".join(str(x) for x in sorted(s.unique()))))
    m["Família"] = m.Material.map(familia)
    m = m.reset_index().sort_values(["n_os", "qtd"], ascending=False)
    ws = wb.create_sheet("Materiais únicos")
    cols = ["Cod. Material (MEGA)", "Material", "Família", "Nº de OS", "Qtd. baixada total", "OS onde aparece",
            "Cadastrado no SAP?", "Código SAP", "Observação"]
    larg = [12, 70, 30, 9, 12, 40, 12, 14, 40]
    F = "Arial"; pend = PatternFill("solid", fgColor="FFF4CC")
    for j, (h, w) in enumerate(zip(cols, larg), 1):
        x = ws.cell(1, j, h); x.font = Font(name=F, bold=True, color="FFFFFF")
        x.fill = PatternFill("solid", fgColor="1C5CAB"); x.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[L(j)].width = w
    ws.row_dimensions[1].height = 30; ws.freeze_panes = "C2"
    for i, r in enumerate(m.itertuples(index=False), 2):
        vals = [int(r[0]), r.Material, r.Família, int(r.n_os), float(r.qtd), r.lista]
        for j, v in enumerate(vals, 1): ws.cell(i, j, v).font = Font(name=F, size=10)
        ws.cell(i, 5).number_format = "#,##0.##"
        for j in (7, 8, 9): ws.cell(i, j).font = Font(name=F, size=10)
        ws.cell(i, 7).fill = pend; ws.cell(i, 8).fill = pend
    n = len(m) + 1
    dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    dv.promptTitle = "Cadastrado no SAP?"; dv.prompt = "Escolha Sim ou Não"
    ws.add_data_validation(dv); dv.add(f"G2:G{n}")
    t = Table(displayName="MateriaisUnicos", ref=f"A1:I{n}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True); ws.add_table(t)
    ws.cell(n + 2, 1, "Ordenado pelo nº de OS em que o material aparece (os mais usados primeiro). "
            "Preencha 'Cadastrado no SAP?' e 'Código SAP' (em amarelo).").font = Font(name=F, size=9, color="52514E")

if __name__ == "__main__":
    d = gera(sys.argv[2:], sys.argv[1])
    print(len(d), "linhas,", d.OS.nunique(), "OS,", d["Cod. Material"].nunique(), "materiais")
