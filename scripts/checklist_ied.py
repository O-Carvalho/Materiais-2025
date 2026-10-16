"""Gera a planilha de conferência de IED: uma linha por OS com placa frontal com recorte para IED."""
import re, sys, warnings
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from ied import analisa
warnings.filterwarnings("ignore")

PROJ = r"(SPCS[_ ][0-9_]+\d|PSI\.\w+[_ ][0-9_]+\d|PCS[_ ][0-9_]+\d)"

def projeto_tag(desc):
    d = re.sub(r"^SEMI ACAB\w*\s+", "", str(desc).strip(), flags=re.I)
    p = re.search(PROJ, d, re.I)
    proj = p.group(1).strip("_ ") if p else ""
    partes = [x.strip() for x in re.split(r"\s[-–]\s?|\s?[-–]\s", d) if x.strip()]
    ult = partes[-1] if len(partes) > 1 else d
    m = re.search(r"\b(?:PN|PNL|PAINEL)\s+(.+)$", ult, re.I)
    if m:
        tag = m.group(1)
    elif len(partes) > 1 and not re.search(r"\b(SE|ETD|CTEEP)\b", ult):
        tag = ult
    else:
        tag = ult.split()[-1]
    return proj, tag.strip()

def gera(saida, respostas=None):
    """respostas: planilha de conferência já preenchida (carrega 'Faz sentido?' e 'Observação')."""
    c = pd.read_excel("Consolidado OS.xlsx")
    r, cc = analisa(c)
    ied = cc[(cc._sip | cc._outro) & (cc["Situação"] != "Cancelada")]
    bx = ied.groupby("OS")["Quantidade Baixada"].sum()
    r["IED_bx"] = r.OS.map(bx).fillna(0)
    r = r[(r.Vagas_IED > 0) | (r.IED_bx > 0)].copy()
    r["Tem placa"] = (r.Vagas_IED > 0).map({True: "Sim", False: "Não"})
    r["Tem IED"] = ((r.SIPROTEC + r.Outros_IED) > 0).map({True: "Sim", False: "Não"})
    r[["Projeto", "TAG"]] = r.Semiacabado.apply(lambda s: pd.Series(projeto_tag(s) if s else ("", "")))
    resp = {}
    if respostas:
        a = pd.read_excel(respostas)
        resp = {int(o): (f, ob) for o, f, ob in zip(a["OS"], a["Faz sentido?"], a["Observação"])}
    def resposta(row):
        f, ob = resp.get(int(row.OS), (None, None))
        f = None if pd.isna(f) else f; ob = None if pd.isna(ob) else ob
        if row["Tem IED"] == "Não" and row.SIPROTEC_cancelado > 0 and not ob:
            ob = "IED cancelado no sistema (outlier)"; f = f or "Faz sentido"
        if row["Tem IED"] == "Sim" and not f: f = "Faz sentido"
        if row["Tem placa"] == "Não" and not ob: ob = "Placa sem 'IED' na descrição"
        return pd.Series([f, ob])
    r[["Faz", "Obs"]] = r.apply(resposta, axis=1)
    # IED fora do sistema: painel sem IED no sistema e não cancelado -> estimado pelos recortes da placa
    fora = (r["Tem IED"] == "Não") & (r.SIPROTEC_cancelado == 0)
    r["IED_fora"] = r.Vagas_IED.where(fora, 0)
    r = r.sort_values(["Tem IED", "Tem placa", "Projeto", "OS"], ascending=[True, False, True, True])

    wb = Workbook(); ws = wb.active; ws.title = "Conferência IED"
    cols = ["OS", "Descrição", "Projeto", "TAG", "Tem placa p/ IED?", "Qtd. recortes p/ IED", "Tem IED?",
            "Qtd. IED baixados (sistema)", "Qtd. IED fora do sistema", "Faz sentido?", "Observação"]
    larg = [8, 70, 18, 13, 10, 10, 9, 12, 12, 16, 42]
    F = "Arial"; hdr = PatternFill("solid", fgColor="1C5CAB"); pend = PatternFill("solid", fgColor="FFF4CC")
    for j, (h, w) in enumerate(zip(cols, larg), 1):
        x = ws.cell(1, j, h); x.font = Font(name=F, bold=True, color="FFFFFF"); x.fill = hdr
        x.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[chr(64 + j)].width = w
    ws.row_dimensions[1].height = 42; ws.freeze_panes = "B2"
    num = lambda v: int(v) if float(v).is_integer() else float(v)
    for i, row in enumerate(r.itertuples(index=False), 2):
        vals = [row.OS, row.Semiacabado or "(OS sem semiacabado)", row.Projeto, row.TAG, r["Tem placa"].iloc[i-2],
                num(row.Vagas_IED), r["Tem IED"].iloc[i-2], num(row.IED_bx), num(row.IED_fora), r.Faz.iloc[i-2], r.Obs.iloc[i-2]]
        for j, v in enumerate(vals, 1):
            ws.cell(i, j, v).font = Font(name=F, size=10)
        if not vals[9]: ws.cell(i, 10).fill = pend
    n = len(r) + 1; t = n + 1
    ws.cell(t, 1, "Total").font = Font(name=F, bold=True)
    for j in (6, 8, 9):
        col = chr(64 + j)
        x = ws.cell(t, j, f"=SUM({col}2:{col}{n})"); x.font = Font(name=F, bold=True)
    ws.cell(t, 2, f'=COUNTA(A2:A{n})&" OS"').font = Font(name=F, bold=True)
    notas = [
        "Qtd. IED baixados (sistema): IEDs SIPROTEC (7SA, 7SJ, 7SL, 7UT, 7VK…) e de outros fabricantes baixados pelo almoxarifado na OS.",
        "Qtd. IED fora do sistema: painel sem IED no sistema; estimado pela quantidade de recortes na placa (revenda, terceiro, avulso).",
        "Recortes a mais que IED baixado = IED futuro, montado em campo. Não precisa bater.",
        "IED cancelado no sistema conta zero nas duas colunas.",
    ]
    for k, txt in enumerate(notas, t + 2):
        ws.cell(k, 1, txt).font = Font(name=F, size=9, color="52514E")
    dv = DataValidation(type="list", formula1='"Faz sentido,Não faz sentido"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"J2:J{n}")
    tb = Table(displayName="ConferenciaIED", ref=f"A1:K{n}")
    tb.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True); ws.add_table(tb)
    wb.save(saida)
    return r

if __name__ == "__main__":
    r = gera(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(len(r), "OS;", int(r.IED_bx.sum()), "IED baixados;", int(r.IED_fora.sum()), "IED fora do sistema")
