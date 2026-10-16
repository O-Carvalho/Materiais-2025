import pandas as pd, re
SIP = r"\b(7S[AJKLDS]\d{2}|7UT\d{2}|7VK\d{2}|7SR\d{2}|7UM\d{2}|6MD\d{2}|7KE\d{2}|7SE\d{2})\b"
AUX = r"IHM|QDC[AC]|QLCA|PTEL|PCOM|DGO|\bQXC|SINOTICO|ROTAS OPTICAS"
CLIENTES = ["CPFL","CTEEP","NEOENERGIA","ENEL","RME","CERAM","SNEF","TECMON","EDP","ELETROSUL","ELETROBRAS","FURNAS","CHESF","COPEL","CEMIG","ISA"]

def analisa(c):
    c = c.copy(); m = c.Material.str.upper(); vivo = c["Situação"] != "Cancelada"
    is_rele = m.str.startswith(("RELE DIGITAL","RELE DIGTIAL","RELE DE PROTECAO"))
    c["_sip"] = is_rele & m.str.contains("SIEMENS") & m.str.contains(SIP.replace("(", "(?:", 1), regex=True)
    c["_outro"] = is_rele & ~c["_sip"]
    c["_vaga"] = m.str.contains(r"^P[FP]F.*\bIED\b", regex=True)
    c["_nvaga"] = m.str.extract(r"(\d)\s+IED")[0].fillna(1).astype(int) * c.Quantidade
    c["_modelo"] = m.str.extract(SIP)[0]
    semi = c[m.str.startswith("SEMI")].groupby("OS").Material.first()
    rows = []
    for os_, g in c.groupby("OS"):
        gv = g[vivo[g.index]]
        desc = semi.get(os_, "")
        D = desc.upper()
        pref = "SPCS" if "SPCS" in D else "PSI.SIS" if "PSI.SIS" in D else "PSI.EPC" if "PSI.EPC" in D else "PCS" if re.search(r"\bPCS", D) else ("Sem semiacabado" if not desc else "Outro")
        cli = next((k for k in CLIENTES if re.search(rf"\b{k}", D)), "Não identificado")
        sip = gv[gv._sip].Quantidade.sum(); sipbx = gv[gv._sip]["Quantidade Baixada"].sum()
        sipc = g[g._sip & (g["Situação"]=="Cancelada")].Quantidade.sum()
        out = gv[gv._outro].Quantidade.sum()
        vag = gv[gv._vaga]._nvaga.sum()
        mods = ", ".join(sorted(gv[gv._sip]._modelo.dropna().unique()))
        ied = sip + out
        if ied > 0:
            diag = "IED a menos que o previsto" if vag > ied else "OK – tem IED"
        elif cli == "ENEL":
            diag = "Sem IED – ENEL (IED de terceiro)"
        elif vag > 0:
            diag = "ALERTA – painel preparado p/ IED, sem IED"
        elif sipc > 0:
            diag = "IED cancelado"
        elif re.search(AUX, D):
            diag = "Sem IED – painel auxiliar (IHM, SA, telecom, DGO)"
        elif pref == "PSI.EPC":
            diag = "Sem IED – conjunto TC/TP (não leva IED)"
        else:
            diag = "Verificar – painel sem IED no sistema"
        rows.append(dict(OS=int(os_), Projeto=pref, Cliente=cli, Semiacabado=desc, Itens=len(g), Vagas_IED=vag, SIPROTEC=sip,
                         SIPROTEC_baixado=sipbx, SIPROTEC_cancelado=sipc, Outros_IED=out, Modelos=mods, Diagnostico=diag))
    return pd.DataFrame(rows), c
