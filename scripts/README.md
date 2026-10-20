# Scripts de análise de materiais por OS

Rodar na pasta onde estão os arquivos `OS <número>.xls` exportados do ERP, nesta ordem:

| Passo | Comando | O que faz |
|---|---|---|
| 1 | `python3 scripts/add_os.py` | Adiciona a coluna **OS** (coluna A) em cada arquivo, com o número tirado do nome |
| 2 | `python3 scripts/consolida.py` | Junta todos os arquivos em `Consolidado OS.xlsx` |
| 3 | `ANO=2026 PYTHONPATH=scripts python3 scripts/relatorio.py "Relatório Materiais 2026.xlsx"` | Gera o relatório (dashboard, ranking, itens em aberto, por OS e por família) |
| 4 | `PYTHONPATH=scripts python3 scripts/ied_xlsx.py "Verificação IED SIPROTEC.xlsx"` | Cruza IEDs SIPROTEC, placas com recorte para IED e tipo de projeto por OS |
| 5 | `PYTHONPATH=scripts python3 scripts/checklist_ied.py "Conferência IED por OS.xlsx"` | Planilha de conferência: OS com placa para IED, se tem IED e coluna "Faz sentido?" para preencher |

**Painéis em montagem (MEGA → SAP):** `python3 scripts/base_montagem.py "Paineis em montagem/Base Materiais - Painéis em Montagem.xlsx" <arquivos de entrada .xlsx>` junta os arquivos, remove linhas duplicadas e ordena por OS numa aba só.

- `familias.py`: regras que agrupam os materiais em famílias (edite para ajustar).
- `ied.py`: regras da verificação de IED (códigos SIPROTEC, clientes, painéis auxiliares).

Dependências: `pip install xlrd xlwt pandas openpyxl`. Os passos 3 e 4 geram fórmulas; abra no Excel para calcular.
