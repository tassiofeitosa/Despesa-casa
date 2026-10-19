"""Gera planilha/Gastos_Casa.xlsx (para abrir no Google Sheets) a partir de dados/despesas.json.

Uso: python3 scripts/gerar_planilha.py
"""
import json
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

RAIZ = Path(__file__).resolve().parent.parent
SAIDA = RAIZ / "planilha" / "Gastos_Casa.xlsx"
CATEGORIAS = ["Alimentação", "Casa/contas", "Transporte", "Saúde", "Lazer", "Outros"]
LINHAS = 1000  # linhas preparadas para novos lançamentos

FONTE = "Arial"
VERDE = "1E6B52"
VERDE_CLARO = "E3F1EA"
CINZA = "F3F5F4"
BORDA = Side(style="thin", color="D5DBD8")
MOEDA = '"R$" #,##0.00;-"R$" #,##0.00;"-"'
MES = "mmm/yyyy"

titulo = Font(name=FONTE, size=16, bold=True, color=VERDE)
cabec = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
normal = Font(name=FONTE, size=10)
entrada = Font(name=FONTE, size=10, color="0000FF")
nota = Font(name=FONTE, size=9, italic=True, color="5F6B66")
forte = Font(name=FONTE, size=12, bold=True)
fundo_cabec = PatternFill("solid", fgColor=VERDE)
fundo_input = PatternFill("solid", fgColor="FFF2CC")
fundo_destaque = PatternFill("solid", fgColor=VERDE_CLARO)


def cabecalho(ws, linha, nomes, larguras=None):
    for i, nome in enumerate(nomes, start=1):
        c = ws.cell(linha, i, nome)
        c.font, c.fill = cabec, fundo_cabec
        c.alignment = Alignment(horizontal="center", vertical="center")
    if larguras:
        for i, w in enumerate(larguras, start=1):
            ws.column_dimensions[chr(64 + i)].width = w


def lancamentos(wb, despesas):
    ws = wb.active
    ws.title = "Lançamentos"
    cabecalho(ws, 1, ["Data", "Descrição", "Categoria", "Valor", "Pagamento", "Anotado por", "Mês"],
              [12, 34, 16, 14, 13, 14, 12])
    ws.freeze_panes = "A2"
    for r, d in enumerate(despesas, start=2):
        ws.cell(r, 1, date.fromisoformat(d["data"]))
        ws.cell(r, 2, d["descricao"])
        ws.cell(r, 3, d["categoria"])
        ws.cell(r, 4, d["valor"])
        ws.cell(r, 5, d["pagamento"] or None)
        ws.cell(r, 6, d["autor"])
    for r in range(2, LINHAS + 2):
        ws.cell(r, 7, f'=IF(A{r}="","",DATE(YEAR(A{r}),MONTH(A{r}),1))')
        ws.cell(r, 1).number_format = "dd/mm/yyyy"
        ws.cell(r, 4).number_format = MOEDA
        ws.cell(r, 7).number_format = MES
        for c in range(1, 8):
            ws.cell(r, c).font = normal
            ws.cell(r, c).border = Border(bottom=BORDA)
        ws.cell(r, 7).font = Font(name=FONTE, size=10, color="5F6B66")
    ws.auto_filter.ref = f"A1:G{LINHAS + 1}"

    dv = DataValidation(type="list", formula1='"' + ",".join(CATEGORIAS) + '"', allow_blank=True)
    dv.error, dv.errorTitle = "Escolha uma das categorias da lista.", "Categoria"
    ws.add_data_validation(dv)
    dv.add(f"C2:C{LINHAS + 1}")

    # Linha com descrição e sem valor fica amarela: falta o valor
    ws.conditional_formatting.add(
        f"A2:F{LINHAS + 1}",
        FormulaRule(formula=['AND($B2<>"",$D2="")'], fill=PatternFill("solid", fgColor="FCE8B2")))
    ws["H1"] = "Linhas amarelas estão sem valor e não entram nos totais. A coluna Mês é calculada."
    ws["H1"].font = nota
    return ws


def resumo(wb, primeiro_mes):
    ws = wb.create_sheet("Resumo")
    ws["A1"] = "Gastos Casa — resumo por mês"
    ws["A1"].font = titulo
    ws["A2"] = "Totais calculados a partir da aba Lançamentos. Itens sem valor contam só como pendentes."
    ws["A2"].font = nota
    cabecalho(ws, 4, ["Mês", "Total", "Gastos com valor", "Sem valor", "Vs. mês anterior"], [14, 15, 16, 12, 17])
    L = f"Lançamentos!$G$2:$G${LINHAS + 1}"
    V = f"Lançamentos!$D$2:$D${LINHAS + 1}"
    for i in range(12):
        r = 5 + i
        ws.cell(r, 1, primeiro_mes if i == 0 else f"=EDATE(A{r - 1},1)").number_format = MES
        ws.cell(r, 2, f"=SUMIFS({V},{L},A{r})").number_format = MOEDA
        ws.cell(r, 3, f'=COUNTIFS({L},A{r},{V},"<>")')
        ws.cell(r, 4, f'=COUNTIFS({L},A{r},{V},"")')
        ws.cell(r, 5, "" if i == 0 else f'=IF(B{r}=0,"",B{r}-B{r - 1})').number_format = MOEDA
        for c in range(1, 6):
            ws.cell(r, c).font = normal
            ws.cell(r, c).border = Border(bottom=BORDA)
    ws.cell(5, 1).font = entrada
    ws.cell(5, 1).comment = Comment("Primeiro mês do grupo (outubro/2026). Os seguintes são calculados.", "Gastos Casa")

    # Categorias × meses
    ws["G4"] = "Categoria"
    for i in range(12):
        ws.cell(4, 8 + i, f"=A{5 + i}").number_format = MES
    for c in range(7, 20):
        ws.cell(4, c).font, ws.cell(4, c).fill = cabec, fundo_cabec
        ws.cell(4, c).alignment = Alignment(horizontal="center")
        ws.column_dimensions[ws.cell(4, c).column_letter].width = 13
    ws.column_dimensions["G"].width = 16
    C = f"Lançamentos!$C$2:$C${LINHAS + 1}"
    for j, cat in enumerate(CATEGORIAS):
        r = 5 + j
        ws.cell(r, 7, cat).font = normal
        for i in range(12):
            col = ws.cell(4, 8 + i).column_letter
            c = ws.cell(r, 8 + i, f"=SUMIFS({V},{L},{col}$4,{C},$G{r})")
            c.number_format, c.font, c.border = MOEDA, normal, Border(bottom=BORDA)
    r = 5 + len(CATEGORIAS)
    ws.cell(r, 7, "Total").font = Font(name=FONTE, size=10, bold=True)
    for i in range(12):
        col = ws.cell(4, 8 + i).column_letter
        c = ws.cell(r, 8 + i, f"=SUM({col}5:{col}{r - 1})")
        c.number_format, c.font = MOEDA, Font(name=FONTE, size=10, bold=True)

    # Gráficos
    barras = BarChart()
    barras.title = "Total por mês"
    barras.y_axis.title = "R$"
    barras.add_data(Reference(ws, min_col=2, min_row=4, max_row=16), titles_from_data=True)
    barras.set_categories(Reference(ws, min_col=1, min_row=5, max_row=16))
    barras.legend = None
    barras.height, barras.width = 8, 18
    ws.add_chart(barras, "A19")

    pizza = PieChart()
    pizza.title = "Por categoria (mês mais recente com gastos)"
    ws["G13"] = "Mês mais recente"
    ws["G13"].font = nota
    ws["H13"] = f"=MAX({L})"
    ws["H13"].number_format = MES
    for j in range(len(CATEGORIAS)):
        ws.cell(14 + j, 7, f"=G{5 + j}").font = normal
        c = ws.cell(14 + j, 8, f"=INDEX($H{5 + j}:$S{5 + j},MATCH($H$13,$H$4:$S$4,0))")
        c.number_format, c.font = MOEDA, normal
    pizza.add_data(Reference(ws, min_col=8, min_row=14, max_row=13 + len(CATEGORIAS)))
    pizza.set_categories(Reference(ws, min_col=7, min_row=14, max_row=13 + len(CATEGORIAS)))
    pizza.dataLabels = DataLabelList()
    pizza.dataLabels.showPercent = True
    pizza.height, pizza.width = 8, 12
    ws.add_chart(pizza, "K19")
    ws.freeze_panes = "A5"


def planejamento(wb):
    ws = wb.create_sheet("Planejamento")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 14
    ws["A1"] = "Planejamento futuro"
    ws["A1"].font = titulo
    ws["A2"] = "Quanto separar por mês no ritmo atual. Células amarelas são para vocês editarem."
    ws["A2"].font = nota

    D = f"Lançamentos!$A$2:$A${LINHAS + 1}"
    V = f"Lançamentos!$D$2:$D${LINHAS + 1}"
    L = f"Lançamentos!$G$2:$G${LINHAS + 1}"
    linhas = [
        ("Margem de segurança", 0.10, "0%", "Edite: quanto a mais separar além da média."),
        ("Mês mais recente", f"=MAX({L})", MES, None),
        ("Gasto no mês mais recente", f"=SUMIFS({V},{L},B5)", MOEDA, None),
        ("Primeiro dia com anotação no mês", f"=IF(MIN({D})>=B5,DAY(MIN({D})),1)", "0", "Se o grupo começou no meio do mês, conta a partir daí."),
        ("Último dia com anotação", f"=DAY(MAX({D}))", "0", None),
        ("Dias com anotações", "=B8-B7+1", "0", None),
        ("Dias no mês", "=DAY(EOMONTH(B5,0))", "0", None),
        ("Média por dia", "=IF(B9=0,0,B6/B9)", MOEDA, None),
        ("Projeção do mês", "=B6+B11*(B10-B8)", MOEDA, "Gasto até agora + média por dia × dias que faltam."),
        ("Meses completos anteriores usados", f'=COUNTIFS(Resumo!$A$5:$A$16,">="&EDATE(B5,-3),Resumo!$A$5:$A$16,"<"&B5,Resumo!$B$5:$B$16,">0")', "0", "Até 3 meses antes do mês mais recente."),
        ("Média mensal (gastos do grupo)", f'=IF(B13=0,B12,AVERAGEIFS(Resumo!$B$5:$B$16,Resumo!$A$5:$A$16,">="&EDATE(B5,-3),Resumo!$A$5:$A$16,"<"&B5,Resumo!$B$5:$B$16,">0"))', MOEDA, "Sem meses completos ainda, usa a projeção do mês."),
        ("Gastos fixos fora do grupo", "=SUM(F5:F24)", MOEDA, "Soma da tabela ao lado."),
        ("Separar por mês", "=(B14+B15)*(1+B4)", MOEDA, "(Média mensal + fixos) × (1 + margem)."),
    ]
    for i, (rotulo, valor, fmt, comentario) in enumerate(linhas):
        r = 4 + i
        ws.cell(r, 1, rotulo).font = normal
        c = ws.cell(r, 2, valor)
        c.number_format, c.font = fmt, normal
        c.border = Border(bottom=BORDA)
        if comentario:
            c.comment = Comment(comentario, "Gastos Casa")
    ws["B4"].font, ws["B4"].fill = entrada, fundo_input
    for cel in ("A16", "B16"):
        ws[cel].font, ws[cel].fill = forte, fundo_destaque

    ws["A18"] = "Reserva por horizonte"
    ws["A18"].font = Font(name=FONTE, size=12, bold=True, color=VERDE)
    cabecalho(ws, 19, ["Horizonte", "Sem margem", "Com margem"])
    for i, meses in enumerate([1, 3, 6, 12]):
        r = 20 + i
        ws.cell(r, 1, f"{meses} {'mês' if meses == 1 else 'meses'}").font = normal
        ws.cell(r, 4, meses).font = Font(name=FONTE, size=10, color="FFFFFF")
        for col, f in ((2, f"=(B$14+B$15)*D{r}"), (3, f"=B$16*D{r}")):
            c = ws.cell(r, col, f)
            c.number_format, c.font, c.border = MOEDA, normal, Border(bottom=BORDA)
    ws["A25"] = "Estimativa fica mais confiável a cada mês completo anotado."
    ws["A25"].font = nota

    barras = BarChart()
    barras.title = "Quanto juntar (com margem)"
    barras.add_data(Reference(ws, min_col=3, min_row=19, max_row=23), titles_from_data=True)
    barras.set_categories(Reference(ws, min_col=1, min_row=20, max_row=23))
    barras.legend = None
    barras.height, barras.width = 7, 14
    ws.add_chart(barras, "A27")

    # Gastos fixos (para preencher)
    ws["E3"] = "Gastos fixos mensais fora do grupo"
    ws["E3"].font = Font(name=FONTE, size=12, bold=True, color=VERDE)
    for col, nome in ((5, "Descrição"), (6, "Valor por mês")):
        c = ws.cell(4, col, nome)
        c.font, c.fill = cabec, fundo_cabec
    for r in range(5, 25):
        ws.cell(r, 5).fill = fundo_input
        ws.cell(r, 6).fill = fundo_input
        ws.cell(r, 5).font = entrada
        ws.cell(r, 6).font = entrada
        ws.cell(r, 6).number_format = MOEDA
    ws["E25"] = "Preencham aqui aluguel, condomínio, plano de saúde etc. (ex.: Aluguel | 1.800,00)."
    ws["E25"].font = nota


def main():
    despesas = json.loads((RAIZ / "dados" / "despesas.json").read_text(encoding="utf-8"))
    primeiro = min(date.fromisoformat(d["data"]) for d in despesas).replace(day=1)
    wb = Workbook()
    lancamentos(wb, despesas)
    resumo(wb, primeiro)
    planejamento(wb)
    SAIDA.parent.mkdir(exist_ok=True)
    wb.save(SAIDA)
    print(SAIDA)


if __name__ == "__main__":
    main()
