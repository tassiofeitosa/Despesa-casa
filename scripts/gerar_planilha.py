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
PAGAMENTOS = ["Crédito", "Débito", "Pix", "Dinheiro", "Vale-alimentação"]
LINHAS = 1000  # linhas que as fórmulas cobrem

FONTE = "Arial"
VERDE = "1E6B52"
VERDE_CLARO = "E3F1EA"
CINZA = "F3F5F4"
BORDA = Side(style="thin", color="D5DBD8")
MOEDA = '"R$" #,##0.00;-"R$" #,##0.00;"-"'
CICLO = "dd/mm/yyyy"  # cada fatura é identificada pelo dia em que começa

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
    cabecalho(ws, 1, ["Data", "Descrição", "Categoria", "Valor", "Pagamento", "Anotado por"],
              [12, 34, 16, 14, 13, 14])
    ws.freeze_panes = "A2"
    for r, d in enumerate(despesas, start=2):
        ws.cell(r, 1, date.fromisoformat(d["data"]))
        ws.cell(r, 2, d["descricao"])
        ws.cell(r, 3, d["categoria"])
        ws.cell(r, 4, d["valor"])
        ws.cell(r, 5, d["pagamento"] or None)
        ws.cell(r, 6, d["autor"])
    # Formata só as linhas preenchidas e uma folga: o arquivo precisa ser pequeno para subir no Drive
    for r in range(2, len(despesas) + 52):
        ws.cell(r, 1).number_format = "dd/mm/yyyy"
        ws.cell(r, 4).number_format = MOEDA
        for c in range(1, 7):
            ws.cell(r, c).font = normal
            ws.cell(r, c).border = Border(bottom=BORDA)
    ws.auto_filter.ref = f"A1:F{LINHAS + 1}"

    dv = DataValidation(type="list", formula1='"' + ",".join(CATEGORIAS) + '"', allow_blank=True)
    dv.error, dv.errorTitle = "Escolha uma das categorias da lista.", "Categoria"
    ws.add_data_validation(dv)
    dv.add(f"C2:C{LINHAS + 1}")
    dvp = DataValidation(type="list", formula1='"' + ",".join(PAGAMENTOS) + '"', allow_blank=True)
    dvp.error, dvp.errorTitle = "Escolha uma das formas de pagamento da lista.", "Pagamento"
    ws.add_data_validation(dvp)
    dvp.add(f"E2:E{LINHAS + 1}")

    # Linha com descrição e sem valor fica amarela: falta o valor
    ws.conditional_formatting.add(
        f"A2:F{LINHAS + 1}",
        FormulaRule(formula=['AND($B2<>"",$D2="")'], fill=PatternFill("solid", fgColor="FCE8B2")))
    ws["H1"] = "Linhas amarelas estão sem valor e não entram nos totais. (banri) = Vale-alimentação."
    ws["H1"].font = nota
    return ws


def resumo(wb, inicio):
    ws = wb.create_sheet("Resumo")
    ws["A1"] = "Gastos Casa — resumo por fatura"
    ws["A1"].font = titulo
    ws["A2"] = "Cada fatura vai do dia de fechamento (26) ao dia 25 seguinte. Itens sem valor contam só como pendentes."
    ws["A2"].font = nota
    cabecalho(ws, 4, ["Fatura (início)", "Total", "Gastos com valor", "Sem valor", "Vs. fatura anterior", "No vale"],
              [15, 15, 16, 12, 18, 14])
    D = f"Lançamentos!$A$2:$A${LINHAS + 1}"
    V = f"Lançamentos!$D$2:$D${LINHAS + 1}"
    P = f"Lançamentos!$E$2:$E${LINHAS + 1}"
    no_ciclo = lambda m: f'{D},">="&{m},{D},"<"&EDATE({m},1)'
    for i in range(12):
        r = 5 + i
        ws.cell(r, 1, inicio if i == 0 else f"=EDATE(A{r - 1},1)").number_format = CICLO
        ws.cell(r, 2, f"=SUMIFS({V},{no_ciclo(f'A{r}')})").number_format = MOEDA
        ws.cell(r, 3, f'=COUNTIFS({no_ciclo(f"A{r}")},{V},"<>")')
        ws.cell(r, 4, f'=COUNTIFS({no_ciclo(f"A{r}")},{V},"")')
        ws.cell(r, 5, "" if i == 0 else f'=IF(B{r}=0,"",B{r}-B{r - 1})').number_format = MOEDA
        ws.cell(r, 6, f'=SUMIFS({V},{no_ciclo(f"A{r}")},{P},"Vale-alimentação")').number_format = MOEDA
        for c in range(1, 7):
            ws.cell(r, c).font = normal
            ws.cell(r, c).border = Border(bottom=BORDA)
    ws.cell(5, 1).font = entrada
    ws.cell(5, 1).comment = Comment("Início da primeira fatura anotada (26/10/2026, fechamento dos cartões). As seguintes são calculadas.", "Gastos Casa")

    # Categorias × faturas
    ws["G4"] = "Categoria"
    for i in range(12):
        ws.cell(4, 8 + i, f"=A{5 + i}").number_format = CICLO
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
            c = ws.cell(r, 8 + i, f"=SUMIFS({V},{no_ciclo(f'{col}$4')},{C},$G{r})")
            c.number_format, c.font, c.border = MOEDA, normal, Border(bottom=BORDA)
    r = 5 + len(CATEGORIAS)
    ws.cell(r, 7, "Total").font = Font(name=FONTE, size=10, bold=True)
    for i in range(12):
        col = ws.cell(4, 8 + i).column_letter
        c = ws.cell(r, 8 + i, f"=SUM({col}5:{col}{r - 1})")
        c.number_format, c.font = MOEDA, Font(name=FONTE, size=10, bold=True)

    # Gráficos
    barras = BarChart()
    barras.title = "Total por fatura"
    barras.y_axis.title = "R$"
    barras.add_data(Reference(ws, min_col=2, min_row=4, max_row=16), titles_from_data=True)
    barras.set_categories(Reference(ws, min_col=1, min_row=5, max_row=16))
    barras.legend = None
    barras.height, barras.width = 8, 18
    ws.add_chart(barras, "A19")

    pizza = PieChart()
    pizza.title = "Por categoria (fatura atual)"
    ws["G13"] = "Fatura atual"
    ws["G13"].font = nota
    ws["H13"] = "=Planejamento!$B$7"
    ws["H13"].number_format = CICLO
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


def planejamento(wb, config):
    ws = wb.create_sheet("Planejamento")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 14
    ws["A1"] = "Planejamento futuro"
    ws["A1"].font = titulo
    ws["A2"] = "Quanto separar por fatura no ritmo atual. Células amarelas são para vocês editarem."
    ws["A2"].font = nota

    D = f"Lançamentos!$A$2:$A${LINHAS + 1}"
    V = f"Lançamentos!$D$2:$D${LINHAS + 1}"
    P = f"Lançamentos!$E$2:$E${LINHAS + 1}"
    R_A, R_B = "Resumo!$A$5:$A$16", "Resumo!$B$5:$B$16"
    anteriores = f'{R_A},">="&EDATE(B7,-3),{R_A},"<"&B7,{R_B},">0"'
    linhas = [
        ("Margem de segurança", 0.10, "0%", "Edite: quanto a mais separar além da média."),
        ("Dia de fechamento da fatura", config["dia_fechamento"], "0", "Dia em que a fatura dos cartões fecha."),
        ("Vale-alimentação por mês", config["vale_mensal"], MOEDA, "Preencham com o valor que cai no vale por mês para ver o saldo."),
        ("Fatura atual (início)", f"=IF(COUNT({D})=0,Resumo!$A$5,DATE(YEAR(MAX({D})),MONTH(MAX({D}))-IF(DAY(MAX({D}))<B5,1,0),B5))", CICLO, "Fatura do gasto mais recente."),
        ("Fatura atual (fim)", "=EDATE(B7,1)-1", CICLO, None),
        ("Gasto na fatura atual", f'=SUMIFS({V},{D},">="&B7,{D},"<="&B8)', MOEDA, None),
        ("Dias com anotações", f"=IF(COUNT({D})=0,0,MAX({D})-B7+1)", "0", "Do dia 26 até o gasto mais recente."),
        ("Dias na fatura", "=B8-B7+1", "0", None),
        ("Média por dia", "=IF(B10=0,0,B9/B10)", MOEDA, None),
        ("Projeção da fatura", "=B9+B12*(B11-B10)", MOEDA, "Gasto até agora + média por dia × dias que faltam."),
        ("Faturas anteriores usadas", f"=COUNTIFS({anteriores})", "0", "Até 3 faturas antes da atual, com gastos."),
        ("Média por fatura (gastos do grupo)", f"=IF(B14=0,B13,AVERAGEIFS({R_B},{anteriores}))", MOEDA, "Sem faturas anteriores ainda, usa a projeção da atual."),
        ("Gastos fixos fora do grupo", "=SUM(F5:F24)", MOEDA, "Soma da tabela ao lado."),
        ("Separar por fatura", "=(B15+B16)*(1+B4)", MOEDA, "(Média por fatura + fixos) × (1 + margem)."),
        ("Gasto no vale nesta fatura", f'=SUMIFS({V},{D},">="&B7,{D},"<="&B8,{P},"Vale-alimentação")', MOEDA, "Compras marcadas com (banri)."),
        ("Saldo do vale", '=IF(B6="","Preencha o valor do vale",B6-B18)', MOEDA, None),
    ]
    for i, (rotulo, valor, fmt, comentario) in enumerate(linhas):
        r = 4 + i
        ws.cell(r, 1, rotulo).font = normal
        c = ws.cell(r, 2, valor)
        c.number_format, c.font = fmt, normal
        c.border = Border(bottom=BORDA)
        if comentario:
            c.comment = Comment(comentario, "Gastos Casa")
    for cel in ("B4", "B5", "B6"):
        ws[cel].font, ws[cel].fill = entrada, fundo_input
    for cel in ("A17", "B17"):
        ws[cel].font, ws[cel].fill = forte, fundo_destaque

    ws["A21"] = "Reserva por horizonte"
    ws["A21"].font = Font(name=FONTE, size=12, bold=True, color=VERDE)
    cabecalho(ws, 22, ["Horizonte", "Sem margem", "Com margem"])
    for i, meses in enumerate([1, 3, 6, 12]):
        r = 23 + i
        ws.cell(r, 1, f"{meses} {'fatura' if meses == 1 else 'faturas'}").font = normal
        ws.cell(r, 4, meses).font = Font(name=FONTE, size=10, color="FFFFFF")
        for col, f in ((2, f"=(B$15+B$16)*D{r}"), (3, f"=B$17*D{r}")):
            c = ws.cell(r, col, f)
            c.number_format, c.font, c.border = MOEDA, normal, Border(bottom=BORDA)
    ws["A27"] = "Estimativa fica mais confiável a cada fatura completa anotada."
    ws["A27"].font = nota

    barras = BarChart()
    barras.title = "Quanto juntar (com margem)"
    barras.add_data(Reference(ws, min_col=3, min_row=22, max_row=26), titles_from_data=True)
    barras.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    barras.legend = None
    barras.height, barras.width = 7, 14
    ws.add_chart(barras, "A29")

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
    config = json.loads((RAIZ / "scripts" / "config.json").read_text(encoding="utf-8"))
    wb = Workbook()
    lancamentos(wb, despesas)
    resumo(wb, date.fromisoformat(config["inicio"]))
    planejamento(wb, config)
    SAIDA.parent.mkdir(exist_ok=True)
    wb.save(SAIDA)
    print(SAIDA)


if __name__ == "__main__":
    main()
