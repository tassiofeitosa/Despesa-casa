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
ROXO, ROXO_CLARO = "6B4599", "EDE6F5"
AZUL, AZUL_CLARO = "3366A6", "E6EDF7"
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
    ws.conditional_formatting.add(
        f"E2:E{LINHAS + 1}",
        FormulaRule(formula=['AND($B2<>"",$E2="")'], fill=PatternFill("solid", fgColor="F8CBAD")))
    ws["H1"] = "Linhas amarelas estão sem valor e não entram nos totais. (banri) = Vale-alimentação."
    ws["H1"].font = nota
    return ws


def resumo(wb, inicio):
    """Três blocos lado a lado que nunca se somam: fatura do cartão, vale-alimentação e conta."""
    ws = wb.create_sheet("Resumo")
    ws["A1"] = "Gastos Casa — resumo por fatura"
    ws["A1"].font = titulo
    ws["A2"] = "Cada bloco é um dinheiro diferente e não se soma com os outros. Cada período vai do dia 26 ao dia 25."
    ws["A2"].font = nota
    blocos = (("A4:E4", "FATURA DO CARTÃO DE CRÉDITO — o que vocês vão pagar", VERDE, VERDE_CLARO, range(1, 6)),
              ("G4:I4", "VALE-ALIMENTAÇÃO — saldo que a empresa coloca", ROXO, ROXO_CLARO, range(7, 10)),
              ("K4:K4", "CONTA — Pix, débito, dinheiro", AZUL, AZUL_CLARO, range(11, 12)))
    nomes = {1: "Início", 2: "Fim", 3: "Valor da fatura", 4: "Compras no cartão", 5: "Vs. fatura anterior",
             7: "Crédito do vale", 8: "Gasto no vale", 9: "Saldo do vale", 11: "Gasto na conta", 13: "Cartão + conta"}
    for faixa, texto, escuro, claro, cols in blocos:
        if ":" in faixa and faixa.split(":")[0] != faixa.split(":")[1]:
            ws.merge_cells(faixa)
        c = ws[faixa.split(":")[0]]
        c.value, c.font, c.fill = texto, cabec, PatternFill("solid", fgColor=escuro)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col in cols:
            h = ws.cell(5, col, nomes[col])
            h.font, h.fill = Font(name=FONTE, size=10, bold=True), PatternFill("solid", fgColor=claro)
            h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["M4"], ws["M5"] = "Cartão + conta (usado no Planejamento)", "Cartão + conta"
    ws.column_dimensions["M"].hidden = True
    for col, w in zip("ABCDEFGHIJKL", (12, 12, 16, 14, 16, 2, 15, 15, 15, 2, 18, 2)):
        ws.column_dimensions[col].width = w

    D = f"Lançamentos!$A$2:$A${LINHAS + 1}"
    V = f"Lançamentos!$D$2:$D${LINHAS + 1}"
    P = f"Lançamentos!$E$2:$E${LINHAS + 1}"
    C = f"Lançamentos!$C$2:$C${LINHAS + 1}"
    periodo = lambda r: f'{D},">="&$A{r},{D},"<="&$B{r}'
    forma = lambda r, f: f'SUMIFS({V},{periodo(r)},{P},"{f}")'
    for i in range(12):
        r = 6 + i
        ws.cell(r, 1, inicio if i == 0 else f"=EDATE(A{r - 1},1)")
        ws.cell(r, 2, f"=EDATE(A{r},1)-1")
        ws.cell(r, 3, f"={forma(r, 'Crédito')}")
        ws.cell(r, 4, f'=COUNTIFS({periodo(r)},{P},"Crédito")')
        ws.cell(r, 5, "" if i == 0 else f'=IF(C{r}=0,"",C{r}-C{r - 1})')
        ws.cell(r, 7, f'=IF($A{r}>Planejamento!$B$6,"",Planejamento!$B$22)')
        ws.cell(r, 8, f"={forma(r, 'Vale-alimentação')}")
        ws.cell(r, 9, f'=IF(G{r}="","",G{r}-H{r})')
        ws.cell(r, 11, f"={forma(r, 'Pix')}+{forma(r, 'Débito')}+{forma(r, 'Dinheiro')}")
        ws.cell(r, 13, f"=C{r}+K{r}")
        for col in (1, 2, 3, 4, 5, 7, 8, 9, 11, 13):
            ws.cell(r, col).font = normal
            ws.cell(r, col).border = Border(bottom=BORDA)
        for col in (1, 2):
            ws.cell(r, col).number_format = CICLO
        for col in (3, 5, 7, 8, 9, 11, 13):
            ws.cell(r, col).number_format = MOEDA
        ws.cell(r, 3).font = ws.cell(r, 9).font = Font(name=FONTE, size=10, bold=True)
    ws.cell(6, 1).font = entrada
    ws.cell(6, 1).comment = Comment("Início da primeira fatura anotada (26/10/2026, fechamento dos cartões). As seguintes são calculadas.", "Gastos Casa")
    ws.conditional_formatting.add("I6:I17", FormulaRule(formula=["AND(ISNUMBER(I6),I6<0)"], font=Font(color="BF1A1A")))

    # Fatura atual por categoria, uma coluna por dinheiro
    ws["A20"] = "Fatura atual por categoria (cada coluna é um dinheiro diferente)"
    ws["A20"].font = Font(name=FONTE, size=12, bold=True, color=VERDE)
    for col, (nome, cor) in enumerate((("Categoria", "EDF0F0"), ("Cartão", VERDE_CLARO), ("Vale", ROXO_CLARO), ("Conta", AZUL_CLARO)), start=1):
        h = ws.cell(21, col, nome)
        h.font, h.fill = Font(name=FONTE, size=10, bold=True), PatternFill("solid", fgColor=cor)
    atual = f'{D},">="&Planejamento!$B$6,{D},"<="&Planejamento!$B$7'
    soma = lambda r, f: f'SUMIFS({V},{atual},{C},$A{r},{P},"{f}")'
    for j, cat in enumerate(CATEGORIAS):
        r = 22 + j
        ws.cell(r, 1, cat)
        ws.cell(r, 2, f"={soma(r, 'Crédito')}")
        ws.cell(r, 3, f"={soma(r, 'Vale-alimentação')}")
        ws.cell(r, 4, f"={soma(r, 'Pix')}+{soma(r, 'Débito')}+{soma(r, 'Dinheiro')}")
    r = 22 + len(CATEGORIAS)
    ws.cell(r, 1, "Total")
    for col in "BCD":
        ws[f"{col}{r}"] = f"=SUM({col}22:{col}{r - 1})"
    for rr in range(22, r + 1):
        for col in range(1, 5):
            c = ws.cell(rr, col)
            c.font, c.border = Font(name=FONTE, size=10, bold=rr == r), Border(bottom=BORDA)
            if col > 1:
                c.number_format = MOEDA
    ws.cell(r + 1, 1, "Não some as colunas: o cartão vai para a fatura, o vale sai do saldo do vale e a conta sai do banco.").font = nota

    barras = BarChart()
    barras.title = "Fatura do cartão por período"
    barras.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    barras.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    barras.legend = None
    barras.height, barras.width = 7, 13
    ws.add_chart(barras, "A31")

    vale = BarChart()
    vale.title = "Vale-alimentação: gasto e saldo por período"
    vale.grouping, vale.overlap = "stacked", 100
    vale.add_data(Reference(ws, min_col=8, max_col=9, min_row=5, max_row=17), titles_from_data=True)
    vale.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    vale.height, vale.width = 7, 13
    ws.add_chart(vale, "G31")
    ws.freeze_panes = "A6"


def planejamento(wb, config):
    ws = wb.create_sheet("Planejamento")
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 14
    ws["A1"] = "Planejamento futuro"
    ws["A1"].font = titulo
    ws["A2"] = "Cartão e conta saem do bolso; o vale é descontado do saldo do vale. Células amarelas são para vocês editarem."
    ws["A2"].font = nota

    D = f"Lançamentos!$A$2:$A${LINHAS + 1}"
    V = f"Lançamentos!$D$2:$D${LINHAS + 1}"
    P = f"Lançamentos!$E$2:$E${LINHAS + 1}"
    R_A, R_E = "Resumo!$A$6:$A$17", "Resumo!$M$6:$M$17"
    anteriores = f'{R_A},">="&EDATE(B6,-3),{R_A},"<"&B6,{R_E},">0"'
    na_fatura = f'{D},">="&B6,{D},"<="&B7'
    forma = lambda f: f'SUMIFS({V},{na_fatura},{P},"{f}")'
    secao = Font(name=FONTE, size=12, bold=True, color=VERDE)
    linhas = {
        4: ("Margem de segurança", 0.10, "0%", "Edite: quanto a mais separar além da média."),
        5: ("Dia de fechamento da fatura", config["dia_fechamento"], "0", "Dia em que a fatura dos cartões fecha."),
        6: ("Fatura atual (início) — pela data de hoje", "=DATE(YEAR(TODAY()),MONTH(TODAY())-IF(DAY(TODAY())<B5,1,0),B5)", CICLO, "Fatura que contém a data de hoje."),
        7: ("Fatura atual (fim)", "=EDATE(B6,1)-1", CICLO, None),
        8: ("Dias na fatura", "=B7-B6+1", "0", None),
        9: ("Dias já passados na fatura", "=MAX(0,MIN(TODAY(),B7)-B6+1)", "0", "Do dia 26 até hoje."),
        11: ("Cartão (crédito) nesta fatura", f"={forma('Crédito')}", MOEDA, "Vai para a fatura que vocês pagam no mês seguinte."),
        12: ("Conta (Pix, débito, dinheiro)", f"={forma('Pix')}+{forma('Débito')}+{forma('Dinheiro')}", MOEDA, None),
        13: ("Cartão + conta", "=B11+B12", MOEDA, None),
        14: ("Média por dia", "=IF(B9=0,0,B13/B9)", MOEDA, None),
        15: ("Projeção até o dia 25", "=B13+B14*(B8-B9)", MOEDA, "Gasto até agora + média por dia × dias que faltam."),
        16: ("Faturas anteriores usadas", f"=COUNTIFS({anteriores})", "0", "Até 3 faturas antes da atual, com gastos."),
        17: ("Média por fatura (cartão + conta)", f"=IF(B16=0,B15,AVERAGEIFS({R_E},{anteriores}))", MOEDA, "Sem faturas anteriores ainda, usa a projeção da atual."),
        18: ("Gastos fixos fora do grupo", "=SUM(F5:F24)", MOEDA, "Soma da tabela ao lado."),
        19: ("Separar por fatura", "=(B17+B18)*(1+B4)", MOEDA, "(Média cartão + conta + fixos) × (1 + margem). O vale fica de fora."),
        22: ("Vale por mês", config["vale_mensal"], MOEDA, "Valor que cai no vale-alimentação por mês."),
        23: ("Gasto no vale nesta fatura", f"={forma('Vale-alimentação')}", MOEDA, "Compras marcadas com (banri)."),
        24: ("Saldo atual do vale", '=IF(B22="","Preencha o valor do vale",B22-B23)', MOEDA, None),
        25: ("Média por dia no vale", "=IF(B9=0,0,B23/B9)", MOEDA, None),
        26: ("Saldo projetado no dia 25", '=IF(B22="","",B22-(B23+B25*(B8-B9)))', MOEDA, "Valor do vale menos o gasto projetado no ritmo atual."),
        27: ("Situação do vale", '=IF(B22="","",IF(B26<0,"Atenção: no ritmo atual o vale acaba antes do dia 25; faltam cerca de R$ "&ROUND(-B26,0)&", que vão cair no cartão ou na conta.","Dentro do saldo do vale"))', "@", None),
        29: ("Gasto total da casa nesta fatura (referência)", f"=SUMIFS({V},{na_fatura})", MOEDA, "Cartão + vale + conta, só para referência."),
    }
    for r, (rotulo, valor, fmt, comentario) in linhas.items():
        ws.cell(r, 1, rotulo).font = normal
        c = ws.cell(r, 2, valor)
        c.number_format, c.font = fmt, normal
        c.border = Border(bottom=BORDA)
        if comentario:
            c.comment = Comment(comentario, "Gastos Casa")
    for r, texto in ((10, "Fatura do cartão + conta (sai do bolso)"), (21, "Vale-alimentação (sai do saldo do vale)")):
        ws.cell(r, 1, texto).font = secao
    for cel in ("B4", "B5", "B22"):
        ws[cel].font, ws[cel].fill = entrada, fundo_input
    for cel in ("A19", "B19"):
        ws[cel].font, ws[cel].fill = forte, fundo_destaque
    for cel in ("A29", "B29"):
        ws[cel].font = nota
    ws.conditional_formatting.add("A26:B27", FormulaRule(formula=["AND(ISNUMBER($B$26),$B$26<0)"],
                                                          font=Font(bold=True, color="BF1A1A")))

    ws["A31"] = "Reserva por horizonte (cartão + conta)"
    ws["A31"].font = secao
    cabecalho(ws, 32, ["Horizonte", "Sem margem", "Com margem"])
    for i, meses in enumerate([1, 3, 6, 12]):
        r = 33 + i
        ws.cell(r, 1, f"{meses} {'fatura' if meses == 1 else 'faturas'}").font = normal
        ws.cell(r, 4, meses).font = Font(name=FONTE, size=10, color="FFFFFF")
        for col, f in ((2, f"=(B$17+B$18)*D{r}"), (3, f"=B$19*D{r}")):
            c = ws.cell(r, col, f)
            c.number_format, c.font, c.border = MOEDA, normal, Border(bottom=BORDA)
    ws["A37"] = "Estimativa fica mais confiável a cada fatura completa anotada."
    ws["A37"].font = nota

    barras = BarChart()
    barras.title = "Quanto juntar para cartão + conta (com margem)"
    barras.add_data(Reference(ws, min_col=3, min_row=32, max_row=36), titles_from_data=True)
    barras.set_categories(Reference(ws, min_col=1, min_row=33, max_row=36))
    barras.legend = None
    barras.height, barras.width = 7, 14
    ws.add_chart(barras, "A39")

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
