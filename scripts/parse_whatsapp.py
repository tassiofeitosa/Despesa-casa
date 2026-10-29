"""Lê o export do grupo "GASTOS casa!!!!" e extrai as despesas.

Uso: python3 scripts/parse_whatsapp.py [dados/conversa.txt]

Gera dados/despesas.csv, dados/despesas.json e dados/nao_reconhecidas.txt.
"""
import calendar
import csv
import json
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DADOS = RAIZ / "dados"

# Android: "16/10/2026 11:30 - Bruna: texto"  |  iOS: "[16/10/2026, 11:30:12] Bruna: texto"
CABECALHO = re.compile(
    r"^\[?(\d{1,2}/\d{1,2}/\d{2,4}),? (\d{1,2}:\d{2})(?::\d{2})?\]?(?: -)? (.*)$"
)
# valor: 31 | 11,35 | 1.234,56 | R$ 45,90 — nunca parte de uma data (16/10)
VALOR = re.compile(r"(?<![\d/.,])(?:R\$\s*)?(\d{1,3}(?:\.\d{3})+|\d+)(?:,(\d{1,2}))?(?![\d/])")
TAG = re.compile(r"#(\w+)")
PAGAMENTO = re.compile(r"\(([^)]*)\)")
PARCELAS = re.compile(r"^(\d{1,2})\s*x\b")  # (6x) / (6x sem juros)
EDITADA = re.compile(r"<mensagem editada>", re.I)
FORMAS = {
    "credito": "Crédito", "cartao": "Crédito", "debito": "Débito", "pix": "Pix", "dinheiro": "Dinheiro",
    # (banri) é o vale-alimentação do Banrisul
    "banri": "Vale-alimentação", "banrisul": "Vale-alimentação", "vale": "Vale-alimentação", "va": "Vale-alimentação",
}
DIAS = {"segunda": 0, "terca": 1, "quarta": 2, "quinta": 3, "sexta": 4, "sabado": 5, "domingo": 6}
# Avisos com autor que não são gastos (os avisos do grupo em si chegam sem autor).
SISTEMA = ("<midia oculta>", "arquivo de midia oculto", "mensagem apagada", "mensagem foi apagada")


def sem_acento(texto):
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c)).lower()


def carregar_config():
    return json.loads((Path(__file__).parent / "config.json").read_text(encoding="utf-8"))


def inicio_do_ciclo(dia, fechamento):
    """Início do ciclo da fatura que contém `dia` (ex.: fechamento 26 -> 26/10 a 25/11)."""
    if dia.day >= fechamento:
        return dia.replace(day=fechamento)
    anterior = dia.replace(day=1) - timedelta(days=1)
    return anterior.replace(day=fechamento)


def carregar_categorias():
    cats = json.loads((Path(__file__).parent / "categorias.json").read_text(encoding="utf-8"))
    palavras = [(sem_acento(p), chave) for chave, c in cats.items() for p in c["palavras"]]
    return cats, palavras


def ler_mensagens(caminho):
    mensagens = []
    for linha in caminho.read_text(encoding="utf-8-sig").splitlines():
        linha = linha.replace("‎", "").replace("‏", "")
        m = CABECALHO.match(linha)
        if m:
            d, hora, resto = m.groups()
            dia, mes, ano = d.split("/")
            ano = int(ano) + (2000 if len(ano) == 2 else 0)
            quando = datetime(ano, int(mes), int(dia), *map(int, hora.split(":")))
            autor, _, texto = resto.partition(": ") if ": " in resto else ("", "", resto)
            mensagens.append({"quando": quando, "autor": autor, "linhas": [texto]})
        elif mensagens:
            mensagens[-1]["linhas"].append(linha)
    return mensagens


def dia_da_semana(linha, base):
    """'Sábado' / 'Teve sexta:' -> data mais recente com esse dia, até a data da mensagem."""
    palavras = re.findall(r"\w+", sem_acento(linha))
    if not palavras or len(palavras) > 3:
        return None
    for p in palavras:
        p = p.removesuffix("feira")
        if p in DIAS:
            return base - timedelta(days=(base.weekday() - DIAS[p]) % 7)
    return None


def interpretar_item(linha):
    texto = EDITADA.sub("", linha).strip().lstrip("-•*").strip()
    tags = [sem_acento(t) for t in TAG.findall(texto)]
    texto = TAG.sub("", texto)
    pagamento = None
    parcelas = 1
    for conteudo in PAGAMENTO.findall(texto):
        forma = FORMAS.get(sem_acento(conteudo).strip())
        n = PARCELAS.match(sem_acento(conteudo).strip())
        if forma:
            pagamento = forma
            texto = texto.replace(f"({conteudo})", "")
        elif n:
            parcelas = int(n.group(1))
            texto = texto.replace(f"({conteudo})", "")
    # "Lancheria 42 - credito": forma de pagamento no fim, depois de um traço
    m = re.search(r"\s[-–]\s*(\w+)\s*$", texto)
    if not pagamento and m and FORMAS.get(sem_acento(m.group(1))):
        pagamento = FORMAS[sem_acento(m.group(1))]
        texto = texto[:m.start()]
    valores = list(VALOR.finditer(texto))
    valor = None
    if valores:
        v = valores[-1]
        valor = float(v.group(1).replace(".", "") + "." + (v.group(2) or "0"))
        texto = texto[:v.start()] + texto[v.end():]
    descricao = re.sub(r"\s+", " ", texto).strip(" -:–")
    return descricao, valor, tags, pagamento, parcelas


def dividir_parcelas(data, valor, n):
    """Compra parcelada: n lançamentos mensais no mesmo dia; os centavos que sobram vão na última."""
    base = None if valor is None else round(valor / n, 2)
    for i in range(n):
        mes = data.month - 1 + i
        ano, mes = data.year + mes // 12, mes % 12 + 1
        dia = min(data.day, calendar.monthrange(ano, mes)[1])
        v = base if valor is None or i < n - 1 else round(valor - base * (n - 1), 2)
        yield date(ano, mes, dia), v, i + 1


def categorizar(descricao, tags, cats, palavras):
    for t in tags:
        if t in cats:
            return t, "tag"
    # Vence a expressão com mais palavras; empate, a que aparece primeiro ("Uber teatro").
    alvo = sem_acento(descricao)
    achados = []
    for palavra, chave in palavras:
        m = re.search(rf"(?<!\w){re.escape(palavra)}(?!\w)", alvo)
        if m:
            achados.append((-len(palavra.split()), m.start(), chave))
    if achados:
        return min(achados)[2], "palavra-chave"
    return "outros", "sem-pista"


def extrair(mensagens, cats, palavras, config):
    inicio = datetime.fromisoformat(config["inicio"]).date()
    despesas, ignoradas = [], []
    for msg in mensagens:
        corpo = "\n".join(msg["linhas"])
        if not msg["autor"] or any(s in sem_acento(corpo) for s in SISTEMA):
            ignoradas.append((msg, "sistema"))
            continue
        linhas = [l for l in msg["linhas"] if l.strip()]
        # Uma mensagem é "lista de gastos" se tem algum valor ou cabeçalho de dia.
        tem_valor = any(interpretar_item(l)[1] is not None for l in linhas)
        tem_dia = any(dia_da_semana(l, msg["quando"].date()) for l in linhas)
        if not (tem_valor or tem_dia):
            ignoradas.append((msg, "conversa"))
            continue
        data = msg["quando"].date()
        for linha in linhas:
            novo_dia = dia_da_semana(linha, msg["quando"].date())
            if novo_dia:
                data = novo_dia
                continue
            descricao, valor, tags, pagamento, parcelas = interpretar_item(linha)
            if not descricao and valor is None:
                continue
            curta = len(descricao.split()) <= 5 and not descricao.endswith(("?", "."))
            if valor is None and not curta:
                ignoradas.append(({**msg, "linhas": [linha]}, "linha sem valor"))
                continue
            if data < inicio:
                ignoradas.append(({**msg, "linhas": [linha]}, "antes do início"))
                continue
            categoria, pista = categorizar(descricao, tags, cats, palavras)
            if pista == "sem-pista" and pagamento == "Vale-alimentação":
                categoria, pista = "alimentacao", "vale"  # o vale só aceita comida
            descricao = descricao[:1].upper() + descricao[1:]
            for dia, v, n in dividir_parcelas(data, valor, parcelas):
                despesas.append({
                    "id": str(len(despesas) + 1),
                    "data": dia.isoformat(),
                    "ciclo": inicio_do_ciclo(dia, config["dia_fechamento"]).isoformat(),
                    "descricao": descricao if parcelas == 1 else f"{descricao} (parcela {n}/{parcelas})",
                    "valor": v,
                    "categoria": cats[categoria]["nome"],
                    "pagamento": pagamento or "",
                    "autor": msg["autor"],
                    "origem": "padrao" if re.match(r"^\s*(R\$\s*)?\d", linha) else "antigo",
                    "pista_categoria": pista,
                    "pendente": v is None,
                    "mensagem": msg["quando"].strftime("%d/%m/%Y %H:%M"),
                })
    return despesas, ignoradas


def main():
    origem = Path(sys.argv[1]) if len(sys.argv) > 1 else DADOS / "conversa.txt"
    cats, palavras = carregar_categorias()
    mensagens = ler_mensagens(origem)
    despesas, ignoradas = extrair(mensagens, cats, palavras, carregar_config())

    campos = ["id", "data", "ciclo", "descricao", "valor", "categoria", "pagamento", "autor", "origem",
              "pista_categoria", "pendente", "mensagem"]
    with open(DADOS / "despesas.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        for d in despesas:
            w.writerow({**d, "valor": "" if d["valor"] is None else f"{d['valor']:.2f}"})
    (DADOS / "despesas.json").write_text(
        json.dumps(despesas, ensure_ascii=False, indent=1), encoding="utf-8")
    with open(DADOS / "nao_reconhecidas.txt", "w", encoding="utf-8") as f:
        for msg, motivo in ignoradas:
            texto = " / ".join(l for l in msg["linhas"] if l.strip())
            f.write(f"[{motivo}] {msg['quando']:%d/%m/%Y %H:%M} {msg['autor']}: {texto}\n")

    com_valor = [d for d in despesas if d["valor"] is not None]
    print(f"{len(mensagens)} mensagens | {len(despesas)} despesas "
          f"({len(despesas) - len(com_valor)} sem valor) | {len(ignoradas)} ignoradas")
    print(f"Total com valor: R$ {sum(d['valor'] for d in com_valor):.2f}")


if __name__ == "__main__":
    main()
