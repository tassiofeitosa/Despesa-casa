# Despesa casa

Controle de gastos da casa. A fonte é a planilha **Gastos Casa** no Google Sheets (aba Lançamentos). O painel **Gastos Casa** lê essa aba ao vivo e mostra cada fatura (26 a 25) com três dinheiros separados:

- **Cartão**: compras no crédito, que vão para a fatura paga no mês seguinte.
- **Vale**: compras no vale-alimentação (`(banri)`), descontadas do saldo do vale (R$ 2.643,77 por mês).
- **Conta**: Pix, débito e dinheiro.

O planejamento ("Separar por fatura") usa só cartão + conta. O vale tem saldo atual e saldo projetado no dia 25 próprios. O total da casa aparece só como referência.

## Lançar

Na aba **Lançamentos**: data, descrição, categoria, valor, pagamento (lista) e quem anotou. Os totais, o resumo, o planejamento e o painel se atualizam sozinhos. Se a linha tiver descrição mas faltar o valor, ela fica amarela; se faltar a forma de pagamento, a célula fica laranja.

## Importar do WhatsApp (opcional)

1. No WhatsApp: grupo → ⋮ → Mais → **Exportar conversa** → **Sem mídia**.
2. Salve o `.txt` em `dados/conversa.txt` e rode `python3 scripts/parse_whatsapp.py`.
3. Confira `dados/nao_reconhecidas.txt` e leve os gastos de `dados/despesas.csv` para a aba Lançamentos.

## Arquivos

- `PADRAO.md` — como anotar os gastos no grupo.
- `scripts/parse_whatsapp.py` — separa gastos de conversa, acha valor, data, categoria e forma de pagamento.
- `scripts/categorias.json` — palavras-chave de cada categoria (edite para ajustar).
- `scripts/config.json` — dia de fechamento da fatura (26), início das anotações (26/10/2026) e valor mensal do vale-alimentação.
- `dados/despesas.csv` / `dados/despesas.json` — gastos extraídos do WhatsApp.
- `painel/index.html` — código da página do painel (lê a planilha pelo conector Google Sheets).
- `scripts/gerar_planilha.py` — gera `planilha/Gastos_Casa.xlsx` com a mesma estrutura da planilha online (cópia de segurança).
