# Despesa casa

Lê o export do grupo do WhatsApp "GASTOS casa!!!!" e alimenta o painel **Gastos Casa** (gastos por mês, categoria, dia e planejamento de quanto separar por mês).

## Atualizar

1. No WhatsApp: grupo → ⋮ → Mais → **Exportar conversa** → **Sem mídia**.
2. Salve o `.txt` em `dados/conversa.txt`.
3. Rode `python3 scripts/parse_whatsapp.py`.
4. Confira `dados/nao_reconhecidas.txt` (mensagens ignoradas) e os itens sem valor em `dados/despesas.csv`.

## Arquivos

- `PADRAO.md` — como anotar os gastos no grupo.
- `scripts/parse_whatsapp.py` — separa gastos de conversa, acha valor, data, categoria e forma de pagamento.
- `scripts/categorias.json` — palavras-chave de cada categoria (edite para ajustar).
- `scripts/config.json` — dia de fechamento da fatura (26), início das anotações (26/10/2026) e valor mensal do vale-alimentação. Gastos antes do início são descartados; cada gasto entra na fatura de 26 a 25.
- `dados/despesas.csv` / `dados/despesas.json` — gastos extraídos.
- `painel/index.html` — código da página do painel.
- `planilha/Gastos_Casa.xlsx` — planilha para o Google Sheets (gerada por `python3 scripts/gerar_planilha.py`).
