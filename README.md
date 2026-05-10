# irpf-helper

Generates a PDF report and copy-paste-ready text for the Brazilian IRPF declaration (Renda Variável) from B3 export files.

Covers: Bens e Direitos, Rendimentos Isentos, Dividendos, JCP, and Renda Variável (with DARF amounts).

---

## Exports you need from B3

Go to [investidor.b3.com.br](https://investidor.b3.com.br) → Extratos e Informativos.

You need **two files**:

### 1. Negociações (required)

**Extratos e Informativos → Negociações**

Export the full year (01/01/YYYY to 31/12/YYYY). Save as `negociacao.xlsx` inside the `relatorios/` folder.

> **Use negociações, not movimentações, for trade data.**
> The movimentações report records empréstimo de ações (stock lending) returns as
> priced settlement entries — structurally identical to real purchases. If you only
> use movimentações, every time your lent shares come back the script counts them as
> a new buy and your positions will be wildly inflated.

### 2. Movimentações (required for income)

**Extratos e Informativos → Movimentações**

Export the full year. Save as `movimentacao.csv` inside the `relatorios/` folder.

This file is used for dividends, JCP, empréstimo fees, and corporate actions. It is
**not** used for buy/sell trade processing.

---

## Setup

```bash
pip install -r requirements.txt
```

Or with make:

```bash
make run
```

---

## Usage

Place your files in `relatorios/`:

```
relatorios/
  negociacao.xlsx    ← trade executions (negociações export)
  movimentacao.csv   ← income and corporate actions (movimentações export)
```

Run:

```bash
python3 irpf.py
```

Output goes to `output/relatorio_irpf.pdf`. The console also prints a copy-paste block with all the data you need to fill in the Receita Federal's online declaration.

---

## Configuration

At the top of `irpf.py`:

| Constant | Default | Description |
|---|---|---|
| `DECLARACAO_ANO` | `2025` | Declaration year — transactions after 31/12 of this year are ignored |
| `INSTITUTION` | Itaú Corretora | Your broker name for the Discriminação field |
| `INSTITUTION_CNPJ` | 61.194.353/0001-64 | Your broker CNPJ |
| `KNOWN_ETFS` | (list) | Tickers classified as ETF instead of FII — add any missing ones here |

---

## What gets calculated

- **Bens e Direitos**: year-end positions with weighted average cost (preço médio ponderado), grouped by Grupo/Código (AÇÃO = 03/01, FII = 07/03, ETF = 07/01, BDR = 03/01)
- **Isenção AÇÃO swing**: months where total AÇÃO swing sales ≤ R$ 20.000 are exempt (art. 3º Lei 11.033/2004) — gains go to Rendimentos Isentos código 20
- **Day trade detection**: same-day buy+sell on the same ticker triggers the 20% day trade rate instead of swing
- **FII**: always taxable at 20% on swing gains (no R$ 20k exemption)
- **ETF / BDR**: 15% on swing gains
- **JCP**: Tributação Exclusiva na Fonte — 15% IRRF already withheld at source
- **Dividendos / Rendimentos FII**: Rendimentos Isentos
- **DARF código 6015**: shown per month for any taxable renda variável result

---

## Notes

- Fracionário tickers (`PETR4F`, `VALE3F`, etc.) are normalized to the standard lot ticker (`PETR4`, `VALE3`) automatically.
- If `negociacao.xlsx` is not found, the script falls back to using `movimentacao.csv` for trades with a warning. Results may be incorrect if you have lent shares.
- The script only knows about the history in your exported files. If your exports start in January of the declaration year, positions carried over from prior years will start at zero. Export a longer date range (back to your first purchase) for a complete cost basis.
