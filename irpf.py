#!/usr/bin/env python3
"""IRPF Helper — B3 movimentação CSV → PDF (pandas + matplotlib)"""

import os
import re
import textwrap
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# ── Constants ──────────────────────────────────────────────────────────────────

KNOWN_ETFS = {
    "BOVA11", "SMAL11", "IVVB11", "XINA11", "SPXI11", "NASD11",
    "HASH11", "GOLD11", "DIVO11", "FIND11", "BBSD11", "ECOO11",
    "MATB11", "ISUS11", "MVBI11", "TFOF11", "QBTC11", "FIXA11", "AGRI11",
}

TAX_SW = {"AÇÃO": 0.15, "FII": 0.20, "ETF": 0.15, "BDR": 0.15}
TAX_DT = {"AÇÃO": 0.20, "FII": 0.20, "ETF": 0.20, "BDR": 0.20}
EXEMPT = 20_000.0

IRPF_GROUP = {"AÇÃO": "03", "FII": "07", "ETF": "07", "BDR": "03"}
IRPF_CODE  = {"AÇÃO": "01", "FII": "03", "ETF": "01", "BDR": "01"}
DARF_CODE  = "6015"

DECLARACAO_ANO   = 2025
INSTITUTION      = "Itaú Corretora de Valores S/A"
INSTITUTION_CNPJ = "61.194.353/0001-64"

OUT_DIR  = Path("output")
PW, PH   = 11.69, 8.27
MAX_ROWS = 28

C = dict(
    hbg="#1d3461", hfg="#ffffff",
    even="#eef2fb", odd="#ffffff",
    edge="#b8c4d8", title="#1d3461",
    pos="#1a7a1a", neg="#c82020",
    bg="#f5f7fc",
)

# ── Formatting helpers ─────────────────────────────────────────────────────────

def _is_nan(v):
    return isinstance(v, float) and v != v

def brl(v):
    return "—" if v is None or _is_nan(v) else f"R$ {v:,.2f}"

def brl4(v):
    return "—" if v is None or _is_nan(v) else f"R$ {v:,.4f}"

# ── CSV helpers ────────────────────────────────────────────────────────────────

def find_csv():
    for d in (".", "relatorios"):
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.endswith(".csv"):
                return os.path.join(d, f)
    raise FileNotFoundError("No CSV found in . or relatorios/")


def find_xlsx():
    for d in (".", "relatorios"):
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.endswith(".xlsx") and "negoci" in f.lower():
                return os.path.join(d, f)
    return None


def parse_brl(s):
    if pd.isna(s):
        return None
    s = re.sub(r"R\$\s*", "", str(s)).replace(",", "").strip()
    return float(s) if s not in ("-", "") else None


def normalize_ticker(produto):
    t = str(produto).strip().split()[0]
    if t.endswith("F") and len(t) in (5, 6, 7) and t[:-1].isalnum():
        return t[:-1]
    return t


def classify(ticker):
    if ticker in KNOWN_ETFS:
        return "ETF"
    if ticker.endswith("11"):
        return "FII"
    if ticker.endswith(("34", "35", "39")):
        return "BDR"
    return "AÇÃO"

# ── Position model ─────────────────────────────────────────────────────────────

class Position:
    def __init__(self, ticker):
        self.ticker     = ticker
        self.kind       = classify(ticker)
        self.qty        = 0
        self.total_cost = 0.0
        self.pnl        = defaultdict(float)
        self.sells      = []
        self.income     = defaultdict(float)

    @property
    def avg_cost(self):
        return self.total_cost / self.qty if self.qty > 0 else 0.0

    def buy(self, qty, total):
        self.total_cost += total
        self.qty += qty

    def sell(self, date, qty, price, total, op):
        basis = self.avg_cost * qty
        pnl   = total - basis
        self.pnl[op]    += pnl
        self.total_cost  = max(0.0, self.total_cost - basis)
        self.qty         = max(0, self.qty - qty)
        self.sells.append(dict(
            date=date, qty=qty, price=price,
            revenue=total, cost_basis=basis, pnl=pnl, op=op,
        ))

# ── Data loading & processing ──────────────────────────────────────────────────

# Fallback trade names used when no negociacao.xlsx is present
_SALE_MOVS = {"Transferência - Liquidação", "Leilão", "Venda", "Compra"}


def load_df(csv_file):
    df = pd.read_csv(csv_file, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()
    df["Data"] = pd.to_datetime(df["Data"].str.strip(), format="%d/%m/%Y")

    # Drop anything after 31/12 of the declaration year so future purchases
    # don't inflate the year-end balance, while keeping older history for
    # correct average-cost calculation.
    df = df[df["Data"].dt.year <= DECLARACAO_ANO]

    df["Ticker"]    = df["Produto"].map(normalize_ticker)
    df["UnitPrice"] = df["Preço unitário"].map(parse_brl)
    df["OpValue"]   = df["Valor da Operação"].map(parse_brl)
    df["Qty"]       = pd.to_numeric(
        df["Quantidade"].astype(str).str.strip(), errors="coerce"
    ).fillna(0).astype(int)
    df["Direction"] = df["Entrada/Saída"].str.strip()
    df["Mov"]       = df["Movimentação"].str.strip()
    df["Month"]     = df["Data"].dt.strftime("%Y-%m")
    return df.sort_values("Data").reset_index(drop=True)


def load_negocios(xlsx_file):
    """Load negociação (trade execution) report — pure Compra/Venda, trade dates."""
    df = pd.read_excel(xlsx_file, engine="openpyxl")
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "Data do Negócio":     "Data",
        "Tipo de Movimentação": "Direction",
        "Código de Negociação": "Produto",
        "Quantidade":           "Qty",
        "Preço":                "UnitPrice",
        "Valor":                "OpValue",
    })
    df["Data"]      = pd.to_datetime(df["Data"], format="%d/%m/%Y")
    df              = df[df["Data"].dt.year <= DECLARACAO_ANO]
    df["Ticker"]    = df["Produto"].map(normalize_ticker)
    df["Month"]     = df["Data"].dt.strftime("%Y-%m")
    df["Qty"]       = pd.to_numeric(df["Qty"],       errors="coerce").fillna(0).astype(int)
    df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="coerce")
    df["OpValue"]   = pd.to_numeric(df["OpValue"],   errors="coerce")
    # Sort by date then Direction ascending ("Compra" < "Venda") so buys are
    # processed before same-day sells — critical for correct day-trade P&L.
    return df.sort_values(["Data", "Direction"], ascending=[True, True]).reset_index(drop=True)


def find_day_trades(df):
    """Return {(date_str, ticker): fraction_of_sells_that_are_daytrade} from movimentacao."""
    liq = df[
        df["Mov"].isin(_SALE_MOVS) &
        df["UnitPrice"].notna() & (df["Qty"] > 0)
    ].copy()
    liq["DS"] = liq["Data"].dt.strftime("%Y-%m-%d")
    buys  = liq[liq["Direction"] == "Credito"].groupby(["DS", "Ticker"])["Qty"].sum()
    sells = liq[liq["Direction"] == "Debito"].groupby(["DS", "Ticker"])["Qty"].sum()
    fracs = {}
    for key in buys.index.intersection(sells.index):
        fracs[key] = min(buys[key], sells[key]) / sells[key]
    return fracs


def find_day_trades_negocios(negocios_df):
    """Return day-trade fractions from negociacao data (trade dates, Compra/Venda)."""
    n = negocios_df.copy()
    n["DS"] = n["Data"].dt.strftime("%Y-%m-%d")
    buys  = n[n["Direction"] == "Compra"].groupby(["DS", "Ticker"])["Qty"].sum()
    sells = n[n["Direction"] == "Venda"].groupby(["DS", "Ticker"])["Qty"].sum()
    fracs = {}
    for key in buys.index.intersection(sells.index):
        fracs[key] = min(buys[key], sells[key]) / sells[key]
    return fracs


def process(df, negocios_df=None):
    """
    negocios_df: negociacao DataFrame — when provided, used as the sole source of
    buy/sell trades (avoids counting empréstimo returns as purchases).
    df (movimentacao) is always used for income and corporate actions.
    """
    dt_fracs  = (find_day_trades_negocios(negocios_df)
                 if negocios_df is not None
                 else find_day_trades(df))
    positions = {}
    monthly   = defaultdict(lambda: defaultdict(lambda: {"swing": 0.0, "daytrade": 0.0}))

    def pos(t):
        if t not in positions:
            positions[t] = Position(t)
        return positions[t]

    def _sell(p, month, ticker, date, qty, uprice, total):
        key    = (date.strftime("%Y-%m-%d"), ticker)
        dtfrac = dt_fracs.get(key, 0.0)
        if dtfrac > 0:
            dtqty = round(qty * dtfrac)
            swqty = qty - dtqty
            if dtqty > 0:
                p.sell(date, dtqty, uprice, total * dtfrac, "daytrade")
                monthly[month][ticker]["daytrade"] += total * dtfrac
            if swqty > 0:
                p.sell(date, swqty, uprice, total * (1 - dtfrac), "swing")
                monthly[month][ticker]["swing"] += total * (1 - dtfrac)
        else:
            p.sell(date, qty, uprice, total, "swing")
            monthly[month][ticker]["swing"] += total

    # ── Phase 1: Trades ──────────────────────────────────────────────────────
    if negocios_df is not None:
        for _, row in negocios_df.iterrows():
            p         = pos(row["Ticker"])
            month     = row["Month"]
            date      = row["Data"]
            qty       = int(row["Qty"])
            uprice    = float(row["UnitPrice"])
            total     = float(row["OpValue"]) if pd.notna(row["OpValue"]) else qty * uprice
            direction = row["Direction"]

            if direction == "Compra":
                p.buy(qty, total)
            elif direction == "Venda":
                _sell(p, month, row["Ticker"], date, qty, uprice, total)
    else:
        for _, row in df.iterrows():
            p         = pos(row["Ticker"])
            month     = row["Month"]
            date      = row["Data"]
            qty       = int(row["Qty"])
            uprice    = row["UnitPrice"]
            opval     = row["OpValue"]
            direction = row["Direction"]
            mov       = row["Mov"]
            uprice_ok = pd.notna(uprice)
            opval_ok  = pd.notna(opval)

            if mov in _SALE_MOVS and uprice_ok and qty > 0:
                total = float(opval) if opval_ok else qty * float(uprice)
                if direction == "Credito":
                    p.buy(qty, total)
                elif direction == "Debito":
                    _sell(p, month, row["Ticker"], date, qty, float(uprice), total)

    # ── Phase 2: Income & corporate actions (always from movimentacao) ───────
    for _, row in df.iterrows():
        p         = pos(row["Ticker"])
        qty       = int(row["Qty"])
        opval     = row["OpValue"]
        direction = row["Direction"]
        mov       = row["Mov"]
        opval_ok  = pd.notna(opval)

        if mov in ("Dividendo", "Juros Sobre Capital Próprio", "Rendimento", "Reembolso"):
            if direction == "Credito" and opval_ok:
                p.income[mov] += float(opval)

        elif mov == "Amortização" and direction == "Credito" and opval_ok:
            p.total_cost = max(0.0, p.total_cost - float(opval))
            p.income["Amortização"] += float(opval)

        elif mov == "Bonificação em Ativos" and direction == "Credito" and qty > 0:
            p.buy(qty, float(opval) if opval_ok else 0.0)

        elif mov == "Desdobramento" and direction == "Credito" and qty > 0:
            p.qty += qty

        elif mov == "Grupamento" and direction == "Debito" and qty > 0:
            p.qty = max(0, p.qty - qty)

        elif mov == "Empréstimo" and direction == "Credito" and opval_ok:
            p.income["Empréstimo (taxa)"] += float(opval)

        elif mov in ("Imposto de Renda Retido na Fonte", "IRRF", "IRRF Day Trade") \
                and direction == "Debito" and opval_ok:
            p.income["IRRF Retido"] += float(opval)

    return positions, monthly

# ── DataFrame builders ─────────────────────────────────────────────────────────

def _df(rows, columns=None):
    return pd.DataFrame(rows, columns=columns) if rows else pd.DataFrame(
        columns=columns if columns else []
    )


def df_positions(positions):
    cols = ["Ticker", "Tipo", "Qtd", "Preço Médio", "Custo Total"]
    rows = [
        dict(Ticker=t, Tipo=p.kind, Qtd=p.qty,
             **{"Preço Médio": brl4(p.avg_cost), "Custo Total": brl(p.total_cost)})
        for t, p in sorted(positions.items()) if p.qty > 0
    ]
    return _df(rows, cols)


def df_trades(positions):
    cols = ["Ticker", "Tipo", "Data", "Operação", "Qtd", "Preço",
            "Receita", "Custo", "P&L", "_pnl"]
    rows = []
    for t, p in sorted(positions.items()):
        for s in p.sells:
            rows.append({
                "Ticker": t, "Tipo": p.kind,
                "Data": s["date"].strftime("%d/%m/%Y"),
                "Operação": "Day Trade" if s["op"] == "daytrade" else "Swing",
                "Qtd": s["qty"], "Preço": brl4(s["price"]),
                "Receita": brl(s["revenue"]), "Custo": brl(s["cost_basis"]),
                "P&L": brl(s["pnl"]), "_pnl": s["pnl"],
            })
    return _df(rows, cols)


def df_annual_pnl(positions):
    cols = ["Ticker", "Tipo", "P&L Swing", "P&L Day Trade",
            "P&L Total Ano", "Preço Médio Final", "_total"]
    rows = []
    for t, p in sorted(positions.items()):
        sw = p.pnl["swing"]
        dt = p.pnl["daytrade"]
        total = sw + dt
        if sw != 0 or dt != 0:
            rows.append({
                "Ticker": t, "Tipo": p.kind,
                "P&L Swing": brl(sw), "P&L Day Trade": brl(dt),
                "P&L Total Ano": brl(total),
                "Preço Médio Final": brl4(p.avg_cost) if p.qty > 0 else "Zerado",
                "_total": total,
            })
    return _df(rows, cols)


def df_monthly_summary(monthly):
    cols = ["Mês", "Ticker", "Tipo", "Operação",
            "Venda", "Total AÇÃO/mês", "Situação", "Alíq."]
    rows = []
    for month in sorted(monthly):
        swing_acoes = sum(
            v["swing"] for t, v in monthly[month].items() if classify(t) == "AÇÃO"
        )
        for ticker in sorted(monthly[month]):
            kind = classify(ticker)
            for op, val in monthly[month][ticker].items():
                if val == 0:
                    continue
                if op == "daytrade":
                    sit, aliq = "Tributado", f"{TAX_DT[kind]*100:.0f}%"
                elif kind == "AÇÃO" and swing_acoes <= EXEMPT:
                    sit, aliq = "Isento", "0%"
                else:
                    sit, aliq = "Tributado", f"{TAX_SW[kind]*100:.0f}%"
                rows.append({
                    "Mês": month, "Ticker": ticker, "Tipo": kind,
                    "Operação": op.capitalize(), "Venda": brl(val),
                    "Total AÇÃO/mês": brl(swing_acoes),
                    "Situação": sit, "Alíq.": aliq,
                })
    return _df(rows, cols)


def df_income(positions, keys):
    cols = ["Ticker", "Tipo", "Descrição", "Valor", "_val"]
    rows = [
        {"Ticker": t, "Tipo": p.kind, "Descrição": k,
         "Valor": brl(p.income[k]), "_val": p.income[k]}
        for t, p in sorted(positions.items()) for k in keys if k in p.income
    ]
    return _df(rows, cols)


def df_renda_variavel_guide(positions, monthly):
    monthly_pnl = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for _, p in positions.items():
        for s in p.sells:
            m = s["date"].strftime("%Y-%m")
            monthly_pnl[m][p.kind][s["op"]] += s["pnl"]

    cols = ["Mês", "Tipo Ativo", "Operação", "Resultado",
            "Alíq.", "Imposto devido", "Cód. DARF", "_pnl", "_imposto"]
    rows = []
    for month in sorted(monthly_pnl):
        swing_acoes = sum(
            v["swing"] for t, v in monthly[month].items() if classify(t) == "AÇÃO"
        )
        for kind in sorted(monthly_pnl[month]):
            for op in sorted(monthly_pnl[month][kind]):
                pnl_val = monthly_pnl[month][kind][op]
                isento  = (op == "swing" and kind == "AÇÃO" and swing_acoes <= EXEMPT)
                aliq    = 0.0 if isento else (TAX_DT[kind] if op == "daytrade" else TAX_SW[kind])
                imposto = max(0.0, pnl_val * aliq)
                rows.append({
                    "Mês": month, "Tipo Ativo": kind,
                    "Operação": "Day Trade" if op == "daytrade" else "Swing",
                    "Resultado": brl(pnl_val),
                    "Alíq.": "Isento" if isento else f"{aliq*100:.0f}%",
                    "Imposto devido": "—" if isento or pnl_val <= 0 else brl(imposto),
                    "Cód. DARF": "—" if isento or pnl_val <= 0 else DARF_CODE,
                    "_pnl": pnl_val, "_imposto": imposto,
                })
    return _df(rows, cols)


def build_discriminacao(ticker, p):
    unit = {"AÇÃO": "ação(ões) ordinária(s)", "FII": "cota(s)",
            "ETF": "cota(s)", "BDR": "BDR(s)"}[p.kind]
    return (
        f"{p.qty} {unit} de {ticker}. "
        f"Negociadas na B3 S.A. - Brasil, Bolsa, Balcão. "
        f"Custodiante: {INSTITUTION} (CNPJ: {INSTITUTION_CNPJ}). "
        f"Preço médio de custo: R$ {p.avg_cost:,.4f}. "
        f"Custo total de aquisição: R$ {p.total_cost:,.2f}."
    )

# ── PDF rendering ──────────────────────────────────────────────────────────────

def _new_fig(title, subtitle=None):
    fig = plt.figure(figsize=(PW, PH))
    fig.patch.set_facecolor(C["bg"])
    fig.text(0.5, 0.97, title, ha="center", va="top",
             fontsize=13, fontweight="bold", color=C["title"])
    if subtitle:
        fig.text(0.5, 0.925, subtitle, ha="center", va="top",
                 fontsize=8.5, color="#555")
    return fig


def _draw_table(ax, df, pnl_col=None):
    display = df.drop(columns=[c for c in df.columns if c.startswith("_")], errors="ignore")
    n_rows, n_cols = display.shape
    cell_colors = [
        [C["even"] if i % 2 == 0 else C["odd"]] * n_cols
        for i in range(n_rows)
    ]
    tbl = ax.table(
        cellText=display.values.tolist(),
        colLabels=display.columns.tolist(),
        cellColours=cell_colors,
        loc="center", cellLoc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(7.5)
    tbl.auto_set_column_width(range(n_cols))

    for j in range(n_cols):
        cell = tbl[0, j]
        cell.set_facecolor(C["hbg"])
        cell.set_text_props(color=C["hfg"], fontweight="bold")
        cell.set_edgecolor(C["hbg"])

    text_cols = {j for j, col in enumerate(display.columns)
                 if display.dtypes.iloc[j] == object and col not in ("Qtd",)}
    pnl_idx = list(display.columns).index(pnl_col) if pnl_col in list(display.columns) else None

    for i in range(1, n_rows + 1):
        for j in range(n_cols):
            cell = tbl[i, j]
            cell.set_edgecolor(C["edge"])
            if j in text_cols:
                cell.get_text().set_ha("left")
        if pnl_idx is not None and "_pnl" in df.columns:
            raw = df.iloc[i - 1]["_pnl"]
            if pd.notna(raw):
                tbl[i, pnl_idx].get_text().set_color(
                    C["pos"] if float(raw) >= 0 else C["neg"]
                )


def add_pages(pdf, title, df, subtitle=None, pnl_col=None):
    if df is None or df.empty:
        fig = _new_fig(title, subtitle)
        ax  = fig.add_axes([0.03, 0.05, 0.94, 0.82])
        ax.axis("off")
        ax.text(0.5, 0.5, "Sem dados para este período.",
                ha="center", va="center", transform=ax.transAxes,
                fontsize=10, color="gray")
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)
        return

    chunks = [df.iloc[i:i+MAX_ROWS] for i in range(0, len(df), MAX_ROWS)]
    total  = len(chunks)
    for page_num, chunk in enumerate(chunks, 1):
        parts = []
        if subtitle:
            parts.append(subtitle)
        if total > 1:
            parts.append(f"pág. {page_num}/{total}")
        sub = "  —  ".join(parts) if parts else None

        fig = _new_fig(title, sub)
        ax  = fig.add_axes([0.03, 0.05, 0.94, 0.82])
        ax.axis("off")
        ax.set_facecolor(C["bg"])
        _draw_table(ax, chunk.reset_index(drop=True), pnl_col=pnl_col)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)


def cover_page(pdf, positions, monthly):
    fig = _new_fig("Relatório IRPF — Renda Variável B3")
    fig.text(0.5, 0.89, f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
             ha="center", fontsize=9, color="#666")

    total_sw  = sum(p.pnl["swing"]    for p in positions.values())
    total_dt  = sum(p.pnl["daytrade"] for p in positions.values())
    div_total = sum(p.income.get("Dividendo", 0) + p.income.get("Rendimento", 0)
                    for p in positions.values())
    jcp_total = sum(p.income.get("Juros Sobre Capital Próprio", 0)
                    for p in positions.values())
    n_pos     = sum(1 for p in positions.values() if p.qty > 0)
    n_traded  = len(positions)

    metrics = [
        ("P&L Swing",                brl(total_sw),  total_sw),
        ("P&L Day Trade",            brl(total_dt),  total_dt),
        ("Rend. Isentos (Div+FII)",  brl(div_total), div_total),
        ("JCP (tributável)",         brl(jcp_total), None),
        ("Posições em carteira",     str(n_pos),     None),
        ("Ativos negociados",        str(n_traded),  None),
    ]

    ax = fig.add_axes([0.10, 0.10, 0.80, 0.68])
    ax.axis("off")

    box_w, box_h = 0.29, 0.35
    coords = [(0.02, 0.50), (0.36, 0.50), (0.70, 0.50),
              (0.02, 0.05), (0.36, 0.05), (0.70, 0.05)]

    for (label, value, sign), (bx, by) in zip(metrics, coords):
        color = (C["pos"] if sign is not None and sign >= 0
                 else (C["neg"] if sign is not None else C["title"]))
        rect = plt.Rectangle(
            (bx, by), box_w, box_h,
            facecolor="#dde5f7", edgecolor=C["hbg"], linewidth=1.2,
            transform=ax.transAxes, clip_on=False,
        )
        ax.add_patch(rect)
        ax.text(bx + box_w / 2, by + box_h * 0.70, label,
                ha="center", va="center", fontsize=8.5,
                color="#444", transform=ax.transAxes)
        ax.text(bx + box_w / 2, by + box_h * 0.28, value,
                ha="center", va="center", fontsize=12, fontweight="bold",
                color=color, transform=ax.transAxes)

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def irpf_bens_page(pdf, positions):
    entries = [(t, p) for t, p in sorted(positions.items()) if p.qty > 0]
    if not entries:
        return

    items_per_page = 5
    chunks = [entries[i:i+items_per_page] for i in range(0, len(entries), items_per_page)]
    total  = len(chunks)

    for page_num, chunk in enumerate(chunks, 1):
        sub = "Copie a Discriminação no campo correspondente em Bens e Direitos"
        if total > 1:
            sub += f"  —  pág. {page_num}/{total}"
        fig = _new_fig("Guia IRPF — Bens e Direitos (Discriminações)", sub)
        ax  = fig.add_axes([0.04, 0.04, 0.92, 0.82])
        ax.axis("off")

        y = 0.99
        for t, p in chunk:
            disc  = build_discriminacao(t, p)
            grupo = IRPF_GROUP[p.kind]
            code  = IRPF_CODE[p.kind]
            header = (
                f"{t}  ({p.kind})   ·   Grupo {grupo} / Código {code}"
                f"   ·   Situação em 31/12: {brl(p.total_cost)}"
            )
            ax.text(0, y, header, fontsize=9, fontweight="bold",
                    color=C["title"], transform=ax.transAxes, va="top")
            y -= 0.04

            for line in textwrap.wrap(disc, width=115):
                ax.text(0.015, y, line, fontsize=8, color="#222",
                        transform=ax.transAxes, va="top", family="monospace")
                y -= 0.032

            ax.plot([0, 1], [y + 0.012, y + 0.012],
                    color=C["edge"], linewidth=0.7,
                    transform=ax.transAxes, clip_on=False)
            y -= 0.025

        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)


def irpf_rv_guide_page(pdf, rv_df):
    fig = _new_fig(
        "Guia IRPF — Renda Variável (Resultado Mensal)",
        "Preencha na aba 'Renda Variável' do programa IRPF. "
        "Pague DARF código 6015 até o último dia útil do mês seguinte."
    )
    ax = fig.add_axes([0.03, 0.05, 0.94, 0.78])
    ax.axis("off")
    ax.set_facecolor(C["bg"])

    display = rv_df.drop(columns=["_pnl", "_imposto"], errors="ignore")

    if display.empty:
        ax.text(0.5, 0.5, "Sem operações tributáveis.", ha="center", va="center",
                transform=ax.transAxes, color="gray")
    else:
        n_rows, n_cols = display.shape
        cell_colors = [
            [C["even"] if i % 2 == 0 else C["odd"]] * n_cols
            for i in range(n_rows)
        ]
        tbl = ax.table(
            cellText=display.values.tolist(),
            colLabels=display.columns.tolist(),
            cellColours=cell_colors,
            loc="center", cellLoc="center",
        )
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(7.5)
        tbl.auto_set_column_width(range(n_cols))
        for j in range(n_cols):
            tbl[0, j].set_facecolor(C["hbg"])
            tbl[0, j].set_text_props(color=C["hfg"], fontweight="bold")
        pnl_idx = list(display.columns).index("Resultado") if "Resultado" in display.columns else None
        for i in range(1, n_rows + 1):
            for j in range(n_cols):
                tbl[i, j].set_edgecolor(C["edge"])
            if pnl_idx is not None:
                raw = rv_df.iloc[i - 1]["_pnl"]
                tbl[i, pnl_idx].get_text().set_color(
                    C["pos"] if raw >= 0 else C["neg"]
                )

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def notes_page(pdf):
    fig = _new_fig("Notas e Premissas")
    ax  = fig.add_axes([0.06, 0.04, 0.88, 0.84])
    ax.axis("off")

    notes = [
        ("Preço Médio Ponderado",
         "Calculado pelo método da média ponderada (custo total ÷ quantidade). "
         "Ajustado em cada compra, amortização, bonificação, desdobramento e grupamento."),
        ("Isenção mensal — Ações (swing)",
         "Vendas de ações em swing trade são isentas de IR quando o total de vendas "
         "do mês ≤ R$ 20.000 (art. 3º Lei 11.033/2004). ETF, FII e BDR não têm essa isenção."),
        ("FII — Ganho de Capital",
         "Tributado à alíquota de 20% (art. 2º Lei 13.043/2014). Sem isenção mensal. "
         "Rendimentos (cotas mensais) são isentos para PF (IN RFB 1.585/2015)."),
        ("ETF / BDR — Ganho de Capital",
         "Tributado à alíquota de 15% em operações swing. Sem isenção mensal."),
        ("Day Trade",
         "Compra e venda do mesmo ativo no mesmo dia = day trade. "
         "Tributado a 20% independente do ativo ou volume. Sem isenção mensal."),
        ("JCP — Juros Sobre Capital Próprio",
         "IRRF de 15% retido na fonte. Declarar em 'Rendimentos Sujeitos à Tributação Exclusiva', "
         "código 10. O valor bruto aparece neste relatório."),
        ("Fracionário",
         "Tickers terminados em 'F' (ex: PETR4F) são normalizados para o lote padrão (PETR4)."),
        ("Amortização",
         "Reduz o custo médio de aquisição automaticamente. O valor recebido aparece em Outros Rend."),
        ("Classificação FII vs ETF",
         "Lista KNOWN_ETFS define os ETFs conhecidos. Qualquer '11' fora dela = FII (alíq. 20%). "
         "Adicione tickers corretos em KNOWN_ETFS no topo do script se necessário."),
        ("IRRF / Dedo Duro",
         "0,005% retido na fonte sobre o valor bruto das vendas. Pode ser compensado no DARF mensal."),
        (f"Ano de referência",
         f"Dados filtrados até 31/12/{DECLARACAO_ANO}. Altere DECLARACAO_ANO no topo do script "
         f"para declarações de outros anos."),
    ]

    y = 0.98
    for title, body in notes:
        ax.text(0.0, y, f"• {title}:", fontsize=8.5, fontweight="bold",
                color=C["title"], transform=ax.transAxes, va="top")
        y -= 0.044
        for line in textwrap.wrap(body, width=120):
            ax.text(0.02, y, line, fontsize=7.8, color="#333",
                    transform=ax.transAxes, va="top")
            y -= 0.032
        y -= 0.016
        if y < 0.04:
            break

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)

# ── Console declaração text ────────────────────────────────────────────────────

def print_declaracao_text(positions, monthly):
    SEP  = "=" * 72
    LINE = "-" * 72

    print(f"\n{SEP}")
    print(" DADOS PARA DECLARACAO IRPF - COPIE E COLE")
    print(SEP)

    # 1. Bens e Direitos
    print("\n-- BENS E DIREITOS")
    print(LINE)
    held = [(t, p) for t, p in sorted(positions.items()) if p.qty > 0]
    if held:
        for t, p in held:
            grupo = IRPF_GROUP[p.kind]
            code  = IRPF_CODE[p.kind]
            print(f"  {t} ({p.kind}) - Grupo {grupo} / Codigo {code} "
                  f"- Situacao em 31/12: {brl(p.total_cost)}")
            print(f"  {build_discriminacao(t, p)}")
            print()
    else:
        print("  Sem posicoes em aberto.\n")

    # 2. Rendimentos Isentos - isencao acoes swing R$20k
    print("-- RENDIMENTOS ISENTOS E NAO TRIBUTAVEIS")
    print(LINE)
    exempt_pnl = 0.0
    for month in sorted(monthly):
        swing_vol = sum(
            v["swing"] for t, v in monthly[month].items() if classify(t) == "AÇÃO"
        )
        if swing_vol <= EXEMPT:
            for _, p in positions.items():
                if p.kind != "AÇÃO":
                    continue
                for s in p.sells:
                    if s["date"].strftime("%Y-%m") == month and s["op"] == "swing":
                        exempt_pnl += s["pnl"]
    print(f"  Ganho isento c/ acoes swing (vendas mensais <= R$ 20.000): "
          f"{brl(max(0.0, exempt_pnl))}")
    if exempt_pnl < 0:
        print(f"  (meses isentos tiveram prejuizo de {brl(exempt_pnl)} - nao dedutivel aqui)")
    print("  -> Rendimentos Isentos, codigo 20\n")

    # 3. Dividendos (isentos)
    print("-- DIVIDENDOS  (Isentos - Rendimentos Isentos)")
    print(LINE)
    total_div = 0.0
    for t, p in sorted(positions.items()):
        div = p.income.get("Dividendo", 0.0) + p.income.get("Rendimento", 0.0)
        if div > 0:
            total_div += div
            label = ("Rendimento FII"
                     if p.income.get("Rendimento", 0) > p.income.get("Dividendo", 0)
                     else "Dividendo")
            print(f"  {t:<12} {brl(div):>12}  [{label}]")
    if total_div > 0:
        print(f"  {'Total':<12} {brl(total_div):>12}")
        print("  -> Rendimentos Isentos - Dividendos recebidos no Pais")
    else:
        print("  Nenhum dividendo/rendimento registrado.")
    print()

    # 4. JCP
    print("-- JCP - Juros Sobre Capital Proprio  (Tributacao Exclusiva na Fonte)")
    print(LINE)
    total_jcp = 0.0
    for t, p in sorted(positions.items()):
        jcp = p.income.get("Juros Sobre Capital Próprio", 0.0)
        if jcp > 0:
            total_jcp += jcp
            print(f"  {t:<12} bruto {brl(jcp):>12}   IRRF retido (15%): {brl(jcp * 0.15)}")
    if total_jcp > 0:
        print(f"  {'Total bruto':<12} {brl(total_jcp):>12}")
        print("  -> Tributacao Exclusiva na Fonte - codigo 10")
    else:
        print("  Nenhum JCP registrado.")
    print()

    # 5. Renda Variavel - resultado tributavel por mes
    print("-- RENDA VARIAVEL - Resultado tributavel (aba 'Renda Variavel')")
    print(LINE)
    print("  FII: 20% | ETF/BDR/ACAO>20k: 15% | Day Trade: 20% | DARF cod. 6015")
    print()

    monthly_pnl = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for _, p in positions.items():
        for s in p.sells:
            m = s["date"].strftime("%Y-%m")
            monthly_pnl[m][p.kind][s["op"]] += s["pnl"]

    any_row = False
    for month in sorted(monthly_pnl):
        swing_vol = sum(
            v["swing"] for t, v in monthly[month].items() if classify(t) == "AÇÃO"
        )
        for kind in ("AÇÃO", "FII", "ETF", "BDR"):
            for op in ("swing", "daytrade"):
                pnl_val = monthly_pnl[month].get(kind, {}).get(op, 0.0)
                if pnl_val == 0.0:
                    continue
                isento = (op == "swing" and kind == "AÇÃO" and swing_vol <= EXEMPT)
                if isento:
                    continue
                aliq    = TAX_DT[kind] if op == "daytrade" else TAX_SW[kind]
                imposto = max(0.0, pnl_val * aliq)
                op_tag  = "DayTrade" if op == "daytrade" else "Swing   "
                nota    = (f"DARF: {brl(imposto)}"
                           if pnl_val > 0
                           else "Prejuizo (acumule para compensacao futura)")
                print(f"  {month}  {kind:<5}  {op_tag}  {brl(pnl_val):>14}   {nota}")
                any_row = True

    if not any_row:
        print("  Nenhum resultado tributavel no periodo.")

    print(f"\n{SEP}\n")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    csv_file = find_csv()
    print(f"Lendo movimentacao: {csv_file}")
    df = load_df(csv_file)

    xlsx_file   = find_xlsx()
    negocios_df = None
    if xlsx_file:
        print(f"Lendo negociacao:   {xlsx_file}")
        negocios_df = load_negocios(xlsx_file)
    else:
        print("negociacao.xlsx nao encontrado — usando movimentacao para trades")
        print("(emprestimos podem inflar posicoes; exporte negociacao.xlsx do B3 para resultado correto)")

    positions, monthly = process(df, negocios_df)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUT_DIR / "relatorio_irpf.pdf"

    rv_guide = df_renda_variavel_guide(positions, monthly)
    trades   = df_trades(positions)
    annual   = df_annual_pnl(positions)

    with PdfPages(str(out_path)) as pdf:
        cover_page(pdf, positions, monthly)
        add_pages(pdf, "Posição Atual — Bens e Direitos",
                  df_positions(positions))
        add_pages(pdf, "Resumo Anual por Ativo",
                  annual.drop(columns=["_total"], errors="ignore"),
                  pnl_col="P&L Total Ano")
        add_pages(pdf, "Operações Realizadas — Renda Variável",
                  trades, pnl_col="P&L")
        add_pages(pdf, "Resumo Mensal — Isenção e Tributação",
                  df_monthly_summary(monthly))

        df_isento = df_income(positions, ("Dividendo", "Rendimento"))
        total_isento = df_isento["_val"].sum() if not df_isento.empty else 0.0
        add_pages(pdf, "Rendimentos Isentos — Dividendos e FII Rendimentos",
                  df_isento.drop(columns=["_val"], errors="ignore"),
                  subtitle=f"Total: {brl(total_isento)}")

        df_jcp = df_income(positions, ("Juros Sobre Capital Próprio",))
        total_jcp = df_jcp["_val"].sum() if not df_jcp.empty else 0.0
        add_pages(pdf, "Rendimentos Tributáveis — JCP (IRRF 15% retido na fonte)",
                  df_jcp.drop(columns=["_val"], errors="ignore"),
                  subtitle=f"Total bruto: {brl(total_jcp)}  |  IRRF estimado: {brl(total_jcp * 0.15)}")

        df_outros = df_income(positions, ("Empréstimo (taxa)", "Reembolso", "Amortização"))
        total_outros = df_outros["_val"].sum() if not df_outros.empty else 0.0
        add_pages(pdf, "Outros Rendimentos — Empréstimo BTC, Reembolsos, Amortizações",
                  df_outros.drop(columns=["_val"], errors="ignore"),
                  subtitle=f"Total: {brl(total_outros)}")

        df_irrf = df_income(positions, ("IRRF Retido",))
        total_irrf = df_irrf["_val"].sum() if not df_irrf.empty else 0.0
        add_pages(pdf, "IRRF Retido na Fonte — Dedo Duro",
                  df_irrf.drop(columns=["_val"], errors="ignore"),
                  subtitle=f"Total retido (compensável no DARF): {brl(total_irrf)}")

        irpf_bens_page(pdf, positions)
        irpf_rv_guide_page(pdf, rv_guide)
        notes_page(pdf)

        meta = pdf.infodict()
        meta["Title"]   = "Relatório IRPF — Renda Variável B3"
        meta["Author"]  = "irpf-helper"
        meta["Subject"] = "Declaração IR — Renda Variável"

    print(f"PDF gerado: {out_path}")
    print_declaracao_text(positions, monthly)


if __name__ == "__main__":
    main()
