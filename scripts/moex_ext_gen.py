#!/usr/bin/env python3
"""Generate db/zzz_moex_ext_after_init.sql — NEW instruments for the moex-finance MCP.

Plan B: the 11 tasks still on yahoo-finance need instruments the 6-ticker moex seed
lacks: a GOLD commodity, the IMOEX index, and foreign-stock analogs (OZON~Amazon,
YNDX~Google, VTBR~JPM, ROSN~Exxon, PHOR~J&J). The moex MCP is PG-backed with NO
symbol whitelist (pg_adapter resolves any symbol via upper() lookup), so adding an
instrument = pure seed rows — zero server code changes.

We DO NOT touch the existing 6 tickers (would risk the ~50 working moex tasks). New
instruments are seeded over a LONG window 2025-03-03..2026-05-26 (covers the 6M/1Y/
12-month lookbacks several tasks need) into a SEPARATE file mounted after zzz_moex.

Design properties (internally consistent; per-task agents regenerate frozen GT +
hardcoded eval literals from the LIVE extended DB, so we need realism + relations,
not legacy yf numbers):
  - GOLD rises strongly (~+27%): wins vs index in gold-vs-stocks; last close ~5093 RUB/g
    (in the 4000-6000 sanity band some evals use).
  - IMOEX rises mildly (~+7%) so gold beats the index.
  - equity revenue/margin ordering: OZON highest revenue (Amazon role), YNDX highest
    margin (Google role), others lower — for yf-financial-metrics.

Run: uv run --with psycopg2-binary python3 scripts/moex_ext_gen.py   (writes SQL only)
"""
import json, math, datetime as dt

OUT = "db/zzz_moex_ext_after_init.sql"
START = dt.date(2025, 3, 3)
END = dt.date(2026, 5, 26)


def business_days(a, b):
    d, out = a, []
    while d <= b:
        if d.weekday() < 5:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


DAYS = business_days(START, END)
N = len(DAYS)


def path(start, end, amp_pct=0.03, waves=3.0):
    """Deterministic smooth close path start->end with mild sinusoidal wobble."""
    out = []
    for i in range(N):
        t = i / (N - 1)
        trend = start + (end - start) * t
        wob = trend * amp_pct * math.sin(t * waves * 2 * math.pi)
        out.append(round(trend + wob, 4))
    return out


def ohlcv(closes, vol):
    """Build OHLCV rows from a close path (open=prev close; high/low ±0.6%)."""
    rows = []
    prev = closes[0]
    for i, c in enumerate(closes):
        o = prev
        hi = round(max(o, c) * 1.006, 4)
        lo = round(min(o, c) * 0.994, 4)
        rows.append((o, hi, lo, c, vol))
        prev = c
    return rows


# ---- instrument definitions ------------------------------------------------
# symbol, start_close, end_close, volume, amp, info(dict partial -> merged with base)
EQUITY_BASE = {
    "currency": "RUB", "exchange": "MCX", "fullExchangeName": "MOEX",
    "quoteType": "EQUITY", "country": "Russia", "financialCurrency": "RUB",
}

INSTRUMENTS = [
    # GOLD commodity (RUB per gram), strong rise, ends ~5093
    dict(symbol="GLDRUB_TOM", start=4010.0, end=5093.30, vol=0, amp=0.025, waves=2.0,
         info={"shortName": "Gold RUB/gram (MOEX)", "longName": "MOEX Exchange Gold (RUB per gram)",
               "currency": "RUB", "quoteType": "COMMODITY", "exchange": "MCX",
               "fullExchangeName": "MOEX", "sector": "Commodity", "industry": "Precious Metals"}),
    # IMOEX index, mild rise (gold beats it)
    dict(symbol="IMOEX.ME", start=2850.0, end=3052.0, vol=0, amp=0.02, waves=3.0,
         info={"shortName": "MOEX Russia Index", "longName": "MOEX Russia Index (IMOEX)",
               "currency": "points", "quoteType": "INDEX", "exchange": "MCX",
               "fullExchangeName": "MOEX", "sector": "Index", "industry": "Market Index"}),
    # OZON ~ Amazon (e-commerce, HIGHEST revenue)
    dict(symbol="OZON.ME", start=3520.0, end=4180.0, vol=2500000, amp=0.04, waves=3.5,
         info={"shortName": "Ozon Holdings", "longName": "Ozon Holdings PLC",
               "sector": "Consumer Cyclical", "industry": "Internet Retail",
               "industryDisp": "Internet Retail", "sectorDisp": "Consumer Cyclical"}),
    # YNDX ~ Alphabet/Google (tech, HIGHEST margin)
    dict(symbol="YNDX.ME", start=4015.0, end=4620.0, vol=1800000, amp=0.035, waves=3.0,
         info={"shortName": "Yandex", "longName": "Yandex N.V.",
               "sector": "Communication Services", "industry": "Internet Content & Information",
               "industryDisp": "Internet Content & Information", "sectorDisp": "Communication Services"}),
    # VTBR ~ JPMorgan (bank)
    dict(symbol="VTBR.ME", start=92.0, end=108.5, vol=9000000, amp=0.03, waves=3.5,
         info={"shortName": "VTB Bank", "longName": "VTB Bank PJSC",
               "sector": "Financial Services", "industry": "Banks - Regional",
               "industryDisp": "Banks - Regional", "sectorDisp": "Financial Services"}),
    # ROSN ~ Exxon (energy)
    dict(symbol="ROSN.ME", start=545.0, end=618.0, vol=3000000, amp=0.03, waves=3.0,
         info={"shortName": "Rosneft", "longName": "Rosneft Oil Company PJSC",
               "sector": "Energy", "industry": "Oil & Gas Integrated",
               "industryDisp": "Oil & Gas Integrated", "sectorDisp": "Energy"}),
    # PHOR ~ Johnson&Johnson (defensive / materials, LOWEST revenue)
    dict(symbol="PHOR.ME", start=6550.0, end=7180.0, vol=120000, amp=0.025, waves=2.5,
         info={"shortName": "PhosAgro", "longName": "PhosAgro PJSC",
               "sector": "Basic Materials", "industry": "Agricultural Inputs",
               "industryDisp": "Agricultural Inputs", "sectorDisp": "Basic Materials"}),
]

# financial_statements for equities (annual, RUB). Revenue ordering: OZON>YNDX>ROSN>VTBR>PHOR;
# margin (NI/Rev): YNDX highest. Two periods 2024-12-31, 2025-12-31.
FIN = {
    # symbol: (rev2024, ni2024, assets2024, rev2025, ni2025, assets2025)
    "OZON.ME": (716_920_000_000, 28_700_000_000, 818_040_000_000, 812_400_000_000, 41_200_000_000, 905_300_000_000),
    "YNDX.ME": (402_840_000_000, 132_170_000_000, 595_280_000_000, 471_900_000_000, 158_600_000_000, 651_100_000_000),
    "VTBR.ME": (312_500_000_000, 48_300_000_000, 980_400_000_000, 351_200_000_000, 55_900_000_000, 1_044_000_000_000),
    "ROSN.ME": (298_100_000_000, 41_600_000_000, 512_700_000_000, 320_800_000_000, 47_300_000_000, 548_900_000_000),
    "PHOR.ME": (94_190_000_000, 26_800_000_000, 199_210_000_000, 101_500_000_000, 29_400_000_000, 214_600_000_000),
}


def q(s):
    return "'" + s.replace("'", "''") + "'"


def jdump(d):
    return q(json.dumps(d, ensure_ascii=False))


def main():
    lines = [
        "-- AUTO-GENERATED by scripts/moex_ext_gen.py — moex-finance EXTENSION (Plan B).",
        "-- NEW instruments only (gold/IMOEX/OZON/YNDX/VTBR/ROSN/PHOR) over 2025-03-03..2026-05-26.",
        "-- Does NOT touch the existing 6 tickers. Mounted after zzz_moex. Idempotent via ON CONFLICT.",
        "",
    ]
    closes_by = {}
    for ins in INSTRUMENTS:
        sym = ins["symbol"]
        closes = path(ins["start"], ins["end"], ins["amp"], ins["waves"])
        closes_by[sym] = closes
        rows = ohlcv(closes, ins["vol"])
        last = closes[-1]
        info = dict(EQUITY_BASE)
        info.update(ins["info"])
        info.update({
            "symbol": sym, "regularMarketPrice": last, "currentPrice": last,
            "previousClose": closes[-2], "regularMarketPreviousClose": closes[-2],
            "fiftyTwoWeekHigh": round(max(closes), 4), "fiftyTwoWeekLow": round(min(closes), 4),
        })
        if info.get("quoteType") == "EQUITY":
            shares = 200_000_000
            info.update({"marketCap": int(last * shares), "sharesOutstanding": shares,
                         "trailingEps": round(FIN[sym][1] / shares, 4) if sym in FIN else None})
        lines.append(f"-- {sym}: {N} daily bars, close {ins['start']}->{ins['end']}")
        lines.append(f"INSERT INTO moex.stock_info (symbol, data) VALUES ({q(sym)}, {jdump(info)}::jsonb) "
                     f"ON CONFLICT (symbol) DO UPDATE SET data = EXCLUDED.data;")
        for d, (o, hi, lo, c, v) in zip(DAYS, rows):
            lines.append(
                f"INSERT INTO moex.stock_prices (symbol, date, open, high, low, close, volume) "
                f"VALUES ({q(sym)}, '{d.isoformat()}', {o}, {hi}, {lo}, {c}, {v}) "
                f"ON CONFLICT (symbol, date) DO UPDATE SET open=EXCLUDED.open, high=EXCLUDED.high, "
                f"low=EXCLUDED.low, close=EXCLUDED.close, volume=EXCLUDED.volume;")
        lines.append("")

    # financial statements (income + balance sheet) for equities, two annual periods
    for sym, (r24, n24, a24, r25, n25, a25) in FIN.items():
        shares = 200_000_000
        for pe, rev, ni, assets in [("2024-12-31", r24, n24, a24), ("2025-12-31", r25, n25, a25)]:
            inc = {"Total Revenue": float(rev), "Net Income": float(ni),
                   "Net Income Common Stockholders": float(ni),
                   "Gross Profit": round(rev * 0.42, 2), "Operating Income": round(ni * 1.25, 2),
                   "Diluted EPS": round(ni / shares, 4), "Basic EPS": round(ni / shares, 4)}
            bal = {"Total Assets": float(assets),
                   "Total Liabilities Net Minority Interest": round(assets * 0.55, 2),
                   "Stockholders Equity": round(assets * 0.45, 2),
                   "Total Debt": round(assets * 0.30, 2), "Cash And Cash Equivalents": round(assets * 0.12, 2)}
            lines.append(
                f"INSERT INTO moex.financial_statements (symbol, period_end, stmt_type, freq, data) "
                f"VALUES ({q(sym)}, '{pe}', 'income_stmt', 'annual', {jdump(inc)}::jsonb) "
                f"ON CONFLICT (symbol, period_end, stmt_type, freq) DO UPDATE SET data=EXCLUDED.data;")
            lines.append(
                f"INSERT INTO moex.financial_statements (symbol, period_end, stmt_type, freq, data) "
                f"VALUES ({q(sym)}, '{pe}', 'balance_sheet', 'annual', {jdump(bal)}::jsonb) "
                f"ON CONFLICT (symbol, period_end, stmt_type, freq) DO UPDATE SET data=EXCLUDED.data;")
    lines.append("")

    with open(OUT, "w") as f:
        f.write("\n".join(lines) + "\n")

    # report key derived values for sanity
    print(f"Wrote {OUT}: {N} bars/instrument x {len(INSTRUMENTS)} instruments")
    g = closes_by["GLDRUB_TOM"]; idx = closes_by["IMOEX.ME"]
    gret = (g[-1] / g[0] - 1) * 100; iret = (idx[-1] / idx[0] - 1) * 100
    print(f"  GOLD {g[0]}->{g[-1]} ({gret:+.1f}%) | IMOEX {idx[0]}->{idx[-1]} ({iret:+.1f}%) | gold_wins={gret>iret}")
    print(f"  gold 30d avg≈{round(sum(g[-30:])/30,2)}, last={g[-1]}, in[4000,6000]={4000<g[-1]<6000}")
    for s in ("OZON.ME","YNDX.ME","ROSN.ME","VTBR.ME","PHOR.ME"):
        r = FIN[s]; print(f"  {s}: rev2025={r[3]/1e9:.1f}B margin={r[4]/r[3]*100:.1f}%")


if __name__ == "__main__":
    main()
