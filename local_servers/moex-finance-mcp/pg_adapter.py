"""
PostgreSQL adapter that mimics the yfinance module interface.
Reads from moex.* tables in PostgreSQL instead of calling Yahoo Finance API.
"""

import os
from collections import namedtuple
from datetime import datetime, timedelta

import pandas as pd
import psycopg2
import psycopg2.extras


def _get_connection():
    """Create a new PostgreSQL connection using environment variables."""
    return psycopg2.connect(
        host=os.environ.get("PG_HOST", "localhost"),
        port=int(os.environ.get("PG_PORT", "5432")),
        database=os.environ.get("PG_DATABASE", "cowork_gym"),
        user=os.environ.get("PG_USER", "postgres"),
        password=os.environ.get("PG_PASSWORD", "postgres"),
    )


def _query(sql, params=None):
    """Execute a query and return rows as list of dicts."""
    conn = _get_connection()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    finally:
        conn.close()


def _query_single(sql, params=None):
    """Execute a query and return a single row as dict, or None."""
    rows = _query(sql, params)
    return dict(rows[0]) if rows else None


def _period_to_timedelta(period: str) -> timedelta | None:
    """Convert a yfinance period string to a timedelta."""
    mapping = {
        "1d": timedelta(days=1),
        "5d": timedelta(days=5),
        "1mo": timedelta(days=30),
        "3mo": timedelta(days=90),
        "6mo": timedelta(days=182),
        "1y": timedelta(days=365),
        "2y": timedelta(days=730),
        "5y": timedelta(days=1825),
        "10y": timedelta(days=3650),
        "ytd": None,  # handled separately
        "max": timedelta(days=36500),
    }
    return mapping.get(period)


class OptionChain:
    """Mimics yfinance option_chain return value."""

    def __init__(self, calls_df: pd.DataFrame, puts_df: pd.DataFrame):
        self.calls = calls_df
        self.puts = puts_df


class PgTicker:
    """Mimics yfinance.Ticker interface, backed by PostgreSQL."""

    def __init__(self, symbol: str):
        self.ticker = symbol.upper()
        self._info_cache = None

    @property
    def isin(self):
        """Check if the ticker exists in the database. Returns a string or None."""
        row = _query_single(
            "SELECT symbol FROM moex.stock_info WHERE symbol = %s", (self.ticker,)
        )
        if row:
            return row["symbol"]
        # Also check stock_prices as a fallback
        row = _query_single(
            "SELECT symbol FROM moex.stock_prices WHERE symbol = %s LIMIT 1",
            (self.ticker,),
        )
        if row:
            return row["symbol"]
        return None

    @property
    def info(self) -> dict:
        """Return stock info as a dict from moex.stock_info JSONB data column."""
        if self._info_cache is not None:
            return self._info_cache
        row = _query_single(
            "SELECT data FROM moex.stock_info WHERE symbol = %s", (self.ticker,)
        )
        info = dict(row["data"]) if row and row["data"] else {}
        # Align price fields with stock_prices (single source of truth)
        last_rows = _query(
            "SELECT open, high, low, close, volume FROM moex.stock_prices WHERE symbol = %s ORDER BY date DESC LIMIT 2",
            (self.ticker,),
        )
        if last_rows:
            cur = last_rows[0]
            prev = last_rows[1] if len(last_rows) > 1 else last_rows[0]
            if cur["close"] is not None:
                info["currentPrice"] = float(cur["close"])
                info["regularMarketPrice"] = float(cur["close"])
            if prev["close"] is not None:
                info["previousClose"] = float(prev["close"])
                info["regularMarketPreviousClose"] = float(prev["close"])
            if cur["open"] is not None:
                info["regularMarketOpen"] = float(cur["open"])
            if cur["high"] is not None:
                info["regularMarketDayHigh"] = float(cur["high"])
            if cur["low"] is not None:
                info["regularMarketDayLow"] = float(cur["low"])
            if cur["volume"] is not None:
                info["regularMarketVolume"] = int(cur["volume"])
        self._info_cache = info
        return self._info_cache

    def history(
        self,
        period: str = "1mo",
        interval: str = "1d",
        start=None,
        end=None,
        auto_adjust: bool = True,
    ) -> pd.DataFrame:
        """Return historical price data as a DataFrame from moex.stock_prices."""
        conditions = ["symbol = %s"]
        params: list = [self.ticker]

        if start is not None:
            if isinstance(start, str):
                start = pd.to_datetime(start)
            conditions.append("date >= %s")
            params.append(start)
        if end is not None:
            if isinstance(end, str):
                end = pd.to_datetime(end)
            conditions.append("date <= %s")
            params.append(end)

        # If no start/end provided, use period anchored to the symbol's latest date
        if start is None and end is None:
            ref_row = _query_single(
                "SELECT MAX(date) AS max_date FROM moex.stock_prices WHERE symbol = %s",
                (self.ticker,),
            )
            ref_date = (
                ref_row["max_date"]
                if ref_row and ref_row["max_date"] is not None
                else datetime.now().date()
            )
            if period == "ytd":
                year_start = datetime(ref_date.year, 1, 1).date()
                conditions.append("date >= %s")
                params.append(year_start)
            elif period != "max":
                td = _period_to_timedelta(period)
                if td is not None:
                    cutoff = ref_date - td
                    conditions.append("date >= %s")
                    params.append(cutoff)

        where = " AND ".join(conditions)
        sql = f"SELECT date, open, high, low, close, volume, dividends, stock_splits FROM moex.stock_prices WHERE {where} ORDER BY date ASC"

        rows = _query(sql, params)

        if not rows:
            return pd.DataFrame(
                columns=["Open", "High", "Low", "Close", "Volume", "Dividends", "Stock Splits"]
            )

        df = pd.DataFrame(rows)
        df = df.rename(
            columns={
                "date": "Date",
                "open": "Open",
                "high": "High",
                "low": "Low",
                "close": "Close",
                "volume": "Volume",
                "dividends": "Dividends",
                "stock_splits": "Stock Splits",
            }
        )

        # Convert numeric columns
        for col in ["Open", "High", "Low", "Close", "Dividends", "Stock Splits"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        if "Volume" in df.columns:
            df["Volume"] = pd.to_numeric(df["Volume"], errors="coerce").astype("Int64")

        df["Date"] = pd.to_datetime(df["Date"])
        df = df.set_index("Date")

        return df

    @property
    def news(self) -> list:
        """Return news items from moex.news."""
        rows = _query(
            "SELECT data FROM moex.news WHERE symbol = %s", (self.ticker,)
        )
        return [row["data"] for row in rows] if rows else []

    @property
    def actions(self) -> pd.DataFrame:
        """Return dividends and stock splits from moex.stock_prices."""
        rows = _query(
            "SELECT date, dividends, stock_splits FROM moex.stock_prices WHERE symbol = %s AND (dividends != 0 OR stock_splits != 0) ORDER BY date ASC",
            (self.ticker,),
        )
        if not rows:
            return pd.DataFrame(columns=["Dividends", "Stock Splits"])

        df = pd.DataFrame(rows)
        df = df.rename(
            columns={
                "date": "Date",
                "dividends": "Dividends",
                "stock_splits": "Stock Splits",
            }
        )
        for col in ["Dividends", "Stock Splits"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        df["Date"] = pd.to_datetime(df["Date"])
        df = df.set_index("Date")
        return df

    def _get_financial_statement(self, stmt_type: str, freq: str) -> pd.DataFrame:
        """Retrieve a financial statement from moex.financial_statements.

        The DB stores JSONB with metric names as keys. We return a DataFrame
        with dates as columns and metrics as index (matching yfinance format).
        """
        rows = _query(
            "SELECT period_end, data FROM moex.financial_statements WHERE symbol = %s AND stmt_type = %s AND freq = %s ORDER BY period_end DESC",
            (self.ticker, stmt_type, freq),
        )
        if not rows:
            return pd.DataFrame()

        result = {}
        for row in rows:
            period_end = row["period_end"]
            col_key = pd.Timestamp(period_end)
            data = row["data"] if row["data"] else {}
            result[col_key] = data

        df = pd.DataFrame(result)
        # Sort columns (dates) descending
        df = df.reindex(sorted(df.columns, reverse=True), axis=1)
        return df

    @property
    def income_stmt(self) -> pd.DataFrame:
        return self._get_financial_statement("income_stmt", "annual")

    @property
    def quarterly_income_stmt(self) -> pd.DataFrame:
        return self._get_financial_statement("income_stmt", "quarterly")

    @property
    def balance_sheet(self) -> pd.DataFrame:
        return self._get_financial_statement("balance_sheet", "annual")

    @property
    def quarterly_balance_sheet(self) -> pd.DataFrame:
        return self._get_financial_statement("balance_sheet", "quarterly")

    @property
    def cashflow(self) -> pd.DataFrame:
        return self._get_financial_statement("cashflow", "annual")

    @property
    def quarterly_cashflow(self) -> pd.DataFrame:
        return self._get_financial_statement("cashflow", "quarterly")

    def _get_holder_data(self, holder_type: str) -> pd.DataFrame:
        """Retrieve holder data from moex.holders."""
        row = _query_single(
            "SELECT data FROM moex.holders WHERE symbol = %s AND holder_type = %s",
            (self.ticker, holder_type),
        )
        if not row or not row["data"]:
            return pd.DataFrame()

        data = row["data"]
        if isinstance(data, list):
            return pd.DataFrame(data)
        elif isinstance(data, dict):
            return pd.DataFrame(data)
        return pd.DataFrame()

    @property
    def major_holders(self) -> pd.DataFrame:
        return self._get_holder_data("major_holders")

    @property
    def institutional_holders(self) -> pd.DataFrame:
        return self._get_holder_data("institutional_holders")

    @property
    def mutualfund_holders(self) -> pd.DataFrame:
        return self._get_holder_data("mutualfund_holders")

    @property
    def insider_transactions(self) -> pd.DataFrame:
        return self._get_holder_data("insider_transactions")

    @property
    def insider_purchases(self) -> pd.DataFrame:
        return self._get_holder_data("insider_purchases")

    @property
    def insider_roster_holders(self) -> pd.DataFrame:
        return self._get_holder_data("insider_roster_holders")

    @property
    def options(self) -> tuple:
        """Return available option expiration dates as a tuple of strings."""
        rows = _query(
            "SELECT DISTINCT expiration_date FROM moex.options WHERE symbol = %s ORDER BY expiration_date",
            (self.ticker,),
        )
        return tuple(
            d.isoformat() if hasattr(d, "isoformat") else str(d)
            for d in (row["expiration_date"] for row in rows)
        ) if rows else ()

    def option_chain(self, expiration_date: str) -> OptionChain:
        """Return option chain for a given expiration date."""
        calls_row = _query_single(
            "SELECT data FROM moex.options WHERE symbol = %s AND expiration_date = %s AND option_type = %s",
            (self.ticker, expiration_date, "calls"),
        )
        puts_row = _query_single(
            "SELECT data FROM moex.options WHERE symbol = %s AND expiration_date = %s AND option_type = %s",
            (self.ticker, expiration_date, "puts"),
        )

        calls_data = calls_row["data"] if calls_row and calls_row["data"] else []
        puts_data = puts_row["data"] if puts_row and puts_row["data"] else []

        calls_df = pd.DataFrame(calls_data) if calls_data else pd.DataFrame()
        puts_df = pd.DataFrame(puts_data) if puts_data else pd.DataFrame()

        return OptionChain(calls_df, puts_df)

    @property
    def recommendations(self) -> pd.DataFrame:
        """Return recommendations from moex.recommendations."""
        row = _query_single(
            "SELECT data FROM moex.recommendations WHERE symbol = %s AND rec_type = %s",
            (self.ticker, "recommendations"),
        )
        if not row or not row["data"]:
            return pd.DataFrame()
        data = row["data"]
        if isinstance(data, list):
            return pd.DataFrame(data)
        return pd.DataFrame()

    @property
    def upgrades_downgrades(self) -> pd.DataFrame:
        """Return upgrades/downgrades from moex.recommendations."""
        row = _query_single(
            "SELECT data FROM moex.recommendations WHERE symbol = %s AND rec_type = %s",
            (self.ticker, "upgrades_downgrades"),
        )
        if not row or not row["data"]:
            return pd.DataFrame()
        data = row["data"]
        if isinstance(data, list):
            df = pd.DataFrame(data)
            if "GradeDate" in df.columns:
                df["GradeDate"] = pd.to_datetime(df["GradeDate"])
                df = df.set_index("GradeDate")
            return df
        return pd.DataFrame()


class PgYFinance:
    """Module-level object that mimics the yfinance module interface."""

    @staticmethod
    def Ticker(symbol: str) -> PgTicker:
        return PgTicker(symbol)
