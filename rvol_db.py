"""PostgreSQL helpers for RVOL minute history (append-only)."""
from __future__ import annotations

import os
from datetime import date, datetime
from typing import Any, List, Optional, Sequence, Tuple, Union

import psycopg

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:postgres@localhost:5432/playbook",
)


def connect():
    return psycopg.connect(DATABASE_URL)


def ensure_schema() -> None:
    with connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS rvol_ticks (
                id      BIGSERIAL PRIMARY KEY,
                ts      TIMESTAMPTZ NOT NULL,
                ticker  TEXT NOT NULL,
                rvol    DOUBLE PRECISION NOT NULL,
                source  TEXT NOT NULL DEFAULT 'tv'
            )
            """
        )
        # старые БД без source
        conn.execute(
            """
            ALTER TABLE rvol_ticks
            ADD COLUMN IF NOT EXISTS source TEXT NOT NULL DEFAULT 'tv'
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS rvol_ticks_ticker_ts_idx
                ON rvol_ticks (ticker, ts)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS rvol_ticks_source_ticker_ts_idx
                ON rvol_ticks (source, ticker, ts)
            """
        )
        conn.commit()


def insert_ticks(
    rows: Sequence[Tuple[datetime, str, float]],
    source: str = "tv",
) -> int:
    """Insert (ts, ticker, rvol) with source tag. Returns rows inserted."""
    if not rows:
        return 0
    payload = [(ts, t, v, source) for ts, t, v in rows]
    with connect() as conn:
        with conn.cursor() as cur:
            cur.executemany(
                """
                INSERT INTO rvol_ticks (ts, ticker, rvol, source)
                VALUES (%s, %s, %s, %s)
                """,
                payload,
            )
        conn.commit()
    return len(rows)


def _as_day(day: Optional[Union[date, datetime]]) -> Optional[date]:
    if day is None:
        return None
    if isinstance(day, datetime):
        return day.date()
    return day


def fetch_ticker_series(
    ticker: str,
    day: Optional[Union[date, datetime]] = None,
    source: str = "tv",
) -> List[dict[str, Any]]:
    """RVOL points for ticker/source, default = today."""
    d = _as_day(day)
    with connect() as conn:
        with conn.cursor() as cur:
            if d is None:
                cur.execute(
                    """
                    SELECT ts, rvol
                    FROM rvol_ticks
                    WHERE ticker = %s
                      AND source = %s
                      AND ts >= date_trunc('day', now())
                    ORDER BY ts
                    """,
                    (ticker, source),
                )
            else:
                cur.execute(
                    """
                    SELECT ts, rvol
                    FROM rvol_ticks
                    WHERE ticker = %s
                      AND source = %s
                      AND ts >= %s::date
                      AND ts < %s::date + interval '1 day'
                    ORDER BY ts
                    """,
                    (ticker, source, d, d),
                )
            return [{"ts": r[0].isoformat(), "rvol": r[1]} for r in cur.fetchall()]


def list_tickers_today(source: str = "tv") -> List[str]:
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT ticker
                FROM rvol_ticks
                WHERE source = %s
                  AND ts >= date_trunc('day', now())
                ORDER BY ticker
                """,
                (source,),
            )
            return [r[0] for r in cur.fetchall()]


def list_tickers_today_any() -> List[str]:
    """Tickers with any data today; if none, fall back to all known tickers."""
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT ticker
                FROM rvol_ticks
                WHERE ts >= date_trunc('day', now())
                ORDER BY ticker
                """
            )
            today = [r[0] for r in cur.fetchall()]
            if today:
                return today
            cur.execute(
                """
                SELECT DISTINCT ticker
                FROM rvol_ticks
                ORDER BY ticker
                """
            )
            return [r[0] for r in cur.fetchall()]


def latest_day_for_ticker(ticker: str) -> Optional[date]:
    """Newest calendar day with any ticks for ticker, or None."""
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT (MAX(ts))::date
                FROM rvol_ticks
                WHERE ticker = %s
                """,
                (ticker,),
            )
            row = cur.fetchone()
            return row[0] if row and row[0] is not None else None


def adjacent_day_with_data(
    ticker: str,
    day: date,
    direction: int,
) -> Optional[date]:
    """Next/prev calendar day that has ticks. direction: -1 prev, +1 next."""
    with connect() as conn:
        with conn.cursor() as cur:
            if direction < 0:
                cur.execute(
                    """
                    SELECT (MAX(ts))::date
                    FROM rvol_ticks
                    WHERE ticker = %s AND ts < %s::date
                    """,
                    (ticker, day),
                )
            else:
                cur.execute(
                    """
                    SELECT (MIN(ts))::date
                    FROM rvol_ticks
                    WHERE ticker = %s AND ts >= %s::date + interval '1 day'
                    """,
                    (ticker, day),
                )
            row = cur.fetchone()
            return row[0] if row and row[0] is not None else None


def _minute_key(ts_iso: str) -> str:
    """Нормализация к минуте HH:MM для склейки источников."""
    # 2026-10-12T15:33:00+03:00 or ...15:33:00
    if "T" in ts_iso:
        t = ts_iso.split("T", 1)[1]
    else:
        t = ts_iso
    return t.replace("Z", "")[:5]


def fetch_aligned_series(
    ticker: str,
    day: Optional[Union[date, datetime]] = None,
) -> dict[str, Any]:
    """
    Общая шкала минут за день: labels + ib[] + tv[].
    day=None → сегодня. Нет данных у источника → null (пустой бар).
    """
    d = _as_day(day)
    ib = fetch_ticker_series(ticker, day=d, source="ib")
    tv = fetch_ticker_series(ticker, day=d, source="tv")
    ib_map = {_minute_key(p["ts"]): p["rvol"] for p in ib}
    tv_map = {_minute_key(p["ts"]): p["rvol"] for p in tv}
    labels = sorted(set(ib_map) | set(tv_map))
    day_out = d.isoformat() if d is not None else date.today().isoformat()
    return {
        "ticker": ticker,
        "day": day_out,
        "labels": labels,
        "ib": [ib_map.get(m) for m in labels],
        "tv": [tv_map.get(m) for m in labels],
    }
