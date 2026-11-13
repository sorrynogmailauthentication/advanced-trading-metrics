"""
Один символ за раз: qualify + reqHistoricalData, тайминг и errorEvent.

  python debug_ib_hist.py --date 2026-10-23
  python debug_ib_hist.py --date 2026-10-23 --symbols NVDA,NFLX,AAPL
  python debug_ib_hist.py --date 2026-10-23 --symbols NVDA --timeout 20
  python debug_ib_hist.py --date 2026-10-23 --verbose
"""

from __future__ import annotations

import argparse
import asyncio
import time
from datetime import date, datetime
from zoneinfo import ZoneInfo

from ib_insync import IB, Stock, util

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

HOST = "127.0.0.1"
PORT = 7496
CLIENT_ID = 95
ET = ZoneInfo("America/New_York")
DEFAULT_SYMBOLS = ("NVDA", "NFLX", "AAPL")


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Debug IBKR 1m historical bars for a few symbols")
    p.add_argument("--date", required=True, metavar="YYYY-MM-DD", help="Session date ET")
    p.add_argument(
        "--symbols",
        default=",".join(DEFAULT_SYMBOLS),
        help=f"Comma-separated (default {','.join(DEFAULT_SYMBOLS)})",
    )
    p.add_argument(
        "--timeout",
        type=float,
        default=60.0,
        help="Seconds to wait per reqHistoricalData (default 60)",
    )
    p.add_argument(
        "--duration",
        default="1 D",
        help='durationStr (default "1 D")',
    )
    p.add_argument("--verbose", action="store_true", help="ib_insync console log")
    return p.parse_args()


def _connect() -> IB:
    ib = IB()
    for port in (PORT, 7497, 7496):
        try:
            ib.connect(HOST, port, clientId=CLIENT_ID, readonly=True, timeout=5)
            print(f"connected {HOST}:{port} clientId={CLIENT_ID}")
            return ib
        except Exception as e:
            print(f"port {port}: {e}")
    raise SystemExit("Не удалось подключиться к TWS/Gateway")


def _end_datetime(session_day: date) -> str:
    today = datetime.now(ET).date()
    if session_day == today:
        return ""
    return f"{session_day.strftime('%Y%m%d')} 16:05:00 US/Eastern"


def _probe(
    ib: IB,
    symbol: str,
    *,
    end_dt: str,
    duration: str,
    timeout: float,
) -> None:
    print(f"\n=== {symbol} ===")
    t0 = time.perf_counter()
    contract = Stock(symbol, "SMART", "USD")
    try:
        qualified = ib.qualifyContracts(contract)
    except Exception as e:
        print(f"qualify FAIL ({time.perf_counter() - t0:.1f}s): {e}")
        return
    if not qualified:
        print(f"qualify FAIL ({time.perf_counter() - t0:.1f}s): empty result")
        return
    print(
        f"qualify OK ({time.perf_counter() - t0:.1f}s): "
        f"conId={contract.conId} exchange={contract.exchange} "
        f"primary={contract.primaryExchange} local={contract.localSymbol}"
    )

    t1 = time.perf_counter()
    try:
        bars = ib.reqHistoricalData(
            contract,
            endDateTime=end_dt,
            durationStr=duration,
            barSizeSetting="1 min",
            whatToShow="TRADES",
            useRTH=True,
            formatDate=1,
            keepUpToDate=False,
            timeout=timeout,
        )
    except Exception as e:
        print(f"history FAIL ({time.perf_counter() - t1:.1f}s): {type(e).__name__}: {e}")
        return

    elapsed = time.perf_counter() - t1
    n = len(bars)
    if n == 0:
        print(
            f"history EMPTY ({elapsed:.1f}s) — timeout/cancel or no bars. "
            f"Смотри ERR-строки выше (162 = часто cancel после timeout)."
        )
        return

    first, last = bars[0], bars[-1]
    print(f"history OK ({elapsed:.1f}s): {n} bars")
    print(f"  first: {first.date} O={first.open} C={first.close} V={first.volume}")
    print(f"  last:  {last.date} O={last.open} C={last.close} V={last.volume}")


def main() -> None:
    args = _parse_args()
    session_day = date.fromisoformat(args.date)
    symbols = [s.strip().upper() for s in args.symbols.split(",") if s.strip()]
    if not symbols:
        raise SystemExit("пустой --symbols")

    if args.verbose:
        util.logToConsole()

    end_dt = _end_datetime(session_day)
    print(
        f"date={session_day.isoformat()} endDateTime={repr(end_dt) or 'now'} "
        f"duration={args.duration!r} timeout={args.timeout}s symbols={symbols}"
    )

    ib = _connect()
    ib.RaiseRequestErrors = True

    def on_error(reqId, errorCode, errorString, contract):
        sym = getattr(contract, "symbol", None) if contract else None
        print(
            f"  ERR t={time.strftime('%H:%M:%S')} reqId={reqId} "
            f"code={errorCode} {errorString}"
            + (f" | {sym}" if sym else "")
        )

    ib.errorEvent += on_error
    try:
        for sym in symbols:
            _probe(
                ib,
                sym,
                end_dt=end_dt,
                duration=args.duration,
                timeout=args.timeout,
            )
    finally:
        ib.disconnect()
        print("\ndisconnected")


if __name__ == "__main__":
    main()
