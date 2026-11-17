"""
Простое окно для export_corr_bars_hist:
  дата, start/end (dropdown), −10/−20 мин от end → start, Run.

  python export_corr_bars_ui.py
"""

from __future__ import annotations

import calendar
import threading
import tkinter as tk
from datetime import date, datetime, timedelta, time as dtime
from tkinter import ttk
from zoneinfo import ZoneInfo

from export_corr_bars_hist import DEFAULT_END, DEFAULT_START, run_oneshot

ET = ZoneInfo("America/New_York")
# Одно меню на время: 00:00 … 23:59
TIMES = [f"{h:02d}:{m:02d}" for h in range(24) for m in range(60)]
# Пн…Вс — firstweekday=0 в calendar
WEEKDAYS = ("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")
CELL_W = 4


class DatePicker(ttk.Frame):
    """Поле YYYY-MM-DD + календарь с выровненной сеткой дней недели."""

    def __init__(self, master, initial: date, **kwargs):
        super().__init__(master, **kwargs)
        self._var = tk.StringVar(value=initial.isoformat())
        ttk.Entry(self, textvariable=self._var, width=12).pack(side=tk.LEFT)
        ttk.Button(self, text="…", width=3, command=self._open).pack(
            side=tk.LEFT, padx=(4, 0)
        )
        self._popup: tk.Toplevel | None = None

    def get(self) -> date:
        return date.fromisoformat(self._var.get().strip())

    def set(self, day: date) -> None:
        self._var.set(day.isoformat())

    def _open(self) -> None:
        if self._popup and self._popup.winfo_exists():
            self._popup.lift()
            return
        try:
            current = self.get()
        except ValueError:
            current = datetime.now(ET).date()

        pop = tk.Toplevel(self)
        self._popup = pop
        pop.title("Date")
        pop.resizable(False, False)
        pop.transient(self.winfo_toplevel())
        pop.grab_set()

        year = tk.IntVar(value=current.year)
        month = tk.IntVar(value=current.month)

        outer = ttk.Frame(pop, padding=8)
        outer.pack()

        head = ttk.Frame(outer)
        head.pack(fill=tk.X, pady=(0, 6))

        def month_title() -> str:
            return f"{calendar.month_name[month.get()]} {year.get()}"

        title_lbl = ttk.Label(head, text=month_title(), width=18, anchor="center")

        def shift(delta: int) -> None:
            m = month.get() + delta
            y = year.get()
            while m < 1:
                m += 12
                y -= 1
            while m > 12:
                m -= 12
                y += 1
            month.set(m)
            year.set(y)
            title_lbl.config(text=month_title())
            redraw()

        ttk.Button(head, text="<", width=3, command=lambda: shift(-1)).pack(side=tk.LEFT)
        title_lbl.pack(side=tk.LEFT, expand=True, padx=4)
        ttk.Button(head, text=">", width=3, command=lambda: shift(1)).pack(side=tk.RIGHT)

        # Одна сетка: заголовки дней и числа в одних и тех же колонках 0..6
        body = ttk.Frame(outer)
        body.pack()
        for col in range(7):
            body.columnconfigure(col, weight=1, minsize=36)

        def redraw() -> None:
            for w in body.winfo_children():
                w.destroy()
            for col, name in enumerate(WEEKDAYS):
                ttk.Label(
                    body,
                    text=name,
                    width=CELL_W,
                    anchor="center",
                    font=("Segoe UI", 9, "bold"),
                ).grid(row=0, column=col, padx=1, pady=(0, 4), sticky="nsew")

            cal = calendar.Calendar(firstweekday=calendar.MONDAY)
            for row_i, week in enumerate(
                cal.monthdayscalendar(year.get(), month.get()), start=1
            ):
                for col, day in enumerate(week):
                    if day == 0:
                        ttk.Label(body, text="", width=CELL_W).grid(
                            row=row_i, column=col, padx=1, pady=1
                        )
                        continue

                    def pick(d: int = day, y: int = year.get(), m: int = month.get()) -> None:
                        self.set(date(y, m, d))
                        pop.grab_release()
                        pop.destroy()
                        self._popup = None

                    style_btn = ttk.Button(
                        body, text=f"{day:2d}", width=CELL_W, command=pick
                    )
                    style_btn.grid(row=row_i, column=col, padx=1, pady=1, sticky="nsew")
                    if (
                        year.get() == current.year
                        and month.get() == current.month
                        and day == current.day
                    ):
                        style_btn.state(["pressed"])

        redraw()

        def on_close() -> None:
            pop.grab_release()
            pop.destroy()
            self._popup = None

        pop.protocol("WM_DELETE_WINDOW", on_close)


class TimePicker(ttk.Frame):
    """Один Combobox HH:MM."""

    def __init__(self, master, initial: dtime, **kwargs):
        super().__init__(master, **kwargs)
        self._var = tk.StringVar(value=initial.strftime("%H:%M"))
        ttk.Combobox(
            self,
            textvariable=self._var,
            values=TIMES,
            width=6,
            state="readonly",
            justify="center",
        ).pack(side=tk.LEFT)

    def get(self) -> dtime:
        return datetime.strptime(self._var.get().strip(), "%H:%M").time()

    def set(self, t: dtime) -> None:
        self._var.set(t.strftime("%H:%M"))


class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=12)
        self.pack(fill=tk.BOTH, expand=True)
        master.title("ATR corr bars — IBKR export")
        master.resizable(False, False)

        today = datetime.now(ET).date()

        ttk.Label(self, text="Date (ET)").grid(row=0, column=0, sticky="w", pady=4)
        self.date_picker = DatePicker(self, today)
        self.date_picker.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(self, text="Start").grid(row=1, column=0, sticky="w", pady=4)
        self.start_picker = TimePicker(self, DEFAULT_START)
        self.start_picker.grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(self, text="End").grid(row=2, column=0, sticky="w", pady=4)
        self.end_picker = TimePicker(self, DEFAULT_END)
        self.end_picker.grid(row=2, column=1, sticky="w", pady=4)

        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 4))
        ttk.Button(btns, text="−10 min", command=lambda: self._back_from_end(10)).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btns, text="−20 min", command=lambda: self._back_from_end(20)).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        self.run_btn = ttk.Button(btns, text="Run", command=self._run)
        self.run_btn.pack(side=tk.RIGHT)

        self.status = tk.StringVar(value="Idle")
        ttk.Label(self, textvariable=self.status, wraplength=360).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )

        self._running = False

    def _back_from_end(self, minutes: int) -> None:
        try:
            end = self.end_picker.get()
        except ValueError as e:
            self.status.set(f"Idle — {e}")
            return
        base = datetime.combine(date(2000, 1, 1), end)
        start = (base - timedelta(minutes=minutes)).time().replace(second=0, microsecond=0)
        self.start_picker.set(start)
        if not self._running:
            self.status.set("Idle")

    def _run(self) -> None:
        if self._running:
            return
        try:
            session_day = self.date_picker.get()
            t0 = self.start_picker.get()
            t1 = self.end_picker.get()
        except ValueError as e:
            self.status.set(f"Idle — {e}")
            return
        if t1 < t0:
            self.status.set("Idle — end must be >= start")
            return

        self._running = True
        self.run_btn.state(["disabled"])
        self.status.set("In progress")

        def worker() -> None:
            import asyncio

            try:
                asyncio.get_event_loop()
            except RuntimeError:
                asyncio.set_event_loop(asyncio.new_event_loop())

            err: Exception | None = None
            try:
                run_oneshot(session_day, t0, t1)
            except Exception as e:
                err = e

            def done() -> None:
                self._running = False
                self.run_btn.state(["!disabled"])
                self.status.set("Done" if err is None else f"Done — failed: {err}")

            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    root = tk.Tk()
    try:
        root.call("tk", "scaling", 1.25)
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
