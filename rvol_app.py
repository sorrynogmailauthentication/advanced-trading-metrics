"""
Локальный график RVOL: IB сверху, TV снизу, общая временная шкала.

Запуск:  uvicorn rvol_app:app --reload --port 8787
Открыть: http://127.0.0.1:8787/t/NVDA

В Excel:
  =ГИПЕРССЫЛКА("http://127.0.0.1:8787/t/"&L2;"RVOL")
"""
from __future__ import annotations

from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from dotenv import load_dotenv

load_dotenv()

from rvol_db import (  # noqa: E402
    ensure_schema,
    fetch_aligned_series,
    list_tickers_today_any,
)

app = FastAPI(title="RVOL history")
REFRESH_MS = 60_000  # RVOL пишется раз в минуту
WINDOW_MINUTES = 60

ensure_schema()


@app.get("/", response_class=HTMLResponse)
def index():
    tickers = list_tickers_today_any()
    links = "".join(
        f'<li><a href="/t/{t}">{t}</a></li>' for t in tickers
    ) or "<li>(no data today)</li>"
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>RVOL</title>
<style>
  body{{font:12px system-ui;margin:8px;background:#111;color:#ddd}}
  a{{color:#8ab4f8}} ul{{padding-left:16px;margin:4px 0}}
</style></head>
<body><ul>{links}</ul></body></html>"""


@app.get("/t/{ticker}", response_class=HTMLResponse)
def ticker_chart(ticker: str):
    ticker = ticker.strip().upper()
    return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>{ticker} RVOL</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    html, body {{ width: 100%; height: 100%; overflow: hidden; background: #111; color: #ccc; }}
    body {{ font: 12px/1.2 system-ui, sans-serif; }}
    #wrap {{
      width: 440px; height: 380px;
      display: flex; flex-direction: column;
      padding: 2px 6px 4px 2px;
    }}
    .pane {{
      flex: 1 1 0; min-height: 0;
      display: flex; flex-direction: column;
    }}
    .head {{
      display: flex; justify-content: space-between; align-items: baseline;
      flex: 0 0 16px; padding: 0 2px;
    }}
    .sym {{ font-weight: 700; color: #eee; font-size: 12px; }}
    .meta {{ color: #777; font-size: 11px; }}
    .box {{ flex: 1 1 auto; min-height: 0; position: relative; }}
    canvas {{ width: 100% !important; height: 100% !important; cursor: crosshair; }}
    #hbar {{
      flex: 0 0 14px;
      margin: 2px 2px 0 28px;
      overflow-x: auto;
      overflow-y: hidden;
      background: #1a1a1a;
    }}
    #hbar::-webkit-scrollbar {{ height: 10px; }}
    #hbar::-webkit-scrollbar-track {{ background: #1a1a1a; }}
    #hbar::-webkit-scrollbar-thumb {{ background: #555; border-radius: 5px; }}
    #hbar::-webkit-scrollbar-thumb:hover {{ background: #777; }}
    #hbar-inner {{ height: 1px; width: 100%; }}
  </style>
</head>
<body>
  <div id="wrap">
    <div class="pane">
      <div class="head"><span class="sym">{ticker} · IB</span><span class="meta" id="meta-ib">…</span></div>
      <div class="box"><canvas id="c-ib"></canvas></div>
    </div>
    <div class="pane">
      <div class="head"><span class="sym">{ticker} · TV</span><span class="meta" id="meta-tv">…</span></div>
      <div class="box"><canvas id="c-tv"></canvas></div>
    </div>
    <div id="hbar"><div id="hbar-inner"></div></div>
  </div>
  <script>
    const TICKER = {ticker!r};
    const REFRESH_MS = {REFRESH_MS};
    const WINDOW = {WINDOW_MINUTES};
    const Y_DEFAULT_MAX = 3;   // шкала 0..3 по умолчанию
    const Y_HEADROOM = 1.05;   // небольшой запас, если значение > 3
    const MIN_RANGE = 8;
    let followLive = true;
    let initialView = true;  // только первый кадр → last hour
    let pointCount = 0;
    let hoverIdx = null;
    let syncingBar = false;
    let refreshTimer = null;
    const hbar = document.getElementById('hbar');
    const hbarInner = document.getElementById('hbar-inner');

    function clamp(v, a, b) {{ return Math.max(a, Math.min(b, v)); }}

    function makeChart(canvasId, color) {{
      return new Chart(document.getElementById(canvasId), {{
        type: 'bar',
        data: {{
          labels: [],
          datasets: [{{
            data: [],
            backgroundColor: color,
            borderWidth: 0,
            barPercentage: 0.85,
            categoryPercentage: 0.9,
            spanGaps: false,
            skipNull: true
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          transitions: {{ active: {{ animation: {{ duration: 0 }} }} }},
          layout: {{ padding: {{ top: 4, right: 6, left: 2, bottom: 0 }} }},
          onHover: (evt, elements) => {{
            if (elements && elements.length) hoverIdx = elements[0].index;
            else if (evt && evt.x != null && charts[0].scales.x) {{
              const v = charts[0].scales.x.getValueForPixel(evt.x);
              hoverIdx = Number.isFinite(v) ? Math.round(v) : null;
            }}
          }},
          plugins: {{
            legend: {{ display: false }},
            tooltip: {{
              enabled: true,
              animation: false,
              filter: (item) => item.raw != null
            }}
          }},
          scales: {{
            x: {{
              min: 0,
              max: WINDOW - 1,
              ticks: {{
                color: '#888', font: {{ size: 10 }},
                maxRotation: 0, autoSkip: true, maxTicksLimit: 8
              }},
              grid: {{ display: false }}
            }},
            y: {{
              min: 0,
              max: Y_DEFAULT_MAX,
              beginAtZero: true,
              grace: 0,
              ticks: {{
                color: '#bbb',
                font: {{ size: 11, weight: '600' }},
                maxTicksLimit: 5,
                padding: 4,
                callback: (v) => Number(v).toFixed(v < 10 ? 1 : 0)
              }},
              grid: {{ color: '#2a2a2a' }}
            }}
          }}
        }}
      }});
    }}

    const chartIb = makeChart('c-ib', '#f59e0b');
    const chartTv = makeChart('c-tv', '#3b82f6');
    const charts = [chartIb, chartTv];

    function viewRange() {{
      const x = chartIb.options.scales.x;
      let min = Number(x.min);
      let max = Number(x.max);
      if (!Number.isFinite(min)) min = 0;
      if (!Number.isFinite(max)) max = Math.max(0, pointCount - 1);
      return {{ min, max, size: Math.max(1, max - min) }};
    }}

    function setView(min, max, updateBar = true) {{
      if (pointCount <= 0) return;
      const last = pointCount - 1;
      let size = Math.max(MIN_RANGE - 1, max - min);
      size = Math.min(size, Math.max(0, last));
      min = clamp(min, 0, Math.max(0, last - size));
      max = min + size;
      for (const c of charts) {{
        c.options.scales.x.min = min;
        c.options.scales.x.max = max;
      }}
      followLive = max >= last - 0.51;
      applyYScales();
      if (updateBar) syncScrollbar();
      for (const c of charts) c.update('none');
    }}

    function showLastHour() {{
      if (pointCount <= 0) return;
      if (pointCount <= WINDOW) setView(0, pointCount - 1);
      else setView(pointCount - WINDOW, pointCount - 1);
    }}

    function syncScrollbar() {{
      if (syncingBar) return;
      const {{ min, size }} = viewRange();
      const last = Math.max(0, pointCount - 1);
      const viewW = hbar.clientWidth || 1;
      const total = Math.max(size + 1, last + 1);
      const contentW = Math.max(viewW, Math.round(viewW * total / (size + 1)));
      syncingBar = true;
      hbarInner.style.width = contentW + 'px';
      const maxScroll = Math.max(0, contentW - viewW);
      const maxStart = Math.max(0, last - size);
      hbar.scrollLeft = maxStart <= 0 ? 0 : (min / maxStart) * maxScroll;
      syncingBar = false;
    }}

    function lastFinite(arr) {{
      for (let i = arr.length - 1; i >= 0; i--) {{
        if (arr[i] != null && Number.isFinite(arr[i])) return arr[i];
      }}
      return null;
    }}

    function visibleMax(arr) {{
      const {{ min, max }} = viewRange();
      const i0 = Math.max(0, Math.floor(min));
      const i1 = Math.min(arr.length - 1, Math.ceil(max));
      let m = 0;
      for (let i = i0; i <= i1; i++) {{
        const v = arr[i];
        if (v != null && Number.isFinite(v) && v > m) m = v;
      }}
      return m;
    }}

    function yAxisMax(arr) {{
      const m = visibleMax(arr);
      if (m > Y_DEFAULT_MAX) return m * Y_HEADROOM;
      return Y_DEFAULT_MAX;
    }}

    function applyYScales() {{
      const ib = chartIb.data.datasets[0].data || [];
      const tv = chartTv.data.datasets[0].data || [];
      chartIb.options.scales.y.min = 0;
      chartTv.options.scales.y.min = 0;
      chartIb.options.scales.y.max = yAxisMax(ib);
      chartTv.options.scales.y.max = yAxisMax(tv);
      // сбросить suggested*, чтобы Chart.js не раздувал шкалу
      delete chartIb.options.scales.y.suggestedMax;
      delete chartTv.options.scales.y.suggestedMax;
    }}

    function onWheel(e) {{
      e.preventDefault();
      if (pointCount <= 0) return;
      const {{ min, max, size }} = viewRange();
      const last = pointCount - 1;
      const focus = clamp(
        hoverIdx != null ? hoverIdx : Math.round((min + max) / 2),
        0, last
      );
      const factor = e.deltaY < 0 ? 0.85 : 1.18;
      let newSize = clamp(size * factor, MIN_RANGE - 1, Math.max(MIN_RANGE - 1, last));
      const rel = size <= 0 ? 0.5 : (focus - min) / size;
      const newMin = focus - rel * newSize;
      setView(newMin, newMin + newSize);
    }}

    document.getElementById('c-ib').addEventListener('wheel', onWheel, {{ passive: false }});
    document.getElementById('c-tv').addEventListener('wheel', onWheel, {{ passive: false }});

    hbar.addEventListener('scroll', () => {{
      if (syncingBar) return;
      const {{ size }} = viewRange();
      const last = Math.max(0, pointCount - 1);
      const maxStart = Math.max(0, last - size);
      const maxScroll = Math.max(0, hbar.scrollWidth - hbar.clientWidth);
      const start = maxScroll <= 0 ? 0 : (hbar.scrollLeft / maxScroll) * maxStart;
      setView(start, start + size, false);
    }});

    async function refresh() {{
      try {{
        const r = await fetch('/api/t/' + encodeURIComponent(TICKER));
        if (!r.ok) {{
          document.getElementById('meta-ib').textContent = 'no data';
          document.getElementById('meta-tv').textContent = 'no data';
          return;
        }}
        const j = await r.json();
        const labels = j.labels || [];
        const ib = (j.ib || []).map(v => (v == null ? null : Number(v)));
        const tv = (j.tv || []).map(v => (v == null ? null : Number(v)));
        const prev = viewRange();
        pointCount = labels.length;

        chartIb.data.labels = labels;
        chartTv.data.labels = labels;
        chartIb.data.datasets[0].data = ib;
        chartTv.data.datasets[0].data = tv;

        const last = Math.max(0, pointCount - 1);
        const prevSize = prev.size;
        if (initialView || !labels.length) {{
          // первый показ — последние WINDOW минут; дальше зум/скролл не сбрасываем
          showLastHour();
          initialView = false;
        }} else if (followLive) {{
          // у живого края сохраняем ТЕКУЩИЙ размер окна (можно > WINDOW)
          const size = Math.min(Math.max(prevSize, MIN_RANGE - 1), last);
          setView(last - size, last);
        }} else {{
          setView(prev.min, prev.min + prevSize);
        }}
        applyYScales();
        for (const c of charts) c.update('none');
        syncScrollbar();

        const lastIb = lastFinite(ib);
        const lastTv = lastFinite(tv);
        document.getElementById('meta-ib').textContent =
          lastIb == null ? '—' : lastIb.toFixed(2);
        document.getElementById('meta-tv').textContent =
          lastTv == null ? '—' : lastTv.toFixed(2);
      }} catch (e) {{
        document.getElementById('meta-ib').textContent = 'err';
        document.getElementById('meta-tv').textContent = 'err';
      }}
    }}

    function scheduleRefresh() {{
      if (refreshTimer) clearInterval(refreshTimer);
      refreshTimer = setInterval(() => {{
        // фоновые/свёрнутые вкладки не долбят API
        if (document.hidden) return;
        refresh();
      }}, REFRESH_MS);
    }}

    document.addEventListener('visibilitychange', () => {{
      if (!document.hidden) refresh();
    }});

    refresh();
    scheduleRefresh();
  </script>
</body>
</html>"""


@app.get("/api/t/{ticker}")
def api_ticker(ticker: str):
    ticker = ticker.strip().upper()
    return fetch_aligned_series(ticker)
