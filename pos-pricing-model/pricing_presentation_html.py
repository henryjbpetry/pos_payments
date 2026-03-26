"""
Generate pricing-presentation.html — executive-first deck (charts + results before methodology).
Style aligned with gtm-sales-deck; narrative benchmark: PoS GTM investment-style slides.
"""
from __future__ import annotations

import html
import re
from datetime import datetime, timezone
from typing import Any

import pandas as pd


def _esc(x: Any) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return html.escape(str(x))


def _fmt_bc_value(label: str, val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "—"
    try:
        v = float(val)
    except (TypeError, ValueError):
        return _esc(val)
    lab = (label or "").strip().lower()
    if not lab:
        if 0 < abs(v) < 2:
            return f"{v * 100:.2f}%"
        return f"{v:,.4f}"
    if "capex" in lab or ("gp" in lab and "notr" not in lab):
        return f"${v:,.2f}"
    if "notr" in lab or "ntr" in lab or "fee" in lab or "cost" in lab:
        return f"{v * 100:.4f}".rstrip("0").rstrip(".") + "%"
    if "rental" in lab or "logistics" in lab or "taxes" in lab:
        return f"{v * 100:.4f}".rstrip("0").rstrip(".") + "%"
    return f"{v * 100:.4f}".rstrip("0").rstrip(".") + "%"


def _read_bc_matrix(bc_path: str) -> tuple[list[str], list[str], list[tuple[str | None, list[Any]]]]:
    df = pd.read_excel(bc_path, sheet_name="2026 BC", header=None)
    regions = [_esc(df.iloc[13, c]) if not pd.isna(df.iloc[13, c]) else "" for c in range(2, 6)]
    providers = [_esc(df.iloc[14, c]) if not pd.isna(df.iloc[14, c]) else "" for c in range(2, 6)]

    rows_out: list[tuple[str | None, list[Any]]] = []
    for r in range(15, 38):
        label = df.iloc[r, 1]
        vals = [df.iloc[r, c] for c in range(2, 6)]
        if all(pd.isna(v) for v in vals) and (pd.isna(label) or str(label).strip() == ""):
            rows_out.append((None, []))
            continue
        lab = None if pd.isna(label) else str(label).strip()
        rows_out.append((lab if lab else None, vals))
    return regions, providers, rows_out


def _bc_row_by_label(
    bc_rows: list[tuple[str | None, list[Any]]], contains: str
) -> tuple[str | None, list[Any]] | None:
    c = contains.lower()
    for lab, vals in bc_rows:
        if lab and c in lab.lower():
            return (lab, vals)
    return None


def _build_bc_table_html(regions: list[str], providers: list[str], bc_rows: list) -> str:
    parts: list[str] = []
    parts.append("<thead>")
    parts.append("<tr>")
    parts.append('<th rowspan="2" class="corner" scope="col">Only CC</th>')
    for i, reg in enumerate(regions):
        parts.append(f'<th scope="col" class="excel-h2">{reg or "—"}</th>')
    parts.append("</tr><tr>")
    for i, name in enumerate(providers):
        cls = " stone" if i == 1 else ""
        parts.append(f'<th scope="col" class="excel-h{cls}">{name or "—"}</th>')
    parts.append("</tr></thead><tbody>")

    for lab, vals in bc_rows:
        if not vals and lab is None:
            parts.append('<tr class="spacer"><td colspan="5"></td></tr>')
            continue
        parts.append("<tr>")
        if lab:
            parts.append(f'<td class="excel-metric">{_esc(lab)}</td>')
        else:
            parts.append('<td class="excel-metric"></td>')
        if not vals:
            parts.append("</tr>")
            continue
        for j, v in enumerate(vals):
            cell = _fmt_bc_value(lab or "", v)
            cls = " stone-col" if j == 1 else ""
            parts.append(f'<td class="excel-num{cls}">{cell}</td>')
        parts.append("</tr>")
    parts.append("</tbody>")
    return "".join(parts)


def _pct01(x: Any) -> float | None:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    try:
        return float(x) * 100.0
    except (TypeError, ValueError):
        return None


def _hbar(label: str, value_pct: float, max_scale: float, color: str, value_fmt: str | None = None) -> str:
    """Horizontal bar; value_pct and max_scale in same units (e.g. percentage points)."""
    if max_scale <= 0:
        w = 0.0
    else:
        w = min(100.0, max(0.0, (value_pct / max_scale) * 100.0))
    vf = value_fmt if value_fmt is not None else f"{value_pct:.2f}%"
    return f"""<div class="hbar">
  <span class="hbar-label">{_esc(label)}</span>
  <div class="hbar-track" role="img" aria-label="{_esc(label)} {vf}"><span class="hbar-fill" style="width:{w:.1f}%;background:{color}"></span></div>
  <span class="hbar-val">{_esc(vf)}</span>
</div>"""


def _pick_summary_row(summary_non_promo: pd.DataFrame, metric: str) -> dict[str, Any]:
    row = summary_non_promo.loc[summary_non_promo["metric"] == metric]
    if row.empty:
        return {"n": 0, "min": None, "mean": None, "max": None}
    r = row.iloc[0]
    return {"n": int(r["n"]), "min": r["min"], "mean": r["mean"], "max": r["max"]}


def _short_scenario_name(s: str) -> str:
    s = re.sub(r"^suggested_", "", s)
    s = s.replace("_", " ")
    s = re.sub(r" below market (min|avg|max) ", r" \1 ", s)
    if len(s) > 42:
        return s[:39] + "…"
    return s


def write_presentation_html(
    out_path: str,
    bc_path: str,
    summary_non_promo: pd.DataFrame,
    by_prazo_df: pd.DataFrame,
    scen_df: pd.DataFrame,
    suggested_df: pd.DataFrame,
    stone: dict[str, Any],
    n_promo: int,
    n_non_promo: int,
    eps: float,
    build_stamp: str | None = None,
) -> None:
    stamp = build_stamp or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    regions, providers, bc_rows = _read_bc_matrix(bc_path)
    bc_table_inner = _build_bc_table_html(regions, providers, bc_rows)

    notr_row = _bc_row_by_label(bc_rows, "notr %")
    lift_notr_row = _bc_row_by_label(bc_rows, "lift notr")
    gp_row = _bc_row_by_label(bc_rows, "gp usd/month")
    proc_fee_row = _bc_row_by_label(bc_rows, "processing fee")
    ant_fee_row = _bc_row_by_label(bc_rows, "anticipation fee")

    # --- Executive KPIs (Stone = col index 1) ---
    stone_i = 1
    notr_vals = [ _pct01(notr_row[1][i]) if notr_row and len(notr_row[1]) > i else None for i in range(4) ]
    lift_vals = [ _pct01(lift_notr_row[1][i]) if lift_notr_row and len(lift_notr_row[1]) > i else None for i in range(4) ]
    gp_vals = [ float(gp_row[1][i]) if gp_row and len(gp_row[1]) > i and pd.notna(gp_row[1][i]) else None for i in range(4) ]

    stone_notr = notr_vals[stone_i] if notr_vals[stone_i] is not None else _pct01(stone.get("notr_pct"))
    stone_lift = lift_vals[stone_i] if lift_vals[stone_i] is not None else _pct01(stone.get("lift_notr_pct"))
    stone_gp = gp_vals[stone_i] if gp_vals[stone_i] is not None else float(stone.get("gp_usd_month") or 0)

    cred = _pick_summary_row(summary_non_promo, "mdr_credito_vista_rate")
    deb = _pick_summary_row(summary_non_promo, "mdr_debito_rate")
    pix = _pick_summary_row(summary_non_promo, "mdr_pix_rate")

    # Regional NOTR chart scale
    notr_for_scale = [v for v in notr_vals if v is not None]
    max_notr = max(notr_for_scale) if notr_for_scale else 2.5
    max_notr = max(max_notr * 1.15, 0.5)

    colors = ["#3b82f6", "#22c55e", "#a855f7", "#f97316"]  # Getnet-ish, Stone, Stripe, dLocal
    regional_bars = []
    names = ["Getnet", "Stone", "Stripe", "dLocal QR"]
    for i in range(4):
        v = notr_vals[i]
        if v is None:
            continue
        regional_bars.append(_hbar(names[i], v, max_notr, colors[i]))

    # Market band: min/mean/max for debit & credit (clustered visual — 2 groups of 3 mini bars)
    def trio_bars(title: str, lo: Any, mid: Any, hi: Any, cmin: str, cavg: str, cmax: str) -> str:
        lo_p, mid_p, hi_p = _pct01(lo), _pct01(mid), _pct01(hi)
        mx = max(x for x in (lo_p, mid_p, hi_p) if x is not None) or 1
        mx *= 1.1
        parts = [f'<div class="trio"><p class="trio-title">{_esc(title)}</p><div class="trio-bars">']
        for lab, v, col in [("Min", lo_p, cmin), ("Avg", mid_p, cavg), ("Max", hi_p, cmax)]:
            if v is None:
                continue
            w = min(100.0, (v / mx) * 100.0)
            parts.append(
                f'<div class="trio-col"><span class="trio-metric">{lab}</span>'
                f'<div class="trio-track"><span class="trio-fill" style="height:{w:.1f}%;background:{col}"></span></div>'
                f'<span class="trio-pct">{v:.2f}%</span></div>'
            )
        parts.append("</div></div>")
        return "".join(parts)

    market_trios = ""
    market_trios += trio_bars("MDR débito (non-promo)", deb["min"], deb["mean"], deb["max"], "#60a5fa", "#3b82f6", "#1d4ed8")
    market_trios += trio_bars("MDR crédito à vista", cred["min"], cred["mean"], cred["max"], "#86efac", "#22c55e", "#15803d")
    if pix["n"]:
        market_trios += trio_bars("MDR Pix", pix["min"], pix["mean"], pix["max"], "#fcd34d", "#f59e0b", "#b45309")

    # Prazo: mean credit vista by bucket
    prazo_chart = []
    if not by_prazo_df.empty and "mdr_credito_vista_rate_mean" in by_prazo_df.columns:
        sub = by_prazo_df.dropna(subset=["mdr_credito_vista_rate_mean"])
        vals = [_pct01(x) for x in sub["mdr_credito_vista_rate_mean"] if pd.notna(x)]
        vmax = max(vals) if vals else 5.0
        vmax = max(vmax * 1.1, 1.0)
        for _, row in sub.iterrows():
            bucket = str(row.get("prazo_liquidacao_aprox_dias", "?"))
            v = _pct01(row.get("mdr_credito_vista_rate_mean"))
            if v is None:
                continue
            prazo_chart.append(_hbar(f"Prazo {bucket} d", v, vmax, "#38bdf8", f"{v:.2f}%"))

    # Scenarios: NOTR estimated
    scen_bars = []
    scen_rows = []
    for _, r in scen_df.iterrows():
        name = str(r.get("scenario", ""))
        ne = r.get("notr_estimated")
        gp_e = r.get("gp_usd_month_estimated_linear")
        if pd.isna(ne):
            continue
        ne_p = float(ne) * 100.0
        scen_rows.append((name, ne_p, gp_e))
    if scen_rows:
        max_s = max(abs(x[1]) for x in scen_rows) * 1.1
        max_s = max(max_s, 0.01)
        palette = ["#22c55e", "#3b82f6", "#a855f7", "#f97316", "#ec4899", "#14b8a6"]
        for i, (name, ne_p, _) in enumerate(scen_rows):
            c = palette[i % len(palette)]
            scen_bars.append(_hbar(_short_scenario_name(name), ne_p, max_s, c, f"{ne_p:.2f}%"))

    # Executive bullets (plain language)
    cred_mean_p = _pct01(cred["mean"])
    exec_bullets = [
        f"<strong>BR Stone (2026 BC)</strong> shows net operating take rate <strong>~{stone_notr:.2f}%</strong> on credit-card economics — with <strong>lift NOTR ~{stone_lift:.2f}%</strong> vs the modeled baseline in the workbook.",
        f"Non-promo public tables cluster <strong>credit à vista</strong> around <strong>~{cred_mean_p:.2f}%</strong> average (min–max spread informs guardrails, not a single price).",
        f"The pricing view uses <strong>{n_non_promo}</strong> non-promo rows (promos excluded) so targets reflect sustainable list pricing, not acquisition promos.",
        "Charts below translate the spreadsheet into a <strong>few decision-ready pictures</strong>; methodology and full tables stay in the Excel export.",
    ]

    title = "PDV pricing & economics"
    year = "2026"

    prazo_block = (
        "".join(prazo_chart)
        if prazo_chart
        else '<p class="subtitle">No prazo slices with data.</p>'
    )
    scen_block = (
        "".join(scen_bars) if scen_bars else '<p class="subtitle">No scenarios.</p>'
    )

    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
  <meta http-equiv="Pragma" content="no-cache" />
  <meta name="pricing-deck-build" content="{_esc(stamp)}" />
  <title>{_esc(title)} — executive view</title>
  <!-- pricing-deck generated { _esc(stamp) } — if the site looks stale, hard-refresh (e.g. Cmd+Shift+R) -->
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap" rel="stylesheet" />
  <style>
    :root {{
      --color-primary: #0059d5;
      --color-primary-light: #3d7fe0;
      --color-dark: #000b19;
      --color-white: #ffffff;
      --color-gray-100: #f5f7fa;
      --color-gray-400: #9ca3af;
      --excel-header: #4472c4;
      --excel-border: #d0d7de;
      --excel-bg: #ffffff;
      --excel-alt: #f6f8fa;
      --nav-h: 3.25rem;
      --panel: rgba(255,255,255,0.05);
    }}
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior: smooth; }}
    body {{
      margin: 0;
      font-family: "Plus Jakarta Sans", -apple-system, BlinkMacSystemFont, sans-serif;
      background: var(--color-dark);
      color: var(--color-gray-100);
      line-height: 1.5;
      padding-top: var(--nav-h);
    }}
    .deck-nav {{
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      z-index: 100;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: center;
      gap: 0.35rem 0.5rem;
      min-height: var(--nav-h);
      padding: 0.5rem 0.75rem;
      background: rgba(0, 11, 25, 0.92);
      backdrop-filter: blur(10px);
      border-bottom: 1px solid rgba(255, 255, 255, 0.08);
      font-size: 0.72rem;
      font-weight: 600;
    }}
    .deck-nav a {{ color: var(--color-gray-400); text-decoration: none; white-space: nowrap; }}
    .deck-nav a:hover {{ color: var(--color-primary-light); }}
    .deck-nav a[aria-current="true"] {{
      color: var(--color-white);
      text-decoration: underline;
      text-underline-offset: 3px;
    }}
    .deck-nav .sep {{ color: rgba(255, 255, 255, 0.25); user-select: none; }}
    .slide {{
      min-height: calc(100vh - var(--nav-h));
      display: flex;
      align-items: center;
      justify-content: center;
      padding: clamp(1rem, 3vw, 2.5rem) 1rem;
      position: relative;
      overflow: hidden;
    }}
    .slide::before {{
      content: "";
      position: absolute;
      inset: 0;
      background:
        radial-gradient(ellipse 120% 80% at 15% 0%, rgba(0, 89, 213, 0.2) 0%, transparent 55%),
        radial-gradient(ellipse 90% 50% at 100% 100%, rgba(0, 68, 168, 0.15) 0%, transparent 45%),
        var(--color-dark);
      pointer-events: none;
    }}
    .slide-inner {{ position: relative; z-index: 1; width: 100%; max-width: 1100px; }}
    .kicker {{
      font-size: 0.75rem;
      font-weight: 700;
      letter-spacing: 0.1em;
      text-transform: uppercase;
      color: var(--color-primary-light);
      margin: 0 0 0.5rem;
    }}
    h1, h2 {{
      font-weight: 800;
      color: var(--color-white);
      line-height: 1.15;
      margin: 0 0 0.75rem;
    }}
    h1 {{ font-size: clamp(1.35rem, 2.8vw, 2rem); }}
    h2 {{ font-size: clamp(1.1rem, 2vw, 1.55rem); }}
    .subtitle {{
      color: var(--color-gray-400);
      font-size: clamp(0.88rem, 1.25vw, 1.02rem);
      margin: 0 0 1.25rem;
      max-width: 72ch;
    }}
    .insights {{
      margin: 0.5rem 0 1.25rem;
      padding: 0;
      list-style: none;
    }}
    .insights li {{
      position: relative;
      padding-left: 1.1rem;
      margin-bottom: 0.65rem;
      font-size: 0.92rem;
      color: #e8ecf0;
      line-height: 1.45;
    }}
    .insights li::before {{
      content: "";
      position: absolute;
      left: 0;
      top: 0.55em;
      width: 6px;
      height: 6px;
      border-radius: 50%;
      background: var(--color-primary-light);
    }}
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
      gap: 0.75rem;
      margin-top: 0.5rem;
    }}
    .kpi {{
      background: var(--panel);
      border: 1px solid rgba(255,255,255,0.1);
      border-radius: 12px;
      padding: 1rem 1rem 0.85rem;
      text-align: center;
    }}
    .kpi .big {{
      font-size: clamp(1.35rem, 2.5vw, 1.85rem);
      font-weight: 800;
      color: #fff;
      font-variant-numeric: tabular-nums;
      line-height: 1.2;
    }}
    .kpi .small {{ font-size: 0.72rem; color: var(--color-gray-400); margin-top: 0.35rem; }}
    .chart-panel {{
      background: rgba(255,255,255,0.04);
      border: 1px solid rgba(255,255,255,0.1);
      border-radius: 12px;
      padding: 1rem 1.1rem;
      margin-top: 0.75rem;
    }}
    .chart-title {{
      font-size: 0.72rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: var(--color-gray-400);
      margin: 0 0 0.75rem;
    }}
    .hbar {{
      display: grid;
      grid-template-columns: minmax(72px, 100px) 1fr 52px;
      align-items: center;
      gap: 0.5rem;
      margin-bottom: 0.45rem;
      font-size: 0.78rem;
    }}
    .hbar-label {{ color: #c9d1d9; }}
    .hbar-track {{
      height: 14px;
      background: rgba(255,255,255,0.08);
      border-radius: 6px;
      overflow: hidden;
    }}
    .hbar-fill {{ display: block; height: 100%; border-radius: 6px; transition: width 0.3s ease; }}
    .hbar-val {{ text-align: right; font-variant-numeric: tabular-nums; color: #fff; font-weight: 700; }}
    .trio-wrap {{
      display: flex;
      flex-wrap: wrap;
      gap: 1rem;
      justify-content: flex-start;
    }}
    .trio {{
      flex: 1 1 200px;
      min-width: 180px;
      background: rgba(0,0,0,0.2);
      border-radius: 10px;
      padding: 0.75rem;
    }}
    .trio-title {{ font-size: 0.8rem; font-weight: 700; color: #fff; margin: 0 0 0.5rem; }}
    .trio-bars {{
      display: flex;
      align-items: flex-end;
      justify-content: space-around;
      gap: 0.5rem;
      height: 120px;
    }}
    .trio-col {{
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 0.25rem;
    }}
    .trio-metric {{ font-size: 0.65rem; color: var(--color-gray-400); }}
    .trio-track {{
      width: 100%;
      max-width: 36px;
      height: 90px;
      background: rgba(255,255,255,0.06);
      border-radius: 6px 6px 2px 2px;
      display: flex;
      align-items: flex-end;
      justify-content: center;
      overflow: hidden;
    }}
    .trio-fill {{ display: block; width: 100%; border-radius: 4px 4px 0 0; min-height: 4px; }}
    .trio-pct {{ font-size: 0.72rem; font-weight: 700; color: #fff; }}
    .split-2 {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 1rem;
    }}
    @media (max-width: 800px) {{ .split-2 {{ grid-template-columns: 1fr; }} }}
    .footnote {{
      font-size: 0.72rem;
      color: var(--color-gray-400);
      margin-top: 1rem;
      max-width: 85ch;
    }}
    .excel-wrap {{
      background: var(--excel-bg);
      border-radius: 10px;
      padding: 0.65rem 0.75rem;
      overflow-x: auto;
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.35);
      border: 1px solid rgba(255, 255, 255, 0.12);
      max-height: 70vh;
      overflow-y: auto;
    }}
    .excel-wrap .caption {{
      font-size: 0.7rem;
      color: #57606a;
      margin: 0 0 0.5rem;
      font-weight: 600;
    }}
    .excel-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.65rem;
      font-variant-numeric: tabular-nums;
      color: #1f2328;
    }}
    .excel-table th, .excel-table td {{ border: 1px solid var(--excel-border); padding: 0.3rem 0.4rem; }}
    .excel-table thead th {{
      background: var(--excel-header);
      color: #fff;
      font-weight: 700;
      text-align: center;
    }}
    .excel-table thead th.corner {{ background: #3b5998; text-align: center; vertical-align: middle; }}
    .excel-table thead th.excel-h2 {{ background: #4a6fb5; }}
    .excel-table thead th.stone {{ background: #2d6a4f; color: #fff; }}
    .excel-table td.excel-metric {{ background: #f3f4f6; color: #374151; text-align: left; font-weight: 500; }}
    .excel-table tbody tr:nth-child(even) td:not(.excel-metric) {{ background: var(--excel-alt); }}
    .excel-table td.excel-num {{ text-align: right; }}
    .excel-table td.stone-col {{ background: rgba(34, 197, 94, 0.18) !important; font-weight: 600; }}
    .excel-table tr.spacer td {{ height: 0.35rem; border: none; background: transparent !important; }}
    footer.slide {{
      min-height: auto;
      padding: 2rem 1rem;
      font-size: 0.8rem;
      color: var(--color-gray-400);
    }}
  </style>
</head>
<body>
  <nav class="deck-nav" aria-label="Slides">
    <a href="#slide-0" data-nav>0 · Summary</a><span class="sep">|</span>
    <a href="#slide-1" data-nav>1 · Regional NOTR</a><span class="sep">|</span>
    <a href="#slide-2" data-nav>2 · Market bands</a><span class="sep">|</span>
    <a href="#slide-3" data-nav>3 · Prazo</a><span class="sep">|</span>
    <a href="#slide-4" data-nav>4 · Scenarios</a><span class="sep">|</span>
    <a href="#slide-5" data-nav>5 · Appendix</a>
  </nav>

  <section class="slide" id="slide-0" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">{_esc(year)} BC · Executive readout</p>
      <h1>BR Stone economics sit in range — market tables set the guardrails</h1>
      <p class="subtitle">Results first: net take, lift, and gross profit from the business case; then charts for cross-country NOTR, market min/avg/max, settlement horizon, and scenario stress tests. Benchmark narrative style: <a href="https://gopedroso.github.io/PoS/GTM-Sales-investment/" style="color:var(--color-primary-light)">PoS GTM Sales Investment</a> (executive summary + charts).</p>
      <h2 style="font-size:1rem;margin-top:0.5rem;">Key insights</h2>
      <ul class="insights">
        {"".join(f"<li>{b}</li>" for b in exec_bullets)}
      </ul>
      <div class="kpi-grid">
        <div class="kpi"><div class="big">{stone_notr:.2f}%</div><div class="small">Stone NOTR % (BR, BC)</div></div>
        <div class="kpi"><div class="big">{stone_lift:.2f}%</div><div class="small">Lift NOTR % (BC)</div></div>
        <div class="kpi"><div class="big">${stone_gp:,.0f}</div><div class="small">GP USD / month (Stone)</div></div>
        <div class="kpi"><div class="big">{cred_mean_p:.2f}%</div><div class="small">Market avg · crédito à vista</div></div>
      </div>
      <p class="footnote">Promos excluded ({n_promo} rows flagged); pricing base = {n_non_promo} non-promo rows. Suggested anchors use ×(1−{eps}) vs min/mean/max.</p>
    </div>
  </section>

  <section class="slide" id="slide-1" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Only CC · Cross-country</p>
      <h2>Net operating take rate is not uniform — Brazil Stone vs peers in the BC</h2>
      <p class="subtitle">Horizontal bars = NOTR % from the <strong>2026 BC</strong> sheet (same row as your Excel). <span style="color:#86efac">Stone</span> is the strategic BR anchor.</p>
      <div class="chart-panel">
        <p class="chart-title">NOTR % by provider (model export)</p>
        {"".join(regional_bars)}
      </div>
      <p class="footnote">Source: PDV Payments Business Case.xlsx · tab 2026 BC · read-only in the build.</p>
    </div>
  </section>

  <section class="slide" id="slide-2" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Non-promo market</p>
      <h2>MDR spreads show where list pricing actually lives</h2>
      <p class="subtitle">Min / average / max from <code>competitor-pricing.csv</code> (no promos). Use bands for sensitivity — not a single “right” number.</p>
      <div class="chart-panel">
        <p class="chart-title">Tri-band comparison (% of GMV)</p>
        <div class="trio-wrap">{market_trios}</div>
      </div>
    </div>
  </section>

  <section class="slide" id="slide-3" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Settlement horizon</p>
      <h2>Prazo shifts the credit à vista mean — price sensitivity follows</h2>
      <p class="subtitle">Grouped by <code>prazo_liquidacao_aprox_dias</code> (D+0 vs D+14 / D+30 policies, etc.).</p>
      <div class="chart-panel">
        <p class="chart-title">Mean MDR crédito à vista by prazo bucket</p>
        {prazo_block}
      </div>
    </div>
  </section>

  <section class="slide" id="slide-4" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Stress tests</p>
      <h2>Scenario NOTR moves when fees move — costs fixed from BC</h2>
      <p class="subtitle">Illustrative linear GP scaling in the workbook; chart shows <strong>estimated NOTR %</strong> per scenario.</p>
      <div class="chart-panel">
        <p class="chart-title">NOTR % (estimated) by scenario</p>
        {scen_block}
      </div>
      <p class="footnote">Full scenario table and GP estimates: <strong>pricing-analysis-economics.xlsx</strong>.</p>
    </div>
  </section>

  <section class="slide" id="slide-5" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Appendix</p>
      <h2>Methodology + full BC matrix</h2>
      <p class="subtitle"><strong>How we got here:</strong> promo filter on scenario IDs; coalesced anticipation (mensal → pontual); BC Stone column D for fee inputs; NOTR ≈ NTR sum − rent/logistics/taxes. Below: the same <em>Only CC</em> table as your spreadsheet for audit.</p>
      <div class="excel-wrap">
        <p class="caption">2026 BC — Payments take rate (credit card only)</p>
        <table class="excel-table">{bc_table_inner}</table>
      </div>
      <p class="footnote">Regenerate: <code style="color:var(--color-primary-light)">python3 build_pricing_analysis.py</code></p>
    </div>
  </section>

  <footer class="slide">
    <div class="slide-inner">
      <p style="margin:0;">Static HTML — no charting SDK. Bars are CSS. For a comparable executive narrative flow see the PoS GTM reference deck.</p>
    </div>
  </footer>

  <script>
    (function () {{
      var links = document.querySelectorAll("[data-nav]");
      var ids = Array.prototype.map.call(links, function (a) {{ return a.getAttribute("href").slice(1); }});
      function setCurrent() {{
        var y = window.scrollY + 80;
        var current = ids[0];
        for (var i = 0; i < ids.length; i++) {{
          var el = document.getElementById(ids[i]);
          if (el && el.offsetTop <= y) current = ids[i];
        }}
        links.forEach(function (a) {{
          a.setAttribute("aria-current", a.getAttribute("href") === "#" + current ? "true" : "false");
        }});
      }}
      window.addEventListener("scroll", setCurrent, {{ passive: true }});
      setCurrent();
    }})();
  </script>
</body>
</html>
"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc)
