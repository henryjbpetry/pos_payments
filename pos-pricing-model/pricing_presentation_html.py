"""
Generate pricing-presentation.html — same deck styling as gtm-sales-deck + Excel-style BC table.
"""
from __future__ import annotations

import html
from typing import Any

import pandas as pd


def _esc(x: Any) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return html.escape(str(x))


def _fmt_bc_value(label: str, val: Any) -> str:
    """Format BC sheet values: rates as %, Capex/GP as USD."""
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


def write_presentation_html(
    out_path: str,
    bc_path: str,
    summary_non_promo: pd.DataFrame,
    by_prazo_df: pd.DataFrame,
    scen_df: pd.DataFrame,
    suggested_df: pd.DataFrame,
    n_promo: int,
    n_non_promo: int,
    eps: float,
) -> None:
    regions, providers, bc_rows = _read_bc_matrix(bc_path)
    bc_table_inner = _build_bc_table_html(regions, providers, bc_rows)

    metric_labels = {
        "mdr_debito_rate": "MDR débito",
        "mdr_credito_vista_rate": "MDR crédito à vista",
        "mdr_credito_parcelado_12x_total_rate": "MDR crédito 12x (total)",
        "mdr_pix_rate": "MDR Pix",
        "anticipation_rate_coalesced": "Antecipação (mensal → pontual)",
        "aluguel_brl_mes": "Aluguel (BRL/mês)",
    }

    summary_rows = []
    for _, row in summary_non_promo.iterrows():
        m = row["metric"]
        if m not in metric_labels:
            continue
        summary_rows.append(
            (
                metric_labels[m],
                int(row["n"]),
                row["min"],
                row["mean"],
                row["max"],
            )
        )

    def fmt_rate(x: Any) -> str:
        if x is None or pd.isna(x):
            return "—"
        return f"{float(x) * 100:.2f}%"

    sum_html = [
        '<table class="excel-table excel-table--compact"><thead><tr><th>Metric</th><th>n</th><th>Min</th><th>Avg</th><th>Max</th></tr></thead><tbody>'
    ]
    for name, n, lo, mid, hi in summary_rows:
        sum_html.append(
            f"<tr><td class=\"excel-metric\">{_esc(name)}</td>"
            f"<td class=\"excel-num\">{n}</td>"
            f"<td class=\"excel-num\">{fmt_rate(lo)}</td>"
            f"<td class=\"excel-num\">{fmt_rate(mid)}</td>"
            f"<td class=\"excel-num\">{fmt_rate(hi)}</td></tr>"
        )
    sum_html.append("</tbody></table>")
    summary_table = "".join(sum_html)

    prazo_preview = by_prazo_df.head(8).to_html(
        classes="excel-table excel-table--compact",
        index=False,
        float_format=lambda x: f"{x:.4f}" if pd.notna(x) else "",
    )

    sug_sample = suggested_df.head(18).copy()
    sug_html = sug_sample.to_html(
        classes="excel-table excel-table--compact",
        index=False,
        float_format=lambda x: f"{float(x):.4f}" if pd.notna(x) else "",
    )

    econ_html = scen_df.to_html(
        classes="excel-table excel-table--compact",
        index=False,
        float_format=lambda x: f"{float(x):.6f}" if pd.notna(x) else "",
    )

    title = "PDV pricing & economics"
    year = "2026"

    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{_esc(title)}</title>
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
      --color-gray-200: #e8ecf0;
      --color-gray-400: #9ca3af;
      --excel-header: #4472c4;
      --excel-border: #d0d7de;
      --excel-bg: #ffffff;
      --excel-alt: #f6f8fa;
      --nav-h: 3.25rem;
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
    .deck-nav a {{
      color: var(--color-gray-400);
      text-decoration: none;
      white-space: nowrap;
    }}
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
    h1 {{ font-size: clamp(1.45rem, 3vw, 2.1rem); }}
    h2 {{ font-size: clamp(1.2rem, 2.2vw, 1.65rem); }}
    .subtitle {{
      color: var(--color-gray-400);
      font-size: clamp(0.9rem, 1.3vw, 1.05rem);
      margin: 0 0 1.25rem;
      max-width: 70ch;
    }}
    .excel-wrap {{
      background: var(--excel-bg);
      border-radius: 10px;
      padding: 0.65rem 0.75rem;
      overflow-x: auto;
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.35);
      border: 1px solid rgba(255, 255, 255, 0.12);
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
      font-size: 0.72rem;
      font-variant-numeric: tabular-nums;
      color: #1f2328;
    }}
    .excel-table th, .excel-table td {{
      border: 1px solid var(--excel-border);
      padding: 0.35rem 0.45rem;
    }}
    .excel-table thead th {{
      background: var(--excel-header);
      color: #fff;
      font-weight: 700;
      text-align: center;
    }}
    .excel-table thead th.corner {{
      background: #3b5998;
      text-align: center;
      vertical-align: middle;
    }}
    .excel-table thead th.excel-h2 {{ background: #4a6fb5; }}
    .excel-table thead th.stone {{
      background: #2d6a4f;
      color: #fff;
    }}
    .excel-table td.excel-metric {{
      background: #f3f4f6;
      color: #374151;
      text-align: left;
      font-weight: 500;
      white-space: nowrap;
    }}
    .excel-table tbody tr:nth-child(even) td:not(.excel-metric) {{ background: var(--excel-alt); }}
    .excel-table td.excel-num {{ text-align: right; }}
    .excel-table td.stone-col {{
      background: rgba(34, 197, 94, 0.18) !important;
      font-weight: 600;
    }}
    .excel-table tr.spacer td {{
      height: 0.35rem;
      border: none;
      background: transparent !important;
    }}
    .excel-table--compact {{ font-size: 0.68rem; }}
    .excel-table--compact thead th {{ background: #5b7fc7; }}
    .matrix {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.82rem;
      margin-top: 0.5rem;
    }}
    .matrix th, .matrix td {{
      border: 1px solid rgba(255, 255, 255, 0.12);
      padding: 0.45rem 0.55rem;
      text-align: center;
    }}
    .matrix th {{ background: rgba(0, 89, 213, 0.25); color: var(--color-white); font-weight: 700; }}
    .matrix td {{ color: var(--color-gray-200); }}
    footer.slide {{
      min-height: auto;
      padding: 2rem 1rem;
      font-size: 0.8rem;
      color: var(--color-gray-400);
    }}
    footer.slide::before {{ opacity: 0.5; }}
  </style>
</head>
<body>
  <nav class="deck-nav" aria-label="Slides">
    <a href="#slide-0" data-nav>0 · Title</a><span class="sep">|</span>
    <a href="#slide-1" data-nav>1 · {year} BC</a><span class="sep">|</span>
    <a href="#slide-2" data-nav>2 · Market</a><span class="sep">|</span>
    <a href="#slide-3" data-nav>3 · Prazo</a><span class="sep">|</span>
    <a href="#slide-4" data-nav>4 · Suggested</a><span class="sep">|</span>
    <a href="#slide-5" data-nav>5 · Economics</a>
  </nav>

  <section class="slide" id="slide-0" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">POS / PDV · Pricing intelligence</p>
      <h1>{_esc(title)}</h1>
      <p class="subtitle">Presentation-style view aligned with the static GTM deck template (typography, dark slides, top nav). Competitor stats exclude promos; Stone BR economics reference column D on the <strong>{_esc(year)} BC</strong> workbook.</p>
      <table class="matrix" aria-label="Highlights">
        <thead><tr><th>Dataset</th><th>Promo rows</th><th>Non-promo (base)</th><th>Suggested discount rule</th></tr></thead>
        <tbody>
          <tr><td>competitor-pricing.csv</td><td>{n_promo}</td><td>{n_non_promo}</td><td>× (1 − {eps}) vs min / mean / max</td></tr>
        </tbody>
      </table>
    </div>
  </section>

  <section class="slide" id="slide-1" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Business case · Only CC</p>
      <h2>{_esc(year)} BC — spreadsheet layout</h2>
      <p class="subtitle">Two-row header (region + provider). <strong>Stone</strong> uses green header + column tint (Excel column D). Read-only from PDV Payments Business Case.xlsx.</p>
      <div class="excel-wrap">
        <p class="caption">Payments take rate scenarios — credit card only</p>
        <table class="excel-table">{bc_table_inner}</table>
      </div>
    </div>
  </section>

  <section class="slide" id="slide-2" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Non-promo market</p>
      <h2>Summary — min / average / max</h2>
      <p class="subtitle">Core MDR and fee columns (CSV decimals shown as percentages).</p>
      <div class="excel-wrap">
        <p class="caption">competitor-pricing.csv · non-promo rows only</p>
        {summary_table}
      </div>
    </div>
  </section>

  <section class="slide" id="slide-3" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Settlement horizon</p>
      <h2>By prazo (liquidation days)</h2>
      <p class="subtitle">Grouped by <code>prazo_liquidacao_aprox_dias</code> (0 = D+0 style, 14 / 30 = policy days, etc.).</p>
      <div class="excel-wrap">
        <p class="caption">Preview · full table in pricing-analysis-economics.xlsx</p>
        {prazo_preview}
      </div>
    </div>
  </section>

  <section class="slide" id="slide-4" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Pricing anchors</p>
      <h2>Suggested prices (slightly below market)</h2>
      <p class="subtitle">Same rule as the workbook: each anchor × (1 − {eps}).</p>
      <div class="excel-wrap">
        <p class="caption">Sample · full long table in workbook</p>
        {sug_html}
      </div>
    </div>
  </section>

  <section class="slide" id="slide-5" tabindex="-1">
    <div class="slide-inner">
      <p class="kicker">Stone BR · scenarios</p>
      <h2>Economics scenarios (BC costs fixed)</h2>
      <p class="subtitle">Illustrative NOTR and linear GP scaling vs baseline — see methodology_notes in the Excel export.</p>
      <div class="excel-wrap">
        <p class="caption">economics_scenarios</p>
        {econ_html}
      </div>
    </div>
  </section>

  <footer class="slide">
    <div class="slide-inner">
      <p style="margin:0;">Regenerate with <code style="color:var(--color-primary-light)">python3 build_pricing_analysis.py</code> (writes the workbook and this HTML). Deck chrome matches <strong>gtm-sales-deck</strong>; data tables use a light Excel-style panel on dark slides.</p>
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
