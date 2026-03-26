# Pricing & economics (static HTML)

Same presentation pattern as `gtm-sales-deck/`: full-height slides, top nav, Excel-style BC table on a dark theme.

## Live site (GitHub Pages)

If this repo has Pages enabled from **main** / **root**, the deck is at:

`https://<username>.github.io/<repo>/pricing-deck/`

## Regenerate

From `pos-pricing-model/` (requires `PDV Payments Business Case.xlsx` locally — not committed):

```bash
PYTHONPATH=_pydeps python3 build_pricing_analysis.py
```

That updates `pricing-presentation.html` and copies to **`pricing-deck/index.html`**.

## Files

- `index.html` — generated; do not hand-edit for data (edit the Python source instead).
- `.nojekyll` — tells GitHub Pages not to use Jekyll on this folder.
