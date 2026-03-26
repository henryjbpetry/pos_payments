#!/usr/bin/env python3
"""
Build pricing-analysis-economics.xlsx from competitor-pricing.csv + BC assumptions.
Does not modify PDV Payments Business Case.xlsx (read-only).
"""
from __future__ import annotations

import os
import shutil
import sys
from datetime import datetime, timezone

_DEPS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_pydeps")
if os.path.isdir(_DEPS):
    sys.path.insert(0, _DEPS)

import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.normpath(os.path.join(ROOT, ".."))
PRICING_DECK_DIR = os.path.join(REPO_ROOT, "pricing-deck")
PRICING_DECK_INDEX = os.path.join(PRICING_DECK_DIR, "index.html")
CSV_PATH = os.path.join(ROOT, "competitor-pricing.csv")
BC_PATH = os.path.join(ROOT, "PDV Payments Business Case.xlsx")
OUT_PATH = os.path.join(ROOT, "pricing-analysis-economics.xlsx")
OUT_HTML = os.path.join(ROOT, "pricing-presentation.html")

# --- Promo detection (non-promo is base for pricing) ---
PROMO_SCENARIO_SUBSTR = ("_promo", "promo_home", "point_promo")
PROMO_PLAN_SUBSTR = ("promo", "30d ou", "5k faturado")


def is_promo_row(row: pd.Series) -> bool:
    sid = str(row.get("scenario_id", "") or "").lower()
    pl = str(row.get("plan_label", "") or "").lower()
    for s in PROMO_SCENARIO_SUBSTR:
        if s in sid:
            return True
    for s in PROMO_PLAN_SUBSTR:
        if s in pl:
            return True
    if "promo" in pl and "apos promo" not in pl:
        return True
    return False


def prazo_bucket(v) -> str:
    if pd.isna(v):
        return "unknown"
    try:
        d = int(float(v))
    except (TypeError, ValueError):
        return "unknown"
    return str(d)


def read_bc_stone() -> dict:
    """Column D = Stone on '2026 BC' tab (index 3)."""
    df = pd.read_excel(BC_PATH, sheet_name="2026 BC", header=None)
    # Row 15: processing fee, col 3 = Stone
    # Row 16: anticipation fee
    def cell(r: int, c: int):
        return df.iloc[r, c]

    return {
        "processing_fee": float(cell(15, 3)),
        "anticipation_fee": float(cell(16, 3)),
        "processing_cost": float(cell(18, 3)),
        "anticipation_cost": float(cell(19, 3)),
        "ntr_processing": float(cell(21, 3)),
        "ntr_anticipation": float(cell(22, 3)),
        "rental_costs": float(cell(24, 3)),
        "logistics_costs": float(cell(25, 3)),
        "taxes": float(cell(28, 3)),
        "notr_pct": float(cell(33, 3)),
        "lift_notr_pct": float(cell(34, 3)),
        "gp_usd_month": float(cell(35, 3)),
        "lift_gp_usd_month": float(cell(36, 3)),
    }


def main() -> None:
    df = pd.read_csv(CSV_PATH)
    df["is_promo"] = df.apply(is_promo_row, axis=1)
    df["prazo_bucket"] = df["prazo_liquidacao_aprox_dias"].apply(prazo_bucket)

    def _num(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s, errors="coerce")

    # Public tables often put "3.7%/mês" in pontual; use first non-null for market stats.
    df["anticipation_rate_coalesced"] = _num(df["taxa_antecipacao_mensal_rate"]).combine_first(
        _num(df["taxa_antecipacao_pontual_rate"])
    )

    non = df[~df["is_promo"]].copy()
    promo = df[df["is_promo"]].copy()

    stat_cols = [
        "mdr_debito_rate",
        "mdr_credito_vista_rate",
        "mdr_credito_parcelado_12x_total_rate",
        "mdr_credito_parcelado_6x_total_rate",
        "mdr_parcela_12x_rate",
        "mdr_pix_rate",
        "taxa_antecipacao_mensal_rate",
        "taxa_antecipacao_pontual_rate",
        "anticipation_rate_coalesced",
        "aluguel_brl_mes",
    ]

    def stats_table(frame: pd.DataFrame) -> pd.DataFrame:
        rows = []
        for c in stat_cols:
            s = pd.to_numeric(frame[c], errors="coerce").dropna()
            if s.empty:
                rows.append(
                    {
                        "metric": c,
                        "n": 0,
                        "min": np.nan,
                        "mean": np.nan,
                        "max": np.nan,
                    }
                )
            else:
                rows.append(
                    {
                        "metric": c,
                        "n": int(s.shape[0]),
                        "min": float(s.min()),
                        "mean": float(s.mean()),
                        "max": float(s.max()),
                    }
                )
        return pd.DataFrame(rows)

    summary_non_promo = stats_table(non)

    # By prazo: count rows + mean key rates
    agg_cols = [
        "mdr_debito_rate",
        "mdr_credito_vista_rate",
        "mdr_credito_parcelado_12x_total_rate",
        "mdr_pix_rate",
        "anticipation_rate_coalesced",
        "aluguel_brl_mes",
    ]
    by_prazo = []
    for p, g in non.groupby("prazo_bucket", sort=True):
        row = {"prazo_liquidacao_aprox_dias": p, "n_rows": len(g)}
        for c in agg_cols:
            s = pd.to_numeric(g[c], errors="coerce").dropna()
            row[f"{c}_mean"] = float(s.mean()) if not s.empty else np.nan
            row[f"{c}_min"] = float(s.min()) if not s.empty else np.nan
            row[f"{c}_max"] = float(s.max()) if not s.empty else np.nan
        by_prazo.append(row)
    by_prazo_df = pd.DataFrame(by_prazo)

    # Suggested prices: "slightly below" = multiply by (1 - epsilon)
    EPS = 0.002  # 20 bps relative; user can tune
    sug = []
    for _, r in summary_non_promo.iterrows():
        m = r["metric"]
        if r["n"] == 0 or pd.isna(r["min"]):
            sug.append(
                {
                    "metric": m,
                    "anchor": "below_min",
                    "suggested_rate": np.nan,
                    "rule": f"min * (1 - {EPS})",
                }
            )
            sug.append(
                {
                    "metric": m,
                    "anchor": "below_avg",
                    "suggested_rate": np.nan,
                    "rule": f"mean * (1 - {EPS})",
                }
            )
            sug.append(
                {
                    "metric": m,
                    "anchor": "below_max",
                    "suggested_rate": np.nan,
                    "rule": f"max * (1 - {EPS})",
                }
            )
            continue
        sug.append(
            {
                "metric": m,
                "anchor": "below_min",
                "suggested_rate": float(r["min"]) * (1 - EPS),
                "rule": f"min * (1 - {EPS})",
            }
        )
        sug.append(
            {
                "metric": m,
                "anchor": "below_avg",
                "suggested_rate": float(r["mean"]) * (1 - EPS),
                "rule": f"mean * (1 - {EPS})",
            }
        )
        sug.append(
            {
                "metric": m,
                "anchor": "below_max",
                "suggested_rate": float(r["max"]) * (1 - EPS),
                "rule": f"max * (1 - {EPS})",
            }
        )
    suggested_df = pd.DataFrame(sug)

    stone = read_bc_stone()

    # Economics scenarios: vary processing + anticipation fees; costs fixed from BC
    # Map market anchors to fee rows (decimals)
    def pick_summary(metric: str) -> dict:
        row = summary_non_promo.loc[summary_non_promo["metric"] == metric]
        if row.empty:
            return {"n": 0, "min": np.nan, "mean": np.nan, "max": np.nan}
        row = row.iloc[0]
        return {
            "n": int(row["n"]),
            "min": row["min"],
            "mean": row["mean"],
            "max": row["max"],
        }

    debit = pick_summary("mdr_debito_rate")
    cred = pick_summary("mdr_credito_vista_rate")
    ant = pick_summary("anticipation_rate_coalesced")

    # Use debit mean as proxy "processing fee" for CC if single number needed; document split
    proc_candidates = {
        "market_min_debito": debit["min"],
        "market_avg_debito": debit["mean"],
        "market_max_debito": debit["max"],
        "market_min_credito_vista": cred["min"],
        "market_avg_credito_vista": cred["mean"],
        "market_max_credito_vista": cred["max"],
        "market_min_anticipation_coalesced": ant["min"],
        "market_avg_anticipation_coalesced": ant["mean"],
        "market_max_anticipation_coalesced": ant["max"],
    }

    scenarios = []
    base_pf = stone["processing_fee"]
    base_af = stone["anticipation_fee"]

    # Fixed opex from BC (same sheet): NOTR ≈ NTR_proc + NTR_ant - rental - logistics - taxes
    opex_fixed = stone["rental_costs"] + stone["logistics_costs"] + stone["taxes"]

    def append_scenario(name: str, pf: float, af: float) -> None:
        ntr_p = pf - stone["processing_cost"]
        ntr_a = af - stone["anticipation_cost"]
        ntr_sum = ntr_p + ntr_a
        notr_est = ntr_sum - opex_fixed
        notr_bc = stone["notr_pct"]
        gp_bc = stone["gp_usd_month"]
        gp_est = gp_bc * (notr_est / notr_bc) if notr_bc else np.nan
        scenarios.append(
            {
                "scenario": name,
                "processing_fee_assumption": pf,
                "anticipation_fee_assumption": af,
                "processing_cost_bc": stone["processing_cost"],
                "anticipation_cost_bc": stone["anticipation_cost"],
                "ntr_processing": ntr_p,
                "ntr_anticipation": ntr_a,
                "ntr_sum": ntr_sum,
                "opex_fixed_rent_log_tax": opex_fixed,
                "notr_estimated": notr_est,
                "gp_usd_month_estimated_linear": gp_est,
                "notr_bc_reference": notr_bc,
                "gp_usd_month_bc": gp_bc,
            }
        )

    append_scenario("BC_baseline_stone", base_pf, base_af)

    # "Only CC" row: processing fee is a blended headline; we model it as **credito à vista** for market comparison.
    cred_min = float(cred["min"]) * (1 - EPS) if cred["n"] else np.nan
    cred_avg = float(cred["mean"]) * (1 - EPS) if cred["n"] else np.nan
    cred_max = float(cred["max"]) * (1 - EPS) if cred["n"] else np.nan
    ant_min = float(ant["min"]) * (1 - EPS) if ant["n"] else np.nan
    ant_avg = float(ant["mean"]) * (1 - EPS) if ant["n"] else np.nan
    ant_max = float(ant["max"]) * (1 - EPS) if ant["n"] else np.nan

    for label, pf in [("below_market_min", cred_min), ("below_market_avg", cred_avg), ("below_market_max", cred_max)]:
        if np.isnan(pf):
            continue
        append_scenario(f"suggested_{label}_proc_only_hold_bc_ant", pf, base_af)

    for label, af in [("below_market_min", ant_min), ("below_market_avg", ant_avg), ("below_market_max", ant_max)]:
        if np.isnan(af):
            continue
        append_scenario(f"suggested_{label}_ant_only_hold_bc_proc", base_pf, af)

    if cred["n"] and ant["n"]:
        append_scenario(
            "suggested_below_avg_credito_and_avg_anticipation",
            float(cred["mean"]) * (1 - EPS),
            float(ant["mean"]) * (1 - EPS),
        )

    scen_df = pd.DataFrame(scenarios)
    # When market anticipation has n=1, min/avg/max scenarios collapse to identical fees.
    scen_df = scen_df.drop_duplicates(
        subset=["processing_fee_assumption", "anticipation_fee_assumption"], keep="first"
    )

    notes = pd.DataFrame(
        {
            "topic": [
                "Promo filter",
                "Prazo grouping",
                "Rates units",
                "Anticipation stats",
                "BC source",
                "Suggested price rule",
                "Economics scenarios",
            ],
            "detail": [
                "Rows flagged promo via scenario_id/plan_label heuristics (e.g. stone_promo_home, mp_point_promo). Base pricing uses non-promo rows only.",
                "prazo_liquidacao_aprox_dias from CSV (0=D+0 style, 14/30=credit policy days, etc.). Grouping is settlement horizon, not a separate anticipation-tenor field.",
                "All rates are decimals on [0,1] in CSV; BC sheet is also decimal.",
                "anticipation_rate_coalesced = mensal if present else pontual (Getnet example stores 3.7% under pontual in the CSV). n is often small.",
                "Stone BR column D on sheet '2026 BC': processing_fee ~row 16, anticipation_fee ~row 17 (see bc_stone_br_col_D). PDV Payments Business Case.xlsx was not modified.",
                f"Slightly below = multiply competitive anchor by (1 - {EPS}) (~20 bps relative).",
                "notr_estimated = ntr_processing + ntr_anticipation - rental - logistics - taxes (Stone BC opex rows). gp_usd_month_estimated_linear = gp_bc * (notr_estimated / notr_bc); illustrative only (constant GMV, no elasticity).",
            ],
        }
    )

    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="all_rows_flagged", index=False)
        promo.to_excel(writer, sheet_name="promo_only", index=False)
        non.to_excel(writer, sheet_name="non_promo_base", index=False)
        by_prazo_df.to_excel(writer, sheet_name="non_promo_by_prazo", index=False)
        summary_non_promo.to_excel(writer, sheet_name="summary_min_avg_max", index=False)
        suggested_df.to_excel(writer, sheet_name="suggested_prices", index=False)
        pd.DataFrame([stone]).T.rename(columns={0: "value"}).to_excel(
            writer, sheet_name="bc_stone_br_col_D"
        )
        scen_df.to_excel(writer, sheet_name="economics_scenarios", index=False)
        pd.DataFrame([proc_candidates]).T.rename(columns={0: "value"}).to_excel(
            writer, sheet_name="market_anchors_processing"
        )
        notes.to_excel(writer, sheet_name="methodology_notes", index=False)

    print(f"Wrote {OUT_PATH}")

    from pricing_presentation_html import write_presentation_html

    write_presentation_html(
        OUT_HTML,
        BC_PATH,
        summary_non_promo,
        by_prazo_df,
        scen_df,
        suggested_df,
        stone,
        n_promo=len(promo),
        n_non_promo=len(non),
        eps=EPS,
        build_stamp=datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    )
    print(f"Wrote {OUT_HTML}")

    os.makedirs(PRICING_DECK_DIR, exist_ok=True)
    shutil.copy2(OUT_HTML, PRICING_DECK_INDEX)
    nojek = os.path.join(PRICING_DECK_DIR, ".nojekyll")
    if not os.path.isfile(nojek):
        with open(nojek, "w", encoding="utf-8"):
            pass
    print(f"Wrote {PRICING_DECK_INDEX}")


if __name__ == "__main__":
    main()
