#!/usr/bin/env python3
"""
EVM Report - Project Controls Toolkit
- Reads time-phased PV/EV/AC from a CSV
- Computes SPI, CPI, SV, CV, BAC, EAC, ETC, VAC
- Exports a detailed CSV + prints an executive summary

Input CSV columns (required): Date, PV, EV, AC
Date format: YYYY-MM-DD (recommended)
PV/EV/AC: numeric (currency units)

Usage:
  python scripts/evm_report.py --input data/evm_input.csv --output data/evm_output.csv
Optional:
  --bac 120000
"""

from __future__ import annotations

import argparse
import csv
import math
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple


REQUIRED_COLS = ["Date", "PV", "EV", "AC"]


def _to_float(x: str) -> float:
    # Accept "12,345.67" or "12.345,67" style (best effort)
    s = str(x).strip()
    if s == "":
        return float("nan")
    # If it contains both separators, assume last one is decimal
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        # If only comma, treat as decimal comma
        if "," in s and "." not in s:
            s = s.replace(",", ".")
    return float(s)


def _safe_div(a: float, b: float) -> float:
    if b == 0 or math.isclose(b, 0.0):
        return float("nan")
    return a / b


def _parse_date(d: str) -> datetime:
    # Accept YYYY-MM-DD or DD/MM/YYYY
    s = d.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError(
        f"Unrecognized Date format: '{d}'. Use YYYY-MM-DD (recommended).")


@dataclass
class EVMRow:
    date: datetime
    pv: float
    ev: float
    ac: float


def read_evm_csv(path: Path) -> List[EVMRow]:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    rows: List[EVMRow] = []
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("CSV has no header row.")

        missing = [c for c in REQUIRED_COLS if c not in reader.fieldnames]
        if missing:
            raise ValueError(
                f"Missing required columns: {missing}. Found: {reader.fieldnames}")

        for i, r in enumerate(reader, start=2):
            try:
                date = _parse_date(r["Date"])
                pv = _to_float(r["PV"])
                ev = _to_float(r["EV"])
                ac = _to_float(r["AC"])
            except Exception as e:
                raise ValueError(f"Error parsing row {i}: {e}") from e

            if any(math.isnan(v) for v in (pv, ev, ac)):
                raise ValueError(
                    f"Row {i} has empty/non-numeric PV/EV/AC: {r}")

            rows.append(EVMRow(date=date, pv=pv, ev=ev, ac=ac))

    # Sort by date
    rows.sort(key=lambda x: x.date)
    return rows


def compute_evm(rows: List[EVMRow], bac: Optional[float]) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    cum_pv = 0.0
    cum_ev = 0.0
    cum_ac = 0.0

    out: List[Dict[str, Any]] = []

    for r in rows:
        cum_pv += r.pv
        cum_ev += r.ev
        cum_ac += r.ac

        sv = cum_ev - cum_pv
        cv = cum_ev - cum_ac
        spi = _safe_div(cum_ev, cum_pv)
        cpi = _safe_div(cum_ev, cum_ac)

        out.append(
            {
                "Date": r.date.strftime("%Y-%m-%d"),
                "PV_period": r.pv,
                "EV_period": r.ev,
                "AC_period": r.ac,
                "PV_cum": cum_pv,
                "EV_cum": cum_ev,
                "AC_cum": cum_ac,
                "SV": sv,
                "CV": cv,
                "SPI": spi,
                "CPI": cpi,
            }
        )

    # Determine BAC: if not provided, assume BAC = last cumulative PV (common approximation if baseline = PV time-phased)
    inferred_bac = out[-1]["PV_cum"] if out else 0.0
    bac_final = bac if bac is not None else inferred_bac

    # EAC variants:
    # - EAC_cpi = BAC / CPI
    # - EAC_cpi_spi = AC + (BAC - EV) / (CPI * SPI)
    # - ETC = EAC - AC
    last = out[-1] if out else {}
    cpi_last = float(last.get("CPI", float("nan")))
    spi_last = float(last.get("SPI", float("nan")))
    ac_last = float(last.get("AC_cum", 0.0))
    ev_last = float(last.get("EV_cum", 0.0))

    eac_cpi = _safe_div(bac_final, cpi_last) if not math.isnan(
        cpi_last) else float("nan")
    denom = (cpi_last * spi_last) if (not math.isnan(cpi_last)
                                      and not math.isnan(spi_last)) else float("nan")
    eac_cpi_spi = ac_last + \
        _safe_div((bac_final - ev_last),
                  denom) if not math.isnan(denom) else float("nan")

    etc_cpi = eac_cpi - ac_last if not math.isnan(eac_cpi) else float("nan")
    etc_cpi_spi = eac_cpi_spi - \
        ac_last if not math.isnan(eac_cpi_spi) else float("nan")

    vac_cpi = bac_final - eac_cpi if not math.isnan(eac_cpi) else float("nan")
    vac_cpi_spi = bac_final - \
        eac_cpi_spi if not math.isnan(eac_cpi_spi) else float("nan")

    summary = {
        "BAC": bac_final,
        "BAC_source": "user" if bac is not None else "inferred_from_last_PV_cum",
        "PV_cum": float(last.get("PV_cum", 0.0)),
        "EV_cum": float(last.get("EV_cum", 0.0)),
        "AC_cum": float(last.get("AC_cum", 0.0)),
        "SPI": float(last.get("SPI", float("nan"))),
        "CPI": float(last.get("CPI", float("nan"))),
        "SV": float(last.get("SV", float("nan"))),
        "CV": float(last.get("CV", float("nan"))),
        "EAC_CPI": eac_cpi,
        "ETC_CPI": etc_cpi,
        "VAC_CPI": vac_cpi,
        "EAC_CPI_SPI": eac_cpi_spi,
        "ETC_CPI_SPI": etc_cpi_spi,
        "VAC_CPI_SPI": vac_cpi_spi,
    }

    return out, summary


def write_csv(path: Path, rows: List[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        raise ValueError("No rows to write.")
    fieldnames = list(rows[0].keys())
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def fmt_money(x: float) -> str:
    if x is None or math.isnan(x):
        return "NA"
    return f"{x:,.2f}"


def fmt_ratio(x: float) -> str:
    if x is None or math.isnan(x):
        return "NA"
    return f"{x:.3f}"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="EVM Report (PV/EV/AC) for Project Controls.")
    parser.add_argument("--input", required=True,
                        help="Input CSV path (Date, PV, EV, AC).")
    parser.add_argument(
        "--output", default="data/evm_output.csv", help="Output CSV path.")
    parser.add_argument("--bac", type=float, default=None,
                        help="Budget at Completion. If omitted, inferred from last PV_cum.")
    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)

    rows = read_evm_csv(in_path)
    detail, summary = compute_evm(rows, args.bac)
    write_csv(out_path, detail)

    print("\n=== EVM Executive Summary ===")
    print(f"Input file:  {in_path}")
    print(f"Output file: {out_path}")
    print(f"BAC ({summary['BAC_source']}): {fmt_money(summary['BAC'])}")
    print(
        f"PV_cum: {fmt_money(summary['PV_cum'])} | EV_cum: {fmt_money(summary['EV_cum'])} | AC_cum: {fmt_money(summary['AC_cum'])}")
    print(
        f"SPI: {fmt_ratio(summary['SPI'])} | CPI: {fmt_ratio(summary['CPI'])}")
    print(f"SV: {fmt_money(summary['SV'])} | CV: {fmt_money(summary['CV'])}")
    print("\nForecasts:")
    print(
        f"EAC (CPI):      {fmt_money(summary['EAC_CPI'])} | ETC: {fmt_money(summary['ETC_CPI'])} | VAC: {fmt_money(summary['VAC_CPI'])}")
    print(
        f"EAC (CPI*SPI):  {fmt_money(summary['EAC_CPI_SPI'])} | ETC: {fmt_money(summary['ETC_CPI_SPI'])} | VAC: {fmt_money(summary['VAC_CPI_SPI'])}")
    print("\nDone.\n")


if __name__ == "__main__":
    main()
