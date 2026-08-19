"""
Daily refresh pipeline for the TVS Motor equity research dashboard.

Copies the source valuation model, writes today's live market data into the
exact input cells identified in the Phase 1 cell map below, recalculates the
workbook with LibreOffice headless, and extracts every downstream output into
data/outputs.json for the Streamlit dashboard to read.

Every external call (yfinance, web scrape, LibreOffice subprocess) is wrapped
so that one failure never crashes the whole run.
"""

import json
import logging
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger("refresh")

ROOT = Path(__file__).resolve().parent.parent
SOURCE_XLSX = ROOT / "TVS_Motor_Valuation_Model.xlsx"
WORKING_XLSX = ROOT / "model_working.xlsx"
OUTPUTS_JSON = ROOT / "data" / "outputs.json"

# ---------------------------------------------------------------------------
# Cell map — from the Phase 1 openpyxl audit of TVS_Motor_Valuation_Model.xlsx
# ---------------------------------------------------------------------------

# Market inputs written daily
CELL_CMP = ("Debt & Cost of Capital", "C48")          # TVSMOTOR.NS current market price
CELL_RF = ("Guidance & Assumptions", "C20")            # India 10-year G-Sec yield
PEER_CELLS = {
    "bajaj_auto": {"sheet": "Peer Comps", "cell": "H8", "ticker": "BAJAJ-AUTO.NS", "name": "Bajaj Auto"},
    "hero_motocorp": {"sheet": "Peer Comps", "cell": "H9", "ticker": "HEROMOTOCO.NS", "name": "Hero MotoCorp"},
    "eicher_motors": {"sheet": "Peer Comps", "cell": "H10", "ticker": "EICHERMOT.NS", "name": "Eicher Motors"},
}
TVS_TICKER = "TVSMOTOR.NS"
RF_TICKER = "IN10Y.NS"

# Key outputs — all on Control Panel, "B. LIVE OUTPUT" block
OUT_KE = ("Control Panel", "C18")
OUT_WACC = ("Control Panel", "C19")
OUT_TERMINAL_G = ("Control Panel", "C20")
OUT_CORE_AUTO_PS = ("Control Panel", "C21")
OUT_TVS_CREDIT_PS = ("Control Panel", "C22")
OUT_SOTP_PS = ("Control Panel", "C23")
OUT_CONCLUDED_PS = ("Control Panel", "C24")
OUT_CMP = ("Control Panel", "C25")
OUT_UPSIDE = ("Control Panel", "C26")
OUT_RECOMMENDATION = ("Control Panel", "C27")

# Scenario switches — Control Panel!C7:C15, each with a list data validation
SCENARIO_SWITCHES = {
    "Operating case": "C7",
    "Terminal growth case": "C8",
    "Beta basis": "C9",
    "WACC adjustment": "C10",
    "Commodity pass-through": "C11",
    "TVS Credit valuation method": "C12",
    "Peer multiple basis": "C13",
    "SOTP weighting scheme": "C14",
    "Monte Carlo volatility regime": "C15",
}
SCENARIO_SHEET = "Control Panel"
FALLBACK_OPTIONS = {
    "Operating case": ["Conservative", "Base", "Aggressive"],
    "Terminal growth case": ["Conservative", "Base", "Aggressive"],
    "Beta basis": ["Regression beta", "Bottom-up relevered", "Average of both"],
    "WACC adjustment": ["-100 bp", "-50 bp", "As calculated", "+50 bp", "+100 bp"],
    "Commodity pass-through": ["Weak - 40%", "Base - 70%", "Strong - 90%"],
    "TVS Credit valuation method": ["Average of three", "Peer P/B only", "Warranted P/B only", "Transaction only"],
    "Peer multiple basis": ["Peer median", "Peer mean", "Peer low", "Peer high"],
    "SOTP weighting scheme": ["DCF-led", "Balanced", "Market-led"],
    "Monte Carlo volatility regime": ["Low", "Base", "High"],
}

# SOTP waterfall legs — SOTP & Football Field
SOTP_SHEET = "SOTP & Football Field"
SOTP_CORE_AUTO_EV = "C7"          # core auto enterprise value (₹ cr)
SOTP_NET_DEBT = "C8"              # net debt, already negative (₹ cr)
SOTP_SHARES_OUT = "C23"           # shares outstanding (cr)
SOTP_TVS_CREDIT_PS = "D13"
SOTP_OTHER_SUBS_PS = "D15"
SOTP_TOTAL_PS = "D22"

# Football field — SOTP & Football Field!C29:E35 (low, base/mid, high)
FOOTBALL_FIELD_ROWS = {
    "DCF - FCFF": 29,
    "DCF - FCFE": 30,
    "EV/EBITDA": 31,
    "P/E": 32,
    "Precedent Transactions": 33,
    "Asset-Based NAV": 34,
    "DDM": 35,
}

# Sensitivity grid — Sensitivity & Scenarios: WACC rows C27:C31, g cols D26:H26, grid D27:H31
SENS_SHEET = "Sensitivity & Scenarios"
SENS_WACC_CELLS = ["C27", "C28", "C29", "C30", "C31"]
SENS_G_CELLS = ["D26", "E26", "F26", "G26", "H26"]
SENS_GRID_COLS = ["D", "E", "F", "G", "H"]
SENS_GRID_ROWS = [27, 28, 29, 30, 31]

# Monte Carlo
MC_SHEET = "Monte Carlo"
MC_MEAN = "C17"
MC_MEDIAN = "C18"
MC_STD = "C19"
MC_P5 = "C20"
MC_P25 = "C21"
MC_P75 = "C22"
MC_P95 = "C23"
MC_CMP = "C24"
MC_PROB_ABOVE_CMP = "C25"
MC_TRIAL_COL = "J"
MC_TRIAL_FIRST_ROW = 31
MC_TRIAL_LAST_ROW = 1030

# Peer multiples — Peer Comps!C15:D17 (EV/EBITDA, P/E)
PEER_MULTIPLE_ROWS = {"bajaj_auto": 15, "hero_motocorp": 16, "eicher_motors": 17}


def fetch_price(ticker: str):
    """Best-effort last price fetch via yfinance. Returns float or None."""
    try:
        import yfinance as yf

        t = yf.Ticker(ticker)
        try:
            price = t.fast_info.get("lastPrice") or t.fast_info.get("last_price")
            if price:
                return float(price)
        except Exception:
            pass
        info = t.info
        price = info.get("currentPrice") or info.get("regularMarketPrice")
        if price:
            return float(price)
        hist = t.history(period="5d")
        if not hist.empty:
            return float(hist["Close"].iloc[-1])
    except Exception as e:
        log.warning("Price fetch failed for %s: %s", ticker, e)
    return None


def fetch_market_cap(ticker: str):
    """Best-effort market cap fetch via yfinance, in raw INR. Returns float or None."""
    try:
        import yfinance as yf

        t = yf.Ticker(ticker)
        try:
            mc = t.fast_info.get("marketCap") or t.fast_info.get("market_cap")
            if mc:
                return float(mc)
        except Exception:
            pass
        info = t.info
        mc = info.get("marketCap")
        if mc:
            return float(mc)
    except Exception as e:
        log.warning("Market cap fetch failed for %s: %s", ticker, e)
    return None


def fetch_risk_free_rate():
    """
    India 10-year G-Sec yield. Tries yfinance IN10Y.NS first, then scrapes
    FBIL. Returns a decimal (e.g. 0.0678) or None if every source fails.
    """
    try:
        import yfinance as yf

        t = yf.Ticker(RF_TICKER)
        hist = t.history(period="5d")
        if not hist.empty:
            val = float(hist["Close"].iloc[-1])
            # Yahoo yield tickers are usually already in percent terms (e.g. 6.78)
            rate = val / 100 if val > 1 else val
            log.info("Risk-free rate from yfinance %s: %.4f", RF_TICKER, rate)
            return rate
    except Exception as e:
        log.warning("yfinance risk-free fetch failed (%s): %s", RF_TICKER, e)

    try:
        import re

        import requests
        from bs4 import BeautifulSoup

        resp = requests.get("https://www.fbil.org.in/#/home", timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        text = soup.get_text(" ")
        match = re.search(r"(\d{1,2}\.\d{2,4})\s*%", text)
        if match:
            rate = float(match.group(1)) / 100
            log.info("Risk-free rate from FBIL scrape: %.4f", rate)
            return rate
        log.warning("FBIL page fetched but no rate pattern found (likely JS-rendered content)")
    except Exception as e:
        log.warning("FBIL scrape failed: %s", e)

    log.warning("All risk-free rate sources failed — keeping existing workbook value")
    return None


def write_market_inputs(wb, results):
    """Write fetched market data into the workbook. wb is loaded with data_only=False."""
    sheet, cell = CELL_CMP
    if results["cmp"] is not None:
        wb[sheet][cell] = results["cmp"]
        log.info("Wrote CMP %.2f -> %s!%s", results["cmp"], sheet, cell)
    else:
        log.warning("CMP unavailable — leaving %s!%s unchanged", sheet, cell)

    sheet, cell = CELL_RF
    if results["rf_rate"] is not None:
        wb[sheet][cell] = results["rf_rate"]
        log.info("Wrote Rf %.4f -> %s!%s", results["rf_rate"], sheet, cell)
    else:
        log.warning("Rf rate unavailable — leaving %s!%s unchanged", sheet, cell)

    for key, meta in PEER_CELLS.items():
        mc_cr = results["peers"][key]["market_cap_cr"]
        if mc_cr is not None:
            wb[meta["sheet"]][meta["cell"]] = mc_cr
            log.info("Wrote %s market cap %.1f cr -> %s!%s", meta["name"], mc_cr, meta["sheet"], meta["cell"])
        else:
            log.warning("%s market cap unavailable — leaving %s!%s unchanged", meta["name"], meta["sheet"], meta["cell"])


def recalculate_with_libreoffice():
    """Try to recalculate WORKING_XLSX in place via LibreOffice headless."""
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        log.warning("LibreOffice (soffice) not found on PATH — skipping recalculation, will read cached values")
        return False
    try:
        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--calc",
                "--infilter=Calc MS Excel 2007 XML",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(WORKING_XLSX.parent),
                str(WORKING_XLSX),
            ],
            capture_output=True,
            text=True,
            timeout=180,
        )
        if result.returncode != 0:
            log.warning("LibreOffice recalculation returned code %s: %s", result.returncode, result.stderr[:500])
            return False
        log.info("Recalculated workbook with LibreOffice headless")
        return True
    except Exception as e:
        log.warning("LibreOffice recalculation failed: %s", e)
        return False


def safe_get(ws, cell, default=None):
    try:
        val = ws[cell].value
        return val if val is not None else default
    except Exception:
        return default


def extract_outputs(results):
    """Load the recalculated workbook with data_only=True and pull every output."""
    try:
        wbv = openpyxl.load_workbook(WORKING_XLSX, data_only=True)
    except Exception as e:
        log.error("Could not open working workbook for extraction: %s", e)
        return None

    out = {}
    out["valuation_date"] = datetime.now().strftime("%Y-%m-%d")
    out["last_refreshed"] = datetime.now(timezone.utc).isoformat()

    out["market_inputs"] = {
        "cmp": results.get("cmp"),
        "rf_rate": results.get("rf_rate"),
        "peers": {
            key: {
                "ticker": meta["ticker"],
                "name": meta["name"],
                "price": results["peers"][key].get("price"),
                "currency": "INR",
            }
            for key, meta in PEER_CELLS.items()
        },
    }

    ke = safe_get(wbv[OUT_KE[0]], OUT_KE[1])
    wacc = safe_get(wbv[OUT_WACC[0]], OUT_WACC[1])
    terminal_g = safe_get(wbv[OUT_TERMINAL_G[0]], OUT_TERMINAL_G[1])
    out["cost_of_capital"] = {"ke": ke, "wacc": wacc, "terminal_g": terminal_g}

    core_auto_ps = safe_get(wbv[OUT_CORE_AUTO_PS[0]], OUT_CORE_AUTO_PS[1])
    tvs_credit_ps = safe_get(wbv[OUT_TVS_CREDIT_PS[0]], OUT_TVS_CREDIT_PS[1])
    sotp_ps = safe_get(wbv[OUT_SOTP_PS[0]], OUT_SOTP_PS[1])

    sotp_ws = wbv[SOTP_SHEET]
    shares_out = safe_get(sotp_ws, SOTP_SHARES_OUT)
    core_auto_ev = safe_get(sotp_ws, SOTP_CORE_AUTO_EV)
    net_debt = safe_get(sotp_ws, SOTP_NET_DEBT)
    other_subs_ps = safe_get(sotp_ws, SOTP_OTHER_SUBS_PS)

    core_auto_ev_ps = (core_auto_ev / shares_out) if (core_auto_ev is not None and shares_out) else None
    net_debt_ps = (net_debt / shares_out) if (net_debt is not None and shares_out) else None

    out["sotp_breakdown"] = {
        "core_auto_per_share": core_auto_ps,
        "tvs_credit_per_share": tvs_credit_ps,
        "sotp_per_share": sotp_ps,
        "other_subs_per_share": other_subs_ps,
        "net_debt_per_share": net_debt_ps,
    }

    out["concluded_value_per_share"] = safe_get(wbv[OUT_CONCLUDED_PS[0]], OUT_CONCLUDED_PS[1])
    out["upside_pct"] = safe_get(wbv[OUT_UPSIDE[0]], OUT_UPSIDE[1])
    out["recommendation"] = safe_get(wbv[OUT_RECOMMENDATION[0]], OUT_RECOMMENDATION[1])

    # Scenario switches with their valid dropdown options
    wbf = None
    try:
        wbf = openpyxl.load_workbook(WORKING_XLSX, data_only=False)
    except Exception:
        pass

    scenario_switches = {}
    dv_options = {}
    if wbf is not None:
        try:
            ws_formula = wbf[SCENARIO_SHEET]
            for dv in ws_formula.data_validations.dataValidation:
                if dv.type == "list" and dv.formula1:
                    opts = dv.formula1.strip('"').split(",")
                    for coord in str(dv.sqref).split():
                        dv_options[coord] = opts
        except Exception as e:
            log.warning("Could not read data validations: %s", e)

    for name, cell in SCENARIO_SWITCHES.items():
        current_value = safe_get(wbv[SCENARIO_SHEET], cell)
        options = dv_options.get(cell) or FALLBACK_OPTIONS.get(name, [])
        scenario_switches[name] = {"current_value": current_value, "valid_options": options}
    out["scenario_switches"] = scenario_switches

    # Sensitivity grid
    sens_ws = wbv[SENS_SHEET]
    wacc_values = [safe_get(sens_ws, c) for c in SENS_WACC_CELLS]
    g_values = [safe_get(sens_ws, c) for c in SENS_G_CELLS]
    matrix = []
    for r in SENS_GRID_ROWS:
        row_vals = [safe_get(sens_ws, f"{col}{r}") for col in SENS_GRID_COLS]
        matrix.append(row_vals)
    out["sensitivity_grid"] = {"wacc_values": wacc_values, "g_values": g_values, "matrix": matrix}

    # Monte Carlo
    mc_ws = wbv[MC_SHEET]
    mean = safe_get(mc_ws, MC_MEAN)
    std_dev = safe_get(mc_ws, MC_STD)
    median = safe_get(mc_ws, MC_MEDIAN)
    prob_above_cmp = safe_get(mc_ws, MC_PROB_ABOVE_CMP)

    trials = []
    try:
        for r in range(MC_TRIAL_FIRST_ROW, MC_TRIAL_LAST_ROW + 1):
            v = safe_get(mc_ws, f"{MC_TRIAL_COL}{r}")
            if v is not None:
                trials.append(v)
    except Exception as e:
        log.warning("Could not read Monte Carlo trial data: %s", e)

    if trials:
        try:
            import numpy as np

            p10 = float(np.percentile(trials, 10))
            p25 = float(np.percentile(trials, 25))
            p50 = float(np.percentile(trials, 50))
            p75 = float(np.percentile(trials, 75))
            p90 = float(np.percentile(trials, 90))
        except Exception:
            p10 = p25 = p50 = p75 = p90 = None
    else:
        p10 = None
        p25 = safe_get(mc_ws, MC_P25)
        p50 = median
        p75 = safe_get(mc_ws, MC_P75)
        p90 = None

    out["monte_carlo"] = {
        "p10": p10,
        "p25": p25,
        "p50": p50 if p50 is not None else median,
        "p75": p75,
        "p90": p90,
        "mean": mean,
        "std_dev": std_dev,
        "prob_above_cmp": prob_above_cmp,
        "raw_trials": trials if trials else None,
    }

    # SOTP legs — ordered waterfall
    out["sotp_legs"] = [
        {"name": "Core Automotive (EV)", "value_per_share": core_auto_ev_ps, "type": "add"},
        {"name": "Net Debt", "value_per_share": net_debt_ps, "type": "subtract"},
        {"name": "TVS Credit (85.15%)", "value_per_share": tvs_credit_ps, "type": "add"},
        {"name": "Other Subsidiaries", "value_per_share": other_subs_ps, "type": "add"},
    ]

    # Football field
    ff_ws = wbv[SOTP_SHEET]
    football_field = []
    for methodology, row in FOOTBALL_FIELD_ROWS.items():
        low = safe_get(ff_ws, f"C{row}")
        mid = safe_get(ff_ws, f"D{row}")
        high = safe_get(ff_ws, f"E{row}")
        football_field.append({"methodology": methodology, "low": low, "mid": mid, "high": high})
    out["football_field"] = football_field

    # Peer multiples for the snapshot table
    peer_ws = wbv["Peer Comps"]
    for key, row in PEER_MULTIPLE_ROWS.items():
        ev_ebitda = safe_get(peer_ws, f"C{row}")
        pe = safe_get(peer_ws, f"D{row}")
        out["market_inputs"]["peers"][key]["ev_ebitda"] = ev_ebitda
        out["market_inputs"]["peers"][key]["pe"] = pe

    return out


def main():
    log.info("=== TVS Motor dashboard daily refresh starting ===")

    if not SOURCE_XLSX.exists():
        log.error("Source workbook not found at %s — cannot proceed", SOURCE_XLSX)
        sys.exit(1)

    shutil.copy(SOURCE_XLSX, WORKING_XLSX)
    log.info("Copied %s -> %s", SOURCE_XLSX.name, WORKING_XLSX.name)

    results = {"cmp": None, "rf_rate": None, "peers": {}}
    fetch_status = {}

    results["cmp"] = fetch_price(TVS_TICKER)
    fetch_status["TVSMOTOR CMP"] = results["cmp"] is not None

    for key, meta in PEER_CELLS.items():
        price = fetch_price(meta["ticker"])
        mc = fetch_market_cap(meta["ticker"])
        mc_cr = (mc / 1e7) if mc is not None else None
        results["peers"][key] = {"price": price, "market_cap_cr": mc_cr}
        fetch_status[f"{meta['name']} price/mcap"] = price is not None and mc_cr is not None

    results["rf_rate"] = fetch_risk_free_rate()
    fetch_status["India 10Y G-Sec (Rf)"] = results["rf_rate"] is not None

    try:
        wb = openpyxl.load_workbook(WORKING_XLSX, keep_vba=False, data_only=False)
        write_market_inputs(wb, results)
        wb.save(WORKING_XLSX)
        log.info("Saved updated working workbook")
    except Exception as e:
        log.error("Failed to write market inputs into workbook: %s", e)

    recalculated = recalculate_with_libreoffice()
    if not recalculated:
        log.warning("Proceeding with cached formula values from the last time the workbook was opened in Excel")

    outputs = extract_outputs(results)
    if outputs is None:
        log.error("Extraction failed — writing minimal fallback outputs.json")
        outputs = {
            "valuation_date": datetime.now().strftime("%Y-%m-%d"),
            "last_refreshed": datetime.now(timezone.utc).isoformat(),
            "market_inputs": results,
            "error": "extraction_failed",
        }

    OUTPUTS_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUTS_JSON, "w", encoding="utf-8") as f:
        json.dump(outputs, f, indent=2, default=str)
    log.info("Wrote %s", OUTPUTS_JSON)

    log.info("=== RUN SUMMARY ===")
    for label, ok in fetch_status.items():
        log.info("  %-30s %s", label, "OK" if ok else "FAILED (kept existing value)")
    log.info("  LibreOffice recalculation:     %s", "OK" if recalculated else "SKIPPED (used cached values)")
    concluded = outputs.get("concluded_value_per_share")
    rec = outputs.get("recommendation")
    cmp_val = outputs.get("market_inputs", {}).get("cmp")
    log.info("  Concluded value/share (Rs):    %s", concluded)
    log.info("  Current market price (Rs):     %s", cmp_val)
    log.info("  Recommendation:                %s", rec)
    log.info("=== Refresh complete ===")


if __name__ == "__main__":
    main()
