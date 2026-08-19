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


# Hardcoded floor for the India 10-year G-Sec yield, used only if every live
# source fails (yfinance IN10Y.NS, investing.com, FBIL REST). Approximate
# yield as of August 2026 — fallback, update manually if yield shifts significantly.
RF_HARDCODED_FALLBACK = 0.0685

INVESTING_COM_URL = "https://www.investing.com/rates-bonds/india-10-year-bond-yield"
INVESTING_COM_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.google.com/",
}

# Undocumented — FBIL's public site is a JS-rendered Angular SPA with no known
# stable public REST path, so this is a best-effort probe of plausible endpoints.
FBIL_REST_CANDIDATES = [
    "https://www.fbil.org.in/rest/OverNightRateHistory",
    "https://www.fbil.org.in/api/RatesData",
]


def _fetch_rate_from_investing_com():
    """Best-effort scrape of investing.com's India 10Y bond yield page."""
    import re

    import requests
    from bs4 import BeautifulSoup

    resp = requests.get(INVESTING_COM_URL, headers=INVESTING_COM_HEADERS, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    tag = soup.find(attrs={"data-test": "instrument-price-last"})
    if tag and tag.text.strip():
        return float(tag.text.strip().replace(",", "")) / 100

    tag = soup.find(id="last_last")
    if tag and tag.text.strip():
        return float(tag.text.strip().replace(",", "")) / 100

    match = re.search(r'"last"\s*:\s*"?(\d{1,2}\.\d{2,4})"?', resp.text)
    if match:
        return float(match.group(1)) / 100

    match = re.search(r"India\s*10[- ]Year.{0,80}?(\d{1,2}\.\d{2,4})\s*%", soup.get_text(" "))
    if match:
        return float(match.group(1)) / 100

    return None


def _fetch_rate_from_fbil_rest():
    """Best-effort probe of undocumented FBIL REST endpoints for the benchmark yield."""
    import requests

    for url in FBIL_REST_CANDIDATES:
        try:
            resp = requests.get(url, headers={"Accept": "application/json"}, timeout=10)
            if resp.status_code != 200:
                continue
            payload = resp.json()
            # Shape is unknown/undocumented — probe common key names defensively.
            candidates = payload if isinstance(payload, list) else [payload]
            for item in candidates:
                if not isinstance(item, dict):
                    continue
                for key in ("rate", "value", "yield", "Rate", "Value"):
                    if key in item:
                        try:
                            val = float(item[key])
                            return val / 100 if val > 1 else val
                        except (TypeError, ValueError):
                            continue
        except Exception:
            continue
    return None


def fetch_risk_free_rate():
    """
    India 10-year G-Sec yield. Source order: yfinance IN10Y.NS, then
    investing.com, then FBIL's (undocumented) REST API, then a hardcoded
    fallback. Always returns a decimal (e.g. 0.0678) — never None — so the
    workbook always gets a value even if every live source is unreachable.
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
        rate = _fetch_rate_from_investing_com()
        if rate:
            log.info("Risk-free rate from investing.com: %.4f", rate)
            return rate
        log.warning("investing.com fetched but no rate pattern found")
    except Exception as e:
        log.warning("investing.com scrape failed (likely blocked by Cloudflare bot protection): %s", e)

    try:
        rate = _fetch_rate_from_fbil_rest()
        if rate:
            log.info("Risk-free rate from FBIL REST: %.4f", rate)
            return rate
        log.warning("FBIL REST probe returned no usable data (endpoints are undocumented/unstable)")
    except Exception as e:
        log.warning("FBIL REST probe failed: %s", e)

    log.warning(
        "All live risk-free rate sources failed — using hardcoded fallback %.2f%% "
        "(update RF_HARDCODED_FALLBACK manually if the yield shifts significantly)",
        RF_HARDCODED_FALLBACK * 100,
    )
    return RF_HARDCODED_FALLBACK


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
    """Load the recalculated workbook with data_only=True and pull every output.

    LibreOffice's headless conversion doesn't always force a full formula
    recalculation, and openpyxl itself drops cached formula results whenever a
    formula-mode workbook is re-saved (write_market_inputs does exactly that).
    Either way, a formula cell in the recalculated working copy can come back
    None even though the pipeline ran cleanly. As a safety net, any cell that
    reads None from the working copy falls back to the same cell's cached
    value in the original, untouched source workbook — Excel pre-computes and
    stores those values in the file itself, so they're non-None even without
    LibreOffice, just not updated with today's live market inputs.
    """
    try:
        wbv = openpyxl.load_workbook(WORKING_XLSX, data_only=True)
    except Exception as e:
        log.error("Could not open working workbook for extraction: %s", e)
        return None

    wbv_fallback = None
    try:
        wbv_fallback = openpyxl.load_workbook(SOURCE_XLSX, data_only=True)
    except Exception as e:
        log.warning("Could not open source workbook for cached-value fallback: %s", e)

    fallback_used = []

    def get_cell(sheet_name, cell):
        val = None
        if sheet_name in wbv.sheetnames:
            val = safe_get(wbv[sheet_name], cell)
        if val is None and wbv_fallback is not None and sheet_name in wbv_fallback.sheetnames:
            fallback_val = safe_get(wbv_fallback[sheet_name], cell)
            if fallback_val is not None:
                val = fallback_val
                fallback_used.append(f"{sheet_name}!{cell}")
        return val

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

    ke = get_cell(*OUT_KE)
    wacc = get_cell(*OUT_WACC)
    terminal_g = get_cell(*OUT_TERMINAL_G)
    out["cost_of_capital"] = {"ke": ke, "wacc": wacc, "terminal_g": terminal_g}

    core_auto_ps = get_cell(*OUT_CORE_AUTO_PS)
    tvs_credit_ps = get_cell(*OUT_TVS_CREDIT_PS)
    sotp_ps = get_cell(*OUT_SOTP_PS)

    shares_out = get_cell(SOTP_SHEET, SOTP_SHARES_OUT)
    core_auto_ev = get_cell(SOTP_SHEET, SOTP_CORE_AUTO_EV)
    net_debt = get_cell(SOTP_SHEET, SOTP_NET_DEBT)
    other_subs_ps = get_cell(SOTP_SHEET, SOTP_OTHER_SUBS_PS)

    core_auto_ev_ps = (core_auto_ev / shares_out) if (core_auto_ev is not None and shares_out) else None
    net_debt_ps = (net_debt / shares_out) if (net_debt is not None and shares_out) else None

    out["sotp_breakdown"] = {
        "core_auto_per_share": core_auto_ps,
        "tvs_credit_per_share": tvs_credit_ps,
        "sotp_per_share": sotp_ps,
        "other_subs_per_share": other_subs_ps,
        "net_debt_per_share": net_debt_ps,
    }

    out["concluded_value_per_share"] = get_cell(*OUT_CONCLUDED_PS)
    out["upside_pct"] = get_cell(*OUT_UPSIDE)
    out["recommendation"] = get_cell(*OUT_RECOMMENDATION)

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
        current_value = get_cell(SCENARIO_SHEET, cell)
        options = dv_options.get(cell) or FALLBACK_OPTIONS.get(name, [])
        scenario_switches[name] = {"current_value": current_value, "valid_options": options}
    out["scenario_switches"] = scenario_switches

    # Sensitivity grid
    wacc_values = [get_cell(SENS_SHEET, c) for c in SENS_WACC_CELLS]
    g_values = [get_cell(SENS_SHEET, c) for c in SENS_G_CELLS]
    matrix = []
    for r in SENS_GRID_ROWS:
        row_vals = [get_cell(SENS_SHEET, f"{col}{r}") for col in SENS_GRID_COLS]
        matrix.append(row_vals)
    out["sensitivity_grid"] = {"wacc_values": wacc_values, "g_values": g_values, "matrix": matrix}

    # Monte Carlo
    mean = get_cell(MC_SHEET, MC_MEAN)
    std_dev = get_cell(MC_SHEET, MC_STD)
    median = get_cell(MC_SHEET, MC_MEDIAN)
    prob_above_cmp = get_cell(MC_SHEET, MC_PROB_ABOVE_CMP)

    trials = []
    try:
        for r in range(MC_TRIAL_FIRST_ROW, MC_TRIAL_LAST_ROW + 1):
            v = get_cell(MC_SHEET, f"{MC_TRIAL_COL}{r}")
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
    football_field = []
    for methodology, row in FOOTBALL_FIELD_ROWS.items():
        low = get_cell(SOTP_SHEET, f"C{row}")
        mid = get_cell(SOTP_SHEET, f"D{row}")
        high = get_cell(SOTP_SHEET, f"E{row}")
        football_field.append({"methodology": methodology, "low": low, "mid": mid, "high": high})
    out["football_field"] = football_field

    # Peer multiples for the snapshot table
    for key, row in PEER_MULTIPLE_ROWS.items():
        ev_ebitda = get_cell("Peer Comps", f"C{row}")
        pe = get_cell("Peer Comps", f"D{row}")
        out["market_inputs"]["peers"][key]["ev_ebitda"] = ev_ebitda
        out["market_inputs"]["peers"][key]["pe"] = pe

    if fallback_used:
        log.warning(
            "%d cell(s) fell back to the original source workbook's cached values "
            "(recalculated working copy returned None): %s",
            len(fallback_used),
            ", ".join(fallback_used[:10]) + (", ..." if len(fallback_used) > 10 else ""),
        )

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
