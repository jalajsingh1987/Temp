#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EM Asia swap-rate PCA framework  (offshore swap curves only - no bond yields, no onshore curves)

Builds two linked Excel workbooks from a Bloomberg export of par swap rates:
    EM_Asia_Swaps_PCA_Results.xlsx   (results, static PCA outputs + live z-score / percentile / rank formulas)
    EM_Asia_Swaps_PCA_Charts.xlsx    (native Excel charts on a Chart_Data sheet that links to the Results file)

USAGE
    python build_all.py --data swap_rates.csv     # real Bloomberg export: Date, Market, Tenor, Rate [, Source, Unit]
    python build_all.py --demo                    # synthetic data, every sheet/chart labelled ILLUSTRATIVE ONLY
    python build_all.py --templates               # only (re)write the CSV templates

Nothing here calls Bloomberg. You export the data (BDH / Excel API / Terminal export) and fill Ticker_Map yourself.
Edit the CONFIG block below (markets, tenors, dates, parameters). Fixed seeds, no hidden state.
"""
import os, sys, math, re, zipfile, shutil, argparse, warnings, itertools, datetime as dt, subprocess, tempfile, textwrap
from collections import OrderedDict
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# =====================================================================================
# 0. CONFIG BLOCK  (everything tunable lives here and is echoed on the Assumptions sheet)
# =====================================================================================
HERE = os.path.dirname(os.path.abspath(__file__))
RESULTS_NAME = "EM_Asia_Swaps_PCA_Results.xlsx"
CHARTS_NAME = "EM_Asia_Swaps_PCA_Charts.xlsx"

PARAMS = OrderedDict([
    ("SEED", (42, "Random seed (synthetic demo generator; nothing else in the pipeline is random)")),
    ("LAST_N_YEARS", (5, "Default analysis window in years; a full-history run is added to PCA_Variance")),
    ("SAMPLE_START", ("", "Optional explicit start date YYYY-MM-DD (blank = end date minus LAST_N_YEARS)")),
    ("SAMPLE_END", ("", "Optional explicit end date YYYY-MM-DD (blank = last date in data)")),
    ("FFILL_MAX_DAYS", (2, "Stale/missing fixings forward-filled for at most this many consecutive days (flagged in Data_QC)")),
    ("BADPRINT_SD", (6.0, "Bad print = |daily change| above this many robust SDs (1.4826*MAD) AND reversed next day")),
    ("BADPRINT_REVERT", (0.30, "Reversal test: |d(t)+d(t+1)| < this x |d(t)| with opposite sign")),
    ("ZERO_CHANGE_DROP", (0.30, "Tenor dropped from a curve if share of zero-change days exceeds this")),
    ("ZERO_CHANGE_FLAG", (0.15, "Tenor flagged 'illiquid' (kept) if share of zero-change days exceeds this")),
    ("MIN_COVERAGE", (0.50, "Tenor dropped if it has quotes on less than this share of the curve's calendar")),
    ("MIN_TENORS", (4, "Curve skipped if fewer tenors than this survive cleaning")),
    ("ROLL_WIN", (252, "Rolling PCA window (trading days, ~1 year)")),
    ("ROLL_STEP", (21, "Rolling PCA step (trading days, ~1 month)")),
    ("CONGR_BREAK", (0.90, "Rolling PC congruence coefficient below this flags a stability break")),
    ("Z_WIN", (252, "Trailing window for z-scores / percentiles (trading days)")),
    ("Z_OUTLIER", (1.5, "RV outlier threshold |z|")),
    ("HL_MAX", (60, "Drop RV signals with AR(1) half-life above this many trading days")),
    ("HL_WIN", (504, "Window for AR(1) half-life estimation (trading days)")),
    ("SPILL_WINS", ("126,252", "Rolling spillover windows (trading days): 6 months, 12 months")),
    ("DECOUPLE_LO", (0.5, "Decoupled if rolling 6M R2 < this x full-sample R2")),
    ("DECOUPLE_HI", (1.5, "Re-coupled/high if rolling 6M R2 > this x full-sample R2")),
    ("FWD_JUMP_Z", (6.0, "Forward point flagged if |daily change / robust SD| exceeds this while at most 2 points exceed 40% of it (isolated jump, not a broad move)")),
    ("ROLL_BREAK_SD", (6.0, "Roll/ticker break: persistent level shift of one tenor vs its neighbours (butterfly residual) above this many robust SDs")),
    ("MIN_OBS_PCA", (100, "Minimum observations for any PCA/regression")),
    ("POOL_MIN_MARKETS", (4, "Minimum markets for a pooled PCA")),
    ("PAIR_TOP_N", (10, "Rows shown in the top cross-market pair tables")),
    ("REPRICE_TOL_BP", (0.1, "Bootstrapping re-pricing tolerance (bp)")),
    ("RV_INCLUDE_BLOCKS", (True, "Include labelled block forwards (beyond the last supported 1Y forward) in the RV PCA panel")),
    ("BLOCK_PAIRS", ("1-3;2-5;3-7;5-10;2-10", "Extra block forwards p-q (from quoted pars p and q), e.g. 2-5 = 2y3y, 5-10 = 5y5y; used only if both pillars quoted")),
    ("PCT_WINS", ("252,756,1260", "Windows (trading days) for per-tenor level percentiles: 1y, 3y, 5y")),
    ("MONTH_DAYS", (21, "Trading days per month used for 1M roll-down")),
    ("ROLLDOWN_NOTE", ("linear slope between adjacent supported points", "Roll-down definition (1M, unchanged curve)")),
])
def P(name):
    return PARAMS[name][0]

# Curve conventions.  Only what the brief states is filled in; everything else is 'TO CONFIRM' (not guessed).
# dc_factor = accrual per annual period used in the bootstrap (ACT/360 -> 365/360, ACT/365F -> 1.0). Assumed - confirm.
def _c(market, label, index, settle, fixing, dc, dcf, alt, is_global=False):
    return dict(market=market, label=label, floating_index=index, settlement=settle, fixing_source=fixing,
                day_count=dc, dc_factor=dcf, alternative=alt, is_global=is_global)
CURVES = OrderedDict([
    ("INR_NDIRS", _c("INR", "INR NDIRS (offshore, MIBOR-linked)", "MIBOR (per brief)", "USD-settled (NDIRS)", "TO CONFIRM (Bloomberg source)", "ACT/365F (assumed - confirm)", 1.0,
                     "None listed in brief; confirm whether another offshore INR curve is quoted")),
    ("KRW_NDIRS", _c("KRW", "KRW NDIRS (offshore, CD-linked)", "KRW CD 91D (chosen as most liquid; brief: CD or KOFR)", "USD-settled (NDIRS)", "TO CONFIRM", "ACT/365F (assumed - confirm)", 1.0,
                     "KOFR-linked offshore KRW IRS")),
    ("CNH_NDIRS", _c("CNH", "CNY FR007 NDIRS (offshore, USD-settled)", "FR007 (per brief)", "USD-settled (NDIRS)", "TO CONFIRM", "ACT/365F (assumed - confirm)", 1.0,
                     "Other CNH offshore swap curve, if quoted (brief: 'CNH offshore swaps')")),
    ("TWD_NDIRS", _c("TWD", "TWD NDIRS (offshore)", "TO CONFIRM (not stated in brief)", "USD-settled (NDIRS)", "TO CONFIRM", "ACT/365F (assumed - confirm)", 1.0, "None listed; confirm")),
    ("IDR_NDIRS", _c("IDR", "IDR NDIRS (offshore)", "TO CONFIRM (not stated in brief)", "USD-settled (NDIRS)", "TO CONFIRM", "ACT/360 (assumed - confirm)", 365.0 / 360.0, "None listed; confirm")),
    ("PHP_NDIRS", _c("PHP", "PHP NDIRS (offshore)", "TO CONFIRM (not stated in brief)", "USD-settled (NDIRS)", "TO CONFIRM", "ACT/360 (assumed - confirm)", 365.0 / 360.0, "None listed; confirm")),
    ("MYR_NDIRS", _c("MYR", "MYR NDIRS (offshore, KLIBOR-linked)", "KLIBOR (per brief)", "USD-settled (NDIRS)", "TO CONFIRM", "ACT/365F (assumed - confirm)", 1.0, "None listed; confirm")),
    ("THB_IRS", _c("THB", "THB offshore IRS", "TO CONFIRM (not stated in brief)", "TO CONFIRM (deliverable vs USD-settled)", "TO CONFIRM", "ACT/365F (assumed - confirm)", 1.0, "None listed; confirm")),
    ("SGD_SORA", _c("SGD", "SGD SORA OIS", "SORA (per brief)", "SGD (OIS)", "TO CONFIRM", "ACT/365F (assumed - confirm)", 1.0, "None listed; confirm")),
    ("SOFR", _c("USD", "USD SOFR OIS (global benchmark)", "SOFR (per brief)", "USD", "TO CONFIRM", "ACT/360 (assumed - confirm)", 365.0 / 360.0, "n/a", True)),
])
EM_CURVES = [c for c, m in CURVES.items() if not m["is_global"]]
GLOBAL_CURVE = "SOFR"
BOOTSTRAP_METHOD = ("Log-linear interpolation of discount factors between quoted par pillars (piecewise-constant continuous forwards); "
                    "annual fixed-leg payments, accrual per period = dc_factor of the curve (assumed day count); "
                    "each pillar's discount factor solved so the quoted par rate re-prices exactly; no extrapolation beyond the longest quoted tenor; "
                    "same method for every curve. 1Y forward k = (D_k/D_k+1 - 1)/accrual; block forward p->q = ((D_p/D_q)^(1/(q-p)) - 1)/accrual.")

# Per-curve settings used ONLY by the synthetic demo generator.
SYN = OrderedDict([
    ("INR_NDIRS", dict(tenors=[1, 2, 3, 4, 5, 7, 10], base=640, vol=(3.0, 1.2, 0.6), beta=0.25, reg=0.3, unit="percent")),
    ("KRW_NDIRS", dict(tenors=[1, 2, 3, 4, 5, 7, 10], base=270, vol=(3.5, 1.5, 0.7), beta=0.55, reg=1.0, unit="percent")),
    ("CNH_NDIRS", dict(tenors=[1, 2, 3, 4, 5, 7, 10], base=240, vol=(2.5, 1.2, 0.6), beta=0.10, reg=0.4, unit="percent")),
    ("TWD_NDIRS", dict(tenors=[1, 2, 3, 4, 5, 7, 10], base=160, vol=(1.5, 0.6, 0.3), beta=0.45, reg=0.8, unit="percent")),
    ("IDR_NDIRS", dict(tenors=[1, 2, 3, 5, 7, 10], base=650, vol=(4.5, 2.0, 0.9), beta=0.40, reg=0.6, unit="percent")),
    ("PHP_NDIRS", dict(tenors=[1, 2, 3, 5, 7, 10], base=560, vol=(4.0, 1.8, 0.8), beta=0.35, reg=0.5, unit="percent")),
    ("MYR_NDIRS", dict(tenors=[1, 2, 3, 4, 5, 7, 10], base=360, vol=(2.0, 0.9, 0.4), beta=0.45, reg=0.7, unit="percent")),
    ("THB_IRS", dict(tenors=[1, 2, 3, 5, 7, 10], base=210, vol=(3.0, 1.4, 0.6), beta=0.50, reg=0.9, unit="bp")),
    ("SGD_SORA", dict(tenors=[1, 2, 3, 4, 5, 7, 10], base=270, vol=(5.0, 2.0, 0.8), beta=0.85, reg=0.9, unit="percent")),
    ("SOFR", dict(tenors=[1, 2, 3, 4, 5, 7, 10], base=400, vol=(0, 0, 0), beta=0, reg=0, unit="percent")),
])
SYN_STALE = {("TWD_NDIRS", 10): (0.25, 0.70), ("IDR_NDIRS", 10): (0.10, 0.60), ("PHP_NDIRS", 10): (0.10, 0.60),
             ("THB_IRS", 10): (0.08, 0.55), ("MYR_NDIRS", 10): (0.06, 0.50), ("IDR_NDIRS", 7): (0.05, 0.45)}
SYNTH_LABEL = "SYNTHETIC DATA - ILLUSTRATIVE ONLY"

# =====================================================================================
# 1. TENOR / UNIT HELPERS AND DATA LOADING
# =====================================================================================
def parse_tenor(s):
    s = str(s).strip().upper().replace(" ", "")
    m = re.fullmatch(r"(\d+(?:\.\d+)?)([YMW]?)", s)
    if not m:
        return np.nan
    v, u = float(m.group(1)), (m.group(2) or "Y")
    return v if u == "Y" else v / 12.0 if u == "M" else v / 52.0

def tenor_label(t):
    t = float(t)
    if abs(t - round(t)) < 1e-9:
        return f"{int(round(t))}Y"
    m = t * 12.0
    return f"{int(round(m))}M" if abs(m - round(m)) < 1e-6 else f"{t:g}Y"

def is_int_year(t):
    return abs(t - round(t)) < 1e-9 and t >= 1

def robust_sd(a):
    a = np.asarray(a, float)
    a = a[np.isfinite(a)]
    if len(a) < 10:
        return np.nan
    sd = 1.4826 * np.median(np.abs(a - np.median(a)))
    if not np.isfinite(sd) or sd <= 0:
        sd = np.std(a)
    return sd

def make_synthetic():
    """Synthetic long-format data with a global/regional/local factor structure and deliberate data problems
    (stale tenors, holiday fills, weekend fills, bad prints, a ticker-roll break, a pricing-source change, mixed units)."""
    rng = np.random.default_rng(int(P("SEED")))
    dates = pd.bdate_range("2019-01-01", "2026-09-28")
    T = len(dates)
    tdraw = lambda size, df=5: rng.standard_t(df, size=size) * math.sqrt((df - 2) / df)
    ou = np.zeros(T)
    for t in range(1, T):
        ou[t] = 0.995 * ou[t - 1] + 0.1 * rng.standard_normal()
    volm = np.exp(0.35 * ou / ou.std())
    def shp(t):
        t = np.asarray(t, float)
        return np.column_stack([1.0 - 0.035 * (t - 1), (t - 5.0) / 5.0, ((t - 5.0) / 5.0) ** 2 - 0.45])
    G = tdraw((T, 3)) * np.array([4.5, 1.6, 1.5]) * volm[:, None]         # global (SOFR) factors, bp
    Rg = tdraw((T, 2)) * np.array([2.0, 0.8]) * volm[:, None]              # regional EM Asia factors
    kappa = 0.002
    long_rows = []
    hol_all = {}
    bad_targets = []
    for ci, (cid, sp) in enumerate(SYN.items()):
        ten = np.array(sp["tenors"], float)
        S = shp(ten)
        if cid == GLOBAL_CURVE:
            f = G.copy()
        else:
            loc = tdraw((T, 3)) * np.array(sp["vol"]) * volm[:, None]
            Gl = np.vstack([np.zeros((1, 3)), G[:-1]])                       # previous US business day
            gm = sp["beta"] * (0.85 * Gl + 0.15 * G) * np.array([1.0, 0.7, 0.4])
            rr = np.column_stack([sp["reg"] * Rg, np.zeros(T)])
            f = loc + gm + rr
        cum = np.zeros((T, 3))
        for t in range(1, T):
            cum[t] = (1 - kappa) * cum[t - 1] + f[t]
        phi = math.exp(-math.log(2) / 15.0)
        su = 0.4 if cid == GLOBAL_CURVE else 1.0
        u = np.zeros((T, len(ten)))
        for t in range(1, T):
            u[t] = phi * u[t - 1] + su * math.sqrt(1 - phi ** 2) * rng.standard_normal(len(ten))
        base = sp["base"] + 9.0 * (ten - 1) ** 0.7
        Y = base[None, :] + cum @ S.T + u
        if cid == "PHP_NDIRS":                                               # ticker-roll break on 7Y
            j = list(sp["tenors"]).index(7)
            Y[dates >= pd.Timestamp("2023-06-15"), j] += 22.0
        Q = np.zeros_like(Y)
        for j, tn in enumerate(sp["tenors"]):
            step = 0.5 if (cid, tn) in SYN_STALE else 0.1
            q = np.round(Y[:, j] / step) * step
            pe, ps = SYN_STALE.get((cid, tn), (0.01, 0.30))
            stale = False
            for t in range(1, T):
                stale = (rng.random() < ps) if stale else (rng.random() < pe)
                q[t] = q[t - 1] if stale else q[t]
            Q[:, j] = q
        nh = int(T * 10 / 260)
        hol = np.sort(rng.choice(np.arange(5, T), size=nh, replace=False))
        for h in hol:
            Q[h] = Q[h - 1]
        absent = set(h for h in hol if rng.random() < 0.5)
        keep = np.array([t not in absent for t in range(T)])
        frame = pd.DataFrame(Q, index=dates, columns=sp["tenors"])
        frame = frame.mask(rng.random(frame.shape) < 0.002)                     # random single missing quotes
        frame = frame[keep]
        if cid == "KRW_NDIRS":                                              # export includes weekends with previous value
            wk = pd.date_range(frame.index[0], frame.index[-1], freq="D")
            wk = wk[wk.dayofweek >= 5]
            frame = pd.concat([frame, pd.DataFrame(np.nan, index=wk, columns=frame.columns)]).sort_index().ffill()
        if cid != GLOBAL_CURVE:
            for _ in range(3):
                bad_targets.append((cid, int(rng.choice(sp["tenors"])), frame.index[int(rng.integers(300, len(frame) - 5))], float(rng.choice([-1, 1]) * rng.uniform(40, 90))))
            for _ in range(2):
                bad_targets.append((cid, int(rng.choice(sp["tenors"][:4])), frame.index[int(rng.integers(300, len(frame) - 5))], float(rng.choice([-1, 1]) * rng.uniform(12, 17))))
        for (c2, tn, d0, jump) in [b for b in bad_targets if b[0] == cid]:
            if d0 in frame.index and tn in frame.columns:
                frame.loc[d0, tn] = frame.loc[d0, tn] + jump
        unit = sp["unit"]
        src = pd.Series("BGN", index=frame.index)
        if cid == "MYR_NDIRS":
            src[src.index >= pd.Timestamp("2024-03-01")] = "CMPN"
        for tn in sp["tenors"]:
            col = frame[tn]
            val = col / 100.0 if unit == "percent" else col
            d = pd.DataFrame({"Date": frame.index, "Market": cid, "Tenor": f"{tn}Y", "Rate": val.values, "Source": src.values, "Unit": unit})
            long_rows.append(d.dropna(subset=["Rate"]))
    df = pd.concat(long_rows, ignore_index=True)
    df["Rate"] = df["Rate"].round(4)
    return df

def normalise_columns(df):
    cols = {c.strip().lower(): c for c in df.columns}
    def pick(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None
    ren = {}
    for target, cands in [("Date", ("date", "dates", "asof")), ("Market", ("market", "curve", "curve_id")),
                          ("Tenor", ("tenor", "term")), ("Rate", ("rate", "px_last", "value", "px_mid", "last_price")),
                          ("Source", ("source", "pricing_source", "bpk")), ("Unit", ("unit", "units"))]:
        c = pick(*cands)
        if c is not None:
            ren[c] = target
    out = df.rename(columns=ren)
    missing = [c for c in ("Date", "Market", "Tenor", "Rate") if c not in out.columns]
    if missing:
        raise ValueError(f"Input file must be long format with columns Date, Market, Tenor, Rate (optional Source, Unit). Missing: {missing}. "
                         f"Found columns: {list(df.columns)}. See swap_rates_template.csv.")
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out["Rate"] = pd.to_numeric(out["Rate"], errors="coerce")
    out["Market"] = out["Market"].astype(str).str.strip()
    return out

def inspect_data(df):
    info = OrderedDict()
    info["rows"] = len(df)
    info["date_min"] = df["Date"].min()
    info["date_max"] = df["Date"].max()
    info["bad_dates"] = int(df["Date"].isna().sum())
    info["bad_rates"] = int(df["Rate"].isna().sum())
    info["columns"] = ", ".join(df.columns)
    rows = []
    for m, g in df.groupby("Market"):
        tn = sorted(set(parse_tenor(t) for t in g["Tenor"].unique()))
        rows.append(dict(Market=m, Rows=len(g), FirstDate=g["Date"].min(), LastDate=g["Date"].max(),
                         Tenors=", ".join(tenor_label(t) for t in tn if np.isfinite(t)),
                         KnownCurve="yes" if m in CURVES else "NO - not in CURVES config (ignored)"))
    info["by_market"] = pd.DataFrame(rows)
    info["duplicates"] = int(df.duplicated(["Date", "Market", "Tenor"]).sum())
    return info

def load_data(path=None, demo=False):
    """Return (long_df, inspection).  long_df columns: Date, Market, Tenor, Rate [, Source, Unit].
    Bloomberg export must be long format (see swap_rates_template.csv).  demo=True -> synthetic data."""
    if demo:
        df = make_synthetic()
        src = "SYNTHETIC (generated in-script, seed %s)" % P("SEED")
    else:
        if not path or not os.path.exists(path):
            raise FileNotFoundError(f"No swap-rate file found at {path!r}. Export par swap rates from Bloomberg (BDH / Excel API / Terminal) "
                                    f"in the layout of swap_rates_template.csv, or run with --demo.")
        ext = os.path.splitext(path)[1].lower()
        raw = pd.read_csv(path) if ext == ".csv" else pd.read_excel(path)
        df = normalise_columns(raw)
        src = os.path.abspath(path)
    df = df.dropna(subset=["Date", "Rate"]).copy()
    insp = inspect_data(df)
    insp["source"] = src
    return df, insp

# =====================================================================================
# 1b. TEMPLATES
# =====================================================================================
def write_templates(out_dir):
    pd.DataFrame({"Date": ["2026-01-02", "2026-01-02", "2026-01-05"], "Market": ["<curve id e.g. INR_NDIRS>"] * 3,
                  "Tenor": ["1Y", "2Y", "1Y"], "Rate": ["<value>"] * 3, "Source": ["<pricing source e.g. BGN>"] * 3,
                  "Unit": ["<percent|bp|decimal>"] * 3}).to_csv(os.path.join(out_dir, "swap_rates_template.csv"), index=False)
    rows = []
    for cid in CURVES:
        for tn in ["1Y", "2Y", "3Y", "4Y", "5Y", "7Y", "10Y"]:
            rows.append(dict(Market=CURVES[cid]["market"], Curve=cid, **{"Tenor / forward point": tn}, **{"Bloomberg ticker": ""},
                             Field="", **{"Pricing source": ""}, **{"Quote convention": ""}))
    pd.DataFrame(rows).to_csv(os.path.join(out_dir, "ticker_map_template.csv"), index=False)
    pd.DataFrame({"Date": ["<YYYY-MM-DD>"], "Market": ["<curve id>"], "Event": ["<e.g. RBI MPC decision>"]}).to_csv(os.path.join(out_dir, "events_template.csv"), index=False)
    pd.DataFrame({"Date": ["<YYYY-MM-DD>"], "Market": ["<curve id>"], "Rate": ["<USD/Asia spot>"]}).to_csv(os.path.join(out_dir, "fx_template.csv"), index=False)
    pd.DataFrame({"Date": ["<YYYY-MM-DD>"], "Market": ["<curve id>"], "Point": ["<e.g. 1y1y>"], "Rate": ["<Bloomberg fwd quote, same unit convention>"]}).to_csv(os.path.join(out_dir, "bbg_forwards_template.csv"), index=False)

# =====================================================================================
# 2. CLEANING / BLOOMBERG-SPECIFIC QC
# =====================================================================================
def detect_unit(values):
    v = np.abs(pd.Series(values).dropna().values)
    if len(v) == 0:
        return "percent", "no data"
    med = float(np.median(v))
    if med < 0.4:
        return "decimal", f"median |rate|={med:.4g} -> decimal"
    if med < 40:
        return "percent", f"median |rate|={med:.4g} -> percent"
    return "bp", f"median |rate|={med:.4g} -> bp"

class QC:
    """Collects every QC finding; becomes the Data_QC sheet."""
    def __init__(self):
        self.rows = []
    def add(self, curve, check, date="", tenor="", detail="", value=""):
        self.rows.append(dict(Curve=curve, Check=check, Date=date, Tenor=tenor, Detail=detail, Value=value))
    def df(self):
        return pd.DataFrame(self.rows, columns=["Curve", "Check", "Date", "Tenor", "Detail", "Value"])

def zero_run_stats(d):
    """d: Series of daily changes (NaN where not observed). Returns share of zero-change days, count of stale runs (> 2 days), longest run."""
    obs = d.dropna()
    if len(obs) == 0:
        return np.nan, 0, 0, 0
    z = (obs.abs() < 1e-9).astype(int).values
    runs, cur = [], 0
    for x in z:
        if x:
            cur += 1
        else:
            if cur:
                runs.append(cur)
            cur = 0
    if cur:
        runs.append(cur)
    stale_runs = [r for r in runs if r > 2]
    return float(z.mean()), len(stale_runs), (max(runs) if runs else 0), int(sum(stale_runs))

def clean_curve(cid, g, qc):
    """g: long rows for one curve. Returns dict with cleaned levels in bp (retained tenors), flags and per-tenor stats."""
    meta = CURVES[cid]
    g = g.copy()
    g["T"] = g["Tenor"].map(parse_tenor)
    bad_t = g[g["T"].isna()]["Tenor"].unique()
    if len(bad_t):
        qc.add(cid, "Unparseable tenor labels ignored", detail=", ".join(map(str, bad_t)))
    g = g.dropna(subset=["T"])
    dup = int(g.duplicated(["Date", "T"]).sum())
    if dup:
        qc.add(cid, "Duplicate (Date, Tenor) rows - last kept", value=dup)
    # ---- units (Bloomberg exports mix percent / bp / decimal) -> bp
    if "Unit" in g.columns and g["Unit"].notna().any() and g["Unit"].astype(str).str.lower().isin(["percent", "bp", "decimal"]).all():
        unit_by = g.groupby("T")["Unit"].agg(lambda s: s.astype(str).str.lower().mode().iloc[0]).to_dict()
        how = "Unit column supplied"
    else:
        unit_by, how = {}, "inferred from magnitude (override with a Unit column if a tenor is ambiguous)"
    mult = {}
    for t, gg in g.groupby("T"):
        u = unit_by.get(t)
        if u is None:
            u, why = detect_unit(gg["Rate"])
        else:
            why = how
        mult[t] = {"percent": 100.0, "decimal": 10000.0, "bp": 1.0}[u]
        qc.add(cid, "Unit check", tenor=tenor_label(t), detail=f"{u} -> bp ({why})", value=mult[t])
    g["bp"] = g["Rate"] * g["T"].map(mult)
    w = g.pivot_table(index="Date", columns="T", values="bp", aggfunc="last").sort_index()
    n_raw_dates = len(w)
    # ---- pricing source: one source / close time per curve, else series break
    break_dates = []
    if "Source" in g.columns and g["Source"].notna().any():
        s = g.groupby("Date")["Source"].agg(lambda x: "/".join(sorted(set(map(str, x)))))
        srcs = sorted(set(g["Source"].dropna().astype(str)))
        if len(srcs) > 1:
            chg = s[s != s.shift(1)].iloc[1:]
            for d0, sv in chg.items():
                break_dates.append(pd.Timestamp(d0))
                qc.add(cid, "Mixed pricing source -> series break (change on this date treated as missing)", date=d0.date(), detail=f"source now {sv}; sources in curve: {', '.join(srcs)}")
        else:
            qc.add(cid, "Pricing source", detail=f"single source {srcs[0]}; close time not in export - confirm one close time per curve")
    else:
        qc.add(cid, "Pricing source", detail="no Source column supplied - cannot verify single pricing source / close time (confirm manually)")
    # ---- Bloomberg fill rows: weekends and rows identical to previous row across all tenors
    wk = w.index.dayofweek >= 5
    d1 = w.diff()
    nobs = d1.notna().sum(axis=1)
    same = ((d1.abs() < 1e-9).sum(axis=1) == nobs) & (nobs >= min(4, w.shape[1]))
    fill_rows = pd.Series(wk, index=w.index) | same
    for d0 in w.index[wk]:
        qc.add(cid, "Weekend row removed (non-trading day)", date=d0.date())
    for d0 in w.index[same & ~pd.Series(wk, index=w.index)]:
        qc.add(cid, "Fill-forward row removed (all tenors identical to prior day -> missing, not zero change)", date=d0.date())
    w = w[~fill_rows.values]
    # ---- bad prints: > k robust SD AND reversed next day -> removed
    k, rev = P("BADPRINT_SD"), P("BADPRINT_REVERT")
    for tn in list(w.columns):
        s = w[tn]
        d = s.diff()
        sd = robust_sd(d)
        dn = d.shift(-1)
        m = (d.abs() > k * sd) & ((d + dn).abs() < rev * d.abs()) & (np.sign(d) * np.sign(dn) < 0)
        for d0 in s.index[m.fillna(False).values]:
            qc.add(cid, "Bad print removed (>%.0f SD, reverted next day)" % k, date=d0.date(), tenor=tenor_label(tn),
                   detail=f"level {s.loc[d0]:.2f}bp, change {d.loc[d0]:.2f}bp vs robust SD {sd:.2f}bp; next-day change {dn.loc[d0]:.2f}bp", value=round(float(s.loc[d0]), 3))
            w.loc[d0, tn] = np.nan
    # ---- stale / missing fixings: forward-fill <= N days, flagged
    wf = w.ffill(limit=int(P("FFILL_MAX_DAYS")))
    filled = w.isna() & wf.notna()
    for tn in wf.columns:
        n = int(filled[tn].sum())
        if n:
            qc.add(cid, "Stale/missing fixings forward-filled (<=%d days, flagged)" % P("FFILL_MAX_DAYS"), tenor=tenor_label(tn), value=n,
                   detail="dates: " + ", ".join(str(x.date()) for x in filled.index[filled[tn].values][:8]) + (" ..." if n > 8 else ""))
    # ---- per-tenor liquidity stats (observed pairs only: exclude filled days)
    obsd = w.diff()
    rows = []
    cover = w.notna().mean()
    for tn in w.columns:
        zs, nrun, mrun, drun = zero_run_stats(obsd[tn])
        rows.append(dict(Curve=cid, Tenor=tenor_label(tn), T=tn, Quotes=int(w[tn].notna().sum()), Coverage=float(cover[tn]),
                         FirstDate=w[tn].first_valid_index(), LastDate=w[tn].last_valid_index(), ZeroChangeShare=zs, StaleRuns_gt2d=nrun,
                         LongestUnchangedRun=mrun, DaysInStaleRuns=drun, FilledDays=int(filled[tn].sum())))
    ts = pd.DataFrame(rows)
    drop_reason = {}
    for _, r in ts.iterrows():
        if r["Coverage"] < P("MIN_COVERAGE"):
            drop_reason[r["T"]] = f"coverage {r['Coverage']:.0%} < {P('MIN_COVERAGE'):.0%}"
        elif r["ZeroChangeShare"] > P("ZERO_CHANGE_DROP"):
            drop_reason[r["T"]] = f"zero-change share {r['ZeroChangeShare']:.0%} > {P('ZERO_CHANGE_DROP'):.0%}"
    ts["Status"] = ts["T"].map(lambda t: "DROPPED: " + drop_reason[t] if t in drop_reason else
                              ("kept - illiquid flag" if ts.loc[ts["T"] == t, "ZeroChangeShare"].iloc[0] > P("ZERO_CHANGE_FLAG") else "kept"))
    for t, why in drop_reason.items():
        qc.add(cid, "Tenor dropped", tenor=tenor_label(t), detail=why)
    keep = [t for t in wf.columns if t not in drop_reason]
    # version including illiquid tenors (for the with/without comparison)
    incl_cols = [t for t in wf.columns if ts.loc[ts["T"] == t, "Coverage"].iloc[0] >= P("MIN_COVERAGE")]
    lv_incl = wf[incl_cols].dropna(how="any")
    lv = wf[keep].dropna(how="any")
    n_drop_rows = int(len(wf.dropna(how="all")) - len(lv))
    if n_drop_rows:
        qc.add(cid, "Dates dropped (a retained tenor still missing after <=%d-day fill)" % P("FFILL_MAX_DAYS"), value=n_drop_rows)
    filled = filled[keep].reindex(lv.index)
    status = "ok"
    if len(keep) < P("MIN_TENORS") or len(lv) < P("MIN_OBS_PCA"):
        status = f"SKIPPED: {len(keep)} tenors / {len(lv)} dates after cleaning"
        qc.add(cid, "Curve skipped", detail=status)
    # ---- ticker-roll / level-shift breaks: one tenor shifts persistently vs its neighbours (butterfly residual) -> back-adjust.
    #      Iterative: apply the largest qualifying shift first (a break in one tenor also moves its neighbours' residuals).
    if status == "ok":
        cols = list(lv.columns)
        n_t = len(cols)
        def resid_series(jx):
            if 0 < jx < n_t - 1:
                return lv[cols[jx]] - 0.5 * (lv[cols[jx - 1]] + lv[cols[jx + 1]])
            return lv[cols[jx]] - lv[cols[1 if jx == 0 else jx - 1]]
        for _ in range(4):
            best = None
            for jx in range(n_t):
                S = resid_series(jx)
                sd = robust_sd(S - S.rolling(60, center=True, min_periods=20).median())
                dS = S.diff()
                for i in np.where(dS.abs().values > P("ROLL_BREAK_SD") * sd)[0]:
                    if i < 10 or i + 30 > len(S):
                        continue
                    a, b, c = S.iloc[i - 10:i].median(), S.iloc[i:i + 10].median(), S.iloc[i + 10:i + 30].median()
                    if abs(b - a) > 0.8 * abs(dS.iloc[i]) and abs(b - a) > P("ROLL_BREAK_SD") * sd and abs(c - a) > 0.6 * abs(b - a):
                        score = abs(b - a) / sd
                        if best is None or score > best[0]:
                            best = (score, jx, i, b - a, sd)
            if best is None:
                break
            _, jx, i, shift, sd = best
            tn = cols[jx]
            qc.add(cid, "Possible roll / ticker-change break (persistent shift vs neighbouring tenors) -> history back-adjusted", date=lv.index[i].date(),
                   tenor=tenor_label(tn), detail=f"shift {shift:+.2f}bp ({abs(shift) / sd:.1f}x the tenor's robust level-residual SD {sd:.2f}bp); levels before this date shifted by {shift:+.2f}bp", value=round(float(shift), 3))
            lv.iloc[:i, jx] += shift
    qc.add(cid, "Coverage", detail=f"{len(lv)} clean dates from {lv.index.min().date() if len(lv) else 'n/a'} to {lv.index.max().date() if len(lv) else 'n/a'}; raw dates {n_raw_dates}",
           value=len(lv))
    return dict(cid=cid, levels=lv, levels_incl=lv_incl, filled=filled, tstats=ts, dropped=[tenor_label(t) for t in drop_reason],
                keep=keep, break_dates=break_dates, status=status)

# =====================================================================================
# 3. BOOTSTRAPPING (same method for every curve)
# =====================================================================================
def bootstrap_dfs(par_bp, pillars, acc):
    """par_bp: DataFrame (dates x pillars int) in bp. Returns discount factors DataFrame (dates x 0..maxT)."""
    n = len(par_bp)
    maxT = max(pillars)
    lnD = {0: np.zeros(n)}
    Dsum_known = np.zeros(n)          # sum of D_i for annual dates i <= previous pillar
    prev = 0
    grid = {0: np.ones(n)}
    for T in pillars:
        S = par_bp[T].values / 1e4
        lnP = lnD[prev]
        lo, hi = np.full(n, -4.0), np.full(n, 1.0)
        def g(x):
            tot = Dsum_known.copy()
            for i in range(prev + 1, T + 1):
                tot += np.exp(lnP + (x - lnP) * (i - prev) / (T - prev))
            return 1.0 - np.exp(x) - S * acc * tot
        for _ in range(90):
            mid = 0.5 * (lo + hi)
            pos = g(mid) > 0
            lo = np.where(pos, mid, lo)
            hi = np.where(pos, hi, mid)
        x = 0.5 * (lo + hi)
        for i in range(prev + 1, T + 1):
            lnD[i] = lnP + (x - lnP) * (i - prev) / (T - prev)
            grid[i] = np.exp(lnD[i])
            Dsum_known = Dsum_known + grid[i]
        prev = T
    D = pd.DataFrame({i: grid[i] for i in range(0, maxT + 1)}, index=par_bp.index)
    return D

def parse_block_pairs():
    out = []
    for tok in str(P("BLOCK_PAIRS")).split(";"):
        tok = tok.strip()
        if "-" in tok:
            a, b = tok.split("-")
            out.append((int(a), int(b)))
    return out

def build_curve(cid, levels_bp):
    """Bootstrap and derive forwards. levels_bp: cleaned par levels (bp) for retained tenors."""
    acc = CURVES[cid]["dc_factor"]
    int_ten = sorted(t for t in levels_bp.columns if is_int_year(t))
    pill = [int(round(t)) for t in int_ten]
    par = levels_bp[int_ten].copy()
    par.columns = pill
    D = bootstrap_dfs(par, pill, acc)
    # re-price every input par rate
    err = {}
    for T in pill:
        annuity = acc * D[[i for i in range(1, T + 1)]].sum(axis=1)
        rep = (1.0 - D[T]) / annuity * 1e4
        err[T] = float((rep - par[T]).abs().max())
    maxT = max(pill)
    supported = [k for k in range(0, maxT) if (k == 0 or k in pill) and (k + 1 in pill)]
    interp = [k for k in range(0, maxT) if k not in supported]
    fw = lambda a, b: ((D[a] / D[b]) ** (1.0 / (b - a)) - 1.0) / acc * 1e4
    sup = pd.DataFrame({f"{k}y1y": fw(k, k + 1) for k in supported}, index=par.index)
    blocks, seen = OrderedDict(), set()
    for a, b in zip(pill[:-1], pill[1:]):
        if b - a > 1:
            blocks[f"{a}y{b - a}y"] = (a, b, "gap between adjacent quoted pillars")
    for a, b in parse_block_pairs():
        if a in pill and b in pill and b - a >= 1 and f"{a}y{b - a}y" not in blocks and b - a > 1:
            blocks[f"{a}y{b - a}y"] = (a, b, "derived from quoted pillars %dY and %dY" % (a, b))
    blk = pd.DataFrame({k: fw(a, b) for k, (a, b, _) in blocks.items()}, index=par.index)
    interp_df = pd.DataFrame({f"{k}y1y": fw(k, k + 1) for k in interp}, index=par.index)
    return dict(D=D, pillars=pill, par=par, acc=acc, reprice_err=err, max_err=max(err.values()), supported_starts=supported,
                supported=sup, blocks=blk, block_meta=blocks, interp_starts=interp, interp=interp_df)

# =====================================================================================
# 4. PCA CORE
# =====================================================================================
def _norm_signs(V):
    """PC1 positive on all tenors (sum>0); PC2 long-minus-short >0; PC3 belly-minus-wings >0; PC4+: largest |loading| positive."""
    V = V.copy()
    n = V.shape[0]
    for k in range(V.shape[1]):
        v = V[:, k]
        if k == 0:
            s = np.sum(v)
        elif k == 1:
            s = v[-1] - v[0]
        elif k == 2:
            s = v[n // 2] - 0.5 * (v[0] + v[-1])
        else:
            s = v[np.argmax(np.abs(v))]
        if s < 0:
            V[:, k] = -v
    return V

def pca_fit(X, standardise=False):
    """PCA on daily changes (bp). standardise=False: covariance matrix on bp changes (MAIN); True: correlation matrix (robustness)."""
    X = np.asarray(X, float)
    T, N = X.shape
    mu = X.mean(0)
    Xc = X - mu
    sd = Xc.std(0, ddof=1) if standardise else np.ones(N)
    sd[sd == 0] = 1.0
    Z = Xc / sd
    C = np.atleast_2d(np.cov(Z, rowvar=False, ddof=1))
    w, V = np.linalg.eigh(C)
    o = np.argsort(w)[::-1]
    w, V = np.clip(w[o], 0, None), V[:, o]
    V = _norm_signs(V)
    return dict(mu=mu, sd=sd, eig=w, ve=w / w.sum(), V=V, scores=Z @ V, standardise=standardise, n=T, N=N)

def sign_changes(v, tol=0.0):
    s = np.sign(np.where(np.abs(v) < tol, 0, v))
    s = s[s != 0]
    return int(np.sum(s[1:] != s[:-1])) if len(s) > 1 else 0

def pc_labels(V):
    """Label PCs by loading shape only if the shape supports it."""
    N, K = V.shape
    out = []
    for k in range(min(K, 3)):
        v = V[:, k]
        sc = sign_changes(v, tol=0.05 * np.max(np.abs(v)))
        if k == 0:
            clean = (sc == 0) and (np.min(np.abs(v)) >= 0.25 * np.max(np.abs(v)))
            out.append("Level" if clean else ("Level-like (uneven loadings)" if sc == 0 else "NOT a clean level factor"))
        elif k == 1:
            out.append("Slope" if sc == 1 else "Slope-like (monotone)" if sc == 0 and abs(v[-1] - v[0]) > 0.3 * np.max(np.abs(v)) else "Unlabelled (shape not slope)")
        else:
            out.append("Curvature" if sc == 2 else "Unlabelled (shape not curvature)")
    return out

def cong(a, b):
    """Tucker congruence coefficient (absolute value: eigenvector sign is arbitrary)."""
    a, b = np.asarray(a, float), np.asarray(b, float)
    return float(abs(np.dot(a, b)) / (np.linalg.norm(a) * np.linalg.norm(b)))

def rolling_pca(dX, full, win, step):
    """Rolling covariance PCA. Returns DataFrame indexed by window-end date: ve1..ve3, cong1..cong3, break flag."""
    rows = []
    idx = dX.index
    for end in range(win, len(dX) + 1, step):
        seg = dX.iloc[end - win:end]
        p = pca_fit(seg.values)
        nk = min(3, p["V"].shape[1])
        r = dict(Date=idx[end - 1])
        for k in range(3):
            r[f"VE_PC{k + 1}"] = p["ve"][k] if k < nk else np.nan
        for k in range(3):
            r[f"Cong_PC{k + 1}"] = cong(p["V"][:, k], full["V"][:, k]) if k < nk else np.nan
        rows.append(r)
    df = pd.DataFrame(rows).set_index("Date")
    for k in (1, 2, 3):
        df[f"Break_PC{k}"] = df[f"Cong_PC{k}"] < P("CONGR_BREAK")
    df["Break"] = df["Break_PC1"]                      # headline break = PC1 (PC2/PC3 rotations reported separately)
    return df

def pct_inc(arr, x):
    arr = np.asarray(arr, float)
    arr = arr[np.isfinite(arr)]
    n = len(arr)
    return float(np.sum(arr < x) / (n - 1)) if n > 1 else np.nan

def trail_stats(series, win=None):
    win = win or int(P("Z_WIN"))
    a = np.asarray(series.dropna().values if hasattr(series, "dropna") else series, float)[-win:]
    x = a[-1]
    mu, sd = a.mean(), a.std(ddof=1)
    return dict(latest=float(x), mean=float(mu), sd=float(sd), z=float((x - mu) / sd) if sd > 0 else np.nan, pct=pct_inc(a, x), n=len(a))

def ols(y, X, const=True):
    y = np.asarray(y, float)
    X = np.asarray(X, float)
    if X.ndim == 1:
        X = X[:, None]
    A = np.column_stack([np.ones(len(y)), X]) if const else X
    b, *_ = np.linalg.lstsq(A, y, rcond=None)
    res = y - A @ b
    sst = np.sum((y - y.mean()) ** 2)
    r2 = 1 - np.sum(res ** 2) / sst if sst > 0 else np.nan
    dof = max(len(y) - A.shape[1], 1)
    s2 = np.sum(res ** 2) / dof
    try:
        se = np.sqrt(np.diag(s2 * np.linalg.inv(A.T @ A)))
    except np.linalg.LinAlgError:
        se = np.full(A.shape[1], np.nan)
    return dict(b=b, se=se, r2=r2, res=res, n=len(y), resid_sd=float(np.sqrt(s2)))

def ar1_halflife(r):
    """AR(1) on the residual level; returns (half-life days, stable flag, phi). Unstable if phi>=1, phi<=0 or halves disagree >3x."""
    r = np.asarray(r, float)
    r = r[np.isfinite(r)]
    if len(r) < 80:
        return np.nan, False, np.nan
    def est(x):
        A = np.column_stack([np.ones(len(x) - 1), x[:-1]])
        return np.linalg.lstsq(A, x[1:], rcond=None)[0][1]
    hl = lambda p: -math.log(2) / math.log(p) if 0 < p < 1 else np.inf
    phi = est(r)
    h = len(r) // 2
    hs = [hl(phi), hl(est(r[:h])), hl(est(r[h:]))]
    stable = bool(all(np.isfinite(hs)) and max(hs[1:]) / min(hs[1:]) <= 3.0)
    return (hs[0] if np.isfinite(hs[0]) else np.nan), stable, float(phi)

# =====================================================================================
# 5. PER-CURVE ANALYSIS
# =====================================================================================
class Hist:
    """Registry of trailing-window series written to the Hist sheet; result-sheet formulas reference these columns."""
    FIRST_ROW = 5                     # 0-based first observation row on the sheet
    def __init__(self):
        self.cols = OrderedDict()
        self.n = int(P("Z_WIN"))
    def add(self, name, series):
        s = pd.Series(series).dropna()
        vals = s.values[-self.n:]
        last = s.index[-1] if len(s) else ""
        self.cols[name] = dict(vals=vals, last=last, col=len(self.cols) + 1)
        return name
    def _letter(self, c):
        from xlsxwriter.utility import xl_col_to_name
        return xl_col_to_name(c)
    def rng(self, name):
        c = self._letter(self.cols[name]["col"])
        return f"Hist!${c}${self.FIRST_ROW + 1}:${c}${self.FIRST_ROW + self.n}"
    def latest(self, name):
        c = self._letter(self.cols[name]["col"])
        return f"Hist!${c}${self.FIRST_ROW + self.n}"
    def stats(self, name):
        return trail_stats(self.cols[name]["vals"])
    def f_z(self, name):
        return f"=IF(STDEV({self.rng(name)})>0,({self.latest(name)}-AVERAGE({self.rng(name)}))/STDEV({self.rng(name)}),0)"
    def f_pct(self, name):
        return f"=PERCENTRANK({self.rng(name)},{self.latest(name)},6)"
    def f_mean(self, name):
        return f"=AVERAGE({self.rng(name)})"
    def f_sd(self, name):
        return f"=STDEV({self.rng(name)})"
    def f_latest(self, name):
        return f"={self.latest(name)}"

def cum_roll(series, w=20):
    return series.rolling(w).sum()

def fwd_consistency(dF, cid, qc):
    """Flag dates where one derived forward point jumps while the others do not (usually a bad par quote)."""
    sd = dF.apply(robust_sd)
    z = dF.div(sd)
    thr = P("FWD_JUMP_Z")
    flagged = []
    for d0, row in z.iterrows():
        big = row.abs() > thr
        if big.any() and (row.abs() > thr / 2.5).sum() <= 2:
            pts = list(row.index[big.values])
            flagged.append(d0)
            qc.add(cid, "Forward-consistency flag (isolated jump in derived forward - check par quote)", date=d0.date(),
                   tenor=", ".join(pts), detail="z=" + ", ".join(f"{row[p]:+.1f}" for p in pts), value=round(float(row.abs().max()), 2))
    return flagged

def analyse_curve(cid, C, B, start, end, qc, hist):
    """All per-curve results on the analysis sample [start,end]."""
    lv_full = C["levels"]
    lv = lv_full.loc[start:end]
    R = dict(cid=cid, levels=lv, levels_full=lv_full)
    dY = lv.diff().dropna()
    dY = dY.drop([d for d in C["break_dates"] if d in dY.index])
    dY.columns = [tenor_label(t) for t in dY.columns]
    R["dY"] = dY
    R["tenors"] = list(dY.columns)
    R["par_cov"] = pca_fit(dY.values)
    R["par_corr"] = pca_fit(dY.values, True)
    R["labels"] = pc_labels(R["par_cov"]["V"])
    K = R["par_cov"]["V"].shape[1]
    kk = min(4, K)
    R["scores"] = pd.DataFrame(R["par_cov"]["scores"][:, :kk], index=dY.index, columns=[f"PC{i + 1}" for i in range(kk)])
    # full history (par cov)
    dYf = lv_full.diff().dropna()
    dYf = dYf.drop([d for d in C["break_dates"] if d in dYf.index])
    R["par_cov_full"] = pca_fit(dYf.values)
    R["n_full"] = len(dYf)
    # rolling stability
    R["rolling"] = rolling_pca(dY, R["par_cov"], int(P("ROLL_WIN")), int(P("ROLL_STEP")))
    # latest scores, 1y percentile / z (daily and 20d cumulative)
    sc = R["scores"]
    cum = sc.rolling(20).sum()
    R["cum20"] = cum
    lat = OrderedDict()
    for k, nm in enumerate(["Level (PC1)", "Slope (PC2)", "Curvature (PC3)"][:min(3, sc.shape[1])]):
        col = f"PC{k + 1}"
        hist.add(f"{cid}|{col}|daily", sc[col])
        hist.add(f"{cid}|{col}|cum20", cum[col].dropna())
        lat[nm] = dict(daily=trail_stats(sc[col]), cum20=trail_stats(cum[col].dropna()), col=col)
    R["latest_scores"] = lat
    # 3-PC fit residual by tenor (bp)
    V3 = R["par_cov"]["V"][:, :3]
    Xhat = R["par_cov"]["mu"] + sc.values[:, :3] @ V3.T
    resid = dY.values - Xhat
    R["fit_resid_sd"] = pd.Series(resid.std(0, ddof=1), index=dY.columns)
    R["fit_resid_last"] = pd.Series(resid[-1], index=dY.columns)
    R["fit_r2_3pc"] = float(R["par_cov"]["ve"][:3].sum())
    # forwards on the sample
    sup = B["supported"].loc[start:end]
    blk = B["blocks"].loc[start:end]
    R["sup"], R["blk"], R["B"] = sup, blk, B
    dF = sup.diff().dropna()
    dF = dF.drop([d for d in C["break_dates"] if d in dF.index])
    R["dF"] = dF
    R["fwd_flags"] = fwd_consistency(dF, cid, qc) if len(dF) > 20 else []
    if sup.shape[1] >= 3 and len(dF) >= P("MIN_OBS_PCA"):
        R["fwd_cov"] = pca_fit(dF.values)
        R["fwd_corr"] = pca_fit(dF.values, True)
    else:
        R["fwd_cov"] = R["fwd_corr"] = None
    # with vs without illiquid tenors (par cov)
    inc = C["levels_incl"].loc[start:end]
    if inc.shape[1] > lv.shape[1] and len(inc) > P("MIN_OBS_PCA"):
        dI = inc.diff().dropna()
        R["par_incl"] = pca_fit(dI.values)
        R["incl_tenors"] = [tenor_label(t) for t in dI.columns]
    else:
        R["par_incl"], R["incl_tenors"] = None, None
    return R

# =====================================================================================
# 6. GLOBAL vs LOCAL
# =====================================================================================
def interval_sum(s, starts, ends, mode):
    """Sum of series s over its own dates: mode 'same': start < d <= end ; mode 'lag': start <= d < end."""
    idx = s.index.values
    cs = np.concatenate([[0.0], np.cumsum(s.values)])
    st, en = np.asarray(starts, dtype="datetime64[ns]"), np.asarray(ends, dtype="datetime64[ns]")
    if mode == "same":
        lo, hi = np.searchsorted(idx, st, side="right"), np.searchsorted(idx, en, side="right")
    else:
        lo, hi = np.searchsorted(idx, st, side="left"), np.searchsorted(idx, en, side="left")
    return cs[hi] - cs[lo]

def global_regressors(R_loc, R_glob, mode):
    """Global (SOFR) PC1-3 scores aggregated over each local return interval. mode: same | lag | avg."""
    lv_idx = R_loc["levels"].index
    prev = pd.Series(lv_idx[:-1], index=lv_idx[1:])
    sc = R_loc["scores"]
    starts = prev.reindex(sc.index).values
    ends = sc.index.values
    G = R_glob["scores"]
    out = {}
    for k in range(3):
        col = f"PC{k + 1}"
        a = interval_sum(G[col], starts, ends, "same")
        b = interval_sum(G[col], starts, ends, "lag")
        out[col] = {"same": a, "lag": b, "avg": 0.5 * (a + b)}[mode]
    return pd.DataFrame(out, index=sc.index)

def global_vs_local(res):
    """Returns dict: per-curve regression results for the three timing versions + chosen version."""
    g = res[GLOBAL_CURVE]
    out = OrderedDict()
    for cid in EM_CURVES:
        if cid not in res:
            continue
        r = res[cid]
        per = OrderedDict()
        for mode in ("same", "lag", "avg"):
            Gx = global_regressors(r, g, mode)
            regs = OrderedDict()
            for k in range(3):
                col = f"PC{k + 1}"
                y = r["scores"][col].values
                fit = ols(y, Gx.values)
                regs[col] = dict(r2=fit["r2"], b=fit["b"][1:], se=fit["se"][1:], resid_sd=fit["resid_sd"], res=pd.Series(fit["res"], index=r["scores"].index))
            ve = r["par_cov"]["ve"][:3]
            wr2 = float(sum(ve[k] * regs[f"PC{k + 1}"]["r2"] for k in range(3)) / ve.sum())
            per[mode] = dict(regs=regs, wr2=wr2, G=Gx)
        out[cid] = per
    # choice: version with the highest median weighted R2 across markets
    med = {m: np.median([out[c][m]["wr2"] for c in out]) for m in ("same", "lag", "avg")}
    choice = max(med, key=med.get)
    return dict(per=out, median_wr2=med, choice=choice)

def rolling_spill(res, gl):
    """Rolling 6M / 12M beta and R2 of local PC1 on global PC1 (chosen timing)."""
    wins = [int(x) for x in str(P("SPILL_WINS")).split(",")]
    ch = gl["choice"]
    out = OrderedDict()
    for cid, per in gl["per"].items():
        y = res[cid]["scores"]["PC1"]
        x = per[ch]["G"]["PC1"]
        d = {}
        for w in wins:
            cov = y.rolling(w).cov(x)
            var = x.rolling(w).var()
            cor = y.rolling(w).corr(x)
            d[f"beta_{w}d"] = cov / var
            d[f"r2_{w}d"] = cor ** 2
        out[cid] = pd.DataFrame(d)
    return out, wins

def decoupling_episodes(roll, full_r2, w6):
    lo, hi = P("DECOUPLE_LO") * full_r2, P("DECOUPLE_HI") * full_r2
    s = roll[f"r2_{w6}d"].dropna()
    state = pd.Series(0, index=s.index)
    if full_r2 > 0.02:
        state[s < lo] = -1
        state[s > hi] = 1
    eps = []
    for val, name in ((-1, "decoupled"), (1, "re-coupled/high")):
        m = (state == val).astype(int)
        grp = (m.diff().fillna(0) != 0).cumsum()
        for _, seg in m.groupby(grp):
            if seg.iloc[0] == 1 and len(seg) >= 10:
                eps.append((name, seg.index[0], seg.index[-1], len(seg)))
    eps.sort(key=lambda x: x[1])
    cur = {-1: "decoupled", 0: "normal", 1: "re-coupled/high"}[int(state.iloc[-1])] if len(state) else "n/a"
    return eps, cur

# =====================================================================================
# 7. POOLED EM ASIA PCA AND TIE-BACK
# =====================================================================================
def par_series(R, T):
    lv = R["levels"]
    return lv[float(T)] if float(T) in lv.columns else None

def pooled_pca(level_dict, min_mk=None):
    min_mk = min_mk or int(P("POOL_MIN_MARKETS"))
    lv = {k: v for k, v in level_dict.items() if v is not None}
    if len(lv) < min_mk:
        return None
    panel = pd.concat(lv, axis=1, join="inner").dropna()
    d = panel.diff().dropna()
    if len(d) < P("MIN_OBS_PCA"):
        return None
    p, pc = pca_fit(d.values), pca_fit(d.values, True)
    corr1 = {m: float(np.corrcoef(d[m].values, p["scores"][:, 0])[0, 1]) for m in d.columns}
    return dict(names=list(d.columns), p=p, pc=pc, n=len(d), corr1=corr1, dates=d.index, panel=panel)

def pooled_sets(res):
    """Pooled PCA sets: par 2Y/5Y/10Y, slopes 2s5s/5s10s, supported forwards 1y1y/3y1y; each with and without SOFR."""
    def lvl(c, T):
        return par_series(res[c], T)
    def slope(c, a, b):
        x, y = par_series(res[c], a), par_series(res[c], b)
        return None if x is None or y is None else (y - x)
    def fwd(c, lab):
        s = res[c]["sup"]
        return s[lab] if lab in s.columns else None
    specs = OrderedDict([
        ("Par 2Y", lambda c: lvl(c, 2)), ("Par 5Y", lambda c: lvl(c, 5)), ("Par 10Y", lambda c: lvl(c, 10)),
        ("Slope 2s5s", lambda c: slope(c, 2, 5)), ("Slope 5s10s", lambda c: slope(c, 5, 10)),
        ("Fwd 1y1y (comparison)", lambda c: fwd(c, "1y1y")), ("Fwd 3y1y (comparison)", lambda c: fwd(c, "3y1y")),
    ])
    out = OrderedDict()
    for nm, fn in specs.items():
        for var, curves in (("ex-SOFR", EM_CURVES), ("incl. SOFR", EM_CURVES + [GLOBAL_CURVE])):
            d = {c: fn(c) for c in curves if c in res}
            excl = [c for c, v in d.items() if v is None]
            r = pooled_pca(d)
            out[(nm, var)] = dict(res=r, excluded=excl)
    return out

def tieback(res, gl):
    """Share of each market's own PC1-3 explained by pooled EM Asia PCs (leave-one-out, par 5Y, PC1-2) vs SOFR PCs (PC1-2)."""
    rows = []
    g = res[GLOBAL_CURVE]
    mode = gl["choice"]
    for cid in EM_CURVES:
        if cid not in res:
            continue
        others = {c: par_series(res[c], 5) for c in EM_CURVES if c != cid and c in res}
        others = {k: v for k, v in others.items() if v is not None}
        if len(others) < 3:
            continue
        panel = pd.concat(others, axis=1, join="inner").dropna()
        d = panel.diff().dropna()
        if len(d) < P("MIN_OBS_PCA"):
            continue
        p = pca_fit(d.values)
        E = pd.DataFrame(p["scores"][:, :2], index=d.index)
        prev = panel.index.to_series().shift(1).reindex(d.index)
        starts, ends = prev.values, d.index.values
        S = pd.DataFrame({k: (interval_sum(g["scores"][f"PC{k + 1}"], starts, ends, "same") if mode == "same" else
                              interval_sum(g["scores"][f"PC{k + 1}"], starts, ends, "lag") if mode == "lag" else
                              0.5 * (interval_sum(g["scores"][f"PC{k + 1}"], starts, ends, "same") + interval_sum(g["scores"][f"PC{k + 1}"], starts, ends, "lag")))
                          for k in range(2)}, index=d.index)
        L = res[cid]["scores"].cumsum()
        allidx = L.index.union(d.index).union(pd.DatetimeIndex(starts))
        La = L.reindex(allidx).ffill()
        row = dict(Curve=cid, N=len(d))
        for k in range(3):
            y = (La[f"PC{k + 1}"].reindex(d.index).values - La[f"PC{k + 1}"].reindex(pd.DatetimeIndex(starts)).values)
            ok = np.isfinite(y)
            if ok.sum() < P("MIN_OBS_PCA"):
                for tag in ("EM_pool", "SOFR", "Both"):
                    row[f"PC{k + 1}_R2_{tag}"] = np.nan
                continue
            row[f"PC{k + 1}_R2_EM_pool"] = ols(y[ok], E.values[ok])["r2"]
            row[f"PC{k + 1}_R2_SOFR"] = ols(y[ok], S.values[ok])["r2"]
            row[f"PC{k + 1}_R2_Both"] = ols(y[ok], np.column_stack([E.values[ok], S.values[ok]]))["r2"]
        rows.append(row)
    return pd.DataFrame(rows)

# =====================================================================================
# 8. RELATIVE VALUE, FLIES, FORWARDS
# =====================================================================================
def fwd_rolls(f_latest, starts):
    """1M roll-down (bp) of each supported 1Y forward: expected change in the forward when its start date rolls 1M closer,
    curve shape unchanged, linear in start date between adjacent supported points. First point uses the one-sided slope."""
    f = pd.Series(f_latest)
    out = {}
    for i, k in enumerate(starts):
        if i > 0:
            slope = f.iloc[i] - f.iloc[i - 1]
        elif len(starts) > 1:
            slope = f.iloc[1] - f.iloc[0]
        else:
            slope = 0.0
        out[k] = -slope / 12.0
    return out

def par_roll(lv_latest, T):
    """1M roll-down (bp) of the par rate at tenor T using the adjacent lower pillar (upper for the shortest)."""
    ten = sorted(lv_latest.index)
    i = ten.index(T)
    if i > 0:
        slope = (lv_latest[T] - lv_latest[ten[i - 1]]) / (T - ten[i - 1])
    else:
        slope = (lv_latest[ten[1]] - lv_latest[T]) / (ten[1] - T)
    return -slope / 12.0

def rv_analysis(cid, R, hist, qc):
    """PCA on supported 1Y forwards (+ labelled block forwards), 3-PC fitted curve, residuals, z, half-lives, PCA-neutral flies."""
    B = R["B"]
    sup, blk = R["sup"], R["blk"]
    pts = list(sup.columns)
    types = {p: "1Y forward (supported)" for p in pts}
    gaps = [k for k, (a, b, why) in B["block_meta"].items() if why.startswith("gap")] if P("RV_INCLUDE_BLOCKS") else []
    panel = pd.concat([sup] + ([blk[gaps]] if gaps else []), axis=1).dropna()
    for gk in gaps:
        types[gk] = "BLOCK forward (labelled)"
    if panel.shape[1] < 4 or len(panel) < P("MIN_OBS_PCA"):
        return dict(ok=False, reason=f"only {panel.shape[1]} usable forward points - RV PCA needs >=4")
    dP = panel.diff().dropna()
    dP = dP.drop([d for d in R.get("break_dates", []) if d in dP.index])
    p = pca_fit(dP.values)
    K = 3
    # fitted daily changes = UNCENTRED projection on PC1-3 (a centred fit cumulates to exactly zero residual at the sample end by construction)
    V3 = p["V"][:, :K]
    Xhat = dP.values @ V3 @ V3.T
    fitted = pd.DataFrame(np.cumsum(Xhat, axis=0), index=dP.index, columns=panel.columns) + panel.iloc[0].values
    actual = panel.loc[dP.index]
    resid = actual - fitted
    rows = []
    for c in panel.columns:
        hist.add(f"{cid}|RV|{c}", resid[c])
        hl, stable, phi = ar1_halflife(resid[c].values[-int(P("HL_WIN")):])
        st = trail_stats(resid[c])
        rows.append(dict(Curve=cid, Point=c, Type=types[c], Actual=float(actual[c].iloc[-1]), Fitted=float(fitted[c].iloc[-1]),
                         Residual=float(resid[c].iloc[-1]), z=st["z"], HalfLife=hl, Stable=stable, phi=phi))
    tab = pd.DataFrame(rows)
    # PCA-neutral butterflies on consecutive supported 1Y forwards
    flies = []
    starts = R["B"]["supported_starts"]
    latestf = sup.iloc[-1]
    D_last = B["D"].loc[:sup.index[-1]].iloc[-1]
    rolls = fwd_rolls(latestf.values, starts)
    pos = {c: i for i, c in enumerate(panel.columns)}
    for i in range(1, len(starts) - 1):
        if not (starts[i - 1] + 1 == starts[i] and starts[i] + 1 == starts[i + 1]):
            continue
        names = [f"{starts[j]}y1y" for j in (i - 1, i, i + 1)]
        L = np.array([[p["V"][pos[n], 0] for n in names], [p["V"][pos[n], 1] for n in names]])
        w = np.cross(L[0], L[1])
        if abs(w[1]) < 1e-9:
            continue
        w = w / (-w[1])                                   # belly weight = -1
        F = sum(w[j] * sup[names[j]] for j in range(3))
        nm = " / ".join(names)
        hist.add(f"{cid}|FLY|{nm}", F)
        st = trail_stats(F)
        dv01 = np.array([1e-4 * B["acc"] * D_last[starts[j] + 1] for j in (i - 1, i, i + 1)])
        notional = (w / dv01) * dv01[1]
        rollF = float(sum(w[j] * rolls[starts[i - 1 + j]] for j in range(3)))
        spot_leg = starts[i - 1] == 0
        hl, stable, phi = ar1_halflife(F.values[-int(P("HL_WIN")):])
        flies.append(dict(Curve=cid, Fly=nm, w1=w[0], w2=w[1], w3=w[2], DV01_1=dv01[0], DV01_2=dv01[1], DV01_3=dv01[2],
                          N1=notional[0], N2=notional[1], N3=notional[2], Latest=st["latest"], Mean=st["mean"], SD=st["sd"], z=st["z"],
                          Pct=st["pct"], Carry_1M=(np.nan if spot_leg else 0.0), Roll_1M=rollF,
                          RolledZ=((st["latest"] + rollF - st["mean"]) / st["sd"] if st["sd"] > 0 else np.nan),
                          PC1_exposure=float(L[0] @ w), PC2_exposure=float(L[1] @ w), HalfLife=hl, Stable=stable,
                          Direction=("z>0: receive wings / pay belly (short fly)" if st["z"] > 0 else "z<0: pay wings / receive belly (long fly)")))
    return dict(ok=True, table=tab, pca=p, panel=panel, fitted=fitted, resid=resid, actual=actual, flies=flies, types=types)

def forward_tables(cid, R, hist):
    """Slopes / forward-forward spreads between supported points, block forwards with z-scores/percentiles, interpolated (reference only)."""
    sup, blk, B = R["sup"], R["blk"], R["B"]
    rows = []
    st_ = B["supported_starts"]
    for a, b in itertools.combinations(st_, 2):
        gap = b - a
        if gap > 4:
            continue
        nm = f"{b}y1y - {a}y1y"
        s = sup[f"{b}y1y"] - sup[f"{a}y1y"]
        hist.add(f"{cid}|SPR|{nm}", s)
        t = trail_stats(s)
        rows.append(dict(Curve=cid, Series=nm, Type=("Forward-forward spread (adjacent)" if gap == 1 else "Slope between supported points"), Latest=t["latest"],
                         Mean=t["mean"], SD=t["sd"], z=t["z"], Pct=t["pct"], Method="difference of supported 1Y forwards (bp)"))
    brow = []
    for k in blk.columns:
        a, b, why = B["block_meta"][k]
        hist.add(f"{cid}|BLK|{k}", blk[k])
        t = trail_stats(blk[k])
        brow.append(dict(Curve=cid, Point=k, Pillars=f"{a}Y -> {b}Y", Type="BLOCK forward", Latest=t["latest"], Mean=t["mean"], SD=t["sd"], z=t["z"], Pct=t["pct"],
                         Method=f"((D_{a}/D_{b})^(1/{b - a}) - 1)/accrual from bootstrapped DFs; {why}"))
    irow = [dict(Curve=cid, Point=c, Latest_bp=float(B["interp"][c].iloc[-1]), Flag="REFERENCE ONLY - depends on interpolated pillars; not used in any PCA / fly / pair")
            for c in B["interp"].columns]
    return pd.DataFrame(rows), pd.DataFrame(brow), pd.DataFrame(irow)

# =====================================================================================
# 9. CROSS-MARKET PCA-NEUTRAL PAIRS
# =====================================================================================
def pairs_analysis(res, hist):
    instr = [("par", "2Y", 2.0), ("par", "5Y", 5.0), ("par", "10Y", 10.0), ("fwd", "1y1y", None), ("fwd", "3y1y", None)]
    rows, series = [], {}
    for kind, lab, T in instr:
        avail = {}
        for c in EM_CURVES:
            if c not in res:
                continue
            R = res[c]
            if kind == "par" and float(T) in R["levels"].columns:
                avail[c] = (R["levels"][float(T)], R["par_cov"]["V"][R["tenors"].index(lab), 0])
            elif kind == "fwd" and R["fwd_cov"] is not None and lab in R["sup"].columns:
                avail[c] = (R["sup"][lab], R["fwd_cov"]["V"][list(R["sup"].columns).index(lab), 0])
        for a, b in itertools.combinations(avail, 2):
            (sa, la), (sb, lb) = avail[a], avail[b]
            panel = pd.concat([sa, sb], axis=1, join="inner").dropna()
            panel.columns = ["A", "B"]
            if len(panel) < P("MIN_OBS_PCA"):
                continue
            hr = la / lb
            S = panel["A"] - hr * panel["B"]
            t = trail_stats(S)
            d = panel.diff().dropna()
            c1 = float(d.tail(int(P("Z_WIN"))).corr().iloc[0, 1])
            cf = float(d.corr().iloc[0, 1])
            if kind == "par":
                ra = par_roll(res[a]["levels"].iloc[-1].rename(index=lambda x: float(x)), float(T))
                rb = par_roll(res[b]["levels"].iloc[-1].rename(index=lambda x: float(x)), float(T))
            else:
                ra = fwd_rolls(res[a]["sup"].iloc[-1].values, res[a]["B"]["supported_starts"])[int(lab.split("y")[0])]
                rb = fwd_rolls(res[b]["sup"].iloc[-1].values, res[b]["B"]["supported_starts"])[int(lab.split("y")[0])]
            roll = ra - hr * rb
            adj = (t["latest"] + roll - t["mean"]) / t["sd"] if t["sd"] > 0 else np.nan
            key = (kind, lab, a, b)
            series[key] = S
            rows.append(dict(Set=("Par (main)" if kind == "par" else "Forward (comparison)"), Instrument=lab, A=a, B=b, HedgeRatio=hr,
                             Spread=t["latest"], Mean=t["mean"], SD=t["sd"], z=t["z"], Pct=t["pct"], Corr_1y=c1, Corr_full=cf, Roll_1M=roll,
                             Carry_1M=np.nan, AdjZ=adj, N=len(panel),
                             Direction=("z>0: receive %s / pay %s x%.2f" % (a, b, hr) if t["z"] > 0 else "z<0: pay %s / receive %s x%.2f" % (a, b, hr))))
    df = pd.DataFrame(rows)
    top = {}
    for sname in ("Par (main)", "Forward (comparison)"):
        sub = df[df["Set"] == sname].copy()
        sub["absAdj"] = sub["AdjZ"].abs()
        sub = sub.sort_values("absAdj", ascending=False).head(int(P("PAIR_TOP_N"))).drop(columns="absAdj").reset_index(drop=True)
        for _, r in sub.iterrows():
            key = ("par" if sname.startswith("Par") else "fwd", r["Instrument"], r["A"], r["B"])
            r_name = f"PAIR|{r['Set'][:3]}|{r['Instrument']}|{r['A']}|{r['B']}"
            hist.add(r_name, series[key])
        top[sname] = sub
    return df, top

# =====================================================================================
# 10. OFFSHORE LIQUIDITY, FX, EVENTS, BLOOMBERG FORWARD COMPARISON (optional inputs)
# =====================================================================================
def read_optional(path):
    if path and os.path.exists(path):
        try:
            df = pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)
            df.columns = [str(c).strip() for c in df.columns]
            if "Date" in df.columns:
                df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
                df = df.dropna(subset=["Date"])
            return df if len(df) else None
        except Exception:
            return None
    return None

def fx_check(res, fx):
    if fx is None:
        return None
    rows = []
    for cid, g in fx.groupby("Market"):
        if cid not in res:
            continue
        s = pd.to_numeric(g.set_index("Date")["Rate"], errors="coerce").dropna().sort_index()
        ret = np.log(s).diff().dropna() * 100.0
        y = res[cid]["scores"]["PC1"]
        j = pd.concat([y, ret], axis=1, join="inner").dropna()
        if len(j) < P("MIN_OBS_PCA"):
            continue
        j.columns = ["pc1", "fx"]
        f = ols(j["pc1"].values, j["fx"].values)
        rows.append(dict(Curve=cid, N=len(j), Corr_PC1_vs_FX=float(j.corr().iloc[0, 1]), Beta_bp_per_1pct=float(f["b"][1]), R2=f["r2"]))
    return pd.DataFrame(rows) if rows else None

def event_analysis(res, ev):
    if ev is None:
        return None
    rows = []
    for _, e in ev.iterrows():
        cid = str(e.get("Market", "")).strip()
        if cid not in res:
            continue
        sc = res[cid]["scores"]
        i = sc.index.searchsorted(e["Date"])
        if i >= len(sc) or i < 1 or i + 1 >= len(sc):
            continue
        win = sc.iloc[i - 1:i + 2]
        row = dict(Date=e["Date"], Curve=cid, Event=e.get("Event", ""))
        zs = []
        for k in range(min(3, sc.shape[1])):
            col = f"PC{k + 1}"
            mv = float(win[col].sum())
            z = mv / (sc[col].std() * math.sqrt(3))
            row[f"{col}_move_bp"], row[f"{col}_z"] = mv, z
            zs.append(abs(z))
        row["Dominant_factor"] = ["Level", "Slope", "Curvature"][int(np.argmax(zs))]
        rows.append(row)
    return pd.DataFrame(rows) if rows else None

def bbg_fwd_compare(res, bf):
    if bf is None:
        return None
    rows = []
    for (cid, pt), g in bf.groupby(["Market", "Point"]):
        if cid not in res:
            continue
        s = pd.to_numeric(g.set_index("Date")["Rate"], errors="coerce").dropna()
        u, _ = detect_unit(s)
        s = s * {"percent": 100.0, "decimal": 10000.0, "bp": 1.0}[u]
        B = res[cid]["B"]
        mine = pd.concat([B["supported"], B["blocks"]], axis=1)
        if pt not in mine.columns:
            rows.append(dict(Curve=cid, Point=pt, N=0, MeanDiff_bp="n/a", MaxAbsDiff_bp="n/a", Note="point not derivable (not supported / not a block forward)"))
            continue
        j = pd.concat([mine[pt], s], axis=1, join="inner").dropna()
        j.columns = ["mine", "bbg"]
        diff = j["mine"] - j["bbg"]
        rows.append(dict(Curve=cid, Point=pt, N=len(j), MeanDiff_bp=float(diff.mean()), MaxAbsDiff_bp=float(diff.abs().max()), Note="derived - Bloomberg (bp); derived forwards used for analysis"))
    return pd.DataFrame(rows) if rows else None

# =====================================================================================
# 11. RESULTS WORKBOOK WRITER
# =====================================================================================
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

SHEET_ORDER_HEAD = ["README", "Assumptions", "Data_QC", "Curve_Build", "PCA_Variance", "PCA_Loadings", "PCA_Scores"]
SHEET_ORDER_TAIL = ["Pooled_PCA", "Rolling_Stability", "Global_vs_Local", "Rolling_Spillover", "RV_Residuals", "Flies", "Forwards",
                    "Cross_Market_Pairs", "Tenor_Percentiles", "Offshore_Liquidity", "Events", "Dashboard", "Ticker_Map", "Hist"]

class Book:
    def __init__(self, path, synthetic, curves):
        self.wb = xlsxwriter.Workbook(path)
        self.syn = synthetic
        self.REG, self.CELL = {}, {}
        self.na_count = 0
        self.nan_written = 0
        F = self.wb.add_format
        self.f = dict(
            title=F({"bold": True, "font_size": 14}), syn=F({"bold": True, "font_color": "#C00000"}),
            note=F({"italic": True, "font_color": "#595959"}), sec=F({"bold": True, "font_size": 12, "bg_color": "#D9E1F2"}),
            hdr=F({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1, "text_wrap": True, "valign": "vcenter"}),
            num=F({"num_format": "0.00", "bg_color": "#F2F2F2"}), num4=F({"num_format": "0.0000", "bg_color": "#F2F2F2"}),
            pct=F({"num_format": "0.0%", "bg_color": "#F2F2F2"}), int=F({"num_format": "0", "bg_color": "#F2F2F2"}),
            date=F({"num_format": "yyyy-mm-dd", "bg_color": "#F2F2F2"}), na=F({"font_color": "#7F7F7F", "italic": True, "bg_color": "#F2F2F2"}),
            fnum=F({"num_format": "0.00"}), fnum4=F({"num_format": "0.0000"}), fpct=F({"num_format": "0.0%"}), fint=F({"num_format": "0"}), ftxt=F({}),
            inp=F({"font_color": "#0000FF"}), inp_num=F({"font_color": "#0000FF", "num_format": "0.00"}), inp_num4=F({"font_color": "#0000FF", "num_format": "0.0000"}),
            inp_date=F({"font_color": "#0000FF", "num_format": "yyyy-mm-dd"}),
            wrap=F({"text_wrap": True, "valign": "top"}), bold=F({"bold": True}),
        )
        self.ws = OrderedDict()
        names = SHEET_ORDER_HEAD + [f"Curve_{c}" for c in curves] + SHEET_ORDER_TAIL
        for n in names:
            self.ws[n] = self.wb.add_worksheet(n)

    def head(self, name, title, notes=()):
        ws = self.ws[name]
        ws.write(0, 0, title, self.f["title"])
        ws.write(1, 0, SYNTH_LABEL if self.syn else "Run on user-supplied Bloomberg data", self.f["syn"] if self.syn else self.f["note"])
        r = 2
        for n in notes:
            ws.write(r, 0, n, self.f["note"])
            r += 1
        return ws, r + 1

    def sec(self, ws, r, text, width=8):
        for c in range(width):
            ws.write(r, c, text if c == 0 else "", self.f["sec"])
        return r + 1

    def put(self, ws, r, c, v, kind="num"):
        if v is None or (isinstance(v, (float, np.floating)) and not np.isfinite(v)):
            ws.write_string(r, c, "n/a", self.f["na"])
            self.na_count += 1
        elif isinstance(v, (pd.Timestamp, dt.datetime, dt.date, np.datetime64)):
            ws.write_datetime(r, c, pd.Timestamp(v).to_pydatetime(), self.f["date"])
        elif isinstance(v, (bool, np.bool_)):
            ws.write_string(r, c, "Y" if v else "N")
        elif isinstance(v, (int, float, np.integer, np.floating)):
            ws.write_number(r, c, float(v), self.f[kind if kind in self.f else "num"])
        else:
            ws.write_string(r, c, str(v))

    def table(self, ws, r0, c0, df, key=None, kinds=None, freeze=False):
        kinds = kinds or {}
        for j, col in enumerate(df.columns):
            ws.write(r0, c0 + j, str(col), self.f["hdr"])
        for i, row in enumerate(df.itertuples(index=False)):
            for j, v in enumerate(row):
                self.put(ws, r0 + 1 + i, c0 + j, v, kinds.get(df.columns[j], "num"))
        if freeze:
            ws.freeze_panes(r0 + 1, 0)
        info = dict(sheet=ws.get_name(), hdr=r0, r0=r0 + 1, r1=r0 + len(df), c0=c0, c1=c0 + len(df.columns) - 1, cols=list(df.columns))
        if key is not None:
            self.REG[key] = info
        return info

    def form(self, ws, r, c, formula, value, kind="fnum"):
        v = 0.0 if value is None or (isinstance(value, (float, np.floating)) and not np.isfinite(value)) else value
        ws.write_formula(r, c, formula, self.f[kind], v)

    def addr(self, ws, r, c, absolute=True):
        return f"'{ws.get_name()}'!{xl_rowcol_to_cell(r, c, row_abs=absolute, col_abs=absolute)}"

    def widths(self, ws, first=24, rest=13, n=40):
        ws.set_column(0, 0, first)
        ws.set_column(1, n, rest)

def stat_row(bk, ws, r, c, hist, name):
    """latest, mean, sd, z, percentile as live formulas on Hist columns (cached values from Python)."""
    st = hist.stats(name)
    bk.form(ws, r, c, hist.f_latest(name), st["latest"])
    bk.form(ws, r, c + 1, hist.f_mean(name), st["mean"])
    bk.form(ws, r, c + 2, hist.f_sd(name), st["sd"])
    bk.form(ws, r, c + 3, hist.f_z(name), st["z"])
    bk.form(ws, r, c + 4, hist.f_pct(name), st["pct"], "fpct")

def pca_rows(label, p):
    ve = list(p["ve"][:4]) + [np.nan] * max(0, 4 - len(p["ve"][:4]))
    return dict(Basis=label, PC1=ve[0], PC2=ve[1], PC3=ve[2], PC4=ve[3], Cum_PC1_3=float(np.sum(p["ve"][:3])),
                N_obs=p["n"], N_vars=p["N"])

# ---------------------------------------------------------------- README / Assumptions
def w_readme(bk, X):
    ws, r = bk.head("README", "EM Asia swap-rate PCA - Results workbook")
    L = [
        ("Purpose", "PCA framework on offshore interest-rate swap curves only (no bond yields, no onshore curves) for EM Asia + USD SOFR OIS benchmark."),
        ("Run date", str(X.run_date)),
        ("Data source", X.insp.get("source", "")),
        ("Sample", f"{X.start.date()} to {X.end.date()} (last {P('LAST_N_YEARS')} years). Full-history PCA variance is also reported on PCA_Variance."),
        ("Curves", ", ".join(X.res.keys())),
        ("Units", "All rates in bp. PCA is on DAILY CHANGES in bp (never levels), covariance matrix (main) with a correlation-matrix robustness run."),
        ("Par vs forwards", "Par swap rates are the main input for factor analysis. 1Y forwards (bootstrapped, same method for all curves) are secondary, for RV/flies/slopes/pairs. "
                            "Only 'supported' forwards (both bounding par pillars quoted) enter any PCA/fly/pair; block forwards are labelled; interpolated points are reference only."),
        ("Bootstrapping", BOOTSTRAP_METHOD),
        ("Static vs live", "Grey-shaded cells are static values written by Python (all PCA outputs, loadings, scores, regression results). White cells with black text are LIVE Excel formulas: "
                           "z-scores, percentiles, means, SDs, ranks, residual = actual - fitted, RV signal flags and the Dashboard links. Blue text = hard-coded inputs (Assumptions parameters, Hist data, Ticker_Map)."),
        ("Hist sheet", "Trailing 252-observation series (each column on its own curve calendar; last row = latest date) that the live z-score / percentile formulas read."),
    ]
    if X.synthetic:
        L.insert(1, ("WARNING", "SYNTHETIC DATA - ILLUSTRATIVE ONLY. Nothing in this workbook reflects real market data. Replace the input with your Bloomberg export and re-run build_all.py."))
    for k, v in L:
        ws.write(r, 0, k, bk.f["bold"])
        ws.write(r, 1, v, bk.f["wrap"])
        r += 1
    r += 1
    r = bk.sec(ws, r, "Built-in tests (run before writing)", 2)
    for t in X.tests:
        ws.write(r, 0, t["Result"], bk.f["bold"])
        ws.write(r, 1, t["Test"] + " - " + t["Detail"])
        r += 1
    ws.write(r, 0, "INFO", bk.f["bold"])
    ws.write(r, 1, f"{bk.na_count} cells carry a documented 'n/a' placeholder (e.g. carry where floating fixings are not supplied); no NaN / inf is written to any cell.")
    r += 2
    r = bk.sec(ws, r, "Sheets", 2)
    desc = [("Assumptions", "Parameters, curve conventions, tenor sets, assumptions to double check"), ("Data_QC", "Data inspection, coverage, every removal / flag (units, sources, fill rows, bad prints, roll breaks, forward consistency)"),
            ("Curve_Build", "Bootstrapping method, re-pricing errors, supported / block / interpolated forward points"), ("PCA_Variance", "Variance explained PC1-PC4 (par main, forwards comparison, correlation and full-history variants)"),
            ("PCA_Loadings", "Loadings by tenor / forward point"), ("PCA_Scores", "Daily PC scores per curve"), ("Curve_<id>", "One sheet per individual curve: variance, loadings, latest scores, fit residuals, rolling stability"),
            ("Pooled_PCA", "Pooled EM Asia PCA with / without SOFR and tie-back to individual PCs"), ("Rolling_Stability", "Rolling 1Y PCA variance and congruence to full-sample PCs"),
            ("Global_vs_Local", "Spillover regressions on SOFR PCs, three timing versions, ranking, decoupling"), ("Rolling_Spillover", "Rolling 6M / 12M beta and R2"),
            ("RV_Residuals", "3-PC fitted vs actual, residuals, z-scores, half-lives, signals"), ("Flies", "PCA-neutral butterflies on supported 1Y forwards"),
            ("Forwards", "Slopes, forward-forward spreads, block forwards, interpolated points (reference only)"), ("Cross_Market_Pairs", "PC1-neutral cross-market pairs, top 10"),
            ("Tenor_Percentiles", "Per-tenor percentile and dislocation of rate levels over 1y / 3y / 5y"), ("Offshore_Liquidity", "Stale / zero-change stats, PCA with vs without illiquid tenors, SOFR / FX cross-check"), ("Events", "Central-bank event analysis (template if no calendar supplied)"),
            ("Dashboard", "One row per curve"), ("Ticker_Map", "Template for your Bloomberg tickers (not filled - no tickers are guessed)"), ("Hist", "Series behind the live formulas")]
    for k, v in desc:
        ws.write(r, 0, k, bk.f["bold"])
        ws.write(r, 1, v)
        r += 1
    r += 1
    r = bk.sec(ws, r, "Caveats", 2)
    cav = ["Offshore curves only; no onshore comparison is included or invented.",
           "Carry of spot-starting legs and pair carry are n/a: floating-index fixings / funding are not in a par-rate export. Only 1M roll-down is reported.",
           "Day counts / floating indices marked 'assumed' or 'TO CONFIRM' on Assumptions are placeholders; they only affect the bootstrap accrual factor.",
           "Bootstrapping uses annual fixed payments on an integer-year grid; non-integer tenors (e.g. 6M) stay in par PCA but are not bootstrapped.",
           "Forward-starting 1Y forwards inherit any error in the two par pillars; check Data_QC forward-consistency flags before trading a signal.",
           "PC scores are daily changes; percentiles / z-scores are trailing 252 observations. 20-day cumulative scores are shown because a single-day score is noisy.",
           "Residual level (RV) is the cumulated 3-PC fit error since the sample start; only mean-reverting, stable, half-life <= limit residuals are flagged as signals."]
    for t in cav:
        ws.write(r, 0, "-")
        ws.write(r, 1, t, bk.f["wrap"])
        r += 1
    ws.set_column(0, 0, 22)
    ws.set_column(1, 1, 130)

def w_assumptions(bk, X):
    ws, r = bk.head("Assumptions", "Assumptions and tunable parameters", ["Blue = editable input in build_all.py CONFIG block; re-run to apply. Values here are echoed, not read back."])
    df = pd.DataFrame([dict(Parameter=k, Value=str(v[0]) if not isinstance(v[0], (int, float)) else v[0], Description=v[1]) for k, v in PARAMS.items()])
    ws.write(r, 0, "Parameter", bk.f["hdr"]); ws.write(r, 1, "Value", bk.f["hdr"]); ws.write(r, 2, "Description", bk.f["hdr"])
    ws.freeze_panes(r + 1, 0)
    for i, row in df.iterrows():
        ws.write(r + 1 + i, 0, row["Parameter"])
        (ws.write_number if isinstance(row["Value"], (int, float)) else ws.write_string)(r + 1 + i, 1, row["Value"], bk.f["inp"])
        ws.write(r + 1 + i, 2, row["Description"])
    r += len(df) + 3
    r = bk.sec(ws, r, "Curve conventions (from the brief; anything not stated is TO CONFIRM, never guessed)", 9)
    cv = pd.DataFrame([dict(Curve=c, Market=m["market"], Description=m["label"], FloatingIndex=m["floating_index"], Settlement=m["settlement"], FixingSource=m["fixing_source"],
                            DayCount=m["day_count"], AccrualFactor=m["dc_factor"], AlternativeCurve=m["alternative"]) for c, m in CURVES.items()])
    bk.table(ws, r, 0, cv, kinds={"AccrualFactor": "num4"})
    r += len(cv) + 3
    r = bk.sec(ws, r, "Tenor sets per curve", 9)
    rows = []
    for c, C in X.C.items():
        B = X.Bd.get(c)
        rows.append(dict(Curve=c, Status=C["status"], Quoted=", ".join(tenor_label(t) for t in sorted(C["tstats"]["T"])),
                         Retained=", ".join(tenor_label(t) for t in C["keep"]), Dropped=", ".join(C["dropped"]) or "none",
                         Supported_1Y_fwd=", ".join(B["supported"].columns) if B else "n/a", Block_fwd=", ".join(B["blocks"].columns) if B else "n/a",
                         Interpolated_reference_only=", ".join(B["interp"].columns) if B else "n/a"))
    bk.table(ws, r, 0, pd.DataFrame(rows))
    r += len(rows) + 3
    r = bk.sec(ws, r, "Assumptions to double check", 9)
    A = [f"Offshore curve chosen per market: see 'Curve conventions'. Unconfirmed markets use the most liquid offshore curve; alternatives listed in AlternativeCurve.",
         f"Lag choice for global vs local: '{X.gl['choice']}' (highest median R2 across markets; medians same/lag/avg = " + ", ".join(f"{k}:{v:.3f}" for k, v in X.gl["median_wr2"].items()) + ").",
         "Timing rule: 'lag' regresses a local return on SOFR PC scores from the last US close before the local close, aggregated over each local return interval; 'same' uses same-calendar-date US scores.",
         "Illiquidity rule: tenor dropped if > 30% of observed days are zero-change (or coverage < 50%); flagged if > 15%.",
         "Stale fixings forward-filled for at most 2 consecutive days (flagged). Longer gaps drop the date.",
         "Bad print: > 6 robust SD change reversed next day (|d1+d2| < 0.3|d1|). Fill rows: all tenors identical to prior row (or weekend) -> treated as missing.",
         "PCA on covariance of daily bp changes (main); correlation-matrix run as robustness. PC1 positive on all tenors, PC2 long-minus-short positive, PC3 belly-minus-wings positive.",
         "Fly weights neutralise PC1 and PC2 of the forward-curve PCA (supported 1Y forwards + labelled block forwards); belly weight = -1. DV01 of a forward-starting 1Y swap = accrual x D(end) x 1bp per unit notional.",
         "Roll-down (1M) = expected change in a forward / par rate when its start / maturity rolls one month closer, linear between adjacent points, curve unchanged. Carry is n/a (no floating fixings).",
         "z-scores / percentiles use the trailing 252 observations including the latest; percentile = PERCENTRANK.INC.",
         "Cross-market pairs: EM-vs-EM only; hedge ratio = PC1 loading of A / PC1 loading of B; spread = A - ratio x B in bp.",
         "Half-life from AR(1) on the last 504 observations of the residual level; unstable if phi<=0 or >=1, or the two half-samples' half-lives differ by more than 3x."]
    if X.synthetic:
        A.insert(0, "SYNTHETIC RUN: data are simulated with a known global/regional/local factor structure and deliberate data errors; findings are illustrative of the pipeline only.")
    for a in A:
        ws.write(r, 0, "-")
        ws.write(r, 1, a)
        r += 1
    ws.set_column(0, 0, 30); ws.set_column(1, 1, 40); ws.set_column(2, 2, 70); ws.set_column(3, 12, 26)

# ---------------------------------------------------------------- Data_QC / Curve_Build
def w_dataqc(bk, X):
    ws, r = bk.head("Data_QC", "Data quality log", ["Units, pricing source, fill rows, bad prints, stale fixings, roll breaks and forward-consistency flags. Every removal is logged."])
    r = bk.sec(ws, r, "A. Data inspection", 8)
    ins = X.insp
    for k in ("source", "rows", "date_min", "date_max", "columns", "duplicates", "bad_dates", "bad_rates"):
        ws.write(r, 0, k, bk.f["bold"])
        v = ins.get(k)
        ws.write(r, 1, str(v.date()) if isinstance(v, pd.Timestamp) else str(v))
        r += 1
    r += 1
    bk.table(ws, r, 0, ins["by_market"], key="qc_inspect", kinds={"Rows": "int"})
    ws.freeze_panes(r + 1, 0)
    r += len(ins["by_market"]) + 3
    r = bk.sec(ws, r, "B. Coverage and liquidity per curve and tenor", 12)
    ts = pd.concat([C["tstats"].drop(columns="T") for C in X.C.values()], ignore_index=True)
    bk.table(ws, r, 0, ts, kinds={"Quotes": "int", "Coverage": "pct", "ZeroChangeShare": "pct", "StaleRuns_gt2d": "int", "LongestUnchangedRun": "int", "DaysInStaleRuns": "int", "FilledDays": "int"})
    r += len(ts) + 3
    r = bk.sec(ws, r, "C. Summary of checks by curve", 8)
    q = X.qc.df()
    summ = q.groupby(["Curve", "Check"]).size().reset_index(name="Count")
    summ["Check"] = summ["Check"].str.slice(0, 120)
    bk.table(ws, r, 0, summ, kinds={"Count": "int"})
    r += len(summ) + 3
    r = bk.sec(ws, r, "D. Full log (every removal, flag and check)", 8)
    q2 = q.copy()
    q2["Date"] = q2["Date"].astype(str)
    q2["Value"] = q2["Value"].map(lambda v: v if isinstance(v, (int, float, np.integer, np.floating)) and not (isinstance(v, float) and not np.isfinite(v)) else (str(v) if v != "" else ""))
    bk.table(ws, r, 0, q2, kinds={"Value": "num"})
    ws.set_column(0, 0, 20); ws.set_column(1, 1, 60); ws.set_column(2, 3, 14); ws.set_column(4, 4, 80); ws.set_column(5, 12, 14)

def w_curvebuild(bk, X):
    ws, r = bk.head("Curve_Build", "Curve construction (bootstrapping) and forward-point support",
                    ["Same method for every curve. Par pillars re-priced from the bootstrapped curve; max error must be < %.2f bp." % P("REPRICE_TOL_BP")])
    rows = []
    for c, B in X.Bd.items():
        rows.append(dict(Curve=c, Quoted_pillars=", ".join(f"{p}Y" for p in B["pillars"]), Accrual_factor=B["acc"], Max_reprice_error_bp=B["max_err"],
                         Check=("PASS" if B["max_err"] < P("REPRICE_TOL_BP") else "FAIL"), Supported_1Y_fwd=", ".join(B["supported"].columns),
                         Block_fwd=", ".join(f"{k} ({a}Y->{b}Y)" for k, (a, b, w) in B["block_meta"].items()) or "none",
                         Interpolated_reference_only=", ".join(B["interp"].columns) or "none"))
    df = pd.DataFrame(rows)
    ws.write(r, 0, "Method: " + BOOTSTRAP_METHOD, bk.f["wrap"])
    r += 2
    bk.table(ws, r, 0, df, key="build", kinds={"Accrual_factor": "num4", "Max_reprice_error_bp": "num4"}, freeze=True)
    r += len(df) + 3
    r = bk.sec(ws, r, "Re-pricing error by pillar (bp, max over dates)", 8)
    rows = [dict(Curve=c, **{f"{p}Y": e for p, e in B["reprice_err"].items()}) for c, B in X.Bd.items()]
    df2 = pd.DataFrame(rows)
    bk.table(ws, r, 0, df2, kinds={c: "num4" for c in df2.columns})
    r += len(df2) + 3
    r = bk.sec(ws, r, "Interpolated forward points - REFERENCE ONLY (not used in any PCA, fly or pair)", 8)
    ip = pd.concat([X.fwdtabs[c][2] for c in X.fwdtabs if len(X.fwdtabs[c][2])], ignore_index=True)
    bk.table(ws, r, 0, ip if len(ip) else pd.DataFrame({"Note": ["none"]}))
    r += max(len(ip), 1) + 3
    r = bk.sec(ws, r, "Comparison with Bloomberg forward quotes (bp, derived - Bloomberg)", 8)
    if X.bbgc is None:
        ws.write(r, 0, "NOT PROVIDED - no Bloomberg forward quotes supplied (see bbg_forwards_template.csv). Derived forwards are used for all analysis.")
    else:
        bk.table(ws, r, 0, X.bbgc)
    ws.set_column(0, 0, 20); ws.set_column(1, 12, 26)

# ---------------------------------------------------------------- PCA variance / loadings / scores
def w_pca_variance(bk, X):
    ws, r = bk.head("PCA_Variance", "Variance explained by principal component (share of total variance)",
                    ["MAIN: PCA on daily par-rate changes (bp), covariance matrix. Comparison: supported 1Y forwards. Robustness: correlation (vol-standardised) matrix.",
                     "Shares of total variance (0-100%). PC labels only where the loading shape supports them."])
    main = []
    for c, R in X.res.items():
        p = R["par_cov"]
        main.append(dict(Curve=c, PC1=p["ve"][0], PC2=p["ve"][1], PC3=p["ve"][2], PC4=(p["ve"][3] if len(p["ve"]) > 3 else np.nan),
                         Cum_PC1_3=float(p["ve"][:3].sum()), N_obs=p["n"], N_tenors=p["N"], PC1_shape=R["labels"][0], PC2_shape=R["labels"][1], PC3_shape=R["labels"][2]))
    dfm = pd.DataFrame(main)
    r = bk.sec(ws, r, "1. MAIN - par rates, covariance, analysis sample", 11)
    bk.table(ws, r, 0, dfm, key="var_main", kinds={k: "pct" for k in ("PC1", "PC2", "PC3", "PC4", "Cum_PC1_3")} | {"N_obs": "int", "N_tenors": "int"}, freeze=True)
    r += len(dfm) + 3
    r = bk.sec(ws, r, "2. All variants", 11)
    rows = []
    for c, R in X.res.items():
        for lab, p in (("Par - covariance (MAIN)", R["par_cov"]), ("Par - correlation", R["par_corr"]), ("Forwards - covariance (comparison)", R["fwd_cov"]),
                       ("Forwards - correlation", R["fwd_corr"]), ("Par - covariance, FULL HISTORY", R["par_cov_full"]),
                       ("Par - covariance incl. illiquid tenors", R["par_incl"])):
            if p is not None:
                rows.append(dict(Curve=c, **pca_rows(lab, p)))
    dfa = pd.DataFrame(rows)
    bk.table(ws, r, 0, dfa, kinds={k: "pct" for k in ("PC1", "PC2", "PC3", "PC4", "Cum_PC1_3")} | {"N_obs": "int", "N_vars": "int"})
    r += len(dfa) + 3
    r = bk.sec(ws, r, "3. Par vs forwards: change in variance explained (forwards minus par, percentage points)", 11)
    rows = []
    for c, R in X.res.items():
        if R["fwd_cov"] is None:
            rows.append(dict(Curve=c, Note="forward PCA not available (fewer than 3 supported forward points)"))
            continue
        pv, fv = R["par_cov"]["ve"], R["fwd_cov"]["ve"]
        rows.append(dict(Curve=c, dPC1_pp=100 * (fv[0] - pv[0]), dPC2_pp=100 * (fv[1] - pv[1]), dPC3_pp=100 * (fv[2] - pv[2]),
                         Par_PC1_3=float(pv[:3].sum()), Fwd_PC1_3=float(fv[:3].sum()), Note="loadings are on different variables (tenors vs forward points): compare shapes on PCA_Loadings"))
    dfc = pd.DataFrame(rows).fillna("")
    bk.table(ws, r, 0, dfc, kinds={"Par_PC1_3": "pct", "Fwd_PC1_3": "pct"})
    ws.set_column(0, 0, 24); ws.set_column(1, 12, 14); ws.set_column(9, 11, 26)

def w_pca_loadings(bk, X):
    ws, r = bk.head("PCA_Loadings", "Loadings by tenor / forward point (unit-length eigenvectors)",
                    ["Signs normalised: PC1 positive on all tenors; PC2 long-minus-short positive; PC3 belly-minus-wings positive.",
                     "Par covariance loadings are in bp-per-unit-score; correlation loadings apply to vol-standardised changes."])
    rows = []
    for c, R in X.res.items():
        for basis, p, names in (("Par - covariance", R["par_cov"], R["tenors"]), ("Par - correlation", R["par_corr"], R["tenors"]),
                                ("Forwards - covariance", R["fwd_cov"], list(R["sup"].columns) if R["fwd_cov"] is not None else None)):
            if p is None:
                continue
            for i, nm in enumerate(names):
                row = dict(Curve=c, Basis=basis, Point=nm)
                for k in range(4):
                    row[f"PC{k + 1}"] = p["V"][i, k] if k < p["V"].shape[1] else np.nan
                rows.append(row)
    df = pd.DataFrame(rows)
    info = bk.table(ws, r, 0, df, key="loadings_all", kinds={f"PC{k}": "num4" for k in range(1, 5)}, freeze=True)
    # per (curve, basis) contiguous blocks for chart linking
    i = 0
    for (c, basis), g in df.groupby(["Curve", "Basis"], sort=False):
        bk.REG[("load", c, basis)] = dict(sheet="PCA_Loadings", hdr=info["hdr"], r0=info["r0"] + g.index[0], r1=info["r0"] + g.index[-1], c0=info["c0"], c1=info["c1"], cols=info["cols"])
    ws.set_column(0, 2, 24); ws.set_column(3, 8, 12)

def w_pca_scores(bk, X):
    ws, r = bk.head("PCA_Scores", "Daily PC scores per curve (bp; score = centred daily change projected on the loading vector)",
                    ["Each curve block has its own trading calendar. PC1 level / PC2 slope / PC3 curvature only where PCA_Variance shape flags support the label."])
    c0 = 0
    for c, R in X.res.items():
        df = R["scores"].reset_index()
        df.columns = ["Date"] + [f"{c} {x}" for x in R["scores"].columns]
        bk.table(ws, r, c0, df, key=("scores", c), kinds={col: "num4" for col in df.columns[1:]})
        c0 += df.shape[1] + 1
    ws.freeze_panes(r + 1, 0)
    ws.set_column(0, c0, 12)

# ---------------------------------------------------------------- Individual curve sheets
def w_curve(bk, X, cid):
    R, meta = X.res[cid], CURVES[cid]
    ws, r = bk.head(f"Curve_{cid}", f"Individual curve PCA - {meta['label']}",
                    [f"Floating index: {meta['floating_index']} | Settlement: {meta['settlement']} | Fixing source: {meta['fixing_source']} | Day count: {meta['day_count']}",
                     "Standalone PCA, separate from pooled analysis. Grey = static Python values; white = live formulas on the Hist sheet."])
    r = bk.sec(ws, r, "A. Variance explained", 8)
    rows = [pca_rows("Par - covariance (MAIN)", R["par_cov"]), pca_rows("Par - correlation", R["par_corr"])]
    if R["fwd_cov"] is not None:
        rows += [pca_rows("Forwards - covariance (comparison)", R["fwd_cov"]), pca_rows("Forwards - correlation", R["fwd_corr"])]
    rows.append(pca_rows("Par - covariance, FULL HISTORY", R["par_cov_full"]))
    if R["par_incl"] is not None:
        rows.append(pca_rows("Par - covariance incl. illiquid tenors", R["par_incl"]))
    df = pd.DataFrame(rows)
    bk.table(ws, r, 0, df, key=("cv_var", cid), kinds={k: "pct" for k in ("PC1", "PC2", "PC3", "PC4", "Cum_PC1_3")} | {"N_obs": "int", "N_vars": "int"}, freeze=True)
    r += len(df) + 2
    ws.write(r, 0, "PC shapes: " + " | ".join(f"PC{i + 1} = {l}" for i, l in enumerate(R["labels"])), bk.f["note"])
    r += 2
    r = bk.sec(ws, r, "B. Loadings - par rates (covariance)", 8)
    L = pd.DataFrame(R["par_cov"]["V"][:, :4], columns=[f"PC{k + 1}" for k in range(min(4, R["par_cov"]["V"].shape[1]))])
    L.insert(0, "Tenor", R["tenors"])
    L.insert(1, "Tenor_yrs", [parse_tenor(t) for t in R["tenors"]])
    bk.table(ws, r, 0, L, key=("cv_load", cid), kinds={c: "num4" for c in L.columns[2:]} | {"Tenor_yrs": "num"})
    r += len(L) + 2
    if R["fwd_cov"] is not None:
        r = bk.sec(ws, r, "B2. Loadings - supported 1Y forwards (covariance, comparison)", 8)
        L2 = pd.DataFrame(R["fwd_cov"]["V"][:, :4], columns=[f"PC{k + 1}" for k in range(min(4, R["fwd_cov"]["V"].shape[1]))])
        L2.insert(0, "Point", list(R["sup"].columns))
        bk.table(ws, r, 0, L2, kinds={c: "num4" for c in L2.columns[1:]})
        r += len(L2) + 2
    r = bk.sec(ws, r, "C. Latest PC scores (bp), 1-year percentile and z-score", 12)
    hdr = ["Factor", "Shape", "Daily score (bp)", "1y mean", "1y SD", "1y z", "1y pctile", "20d cum (bp)", "20d mean", "20d SD", "20d z", "20d pctile"]
    for j, h in enumerate(hdr):
        ws.write(r, j, h, bk.f["hdr"])
    for k, (nm, d) in enumerate(R["latest_scores"].items()):
        ws.write(r + 1 + k, 0, nm)
        ws.write(r + 1 + k, 1, R["labels"][k])
        stat_row(bk, ws, r + 1 + k, 2, X.hist, f"{cid}|{d['col']}|daily")
        stat_row(bk, ws, r + 1 + k, 7, X.hist, f"{cid}|{d['col']}|cum20")
        bk.CELL[("score_daily", cid, k)] = bk.addr(ws, r + 1 + k, 2)
        bk.CELL[("score_daily_pct", cid, k)] = bk.addr(ws, r + 1 + k, 6)
        bk.CELL[("score_cum_pct", cid, k)] = bk.addr(ws, r + 1 + k, 11)
        bk.CELL[("score_cum_z", cid, k)] = bk.addr(ws, r + 1 + k, 10)
        bk.CELL[("score_cum", cid, k)] = bk.addr(ws, r + 1 + k, 7)
    bk.REG[("cv_scores", cid)] = dict(sheet=ws.get_name(), hdr=r, r0=r + 1, r1=r + len(R["latest_scores"]), c0=0, c1=11, cols=hdr)
    r += len(R["latest_scores"]) + 3
    r = bk.sec(ws, r, "D. 3-PC fit residual by tenor (bp) - what PC1-3 do not explain", 8)
    D = pd.DataFrame({"Tenor": R["tenors"], "Residual SD (bp)": R["fit_resid_sd"].values, "Latest-day residual (bp)": R["fit_resid_last"].values})
    bk.table(ws, r, 0, D)
    r += len(D) + 1
    ws.write(r, 0, f"3-PC variance explained (par, cov): {R['fit_r2_3pc']:.1%}", bk.f["note"])
    r += 3
    r = bk.sec(ws, r, "E. Rolling stability (1Y windows, 1M step)", 8)
    ro = R["rolling"]
    brk = ro[ro["Break_PC1"]]
    S = pd.DataFrame([dict(Windows=len(ro), Mean_PC1_VE=ro["VE_PC1"].mean(), Min_PC1_VE=ro["VE_PC1"].min(), Max_PC1_VE=ro["VE_PC1"].max(),
                           Min_Cong_PC1=ro["Cong_PC1"].min(), Min_Cong_PC2=ro["Cong_PC2"].min(), Min_Cong_PC3=ro["Cong_PC3"].min(),
                           Breaks_PC1=int(ro["Break_PC1"].sum()), Breaks_PC2=int(ro["Break_PC2"].sum()), Breaks_PC3=int(ro["Break_PC3"].sum()),
                           PC1_break_windows_ending=(", ".join(str(d.date()) for d in brk.index[:6]) + (" ..." if len(brk) > 6 else "")) if len(brk) else "none")])
    bk.table(ws, r, 0, S, kinds={"Windows": "int", "Breaks_PC1": "int", "Breaks_PC2": "int", "Breaks_PC3": "int", "Mean_PC1_VE": "pct", "Min_PC1_VE": "pct", "Max_PC1_VE": "pct", "Min_Cong_PC1": "num4", "Min_Cong_PC2": "num4", "Min_Cong_PC3": "num4"})
    r += 2
    ws.write(r, 0, "Congruence = |cosine| of rolling vs full-sample loading vector. Low PC2/PC3 congruence usually means near-equal eigenvalues rotating, not a structural break; PC1 breaks are the headline.", bk.f["note"])
    ws.set_column(0, 0, 34); ws.set_column(1, 12, 15)

# ---------------------------------------------------------------- Pooled
def w_pooled(bk, X):
    ws, r = bk.head("Pooled_PCA", "Pooled EM Asia PCA (same tenor stacked across markets) with and without SOFR",
                    ["Variables = markets; observations = common dates; daily changes in bp; covariance matrix. Markets lacking the tenor / supported forward point are excluded (listed).",
                     "Forward sets (1y1y, 3y1y) are comparison only."])
    rows = []
    for (nm, var), d in X.pooled.items():
        p = d["res"]
        if p is None:
            rows.append(dict(Set=nm, Variant=var, N_obs="n/a", N_markets="n/a", Excluded=", ".join(d["excluded"]) or "none", PC1="n/a", PC2="n/a", PC3="n/a", PC4="n/a",
                             PC1_corr_matrix="n/a"))
            continue
        ve = list(p["p"]["ve"][:4]) + [np.nan] * 3
        rows.append(dict(Set=nm, Variant=var, N_obs=p["n"], N_markets=len(p["names"]), Excluded=", ".join(d["excluded"]) or "none", PC1=ve[0], PC2=ve[1], PC3=ve[2], PC4=ve[3],
                         PC1_corr_matrix=float(p["pc"]["ve"][0])))
    df = pd.DataFrame(rows)
    r = bk.sec(ws, r, "1. Variance explained by pooled PCs", 10)
    bk.table(ws, r, 0, df, key="pool_sum", kinds={k: "pct" for k in ("PC1", "PC2", "PC3", "PC4", "PC1_corr_matrix")} | {"N_obs": "int", "N_markets": "int"}, freeze=True)
    r += len(df) + 3
    r = bk.sec(ws, r, "2. With vs without SOFR - change in pooled PC1 variance", 10)
    cmp_ = []
    for nm in dict.fromkeys(k[0] for k in X.pooled):
        a, b = X.pooled[(nm, "ex-SOFR")]["res"], X.pooled[(nm, "incl. SOFR")]["res"]
        if a is None or b is None:
            continue
        cmp_.append(dict(Set=nm, PC1_ex_SOFR=a["p"]["ve"][0], PC1_incl_SOFR=b["p"]["ve"][0], Diff_pp=100 * (b["p"]["ve"][0] - a["p"]["ve"][0]),
                         PC2_ex_SOFR=a["p"]["ve"][1], PC2_incl_SOFR=b["p"]["ve"][1], SOFR_loading_on_PC1_incl=b["p"]["V"][b["names"].index(GLOBAL_CURVE), 0] if GLOBAL_CURVE in b["names"] else np.nan,
                         SOFR_corr_with_PC1_incl=b["corr1"].get(GLOBAL_CURVE, np.nan)))
    bk.table(ws, r, 0, pd.DataFrame(cmp_), kinds={"PC1_ex_SOFR": "pct", "PC1_incl_SOFR": "pct", "PC2_ex_SOFR": "pct", "PC2_incl_SOFR": "pct", "SOFR_loading_on_PC1_incl": "num4", "SOFR_corr_with_PC1_incl": "num4"})
    r += len(cmp_) + 3
    r = bk.sec(ws, r, "3. Market loadings on pooled PCs and correlation with pooled PC1", 10)
    for (nm, var), d in X.pooled.items():
        p = d["res"]
        if p is None:
            continue
        ws.write(r, 0, f"{nm} - {var}   (N={p['n']}, markets={len(p['names'])})", bk.f["bold"])
        r += 1
        L = pd.DataFrame({"Market": p["names"], "PC1_loading": p["p"]["V"][:, 0], "PC2_loading": p["p"]["V"][:, 1] if p["p"]["V"].shape[1] > 1 else np.nan,
                          "PC3_loading": p["p"]["V"][:, 2] if p["p"]["V"].shape[1] > 2 else np.nan, "Corr_with_pooled_PC1": [p["corr1"][m] for m in p["names"]]})
        bk.table(ws, r, 0, L, key=("pool_load", nm, var), kinds={c: "num4" for c in L.columns[1:]})
        r += len(L) + 2
    r = bk.sec(ws, r, "4. Tie-back: share of each market's own PC1-3 explained by pooled EM Asia PCs (leave-one-out, par 5Y, PC1-2) vs SOFR PCs (PC1-2)", 10)
    tb = X.tie.copy()
    bk.table(ws, r, 0, tb, key="tieback", kinds={c: "pct" for c in tb.columns if "R2" in c} | {"N": "int"})
    r += len(tb) + 1
    ws.write(r, 0, "R2 of the market's own PC score changes (measured over common-date intervals) on the regressors named. 'Both' = EM pool + SOFR together.", bk.f["note"])
    ws.set_column(0, 0, 24); ws.set_column(1, 14, 16); ws.set_column(4, 4, 30)

# ---------------------------------------------------------------- Rolling stability / global vs local / spillover
def w_rolling(bk, X):
    ws, r = bk.head("Rolling_Stability", "Rolling 1Y PCA stability (window %d obs, step %d obs)" % (P("ROLL_WIN"), P("ROLL_STEP")),
                    ["Congruence coefficient = |cosine| of the rolling PC loading vs the full-sample PC loading. Break flagged when below %.2f." % P("CONGR_BREAK"),
                     "PC1 breaks are the headline; PC2/PC3 congruence dips usually reflect near-equal eigenvalues rotating (few tenors), so they are counted separately."])
    rows = []
    for c, R in X.res.items():
        ro = R["rolling"]
        b = ro[ro["Break_PC1"]]
        rows.append(dict(Curve=c, Windows=len(ro), Mean_PC1_VE=ro["VE_PC1"].mean(), Min_PC1_VE=ro["VE_PC1"].min(), Max_PC1_VE=ro["VE_PC1"].max(),
                         Min_Cong_PC1=ro["Cong_PC1"].min(), Min_Cong_PC2=ro["Cong_PC2"].min(), Min_Cong_PC3=ro["Cong_PC3"].min(),
                         Breaks_PC1=int(ro["Break_PC1"].sum()), Breaks_PC2=int(ro["Break_PC2"].sum()), Breaks_PC3=int(ro["Break_PC3"].sum()),
                         PC1_break_windows_ending=", ".join(str(d.date()) for d in b.index[:5]) + (" ..." if len(b) > 5 else "") if len(b) else "none"))
    df = pd.DataFrame(rows)
    r = bk.sec(ws, r, "1. Summary and flagged breaks", 12)
    bk.table(ws, r, 0, df, kinds={"Windows": "int", "Breaks_PC1": "int", "Breaks_PC2": "int", "Breaks_PC3": "int", "Mean_PC1_VE": "pct", "Min_PC1_VE": "pct", "Max_PC1_VE": "pct", "Min_Cong_PC1": "num4", "Min_Cong_PC2": "num4", "Min_Cong_PC3": "num4"}, freeze=True)
    r += len(df) + 3
    r = bk.sec(ws, r, "2. PC1 variance explained by month (window-end month; months where every curve has a full window) - feeds chart 3", 10)
    wide = pd.DataFrame({c: R["rolling"]["VE_PC1"].groupby(R["rolling"].index.to_period("M")).last() for c, R in X.res.items()})
    wide.index = wide.index.astype(str)
    wide = wide.ffill().dropna(how="any").reset_index().rename(columns={"index": "Month"})
    bk.table(ws, r, 0, wide, key="roll_ve1", kinds={c: "pct" for c in wide.columns[1:]})
    r += len(wide) + 3
    r = bk.sec(ws, r, "3. Full rolling detail", 10)
    rows = []
    for c, R in X.res.items():
        t = R["rolling"].reset_index()
        t.insert(1, "Curve", c)
        rows.append(t)
    det = pd.concat(rows, ignore_index=True)
    bk.table(ws, r, 0, det, kinds={"VE_PC1": "pct", "VE_PC2": "pct", "VE_PC3": "pct", "Cong_PC1": "num4", "Cong_PC2": "num4", "Cong_PC3": "num4"})
    ws.set_column(0, 0, 20); ws.set_column(1, 12, 14); ws.set_column(10, 10, 60)

def w_gl(bk, X):
    gl = X.gl
    ws, r = bk.head("Global_vs_Local", "Global (SOFR) vs local decomposition",
                    [f"Local PC1-3 scores regressed on SOFR PC1-3 scores. Timing versions: same (same-day US), lag (previous US close), avg. CHOSEN: '{gl['choice']}' - highest median weighted R2 across markets.",
                     "Weighted R2 = sum(local PC variance share x R2 of that PC) / sum(shares of PC1-3) = share of the curve's PC1-3 variance driven by SOFR factors. Idiosyncratic = 1 - weighted R2."])
    r = bk.sec(ws, r, "1. Timing versions (weighted R2)", 8)
    rows = []
    for cid, per in gl["per"].items():
        best = max(per, key=lambda m: per[m]["wr2"])
        rows.append(dict(Curve=cid, Same_day=per["same"]["wr2"], Prev_US_close=per["lag"]["wr2"], Average=per["avg"]["wr2"], Best_for_this_market=best))
    rows.append(dict(Curve="MEDIAN (selects overall choice)", Same_day=gl["median_wr2"]["same"], Prev_US_close=gl["median_wr2"]["lag"], Average=gl["median_wr2"]["avg"], Best_for_this_market=gl["choice"]))
    bk.table(ws, r, 0, pd.DataFrame(rows), kinds={"Same_day": "pct", "Prev_US_close": "pct", "Average": "pct"}, freeze=True)
    r += len(rows) + 3
    r = bk.sec(ws, r, f"2. Spillover betas, R2 and residual - chosen timing ({gl['choice']})", 14)
    rows = []
    for cid, per in gl["per"].items():
        g = per[gl["choice"]]
        rg = g["regs"]
        rows.append(dict(Curve=cid, Beta_L1_on_G1=rg["PC1"]["b"][0], t_L1_on_G1=rg["PC1"]["b"][0] / rg["PC1"]["se"][0] if rg["PC1"]["se"][0] > 0 else np.nan,
                         Beta_L2_on_G2=rg["PC2"]["b"][1], Beta_L3_on_G3=rg["PC3"]["b"][2], R2_PC1=rg["PC1"]["r2"], R2_PC2=rg["PC2"]["r2"], R2_PC3=rg["PC3"]["r2"],
                         Weighted_R2=g["wr2"], Idiosyncratic_share=1 - g["wr2"], Resid_SD_PC1_bp=rg["PC1"]["resid_sd"], Local_PC1_SD_bp=float(X.res[cid]["scores"]["PC1"].std()),
                         Current_state=X.decoup[cid][1]))
    df = pd.DataFrame(rows)
    info = bk.table(ws, r, 0, df, key="gl_main", kinds={"Beta_L1_on_G1": "num4", "t_L1_on_G1": "num", "Beta_L2_on_G2": "num4", "Beta_L3_on_G3": "num4", "R2_PC1": "pct", "R2_PC2": "pct", "R2_PC3": "pct",
                                                         "Weighted_R2": "pct", "Idiosyncratic_share": "pct"})
    wc = info["cols"].index("Weighted_R2")
    rc = len(info["cols"])
    ws.write(info["hdr"], rc, "Rank (1 = most global)", bk.f["hdr"])
    rng = f"{xl_rowcol_to_cell(info['r0'], wc, True, True)}:{xl_rowcol_to_cell(info['r1'], wc, True, True)}"
    order = df["Weighted_R2"].rank(ascending=False, method="min")
    for i in range(len(df)):
        bk.form(ws, info["r0"] + i, rc, f"=RANK({xl_rowcol_to_cell(info['r0'] + i, wc)},{rng},0)", order.iloc[i], "fint")
        bk.CELL[("gl_beta", df["Curve"].iloc[i])] = bk.addr(ws, info["r0"] + i, 1)
        bk.CELL[("gl_wr2", df["Curve"].iloc[i])] = bk.addr(ws, info["r0"] + i, wc)
        bk.CELL[("gl_state", df["Curve"].iloc[i])] = bk.addr(ws, info["r0"] + i, info["cols"].index("Current_state"))
        bk.CELL[("gl_rank", df["Curve"].iloc[i])] = bk.addr(ws, info["r0"] + i, rc)
    r += len(df) + 3
    r = bk.sec(ws, r, "3. Ranking: most global-driven -> most locally driven", 8)
    rk = df.sort_values("Weighted_R2", ascending=False)[["Curve", "Weighted_R2", "Idiosyncratic_share", "Beta_L1_on_G1", "Current_state"]].reset_index(drop=True)
    rk.insert(0, "Rank", range(1, len(rk) + 1))
    bk.table(ws, r, 0, rk, kinds={"Rank": "int", "Weighted_R2": "pct", "Idiosyncratic_share": "pct", "Beta_L1_on_G1": "num4"})
    r += len(rk) + 3
    r = bk.sec(ws, r, "4. Decoupling / re-coupling episodes (rolling 6M PC1 R2 vs full-sample R2; min 10 days)", 8)
    rows = []
    for cid, (eps, cur) in X.decoup.items():
        if not eps:
            rows.append(dict(Curve=cid, Episode="none", From="", To="", Days=0))
        for (nm, a, b, n) in eps:
            rows.append(dict(Curve=cid, Episode=nm, From=str(a.date()), To=str(b.date()), Days=n))
    bk.table(ws, r, 0, pd.DataFrame(rows), kinds={"Days": "int"})
    ws.set_column(0, 0, 30); ws.set_column(1, 14, 16)

def w_spill(bk, X):
    ws, r = bk.head("Rolling_Spillover", "Rolling 6M / 12M beta and R2 of local PC1 on SOFR PC1 (chosen timing: %s)" % X.gl["choice"],
                    ["Union of local calendars; last known value carried across non-trading days. 'n/a' before the first full window."])
    wins = X.spill_wins
    idx = sorted(set().union(*[set(d.index) for d in X.spill.values()]))
    cols = OrderedDict()
    for cid, d in X.spill.items():
        for w in wins:
            cols[f"{cid} R2 {w}d"] = d[f"r2_{w}d"].reindex(idx).ffill()
            cols[f"{cid} beta {w}d"] = d[f"beta_{w}d"].reindex(idx).ffill()
    df = pd.DataFrame(cols, index=idx)
    df = df.dropna(how="all")
    df.index.name = "Date"
    df = df.reset_index()
    bk.table(ws, r, 0, df, key="spill", kinds={c: "num4" for c in df.columns[1:]}, freeze=True)
    ws.set_column(0, 0, 12); ws.set_column(1, 60, 15)

# ---------------------------------------------------------------- RV / flies / forwards / pairs
def w_rv(bk, X):
    ws, r = bk.head("RV_Residuals", "Relative value on the curve: 3-PC fitted vs actual, residuals, z-scores, half-lives",
                    ["PCA on supported 1Y forwards (+ labelled block forwards); fitted level = start level + cumulated 3-PC fitted daily changes (uncentred projection on PC1-3); residual = actual - fitted (bp, live formula).",
                     f"Signal (live formula) = |z| > {P('Z_OUTLIER')} AND stable AND half-life <= {P('HL_MAX')} trading days. z / half-life on trailing windows (Assumptions). Rate ABOVE fit -> receive; BELOW fit -> pay."])
    hdr = ["Curve", "Point", "Type", "Actual (bp)", "Fitted (bp)", "Residual (bp)", "Residual z (1y)", "Half-life (days)", "Stable AR(1)", "Signal"]
    for j, h in enumerate(hdr):
        ws.write(r, j, h, bk.f["hdr"])
    ws.freeze_panes(r + 1, 0)
    r += 1
    zthr, hlmax = P("Z_OUTLIER"), P("HL_MAX")
    skipped = []
    for cid, d in X.rv.items():
        if not d["ok"]:
            skipped.append((cid, d["reason"]))
            continue
        first = r
        for _, row in d["table"].iterrows():
            ws.write(r, 0, cid); ws.write(r, 1, row["Point"]); ws.write(r, 2, row["Type"])
            bk.put(ws, r, 3, row["Actual"]); bk.put(ws, r, 4, row["Fitted"])
            bk.form(ws, r, 5, f"={xl_rowcol_to_cell(r, 3)}-{xl_rowcol_to_cell(r, 4)}", row["Actual"] - row["Fitted"])
            name = f"{cid}|RV|{row['Point']}"
            st = X.hist.stats(name)
            bk.form(ws, r, 6, X.hist.f_z(name), st["z"])
            bk.put(ws, r, 7, row["HalfLife"]); bk.put(ws, r, 8, row["Stable"])
            sig = ""
            if np.isfinite(row["HalfLife"]) and abs(st["z"]) > zthr and row["HalfLife"] <= hlmax and row["Stable"]:
                sig = "RECEIVE (rate above fit)" if st["z"] > 0 else "PAY (rate below fit)"
            zc, hc, sc_ = xl_rowcol_to_cell(r, 6), xl_rowcol_to_cell(r, 7), xl_rowcol_to_cell(r, 8)
            ws.write_formula(r, 9, f'=IF(AND(ISNUMBER({hc}),ABS({zc})>{zthr},{hc}<={hlmax},{sc_}="Y"),IF({zc}>0,"RECEIVE (rate above fit)","PAY (rate below fit)"),"")', bk.f["ftxt"], sig)
            r += 1
        bk.REG[("rv", cid)] = dict(sheet="RV_Residuals", hdr=first - 1, r0=first, r1=r - 1, c0=0, c1=9, cols=hdr)
    if skipped:
        r += 1
        ws.write(r, 0, "RV not computed for:", bk.f["bold"]); r += 1
        for cid, why in skipped:
            ws.write(r, 0, cid); ws.write(r, 1, why); r += 1
    ws.set_column(0, 0, 14); ws.set_column(1, 1, 12); ws.set_column(2, 2, 24); ws.set_column(3, 8, 14); ws.set_column(9, 9, 26)

def w_flies(bk, X):
    ws, r = bk.head("Flies", "PCA-neutral butterflies on supported 1Y forwards",
                    ["Weights w (rate space) solve PC1 and PC2 exposure = 0 on the forward-curve PCA; belly weight = -1. Fly value F = sum w_i x forward_i (bp). Short fly (z>0) = receive wings / pay belly.",
                     "DV01 of a forward-starting 1Y swap = 1bp x accrual x discount factor to the forward's end date, per unit notional. Notional multiples = w_i / DV01_i scaled to belly notional = 1 (negative = belly).",
                     "Carry n/a when a leg is spot-starting (0y1y): floating fixing not in a par export; forward-starting legs carry 0. Roll = sum w_i x 1M roll-down of leg i. Rolled z = z of F after 1M roll at unchanged curve."])
    rows = [f for d in X.rv.values() if d["ok"] for f in d["flies"]]
    if not rows:
        ws.write(r, 0, "No butterflies available: need three consecutive supported 1Y forward points in an RV-eligible curve.")
        return
    hdr = ["Curve", "Fly (start points)", "w1", "w2 (belly)", "w3", "DV01 leg1", "DV01 leg2", "DV01 leg3", "Notional leg1", "Notional leg2 (belly)", "Notional leg3",
           "PC1 exposure", "PC2 exposure", "Latest F (bp)", "Mean 1y", "SD 1y", "z-score", "Percentile", "Carry 1M (bp)", "Roll 1M (bp)", "Rolled z", "Half-life (d)", "Direction"]
    for j, h in enumerate(hdr):
        ws.write(r, j, h, bk.f["hdr"])
    ws.freeze_panes(r + 1, 2)
    r += 1
    first = r
    for f in rows:
        name = f"{f['Curve']}|FLY|{f['Fly']}"
        ws.write(r, 0, f["Curve"]); ws.write(r, 1, f["Fly"])
        for j, k in enumerate(["w1", "w2", "w3"]):
            bk.put(ws, r, 2 + j, f[k], "num4")
        for j, k in enumerate(["DV01_1", "DV01_2", "DV01_3"]):
            bk.put(ws, r, 5 + j, f[k], "num4")
        for j, k in enumerate(["N1", "N2", "N3"]):
            bk.put(ws, r, 8 + j, f[k], "num4")
        bk.put(ws, r, 11, f["PC1_exposure"], "num4"); bk.put(ws, r, 12, f["PC2_exposure"], "num4")
        st = X.hist.stats(name)
        bk.form(ws, r, 13, X.hist.f_latest(name), st["latest"]); bk.form(ws, r, 14, X.hist.f_mean(name), st["mean"]); bk.form(ws, r, 15, X.hist.f_sd(name), st["sd"])
        bk.form(ws, r, 16, X.hist.f_z(name), st["z"]); bk.form(ws, r, 17, X.hist.f_pct(name), st["pct"], "fpct")
        bk.put(ws, r, 18, f["Carry_1M"]); bk.put(ws, r, 19, f["Roll_1M"])
        bk.form(ws, r, 20, f"=IF({xl_rowcol_to_cell(r, 15)}>0,({xl_rowcol_to_cell(r, 13)}+{xl_rowcol_to_cell(r, 19)}-{xl_rowcol_to_cell(r, 14)})/{xl_rowcol_to_cell(r, 15)},0)", f["RolledZ"])
        bk.put(ws, r, 21, f["HalfLife"]); ws.write(r, 22, f["Direction"])
        r += 1
    bk.REG["fly"] = dict(sheet="Flies", hdr=first - 1, r0=first, r1=r - 1, c0=0, c1=22, cols=hdr)
    bk.FLY_ROWS = rows
    ws.set_column(0, 1, 26); ws.set_column(2, 21, 12); ws.set_column(22, 22, 44)

def w_forwards(bk, X):
    ws, r = bk.head("Forwards", "Slopes, forward-forward spreads and block forwards (bp)",
                    ["Slopes / spreads use supported 1Y forwards only. Block forwards are derived between QUOTED pars (never interpolated) and are labelled. z / percentile = live formulas on trailing 252 obs.",
                     "Interpolated forward points are shown only in the reference table at the bottom and never enter any PCA, fly or pair."])
    r = bk.sec(ws, r, "1. Slopes between supported points and forward-forward spreads", 10)
    hdr = ["Curve", "Series", "Type", "Latest (bp)", "Mean 1y", "SD 1y", "z-score", "Percentile", "Method"]
    for j, h in enumerate(hdr):
        ws.write(r, j, h, bk.f["hdr"])
    ws.freeze_panes(r + 1, 0)
    r += 1
    first = r
    for cid, (sp, bl, ip) in X.fwdtabs.items():
        for _, row in sp.iterrows():
            nm = f"{cid}|SPR|{row['Series']}"
            ws.write(r, 0, cid); ws.write(r, 1, row["Series"]); ws.write(r, 2, row["Type"])
            stat_row(bk, ws, r, 3, X.hist, nm)
            ws.write(r, 8, row["Method"])
            r += 1
    bk.REG["fwd_spreads"] = dict(sheet="Forwards", hdr=first - 1, r0=first, r1=r - 1, c0=0, c1=8, cols=hdr)
    r += 2
    r = bk.sec(ws, r, "2. Block forwards between quoted pillars (labelled BLOCK)", 10)
    hdr2 = ["Curve", "Point", "Pillars", "Type", "Latest (bp)", "Mean 1y", "SD 1y", "z-score", "Percentile", "Method"]
    for j, h in enumerate(hdr2):
        ws.write(r, j, h, bk.f["hdr"])
    r += 1
    for cid, (sp, bl, ip) in X.fwdtabs.items():
        for _, row in bl.iterrows():
            nm = f"{cid}|BLK|{row['Point']}"
            ws.write(r, 0, cid); ws.write(r, 1, row["Point"]); ws.write(r, 2, row["Pillars"]); ws.write(r, 3, row["Type"])
            stat_row(bk, ws, r, 4, X.hist, nm)
            ws.write(r, 9, row["Method"])
            r += 1
    r += 2
    r = bk.sec(ws, r, "3. Interpolated points - REFERENCE ONLY", 10)
    ip = pd.concat([v[2] for v in X.fwdtabs.values() if len(v[2])], ignore_index=True)
    bk.table(ws, r, 0, ip if len(ip) else pd.DataFrame({"Note": ["none"]}))
    ws.set_column(0, 0, 14); ws.set_column(1, 1, 18); ws.set_column(2, 3, 26); ws.set_column(4, 8, 13); ws.set_column(9, 9, 90)

def w_pairs(bk, X):
    ws, r = bk.head("Cross_Market_Pairs", "Cross-market PC1-neutral pairs (EM vs EM)",
                    ["Spread = A - hedge ratio x B (bp); hedge ratio = PC1 loading of A / PC1 loading of B at the same tenor / forward point. Ranked by |carry/roll-adjusted z|.",
                     "Adjusted z = (spread + 1M roll of A - ratio x 1M roll of B - mean) / SD. Carry is n/a (floating fixings / funding not in a par export). Top %d per set; full table below." % P("PAIR_TOP_N")])
    hdr = ["Rank", "Pair", "Set", "Instrument", "A", "B", "Hedge ratio", "Spread (bp)", "Mean 1y", "SD 1y", "z-score", "Percentile", "Corr 1y", "Corr full", "Roll 1M (bp)", "Carry 1M", "Adj z", "Direction"]
    first_table = True
    for sname in ("Par (main)", "Forward (comparison)"):
        r = bk.sec(ws, r, f"Top {P('PAIR_TOP_N')} - {sname}", 12)
        for j, h in enumerate(hdr):
            ws.write(r, j, h, bk.f["hdr"])
        if first_table:
            ws.freeze_panes(r + 1, 0)
            first_table = False
        r += 1
        t = X.top[sname]
        r0 = r
        for i, row in t.iterrows():
            name = f"PAIR|{row['Set'][:3]}|{row['Instrument']}|{row['A']}|{row['B']}"
            ws.write(r, 0, i + 1, bk.f["int"]); ws.write(r, 1, f"{row['A']} vs {row['B']} {row['Instrument']}")
            ws.write(r, 2, row["Set"]); ws.write(r, 3, row["Instrument"]); ws.write(r, 4, row["A"]); ws.write(r, 5, row["B"])
            bk.put(ws, r, 6, row["HedgeRatio"], "num4")
            stat_row(bk, ws, r, 7, X.hist, name)
            bk.put(ws, r, 12, row["Corr_1y"], "num4"); bk.put(ws, r, 13, row["Corr_full"], "num4"); bk.put(ws, r, 14, row["Roll_1M"]); bk.put(ws, r, 15, row["Carry_1M"])
            adj = f"=IF({xl_rowcol_to_cell(r, 9)}>0,({xl_rowcol_to_cell(r, 7)}+{xl_rowcol_to_cell(r, 14)}-{xl_rowcol_to_cell(r, 8)})/{xl_rowcol_to_cell(r, 9)},0)"
            bk.form(ws, r, 16, adj, row["AdjZ"])
            ws.write(r, 17, row["Direction"])
            r += 1
        bk.REG[("pairs_top", sname[:3])] = dict(sheet="Cross_Market_Pairs", hdr=r0 - 1, r0=r0, r1=r - 1, c0=0, c1=17, cols=hdr)
        r += 2
    r = bk.sec(ws, r, "Full pair table (static, ranked by |adjusted z|)", 12)
    full = X.pairs_all.copy()
    full["absAdj"] = full["AdjZ"].abs()
    full = full.sort_values(["Set", "absAdj"], ascending=[True, False]).drop(columns="absAdj")
    bk.table(ws, r, 0, full, kinds={"HedgeRatio": "num4", "Pct": "pct", "Corr_1y": "num4", "Corr_full": "num4", "N": "int"})
    ws.set_column(0, 0, 8); ws.set_column(1, 1, 34); ws.set_column(2, 5, 16); ws.set_column(6, 16, 12); ws.set_column(17, 17, 46)

# ---------------------------------------------------------------- Offshore liquidity / events / dashboard / templates / hist
def w_liquidity(bk, X):
    ws, r = bk.head("Offshore_Liquidity", "Offshore curve liquidity and quality",
                    [f"Zero-change share = share of OBSERVED daily changes exactly equal to zero (filled days excluded). Drop rule: > {P('ZERO_CHANGE_DROP'):.0%}; illiquid flag: > {P('ZERO_CHANGE_FLAG'):.0%}.",
                     "Stale run = quote unchanged for more than 2 consecutive days."])
    ts = pd.concat([C["tstats"].drop(columns="T") for C in X.C.values()], ignore_index=True)
    r = bk.sec(ws, r, "1. Stale-quote and zero-change statistics by tenor", 12)
    bk.table(ws, r, 0, ts, key="liq", kinds={"Quotes": "int", "Coverage": "pct", "ZeroChangeShare": "pct", "StaleRuns_gt2d": "int", "LongestUnchangedRun": "int", "DaysInStaleRuns": "int", "FilledDays": "int"}, freeze=True)
    r += len(ts) + 3
    r = bk.sec(ws, r, "2. PCA with vs without the illiquid (dropped) tenors", 12)
    rows = []
    for c, R in X.res.items():
        if R["par_incl"] is None:
            rows.append(dict(Curve=c, Dropped_tenors=", ".join(X.C[c]["dropped"]) or "none", Note="no tenor dropped - identical to MAIN PCA"))
            continue
        pi, pm = R["par_incl"], R["par_cov"]
        common = [t for t in R["tenors"] if t in R["incl_tenors"]]
        a = [R["incl_tenors"].index(t) for t in common]
        b = [R["tenors"].index(t) for t in common]
        rows.append(dict(Curve=c, Dropped_tenors=", ".join(X.C[c]["dropped"]), PC1_incl=pi["ve"][0], PC1_retained=pm["ve"][0], PC2_incl=pi["ve"][1], PC2_retained=pm["ve"][1],
                         PC3_incl=pi["ve"][2], PC3_retained=pm["ve"][2], Cong_PC1_common=abs(cong(pi["V"][a, 0], pm["V"][b, 0])), Cong_PC2_common=abs(cong(pi["V"][a, 1], pm["V"][b, 1])),
                         Cong_PC3_common=abs(cong(pi["V"][a, 2], pm["V"][b, 2])), Note="loading congruence on the common tenors"))
    df = pd.DataFrame(rows).fillna("")
    bk.table(ws, r, 0, df, kinds={c: "pct" for c in df.columns if c.startswith("PC")} | {c: "num4" for c in df.columns if c.startswith("Cong")})
    r += len(df) + 3
    r = bk.sec(ws, r, "3. Offshore PC1 vs SOFR PC1 (USD funding factor check)", 12)
    rows = []
    for cid, per in X.gl["per"].items():
        g = per[X.gl["choice"]]
        y = X.res[cid]["scores"]["PC1"]
        rows.append(dict(Curve=cid, Corr_PC1_vs_SOFR_PC1=float(np.corrcoef(y.values, g["G"]["PC1"].values)[0, 1]), Beta_PC1_on_SOFR_PC1=g["regs"]["PC1"]["b"][0], R2_PC1=g["regs"]["PC1"]["r2"],
                         Timing=X.gl["choice"]))
    bk.table(ws, r, 0, pd.DataFrame(rows), kinds={"Corr_PC1_vs_SOFR_PC1": "num4", "Beta_PC1_on_SOFR_PC1": "num4", "R2_PC1": "pct"})
    r += len(rows) + 3
    r = bk.sec(ws, r, "4. Offshore PC1 vs USD/Asia FX move", 12)
    if X.fxc is None:
        ws.write(r, 0, "NOT PROVIDED - no USD/Asia FX data supplied (see fx_template.csv). Skipped, nothing invented.")
    else:
        bk.table(ws, r, 0, X.fxc, kinds={"Corr_PC1_vs_FX": "num4", "R2": "pct", "N": "int"})
    ws.set_column(0, 0, 22); ws.set_column(1, 14, 15); ws.set_column(11, 11, 40)

def w_events(bk, X):
    ws, r = bk.head("Events", "Central-bank event analysis: PC1-PC3 score move over a [-1,+1] trading-day window")
    if X.evc is None:
        ws.write(r, 0, "NOT PROVIDED - no event calendar supplied. Template below (also events_template.csv); populate and re-run.", bk.f["bold"])
        r += 2
        t = pd.DataFrame({"Date": ["<YYYY-MM-DD>"], "Market": ["<curve id e.g. INR_NDIRS>"], "Event": ["<e.g. RBI MPC decision>"]})
        bk.table(ws, r, 0, t)
    else:
        bk.table(ws, r, 0, X.evc, kinds={c: "num" for c in X.evc.columns if c.endswith("_bp") or c.endswith("_z")}, freeze=True)
    ws.set_column(0, 12, 22)

def w_tenor_pct(bk, X):
    wins = [int(x) for x in str(P("PCT_WINS")).split(",")]
    ws, r = bk.head("Tenor_Percentiles", "Percentile dislocation of each tenor's rate level over the past 1y / 3y / 5y",
                    ["Percentile (PERCENTRANK.INC style) of the latest par rate / supported 1Y forward within its own trailing window; 0% = lowest, 100% = highest. Dislocation = latest minus window median (bp).",
                     "Windows = %s trading days. Static values written by Python; a window is 'n/a' if the curve has fewer observations than the window." % ", ".join(map(str, wins))])
    rows = []
    for c, R in X.res.items():
        lv = R["levels_full"]
        series = [(tenor_label(t), "Par", lv[t]) for t in lv.columns] + [(p, "1Y forward (supported)", R["sup"][p]) for p in R["sup"].columns]
        for nm, typ, s in series:
            s = s.dropna()
            row = dict(Curve=c, Tenor=nm, Type=typ, Latest_bp=float(s.iloc[-1]))
            for w, lab in zip(wins, ("1y", "3y", "5y")):
                if len(s) >= w:
                    a = s.values[-w:]
                    row[f"Pctile_{lab}"] = pct_inc(a, a[-1])
                    row[f"Dislocation_{lab}_bp"] = float(a[-1] - np.median(a))
                else:
                    row[f"Pctile_{lab}"], row[f"Dislocation_{lab}_bp"] = np.nan, np.nan
            rows.append(row)
    df = pd.DataFrame(rows)
    bk.table(ws, r, 0, df, key="tenor_pct", kinds={c: "pct" for c in df.columns if c.startswith("Pctile")}, freeze=True)
    n = len(df)
    ws.conditional_format(r + 1, 4, r + n, 4, {"type": "3_color_scale", "min_color": "#5B9BD5", "mid_color": "#FFFFFF", "max_color": "#F8696B", "min_type": "num", "min_value": 0, "mid_type": "num", "mid_value": 0.5, "max_type": "num", "max_value": 1})
    for j, col in enumerate(df.columns):
        if col.startswith("Pctile"):
            ws.conditional_format(r + 1, j, r + n, j, {"type": "3_color_scale", "min_color": "#5B9BD5", "mid_color": "#FFFFFF", "max_color": "#F8696B", "min_type": "num", "min_value": 0, "mid_type": "num", "mid_value": 0.5, "max_type": "num", "max_value": 1})
    ws.set_column(0, 2, 22); ws.set_column(3, 12, 15)

def w_dashboard(bk, X):
    ws, r = bk.head("Dashboard", "Dashboard - one row per curve",
                    ["All numbers are live links to the Curve_<id> and Global_vs_Local sheets. Percentiles are 1-year (252 obs); 20d = 20-day cumulative score."])
    hdr = ["Curve", "Level score (bp)", "Level 20d pctile", "Level 20d z", "Slope score (bp)", "Slope 20d pctile", "Slope 20d z", "Curv score (bp)", "Curv 20d pctile", "Curv 20d z",
           "Global beta (L1 on G1)", "Global-driven share (R2)", "Rank (1=most global)", "State", "Top RV signals", "Top fly (|z|)", "Top fly z"]
    for j, h in enumerate(hdr):
        ws.write(r, j, h, bk.f["hdr"])
    ws.freeze_panes(r + 1, 1)
    r += 1
    first = r
    st = X.hist
    for cid, R in X.res.items():
        ws.write(r, 0, cid)
        nf = len(R["latest_scores"])
        for k in range(3):
            if k < nf:
                bk.form(ws, r, 1 + 3 * k, "=" + bk.CELL[("score_daily", cid, k)], st.stats(f"{cid}|PC{k + 1}|daily")["latest"])
                bk.form(ws, r, 2 + 3 * k, "=" + bk.CELL[("score_cum_pct", cid, k)], st.stats(f"{cid}|PC{k + 1}|cum20")["pct"], "fpct")
                bk.form(ws, r, 3 + 3 * k, "=" + bk.CELL[("score_cum_z", cid, k)], st.stats(f"{cid}|PC{k + 1}|cum20")["z"])
        if cid in X.gl["per"]:
            g = X.gl["per"][cid][X.gl["choice"]]
            bk.form(ws, r, 10, "=" + bk.CELL[("gl_beta", cid)], g["regs"]["PC1"]["b"][0], "fnum4")
            bk.form(ws, r, 11, "=" + bk.CELL[("gl_wr2", cid)], g["wr2"], "fpct")
            order = pd.Series({c: v[X.gl["choice"]]["wr2"] for c, v in X.gl["per"].items()}).rank(ascending=False, method="min")[cid]
            bk.form(ws, r, 12, "=" + bk.CELL[("gl_rank", cid)], order, "fint")
            ws.write_formula(r, 13, "=" + bk.CELL[("gl_state", cid)], bk.f["ftxt"], X.decoup[cid][1])
        else:
            ws.write(r, 10, "SOFR = global benchmark")
        d = X.rv.get(cid)
        txt, topfly, topz = "", "", 0.0
        if d and d["ok"]:
            t = d["table"]
            ok = t[(t["z"].abs() > P("Z_OUTLIER")) & t["Stable"] & (t["HalfLife"] <= P("HL_MAX"))].copy()
            ok["a"] = ok["z"].abs()
            txt = "; ".join(f"{p} z={z:+.1f}" for p, z in zip(ok.sort_values("a", ascending=False)["Point"][:3], ok.sort_values("a", ascending=False)["z"][:3])) or "none"
            if d["flies"]:
                f = max(d["flies"], key=lambda x: abs(x["z"]))
                topfly, topz = f["Fly"], f["z"]
        ws.write(r, 14, txt or "n/a"); ws.write(r, 15, topfly or "n/a")
        bk.put(ws, r, 16, topz)
        r += 1
    bk.REG["dash"] = dict(sheet="Dashboard", hdr=first - 1, r0=first, r1=r - 1, c0=0, c1=16, cols=hdr)
    r += 2
    r = bk.sec(ws, r, "Heatmap: 1-year percentile of PC scores (daily and 20-day cumulative)", 8)
    hh = ["Curve", "Level daily", "Slope daily", "Curv daily", "Level 20d", "Slope 20d", "Curv 20d"]
    for j, h in enumerate(hh):
        ws.write(r, j, h, bk.f["hdr"])
    r += 1
    h0 = r
    for cid, R in X.res.items():
        ws.write(r, 0, cid)
        nf = len(R["latest_scores"])
        for k in range(3):
            if k < nf:
                bk.form(ws, r, 1 + k, "=" + bk.CELL[("score_daily_pct", cid, k)], st.stats(f"{cid}|PC{k + 1}|daily")["pct"], "fpct")
                bk.form(ws, r, 4 + k, "=" + bk.CELL[("score_cum_pct", cid, k)], st.stats(f"{cid}|PC{k + 1}|cum20")["pct"], "fpct")
        r += 1
    ws.conditional_format(h0, 1, r - 1, 6, {"type": "3_color_scale", "min_color": "#5B9BD5", "mid_color": "#FFFFFF", "max_color": "#F8696B", "min_type": "num", "min_value": 0, "mid_type": "num", "mid_value": 0.5, "max_type": "num", "max_value": 1})
    bk.REG["heat"] = dict(sheet="Dashboard", hdr=h0 - 1, r0=h0, r1=r - 1, c0=0, c1=6, cols=hh)
    ws.set_column(0, 0, 16); ws.set_column(1, 13, 13); ws.set_column(14, 15, 40); ws.set_column(16, 16, 12)

def w_ticker(bk, X):
    ws, r = bk.head("Ticker_Map", "Ticker map TEMPLATE - fill from Bloomberg (no tickers are guessed or recalled)",
                    ["Columns: Market, Curve, Tenor / forward point, Bloomberg ticker, Field, Pricing source, Quote convention. Blue = your inputs."])
    hdr = ["Market", "Curve", "Tenor / forward point", "Bloomberg ticker", "Field", "Pricing source", "Quote convention"]
    for j, h in enumerate(hdr):
        ws.write(r, j, h, bk.f["hdr"])
    ws.freeze_panes(r + 1, 0)
    r += 1
    for cid, m in CURVES.items():
        for tn in ["1Y", "2Y", "3Y", "4Y", "5Y", "7Y", "10Y"]:
            ws.write(r, 0, m["market"], bk.f["inp"]); ws.write(r, 1, cid, bk.f["inp"]); ws.write(r, 2, tn, bk.f["inp"])
            for j in range(3, 7):
                ws.write_blank(r, j, None, bk.f["inp"])
            r += 1
    ws.set_column(0, 6, 24)

def w_hist(bk, X):
    ws, r = bk.head("Hist", "Trailing series behind the live formulas (each column on its own calendar; last row = latest date)",
                    ["Row 4 = last date of each series; values are inputs (blue). Result-sheet z-scores, percentiles, means and SDs are formulas on these ranges."])
    H = X.hist
    fr = Hist.FIRST_ROW
    ws.write(fr - 2, 0, "Last date", bk.f["hdr"])
    ws.write(fr - 1, 0, "Obs # / series", bk.f["hdr"])
    for i in range(H.n):
        ws.write_number(fr + i, 0, i + 1)
    for name, d in H.cols.items():
        c = d["col"]
        ws.write(fr - 1, c, name, bk.f["hdr"])
        if d["last"] != "":
            ws.write_datetime(fr - 2, c, pd.Timestamp(d["last"]).to_pydatetime(), bk.f["inp_date"])
        vals = d["vals"]
        off = H.n - len(vals)
        for i, v in enumerate(vals):
            ws.write_number(fr + off + i, c, float(v), bk.f["inp_num4"])
    ws.freeze_panes(fr, 1)
    ws.set_column(0, 0, 14); ws.set_column(1, max(len(H.cols), 1), 18)

def build_results(X, path):
    bk = Book(path, X.synthetic, list(X.res.keys()))
    for fn in (w_assumptions, w_dataqc, w_curvebuild, w_pca_variance, w_pca_loadings, w_pca_scores):
        fn(bk, X)
    for cid in X.res:
        w_curve(bk, X, cid)
    for fn in (w_pooled, w_rolling, w_gl, w_spill, w_rv, w_flies, w_forwards, w_pairs, w_tenor_pct, w_liquidity, w_events, w_dashboard, w_ticker, w_hist, w_readme):
        fn(bk, X)
    bk.wb.close()
    return bk

# =====================================================================================
# 12. CHARTS WORKBOOK (native Excel charts on Chart_Data linked to the Results workbook)
# =====================================================================================
PALETTE = ["#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD", "#8C564B", "#E377C2", "#17BECF", "#BCBD22", "#000000"]

class ChartData:
    """Writes external-reference link cells on Chart_Data, keeps cached values for the externalLink part and a Link_Map."""
    def __init__(self, wb, ws, reader, target_name):
        self.wb, self.ws, self.reader, self.target = wb, ws, reader, target_name
        self.row = 2
        self.cache = {}
        self.linkmap = []
        self.fmt = wb.add_format({"num_format": "0.0000"})
        self.fmt_txt = wb.add_format({})
        self.fmt_h = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})

    def block(self, label, sheet, rows, cols, group):
        """rows: source rows (0-based) incl. header row if wanted; cols: source columns. Returns dict with destination coordinates."""
        ws = self.ws
        ws.write(self.row, 0, f"{label}   <- {sheet}", self.fmt_h)
        self.row += 1
        r0 = self.row
        for i, sr in enumerate(rows):
            for j, sc in enumerate(cols):
                v = self.reader.val(sheet, sr, sc)
                ref = xl_rowcol_to_cell(sr, sc, row_abs=True, col_abs=True)
                self.cache.setdefault(sheet, {})[xl_rowcol_to_cell(sr, sc)] = v
                f = f"=[1]{sheet}!{ref}"
                if v is None:
                    v = ""
                ws.write_formula(r0 + i, j, f, self.fmt if isinstance(v, (int, float)) else self.fmt_txt, v)
        blk = dict(r0=r0, r1=r0 + len(rows) - 1, ncols=len(cols))
        src = f"{sheet}!{xl_rowcol_to_cell(min(rows), min(cols))}:{xl_rowcol_to_cell(max(rows), max(cols))}" + ("" if list(rows) == list(range(min(rows), max(rows) + 1)) else " (sampled rows)")
        self.linkmap.append((group, label, f"Chart_Data!{xl_rowcol_to_cell(r0, 0)}:{xl_rowcol_to_cell(r0 + len(rows) - 1, len(cols) - 1)}", self.target, src))
        self.row = r0 + len(rows) + 2
        return blk

class ResultsReader:
    def __init__(self, path):
        import openpyxl
        self.wb = openpyxl.load_workbook(path, data_only=True)
    def val(self, sheet, r, c):
        v = self.wb[sheet].cell(row=r + 1, column=c + 1).value
        if isinstance(v, (dt.datetime, dt.date)):
            v = pd.Timestamp(v).strftime("%Y-%m-%d")
        return v

def build_charts(X, bk, results_path, charts_path):
    reader = ResultsReader(results_path)
    wb = xlsxwriter.Workbook(charts_path)
    target = os.path.basename(results_path)
    tabs = OrderedDict((n, wb.add_worksheet(n)) for n in ["Chart_Data"] + [f"Ch{i}_{nm}" for i, nm in
                                                                        enumerate(["Variance", "Loadings", "RollingPC1", "GlobalVsLocal", "RV", "Flies", "Pairs", "IndividualCurves", "PooledPCA", "Heatmap"], 1)] + ["Link_Map"])
    cdws = tabs["Chart_Data"]
    cd = ChartData(wb, cdws, reader, target)
    cdws.write(0, 0, "Chart_Data - every cell is an external-reference formula into " + target + " (cached values stored so charts render before links refresh). " + (SYNTH_LABEL if X.synthetic else ""), wb.add_format({"bold": True}))
    curve_col = {c: PALETTE[i % len(PALETTE)] for i, c in enumerate(X.res.keys())}
    sfx = f" [{SYNTH_LABEL}]" if X.synthetic else ""
    note_f = wb.add_format({"italic": True, "font_color": "#595959"})
    syn_f = wb.add_format({"bold": True, "font_color": "#C00000"})

    def tab_head(name, title, source):
        ws = tabs[name]
        ws.write(0, 0, title, wb.add_format({"bold": True, "font_size": 14}))
        ws.write(1, 0, SYNTH_LABEL if X.synthetic else "Run on user-supplied Bloomberg data", syn_f if X.synthetic else note_f)
        ws.write(2, 0, "Source: " + source + " (linked through Chart_Data)", note_f)
        return ws
    def new_chart(kind, sub=None):
        ch = wb.add_chart({"type": kind, "subtype": sub} if sub else {"type": kind})
        return ch
    def style(ch, title, xt, yt, w=640, h=330, legend="bottom", yfmt=None, bar=False, reverse=False):
        ch.set_title({"name": title + sfx, "name_font": {"size": 10}})
        xa = {"name": xt, "name_font": {"size": 9}, "num_font": {"size": 8}, "label_position": "low"}
        ya = {"name": yt, "name_font": {"size": 9}, "num_font": {"size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}}
        if yfmt:
            ya["num_format"] = yfmt
        if bar:                                   # xlsxwriter bar charts: x_axis = value axis (horizontal), y_axis = category axis (vertical)
            if reverse:
                xa["reverse"] = True
            ch.set_x_axis(ya)
            ch.set_y_axis(xa)
        else:
            ch.set_x_axis(xa)
            ch.set_y_axis(ya)
        ch.set_size({"width": w, "height": h})
        ch.set_legend({"position": legend, "font": {"size": 8}} if legend else {"none": True})
    def rng(blk, c0, c1=None, r_off0=0, r_off1=0):
        return ["Chart_Data", blk["r0"] + r_off0, c0, blk["r1"] - r_off1, c1 if c1 is not None else c0]
    def rows_of(info, include_header=True):
        return list(range(info["hdr"] if include_header else info["r0"], info["r1"] + 1))

    # ---- 1. variance explained stacked
    ws = tab_head("Ch1_Variance", "1. Variance explained by PC per curve (par rates, covariance, analysis sample)", "PCA_Variance")
    info = bk.REG["var_main"]
    cols = [info["cols"].index(k) for k in ("Curve", "PC1", "PC2", "PC3", "PC4")]
    blk = cd.block("Variance explained (main)", "PCA_Variance", rows_of(info), cols, "1")
    ch = new_chart("column", "stacked")
    for j, nm in enumerate(["PC1", "PC2", "PC3", "PC4"], 1):
        ch.add_series({"name": ["Chart_Data", blk["r0"], j], "categories": rng(blk, 0, r_off0=1), "values": rng(blk, j, r_off0=1), "fill": {"color": PALETTE[j - 1]}})
    style(ch, "Variance explained by PC", "Curve", "Share of total variance", w=900, h=380, yfmt="0%")
    ch.set_y_axis({"name": "Share of total variance (%)", "num_format": "0%", "min": 0, "max": 1})
    ws.insert_chart(4, 0, ch)

    # ---- 2. loadings by PC (scatter lines, all curves; numeric tenor axis)
    ws = tab_head("Ch2_Loadings", "2. Loadings by tenor - one chart per PC, one line per curve", "Curve_<id> sheets (loadings, par covariance)")
    blks = {}
    for c in X.res:
        info = bk.REG[("cv_load", c)]
        blks[c] = cd.block(f"Loadings {c}", f"Curve_{c}", rows_of(info), [info["cols"].index("Tenor_yrs")] + [info["cols"].index(f"PC{k}") for k in (1, 2, 3)], "2")
    for k in (1, 2, 3):
        ch = new_chart("scatter", "straight_with_markers")
        for c, b in blks.items():
            ch.add_series({"name": c, "categories": rng(b, 0, r_off0=1), "values": rng(b, k, r_off0=1), "line": {"color": curve_col[c], "width": 1.75}, "marker": {"type": "circle", "size": 4, "fill": {"color": curve_col[c]}, "border": {"color": curve_col[c]}}})
        style(ch, f"PC{k} loadings by tenor", "Tenor (years)", "Loading (unit-length eigenvector)", w=640, h=340, legend="right")
        ws.insert_chart(4 + (k - 1) * 18, 0, ch)
    # ---- 3. rolling PC1 variance
    ws = tab_head("Ch3_RollingPC1", "3. Rolling 1Y PC1 variance explained per curve", "Rolling_Stability (month-end windows)")
    info = bk.REG["roll_ve1"]
    blk = cd.block("Rolling PC1 variance by month", "Rolling_Stability", rows_of(info), list(range(info["c0"], info["c1"] + 1)), "3")
    ch = new_chart("line")
    for j, nm in enumerate(info["cols"][1:], 1):
        ch.add_series({"name": ["Chart_Data", blk["r0"], j], "categories": rng(blk, 0, r_off0=1), "values": rng(blk, j, r_off0=1), "line": {"color": curve_col.get(nm, "#808080"), "width": 1.5}})
    style(ch, "Rolling 1Y PC1 variance explained", "Window-end month", "Share of total variance (%)", w=1000, h=400, legend="right", yfmt="0%")
    ws.insert_chart(4, 0, ch)
    # ---- 4. global R2 / beta
    ws = tab_head("Ch4_GlobalVsLocal", "4. Global (SOFR) R-squared and beta per curve, and rolling 6M R-squared / beta", "Global_vs_Local, Rolling_Spillover")
    info = bk.REG["gl_main"]
    cols = [info["cols"].index(k) for k in ("Curve", "Weighted_R2", "Beta_L1_on_G1")]
    blk = cd.block("Global-driven share and beta", "Global_vs_Local", rows_of(info), cols, "4")
    ch = new_chart("column")
    ch.add_series({"name": "Weighted R-squared (share of PC1-3 variance driven by SOFR)", "categories": rng(blk, 0, r_off0=1), "values": rng(blk, 1, r_off0=1), "fill": {"color": PALETTE[0]}})
    style(ch, "Global-driven share of local PC1-3 variance (R2)", "Curve", "R-squared (%)", w=640, h=330, legend=None, yfmt="0%")
    ws.insert_chart(4, 0, ch)
    ch = new_chart("column")
    ch.add_series({"name": "Beta local PC1 on SOFR PC1", "categories": rng(blk, 0, r_off0=1), "values": rng(blk, 2, r_off0=1), "fill": {"color": PALETTE[1]}})
    style(ch, "Spillover beta: local PC1 on SOFR PC1", "Curve", "Beta (local PC1 bp per SOFR PC1 bp)", w=640, h=330, legend=None)
    ws.insert_chart(4, 11, ch)
    sinfo = bk.REG["spill"]
    srows = [sinfo["hdr"]] + list(range(sinfo["r0"], sinfo["r1"] + 1, int(P("MONTH_DAYS"))))
    wins = X.spill_wins
    for m, (metric, yt, fmt) in enumerate((("R2", "R-squared (%)", "0%"), ("beta", "Beta", None))):
        cols = [0] + [sinfo["cols"].index(f"{c} {metric} {wins[0]}d") for c in X.spill]
        b2 = cd.block(f"Rolling {wins[0]}d {metric} (monthly sample)", "Rolling_Spillover", srows, cols, "4")
        ch = new_chart("line")
        for j, c in enumerate(X.spill, 1):
            ch.add_series({"name": c, "categories": rng(b2, 0, r_off0=1), "values": rng(b2, j, r_off0=1), "line": {"color": curve_col[c], "width": 1.25}})
        style(ch, f"Rolling {wins[0]}-day {metric} of local PC1 on SOFR PC1", "Date", yt, w=1000, h=380, legend="right", yfmt=fmt)
        ws.insert_chart(22 + m * 20, 0, ch)
    # ---- 5. RV
    ws = tab_head("Ch5_RV", "5. RV: fitted vs actual forward curve and residual bars, latest date, per curve", "RV_Residuals")
    k = 0
    for c in X.res:
        if ("rv", c) not in bk.REG:
            continue
        info = bk.REG[("rv", c)]
        cols = [info["cols"].index(x) for x in ("Point", "Actual (bp)", "Fitted (bp)", "Residual (bp)")]
        b = cd.block(f"RV {c}", "RV_Residuals", rows_of(info), cols, "5")
        ch = new_chart("line")
        for j, (nm, dash) in enumerate((("Actual", "solid"), ("Fitted (3 PCs)", "dash")), 1):
            ch.add_series({"name": nm, "categories": rng(b, 0, r_off0=1), "values": rng(b, j, r_off0=1), "line": {"color": curve_col[c] if j == 1 else "#7F7F7F", "width": 2, "dash_type": dash}, "marker": {"type": "circle", "size": 4}})
        style(ch, f"{c}: actual vs fitted forward curve", "Forward point (BLOCK = block forward)", "Rate (bp)", w=520, h=300)
        ws.insert_chart(4 + k * 16, 0, ch)
        ch = new_chart("column")
        ch.add_series({"name": "Residual (actual - fitted)", "categories": rng(b, 0, r_off0=1), "values": rng(b, 3, r_off0=1), "fill": {"color": curve_col[c]}})
        style(ch, f"{c}: residual by point", "Forward point", "Residual (bp)", w=520, h=300, legend=None)
        ws.insert_chart(4 + k * 16, 9, ch)
        k += 1
    # ---- 6. flies
    ws = tab_head("Ch6_Flies", "6. Fly z-scores and history of the top flies", "Flies, Hist")
    if "fly" in bk.REG:
        info = bk.REG["fly"]
        cols = [info["cols"].index(x) for x in ("Curve", "Fly (start points)", "z-score")]
        b = cd.block("Fly z-scores", "Flies", rows_of(info), cols, "6")
        for i in range(b["r0"] + 1, b["r1"] + 1):
            lab = f"{cdws.table_cache[(i, 0)] if hasattr(cdws, 'table_cache') else ''}"
            cur, fly = reader.val("Flies", info["r0"] + (i - b["r0"] - 1), cols[0]), reader.val("Flies", info["r0"] + (i - b["r0"] - 1), cols[1])
            cdws.write_formula(i, 3, f'={xl_rowcol_to_cell(i, 0)}&" | "&{xl_rowcol_to_cell(i, 1)}', cd.fmt_txt, f"{cur} | {fly}")
        ch = new_chart("bar")
        ch.add_series({"name": "z-score (1y)", "categories": ["Chart_Data", b["r0"] + 1, 3, b["r1"], 3], "values": rng(b, 2, r_off0=1), "fill": {"color": PALETTE[0]}, "invert_if_negative": False})
        style(ch, "Fly z-scores (PCA-neutral butterflies on supported 1Y forwards)", "Fly", "z-score (SD)", w=760, h=max(340, 30 * (info["r1"] - info["r0"] + 2)), legend=None, bar=True, reverse=True)
        ws.insert_chart(4, 0, ch)
        top = sorted(bk.FLY_ROWS, key=lambda f: -abs(f["z"]))[:5]
        names = [f"{f['Curve']}|FLY|{f['Fly']}" for f in top]
        fr = Hist.FIRST_ROW
        hrows = [fr - 1] + list(range(fr, fr + X.hist.n))
        cols = [0] + [X.hist.cols[n]["col"] for n in names]
        b2 = cd.block("History of top flies (trailing 252 obs)", "Hist", hrows, cols, "6")
        ch = new_chart("line")
        for j in range(1, len(cols)):
            ch.add_series({"name": ["Chart_Data", b2["r0"], j], "categories": rng(b2, 0, r_off0=1), "values": rng(b2, j, r_off0=1), "line": {"width": 1.5, "color": PALETTE[j - 1]}})
        style(ch, "History of top-5 flies by |z| (value in bp)", "Observation # (252 = latest)", "Fly value (bp)", w=900, h=380, legend="right")
        ws.insert_chart(4 + int(max(340, 30 * (info["r1"] - info["r0"] + 2)) / 20) + 2, 0, ch)
    else:
        ws.write(4, 0, "No flies available for this run (need three consecutive supported 1Y forward points).")
    # ---- 7. pairs
    ws = tab_head("Ch7_Pairs", "7. Cross-market PCA-neutral pair z-score ranking (top 10 per set)", "Cross_Market_Pairs")
    for i, sname in enumerate(("Par", "For")):
        info = bk.REG[("pairs_top", sname)]
        cols = [info["cols"].index(x) for x in ("Pair", "z-score", "Adj z")]
        b = cd.block(f"Top pairs {sname}", "Cross_Market_Pairs", rows_of(info), cols, "7")
        ch = new_chart("bar")
        ch.add_series({"name": "z-score", "categories": rng(b, 0, r_off0=1), "values": rng(b, 1, r_off0=1), "fill": {"color": PALETTE[0]}, "invert_if_negative": False})
        ch.add_series({"name": "Roll-adjusted z", "categories": rng(b, 0, r_off0=1), "values": rng(b, 2, r_off0=1), "fill": {"color": PALETTE[1]}, "invert_if_negative": False})
        style(ch, "Top pairs (par 2Y/5Y/10Y)" if sname == "Par" else "Top pairs (supported 1y1y / 3y1y, comparison)", "Pair", "z-score (SD)", w=860, h=420, legend="bottom", bar=True, reverse=True)
        ws.insert_chart(4 + i * 24, 0, ch)
    # ---- 8. individual curves
    ws = tab_head("Ch8_IndividualCurves", "8. Individual curves: loadings (PC1-PC3 by tenor) and variance explained", "Curve_<id> sheets")
    for i, c in enumerate(X.res):
        info = bk.REG[("cv_load", c)]
        b = blks[c]
        ch = new_chart("line")
        for k2, col in zip((1, 2, 3), PALETTE[:3]):
            ch.add_series({"name": ["Chart_Data", b["r0"], k2], "categories": ["Chart_Data", b["r0"] + 1, 0, b["r1"], 0], "values": ["Chart_Data", b["r0"] + 1, k2, b["r1"], k2], "line": {"color": col, "width": 2}, "marker": {"type": "circle", "size": 4}})
        # categories: numeric tenor years
        style(ch, f"{c}: loadings by tenor", "Tenor (years)", "Loading", w=480, h=280)
        ws.insert_chart(4 + i * 15, 0, ch)
        vi = bk.REG[("cv_var", c)]
        vb = cd.block(f"Variance {c}", f"Curve_{c}", [vi["hdr"], vi["r0"]], [vi["cols"].index(k) for k in ("PC1", "PC2", "PC3", "PC4")], "8")
        ch = new_chart("column")
        ch.add_series({"name": f"{c} variance explained", "categories": ["Chart_Data", vb["r0"], 0, vb["r0"], 3], "values": ["Chart_Data", vb["r0"] + 1, 0, vb["r0"] + 1, 3], "fill": {"color": curve_col[c]}})
        style(ch, f"{c}: variance explained (par, covariance)", "Principal component", "Share of total variance (%)", w=480, h=280, legend=None, yfmt="0%")
        ws.insert_chart(4 + i * 15, 8, ch)
    # ---- 9. pooled
    ws = tab_head("Ch9_PooledPCA", "9. Pooled EM Asia PCA: variance explained, market loadings and correlation with pooled PC1", "Pooled_PCA")
    info = bk.REG["pool_sum"]
    cols = [info["cols"].index(k) for k in ("Set", "Variant", "PC1", "PC2", "PC3", "PC4")]
    rows = [info["hdr"]] + [r for r in range(info["r0"], info["r1"] + 1) if isinstance(reader.val("Pooled_PCA", r, info["cols"].index("PC1")), (int, float))]
    b = cd.block("Pooled variance explained", "Pooled_PCA", rows, cols, "9")
    ch = new_chart("column", "stacked")
    for j, nm in enumerate(["PC1", "PC2", "PC3", "PC4"], 2):
        ch.add_series({"name": nm, "categories": ["Chart_Data", b["r0"] + 1, 0, b["r1"], 1], "values": rng(b, j, r_off0=1), "fill": {"color": PALETTE[j - 2]}})
    style(ch, "Pooled EM Asia PCA: variance explained (with / without SOFR)", "Set / variant", "Share of total variance (%)", w=1000, h=420, legend="bottom", yfmt="0%")
    ch.set_y_axis({"num_format": "0%", "min": 0, "max": 1, "name": "Share of total variance (%)"})
    ws.insert_chart(4, 0, ch)
    for j, var in enumerate(("ex-SOFR", "incl. SOFR")):
        key = ("pool_load", "Par 5Y", var)
        if key not in bk.REG:
            continue
        li = bk.REG[key]
        cols = [li["cols"].index(k) for k in ("Market", "PC1_loading", "PC2_loading", "Corr_with_pooled_PC1")]
        b = cd.block(f"Pooled 5Y loadings ({var})", "Pooled_PCA", rows_of(li), cols, "9")
        ch = new_chart("column")
        ch.add_series({"name": "PC1 loading", "categories": rng(b, 0, r_off0=1), "values": rng(b, 1, r_off0=1), "fill": {"color": PALETTE[0]}})
        ch.add_series({"name": "PC2 loading", "categories": rng(b, 0, r_off0=1), "values": rng(b, 2, r_off0=1), "fill": {"color": PALETTE[1]}})
        style(ch, f"Pooled 5Y par PCA ({var}): market loadings on PC1 and PC2", "Market", "Loading", w=640, h=340)
        ws.insert_chart(26 + j * 18, 0, ch)
        ch = new_chart("column")
        ch.add_series({"name": "Correlation with pooled PC1", "categories": rng(b, 0, r_off0=1), "values": rng(b, 3, r_off0=1), "fill": {"color": PALETTE[2]}})
        style(ch, f"Pooled 5Y par PCA ({var}): correlation with pooled PC1", "Market", "Correlation", w=640, h=340, legend=None)
        ws.insert_chart(26 + j * 18, 11, ch)
    # ---- 10. heatmap
    ws = tab_head("Ch10_Heatmap", "10. Dashboard heatmap: 1-year percentile of PC scores", "Dashboard")
    hi = bk.REG["heat"]
    b = cd.block("PC score percentiles", "Dashboard", rows_of(hi), list(range(hi["c0"], hi["c1"] + 1)), "10")
    hdrf = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1})
    pf = wb.add_format({"num_format": "0%", "border": 1, "align": "center"})
    for j, h in enumerate(hi["cols"]):
        ws.write(4, j, h, hdrf)
    for i in range(hi["r1"] - hi["r0"] + 1):
        ws.write_formula(5 + i, 0, f"=Chart_Data!{xl_rowcol_to_cell(b['r0'] + 1 + i, 0)}", wb.add_format({"border": 1}), reader.val("Dashboard", hi["r0"] + i, 0))
        for j in range(1, len(hi["cols"])):
            v = reader.val("Dashboard", hi["r0"] + i, j)
            ws.write_formula(5 + i, j, f"=Chart_Data!{xl_rowcol_to_cell(b['r0'] + 1 + i, j)}", pf, v if isinstance(v, (int, float)) else 0)
    ws.conditional_format(5, 1, 5 + hi["r1"] - hi["r0"], len(hi["cols"]) - 1, {"type": "3_color_scale", "min_color": "#5B9BD5", "mid_color": "#FFFFFF", "max_color": "#F8696B", "min_type": "num", "min_value": 0, "mid_type": "num", "mid_value": 0.5, "max_type": "num", "max_value": 1})
    ws.write(6 + hi["r1"] - hi["r0"], 0, "Blue = low percentile (score weak vs last 1y), red = high. Daily = latest daily score; 20d = 20-day cumulative score.", note_f)
    ws.set_column(0, 0, 18); ws.set_column(1, 8, 12)
    # ---- Link_Map
    lm = tabs["Link_Map"]
    lm.write(0, 0, "Link_Map - every Chart_Data range and the Results-workbook range it mirrors", wb.add_format({"bold": True, "font_size": 14}))
    for j, h in enumerate(["Chart group", "Block", "Chart_Data range", "Source workbook", "Source sheet and range"]):
        lm.write(2, j, h, hdrf)
    for i, row in enumerate(cd.linkmap):
        for j, v in enumerate(row):
            lm.write(3 + i, j, v)
    lm.set_column(0, 0, 12); lm.set_column(1, 1, 44); lm.set_column(2, 4, 44)
    cdws.set_column(0, 0, 30); cdws.set_column(1, 40, 14)
    wb.close()
    sheet_names = reader.wb.sheetnames
    inject_external_links(charts_path, target, sheet_names, cd.cache)
    return cd

def _xml_esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def inject_external_links(xlsx_path, target_name, sheet_names, cache):
    """Add an externalLink part (with cached cell values) so the [1]Sheet!Cell formulas on Chart_Data are genuine external references."""
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    data_sets = []
    for si, sn in enumerate(sheet_names):
        cells = cache.get(sn)
        if not cells:
            continue
        rows = {}
        for a1, v in cells.items():
            col, row = coordinate_from_string(a1)
            rows.setdefault(row, []).append((column_index_from_string(col), a1, v))
        xs = [f'<sheetData sheetId="{si}">']
        for row in sorted(rows):
            xs.append(f'<row r="{row}">')
            for _, a1, v in sorted(rows[row]):
                if v is None:
                    xs.append(f'<cell r="{a1}"/>')
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    xs.append(f'<cell r="{a1}"><v>{repr(float(v))}</v></cell>')
                else:
                    xs.append(f'<cell r="{a1}" t="str"><v>{_xml_esc(v)}</v></cell>')
            xs.append("</row>")
        xs.append("</sheetData>")
        data_sets.append("".join(xs))
    ext = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<externalBook r:id="rId1"><sheetNames>' + "".join(f'<sheetName val="{_xml_esc(n)}"/>' for n in sheet_names) + "</sheetNames>"
           "<sheetDataSet>" + "".join(data_sets) + "</sheetDataSet></externalBook></externalLink>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="{_xml_esc(target_name)}" TargetMode="External"/></Relationships>')
    tmp = xlsx_path + ".tmp"
    with zipfile.ZipFile(xlsx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                s = data.decode("utf8")
                s = s.replace("</Types>", '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>')
                data = s.encode("utf8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                s = data.decode("utf8")
                nid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', s)) + 1
                s = s.replace("</Relationships>", f'<Relationship Id="rId{nid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>')
                data = s.encode("utf8")
                rid = f"rId{nid}"
            elif item.filename == "xl/workbook.xml":
                wbxml = data.decode("utf8")
                continue_wb = True
            zout.writestr(item, data)
        # workbook.xml needs the rId assigned above -> rewrite after the loop
    with zipfile.ZipFile(tmp) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item, data in items:
            if item.filename == "xl/workbook.xml":
                s = data.decode("utf8")
                s = s.replace("</sheets>", f'</sheets><externalReferences><externalReference r:id="{rid}"/></externalReferences>', 1)
                data = s.encode("utf8")
            zout.writestr(item, data)
        zout.writestr("xl/externalLinks/externalLink1.xml", ext)
        zout.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", rels)
    os.remove(tmp)

# =====================================================================================
# 13. TESTS
# =====================================================================================
def run_tests(X):
    T = []
    def add(name, ok, detail=""):
        T.append(dict(Test=name, Result="PASS" if ok else "FAIL", Detail=detail))
    pcs = []
    for c, R in X.res.items():
        for nm in ("par_cov", "par_corr", "fwd_cov", "fwd_corr", "par_cov_full", "par_incl"):
            if R.get(nm) is not None:
                pcs.append((f"{c}:{nm}", R[nm]))
    for d in X.pooled.values():
        if d["res"] is not None:
            pcs.append(("pooled", d["res"]["p"]))
    err_orth = max(np.abs(p["V"].T @ p["V"] - np.eye(p["V"].shape[1])).max() for _, p in pcs)
    add("PC loadings orthonormal (V'V = I)", err_orth < 1e-9, f"max abs deviation {err_orth:.2e} over {len(pcs)} PCAs")
    err_ve = max(abs(p["ve"].sum() - 1) for _, p in pcs)
    add("Variance explained sums to 1", err_ve < 1e-9, f"max deviation {err_ve:.2e}")
    # reconstruction sanity (full rank exact; 3-PC error share equals 1 - cumulative variance)
    worst_full, worst3 = 0.0, 0.0
    for c, R in X.res.items():
        p, dY = R["par_cov"], R["dY"].values
        Xc = dY - p["mu"]
        full = np.linalg.norm(Xc - p["scores"] @ p["V"].T) / np.linalg.norm(Xc)
        rec3 = p["scores"][:, :3] @ p["V"][:, :3].T
        share = np.sum((Xc - rec3) ** 2) / np.sum(Xc ** 2)
        worst_full = max(worst_full, full)
        worst3 = max(worst3, abs(share - (1 - p["ve"][:3].sum())))
    add("Reconstruction: full-rank error ~ 0", worst_full < 1e-9, f"max relative error {worst_full:.2e}")
    add("Reconstruction: 3-PC residual share = 1 - cumulative variance", worst3 < 1e-9, f"max deviation {worst3:.2e}")
    # independent re-computation of one market's PCA (scikit-learn) and via SVD
    c0 = "INR_NDIRS" if "INR_NDIRS" in X.res else list(X.res)[0]
    R0 = X.res[c0]
    try:
        from sklearn.decomposition import PCA as SKPCA
        sk = SKPCA().fit(R0["dY"].values)
        d1 = np.abs(sk.explained_variance_ratio_[:4] - R0["par_cov"]["ve"][:4]).max()
        d2 = np.abs(np.abs(sk.components_[:3].T) - np.abs(R0["par_cov"]["V"][:, :3])).max()
        add(f"Independent check ({c0}): scikit-learn PCA variance and |loadings|", d1 < 1e-9 and d2 < 1e-6, f"max diff variance {d1:.1e}, loadings {d2:.1e}")
    except ImportError:
        add(f"Independent check ({c0}): SVD", True, "scikit-learn not installed - see SVD check")
    Xc = R0["dY"].values - R0["dY"].values.mean(0)
    s = np.linalg.svd(Xc, compute_uv=False) ** 2
    d3 = np.abs(s / s.sum() - R0["par_cov"]["ve"]).max()
    add(f"Independent check ({c0}): SVD of the data matrix", d3 < 1e-9, f"max diff variance {d3:.1e}")
    # sign conventions
    bad = []
    for c, R in X.res.items():
        V = R["par_cov"]["V"]
        n = V.shape[0]
        if V[:, 0].sum() <= 0 or (V[-1, 1] - V[0, 1]) <= 0 or (V[n // 2, 2] - 0.5 * (V[0, 2] + V[-1, 2])) <= 0:
            bad.append(c)
    add("Sign normalisation (PC1 sum>0, PC2 long-short>0, PC3 belly-wings>0)", not bad, ", ".join(bad) or "all curves")
    # re-pricing
    worst = max(B["max_err"] for B in X.Bd.values())
    add("Bootstrap re-prices every input par rate (< %.2f bp)" % P("REPRICE_TOL_BP"), worst < P("REPRICE_TOL_BP"), f"max error {worst:.2e} bp")
    # 0y1y equals 1Y par
    d0 = max(float((B["supported"]["0y1y"] - B["par"][1]).abs().max()) for B in X.Bd.values() if "0y1y" in B["supported"].columns)
    add("0y1y forward equals 1Y par rate", d0 < 1e-6, f"max diff {d0:.1e} bp")
    # flies neutral
    fl = [f for d in X.rv.values() if d["ok"] for f in d["flies"]]
    if fl:
        e = max(max(abs(f["PC1_exposure"]), abs(f["PC2_exposure"])) for f in fl)
        add("Flies PC1/PC2-neutral", e < 1e-9, f"max exposure {e:.1e} over {len(fl)} flies")
    # no NaNs in delivered core tables
    nan_ct = 0
    for c, R in X.res.items():
        nan_ct += int(np.isnan(R["scores"].values).sum()) + int(np.isnan(R["par_cov"]["V"]).sum()) + int(np.isnan(R["par_cov"]["ve"]).sum())
        nan_ct += int(R["rolling"][["VE_PC1", "VE_PC2", "VE_PC3", "Cong_PC1", "Cong_PC2", "Cong_PC3"]].isna().sum().sum())
    add("No NaNs in variance / loadings / scores / rolling tables", nan_ct == 0, f"{nan_ct} NaNs")
    # unforeseen: unsupported forwards never used
    used_bad = []
    for c, R in X.res.items():
        interp = set(R["B"]["interp"].columns)
        if interp & set(R["sup"].columns):
            used_bad.append(c)
    add("No interpolated forward point in any PCA / fly / pair set", not used_bad, ", ".join(used_bad) or "supported sets contain no interpolated points")
    return T

# =====================================================================================
# 14. MAIN PIPELINE
# =====================================================================================
class Ctx:
    pass

def run_pipeline(df, insp, demo, out_dir, log=print):
    X = Ctx()
    X.synthetic, X.insp, X.run_date = demo, insp, dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    qc = QC()
    X.qc = qc
    unknown = sorted(set(df["Market"]) - set(CURVES))
    for u in unknown:
        qc.add(u, "Market not in CURVES config - ignored")
    df = df[df["Market"].isin(CURVES)]
    C_all = OrderedDict()
    for cid in CURVES:
        g = df[df["Market"] == cid]
        if g.empty:
            qc.add(cid, "No data supplied for this curve - not analysed")
            continue
        C_all[cid] = clean_curve(cid, g, qc)
        log(f"  cleaned {cid}: {C_all[cid]['status']}, tenors kept {[tenor_label(t) for t in C_all[cid]['keep']]}, dropped {C_all[cid]['dropped']}")
    C = OrderedDict((c, v) for c, v in C_all.items() if v["status"] == "ok")
    if GLOBAL_CURVE not in C:
        raise SystemExit("SOFR curve is required for the global-vs-local section but is missing / failed cleaning.")
    X.C, X.C_all = C, C_all
    end = pd.Timestamp(P("SAMPLE_END")) if P("SAMPLE_END") else max(v["levels"].index.max() for v in C.values())
    start = pd.Timestamp(P("SAMPLE_START")) if P("SAMPLE_START") else end - pd.DateOffset(years=int(P("LAST_N_YEARS")))
    X.start, X.end = start, end
    hist = Hist()
    X.hist = hist
    X.Bd = OrderedDict((c, build_curve(c, v["levels"])) for c, v in C.items())
    res = OrderedDict()
    for c in C:
        R = analyse_curve(c, C[c], X.Bd[c], start, end, qc, hist)
        R["break_dates"] = C[c]["break_dates"]
        res[c] = R
    X.res = res
    X.gl = global_vs_local(res)
    X.spill, X.spill_wins = rolling_spill(res, X.gl)
    X.decoup = {c: decoupling_episodes(X.spill[c], X.gl["per"][c][X.gl["choice"]]["regs"]["PC1"]["r2"], X.spill_wins[0]) for c in X.spill}
    X.pooled = pooled_sets(res)
    X.tie = tieback(res, X.gl)
    X.rv = OrderedDict((c, rv_analysis(c, res[c], hist, qc)) for c in res)
    X.fwdtabs = OrderedDict((c, forward_tables(c, res[c], hist)) for c in res)
    X.pairs_all, X.top = pairs_analysis(res, hist)
    X.fxc = fx_check(res, read_optional(os.path.join(out_dir, "fx_rates.csv")))
    X.evc = event_analysis(res, read_optional(os.path.join(out_dir, "events.csv")))
    X.bbgc = bbg_fwd_compare(res, read_optional(os.path.join(out_dir, "bbg_forwards.csv")))
    X.tests = run_tests(X)
    for t in X.tests:
        log(f"  [{t['Result']}] {t['Test']} - {t['Detail']}")
    return X

def summary_lines(X):
    L = []
    pc1 = {c: R["par_cov"]["ve"][0] for c, R in X.res.items()}
    L.append("PC1 variance (par, cov): " + ", ".join(f"{c} {v:.1%}" for c, v in pc1.items()))
    wr = {c: v[X.gl["choice"]]["wr2"] for c, v in X.gl["per"].items()}
    rk = sorted(wr.items(), key=lambda kv: -kv[1])
    L.append(f"Global-vs-local timing choice: {X.gl['choice']} (median weighted R2 {X.gl['median_wr2']})")
    L.append("Most -> least global-driven: " + ", ".join(f"{c} {v:.1%}" for c, v in rk))
    sig = []
    for c, d in X.rv.items():
        if d["ok"]:
            t = d["table"]
            ok = t[(t["z"].abs() > P("Z_OUTLIER")) & t["Stable"] & (t["HalfLife"] <= P("HL_MAX"))]
            for _, r in ok.iterrows():
                sig.append((abs(r["z"]), c, r["Point"], r["z"], r["HalfLife"]))
    sig.sort(reverse=True)
    L.append("RV signals passing all filters: " + (", ".join(f"{c} {p} z={z:+.1f} HL={h:.0f}d" for _, c, p, z, h in sig[:6]) or "none"))
    for sname, t in X.top.items():
        if len(t):
            r = t.iloc[0]
            L.append(f"Top pair ({sname}): {r['A']} vs {r['B']} {r['Instrument']} z={r['z']:+.2f} adj z={r['AdjZ']:+.2f}")
    L.append(f"Curves: {len(X.res)}; sample {X.start.date()} to {X.end.date()}")
    return L

def main():
    ap = argparse.ArgumentParser(description="EM Asia swap-rate PCA framework")
    ap.add_argument("--data", help="Bloomberg export (long format CSV/XLSX: Date, Market, Tenor, Rate [, Source, Unit])")
    ap.add_argument("--demo", action="store_true", help="run on synthetic data (all outputs labelled ILLUSTRATIVE ONLY)")
    ap.add_argument("--out", default=HERE, help="output folder (both workbooks + templates)")
    ap.add_argument("--templates", action="store_true", help="only write CSV templates")
    ap.add_argument("--no-charts", action="store_true")
    ap.add_argument("--no-verify", action="store_true")
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    os.makedirs(out, exist_ok=True)
    write_templates(out)
    if a.templates:
        print("Templates written to", out)
        return
    data_path = a.data or (os.path.join(out, "swap_rates.csv") if os.path.exists(os.path.join(out, "swap_rates.csv")) else None)
    if not a.demo and not data_path:
        print("No swap-rate data supplied. Export par swap rates from Bloomberg (BDH / Excel API / Terminal) in the layout of swap_rates_template.csv,\n"
              "fill ticker_map_template.csv, then run:  python build_all.py --data swap_rates.csv\n(or run  python build_all.py --demo  to see the pipeline on clearly labelled synthetic data).")
        sys.exit(2)
    print("Loading data ...")
    df, insp = load_data(data_path, a.demo)
    print(f"  {insp['rows']:,} rows, {insp['date_min'].date()} to {insp['date_max'].date()}; markets: {', '.join(insp['by_market']['Market'])}")
    X = run_pipeline(df, insp, a.demo, out)
    print("Writing Results workbook ...")
    rpath, cpath = os.path.join(out, RESULTS_NAME), os.path.join(out, CHARTS_NAME)
    bk = build_results(X, rpath)
    print(f"  {RESULTS_NAME}: {len(bk.ws)} sheets, {bk.na_count} documented 'n/a' cells")
    if not a.no_charts:
        print("Writing Charts workbook ...")
        build_charts(X, bk, rpath, cpath)
        print(f"  {CHARTS_NAME} written (external links to {RESULTS_NAME})")
    lines = summary_lines(X)
    with open(os.path.join(out, "run_summary.txt"), "w") as f:
        f.write(("SYNTHETIC DATA - ILLUSTRATIVE ONLY\n" if a.demo else "") + "\n".join(lines) + "\n\nTESTS\n" + "\n".join(f"[{t['Result']}] {t['Test']} - {t['Detail']}" for t in X.tests) + "\n")
    print("\n".join(lines))
    return X, bk

if __name__ == "__main__":
    main()
