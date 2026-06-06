#!/usr/bin/env python3
"""
visit_analysis.py
─────────────────
Analýza návštěvnosti a kapacity poboček 2025.

Vstupy : ../in/tables/VISITS_2025.csv
         kpis_grouped_2026.pkl
         export_specialiste.pkl
         ../vypocet_ir_2026/zdroje/Pobockova_profitabilita_4Q2025.xlsx
         ../vypocet_ir_2026/zdroje/report_od_pobocky_dbs_04_2026.xlsx
Výstup : report_navstevnost.html
"""

import os, sys, json, math, unicodedata, re, subprocess
import pandas as pd
import numpy as np

for _pkg in ['openpyxl']:
    try: __import__(_pkg)
    except ImportError:
        print(f"Instaluji {_pkg}…")
        subprocess.check_call([sys.executable, "-m", "pip", "install", _pkg])

# ─── Paths ─────────────────────────────────────────────────────────────────────
VISITS_PATH  = '../in/tables/VISITS_2025.csv'
KPIS_PATH    = 'kpis_grouped_2026.pkl'
SPEC_PATH    = 'export_specialiste.pkl'
PROF_PATH    = '../vypocet_ir_2026/zdroje/Pobockova_profitabilita_4Q2025.xlsx'
OD_PATH      = '../vypocet_ir_2026/zdroje/report_od_pobocky_dbs_04_2026.xlsx'
SALES_PATH   = '../in/tables/VYNOSY_NOVYCH_OBCHODU_BASED_ON_LTV.csv'
OUTPUT_FILE  = 'report_navstevnost.html'

# ─── Colors (shades of blue) ────────────────────────────────────────────────────
VISIT_TYPES = [
    ('online',   'Online schůzky',    '#1d4ed8', '#eff6ff'),
    ('fyzicka',  'Fyzické schůzky',   '#3b82f6', '#dbeafe'),
    ('bezhot',   'Bezhot. walkin',    '#7dd3fc', '#e0f2fe'),
    ('hotovost', 'Hotovostní walkin', '#bfdbfe', '#f0f9ff'),
]
TYPE_KEYS  = [t[0] for t in VISIT_TYPES]
TYPE_LABEL = {t[0]: t[1] for t in VISIT_TYPES}
TYPE_COLOR = {t[0]: t[2] for t in VISIT_TYPES}
TYPE_BG    = {t[0]: t[3] for t in VISIT_TYPES}

MONTH_NAMES = ['Led','Úno','Bře','Dub','Kvě','Čvn','Čvc','Srp','Zář','Říj','Lis','Pro']
WD_NAMES    = ['Po','Út','St','Čt','Pá','So','Ne']
MC_HOURS    = list(range(6, 22))   # 6..21 (16 hourly slots)

_OD_DAYS = [
    ('PONDELI','PO','Pondělí',False), ('UTERY','UT','Úterý',False),
    ('STREDA','ST','Středa',False),   ('CTVRTEK','CT','Čtvrtek',False),
    ('PATEK','PA','Pátek',False),     ('SOBOTA','SO','Sobota',True),
    ('NEDELE','NE','Neděle',True),
]

# ─── Capacity constants ─────────────────────────────────────────────────────────
WORKING_DAYS        = 252
WORK_MINS_DAY       = 450   # 7.5 h
TARGET_PORTFOLIO    = 1500
TARGET_MTGS_MIN     = 4
TARGET_MTGS_MAX     = 5
MEETING_MINS        = 45
WALKIN_SHORT_PCT    = 0.80; WALKIN_SHORT_MINS = 15
WALKIN_LONG_PCT     = 0.20; WALKIN_LONG_MINS  = 30
WALKIN_CONVERT_PCT  = 0.20   # % walkinů přeroste ve schůzku 45 min
WALKIN_AVG_MINS     = 15.0   # průměrná délka bezhotovostního walkin (min)
ABSENCE_RATE        = 0.229   # 22,9 % — nemoci, dovolená, školení (bankéři + BKP)

MC_ITERATIONS       = 1000   # počet Monte Carlo iterací

# ─── Branch format thresholds (same logic as main script) ──────────────────────
FORMAT_ORDER = ['flagship', 'medium', 'medium economy', 'small']
FORMAT_LABEL_PY = {
    'flagship':       'Flagship ≥25 FTE',
    'medium':         'Střední 10–25 FTE',
    'medium economy': 'Economy 5–10 FTE',
    'small':          'Malá <5 FTE',
}
FORMAT_BG_PY = {
    'flagship':       '#1e3a8a',
    'medium':         '#2563eb',
    'medium economy': '#7c3aed',
    'small':          '#64748b',
}

# ─── Staff position keys (normalized) ──────────────────────────────────────────
BANKER_COLS  = {'OSOBNI_BANKER_-_JUNIOR', 'OSOBNI_BANKER_-_MEDIOR', 'OSOBNI_BANKER_-_SENIOR'}
SERVICE_COL  = 'BANKER_KLIENTSKE_PECE_-_MEDIOR'
CASHIER_COL  = 'BANKER_KLIENTSKE_PECE_-_JUNIOR'
POSITION_LABELS = {
    'OSOBNI_BANKER_-_JUNIOR':          'OB Junior',
    'OSOBNI_BANKER_-_MEDIOR':          'OB Medior',
    'OSOBNI_BANKER_-_SENIOR':          'OB Senior',
    'BANKER_KLIENTSKE_PECE_-_JUNIOR':  'BKP Junior',
    'BANKER_KLIENTSKE_PECE_-_MEDIOR':  'BKP Medior',
    'PREMIER_BANKAR_-_MEDIOR':         'Premier Medior',
    'PREMIER_BANKAR_-_SENIOR':         'Premier Senior',
    'HYPOTECNI_SPECIALISTA_-_MEDIOR':  'Hypoteční spec.',
    'INVESTICNI_SPECIALISTA_-_MEDIOR': 'Investiční spec.',
    'POJISTOVACI_SPECIALISTA_-_MEDIOR':'Pojišťovací spec.',
}


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _nc(col):
    if not isinstance(col, str): col = str(col)
    col = ''.join(c for c in unicodedata.normalize('NFD', col)
                  if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', '_', col.upper().strip())

def _map_type(val):
    if not isinstance(val, str): return None
    v = ''.join(c for c in unicodedata.normalize('NFD', val.lower().replace('-',''))
                if unicodedata.category(c) != 'Mn')
    if 'online' in v:                                       return 'online'
    if 'bezhot' in v:                                       return 'bezhot'
    if 'hotov' in v:                                        return 'hotovost'
    if any(x in v for x in ('schuzk','fyzick','schu')):    return 'fyzicka'
    return None

def _odv(row, key): return str(row.get(key,'') or '').strip()
def _od_closed(v): return v in ('00:00','0:00','','nan','None','0')


# ─── Capacity calculation ───────────────────────────────────────────────────────

def _cap_model(online, fyzicka, bezhot, bankers, svc_fte, n_days, walkin_conv_pct=None):
    """Annual capacity metrics dict for given visit counts and staffing."""
    if bankers <= 0: return None
    if walkin_conv_pct is None: walkin_conv_pct = WALKIN_CONVERT_PCT
    eff       = 1.0 - ABSENCE_RATE          # effective presence factor (0.771)
    avail_ob  = bankers * n_days * WORK_MINS_DAY * eff
    avail_svc = svc_fte * n_days * WORK_MINS_DAY * eff

    online_mins  = online   * MEETING_MINS
    fyzicka_mins = fyzicka  * MEETING_MINS
    bezhot_base  = bezhot   * WALKIN_AVG_MINS
    bezhot_conv  = bezhot   * walkin_conv_pct * MEETING_MINS

    if svc_fte > 0:
        ob_used  = online_mins + fyzicka_mins + bezhot_conv
        svc_used = bezhot_base
    else:
        ob_used  = online_mins + fyzicka_mins + bezhot_base + bezhot_conv
        svc_used = 0.0

    return {
        'online_mins':  round(online_mins),
        'fyzicka_mins': round(fyzicka_mins),
        'bezhot_base':  round(bezhot_base),
        'bezhot_conv':  round(bezhot_conv),
        'ob_used':      round(ob_used),
        'avail_ob':     round(avail_ob),
        'util_ob':      round(ob_used / avail_ob * 100, 1) if avail_ob > 0 else 0,
        'svc_used':     round(svc_used),
        'avail_svc':    round(avail_svc),
        'util_svc':     round(svc_used / avail_svc * 100, 1) if avail_svc > 0 else None,
        'n_days':       n_days,
        'bankers':      bankers,
        'svc_fte':      svc_fte,
    }


# ─── Monte Carlo ───────────────────────────────────────────────────────────────

def run_monte_carlo(by_hour, bankers, svc_fte, n_iter=MC_ITERATIONS, seed=42):
    """
    Simulate n_iter random working days using Poisson hourly visit counts.

    For each hour slot (6–21), visits per type are drawn from Poisson(λ)
    where λ = recorded average visits per day in that hour.
    Capacity per hour: bankers × 60 min (OB), svc_fte × 60 min (service).

    Returns dict with per-hour overload probability, P50/P95 utilization,
    overall coverage %, and minimum bankers needed for 95 % coverage.
    """
    if bankers <= 0:
        return None

    # Lambda matrix: (16 hours, 3 types) = [online, fyzicka, bezhot]
    lam = np.array([
        [max(float(by_hour.get('online',  [0]*24)[h] or 0), 0),
         max(float(by_hour.get('fyzicka', [0]*24)[h] or 0), 0),
         max(float(by_hour.get('bezhot',  [0]*24)[h] or 0), 0)]
        for h in MC_HOURS
    ])  # (16, 3)

    rng = np.random.default_rng(seed)
    # Sample once: (n_iter, 16 hours, 3 types)
    samples = rng.poisson(lam[None, :, :].clip(0), size=(n_iter, len(MC_HOURS), 3))

    # Minutes of OB work per (iteration, hour)
    #   online: 45 min fully blocks 1 banker slot
    #   fyzicka: 45 min
    #   bezhot converted (10%): 45 min  → contribution = 0.1 × 45 = 4.5 min per walkin
    ob_needed  = (45 * samples[:, :, 0] +
                  45 * samples[:, :, 1] +
                  WALKIN_CONVERT_PCT * MEETING_MINS * samples[:, :, 2])  # (n_iter, 16)
    svc_needed_raw = WALKIN_AVG_MINS * samples[:, :, 2]   # (n_iter, 16) — original SVC demand

    eff = 1.0 - ABSENCE_RATE   # 0.771 — effective presence per hour

    if svc_fte <= 0:
        ob_needed    = ob_needed + svc_needed_raw   # OB-junior handles bezhot
        svc_needed   = svc_needed_raw               # keep for histograms (no BKP capacity)
        svc_overload = np.zeros((n_iter, len(MC_HOURS)), dtype=bool)
    else:
        svc_needed   = svc_needed_raw
        svc_overload = svc_needed > svc_fte * 60 * eff

    # Utilization of OB team at current bankers (absence-adjusted)
    ob_cap = bankers * 60.0 * eff
    util   = ob_needed / ob_cap           # (n_iter, 16)   can be > 1 = overloaded
    ob_overload = ob_needed > ob_cap
    any_overload = ob_overload | svc_overload   # (n_iter, 16)
    day_overload = any_overload.any(axis=1)     # (n_iter,)

    coverage_pct  = round(float(1.0 - day_overload.mean()) * 100, 1)
    overload_prob = [round(float(any_overload[:, i].mean()) * 100, 1)
                     for i in range(len(MC_HOURS))]
    p50_util = [round(float(np.percentile(util[:, i], 50)) * 100, 1)
                for i in range(len(MC_HOURS))]
    p95_util = [round(float(np.percentile(util[:, i], 95)) * 100, 1)
                for i in range(len(MC_HOURS))]

    # Find minimum bankers for 95 % day-coverage (try increments of 0.5)
    bankers_for_95 = None
    for extra in [0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0]:
        b = bankers + extra
        ov = (ob_needed > b * 60 * eff) | svc_overload
        cov = 1.0 - ov.any(axis=1).mean()
        if cov >= 0.95:
            bankers_for_95 = round(b, 1)
            break

    # ── FTE visualization data ──────────────────────────────────────────────────
    # Hourly P95 FTE demand: minutes / 60 → FTE-hours; capacity in FTE = bankers/svc_fte × eff
    p95_ob_fte  = [round(float(np.percentile(ob_needed[:, i],  95)) / 60, 3) for i in range(len(MC_HOURS))]
    p95_svc_fte = [round(float(np.percentile(svc_needed[:, i], 95)) / 60, 3) for i in range(len(MC_HOURS))]
    ob_cap_fte  = round(bankers * eff, 3)      # effective FTE capacity per hour
    svc_cap_fte = round(svc_fte * eff, 3)      # effective SVC FTE capacity per hour

    # Daily total FTE demand distribution (sum across hours, then /60 → banker-hours)
    ob_day  = ob_needed.sum(axis=1) / 60
    svc_day = svc_needed.sum(axis=1) / 60
    ob_cap_day  = round(bankers * WORK_MINS_DAY / 60 * eff, 2)   # effective banker-hours/day
    svc_cap_day = round(svc_fte  * WORK_MINS_DAY / 60 * eff, 2)
    ob_p95_day  = round(float(np.percentile(ob_day,  95)), 2)
    svc_p95_day = round(float(np.percentile(svc_day, 95)), 2)

    n_bins = 24
    ob_c,  ob_e  = np.histogram(ob_day.clip(0),  bins=n_bins)
    svc_c, svc_e = np.histogram(svc_day.clip(0), bins=n_bins)

    return {
        'coverage_pct':    coverage_pct,
        'overload_prob':   overload_prob,   # per hour [6..21]
        'p50_util':        p50_util,        # median utilization % per hour
        'p95_util':        p95_util,        # 95th-pct utilization % per hour
        'bankers_for_95':  bankers_for_95,
        'current_bankers': round(bankers, 1),
        'svc_fte':         round(svc_fte, 1),
        'n_iter':          n_iter,
        'lam_online':      [round(lam[i, 0], 2) for i in range(len(MC_HOURS))],
        'lam_fyzicka':     [round(lam[i, 1], 2) for i in range(len(MC_HOURS))],
        'lam_bezhot':      [round(lam[i, 2], 2) for i in range(len(MC_HOURS))],
        # FTE visualization
        'p95_ob_fte':    p95_ob_fte,
        'p95_svc_fte':   p95_svc_fte,
        'ob_cap_fte':    ob_cap_fte,
        'svc_cap_fte':   svc_cap_fte,
        'ob_hist':       ob_c.tolist(),
        'ob_edges':      [round(x, 2) for x in ob_e.tolist()],
        'svc_hist':      svc_c.tolist(),
        'svc_edges':     [round(x, 2) for x in svc_e.tolist()],
        'ob_cap_day':    ob_cap_day,
        'svc_cap_day':   svc_cap_day,
        'ob_p95_day':    ob_p95_day,
        'svc_p95_day':   svc_p95_day,
    }


# ─── Format & rooms ────────────────────────────────────────────────────────────

def compute_format(fte):
    if fte is None or fte <= 0: return None
    if fte >= 25: return 'flagship'
    if fte >= 10: return 'medium'
    if fte >= 5:  return 'medium economy'
    return 'small'

def compute_rooms(mc):
    """Recommend meeting rooms and service desks using P95 Poisson peak-hour estimate."""
    if mc is None: return None
    mtg_lam  = [mc['lam_online'][i] + mc['lam_fyzicka'][i] for i in range(len(MC_HOURS))]
    peak_mtg = max(mtg_lam) if mtg_lam else 0
    peak_beh = max(mc.get('lam_bezhot', [0])) if mc.get('lam_bezhot') else 0
    def p95(lam): return lam + 1.645 * math.sqrt(lam) if lam > 0 else 0
    pm = p95(peak_mtg); pb = p95(peak_beh)
    return {
        'meeting_rooms':   math.ceil(pm * 45 / 60) if pm > 0 else 0,
        'service_desks':   math.ceil(pb * WALKIN_AVG_MINS / 60) if pb > 0 else 0,
        'peak_mtg_lam':    round(peak_mtg, 2),
        'p95_mtg':         round(pm, 1),
        'delta_mtg':       round(pm - peak_mtg, 1),   # bezpečnostní rezerva nad λ
        'peak_bezhot_lam': round(peak_beh, 2),
        'p95_bezhot':      round(pb, 1),
        'delta_bezhot':    round(pb - peak_beh, 1),
    }


# ─── Data loading ──────────────────────────────────────────────────────────────

def load_visits():
    print(f"📂 Návštěvy: {VISITS_PATH}")
    if not os.path.exists(VISITS_PATH):
        print(f"❌ Soubor nenalezen: {VISITS_PATH}", file=sys.stderr); sys.exit(1)
    df = pd.read_csv(VISITS_PATH, low_memory=False)
    df.columns = [_nc(c) for c in df.columns]
    seg_c = next((c for c in df.columns if c == 'CLIENT_SEGMENT'), None)
    if seg_c:
        n_before = len(df)
        df = df[df[seg_c].isin(['MM', 'MA'])]
        print(f"   CLIENT_SEGMENT filter MM/MA: {len(df):,} z {n_before:,} řádků")
    print(f"   {len(df):,} řádků · sloupce: {list(df.columns)}")
    return df


def load_kpis():
    """branch_id → {fte, pocet_klientu?, name?}"""
    out = {}
    if not os.path.exists(KPIS_PATH): print(f"⚠️  kpis nenalezen"); return out
    try:
        kp = pd.read_pickle(KPIS_PATH)
        kp.columns = [_nc(c) for c in kp.columns]
        id_c  = next((c for c in ['POBOCKA_ID','BRANCH_CODE','ID_POBOCKY'] if c in kp.columns), None)
        nm_c  = next((c for c in ['POBOCKA_NAZEV','BRANCH_NAME','NAZEV'] if c in kp.columns), None)
        fte_c = next((c for c in ['FTE','POCET_BANKERU'] if c in kp.columns), None)
        cli_c = next((c for c in ['POCET_KLIENTU','PRIMARNI_KLIENTI','AKTIVNI_KLIENTI']
                      if c in kp.columns), None)
        if id_c is None: return out
        for _, row in kp.iterrows():
            bid = pd.to_numeric(row[id_c], errors='coerce')
            if pd.isna(bid): continue
            bid = int(bid)
            out[bid] = {
                'name': str(row[nm_c]) if nm_c else None,
                'fte':  float(row[fte_c]) if fte_c and pd.notna(row.get(fte_c)) else None,
                'pocet_klientu': (int(row[cli_c]) if cli_c and pd.notna(row.get(cli_c))
                                  and float(row[cli_c]) > 0 else None),
            }
        print(f"   kpis: {len(out)} poboček · fte={'✓' if fte_c else '✗'} · "
              f"klienti={'✓' if cli_c else '✗'}")
    except Exception as e:
        print(f"⚠️  kpis: {e}")
    return out


def load_profitabilita():
    """branch_id → pocet_klientu; also returns branch_id → fte as second dict."""
    poc_kli, fte_map = {}, {}
    if not os.path.exists(PROF_PATH): print(f"⚠️  Profitabilita nenalezena"); return poc_kli, fte_map
    try:
        pf = pd.read_excel(PROF_PATH, header=2, usecols='A:AB')
        pf.columns = [_nc(c) for c in pf.columns]
        id_c  = next((c for c in ['ID_POBOCKY','BRANCH_CODE','ID'] if c in pf.columns), None)
        cli_c = next((c for c in ['POCET_KLIENTU','PRIMARNI_KLIENTI','AKTIVNI_KLIENTI']
                      if c in pf.columns), None)
        fte_c = next((c for c in ['FTE','CELKOVA_FTE','TOTAL_FTE'] if c in pf.columns), None)
        if id_c:
            for _, row in pf.iterrows():
                bid = pd.to_numeric(row[id_c], errors='coerce')
                if pd.isna(bid): continue
                bid = int(bid)
                if cli_c:
                    v = pd.to_numeric(row[cli_c], errors='coerce')
                    if pd.notna(v) and v > 0: poc_kli[bid] = int(v)
                if fte_c:
                    f = pd.to_numeric(row[fte_c], errors='coerce')
                    if pd.notna(f) and f > 0: fte_map[bid] = round(float(f), 1)
        print(f"   Profitabilita: {len(poc_kli)} poboček · "
              f"klienti={'✓' if cli_c else '✗'} · FTE={'✓' if fte_c else '✗'}")
    except Exception as e:
        print(f"⚠️  Profitabilita: {e}")
    return poc_kli, fte_map


def load_specialiste():
    """branch_id → {name, bankers, svc_fte, has_svc, positions}"""
    out = {}
    if not os.path.exists(SPEC_PATH): print(f"⚠️  Specialisté nenalezeni"); return out
    try:
        sp = pd.read_pickle(SPEC_PATH)
        sp.columns = [_nc(c) for c in sp.columns]
        bid_c = next((c for c in ['BRANCH_ID','BRANCH_CODE','ID'] if c in sp.columns), None)
        nm_c  = next((c for c in ['BRANCH_NAME','POBOCKA_NAZEV','NAZEV'] if c in sp.columns), None)
        id_cols = {'BRANCH_ID','BRANCH_CODE','BRANCH_NAME','POBOCKA_NAZEV',
                   'GPS_X','GPS_Y','EVIDENCNI_STAV','ID','NAZEV'}
        pos_cols = [c for c in sp.columns if c not in id_cols]
        for c in pos_cols:
            sp[c] = pd.to_numeric(sp[c], errors='coerce').fillna(0)

        for _, row in sp.iterrows():
            bid = pd.to_numeric(row[bid_c] if bid_c else None, errors='coerce')
            if pd.isna(bid): continue
            bid = int(bid)
            bankers     = sum(float(row.get(c, 0) or 0) for c in pos_cols if c in BANKER_COLS)
            svc_fte     = float(row.get(SERVICE_COL, 0) or 0)
            cashier_fte = float(row.get(CASHIER_COL, 0) or 0)
            positions = {}
            for c in pos_cols:
                v = float(row.get(c, 0) or 0)
                if v > 0:
                    lbl = POSITION_LABELS.get(c, c.replace('_',' ').title())
                    positions[lbl] = round(v, 1)
            total_fte = sum(float(row.get(c, 0) or 0) for c in pos_cols)
            out[bid] = {
                'name':        str(row[nm_c]) if nm_c and pd.notna(row.get(nm_c,'')) else None,
                'bankers':     round(bankers, 1),
                'svc_fte':     round(svc_fte, 1),
                'has_svc':     svc_fte > 0,
                'cashier_fte': round(cashier_fte, 1),
                'has_cash':    cashier_fte > 0,
                'positions':   positions,
                'total_fte':   round(total_fte, 1),
            }
        print(f"   Specialisté: {len(out)} poboček")
    except Exception as e:
        print(f"⚠️  Specialisté: {e}")
    return out


def load_oteviraci():
    out = {}
    if not os.path.exists(OD_PATH): print(f"⚠️  Ot.doba nenalezena"); return out
    try:
        od = pd.read_excel(OD_PATH, dtype=str)
        if 'KOD_POBOCKY' not in od.columns: print("⚠️  OD: chybí KOD_POBOCKY"); return out
        for _, row in od.iterrows():
            bid = pd.to_numeric(row.get('KOD_POBOCKY',''), errors='coerce')
            if pd.isna(bid): continue
            bid = int(bid)
            days = []; is_vikend = False
            for dn, ds, dcz, wknd in _OD_DAYS:
                tot = _odv(row, ds); closed = _od_closed(tot)
                dop = (f"{_odv(row,f'{dn}_DOP._OD')}–{_odv(row,f'{dn}_DOP._DO')}"
                       if not _od_closed(_odv(row,f'{dn}_DOP._OD')) else '')
                odp = (f"{_odv(row,f'{dn}_ODP._OD')}–{_odv(row,f'{dn}_ODP._DO')}"
                       if not _od_closed(_odv(row,f'{dn}_ODP._OD')) else '')
                if wknd and not closed: is_vikend = True
                days.append({'lbl':dcz,'wknd':wknd,'closed':closed,
                             'dop':dop,'odp':odp,'tot':tot if not closed else ''})
            try: ph = float(str(row.get('PH',0) or 0).replace(',','.'))
            except: ph = 0.0
            wd_open = sum(1 for d in days if not d['wknd'] and not d['closed'])
            we_open = sum(1 for d in days if d['wknd']     and not d['closed'])
            annual_open_days = wd_open * 52 + we_open * 52
            out[bid] = {'is_vikend':is_vikend,'ph_tyden':ph,'od_days':days,
                        'annual_open_days': annual_open_days}
        print(f"   Ot.doba: {len(out)} poboček")
    except Exception as e:
        print(f"⚠️  Ot.doba: {e}")
    return out


def load_sales():
    """branch_id → {products: [{name, pocet, objem}], total_pocet, total_objem}"""
    out = {}
    if not os.path.exists(SALES_PATH):
        print(f"⚠️  Prodejní data nenalezena: {SALES_PATH}")
        return out
    try:
        sl = pd.read_csv(SALES_PATH, low_memory=False)
        sl.columns = [_nc(c) for c in sl.columns]
        bid_c  = next((c for c in ['BRANCH_CODE', 'POBOCKA_ID', 'ID_POBOCKY'] if c in sl.columns), None)
        prod_c = next((c for c in ['OKOPRODG_SOURCE_ID', 'PRODUKT', 'PRODUCT'] if c in sl.columns), None)
        cnt_c  = next((c for c in ['POCET_PRODEJU', 'POCET'] if c in sl.columns), None)
        vol_c  = next((c for c in ['OBJEM_VYNOSU_CZK', 'OBJEM', 'REVENUE'] if c in sl.columns), None)
        if bid_c is None or prod_c is None:
            print(f"⚠️  Prodejní data: nelze najít sloupce — nalezeno: {list(sl.columns[:10])}")
            return out
        sl[bid_c] = pd.to_numeric(sl[bid_c], errors='coerce')
        sl = sl.dropna(subset=[bid_c])
        sl[bid_c] = sl[bid_c].astype(int)
        agg = {}
        if cnt_c: agg[cnt_c] = 'sum'
        if vol_c: agg[vol_c] = 'sum'
        if agg:
            grp = sl.groupby([bid_c, prod_c]).agg(agg).reset_index()
        else:
            grp = sl.groupby([bid_c, prod_c]).size().reset_index(name='_cnt')
            cnt_c = '_cnt'
        for bid, sub in grp.groupby(bid_c):
            sort_col = cnt_c if cnt_c in sub.columns else prod_c
            products = []
            for _, row in sub.sort_values(sort_col, ascending=False).iterrows():
                p = {'name': str(row[prod_c])}
                if cnt_c and cnt_c in row.index:
                    p['pocet'] = int(row[cnt_c]) if pd.notna(row[cnt_c]) else 0
                if vol_c and vol_c in row.index:
                    p['objem'] = int(round(float(row[vol_c]))) if pd.notna(row[vol_c]) else 0
                products.append(p)
            out[bid] = {
                'products':    products,
                'total_pocet': sum(p.get('pocet', 0) for p in products),
                'total_objem': sum(p.get('objem', 0) for p in products),
            }
        print(f"   Prodeje: {len(out)} poboček · {len(sl)} řádků · "
              f"produkt={prod_c} · počet={'✓' if cnt_c else '✗'} · objem={'✓' if vol_c else '✗'}")
    except Exception as e:
        print(f"⚠️  Prodejní data: {e}")
    return out


# ─── Build data ────────────────────────────────────────────────────────────────

def build_data(df, kpis, prof_kli, prof_fte, spec, od, sales=None):
    bid_c  = next((c for c in ['BRANCH_ID','BRANCH_CODE','POBOCKA'] if c in df.columns), None)
    bname_c= next((c for c in ['BRANCH_NAME','POBOCKA_NAZEV'] if c in df.columns), None)
    att_c  = next((c for c in ['ATTENDANCE_TYPE','VISIT_TYPE','TYP_NAVSTEVY'] if c in df.columns), None)
    date_c = next((c for c in ['VISIT_DATE','DATE','DATUM'] if c in df.columns), None)
    time_c = next((c for c in ['VISIT_TIME','TIME','CAS'] if c in df.columns), None)
    if bid_c is None: print("❌ ID pobočky nenalezeno"); sys.exit(1)

    df = df.copy()
    df[bid_c] = pd.to_numeric(df[bid_c], errors='coerce')
    df = df.dropna(subset=[bid_c]); df[bid_c] = df[bid_c].astype(int)

    if att_c:
        df['_t'] = df[att_c].apply(_map_type)
        print(f"   Typy: {df['_t'].value_counts().to_dict()}")
    else:
        df['_t'] = None

    has_date = date_c is not None
    if has_date:
        df['_dt']  = pd.to_datetime(df[date_c], errors='coerce')
        df['_mon'] = df['_dt'].dt.month
        df['_wd']  = df['_dt'].dt.weekday
    else:
        df['_mon'] = None; df['_wd'] = None

    has_time = time_c is not None
    if has_time:
        df['_hr'] = pd.to_numeric(
            df[time_c].astype(str).str.split(':').str[0], errors='coerce')
    else:
        df['_hr'] = None

    result = {}
    branches = sorted(df[bid_c].unique())
    print(f"   {len(branches)} poboček — počítám statistiky + Monte Carlo…")

    for bid in branches:
        vb = df[df[bid_c] == bid]
        k  = kpis.get(bid, {})
        s  = spec.get(bid, {})
        o  = od.get(bid, {})
        sl = (sales or {}).get(bid)

        name = (s.get('name') or k.get('name') or
                (str(vb[bname_c].iloc[0]) if bname_c else None) or f"Pobočka {bid}")

        total   = len(vb)
        by_type = {k2: int((vb['_t'] == k2).sum()) for k2 in TYPE_KEYS}
        unknown = int(vb['_t'].isna().sum())

        by_month   = {k2: [0]*12 for k2 in TYPE_KEYS}
        by_weekday = {k2: [0]*7  for k2 in TYPE_KEYS}
        by_hour    = {k2: [0]*24 for k2 in TYPE_KEYS}
        heatmap    = [[0]*24 for _ in range(7)]
        n_days = 1

        n_days_wd = [0] * 7   # unique dates per weekday
        if has_date:
            n_days = max(int(vb['_dt'].dt.date.nunique()), 1)
            for wd2 in range(7):
                mask = vb['_wd'] == wd2
                if mask.any():
                    n_days_wd[wd2] = int(vb.loc[mask, '_dt'].dt.date.nunique())
            for k2 in TYPE_KEYS:
                sub = vb[vb['_t'] == k2]
                if not sub.empty:
                    by_month[k2]   = [int(v) for v in sub['_mon'].value_counts()
                                      .reindex(range(1,13), fill_value=0)]
                    by_weekday[k2] = [int(v) for v in sub['_wd'].value_counts()
                                      .reindex(range(7), fill_value=0)]

        if has_time and has_date:
            for k2 in TYPE_KEYS:
                sub = vb[vb['_t'] == k2]
                if not sub.empty:
                    hc = sub['_hr'].dropna().astype(int).value_counts().reindex(range(24), fill_value=0)
                    by_hour[k2] = [round(float(v)/n_days, 2) for v in hc]
            # Heatmap (all types combined)
            valid = vb.dropna(subset=['_wd','_hr']).copy()
            if not valid.empty:
                valid['_wd'] = valid['_wd'].astype(int); valid['_hr'] = valid['_hr'].astype(int)
                hm = (valid.groupby(['_wd','_hr']).size().unstack(fill_value=0)
                      .reindex(index=range(7), columns=range(24), fill_value=0))
                heatmap = [[int(hm.loc[wd2, hr]) for hr in range(24)] for wd2 in range(7)]

        # Normalized heatmap: avg visits per day of week × hour
        heatmap_avg = [
            [round(heatmap[w][h] / max(n_days_wd[w], 1), 2) for h in range(24)]
            for w in range(7)
        ]

        # Staff
        bankers     = float(s.get('bankers', 0) or 0)
        svc_fte     = float(s.get('svc_fte', 0) or 0)
        has_svc     = bool(s.get('has_svc', False))
        cashier_fte = float(s.get('cashier_fte', 0) or 0)
        has_cash    = bool(s.get('has_cash', False))
        positions   = s.get('positions', {})
        # FTE priority: profitabilita xlsx → kpis pkl → specialiste total (fallback)
        fte = prof_fte.get(bid) or k.get('fte') or s.get('total_fte')

        # Client count: profitabilita first, then kpis fallback
        poc_kli = prof_kli.get(bid) or k.get('pocet_klientu') or 0

        # Annual open days for this branch (from opening hours; fallback = generic WORKING_DAYS)
        annual_open_days = o.get('annual_open_days', WORKING_DAYS) or WORKING_DAYS

        # Metrics
        portfolio_pb = round(poc_kli / bankers) if bankers > 0 and poc_kli > 0 else None
        total_mtgs   = by_type['online'] + by_type['fyzicka']
        mtgs_pb_day  = round(total_mtgs / bankers / n_days, 1) if bankers > 0 and n_days > 0 else None

        # Capacity Model 1 (real visits)
        cap1 = _cap_model(by_type['online'], by_type['fyzicka'], by_type['bezhot'],
                          bankers, svc_fte, n_days)

        # Capacity Model 4 (35 % walkin → meeting conversion, same base data as M1)
        cap_m4 = _cap_model(by_type['online'], by_type['fyzicka'], by_type['bezhot'],
                            bankers, svc_fte, n_days, walkin_conv_pct=0.35)

        # Capacity Model 2 (client-based) — uses branch-specific annual open days
        cap2 = None
        total_excl_hot = sum(by_type[k2] for k2 in ['online','fyzicka','bezhot'])
        if poc_kli > 0 and total_excl_hot > 0 and bankers > 0:
            r_on  = by_type['online']  / total_excl_hot
            r_fyz = by_type['fyzicka'] / total_excl_hot
            r_beh = by_type['bezhot']  / total_excl_hot
            cap2  = _cap_model(poc_kli * r_on, poc_kli * r_fyz, poc_kli * r_beh,
                               bankers, svc_fte, annual_open_days)

        # Annual "1x per client" scenario — uses branch-specific annual open days
        annual_1x = None
        if poc_kli > 0 and bankers > 0:
            eff = 1.0 - ABSENCE_RATE
            avail_ob_yr = bankers * annual_open_days * WORK_MINS_DAY * eff
            mins_needed = poc_kli * MEETING_MINS
            annual_1x = {
                'poc_kli':        poc_kli,
                'mins_needed':    round(mins_needed),
                'avail_ob':       round(avail_ob_yr),
                'util_ob':        round(mins_needed / avail_ob_yr * 100, 1) if avail_ob_yr > 0 else 0,
                'mtgs_pb_day':    round(poc_kli / bankers / annual_open_days, 1),
                'bankers':        round(bankers, 1),
                'bankers_needed': round(mins_needed / (annual_open_days * WORK_MINS_DAY * eff), 1),
                'open_days':      annual_open_days,
            }

        # Monte Carlo (only when hourly data is available)
        mc = None
        if has_time and has_date and bankers > 0:
            mc = run_monte_carlo(by_hour, bankers, svc_fte)

        mc2 = None
        if mc is not None:
            mc2 = run_monte_carlo(by_hour, bankers, svc_fte, n_iter=2000, seed=42)

        # MC boost: +20% all visit lambdas
        mc_boost = None
        if has_time and has_date and bankers > 0:
            by_hour_boost = {k2: [v * 1.2 for v in vals] for k2, vals in by_hour.items()}
            mc_boost = run_monte_carlo(by_hour_boost, bankers, svc_fte)

        branch_format = compute_format(fte)
        rooms         = compute_rooms(mc)

        result[str(bid)] = {
            'name':       name,
            'fte':        fte,
            'bankers':    bankers,
            'svc_fte':    svc_fte,
            'has_svc':      has_svc,
            'cashier_fte':  cashier_fte,
            'has_cash':     has_cash,
            'positions':    positions,
            'poc_kli':    poc_kli,
            'total':      total,
            'by_type':    by_type,
            'unknown':    unknown,
            'by_month':   by_month,
            'by_weekday': by_weekday,
            'by_hour':       by_hour,
            'heatmap':       heatmap,
            'heatmap_avg':   heatmap_avg,
            'n_days_wd':     n_days_wd,
            'has_time':      has_time and has_date,
            'n_days':        n_days,
            'annual_open_days': annual_open_days,
            'portfolio_pb': portfolio_pb,
            'mtgs_pb_day':  mtgs_pb_day,
            'cap1':       cap1,
            'cap2':       cap2,
            'cap_m4':     cap_m4,
            'annual_1x':  annual_1x,
            'mc':           mc,
            'mc2':          mc2,
            'mc_boost':     mc_boost,
            'branch_format': branch_format,
            'rooms':         rooms,
            'is_vikend':  o.get('is_vikend', False),
            'ph_tyden':   o.get('ph_tyden',  0.0),
            'od_days':    o.get('od_days',   []),
            'sales':      sl,
        }

    order = sorted(result.keys(), key=lambda x: result[x]['name'])
    return result, order, att_c is not None


# ─── Benchmarks ────────────────────────────────────────────────────────────────

def compute_benchmarks(result):
    """Add benchmark comparisons (network + same-format medians) to each branch."""
    def _med(vals):
        vals = sorted(float(v) for v in vals if v is not None and not math.isnan(float(v)))
        if not vals: return None
        n = len(vals)
        return round((vals[n//2-1]+vals[n//2])/2 if n%2==0 else vals[n//2], 1)

    def _own(d):
        tot = d['total']; nd = max(d['n_days'], 1); bt = d['by_type']
        cap_util = d['cap1']['util_ob'] if d.get('cap1') else None
        return {
            'visits_pd':    round(tot / nd, 1),
            'mtgs_pb_day':  d.get('mtgs_pb_day'),
            'cap_util_ob':  cap_util,
            'online_pct':   round(bt.get('online',   0) / max(tot, 1) * 100, 1),
            'fyzicka_pct':  round(bt.get('fyzicka',  0) / max(tot, 1) * 100, 1),
            'bezhot_pct':   round(bt.get('bezhot',   0) / max(tot, 1) * 100, 1),
            'hotovost_pct': round(bt.get('hotovost', 0) / max(tot, 1) * 100, 1),
        }

    BKEYS = ['visits_pd', 'mtgs_pb_day', 'cap_util_ob',
             'online_pct', 'fyzicka_pct', 'bezhot_pct', 'hotovost_pct']
    own = {bid: _own(d) for bid, d in result.items()}
    network = {k: _med([m[k] for m in own.values()]) for k in BKEYS}

    fmt_groups = {}
    for bid, d in result.items():
        f = d.get('branch_format')
        if f: fmt_groups.setdefault(f, []).append(bid)
    fmt_medians = {f: {k: _med([own[bid][k] for bid in ids]) for k in BKEYS}
                   for f, ids in fmt_groups.items()}

    for bid, d in result.items():
        d['benchmark'] = {
            'own':     own[bid],
            'network': network,
            'format':  fmt_medians.get(d.get('branch_format')),
        }
    return result


# ─── HTML ───────────────────────────────────────────────────────────────────────

def render_html(data, order, has_type_col):
    data_js  = json.dumps(data,  ensure_ascii=False, separators=(',',':'))
    order_js = json.dumps(order, ensure_ascii=False)
    types_js = json.dumps(
        [{'key':k,'label':TYPE_LABEL[k],'color':TYPE_COLOR[k]} for k in TYPE_KEYS],
        ensure_ascii=False)
    mc_hours_js = json.dumps(MC_HOURS)
    consts_js = json.dumps({
        'TARGET_PORTFOLIO': TARGET_PORTFOLIO,
        'TARGET_MTGS_MIN':  TARGET_MTGS_MIN,
        'TARGET_MTGS_MAX':  TARGET_MTGS_MAX,
        'MEETING_MINS':     MEETING_MINS,
        'WALKIN_SHORT_PCT': WALKIN_SHORT_PCT,
        'WALKIN_SHORT_MINS':WALKIN_SHORT_MINS,
        'WALKIN_LONG_PCT':  WALKIN_LONG_PCT,
        'WALKIN_LONG_MINS': WALKIN_LONG_MINS,
        'WALKIN_AVG_MINS':  WALKIN_AVG_MINS,
        'WALKIN_CONVERT_PCT': WALKIN_CONVERT_PCT,
        'WORK_MINS_DAY':    WORK_MINS_DAY,
        'WORKING_DAYS':     WORKING_DAYS,
        'MC_ITERATIONS':    MC_ITERATIONS,
        'ABSENCE_RATE':     ABSENCE_RATE,
    })
    no_type_warn = ('' if has_type_col else
        '<div class="warn">⚠️ Sloupec ATTENDANCE_TYPE nenalezen — typové grafy nejsou dostupné.</div>')

    return f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Analýza návštěvnosti 2025</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',system-ui,sans-serif;background:#f0f4ff;color:#1e293b;font-size:15px;line-height:1.5;}}
.wrap{{max-width:1100px;margin:0 auto;padding:20px 16px;}}
h1{{font-size:1.4rem;font-weight:800;color:#1e3a8a;margin-bottom:3px;}}
.subtitle{{font-size:.82rem;color:#94a3b8;margin-bottom:20px;}}
.warn{{background:#fef9c3;border:1px solid #fde047;border-radius:8px;padding:10px 14px;font-size:.82rem;color:#713f12;margin-bottom:14px;}}
.sw{{position:relative;margin-bottom:20px;}}
.sw input{{width:100%;padding:11px 14px 11px 40px;border:1.5px solid #bfdbfe;border-radius:10px;font-size:16px;background:#fff;outline:none;transition:border .15s;-webkit-tap-highlight-color:transparent;}}
.sw input:focus{{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.15);}}
.sw .ico{{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#93c5fd;pointer-events:none;}}
.bl{{display:none;position:absolute;top:calc(100% + 4px);left:0;right:0;background:#fff;border:1.5px solid #bfdbfe;border-radius:10px;max-height:280px;overflow-y:auto;z-index:100;box-shadow:0 8px 32px rgba(30,58,138,.12);-webkit-overflow-scrolling:touch;}}
.bl.open{{display:block;}}
.bi{{padding:9px 14px;cursor:pointer;font-size:.88rem;border-bottom:1px solid #eff6ff;}}
.bi:hover,.bi.sel{{background:#eff6ff;color:#1d4ed8;font-weight:600;}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;}}
.grid3{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px;}}
.grid4{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px;}}
@media(max-width:700px){{.grid3,.grid4{{grid-template-columns:1fr 1fr;}}}}
@media(max-width:420px){{.grid2,.grid3,.grid4{{grid-template-columns:1fr;}}}}
.card{{background:#fff;border-radius:12px;border:1.5px solid #dbeafe;padding:13px 15px;}}
.card.ok{{border-color:#bbf7d0;background:#f0fdf4;}}
.card.warn-c{{border-color:#fca5a5;background:#fff5f5;}}
.cl{{font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#64748b;margin-bottom:3px;}}
.cv{{font-size:1.35rem;font-weight:800;line-height:1.1;color:#1e3a8a;}}
.cs{{font-size:.72rem;color:#94a3b8;margin-top:2px;}}
.sec{{background:#fff;border-radius:12px;border:1.5px solid #dbeafe;padding:16px;margin-bottom:14px;}}
.st{{font-size:.76rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#475569;margin-bottom:12px;}}
/* bars */
.bars{{display:flex;gap:3px;align-items:flex-end;}}
.bw{{display:flex;flex-direction:column;align-items:center;gap:2px;flex:1;min-width:0;}}
.bs{{display:flex;flex-direction:column-reverse;width:100%;border-radius:4px 4px 0 0;overflow:hidden;}}
.bseg{{width:100%;flex-shrink:0;}}
.blbl{{font-size:.58rem;color:#94a3b8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%;text-align:center;}}
.bnum{{font-size:.58rem;color:#475569;font-weight:600;text-align:center;}}
.leg{{display:flex;flex-wrap:wrap;gap:9px;margin-top:10px;}}
.li{{display:flex;align-items:center;gap:5px;font-size:.74rem;color:#475569;}}
.ld{{width:10px;height:10px;border-radius:3px;flex-shrink:0;}}
/* heatmap */
.hmt{{border-collapse:collapse;width:100%;font-size:.63rem;}}
.hmt td,.hmt th{{padding:2px 2px;text-align:center;border-radius:2px;min-width:20px;}}
.hmt th{{color:#94a3b8;font-weight:600;font-size:.6rem;}}
/* capacity bar */
.cbwrap{{position:relative;height:26px;border-radius:6px;overflow:hidden;background:#e0e7ef;margin:6px 0 2px;}}
.cbfill{{height:100%;border-radius:6px;display:flex;align-items:center;padding:0 7px;font-size:.7rem;font-weight:700;color:#fff;white-space:nowrap;overflow:hidden;min-width:40px;}}
/* MC bars */
.mcbar-wrap{{display:flex;gap:3px;align-items:flex-end;margin-bottom:8px;}}
.mcbar-col{{display:flex;flex-direction:column;align-items:center;gap:2px;flex:1;min-width:0;}}
.mcbar-inner{{width:100%;border-radius:3px 3px 0 0;overflow:hidden;position:relative;}}
/* OD table */
.odt{{width:100%;border-collapse:collapse;font-size:.83rem;}}
.odt td,.odt th{{padding:5px 9px;}}
.odt th{{font-size:.67rem;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#94a3b8;border-bottom:1px solid #dbeafe;text-align:left;}}
.od-wknd{{background:#fffbf0!important;}}
/* positions table */
.post{{width:100%;border-collapse:collapse;font-size:.82rem;}}
.post td{{padding:4px 9px;border-bottom:1px solid #f0f4ff;}}
.post tr:last-child td{{border-bottom:none;}}
.badge{{display:inline-block;border-radius:10px;padding:2px 8px;font-size:.68rem;font-weight:700;}}
/* methodology */
.meth{{background:#f8faff;border:1px solid #dbeafe;border-radius:10px;padding:14px 16px;font-size:.8rem;color:#475569;}}
.meth h4{{font-size:.78rem;font-weight:700;color:#1e3a8a;margin:10px 0 4px;}}
.meth h4:first-child{{margin-top:0;}}
.meth ul{{padding-left:18px;}}
.meth li{{margin-bottom:3px;}}
details summary{{cursor:pointer;font-size:.76rem;font-weight:700;color:#2563eb;padding:8px 0;user-select:none;}}
.placeholder{{text-align:center;color:#94a3b8;padding:50px 0;font-size:.9rem;}}
/* Kapacitní karta */
.kap-card{{background:#fff;border:2px solid #1e3a8a;border-radius:14px;margin-bottom:14px;overflow:hidden;}}
.kap-title{{background:#1e3a8a;color:#fff;padding:11px 16px;font-size:.82rem;font-weight:700;letter-spacing:.3px;}}
.kap-model{{padding:14px 16px;border-top:1px solid #eff6ff;}}
.kap-model-hd{{font-size:.84rem;font-weight:700;color:#1d4ed8;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #eff6ff;}}
/* Day bar */
.daybar{{height:30px;display:flex;border-radius:6px;overflow:hidden;margin:8px 0;}}
.daybar-seg{{height:100%;display:flex;align-items:center;justify-content:center;font-size:.62rem;font-weight:700;overflow:hidden;white-space:nowrap;padding:0 3px;flex-shrink:0;}}
/* Staff */
.staff-big-n{{font-size:2rem;font-weight:800;line-height:1;}}
.staff-big-lbl{{font-size:.75rem;color:#475569;font-weight:600;}}
/* Visit card */
.v-card{{border:1.5px solid #1e3a8a;border-radius:14px;overflow:hidden;margin-bottom:14px;}}
.v-title{{background:#1e3a8a;color:#fff;padding:10px 16px;font-size:.8rem;font-weight:700;}}
.v-body{{padding:14px;}}
/* Capacity gauge ring */
.cap-ring{{width:76px;height:76px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0;}}
.cap-ring-in{{width:52px;height:52px;border-radius:50%;background:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;line-height:1.1;}}
.cap-gauge-row{{display:flex;align-items:flex-start;gap:14px;margin-bottom:10px;}}
.cap-gauge-meta{{flex:1;min-width:0;}}
/* MC FTE charts */
.mc-chart-wrap{{border:1px solid #dbeafe;border-radius:6px;padding:6px;background:#fff;margin-bottom:4px;}}
.mc-chart-title{{font-size:.68rem;font-weight:700;color:#1e3a8a;margin:8px 0 3px;}}
.mc-charts-row{{display:grid;grid-template-columns:1fr 1fr;gap:10px;}}
@media(max-width:560px){{.mc-charts-row{{grid-template-columns:1fr;}}}}
/* Visit toggle */
.vtog{{display:flex;background:#f1f5f9;border-radius:6px;padding:2px;gap:2px;}}
.vt-btn{{padding:4px 10px;border:none;background:transparent;border-radius:4px;font-size:.72rem;font-weight:600;color:#475569;cursor:pointer;white-space:nowrap;}}
.vt-btn.vt-on{{background:#fff;color:#1d4ed8;box-shadow:0 1px 3px rgba(0,0,0,.12);}}
/* Capacity sub-panels */
.cap-sub{{background:#f8faff;border-radius:10px;padding:12px;border:1px solid #dbeafe;}}
.cap-sub-hd{{font-size:.72rem;font-weight:700;margin-bottom:8px;}}
.hbar-row{{display:flex;align-items:center;gap:6px;margin-bottom:5px;}}
.hbar-lbl{{width:72px;font-size:.63rem;color:#475569;text-align:right;flex-shrink:0;}}
.hbar-track{{flex:1;height:16px;background:#f1f5f9;border-radius:3px;overflow:hidden;}}
.hbar-fill{{height:100%;border-radius:3px;display:flex;align-items:center;padding-left:5px;min-width:0;transition:width .2s;}}
.hbar-val{{font-size:.62rem;font-weight:700;white-space:nowrap;}}
.gap90{{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:8px 12px;margin-top:10px;font-size:.72rem;}}
.mc-2col{{display:grid;grid-template-columns:1fr 1fr;gap:12px;}}
@media(max-width:600px){{.mc-2col{{grid-template-columns:1fr;}}}}
/* Ranking table */
.rank-tbl{{width:100%;border-collapse:collapse;font-size:.82rem;}}
.rank-tbl th{{font-size:.64rem;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#94a3b8;border-bottom:2px solid #dbeafe;padding:7px 10px;text-align:left;white-space:nowrap;}}
.rank-tbl td{{padding:7px 10px;border-bottom:1px solid #f0f4ff;vertical-align:middle;}}
.rank-tbl tbody tr{{cursor:pointer;}}
.rank-tbl tbody tr:hover{{background:#f8faff;}}
.rank-score{{display:inline-flex;align-items:center;justify-content:center;border-radius:6px;padding:2px 8px;font-weight:700;font-size:.85rem;min-width:46px;}}
</style>
</head>
<body>
<div class="wrap">
  <h1>Analýza návštěvnosti poboček 2025</h1>
  <div class="subtitle">Zdroj: {VISITS_PATH}</div>
  {no_type_warn}
  <div class="sw" id="sw">
    <svg class="ico" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
      <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
    <input type="text" id="si" placeholder="Vyhledat pobočku…" autocomplete="off" autocorrect="off" spellcheck="false">
    <div class="bl" id="bl"></div>
  </div>
  <div id="mc"><div class="placeholder">← Vyberte pobočku výše</div></div>
</div>

<script>
let _vtAgg='hour',_vtVal='total';
function _vtsw(a){{_vtAgg=a;_vtupd();}}
function _vtsv(v){{_vtVal=v;_vtupd();}}
function _vtupd(){{
  ['hour','wd','month'].forEach(a=>['total','pb'].forEach(v=>{{
    const el=document.getElementById(`vt-${{a}}-${{v}}`);
    if(el)el.style.display=(a===_vtAgg&&v===_vtVal)?'block':'none';
  }}));
  document.querySelectorAll('[data-vta]').forEach(b=>b.classList.toggle('vt-on',b.dataset.vta===_vtAgg));
  document.querySelectorAll('[data-vtv]').forEach(b=>b.classList.toggle('vt-on',b.dataset.vtv===_vtVal));
}}
const DATA  = {data_js};
const ORDER = {order_js};
const TYPES = {types_js};
const MCH   = {mc_hours_js};
const C     = {consts_js};
const WD    = ['Po','Út','St','Čt','Pá','So','Ne'];
const MON   = ['Led','Úno','Bře','Dub','Kvě','Čvn','Čvc','Srp','Zář','Říj','Lis','Pro'];

let cur=null;
const si=document.getElementById('si'),bl=document.getElementById('bl'),sw=document.getElementById('sw');

// ── Score helpers ─────────────────────────────────────────────────────────────
function _scoreClr(s){{return s>=90?'#ef4444':s>=70?'#f59e0b':'#2563eb';}}
function _scoreBg(s){{return s>=90?'#fee2e2':s>=70?'#fef3c7':'#dbeafe';}}
function _branchScore(d){{
  if(!d.cap1)return null;
  return Math.round(d.cap1.util_ob||0);
}}
function _rankClick(id){{
  si.value=DATA[id].name;bl.classList.remove('open');render(id);
}}

// ── Ranking table (home screen) ───────────────────────────────────────────────
function renderRanking(){{
  const rows=ORDER.map(id=>{{
    const d=DATA[id];
    const score=_branchScore(d);
    if(score===null)return null;
    const mtgPD=d.mtgs_pb_day;
    const wkPD=d.has_svc&&d.cap1&&d.svc_fte>0?d.cap1.bezhot_base/C.WALKIN_AVG_MINS/(d.n_days||1)/d.svc_fte:null;
    const totalMtgs=(d.by_type?.online||0)+(d.by_type?.fyzicka||0);
    const revMtg=d.sales&&d.sales.total_objem>0&&totalMtgs>0?d.sales.total_objem/totalMtgs:null;
    const revBanker=d.sales&&d.sales.total_objem>0&&d.bankers>0?d.sales.total_objem/d.bankers:null;
    return{{id,d,score,mtgPD,wkPD,revMtg,revBanker}};
  }}).filter(Boolean).sort((a,b)=>(b.score||0)-(a.score||0));
  if(!rows.length){{document.getElementById('mc').innerHTML='<div class="placeholder">← Vyberte pobočku výše</div>';return;}}
  const hasSvc=rows.some(r=>r.wkPD!=null);
  const hasSales=rows.some(r=>r.revMtg!=null);
  const scoreCol=(s)=>`<td><div class="rank-score" style="background:${{_scoreBg(s)}};color:${{_scoreClr(s)}};">${{s}}%</div></td>`;
  const thead=`<tr>
    <th style="width:32px;">#</th>
    <th>Pobočka</th>
    <th>Skóre</th>
    <th title="Průměrné schůzky OB / bankéř / den">Sch./bank./den</th>
    ${{hasSvc?'<th title="Walkins / BKP Medior / den">Walk./BKP/den</th>':''}}
    ${{hasSales?'<th title="Výnos z prodejů na jednu schůzku OB (CZK)">Výnos/schůzku</th><th title="Roční výnos z prodejů na bankéře (CZK)">Výnos/bankéř</th>':''}}
  </tr>`;
  const tbody=rows.map((r,i)=>`<tr onclick="_rankClick('${{r.id}}')">
    <td style="font-weight:700;color:#94a3b8;font-size:.78rem;">${{i+1}}</td>
    <td>
      <div style="font-weight:600;color:#1e293b;">${{r.d.name}}</div>
      <div style="font-size:.68rem;color:#94a3b8;">#${{r.id}}${{r.d.branch_format?' · '+r.d.branch_format:''}}</div>
    </td>
    ${{scoreCol(r.score)}}
    <td style="font-weight:700;color:#1e3a8a;">${{r.mtgPD!=null?fmt1(r.mtgPD):'—'}}</td>
    ${{hasSvc?`<td style="color:#0891b2;">${{r.wkPD!=null?fmt1(r.wkPD):'—'}}</td>`:''}}
    ${{hasSales?`<td style="font-weight:600;color:#15803d;">${{r.revMtg!=null?fmtN(Math.round(r.revMtg))+' Kč':'—'}}</td>
               <td style="font-weight:600;color:#15803d;">${{r.revBanker!=null?fmtN(Math.round(r.revBanker))+' Kč':'—'}}</td>`:''}}
  </tr>`).join('');
  const note=`<div style="font-size:.64rem;color:#94a3b8;padding:8px 12px 4px;">
    Skóre = OB utilizace (%) · Výnosy z ${{rows.filter(r=>r.revMtg!=null).length}} poboček s dostupnými prodejními daty
  </div>`;
  document.getElementById('mc').innerHTML=`
    <div style="background:#fff;border:1.5px solid #dbeafe;border-radius:14px;overflow:hidden;margin-bottom:14px;">
      <div style="background:#1e3a8a;padding:11px 16px;">
        <div style="font-size:.84rem;font-weight:700;color:#fff;">🏆 Přehled vytíženosti poboček</div>
        <div style="font-size:.68rem;color:#93c5fd;margin-top:2px;">Seřazeno dle skóre vytíženosti (OB utilizace) · kliknutím zobrazíte detail</div>
      </div>
      <div style="overflow-x:auto;"><table class="rank-tbl"><thead>${{thead}}</thead><tbody>${{tbody}}</tbody></table></div>
      ${{note}}
    </div>`;
}}

function renderList(q){{
  q=q.toLowerCase();
  const hits=ORDER.filter(id=>DATA[id].name.toLowerCase().includes(q)||id.includes(q)).slice(0,120);
  bl.innerHTML=hits.map(id=>`<div class="bi${{id===cur?' sel':''}}" data-id="${{id}}">
    ${{DATA[id].name}} <span style="color:#bbb;font-size:.77em">#${{id}}</span></div>`).join('');
  bl.classList.toggle('open',hits.length>0);
}}
si.addEventListener('input',()=>renderList(si.value));
si.addEventListener('focus',()=>renderList(si.value));
document.addEventListener('click',e=>{{if(!sw.contains(e.target))bl.classList.remove('open');}});
bl.addEventListener('click',e=>{{
  const it=e.target.closest('.bi');if(!it)return;
  si.value=DATA[it.dataset.id].name;bl.classList.remove('open');render(it.dataset.id);
}});

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmt1(n){{if(n==null||isNaN(n))return'—';const f=parseFloat(n);
  return f.toLocaleString('cs',{{minimumFractionDigits:f%1?1:0,maximumFractionDigits:1}});}}
function fmtI(n){{return n==null?'—':Math.round(n).toLocaleString('cs');}}
function fmtN(n){{if(!n&&n!==0)return'0';n=Math.round(n);
  if(n>=1000000)return(n/1000000).toFixed(1).replace('.0','')+'M';
  if(n>=1000)return(n/1000).toFixed(1).replace('.0','')+'k';return String(n);}}
function fmtH(mins){{if(!mins)return'0h';const h=mins/60;
  return h>=1?Math.round(h)+'h':Math.round(mins)+'min';}}

function stackedBars(labels,typeArr,maxH){{
  const totals=labels.map((_,i)=>typeArr.reduce((s,t)=>s+(t.values[i]||0),0));
  const maxV=Math.max(...totals,1);
  return labels.map((lbl,i)=>{{
    const tot=totals[i];
    const segs=typeArr.filter(t=>(t.values[i]||0)>0).map(t=>{{
      const h=(t.values[i]/maxV*maxH).toFixed(1);
      return`<div class="bseg" style="height:${{h}}px;background:${{t.color}};"
                  title="${{t.label}}: ${{fmtI(t.values[i])}}"></div>`;
    }}).join('');
    return`<div class="bw">
      <div class="bnum">${{tot>0?fmtN(tot):''}}</div>
      <div class="bs" style="height:${{(tot/maxV*maxH).toFixed(1)}}px;max-height:${{maxH}}px;">${{segs}}</div>
      <div class="blbl">${{lbl}}</div></div>`;
  }}).join('');
}}
function legend(){{return`<div class="leg">${{TYPES.map(t=>`<div class="li">
  <div class="ld" style="background:${{t.color}}"></div>${{t.label}}</div>`).join('')}}</div>`;}}

// ── Heatmap ───────────────────────────────────────────────────────────────────
function heatmapSec(d){{
  if(!d.has_time||!d.heatmap)return'';
  const rows=d.is_vikend?[0,1,2,3,4,5,6]:[0,1,2,3,4];
  const cols=MCH;
  const maxV=Math.max(...rows.flatMap(w=>cols.map(h=>d.heatmap[w]?.[h]||0)),1);
  const th='<th></th>'+cols.map(h=>`<th>${{h}}h</th>`).join('');
  const trs=rows.map(w=>{{
    const tds=cols.map(h=>{{
      const v=d.heatmap[w]?.[h]||0,p=v/maxV;
      const r=Math.round(239-(239-29)*p),g=Math.round(246-(246-78)*p),b2=Math.round(255-(255-216)*p);
      const clr=v>0?`rgb(${{r}},${{g}},${{b2}})`:'#f8faff';
      const txt=p>.5?'#1e3a8a':'#94a3b8';
      return`<td style="background:${{clr}};color:${{txt}};"
                  title="${{WD[w]}} ${{h}}h: ${{fmtI(v)}}">${{v>0?fmtN(v):''}}</td>`;
    }}).join('');
    return`<tr style="${{w>=5?'background:#fffbf0':''}}">
      <td style="font-size:.66rem;font-weight:600;color:${{w>=5?'#d97706':'#475569'}};
          white-space:nowrap;padding:2px 5px;">${{WD[w]}}</td>${{tds}}</tr>`;
  }}).join('');
  return`<div class="sec"><div class="st">🗓️ Heatmapa den × hodina (celkové počty návštěv)</div>
    <div style="overflow-x:auto;"><table class="hmt">
      <thead><tr>${{th}}</tr></thead><tbody>${{trs}}</tbody></table></div></div>`;
}}

// ── Capacity bar ──────────────────────────────────────────────────────────────
function capBar(pct,label){{
  if(pct==null)return'';
  const clr=pct<70?'#2563eb':pct<90?'#f59e0b':'#ef4444';
  const bg =pct<70?'#dbeafe':pct<90?'#fef3c7':'#fee2e2';
  const ico=pct<70?'✅':pct<90?'⚠️':'🔴';
  return`<div style="margin-bottom:8px;">
    <div style="font-size:.7rem;color:#64748b;margin-bottom:2px;">${{label}}</div>
    <div class="cbwrap" style="background:${{bg}};">
      <div class="cbfill" style="width:${{Math.min(pct,100)}}%;background:${{clr}};">${{pct.toFixed(1)}}%</div>
    </div>
    <div style="font-size:.67rem;color:#94a3b8;text-align:right;">${{ico}} ${{
      pct<70?'kapacita postačuje':pct<90?'blíží se limitu':'přetíženo'}}</div></div>`;
}}



// ── Interactive visit trend chart ─────────────────────────────────────────────
function visitTrendSec(d){{
  if(!d.total)return'';
  const b=d.bankers||1,nd=d.n_days||1;
  const wdLen=d.is_vikend?7:5,wdLabels=WD.slice(0,wdLen);
  const hrLabels=MCH.map(h=>h+'h');
  const hrTot=TYPES.map(t=>({{label:t.label,color:t.color,values:MCH.map(h=>d.by_hour[t.key]?.[h]||0)}}));
  const hrPB=TYPES.map(t=>({{label:t.label,color:t.color,values:MCH.map(h=>(d.by_hour[t.key]?.[h]||0)/b)}}));
  const wdTot=TYPES.map(t=>({{label:t.label,color:t.color,values:(d.by_weekday[t.key]||Array(7).fill(0)).slice(0,wdLen)}}));
  const wdPB=TYPES.map(t=>({{label:t.label,color:t.color,values:(d.by_weekday[t.key]||Array(7).fill(0)).slice(0,wdLen).map((v,i)=>d.n_days_wd&&d.n_days_wd[i]>0?v/d.n_days_wd[i]/b:0)}}));
  const monTot=TYPES.map(t=>({{label:t.label,color:t.color,values:d.by_month[t.key]||Array(12).fill(0)}}));
  const monPB=TYPES.map(t=>({{label:t.label,color:t.color,values:(d.by_month[t.key]||Array(12).fill(0)).map(v=>nd>0?v/(nd/12)/b:0)}}));
  const mkV=(sfx,show,labels,arr,ht)=>`<div id="vt-${{sfx}}" style="display:${{show?'block':'none'}};"><div class="bars" style="height:${{ht}}px;">${{stackedBars(labels,arr,ht-10)}}</div>${{legend()}}</div>`;
  const views=[
    mkV('hour-total',true,hrLabels,hrTot,90),
    mkV('hour-pb',false,hrLabels,hrPB,90),
    mkV('wd-total',false,wdLabels,wdTot,90),
    mkV('wd-pb',false,wdLabels,wdPB,90),
    mkV('month-total',false,MON,monTot,100),
    mkV('month-pb',false,MON,monPB,100),
  ].join('');
  const wdNote=!d.is_vikend?`<div style="font-size:.68rem;color:#94a3b8;margin-top:4px;">So/Ne skryty — nevíkendová pobočka</div>`:'';
  return`<div class="sec">
    <div style="display:flex;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:12px;">
      <div class="st" style="margin-bottom:0;">📈 Průběh návštěv</div>
      <div style="display:flex;gap:6px;flex-wrap:wrap;margin-left:auto;">
        <div class="vtog">
          <button class="vt-btn vt-on" data-vta="hour" onclick="_vtsw('hour')">Hodina</button>
          <button class="vt-btn" data-vta="wd" onclick="_vtsw('wd')">Den týdne</button>
          <button class="vt-btn" data-vta="month" onclick="_vtsw('month')">Měsíc</button>
        </div>
        <div class="vtog">
          <button class="vt-btn vt-on" data-vtv="total" onclick="_vtsv('total')">Celkem</button>
          <button class="vt-btn" data-vtv="pb" onclick="_vtsv('pb')">/ bankéř / den</button>
        </div>
      </div>
    </div>
    ${{views}}
    ${{wdNote}}
  </div>`;
}}

// ── Staff section ─────────────────────────────────────────────────────────────
function staffSec(d){{
  if(!d.bankers&&!Object.keys(d.positions||{{}}).length)return'';
  const svcNote=d.has_svc
    ?`<span class="badge" style="background:#dbeafe;color:#1d4ed8;">✓ BKP Medior zajišťuje servis</span>`
    :`<span class="badge" style="background:#fef9c3;color:#854d0e;">⚠ OB Junior — fallback servis</span>`;
  const cashNote=d.has_cash
    ?`<span class="badge" style="background:#fef3c7;color:#92400e;margin-top:4px;display:inline-block;">🏦 S pokladnou</span>`
    :`<span class="badge" style="background:#f0f4ff;color:#64748b;margin-top:4px;display:inline-block;">💳 Cashless</span>`;
  const obDtl=Object.entries(d.positions||{{}}).filter(([l])=>l.startsWith('OB'))
    .map(([l,v])=>`${{l}}: <b>${{fmt1(v)}}</b>`).join(' · ');
  const otherRows=Object.entries(d.positions||{{}})
    .filter(([l])=>!l.startsWith('OB')&&l!=='BKP Medior'&&l!=='BKP Junior')
    .map(([l,v])=>`<tr><td style="color:#64748b;font-size:.8rem;">${{l}}</td>
        <td style="font-weight:600;color:#475569;text-align:right;">${{fmt1(v)}}</td></tr>`).join('');
  const porClr=d.poc_kli>0&&d.bankers>0?(Math.round(d.poc_kli/d.bankers)<=C.TARGET_PORTFOLIO?'#16a34a':'#dc2626'):'#64748b';
  const cols=d.has_cash?'1fr 1fr 1fr':'1fr 1fr';
  return`<div class="sec"><div class="st">👤 Personální obsazení</div>
    <div style="display:grid;grid-template-columns:${{cols}};gap:12px;margin-bottom:12px;">
      <div>
        <div style="display:flex;align-items:baseline;gap:6px;margin-bottom:2px;">
          <span class="staff-big-n" style="color:#1d4ed8;">${{fmt1(d.bankers)}}</span>
          <span class="staff-big-lbl">Bankéři OB</span>
        </div>
        <div style="font-size:.7rem;color:#94a3b8;margin-bottom:4px;">${{obDtl||'—'}}</div>
        ${{d.poc_kli>0&&d.bankers>0?`<div style="font-size:.72rem;color:#475569;">
          Portfolio/bankéř: <b style="color:${{porClr}};">${{fmtI(Math.round(d.poc_kli/d.bankers))}}</b>
          <span style="color:#94a3b8;">(cíl ≤${{C.TARGET_PORTFOLIO}})</span></div>`:''}}
      </div>
      <div>
        <div style="display:flex;align-items:baseline;gap:6px;margin-bottom:2px;">
          <span class="staff-big-n" style="color:${{d.has_svc?'#0891b2':'#94a3b8'}};">${{fmt1(d.svc_fte)}}</span>
          <span class="staff-big-lbl">BKP Medior</span>
        </div>
        <div style="font-size:.7rem;color:#94a3b8;margin-bottom:6px;">Servis (bezhot. walkin)</div>
        ${{svcNote}}
      </div>
      ${{d.has_cash?`<div>
        <div style="display:flex;align-items:baseline;gap:6px;margin-bottom:2px;">
          <span class="staff-big-n" style="color:#d97706;">${{fmt1(d.cashier_fte)}}</span>
          <span class="staff-big-lbl">BKP Junior</span>
        </div>
        <div style="font-size:.7rem;color:#94a3b8;margin-bottom:6px;">Pokladník (hotov. op.)</div>
        ${{cashNote}}
      </div>`:''}}
    </div>
    ${{otherRows?`<table class="post" style="margin-bottom:8px;"><tbody>${{otherRows}}</tbody></table>`:''}}
    ${{d.fte!=null?`<div style="font-size:.7rem;color:#94a3b8;padding-top:8px;border-top:1px solid #f0f4ff;">FTE celkem: <b>${{fmt1(d.fte)}}</b> · ${{d.has_cash?'S pokladnou':'Cashless'}}</div>`:''}}
  </div>`;
}}


// ── Opening hours ─────────────────────────────────────────────────────────────
function odSec(d){{
  if(!d.od_days||!d.od_days.length)return'';
  const vBadge=d.is_vikend?`<span class="badge" style="background:#d97706;color:#fff;margin-bottom:8px;
    display:inline-block;">🌅 Víkendová pobočka</span><br>`:'';
  const rows=d.od_days.map(day=>day.closed
    ?`<tr class="${{day.wknd?'od-wknd':''}}">
       <td style="color:#cbd5e1;font-weight:600;">${{day.wknd?'🌅 ':''}}${{day.lbl}}</td>
       <td colspan="3" style="color:#cbd5e1;font-style:italic;text-align:center;">Zavřeno</td></tr>`
    :`<tr class="${{day.wknd?'od-wknd':''}}">
       <td style="font-weight:600;color:${{day.wknd?'#d97706':'#334155'}};white-space:nowrap;">
         ${{day.wknd?'🌅 ':''}}${{day.lbl}}</td>
       <td style="color:#475569;text-align:center;">${{day.dop||'—'}}</td>
       <td style="color:#475569;text-align:center;">${{day.odp||'—'}}</td>
       <td style="font-weight:700;color:#2563eb;text-align:center;">${{day.tot?day.tot+'h':''}}</td></tr>`
  ).join('');
  return`<div class="sec"><div class="st">🕐 Otevírací doba</div>
    ${{vBadge}}
    <table class="odt">
      <thead><tr><th>Den</th><th>Dopoledne</th><th>Odpoledne</th><th>Celkem</th></tr></thead>
      <tbody>${{rows}}</tbody></table>
    ${{d.ph_tyden>0?`<div style="font-size:.76rem;color:#64748b;margin-top:7px;">
      Týdenní ot. hodiny: <b style="color:#2563eb;">${{d.ph_tyden}}h</b></div>`:''}}</div>`;
}}

// ── Methodology ───────────────────────────────────────────────────────────────
function methSec(d){{
  const ob=fmt1(d.bankers), svc=d.has_svc?`BKP Medior (${{fmt1(d.svc_fte)}} FTE)`:'OB Junior (fallback)';
  return`<div class="sec"><div class="st">📐 Metodika výpočtu</div>
  <div class="meth">
    <h4>Staffing použitý ve výpočtu</h4>
    <ul>
      <li><b>OB bankéři:</b> ${{ob}} FTE (osobní bankéř junior + medior + senior)</li>
      <li><b>Servisní bankéř:</b> ${{svc}}</li>
      <li><b>Nominální kapacita:</b> ${{C.WORK_MINS_DAY/60}}h = ${{C.WORK_MINS_DAY}} min/den/bankéř</li>
      <li><b>Absence (nemoci, dovolená, školení):</b> ${{Math.round(C.ABSENCE_RATE*100)}}% → efektivní přítomnost
        ${{Math.round((1-C.ABSENCE_RATE)*100)}}% → efektivní kapacita = ${{C.WORK_MINS_DAY}} × ${{(1-C.ABSENCE_RATE).toFixed(3)}}
        = <b>${{Math.round(C.WORK_MINS_DAY*(1-C.ABSENCE_RATE))}} min/den/bankéř</b></li>
    </ul>
    <h4>Konverze návštěv na čas</h4>
    <ul>
      <li><b>Online schůzka:</b> ${{C.MEETING_MINS}} min — bankéř plně obsazen, nemůže obsluhovat walk-iny</li>
      <li><b>Fyzická schůzka:</b> ${{C.MEETING_MINS}} min — totéž</li>
      <li><b>Bezhotovostní walkin:</b> ${{C.WALKIN_SHORT_PCT*100}}% × ${{C.WALKIN_SHORT_MINS}} min
        + ${{C.WALKIN_LONG_PCT*100}}% × ${{C.WALKIN_LONG_MINS}} min = průměr ${{C.WALKIN_AVG_MINS}} min;
        ${{C.WALKIN_CONVERT_PCT*100}}% se přemění na schůzku (${{C.MEETING_MINS}} min)</li>
      <li><b>Hotovostní walkin:</b> nezahrnut do kapacitního výpočtu</li>
    </ul>
    <h4>Model 1 — reálná data</h4>
    <ul>
      <li>Vstup: skutečné roční počty návštěv z dat (${{fmtI(d.n_days)}} dnů)</li>
      <li>Dostupné OB hodiny: ${{ob}} bankéřů × ${{fmtI(d.n_days)}} dnů × ${{C.WORK_MINS_DAY/60}}h
        = ${{d.cap1?fmtH(d.cap1.avail_ob):'N/A'}}</li>
    </ul>
    <h4>Model 2 — klientský model</h4>
    <ul>
      <li>Vstup: ${{fmtI(d.poc_kli)}} klientů (ze zdrojů profitabilita/kpis)</li>
      <li>Distribuce typů návštěv: poměr online/fyzické/bezhot. z reálných dat</li>
      <li>Počítá se s ${{C.WORKING_DAYS}} pracovními dny ročně</li>
      <li>Interpretace: každý klient = 1 návštěva/rok, rozdělená dle reálných poměrů</li>
    </ul>
    <h4>Monte Carlo simulace</h4>
    <ul>
      <li>${{C.MC_ITERATIONS}} simulací průměrného pracovního dne</li>
      <li>Pro každou hodinu (6–21h): počet příchozích ~ Poisson(λ), kde λ = průměrný denní počet</li>
      <li>Hodinová kapacita OB = počet bankéřů × 60 min × ${{(1-C.ABSENCE_RATE).toFixed(3)}} (efektivní přítomnost)</li>
      <li>Pokrytí: % simulací, kdy v žádné hodině nebyla překročena kapacita</li>
      <li>P95: v 95 % simulovaných dnů bylo využití ≤ P95 hodnota</li>
    </ul>
    <h4>Scénář 1× ročně s každým klientem</h4>
    <ul>
      <li>Každý klient portfolia = 1 schůzka × ${{C.MEETING_MINS}} min ročně</li>
      <li>Dostupné OB hodiny = bankéři × ${{C.WORKING_DAYS}} dnů × ${{C.WORK_MINS_DAY}} min × ${{(1-C.ABSENCE_RATE).toFixed(3)}}</li>
      <li>Potřební bankéři = (klienti × ${{C.MEETING_MINS}} min) ÷ (${{C.WORKING_DAYS}} × ${{C.WORK_MINS_DAY}} × ${{(1-C.ABSENCE_RATE).toFixed(3)}})</li>
    </ul>
  </div></div>`;
}}

// ── MC avg heatmap day × hour ─────────────────────────────────────────────────
function mcHeatmapSec(d){{
  if(!d.has_time||!d.heatmap_avg)return'';
  const rows=d.is_vikend?[0,1,2,3,4,5,6]:[0,1,2,3,4];
  const cols=MCH;
  const maxV=Math.max(...rows.flatMap(w=>cols.map(h=>d.heatmap_avg[w]?.[h]||0)),0.01);
  const th='<th></th>'+cols.map(h=>`<th>${{h}}h</th>`).join('');
  const trs=rows.map(w=>{{
    const nd=d.n_days_wd?.[w]||0;
    const tds=cols.map(h=>{{
      const v=d.heatmap_avg[w]?.[h]||0,p=v/maxV;
      const ri=Math.round(239-(239-29)*p),gi=Math.round(246-(246-78)*p),bi=Math.round(255-(255-216)*p);
      const clr=v>0?`rgb(${{ri}},${{gi}},${{bi}})`:'#f8faff';
      const txt=p>.5?'#1e3a8a':'#94a3b8';
      return`<td style="background:${{clr}};color:${{txt}};"
                  title="${{WD[w]}} ${{h}}h: ø ${{v.toFixed(1)}}/den (${{nd}} dnů)">${{v>=0.1?v.toFixed(1):''}}</td>`;
    }}).join('');
    const ndLbl=nd>0?`<span style="font-size:.56rem;color:#94a3b8;margin-left:2px;">(${{nd}}d)</span>`:'';
    return`<tr style="${{w>=5?'background:#fffbf0':''}}">
      <td style="font-size:.66rem;font-weight:600;color:${{w>=5?'#d97706':'#475569'}};
          white-space:nowrap;padding:2px 5px;">${{WD[w]}}${{ndLbl}}</td>${{tds}}</tr>`;
  }}).join('');
  return`<div class="sec"><div class="st">🗓️ MC vstup — průměrné návštěvy den × hodina (ø/den)</div>
    <div style="font-size:.72rem;color:#64748b;margin-bottom:6px;">
      Základ pro Poisson λ v Monte Carlo simulaci.
      Čísla = průměrný počet návštěv (všech typů) v dané hodině za daný den v týdnu.
      Závorka = počet takových dnů v datasetu.
    </div>
    <div style="overflow-x:auto;"><table class="hmt">
      <thead><tr>${{th}}</tr></thead><tbody>${{trs}}</tbody></table></div></div>`;
}}

// ── Meeting mix per hour (skladba schůzek) ────────────────────────────────────
function mixSec(d){{
  if(!d.has_time)return'';
  const hrs=MCH;
  // Filter to hours with any visits
  const totals=hrs.map(h=>TYPES.reduce((s,t)=>s+(d.by_hour[t.key]?.[h]||0),0));
  const active=hrs.filter((_,i)=>totals[i]>0);
  if(!active.length)return'';

  const maxV=Math.max(...active.map(h=>totals[hrs.indexOf(h)]),0.01);

  // Stacked 100% bars (percentage composition)
  const pctBars=active.map(h=>{{
    const i=hrs.indexOf(h);
    const tot=totals[i];
    const segs=TYPES.map(t=>{{
      const v=d.by_hour[t.key]?.[h]||0;
      const pct=tot>0?v/tot*100:0;
      return pct>0?`<div style="height:${{pct.toFixed(1)}}%;background:${{t.color}};flex-shrink:0;"
                        title="${{t.label}} ${{h}}h: ${{pct.toFixed(0)}}% (ø ${{v.toFixed(2)}}/den)"></div>`:'';
    }}).join('');
    return`<div class="bw">
      <div style="height:70px;display:flex;flex-direction:column-reverse;
                  width:100%;border-radius:3px 3px 0 0;overflow:hidden;">${{segs}}</div>
      <div class="blbl">${{h}}h</div>
    </div>`;
  }}).join('');

  // Absolute avg bars (volume)
  const absBars=active.map(h=>{{
    const i=hrs.indexOf(h);
    const tot=totals[i];
    const ht=(tot/maxV*70).toFixed(1);
    const segs=TYPES.map(t=>{{
      const v=d.by_hour[t.key]?.[h]||0,pct=tot>0?v/tot*100:0;
      return pct>0?`<div style="height:${{pct.toFixed(1)}}%;background:${{t.color}};flex-shrink:0;"
                        title="${{t.label}} ${{h}}h: ø ${{v.toFixed(2)}}/den"></div>`:'';
    }}).join('');
    return`<div class="bw">
      <div class="bnum">${{tot>0?tot.toFixed(1):''}}</div>
      <div style="height:${{ht}}px;display:flex;flex-direction:column-reverse;
                  width:100%;border-radius:3px 3px 0 0;overflow:hidden;">${{segs}}</div>
      <div class="blbl">${{h}}h</div>
    </div>`;
  }}).join('');

  return`<div class="sec"><div class="st">🔀 Skladba schůzek podle hodiny</div>
    <div style="font-size:.72rem;color:#475569;margin-bottom:8px;font-weight:600;">
      Procentuální složení (100 % = všechny typy v dané hodině)</div>
    <div class="bars" style="height:80px;">${{pctBars}}</div>
    ${{legend()}}
    <div style="font-size:.72rem;color:#475569;margin:12px 0 8px;font-weight:600;">
      Průměrný denní objem dle hodiny (abs. hodnoty)</div>
    <div class="bars" style="height:80px;">${{absBars}}</div>
    ${{legend()}}
  </div>`;
}}

// ── Shared capacity panel (Models 1 & 2) ──────────────────────────────────────
function _capPanel(cap,nd,b,has_svc,title,subtitle,extra){{
  const eff=1-C.ABSENCE_RATE,effM=C.WORK_MINS_DAY*eff;
  const oM=cap.online_mins/b/nd,fM=cap.fyzicka_mins/b/nd;
  const bCnv=cap.bezhot_conv/b/nd,bSvc=has_svc?0:cap.bezhot_base/b/nd;
  const usedM=oM+fM+bCnv+bSvc,freeM=Math.max(effM-usedM,0);
  const segs=[
    {{m:oM,   bg:'#1d4ed8',clr:'#fff',   lbl:'Online'}},
    {{m:fM,   bg:'#3b82f6',clr:'#fff',   lbl:'Fyzické'}},
    {{m:bCnv, bg:'#7dd3fc',clr:'#1e3a8a',lbl:'WI→mtg'}},
    {{m:bSvc, bg:'#bae6fd',clr:'#1e3a8a',lbl:'WI servis'}},
    {{m:freeM,bg:'#e2e8f0',clr:'#64748b',lbl:'Volný čas'}},
  ].filter(s=>s.m>=1);
  const dayBar=segs.map(s=>{{
    const w=(s.m/effM*100).toFixed(1),lbl=s.m>18?`${{Math.round(s.m)}}m`:'';
    return`<div class="daybar-seg" style="width:${{w}}%;background:${{s.bg}};color:${{s.clr}};"
      title="${{s.lbl}}: ${{Math.round(s.m)}} min (${{w}}%)">${{lbl}}</div>`;
  }}).join('');
  const legItems=segs.map(s=>`<span style="font-size:.66rem;display:inline-flex;align-items:center;gap:3px;margin-right:8px;">
    <span style="width:8px;height:8px;background:${{s.bg}};border-radius:2px;display:inline-block;flex-shrink:0;"></span>
    ${{s.lbl}} ${{Math.round(s.m)}}m</span>`).join('');
  const util=cap.util_ob;
  const utilClr=util<70?'#2563eb':util<90?'#f59e0b':'#ef4444';
  const utilBg =util<70?'#dbeafe':util<90?'#fef3c7':'#fee2e2';
  const statusTxt=util<70?'✅ kapacita OK':util<90?'⚠️ blíží se limitu':'🔴 přetíženo';
  const deg=(Math.min(util,100)*3.6).toFixed(0);
  const ringStyle=`background:conic-gradient(${{utilClr}} ${{deg}}deg,${{utilBg}} ${{deg}}deg)`;
  const mtgTot=(cap.online_mins+cap.fyzicka_mins)/b/nd/C.MEETING_MINS;
  const mtgClr=mtgTot>=C.TARGET_MTGS_MIN&&mtgTot<=C.TARGET_MTGS_MAX?'#16a34a':mtgTot<C.TARGET_MTGS_MIN?'#64748b':'#dc2626';
  // Service section variables
  const svcNeeded=cap.svc_fte>0&&cap.svc_used>0?cap.svc_used/(nd*C.WORK_MINS_DAY*eff):0;
  const walkinPD=cap.svc_fte>0&&cap.bezhot_base>0?cap.bezhot_base/C.WALKIN_AVG_MINS/nd/cap.svc_fte:0;
  const svcClr=cap.util_svc!=null?(cap.util_svc<70?'#0891b2':cap.util_svc<90?'#f59e0b':'#ef4444'):'#94a3b8';
  const svcOk=svcNeeded<=cap.svc_fte;
  return`<div class="kap-model">
    <div class="kap-model-hd">${{title}}</div>
    <div style="font-size:.72rem;color:#64748b;margin-bottom:10px;">${{subtitle}}</div>
    <div class="cap-gauge-row">
      <div class="cap-ring" style="${{ringStyle}}">
        <div class="cap-ring-in">
          <span style="font-size:.88rem;font-weight:800;color:${{utilClr}};">${{util.toFixed(0)}}%</span>
          <span style="font-size:.48rem;color:#64748b;line-height:1.2;text-align:center;">OB<br>kapc.</span>
        </div>
      </div>
      <div class="cap-gauge-meta">
        <div style="font-size:.67rem;color:#64748b;margin-bottom:3px;">
          Průměrný den / 1 bankéř — efekt. kapacita ${{Math.round(effM)}} min
          (${{C.WORK_MINS_DAY}}min × ${{Math.round(eff*100)}}%)
        </div>
        <div class="daybar">${{dayBar}}</div>
        <div style="margin-top:3px;line-height:1.8;">${{legItems}}</div>
      </div>
    </div>
    <div class="grid3" style="margin-top:8px;">
      <div class="card">
        <div class="cl">Schůzky / bankéř / den</div>
        <div class="cv" style="color:${{mtgClr}};">${{fmt1(mtgTot)}}</div>
        <div class="cs">cíl ${{C.TARGET_MTGS_MIN}}–${{C.TARGET_MTGS_MAX}} · on ${{fmt1(oM/C.MEETING_MINS)}} fyz ${{fmt1(fM/C.MEETING_MINS)}}</div>
      </div>
      <div class="card">
        <div class="cl">Obsazenost OB</div>
        <div class="cv" style="color:${{utilClr}};">${{util.toFixed(1)}}%</div>
        <div class="cs">${{statusTxt}}</div>
      </div>
      <div class="card">
        <div class="cl">Dostupné OB / rok</div>
        <div class="cv">${{fmtH(cap.avail_ob)}}</div>
        <div class="cs">${{fmt1(b)}} bank. × ${{fmtI(nd)}}d × ${{Math.round(eff*100)}}%</div>
      </div>
    </div>
    ${{has_svc&&cap.util_svc!=null?`
    <div style="margin-top:10px;border-top:1px solid #eff6ff;padding-top:8px;">
      <div style="font-size:.7rem;font-weight:700;color:#0891b2;margin-bottom:6px;">BKP Medior — Servisní kapacita</div>
      <div class="grid3">
        <div class="card">
          <div class="cl">Walkin / BKP / den</div>
          <div class="cv" style="color:#0891b2;">${{fmt1(walkinPD)}}</div>
          <div class="cs">${{fmtI(Math.round(cap.bezhot_base/C.WALKIN_AVG_MINS/nd))}} walk./den · ${{fmt1(cap.svc_fte)}} BKP</div>
        </div>
        <div class="card">
          <div class="cl">BKP potřební</div>
          <div class="cv" style="color:${{svcOk?'#15803d':'#b91c1c'}}">${{fmt1(svcNeeded)}}</div>
          <div class="cs">k disp. ${{fmt1(cap.svc_fte)}} · ef. ${{Math.round(eff*100)}}%</div>
        </div>
        <div class="card">
          <div class="cl">Utilizace BKP</div>
          <div class="cv" style="color:${{svcClr}}">${{cap.util_svc.toFixed(1)}}%</div>
          <div class="cs">${{cap.util_svc<70?'OK':cap.util_svc<90?'blíží se':'přetíženo'}}</div>
        </div>
      </div>
    </div>`:''}}
    ${{extra||''}}
  </div>`;
}}

// ── SVG bar chart: hourly FTE demand for deterministic models ─────────────────
function _svgFteBar(obArr,svcArr,capOB,capSVC,hasSvc,W,H){{
  if(!obArr||!obArr.length)return'';
  const pl=32,pr=14,pt=8,pb=18,pw=W-pl-pr,ph=H-pt-pb;
  const n=obArr.length;
  const maxY=Math.max(...obArr,...(hasSvc&&svcArr.length?svcArr:[0]),capOB,hasSvc?capSVC:0)*1.2||1;
  const ys=v=>pt+ph*(1-Math.min(v,maxY)/maxY);
  const bW=pw/n;
  const step=maxY<=1.5?0.25:maxY<=3?0.5:maxY<=6?1:maxY<=12?2:5;
  let yts=[];for(let v=0;v<=maxY*1.05;v+=step)yts.push(parseFloat(v.toFixed(2)));
  const ytkEl=yts.map(v=>`<line x1="${{pl}}" y1="${{ys(v).toFixed(1)}}" x2="${{pl+pw}}" y2="${{ys(v).toFixed(1)}}" stroke="#e2e8f0" stroke-width="0.5"/><text x="${{pl-3}}" y="${{ys(v).toFixed(1)}}" text-anchor="end" dominant-baseline="middle" font-size="7" fill="#94a3b8">${{v}}</text>`).join('');
  const xtkEl=MCH.map((h,i)=>i%2===0?`<text x="${{(pl+(i+0.5)*bW).toFixed(1)}}" y="${{H-3}}" text-anchor="middle" font-size="7" fill="#94a3b8">${{h}}h</text>`:'').join('');
  const capOBY=ys(capOB).toFixed(1);
  const capLines=`<line x1="${{pl}}" y1="${{capOBY}}" x2="${{pl+pw}}" y2="${{capOBY}}" stroke="#2563eb" stroke-dasharray="5,3" stroke-width="1.5"/>
    <text x="${{pl+pw+2}}" y="${{capOBY}}" dominant-baseline="middle" font-size="7" fill="#2563eb">OB ${{fmt1(capOB)}}</text>
    ${{hasSvc?`<line x1="${{pl}}" y1="${{ys(capSVC).toFixed(1)}}" x2="${{pl+pw}}" y2="${{ys(capSVC).toFixed(1)}}" stroke="#d97706" stroke-dasharray="5,3" stroke-width="1.5"/><text x="${{pl+pw+2}}" y="${{ys(capSVC).toFixed(1)}}" dominant-baseline="middle" font-size="7" fill="#d97706">SVC ${{fmt1(capSVC)}}</text>`:''}}`;
  const obBars=obArr.map((v,i)=>{{
    const pct=capOB>0?v/capOB:0;
    const clr=pct<0.7?'#2563eb':pct<1.0?'#f59e0b':'#ef4444';
    const x=pl+i*bW+0.5,bh=Math.max(ph-(ys(v)-pt),0),by=ys(v);
    return bh>0?`<rect x="${{x.toFixed(1)}}" y="${{by.toFixed(1)}}" width="${{(bW-1).toFixed(1)}}" height="${{bh.toFixed(1)}}" fill="${{clr}}" fill-opacity="0.8"/>`:'';}}).join('');
  const svcBars=hasSvc&&svcArr.length?svcArr.map((v,i)=>{{
    const pct=capSVC>0?v/capSVC:0;
    const clr=pct<0.7?'#d97706':'#ef4444';
    const x=pl+i*bW+bW*0.55,sw=bW*0.38,bh=Math.max(ph-(ys(v)-pt),0),by=ys(v);
    return bh>0?`<rect x="${{x.toFixed(1)}}" y="${{by.toFixed(1)}}" width="${{sw.toFixed(1)}}" height="${{bh.toFixed(1)}}" fill="${{clr}}" fill-opacity="0.7"/>`:'';}}).join(''):'';
  return`<svg viewBox="0 0 ${{W}} ${{H}}" style="width:100%;display:block;">
    <rect x="${{pl}}" y="${{pt}}" width="${{pw}}" height="${{ph}}" fill="#f8faff" rx="2"/>
    ${{ytkEl}}${{xtkEl}}${{obBars}}${{svcBars}}${{capLines}}
    <line x1="${{pl}}" y1="${{pt}}" x2="${{pl}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
    <line x1="${{pl}}" y1="${{pt+ph}}" x2="${{pl+pw}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
  </svg>`;
}}

// ── Workspace recommendation ──────────────────────────────────────────────────
function _workRec(peakMtgLam,peakBehLam,bankers,svcFte,eff,hasSvc){{
  function p95(l){{return l>0?l+1.645*Math.sqrt(l):0;}}
  const pm=p95(peakMtgLam),pb=hasSvc?p95(peakBehLam):0;
  const rooms=peakMtgLam>0?Math.ceil(pm*C.MEETING_MINS/60):0;
  const desks=hasSvc&&peakBehLam>0?Math.ceil(pb*C.WALKIN_AVG_MINS/60):0;
  const bDesks=Math.ceil(bankers*eff);
  const cs='flex:1;min-width:90px;padding:7px 10px;';
  return`<div style="margin-top:10px;border-top:1px solid #eff6ff;padding-top:10px;">
    <div style="font-size:.7rem;font-weight:700;color:#1e3a8a;margin-bottom:7px;">🏢 Doporučení pracovišť</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;">
      <div class="card" style="${{cs}}">
        <div class="cl">Zasedací místnosti</div>
        <div class="cv" style="color:#1d4ed8;">${{rooms}}</div>
        <div class="cs">P95(${{fmt1(peakMtgLam)}})×${{C.MEETING_MINS}}min</div>
      </div>
      ${{hasSvc&&peakBehLam>0?`<div class="card" style="${{cs}}">
        <div class="cl">Servisní místa</div>
        <div class="cv" style="color:#0891b2;">${{desks}}</div>
        <div class="cs">P95(${{fmt1(peakBehLam)}})×${{C.WALKIN_AVG_MINS}}min</div>
      </div>`:''}}
      <div class="card" style="${{cs}}">
        <div class="cl">Bankéřská místa</div>
        <div class="cv" style="color:#7c3aed;">${{bDesks}}</div>
        <div class="cs">${{fmt1(bankers)}} bank.×${{Math.round(eff*100)}}%</div>
      </div>
    </div>
    <div style="font-size:.63rem;color:#94a3b8;margin-top:4px;">
      P95=λ+1.645√λ · místn.=⌈P95×${{C.MEETING_MINS}}/60⌉ · serv.=⌈P95×${{C.WALKIN_AVG_MINS}}/60⌉ · bank.=⌈bank.×ef.%⌉
    </div>
  </div>`;
}}

// ── Horizontal bar helper ─────────────────────────────────────────────────────
function _hbar(items,maxV){{
  return items.filter(it=>it.v>0).map(it=>{{
    const w=(it.v/maxV*100).toFixed(1);
    const light=it.v/maxV<0.35;
    return`<div class="hbar-row">
      <div class="hbar-lbl">${{it.l}}</div>
      <div class="hbar-track">
        <div class="hbar-fill" style="width:${{w}}%;background:${{it.c}};">
          <span class="hbar-val" style="color:${{light?it.c:'#fff'}};${{light?'margin-left:4px;':''}}">
            ${{it.v.toFixed(1)}}</span>
        </div>
      </div>
    </div>`;
  }}).join('');
}}

// ── Capacity section ──────────────────────────────────────────────────────────
function capacitySec(d){{
  if(!d.cap1&&!d.cap2&&!d.mc)return'';
  const eff=1-C.ABSENCE_RATE,nd=d.n_days||1,b=d.bankers||1;

  // ── helper: OB daybar+ring for a cap object ──────────────────────────────
  function _obPanel(cap,ndays,bankers,hasSvc,lbl){{
    const effM=C.WORK_MINS_DAY*eff;
    const oM=cap.online_mins/bankers/ndays,fM=cap.fyzicka_mins/bankers/ndays;
    const bCnv=cap.bezhot_conv/bankers/ndays,bSvc=hasSvc?0:cap.bezhot_base/bankers/ndays;
    const usedM=oM+fM+bCnv+bSvc,freeM=Math.max(effM-usedM,0);
    const segs=[
      {{m:oM,   bg:'#1d4ed8',clr:'#fff',   lbl:'Online'}},
      {{m:fM,   bg:'#3b82f6',clr:'#fff',   lbl:'Fyzické'}},
      {{m:bCnv, bg:'#7dd3fc',clr:'#1e3a8a',lbl:'WI→mtg'}},
      {{m:bSvc, bg:'#bae6fd',clr:'#1e3a8a',lbl:'WI serv.'}},
      {{m:freeM,bg:'#e2e8f0',clr:'#64748b',lbl:'Volný'}},
    ].filter(s=>s.m>=1);
    const dayBar=segs.map(s=>{{
      const w=(s.m/effM*100).toFixed(1),lbl=s.m>18?`${{Math.round(s.m)}}m`:'';
      return`<div class="daybar-seg" style="width:${{w}}%;background:${{s.bg}};color:${{s.clr}};" title="${{s.lbl}}: ${{Math.round(s.m)}}min">${{lbl}}</div>`;
    }}).join('');
    const util=cap.util_ob;
    const utilClr=util<70?'#2563eb':util<90?'#f59e0b':'#ef4444';
    const utilBg=util<70?'#dbeafe':util<90?'#fef3c7':'#fee2e2';
    const deg=(Math.min(util,100)*3.6).toFixed(0);
    const rs=`background:conic-gradient(${{utilClr}} ${{deg}}deg,${{utilBg}} ${{deg}}deg)`;
    // meeting count bars
    const oPerDay=cap.online_mins/ndays/C.MEETING_MINS;
    const fPerDay=cap.fyzicka_mins/ndays/C.MEETING_MINS;
    const wConvPerDay=cap.bezhot_conv/ndays/C.MEETING_MINS;
    const maxMtg=Math.max(oPerDay,fPerDay,wConvPerDay,0.01);
    const mtgBars=_hbar([
      {{l:'Online',v:oPerDay,c:'#1d4ed8'}},
      {{l:'Fyzické',v:fPerDay,c:'#3b82f6'}},
      {{l:'WI→schůzka',v:wConvPerDay,c:'#7dd3fc'}},
    ],maxMtg+0.01);
    return`<div class="cap-sub" style="min-width:0;">
      <div class="cap-sub-hd" style="color:#1d4ed8;">👤 OB Schůzky${{lbl?' · '+lbl:''}}</div>
      <div class="cap-gauge-row">
        <div class="cap-ring" style="${{rs}}">
          <div class="cap-ring-in">
            <span style="font-size:.88rem;font-weight:800;color:${{utilClr}};">${{util.toFixed(0)}}%</span>
            <span style="font-size:.48rem;color:#64748b;line-height:1.2;text-align:center;">OB<br>util.</span>
          </div>
        </div>
        <div class="cap-gauge-meta">
          <div style="font-size:.62rem;color:#64748b;margin-bottom:2px;">Průměrný bankéř / den · ${{Math.round(effM)}} min</div>
          <div class="daybar">${{dayBar}}</div>
        </div>
      </div>
      <div style="margin-top:8px;border-top:1px solid #eff6ff;padding-top:6px;">
        <div style="font-size:.63rem;font-weight:700;color:#475569;margin-bottom:4px;">Schůzky / den (průměr)</div>
        ${{mtgBars}}
        <div style="font-size:.62rem;color:#94a3b8;margin-top:3px;">
          Celkem ${{fmt1(oPerDay+fPerDay+wConvPerDay)}} sch./den · ${{fmtI(Math.round(cap.online_mins/ndays+cap.fyzicka_mins/ndays+cap.bezhot_conv/ndays))}} min/den OB
        </div>
      </div>
    </div>`;
  }}

  // ── helper: BKP service panel ────────────────────────────────────────────
  function _bkpPanel(cap,ndays,hasSvc){{
    if(!hasSvc||cap.util_svc==null)return'';
    const effM=C.WORK_MINS_DAY*eff;
    const util=cap.util_svc,svcFte=cap.svc_fte;
    const svcClr=util<70?'#0891b2':util<90?'#f59e0b':'#ef4444';
    const svcBg=util<70?'#cffafe':util<90?'#fef3c7':'#fee2e2';
    const deg=(Math.min(util,100)*3.6).toFixed(0);
    const rs=`background:conic-gradient(${{svcClr}} ${{deg}}deg,${{svcBg}} ${{deg}}deg)`;
    const wPerDay=cap.bezhot_base/C.WALKIN_AVG_MINS/ndays;
    const wPerBkpDay=svcFte>0?wPerDay/svcFte:0;
    const svcUsedM=cap.svc_used/ndays/svcFte;
    const freeM=Math.max(effM-svcUsedM,0);
    const wBars=_hbar([
      {{l:'Walk./BKP/den',v:wPerBkpDay,c:'#0891b2'}},
      {{l:'Voln. čas BKP',v:freeM/C.WALKIN_AVG_MINS,c:'#e2e8f0'}},
    ],Math.max(wPerBkpDay,freeM/C.WALKIN_AVG_MINS)+0.01);
    return`<div class="cap-sub" style="min-width:0;">
      <div class="cap-sub-hd" style="color:#0891b2;">🧑 Servis BKP Medior</div>
      <div class="cap-gauge-row">
        <div class="cap-ring" style="${{rs}}">
          <div class="cap-ring-in">
            <span style="font-size:.88rem;font-weight:800;color:${{svcClr}};">${{util.toFixed(0)}}%</span>
            <span style="font-size:.48rem;color:#64748b;line-height:1.2;text-align:center;">BKP<br>util.</span>
          </div>
        </div>
        <div class="cap-gauge-meta">
          <div style="font-size:.67rem;color:#64748b;margin-bottom:4px;">
            ${{fmt1(svcFte)}} BKP Medior · ${{fmtI(Math.round(wPerDay))}} walkinů/den celkem
          </div>
          <div style="font-size:.7rem;">Walkin / BKP / den: <b style="color:#0891b2;">${{fmt1(wPerBkpDay)}}</b></div>
        </div>
      </div>
      <div style="margin-top:8px;border-top:1px solid #eff6ff;padding-top:6px;">
        <div style="font-size:.63rem;font-weight:700;color:#475569;margin-bottom:4px;">Servisní kapacita / BKP / den</div>
        ${{wBars}}
      </div>
    </div>`;
  }}

  // ── helper: 90% gap block ────────────────────────────────────────────────
  function _gap90block(cap1,mc,hasSvc,totalNH,ndays){{
    const c1=cap1;
    const ob90ok=c1.util_ob>=90;
    const obSF1=!ob90ok&&c1.ob_used>0?c1.avail_ob*0.9/c1.ob_used:0;
    const addNH1=obSF1>1?Math.ceil(totalNH*(obSF1-1)/ndays):0;
    const nowNH1=totalNH/ndays;
    const s90ok=c1.util_svc!=null&&c1.util_svc>=90;
    const addWB1=hasSvc&&c1.avail_svc>0&&!s90ok&&(c1.avail_svc*0.9-c1.svc_used)>0?Math.ceil((c1.avail_svc*0.9-c1.svc_used)/C.WALKIN_AVG_MINS/ndays):0;
    const nowWB1=(totalNH-((c1.online_mins+c1.fyzicka_mins)/ndays/C.MEETING_MINS*(1/(c1.ob_used/totalNH||1))))||totalNH/ndays*0.4;
    let obLine='',svcLine='';
    if(ob90ok){{obLine=`<span style="color:#16a34a;">✅ OB již ≥ 90 % (${{c1.util_ob.toFixed(1)}} %)</span>`;}}
    else if(addNH1>0){{obLine=`OB: Nyní <b>${{fmt1(nowNH1)}}</b> → potřeba <b>${{fmt1(nowNH1+addNH1)}}</b> nav./den (+${{fmtI(addNH1)}})`;}}
    if(hasSvc){{
      if(s90ok){{svcLine=`<span style="color:#16a34a;"> · ✅ BKP již ≥ 90 %</span>`;}}
      else if(addWB1>0){{svcLine=` · BKP: +${{fmtI(addWB1)}} walkinů/den`;}}
    }}
    if(!ob90ok&&addNH1===0&&(!hasSvc||s90ok))return'';
    return`<div class="gap90">
      <span style="font-weight:700;color:#92400e;">Pro 90% utilizaci: </span>${{obLine}}${{svcLine}}
    </div>`;
  }}

  // ── M1 variant extra (35% walkin→meeting) ────────────────────────────────
  function _m4extra(cap1,cap_m4,bankers2,hasSvc,ndays2){{
    if(!cap_m4)return'';
    const c4=cap_m4,c1=cap1;
    const u4Clr=c4.util_ob<70?'#2563eb':c4.util_ob<90?'#f59e0b':'#ef4444';
    const s4Clr=c4.util_svc!=null?(c4.util_svc<70?'#0891b2':c4.util_svc<90?'#f59e0b':'#ef4444'):'#94a3b8';
    const mtg4=(c4.online_mins+c4.fyzicka_mins)/bankers2/ndays2/C.MEETING_MINS;
    const mtg1=(c1.online_mins+c1.fyzicka_mins)/bankers2/ndays2/C.MEETING_MINS;
    return`<details style="margin-top:10px;border-top:1px solid #eff6ff;padding-top:8px;">
      <summary style="font-size:.72rem;font-weight:700;color:#7c3aed;cursor:pointer;user-select:none;list-style:none;">▶ Scénář: 35 % walkinů → schůzka OB</summary>
      <div style="margin-top:8px;">
        <div style="font-size:.67rem;color:#64748b;margin-bottom:7px;">Stejná reálná data · 35 % bezhotovostních walkinů → OB schůzka ${{C.MEETING_MINS}} min</div>
        <div class="grid3" style="margin-bottom:0;">
          <div class="card"><div class="cl">OB Utilizace</div>
            <div class="cv" style="color:${{u4Clr}};">${{c4.util_ob.toFixed(1)}} %</div>
            <div class="cs">reálná data: ${{c1.util_ob.toFixed(1)}} %</div></div>
          <div class="card"><div class="cl">Schůzky / bankéř / den</div>
            <div class="cv">${{fmt1(mtg4)}}</div>
            <div class="cs">reálná data: ${{fmt1(mtg1)}}</div></div>
          ${{hasSvc&&c4.util_svc!=null?`<div class="card"><div class="cl">BKP Utilizace</div>
            <div class="cv" style="color:${{s4Clr}};">${{c4.util_svc.toFixed(1)}} %</div>
            <div class="cs">reálná data: ${{c1.util_svc!=null?c1.util_svc.toFixed(1)+' %':'N/A'}}</div></div>`:''}}
        </div>
      </div>
    </details>`;
  }}

  // ── M3 (1×/rok) variant extra ─────────────────────────────────────────────
  function _m3extra(annual_1x){{
    if(!annual_1x)return'';
    const a=annual_1x;
    const aClr=a.util_ob<70?'#2563eb':a.util_ob<90?'#f59e0b':'#ef4444';
    const aOk=a.bankers_needed<=a.bankers;
    return`<details style="margin-top:10px;border-top:1px solid #eff6ff;padding-top:8px;">
      <summary style="font-size:.72rem;font-weight:700;color:#7c3aed;cursor:pointer;user-select:none;list-style:none;">▶ Scénář: 1× ročně s každým klientem</summary>
      <div style="margin-top:8px;">
        <div style="font-size:.67rem;color:#64748b;margin-bottom:7px;">Každý z ${{fmtI(a.poc_kli)}} klientů × 1 schůzka × ${{C.MEETING_MINS}} min · ${{fmtI(a.open_days||C.WORKING_DAYS)}} dnů/rok</div>
        <div class="grid3" style="margin-bottom:0;">
          <div class="card"><div class="cl">Schůzky / bankéř / den</div>
            <div class="cv">${{fmt1(a.mtgs_pb_day)}}</div>
            <div class="cs">cíl ${{C.TARGET_MTGS_MIN}}–${{C.TARGET_MTGS_MAX}}/den</div></div>
          <div class="card"><div class="cl">Bankéři potřební</div>
            <div class="cv" style="color:${{aOk?'#15803d':'#b91c1c'}}">${{fmt1(a.bankers_needed)}}</div>
            <div class="cs">k disp. ${{fmt1(a.bankers)}}</div></div>
          <div class="card"><div class="cl">Utilizace OB</div>
            <div class="cv" style="color:${{aClr}}">${{a.util_ob.toFixed(1)}}%</div>
            <div class="cs">${{aOk?'✅ kapacita OK':'⚠️ přetíženo'}}</div></div>
        </div>
      </div>
    </details>`;
  }}

  // ── MC 90% gap helper ─────────────────────────────────────────────────────
  function _mcGap(mc,hasSvc,total,ndays){{
    if(!mc)return'';
    const obP=mc.ob_cap_day>0?mc.ob_p95_day/mc.ob_cap_day*100:null;
    const obSFmc=obP!=null&&obP<90&&mc.ob_p95_day>0?mc.ob_cap_day*0.9/mc.ob_p95_day:0;
    const addNHmc=obSFmc>1?Math.ceil((obSFmc-1)*total/ndays):0;
    const svcP=mc.svc_cap_day>0&&mc.svc_p95_day>0?mc.svc_p95_day/mc.svc_cap_day*100:null;
    const addWBmc=hasSvc&&svcP!=null&&svcP<90&&mc.svc_p95_day>0?Math.ceil((mc.svc_cap_day*0.9/mc.svc_p95_day-1)*(d.by_type?.bezhot||0)/ndays):0;
    const nowPD=total/ndays;
    let obLine='',svcLine='';
    if(obP>=90){{obLine=`<span style="color:#16a34a;">✅ OB P95 již ≥ 90 % (${{obP.toFixed(1)}} %)</span>`;}}
    else if(addNHmc>0){{obLine=`OB: Nyní <b>${{fmt1(nowPD)}}</b> → potřeba <b>${{fmt1(nowPD+addNHmc)}}</b> nav./den (+${{fmtI(addNHmc)}})`;}}
    if(hasSvc){{
      if(svcP!=null&&svcP>=90){{svcLine=`<span style="color:#16a34a;"> · ✅ BKP P95 ≥ 90 %</span>`;}}
      else if(addWBmc>0){{svcLine=` · BKP: +${{fmtI(addWBmc)}} walkinů/den`;}}
    }}
    if(!obLine&&!svcLine)return'';
    return`<div class="gap90" style="margin-top:6px;">
      <span style="font-weight:700;color:#92400e;">Pro 90% P95: </span>${{obLine}}${{svcLine}}
    </div>`;
  }}

  // ── Room recommendation inline ─────────────────────────────────────────────
  function _roomsInline(){{
    if(!d.rooms)return'';
    const r=d.rooms;
    return`<div class="kap-model">
      <div class="kap-model-hd">🏢 Doporučení prostor</div>
      <div class="grid2" style="margin-bottom:8px;">
        <div class="card">
          <div class="cl">Zasedací místnosti</div>
          <div class="cv" style="color:#1d4ed8;">${{r.meeting_rooms}}</div>
          <div class="cs">P95 schůzek v špičce</div>
        </div>
        <div class="card">
          <div class="cl">Servisní místa</div>
          <div class="cv" style="color:#0891b2;">${{r.service_desks}}</div>
          <div class="cs">P95 walkinů v špičce</div>
        </div>
      </div>
      <div style="font-size:.68rem;color:#64748b;line-height:1.6;">
        λ = průměrné příchody v nejfrekventovanější hodině · P95 = λ + 1.645·√λ<br>
        Místnosti = ⌈P95 × ${{C.MEETING_MINS}} min ÷ 60⌉ · Servis = ⌈P95 × ${{C.WALKIN_AVG_MINS}} min ÷ 60⌉
      </div>
    </div>`;
  }}

  // ── REAL DATA (M1) ────────────────────────────────────────────────────────
  const totalNH=(d.by_type?.online||0)+(d.by_type?.fyzicka||0)+(d.by_type?.bezhot||0);
  const m1HTML=d.cap1?`<div class="kap-model">
    <div class="kap-model-hd">Reálná data</div>
    <div style="font-size:.72rem;color:#64748b;margin-bottom:10px;">
      Skutečné návštěvy · ${{fmtI(nd)}} dnů · ${{fmtI(d.total)}} celkem · ${{fmtI(d.bankers)}} bankéřů
    </div>
    <div class="grid2" style="align-items:start;">
      ${{_obPanel(d.cap1,nd,b,d.has_svc,'')}}
      ${{d.has_svc?_bkpPanel(d.cap1,nd,d.has_svc):`<div class="cap-sub"><div class="cap-sub-hd" style="color:#94a3b8;">Bez servisní zóny</div><div style="font-size:.72rem;color:#94a3b8;">Pobočka bez BKP Medior</div></div>`}}
    </div>
    ${{_m4extra(d.cap1,d.cap_m4,b,d.has_svc,nd)}}
    ${{_gap90block(d.cap1,d.mc,d.has_svc,totalNH,nd)}}
  </div>`:'';

  // ── CLIENT MODEL (M2) ─────────────────────────────────────────────────────
  const m2HTML=d.cap2?`<div class="kap-model">
    <div class="kap-model-hd">Klientský model</div>
    <div style="font-size:.72rem;color:#64748b;margin-bottom:10px;">
      ${{fmtI(d.poc_kli)}} klientů · ${{fmtI(d.annual_open_days||252)}} ot. dnů/rok · typy z reálných dat
    </div>
    <div class="grid2" style="align-items:start;">
      ${{_obPanel(d.cap2,d.annual_open_days||252,d.cap2.bankers,d.has_svc,'')}}
      ${{d.has_svc?_bkpPanel(d.cap2,d.annual_open_days||252,d.has_svc):`<div class="cap-sub"><div style="font-size:.72rem;color:#94a3b8;">Bez servisní zóny</div></div>`}}
    </div>
    ${{_m3extra(d.annual_1x)}}
  </div>`:(d.poc_kli===0?`<div class="kap-model"><div class="kap-model-hd">Klientský model</div><div style="font-size:.78rem;color:#94a3b8;">⚠️ Data o počtu klientů nejsou dostupná.</div></div>`:'');

  // ── MONTE CARLO ────────────────────────────────────────────────────────────
  const mcHTML=d.mc?`<div class="kap-model">
    <div class="kap-model-hd">Model 3 — Monte Carlo simulace průměrného dne</div>
    <div style="font-size:.72rem;color:#64748b;margin-bottom:10px;">
      Poisson(λ) per hodina · efektivní přítomnost ${{Math.round(eff*100)}}% · 2 varianty
    </div>
    <div class="mc-2col">
      ${{_mcPanel(d.mc,C.MC_ITERATIONS,d.has_svc)}}
      ${{d.mc_boost?_mcPanel(d.mc_boost,C.MC_ITERATIONS,d.has_svc):`<div style="display:flex;align-items:center;justify-content:center;padding:20px;color:#94a3b8;font-size:.78rem;">Varianta +20% není dostupná</div>`}}
    </div>
    <div class="grid2" style="margin-top:4px;">
      <div>
        <div style="font-size:.68rem;font-weight:700;color:#475569;margin-bottom:2px;">Varianta 1 — reálné návštěvy</div>
        ${{_mcGap(d.mc,d.has_svc,d.total,nd)}}
      </div>
      <div>
        <div style="font-size:.68rem;font-weight:700;color:#475569;margin-bottom:2px;">Varianta 2 — +20 % návštěv</div>
        ${{d.mc_boost?_mcGap(d.mc_boost,d.has_svc,Math.round(d.total*1.2),nd):'<span style="font-size:.68rem;color:#94a3b8;">N/A</span>'}}
      </div>
    </div>
    ${{d.mc.bankers_for_95!=null?`<div style="font-size:.75rem;padding:8px 12px;border-radius:8px;margin-top:8px;
        background:${{d.mc.bankers_for_95===d.mc.current_bankers?'#f0fdf4':'#f0f4ff'}};
        color:${{d.mc.bankers_for_95===d.mc.current_bankers?'#15803d':'#475569'}};">
      ${{d.mc.bankers_for_95===d.mc.current_bankers?`✅ Současný počet bankéřů (${{fmt1(d.mc.current_bankers)}}) postačuje pro 95% pokrytí.`:`Doporučení: <b>${{fmt1(d.mc.bankers_for_95)}} bankéřů</b> pro 95% pokrytí (nyní ${{fmt1(d.mc.current_bankers)}})`}}
    </div>`:''}}
    ${{hourlyModelsSec(d)}}
  </div>`:'';

  return`<div class="kap-card">
    <div class="kap-title">⚡ Kapacita pobočky &nbsp;·&nbsp; efektivní přítomnost ${{Math.round(eff*100)}}% &nbsp;·&nbsp; absence ${{Math.round(C.ABSENCE_RATE*100)}}%</div>
    ${{m1HTML}}${{m2HTML}}${{mcHTML}}
    ${{_roomsInline()}}
  </div>`;
}}

// ── SVG FTE line chart: P95 hourly FTE demand vs. capacity ───────────────────
function _svgFteLine(mc,hasSvc,W,H){{
  if(!mc||!mc.p95_ob_fte)return'';
  const pl=32,pr=12,pt=8,pb=18,pw=W-pl-pr,ph=H-pt-pb;
  const n=MCH.length;
  const xs=i=>pl+i/(n-1)*pw;
  const obFte=mc.p95_ob_fte,svcFte=mc.p95_svc_fte||[];
  const capOB=mc.ob_cap_fte||0,capSVC=mc.svc_cap_fte||0;
  const maxY=Math.max(...obFte,...(hasSvc?svcFte:[]),capOB,hasSvc?capSVC:0)*1.2||1;
  const ys=v=>pt+ph*(1-Math.min(v,maxY)/maxY);
  // Y ticks
  const step=maxY<=1.5?0.25:maxY<=3?0.5:maxY<=6?1:maxY<=12?2:5;
  let yts=[]; for(let v=0;v<=maxY*1.05;v+=step)yts.push(parseFloat(v.toFixed(2)));
  const ytkEl=yts.map(v=>`<line x1="${{pl}}" y1="${{ys(v).toFixed(1)}}" x2="${{pl+pw}}" y2="${{ys(v).toFixed(1)}}" stroke="#e2e8f0" stroke-width="0.5"/><text x="${{pl-3}}" y="${{ys(v).toFixed(1)}}" text-anchor="end" dominant-baseline="middle" font-size="7" fill="#94a3b8">${{v}}</text>`).join('');
  // X labels (every 2h)
  const xtkEl=MCH.map((h,i)=>i%2===0?`<text x="${{xs(i).toFixed(1)}}" y="${{H-3}}" text-anchor="middle" font-size="7" fill="#94a3b8">${{h}}h</text>`:'').join('');
  // Capacity dashed lines
  const capOBY=ys(capOB);
  const capLines=`<line x1="${{pl}}" y1="${{capOBY.toFixed(1)}}" x2="${{pl+pw}}" y2="${{capOBY.toFixed(1)}}" stroke="#2563eb" stroke-dasharray="5,3" stroke-width="1.5"/>
    <text x="${{pl+pw+2}}" y="${{capOBY.toFixed(1)}}" dominant-baseline="middle" font-size="7" fill="#2563eb">OB ${{fmt1(capOB)}}</text>
    ${{hasSvc?`<line x1="${{pl}}" y1="${{ys(capSVC).toFixed(1)}}" x2="${{pl+pw}}" y2="${{ys(capSVC).toFixed(1)}}" stroke="#d97706" stroke-dasharray="5,3" stroke-width="1.5"/><text x="${{pl+pw+2}}" y="${{ys(capSVC).toFixed(1)}}" dominant-baseline="middle" font-size="7" fill="#d97706">SVC ${{fmt1(capSVC)}}</text>`:''}}`;
  // Lines
  const obPts=obFte.map((v,i)=>`${{i===0?'M':'L'}}${{xs(i).toFixed(1)}},${{ys(v).toFixed(1)}}`).join(' ');
  const svcPts=hasSvc?svcFte.map((v,i)=>`${{i===0?'M':'L'}}${{xs(i).toFixed(1)}},${{ys(v).toFixed(1)}}`).join(' '):'';
  // Fill area under OB line
  const areaOB=`${{obPts}} L${{xs(n-1).toFixed(1)}},${{(pt+ph).toFixed(1)}} L${{pl}},${{(pt+ph).toFixed(1)}} Z`;
  return`<svg viewBox="0 0 ${{W}} ${{H}}" style="width:100%;display:block;">
    <rect x="${{pl}}" y="${{pt}}" width="${{pw}}" height="${{ph}}" fill="#f8faff" rx="2"/>
    ${{ytkEl}}${{xtkEl}}${{capLines}}
    <path d="${{areaOB}}" fill="#2563eb" fill-opacity="0.06"/>
    ${{hasSvc?`<path d="${{svcPts}}" fill="none" stroke="#d97706" stroke-width="1.5" stroke-linejoin="round"/>`:''}}
    <path d="${{obPts}}" fill="none" stroke="#2563eb" stroke-width="2" stroke-linejoin="round"/>
    <line x1="${{pl}}" y1="${{pt}}" x2="${{pl}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
    <line x1="${{pl}}" y1="${{pt+ph}}" x2="${{pl+pw}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
  </svg>`;
}}

// ── SVG histogram: daily FTE demand distribution ───────────────────────────────
function _svgHist(counts,edges,capV,p95V,color,W,H){{
  if(!counts||!counts.length||!edges||edges.length<2)return'';
  const pl=32,pr=6,pt=8,pb=18,pw=W-pl-pr,ph=H-pt-pb;
  const maxX=edges[edges.length-1]||1,maxC=Math.max(...counts,1);
  const xs=x=>pl+Math.min(x,maxX)/maxX*pw;
  const ys=v=>pt+ph*(1-v/maxC);
  // Bars
  const bars=counts.map((c,i)=>{{
    const x1=xs(edges[i]),x2=xs(edges[i+1]),w=Math.max(x2-x1-0.5,0.5),y=ys(c),h=ph-(y-pt);
    return`<rect x="${{x1.toFixed(1)}}" y="${{y.toFixed(1)}}" width="${{w.toFixed(1)}}" height="${{h.toFixed(1)}}" fill="${{color}}" fill-opacity="0.75"/>`;
  }}).join('');
  // Vertical reference lines
  const capX=xs(Math.min(capV,maxX)),p95X=xs(Math.min(p95V,maxX));
  const capLn=`<line x1="${{capX.toFixed(1)}}" y1="${{pt}}" x2="${{capX.toFixed(1)}}" y2="${{pt+ph}}" stroke="#475569" stroke-dasharray="4,3" stroke-width="1.5"/>`;
  const p95Ln=p95V>0?`<line x1="${{p95X.toFixed(1)}}" y1="${{pt}}" x2="${{p95X.toFixed(1)}}" y2="${{pt+ph}}" stroke="#16a34a" stroke-dasharray="4,3" stroke-width="1.5"/>`:'';
  // X ticks (5 evenly spaced)
  const xTk=[0,0.25,0.5,0.75,1.0].map(t=>{{const v=t*maxX,x=xs(v);return`<text x="${{x.toFixed(1)}}" y="${{H-3}}" text-anchor="middle" font-size="7" fill="#94a3b8">${{v.toFixed(1)}}</text>`;}}).join('');
  return`<svg viewBox="0 0 ${{W}} ${{H}}" style="width:100%;display:block;">
    <rect x="${{pl}}" y="${{pt}}" width="${{pw}}" height="${{ph}}" fill="#f8faff" rx="2"/>
    ${{bars}}${{capLn}}${{p95Ln}}
    <line x1="${{pl}}" y1="${{pt}}" x2="${{pl}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
    <line x1="${{pl}}" y1="${{pt+ph}}" x2="${{pl+pw}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
    ${{xTk}}
  </svg>`;
}}

// ── SVG multi-model hourly visits line chart ──────────────────────────────────
function _svgHourlyVisits(d,W,H){{
  if(!d.has_time)return'';
  const pl=34,pr=12,pt=8,pb=20,pw=W-pl-pr,ph=H-pt-pb;
  const n=MCH.length;
  const xs=i=>pl+i/(n-1)*pw;
  const actS=MCH.map(h=>TYPES.reduce((s,t)=>s+(d.by_hour[t.key]?.[h]||0),0));
  const aPD=d.total/(d.n_days||1);
  const m2PD=d.poc_kli>0&&d.annual_open_days>0?d.poc_kli/d.annual_open_days:0;
  const m2SF=aPD>0&&m2PD>0?m2PD/aPD:0;
  const m2S=m2SF>0?actS.map(v=>v*m2SF):null;
  const mcS=d.mc?MCH.map((_,i)=>(d.mc.lam_online[i]||0)+(d.mc.lam_fyzicka[i]||0)+(d.mc.lam_bezhot[i]||0)):null;
  const allV=[...actS,...(m2S||[]),...(mcS||[])];
  const maxY=Math.max(...allV,0.01)*1.2;
  const ys=v=>pt+ph*(1-Math.min(v,maxY)/maxY);
  const step=maxY<=2?0.5:maxY<=5?1:maxY<=10?2:maxY<=20?5:10;
  let yts=[];for(let v=0;v<=maxY*1.05;v+=step)yts.push(parseFloat(v.toFixed(1)));
  const ytkEl=yts.map(v=>`<line x1="${{pl}}" y1="${{ys(v).toFixed(1)}}" x2="${{pl+pw}}" y2="${{ys(v).toFixed(1)}}" stroke="#e2e8f0" stroke-width="0.5"/><text x="${{pl-3}}" y="${{ys(v).toFixed(1)}}" text-anchor="end" dominant-baseline="middle" font-size="7" fill="#94a3b8">${{v}}</text>`).join('');
  const xtkEl=MCH.map((h,i)=>i%2===0?`<text x="${{xs(i).toFixed(1)}}" y="${{H-3}}" text-anchor="middle" font-size="7" fill="#94a3b8">${{h}}h</text>`:'').join('');
  const lp=s=>s.map((v,i)=>`${{i===0?'M':'L'}}${{xs(i).toFixed(1)}},${{ys(v).toFixed(1)}}`).join(' ');
  const areaAct=`${{lp(actS)}} L${{xs(n-1).toFixed(1)}},${{(pt+ph).toFixed(1)}} L${{pl}},${{(pt+ph).toFixed(1)}} Z`;
  return`<svg viewBox="0 0 ${{W}} ${{H}}" style="width:100%;display:block;">
    <rect x="${{pl}}" y="${{pt}}" width="${{pw}}" height="${{ph}}" fill="#f8faff" rx="2"/>
    ${{ytkEl}}${{xtkEl}}
    <path d="${{areaAct}}" fill="#2563eb" fill-opacity="0.07"/>
    ${{m2S?`<path d="${{lp(m2S)}}" fill="none" stroke="#16a34a" stroke-width="1.8" stroke-dasharray="6,3" stroke-linejoin="round"/>`:''}}
    ${{mcS?`<path d="${{lp(mcS)}}" fill="none" stroke="#9333ea" stroke-width="1.5" stroke-dasharray="2,2" stroke-linejoin="round"/>`:''}}
    <path d="${{lp(actS)}}" fill="none" stroke="#2563eb" stroke-width="2" stroke-linejoin="round"/>
    <line x1="${{pl}}" y1="${{pt}}" x2="${{pl}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
    <line x1="${{pl}}" y1="${{pt+ph}}" x2="${{pl+pw}}" y2="${{pt+ph}}" stroke="#94a3b8" stroke-width="0.8"/>
  </svg>`;
}}

// ── Model 4 — Monte Carlo with FTE visualization ──────────────────────────────
function _mcPanel(mc,n,hasSvc){{
  if(!mc)return'';
  const ovHPD=mc.overload_prob.reduce((a,b)=>a+b,0)/100;
  const ovClr=ovHPD<0.5?'#15803d':ovHPD<2?'#d97706':'#b91c1c';
  return`<div style="flex:1;min-width:260px;background:#f8faff;border-radius:10px;padding:12px;">
    <div style="font-size:.76rem;font-weight:700;color:#1e3a8a;margin-bottom:8px;">n = ${{fmtI(n)}} simulací</div>
    <div style="display:flex;gap:6px;margin-bottom:10px;">
      <div class="card" style="flex:1;padding:6px 9px;">
        <div class="cl">Přetížení</div>
        <div style="font-size:1.4rem;font-weight:800;color:${{ovClr}};">${{fmt1(ovHPD)}} h</div>
        <div class="cs">/ den průměrně</div>
      </div>
      <div class="card" style="flex:1;padding:6px 9px;">
        <div class="cl">Pro 95% pokrytí</div>
        <div style="font-size:1.4rem;font-weight:800;">${{mc.bankers_for_95!=null?fmt1(mc.bankers_for_95):'>+3'}}</div>
        <div class="cs">bankéřů OB</div>
      </div>
      <div class="card" style="flex:1;padding:6px 9px;">
        <div class="cl">OB P95 špička</div>
        <div style="font-size:1.4rem;font-weight:800;">${{fmt1(Math.max(...mc.p95_ob_fte))}}</div>
        <div class="cs">FTE / hod</div>
      </div>
    </div>

    <div class="mc-chart-title">P95 FTE poptávka po hodinách (6h–21h)</div>
    <div style="font-size:.62rem;color:#64748b;margin-bottom:3px;">
      OB (modrá) ${{hasSvc?'· Servis BKP (oranžová) ':''}}· kapacita = přerušovaná čára (FTE = bankéř × ${{Math.round((1-C.ABSENCE_RATE)*100)}}% přítomnost)
    </div>
    <div class="mc-chart-wrap">${{_svgFteLine(mc,hasSvc,480,130)}}</div>

    <div class="mc-chart-title">Distribuce celkové denní FTE poptávky (bankéř-hodiny/den)</div>
    <div style="font-size:.62rem;color:#64748b;margin-bottom:3px;">Frekvence · šedá = kapacita · zelená = P95 poptávky</div>
    <div class="${{hasSvc?'mc-charts-row':''}}">
      <div>
        <div style="font-size:.65rem;color:#2563eb;font-weight:700;margin-bottom:2px;">
          OB tým (schůzky) — kapacita ${{fmt1(mc.ob_cap_day)}}h · P95 ${{fmt1(mc.ob_p95_day)}}h
        </div>
        <div class="mc-chart-wrap">${{_svgHist(mc.ob_hist,mc.ob_edges,mc.ob_cap_day,mc.ob_p95_day,'#2563eb',hasSvc?230:470,100)}}</div>
      </div>
      ${{hasSvc?`<div>
        <div style="font-size:.65rem;color:#d97706;font-weight:700;margin-bottom:2px;">
          Servisní zóna (BKP) — kapacita ${{fmt1(mc.svc_cap_day)}}h · P95 ${{fmt1(mc.svc_p95_day)}}h
        </div>
        <div class="mc-chart-wrap">${{_svgHist(mc.svc_hist,mc.svc_edges,mc.svc_cap_day,mc.svc_p95_day,'#d97706',230,100)}}</div>
      </div>`:''}}
    </div>
  </div>`;
}}

function hourlyModelsSec(d){{
  if(!d.has_time)return'';
  const aPD=d.total/(d.n_days||1);
  const m2PD=d.poc_kli>0&&d.annual_open_days>0?d.poc_kli/d.annual_open_days:0;
  const m2SF=aPD>0&&m2PD>0?m2PD/aPD:0;
  return`<div class="kap-model">
    <div class="kap-model-hd">📊 Hodinový profil — průměrné návštěvy / hodinu</div>
    <div style="font-size:.72rem;color:#64748b;margin-bottom:8px;">Srovnání skutečného profilu s projekcemi modelů (6–21h)</div>
    <div class="mc-chart-wrap">${{_svgHourlyVisits(d,520,140)}}</div>
    <div style="display:flex;gap:14px;flex-wrap:wrap;margin-top:6px;font-size:.68rem;color:#475569;">
      <span style="display:inline-flex;align-items:center;gap:5px;">
        <svg width="22" height="8" style="flex-shrink:0;"><line x1="0" y1="4" x2="22" y2="4" stroke="#2563eb" stroke-width="2"/></svg>
        Reálná data (M1, M4)</span>
      ${{m2SF>0?`<span style="display:inline-flex;align-items:center;gap:5px;">
        <svg width="22" height="8" style="flex-shrink:0;"><line x1="0" y1="4" x2="22" y2="4" stroke="#16a34a" stroke-width="1.8" stroke-dasharray="6,3"/></svg>
        M2/M3 projekce (${{fmt1(m2PD)}} nav./den)</span>`:''}}
      ${{d.mc?`<span style="display:inline-flex;align-items:center;gap:5px;">
        <svg width="22" height="8" style="flex-shrink:0;"><line x1="0" y1="4" x2="22" y2="4" stroke="#9333ea" stroke-width="1.5" stroke-dasharray="2,2"/></svg>
        MC λ střed</span>`:''}}
    </div>
  </div>`;
}}

// ── Format badge ─────────────────────────────────────────────────────────────
function formatBadge(d){{
  if(!d.branch_format)return'';
  const FM={{'flagship':{{bg:'#1e3a8a',lbl:'Flagship ≥25 FTE'}},
    'medium':{{bg:'#2563eb',lbl:'Střední 10–25 FTE'}},
    'medium economy':{{bg:'#7c3aed',lbl:'Economy 5–10 FTE'}},
    'small':{{bg:'#64748b',lbl:'Malá <5 FTE'}}}};
  const f=FM[d.branch_format]||{{bg:'#94a3b8',lbl:d.branch_format}};
  return`<span class="badge" style="background:${{f.bg}};color:#fff;font-size:.7rem;
      vertical-align:middle;margin-left:6px;">${{f.lbl}}</span>`;
}}

// ── Benchmark ─────────────────────────────────────────────────────────────────
function benchmarkSec(d){{
  if(!d.benchmark)return'';
  const bm=d.benchmark,own=bm.own,net=bm.network,fmtM=bm.format;
  function bRow(lbl,key,unit,hiGood){{
    const v=own[key],vn=net[key],vf=fmtM?fmtM[key]:null;
    const f=x=>x!=null?fmt1(x)+(unit?' '+unit:''):'—';
    function dir(a,b){{
      if(a==null||b==null||hiGood==null)return'';
      const diff=a-b;if(Math.abs(diff)<0.5)return'<span style="color:#94a3b8">≈</span>';
      return(hiGood?diff>0:diff<0)?'<span style="color:#16a34a;font-size:.82em">▲</span>'
                                  :'<span style="color:#dc2626;font-size:.82em">▼</span>';
    }}
    return`<tr>
      <td style="color:#475569;font-size:.82rem;padding:5px 9px;border-bottom:1px solid #f0f4ff;">${{lbl}}</td>
      <td style="font-weight:700;color:#1e3a8a;text-align:right;padding:5px 9px;border-bottom:1px solid #f0f4ff;">${{f(v)}}</td>
      <td style="text-align:right;padding:5px 9px;font-size:.82rem;color:#475569;border-bottom:1px solid #f0f4ff;">
        ${{f(vn)}} ${{dir(v,vn)}}</td>
      <td style="text-align:right;padding:5px 9px;font-size:.82rem;color:#7c3aed;border-bottom:1px solid #f0f4ff;">
        ${{vf!=null?f(vf)+' '+dir(v,vf):'—'}}</td></tr>`;
  }}
  const fmtLbl=d.branch_format?' ('+d.branch_format+')':'';
  return`<div class="sec"><div class="st">📊 Benchmark — porovnání s mediánem sítě</div>
    <table style="width:100%;border-collapse:collapse;">
      <thead><tr>
        <th style="text-align:left;font-size:.66rem;font-weight:700;text-transform:uppercase;
            letter-spacing:.4px;color:#94a3b8;border-bottom:1.5px solid #dbeafe;
            padding:4px 9px;">Metrika</th>
        <th style="text-align:right;font-size:.66rem;font-weight:700;text-transform:uppercase;
            letter-spacing:.4px;color:#1e3a8a;border-bottom:1.5px solid #dbeafe;
            padding:4px 9px;">Tato pobočka</th>
        <th style="text-align:right;font-size:.66rem;font-weight:700;text-transform:uppercase;
            letter-spacing:.4px;color:#64748b;border-bottom:1.5px solid #dbeafe;
            padding:4px 9px;">Síť (medián)</th>
        <th style="text-align:right;font-size:.66rem;font-weight:700;text-transform:uppercase;
            letter-spacing:.4px;color:#7c3aed;border-bottom:1.5px solid #dbeafe;
            padding:4px 9px;">Formát${{fmtLbl}}</th>
      </tr></thead>
      <tbody>
        ${{bRow('Návštěvy / den','visits_pd','',true)}}
        ${{bRow('Schůzky / bankéř / den','mtgs_pb_day','',true)}}
        ${{bRow('Kapacita OB','cap_util_ob','%',false)}}
        ${{bRow('Online schůzky','online_pct','%',null)}}
        ${{bRow('Fyzické schůzky','fyzicka_pct','%',null)}}
        ${{bRow('Bezhot. walkin','bezhot_pct','%',null)}}
        ${{bRow('Hotovostní walkin','hotovost_pct','%',null)}}
      </tbody>
    </table>
    <div style="font-size:.67rem;color:#94a3b8;margin-top:6px;">
      ▲ lepší než medián &nbsp;▼ horší &nbsp;≈ blízko mediánu (±0,5)
    </div></div>`;
}}

// ── Main render ───────────────────────────────────────────────────────────────
function render(id){{
  cur=id; const d=DATA[id];
  const score=_branchScore(d);
  const scoreCard=score!=null?`<div class="card" style="border-color:${{_scoreClr(score)}}40;background:${{_scoreBg(score)}};">
    <div class="cl">Skóre vytíženosti</div>
    <div class="cv" style="color:${{_scoreClr(score)}};">${{score}}%</div>
    <div class="cs">${{score>=90?'🔴 přetíženo':score>=70?'⚠️ blíží se':'✅ OK'}} · OB utilizace</div></div>`:'';
  const totCard=`<div class="card"><div class="cl">Celkem návštěv</div>
    <div class="cv">${{fmtI(d.total)}}</div>
    <div class="cs">rok 2025 · ${{d.n_days}} dnů · ${{d.annual_open_days||0}} ot. dnů/rok</div></div>`;
  const cliCard=d.poc_kli>0?`<div class="card"><div class="cl">Počet klientů</div>
    <div class="cv">${{fmtI(d.poc_kli)}}</div><div class="cs">portfoliová data</div></div>`:'';
  const totalMtgs=(d.by_type?.online||0)+(d.by_type?.fyzicka||0);
  const revMtgCard=d.sales&&d.sales.total_objem>0&&totalMtgs>0?`<div class="card" style="border-color:#bbf7d0;background:#f0fdf4;">
    <div class="cl">Výnos / schůzku</div>
    <div class="cv" style="color:#15803d;">${{fmtN(Math.round(d.sales.total_objem/totalMtgs))}}</div>
    <div class="cs">Kč · ${{fmtI(totalMtgs)}} schůzek</div></div>`:'';
  const revBankerCard=d.sales&&d.sales.total_objem>0&&d.bankers>0?`<div class="card" style="border-color:#bbf7d0;background:#f0fdf4;">
    <div class="cl">Výnos / bankéř</div>
    <div class="cv" style="color:#15803d;">${{fmtN(Math.round(d.sales.total_objem/d.bankers))}}</div>
    <div class="cs">Kč · ${{fmt1(d.bankers)}} bankéřů</div></div>`:'';
  const typeCards=TYPES.map(t=>{{
    const v=d.by_type[t.key]||0,pct=d.total>0?(v/d.total*100).toFixed(1):'0.0';
    return`<div class="card" style="border-color:${{t.color}}50;background:${{t.color}}0a;">
      <div class="cl" style="color:${{t.color}};">${{t.label}}</div>
      <div class="cv" style="color:${{t.color}};">${{fmtI(v)}}</div>
      <div class="cs">${{pct}} %</div></div>`;
  }}).join('');
  const unkNote=d.unknown>0
    ?`<div style="font-size:.7rem;color:#94a3b8;margin:3px 0 10px;">${{fmtI(d.unknown)}} návštěv bez určeného typu</div>`:'';

  document.getElementById('mc').innerHTML=`
<div style="font-size:1.1rem;font-weight:800;color:#1e3a8a;margin-bottom:12px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
  ${{d.name}}
  <span style="font-size:.78rem;font-weight:400;color:#94a3b8;">#${{id}}</span>
  ${{formatBadge(d)}}
</div>
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-bottom:10px;">
  ${{scoreCard}}${{totCard}}${{cliCard}}${{revMtgCard}}${{revBankerCard}}${{typeCards}}
</div>
${{unkNote}}
${{staffSec(d)}}
${{visitTrendSec(d)}}
${{capacitySec(d)}}
${{odSec(d)}}
${{benchmarkSec(d)}}
<details style="margin-bottom:14px;"><summary>📐 Metodika výpočtu (rozbalit)</summary>
  ${{methSec(d)}}</details>
`;}}
renderRanking();
</script>
</body>
</html>"""


# ─── Generování ────────────────────────────────────────────────────────────────

_df_visits = load_visits()
_kpis      = load_kpis()
_prof_kli, _prof_fte = load_profitabilita()
_spec      = load_specialiste()
_od        = load_oteviraci()
_sales     = load_sales()

_visit_data, _order, _has_type = build_data(_df_visits, _kpis, _prof_kli, _prof_fte, _spec, _od, _sales)
_visit_data = compute_benchmarks(_visit_data)
_html = render_html(_visit_data, _order, _has_type)

with open(OUTPUT_FILE, 'w', encoding='utf-8') as _f:
    _f.write(_html)

_n_names  = sum(1 for v in _visit_data.values() if not v['name'].startswith('Pobočka'))
_n_cap2   = sum(1 for v in _visit_data.values() if v.get('cap2') is not None)
_n_mc     = sum(1 for v in _visit_data.values() if v.get('mc') is not None)
_n_fmt    = sum(1 for v in _visit_data.values() if v.get('branch_format'))
_n_rooms  = sum(1 for v in _visit_data.values() if v.get('rooms'))
print(f"\n✅ Report uložen: {OUTPUT_FILE}")
print(f"   {len(_visit_data)} poboček · {_n_names} s názvem · typy={'✓' if _has_type else '✗'}")
print(f"   Formát: {_n_fmt} · MC: {_n_mc} · Místnosti: {_n_rooms} · Model 2: {_n_cap2}")
