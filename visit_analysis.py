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
WALKIN_CONVERT_PCT  = 0.10   # % walkinů přeroste ve schůzku 45 min
WALKIN_AVG_MINS     = WALKIN_SHORT_PCT * WALKIN_SHORT_MINS + WALKIN_LONG_PCT * WALKIN_LONG_MINS  # = 18
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
BANKER_COLS = {'OSOBNI_BANKER_-_JUNIOR', 'OSOBNI_BANKER_-_MEDIOR', 'OSOBNI_BANKER_-_SENIOR'}
SERVICE_COL = 'BANKER_KLIENTSKE_PECE_-_MEDIOR'
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

def _cap_model(online, fyzicka, bezhot, bankers, svc_fte, n_days):
    """Annual capacity metrics dict for given visit counts and staffing."""
    if bankers <= 0: return None
    eff       = 1.0 - ABSENCE_RATE          # effective presence factor (0.771)
    avail_ob  = bankers * n_days * WORK_MINS_DAY * eff
    avail_svc = svc_fte * n_days * WORK_MINS_DAY * eff

    online_mins  = online   * MEETING_MINS
    fyzicka_mins = fyzicka  * MEETING_MINS
    bezhot_base  = bezhot   * WALKIN_AVG_MINS
    bezhot_conv  = bezhot   * WALKIN_CONVERT_PCT * MEETING_MINS

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
    svc_needed = WALKIN_AVG_MINS * samples[:, :, 2]   # (n_iter, 16)

    eff = 1.0 - ABSENCE_RATE   # 0.771 — effective presence per hour

    if svc_fte <= 0:
        ob_needed = ob_needed + svc_needed   # OB-junior handles bezhot
        svc_overload = np.zeros((n_iter, len(MC_HOURS)), dtype=bool)
    else:
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

    return {
        'coverage_pct':    coverage_pct,
        'overload_prob':   overload_prob,   # per hour [6..21]
        'p50_util':        p50_util,        # median utilization % per hour
        'p95_util':        p95_util,        # 95th-pct utilization % per hour
        'bankers_for_95':  bankers_for_95,
        'current_bankers': round(bankers, 1),
        'n_iter':          n_iter,
        'lam_online':      [round(lam[i, 0], 2) for i in range(len(MC_HOURS))],
        'lam_fyzicka':     [round(lam[i, 1], 2) for i in range(len(MC_HOURS))],
        'lam_bezhot':      [round(lam[i, 2], 2) for i in range(len(MC_HOURS))],
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
        'service_desks':   math.ceil(pb * 18 / 60) if pb > 0 else 0,
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
            bankers = sum(float(row.get(c, 0) or 0) for c in pos_cols if c in BANKER_COLS)
            svc_fte = float(row.get(SERVICE_COL, 0) or 0)
            positions = {}
            for c in pos_cols:
                v = float(row.get(c, 0) or 0)
                if v > 0:
                    lbl = POSITION_LABELS.get(c, c.replace('_',' ').title())
                    positions[lbl] = round(v, 1)
            total_fte = sum(float(row.get(c, 0) or 0) for c in pos_cols)
            out[bid] = {
                'name':      str(row[nm_c]) if nm_c and pd.notna(row.get(nm_c,'')) else None,
                'bankers':   round(bankers, 1),
                'svc_fte':   round(svc_fte, 1),
                'has_svc':   svc_fte > 0,
                'positions': positions,
                'total_fte': round(total_fte, 1),
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
            out[bid] = {'is_vikend':is_vikend,'ph_tyden':ph,'od_days':days}
        print(f"   Ot.doba: {len(out)} poboček")
    except Exception as e:
        print(f"⚠️  Ot.doba: {e}")
    return out


# ─── Build data ────────────────────────────────────────────────────────────────

def build_data(df, kpis, prof_kli, prof_fte, spec, od):
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

        if has_date:
            n_days = max(int(vb['_dt'].dt.date.nunique()), 1)
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

        # Staff
        bankers = float(s.get('bankers', 0) or 0)
        svc_fte = float(s.get('svc_fte', 0) or 0)
        has_svc = bool(s.get('has_svc', False))
        positions = s.get('positions', {})
        # FTE priority: profitabilita xlsx → kpis pkl → specialiste total (fallback)
        fte = prof_fte.get(bid) or k.get('fte') or s.get('total_fte')

        # Client count: profitabilita first, then kpis fallback
        poc_kli = prof_kli.get(bid) or k.get('pocet_klientu') or 0

        # Metrics
        portfolio_pb = round(poc_kli / bankers) if bankers > 0 and poc_kli > 0 else None
        total_mtgs   = by_type['online'] + by_type['fyzicka']
        mtgs_pb_day  = round(total_mtgs / bankers / n_days, 1) if bankers > 0 and n_days > 0 else None

        # Capacity Model 1 (real visits)
        cap1 = _cap_model(by_type['online'], by_type['fyzicka'], by_type['bezhot'],
                          bankers, svc_fte, n_days)

        # Capacity Model 2 (client-based)
        cap2 = None
        total_excl_hot = sum(by_type[k2] for k2 in ['online','fyzicka','bezhot'])
        if poc_kli > 0 and total_excl_hot > 0 and bankers > 0:
            r_on  = by_type['online']  / total_excl_hot
            r_fyz = by_type['fyzicka'] / total_excl_hot
            r_beh = by_type['bezhot']  / total_excl_hot
            cap2  = _cap_model(poc_kli * r_on, poc_kli * r_fyz, poc_kli * r_beh,
                               bankers, svc_fte, WORKING_DAYS)

        # Annual "1x per client" scenario: every client needs 1 meeting/year
        annual_1x = None
        if poc_kli > 0 and bankers > 0:
            eff = 1.0 - ABSENCE_RATE
            avail_ob_yr = bankers * WORKING_DAYS * WORK_MINS_DAY * eff
            mins_needed = poc_kli * MEETING_MINS
            annual_1x = {
                'poc_kli':        poc_kli,
                'mins_needed':    round(mins_needed),
                'avail_ob':       round(avail_ob_yr),
                'util_ob':        round(mins_needed / avail_ob_yr * 100, 1) if avail_ob_yr > 0 else 0,
                'mtgs_pb_day':    round(poc_kli / bankers / WORKING_DAYS, 1),
                'bankers':        round(bankers, 1),
                'bankers_needed': round(mins_needed / (WORKING_DAYS * WORK_MINS_DAY * eff), 1),
            }

        # Monte Carlo (only when hourly data is available)
        mc = None
        if has_time and has_date and bankers > 0:
            mc = run_monte_carlo(by_hour, bankers, svc_fte)

        branch_format = compute_format(fte)
        rooms         = compute_rooms(mc)

        result[str(bid)] = {
            'name':       name,
            'fte':        fte,
            'bankers':    bankers,
            'svc_fte':    svc_fte,
            'has_svc':    has_svc,
            'positions':  positions,
            'poc_kli':    poc_kli,
            'total':      total,
            'by_type':    by_type,
            'unknown':    unknown,
            'by_month':   by_month,
            'by_weekday': by_weekday,
            'by_hour':    by_hour,
            'heatmap':    heatmap,
            'has_time':   has_time and has_date,
            'n_days':     n_days,
            'portfolio_pb': portfolio_pb,
            'mtgs_pb_day':  mtgs_pb_day,
            'cap1':       cap1,
            'cap2':       cap2,
            'annual_1x':  annual_1x,
            'mc':           mc,
            'branch_format': branch_format,
            'rooms':         rooms,
            'is_vikend':  o.get('is_vikend', False),
            'ph_tyden':   o.get('ph_tyden',  0.0),
            'od_days':    o.get('od_days',   []),
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
const DATA  = {data_js};
const ORDER = {order_js};
const TYPES = {types_js};
const MCH   = {mc_hours_js};
const C     = {consts_js};
const WD    = ['Po','Út','St','Čt','Pá','So','Ne'];
const MON   = ['Led','Úno','Bře','Dub','Kvě','Čvn','Čvc','Srp','Zář','Říj','Lis','Pro'];

let cur=null;
const si=document.getElementById('si'),bl=document.getElementById('bl'),sw=document.getElementById('sw');

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

function capDetails(cap,title){{
  if(!cap)return'';
  const svc=cap.util_svc!=null?capBar(cap.util_svc,'Servisní bankéř — bezhotovostní walkin'):'';
  return`<div style="margin-bottom:14px;">
    <b style="font-size:.85rem;color:#1e3a8a;">${{title}}</b>
    <div style="margin-top:8px;">${{capBar(cap.util_ob,'OB tým — schůzky + online + konverze walkin')}}${{svc}}</div>
    <div style="font-size:.7rem;color:#64748b;margin-top:6px;display:flex;flex-wrap:wrap;gap:12px;">
      <span>📅 Dnů: <b>${{cap.n_days}}</b></span>
      <span>👤 OB bankéři: <b>${{fmt1(cap.bankers)}}</b></span>
      <span>⏱️ OB dostupné: <b>${{fmtH(cap.avail_ob)}}</b></span>
      <span>🔵 Online: <b>${{fmtH(cap.online_mins)}}</b></span>
      <span>🤝 Fyzické: <b>${{fmtH(cap.fyzicka_mins)}}</b></span>
      <span>🚶 Walkin základ: <b>${{fmtH(cap.bezhot_base)}}</b></span>
      <span>🔄 Walkin → schůzka: <b>${{fmtH(cap.bezhot_conv)}}</b></span>
    </div></div>`;
}}

function capSec(d){{
  if(!d.cap1&&!d.cap2)return'';
  return`<div class="sec"><div class="st">⚡ Kapacitní analýza</div>
    ${{capDetails(d.cap1,'Model 1 — reálná data (skutečné návštěvy '+fmtI(d.n_days)+' dnů)')}}
    ${{d.cap2
        ? capDetails(d.cap2,'Model 2 — klientský model ('+fmtI(d.poc_kli)+' klientů × poměry typů návštěv)')
        : d.poc_kli===0
          ? '<div style="font-size:.78rem;color:#94a3b8;padding:8px 0;">'+
            '⚠️ Model 2 není dostupný — data o počtu klientů nebyla načtena (profitabilita.xlsx).</div>'
          : ''}}
    </div>`;
}}

// ── Monte Carlo ───────────────────────────────────────────────────────────────
function mcSec(d){{
  if(!d.mc)return'';
  const mc=d.mc;
  const hours=MCH;
  const n=hours.length;

  // Coverage card color
  const cov=mc.coverage_pct;
  const covClr=cov>=95?'#15803d':cov>=80?'#d97706':'#b91c1c';
  const covBg =cov>=95?'#f0fdf4':cov>=80?'#fffbeb':'#fff5f5';

  // Per-hour bar: overload probability
  const maxProb=Math.max(...mc.overload_prob,1);
  const bars=hours.map((h,i)=>{{
    const prob=mc.overload_prob[i]||0;
    const p50 =mc.p50_util[i]||0;
    const p95 =mc.p95_util[i]||0;
    const barClr=prob>25?'#ef4444':prob>10?'#f59e0b':prob>2?'#3b82f6':'#93c5fd';
    const h_bar=Math.max(prob/Math.max(maxProb,1)*80,1).toFixed(1);
    return`<div class="mcbar-col">
      <div style="font-size:.57rem;color:#94a3b8;text-align:center;">${{prob>0?prob.toFixed(0)+'%':''}}</div>
      <div class="mcbar-inner" style="height:${{h_bar}}px;background:${{barClr}};border-radius:3px 3px 0 0;"
           title="${{h}}h — přetížení: ${{prob}}% | P50 využití: ${{p50}}% | P95: ${{p95}}%"></div>
      <div style="font-size:.58rem;color:#64748b;">${{h}}h</div>
    </div>`;
  }}).join('');

  // P95 utilization line chart as inline bar
  const maxP95=Math.max(...mc.p95_util,1);
  const p95bars=hours.map((h,i)=>{{
    const v=mc.p95_util[i]||0;
    const clr=v>100?'#ef4444':v>80?'#f59e0b':'#2563eb';
    const hw=Math.max(v/Math.max(maxP95,1)*80,1).toFixed(1);
    const border=v>100?' outline:2px solid #1e3a8a;outline-offset:-2px;':'';
    return`<div class="mcbar-col">
      <div style="font-size:.57rem;color:#94a3b8;text-align:center;">${{v>0?Math.round(v)+'%':''}}</div>
      <div style="height:${{hw}}px;background:${{clr}};border-radius:3px 3px 0 0;width:100%;${{border}}"
           title="${{h}}h — P95 využití: ${{v}}%"></div>
      <div style="font-size:.58rem;color:#64748b;">${{h}}h</div>
    </div>`;
  }}).join('');

  const b95note = mc.bankers_for_95!=null
    ? (mc.bankers_for_95===mc.current_bankers
        ? `✅ Současný počet bankéřů (${{fmt1(mc.current_bankers)}}) postačuje pro 95% pokrytí.`
        : `Doporučení: <b>${{fmt1(mc.bankers_for_95)}} bankéřů</b> pro 95% pokrytí (nyní ${{fmt1(mc.current_bankers)}})`)
    : `Ani s +3 bankéři nedosaženo 95 % pokrytí.`;

  return`<div class="sec"><div class="st">🎲 Monte Carlo simulace průměrného dne (n=${{mc.n_iter}})</div>
    <div class="grid3" style="margin-bottom:14px;">
      <div class="card" style="background:${{covBg}};border-color:${{covClr}}30;">
        <div class="cl">Pokrytí dne</div>
        <div class="cv" style="color:${{covClr}};font-size:1.8rem;">${{cov}}%</div>
        <div class="cs">dnů bez přetížení / ${{mc.n_iter}} simulací</div>
      </div>
      <div class="card"><div class="cl">P95 využití (špička)</div>
        <div class="cv">${{fmt1(Math.max(...mc.p95_util))}}%</div>
        <div class="cs">max. P95 využití v hodině</div></div>
      <div class="card"><div class="cl">Staffing pro 95 %</div>
        <div class="cv">${{mc.bankers_for_95!=null?fmt1(mc.bankers_for_95):'> +3'}}</div>
        <div class="cs">bankéřů OB</div></div>
    </div>

    <div style="font-size:.76rem;color:#475569;margin-bottom:6px;font-weight:600;">
      📊 Pravděpodobnost přetížení pobočky v dané hodině</div>
    <div class="mcbar-wrap" style="height:90px;align-items:flex-end;">${{bars}}</div>
    <div style="font-size:.68rem;color:#94a3b8;margin-bottom:14px;">
      🔴 &gt;25%  🟡 10–25%  🔵 2–10%  <span style="color:#93c5fd;">■</span> &lt;2%  —
      výška = P(přetížení) v dané hodině</div>

    <div style="font-size:.76rem;color:#475569;margin-bottom:6px;font-weight:600;">
      📈 P95 využití OB kapacity v dané hodině</div>
    <div class="mcbar-wrap" style="height:90px;align-items:flex-end;">${{p95bars}}</div>
    <div style="font-size:.68rem;color:#94a3b8;margin-bottom:14px;">
      🔴 &gt;100 % (přetíženo)  🟡 80–100 %  🔵 &lt;80 % —
      výška = 95. percentil využití za ${{mc.n_iter}} simulovaných dnů</div>

    <div style="font-size:.75rem;color:#475569;padding:8px 12px;background:#f0f4ff;
                border-radius:8px;">${{b95note}}</div>
  </div>`;
}}

// ── Staff section ─────────────────────────────────────────────────────────────
function staffSec(d){{
  if(!d.bankers&&!Object.keys(d.positions||{{}}).length)return'';
  const rows=Object.entries(d.positions||{{}}).map(([l,v])=>
    `<tr><td style="color:#475569;">${{l}}</td>
     <td style="font-weight:700;color:#1d4ed8;text-align:right;min-width:40px;">${{fmt1(v)}}</td></tr>`
  ).join('');
  const svcN=d.has_svc
    ?`<span class="badge" style="background:#dbeafe;color:#1d4ed8;">BKP Medior (servis)</span>`
    :`<span class="badge" style="background:#fef9c3;color:#854d0e;">OB Junior — fallback servis</span>`;
  return`<div class="sec"><div class="st">👤 Personální obsazení</div>
    <div class="grid2" style="margin-bottom:10px;">
      <div class="card"><div class="cl">Bankéři OB</div>
        <div class="cv" style="color:#1d4ed8;">${{fmt1(d.bankers)}}</div>
        <div class="cs">OB junior / medior / senior</div></div>
      <div class="card"><div class="cl">Servisní bankéř</div>
        <div style="margin-top:4px;">${{svcN}}</div>
        ${{d.svc_fte>0?`<div class="cs">${{fmt1(d.svc_fte)}} FTE</div>`:''}}</div>
    </div>
    ${{rows?`<table class="post"><tbody>${{rows}}</tbody></table>`:''}}
  </div>`;
}}

// ── Metrics ───────────────────────────────────────────────────────────────────
function metricsSec(d){{
  const por=d.portfolio_pb, mtg=d.mtgs_pb_day;
  if(por==null&&mtg==null)return'';
  const porCard=por!=null?`<div class="card ${{por<=C.TARGET_PORTFOLIO?'ok':'warn-c'}}">
    <div class="cl">Portfolio / bankéř</div>
    <div class="cv" style="color:${{por<=C.TARGET_PORTFOLIO?'#15803d':'#b91c1c'}}">${{fmtI(por)}}</div>
    <div class="cs">cíl ≤${{C.TARGET_PORTFOLIO}} kl. · ${{por<=C.TARGET_PORTFOLIO
      ?'✅ volná kap. '+(C.TARGET_PORTFOLIO-por)+' kl.'
      :'⚠️ překračuje o '+(por-C.TARGET_PORTFOLIO)+' kl.'}}</div></div>`:'';
  const mtgCard=mtg!=null?`<div class="card ${{mtg>=C.TARGET_MTGS_MIN&&mtg<=C.TARGET_MTGS_MAX?'ok':''}}">
    <div class="cl">Schůzky / bankéř / den</div>
    <div class="cv" style="color:${{mtg>=C.TARGET_MTGS_MIN&&mtg<=C.TARGET_MTGS_MAX?'#15803d':
      mtg<C.TARGET_MTGS_MIN?'#64748b':'#b91c1c'}}">${{fmt1(mtg)}}</div>
    <div class="cs">cíl ${{C.TARGET_MTGS_MIN}}–${{C.TARGET_MTGS_MAX}} · ${{
      mtg>=C.TARGET_MTGS_MIN&&mtg<=C.TARGET_MTGS_MAX?'✅ v normě':
      mtg<C.TARGET_MTGS_MIN?'⬇️ pod cílem':'⬆️ nad cílem'}}</div></div>`:'';
  return`<div class="grid2">${{porCard}}${{mtgCard}}</div>`;
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

// ── Rooms recommendation ──────────────────────────────────────────────────────
function roomSec(d){{
  if(!d.rooms)return'';
  const r=d.rooms;
  // λ = průměr příchodů v nejfrekventovanější hodině (z MC dat)
  // P95 = λ + 1,645·√λ  (95. percentil Poissonova rozdělení)
  // Δ   = P95 − λ        (bezpečnostní rezerva nad průměrem, pokrývá poptávkové špičky)
  const mtgNote=r.peak_mtg_lam>0
    ?`λ = ${{fmt1(r.peak_mtg_lam)}}/hod → P95 = ${{fmt1(r.p95_mtg)}} (+Δ ${{fmt1(r.delta_mtg)}}) → ⌈${{fmt1(r.p95_mtg)}}×45/60⌉`:'';
  const behNote=r.peak_bezhot_lam>0
    ?`λ = ${{fmt1(r.peak_bezhot_lam)}}/hod → P95 = ${{fmt1(r.p95_bezhot)}} (+Δ ${{fmt1(r.delta_bezhot)}}) → ⌈${{fmt1(r.p95_bezhot)}}×18/60⌉`:'';
  return`<div class="sec"><div class="st">🏢 Doporučení prostor</div>
    <div class="grid2" style="margin-bottom:10px;">
      <div class="card">
        <div class="cl">Zasedací místnosti</div>
        <div class="cv" style="color:#1d4ed8;">${{r.meeting_rooms}}</div>
        <div class="cs" style="font-size:.68rem;">${{mtgNote}}</div>
      </div>
      <div class="card">
        <div class="cl">Servisní místa</div>
        <div class="cv" style="color:#0891b2;">${{r.service_desks}}</div>
        <div class="cs" style="font-size:.68rem;">${{behNote}}</div>
      </div>
    </div>
    <div style="font-size:.72rem;color:#64748b;padding:10px 12px;background:#f0f4ff;border-radius:8px;line-height:1.6;">
      <b>Jak číst výpočet:</b><br>
      <b>λ</b> (lambda) = průměrný počet příchodů v <i>nejfrekventovanější</i> hodině dne (z Monte Carlo dat).<br>
      <b>P95</b> = λ + 1,645·√λ — 95. percentil Poissonova rozdělení. V 95 % dnů přijde nejvýše P95 klientů za tuto hodinu.<br>
      <b>Δ = P95 − λ</b> = bezpečnostní rezerva nad průměrem (pokrývá náhodné špičky v poptávce).<br>
      <b>Místnosti</b> = ⌈P95 × 45 min ÷ 60 min⌉ — kolik místností musí být souběžně dostupných.<br>
      <b>Servisní místa</b> = ⌈P95 × 18 min ÷ 60 min⌉ — analogicky pro bezhot. walkin.
    </div></div>`;
}}

// ── Annual 1× per client ──────────────────────────────────────────────────────
function annualMtgSec(d){{
  if(!d.annual_1x)return'';
  const a=d.annual_1x, eff=1-C.ABSENCE_RATE;
  const clr=a.util_ob<70?'#2563eb':a.util_ob<90?'#f59e0b':'#ef4444';
  const bg =a.util_ob<70?'#dbeafe':a.util_ob<90?'#fef3c7':'#fee2e2';
  const ico=a.util_ob<70?'✅':a.util_ob<90?'⚠️':'🔴';
  const suffix=a.util_ob<70?'kapacita postačuje':a.util_ob<90?'blíží se limitu':'přetíženo';
  return`<div class="sec">
    <div class="st">📋 Scénář — 1 schůzka ročně s každým klientem</div>
    <div style="font-size:.78rem;color:#475569;margin-bottom:10px;">
      Pokud by se každý z <b>${{fmtI(a.poc_kli)}}</b> klientů setkal s bankéřem právě 1× ročně
      na ${{C.MEETING_MINS}}&thinsp;min, jakou kapacitu OB týmu by to vyžadovalo?
    </div>
    <div class="grid3" style="margin-bottom:12px;">
      <div class="card">
        <div class="cl">Schůzky / bankéř / den</div>
        <div class="cv" style="color:${{a.mtgs_pb_day>=C.TARGET_MTGS_MIN&&a.mtgs_pb_day<=C.TARGET_MTGS_MAX?'#15803d':a.mtgs_pb_day>C.TARGET_MTGS_MAX?'#b91c1c':'#64748b'}};">${{fmt1(a.mtgs_pb_day)}}</div>
        <div class="cs">cíl ${{C.TARGET_MTGS_MIN}}–${{C.TARGET_MTGS_MAX}}/den · ${{C.WORKING_DAYS}} prac. dnů</div>
      </div>
      <div class="card">
        <div class="cl">Potřebný čas OB</div>
        <div class="cv">${{fmtH(a.mins_needed)}}</div>
        <div class="cs">${{fmtI(a.poc_kli)}} kl. × ${{C.MEETING_MINS}} min</div>
      </div>
      <div class="card">
        <div class="cl">Bankéři potřební</div>
        <div class="cv" style="color:${{a.bankers_needed<=a.bankers?'#15803d':'#b91c1c'}};">${{fmt1(a.bankers_needed)}}</div>
        <div class="cs">k dispozici ${{fmt1(a.bankers)}} · efekt. přítomnost ${{Math.round(eff*100)}}%</div>
      </div>
    </div>
    <div style="margin-bottom:6px;font-size:.7rem;color:#64748b;">Využití OB kapacity pro 1×/rok scénář</div>
    <div class="cbwrap" style="background:${{bg}};">
      <div class="cbfill" style="width:${{Math.min(a.util_ob,100)}}%;background:${{clr}};">${{a.util_ob.toFixed(1)}}%</div>
    </div>
    <div style="font-size:.67rem;color:#94a3b8;text-align:right;margin-top:2px;">
      ${{ico}} ${{suffix}} &nbsp;·&nbsp; dostupné OB hodiny: ${{fmtH(a.avail_ob)}}
      (s ${{Math.round(C.ABSENCE_RATE*100)}}% absencí)
    </div>
  </div>`;
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
  const wdLen=d.is_vikend?7:5, wdLabels=WD.slice(0,wdLen);

  const totCard=`<div class="card"><div class="cl">Celkem návštěv</div>
    <div class="cv">${{fmtI(d.total)}}</div><div class="cs">rok 2025 · ${{d.n_days}} dnů</div></div>`;
  const fteCard=d.fte!=null?`<div class="card"><div class="cl">FTE celkem</div>
    <div class="cv">${{fmt1(d.fte)}}</div><div class="cs">přepočtené úvazky</div></div>`:'';
  const cliCard=d.poc_kli>0?`<div class="card"><div class="cl">Počet klientů</div>
    <div class="cv">${{fmtI(d.poc_kli)}}</div><div class="cs">portfoliová data</div></div>`:'';
  const typeCards=TYPES.map(t=>{{
    const v=d.by_type[t.key]||0,pct=d.total>0?(v/d.total*100).toFixed(1):'0.0';
    return`<div class="card" style="border-color:${{t.color}}50;background:${{t.color}}0a;">
      <div class="cl" style="color:${{t.color}};">${{t.label}}</div>
      <div class="cv" style="color:${{t.color}};">${{fmtI(v)}}</div>
      <div class="cs">${{pct}} %</div></div>`;
  }}).join('');

  const tArr=TYPES.map(t=>({{label:t.label,color:t.color,values:d.by_month[t.key]||Array(12).fill(0)}}));
  const wdArr=TYPES.map(t=>({{label:t.label,color:t.color,
    values:(d.by_weekday[t.key]||Array(7).fill(0)).slice(0,wdLen)}}));
  let hrSec='';
  if(d.has_time){{
    const hrs=MCH;
    const hArr=TYPES.map(t=>({{label:t.label,color:t.color,
      values:hrs.map(h=>d.by_hour[t.key]?.[h]||0)}}));
    hrSec=`<div class="sec"><div class="st">⏱️ Průměrný den — avg návštěv/hodinu (6–21h)</div>
      <div class="bars" style="height:90px;">${{stackedBars(hrs.map(h=>h+'h'),hArr,80)}}</div>
      ${{legend()}}</div>`;
  }}
  const unkNote=d.unknown>0
    ?`<div style="font-size:.7rem;color:#94a3b8;margin:3px 0 10px;">
       ${{fmtI(d.unknown)}} návštěv bez určeného typu</div>`:'';
  const wdNote=!d.is_vikend
    ?`<div style="font-size:.68rem;color:#cbd5e1;margin-top:4px;">So/Ne skryty — nevíkendová pobočka</div>`:'';

  document.getElementById('mc').innerHTML=`
<div style="font-size:1.05rem;font-weight:700;color:#1e3a8a;margin-bottom:12px;">
  ${{d.name}} <span style="font-size:.78rem;font-weight:400;color:#94a3b8;">#${{id}}</span>
  ${{formatBadge(d)}}</div>
<div class="grid4">${{totCard}}${{fteCard}}${{cliCard}}${{typeCards}}</div>
${{unkNote}}
${{metricsSec(d)}}
${{staffSec(d)}}
${{benchmarkSec(d)}}
${{capSec(d)}}
${{annualMtgSec(d)}}
${{mcSec(d)}}
${{roomSec(d)}}
<div class="sec"><div class="st">📅 Návštěvy dle měsíce</div>
  <div class="bars" style="height:110px;">${{stackedBars(MON,tArr,100)}}</div>${{legend()}}</div>
<div class="sec"><div class="st">📆 Návštěvy dle dne v týdnu</div>
  <div class="bars" style="height:100px;">${{stackedBars(wdLabels,wdArr,90)}}</div>
  ${{legend()}}${{wdNote}}</div>
${{hrSec}}
${{heatmapSec(d)}}
${{odSec(d)}}
<details style="margin-bottom:14px;"><summary>📐 Metodika výpočtu (rozbalit)</summary>
  ${{methSec(d)}}</details>
`;}}
</script>
</body>
</html>"""


# ─── Generování ────────────────────────────────────────────────────────────────

_df_visits = load_visits()
_kpis      = load_kpis()
_prof_kli, _prof_fte = load_profitabilita()
_spec      = load_specialiste()
_od        = load_oteviraci()

_visit_data, _order, _has_type = build_data(_df_visits, _kpis, _prof_kli, _prof_fte, _spec, _od)
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
