#!/usr/bin/env python3
"""
visit_analysis.py
─────────────────
Analýza návštěvnosti a kapacity poboček 2025.

Vstupy : ../in/tables/VISITS_2025.csv
         kpis_grouped_2026.pkl
         export_specialiste.pkl
         ../vypocet_ir_2026/zdroje/report_od_pobocky_dbs_04_2026.xlsx
Výstup : report_navstevnost.html
"""

import os, sys, json, unicodedata, re, subprocess
import pandas as pd
import numpy as np

for _pkg in ['openpyxl']:
    try:
        __import__(_pkg)
    except ImportError:
        print(f"Instaluji {_pkg}…")
        subprocess.check_call([sys.executable, "-m", "pip", "install", _pkg])

# ─── Paths ─────────────────────────────────────────────────────────────────────
VISITS_PATH = '../in/tables/VISITS_2025.csv'
KPIS_PATH   = 'kpis_grouped_2026.pkl'
SPEC_PATH   = 'export_specialiste.pkl'
OD_PATH     = '../vypocet_ir_2026/zdroje/report_od_pobocky_dbs_04_2026.xlsx'
OUTPUT_FILE = 'report_navstevnost.html'

# ─── Color palette (shades of blue) ────────────────────────────────────────────
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

_OD_DAYS = [
    ('PONDELI','PO','Pondělí',False), ('UTERY','UT','Úterý',False),
    ('STREDA','ST','Středa',False),   ('CTVRTEK','CT','Čtvrtek',False),
    ('PATEK','PA','Pátek',False),     ('SOBOTA','SO','Sobota',True),
    ('NEDELE','NE','Neděle',True),
]

# ─── Capacity constants ─────────────────────────────────────────────────────────
WORKING_DAYS        = 252
WORK_MINS_DAY       = 450   # 7.5 h
TARGET_PORTFOLIO    = 1500  # clients per banker
TARGET_MTGS_MIN     = 4
TARGET_MTGS_MAX     = 5
MEETING_MINS        = 45
WALKIN_SHORT_PCT    = 0.80; WALKIN_SHORT_MINS = 15
WALKIN_LONG_PCT     = 0.20; WALKIN_LONG_MINS  = 30
WALKIN_CONVERT_PCT  = 0.10  # 10 % walkin → schůzka 45 min
WALKIN_AVG_MINS     = WALKIN_SHORT_PCT * WALKIN_SHORT_MINS + WALKIN_LONG_PCT * WALKIN_LONG_MINS  # 18

# Staff positions (normalized uppercase key)
BANKER_COLS = {'OSOBNI_BANKER_-_JUNIOR', 'OSOBNI_BANKER_-_MEDIOR', 'OSOBNI_BANKER_-_SENIOR'}
SERVICE_COL   = 'BANKER_KLIENTSKE_PECE_-_MEDIOR'
SVC_FALLBACK  = 'OSOBNI_BANKER_-_JUNIOR'   # fallback when no BKP-medior
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
    if 'online' in v:        return 'online'
    if 'bezhot' in v:        return 'bezhot'
    if 'hotov' in v:         return 'hotovost'
    if any(x in v for x in ('schuzk','fyzick','schu')): return 'fyzicka'
    return None

def _odv(row, key): return str(row.get(key,'') or '').strip()
def _od_closed(v): return v in ('00:00','0:00','','nan','None','0')


# ─── Capacity calculation ───────────────────────────────────────────────────────

def _cap_model(online, fyzicka, bezhot, bankers, service_fte, n_days):
    """Return capacity metrics dict for a given visit scenario."""
    if bankers <= 0:
        return None

    avail_ob  = bankers   * n_days * WORK_MINS_DAY
    avail_svc = service_fte * n_days * WORK_MINS_DAY

    # Time used
    online_mins   = online   * MEETING_MINS
    fyzicka_mins  = fyzicka  * MEETING_MINS
    bezhot_base   = bezhot   * WALKIN_AVG_MINS
    bezhot_conv   = bezhot   * WALKIN_CONVERT_PCT * MEETING_MINS  # 10 % → schůzka

    # Allocation
    if service_fte > 0:
        # OB handles: online + fyzicka + converted walkin
        # BKP handles: bezhot base time
        ob_used  = online_mins + fyzicka_mins + bezhot_conv
        svc_used = bezhot_base
    else:
        # OB handles everything (junior handles bezhot in addition)
        ob_used  = online_mins + fyzicka_mins + bezhot_base + bezhot_conv
        svc_used = 0

    util_ob  = ob_used  / avail_ob  * 100 if avail_ob  > 0 else 0
    util_svc = svc_used / avail_svc * 100 if avail_svc > 0 else 0

    return {
        'online_mins':  round(online_mins),
        'fyzicka_mins': round(fyzicka_mins),
        'bezhot_base':  round(bezhot_base),
        'bezhot_conv':  round(bezhot_conv),
        'ob_used':      round(ob_used),
        'avail_ob':     round(avail_ob),
        'util_ob':      round(util_ob, 1),
        'svc_used':     round(svc_used),
        'avail_svc':    round(avail_svc),
        'util_svc':     round(util_svc, 1) if service_fte > 0 else None,
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
    """branch_id → {pocet_klientu, fte, name?}"""
    out = {}
    if not os.path.exists(KPIS_PATH): print(f"⚠️  kpis nenalezen: {KPIS_PATH}"); return out
    try:
        kp = pd.read_pickle(KPIS_PATH)
        kp.columns = [_nc(c) for c in kp.columns]
        id_c  = next((c for c in ['POBOCKA_ID','BRANCH_CODE','ID_POBOCKY'] if c in kp.columns), None)
        nm_c  = next((c for c in ['POBOCKA_NAZEV','BRANCH_NAME','NAZEV'] if c in kp.columns), None)
        fte_c = next((c for c in ['FTE','POCET_BANKERU'] if c in kp.columns), None)
        cli_c = next((c for c in ['POCET_KLIENTU','PRIMARNI_KLIENTI','AKTIVNI_KLIENTI'] if c in kp.columns), None)
        if id_c is None: print("⚠️  kpis: ID sloupec nenalezen"); return out
        for _, row in kp.iterrows():
            bid = pd.to_numeric(row[id_c], errors='coerce')
            if pd.isna(bid): continue
            bid = int(bid)
            out[bid] = {
                'name': str(row[nm_c]) if nm_c else None,
                'fte':  float(row[fte_c]) if fte_c and pd.notna(row[fte_c]) else None,
                'pocet_klientu': int(row[cli_c]) if cli_c and pd.notna(row[cli_c]) else None,
            }
        print(f"   kpis: {len(out)} poboček")
    except Exception as e:
        print(f"⚠️  kpis chyba: {e}")
    return out


def load_specialiste():
    """branch_id → {name, bankers, service_fte, positions}"""
    out = {}
    if not os.path.exists(SPEC_PATH): print(f"⚠️  Specialisté nenalezeni: {SPEC_PATH}"); return out
    try:
        sp = pd.read_pickle(SPEC_PATH)
        # Normalize: main script renames to lowercase, we normalize to UPPER
        sp.columns = [_nc(c) for c in sp.columns]
        bid_c  = next((c for c in ['BRANCH_ID','BRANCH_CODE','ID'] if c in sp.columns), None)
        nm_c   = next((c for c in ['BRANCH_NAME','POBOCKA_NAZEV','NAZEV'] if c in sp.columns), None)
        id_cols = {'BRANCH_ID','BRANCH_CODE','BRANCH_NAME','POBOCKA_NAZEV',
                   'GPS_X','GPS_Y','EVIDENCNI_STAV','ID','NAZEV'}
        pos_cols = [c for c in sp.columns if c not in id_cols]
        # numeric
        for c in pos_cols:
            sp[c] = pd.to_numeric(sp[c], errors='coerce').fillna(0)

        for _, row in sp.iterrows():
            bid = pd.to_numeric(row[bid_c] if bid_c else None, errors='coerce')
            if pd.isna(bid): continue
            bid = int(bid)

            bankers = sum(float(row.get(c, 0) or 0) for c in pos_cols if c in BANKER_COLS)
            svc_fte = float(row.get(SERVICE_COL, 0) or 0)
            has_svc = svc_fte > 0

            positions = {}
            for c in pos_cols:
                v = float(row.get(c, 0) or 0)
                if v > 0:
                    lbl = POSITION_LABELS.get(c, c.replace('_',' ').title())
                    positions[lbl] = v

            out[bid] = {
                'name':        str(row[nm_c]) if nm_c and pd.notna(row.get(nm_c,'')) else None,
                'bankers':     round(bankers, 1),
                'service_fte': round(svc_fte, 1),
                'has_svc':     has_svc,
                'positions':   positions,
            }
        print(f"   Specialisté: {len(out)} poboček, {len(pos_cols)} pozic")
    except Exception as e:
        print(f"⚠️  Specialisté chyba: {e}")
    return out


def load_oteviraci():
    out = {}
    if not os.path.exists(OD_PATH): print(f"⚠️  Ot.doba nenalezena: {OD_PATH}"); return out
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
                days.append({'lbl':dcz,'wknd':wknd,'closed':closed,'dop':dop,'odp':odp,
                             'tot':tot if not closed else ''})
            try: ph = float(str(row.get('PH',0) or 0).replace(',','.'))
            except: ph = 0.0
            out[bid] = {'is_vikend':is_vikend,'ph_tyden':ph,'od_days':days}
        print(f"   Ot.doba: {len(out)} poboček")
    except Exception as e:
        print(f"⚠️  Ot.doba chyba: {e}")
    return out


# ─── Aggregation ───────────────────────────────────────────────────────────────

def build_data(df, kpis, spec, od):
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

    result  = {}
    for bid in sorted(df[bid_c].unique()):
        vb  = df[df[bid_c] == bid]
        k   = kpis.get(bid, {})
        s   = spec.get(bid, {})
        o   = od.get(bid, {})

        # Branch name — specialiste first (most reliable), then kpis, then visits CSV
        name = (s.get('name') or k.get('name') or
                (str(vb[bname_c].iloc[0]) if bname_c and bname_c in vb.columns else None) or
                f"Pobočka {bid}")

        total   = len(vb)
        by_type = {k2: int((vb['_t'] == k2).sum()) for k2 in TYPE_KEYS}
        unknown = int(vb['_t'].isna().sum())

        by_month   = {k2: [0]*12 for k2 in TYPE_KEYS}
        by_weekday = {k2: [0]*7  for k2 in TYPE_KEYS}
        by_hour    = {k2: [0]*24 for k2 in TYPE_KEYS}
        heatmap    = [[0]*24 for _ in range(7)]  # [weekday][hour]

        n_days = 1
        if has_date:
            n_days = max(int(vb['_dt'].dt.date.nunique()), 1)
            for k2 in TYPE_KEYS:
                sub = vb[vb['_t'] == k2]
                if not sub.empty:
                    m  = sub['_mon'].value_counts().reindex(range(1,13), fill_value=0)
                    wd = sub['_wd'].value_counts().reindex(range(7),  fill_value=0)
                    by_month[k2]   = [int(v) for v in m]
                    by_weekday[k2] = [int(v) for v in wd]

        if has_time and has_date:
            for k2 in TYPE_KEYS:
                sub = vb[vb['_t'] == k2]
                if not sub.empty:
                    h  = sub['_hr'].dropna().astype(int)
                    hc = h.value_counts().reindex(range(24), fill_value=0)
                    by_hour[k2] = [round(float(v)/n_days, 1) for v in hc]

            # Heatmap: ALL types combined, wd × hour (totals, not averages)
            valid = vb.dropna(subset=['_wd','_hr'])
            if not valid.empty:
                valid = valid.copy()
                valid['_wd'] = valid['_wd'].astype(int)
                valid['_hr'] = valid['_hr'].astype(int)
                hm = (valid.groupby(['_wd','_hr']).size()
                      .unstack(fill_value=0)
                      .reindex(index=range(7), columns=range(24), fill_value=0))
                heatmap = [[int(hm.loc[wd2, hr]) for hr in range(24)] for wd2 in range(7)]

        # Staff
        bankers   = float(s.get('bankers', 0) or 0)
        svc_fte   = float(s.get('service_fte', 0) or 0)
        has_svc   = bool(s.get('has_svc', False))
        positions = s.get('positions', {})
        poc_kli   = k.get('pocet_klientu') or 0
        fte       = k.get('fte') or s.get('total_fte')

        # Metrics
        portfolio_per_banker = round(poc_kli / bankers, 0) if bankers > 0 and poc_kli > 0 else None
        online_yearly  = by_type['online']
        fyzicka_yearly = by_type['fyzicka']
        total_mtgs     = online_yearly + fyzicka_yearly
        mtgs_pb_day    = round(total_mtgs / bankers / n_days, 1) if bankers > 0 and n_days > 0 else None

        # Capacity models
        cap1 = _cap_model(online_yearly, fyzicka_yearly, by_type['bezhot'],
                          bankers, svc_fte, n_days)
        # Model 2: scale visit counts to pocet_klientu using actual type ratios
        cap2 = None
        total_excl_hot = sum(by_type[k2] for k2 in ['online','fyzicka','bezhot'])
        if poc_kli > 0 and total_excl_hot > 0 and bankers > 0:
            r_onl = by_type['online']  / total_excl_hot
            r_fyz = by_type['fyzicka'] / total_excl_hot
            r_beh = by_type['bezhot']  / total_excl_hot
            cap2 = _cap_model(poc_kli * r_onl, poc_kli * r_fyz, poc_kli * r_beh,
                              bankers, svc_fte, WORKING_DAYS)

        result[str(bid)] = {
            'name':       name,
            'fte':        fte,
            'bankers':    bankers,
            'svc_fte':    svc_fte,
            'has_svc':    has_svc,
            'positions':  positions,
            'pocet_klientu': poc_kli,
            'total':      total,
            'by_type':    by_type,
            'unknown':    unknown,
            'by_month':   by_month,
            'by_weekday': by_weekday,
            'by_hour':    by_hour,
            'heatmap':    heatmap,
            'has_time':   has_time and has_date,
            'n_days':     n_days,
            'portfolio_per_banker': portfolio_per_banker,
            'mtgs_pb_day':  mtgs_pb_day,
            'cap1':       cap1,
            'cap2':       cap2,
            'is_vikend':  o.get('is_vikend', False),
            'ph_tyden':   o.get('ph_tyden',  0.0),
            'od_days':    o.get('od_days',   []),
        }

    order = sorted(result.keys(), key=lambda x: result[x]['name'])
    return result, order, att_c is not None


# ─── HTML ───────────────────────────────────────────────────────────────────────

def render_html(data, order, has_type_col):
    data_js  = json.dumps(data,  ensure_ascii=False, separators=(',',':'))
    order_js = json.dumps(order, ensure_ascii=False)
    types_js = json.dumps(
        [{'key':k,'label':TYPE_LABEL[k],'color':TYPE_COLOR[k]} for k in TYPE_KEYS],
        ensure_ascii=False)
    consts_js = json.dumps({
        'TARGET_PORTFOLIO': TARGET_PORTFOLIO,
        'TARGET_MTGS_MIN':  TARGET_MTGS_MIN,
        'TARGET_MTGS_MAX':  TARGET_MTGS_MAX,
        'MEETING_MINS':     MEETING_MINS,
        'WALKIN_AVG_MINS':  WALKIN_AVG_MINS,
        'WALKIN_CONVERT_PCT': WALKIN_CONVERT_PCT,
        'WORK_MINS_DAY':    WORK_MINS_DAY,
    })

    no_type_warn = ('' if has_type_col else
        '<div class="warn">⚠️ Sloupec ATTENDANCE_TYPE nenalezen — typové grafy nejsou dostupné.</div>')

    return f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Analýza návštěvnosti 2025</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',system-ui,sans-serif;background:#f0f4ff;color:#1e293b;font-size:15px;line-height:1.5;}}
.wrap{{max-width:1100px;margin:0 auto;padding:20px 16px;}}
h1{{font-size:1.45rem;font-weight:800;color:#1e3a8a;margin-bottom:3px;}}
.subtitle{{font-size:.82rem;color:#94a3b8;margin-bottom:20px;}}
.warn{{background:#fef9c3;border:1px solid #fde047;border-radius:8px;padding:10px 14px;
       font-size:.82rem;color:#713f12;margin-bottom:14px;}}
/* Search */
.sw{{position:relative;margin-bottom:20px;}}
.sw input{{width:100%;padding:11px 14px 11px 40px;border:1.5px solid #bfdbfe;border-radius:10px;
           font-size:16px;background:#fff;outline:none;transition:border .15s;
           -webkit-tap-highlight-color:transparent;}}
.sw input:focus{{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.15);}}
.sw .ico{{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#93c5fd;pointer-events:none;}}
.bl{{display:none;position:absolute;top:calc(100% + 4px);left:0;right:0;background:#fff;
     border:1.5px solid #bfdbfe;border-radius:10px;max-height:280px;overflow-y:auto;
     z-index:100;box-shadow:0 8px 32px rgba(30,58,138,.12);-webkit-overflow-scrolling:touch;}}
.bl.open{{display:block;}}
.bi{{padding:9px 14px;cursor:pointer;font-size:.88rem;border-bottom:1px solid #eff6ff;}}
.bi:hover,.bi.sel{{background:#eff6ff;color:#1d4ed8;font-weight:600;}}
/* Layout */
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;}}
.grid3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:14px;}}
@media(max-width:700px){{.grid2,.grid3{{grid-template-columns:1fr 1fr;}} }}
@media(max-width:400px){{.grid2,.grid3{{grid-template-columns:1fr;}} }}
/* Cards */
.card{{background:#fff;border-radius:12px;border:1.5px solid #dbeafe;padding:14px 16px;}}
.card.warn-card{{border-color:#fca5a5;background:#fff5f5;}}
.card.ok-card{{border-color:#bbf7d0;background:#f0fdf4;}}
.cl{{font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#64748b;margin-bottom:3px;}}
.cv{{font-size:1.4rem;font-weight:800;line-height:1.1;color:#1e3a8a;}}
.cs{{font-size:.73rem;color:#94a3b8;margin-top:2px;}}
/* Section */
.sec{{background:#fff;border-radius:12px;border:1.5px solid #dbeafe;padding:16px;margin-bottom:14px;}}
.st{{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#475569;margin-bottom:14px;}}
/* Bar charts */
.bars{{display:flex;gap:4px;align-items:flex-end;}}
.bw{{display:flex;flex-direction:column;align-items:center;gap:2px;flex:1;min-width:0;}}
.bs{{display:flex;flex-direction:column-reverse;width:100%;border-radius:4px 4px 0 0;overflow:hidden;}}
.bseg{{width:100%;flex-shrink:0;}}
.blbl{{font-size:.6rem;color:#94a3b8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
       max-width:100%;text-align:center;}}
.bnum{{font-size:.6rem;color:#475569;font-weight:600;text-align:center;}}
/* Legend */
.leg{{display:flex;flex-wrap:wrap;gap:10px;margin-top:10px;}}
.li{{display:flex;align-items:center;gap:5px;font-size:.75rem;color:#475569;}}
.ld{{width:10px;height:10px;border-radius:3px;flex-shrink:0;}}
/* Heatmap */
.hm-table{{border-collapse:collapse;width:100%;font-size:.65rem;}}
.hm-table td,.hm-table th{{padding:2px 3px;text-align:center;border-radius:3px;}}
.hm-table th{{color:#94a3b8;font-weight:600;font-size:.62rem;}}
/* Capacity bar */
.cap-bar-wrap{{position:relative;height:28px;border-radius:6px;overflow:hidden;
               background:#e0e7ef;margin:8px 0;}}
.cap-bar-fill{{height:100%;border-radius:6px;transition:width .4s ease;
               display:flex;align-items:center;padding:0 8px;
               font-size:.72rem;font-weight:700;color:#fff;white-space:nowrap;overflow:hidden;}}
.cap-zone{{position:absolute;top:0;height:100%;opacity:.18;pointer-events:none;}}
/* OD table */
.od-t{{width:100%;border-collapse:collapse;font-size:.84rem;}}
.od-t td,.od-t th{{padding:5px 10px;}}
.od-t th{{font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.4px;
           color:#94a3b8;border-bottom:1px solid #dbeafe;text-align:left;}}
.od-wknd{{background:#fffbf0!important;}}
/* Positions table */
.pos-t{{width:100%;border-collapse:collapse;font-size:.82rem;}}
.pos-t td{{padding:4px 10px;border-bottom:1px solid #f0f4ff;}}
.pos-t tr:last-child td{{border-bottom:none;}}
.badge{{display:inline-block;border-radius:10px;padding:2px 9px;font-size:.7rem;font-weight:700;}}
.placeholder{{text-align:center;color:#94a3b8;padding:50px 0;font-size:.9rem;}}
</style>
</head>
<body>
<div class="wrap">
  <h1>Analýza návštěvnosti poboček 2025</h1>
  <div class="subtitle">Zdroj: {VISITS_PATH}</div>
  {no_type_warn}

  <div class="sw" id="sw">
    <svg class="ico" width="15" height="15" viewBox="0 0 24 24" fill="none"
         stroke="currentColor" stroke-width="2.5">
      <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
    </svg>
    <input type="text" id="si" placeholder="Vyhledat pobočku…"
           autocomplete="off" autocorrect="off" spellcheck="false">
    <div class="bl" id="bl"></div>
  </div>
  <div id="mc"><div class="placeholder">← Vyberte pobočku výše</div></div>
</div>

<script>
const DATA   = {data_js};
const ORDER  = {order_js};
const TYPES  = {types_js};
const C      = {consts_js};
const WD     = ['Po','Út','St','Čt','Pá','So','Ne'];
const MON    = ['Led','Úno','Bře','Dub','Kvě','Čvn','Čvc','Srp','Zář','Říj','Lis','Pro'];
const HM_HOURS = Array.from({{length:16}},(_,i)=>i+6);  // 6..21

let cur=null;
const si=document.getElementById('si'),bl=document.getElementById('bl'),sw=document.getElementById('sw');

function renderList(q) {{
  q=q.toLowerCase();
  const hits=ORDER.filter(id=>DATA[id].name.toLowerCase().includes(q)||id.includes(q)).slice(0,100);
  bl.innerHTML=hits.map(id=>`<div class="bi${{id===cur?' sel':''}}" data-id="${{id}}">
    ${{DATA[id].name}} <span style="color:#bbb;font-size:.78em">#${{id}}</span></div>`).join('');
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
function fmt1(n){{return typeof n==='number'?n.toLocaleString('cs',{{minimumFractionDigits:n%1?1:0,maximumFractionDigits:1}}):'—';}}
function fmtI(n){{return typeof n==='number'?Math.round(n).toLocaleString('cs'):'—';}}
function fmtN(n){{if(!n)return'0';if(n>=1000)return(n/1000).toFixed(1).replace('.0','')+'k';return String(Math.round(n));}}

function stackedBars(labels, typeArr, maxH) {{
  const totals=labels.map((_,i)=>typeArr.reduce((s,t)=>s+(t.values[i]||0),0));
  const maxVal=Math.max(...totals,1);
  return labels.map((lbl,i)=>{{
    const tot=totals[i];
    const segs=typeArr.filter(t=>(t.values[i]||0)>0).map(t=>{{
      const h=(t.values[i]/maxVal*maxH).toFixed(1);
      return `<div class="bseg" style="height:${{h}}px;background:${{t.color}};"
                   title="${{t.label}}: ${{fmtI(t.values[i])}}"></div>`;
    }}).join('');
    return `<div class="bw">
      <div class="bnum">${{tot>0?fmtN(tot):''}}</div>
      <div class="bs" style="height:${{(tot/maxVal*maxH).toFixed(1)}}px;max-height:${{maxH}}px;">${{segs}}</div>
      <div class="blbl">${{lbl}}</div></div>`;
  }}).join('');
}}

function legend(){{
  return`<div class="leg">${{TYPES.map(t=>`<div class="li">
    <div class="ld" style="background:${{t.color}}"></div>${{t.label}}</div>`).join('')}}</div>`;
}}

// ── Heatmap ───────────────────────────────────────────────────────────────────
function heatmapHtml(d, isVikend) {{
  if(!d.has_time||!d.heatmap) return '';
  const rows=(isVikend?[0,1,2,3,4,5,6]:[0,1,2,3,4]);
  const maxV=Math.max(...rows.flatMap(wd=>HM_HOURS.map(h=>d.heatmap[wd]?.[h]||0)),1);
  const th='<th></th>'+HM_HOURS.map(h=>`<th>${{h}}h</th>`).join('');
  const trs=rows.map(wd=>{{
    const wknd=wd>=5;
    const tds=HM_HOURS.map(h=>{{
      const v=d.heatmap[wd]?.[h]||0;
      const p=v/maxV;
      const r=Math.round(239-(239-29)*p),g=Math.round(246-(246-78)*p),b=Math.round(255-(255-216)*p);
      const clr=v>0?`rgb(${{r}},${{g}},${{b}})`:'#f8faff';
      const txt=v>maxV*.5?'#1e3a8a':'#94a3b8';
      return `<td style="background:${{clr}};color:${{txt}};min-width:22px;"
                   title="${{WD[wd]}} ${{h}}h: ${{fmtI(v)}}">${{v>0?fmtN(v):''}}</td>`;
    }}).join('');
    return `<tr style="${{wknd?'background:#fffbf0':''}}"><td style="font-size:.68rem;font-weight:600;
      color:${{wknd?'#d97706':'#475569'}};white-space:nowrap;padding:2px 6px;">${{WD[wd]}}</td>${{tds}}</tr>`;
  }}).join('');
  return`<div class="sec"><div class="st">🗓️ Heatmapa den × hodina (počty návštěv)</div>
    <div style="overflow-x:auto;">
    <table class="hm-table"><thead><tr>${{th}}</tr></thead><tbody>${{trs}}</tbody></table>
    </div></div>`;
}}

// ── Capacity bar ──────────────────────────────────────────────────────────────
function capBar(pct, label) {{
  if(pct==null) return '';
  const p=Math.min(pct,150);
  const clr=pct<70?'#2563eb':pct<90?'#f59e0b':'#ef4444';
  const bg=pct<70?'#dbeafe':pct<90?'#fef3c7':'#fee2e2';
  return `<div style="margin-bottom:6px;">
    <div style="font-size:.72rem;color:#64748b;margin-bottom:3px;">${{label}}</div>
    <div class="cap-bar-wrap" style="background:${{bg}};">
      <div class="cap-bar-fill" style="width:${{Math.min(p,100)}}%;background:${{clr}};">
        ${{pct.toFixed(1)}}%</div>
    </div>
    <div style="font-size:.68rem;color:#94a3b8;text-align:right;">${{
      pct<70?'✅ v normě':pct<90?'⚠️ blíží se limitě':'🔴 přetíženo'}}</div>
  </div>`;
}}

function capSection(d) {{
  const c1=d.cap1, c2=d.cap2;
  if(!c1&&!c2) return '';
  let html='<div class="sec"><div class="st">⚡ Kapacitní analýza</div>';

  function capDetail(cap, title, n_days) {{
    if(!cap) return '';
    const ob_h=(cap.ob_used/60).toFixed(1), avail_h=(cap.avail_ob/60).toFixed(1);
    const svc = cap.util_svc!=null
      ? capBar(cap.util_svc,`BKP Medior (servisní návštěvy)`):'';
    return `<div style="margin-bottom:14px;">
      <div style="font-size:.82rem;font-weight:700;color:#1e3a8a;margin-bottom:8px;">${{title}}</div>
      ${{capBar(cap.util_ob,'OB tým (schůzky + online + walkin konverze)')}}
      ${{svc}}
      <div style="font-size:.72rem;color:#64748b;margin-top:6px;display:flex;gap:16px;flex-wrap:wrap;">
        <span>⏱️ OB využito: <b>${{ob_h}}h</b> / ${{avail_h}}h</span>
        <span>📅 Dnů v datech: <b>${{n_days}}</b></span>
        <span>🕐 Online: <b>${{fmtI(cap.online_mins/60)}}h</b></span>
        <span>🤝 Fyzické: <b>${{fmtI(cap.fyzicka_mins/60)}}h</b></span>
        <span>🚶 Walkin základ: <b>${{fmtI(cap.bezhot_base/60)}}h</b>
              + konverze <b>${{fmtI(cap.bezhot_conv/60)}}h</b></span>
      </div></div>`;
  }}
  html+=capDetail(c1,'Model 1 — reálná data (skutečné návštěvy)',d.n_days);
  html+=capDetail(c2,'Model 2 — klientský model ('+fmtI(d.pocet_klientu)+' klientů → přepočet)',240);
  html+='</div>';
  return html;
}}

// ── Staff section ─────────────────────────────────────────────────────────────
function staffSection(d) {{
  if(!d.bankers&&!Object.keys(d.positions||{{}}).length) return '';
  const posBadges=Object.entries(d.positions||{{}}).map(([lbl,cnt])=>
    `<tr><td style="color:#475569;">${{lbl}}</td>
     <td style="font-weight:700;color:#1d4ed8;text-align:right;">${{fmt1(cnt)}}</td></tr>`
  ).join('');
  const svcNote=d.has_svc
    ?`<span class="badge" style="background:#dbeafe;color:#1d4ed8;">BKP Medior (servis)</span>`
    :`<span class="badge" style="background:#fef9c3;color:#854d0e;">OB Junior (servis — fallback)</span>`;
  return `<div class="sec"><div class="st">👤 Personální obsazení</div>
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px;">
      <div class="card" style="flex:0 0 auto;">
        <div class="cl">Bankéři (OB j/m/s)</div>
        <div class="cv" style="color:#1d4ed8;">${{fmt1(d.bankers)}}</div>
        <div class="cs">osobní bankéř junior/medior/senior</div>
      </div>
      <div class="card" style="flex:0 0 auto;">
        <div class="cl">Servisní bankéř</div>
        <div class="cv" style="font-size:1rem;">${{svcNote}}</div>
      </div>
    </div>
    ${{posBadges?`<table class="pos-t"><tbody>${{posBadges}}</tbody></table>`:''}}
  </div>`;
}}

// ── Metrics cards ─────────────────────────────────────────────────────────────
function metricsSection(d) {{
  const por=d.portfolio_per_banker, mtg=d.mtgs_pb_day;
  if(por==null&&mtg==null) return '';

  let porCard='', mtgCard='';
  if(por!=null) {{
    const ok=por<=C.TARGET_PORTFOLIO;
    const diff=Math.abs(por-C.TARGET_PORTFOLIO);
    porCard=`<div class="card ${{ok?'ok-card':'warn-card'}}">
      <div class="cl">Portfolio / bankéř</div>
      <div class="cv" style="color:${{ok?'#15803d':'#b91c1c'}}">${{fmtI(por)}}</div>
      <div class="cs">cíl ≤ ${{C.TARGET_PORTFOLIO}} · ${{ok
        ?'✅ kapacita volná '+fmtI(diff)+' kl.'
        :'⚠️ převyšuje o '+fmtI(diff)+' kl.'}}</div>
    </div>`;
  }}
  if(mtg!=null) {{
    const ok=mtg>=C.TARGET_MTGS_MIN&&mtg<=C.TARGET_MTGS_MAX;
    const under=mtg<C.TARGET_MTGS_MIN;
    mtgCard=`<div class="card ${{ok?'ok-card':''}}">
      <div class="cl">Schůzky / bankéř / den</div>
      <div class="cv" style="color:${{ok?'#15803d':under?'#64748b':'#b91c1c'}}">${{fmt1(mtg)}}</div>
      <div class="cs">cíl ${{C.TARGET_MTGS_MIN}}–${{C.TARGET_MTGS_MAX}} · ${{ok
        ?'✅ v normě':under
        ?'⬇️ pod cílem':'⬆️ nad cílem'}}</div>
    </div>`;
  }}
  return`<div class="grid2">${{porCard}}${{mtgCard}}</div>`;
}}

// ── Opening hours ─────────────────────────────────────────────────────────────
function odSection(d) {{
  if(!d.od_days||!d.od_days.length) return '';
  const vBadge=d.is_vikend?`<span class="badge" style="background:#d97706;color:#fff;margin-bottom:8px;">🌅 Víkendová pobočka</span>`:'';
  const rows=d.od_days.map(day=>day.closed
    ?`<tr class="${{day.wknd?'od-wknd':''}}">
       <td style="color:#cbd5e1;font-weight:600;">${{day.wknd?'🌅 ':''}}${{day.lbl}}</td>
       <td colspan="3" style="color:#cbd5e1;font-style:italic;text-align:center;">Zavřeno</td></tr>`
    :`<tr class="${{day.wknd?'od-wknd':''}}">
       <td style="font-weight:600;color:${{day.wknd?'#d97706':'#334155'}};white-space:nowrap;">${{day.wknd?'🌅 ':''}}${{day.lbl}}</td>
       <td style="color:#475569;text-align:center;">${{day.dop||'—'}}</td>
       <td style="color:#475569;text-align:center;">${{day.odp||'—'}}</td>
       <td style="font-weight:700;color:#2563eb;text-align:center;">${{day.tot?day.tot+'h':''}}</td></tr>`
  ).join('');
  return`<div class="sec"><div class="st">🕐 Otevírací doba</div>
    ${{vBadge}}
    <table class="od-t">
      <thead><tr><th>Den</th><th>Dopoledne</th><th>Odpoledne</th><th>Celkem</th></tr></thead>
      <tbody>${{rows}}</tbody>
    </table>
    ${{d.ph_tyden>0?`<div style="font-size:.78rem;color:#64748b;margin-top:8px;">
      Týdenní ot. hodiny: <b style="color:#2563eb;">${{d.ph_tyden}}h</b></div>`:''}}</div>`;
}}

// ── Main render ───────────────────────────────────────────────────────────────
function render(id) {{
  cur=id; const d=DATA[id]; const total=d.total;
  const wdLen=d.is_vikend?7:5;
  const wdLabels=WD.slice(0,wdLen);

  // KPI row
  const totCard=`<div class="card"><div class="cl">Celkem návštěv</div>
    <div class="cv">${{fmtI(total)}}</div><div class="cs">rok 2025 · ${{d.n_days}} dnů</div></div>`;
  const fteCard=d.fte!=null?`<div class="card"><div class="cl">FTE celkem</div>
    <div class="cv">${{fmt1(d.fte)}}</div><div class="cs">přepočtené úvazky</div></div>`:'';
  const typeCards=TYPES.map(t=>{{
    const v=d.by_type[t.key]||0, pct=total>0?(v/total*100).toFixed(1):'0.0';
    return`<div class="card" style="border-color:${{t.color}}50;background:${{t.color}}0a;">
      <div class="cl" style="color:${{t.color}};">${{t.label}}</div>
      <div class="cv" style="color:${{t.color}};">${{fmtI(v)}}</div>
      <div class="cs">${{pct}} %</div></div>`;
  }}).join('');

  // Charts
  const tArr=TYPES.map(t=>({{label:t.label,color:t.color,values:d.by_month[t.key]||Array(12).fill(0)}}));
  const wdArr=TYPES.map(t=>({{label:t.label,color:t.color,
    values:(d.by_weekday[t.key]||Array(7).fill(0)).slice(0,wdLen)}}));

  let hrSec='';
  if(d.has_time){{
    const hrs=HM_HOURS;
    const hArr=TYPES.map(t=>({{label:t.label,color:t.color,
      values:hrs.map(h=>d.by_hour[t.key]?.[h]||0)}}));
    hrSec=`<div class="sec"><div class="st">⏱️ Průměrný den — návštěv/hod (6–21h)</div>
      <div class="bars" style="height:100px;">${{stackedBars(hrs.map(h=>h+'h'),hArr,90)}}</div>
      ${{legend()}}</div>`;
  }}

  const unkNote=d.unknown>0
    ?`<div style="font-size:.72rem;color:#94a3b8;margin:4px 0 10px;">
       ${{fmtI(d.unknown)}} návštěv bez určeného typu</div>`:'';
  const wdNote=!d.is_vikend
    ?`<div style="font-size:.7rem;color:#cbd5e1;margin-top:5px;">So/Ne skryty — nevíkendová pobočka</div>`:'';

  document.getElementById('mc').innerHTML=`
<div style="font-size:1.05rem;font-weight:700;color:#1e3a8a;margin-bottom:12px;">
  ${{d.name}} <span style="font-size:.78rem;font-weight:400;color:#94a3b8;">#${{id}}</span>
</div>
<div class="grid3">${{totCard}}${{fteCard}}${{typeCards}}</div>
${{unkNote}}
${{metricsSection(d)}}
${{staffSection(d)}}
${{capSection(d)}}
<div class="sec"><div class="st">📅 Návštěvy dle měsíce</div>
  <div class="bars" style="height:120px;">${{stackedBars(MON,tArr,110)}}</div>
  ${{legend()}}</div>
<div class="sec"><div class="st">📆 Návštěvy dle dne v týdnu</div>
  <div class="bars" style="height:110px;">${{stackedBars(wdLabels,wdArr,100)}}</div>
  ${{legend()}}${{wdNote}}</div>
${{hrSec}}
${{heatmapHtml(d,d.is_vikend)}}
${{odSection(d)}}
`; }}
</script>
</body>
</html>"""


# ─── Generování ────────────────────────────────────────────────────────────────

_df_visits = load_visits()
_kpis      = load_kpis()
_spec      = load_specialiste()
_od        = load_oteviraci()

_visit_data, _order, _has_type = build_data(_df_visits, _kpis, _spec, _od)
_html = render_html(_visit_data, _order, _has_type)

with open(OUTPUT_FILE, 'w', encoding='utf-8') as _f:
    _f.write(_html)

_n_with_names = sum(1 for v in _visit_data.values() if not v['name'].startswith('Pobočka'))
print(f"\n✅ Report uložen: {OUTPUT_FILE}")
print(f"   {len(_visit_data)} poboček · {_n_with_names} s názvem · "
      f"{'typy ✓' if _has_type else 'typy ✗'} · "
      f"spec={'✓' if _spec else '✗'} · od={'✓' if _od else '✗'}")
