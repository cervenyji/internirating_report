"""
generate_network_analysis.py
Analýza optimalizace sítě poboček — tabulky, scénář 1, interaktivní model, dopady

Spuštění:
  Standalone: df = pd.read_pickle("...pkl") na začátku souboru
  Z hlavního skriptu: from generate_network_analysis import generate_network_analysis_report
"""

import json, math, warnings
import numpy as np
import pandas as pd

# ── Konstanty ─────────────────────────────────────────────────────────────────
BANKER_CAPACITY    = 1_500
AVAIL_N_NEAREST    = 5
MAX_METRO_FLAGSHIP = 3
METRO_CITIES       = {'Praha', 'Brno'}
CLIENT_CHURN_RATE  = 0.05

MAPBOX_TOKEN = (
    'pk.eyJ1IjoiY2VydmVueWppIiwiYSI6ImNqM2tsbTl6ajA'
    'wazMyd3FzeGZxa2VxZzcifQ.z-Ruzxhj76p-f84Ti4r3Gw'
)


# ── Pomocné funkce ────────────────────────────────────────────────────────────

def _city_base(city):
    import re
    if not isinstance(city, str):
        return ''
    c = city.strip()
    if re.match(r'^Praha[\s\-–]?\d*', c, re.I) or c.lower() == 'praha':
        return 'Praha'
    if re.match(r'^Brno[\s\-–]', c, re.I) or c.lower() == 'brno':
        return 'Brno'
    return c


def _haversine(lat1, lon1, lat2, lon2):
    R = 6_371_000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl  = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    a = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2 * R * math.asin(math.sqrt(max(0.0, min(1.0, a))))


def _geo_circle(lat, lon, radius_m=10_000, n_pts=28):
    R = 6_371_000
    pts = []
    for i in range(n_pts + 1):
        a    = 2 * math.pi * i / n_pts
        dlat = math.degrees(radius_m / R * math.cos(a))
        dlon = math.degrees(radius_m / R * math.sin(a) / max(1e-9, math.cos(math.radians(lat))))
        pts.append([lon + dlon, lat + dlat])
    return {'type': 'Feature',
            'geometry': {'type': 'Polygon', 'coordinates': [pts]},
            'properties': {}}


# ── Příprava dat ──────────────────────────────────────────────────────────────

def _prep_df(rating_status):
    df = rating_status.copy()
    if 'IR_FLAG' in df.columns:
        df = df[df['IR_FLAG'].eq('Y')].copy()
    if 'BRANCH_CLOSED' in df.columns:
        df = df[~df['BRANCH_CLOSED'].eq(True)].copy()
    df = df.reset_index(drop=True)

    for xc, yc in [('GPS_X', 'GPS_Y'), ('GPS_Y', 'GPS_X')]:
        if xc in df.columns and yc in df.columns:
            mx = pd.to_numeric(df[xc], errors='coerce').median()
            if 48 <= mx <= 52:
                df['_lat'] = pd.to_numeric(df[xc], errors='coerce')
                df['_lon'] = pd.to_numeric(df[yc], errors='coerce')
            else:
                df['_lat'] = pd.to_numeric(df[yc], errors='coerce')
                df['_lon'] = pd.to_numeric(df[xc], errors='coerce')
            break
    else:
        df['_lat'] = np.nan
        df['_lon'] = np.nan

    df['_city'] = df.get('CITY', pd.Series('', index=df.index)).apply(_city_base)

    _nums = {
        'PRIMARNI_KLIENTI': 0.0, 'AKTIVNI_KLIENTI': 0.0,
        'POCET_SCHUZEK_FYZICKY': 0.0, 'CELK_PLOCHA_POBOCKY_2026': 0.0,
        'BANKERS_COUNT': 0.0, 'OBJEM_VYNOSU_CZK': 0.0, 'VYNOSY': 0.0,
        'PRIME_NAKLADY/VYNOSY': np.nan, 'ROCNI_SPLATKY_S_DPH_CZK': 0.0,
        'IR_Q': np.nan,
    }
    for col, dflt in _nums.items():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(dflt)
        else:
            df[col] = dflt

    if 'BRANCH_NAME' not in df.columns:
        df['BRANCH_NAME'] = df.get('BRANCH_CODE', pd.Series('', index=df.index)).astype(str)
    if 'BRANCH_FORMAT' not in df.columns:
        df['BRANCH_FORMAT'] = 'unknown'
    df['BRANCH_FORMAT'] = df['BRANCH_FORMAT'].fillna('unknown').str.lower().str.strip()

    reg = next((c for c in ['REGION_FIXED', 'REGION_NAME', 'REGION'] if c in df.columns), None)
    df['_region'] = df[reg].fillna('—') if reg else '—'

    return df


# ── Výpočty metrik ────────────────────────────────────────────────────────────

def compute_network_availability(df):
    valid  = df.dropna(subset=['_lat', '_lon'])
    result = pd.Series(np.nan, index=df.index)
    lats, lons, idxs = valid['_lat'].values, valid['_lon'].values, valid.index.values
    for i in range(len(idxs)):
        dists = sorted(
            _haversine(lats[i], lons[i], lats[j], lons[j])
            for j in range(len(idxs)) if i != j
        )
        if dists:
            result[idxs[i]] = float(np.mean(dists[:AVAIL_N_NEAREST]))
    return result


def compute_capacity_utilization(df):
    bc = df['BANKERS_COUNT'].clip(lower=0.01)
    return (df['PRIMARNI_KLIENTI'] / (bc * BANKER_CAPACITY)).clip(upper=5.0)


# ── Scénář 1 ─────────────────────────────────────────────────────────────────

def apply_scenario_1(df):
    df = df.copy()
    df['sc1_keep'] = False
    _fmt  = {'flagship': 0, 'medium': 1, 'medium economy': 2, 'small': 3}
    df['_fmt_rank'] = df['BRANCH_FORMAT'].map(_fmt).fillna(9).astype(int)
    df['_iq']       = df['IR_Q'].fillna(3.0)
    for city, grp in df.groupby('_city'):
        if grp.empty:
            continue
        if city in METRO_CITIES:
            flg      = grp[grp['BRANCH_FORMAT'] == 'flagship'].sort_values('_iq')
            keep_idx = flg.index[:MAX_METRO_FLAGSHIP]
            df.loc[keep_idx, 'sc1_keep'] = True
        else:
            best = grp.sort_values(['_fmt_rank', '_iq']).index[0]
            df.loc[best, 'sc1_keep'] = True
    return df


# ── HTML pomocníci ────────────────────────────────────────────────────────────

def _fv(v, fmt):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return '—'
    if fmt == 'int':   return f'{int(round(v)):,}'.replace(',', ' ')
    if fmt == 'f1':    return f'{v:.1f}'
    if fmt == 'f2':    return f'{v:.2f}'
    if fmt == 'pct':   return f'{v:.1f} %'
    if fmt == 'mczk':  return f'{v/1e6:.1f} M'
    if fmt == 'km':    return f'{v:.1f} km'
    if fmt == 'mkczk': return f'{v/1e6:.0f} M Kč'
    if fmt == 'm2':    return f'{v:.0f} m²'
    return str(v)


def _sym_row(label, before, after, fmt, good='neutral'):
    try:
        b, a = float(before), float(after)
        pct   = (a - b) / abs(b) * 100 if b != 0 else 0.0
        pct_s = f'{abs(pct):.1f} %'
        delta = a - b
        if abs(pct) < 0.3:
            sym, col = '→', '#64748b'
        elif delta > 0:
            sym = '▲'
            col = '#16a34a' if good == 'up' else ('#dc2626' if good == 'down' else '#64748b')
        else:
            sym = '▼'
            col = '#16a34a' if good == 'down' else ('#dc2626' if good == 'up' else '#64748b')
        bv = _fv(b, fmt); av = _fv(a, fmt); dv = _fv(delta, fmt) if delta != 0 else '0'
    except Exception:
        sym, col, pct_s, bv, av, dv = '—', '#64748b', '—', '—', '—', '—'
    return (
        f'<tr>'
        f'<td style="padding:5px 10px;font-size:0.81rem;font-weight:600;">{label}</td>'
        f'<td style="padding:5px 10px;text-align:right;font-size:0.81rem;color:#64748b;">{bv}</td>'
        f'<td style="padding:5px 10px;text-align:right;font-size:0.81rem;font-weight:700;color:{col};">{av}</td>'
        f'<td style="padding:5px 10px;text-align:right;font-size:0.81rem;color:{col};">{dv}</td>'
        f'<td style="padding:5px 10px;text-align:center;font-size:0.9rem;font-weight:800;color:{col};">{sym} {pct_s}</td>'
        f'</tr>'
    )


def _calc_row(label, value, note='', col='#1e2a38'):
    note_html = (f'<div style="font-size:0.71rem;color:#94a3b8;">{note}</div>'
                 if note else '')
    return (
        f'<div style="display:flex;align-items:baseline;gap:10px;padding:7px 0;'
        f'border-bottom:1px solid #f1f5f9;">'
        f'<div style="flex:1;font-size:0.81rem;color:#374151;">{label}</div>'
        f'<div style="font-size:1.0rem;font-weight:800;color:{col};white-space:nowrap;">{value}</div>'
        f'{note_html}'
        f'</div>'
    )


# ── Hlavní funkce ─────────────────────────────────────────────────────────────

def generate_network_analysis_report(
    rating_status,
    output_path='report_network_analysis.html',
):
    warnings.filterwarnings('ignore')

    print('  📊 Příprava dat...')
    df = _prep_df(rating_status)
    n_total = len(df)
    print(f'     {n_total} poboček v perimetru')

    print('  📍 Výpočet dostupnosti a kapacity...')
    df['network_availability'] = compute_network_availability(df)
    df['avail_km']             = df['network_availability'] / 1000
    df['capacity_utilization'] = compute_capacity_utilization(df)

    print('  🔀 Scénář 1...')
    df_sc1  = apply_scenario_1(df)
    kdf     = df_sc1[df_sc1['sc1_keep']].copy()
    cdf     = df_sc1[~df_sc1['sc1_keep']].copy()
    n_keep  = len(kdf)
    n_close = len(cdf)

    print('  📍 Dostupnost po Scénáři 1...')
    kdf_avail          = compute_network_availability(kdf)
    avg_avail_after_km = float(kdf_avail.dropna().mean() / 1000) if kdf_avail.notna().any() else np.nan

    # ── Souhrné statistiky ─────────────────────────────────────────────────────
    total_cli    = float(df['PRIMARNI_KLIENTI'].sum())
    total_aktiv  = float(df['AKTIVNI_KLIENTI'].sum())
    total_ban    = float(df['BANKERS_COUNT'].sum())
    total_rev    = float(df['VYNOSY'].sum())
    total_rent   = float(df['ROCNI_SPLATKY_S_DPH_CZK'].sum())
    avg_avail_km = float(df['avail_km'].dropna().mean())
    avg_ci       = float(df['PRIME_NAKLADY/VYNOSY'].dropna().mean())
    _plocha_base = df[df['CELK_PLOCHA_POBOCKY_2026'] > 0]['CELK_PLOCHA_POBOCKY_2026']
    avg_plocha   = float(_plocha_base.mean()) if not _plocha_base.empty else 0.0

    k_cli    = float(kdf['PRIMARNI_KLIENTI'].sum())
    k_aktiv  = float(kdf['AKTIVNI_KLIENTI'].sum())
    k_ban    = float(kdf['BANKERS_COUNT'].sum())
    k_rev    = float(kdf['VYNOSY'].sum())
    k_rent   = float(kdf['ROCNI_SPLATKY_S_DPH_CZK'].sum())
    k_ci     = float(kdf['PRIME_NAKLADY/VYNOSY'].dropna().mean())
    _plocha_k = kdf[kdf['CELK_PLOCHA_POBOCKY_2026'] > 0]['CELK_PLOCHA_POBOCKY_2026']
    k_plocha  = float(_plocha_k.mean()) if not _plocha_k.empty else 0.0

    c_cli    = float(cdf['PRIMARNI_KLIENTI'].sum())
    c_ban    = float(cdf['BANKERS_COUNT'].sum())
    c_rent   = float(cdf['ROCNI_SPLATKY_S_DPH_CZK'].sum())

    rev_per_cli     = total_rev / total_cli if total_cli > 0 else 0.0
    churn_cli       = c_cli * CLIENT_CHURN_RATE
    churn_rev       = churn_cli * rev_per_cli
    cli_per_ban_now = total_cli / total_ban if total_ban > 0 else 0.0
    cli_per_ban_all = total_cli / k_ban     if k_ban     > 0 else 0.0
    cli_per_ban_own = k_cli     / k_ban     if k_ban     > 0 else 0.0
    rent_rev_pct    = total_rent / total_rev * 100 if total_rev > 0 else 0.0

    # ── Korelace ───────────────────────────────────────────────────────────────
    _corr_map = {
        'Primární klienti': 'PRIMARNI_KLIENTI',  'Aktivní klienti': 'AKTIVNI_KLIENTI',
        'Fyzické schůzky':  'POCET_SCHUZEK_FYZICKY', 'Plocha (m²)': 'CELK_PLOCHA_POBOCKY_2026',
        'Bankéři':          'BANKERS_COUNT',      'Nové výnosy':   'OBJEM_VYNOSU_CZK',
        'Výnosy celkem':    'VYNOSY',             'C/I ratio':     'PRIME_NAKLADY/VYNOSY',
        'Nájemné':          'ROCNI_SPLATKY_S_DPH_CZK', 'IR kvintil': 'IR_Q',
        'Dostupnost (km)':  'avail_km',           'Kapacita (%)':  'capacity_utilization',
    }
    avail_cm = {k: v for k, v in _corr_map.items() if v in df.columns}
    corr_sub = df[[v for v in avail_cm.values()]].rename(columns={v: k for k, v in avail_cm.items()})
    corr_sub = corr_sub.apply(pd.to_numeric, errors='coerce')
    corr_mat = corr_sub.corr()
    names    = corr_mat.columns.tolist()

    pairs = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            v = corr_mat.iloc[i, j]
            if not np.isnan(v) and abs(v) > 0.2:
                pairs.append((abs(v), v, names[i], names[j]))
    pairs.sort(key=lambda x: -x[0])

    def _corr_item(absv, v, a, b):
        col   = '#dc2626' if v > 0 else '#2563eb'
        sign  = '+' if v > 0 else ''
        strng = 'silná' if absv > 0.65 else ('střední' if absv > 0.4 else 'slabá')
        bw    = f'{absv*100:.0f}%'
        bg    = col + '18'
        return (
            f'<div style="display:flex;align-items:center;gap:9px;padding:4px 0;'
            f'border-bottom:1px solid #f8fafc;">'
            f'<div style="min-width:46px;text-align:center;padding:2px 5px;'
            f'background:{bg};border-radius:5px;font-size:0.8rem;'
            f'font-weight:800;color:{col};font-variant-numeric:tabular-nums;">'
            f'{sign}{v:.2f}</div>'
            f'<div style="flex:1;min-width:0;">'
            f'<div style="font-size:0.77rem;color:#1e2a38;font-weight:500;'
            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'
            f'{a}<span style="color:#cbd5e1;margin:0 4px;">↔</span>{b}</div>'
            f'<div style="background:#f1f5f9;border-radius:2px;height:3px;margin-top:3px;">'
            f'<div style="width:{bw};background:{col};height:100%;'
            f'border-radius:2px;opacity:.45;"></div></div>'
            f'</div>'
            f'<span style="font-size:0.62rem;color:#94a3b8;white-space:nowrap;'
            f'padding:1px 5px;background:#f8fafc;border-radius:3px;">{strng}</span>'
            f'</div>'
        )

    top_pairs  = pairs[:12]
    mid        = (len(top_pairs) + 1) // 2
    corr_left  = ''.join(_corr_item(*p) for p in top_pairs[:mid])
    corr_right = ''.join(_corr_item(*p) for p in top_pairs[mid:])

    # ── Tabulka poboček ────────────────────────────────────────────────────────
    _tbl_cols = [
        ('Pobočka',             'BRANCH_NAME',              'text', 'left'),
        ('Region',              '_region',                  'text', 'left'),
        ('Formát',              'BRANCH_FORMAT',            'text', 'left'),
        ('IR Q',                'IR_Q',                     'int',  'center'),
        ('Prim. klienti',       'PRIMARNI_KLIENTI',         'int',  'right'),
        ('Akt. klienti',        'AKTIVNI_KLIENTI',          'int',  'right'),
        ('Schůzky',             'POCET_SCHUZEK_FYZICKY',    'int',  'right'),
        ('Bankéři',             'BANKERS_COUNT',            'f1',   'right'),
        ('Výnosy (M Kč)',       'VYNOSY',                   'mczk', 'right'),
        ('Nové výn. (M Kč)',    'OBJEM_VYNOSU_CZK',         'mczk', 'right'),
        ('C/I (%)',             'PRIME_NAKLADY/VYNOSY',     'pct',  'right'),
        ('Plocha (m²)',         'CELK_PLOCHA_POBOCKY_2026', 'm2',   'right'),
        ('Nájemné (M Kč/rok)', 'ROCNI_SPLATKY_S_DPH_CZK', 'mczk', 'right'),
        ('Kapacita (%)',        'capacity_utilization',     'cap',  'right'),
        ('Dostupnost (km)',     'avail_km',                 'km',   'right'),
    ]

    df_sorted = df.sort_values('VYNOSY', ascending=False).reset_index(drop=True)

    def _ir_color(v):
        try:
            return {1:'#16a34a',2:'#65a30d',3:'#f59e0b',4:'#ea580c',5:'#dc2626'}.get(int(float(v)),'#64748b')
        except Exception:
            return '#64748b'

    def _cap_color(v):
        try:
            f = float(v)
            return '#dc2626' if f > 1.2 else ('#f59e0b' if f > 0.9 else '#16a34a')
        except Exception:
            return '#64748b'

    def _cell(v, fmt, align):
        s = f'padding:5px 9px;text-align:{align};font-size:0.79rem;'
        if fmt == 'text':
            return f'<td style="{s}">{v}</td>'
        if fmt == 'int':
            sv = f'{int(round(float(v))):,}'.replace(',', ' ') if pd.notna(v) else '—'
            return f'<td style="{s}">{sv}</td>'
        if fmt == 'mczk':
            sv = f'{float(v)/1e6:.1f}' if pd.notna(v) and float(v) != 0 else '—'
            return f'<td style="{s}">{sv}</td>'
        if fmt == 'pct':
            sv = f'{float(v):.1f}%' if pd.notna(v) else '—'
            return f'<td style="{s}">{sv}</td>'
        if fmt == 'f1':
            sv = f'{float(v):.1f}' if pd.notna(v) else '—'
            return f'<td style="{s}">{sv}</td>'
        if fmt == 'm2':
            sv = f'{float(v):.0f}' if pd.notna(v) and float(v) > 0 else '—'
            return f'<td style="{s}">{sv}</td>'
        if fmt == 'km':
            sv = f'{float(v):.1f}' if pd.notna(v) else '—'
            return f'<td style="{s}">{sv}</td>'
        if fmt == 'cap':
            try:
                fv = float(v); c = _cap_color(fv)
                return (f'<td style="{s}font-weight:700;color:{c};">{fv:.0%}</td>')
            except Exception:
                return f'<td style="{s}">—</td>'
        return f'<td style="{s}">{v}</td>'

    thead_cells = ''.join(
        f'<th style="padding:6px 9px;font-size:0.69rem;font-weight:700;color:#64748b;'
        f'background:#f8fafc;border-bottom:2px solid #e2e8f0;white-space:nowrap;'
        f'text-align:{al};">{lbl}</th>'
        for lbl, _, _, al in _tbl_cols
    )

    def _build_rows(df_part):
        rows = ''
        for _, row in df_part.iterrows():
            cells = ''
            for lbl, col_name, fmt, align in _tbl_cols:
                v = row.get(col_name, '—')
                if col_name == 'IR_Q':
                    try:
                        iv = int(float(v)); c = _ir_color(iv)
                        cells += (f'<td style="padding:5px 9px;text-align:center;'
                                  f'font-size:0.79rem;font-weight:700;color:{c};">{iv}</td>')
                    except Exception:
                        cells += '<td style="padding:5px 9px;text-align:center;">—</td>'
                elif col_name == 'BRANCH_FORMAT':
                    fmt_colors = {'flagship':'#2563eb','medium':'#16a34a',
                                  'medium economy':'#65a30d','small':'#64748b'}
                    c = fmt_colors.get(str(v).lower(), '#94a3b8')
                    cells += (f'<td style="padding:5px 9px;font-size:0.74rem;'
                              f'font-weight:600;color:{c};">{str(v).capitalize()}</td>')
                else:
                    cells += _cell(v, fmt, align)
            rows += f'<tr style="border-bottom:1px solid #f1f5f9;">{cells}</tr>\n'
        return rows

    preview_rows = _build_rows(df_sorted.head(5))
    rest_rows    = _build_rows(df_sorted.iloc[5:])
    n_rest       = len(df_sorted) - 5

    # ── Impact tabulka ─────────────────────────────────────────────────────────
    impact_html = (
        _sym_row('Poboček',                      n_total,      n_keep,             'int',  'neutral') +
        _sym_row('Primárních klientů',            total_cli,    k_cli,              'int',  'up') +
        _sym_row('Aktivních klientů',             total_aktiv,  k_aktiv,            'int',  'up') +
        _sym_row('Výnosy celkem (M Kč)',          total_rev,    k_rev,              'mczk', 'up') +
        _sym_row('Bankéřů celkem',                total_ban,    k_ban,              'int',  'neutral') +
        _sym_row('Průměrná plocha pobočky (m²)',  avg_plocha,   k_plocha,           'f1',   'up') +
        _sym_row('Průměrná dostupnost sítě (km)', avg_avail_km, avg_avail_after_km, 'km',   'down') +
        _sym_row('Nájemné celkem (M Kč/rok)',     total_rent,   k_rent,             'mczk', 'down') +
        _sym_row('C/I ratio průměr (%)',          avg_ci,       k_ci,               'pct',  'down')
    )

    calcs_html = (
        _calc_row('Neobsloužených klientů (primárních)',
                  _fv(c_cli, 'int'), 'z uzavřených poboček', '#dc2626') +
        _calc_row('Klientů na bankéře — zachovaní bankéři, jen vlastní klienti',
                  _fv(cli_per_ban_own, 'f1'), f'dnes {cli_per_ban_now:.0f} / bankéř', '#2563eb') +
        _calc_row('Klientů na bankéře — zachovaní bankéři, všichni klienti sítě',
                  _fv(cli_per_ban_all, 'f1'), f'(+{c_cli:,.0f} redistributováno)', '#7c3aed') +
        _calc_row('Bankéřů dotčených uzavřením',
                  _fv(c_ban, 'f1'), f'= {c_ban/total_ban*100:.1f} % celku', '#dc2626') +
        _calc_row(f'Dopad do výnosů (odchodovost {CLIENT_CHURN_RATE*100:.0f} % klientů uzavřených poboček)',
                  _fv(churn_rev, 'mkczk') + '/rok',
                  f'{churn_cli:,.0f} klientů × {rev_per_cli:,.0f} Kč/klient', '#dc2626') +
        _calc_row('Úspora nájmů uzavřených poboček',
                  _fv(c_rent, 'mkczk') + '/rok', '', '#16a34a') +
        _calc_row('Průměrná plocha poboček — před / po',
                  f'{avg_plocha:.0f} → {k_plocha:.0f} m²',
                  f'změna {k_plocha - avg_plocha:+.0f} m²', '#02A3A4')
    )

    close_rows_html = ''
    for _, row in cdf.sort_values(['_region', 'PRIMARNI_KLIENTI'], ascending=[True, False]).iterrows():
        iq = f'{int(float(row["IR_Q"]))}' if pd.notna(row['IR_Q']) else '—'
        close_rows_html += (
            f'<tr style="border-bottom:1px solid #fef2f2;">'
            f'<td style="padding:5px 9px;font-size:0.78rem;font-weight:600;">{row["BRANCH_NAME"]}</td>'
            f'<td style="padding:5px 9px;font-size:0.73rem;color:#64748b;">{row["_city"]}</td>'
            f'<td style="padding:5px 9px;font-size:0.73rem;">{str(row["BRANCH_FORMAT"]).capitalize()}</td>'
            f'<td style="padding:5px 9px;text-align:center;font-size:0.79rem;'
            f'font-weight:700;color:{_ir_color(row["IR_Q"])};">{iq}</td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.79rem;">'
            f'{int(row["PRIMARNI_KLIENTI"]):,}</td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.79rem;">'
            f'{row["ROCNI_SPLATKY_S_DPH_CZK"]/1e6:.1f} M</td>'
            f'</tr>\n'
        )

    # ── Klientský pohled ──────────────────────────────────────────────────────
    print('  👥 Analýza klientského pohledu...')

    BRANCH_DEP_CHURN = 0.25   # odhadovaná odchodovost pobočkových klientů při uzavření

    # Proxy: fyzické schůzky / primární klienti = míra závislosti na pobočce (0..1)
    _meet = df['POCET_SCHUZEK_FYZICKY'] / df['PRIMARNI_KLIENTI'].clip(lower=1)
    df['_meet_ratio'] = _meet.clip(0.0, 1.0).fillna(0.0)

    cdf_c = cdf.copy()
    cdf_c['_meet_ratio'] = df.loc[cdf.index, '_meet_ratio'].values

    visits_loaded = False
    try:
        _vdf = pd.read_csv('../in/tables/VISITS_2025.csv', low_memory=False)
        _vdf.columns = [c.strip().upper().replace(' ', '_') for c in _vdf.columns]
        _bid_c = next((c for c in ['BRANCH_ID', 'BRANCH_CODE', 'ID_POBOCKY', 'POBOCKA'] if c in _vdf.columns), None)
        _cid_c = next((c for c in ['PT_UNIFIED_KEY', 'CLIENT_ID', 'KLIENT_ID'] if c in _vdf.columns), None)
        if _bid_c and _cid_c:
            _vis_cnt = _vdf.groupby(_bid_c)[_cid_c].nunique().to_dict()
            _bc_col = next((c for c in ['BRANCH_CODE', 'BRANCH_ID'] if c in cdf_c.columns), None)
            if _bc_col:
                cdf_c['_visitors'] = cdf_c[_bc_col].map(_vis_cnt).fillna(0)
                cdf_c['_meet_ratio'] = (cdf_c['_visitors'] / cdf_c['PRIMARNI_KLIENTI'].clip(lower=1)).clip(0.0, 1.0)
                visits_loaded = True
                print(f'     Visits CSV: {len(_vdf):,} řádků, reálná data návštěv')
    except Exception:
        pass

    cdf_c['_est_branch_dep'] = (cdf_c['PRIMARNI_KLIENTI'] * cdf_c['_meet_ratio']).round().astype(int).clip(lower=0)
    cdf_c['_est_digital']    = (cdf_c['PRIMARNI_KLIENTI'] - cdf_c['_est_branch_dep']).clip(lower=0)

    tot_cli_closed   = int(cdf_c['PRIMARNI_KLIENTI'].sum())
    tot_branch_dep   = int(cdf_c['_est_branch_dep'].sum())
    tot_digital_c    = int(cdf_c['_est_digital'].sum())
    tot_churn_branch = int(round(tot_branch_dep * BRANCH_DEP_CHURN))
    tot_churn_dig    = int(round(tot_digital_c * CLIENT_CHURN_RATE))
    tot_churn_total  = tot_churn_branch + tot_churn_dig
    churn_rev_branch = tot_churn_branch * rev_per_cli
    churn_rev_dig    = tot_churn_dig    * rev_per_cli
    churn_rev_total  = churn_rev_branch + churn_rev_dig

    # Věkové skupiny — pokud jsou dostupné ve vstupních datech
    _AGE_GROUPS = ['1-15', '16-25', '26-45', '46-65', '65+']
    _AGE_COLORS = {'1-15':'#2770ef','16-25':'#0bb440','26-45':'#00a3a5','46-65':'#fd6230','65+':'#9b59b6'}
    _age_avail  = [g for g in _AGE_GROUPS if f'{g}_POCET_KLIENTU' in cdf_c.columns]
    age_bars_html = ''
    if _age_avail:
        _age_tots  = {g: int(cdf_c[f'{g}_POCET_KLIENTU'].sum()) for g in _age_avail}
        _age_total = sum(_age_tots.values()) or 1
        for g, cnt in _age_tots.items():
            pct = cnt / _age_total * 100
            c   = _AGE_COLORS.get(g, '#64748b')
            age_bars_html += (
                f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">'
                f'<div style="min-width:42px;font-size:0.75rem;font-weight:700;color:{c};">{g}</div>'
                f'<div style="flex:1;background:#f1f5f9;border-radius:3px;height:7px;overflow:hidden;">'
                f'<div style="width:{pct:.1f}%;background:{c};height:100%;border-radius:3px;"></div></div>'
                f'<div style="min-width:82px;text-align:right;font-size:0.74rem;color:#64748b;">'
                f'{cnt:,} &nbsp;({pct:.0f}%)</div>'
                f'</div>'
            )

    # Stacked bar (pobočkoví vs digitální) na úrovni sítě
    _bd_pct_net = tot_branch_dep / max(tot_cli_closed, 1) * 100
    _dg_pct_net = 100.0 - _bd_pct_net

    # Tabulka: top uzavřené pobočky dle počtu klientů
    _cli_rows_html = ''
    for _, row in cdf_c.sort_values('PRIMARNI_KLIENTI', ascending=False).head(12).iterrows():
        pc  = int(row['PRIMARNI_KLIENTI'])
        bd  = int(row['_est_branch_dep'])
        dg  = int(row['_est_digital'])
        mr  = float(row['_meet_ratio'])
        bd_w = max(1, int(mr * 70))
        dg_w = max(1, 70 - bd_w)
        est_churn = int(round(bd * BRANCH_DEP_CHURN)) + int(round(dg * CLIENT_CHURN_RATE))
        _cli_rows_html += (
            f'<tr style="border-bottom:1px solid #f1f5f9;">'
            f'<td style="padding:5px 9px;font-size:0.78rem;font-weight:600;">{row["BRANCH_NAME"]}</td>'
            f'<td style="padding:5px 9px;font-size:0.73rem;color:#64748b;">{row["_city"]}</td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.79rem;">{pc:,}</td>'
            f'<td style="padding:5px 9px;">'
            f'<div style="display:flex;gap:2px;align-items:center;">'
            f'<div style="width:{bd_w}px;background:#dc2626;height:7px;border-radius:2px 0 0 2px;opacity:.65;" title="Pobočkoví: {bd:,}"></div>'
            f'<div style="width:{dg_w}px;background:#2563eb;height:7px;border-radius:0 2px 2px 0;opacity:.35;" title="Digitální: {dg:,}"></div>'
            f'</div></td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.79rem;color:#dc2626;font-weight:700;">{bd:,}</td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.79rem;color:#2563eb;">{dg:,}</td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.79rem;color:#ea580c;font-weight:700;">{est_churn:,}</td>'
            f'</tr>\n'
        )

    _cli_source_note = ('Počet unikátních návštěvníků z VISITS_2025.csv'
                        if visits_loaded else
                        'Proxy: fyzické schůzky ÷ primární klienti (VISITS_2025.csv není dostupný)')

    # ── JSON dat pro mapu ──────────────────────────────────────────────────────
    print('  🗂️  Serializace dat pro JS...')
    lat_c = float(df_sc1['_lat'].dropna().mean()) if df_sc1['_lat'].notna().any() else 49.8
    lon_c = float(df_sc1['_lon'].dropna().mean()) if df_sc1['_lon'].notna().any() else 15.5

    branches_js = []
    for idx, row in df_sc1.iterrows():
        lat = row.get('_lat')
        lon = row.get('_lon')
        has_gps = pd.notna(lat) and pd.notna(lon)
        circ_coords = []
        if has_gps:
            circ_coords = _geo_circle(float(lat), float(lon))['geometry']['coordinates']
        raw_ci = row.get('PRIME_NAKLADY/VYNOSY')
        branches_js.append({
            'id':          int(idx),
            'name':        str(row.get('BRANCH_NAME', '—')),
            'city':        str(row.get('_city', '—')),
            'format':      str(row.get('BRANCH_FORMAT', '—')),
            'ir_q':        int(float(row['IR_Q'])) if pd.notna(row.get('IR_Q')) else 0,
            'clients':     int(row.get('PRIMARNI_KLIENTI', 0)),
            'revenue':     float(row.get('VYNOSY', 0)),
            'rent':        float(row.get('ROCNI_SPLATKY_S_DPH_CZK', 0)),
            'bankers':     float(row.get('BANKERS_COUNT', 0)),
            'ci':          round(float(raw_ci), 3) if pd.notna(raw_ci) else None,
            'avail_km':    round(float(row['avail_km']), 3) if pd.notna(row.get('avail_km')) else 0.0,
            'cap_pct':     round(float(row['capacity_utilization']), 4) if pd.notna(row.get('capacity_utilization')) else 0.0,
            'sc1_keep':    bool(row.get('sc1_keep', False)),
            'lat':         round(float(lat), 6) if has_gps else None,
            'lon':         round(float(lon), 6) if has_gps else None,
            'circle_coords': circ_coords,
        })

    branches_json_str = json.dumps(branches_js, ensure_ascii=False, separators=(',', ':'))

    # ── JS data (Python values → JS constants, plain string concatenation) ────
    js_data = (
        'const MAPBOX_TOKEN=' + json.dumps(MAPBOX_TOKEN) + ';\n'
        'const BRANCHES='     + branches_json_str + ';\n'
        'const LAT_C='        + str(round(lat_c, 4)) + ';\n'
        'const LON_C='        + str(round(lon_c, 4)) + ';\n'
        'const MAX_METRO_FLAGSHIP=' + str(MAX_METRO_FLAGSHIP) + ';\n'
        'const METRO_CITIES_SET=new Set(' + json.dumps(sorted(METRO_CITIES)) + ');\n'
        'const BANKER_CAPACITY=' + str(BANKER_CAPACITY) + ';\n'
        'const CLIENT_CHURN_RATE=' + str(CLIENT_CHURN_RATE) + ';\n'
        'const AVAIL_N_NEAREST=' + str(AVAIL_N_NEAREST) + ';\n'
        'const BASE_CLI_PER_BAN=' + str(round(cli_per_ban_now, 2)) + ';\n'
        'const BASE_RENT_REV_PCT=' + str(round(rent_rev_pct, 2)) + ';\n'
        'const BASE_AVG_CI=' + str(round(avg_ci, 2)) + ';\n'
        'const BASE_AVAIL_KM=' + str(round(avg_avail_km, 3)) + ';\n'
    )

    # ── JS logic (raw string — no f-string escaping) ──────────────────────────
    js_logic = r"""
// ── Utils ──────────────────────────────────────────────────────────────────
function fmtInt(v){return v==null?'—':Math.round(v).toLocaleString('cs-CZ');}
function fmtMczk(v){return v==null?'—':(v/1e6).toFixed(1)+' M';}
function fmtKm(v){return v==null?'—':v.toFixed(1)+' km';}
function fmtPct(v){return v==null?'—':v.toFixed(1)+' %';}
function fmtF1(v){return v==null?'—':v.toFixed(1);}

// ── GeoJSON builders ────────────────────────────────────────────────────────
function buildPoints(stObj){
  return {type:'FeatureCollection',features:BRANCHES.filter(b=>b.lat&&b.lon).map(b=>({
    type:'Feature',geometry:{type:'Point',coordinates:[b.lon,b.lat]},
    properties:{id:b.id,name:b.name,city:b.city,format:b.format,
                ir_q:b.ir_q,clients:b.clients,revenue:b.revenue,
                state:stObj[b.id]||'close'}
  }))};
}
function buildCircles(stObj){
  return {type:'FeatureCollection',features:BRANCHES.filter(b=>b.lat&&b.lon&&b.circle_coords.length).map(b=>({
    type:'Feature',geometry:{type:'Polygon',coordinates:b.circle_coords},
    properties:{id:b.id,state:stObj[b.id]||'close'}
  }))};
}

// ── Map data updater — přímé volání bez closure ────────────────────────────
// Použijeme globální reference sc1Map / modelMap; žádné closure potřeba.
function updateMapData(map,stObj){
  if(!map)return;
  // Pokud mapa ještě není načtená, zaregistrujeme jednorázový pokus po 'idle'
  try{
    const ps=map.getSource('pts'),cs=map.getSource('cir');
    if(ps&&cs){ps.setData(buildPoints(stObj));cs.setData(buildCircles(stObj));return;}
  }catch(e){}
  // Map source ještě neexistuje — počkáme na 'idle' (nastane po load)
  const retry=()=>{
    try{
      const ps=map.getSource('pts'),cs=map.getSource('cir');
      if(ps)ps.setData(buildPoints(stObj));
      if(cs)cs.setData(buildCircles(stObj));
    }catch(e2){}
  };
  map.once('idle',retry);
}

// ── Stats chips ──────────────────────────────────────────────────────────────
function calcStats(stObj){
  const keep=BRANCHES.filter(b=>stObj[b.id]==='keep');
  const close=BRANCHES.filter(b=>stObj[b.id]!=='keep');
  const totCli=BRANCHES.reduce((s,b)=>s+b.clients,0);
  const totRev=BRANCHES.reduce((s,b)=>s+b.revenue,0);
  const kCli=keep.reduce((s,b)=>s+b.clients,0);
  const kRev=keep.reduce((s,b)=>s+b.revenue,0);
  const kRent=keep.reduce((s,b)=>s+b.rent,0);
  const cCli=close.reduce((s,b)=>s+b.clients,0);
  const cRent=close.reduce((s,b)=>s+b.rent,0);
  const rpc=totCli>0?totRev/totCli:0;
  return {nKeep:keep.length,nClose:close.length,kCli,kRev,kRent,cCli,cRent,
    churnRev:cCli*CLIENT_CHURN_RATE*rpc,totCli,totRev};
}

function renderStats(stObj,pfx){
  const s=calcStats(stObj);
  const el=id=>document.getElementById(pfx+'-'+id);
  if(!el('nk'))return;
  el('nk').textContent=s.nKeep;
  el('nc').textContent=s.nClose;
  el('kcli').textContent=fmtInt(s.kCli);
  el('krev').textContent=fmtMczk(s.kRev);
  el('crent').textContent=fmtMczk(s.cRent);
  el('churn').textContent=fmtMczk(s.churnRev);
}

// ── Layer toggle (Sc1 map) ───────────────────────────────────────────────────
function toggleSc1Layer(layerIds,btn){
  if(!sc1Map)return;
  const on=btn.classList.contains('lyr-on');
  const vis=on?'none':'visible';
  layerIds.forEach(id=>{try{sc1Map.setLayoutProperty(id,'visibility',vis);}catch(e){}});
  btn.classList.toggle('lyr-on',!on);
  btn.classList.toggle('lyr-off',on);
}

// ── Map factory ──────────────────────────────────────────────────────────────
function createMap(containerId,initState,onToggle){
  mapboxgl.accessToken=MAPBOX_TOKEN;
  const map=new mapboxgl.Map({
    container:containerId,
    style:'mapbox://styles/mapbox/light-v11',
    center:[LON_C,LAT_C],zoom:6.4,
    attributionControl:false
  });
  map.addControl(new mapboxgl.NavigationControl({showCompass:false}),'top-right');
  map.addControl(new mapboxgl.AttributionControl({compact:true}));

  map.on('load',()=>{
    // Snapshot stavu v momentě načtení (může být novější než initState)
    const curState=containerId==='sc1-map'?sc1State:modelState;
    map.addSource('pts',{type:'geojson',data:buildPoints(curState)});
    map.addSource('cir',{type:'geojson',data:buildCircles(curState)});

    map.addLayer({id:'cir-fill-c',type:'fill',source:'cir',
      filter:['==',['get','state'],'close'],
      paint:{'fill-color':'#dc2626','fill-opacity':0.07}});
    map.addLayer({id:'cir-line-c',type:'line',source:'cir',
      filter:['==',['get','state'],'close'],
      paint:{'line-color':'#dc2626','line-opacity':0.28,'line-width':1}});
    map.addLayer({id:'cir-fill-k',type:'fill',source:'cir',
      filter:['==',['get','state'],'keep'],
      paint:{'fill-color':'#16a34a','fill-opacity':0.09}});
    map.addLayer({id:'cir-line-k',type:'line',source:'cir',
      filter:['==',['get','state'],'keep'],
      paint:{'line-color':'#16a34a','line-opacity':0.40,'line-width':1.5}});
    map.addLayer({id:'pts-c',type:'circle',source:'pts',
      filter:['==',['get','state'],'close'],
      paint:{'circle-radius':7,'circle-color':'#dc2626',
             'circle-stroke-width':1.5,'circle-stroke-color':'#fff'}});
    map.addLayer({id:'pts-k',type:'circle',source:'pts',
      filter:['==',['get','state'],'keep'],
      paint:{'circle-radius':9,'circle-color':'#16a34a',
             'circle-stroke-width':1.5,'circle-stroke-color':'#fff'}});

    const popup=new mapboxgl.Popup({closeButton:false,closeOnClick:false,offset:12});
    ['pts-k','pts-c'].forEach(lyr=>{
      map.on('mouseenter',lyr,e=>{
        map.getCanvas().style.cursor='pointer';
        const p=e.features[0].properties;
        // Stav čteme přímo z globálního objektu — vždy aktuální
        const liveState=containerId==='sc1-map'?sc1State:modelState;
        const cur=liveState[p.id]||'close';
        const hint=onToggle?('<br><em style="color:#94a3b8;font-size:0.75rem">Klikni: '+(cur==='keep'?'uzavřít':'zachovat')+'</em>'):'';
        popup.setLngLat(e.lngLat).setHTML(
          '<div style="font-family:system-ui;font-size:0.82rem;line-height:1.5;max-width:210px;">'+
          '<strong>'+p.name+'</strong><br>'+p.city+' · '+p.format+'<br>'+
          'IR Q '+p.ir_q+' · '+fmtInt(p.clients)+' klientů'+hint+'</div>'
        ).addTo(map);
      });
      map.on('mouseleave',lyr,()=>{map.getCanvas().style.cursor='';popup.remove();});
      if(onToggle){
        map.on('click',lyr,e=>{
          const id=e.features[0].properties.id;
          onToggle(map,id);
        });
      }
    });
  });
  return map;
}

// ── Haversine (km) ────────────────────────────────────────────────────────────
function haversineKm(lat1,lon1,lat2,lon2){
  const R=6371,p=Math.PI/180;
  const dLat=(lat2-lat1)*p,dLon=(lon2-lon1)*p;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*p)*Math.cos(lat2*p)*Math.sin(dLon/2)**2;
  return 2*R*Math.asin(Math.min(1,Math.sqrt(Math.max(0,a))));
}

function computeAvailKm(bs){
  const vld=bs.filter(b=>b.lat&&b.lon);
  if(vld.length<2)return null;
  const n=Math.min(AVAIL_N_NEAREST,vld.length-1);
  let tot=0;
  vld.forEach((b,i)=>{
    const dists=vld.filter((_,j)=>j!==i)
      .map(b2=>haversineKm(b.lat,b.lon,b2.lat,b2.lon))
      .sort((a,b)=>a-b).slice(0,n);
    tot+=dists.reduce((s,d)=>s+d,0)/(dists.length||1);
  });
  return tot/vld.length;
}

// ── Insight card ──────────────────────────────────────────────────────────────
function calcInsights(stObj){
  const keep=BRANCHES.filter(b=>stObj[b.id]==='keep');
  const kBan=keep.reduce((s,b)=>s+b.bankers,0);
  const totCli=BRANCHES.reduce((s,b)=>s+b.clients,0);
  const kRev=keep.reduce((s,b)=>s+b.revenue,0);
  const kRent=keep.reduce((s,b)=>s+b.rent,0);
  const cliPerBan=kBan>0?totCli/kBan:null;
  const rentRevPct=kRev>0?kRent/kRev*100:null;
  let ciN=0,ciD=0;
  keep.forEach(b=>{if(b.ci!=null&&b.revenue>0){ciN+=b.ci*b.revenue;ciD+=b.revenue;}});
  const avgCI=ciD>0?ciN/ciD:null;
  const availKm=computeAvailKm(keep);
  return {cliPerBan,rentRevPct,avgCI,availKm};
}

function insightRow(label,context,base,s1val,mdlval,fmtFn,goodDir){
  function cmpCell(val){
    if(val==null||base==null)return '<td class="ic-val ic-na">—</td>';
    const diff=val-base;
    const pct=Math.abs(base)>0.001?Math.abs(diff/base)*100:0;
    const better=goodDir==='down'?diff<-0.3:diff>0.3;
    const worse =goodDir==='down'?diff>0.3:diff<-0.3;
    const col=pct<0.3?'#64748b':(better?'#16a34a':(worse?'#dc2626':'#64748b'));
    const sym=pct<0.3?'→':(diff>0?'▲':'▼');
    return '<td class="ic-val" style="color:'+col+';font-weight:700;">'+fmtFn(val)+
           ' <span class="ic-sym">'+sym+'</span></td>';
  }
  return '<tr class="ic-row">'+
    '<td class="ic-lbl">'+label+'<div class="ic-ctx">'+context+'</div></td>'+
    '<td class="ic-base">'+fmtFn(base)+'</td>'+
    cmpCell(s1val)+
    cmpCell(mdlval)+
    '</tr>';
}

function _renderInsightNow(){
  const tbody=document.getElementById('insight-body');
  if(!tbody)return;
  const s1=calcInsights(sc1State);
  const md=calcInsights(modelState);
  tbody.innerHTML=
    insightRow('Klientů na bankéře',
      'Všichni klienti sítě ÷ zbývající bankéři (cílový limit: '+BANKER_CAPACITY+')',
      BASE_CLI_PER_BAN,s1.cliPerBan,md.cliPerBan,fmtF1,'down')+
    insightRow('Nájemné / výnosy',
      'Podíl ročních nájmů zachovaných poboček na jejich celkových výnosech',
      BASE_RENT_REV_PCT,s1.rentRevPct,md.rentRevPct,fmtPct,'down')+
    insightRow('C/I ratio průměr',
      'Výnosově vážený průměr nákladové efektivity (prime náklady / výnosy)',
      BASE_AVG_CI,s1.avgCI,md.avgCI,fmtPct,'down')+
    insightRow('Dostupnost sítě (km)',
      'Průměrná vzdálenost pobočky k '+AVAIL_N_NEAREST+' nejbližším — výpočet v prohlížeči',
      BASE_AVAIL_KM,s1.availKm,md.availKm,fmtKm,'down');
}

let _insightDebounce=null;
function renderInsightCard(){
  clearTimeout(_insightDebounce);
  _insightDebounce=setTimeout(_renderInsightNow,120);
}

// ── Scénář 1 ─────────────────────────────────────────────────────────────────
const sc1State={};
BRANCHES.forEach(b=>{sc1State[b.id]=b.sc1_keep?'keep':'close';});
let sc1Map=null;

// ── Interaktivní model ────────────────────────────────────────────────────────
let modelScored=[];
let modelState={};
let modelMap=null;

const maxAvail=Math.max(...BRANCHES.map(b=>b.avail_km||0))||1;
const maxRev  =Math.max(...BRANCHES.map(b=>b.revenue  ||0))||1;
const maxCli  =Math.max(...BRANCHES.map(b=>b.clients  ||0))||1;

function scoreAndBuild(wA,wR,wC){
  const tot=wA+wR+wC||1;
  const wa=wA/tot,wr=wR/tot,wc=wC/tot;
  modelScored=BRANCHES.map(b=>({
    ...b,
    score:wa*(b.avail_km/maxAvail)+wr*(b.revenue/maxRev)+wc*(b.clients/maxCli)
  }));
  const state={};
  BRANCHES.forEach(b=>{state[b.id]='close';});
  const byCity={};
  modelScored.forEach(b=>{if(!byCity[b.city])byCity[b.city]=[];byCity[b.city].push(b);});
  Object.entries(byCity).forEach(([city,bs])=>{
    const sorted=[...bs].sort((a,b)=>b.score-a.score);
    if(METRO_CITIES_SET.has(city)){
      sorted.slice(0,MAX_METRO_FLAGSHIP).forEach(b=>{state[b.id]='keep';});
    }else{
      if(sorted.length>0)state[sorted[0].id]='keep';
    }
  });
  return state;
}

function renderModelList(stObj){
  const tbody=document.getElementById('ml-body');
  if(!tbody)return;
  const keeps=modelScored.filter(b=>stObj[b.id]==='keep').sort((a,b)=>b.score-a.score);
  tbody.innerHTML=keeps.map(b=>(
    '<tr style="border-bottom:1px solid #f1f5f9;">'+
    '<td style="padding:4px 9px;font-size:0.79rem;font-weight:600;">'+b.name+'</td>'+
    '<td style="padding:4px 9px;font-size:0.73rem;color:#64748b;">'+b.city+'</td>'+
    '<td style="padding:4px 9px;font-size:0.73rem;">'+b.format+'</td>'+
    '<td style="padding:4px 9px;text-align:right;font-size:0.79rem;">'+fmtInt(b.clients)+'</td>'+
    '<td style="padding:4px 9px;text-align:right;font-size:0.79rem;">'+fmtMczk(b.revenue)+'</td>'+
    '<td style="padding:4px 9px;text-align:right;font-size:0.79rem;">'+fmtKm(b.avail_km)+'</td>'+
    '<td style="padding:4px 9px;text-align:right;font-size:0.77rem;color:#2563eb;font-weight:700;">'+
    (b.score*100).toFixed(1)+'</td>'+
    '</tr>'
  )).join('');
}

// ── Aplikace modelu (slider input) ───────────────────────────────────────────
function applyModel(){
  const wA=+document.getElementById('w-avail').value||33;
  const wR=+document.getElementById('w-rev').value||33;
  const wC=+document.getElementById('w-cli').value||33;
  document.getElementById('lbl-avail').textContent=Math.round(wA);
  document.getElementById('lbl-rev').textContent=Math.round(wR);
  document.getElementById('lbl-cli').textContent=Math.round(wC);
  modelState=scoreAndBuild(wA,wR,wC);
  updateMapData(modelMap,modelState);
  renderStats(modelState,'model');
  renderModelList(modelState);
  renderInsightCard();
}

// ── Propsat hodnoty (tlačítko — okamžitá aktualizace) ────────────────────────
function propsatHodnoty(){
  const wA=+document.getElementById('w-avail').value||33;
  const wR=+document.getElementById('w-rev').value||33;
  const wC=+document.getElementById('w-cli').value||33;
  modelState=scoreAndBuild(wA,wR,wC);
  updateMapData(modelMap,modelState);
  renderStats(modelState,'model');
  renderModelList(modelState);
  clearTimeout(_insightDebounce);
  _renderInsightNow();
  const btn=document.getElementById('propsat-btn');
  if(btn){btn.textContent='✓ Propsat';btn.style.background='#16a34a';
    setTimeout(()=>{btn.textContent='🔄 Propsat hodnoty';btn.style.background='#7c3aed';},1400);}
}

// ── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded',()=>{
  sc1Map=createMap('sc1-map',sc1State,(map,id)=>{
    sc1State[id]=sc1State[id]==='keep'?'close':'keep';
    updateMapData(sc1Map,sc1State);
    renderStats(sc1State,'sc1');
    renderInsightCard();
  });
  renderStats(sc1State,'sc1');

  modelState=scoreAndBuild(33,33,33);
  modelMap=createMap('model-map',modelState,null);
  renderStats(modelState,'model');
  renderModelList(modelState);
  setTimeout(()=>_renderInsightNow(),200);

  ['w-avail','w-rev','w-cli'].forEach(id=>{
    document.getElementById(id).addEventListener('input',applyModel);
  });
});
"""

    js_code = js_data + js_logic

    # ── CSS — ČS brand guidelines (Bright Blue accent) ────────────────────────
    CSS = (
        ':root{'
        '--cs-blue:#2870ED;--cs-teal:#02A3A4;--cs-forest:#028661;--cs-apple:#0CB43F;'
        '--cs-orange:#FF6130;--cs-pink:#EB4C79;--cs-aubergine:#721C7A;--cs-stone:#245375;'
        '--cs-accent:#2870ED;'  # Bright Blue — zvolená barva
        '--cs-anthracite:#202020;--cs-gray-dark:#4A4A4A;--cs-gray:#9B9B9B;'
        '--cs-gray-light:#E6E6E6;--cs-bg:#F4F6FA;'
        '}\n'
        '*{box-sizing:border-box;margin:0;padding:0;}\n'
        'body{font-family:"Inter",Arial,"Helvetica Neue",sans-serif;'
        'background:var(--cs-bg);color:var(--cs-anthracite);line-height:1.5;font-size:14px;}\n'
        # Hero
        '.hdr{background:var(--cs-accent);color:#fff;padding:40px 48px 28px;position:relative;overflow:hidden;}\n'
        '.hdr::before{content:"";position:absolute;top:0;left:0;width:44px;height:3px;'
        'background:rgba(255,255,255,.65);}\n'
        '.hdr::after{content:"";position:absolute;top:0;left:0;width:3px;height:44px;'
        'background:rgba(255,255,255,.65);}\n'
        '.hdr-top{font-size:0.66rem;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;'
        'opacity:.7;margin-bottom:14px;}\n'
        '.hdr h1{font-size:1.95rem;font-weight:800;line-height:1.18;margin-bottom:10px;'
        'text-wrap:balance;max-width:820px;}\n'
        '.hdr-sub{font-size:0.82rem;opacity:.72;margin-bottom:22px;}\n'
        '.hdr-meta{font-size:0.74rem;opacity:.6;border-top:1px solid rgba(255,255,255,.22);'
        'padding-top:12px;display:flex;gap:20px;align-items:center;flex-wrap:wrap;}\n'
        '.hdr-claim{font-size:0.88rem;font-weight:700;opacity:.9;margin-left:auto;letter-spacing:.4px;}\n'
        # Wrap
        '.wrap{max-width:1440px;margin:0 auto;padding:24px 20px 16px;}\n'
        # Cards — each section gets its own border-top color
        '.card{background:#fff;border-radius:10px;padding:22px 24px;'
        'box-shadow:0 1px 6px rgba(0,0,0,.06);margin-bottom:20px;'
        'border-top:3px solid var(--cs-accent);}\n'
        '.card-teal{border-top-color:var(--cs-teal);}\n'
        '.card-forest{border-top-color:var(--cs-forest);}\n'
        '.card-stone{border-top-color:var(--cs-stone);}\n'
        '.card-aubergine{border-top-color:var(--cs-aubergine);}\n'
        '.card-orange{border-top-color:var(--cs-orange);}\n'
        # Section label (CT)
        '.ct{font-size:0.66rem;font-weight:700;text-transform:uppercase;letter-spacing:.9px;'
        'margin-bottom:14px;color:var(--cs-accent);}\n'
        '.ct-teal{color:var(--cs-teal);}\n'
        '.ct-forest{color:var(--cs-forest);}\n'
        '.ct-stone{color:var(--cs-stone);}\n'
        '.ct-aubergine{color:var(--cs-aubergine);}\n'
        '.ct-orange{color:var(--cs-orange);}\n'
        # Layout
        '.two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px;}\n'
        '@media(max-width:860px){.two-col{grid-template-columns:1fr;}}\n'
        '.map-box{border-radius:8px;overflow:hidden;border:1px solid var(--cs-gray-light);margin-bottom:12px;}\n'
        # Tables
        '.impact-tbl{width:100%;border-collapse:collapse;}\n'
        '.impact-tbl th{background:#f6f8fb;padding:6px 10px;font-size:0.67rem;font-weight:700;'
        'color:var(--cs-gray-dark);border-bottom:2px solid var(--cs-gray-light);white-space:nowrap;}\n'
        '.impact-tbl tr:hover td{background:#f6f8fb;}\n'
        # Chips
        '.stats-row{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 16px;}\n'
        '.stat-chip{background:#f6f8fb;border:1px solid var(--cs-gray-light);border-radius:8px;'
        'padding:7px 14px;font-size:0.77rem;color:var(--cs-gray-dark);}\n'
        '.stat-chip b{font-size:0.91rem;color:var(--cs-anthracite);display:block;margin-bottom:2px;}\n'
        '.stat-chip.green b{color:var(--cs-forest);}\n'
        '.stat-chip.red b{color:#C0392B;}\n'
        # Badges
        '.badge{display:inline-block;border-radius:4px;padding:2px 8px;font-size:0.68rem;font-weight:700;}\n'
        '.b-keep{background:#E8F5EF;color:var(--cs-forest);border:1px solid #b3d9c6;}\n'
        '.b-close{background:#FCECEA;color:#C0392B;border:1px solid #f5beba;}\n'
        # Show more
        '.show-more-btn{display:inline-block;margin-top:10px;padding:7px 18px;'
        'background:#f6f8fb;border:1.5px solid var(--cs-gray-light);border-radius:6px;'
        'font-size:0.77rem;font-weight:600;color:var(--cs-accent);cursor:pointer;'
        'transition:background .15s;font-family:inherit;}\n'
        '.show-more-btn:hover{background:#e0eafc;}\n'
        '.scroll-tbl{max-height:320px;overflow-y:auto;}\n'
        # Sliders
        '.slider-row{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-bottom:18px;}\n'
        '@media(max-width:700px){.slider-row{grid-template-columns:1fr;}}\n'
        '.slider-label{font-size:0.78rem;font-weight:600;color:var(--cs-gray-dark);margin-bottom:6px;'
        'display:flex;justify-content:space-between;}\n'
        '.slider-label span{font-size:0.9rem;font-weight:800;color:var(--cs-aubergine);}\n'
        'input[type=range]{width:100%;accent-color:var(--cs-aubergine);height:4px;cursor:pointer;}\n'
        '.sc1-desc{font-size:0.78rem;color:var(--cs-gray-dark);margin-bottom:14px;line-height:1.55;}\n'
        '.corr-grid{display:grid;grid-template-columns:1fr 1fr;gap:0 28px;}\n'
        '@media(max-width:700px){.corr-grid{grid-template-columns:1fr;}}\n'
        # Layer toggles
        '.lyr-toggles{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px;}\n'
        '.lyr-btn{display:flex;align-items:center;gap:5px;padding:5px 13px;cursor:pointer;'
        'border-radius:20px;border:1.5px solid var(--cs-gray-light);background:#fff;font-size:0.74rem;'
        'font-weight:600;transition:all .15s;user-select:none;font-family:inherit;}\n'
        '.lyr-btn.lyr-on{opacity:1;}\n'
        '.lyr-btn.lyr-off{opacity:.4;text-decoration:line-through;background:#f6f8fb;}\n'
        '.lyr-btn.keep-k.lyr-on{border-color:var(--cs-forest);color:var(--cs-forest);}\n'
        '.lyr-btn.close-k.lyr-on{border-color:#C0392B;color:#C0392B;}\n'
        # Insight table
        '.insight-tbl{width:100%;border-collapse:collapse;}\n'
        '.ic-row{border-bottom:1px solid #eef0f4;}\n'
        '.ic-row:hover td{background:#f6f8fb;}\n'
        '.ic-lbl{padding:10px 12px;font-size:0.82rem;font-weight:600;min-width:220px;}\n'
        '.ic-ctx{font-size:0.69rem;color:var(--cs-gray);font-weight:400;margin-top:2px;}\n'
        '.ic-base{padding:10px 12px;text-align:right;font-size:0.85rem;color:var(--cs-gray-dark);white-space:nowrap;}\n'
        '.ic-val{padding:10px 12px;text-align:right;font-size:0.88rem;white-space:nowrap;}\n'
        '.ic-na{color:var(--cs-gray);}\n'
        '.ic-sym{font-size:0.72rem;}\n'
        '.insight-th{padding:8px 12px;font-size:0.67rem;font-weight:700;color:var(--cs-gray-dark);'
        'background:#f6f8fb;border-bottom:2px solid var(--cs-gray-light);text-align:right;white-space:nowrap;}\n'
        '.insight-th.left{text-align:left;}\n'
        # Footer
        '.ftr{background:var(--cs-accent);color:#fff;padding:28px 48px;display:flex;'
        'align-items:center;justify-content:space-between;gap:20px;position:relative;overflow:hidden;}\n'
        '.ftr::before{content:"";position:absolute;top:0;left:0;width:44px;height:3px;'
        'background:rgba(255,255,255,.55);}\n'
        '.ftr::after{content:"";position:absolute;top:0;left:0;width:3px;height:44px;'
        'background:rgba(255,255,255,.55);}\n'
        '.ftr-left{font-size:0.77rem;opacity:.75;line-height:1.7;}\n'
        '.ftr-claim{font-size:1.15rem;font-weight:800;letter-spacing:.5px;}\n'
        '.ftr-hash{font-size:0.82rem;font-weight:600;opacity:.72;margin-top:3px;}\n'
        '@media(max-width:600px){.ftr{flex-direction:column;align-items:flex-start;padding:24px 20px;}}\n'
    )

    # ── Assemble HTML ──────────────────────────────────────────────────────────
    html = (
        '<!DOCTYPE html>\n'
        '<html lang="cs">\n'
        '<head>\n'
        '<meta charset="UTF-8">\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0">\n'
        '<title>Analýza sítě poboček</title>\n'
        "<link rel='preconnect' href='https://fonts.googleapis.com'>\n"
        "<link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>\n"
        "<link href='https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap' rel='stylesheet'>\n"
        "<link href='https://api.mapbox.com/mapbox-gl-js/v2.15.0/mapbox-gl.css' rel='stylesheet'/>\n"
        "<script src='https://api.mapbox.com/mapbox-gl-js/v2.15.0/mapbox-gl.js'></script>\n"
        '<style>\n' + CSS + '\n</style>\n'
        '</head>\n'
        '<body>\n'
        '<div class="hdr">\n'
        '  <div class="hdr-top">Česká spořitelna · Interní analýza</div>\n'
        f'  <h1>Analýza sítě poboček — optimalizace</h1>\n'
        f'  <div class="hdr-sub">{n_total} poboček v perimetru · Dostupnost, kapacita, scénáře</div>\n'
        '  <div class="hdr-meta">\n'
        f'    <span>Network Analysis Report</span>\n'
        '    <span class="hdr-claim">Ať se daří</span>\n'
        '  </div>\n'
        '</div>\n'
        '<div class="wrap">\n'

        # ── Tabulka poboček ────────────────────────────────────────────────────
        '\n<!-- Tabulka poboček -->\n'
        '<div class="card">\n'
        f'  <div class="ct">📊 Přehled poboček — {n_total} poboček (seřazeno dle výnosů)</div>\n'
        '  <div style="overflow-x:auto;">\n'
        '  <table style="width:100%;border-collapse:collapse;min-width:1100px;">\n'
        f'    <thead><tr>{thead_cells}</tr></thead>\n'
        f'    <tbody>{preview_rows}</tbody>\n'
        '  </table>\n'
        '  </div>\n'
        '  <div id="branch-more" style="display:none;overflow-x:auto;margin-top:2px;">\n'
        '  <table style="width:100%;border-collapse:collapse;min-width:1100px;">\n'
        f'    <thead><tr>{thead_cells}</tr></thead>\n'
        f'    <tbody>{rest_rows}</tbody>\n'
        '  </table>\n'
        '  </div>\n'
        f'  <button class="show-more-btn" id="more-btn"\n'
        "          onclick=\"(function(){"
        "var el=document.getElementById('branch-more');"
        "var btn=document.getElementById('more-btn');"
        "var open=el.style.display!=='none';"
        "el.style.display=open?'none':'block';"
        f"btn.textContent=open?'Zobrazit všech {n_rest} dalších poboček ▼':'Skrýt ▲';"
        "})()\">\n"
        f'    Zobrazit všech {n_rest} dalších poboček ▼\n'
        '  </button>\n'
        '</div>\n'

        # ── Korelace ───────────────────────────────────────────────────────────
        '\n<!-- Korelace -->\n'
        '<div class="card card-teal">\n'
        '  <div class="ct ct-teal">🔗 Nejvýznamnější korelace mezi metrikami</div>\n'
        '  <div class="corr-grid">\n'
        f'    <div>{corr_left}</div>\n'
        f'    <div>{corr_right}</div>\n'
        '  </div>\n'
        '  <p style="margin-top:10px;font-size:0.69rem;color:#94a3b8;">'
        'Pearsonův r · páry s |r| &gt; 0.2 · ±1 = perfektní lineární vztah</p>\n'
        '</div>\n'

        # ── Scénář 1 ───────────────────────────────────────────────────────────
        '\n<!-- Scénář 1 -->\n'
        '<div class="card card-forest">\n'
        '  <div style="font-size:1rem;font-weight:700;color:var(--cs-forest);margin-bottom:6px;">'
        '📋 Scénář 1 — nejpřísnější optimalizace</div>\n'
        f'  <div class="sc1-desc">'
        f'<strong>Pravidlo:</strong> Mimo Prahu a Brno zůstane v každém městě <strong>jedna pobočka</strong>'
        f' — nejlepší dle IR kvintilu, přednost flagship formátu.'
        f' V Praze a Brně max {MAX_METRO_FLAGSHIP} flagship.'
        f' <span style="color:#94a3b8;">Kliknutím na bod přepnete stav pobočky.</span>'
        f'</div>\n'

        # stats chips
        '  <div class="stats-row">\n'
        f'    <div class="stat-chip green"><b id="sc1-nk">{n_keep}</b>Zachovat</div>\n'
        f'    <div class="stat-chip red"><b id="sc1-nc">{n_close}</b>Uzavřít</div>\n'
        '    <div class="stat-chip"><b id="sc1-kcli">—</b>Primárních klientů zachováno</div>\n'
        '    <div class="stat-chip"><b id="sc1-krev">—</b>Výnosy zachováno</div>\n'
        '    <div class="stat-chip green"><b id="sc1-crent">—</b>Úspora nájmů</div>\n'
        '    <div class="stat-chip red"><b id="sc1-churn">—</b>Odhad ztráty výnosů (5% odchod)</div>\n'
        '  </div>\n'

        # layer toggle buttons
        '  <div class="lyr-toggles">\n'
        '    <button class="lyr-btn keep-k lyr-on" onclick="toggleSc1Layer([\'pts-k\'],this)">🟢 Zachovat — body</button>\n'
        '    <button class="lyr-btn keep-k lyr-on" onclick="toggleSc1Layer([\'cir-fill-k\',\'cir-line-k\'],this)">🟢 Zachovat — kružnice</button>\n'
        '    <button class="lyr-btn close-k lyr-on" onclick="toggleSc1Layer([\'pts-c\'],this)">🔴 Uzavřít — body</button>\n'
        '    <button class="lyr-btn close-k lyr-on" onclick="toggleSc1Layer([\'cir-fill-c\',\'cir-line-c\'],this)">🔴 Uzavřít — kružnice</button>\n'
        '  </div>\n'

        # map
        '  <div class="map-box" style="height:520px;" id="sc1-map"></div>\n'

        # impact + calcs
        '  <div class="two-col">\n'
        '    <div>\n'
        '      <div class="ct">📉 Dopad na parametry sítě</div>\n'
        '      <div style="overflow-x:auto;">\n'
        '      <table class="impact-tbl">\n'
        '        <thead><tr>\n'
        '          <th style="text-align:left;">Metrika</th>\n'
        '          <th style="text-align:right;">Před</th>\n'
        '          <th style="text-align:right;">Scénář 1</th>\n'
        '          <th style="text-align:right;">Změna</th>\n'
        '          <th style="text-align:center;">Trend</th>\n'
        '        </tr></thead>\n'
        f'        <tbody>{impact_html}</tbody>\n'
        '      </table>\n'
        '      </div>\n'
        '    </div>\n'
        '    <div>\n'
        '      <div class="ct">🔢 Detailní výpočet dopadů</div>\n'
        f'      {calcs_html}\n'
        '    </div>\n'
        '  </div>\n'

        # closed branches
        '  <div style="margin-top:18px;">\n'
        f'    <div class="ct">Pobočky navržené k uzavření ({n_close})</div>\n'
        '    <div class="scroll-tbl">\n'
        '    <table style="width:100%;border-collapse:collapse;">\n'
        '      <thead><tr style="background:#fef2f2;">\n'
        '        <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#b91c1c;'
        'border-bottom:1px solid #fecaca;text-align:left;">Pobočka</th>\n'
        '        <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#b91c1c;'
        'border-bottom:1px solid #fecaca;text-align:left;">Město</th>\n'
        '        <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#b91c1c;'
        'border-bottom:1px solid #fecaca;text-align:left;">Formát</th>\n'
        '        <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#b91c1c;'
        'border-bottom:1px solid #fecaca;text-align:center;">IR Q</th>\n'
        '        <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#b91c1c;'
        'border-bottom:1px solid #fecaca;text-align:right;">Primárních klientů</th>\n'
        '        <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#b91c1c;'
        'border-bottom:1px solid #fecaca;text-align:right;">Nájemné (M Kč/rok)</th>\n'
        '      </tr></thead>\n'
        f'      <tbody>{close_rows_html}</tbody>\n'
        '    </table>\n'
        '    </div>\n'
        '  </div>\n'
        '</div>\n'   # /sc1

        # ── Klíčové dopady ─────────────────────────────────────────────────────
        '\n<!-- Klíčové dopady -->\n'
        '<div class="card card-stone">\n'
        '  <div style="font-size:1rem;font-weight:700;color:var(--cs-stone);margin-bottom:6px;">'
        '💡 Klíčové dopady — živá srovnávací analýza</div>\n'
        '  <div style="font-size:0.77rem;color:var(--cs-gray-dark);margin-bottom:14px;line-height:1.5;">'
        'Hodnoty se aktualizují při každé změně scénáře 1 (klik na pobočku) '
        'i interaktivního modelu (posun slideru). '
        '<span style="color:#94a3b8;">Dostupnost = průměr vzdáleností k 5 nejbližším pobočkám, '
        'počítáno v prohlížeči pomocí haversine.</span>'
        '</div>\n'
        '  <div style="overflow-x:auto;">\n'
        '  <table class="insight-tbl">\n'
        '    <thead><tr>\n'
        '      <th class="insight-th left" style="min-width:260px;">Metrika</th>\n'
        '      <th class="insight-th">Baseline (nyní)</th>\n'
        '      <th class="insight-th" style="color:var(--cs-forest);">📋 Scénář 1</th>\n'
        '      <th class="insight-th" style="color:var(--cs-aubergine);">⚖️ Interaktivní model</th>\n'
        '    </tr></thead>\n'
        '    <tbody id="insight-body">\n'
        '      <tr><td colspan="4" style="padding:20px;text-align:center;color:#94a3b8;'
        'font-size:0.8rem;">Načítám…</td></tr>\n'
        '    </tbody>\n'
        '  </table>\n'
        '  </div>\n'
        f'  <p style="margin-top:10px;font-size:0.69rem;color:#94a3b8;">'
        f'Cílový limit bankéřů: {BANKER_CAPACITY:,} klientů/bankéř · '
        f'Odchodovost: {CLIENT_CHURN_RATE*100:.0f} % · '
        f'Dostupnost dle {AVAIL_N_NEAREST} nejbližších poboček</p>\n'
        '</div>\n'

        # ── Interaktivní model ─────────────────────────────────────────────────
        '\n<!-- Interaktivní model -->\n'
        '<div class="card card-aubergine">\n'
        '  <div style="font-size:1rem;font-weight:700;color:var(--cs-aubergine);margin-bottom:6px;">'
        '⚖️ Interaktivní model — vyvažování priorit</div>\n'
        '  <div class="sc1-desc">'
        'Posun váhy určuje, co má mít přednost při výběru zachované pobočky ve městě.'
        f' Praha/Brno: max {MAX_METRO_FLAGSHIP} flagshipy.'
        '</div>\n'

        # sliders
        '  <div class="slider-row">\n'
        '    <div>\n'
        '      <div class="slider-label">🗺️ Zachování dostupnosti<span id="lbl-avail">33</span></div>\n'
        '      <input type="range" id="w-avail" min="0" max="100" value="33">\n'
        '      <div style="font-size:0.69rem;color:#94a3b8;margin-top:4px;">'
        'Preferuje pobočky s vysokou avail_km (izolované lokality)</div>\n'
        '    </div>\n'
        '    <div>\n'
        '      <div class="slider-label">💰 Zachování výnosů<span id="lbl-rev">33</span></div>\n'
        '      <input type="range" id="w-rev" min="0" max="100" value="33">\n'
        '      <div style="font-size:0.69rem;color:#94a3b8;margin-top:4px;">'
        'Preferuje pobočky s vyššími celkovými výnosy</div>\n'
        '    </div>\n'
        '    <div>\n'
        '      <div class="slider-label">👥 Počet klientů<span id="lbl-cli">33</span></div>\n'
        '      <input type="range" id="w-cli" min="0" max="100" value="33">\n'
        '      <div style="font-size:0.69rem;color:#94a3b8;margin-top:4px;">'
        'Preferuje pobočky s nejvyšším počtem primárních klientů</div>\n'
        '    </div>\n'
        '  </div>\n'
        '  <div style="margin-bottom:14px;">\n'
        '    <button id="propsat-btn" onclick="propsatHodnoty()"\n'
        '      style="padding:9px 22px;background:#7c3aed;color:white;border:none;'
        'border-radius:8px;font-size:0.84rem;font-weight:700;cursor:pointer;'
        'transition:background .2s;">'
        '🔄 Propsat hodnoty</button>\n'
        '    <span style="margin-left:12px;font-size:0.73rem;color:#94a3b8;">'
        'Aplikuje aktuální nastavení sliderů do mapy a srovnávací analýzy</span>\n'
        '  </div>\n'

        # model stats
        '  <div class="stats-row">\n'
        '    <div class="stat-chip green"><b id="model-nk">—</b>Zachovat</div>\n'
        '    <div class="stat-chip red"><b id="model-nc">—</b>Uzavřít</div>\n'
        '    <div class="stat-chip"><b id="model-kcli">—</b>Primárních klientů zachováno</div>\n'
        '    <div class="stat-chip"><b id="model-krev">—</b>Výnosy zachováno</div>\n'
        '    <div class="stat-chip green"><b id="model-crent">—</b>Úspora nájmů</div>\n'
        '    <div class="stat-chip red"><b id="model-churn">—</b>Odhad ztráty výnosů (5% odchod)</div>\n'
        '  </div>\n'

        # model map + list
        '  <div class="two-col" style="align-items:start;">\n'
        '    <div>\n'
        '      <div class="map-box" style="height:440px;" id="model-map"></div>\n'
        '      <div style="font-size:0.72rem;color:#94a3b8;margin-top:4px;">'
        '🟢 Zachovat &nbsp;🔴 Uzavřít &nbsp;Kružnice = 10 km</div>\n'
        '    </div>\n'
        '    <div>\n'
        '      <div class="ct">Zachované pobočky (seřazeno dle skóre)</div>\n'
        '      <div class="scroll-tbl" style="max-height:440px;">\n'
        '      <table style="width:100%;border-collapse:collapse;">\n'
        '        <thead><tr style="background:#f8fafc;">\n'
        '          <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#64748b;'
        'border-bottom:2px solid #e2e8f0;text-align:left;">Pobočka</th>\n'
        '          <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#64748b;'
        'border-bottom:2px solid #e2e8f0;text-align:left;">Město</th>\n'
        '          <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#64748b;'
        'border-bottom:2px solid #e2e8f0;text-align:left;">Formát</th>\n'
        '          <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#64748b;'
        'border-bottom:2px solid #e2e8f0;text-align:right;">Klienti</th>\n'
        '          <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#64748b;'
        'border-bottom:2px solid #e2e8f0;text-align:right;">Výnosy</th>\n'
        '          <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#64748b;'
        'border-bottom:2px solid #e2e8f0;text-align:right;">Avail. km</th>\n'
        '          <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#2563eb;'
        'border-bottom:2px solid #e2e8f0;text-align:right;">Skóre</th>\n'
        '        </tr></thead>\n'
        '        <tbody id="ml-body"></tbody>\n'
        '      </table>\n'
        '      </div>\n'
        '    </div>\n'
        '  </div>\n'
        '</div>\n'   # /model

        # ── Pohled přes klienty ────────────────────────────────────────────────
        '\n<!-- Klientský pohled -->\n'
        '<div class="card card-orange">\n'
        '  <div style="font-size:1rem;font-weight:700;color:var(--cs-orange);margin-bottom:6px;">'
        '👥 Pohled přes klienty — dopad uzavření poboček</div>\n'
        f'  <div style="font-size:0.76rem;color:#64748b;margin-bottom:16px;">'
        f'Analýza {n_close} uzavíraných poboček (Scénář 1) &nbsp;·&nbsp; {_cli_source_note}</div>\n'

        # summary chips row
        '  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:18px;">\n'
        f'    <div class="stat-chip"><b style="color:#ea580c;">{tot_cli_closed:,}</b>Klientů v uzavíraných pobočkách</div>\n'
        f'    <div class="stat-chip red"><b>{tot_branch_dep:,}</b>Pobočkově závislí (riziko odchodu)</div>\n'
        f'    <div class="stat-chip"><b style="color:#2563eb;">{tot_digital_c:,}</b>Digitální klienti (nízké riziko)</div>\n'
        f'    <div class="stat-chip red"><b>{tot_churn_total:,}</b>Odhadovaný odchod celkem</div>\n'
        f'    <div class="stat-chip red"><b>{churn_rev_total/1e6:.0f} M Kč</b>Odhadovaná ztráta výnosů/rok</div>\n'
        '  </div>\n'

        # stacked bar — branch-dep vs digital
        '  <div style="margin-bottom:18px;">\n'
        '    <div style="font-size:0.75rem;font-weight:700;color:#374151;margin-bottom:6px;">'
        'Segmentace klientů uzavíraných poboček</div>\n'
        '    <div style="display:flex;height:20px;border-radius:6px;overflow:hidden;'
        'background:#f1f5f9;margin-bottom:6px;">\n'
        f'      <div style="width:{_bd_pct_net:.1f}%;background:#dc2626;opacity:.7;" '
        f'title="Pobočkoví: {tot_branch_dep:,}"></div>\n'
        f'      <div style="width:{_dg_pct_net:.1f}%;background:#2563eb;opacity:.3;" '
        f'title="Digitální: {tot_digital_c:,}"></div>\n'
        '    </div>\n'
        '    <div style="display:flex;gap:18px;font-size:0.72rem;">\n'
        f'      <span><span style="display:inline-block;width:10px;height:10px;'
        f'background:#dc2626;opacity:.7;border-radius:2px;margin-right:4px;"></span>'
        f'Pobočkoví {_bd_pct_net:.0f}% ({tot_branch_dep:,}) — odchod {BRANCH_DEP_CHURN*100:.0f}%</span>\n'
        f'      <span><span style="display:inline-block;width:10px;height:10px;'
        f'background:#2563eb;opacity:.5;border-radius:2px;margin-right:4px;"></span>'
        f'Digitální {_dg_pct_net:.0f}% ({tot_digital_c:,}) — odchod {CLIENT_CHURN_RATE*100:.0f}%</span>\n'
        '    </div>\n'
        '  </div>\n'

        # two-col: churn breakdown + age groups
        '  <div class="two-col" style="margin-bottom:20px;">\n'
        '    <div>\n'
        '      <div class="ct">📉 Odhadovaný odchod a ztráta výnosů</div>\n'
        + (
            _calc_row('Pobočkoví klienti — odchod',
                      f'{tot_churn_branch:,} ({BRANCH_DEP_CHURN*100:.0f}% ze {tot_branch_dep:,})',
                      f'≈ {churn_rev_branch/1e6:.0f} M Kč/rok', '#dc2626') +
            _calc_row('Digitální klienti — odchod',
                      f'{tot_churn_dig:,} ({CLIENT_CHURN_RATE*100:.0f}% z {tot_digital_c:,})',
                      f'≈ {churn_rev_dig/1e6:.0f} M Kč/rok', '#2563eb') +
            _calc_row('Celkový odchad odchodu',
                      f'{tot_churn_total:,} klientů',
                      f'≈ {churn_rev_total/1e6:.0f} M Kč/rok', '#ea580c')
        )
        + '    </div>\n'
        '    <div>\n'
        '      <div class="ct">🎂 Věkové skupiny v uzavíraných pobočkách</div>\n'
        + (age_bars_html if age_bars_html else
           '<div style="font-size:0.78rem;color:#94a3b8;padding:12px 0;">'
           'Data věkových skupin nejsou dostupná ve vstupním datasetu.</div>')
        + '    </div>\n'
        '  </div>\n'

        # top branches table
        '  <div class="ct">Top uzavírané pobočky dle počtu klientů</div>\n'
        '  <div style="overflow-x:auto;">\n'
        '  <table style="width:100%;border-collapse:collapse;min-width:600px;">\n'
        '    <thead><tr style="background:#fff4ed;">\n'
        '      <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#9a3412;'
        'border-bottom:1px solid #fed7aa;text-align:left;">Pobočka</th>\n'
        '      <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#9a3412;'
        'border-bottom:1px solid #fed7aa;text-align:left;">Město</th>\n'
        '      <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#9a3412;'
        'border-bottom:1px solid #fed7aa;text-align:right;">Klientů</th>\n'
        '      <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#9a3412;'
        'border-bottom:1px solid #fed7aa;">Segment</th>\n'
        '      <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#dc2626;'
        'border-bottom:1px solid #fed7aa;text-align:right;">Pobočkoví</th>\n'
        '      <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#2563eb;'
        'border-bottom:1px solid #fed7aa;text-align:right;">Digitální</th>\n'
        '      <th style="padding:5px 9px;font-size:0.68rem;font-weight:700;color:#ea580c;'
        'border-bottom:1px solid #fed7aa;text-align:right;">Odhad odchodu</th>\n'
        '    </tr></thead>\n'
        f'    <tbody>{_cli_rows_html}</tbody>\n'
        '  </table>\n'
        '  </div>\n'
        f'  <p style="margin-top:10px;font-size:0.69rem;color:#94a3b8;">'
        f'Pobočkoví klienti: míra závislosti × {BRANCH_DEP_CHURN*100:.0f}% odchodovost &nbsp;·&nbsp; '
        f'Digitální klienti: {CLIENT_CHURN_RATE*100:.0f}% odchodovost &nbsp;·&nbsp; {_cli_source_note}</p>\n'
        '</div>\n'   # /klientský pohled

        '</div>\n'  # /wrap

        # ── Footer ČS brand ───────────────────────────────────────────────────
        '<div class="ftr">\n'
        '  <div class="ftr-left">\n'
        f'    Česká spořitelna &nbsp;·&nbsp; Analýza optimalizace sítě poboček<br>\n'
        f'    {n_total} poboček v perimetru &nbsp;·&nbsp; Scénář 1 + Interaktivní model\n'
        '  </div>\n'
        '  <div style="text-align:right;">\n'
        '    <div class="ftr-claim">Ať se daří</div>\n'
        '    <div class="ftr-hash">#silnější</div>\n'
        '  </div>\n'
        '</div>\n'

        '<script>\n' + js_code + '\n</script>\n'
        '</body>\n'
        '</html>'
    )

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as fh:
            fh.write(html)
        print(f'  ✅ Report: {output_path}')

    return html


# ── Standalone ────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import os, sys, pickle

    _rs = None
    for _vname in ('df', 'rating_status'):
        try:
            _cand = eval(_vname)                          # noqa: S307
            if hasattr(_cand, 'empty') and not _cand.empty:
                _rs = _cand
                print(f'Používám `{_vname}` ({len(_rs)} řádků).')
                break
        except NameError:
            pass

    if _rs is None:
        for _p in ('rating_status.pkl',
                   '../vypocet_ir_2026/report_rating_2026_staticky.pkl'):
            if os.path.exists(_p):
                print(f'Načítám {_p}...')
                with open(_p, 'rb') as _f:
                    _rs = pickle.load(_f)
                break

    if _rs is None:
        print('DataFrame nenalezen. Přidej na začátek souboru:')
        print('  df = pd.read_pickle("../vypocet_ir_2026/report_rating_2026_staticky.pkl")')
        sys.exit(1)

    generate_network_analysis_report(_rs, output_path='report_network_analysis.html')
