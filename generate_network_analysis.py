"""
generate_network_analysis.py
Analýza optimalizace sítě poboček — tabulky, scénář 1, mapa

Spuštění:
  Standalone: df = pd.read_pickle("...pkl") na začátku souboru
  Z hlavního skriptu: from generate_network_analysis import generate_network_analysis_report
"""

import math, warnings
import numpy as np
import pandas as pd
import plotly.graph_objects as go

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

_CFG = {'displayModeBar': False, 'responsive': True}


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
    """GeoJSON polygon circle (pro Mapbox layers)."""
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

    # GPS orientace
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
    """Průměrná vzdálenost (m) k AVAIL_N_NEAREST nejbližším pobočkám v df."""
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


# ── Scénář 1 (nejpřísnější) ───────────────────────────────────────────────────

def apply_scenario_1(df):
    """
    Scénář 1 — nejpřísnější:
      • Mimo Praha/Brno: přesně 1 pobočka na město
        (přednost flagship, pak nejlepší IR kvintil)
      • Praha/Brno: max MAX_METRO_FLAGSHIP flagship poboček
        (ne-flagship odstraněny)
    """
    df = df.copy()
    df['sc1_keep'] = False

    _fmt  = {'flagship': 0, 'medium': 1, 'medium economy': 2, 'small': 3}
    df['_fmt_rank'] = df['BRANCH_FORMAT'].map(_fmt).fillna(9).astype(int)
    df['_iq']       = df['IR_Q'].fillna(3.0)

    for city, grp in df.groupby('_city'):
        if grp.empty:
            continue
        if city in METRO_CITIES:
            flg = grp[grp['BRANCH_FORMAT'] == 'flagship'].sort_values('_iq')
            keep_idx = flg.index[:MAX_METRO_FLAGSHIP]
            df.loc[keep_idx, 'sc1_keep'] = True
        else:
            best = grp.sort_values(['_fmt_rank', '_iq']).index[0]
            df.loc[best, 'sc1_keep'] = True

    return df


# ── Mapbox mapa se scénářem ───────────────────────────────────────────────────

def _make_sc1_map(df_sc1):
    kdf = df_sc1[df_sc1['sc1_keep']].dropna(subset=['_lat', '_lon'])
    cdf = df_sc1[~df_sc1['sc1_keep']].dropna(subset=['_lat', '_lon'])

    def _fc(rows):
        return {'type': 'FeatureCollection',
                'features': [_geo_circle(r['_lat'], r['_lon'])
                             for _, r in rows.iterrows()]}

    fc_keep  = _fc(kdf)
    fc_close = _fc(cdf)

    def _trace(sub, name, color, symbol, size):
        if sub.empty:
            return None
        iq_vals = sub['IR_Q'].fillna(-1).astype(int).astype(str).replace('-1', '—')
        return go.Scattermapbox(
            lat=sub['_lat'], lon=sub['_lon'],
            mode='markers', name=name,
            marker=dict(size=size, color=color),
            text=sub['BRANCH_NAME'],
            customdata=list(zip(
                sub['_city'],
                sub['BRANCH_FORMAT'],
                iq_vals,
                sub['PRIMARNI_KLIENTI'].astype(int),
            )),
            hovertemplate=(
                '<b>%{text}</b><br>Město: %{customdata[0]}<br>'
                'Formát: %{customdata[1]}<br>IR kvintil: %{customdata[2]}<br>'
                'Primárních klientů: %{customdata[3]:,}<extra></extra>'
            ),
        )

    traces = [t for t in [
        _trace(kdf, 'Zachovat', '#16a34a', 'circle', 12),
        _trace(cdf, 'Zavřít',   '#dc2626', 'circle', 9),
    ] if t is not None]

    lat_c = float(df_sc1['_lat'].dropna().mean()) if df_sc1['_lat'].notna().any() else 49.8
    lon_c = float(df_sc1['_lon'].dropna().mean()) if df_sc1['_lon'].notna().any() else 15.5

    fig = go.Figure(traces)
    fig.update_layout(
        mapbox=dict(
            accesstoken=MAPBOX_TOKEN,
            style='mapbox://styles/mapbox/light-v11',
            center=dict(lat=lat_c, lon=lon_c),
            zoom=6.4,
            layers=[
                dict(sourcetype='geojson', source=fc_keep,  type='fill',
                     color='#16a34a', opacity=0.09),
                dict(sourcetype='geojson', source=fc_keep,  type='line',
                     color='#16a34a', opacity=0.40),
                dict(sourcetype='geojson', source=fc_close, type='fill',
                     color='#dc2626', opacity=0.07),
                dict(sourcetype='geojson', source=fc_close, type='line',
                     color='#dc2626', opacity=0.28),
            ],
        ),
        height=560, margin=dict(l=0, r=0, t=0, b=0),
        legend=dict(x=0.01, y=0.98, bgcolor='rgba(255,255,255,.88)',
                    bordercolor='#e2e8f0', borderwidth=1, font=dict(size=11)),
    )
    return fig


# ── HTML pomocníci ────────────────────────────────────────────────────────────

def _fv(v, fmt):
    """Formátuj číslo dle formátu."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return '—'
    if fmt == 'int':   return f'{int(round(v)):,}'.replace(',', ' ')
    if fmt == 'f1':    return f'{v:.1f}'
    if fmt == 'f2':    return f'{v:.2f}'
    if fmt == 'pct':   return f'{v:.1f} %'
    if fmt == 'mczk':  return f'{v/1e6:.1f} M'
    if fmt == 'km':    return f'{v:.1f} km'
    if fmt == 'mkczk': return f'{v/1e6:.0f} M Kč'
    if fmt == 'm2':    return f'{v:.0f} m²'
    return str(v)


def _sym_row(label, before, after, fmt, good='neutral', bold_after=False):
    """Řádek tabulky dopadů se symbolem ▲▼→ a barevným označením."""
    try:
        b, a = float(before), float(after)
        if b == 0:
            pct, pct_s = 0.0, '—'
        else:
            pct   = (a - b) / abs(b) * 100
            pct_s = f'{abs(pct):.1f} %'
        delta = a - b

        if abs(pct) < 0.3:
            sym, col = '→', '#64748b'
        elif delta > 0:
            sym = '▲'
            col = '#16a34a' if good == 'up' else ('#dc2626' if good == 'down' else '#64748b')
        else:
            sym = '▼'
            col = '#16a34a' if good == 'down' else ('#dc2626' if good == 'up' else '#64748b')

        bv = _fv(b, fmt)
        av = _fv(a, fmt)
        dv = _fv(delta, fmt) if delta != 0 else '0'
    except Exception:
        sym, col, pct_s = '—', '#64748b', '—'
        bv = av = dv = '—'

    aw = 'font-weight:700;' if bold_after else ''
    return (
        f'<tr>'
        f'<td style="padding:6px 10px;font-size:0.82rem;font-weight:600;">{label}</td>'
        f'<td style="padding:6px 10px;text-align:right;font-size:0.82rem;color:#475569;">{bv}</td>'
        f'<td style="padding:6px 10px;text-align:right;font-size:0.82rem;{aw}color:{col};">{av}</td>'
        f'<td style="padding:6px 10px;text-align:right;font-size:0.82rem;color:{col};">{dv}</td>'
        f'<td style="padding:6px 10px;text-align:center;font-size:0.95rem;font-weight:800;'
        f'color:{col};">{sym} {pct_s}</td>'
        f'</tr>'
    )


def _calc_row(label, value, note='', col='#1e2a38'):
    return (
        f'<div style="display:flex;align-items:baseline;gap:10px;padding:8px 0;'
        f'border-bottom:1px solid #f1f5f9;">'
        f'<div style="flex:1;font-size:0.82rem;color:#374151;">{label}</div>'
        f'<div style="font-size:1.05rem;font-weight:800;color:{col};white-space:nowrap;">'
        f'{value}</div>'
        f'{f"<div style=\\"font-size:0.72rem;color:#94a3b8;\\">{note}</div>" if note else ""}'
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

    # Dostupnost po scénáři (jen zachované pobočky)
    print('  📍 Dostupnost po Scénáři 1...')
    kdf_avail = compute_network_availability(kdf)
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

    rev_per_cli  = total_rev / total_cli if total_cli > 0 else 0.0
    churn_cli    = c_cli * CLIENT_CHURN_RATE
    churn_rev    = churn_cli * rev_per_cli

    cli_per_ban_now  = total_cli / total_ban  if total_ban  > 0 else 0.0
    cli_per_ban_all  = total_cli / k_ban       if k_ban       > 0 else 0.0
    cli_per_ban_own  = k_cli     / k_ban       if k_ban       > 0 else 0.0

    # ── Korelace ──────────────────────────────────────────────────────────────
    _corr_map = {
        'Primární klienti': 'PRIMARNI_KLIENTI',  'Aktivní klienti': 'AKTIVNI_KLIENTI',
        'Fyzické schůzky':  'POCET_SCHUZEK_FYZICKY', 'Plocha (m²)': 'CELK_PLOCHA_POBOCKY_2026',
        'Bankéři':          'BANKERS_COUNT',      'Nové výnosy':  'OBJEM_VYNOSU_CZK',
        'Výnosy celkem':    'VYNOSY',             'C/I ratio':    'PRIME_NAKLADY/VYNOSY',
        'Nájemné':          'ROCNI_SPLATKY_S_DPH_CZK', 'IR kvintil': 'IR_Q',
        'Dostupnost (km)':  'avail_km',           'Kapacita (%)': 'capacity_utilization',
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

    corr_html = ''
    for _, v, a, b in pairs[:12]:
        col   = '#dc2626' if v > 0 else '#2563eb'
        sign  = '+' if v > 0 else ''
        strng = ('silná' if abs(v) > 0.65 else ('střední' if abs(v) > 0.4 else 'slabá'))
        bar_w = f'{abs(v)*100:.0f}%'
        bar_c = col
        corr_html += (
            f'<div style="display:flex;align-items:center;gap:12px;padding:7px 0;'
            f'border-bottom:1px solid #f1f5f9;">'
            f'<span style="font-size:1.05rem;font-weight:800;color:{col};'
            f'min-width:52px;text-align:right;">{sign}{v:.2f}</span>'
            f'<div style="flex:1;">'
            f'<div style="font-size:0.81rem;color:#1e2a38;font-weight:500;">'
            f'{a} <span style="color:#94a3b8;">↔</span> {b}</div>'
            f'<div style="background:#f1f5f9;border-radius:2px;height:4px;margin-top:3px;">'
            f'<div style="width:{bar_w};background:{bar_c};height:100%;border-radius:2px;opacity:.6;"></div></div>'
            f'</div>'
            f'<span style="font-size:0.68rem;color:#94a3b8;white-space:nowrap;">{strng}</span>'
            f'</div>'
        )

    # ── Tabulka poboček (top) ──────────────────────────────────────────────────
    _tbl_cols = [
        ('Pobočka',             'BRANCH_NAME',              'text',   'left'),
        ('Region',              '_region',                  'text',   'left'),
        ('Formát',              'BRANCH_FORMAT',            'text',   'left'),
        ('IR Q',                'IR_Q',                     'int',    'center'),
        ('Prim. klienti',       'PRIMARNI_KLIENTI',         'int',    'right'),
        ('Akt. klienti',        'AKTIVNI_KLIENTI',          'int',    'right'),
        ('Schůzky',             'POCET_SCHUZEK_FYZICKY',    'int',    'right'),
        ('Bankéři',             'BANKERS_COUNT',             'f1',    'right'),
        ('Výnosy (M Kč)',       'VYNOSY',                   'mczk',   'right'),
        ('Nové výn. (M Kč)',    'OBJEM_VYNOSU_CZK',         'mczk',   'right'),
        ('C/I (%)',             'PRIME_NAKLADY/VYNOSY',     'pct',    'right'),
        ('Plocha (m²)',         'CELK_PLOCHA_POBOCKY_2026', 'm2',     'right'),
        ('Nájemné (M Kč/rok)', 'ROCNI_SPLATKY_S_DPH_CZK', 'mczk',   'right'),
        ('Kapacita (%)',        'capacity_utilization',     'pct',    'right'),
        ('Dostupnost (km)',     'avail_km',                  'km',    'right'),
    ]

    df_sorted = df.sort_values('VYNOSY', ascending=False).reset_index(drop=True)

    def _cell(v, fmt, align):
        a_style = f'text-align:{align};'
        if fmt == 'text':
            return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{v}</td>'
        if fmt == 'int':
            sv = f'{int(round(float(v))):,}'.replace(',', ' ') if pd.notna(v) and str(v) not in ('—','') else '—'
            return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{sv}</td>'
        if fmt == 'mczk':
            sv = f'{float(v)/1e6:.1f}' if pd.notna(v) and float(v) != 0 else '—'
            return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{sv}</td>'
        if fmt == 'pct':
            sv = f'{float(v):.1f}%' if pd.notna(v) else '—'
            col = ''
            if fmt == 'pct' and 'capacity' in str(v):
                pass
            return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{sv}</td>'
        if fmt == 'f1':
            sv = f'{float(v):.1f}' if pd.notna(v) else '—'
            return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{sv}</td>'
        if fmt == 'm2':
            sv = f'{float(v):.0f}' if pd.notna(v) and float(v) > 0 else '—'
            return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{sv}</td>'
        if fmt == 'km':
            sv = f'{float(v):.1f}' if pd.notna(v) else '—'
            return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{sv}</td>'
        return f'<td style="padding:5px 9px;{a_style}font-size:0.8rem;">{v}</td>'

    def _cap_color(v):
        try:
            f = float(v)
            if f > 1.2: return '#dc2626'
            if f > 0.9: return '#f59e0b'
            return '#16a34a'
        except Exception:
            return '#64748b'

    def _ir_color(v):
        try:
            iv = int(float(v))
            return {1: '#16a34a', 2: '#65a30d', 3: '#f59e0b', 4: '#ea580c', 5: '#dc2626'}.get(iv, '#64748b')
        except Exception:
            return '#64748b'

    thead_cells = ''.join(
        f'<th style="padding:6px 9px;font-size:0.7rem;font-weight:700;color:#64748b;'
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
                if col_name == 'capacity_utilization':
                    try:
                        fv = float(v)
                        color = _cap_color(fv)
                        sv = f'{fv:.0%}'
                        cells += (f'<td style="padding:5px 9px;text-align:right;font-size:0.8rem;'
                                  f'font-weight:700;color:{color};">{sv}</td>')
                    except Exception:
                        cells += '<td style="padding:5px 9px;text-align:right;">—</td>'
                elif col_name == 'IR_Q':
                    try:
                        iv = int(float(v))
                        color = _ir_color(iv)
                        cells += (f'<td style="padding:5px 9px;text-align:center;font-size:0.8rem;'
                                  f'font-weight:700;color:{color};">{iv}</td>')
                    except Exception:
                        cells += '<td style="padding:5px 9px;text-align:center;">—</td>'
                elif col_name == 'BRANCH_FORMAT':
                    fmt_colors = {'flagship': '#2563eb', 'medium': '#16a34a',
                                  'medium economy': '#65a30d', 'small': '#64748b'}
                    c = fmt_colors.get(str(v).lower(), '#94a3b8')
                    cells += (f'<td style="padding:5px 9px;font-size:0.75rem;'
                              f'font-weight:600;color:{c};">{str(v).capitalize()}</td>')
                else:
                    cells += _cell(v, fmt, align)
            rows += f'<tr style="border-bottom:1px solid #f1f5f9;">{cells}</tr>\n'
        return rows

    preview_rows = _build_rows(df_sorted.head(5))
    rest_rows    = _build_rows(df_sorted.iloc[5:])
    n_rest       = len(df_sorted) - 5

    # ── Scénář 1: mapa ─────────────────────────────────────────────────────────
    print('  🗺️  Generování mapy Scénář 1...')
    g_map = _make_sc1_map(df_sc1).to_html(
        full_html=False, include_plotlyjs='cdn', config=_CFG
    )

    # ── Scénář 1: tabulka zavřených ───────────────────────────────────────────
    close_rows_html = ''
    for _, row in cdf.sort_values(['_region', 'PRIMARNI_KLIENTI'], ascending=[True, False]).iterrows():
        iq = f'{int(float(row["IR_Q"]))}' if pd.notna(row['IR_Q']) else '—'
        close_rows_html += (
            f'<tr style="border-bottom:1px solid #fef2f2;">'
            f'<td style="padding:5px 9px;font-size:0.79rem;font-weight:600;">{row["BRANCH_NAME"]}</td>'
            f'<td style="padding:5px 9px;font-size:0.74rem;color:#64748b;">{row["_city"]}</td>'
            f'<td style="padding:5px 9px;font-size:0.74rem;">{row["BRANCH_FORMAT"].capitalize()}</td>'
            f'<td style="padding:5px 9px;text-align:center;font-size:0.8rem;'
            f'font-weight:700;color:{_ir_color(row["IR_Q"])};">{iq}</td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.8rem;">'
            f'{int(row["PRIMARNI_KLIENTI"]):,}</td>'
            f'<td style="padding:5px 9px;text-align:right;font-size:0.8rem;">'
            f'{row["ROCNI_SPLATKY_S_DPH_CZK"]/1e6:.1f} M</td>'
            f'</tr>\n'
        )

    # ── Scénář 1: impact tabulka ──────────────────────────────────────────────
    impact_html = (
        _sym_row('Poboček',                     n_total,      n_keep,             'int',   'neutral') +
        _sym_row('Primárních klientů',           total_cli,    k_cli,              'int',   'up') +
        _sym_row('Aktivních klientů',            total_aktiv,  k_aktiv,            'int',   'up') +
        _sym_row('Výnosy celkem (M Kč)',         total_rev,    k_rev,              'mczk',  'up') +
        _sym_row('Bankéřů celkem',               total_ban,    k_ban,              'int',   'neutral') +
        _sym_row('Průměrná plocha pobočky (m²)', avg_plocha,   k_plocha,           'f1',    'up') +
        _sym_row('Průměrná dostupnost sítě (km)',avg_avail_km, avg_avail_after_km, 'km',    'down') +
        _sym_row('Nájemné celkem (M Kč/rok)',    total_rent,   k_rent,             'mczk',  'down') +
        _sym_row('C/I ratio průměr (%)',         avg_ci,       k_ci,               'pct',   'down')
    )

    # ── Kalkulace dopadů ──────────────────────────────────────────────────────
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
                  f'změna {k_plocha - avg_plocha:+.0f} m²', '#0891b2')
    )

    # ── HTML ──────────────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Analýza sítě poboček</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
     background:#f0f4fb;color:#1e2a38;}}
.hdr{{background:linear-gradient(135deg,#1a3a6c 0%,#2563eb 100%);
     color:white;padding:22px 32px 18px;}}
.hdr h1{{font-size:1.45rem;font-weight:800;margin-bottom:3px;}}
.hdr p{{font-size:0.8rem;opacity:.75;}}
.wrap{{max-width:1420px;margin:0 auto;padding:22px 18px;}}
.card{{background:white;border-radius:12px;padding:18px 20px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);margin-bottom:18px;}}
.ct{{font-size:0.71rem;font-weight:700;color:#2563eb;text-transform:uppercase;
    letter-spacing:.5px;margin-bottom:12px;}}
.two-col{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px;}}
@media(max-width:900px){{.two-col{{grid-template-columns:1fr;}}}}
.impact-tbl{{width:100%;border-collapse:collapse;}}
.impact-tbl th{{background:#f8fafc;padding:6px 10px;font-size:0.7rem;font-weight:700;
               color:#64748b;border-bottom:2px solid #e2e8f0;white-space:nowrap;}}
.impact-tbl tr:hover td{{background:#f8fafc;}}
.sc1-card{{border-top:3px solid #2563eb;}}
.sc1-title{{font-size:1rem;font-weight:700;color:#1d4ed8;margin-bottom:4px;}}
.sc1-desc{{font-size:0.78rem;color:#374151;margin-bottom:16px;}}
.badge{{display:inline-block;border-radius:5px;padding:2px 7px;font-size:0.7rem;font-weight:700;}}
.b-keep{{background:#ecfdf5;color:#059669;border:1px solid #a7f3d0;}}
.b-close{{background:#fef2f2;color:#dc2626;border:1px solid #fecaca;}}
.scroll-tbl{{max-height:340px;overflow-y:auto;}}
.show-more-btn{{
  display:inline-block;margin-top:10px;padding:7px 18px;
  background:#f0f4fb;border:1px solid #e2e8f0;border-radius:8px;
  font-size:0.78rem;font-weight:600;color:#2563eb;cursor:pointer;
  transition:background .15s;
}}
.show-more-btn:hover{{background:#dbeafe;}}
</style>
</head>
<body>
<div class="hdr">
  <h1>🏦 Analýza sítě poboček — optimalizace</h1>
  <p>{n_total} poboček v perimetru · Výpočet dostupnosti a kapacity</p>
</div>
<div class="wrap">

<!-- ══ Tabulka poboček ═════════════════════════════════════════════════════ -->
<div class="card">
  <div class="ct">📊 Přehled poboček — {n_total} poboček (seřazeno dle výnosů)</div>
  <div style="overflow-x:auto;">
  <table style="width:100%;border-collapse:collapse;min-width:1100px;">
    <thead><tr>{thead_cells}</tr></thead>
    <tbody>
{preview_rows}
    </tbody>
  </table>
  </div>
  <div id="branch-more" style="display:none;overflow-x:auto;margin-top:2px;">
  <table style="width:100%;border-collapse:collapse;min-width:1100px;">
    <thead><tr>{thead_cells}</tr></thead>
    <tbody>
{rest_rows}
    </tbody>
  </table>
  </div>
  <button class="show-more-btn" id="more-btn"
          onclick="(function(){{
            var el=document.getElementById('branch-more');
            var btn=document.getElementById('more-btn');
            var open=el.style.display!=='none';
            el.style.display=open?'none':'block';
            btn.textContent=open?'Zobrazit všech {n_rest} dalších poboček ▼':'Skrýt ▲';
          }})()">
    Zobrazit všech {n_rest} dalších poboček ▼
  </button>
</div>

<!-- ══ Korelace ═══════════════════════════════════════════════════════════ -->
<div class="card">
  <div class="ct">🔗 Nejvýznamnější korelace mezi metrikami</div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:0 30px;">
    <div>{corr_html[:len(corr_html)//2]}</div>
    <div>{corr_html[len(corr_html)//2:]}</div>
  </div>
  <p style="margin-top:10px;font-size:0.71rem;color:#94a3b8;">
    Pearsonův r &nbsp;·&nbsp; zobrazeny páry s |r| &gt; 0.2 &nbsp;·&nbsp; ±1 = perfektní lineární vztah
  </p>
</div>

<!-- ══ Scénář 1 ════════════════════════════════════════════════════════════ -->
<div class="card sc1-card">
  <div class="sc1-title">📋 Scénář 1 — nejpřísnější optimalizace</div>
  <div class="sc1-desc">
    <strong>Pravidlo:</strong> Mimo Prahu a Brno zůstane v každém městě <strong>jedna pobočka</strong>
    — vybrána nejlepší dle interního ratingu s preferencí formátu flagship.
    V Praze a Brně jsou zachovány flagshipy (max {MAX_METRO_FLAGSHIP} v každém).
    Ne-flagship pobočky v metropolích jsou uzavřeny.
    &nbsp;·&nbsp; Kružnice = 10 km dostupnosti.
  </div>

  <!-- Mapa -->
  {g_map}
  <div style="display:flex;gap:16px;margin:8px 0 16px;font-size:0.76rem;flex-wrap:wrap;">
    <span>🟢 <span class="badge b-keep">Zachovat ({n_keep})</span></span>
    <span>🔴 <span class="badge b-close">Zavřít ({n_close})</span></span>
    <span style="color:#94a3b8;">Kružnice = 10 km od pobočky</span>
  </div>

  <div class="two-col">

    <!-- Dopad na parametry -->
    <div>
      <div class="ct">📉 Dopad na parametry sítě</div>
      <div style="overflow-x:auto;">
      <table class="impact-tbl">
        <thead><tr>
          <th style="text-align:left;">Metrika</th>
          <th style="text-align:right;">Před</th>
          <th style="text-align:right;">Scénář 1</th>
          <th style="text-align:right;">Změna</th>
          <th style="text-align:center;">Trend</th>
        </tr></thead>
        <tbody>{impact_html}</tbody>
      </table>
      </div>
    </div>

    <!-- Kalkulace dopadů -->
    <div>
      <div class="ct">🔢 Detailní výpočet dopadů</div>
      {calcs_html}
    </div>

  </div>

  <!-- Tabulka zavřených poboček -->
  <div style="margin-top:16px;">
    <div class="ct">Pobočky navržené k uzavření ({n_close})</div>
    <div class="scroll-tbl">
    <table style="width:100%;border-collapse:collapse;">
      <thead><tr style="background:#fef2f2;">
        <th style="padding:5px 9px;font-size:0.7rem;font-weight:700;color:#b91c1c;
                   border-bottom:1px solid #fecaca;text-align:left;">Pobočka</th>
        <th style="padding:5px 9px;font-size:0.7rem;font-weight:700;color:#b91c1c;
                   border-bottom:1px solid #fecaca;text-align:left;">Město</th>
        <th style="padding:5px 9px;font-size:0.7rem;font-weight:700;color:#b91c1c;
                   border-bottom:1px solid #fecaca;text-align:left;">Formát</th>
        <th style="padding:5px 9px;font-size:0.7rem;font-weight:700;color:#b91c1c;
                   border-bottom:1px solid #fecaca;text-align:center;">IR Q</th>
        <th style="padding:5px 9px;font-size:0.7rem;font-weight:700;color:#b91c1c;
                   border-bottom:1px solid #fecaca;text-align:right;">Primárních klientů</th>
        <th style="padding:5px 9px;font-size:0.7rem;font-weight:700;color:#b91c1c;
                   border-bottom:1px solid #fecaca;text-align:right;">Nájemné (M Kč/rok)</th>
      </tr></thead>
      <tbody>{close_rows_html}</tbody>
    </table>
    </div>
  </div>

</div><!-- /sc1-card -->

</div><!-- /wrap -->
</body>
</html>"""

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
