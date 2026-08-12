"""
generate_network_analysis.py
Analýza optimalizace sítě poboček — scénáře, korelace, kapacita, mapa

Spuštění:
  1. Z internirating_report.py:
       from generate_network_analysis import generate_network_analysis_report
       generate_network_analysis_report(rating_status)
  2. Standalone (potřebuje rating_status.pkl):
       python generate_network_analysis.py
"""

import math
import json
import warnings
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── Konstanty ─────────────────────────────────────────────────────────────────
BANKER_CAPACITY = 1_500     # klientů na bankéře (kapacitní výpočet)
AVAIL_N_NEAREST = 5         # nejbližších poboček pro výpočet dostupnosti
MAX_METRO_FLAGSHIP = 3      # max flagship v Praze a Brně
METRO_CITIES = {'Praha', 'Brno'}

# Barvy formátů (categorical, fixed order)
FORMAT_COLORS = {
    'flagship':       '#2563eb',   # modrá
    'medium':         '#16a34a',   # zelená
    'medium economy': '#65a30d',   # žlutozelená
    'small':          '#64748b',   # šedá
    'unknown':        '#94a3b8',
}

_CFG = {'displayModeBar': False, 'responsive': True}


# ── Příprava dat ──────────────────────────────────────────────────────────────

def _city_base(city: str) -> str:
    """Praha 1–22 → Praha; Brno-* → Brno; ostatní beze změny."""
    import re
    if not isinstance(city, str):
        return ''
    c = city.strip()
    if re.match(r'^Praha\s*(\d+|[-–])', c, re.I):
        return 'Praha'
    if re.match(r'^Brno[\s\-–]', c, re.I) or c.lower() == 'brno':
        return 'Brno'
    return c


def _prep_df(rating_status: pd.DataFrame) -> pd.DataFrame:
    df = rating_status.copy()

    # Filtr: pouze pobočky v ratingovém perimetru a neuzavřené
    if 'IR_FLAG' in df.columns:
        df = df[df['IR_FLAG'].eq('Y')].copy()
    if 'BRANCH_CLOSED' in df.columns:
        df = df[~df['BRANCH_CLOSED'].eq(True)].copy()
    df = df.reset_index(drop=True)

    # GPS — detekce orientace (GPS_X = lat pokud medián v 48–52)
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

    # Normalizace měst
    df['_city'] = df.get('CITY', pd.Series('', index=df.index)).apply(_city_base)

    # Numerické slouce + default při absenci
    _num_defaults = {
        'PRIMARNI_KLIENTI':           0.0,
        'AKTIVNI_KLIENTI':            0.0,
        'POCET_SCHUZEK_FYZICKY':      0.0,
        'CELK_PLOCHA_POBOCKY_2026':   0.0,
        'BANKERS_COUNT':              0.0,
        'OBJEM_VYNOSU_CZK':           0.0,
        'VYNOSY':                     0.0,
        'PRIME_NAKLADY/VYNOSY':       np.nan,
        'ROCNI_SPLATKY_S_DPH_CZK':    0.0,
        'IR_Q':                       np.nan,
    }
    for col, dflt in _num_defaults.items():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(dflt)
        else:
            df[col] = dflt

    # Název a formát
    if 'BRANCH_NAME' not in df.columns:
        df['BRANCH_NAME'] = df.get('BRANCH_CODE', pd.Series('', index=df.index)).astype(str)
    if 'BRANCH_FORMAT' not in df.columns:
        df['BRANCH_FORMAT'] = 'unknown'
    df['BRANCH_FORMAT'] = df['BRANCH_FORMAT'].fillna('unknown').str.lower().str.strip()

    # Region
    reg_col = next((c for c in ['REGION_FIXED', 'REGION_NAME', 'REGION'] if c in df.columns), None)
    df['_region'] = df[reg_col].fillna('—') if reg_col else '—'

    return df


# ── Výpočet odvozených metrik ─────────────────────────────────────────────────

def _haversine(lat1, lon1, lat2, lon2):
    R = 6_371_000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a  = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2 * R * math.asin(math.sqrt(max(0.0, min(1.0, a))))


def compute_network_availability(df: pd.DataFrame) -> pd.Series:
    """Průměrná vzdálenost (m) k AVAIL_N_NEAREST nejbližším pobočkám."""
    valid = df.dropna(subset=['_lat', '_lon'])
    result = pd.Series(np.nan, index=df.index, name='network_availability')

    lats = valid['_lat'].values
    lons = valid['_lon'].values
    idxs = valid.index.values

    for i in range(len(idxs)):
        dists = [
            _haversine(lats[i], lons[i], lats[j], lons[j])
            for j in range(len(idxs)) if i != j
        ]
        if dists:
            dists.sort()
            result[idxs[i]] = float(np.mean(dists[:AVAIL_N_NEAREST]))

    return result


def compute_capacity_utilization(df: pd.DataFrame) -> pd.Series:
    """PRIMARNI_KLIENTI / (BANKERS_COUNT × BANKER_CAPACITY), max 5.0."""
    bc = df['BANKERS_COUNT'].clip(lower=0.01)
    return (df['PRIMARNI_KLIENTI'] / (bc * BANKER_CAPACITY)).clip(upper=5.0)


# ── Scénář 1 ──────────────────────────────────────────────────────────────────

def apply_scenario_1(df: pd.DataFrame) -> pd.DataFrame:
    """
    Scénář 1 — pravidla:
      1. Odebrány pobočky IR_Q ∈ {4, 5} (nejslabší)
      2. Max 1 pobočka na město (preferuje flagship, pak nejlepší IR)
      3. Praha / Brno: max MAX_METRO_FLAGSHIP flagshipů; ne-flagship odstraněny
    Vrací DataFrame s novým sloupcem 'sc1_keep' (bool).
    """
    df = df.copy()
    df['sc1_keep'] = True

    # Krok 1: výkonnostní filtr
    df.loc[df['IR_Q'].isin([4.0, 5.0]), 'sc1_keep'] = False

    # Pomocné sloupce pro řazení
    _fmt_order = {'flagship': 0, 'medium': 1, 'medium economy': 2, 'small': 3}
    df['_fmt_rank']   = df['BRANCH_FORMAT'].map(_fmt_order).fillna(9).astype(int)
    df['_ir_q_sort']  = df['IR_Q'].fillna(3.0)

    # Krok 2 + 3: filtr dle města
    kept_mask = df['sc1_keep'].copy()
    for city, grp in df[kept_mask].groupby('_city'):
        if len(grp) <= 1:
            continue
        is_metro = city in METRO_CITIES
        if is_metro:
            # Ne-flagship v metropoli odstraň
            non_flg = grp[grp['BRANCH_FORMAT'] != 'flagship']
            df.loc[non_flg.index, 'sc1_keep'] = False
            # Flagshipů max MAX_METRO_FLAGSHIP
            flg = grp[grp['BRANCH_FORMAT'] == 'flagship']
            if len(flg) > MAX_METRO_FLAGSHIP:
                remove = flg.sort_values(['_ir_q_sort', '_fmt_rank']).index[MAX_METRO_FLAGSHIP:]
                df.loc[remove, 'sc1_keep'] = False
        else:
            # Mimo metropoli: zachovej nejlepší 1
            keep_one = grp.sort_values(['_fmt_rank', '_ir_q_sort']).index[0]
            df.loc[grp.index.difference([keep_one]), 'sc1_keep'] = False

    return df


# ── Citlivostní skóre ─────────────────────────────────────────────────────────

def _norm_series(s: pd.Series) -> pd.Series:
    mn, mx = s.min(), s.max()
    if mx == mn:
        return pd.Series(0.5, index=s.index)
    return (s - mn) / (mx - mn)


def sensitivity_score(df: pd.DataFrame, w_access: float, w_revenue: float,
                      w_clients: float) -> pd.Series:
    """
    Kompozitní důležitostní skóre pobočky (vyšší = důležitější zachovat).
    Vstupy normovány do [0, 1] před vážením.
    """
    avail_s = _norm_series(
        df.get('network_availability', pd.Series(0.0, index=df.index)).fillna(0.0)
    )
    rev_s   = _norm_series(df['VYNOSY'].fillna(0.0))
    cli_s   = _norm_series(df['PRIMARNI_KLIENTI'].fillna(0.0))

    total = w_access + w_revenue + w_clients or 1.0
    return (w_access * avail_s + w_revenue * rev_s + w_clients * cli_s) / total


# ── Plotly grafy ──────────────────────────────────────────────────────────────

def _html(fig: go.Figure, first: bool = False) -> str:
    return fig.to_html(
        full_html=False,
        include_plotlyjs='cdn' if first else False,
        config=_CFG,
    )


def _make_corr_heatmap(df: pd.DataFrame) -> go.Figure:
    _col_map = {
        'Primární klienti':   'PRIMARNI_KLIENTI',
        'Aktivní klienti':    'AKTIVNI_KLIENTI',
        'Fyzické schůzky':    'POCET_SCHUZEK_FYZICKY',
        'Plocha (m²)':        'CELK_PLOCHA_POBOCKY_2026',
        'Bankéři':            'BANKERS_COUNT',
        'Nové výnosy':        'OBJEM_VYNOSU_CZK',
        'Výnosy celkem':      'VYNOSY',
        'C/I ratio':          'PRIME_NAKLADY/VYNOSY',
        'Nájemné (Kč/rok)':   'ROCNI_SPLATKY_S_DPH_CZK',
        'IR kvintil':         'IR_Q',
        'Dostupnost (m)':     'network_availability',
        'Kapacit. využití':   'capacity_utilization',
    }
    avail = {k: v for k, v in _col_map.items() if v in df.columns}
    sub   = df[[v for v in avail.values()]].rename(columns={v: k for k, v in avail.items()})
    sub   = sub.apply(pd.to_numeric, errors='coerce')
    corr  = sub.corr()
    labs  = corr.columns.tolist()
    z     = corr.values

    ann = [[f'{v:.2f}' if not np.isnan(v) else '' for v in row] for row in z]

    fig = go.Figure(go.Heatmap(
        z=z.tolist(), x=labs, y=labs,
        colorscale=[[0, '#2563eb'], [0.5, '#f8fafc'], [1, '#dc2626']],
        zmid=0, zmin=-1, zmax=1,
        text=ann, texttemplate='%{text}',
        textfont=dict(size=9),
        hovertemplate='<b>%{x}</b> ↔ <b>%{y}</b><br>r = %{z:.3f}<extra></extra>',
        colorbar=dict(title='r', thickness=12, len=0.75, tickvals=[-1, -0.5, 0, 0.5, 1]),
    ))
    fig.update_layout(
        height=500, margin=dict(l=5, r=5, t=5, b=5),
        xaxis=dict(tickangle=-40, automargin=True, side='bottom'),
        yaxis=dict(automargin=True),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(size=10),
    )
    return fig


def _make_map(df: pd.DataFrame, color_col: str = 'BRANCH_FORMAT') -> go.Figure:
    valid = df.dropna(subset=['_lat', '_lon']).copy()
    if valid.empty:
        return go.Figure()

    traces = []

    if color_col == 'sc1_keep':
        groups = [(True, 'Zachovat', '#16a34a', 'circle', 12),
                  (False, 'Zavřít', '#dc2626', 'circle-open', 9)]
        for keep, label, color, symbol, size in groups:
            g = valid[valid['sc1_keep'] == keep]
            if g.empty:
                continue
            traces.append(go.Scattermapbox(
                lat=g['_lat'], lon=g['_lon'],
                mode='markers',
                marker=dict(size=size, color=color),
                name=label,
                text=g['BRANCH_NAME'],
                customdata=np.stack([
                    g['_city'].values,
                    g['BRANCH_FORMAT'].values,
                    g['IR_Q'].fillna(-1).values,
                ], axis=1),
                hovertemplate=(
                    '<b>%{text}</b><br>Město: %{customdata[0]}<br>'
                    'Formát: %{customdata[1]}<br>IR kvintil: %{customdata[2]:.0f}'
                    '<extra></extra>'
                ),
            ))
    else:
        for fmt in ['flagship', 'medium', 'medium economy', 'small', 'unknown']:
            g = valid[valid['BRANCH_FORMAT'] == fmt]
            if g.empty:
                continue
            color = FORMAT_COLORS.get(fmt, '#94a3b8')
            traces.append(go.Scattermapbox(
                lat=g['_lat'], lon=g['_lon'],
                mode='markers',
                marker=dict(
                    size=[max(8, min(18, int(c / 600) + 8)) for c in g['PRIMARNI_KLIENTI']],
                    color=color,
                    opacity=0.85,
                ),
                name=fmt.capitalize(),
                text=g['BRANCH_NAME'],
                customdata=np.stack([
                    g['_city'].values,
                    g['IR_Q'].fillna(-1).values,
                    g['PRIMARNI_KLIENTI'].values,
                    g['VYNOSY'].values / 1e6,
                ], axis=1),
                hovertemplate=(
                    '<b>%{text}</b><br>Město: %{customdata[0]}<br>'
                    'IR kvintil: %{customdata[1]:.0f}<br>'
                    'Primárních klientů: %{customdata[2]:,.0f}<br>'
                    'Výnosy: %{customdata[3]:.1f} M Kč'
                    '<extra></extra>'
                ),
            ))

    lat_c = float(valid['_lat'].mean())
    lon_c = float(valid['_lon'].mean())

    fig = go.Figure(traces)
    fig.update_layout(
        mapbox=dict(style='open-street-map', center=dict(lat=lat_c, lon=lon_c), zoom=6.4),
        height=500, margin=dict(l=0, r=0, t=0, b=0),
        legend=dict(x=0.01, y=0.99, bgcolor='rgba(255,255,255,.88)',
                    bordercolor='#e2e8f0', borderwidth=1, font=dict(size=11)),
    )
    return fig


def _make_capacity_chart(df: pd.DataFrame) -> go.Figure:
    ds = df.sort_values('capacity_utilization', ascending=True).tail(40)
    colors = ['#dc2626' if v > 1.2 else ('#f59e0b' if v > 0.9 else '#16a34a')
              for v in ds['capacity_utilization']]
    fig = go.Figure(go.Bar(
        x=ds['capacity_utilization'], y=ds['BRANCH_NAME'],
        orientation='h', marker_color=colors,
        text=[f'{v:.0%}' for v in ds['capacity_utilization']],
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Kapacita: %{x:.1%}<extra></extra>',
    ))
    fig.add_vline(x=1.0, line=dict(color='#dc2626', width=1.5, dash='dash'),
                  annotation_text='Plná kapacita', annotation_font_size=10)
    fig.update_layout(
        height=max(300, len(ds) * 22 + 60),
        xaxis_title='Kapacitní využití (primárních klientů / bankéři × 1 500)',
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(size=10), margin=dict(l=5, r=80, t=15, b=10),
        xaxis=dict(showgrid=True, gridcolor='#f0f0f0', tickformat='.0%'),
        yaxis=dict(automargin=True),
    )
    return fig


def _make_availability_chart(df: pd.DataFrame) -> go.Figure:
    ds = (df.dropna(subset=['network_availability'])
          .sort_values('network_availability', ascending=False)
          .head(30))
    vals_km = ds['network_availability'] / 1000
    colors = ['#dc2626' if v > 20 else ('#f59e0b' if v > 10 else '#16a34a') for v in vals_km]
    fig = go.Figure(go.Bar(
        x=vals_km, y=ds['BRANCH_NAME'],
        orientation='h', marker_color=colors,
        text=[f'{v:.1f} km' for v in vals_km],
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Dostupnost: %{x:.2f} km<br>Město: %{customdata}<extra></extra>',
        customdata=ds['_city'],
    ))
    fig.update_layout(
        height=max(300, len(ds) * 22 + 60),
        xaxis_title=f'Prům. vzdálenost k {AVAIL_N_NEAREST} nejbližším pobočkám (km)',
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(size=10), margin=dict(l=5, r=80, t=10, b=10),
        xaxis=dict(showgrid=True, gridcolor='#f0f0f0'),
        yaxis=dict(automargin=True),
    )
    return fig


def _make_scenario_impact(df_base: pd.DataFrame, df_sc1: pd.DataFrame) -> go.Figure:
    keep = df_sc1['sc1_keep']
    specs = [
        ('Pobočky',           None,                          'count'),
        ('Primárních klientů', 'PRIMARNI_KLIENTI',           'sum'),
        ('Výnosy (M Kč)',     'VYNOSY',                     'sum_m'),
        ('Nové výnosy (M Kč)', 'OBJEM_VYNOSU_CZK',          'sum_m'),
        ('Fyzické schůzky',   'POCET_SCHUZEK_FYZICKY',      'sum'),
        ('Bankéři',           'BANKERS_COUNT',               'sum'),
        ('Nájemné (M Kč/rok)', 'ROCNI_SPLATKY_S_DPH_CZK',  'sum_m'),
    ]

    labels, before_v, after_v, pcts = [], [], [], []
    for lbl, col, mode in specs:
        if mode == 'count':
            b, a = len(df_base), int(keep.sum())
        elif mode == 'sum':
            if col not in df_base.columns:
                continue
            b = df_base[col].sum()
            a = df_base.loc[keep, col].sum()
        else:  # sum_m
            if col not in df_base.columns:
                continue
            b = df_base[col].sum() / 1e6
            a = df_base.loc[keep, col].sum() / 1e6
        pct = (a - b) / b * 100 if b else 0
        labels.append(lbl)
        before_v.append(float(b))
        after_v.append(float(a))
        pcts.append(float(pct))

    clrs = ['#16a34a' if p >= 0 else '#dc2626' for p in pcts]

    fig = make_subplots(
        rows=1, cols=2,
        column_widths=[0.42, 0.58],
        subplot_titles=['Relativní změna (%)', 'Absolutní srovnání'],
    )
    fig.add_trace(go.Bar(
        y=labels, x=pcts, orientation='h',
        marker_color=clrs,
        text=[f'{p:+.1f}%' for p in pcts], textposition='outside',
        showlegend=False,
        hovertemplate='<b>%{y}</b><br>%{x:+.1f}%<extra></extra>',
    ), row=1, col=1)
    fig.add_trace(go.Bar(
        y=labels, x=before_v, orientation='h', name='Před',
        marker_color='#93c5fd', marker_line_width=0,
        hovertemplate='<b>%{y}</b><br>Před: %{x:,.1f}<extra></extra>',
    ), row=1, col=2)
    fig.add_trace(go.Bar(
        y=labels, x=after_v, orientation='h', name='Po scénáři',
        marker_color='#2563eb', marker_line_width=0,
        hovertemplate='<b>%{y}</b><br>Po: %{x:,.1f}<extra></extra>',
    ), row=1, col=2)

    fig.update_layout(
        height=320, barmode='group',
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(size=10), margin=dict(l=5, r=60, t=30, b=5),
        xaxis=dict(showgrid=True, gridcolor='#f0f0f0',
                   zeroline=True, zerolinecolor='#94a3b8', ticksuffix='%'),
        xaxis2=dict(showgrid=True, gridcolor='#f0f0f0'),
        yaxis=dict(automargin=True), yaxis2=dict(automargin=True),
        legend=dict(orientation='h', x=0.35, y=1.08, font=dict(size=10)),
    )
    return fig


def _make_sensitivity_chart(df: pd.DataFrame, n_steps: int = 12) -> go.Figure:
    """Citlivostní analýza: % zachovaných metrik vs. % zachovaných poboček."""
    n   = len(df)
    lo  = max(1, int(n * 0.35))
    xs  = list(range(lo, n + 1, max(1, (n - lo) // n_steps))) + [n]
    xs  = sorted(set(xs))

    total_rev = df['VYNOSY'].sum() or 1.0
    total_cli = df['PRIMARNI_KLIENTI'].sum() or 1.0

    strategies = [
        ('Dostupnost',   (1.0, 0.0, 0.0), '#7c3aed'),
        ('Výnosy',       (0.0, 1.0, 0.0), '#2563eb'),
        ('Klienti',      (0.0, 0.0, 1.0), '#16a34a'),
        ('Vyvážená',     (1.0, 1.0, 1.0), '#d97706'),
    ]

    fig = go.Figure()

    for sname, (wa, wr, wc), col in strategies:
        score  = sensitivity_score(df, wa, wr, wc)
        xpct, yrev, ycli = [], [], []
        for k in xs:
            top   = score.nlargest(k).index
            xpct.append(round(k / n * 100, 1))
            yrev.append(round(df.loc[top, 'VYNOSY'].sum() / total_rev * 100, 2))
            ycli.append(round(df.loc[top, 'PRIMARNI_KLIENTI'].sum() / total_cli * 100, 2))

        fig.add_trace(go.Scatter(
            x=xpct, y=yrev, name=f'{sname} — výnosy',
            mode='lines+markers', line=dict(color=col, width=2),
            marker=dict(size=5, color=col),
            hovertemplate=f'<b>{sname} · výnosy</b><br>Poboček: %{{x:.0f}}%<br>Výnosy zachovány: %{{y:.1f}}%<extra></extra>',
        ))
        fig.add_trace(go.Scatter(
            x=xpct, y=ycli, name=f'{sname} — klienti',
            mode='lines', line=dict(color=col, width=1.5, dash='dot'),
            hovertemplate=f'<b>{sname} · klienti</b><br>Poboček: %{{x:.0f}}%<br>Klienti zachováni: %{{y:.1f}}%<extra></extra>',
        ))

    # Diagonála jako reference
    fig.add_shape(type='line', x0=35, y0=35, x1=100, y1=100,
                  line=dict(color='#cbd5e1', width=1, dash='dash'))
    fig.add_annotation(x=72, y=66, text='Lineární reference', showarrow=False,
                       font=dict(size=9, color='#94a3b8'), textangle=-38)

    fig.update_layout(
        height=400,
        xaxis_title='Zachovaných poboček (%)',
        yaxis_title='Zachovaných metrik (%)',
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(size=10), margin=dict(l=5, r=10, t=10, b=40),
        xaxis=dict(showgrid=True, gridcolor='#f0f0f0', ticksuffix='%', range=[32, 103]),
        yaxis=dict(showgrid=True, gridcolor='#f0f0f0', ticksuffix='%', range=[32, 103]),
        legend=dict(orientation='h', x=0, y=1.10, font=dict(size=9),
                    tracegroupgap=0),
    )
    return fig


def _make_optimization_scatter(df: pd.DataFrame, to_close: pd.DataFrame) -> go.Figure:
    to_keep = df[~df.index.isin(to_close.index)]

    def _trace(sub, name, color, symbol, size_base):
        sizes = [max(8, min(22, int(b) * 4 + 6)) for b in sub['BANKERS_COUNT']]
        return go.Scatter(
            x=(sub['VYNOSY'] / 1e6).round(2), y=sub['PRIMARNI_KLIENTI'],
            mode='markers', name=name,
            marker=dict(size=sizes, color=color, symbol=symbol, opacity=0.78,
                        line=dict(width=1.5, color='white')),
            text=sub['BRANCH_NAME'],
            customdata=np.stack([
                sub['_city'].values,
                sub['BRANCH_FORMAT'].values,
                sub['IR_Q'].fillna(-1).values,
                (sub['network_availability'].fillna(0) / 1000).round(1).values,
            ], axis=1),
            hovertemplate=(
                '<b>%{text}</b><br>Město: %{customdata[0]}<br>'
                'Formát: %{customdata[1]}<br>IR kvintil: %{customdata[2]:.0f}<br>'
                'Výnosy: %{x:.1f} M Kč<br>Klientů: %{y:,.0f}<br>'
                'Dostupnost: %{customdata[3]:.1f} km<extra></extra>'
            ),
        )

    fig = go.Figure([
        _trace(to_keep, 'Zachovat', '#2563eb', 'circle', 10),
        _trace(to_close, 'Zavřít', '#dc2626', 'x', 8),
    ])
    fig.update_layout(
        height=440,
        xaxis_title='Výnosy celkem (M Kč)',
        yaxis_title='Primárních klientů',
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(size=11), margin=dict(l=5, r=10, t=10, b=40),
        xaxis=dict(showgrid=True, gridcolor='#f0f0f0'),
        yaxis=dict(showgrid=True, gridcolor='#f0f0f0'),
        legend=dict(orientation='h', x=0, y=1.08),
    )
    return fig


# ── HTML helpers ──────────────────────────────────────────────────────────────

def _kpi(val: str, lbl: str, col='#2563eb', bg='#eff6ff', bo='#bfdbfe') -> str:
    return (
        f'<div style="background:{bg};border:1px solid {bo};border-radius:10px;'
        f'padding:12px 18px;text-align:center;min-width:130px;flex:1;">'
        f'<div style="font-size:1.3rem;font-weight:800;color:{col};">{val}</div>'
        f'<div style="font-size:0.67rem;color:#6b7280;text-transform:uppercase;'
        f'letter-spacing:.3px;margin-top:2px;">{lbl}</div>'
        f'</div>'
    )


def _branch_table_rows(df_rows: pd.DataFrame, cols: list) -> str:
    """Generuje <tr> řádky pro tabulku poboček."""
    rows = ''
    for _, r in df_rows.iterrows():
        cells = ''
        for c, fmt, align in cols:
            v = r.get(c, '—')
            if fmt == 'int':
                v = f'{int(v):,}' if pd.notna(v) and v != '—' else '—'
            elif fmt == 'float1':
                v = f'{float(v):.1f}' if pd.notna(v) and v != '—' else '—'
            elif fmt == 'mczk':
                v = f'{float(v)/1e6:.1f} M' if pd.notna(v) and v != '—' and float(v) else '—'
            elif fmt == 'pct':
                v = f'{float(v):.0f}%' if pd.notna(v) and v != '—' else '—'
            a = f'text-align:{align};' if align != 'left' else ''
            cells += f'<td style="padding:4px 8px;font-size:0.78rem;{a}">{v}</td>'
        rows += f'<tr>{cells}</tr>\n'
    return rows


# ── Hlavní funkce ─────────────────────────────────────────────────────────────

def generate_network_analysis_report(
    rating_status: pd.DataFrame,
    output_path: str = 'report_network_analysis.html',
) -> str:
    """
    Generuje kompletní HTML report analýzy sítě poboček.

    Args:
        rating_status: hlavní DataFrame z internirating_report.py
        output_path:   pokud je zadán, zapíše HTML do souboru
    Returns:
        HTML string
    """
    warnings.filterwarnings('ignore')

    print('  📊 Příprava dat pro síťovou analýzu...')
    df = _prep_df(rating_status)
    n_total = len(df)
    print(f'     {n_total} poboček v perimetru')

    print('  📍 Výpočet dostupnosti sítě (O(n²) vzdálenostní matice)...')
    df['network_availability'] = compute_network_availability(df)
    df['capacity_utilization'] = compute_capacity_utilization(df)
    print('  ✓ Dostupnost a kapacita vypočteny')

    print('  🔀 Aplikace Scénáře 1...')
    df_sc1 = apply_scenario_1(df)
    n_keep  = int(df_sc1['sc1_keep'].sum())
    n_close = n_total - n_keep

    # Optimalizace — složené skóre (rovné váhy)
    df['_opt_score']   = sensitivity_score(df, 1.0, 1.0, 1.0)
    n_opt_close        = max(1, n_total // 5)
    df_opt_close       = df.nsmallest(n_opt_close, '_opt_score')

    # ── Statistiky ─────────────────────────────────────────────────────────────
    total_clients = int(df['PRIMARNI_KLIENTI'].sum())
    total_revenue = float(df['VYNOSY'].sum())
    total_bankers = int(df['BANKERS_COUNT'].sum())
    overloaded    = int((df['capacity_utilization'] > 1.0).sum())
    avg_avail_km  = float(df['network_availability'].dropna().mean() / 1000)

    kdf = df_sc1[df_sc1['sc1_keep']]
    cdf = df_sc1[~df_sc1['sc1_keep']]

    # ── Korelace — top páry ────────────────────────────────────────────────────
    _corr_map = {
        'PRIMARNI_KLIENTI': 'Primární klienti',   'AKTIVNI_KLIENTI': 'Aktivní klienti',
        'POCET_SCHUZEK_FYZICKY': 'Fyzické schůzky', 'CELK_PLOCHA_POBOCKY_2026': 'Plocha',
        'BANKERS_COUNT': 'Bankéři',                 'OBJEM_VYNOSU_CZK': 'Nové výnosy',
        'VYNOSY': 'Výnosy celkem',                  'PRIME_NAKLADY/VYNOSY': 'C/I ratio',
        'ROCNI_SPLATKY_S_DPH_CZK': 'Nájemné',      'IR_Q': 'IR kvintil',
        'network_availability': 'Dostupnost sítě',  'capacity_utilization': 'Kapacit. využití',
    }
    avail_cm = {k: v for k, v in _corr_map.items() if k in df.columns}
    corr_sub = df[[k for k in avail_cm]].rename(columns=avail_cm).apply(pd.to_numeric, errors='coerce')
    corr_mat = corr_sub.corr()
    names    = corr_mat.columns.tolist()

    pairs = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            v = corr_mat.iloc[i, j]
            if not np.isnan(v):
                pairs.append((abs(v), v, names[i], names[j]))
    pairs.sort(key=lambda x: -x[0])

    corr_rows_html = ''
    for _, v, a, b in pairs[:10]:
        col   = '#dc2626' if v > 0 else '#2563eb'
        sign  = '+' if v > 0 else ''
        strng = 'silná' if abs(v) > 0.7 else ('střední' if abs(v) > 0.4 else 'slabá')
        corr_rows_html += (
            f'<div style="display:flex;align-items:center;gap:10px;padding:7px 0;'
            f'border-bottom:1px solid #f1f5f9;">'
            f'<span style="font-size:1.05rem;font-weight:800;color:{col};'
            f'min-width:55px;text-align:right;">{sign}{v:.2f}</span>'
            f'<div><div style="font-size:0.8rem;color:#1e2a38;">{a} ↔ {b}</div>'
            f'<div style="font-size:0.68rem;color:#94a3b8;">{strng} korelace</div></div>'
            f'</div>'
        )

    # ── KPI tiles ─────────────────────────────────────────────────────────────
    kpis_base = ''.join([
        _kpi(f'{n_total}', 'Poboček celkem'),
        _kpi(f'{total_clients:,.0f}', 'Primárních klientů', '#7c3aed', '#faf5ff', '#ddd6fe'),
        _kpi(f'{total_revenue/1e6:.0f} M Kč', 'Výnosy celkem', '#059669', '#ecfdf5', '#a7f3d0'),
        _kpi(f'{total_bankers}', 'Bankéřů celkem', '#0891b2', '#e0f2fe', '#bae6fd'),
        _kpi(f'{avg_avail_km:.1f} km', 'Prům. dostupnost', '#d97706', '#fff7ed', '#fed7aa'),
        _kpi(f'{overloaded}', 'Přetížených', '#dc2626', '#fef2f2', '#fecaca'),
    ])

    save_rent_sc1 = float(cdf['ROCNI_SPLATKY_S_DPH_CZK'].sum())
    kpis_sc1 = ''.join([
        _kpi(f'{n_keep}', 'Zachováno poboček', '#059669', '#ecfdf5', '#a7f3d0'),
        _kpi(f'{n_close}', 'Navrženo k zavření', '#dc2626', '#fef2f2', '#fecaca'),
        _kpi(f'{kdf["PRIMARNI_KLIENTI"].sum()/total_clients*100:.0f}%', 'Klientů zachováno', '#7c3aed', '#faf5ff', '#ddd6fe'),
        _kpi(f'{kdf["VYNOSY"].sum()/total_revenue*100:.0f}%', 'Výnosů zachováno', '#059669', '#ecfdf5', '#a7f3d0'),
        _kpi(f'{kdf["BANKERS_COUNT"].sum():.0f}/{total_bankers}', 'Bankéřů zůstane', '#0891b2', '#e0f2fe', '#bae6fd'),
        _kpi(f'{save_rent_sc1/1e6:.0f} M Kč', 'Úspora nájmů/rok', '#d97706', '#fff7ed', '#fed7aa'),
    ])

    # ── Tabulka zavřených poboček — Scénář 1 ─────────────────────────────────
    sc1_closed_rows = ''
    for _, r in cdf.sort_values(['_ir_q_sort', 'PRIMARNI_KLIENTI'], ascending=[False, False]).iterrows():
        ir_q = f'{int(r["IR_Q"])}' if pd.notna(r['IR_Q']) else '—'
        sc1_closed_rows += (
            f'<tr>'
            f'<td style="padding:4px 8px;font-size:0.79rem;font-weight:600;">{r["BRANCH_NAME"]}</td>'
            f'<td style="padding:4px 8px;font-size:0.74rem;color:#64748b;">{r["_city"]}</td>'
            f'<td style="padding:4px 8px;font-size:0.74rem;">{r["BRANCH_FORMAT"].capitalize()}</td>'
            f'<td style="padding:4px 8px;text-align:center;">{ir_q}</td>'
            f'<td style="padding:4px 8px;text-align:right;color:#2563eb;">'
            f'{int(r["PRIMARNI_KLIENTI"]):,}</td>'
            f'</tr>\n'
        )

    # ── Tabulka pro optimalizaci ───────────────────────────────────────────────
    opt_close_rows = ''
    for _, r in df_opt_close.sort_values('_opt_score').iterrows():
        opt_close_rows += (
            f'<tr>'
            f'<td style="padding:4px 8px;font-size:0.79rem;font-weight:600;">{r["BRANCH_NAME"]}</td>'
            f'<td style="padding:4px 8px;font-size:0.74rem;color:#64748b;">{r["_city"]}</td>'
            f'<td style="padding:4px 8px;font-size:0.74rem;">{r["BRANCH_FORMAT"].capitalize()}</td>'
            f'<td style="padding:4px 8px;text-align:right;color:#2563eb;">'
            f'{int(r["PRIMARNI_KLIENTI"]):,}</td>'
            f'<td style="padding:4px 8px;text-align:right;">'
            f'{r["VYNOSY"]/1e6:.1f} M</td>'
            f'<td style="padding:4px 8px;text-align:right;">'
            f'{r["network_availability"]/1000:.1f} km</td>'
            f'</tr>\n'
        )

    opt_save_rent = float(df_opt_close['ROCNI_SPLATKY_S_DPH_CZK'].sum())
    opt_rev_loss  = float(df_opt_close['VYNOSY'].sum())

    # ── Plotly grafy ──────────────────────────────────────────────────────────
    print('  📈 Generování grafů...')
    g_corr   = _html(_make_corr_heatmap(df), first=True)
    g_cap    = _html(_make_capacity_chart(df))
    g_avail  = _html(_make_availability_chart(df))
    g_map    = _html(_make_map(df, 'BRANCH_FORMAT'))
    g_impact = _html(_make_scenario_impact(df, df_sc1))
    g_mapsc1 = _html(_make_map(df_sc1, 'sc1_keep'))
    g_sensit = _html(_make_sensitivity_chart(df))
    g_opt    = _html(_make_optimization_scatter(df, df_opt_close))

    # ── HTML ──────────────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Analýza optimalizace sítě poboček</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
     background:#f0f4fb;color:#1e2a38;line-height:1.5;}}
.hdr{{background:linear-gradient(135deg,#1a3a6c 0%,#2563eb 100%);color:white;
     padding:24px 32px 20px;}}
.hdr h1{{font-size:1.5rem;font-weight:800;margin-bottom:3px;}}
.hdr p{{font-size:0.8rem;opacity:.75;}}
.wrap{{max-width:1380px;margin:0 auto;padding:22px 18px;}}
.card{{background:white;border-radius:12px;padding:18px 20px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);margin-bottom:18px;}}
.ct{{font-size:0.71rem;font-weight:700;color:#2563eb;text-transform:uppercase;
    letter-spacing:.5px;margin-bottom:12px;}}
.krow{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px;}}
.two-col{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px;}}
@media(max-width:860px){{.two-col{{grid-template-columns:1fr;}}}}
.three-col{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;margin-bottom:18px;}}
@media(max-width:1100px){{.three-col{{grid-template-columns:1fr 1fr;}}}}
@media(max-width:640px){{.three-col{{grid-template-columns:1fr;}}}}
table{{width:100%;border-collapse:collapse;font-size:0.8rem;}}
th{{background:#f8fafc;padding:6px 8px;font-size:0.7rem;font-weight:700;
   color:#64748b;border-bottom:2px solid #e2e8f0;white-space:nowrap;}}
td{{padding:5px 8px;border-bottom:1px solid #f1f5f9;vertical-align:middle;}}
tr:last-child td{{border-bottom:none;}}
tr:hover td{{background:#f8fafc;}}
.note{{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;
      padding:8px 14px;font-size:0.75rem;color:#92400e;margin-bottom:12px;}}
.sc-banner{{border-left:4px solid #2563eb;padding-left:14px;margin-bottom:12px;}}
.sc-banner h3{{font-size:0.95rem;font-weight:700;color:#1d4ed8;margin-bottom:3px;}}
.sc-banner p{{font-size:0.78rem;color:#374151;}}
.info-box{{background:#f8fafc;border-radius:8px;padding:14px 16px;font-size:0.82rem;
          line-height:1.7;color:#374151;}}
.info-box strong{{color:#1e2a38;}}
.badge{{display:inline-block;border-radius:5px;padding:2px 7px;font-size:0.7rem;font-weight:700;}}
.b-keep{{background:#ecfdf5;color:#059669;border:1px solid #a7f3d0;}}
.b-close{{background:#fef2f2;color:#dc2626;border:1px solid #fecaca;}}
.scroll-tbl{{max-height:380px;overflow-y:auto;}}
details > summary{{
  cursor:pointer;padding:10px 0;list-style:none;
  font-size:0.71rem;font-weight:700;color:#2563eb;
  text-transform:uppercase;letter-spacing:.5px;
  display:flex;align-items:center;gap:6px;
}}
details > summary::-webkit-details-marker{{display:none;}}
</style>
</head>
<body>
<div class="hdr">
  <h1>🏦 Analýza optimalizace sítě poboček</h1>
  <p>Korelace · Kapacita · Dostupnost · Scénáře · Citlivostní analýza · Mapa</p>
</div>
<div class="wrap">

<!-- ══ KPIs ══════════════════════════════════════════════════════════════════ -->
<div class="krow">{kpis_base}</div>

<!-- ══ Korelace ══════════════════════════════════════════════════════════════ -->
<div class="two-col">
  <div class="card">
    <div class="ct">🔗 Korelační matice klíčových metrik</div>
    <div style="overflow-x:auto;">{g_corr}</div>
  </div>
  <div class="card">
    <div class="ct">🏆 Nejvýznamnější korelace</div>
    {corr_rows_html}
    <p style="margin-top:10px;font-size:0.72rem;color:#94a3b8;">
      Pearsonův r · ±1 = perfektní lineární vztah · ≥0.7 silná korelace
    </p>
  </div>
</div>

<!-- ══ Kapacitní analýza ═══════════════════════════════════════════════════ -->
<div class="card">
  <div class="ct">👥 Kapacitní analýza — vytížení poboček (Top 40)</div>
  <div class="note">
    <strong>Vzorec:</strong> Kapacitní využití = primární klienti ÷ (bankéři × {BANKER_CAPACITY:,})
    &nbsp;·&nbsp; 🔴 &gt;120% přetíženo &nbsp;·&nbsp; 🟡 90–120% na hraně &nbsp;·&nbsp; 🟢 &lt;90% OK
    &nbsp;·&nbsp; Přetížených pobočen: <strong>{overloaded}</strong>
  </div>
  <div style="overflow-x:auto;">{g_cap}</div>
</div>

<!-- ══ Dostupnost sítě ════════════════════════════════════════════════════ -->
<div class="card">
  <div class="ct">📍 Dostupnost vlastní sítě — průměrná vzdálenost k {AVAIL_N_NEAREST} nejbližším pobočkám</div>
  <div class="note">
    Čím vyšší vzdálenost, tím je pobočka <strong>geograficky izolovanější</strong> — zákazníci nemají
    blízkou alternativu. Tyto pobočky jsou při optimalizaci kritické zachovat.
    Zobrazeno top 30 s největší vzdáleností.
  </div>
  <div style="overflow-x:auto;">{g_avail}</div>
</div>

<!-- ══ Mapa — aktuální stav ════════════════════════════════════════════════ -->
<div class="card">
  <div class="ct">🗺️ Mapa sítě poboček — aktuální stav (barva = formát, velikost = počet klientů)</div>
  {g_map}
</div>

<!-- ══ Scénář 1 ═══════════════════════════════════════════════════════════ -->
<div class="card" style="border-top:3px solid #2563eb;">
  <div class="sc-banner">
    <h3>📋 Scénář 1 — výkonnostní filtr + jedna pobočka ve městě</h3>
    <p>
      <strong>Pravidlo 1:</strong> Odebrány pobočky s IR kvintil 4 a 5 (nejslabší výkonnost)
      &nbsp;·&nbsp;
      <strong>Pravidlo 2:</strong> Max 1 pobočka na město — priorita flagship, pak nejlepší IR
      &nbsp;·&nbsp;
      <strong>Pravidlo 3:</strong> Praha a Brno — max {MAX_METRO_FLAGSHIP} flagship; ne-flagship v metropolích odstraněny
    </p>
  </div>

  <div class="krow" style="margin-bottom:16px;">{kpis_sc1}</div>

  <div class="ct">Dopad scénáře na klíčové metriky</div>
  <div style="overflow-x:auto;margin-bottom:18px;">{g_impact}</div>

  <div class="two-col">
    <div>
      <div class="ct">Mapa po Scénáři 1
        <span class="badge b-keep">● Zachovat</span>
        <span class="badge b-close">● Zavřít</span>
      </div>
      {g_mapsc1}
    </div>
    <div>
      <div class="ct">Pobočky navržené k uzavření — Scénář 1 ({n_close})</div>
      <div class="scroll-tbl">
        <table>
          <thead><tr>
            <th>Pobočka</th><th>Město</th><th>Formát</th>
            <th style="text-align:center;">IR Q</th>
            <th style="text-align:right;">Prim. klientů</th>
          </tr></thead>
          <tbody>{sc1_closed_rows}</tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ══ Citlivostní analýza ════════════════════════════════════════════════ -->
<div class="card">
  <div class="ct">⚖️ Citlivostní analýza — vliv optimalizační strategie na zachované metriky</div>
  <div class="note">
    Ukazuje, kolik % výnosů / klientů by bylo zachováno při různém počtu ponechaných poboček
    a třech optimalizačních strategiích.
    <strong>Plná čára = výnosy · Tečkovaná = klienti.</strong>
    Čím více křivka stoupá nad diagonálu, tím efektivnější strategie.
  </div>
  {g_sensit}
  <div style="display:flex;gap:24px;margin-top:10px;flex-wrap:wrap;font-size:0.75rem;color:#374151;">
    <div><span style="color:#7c3aed;font-weight:700;">■</span>
      <strong> Dostupnost</strong> — zachovává geograficky nejdůležitější (izolované) pobočky</div>
    <div><span style="color:#2563eb;font-weight:700;">■</span>
      <strong> Výnosy</strong> — zachovává nejvýnosnější pobočky</div>
    <div><span style="color:#16a34a;font-weight:700;">■</span>
      <strong> Klienti</strong> — zachovává pobočky s největší klientskou základnou</div>
    <div><span style="color:#d97706;font-weight:700;">■</span>
      <strong> Vyvážená</strong> — rovné váhy všech tří dimenzí</div>
  </div>
</div>

<!-- ══ Maximalizace všech parametrů ═══════════════════════════════════════ -->
<div class="card" style="border-top:3px solid #7c3aed;">
  <div class="sc-banner" style="border-left-color:#7c3aed;">
    <h3 style="color:#7c3aed;">🎯 Maximalizace všech parametrů — doporučené uzavření ({n_opt_close} poboček)</h3>
    <p>
      Složené skóre = průměr normalizované dostupnosti + výnosů + počtu klientů.
      Pobočky s nejnižším součtovým skóre jsou kandidáty na uzavření.
      Velikost bodu = počet bankéřů.
    </p>
  </div>

  {g_opt}

  <div class="two-col" style="margin-top:18px;">
    <div>
      <div class="ct">Navrženo k uzavření ({n_opt_close} poboček)</div>
      <div class="scroll-tbl">
        <table>
          <thead><tr>
            <th>Pobočka</th><th>Město</th><th>Formát</th>
            <th style="text-align:right;">Klientů</th>
            <th style="text-align:right;">Výnosy</th>
            <th style="text-align:right;">Dostupnost</th>
          </tr></thead>
          <tbody>{opt_close_rows}</tbody>
        </table>
      </div>
    </div>
    <div class="info-box">
      <div style="font-size:0.71rem;font-weight:700;color:#7c3aed;text-transform:uppercase;
                  letter-spacing:.5px;margin-bottom:12px;">Co tento scénář optimalizuje</div>
      <div style="margin-bottom:8px;">
        ✅ <strong>Dostupnost sítě</strong> — zachovány geograficky izolované pobočky,
        kde zákazníci nemají blízkou alternativu
      </div>
      <div style="margin-bottom:8px;">
        ✅ <strong>Výnosy</strong> — zachovány nejvýnosnější pobočky
      </div>
      <div style="margin-bottom:8px;">
        ✅ <strong>Klientská základna</strong> — zachovány pobočky s největším portfoliem
      </div>
      <div style="margin-bottom:8px;">
        💰 <strong>Úspora na nájmech:</strong>
        <span style="color:#059669;font-weight:700;">{opt_save_rent/1e6:.1f} M Kč/rok</span>
      </div>
      <div style="margin-bottom:8px;">
        📉 <strong>Dopad na výnosy:</strong>
        ztráta <span style="color:#dc2626;font-weight:700;">
        {opt_rev_loss/total_revenue*100:.1f}%</span> výnosů sítě
      </div>
      <div>
        👤 <strong>Klientů bez pobočky:</strong>
        <span style="color:#d97706;font-weight:700;">
        {df_opt_close["PRIMARNI_KLIENTI"].sum():,.0f}</span>
        ({df_opt_close["PRIMARNI_KLIENTI"].sum()/total_clients*100:.1f}% z celku)
      </div>
    </div>
  </div>
</div>

</div><!-- /wrap -->
</body>
</html>"""

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as fh:
            fh.write(html)
        print(f'  ✅ Síťový report: {output_path}')

    return html


# ── Standalone spuštění ───────────────────────────────────────────────────────

if __name__ == '__main__':
    import os
    import sys
    import pickle

    # 1. Použij df nebo rating_status pokud jsou již načteny v tomto souboru
    _rs = None
    for _vname in ('df', 'rating_status'):
        try:
            _cand = eval(_vname)                          # noqa: S307
            if hasattr(_cand, 'empty') and not _cand.empty:
                _rs = _cand
                print(f'Používám `{_vname}` načtený v souboru ({len(_rs)} řádků).')
                break
        except NameError:
            pass

    # 2. Fallback: hledej pkl soubory
    if _rs is None:
        _pkl_candidates = [
            'rating_status.pkl',
            '../vypocet_ir_2026/report_rating_2026_staticky.pkl',
        ]
        for _p in _pkl_candidates:
            if os.path.exists(_p):
                print(f'Načítám {_p}...')
                with open(_p, 'rb') as _f:
                    _rs = pickle.load(_f)
                print(f'  → {len(_rs)} řádků načteno.')
                break

    if _rs is None:
        print('DataFrame nenalezen.')
        print('Přidej na začátek souboru:')
        print('  df = pd.read_pickle("../vypocet_ir_2026/report_rating_2026_staticky.pkl")')
        sys.exit(1)

    generate_network_analysis_report(_rs, output_path='report_network_analysis.html')
