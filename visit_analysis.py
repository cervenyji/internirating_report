#!/usr/bin/env python3
"""
visit_analysis.py
─────────────────
Standalone analýza návštěvnosti poboček.

Vstup  : ../in/tables/VISITS_2025.csv          (jednotlivé návštěvy)
         kpis_grouped_2026.pkl                  (FTE, název pobočky)
         ../vypocet_ir_2026/zdroje/report_od_pobocky_dbs_04_2026.xlsx
Výstup : report_navstevnost.html
"""

import os
import sys
import json
import unicodedata
import re
import subprocess
import pandas as pd
import numpy as np

try:
    import openpyxl
except ImportError:
    print("Instaluji openpyxl…")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# ─── Paths ─────────────────────────────────────────────────────────────────────
VISITS_PATH = '../in/tables/VISITS_2025.csv'
KPIS_PATH   = 'kpis_grouped_2026.pkl'
OD_PATH     = '../vypocet_ir_2026/zdroje/report_od_pobocky_dbs_04_2026.xlsx'
OUTPUT_FILE = 'report_navstevnost.html'

# ─── Visit type definitions ─────────────────────────────────────────────────────
VISIT_TYPES = [
    ('online',   'Online schůzky',     '#2770f0', '#eef4ff'),
    ('fyzicka',  'Fyzické schůzky',    '#45b065', '#eafaf1'),
    ('bezhot',   'Bezhot. walkin',     '#e07020', '#fff3ea'),
    ('hotovost', 'Hotovostní walkin',  '#9b6bbf', '#f5f0ff'),
]
TYPE_KEYS  = [t[0] for t in VISIT_TYPES]
TYPE_LABEL = {t[0]: t[1] for t in VISIT_TYPES}
TYPE_COLOR = {t[0]: t[2] for t in VISIT_TYPES}
TYPE_BG    = {t[0]: t[3] for t in VISIT_TYPES}

MONTH_NAMES = ['Led','Úno','Bře','Dub','Kvě','Čvn','Čvc','Srp','Zář','Říj','Lis','Pro']
WD_NAMES    = ['Po','Út','St','Čt','Pá','So','Ne']

_OD_DAYS = [
    ('PONDELI', 'PO',  'Pondělí', False),
    ('UTERY',   'UT',  'Úterý',   False),
    ('STREDA',  'ST',  'Středa',  False),
    ('CTVRTEK', 'CT',  'Čtvrtek', False),
    ('PATEK',   'PA',  'Pátek',   False),
    ('SOBOTA',  'SO',  'Sobota',  True),
    ('NEDELE',  'NE',  'Neděle',  True),
]


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _nc(col):
    """Normalize column: strip diacritics, upper, spaces→_."""
    if not isinstance(col, str):
        col = str(col)
    col = ''.join(c for c in unicodedata.normalize('NFD', col)
                  if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', '_', col.upper().strip())


def _map_type(val):
    """Map raw ATTENDANCE_TYPE to canonical key.

    Handles variants like 'ON-LINE SCHUZKA', 'FYZICKA SCHUZKA', 'WALKIN BEZHOT.'.
    """
    if not isinstance(val, str):
        return None
    # Strip diacritics + hyphens before matching
    v = ''.join(c for c in unicodedata.normalize('NFD', val.lower().replace('-', ''))
                if unicodedata.category(c) != 'Mn')
    if 'online' in v:
        return 'online'
    if 'bezhot' in v:
        return 'bezhot'
    if 'hotov' in v:
        return 'hotovost'
    if any(x in v for x in ('schuzk', 'fyzick', 'schu')):
        return 'fyzicka'
    return None


def _odv(row, key):
    """Get opening-hours cell value as stripped string."""
    return str(row.get(key, '') or '').strip()


def _od_closed(v):
    return v in ('00:00', '0:00', '', 'nan', 'None', '0')


# ─── Data loading ──────────────────────────────────────────────────────────────

def load_visits():
    print(f"📂 Načítám návštěvy: {VISITS_PATH}")
    if not os.path.exists(VISITS_PATH):
        print(f"❌ Soubor nenalezen: {VISITS_PATH}", file=sys.stderr)
        sys.exit(1)
    df = pd.read_csv(VISITS_PATH, low_memory=False)
    df.columns = [_nc(c) for c in df.columns]
    print(f"   {len(df):,} řádků · sloupce: {list(df.columns)}")
    return df


def load_kpis():
    """Load kpis pkl → dict branch_id (int) → {name, fte}."""
    result = {}
    if not os.path.exists(KPIS_PATH):
        print(f"⚠️  kpis pkl nenalezen: {KPIS_PATH}")
        return result
    try:
        kp = pd.read_pickle(KPIS_PATH)
        kp.columns = [_nc(c) for c in kp.columns]
        id_c  = next((c for c in ['POBOCKA_ID','BRANCH_CODE','ID_POBOCKY'] if c in kp.columns), None)
        nm_c  = next((c for c in ['POBOCKA_NAZEV','BRANCH_NAME','NAZEV','POBOCKA'] if c in kp.columns), None)
        fte_c = next((c for c in ['FTE','POCET_BANKERU','BANKERI'] if c in kp.columns), None)
        if id_c is None:
            print("⚠️  kpis: sloupec s ID pobočky nenalezen")
            return result
        for _, row in kp.iterrows():
            bid = pd.to_numeric(row[id_c], errors='coerce')
            if pd.isna(bid):
                continue
            bid = int(bid)
            result[bid] = {
                'name': str(row[nm_c]) if nm_c else f"Pobočka {bid}",
                'fte':  float(row[fte_c]) if fte_c and pd.notna(row[fte_c]) else None,
            }
        print(f"   kpis: {len(result)} poboček · FTE={'ano' if fte_c else 'ne'} · název={'ano' if nm_c else 'ne'}")
    except Exception as e:
        print(f"⚠️  Chyba při načítání kpis: {e}")
    return result


def load_oteviraci_doba():
    """Load opening hours xlsx → dict branch_id (int) → opening-hours payload."""
    result = {}
    if not os.path.exists(OD_PATH):
        print(f"⚠️  Otevírací doba nenalezena: {OD_PATH}")
        return result
    try:
        od = pd.read_excel(OD_PATH, dtype=str)
        if 'KOD_POBOCKY' not in od.columns:
            print("⚠️  Otevírací doba: chybí sloupec KOD_POBOCKY")
            return result

        for _, row in od.iterrows():
            bid = pd.to_numeric(row.get('KOD_POBOCKY', ''), errors='coerce')
            if pd.isna(bid):
                continue
            bid = int(bid)

            days = []
            is_vikend = False
            for dn, ds, dcz, wknd in _OD_DAYS:
                tot = _odv(row, ds)
                closed = _od_closed(tot)
                dop_od = _odv(row, f'{dn}_DOP._OD')
                dop_do = _odv(row, f'{dn}_DOP._DO')
                odp_od = _odv(row, f'{dn}_ODP._OD')
                odp_do = _odv(row, f'{dn}_ODP._DO')

                if wknd and not closed:
                    is_vikend = True

                dop_s = f"{dop_od}–{dop_do}" if not _od_closed(dop_od) else ''
                odp_s = f"{odp_od}–{odp_do}" if not _od_closed(odp_od) else ''
                days.append({
                    'lbl':    dcz,
                    'wknd':   wknd,
                    'closed': closed,
                    'dop':    dop_s,
                    'odp':    odp_s,
                    'tot':    tot if not closed else '',
                })

            # Weekly hours
            ph_raw = str(row.get('PH', '') or '').replace(',', '.')
            try:
                ph = float(ph_raw)
            except (ValueError, TypeError):
                ph = 0.0

            result[bid] = {
                'is_vikend': is_vikend,
                'ph_tyden':  ph,
                'od_days':   days,
            }
        print(f"   Otevírací doba: {len(result)} poboček")
    except Exception as e:
        print(f"⚠️  Chyba při načítání otevírací doby: {e}")
    return result


# ─── Aggregation ───────────────────────────────────────────────────────────────

def build_data(df, kpis, oteviraci):
    """Pre-compute per-branch stats → dict branch_id_str → payload."""

    bid_c   = next((c for c in ['BRANCH_ID','BRANCH_CODE','POBOCKA'] if c in df.columns), None)
    bname_c = next((c for c in ['BRANCH_NAME','POBOCKA_NAZEV','BRANCH_NAZEV'] if c in df.columns), None)
    att_c   = next((c for c in ['ATTENDANCE_TYPE','VISIT_TYPE','TYP_NAVSTEVY'] if c in df.columns), None)
    date_c  = next((c for c in ['VISIT_DATE','DATE','DATUM'] if c in df.columns), None)
    time_c  = next((c for c in ['VISIT_TIME','TIME','CAS'] if c in df.columns), None)

    if bid_c is None:
        print("❌ Sloupec s ID pobočky nenalezen (BRANCH_ID / BRANCH_CODE / POBOCKA)")
        sys.exit(1)

    df = df.copy()
    df[bid_c] = pd.to_numeric(df[bid_c], errors='coerce')
    df = df.dropna(subset=[bid_c])
    df[bid_c] = df[bid_c].astype(int)

    # Visit type
    if att_c:
        df['_t'] = df[att_c].apply(_map_type)
        print(f"   Rozpoznané typy: {df['_t'].value_counts().to_dict()}")
    else:
        df['_t'] = None
        print("⚠️  ATTENDANCE_TYPE sloupec nenalezen — grafy dle typu nebudou dostupné")

    # Date
    has_date = date_c is not None
    if has_date:
        df['_dt']  = pd.to_datetime(df[date_c], errors='coerce')
        df['_mon'] = df['_dt'].dt.month
        df['_wd']  = df['_dt'].dt.weekday
    else:
        df['_mon'] = None; df['_wd'] = None

    # Hour
    has_time = time_c is not None
    if has_time:
        df['_hr'] = pd.to_numeric(
            df[time_c].astype(str).str.split(':').str[0], errors='coerce')
    else:
        df['_hr'] = None

    result  = {}
    branches = sorted(df[bid_c].unique())
    print(f"   {len(branches)} poboček — počítám statistiky...")

    for bid in branches:
        vb  = df[df[bid_c] == bid]
        kpi = kpis.get(bid, {})
        od  = oteviraci.get(bid, {})

        # Name
        if bname_c and bname_c in vb.columns:
            name = str(vb[bname_c].iloc[0])
        elif kpi.get('name') and not kpi['name'].startswith('Pobočka'):
            name = kpi['name']
        else:
            name = kpi.get('name') or f"Pobočka {bid}"

        total   = len(vb)
        by_type = {k: int((vb['_t'] == k).sum()) for k in TYPE_KEYS}
        unknown = int(vb['_t'].isna().sum())

        by_month   = {k: [0]*12 for k in TYPE_KEYS}
        by_weekday = {k: [0]*7  for k in TYPE_KEYS}
        by_hour    = {k: [0]*24 for k in TYPE_KEYS}

        if has_date:
            for k in TYPE_KEYS:
                sub = vb[vb['_t'] == k]
                if not sub.empty:
                    m  = sub['_mon'].value_counts().reindex(range(1,13), fill_value=0)
                    wd = sub['_wd'].value_counts().reindex(range(7),     fill_value=0)
                    by_month[k]   = [int(v) for v in m]
                    by_weekday[k] = [int(v) for v in wd]

        n_days = 1
        if has_time and has_date:
            n_days = max(int(vb['_dt'].dt.date.nunique()), 1)
            for k in TYPE_KEYS:
                sub = vb[vb['_t'] == k]
                if not sub.empty:
                    h  = sub['_hr'].dropna().astype(int)
                    hc = h.value_counts().reindex(range(24), fill_value=0)
                    by_hour[k] = [round(float(v)/n_days, 2) for v in hc]

        result[str(bid)] = {
            'name':       name,
            'fte':        kpi.get('fte'),
            'total':      total,
            'by_type':    by_type,
            'unknown':    unknown,
            'by_month':   by_month,
            'by_weekday': by_weekday,
            'by_hour':    by_hour,
            'has_time':   has_time and has_date,
            'n_days':     n_days,
            # Opening hours
            'is_vikend':  od.get('is_vikend', False),
            'ph_tyden':   od.get('ph_tyden', 0.0),
            'od_days':    od.get('od_days', []),
        }

    order = sorted(result.keys(), key=lambda x: result[x]['name'])
    return result, order, att_c is not None


# ─── HTML / JS rendering ────────────────────────────────────────────────────────

def render_html(data, order, has_type_col):
    data_json  = json.dumps(data,  ensure_ascii=False, separators=(',',':'))
    order_json = json.dumps(order, ensure_ascii=False)
    types_json = json.dumps(
        [{'key':k,'label':TYPE_LABEL[k],'color':TYPE_COLOR[k],'bg':TYPE_BG[k]}
         for k in TYPE_KEYS],
        ensure_ascii=False
    )
    no_type_warn = ('' if has_type_col else
        '<div class="warn">⚠️ Sloupec <code>ATTENDANCE_TYPE</code> nenalezen — '
        'rozlišení typů není dostupné.</div>')

    return f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Analýza návštěvnosti 2025</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',system-ui,sans-serif;background:#f4f6fa;color:#222;font-size:15px;}}
.wrap{{max-width:1080px;margin:0 auto;padding:20px 16px;}}
h1{{font-size:1.4rem;font-weight:800;color:#1a2540;margin-bottom:3px;}}
.subtitle{{font-size:0.82rem;color:#999;margin-bottom:20px;}}
.warn{{background:#fff8e1;border:1px solid #ffe082;border-radius:8px;
       padding:10px 14px;font-size:0.82rem;color:#795548;margin-bottom:16px;}}
/* Search */
.search-wrap{{position:relative;margin-bottom:20px;}}
.search-wrap input{{
  width:100%;padding:11px 14px 11px 40px;border:1.5px solid #dde3ef;
  border-radius:10px;font-size:16px;background:#fff;outline:none;
  transition:border .15s;-webkit-tap-highlight-color:transparent;
}}
.search-wrap input:focus{{border-color:#2770f0;box-shadow:0 0 0 3px rgba(39,112,240,.12);}}
.search-icon{{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#aaa;pointer-events:none;}}
.branch-list{{
  display:none;position:absolute;top:calc(100% + 4px);left:0;right:0;
  background:#fff;border:1.5px solid #dde3ef;border-radius:10px;
  max-height:260px;overflow-y:auto;z-index:100;box-shadow:0 6px 24px rgba(0,0,0,.1);
  -webkit-overflow-scrolling:touch;
}}
.branch-list.open{{display:block;}}
.branch-item{{padding:9px 14px;cursor:pointer;font-size:.9rem;border-bottom:1px solid #f0f0f0;}}
.branch-item:last-child{{border-bottom:none;}}
.branch-item:hover,.branch-item.sel{{background:#eef4ff;color:#2770f0;font-weight:600;}}
/* Cards */
.kpi-row{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;}}
.kpi-card{{flex:1;min-width:140px;background:#fff;border-radius:10px;
           padding:12px 14px;border:1.5px solid #eee;}}
.kpi-lbl{{font-size:.7rem;font-weight:700;text-transform:uppercase;
          letter-spacing:.5px;color:#888;margin-bottom:3px;}}
.kpi-val{{font-size:1.45rem;font-weight:800;line-height:1.1;}}
.kpi-sub{{font-size:.75rem;color:#aaa;margin-top:2px;}}
/* Sections */
.sec{{background:#fff;border-radius:12px;border:1.5px solid #eee;
      padding:16px;margin-bottom:14px;}}
.sec-t{{font-size:.78rem;font-weight:700;text-transform:uppercase;
        letter-spacing:.5px;color:#666;margin-bottom:14px;}}
/* Bar charts */
.bars{{display:flex;gap:4px;align-items:flex-end;height:120px;}}
.bar-w{{display:flex;flex-direction:column;align-items:center;gap:3px;flex:1;min-width:0;}}
.bar-stack{{display:flex;flex-direction:column-reverse;width:100%;
            border-radius:4px 4px 0 0;overflow:hidden;}}
.bar-seg{{width:100%;flex-shrink:0;}}
.bar-lbl{{font-size:.6rem;color:#888;white-space:nowrap;overflow:hidden;
          text-overflow:ellipsis;max-width:100%;text-align:center;}}
.bar-num{{font-size:.6rem;color:#555;font-weight:600;text-align:center;}}
/* Legend */
.legend{{display:flex;flex-wrap:wrap;gap:10px;margin-top:12px;}}
.leg-i{{display:flex;align-items:center;gap:5px;font-size:.76rem;color:#555;}}
.leg-d{{width:10px;height:10px;border-radius:3px;flex-shrink:0;}}
/* Opening hours table */
.od-table{{width:100%;border-collapse:collapse;font-size:.84rem;}}
.od-table td,.od-table th{{padding:6px 10px;}}
.od-table thead th{{font-size:.7rem;font-weight:700;text-transform:uppercase;
                    letter-spacing:.4px;color:#888;text-align:left;
                    border-bottom:1px solid #eee;}}
.od-table tbody tr:nth-child(odd){{background:#fafafa;}}
.od-wknd{{background:#fffbf0!important;}}
.od-closed{{color:#ccc;font-style:italic;}}
.badge{{display:inline-block;border-radius:12px;padding:3px 10px;
        font-size:.72rem;font-weight:700;margin-right:6px;}}
.placeholder{{text-align:center;color:#bbb;font-size:.85rem;padding:40px 0;}}
@media(max-width:600px){{
  .kpi-val{{font-size:1.2rem;}}
  .kpi-card{{min-width:120px;}}
}}
</style>
</head>
<body>
<div class="wrap">
  <h1>Analýza návštěvnosti poboček 2025</h1>
  <div class="subtitle">Zdroj: {VISITS_PATH}</div>
  {no_type_warn}

  <div class="search-wrap" id="sw">
    <svg class="search-icon" width="16" height="16" viewBox="0 0 24 24" fill="none"
         stroke="currentColor" stroke-width="2">
      <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
    </svg>
    <input type="text" id="si" placeholder="Vyhledat pobočku…"
           autocomplete="off" autocorrect="off" spellcheck="false">
    <div class="branch-list" id="bl"></div>
  </div>

  <div id="mc"><div class="placeholder">← Vyberte pobočku výše</div></div>
</div>

<script>
const DATA  = {data_json};
const ORDER = {order_json};
const TYPES = {types_json};
const WD    = ['Po','Út','St','Čt','Pá','So','Ne'];
const MON   = ['Led','Úno','Bře','Dub','Kvě','Čvn','Čvc','Srp','Zář','Říj','Lis','Pro'];

let cur = null;
const si = document.getElementById('si');
const bl = document.getElementById('bl');
const sw = document.getElementById('sw');

// ── Dropdown ─────────────────────────────────────────────────────────────────
function renderList(q) {{
  q = q.toLowerCase();
  const hits = ORDER.filter(id => {{
    const d = DATA[id];
    return d.name.toLowerCase().includes(q) || id.includes(q);
  }}).slice(0, 100);
  bl.innerHTML = hits.map(id => {{
    const d = DATA[id];
    return `<div class="branch-item${{id===cur?' sel':''}}" data-id="${{id}}">
      ${{d.name}} <span style="color:#bbb;font-size:.78em">#${{id}}</span></div>`;
  }}).join('');
  bl.classList.toggle('open', hits.length > 0);
}}
si.addEventListener('input', () => renderList(si.value));
si.addEventListener('focus', () => renderList(si.value));
document.addEventListener('click', e => {{
  if (!sw.contains(e.target)) bl.classList.remove('open');
}});
bl.addEventListener('click', e => {{
  const it = e.target.closest('.branch-item');
  if (!it) return;
  si.value = DATA[it.dataset.id].name;
  bl.classList.remove('open');
  render(it.dataset.id);
}});

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmtN(n) {{
  if (n >= 1000) return (n/1000).toFixed(1).replace('.0','') + 'k';
  return String(n);
}}
function fmtF(f) {{
  return f >= 10 ? Math.round(f).toLocaleString('cs') : f.toFixed(1);
}}

function stackedBars(labels, typeArr, maxH) {{
  // typeArr: [{{label,color,values[]}}]
  const totals = labels.map((_,i) => typeArr.reduce((s,t) => s + (t.values[i]||0), 0));
  const maxVal = Math.max(...totals, 1);
  return labels.map((lbl, i) => {{
    const tot = totals[i];
    const pctH = (tot/maxVal*maxH).toFixed(1);
    const segs = typeArr
      .filter(t => (t.values[i]||0) > 0)
      .map(t => {{
        const h = (t.values[i]/maxVal*maxH).toFixed(1);
        return `<div class="bar-seg" style="height:${{h}}px;background:${{t.color}};"
                     title="${{t.label}}: ${{(t.values[i]||0).toLocaleString('cs')}}"></div>`;
      }}).join('');
    return `<div class="bar-w">
      <div class="bar-num">${{tot>0?fmtN(tot):''}}</div>
      <div class="bar-stack" style="height:${{pctH}}px;max-height:${{maxH}}px;">${{segs}}</div>
      <div class="bar-lbl">${{lbl}}</div>
    </div>`;
  }}).join('');
}}

function legend() {{
  return `<div class="legend">${{TYPES.map(t =>
    `<div class="leg-i"><div class="leg-d" style="background:${{t.color}}"></div>${{t.label}}</div>`
  ).join('')}}</div>`;
}}

function odTable(d) {{
  if (!d.od_days || d.od_days.length === 0) return '';
  const phLine = d.ph_tyden > 0
    ? `<div style="font-size:.78rem;color:#888;margin-top:8px;">
         Týdenní ot. hodiny: <b style="color:#2770f0;">${{d.ph_tyden}}h</b>
       </div>` : '';
  const rows = d.od_days.map(day => {{
    if (day.closed) return `
      <tr class="${{day.wknd?'od-wknd':''}}">
        <td style="font-weight:600;color:#ccc;">${{day.wknd?'🌅 ':''}}${{day.lbl}}</td>
        <td colspan="3" class="od-closed" style="text-align:center;">Zavřeno</td>
      </tr>`;
    return `
      <tr class="${{day.wknd?'od-wknd':''}}">
        <td style="font-weight:600;color:${{day.wknd?'#d97706':'#1a2a4a'}};white-space:nowrap;">
          ${{day.wknd?'🌅 ':''}}${{day.lbl}}</td>
        <td style="color:#555;text-align:center;">${{day.dop||'—'}}</td>
        <td style="color:#555;text-align:center;">${{day.odp||'—'}}</td>
        <td style="font-weight:700;color:#2770f0;text-align:center;">
          ${{day.tot?day.tot+'h':''}}</td>
      </tr>`;
  }}).join('');
  return `
<div class="sec">
  <div class="sec-t">🕐 Otevírací doba</div>
  ${{d.is_vikend
    ? '<span class="badge" style="background:#d97706;color:#fff;">🌅 Víkendová pobočka</span>'
    : ''}}
  <table class="od-table" style="margin-top:${{d.is_vikend?'10px':'0'}};">
    <thead><tr>
      <th>Den</th><th>Dopoledne</th><th>Odpoledne</th><th>Celkem</th>
    </tr></thead>
    <tbody>${{rows}}</tbody>
  </table>
  ${{phLine}}
</div>`;
}}

// ── Main render ───────────────────────────────────────────────────────────────
function render(id) {{
  cur = id;
  const d = DATA[id];
  const total = d.total;

  // KPI row
  const fte = d.fte != null
    ? `<div class="kpi-card"><div class="kpi-lbl">Bankéři (FTE)</div>
         <div class="kpi-val">${{d.fte % 1 === 0 ? d.fte : d.fte.toFixed(1)}}</div>
         <div class="kpi-sub">přepočtené úvazky</div></div>` : '';

  const phCard = d.ph_tyden > 0
    ? `<div class="kpi-card"><div class="kpi-lbl">Ot. hod. / týden</div>
         <div class="kpi-val">${{d.ph_tyden}}h</div>
         <div class="kpi-sub">${{d.is_vikend ? '🌅 víkendová' : 'pracovní dny'}}</div></div>` : '';

  const typeCards = TYPES.map(t => {{
    const v = d.by_type[t.key] || 0;
    const pct = total > 0 ? (v/total*100).toFixed(1) : '0.0';
    return `<div class="kpi-card" style="border-color:${{t.color}}30;">
      <div class="kpi-lbl" style="color:${{t.color}};">${{t.label}}</div>
      <div class="kpi-val" style="color:${{t.color}};">${{v.toLocaleString('cs')}}</div>
      <div class="kpi-sub">${{pct}} % z celku</div>
    </div>`;
  }}).join('');

  const totalCard = `<div class="kpi-card">
    <div class="kpi-lbl">Celkem návštěv</div>
    <div class="kpi-val">${{total.toLocaleString('cs')}}</div>
    <div class="kpi-sub">rok 2025</div>
  </div>`;

  // Month chart
  const tArr = TYPES.map(t => ({{label:t.label, color:t.color,
    values: d.by_month[t.key] || Array(12).fill(0)}}));
  const monthBars = stackedBars(MON, tArr, 110);

  // Weekday chart — So/Ne only for weekend branches
  const wdLabels = d.is_vikend ? WD : WD.slice(0,5);
  const wdArr = TYPES.map(t => ({{label:t.label, color:t.color,
    values: (d.by_weekday[t.key] || Array(7).fill(0)).slice(0, wdLabels.length)}}));
  const wdBars = stackedBars(wdLabels, wdArr, 100);
  const wdNote = !d.is_vikend
    ? '<div style="font-size:.72rem;color:#bbb;margin-top:6px;">So/Ne skryty — nevíkendová pobočka</div>'
    : '';

  // Hourly chart
  let hourSec = '';
  if (d.has_time) {{
    const hours = Array.from({{length:16}},(_,i)=>i+6);
    const hArr = TYPES.map(t => ({{label:t.label, color:t.color,
      values: hours.map(h => d.by_hour[t.key]?.[h] || 0)}}));
    hourSec = `<div class="sec">
      <div class="sec-t">⏱️ Průměrný den — návštěv na hodinu (6–21h)</div>
      <div class="bars">${{stackedBars(hours.map(h=>h+'h'), hArr, 90)}}</div>
      ${{legend()}}
    </div>`;
  }}

  const unkNote = d.unknown > 0
    ? `<div style="font-size:.73rem;color:#bbb;margin:4px 0 10px;">
         ${{d.unknown.toLocaleString('cs')}} návštěv bez určeného typu</div>` : '';

  document.getElementById('mc').innerHTML = `
<div style="font-size:1rem;font-weight:700;color:#1a2540;margin-bottom:12px;">
  ${{d.name}}
  <span style="font-size:.78rem;font-weight:400;color:#bbb;margin-left:6px;">#${{id}}</span>
</div>
<div class="kpi-row">${{totalCard}}${{fte}}${{phCard}}${{typeCards}}</div>
${{unkNote}}
<div class="sec">
  <div class="sec-t">📅 Návštěvy dle měsíce</div>
  <div class="bars">${{monthBars}}</div>
  ${{legend()}}
</div>
<div class="sec">
  <div class="sec-t">📆 Návštěvy dle dne v týdnu</div>
  <div class="bars">${{wdBars}}</div>
  ${{legend()}}
  ${{wdNote}}
</div>
${{hourSec}}
${{odTable(d)}}
`;
}}
</script>
</body>
</html>"""


# ─── Generování ────────────────────────────────────────────────────────────────

_df_visits = load_visits()
_kpis      = load_kpis()
_oteviraci = load_oteviraci_doba()

_visit_data, _order, _has_type = build_data(_df_visits, _kpis, _oteviraci)
_html = render_html(_visit_data, _order, _has_type)

with open(OUTPUT_FILE, 'w', encoding='utf-8') as _f:
    _f.write(_html)

print(f"\n✅ Report uložen: {OUTPUT_FILE}")
print(f"   {len(_visit_data)} poboček · "
      f"{'s rozlišením typů' if _has_type else 'bez typů'} · "
      f"kpis={'ano' if _kpis else 'ne'} · "
      f"ot.doba={'ano' if _oteviraci else 'ne'}")
