#!/usr/bin/env python3
"""
visit_analysis.py
─────────────────
Standalone analýza návštěvnosti poboček.

Vstup  : ../in/tables/VISITS_2025.csv  (jednotlivé návštěvy)
Výstup : report_navstevnost.html  (self-contained HTML s JS grafy)

Rozlišuje 4 typy návštěv (dle ATTENDANCE_TYPE):
  • Online schůzky        – 'online'
  • Fyzické schůzky       – 'schuzk' / 'fyzick' (ne online)
  • Bezhotovostní walkin  – 'bezhot'
  • Hotovostní walkin     – 'hotov' (ne bezhotov)
"""

import os
import sys
import json
import unicodedata
import re
import pandas as pd
import numpy as np

# ─── Paths ─────────────────────────────────────────────────────────────────────
VISITS_PATH = '../in/tables/VISITS_2025.csv'
KPIS_PATH   = 'kpis_grouped_2026.pkl'
OUTPUT_FILE = 'report_navstevnost.html'

# ─── Visit type definitions ─────────────────────────────────────────────────────
VISIT_TYPES = [
    ('online',    'Online schůzky',     '#2770f0', '#eef4ff'),
    ('fyzicka',   'Fyzické schůzky',    '#45b065', '#eafaf1'),
    ('bezhot',    'Bezhot. walkin',     '#e07020', '#fff3ea'),
    ('hotovost',  'Hotovostní walkin',  '#9b6bbf', '#f5f0ff'),
]
TYPE_KEYS   = [t[0] for t in VISIT_TYPES]
TYPE_LABEL  = {t[0]: t[1] for t in VISIT_TYPES}
TYPE_COLOR  = {t[0]: t[2] for t in VISIT_TYPES}
TYPE_BG     = {t[0]: t[3] for t in VISIT_TYPES}

MONTH_NAMES = ['Led', 'Úno', 'Bře', 'Dub', 'Kvě', 'Čvn',
               'Čvc', 'Srp', 'Zář', 'Říj', 'Lis', 'Pro']
WD_NAMES    = ['Po', 'Út', 'St', 'Čt', 'Pá', 'So', 'Ne']


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _nc(col):
    """Normalize column name: strip diacritics, upper, spaces→_."""
    if not isinstance(col, str): col = str(col)
    col = ''.join(c for c in unicodedata.normalize('NFD', col)
                  if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', '_', col.upper().strip())


def _map_type(val):
    """Map raw ATTENDANCE_TYPE string to canonical key, or None."""
    if not isinstance(val, str):
        return None
    v = ''.join(c for c in unicodedata.normalize('NFD', val.lower())
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


# ─── Data loading ──────────────────────────────────────────────────────────────

def load_visits():
    print(f"📂 Načítám: {VISITS_PATH}")
    if not os.path.exists(VISITS_PATH):
        print(f"❌ Soubor nenalezen: {VISITS_PATH}", file=sys.stderr)
        sys.exit(1)
    df = pd.read_csv(VISITS_PATH, low_memory=False)
    df.columns = [_nc(c) for c in df.columns]
    print(f"   {len(df):,} řádků | sloupce: {list(df.columns)}")
    return df


def load_branch_names():
    """Branch id→name from kpis pkl; fallback empty dict."""
    try:
        kp = pd.read_pickle(KPIS_PATH)
        kp.columns = [_nc(c) for c in kp.columns]
        id_c  = next((c for c in ['POBOCKA_ID','BRANCH_CODE','ID_POBOCKY'] if c in kp.columns), None)
        nm_c  = next((c for c in ['POBOCKA_NAZEV','BRANCH_NAME','NAZEV'] if c in kp.columns), None)
        if id_c and nm_c:
            return {int(k): str(v)
                    for k, v in zip(pd.to_numeric(kp[id_c], errors='coerce'), kp[nm_c])
                    if pd.notna(k)}
    except Exception:
        pass
    return {}


# ─── Aggregation ───────────────────────────────────────────────────────────────

def build_data(df):
    """Compute per-branch stats → dict branch_id_str → payload."""

    bid_c   = next((c for c in ['BRANCH_ID','BRANCH_CODE','POBOCKA'] if c in df.columns), None)
    bname_c = next((c for c in ['BRANCH_NAME','POBOCKA_NAZEV','BRANCH_NAZEV'] if c in df.columns), None)
    att_c   = next((c for c in ['ATTENDANCE_TYPE','VISIT_TYPE','TYP_NAVSTEVY'] if c in df.columns), None)
    date_c  = next((c for c in ['VISIT_DATE','DATE','DATUM'] if c in df.columns), None)
    time_c  = next((c for c in ['VISIT_TIME','TIME','CAS'] if c in df.columns), None)

    if bid_c is None:
        print("❌ Sloupec s ID pobočky nenalezen (hledáno: BRANCH_ID, BRANCH_CODE, POBOCKA)")
        sys.exit(1)

    df = df.copy()
    df[bid_c] = pd.to_numeric(df[bid_c], errors='coerce')
    df = df.dropna(subset=[bid_c])
    df[bid_c] = df[bid_c].astype(int)

    # Visit type
    if att_c:
        df['_t'] = df[att_c].apply(_map_type)
    else:
        df['_t'] = None

    # Date parsing
    has_date = date_c is not None
    if has_date:
        df['_dt']  = pd.to_datetime(df[date_c], errors='coerce')
        df['_mon'] = df['_dt'].dt.month
        df['_wd']  = df['_dt'].dt.weekday
    else:
        df['_mon'] = None
        df['_wd']  = None

    # Hour parsing
    has_time = time_c is not None
    if has_time:
        df['_hr'] = pd.to_numeric(
            df[time_c].astype(str).str.split(':').str[0], errors='coerce')
    else:
        df['_hr'] = None

    branch_names = load_branch_names()
    result       = {}
    branches     = sorted(df[bid_c].unique())
    print(f"   {len(branches)} poboček — počítám statistiky...")

    for bid in branches:
        vb = df[df[bid_c] == bid]

        if bname_c and bname_c in vb.columns:
            name = str(vb[bname_c].iloc[0])
        elif bid in branch_names:
            name = branch_names[bid]
        else:
            name = f"Pobočka {bid}"

        total    = len(vb)
        by_type  = {k: int((vb['_t'] == k).sum()) for k in TYPE_KEYS}
        unknown  = int(vb['_t'].isna().sum())

        by_month   = {k: [0]*12 for k in TYPE_KEYS}
        by_weekday = {k: [0]*7  for k in TYPE_KEYS}
        by_hour    = {k: [0]*24 for k in TYPE_KEYS}

        if has_date:
            for k in TYPE_KEYS:
                sub = vb[vb['_t'] == k]
                if not sub.empty:
                    m  = sub['_mon'].value_counts().reindex(range(1, 13), fill_value=0)
                    wd = sub['_wd'].value_counts().reindex(range(7),  fill_value=0)
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
                    by_hour[k] = [round(float(v) / n_days, 2) for v in hc]

        result[str(bid)] = {
            'name':      name,
            'total':     total,
            'by_type':   by_type,
            'unknown':   unknown,
            'by_month':  by_month,
            'by_weekday': by_weekday,
            'by_hour':   by_hour,
            'has_time':  has_time and has_date,
            'n_days':    n_days,
        }

    order = sorted(result.keys(), key=lambda x: result[x]['name'])
    return result, order, att_c is not None


# ─── HTML generation ────────────────────────────────────────────────────────────

def render_html(data, order, has_type_col):
    data_json  = json.dumps(data,  ensure_ascii=False, separators=(',', ':'))
    order_json = json.dumps(order, ensure_ascii=False)
    types_json = json.dumps(
        [{'key': k, 'label': TYPE_LABEL[k], 'color': TYPE_COLOR[k], 'bg': TYPE_BG[k]}
         for k in TYPE_KEYS],
        ensure_ascii=False
    )

    # Language for "no type info" warning
    no_type_note = ('' if has_type_col else
        '<div style="background:#fff8e1;border:1px solid #ffe082;border-radius:8px;'
        'padding:10px 14px;font-size:0.82rem;color:#795548;margin-bottom:16px;">'
        '⚠️ Zdroj neobsahuje sloupec <code>ATTENDANCE_TYPE</code> — '
        'grafy dle typu návštěvy nejsou dostupné.</div>')

    return f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Analýza návštěvnosti poboček 2025</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',system-ui,sans-serif;background:#f4f6fa;color:#222;font-size:15px;}}
.wrap{{max-width:1100px;margin:0 auto;padding:20px 16px;}}
h1{{font-size:1.45rem;font-weight:800;color:#1a2540;margin-bottom:4px;}}
.subtitle{{font-size:0.83rem;color:#888;margin-bottom:20px;}}
.search-wrap{{position:relative;margin-bottom:20px;}}
.search-wrap input{{
  width:100%;padding:11px 14px 11px 40px;border:1.5px solid #dde3ef;border-radius:10px;
  font-size:16px;background:#fff;outline:none;transition:border .15s;
  -webkit-tap-highlight-color:transparent;
}}
.search-wrap input:focus{{border-color:#2770f0;box-shadow:0 0 0 3px rgba(39,112,240,.12);}}
.search-wrap svg{{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#aaa;}}
.branch-list{{
  display:none;position:absolute;top:calc(100% + 4px);left:0;right:0;
  background:#fff;border:1.5px solid #dde3ef;border-radius:10px;
  max-height:260px;overflow-y:auto;z-index:100;box-shadow:0 6px 24px rgba(0,0,0,.1);
}}
.branch-list.open{{display:block;}}
.branch-item{{
  padding:9px 14px;cursor:pointer;font-size:0.9rem;border-bottom:1px solid #f0f0f0;
}}
.branch-item:last-child{{border-bottom:none;}}
.branch-item:hover,.branch-item.active{{background:#eef4ff;color:#2770f0;font-weight:600;}}
.kpi-row{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;}}
.kpi-card{{
  flex:1;min-width:160px;background:#fff;border-radius:12px;
  padding:14px 16px;border:1.5px solid #eee;
}}
.kpi-label{{font-size:0.72rem;font-weight:700;text-transform:uppercase;
            letter-spacing:.5px;color:#888;margin-bottom:4px;}}
.kpi-val{{font-size:1.6rem;font-weight:800;line-height:1.1;}}
.kpi-pct{{font-size:0.78rem;color:#999;margin-top:2px;}}
.section{{background:#fff;border-radius:12px;border:1.5px solid #eee;
          padding:16px;margin-bottom:16px;}}
.sec-title{{font-size:0.8rem;font-weight:700;text-transform:uppercase;
            letter-spacing:.5px;color:#666;margin-bottom:14px;}}
.bar-group{{display:flex;gap:4px;align-items:flex-end;height:120px;}}
.bar-wrap{{display:flex;flex-direction:column;align-items:center;gap:3px;flex:1;min-width:0;}}
.bar-stack{{display:flex;flex-direction:column-reverse;width:100%;border-radius:4px 4px 0 0;
            overflow:hidden;cursor:default;}}
.bar-seg{{width:100%;transition:height .3s ease;}}
.bar-lbl{{font-size:0.6rem;color:#888;white-space:nowrap;overflow:hidden;
          text-overflow:ellipsis;max-width:100%;text-align:center;}}
.bar-val{{font-size:0.62rem;color:#555;font-weight:600;}}
.legend{{display:flex;flex-wrap:wrap;gap:10px;margin-top:12px;}}
.leg-item{{display:flex;align-items:center;gap:5px;font-size:0.76rem;color:#555;}}
.leg-dot{{width:10px;height:10px;border-radius:3px;flex-shrink:0;}}
.placeholder{{text-align:center;color:#bbb;font-size:0.85rem;padding:30px 0;}}
@media(max-width:600px){{
  .kpi-val{{font-size:1.3rem;}}
  .kpi-card{{min-width:130px;}}
}}
</style>
</head>
<body>
<div class="wrap">
  <h1>Analýza návštěvnosti poboček 2025</h1>
  <div class="subtitle">Zdroj: {VISITS_PATH} &nbsp;·&nbsp; Typy: Online schůzky · Fyzické schůzky · Bezhot. walkin · Hotov. walkin</div>

  {no_type_note}

  <div class="search-wrap" id="searchWrap">
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
    </svg>
    <input type="text" id="searchInput" placeholder="Vyhledat pobočku…" autocomplete="off" autocorrect="off" spellcheck="false">
    <div class="branch-list" id="branchList"></div>
  </div>

  <div id="mainContent">
    <div class="placeholder">← Vyberte pobočku výše</div>
  </div>
</div>

<script>
const DATA  = {data_json};
const ORDER = {order_json};
const TYPES = {types_json};
const WD    = ['Po','Út','St','Čt','Pá','So','Ne'];
const MON   = ['Led','Úno','Bře','Dub','Kvě','Čvn','Čvc','Srp','Zář','Říj','Lis','Pro'];

let currentBranch = null;

// ── Search / dropdown ─────────────────────────────────────────────────────────
const inp    = document.getElementById('searchInput');
const list   = document.getElementById('branchList');
const wrap   = document.getElementById('searchWrap');

function renderList(filter) {{
  const q = filter.toLowerCase().trim();
  const hits = ORDER.filter(id => {{
    const d = DATA[id];
    return d.name.toLowerCase().includes(q) || id.includes(q);
  }}).slice(0, 80);
  list.innerHTML = hits.map(id => {{
    const d = DATA[id];
    const active = id === currentBranch ? ' active' : '';
    return `<div class="branch-item${{active}}" data-id="${{id}}">${{d.name}} <span style="color:#aaa;font-size:.8em">#${{id}}</span></div>`;
  }}).join('');
  hits.length ? list.classList.add('open') : list.classList.remove('open');
}}

inp.addEventListener('input', () => renderList(inp.value));
inp.addEventListener('focus', () => renderList(inp.value));
document.addEventListener('click', e => {{
  if (!wrap.contains(e.target)) list.classList.remove('open');
}});
list.addEventListener('click', e => {{
  const item = e.target.closest('.branch-item');
  if (!item) return;
  const id = item.dataset.id;
  inp.value = DATA[id].name;
  list.classList.remove('open');
  selectBranch(id);
}});

// ── Chart helpers ─────────────────────────────────────────────────────────────
function fmtN(n) {{
  if (n >= 1000) return (n/1000).toFixed(1).replace('.0','') + 'k';
  return String(n);
}}
function fmtF(f) {{
  if (f >= 100) return Math.round(f).toLocaleString('cs');
  return f.toFixed(1);
}}

function stackedBars(labels, dataByType, maxH, labelFn) {{
  // dataByType: array of {{key,color,values[]}}
  const n = labels.length;
  const totals = labels.map((_,i) => dataByType.reduce((s,t)=>s+t.values[i],0));
  const maxVal = Math.max(...totals, 1);

  return labels.map((lbl, i) => {{
    const tot = totals[i];
    const pct = tot/maxVal*100;
    const segs = dataByType.filter(t=>t.values[i]>0).map(t => {{
      const h = (t.values[i]/maxVal*maxH).toFixed(1);
      return `<div class="bar-seg" style="height:${{h}}px;background:${{t.color}};
              flex-shrink:0;" title="${{t.label}}: ${{t.values[i].toLocaleString('cs')}}"></div>`;
    }}).join('');
    return `<div class="bar-wrap">
      <div class="bar-val">${{tot>0?fmtN(tot):''}}</div>
      <div class="bar-stack" style="height:${{(pct/100*maxH).toFixed(1)}}px;max-height:${{maxH}}px;">
        ${{segs}}
      </div>
      <div class="bar-lbl" title="${{labelFn(lbl)}}">${{labelFn(lbl)}}</div>
    </div>`;
  }}).join('');
}}

function legend() {{
  return `<div class="legend">${{TYPES.map(t=>
    `<div class="leg-item"><div class="leg-dot" style="background:${{t.color}}"></div>${{t.label}}</div>`
  ).join('')}}</div>`;
}}

// ── Main render ───────────────────────────────────────────────────────────────
function selectBranch(id) {{
  currentBranch = id;
  const d = DATA[id];
  const total = d.total;

  // KPI cards
  const kpiCards = TYPES.map(t => {{
    const v = d.by_type[t.key] || 0;
    const pct = total > 0 ? (v/total*100).toFixed(1) : '0.0';
    return `<div class="kpi-card" style="border-color:${{t.color}}20;">
      <div class="kpi-label" style="color:${{t.color}};">${{t.label}}</div>
      <div class="kpi-val" style="color:${{t.color}};">${{v.toLocaleString('cs')}}</div>
      <div class="kpi-pct">${{pct}}% z celku</div>
    </div>`;
  }}).join('');

  const totalCard = `<div class="kpi-card" style="border-color:#2770f020;">
    <div class="kpi-label">Celkem návštěv</div>
    <div class="kpi-val">${{total.toLocaleString('cs')}}</div>
    <div class="kpi-pct">za rok 2025</div>
  </div>`;

  // Monthly chart
  const typeData = TYPES.map(t => ({{
    key: t.key, label: t.label, color: t.color,
    values: d.by_month[t.key] || Array(12).fill(0)
  }}));
  const monthBars = stackedBars(MON, typeData, 110, l => l);

  // Weekday chart
  const wdData = TYPES.map(t => ({{
    key: t.key, label: t.label, color: t.color,
    values: d.by_weekday[t.key] || Array(7).fill(0)
  }}));
  const wdBars = stackedBars(WD, wdData, 100, l => l);

  // Hourly chart (průměrný den)
  let hourSection = '';
  if (d.has_time) {{
    const hours = Array.from({{length:16}}, (_,i) => i+6);  // 6..21
    const hrData = TYPES.map(t => ({{
      key: t.key, label: t.label, color: t.color,
      values: hours.map(h => d.by_hour[t.key]?.[h] || 0)
    }}));
    const hrBars = stackedBars(hours, hrData, 90, h => h+'h');
    hourSection = `
<div class="section">
  <div class="sec-title">⏱️ Průměrný den (6–21 hod) — průměr na den</div>
  <div class="bar-group">${{hrBars}}</div>
  ${{legend()}}
</div>`;
  }}

  // Unknown type note
  const unkNote = d.unknown > 0
    ? `<div style="font-size:.75rem;color:#aaa;margin-top:6px;">
         ${{d.unknown.toLocaleString('cs')}} návštěv bez určeného typu</div>`
    : '';

  document.getElementById('mainContent').innerHTML = `
<div style="font-size:1.05rem;font-weight:700;color:#1a2540;margin-bottom:12px;">
  ${{d.name}} <span style="font-size:.8rem;font-weight:400;color:#aaa;">#${{id}}</span>
</div>
<div class="kpi-row">${{totalCard}}${{kpiCards}}</div>
${{unkNote}}
<div class="section">
  <div class="sec-title">📅 Vytížení dle měsíce</div>
  <div class="bar-group">${{monthBars}}</div>
  ${{legend()}}
</div>
<div class="section">
  <div class="sec-title">📆 Vytížení dle dne v týdnu</div>
  <div class="bar-group">${{wdBars}}</div>
  ${{legend()}}
</div>
${{hourSection}}
`;
}}
</script>
</body>
</html>"""


# ─── Entry point ───────────────────────────────────────────────────────────────

def main():
    df = load_visits()
    data, order, has_type = build_data(df)
    html = render_html(data, order, has_type)

    out_path = os.path.join(os.path.dirname(globals().get('__file__', '') or ''), OUTPUT_FILE) or OUTPUT_FILE
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n✅ Report uložen: {out_path}")
    print(f"   {len(data)} poboček, {'s' if has_type else 'bez'} rozlišením typu návštěvy")


if __name__ == '__main__':
    main()
