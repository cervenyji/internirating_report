"""
CS brand report – pouze hlavní tabulka + přehled obsazení pozic.
Vizuální design striktně podle značkového manuálu České spořitelny.

Použití (standalone):
    python internirating_report_cs.py

Nebo volej generate_cs_brand_report() přímo z internirating_report.py.
"""

import os
import sys


def generate_cs_brand_report(
    df_sorted,
    dbs_date,
    output_prefix,
    fn_filterable_table,
    fn_specialiste_summary,
    df_specialiste_detail=None,
    accent="#245375",           # Stone jako výchozí dominantní barva
):
    """
    Vygeneruje HTML report ve stylu CS brand manuálu.

    df_sorted               – DataFrame s pobočkami (z generate_report)
    dbs_date                – řetězec data, např. "2026-02-10"
    output_prefix           – prefix výstupního souboru
    fn_filterable_table     – reference na generate_filterable_table()
    fn_specialiste_summary  – reference na generate_specialiste_summary_table()
    df_specialiste_detail   – DataFrame se zaměstnanci (může být None)
    accent                  – hex barva dominantního akcentu
    """

    # Injektuj df_specialiste_detail do globálního prostoru funkce,
    # protože generate_specialiste_summary_table() čte globals()
    if df_specialiste_detail is not None:
        fn_specialiste_summary.__globals__['df_specialiste_detail'] = df_specialiste_detail

    # ── Filtrovatelná tabulka ────────────────────────────────────────────────
    html_table = fn_filterable_table(df_sorted, "cs_main_table")

    # ── Přehled obsazení pozic ───────────────────────────────────────────────
    branch_codes = list(df_sorted['BRANCH_CODE'].dropna().unique())
    html_specialiste = fn_specialiste_summary(branch_codes, title="👷 Přehled obsazení pozic")

    if html_specialiste:
        specialiste_section = f"""
<section class="cs-section">
  <details class="cs-collapsible" open>
    <summary>👷 Přehled obsazení pozic</summary>
    <div class="cs-collapsible-body">
      {html_specialiste}
    </div>
  </details>
</section>"""
    else:
        specialiste_section = ""

    n_pobocek = len(df_sorted)

    # ── CS brand paleta ──────────────────────────────────────────────────────
    css = f"""
:root {{
  --cs-accent:    {accent};
  --cs-bright-blue: #2870ED;
  --cs-teal:      #02A3A4;
  --cs-forest:    #028661;
  --cs-apple:     #0CB43F;
  --cs-orange:    #FF6130;
  --cs-pink:      #EB4C79;
  --cs-aubergine: #721C7A;
  --cs-stone:     #245375;
  --cs-white:     #FFFFFF;
  --cs-anthracite:#202020;
  --cs-gray-dark: #4A4A4A;
  --cs-gray:      #9B9B9B;
  --cs-gray-light:#E6E6E6;
  --cs-bg:        #F4F6FA;
}}

*,*::before,*::after {{ box-sizing:border-box; margin:0; padding:0; }}

body {{
  font-family:"Inter",Arial,"Helvetica Neue",sans-serif;
  background:var(--cs-bg);
  color:var(--cs-anthracite);
  line-height:1.5;
}}

/* ── Hero ─────────────────────────────────────────────────────────────── */
.cs-hero {{
  position:relative;
  background:var(--cs-accent);
  color:#fff;
  padding:44px 56px 32px;
  overflow:hidden;
}}
.cs-hero::before {{
  content:"";
  position:absolute;
  top:0; left:0;
  width:64px; height:4px;
  background:rgba(255,255,255,.70);
}}
.cs-hero::after {{
  content:"";
  position:absolute;
  top:0; left:0;
  width:4px; height:64px;
  background:rgba(255,255,255,.70);
}}
.cs-hero-topline {{
  font-size:0.68rem;
  text-transform:uppercase;
  letter-spacing:2.8px;
  opacity:.72;
  font-weight:600;
  margin-bottom:12px;
}}
.cs-hero h1 {{
  font-size:2.4rem;
  font-weight:800;
  line-height:1.1;
  letter-spacing:-.3px;
  margin-bottom:10px;
  max-width:800px;
}}
.cs-hero-meta {{
  font-size:0.86rem;
  opacity:.82;
}}
.cs-hero-meta strong {{
  font-weight:700;
}}

/* ── Wrapper ──────────────────────────────────────────────────────────── */
.cs-main {{
  max-width:1700px;
  margin:0 auto;
  padding:36px 40px;
}}

/* ── Section headings ─────────────────────────────────────────────────── */
.cs-section {{
  margin-bottom:36px;
}}
.cs-section-title {{
  font-size:1.05rem;
  font-weight:700;
  color:var(--cs-accent);
  border-bottom:2px solid var(--cs-accent);
  padding-bottom:6px;
  margin-bottom:20px;
  text-transform:uppercase;
  letter-spacing:.6px;
}}

/* ── Collapsible ──────────────────────────────────────────────────────── */
.cs-collapsible {{
  border:1px solid var(--cs-gray-light);
  border-radius:10px;
  overflow:hidden;
  background:#fff;
}}
.cs-collapsible > summary {{
  background:color-mix(in srgb,var(--cs-accent) 8%,#fff);
  color:var(--cs-accent);
  font-weight:700;
  font-size:0.97rem;
  padding:14px 20px;
  cursor:pointer;
  list-style:none;
  display:flex;
  align-items:center;
  gap:8px;
  border-bottom:1px solid var(--cs-gray-light);
  user-select:none;
}}
.cs-collapsible > summary:hover {{
  background:color-mix(in srgb,var(--cs-accent) 14%,#fff);
}}
.cs-collapsible > summary::marker,
.cs-collapsible > summary::-webkit-details-marker {{ display:none; }}
.cs-collapsible > summary::before {{
  content:"▶";
  font-size:.65rem;
  transition:transform .2s;
}}
.cs-collapsible[open] > summary::before {{
  transform:rotate(90deg);
}}
.cs-collapsible-body {{
  padding:20px;
}}

/* ── Footer ───────────────────────────────────────────────────────────── */
.cs-footer {{
  position:relative;
  background:var(--cs-accent);
  color:#fff;
  padding:24px 56px;
  display:flex;
  justify-content:space-between;
  align-items:center;
  overflow:hidden;
  margin-top:8px;
}}
.cs-footer::before {{
  content:"";
  position:absolute;
  top:0; left:0;
  width:64px; height:4px;
  background:rgba(255,255,255,.70);
}}
.cs-footer::after {{
  content:"";
  position:absolute;
  top:0; left:0;
  width:4px; height:64px;
  background:rgba(255,255,255,.70);
}}
.cs-footer-left {{
  font-size:0.83rem;
  opacity:.85;
}}
.cs-footer-right {{
  text-align:right;
}}
.cs-footer-claim {{
  font-size:1.1rem;
  font-weight:800;
  letter-spacing:.4px;
}}
.cs-footer-hash {{
  font-size:0.75rem;
  font-weight:600;
  opacity:.72;
}}

/* ── Embed: přepiš barvy generované tabulky do CS stylu ──────────────── */
.cs-main .ft-toolbar {{ border-radius:8px; }}
.cs-main table thead th {{
  background:var(--cs-accent) !important;
  color:#fff !important;
  border-color:rgba(255,255,255,.2) !important;
}}
.cs-main details.collapsible-section > summary {{
  background:color-mix(in srgb,var(--cs-accent) 9%,#fff) !important;
  color:var(--cs-accent) !important;
  border-color:var(--cs-gray-light) !important;
}}
.cs-main details.collapsible-section > summary:hover {{
  background:color-mix(in srgb,var(--cs-accent) 16%,#fff) !important;
}}

@media(max-width:768px) {{
  .cs-hero {{ padding:28px 24px 20px; }}
  .cs-hero h1 {{ font-size:1.6rem; }}
  .cs-main {{ padding:20px 16px; }}
  .cs-footer {{ flex-direction:column; gap:10px; text-align:center; padding:20px 24px; }}
  .cs-footer-right {{ text-align:center; }}
}}
"""

    html = f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Rating poboček 2026 — Přehled · Česká spořitelna</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
{css}
</style>
</head>
<body>

<header class="cs-hero">
  <div class="cs-hero-topline">Česká spořitelna &mdash; Rating poboček</div>
  <h1>Přehled ratingů poboček 2026</h1>
  <p class="cs-hero-meta">
    <strong>{n_pobocek}</strong> poboček v přehledu
    &nbsp;&middot;&nbsp; DBS ke dni <strong>{dbs_date}</strong>
  </p>
</header>

<main class="cs-main">

  <!-- Hlavní tabulka -->
  <section class="cs-section">
    <div class="cs-section-title">Hlavní přehled poboček</div>
    {html_table}
  </section>

  <!-- Obsazení pozic -->
  {specialiste_section}

</main>

<footer class="cs-footer">
  <div class="cs-footer-left">
    Česká spořitelna &nbsp;&middot;&nbsp; Rating poboček 2026
  </div>
  <div class="cs-footer-right">
    <div class="cs-footer-claim">Ať se daří</div>
    <div class="cs-footer-hash">#silnější</div>
  </div>
</footer>

</body>
</html>"""

    filename = f"{output_prefix}_cs.html"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ CS brand report uložen: {filename}")
    return filename


# ── Standalone spuštění ──────────────────────────────────────────────────────
if __name__ == "__main__":
    import pickle
    import pandas as pd

    # Hledej pkl soubor v aktuálním adresáři
    pkl_candidates = [
        "report_rating_2026_staticky.pkl",
        "report_staticky.pkl",
    ]
    pkl_path = None
    for c in pkl_candidates:
        if os.path.exists(c):
            pkl_path = c
            break

    if pkl_path is None:
        print("❌ Nenalezen pkl soubor. Spusť nejdříve internirating_report.py.")
        print(f"   Hledal: {pkl_candidates}")
        sys.exit(1)

    print(f"📂 Načítám data z {pkl_path} ...")
    df_sorted = pd.read_pickle(pkl_path)

    # Připoj generate_filterable_table a generate_specialiste_summary_table
    # z hlavního modulu – spustíme import s potlačením vedlejších efektů
    _orig_dir = os.getcwd()
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    if _script_dir not in sys.path:
        sys.path.insert(0, _script_dir)

    # Načíst funkce bez spuštění module-level kódu není možné přímo;
    # proto definujeme lokální stub, který importuje jen definice.
    try:
        import importlib.util as _ilu
        _spec = _ilu.spec_from_file_location(
            "ir_funcs",
            os.path.join(_script_dir, "internirating_report.py")
        )
        # Module-level kód v IR souboru se spustí při importu –
        # zachyťme výjimky způsobené chybějícími daty gracefully.
        import io as _io
        _mod = _ilu.module_from_spec(_spec)
        _old_stdout = sys.stdout
        sys.stdout = _io.StringIO()
        try:
            _spec.loader.exec_module(_mod)
        except Exception:
            pass
        finally:
            sys.stdout = _old_stdout

        _fn_table = _mod.generate_filterable_table
        _fn_spec  = _mod.generate_specialiste_summary_table

        # Zkus načíst df_specialiste_detail z pkl vedlejšího souboru
        _spec_pkl = "report_specialiste_detail.pkl"
        _df_spec = None
        if os.path.exists(_spec_pkl):
            try:
                _df_spec = pd.read_pickle(_spec_pkl)
                print(f"📂 Načten specialiste detail z {_spec_pkl}")
            except Exception as e:
                print(f"⚠ Specialiste detail: {e}")

        from datetime import date as _date
        _dbs_date = getattr(_mod, "DBS_DATE", str(_date.today()))

        generate_cs_brand_report(
            df_sorted=df_sorted,
            dbs_date=_dbs_date,
            output_prefix="report_rating_2026",
            fn_filterable_table=_fn_table,
            fn_specialiste_summary=_fn_spec,
            df_specialiste_detail=_df_spec,
        )

    except Exception as e:
        print(f"❌ Chyba: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)
