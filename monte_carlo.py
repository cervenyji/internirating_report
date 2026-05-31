"""
Monte Carlo simulace sítě poboček
===================================
Opakovaně spouští simulaci přesunu klientů ze zavřených poboček s náhodnou
variací vstupních parametrů a odhaduje robustnost výsledků:

  • Stochastické redistribuční váhy (klienti, kteří přejdou na spádovou pobočku)
  • Variace výnosových trendů (CAGR ± σ)
  • Výsledky: distribuce SIM_REV_2030, kapacitní přetížení, percentily

Použití
--------
    from monte_carlo import run_monte_carlo

    results = run_monte_carlo(
        df            = rating_status,      # hlavní DataFrame (z notebooku)
        spadovky_df   = spadovky,           # spádová tabulka
        n_iter        = 500,                # počet iterací
        target_n      = 250,               # cílový počet poboček
        seed          = 42,
    )
    print(results['summary'])
    results['html']  # HTML report pro vložení do notebooku

Závislosti: numpy, pandas, scipy (volitelně pro KDE)
"""

from __future__ import annotations

import warnings
import math
from collections import defaultdict
from typing import Optional

import numpy as np
import pandas as pd

# ── Konstanty (shodné s generate_network_simulation_200) ─────────────────────
FORMAT_CAP: dict[str, float] = {
    'small': 0.10, 'medium economy': 0.25, 'medium': 0.40, 'flagship': 0.70,
}
MEETINGS_PER_DAY   = 5
WORKING_DAYS       = 248
BANKER_CAP_YEAR    = MEETINGS_PER_DAY * WORKING_DAYS   # 1 240 sch./bankéř/rok
BEZHOT_TO_SCH_RATE = 0.30
BEZHOT_SCH_EQUIV   = 20 / 45
PORTFOLIO_PER_BANKER = 1500

# Variace (směrodatné odchylky pro náhodné výběry)
PRESUN_SIGMA       = 0.10   # ± 10 % relativní odchylka redistribučních vah
CAGR_SIGMA         = 0.01   # ± 1 pp absolutní odchylka CAGR
CI_SIGMA           = 0.02   # ± 2 pp absolutní odchylka C/I trendu


def _sf(v: object) -> float:
    """NaN-safe float konverze."""
    try:
        f = float(v)
        return 0.0 if f != f else f
    except (TypeError, ValueError):
        return 0.0


def _one_iteration(
    d_remain: pd.DataFrame,
    d_close: pd.DataFrame,
    sp: Optional[pd.DataFrame],
    has_presun_pct: bool,
    rng: np.random.Generator,
    target_n: int,
) -> dict:
    """
    Spustí jednu iteraci simulace s náhodně perturbovanými váhami.
    Vrací dict s klíčovými metrikami.
    """
    d = d_remain.copy()

    d['_delta_cli']        = 0.0
    d['_delta_cli_online'] = 0.0
    d['_delta_rev']        = 0.0
    d['_max_extra_cli'] = (
        d['POCET_KLIENTU'].fillna(0) *
        d['BRANCH_FORMAT'].str.lower().str.strip()
         .map(lambda f: FORMAT_CAP.get(f, FORMAT_CAP['medium']))
    )

    # ── Přesun klientů ────────────────────────────────────────────────────────
    if sp is not None and not sp.empty:
        for _, crow in d_close.iterrows():
            bc     = int(crow['BRANCH_CODE'])
            bc_cli = _sf(crow.get('POCET_KLIENTU'))
            sp_r   = sp[sp['BRANCH_ID'] == bc]
            if sp_r.empty or bc_cli <= 0:
                continue
            sp_r = sp_r.iloc[0]

            candidates: list[tuple[int, float]] = []
            if has_presun_pct:
                raw_weights = []
                tids = []
                for _i in [1, 2, 3]:
                    _tid = sp_r.get(f'TOP_POBOCKA_ID_{_i}')
                    _pct = _sf(sp_r.get(f'ODHADOVANY_PRESUN_PCT_{_i}', 0))
                    if pd.isna(_tid) or _pct <= 0:
                        continue
                    try:
                        _tid = int(float(_tid))
                    except (ValueError, TypeError):
                        continue
                    if _tid not in d.index:
                        continue
                    # Perturbuj váhu: multiplikativní šum (log-normální)
                    noise = float(rng.lognormal(0.0, PRESUN_SIGMA))
                    raw_weights.append(_pct * noise)
                    tids.append(_tid)

                if not tids:
                    continue
                # Renormalizuj na 95 % (zachovej 5% úbytek)
                total_w = sum(raw_weights)
                if total_w > 0:
                    candidates = [(tid, w / total_w * 0.95) for tid, w in zip(tids, raw_weights)]
            else:
                total_vis = sum(
                    _sf(sp_r.get(f'TOP_POBOCKA_VISITS_{i}', 0)) for i in [1, 2, 3]
                )
                if total_vis <= 0:
                    continue
                for _i in [1, 2, 3]:
                    _tid  = sp_r.get(f'TOP_POBOCKA_ID_{_i}')
                    _tvis = _sf(sp_r.get(f'TOP_POBOCKA_VISITS_{_i}', 0))
                    if pd.isna(_tid) or _tvis <= 0:
                        continue
                    try:
                        _tid = int(float(_tid))
                    except (ValueError, TypeError):
                        continue
                    if _tid not in d.index:
                        continue
                    noise = float(rng.lognormal(0.0, PRESUN_SIGMA))
                    candidates.append((_tid, _tvis / total_vis * noise))

                # Renormalizuj
                total_w = sum(w for _, w in candidates)
                if total_w > 0:
                    candidates = [(_tid, w / total_w) for _tid, w in candidates]

            _rpv_close = _sf(crow.get('OBJEM_VYNOSU_CZK')) / max(_sf(crow.get('POCET_KLIENTU')), 1)
            for _orig_tid, _weight in candidates:
                _cli_wave = bc_cli * _weight
                _cap_rem  = max(0.0, d.at[_orig_tid, '_max_extra_cli'] - d.at[_orig_tid, '_delta_cli'])
                _cli_phys = min(_cli_wave, _cap_rem)
                _cli_onl  = _cli_wave - _cli_phys
                d.at[_orig_tid, '_delta_cli']        += _cli_phys
                d.at[_orig_tid, '_delta_cli_online'] += _cli_onl
                d.at[_orig_tid, '_delta_rev']        += (
                    _cli_phys * _rpv_close + _cli_onl * _rpv_close * 0.60
                )

    d['SIM_KLIENTI'] = d['POCET_KLIENTU'].fillna(0) + d['_delta_cli']
    d['SIM_REV']     = d['OBJEM_VYNOSU_CZK'].fillna(0) + d['_delta_rev']

    # ── Výnosový trend 2030 (s perturbací CAGR a C/I) ────────────────────────
    _cli    = d['POCET_KLIENTU'].fillna(0).values
    _rev_now = d['OBJEM_VYNOSU_CZK'].fillna(0).values
    _rev_per_cli_now = np.where(_cli > 0, _rev_now / np.where(_cli > 0, _cli, 1), 0.0)

    # Perturbace CAGR (±CAGR_SIGMA absolutně)
    _cagr_base = d['_cagr'].fillna(0).values if '_cagr' in d.columns else np.zeros(len(d))
    _cagr_noise = rng.normal(0.0, CAGR_SIGMA, size=len(d))
    _cagr = np.clip(_cagr_base + _cagr_noise, -0.20, 0.30)

    years_ahead = 5
    _rpc_2030 = _rev_per_cli_now * (1 + _cagr) ** years_ahead
    d['SIM_REV_2030'] = d['SIM_KLIENTI'] * _rpc_2030

    # ── Rating 2030 ───────────────────────────────────────────────────────────
    _mask = d['IR'].notna() & (d['IR'] > 0) if 'IR' in d.columns else pd.Series(True, index=d.index)
    d_sim = d[_mask].copy()

    if d_sim.empty or 'SIM_REV_2030' not in d_sim.columns:
        return {}

    d_sim['_rev_rnk'] = d_sim['SIM_REV_2030'].rank(ascending=False, method='first')

    ci_base = d_sim['PRIME_NAKLADY/VYNOSY'].fillna(0.5).values if 'PRIME_NAKLADY/VYNOSY' in d_sim.columns else np.full(len(d_sim), 0.5)
    ci_noise = rng.normal(0.0, CI_SIGMA, size=len(d_sim))
    ci_proj = np.clip(ci_base + ci_noise, 0.01, 3.0)
    d_sim['_ci_proj'] = ci_proj
    d_sim['_ci_perc'] = (
        pd.Series(ci_proj, index=d_sim.index)
         .rank(method='first', pct=True, ascending=True) * 100
    ).apply(lambda x: max(1, int(round(x))))
    d_sim['SIM_IR_SCORE'] = d_sim['_ci_perc'] * 0.4 + d_sim['_rev_rnk'] * 0.6
    d_sim['SIM_IR_2030']  = d_sim['SIM_IR_SCORE'].rank(ascending=True, method='first').astype(int)

    top = d_sim.nsmallest(target_n, 'SIM_IR_2030')

    # ── Metriky ───────────────────────────────────────────────────────────────
    total_rev_2030  = float(top['SIM_REV_2030'].fillna(0).sum())
    total_cli_sim   = float(top['SIM_KLIENTI'].fillna(0).sum())
    n_overloaded    = int((top['_delta_cli_online'].fillna(0) > 0).sum())
    total_overflow  = float(top['_delta_cli_online'].fillna(0).sum())
    keep_codes      = set(top['BRANCH_CODE'].dropna().astype(int).tolist())

    return {
        'total_rev_2030':  total_rev_2030,
        'total_cli_sim':   total_cli_sim,
        'n_overloaded':    n_overloaded,
        'total_overflow':  total_overflow,
        'keep_codes':      keep_codes,
    }


def run_monte_carlo(
    df: pd.DataFrame,
    spadovky_df: Optional[pd.DataFrame] = None,
    n_iter: int = 500,
    target_n: int = 250,
    seed: int = 42,
    close_col: str = 'FOOTPRINT_STRATEGIE_(2030)',
    close_year_col: str = 'PLAN_CLOSE_(ROK)',
    sim_year: int = 2030,
    verbose: bool = True,
) -> dict:
    """
    Spustí Monte Carlo simulaci sítě poboček.

    Parametry
    ---------
    df            : hlavní DataFrame (rating_status)
    spadovky_df   : spádová tabulka (surová, před merge)
    n_iter        : počet Monte Carlo iterací
    target_n      : cílový počet poboček v síti
    seed          : seed pro reprodukovatelnost
    close_col     : sloupec strategie (obsahuje 'close')
    close_year_col: sloupec roku uzavření
    sim_year      : rok simulace (pro filtr close)
    verbose       : tisk průběhu

    Vrací
    -----
    dict s klíči:
      'summary'     : pd.DataFrame se souhrnnou statistikou
      'keep_freq'   : pd.Series — frekvence výskytu každé pobočky v top N
      'rev_dist'    : np.ndarray — distribuce celkových výnosů 2030
      'overflow_dist': np.ndarray — distribuce počtu přetížených poboček
      'html'        : str — HTML report
    """
    rng = np.random.default_rng(seed)

    # ── Příprava dat ─────────────────────────────────────────────────────────
    df = df.copy()
    required_cols = ['BRANCH_CODE', 'POCET_KLIENTU', 'OBJEM_VYNOSU_CZK']
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"Chybí sloupec: {c}")

    if 'BRANCH_CODE' in df.columns:
        df = df.set_index('BRANCH_CODE') if df.index.name != 'BRANCH_CODE' else df
        if df.index.name != 'BRANCH_CODE':
            df = df.set_index('BRANCH_CODE')

    # Identifikace zavíraných poboček (ke dni sim_year)
    _is_close = (
        df[close_col].astype(str).str.lower().str.contains('close', na=False)
        if close_col in df.columns else pd.Series(False, index=df.index)
    )
    _close_yr = pd.to_numeric(
        df[close_year_col] if close_year_col in df.columns else pd.Series(dtype=float),
        errors='coerce'
    ).reindex(df.index)
    _close_mask = _is_close & (_close_yr.fillna(9999) <= sim_year)

    d_close  = df[_close_mask].copy()
    d_remain = df[~_close_mask].copy()

    if d_remain.empty:
        raise ValueError("d_remain je prázdné — zkontroluj close_col a close_year_col.")

    # Příprava spádové tabulky
    sp_prep: Optional[pd.DataFrame] = None
    has_presun_pct = False
    if spadovky_df is not None and not spadovky_df.empty:
        sp_prep = spadovky_df.copy()
        sp_prep.columns = [c.strip().upper() for c in sp_prep.columns]
        sp_prep['BRANCH_ID'] = pd.to_numeric(sp_prep['BRANCH_ID'], errors='coerce')
        has_presun_pct = all(f'ODHADOVANY_PRESUN_PCT_{i}' in sp_prep.columns for i in [1, 2, 3])

    if verbose:
        print(f"Monte Carlo: {n_iter} iteraci, target_n={target_n}, "
              f"close={len(d_close)}, remain={len(d_remain)}, "
              f"spadovky={'ano (pct)' if has_presun_pct else ('ano (visits)' if sp_prep is not None else 'ne')}")

    # ── Iterace ───────────────────────────────────────────────────────────────
    rev_dist:      list[float] = []
    overflow_dist: list[int]   = []
    ovf_cli_dist:  list[float] = []
    keep_counter:  dict[int, int] = defaultdict(int)

    for i in range(n_iter):
        if verbose and (i + 1) % 100 == 0:
            print(f"  iterace {i + 1}/{n_iter} …")
        res = _one_iteration(d_remain, d_close, sp_prep, has_presun_pct, rng, target_n)
        if not res:
            continue
        rev_dist.append(res['total_rev_2030'])
        overflow_dist.append(res['n_overloaded'])
        ovf_cli_dist.append(res['total_overflow'])
        for bc in res['keep_codes']:
            keep_counter[bc] += 1

    n_valid = len(rev_dist)
    if n_valid == 0:
        raise RuntimeError("Žádná platná iterace — zkontroluj vstupní data.")

    rev_arr = np.array(rev_dist)
    ovf_arr = np.array(overflow_dist)

    # ── Souhrnná statistika ───────────────────────────────────────────────────
    def _pct(arr, q):
        return float(np.percentile(arr, q))

    summary = pd.DataFrame({
        'metrika': [
            'Celkové výnosy sítě 2030 (Kč)',
            'Přetížené pobočky (počet)',
            'Klienti přesměrovaní online',
        ],
        'medián': [
            _pct(rev_arr, 50),
            _pct(ovf_arr, 50),
            _pct(np.array(ovf_cli_dist), 50),
        ],
        'P10': [
            _pct(rev_arr, 10),
            _pct(ovf_arr, 10),
            _pct(np.array(ovf_cli_dist), 10),
        ],
        'P90': [
            _pct(rev_arr, 90),
            _pct(ovf_arr, 90),
            _pct(np.array(ovf_cli_dist), 90),
        ],
        'σ': [
            float(np.std(rev_arr)),
            float(np.std(ovf_arr)),
            float(np.std(np.array(ovf_cli_dist))),
        ],
    })

    # Stabilita výběru: pobočky v top N ve více než 80 % iterací
    keep_freq = pd.Series(keep_counter).sort_values(ascending=False)
    keep_freq_pct = (keep_freq / n_valid * 100).round(1)
    n_stable   = int((keep_freq_pct >= 80).sum())
    n_volatile = int((keep_freq_pct < 50).sum())

    # ── HTML report ───────────────────────────────────────────────────────────
    def _fmt_m(v: float) -> str:
        if abs(v) >= 1e9:  return f"{v/1e9:.2f} mld. Kč"
        if abs(v) >= 1e6:  return f"{v/1e6:.1f} M Kč"
        return f"{v:,.0f} Kč"

    _rev_med = _pct(rev_arr, 50)
    _rev_p10 = _pct(rev_arr, 10)
    _rev_p90 = _pct(rev_arr, 90)
    _ovf_med = _pct(ovf_arr, 50)
    _ovf_p10 = _pct(ovf_arr, 10)
    _ovf_p90 = _pct(ovf_arr, 90)

    # Top 20 nejvolatilnějších poboček (v rozmezí 40–80 % stabilita)
    _volatile_bcs = keep_freq_pct[(keep_freq_pct >= 20) & (keep_freq_pct < 80)].head(20)
    _volatile_rows = ''
    for bc, pct_val in _volatile_bcs.items():
        _brow = df[df.index == bc]
        _bname = str(_brow['BRANCH_NAME'].iloc[0]) if 'BRANCH_NAME' in _brow.columns and len(_brow) > 0 else str(bc)
        _breg  = str(_brow['REGION_NAME'].iloc[0]) if 'REGION_NAME' in _brow.columns and len(_brow) > 0 else '—'
        _bclr  = '#f59e0b' if pct_val >= 60 else ('#ea580c' if pct_val >= 40 else '#eb4d79')
        _bw    = int(pct_val)
        _volatile_rows += (
            f'<tr style="border-bottom:1px solid #f0f2f8;">'
            f'<td style="padding:4px 8px;font-size:0.78rem;font-weight:600;">{_bname}</td>'
            f'<td style="padding:4px 8px;font-size:0.72rem;color:#666;">{_breg}</td>'
            f'<td style="padding:4px 8px;">'
            f'<div style="display:flex;align-items:center;gap:4px;">'
            f'<div style="background:#f3f4f6;border-radius:3px;height:6px;width:70px;overflow:hidden;">'
            f'<div style="background:{_bclr};width:{_bw}%;height:100%;border-radius:3px;"></div></div>'
            f'<span style="font-size:0.75rem;font-weight:700;color:{_bclr};">{pct_val:.0f} %</span>'
            f'</div></td>'
            f'</tr>'
        )

    # Pre-compute HTML parts to avoid nested f-string issues (Python 3.11)
    _kpi_items = [
        ('Vynosy site 2030 - median', _fmt_m(_rev_med), _fmt_m(_rev_p10), _fmt_m(_rev_p90)),
        ('Pretizene pobocky - median', f"{_ovf_med:.0f}", f"{_ovf_p10:.0f}", f"{_ovf_p90:.0f}"),
        ('Stabilnich pobocek (>=80 %)', str(n_stable), '-', '-'),
        ('Volatilnich pobocek (<50 %)', str(n_volatile), '-', '-'),
    ]
    _kpi_html = ''
    for _lbl, _med, _p10, _p90 in _kpi_items:
        _kpi_html += (
            f'<div style="background:#f8f7ff;border-radius:8px;padding:10px 16px;'
            f'border:1px solid #ede9fe;text-align:center;min-width:160px;">'
            f'<div style="font-size:0.65rem;color:#888;text-transform:uppercase;'
            f'letter-spacing:.3px;margin-bottom:4px;">{_lbl}</div>'
            f'<div style="font-size:1.2rem;font-weight:800;color:#7c3aed;">{_med}</div>'
            f'<div style="font-size:0.68rem;color:#aaa;">P10: {_p10} / P90: {_p90}</div>'
            f'</div>'
        )

    if _volatile_rows:
        _vol_section = (
            f'<div style="margin-bottom:14px;">'
            f'<div style="font-size:0.8rem;font-weight:700;color:#374151;margin-bottom:6px;">'
            f'Volatilni pobocky (20-80 % vyskytu v top {target_n})</div>'
            f'<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;">'
            f'<thead><tr style="background:#f5f3ff;">'
            f'<th style="padding:4px 8px;font-size:0.72rem;font-weight:700;color:#6d28d9;text-align:left;">Pobocka</th>'
            f'<th style="padding:4px 8px;font-size:0.72rem;font-weight:700;color:#6d28d9;text-align:left;">Region</th>'
            f'<th style="padding:4px 8px;font-size:0.72rem;font-weight:700;color:#6d28d9;">Stabilita</th>'
            f'</tr></thead><tbody>{_volatile_rows}</tbody></table></div></div>'
        )
    else:
        _vol_section = (
            '<div style="color:#aaa;font-size:0.78rem;font-style:italic;">'
            'Vsechny pobocky jsou stabilni (&gt;=80 % nebo &lt;=20 % vyskytu).</div>'
        )

    _params_str = (
        f'PRESUN_SIGMA={PRESUN_SIGMA:.0%}, CAGR_SIGMA={CAGR_SIGMA:.0%}, '
        f'CI_SIGMA={CI_SIGMA:.0%}, seed={seed}, iteraci={n_valid}/{n_iter}'
    )
    _title_str = (
        f'Monte Carlo - simulace site {target_n} pobocek '
        f'({n_valid} iteraci, sigma presunu={int(PRESUN_SIGMA*100)} %, '
        f'sigma CAGR={int(CAGR_SIGMA*100)} pp)'
    )

    html = (
        f'<div style="background:#fff;border-radius:10px;padding:16px 20px;'
        f'border:1px solid #e8edf5;box-shadow:0 2px 8px rgba(0,0,0,0.05);">'
        f'<div style="font-size:0.72rem;color:#7c3aed;font-weight:700;text-transform:uppercase;'
        f'letter-spacing:.5px;margin-bottom:8px;">{_title_str}</div>'
        f'<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;">{_kpi_html}</div>'
        f'{_vol_section}'
        f'<div style="font-size:0.68rem;color:#aaa;margin-top:8px;line-height:1.6;">{_params_str}</div>'
        f'</div>'
    )

    if verbose:
        print(f"\n✅ Hotovo: {n_valid} platných iterací")
        print(f"   Výnosy 2030 — medián: {_fmt_m(_rev_med)}  (P10: {_fmt_m(_rev_p10)} — P90: {_fmt_m(_rev_p90)})")
        print(f"   Přetížené pobočky — medián: {_ovf_med:.0f}  (P10: {_ovf_p10:.0f} — P90: {_ovf_p90:.0f})")
        print(f"   Stabilní (≥80 %): {n_stable}  |  Volatilní (<50 %): {n_volatile}")

    return {
        'summary':      summary,
        'keep_freq':    keep_freq_pct,
        'rev_dist':     rev_arr,
        'overflow_dist': ovf_arr,
        'html':         html,
        'n_valid':      n_valid,
        'n_stable':     n_stable,
        'n_volatile':   n_volatile,
    }


# ── Samostatné spuštění (ukázka) ─────────────────────────────────────────────
if __name__ == '__main__':
    print("monte_carlo.py — importuj run_monte_carlo() z notebooku.")
    print("Příklad:")
    print("  from monte_carlo import run_monte_carlo")
    print("  res = run_monte_carlo(rating_status, spadovky, n_iter=200)")
    print("  display(HTML(res['html']))")
