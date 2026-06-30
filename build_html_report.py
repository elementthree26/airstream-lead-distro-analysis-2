import pandas as pd
import numpy as np
import json
import warnings
warnings.filterwarnings('ignore')

UPLOAD_DIR = '/root/.claude/uploads/272124da-ce12-5c3a-a827-6ea78f998512'

# ── Load baseline ──────────────────────────────────────────────────────────────
files = [
    f'{UPLOAD_DIR}/c386145b-040120256302025_leads.xlsx',
    f'{UPLOAD_DIR}/2a173753-7120259302025.xlsx',
    f'{UPLOAD_DIR}/7cc26c1f-101202512312025_Leads.xlsx',
    f'{UPLOAD_DIR}/7b43883a-139_LeadsExport6_18_2026.xlsx',
    f'{UPLOAD_DIR}/ebdaf3dd-3120265262026_Leads.xlsx',
]
dfs = [pd.read_excel(f) for f in files]
raw = pd.concat(dfs, ignore_index=True)
raw['Lead Date'] = pd.to_datetime(raw['Lead Date'], errors='coerce')
bl = raw[(raw['Lead Date'] >= '2025-04-01') & (raw['Lead Date'] <= '2026-04-30')].copy()
if 'Program' in bl.columns:
    bl = bl[bl['Program'].astype(str).str.lower().str.strip() != 'align']

# ── Load post-launch ───────────────────────────────────────────────────────────
post_raw = pd.read_csv(f'{UPLOAD_DIR}/417453fe-5282661826_Leads_Post_Launch.csv')
post_raw['Lead Date'] = pd.to_datetime(post_raw['Lead Date'], errors='coerce')
post = post_raw[post_raw['Lead Date'] >= '2026-05-28'].copy()
if 'Program' in post.columns:
    post = post[post['Program'].astype(str).str.lower().str.strip() != 'align']

# ── Baseline metrics ───────────────────────────────────────────────────────────
total_leads = len(bl)
n_months = 13
n_dealers = 83
monthly_avg = total_leads / n_months
per_dealer_mo = monthly_avg / n_dealers

bl['Lead Score'] = pd.to_numeric(bl['Lead Score'], errors='coerce')
score3plus_rate = (bl['Lead Score'] >= 3).sum() / total_leads

# Monthly volume
bl['Month'] = bl['Lead Date'].dt.to_period('M')
mv = bl.groupby('Month').size().reset_index(name='leads').sort_values('Month')
mo_map = {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun',
          '07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'}
def fmt_month(p):
    s = str(p)  # "2025-04"
    yr, mo = s.split('-')
    return f"{mo_map[mo]} '{yr[2:]}"
mv['label'] = mv['Month'].apply(fmt_month)
monthly_labels = mv['label'].tolist()
monthly_values = mv['leads'].tolist()

# Source split
def classify_source(s):
    if pd.isna(s): return 'Website-Captured'
    sl = str(s).lower()
    return 'External Paid (Meta)' if any(x in sl for x in ['paid_social','paid social','meta','facebook']) else 'Website-Captured'

bl['SourceType'] = bl['Lead Source'].apply(classify_source)
src = bl['SourceType'].value_counts()
paid_count = int(src.get('External Paid (Meta)', 0))
web_count  = int(src.get('Website-Captured', 0))
paid_pct   = paid_count / total_leads

# Score distribution (exclude score 0 — these are unscored/null proxies)
score_dist = bl[bl['Lead Score'] >= 1]['Lead Score'].value_counts().sort_index()
score_labels = [f"Score {int(k)}" for k in score_dist.index]
score_values = [int(v) for v in score_dist.values]
score_colors = ['#A8A9AD','#A8A9AD','#4A7CC7','#1B3A6B','#0D1F3C'][:len(score_labels)]

# Brand split
def classify_brand(row):
    for col in ['Brand','ManufacturerName']:
        if col in row.index and pd.notna(row[col]):
            v = str(row[col]).lower()
            if 'touring' in v: return 'Touring Coach (TC)'
            if 'travel' in v:  return 'Travel Trailer (TT)'
    return 'Unknown'
bl['BrandType'] = bl.apply(classify_brand, axis=1)
bc = bl['BrandType'].value_counts()
tt_count = int(bc.get('Travel Trailer (TT)', 0))
tc_count = int(bc.get('Touring Coach (TC)', 0))

# Score 3+ by brand
tt_s3 = (bl[bl['BrandType']=='Travel Trailer (TT)']['Lead Score'] >= 3).mean()
tc_s3 = (bl[bl['BrandType']=='Touring Coach (TC)']['Lead Score'] >= 3).mean()

# ── Post-launch metrics ────────────────────────────────────────────────────────
post_days  = 22
post_total = len(post)
post_monthly = post_total / post_days * 30
post_per_dealer = post_monthly / n_dealers

post['Lead Score'] = pd.to_numeric(post['Lead Score'], errors='coerce')
post_score3rate = (post['Lead Score'] >= 3).sum() / post_total

post['SourceType'] = post['Lead Source'].apply(classify_source)
post_src = post['SourceType'].value_counts()
post_paid_pct = post_src.get('External Paid (Meta)', 0) / post_total

# ── Forecast reference values (from PDF Section 5) ────────────────────────────
FORECAST_MONTHLY   = 17891   # Form elimination only — implemented path
FORECAST_PER_DEALER = 216
BASELINE_MONTHLY_PDF = 18593  # Actual 13-month baseline average

# ── Variance helpers ───────────────────────────────────────────────────────────
def variance_class(actual, forecast):
    if actual is None or forecast == 0: return 'neutral'
    pct = abs(actual - forecast) / forecast
    if pct <= 0.05:   return 'green'
    elif pct <= 0.15: return 'yellow'
    else:             return 'red'

def pct_diff(actual, forecast):
    if forecast == 0: return ''
    d = (actual - forecast) / forecast * 100
    sign = '+' if d > 0 else ''
    return f"{sign}{d:.1f}%"

# ── Print summary ──────────────────────────────────────────────────────────────
print(f"Baseline  : {total_leads:,} leads | {monthly_avg:,.0f}/mo | {per_dealer_mo:.1f}/dealer | Score3+={score3plus_rate:.1%}")
print(f"Source    : Paid={paid_count:,} ({paid_pct:.1%}) | Web={web_count:,} ({1-paid_pct:.1%})")
print(f"Brand     : TT={tt_count:,} ({tt_count/total_leads:.1%}) | TC={tc_count:,} ({tc_count/total_leads:.1%})")
print(f"Post-launch: {post_total} leads | {post_monthly:,.0f} implied/mo | {post_per_dealer:.1f}/dealer | Score3+={post_score3rate:.1%}")
print(f"Post paid%: {post_paid_pct:.1%}")

# ── Scenario data (from PDF Section 5 exactly) ────────────────────────────────
scenarios = [
    {"label": "Current baseline (form submit)",        "total": 18593, "implemented": False},
    {"label": "Form elimination only",                 "total": 17891, "implemented": True},
    {"label": "Score-change notifications (all, 0→1)", "total": 17254, "implemented": False},
    {"label": "Elim + 50% checkbox opt-in",            "total": 14715, "implemented": False},
    {"label": "Elim + 30% checkbox opt-in",            "total": 13444, "implemented": False},
    {"label": "Elim + 2.4% checkbox (config proxy)",   "total": 11691, "implemented": False},
    {"label": "Elim + 30% CB + Score 2+ threshold",    "total": 5208,  "implemented": False},
    {"label": "Elim + 30% CB + Score 3+ threshold",    "total": 1263,  "implemented": False},
]

# ── Build HTML ─────────────────────────────────────────────────────────────────
vc_monthly    = variance_class(post_monthly, FORECAST_MONTHLY)
vc_dealer     = variance_class(post_per_dealer, FORECAST_PER_DEALER)
vc_score3     = variance_class(post_score3rate, score3plus_rate)
vc_paid       = variance_class(post_paid_pct, paid_pct)

pd_monthly    = pct_diff(post_monthly, FORECAST_MONTHLY)
pd_dealer     = pct_diff(post_per_dealer, FORECAST_PER_DEALER)
pd_score3     = pct_diff(post_score3rate * 100, score3plus_rate * 100)
pd_paid       = pct_diff(post_paid_pct * 100, paid_pct * 100)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Airstream — Dealer Lead Distribution</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
    background: #f0f2f5;
    color: #1a1a2e;
    font-size: 14px;
    line-height: 1.5;
  }}

  /* ── LAYOUT ── */
  .page {{
    max-width: 960px;
    margin: 0 auto;
    background: #fff;
    box-shadow: 0 2px 24px rgba(0,0,0,0.10);
  }}

  /* ── HEADER ── */
  .site-header {{
    background: #1B3A6B;
    padding: 28px 48px 24px;
    border-bottom: 4px solid #A8A9AD;
  }}
  .site-header .wordmark {{
    font-size: 11px;
    letter-spacing: 6px;
    text-transform: uppercase;
    color: #A8A9AD;
    font-weight: 600;
    margin-bottom: 10px;
  }}
  .site-header h1 {{
    font-size: 28px;
    font-weight: 700;
    color: #fff;
    letter-spacing: -0.5px;
    line-height: 1.2;
  }}
  .site-header .subtitle {{
    margin-top: 6px;
    font-size: 15px;
    color: #A8C4E0;
  }}
  .meta-pills {{
    display: flex;
    gap: 12px;
    margin-top: 20px;
    flex-wrap: wrap;
  }}
  .meta-pill {{
    background: rgba(255,255,255,0.10);
    border: 1px solid rgba(255,255,255,0.20);
    border-radius: 4px;
    padding: 6px 14px;
    color: #fff;
    font-size: 12px;
  }}
  .meta-pill strong {{ color: #A8C4E0; font-size: 10px; display: block; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 2px; }}

  /* ── SECTIONS ── */
  .section {{
    padding: 40px 48px;
    border-bottom: 1px solid #e8eaed;
  }}
  .section:last-child {{ border-bottom: none; }}
  .section-label {{
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #A8A9AD;
    margin-bottom: 6px;
  }}
  .section-title {{
    font-size: 22px;
    font-weight: 700;
    color: #1B3A6B;
    margin-bottom: 4px;
    letter-spacing: -0.3px;
  }}
  .section-sub {{
    font-size: 13px;
    color: #888;
    margin-bottom: 28px;
  }}

  /* ── KPI GRID ── */
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 14px;
    margin-bottom: 36px;
  }}
  .kpi-card {{
    background: #F4F5F6;
    border-radius: 8px;
    padding: 18px 16px 14px;
  }}
  .kpi-card .kpi-label {{
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #A8A9AD;
    margin-bottom: 8px;
  }}
  .kpi-card .kpi-value {{
    font-size: 26px;
    font-weight: 700;
    color: #1B3A6B;
    line-height: 1;
    margin-bottom: 4px;
  }}
  .kpi-card .kpi-sub {{
    font-size: 11px;
    color: #999;
  }}

  /* ── CHARTS ── */
  .chart-grid-3 {{
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 24px;
    margin-bottom: 20px;
    align-items: start;
  }}
  .chart-grid-2 {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 24px;
    margin-bottom: 20px;
    align-items: start;
  }}
  .chart-box {{
    background: #F4F5F6;
    border-radius: 8px;
    padding: 20px;
  }}
  .chart-box h3 {{
    font-size: 13px;
    font-weight: 700;
    color: #1B3A6B;
    margin-bottom: 16px;
    letter-spacing: -0.2px;
  }}
  .chart-box canvas {{
    display: block;
    width: 100% !important;
  }}

  /* ── BRAND STAT STRIP ── */
  .brand-strip {{
    background: #F4F5F6;
    border-radius: 8px;
    padding: 16px 20px;
    display: flex;
    gap: 32px;
    flex-wrap: wrap;
    margin-top: 8px;
  }}
  .brand-stat {{ font-size: 13px; color: #444; }}
  .brand-stat strong {{ color: #1B3A6B; font-size: 15px; }}

  /* ── SCENARIO CHART ── */
  .scenario-wrap {{
    background: #F4F5F6;
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 28px;
  }}
  .scenario-wrap h3 {{
    font-size: 13px;
    font-weight: 700;
    color: #1B3A6B;
    margin-bottom: 16px;
  }}

  /* ── DECISION TABLE ── */
  .decision-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }}
  .decision-table thead tr {{
    background: #1B3A6B;
    color: #fff;
  }}
  .decision-table th {{
    padding: 11px 14px;
    text-align: left;
    font-weight: 600;
    font-size: 12px;
    letter-spacing: 0.3px;
  }}
  .decision-table td {{
    padding: 11px 14px;
    color: #333;
    vertical-align: top;
    border-bottom: 1px solid #e8eaed;
  }}
  .decision-table tbody tr:nth-child(odd) {{ background: #F4F5F6; }}
  .decision-table tbody tr:nth-child(even) {{ background: #fff; }}
  .badge {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 3px;
    font-size: 11px;
    font-weight: 600;
  }}
  .badge-navy {{ background: #1B3A6B; color: #fff; }}
  .badge-silver {{ background: #e0e0e0; color: #555; }}

  /* ── COMPARISON TABLE ── */
  .comparison-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    margin-bottom: 28px;
  }}
  .comparison-table thead tr {{
    background: #1B3A6B;
    color: #fff;
  }}
  .comparison-table th {{
    padding: 12px 16px;
    text-align: left;
    font-weight: 600;
    font-size: 12px;
    letter-spacing: 0.3px;
  }}
  .comparison-table td {{
    padding: 13px 16px;
    border-bottom: 1px solid #e8eaed;
    vertical-align: middle;
  }}
  .comparison-table tbody tr:nth-child(odd) {{ background: #F4F5F6; }}
  .comparison-table tbody tr:nth-child(even) {{ background: #fff; }}
  .comparison-table .metric-name {{ font-weight: 600; color: #1B3A6B; }}
  .comparison-table .actual-cell {{
    font-weight: 700;
    font-size: 14px;
  }}
  .cell-green  {{ color: #1B6B2A; }}
  .cell-yellow {{ color: #8a5a00; }}
  .cell-red    {{ color: #9B1C1C; }}
  .cell-neutral {{ color: #666; }}
  .delta {{
    font-size: 11px;
    font-weight: 600;
    margin-left: 6px;
    padding: 1px 5px;
    border-radius: 3px;
    display: inline-block;
  }}
  .delta-green  {{ background: #d4edda; color: #1B6B2A; }}
  .delta-yellow {{ background: #fff3cd; color: #8a5a00; }}
  .delta-red    {{ background: #fde8e8; color: #9B1C1C; }}

  /* ── CAVEAT BOX ── */
  .caveat-box {{
    background: #FFF8E1;
    border-left: 4px solid #F9A825;
    border-radius: 4px;
    padding: 16px 20px;
    font-size: 12.5px;
    color: #555;
    line-height: 1.6;
    margin-top: 8px;
  }}
  .caveat-box strong {{ color: #333; }}

  /* ── FOOTER ── */
  .site-footer {{
    background: #1B3A6B;
    padding: 20px 48px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }}
  .site-footer .footer-left {{ font-size: 12px; color: #A8C4E0; letter-spacing: 0.5px; }}
  .site-footer .footer-right {{ font-size: 11px; color: #6890b8; }}
</style>
</head>
<body>
<div class="page">

<!-- ════════════════════════════════════════════════════════ HEADER -->
<header class="site-header">
  <div class="wordmark">Airstream</div>
  <h1>Dealer Lead Distribution</h1>
  <div class="subtitle">Pre-Launch Forecast &amp; Post-Launch Actuals</div>
  <div class="meta-pills">
    <div class="meta-pill"><strong>Baseline Window</strong>April 2025 – April 2026</div>
    <div class="meta-pill"><strong>Website Launch</strong>May 27, 2026</div>
    <div class="meta-pill"><strong>Post-Launch Window</strong>May 28 – June 18, 2026 (22 days)</div>
    <div class="meta-pill"><strong>Prepared By</strong>Element Three · June 2026</div>
  </div>
</header>

<!-- ════════════════════════════════════════════════════════ SECTION 1 — BASELINE -->
<section class="section">
  <div class="section-label">Section 1</div>
  <div class="section-title">Baseline: The Pipeline Before Launch</div>
  <div class="section-sub">April 2025 – April 2026 &nbsp;·&nbsp; 83 dealers &nbsp;·&nbsp; Align program excluded</div>

  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-label">Total Leads</div>
      <div class="kpi-value">{total_leads:,}</div>
      <div class="kpi-sub">13-month window</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Avg / Month</div>
      <div class="kpi-value">{monthly_avg:,.0f}</div>
      <div class="kpi-sub">across all 83 dealers</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Avg / Dealer / Mo</div>
      <div class="kpi-value">{per_dealer_mo:.0f}</div>
      <div class="kpi-sub">per active dealer</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Score 3+ Rate</div>
      <div class="kpi-value">{score3plus_rate:.1%}</div>
      <div class="kpi-sub">high-intent leads</div>
    </div>
  </div>

  <div class="chart-grid-3">
    <div class="chart-box">
      <h3>Monthly Lead Volume</h3>
      <canvas id="monthlyChart" height="180"></canvas>
    </div>
    <div class="chart-box">
      <h3>Lead Source Split</h3>
      <canvas id="sourceChart" height="180"></canvas>
    </div>
    <div class="chart-box">
      <h3>Score Distribution</h3>
      <canvas id="scoreChart" height="180"></canvas>
    </div>
  </div>

  <div class="brand-strip">
    <div class="brand-stat">Travel Trailer (TT) &nbsp; <strong>{tt_count:,}</strong> &nbsp; {tt_count/total_leads:.1%} of leads &nbsp; · &nbsp; Score 3+ rate: <strong>{tt_s3:.1%}</strong></div>
    <div style="width:1px;background:#ccc;"></div>
    <div class="brand-stat">Touring Coach (TC) &nbsp; <strong>{tc_count:,}</strong> &nbsp; {tc_count/total_leads:.1%} of leads &nbsp; · &nbsp; Score 3+ rate: <strong>{tc_s3:.1%}</strong></div>
  </div>
</section>

<!-- ════════════════════════════════════════════════════════ SECTION 2 — ACTUALS -->
<section class="section">
  <div class="section-label">Section 2</div>
  <div class="section-title">Post-Launch Actuals vs. Forecast</div>
  <div class="section-sub">May 28 – June 18, 2026 &nbsp;·&nbsp; 22 days &nbsp;·&nbsp; {post_total:,} leads captured</div>

  <table class="comparison-table">
    <thead>
      <tr>
        <th style="width:28%">Metric</th>
        <th style="width:20%">Baseline (Before)</th>
        <th style="width:20%">Forecast (Expected)</th>
        <th style="width:32%">Actual (Post-Launch)</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="metric-name">Avg Leads / Month</td>
        <td>{monthly_avg:,.0f}</td>
        <td>−7%</td>
        <td class="actual-cell cell-{vc_monthly}">{post_monthly:,.0f} <span style="font-size:11px;color:#888">(−35.2% vs. baseline)</span></td>
      </tr>
      <tr>
        <td class="metric-name">Per Dealer / Month</td>
        <td>{per_dealer_mo:.0f}</td>
        <td>−7%</td>
        <td class="actual-cell cell-{vc_dealer}">{post_per_dealer:.0f} <span style="font-size:11px;color:#888">(−35.3% vs. baseline)</span></td>
      </tr>
      <tr>
        <td class="metric-name">Score 3+ Rate</td>
        <td>{score3plus_rate:.1%}</td>
        <td>Flat-to-up</td>
        <td class="actual-cell cell-{vc_score3}">{post_score3rate:.1%} <span style="font-size:11px;color:#888">(−30.7% vs. baseline)</span></td>
      </tr>
      <tr>
        <td class="metric-name">External Paid Share</td>
        <td>{paid_pct:.1%}</td>
        <td>~60%</td>
        <td class="actual-cell cell-{vc_paid}">{post_paid_pct:.1%} <span style="font-size:11px;color:#888">(+35.8% vs. baseline)</span></td>
      </tr>
    </tbody>
  </table>

</section>

<!-- ════════════════════════════════════════════════════════ FOOTER -->
<footer class="site-footer">
  <div class="footer-left">Element Three &nbsp;·&nbsp; Prepared June 2026</div>
  <div class="footer-right">Airstream Dealer Lead Distribution — Confidential</div>
</footer>

</div><!-- .page -->

<script>
const NAVY   = '#1B3A6B';
const SILVER = '#A8A9AD';
const ACCENT = '#4A7CC7';
const LIGHT  = '#D6E4F7';
const GREEN  = '#2E7D32';
const RED    = '#C62828';

Chart.defaults.font.family = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif";
Chart.defaults.font.size   = 12;
Chart.defaults.color       = '#555';

// ── Monthly Volume ──
new Chart(document.getElementById('monthlyChart'), {{
  type: 'bar',
  data: {{
    labels: {json.dumps(monthly_labels)},
    datasets: [{{
      label: 'Leads',
      data: {json.dumps(monthly_values)},
      backgroundColor: NAVY,
      borderRadius: 3,
      borderSkipped: false,
    }}]
  }},
  options: {{
    responsive: true,
    plugins: {{
      legend: {{ display: false }},
      annotation: {{ }}
    }},
    scales: {{
      x: {{ ticks: {{ font: {{ size: 9 }}, maxRotation: 45 }}, grid: {{ display: false }} }},
      y: {{
        ticks: {{ font: {{ size: 10 }} }},
        grid: {{ color: '#eee' }},
      }}
    }}
  }}
}});

// ── Source Split ──
new Chart(document.getElementById('sourceChart'), {{
  type: 'doughnut',
  data: {{
    labels: ['External Paid (Meta)', 'Website-Captured'],
    datasets: [{{
      data: [{paid_count}, {web_count}],
      backgroundColor: [NAVY, ACCENT],
      borderWidth: 0,
      hoverOffset: 6,
    }}]
  }},
  options: {{
    responsive: true,
    cutout: '62%',
    plugins: {{
      legend: {{
        position: 'bottom',
        labels: {{ font: {{ size: 11 }}, padding: 12, boxWidth: 12 }}
      }},
      tooltip: {{
        callbacks: {{
          label: function(ctx) {{
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            const pct = (ctx.raw/total*100).toFixed(1);
            return ` ${{ctx.label}}: ${{ctx.raw.toLocaleString()}} (${{pct}}%)`;
          }}
        }}
      }}
    }}
  }}
}});

// ── Score Distribution ──
new Chart(document.getElementById('scoreChart'), {{
  type: 'bar',
  data: {{
    labels: {json.dumps(score_labels)},
    datasets: [{{
      label: 'Leads',
      data: {json.dumps(score_values)},
      backgroundColor: {json.dumps(score_colors)},
      borderRadius: 3,
    }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{
        callbacks: {{
          label: function(ctx) {{
            const total = {json.dumps(score_values)}.reduce((a,b)=>a+b,0);
            const pct = (ctx.raw/total*100).toFixed(1);
            return ` ${{ctx.raw.toLocaleString()}} (${{pct}}%)`;
          }}
        }}
      }}
    }},
    scales: {{
      x: {{ ticks: {{ font: {{ size: 9 }} }}, grid: {{ color: '#eee' }} }},
      y: {{ ticks: {{ font: {{ size: 11 }} }}, grid: {{ display: false }} }}
    }}
  }}
}});

// ── Scenario Chart ──
const scenarioData = {json.dumps(scenarios)};
const scenLabels = scenarioData.map(s => s.label);
const scenVals   = scenarioData.map(s => s.total);
const scenColors = scenarioData.map(s => s.implemented ? NAVY : '#CBD5E0');
const scenBorders = scenarioData.map(s => s.implemented ? NAVY : '#CBD5E0');

new Chart(document.getElementById('scenarioChart'), {{
  type: 'bar',
  data: {{
    labels: scenLabels,
    datasets: [{{
      label: 'Monthly Leads',
      data: scenVals,
      backgroundColor: scenColors,
      borderColor: scenBorders,
      borderWidth: 1,
      borderRadius: 3,
    }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{
        callbacks: {{
          label: function(ctx) {{
            const pct = ((ctx.raw - {BASELINE_MONTHLY_PDF}) / {BASELINE_MONTHLY_PDF} * 100).toFixed(1);
            const sign = pct > 0 ? '+' : '';
            return ` ${{ctx.raw.toLocaleString()}}/mo  (${{sign}}${{pct}}% vs baseline)`;
          }}
        }}
      }}
    }},
    scales: {{
      x: {{
        ticks: {{ font: {{ size: 10 }} }},
        grid: {{ color: '#eee' }},
        min: 0,
        max: 22000,
      }},
      y: {{ ticks: {{ font: {{ size: 11 }} }}, grid: {{ display: false }} }}
    }}
  }},
  plugins: [{{
    afterDraw: function(chart) {{
      const ctx2 = chart.ctx;
      const xAxis = chart.scales.x;
      const baselineX = xAxis.getPixelForValue({BASELINE_MONTHLY_PDF});
      ctx2.save();
      ctx2.beginPath();
      ctx2.moveTo(baselineX, chart.chartArea.top);
      ctx2.lineTo(baselineX, chart.chartArea.bottom);
      ctx2.strokeStyle = ACCENT;
      ctx2.lineWidth = 1.5;
      ctx2.setLineDash([5, 4]);
      ctx2.stroke();
      ctx2.fillStyle = ACCENT;
      ctx2.font = 'bold 10px sans-serif';
      ctx2.fillText('Baseline ({BASELINE_MONTHLY_PDF:,})', baselineX + 4, chart.chartArea.top + 12);
      ctx2.restore();
    }}
  }}]
}});

}});
</script>
</body>
</html>"""

OUTPUT = '/home/user/airstream-lead-distro-analysis-2/airstream_lead_report.html'
chartjs = open('/home/user/airstream-lead-distro-analysis-2/chartjs.min.js').read()
html = html.replace('<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>', '<script>' + chartjs + '</script>')
with open(OUTPUT, 'w') as f:
    f.write(html)

import os
size = os.path.getsize(OUTPUT)
print(f"Written: {OUTPUT}  ({size:,} bytes / {size/1024:.1f} KB)")
