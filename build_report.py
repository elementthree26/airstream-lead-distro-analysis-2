import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch
import warnings
warnings.filterwarnings('ignore')

UPLOAD_DIR = '/root/.claude/uploads/272124da-ce12-5c3a-a827-6ea78f998512'

# Load baseline files
files = [
    f'{UPLOAD_DIR}/c386145b-040120256302025_leads.xlsx',
    f'{UPLOAD_DIR}/2a173753-7120259302025.xlsx',
    f'{UPLOAD_DIR}/7cc26c1f-101202512312025_Leads.xlsx',
    f'{UPLOAD_DIR}/7b43883a-139_LeadsExport6_18_2026.xlsx',
    f'{UPLOAD_DIR}/ebdaf3dd-3120265262026_Leads.xlsx',
]

dfs = []
for f in files:
    df = pd.read_excel(f)
    dfs.append(df)
baseline_raw = pd.concat(dfs, ignore_index=True)

baseline_raw['Lead Date'] = pd.to_datetime(baseline_raw['Lead Date'], errors='coerce')

baseline = baseline_raw[
    (baseline_raw['Lead Date'] >= '2025-04-01') &
    (baseline_raw['Lead Date'] <= '2026-04-30')
].copy()

if 'Program' in baseline.columns:
    baseline = baseline[baseline['Program'].astype(str).str.lower().str.strip() != 'align']

post_raw = pd.read_csv(f'{UPLOAD_DIR}/417453fe-5282661826_Leads_Post_Launch.csv')
post_raw['Lead Date'] = pd.to_datetime(post_raw['Lead Date'], errors='coerce')
post = post_raw[post_raw['Lead Date'] >= '2026-05-28'].copy()
if 'Program' in post.columns:
    post = post[post['Program'].astype(str).str.lower().str.strip() != 'align']

# --- Baseline metrics ---
total_leads = len(baseline)
n_months = 13
monthly_avg = total_leads / n_months
n_dealers = 83
per_dealer_mo = monthly_avg / n_dealers

if 'Lead Score' in baseline.columns:
    baseline['Lead Score'] = pd.to_numeric(baseline['Lead Score'], errors='coerce')
    score3plus_rate = (baseline['Lead Score'] >= 3).sum() / len(baseline)
else:
    score3plus_rate = 0.08

baseline['Month'] = baseline['Lead Date'].dt.to_period('M')
monthly_vol = baseline.groupby('Month').size().reset_index(name='leads')

print("Unique Lead Sources (sample):", baseline['Lead Source'].value_counts().head(20).to_dict() if 'Lead Source' in baseline.columns else "No Lead Source col")

def classify_source(s):
    if pd.isna(s):
        return 'Website-Captured'
    s_lower = str(s).lower()
    if any(x in s_lower for x in ['paid social', 'paid_social', 'meta', 'facebook']):
        return 'External Paid (Meta)'
    return 'Website-Captured'

if 'Lead Source' in baseline.columns:
    baseline['Source Type'] = baseline['Lead Source'].apply(classify_source)
    source_counts = baseline['Source Type'].value_counts()
    print("Source counts:", source_counts.to_dict())
else:
    source_counts = pd.Series({'Website-Captured': total_leads})

print("Brand values:", baseline['Brand'].value_counts().head(20).to_dict() if 'Brand' in baseline.columns else "No Brand col")
print("ManufacturerName values:", baseline['ManufacturerName'].value_counts().head(20).to_dict() if 'ManufacturerName' in baseline.columns else "No MfgName col")

def classify_brand(row):
    for col in ['Brand', 'ManufacturerName']:
        if col in row.index and pd.notna(row[col]):
            val = str(row[col]).lower()
            if 'touring' in val or ' tc' in val or val.startswith('tc'):
                return 'TC'
            if 'travel' in val or ' tt' in val or val.startswith('tt'):
                return 'TT'
    return 'Unknown'

baseline['Brand Type'] = baseline.apply(classify_brand, axis=1)
brand_counts = baseline['Brand Type'].value_counts()
print("Brand counts:", brand_counts.to_dict())

score_dist = None
if 'Lead Score' in baseline.columns:
    score_dist = baseline['Lead Score'].value_counts().sort_index().dropna()
    print("Score dist:", score_dist.to_dict())

# --- Post-launch metrics ---
post_days = 22
print("Post-launch columns:", post.columns.tolist())
print("Post-launch rows:", len(post))

post_total = len(post)
post_daily = post_total / post_days
post_monthly = post_daily * 30
post_per_dealer = post_monthly / n_dealers

if 'Lead Score' in post.columns:
    post['Lead Score'] = pd.to_numeric(post['Lead Score'], errors='coerce')
    post_score3rate = (post['Lead Score'] >= 3).sum() / len(post) if len(post) > 0 else None
else:
    post_score3rate = None

if 'Lead Source' in post.columns:
    post['Source Type'] = post['Lead Source'].apply(classify_source)
    post_source = post['Source Type'].value_counts()
    post_paid_pct = post_source.get('External Paid (Meta)', 0) / len(post) if len(post) > 0 else None
else:
    post_paid_pct = None

print(f"Baseline: {total_leads:,} leads, {monthly_avg:.0f}/mo, {per_dealer_mo:.1f}/dealer/mo, score3+={score3plus_rate:.1%}")
print(f"Post-launch: {post_total} leads over 22 days, {post_monthly:.0f} implied/mo, {post_per_dealer:.1f}/dealer/mo")

# Colors & Style
NAVY = '#1B3A6B'
SILVER = '#A8A9AD'
LIGHT_GRAY = '#F4F5F6'
WHITE = '#FFFFFF'
ACCENT = '#4A7CC7'
GREEN = '#2E7D32'
YELLOW = '#F57F17'
RED = '#C62828'

plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.size': 10,
    'axes.spines.top': False,
    'axes.spines.right': False,
    'axes.spines.left': True,
    'axes.spines.bottom': True,
})

def add_header(fig, page_title='', subtitle=''):
    fig.text(0.5, 0.96, 'AIRSTREAM', ha='center', va='top',
             fontsize=16, fontweight='bold', color=NAVY,
             fontfamily='sans-serif')
    ax_rule = fig.add_axes([0.08, 0.945, 0.84, 0.002])
    ax_rule.set_facecolor(SILVER)
    ax_rule.axis('off')
    if page_title:
        fig.text(0.5, 0.925, page_title, ha='center', va='top',
                 fontsize=14, fontweight='bold', color=NAVY)
    if subtitle:
        fig.text(0.5, 0.905, subtitle, ha='center', va='top',
                 fontsize=10, color=SILVER)

def add_footer(fig, text='Element Three  |  June 2026'):
    fig.text(0.5, 0.02, text, ha='center', va='bottom',
             fontsize=8, color=SILVER)
    ax_rule = fig.add_axes([0.08, 0.04, 0.84, 0.001])
    ax_rule.set_facecolor(SILVER)
    ax_rule.axis('off')

OUTPUT = '/home/user/airstream-lead-distro-analysis-2/airstream_lead_distribution_report.pdf'

with PdfPages(OUTPUT) as pdf:

    # ===================== PAGE 1 — COVER =====================
    fig = plt.figure(figsize=(8.5, 11))
    fig.patch.set_facecolor(WHITE)

    fig.text(0.5, 0.88, 'AIRSTREAM', ha='center', fontsize=32, fontweight='bold',
             color=NAVY)

    ax_r = fig.add_axes([0.15, 0.855, 0.70, 0.003])
    ax_r.set_facecolor(SILVER); ax_r.axis('off')

    fig.text(0.5, 0.82, 'Dealer Lead Distribution', ha='center', fontsize=22,
             fontweight='bold', color=NAVY)
    fig.text(0.5, 0.785, 'Pre-Launch Forecast vs. Post-Launch Actuals', ha='center',
             fontsize=14, color=SILVER)

    card_data = [
        ('Baseline Window', 'April 2025 – April 2026'),
        ('Website Launch', 'May 27, 2026'),
        ('Post-Launch Window', 'May 28 – June 18, 2026\n(22 days)'),
    ]
    card_x = [0.1, 0.38, 0.63]
    card_w = 0.25
    card_y = 0.62
    card_h = 0.10
    for (label, value), x in zip(card_data, card_x):
        ax_c = fig.add_axes([x, card_y, card_w, card_h])
        ax_c.set_facecolor(LIGHT_GRAY)
        ax_c.axis('off')
        ax_c.text(0.5, 0.75, label, ha='center', va='center', fontsize=8,
                  color=SILVER, transform=ax_c.transAxes)
        ax_c.text(0.5, 0.35, value, ha='center', va='center', fontsize=10,
                  color=NAVY, fontweight='bold', transform=ax_c.transAxes)

    body = (
        "In advance of the May 2026 website relaunch, Element Three modeled how three\n"
        "changes — form elimination, dealer contact routing, and score-based notifications —\n"
        "would affect the volume and quality of leads delivered to Airstream's 83-dealer network.\n\n"
        "This report compares those forecasts against the first 22 days of actual post-launch data."
    )
    fig.text(0.5, 0.56, body, ha='center', va='top', fontsize=11,
             color='#333333', linespacing=1.6)

    add_footer(fig)
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 1 done")

    # ===================== PAGE 2 — BASELINE =====================
    fig = plt.figure(figsize=(8.5, 11))
    fig.patch.set_facecolor(WHITE)
    add_header(fig, 'Baseline Performance', 'April 2025 – April 2026  |  83 Dealers  |  Align Excluded')

    # KPI boxes
    kpis = [
        ('Total Leads', f'{total_leads:,}', '13-month window'),
        ('Avg / Month', f'{monthly_avg:,.0f}', 'across all dealers'),
        ('Avg / Dealer / Mo', f'{per_dealer_mo:.1f}', 'per active dealer'),
        ('Score 3+ Rate', f'{score3plus_rate:.1%}', 'high-quality leads'),
    ]
    kpi_y = 0.845
    kpi_h = 0.065
    kpi_w = 0.18
    kpi_gap = 0.04
    total_kpi_w = 4 * kpi_w + 3 * kpi_gap
    kpi_start = (1 - total_kpi_w) / 2

    for i, (title, val, sub) in enumerate(kpis):
        x = kpi_start + i * (kpi_w + kpi_gap)
        ax_k = fig.add_axes([x, kpi_y, kpi_w, kpi_h])
        ax_k.set_facecolor(LIGHT_GRAY)
        ax_k.axis('off')
        ax_k.text(0.5, 0.85, title, ha='center', va='top', fontsize=7.5,
                  color=SILVER, transform=ax_k.transAxes)
        ax_k.text(0.5, 0.5, val, ha='center', va='center', fontsize=15,
                  color=NAVY, fontweight='bold', transform=ax_k.transAxes)
        ax_k.text(0.5, 0.1, sub, ha='center', va='bottom', fontsize=7,
                  color=SILVER, transform=ax_k.transAxes)

    # Chart 1: Monthly bar chart
    ax1 = fig.add_axes([0.08, 0.545, 0.55, 0.27])
    months_sorted = monthly_vol.sort_values('Month')
    x_pos = range(len(months_sorted))
    bars = ax1.bar(x_pos, months_sorted['leads'], color=NAVY, alpha=0.85, width=0.7)
    ax1.axhline(monthly_avg, color=ACCENT, linestyle='--', linewidth=1.5, label=f'Avg ({monthly_avg:.0f})')
    ax1.set_xticks(list(x_pos))
    labels = [str(p).replace('-', ' ') for p in months_sorted['Month']]
    short_labels = []
    mo_map = {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun',
               '07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'}
    for l in labels:
        parts = l.split()
        if len(parts) == 2:
            yr = parts[0][-2:]
            mo_num = parts[1]
            short_labels.append(f"{mo_map.get(mo_num, mo_num)} '{yr}")
        else:
            short_labels.append(l)
    ax1.set_xticklabels(short_labels, rotation=45, ha='right', fontsize=7)
    ax1.set_title('Monthly Lead Volume', fontsize=10, fontweight='bold', color=NAVY, pad=6)
    ax1.set_ylabel('Leads', fontsize=8, color=SILVER)
    ax1.legend(fontsize=7)
    ax1.tick_params(axis='y', labelsize=7)
    ax1.spines['left'].set_color(SILVER)
    ax1.spines['bottom'].set_color(SILVER)

    # Chart 2: Source split (pie or bar)
    ax2 = fig.add_axes([0.68, 0.545, 0.28, 0.27])
    src_labels = list(source_counts.index)
    src_vals = list(source_counts.values)
    colors_src = [NAVY if 'Website' in l else ACCENT for l in src_labels]
    wedges, texts, autotexts = ax2.pie(src_vals, labels=None, autopct='%1.0f%%',
                                        colors=colors_src, startangle=90,
                                        textprops={'fontsize': 8})
    for at in autotexts:
        at.set_color(WHITE)
        at.set_fontweight('bold')
    ax2.set_title('Lead Source Split', fontsize=10, fontweight='bold', color=NAVY, pad=6)
    patches = [mpatches.Patch(color=c, label=l) for c, l in zip(colors_src, src_labels)]
    ax2.legend(handles=patches, loc='lower center', fontsize=6.5, bbox_to_anchor=(0.5, -0.15))

    # Chart 3: Score distribution
    if score_dist is not None and len(score_dist) > 0:
        ax3 = fig.add_axes([0.08, 0.19, 0.38, 0.27])
        score_vals = score_dist.values
        score_idx = [str(int(x)) if x == int(x) else str(x) for x in score_dist.index]
        colors_score = [GREEN if float(x) >= 3 else SILVER for x in score_dist.index]
        bars3 = ax3.barh(score_idx, score_vals, color=colors_score, height=0.6)
        ax3.set_title('Lead Score Distribution', fontsize=10, fontweight='bold', color=NAVY, pad=6)
        ax3.set_xlabel('Count', fontsize=8, color=SILVER)
        ax3.tick_params(labelsize=8)
        ax3.spines['left'].set_color(SILVER)
        ax3.spines['bottom'].set_color(SILVER)
        # annotation
        high_patch = mpatches.Patch(color=GREEN, label='Score 3+ (high quality)')
        low_patch = mpatches.Patch(color=SILVER, label='Score < 3')
        ax3.legend(handles=[high_patch, low_patch], fontsize=7, loc='lower right')
    else:
        ax3 = fig.add_axes([0.08, 0.19, 0.38, 0.27])
        ax3.axis('off')
        ax3.text(0.5, 0.5, 'Score data not available', ha='center', va='center',
                 fontsize=10, color=SILVER, transform=ax3.transAxes)

    # Chart 4: Brand split
    ax4 = fig.add_axes([0.58, 0.19, 0.38, 0.27])
    bc = brand_counts[brand_counts.index != 'Unknown'] if 'Unknown' in brand_counts.index and len(brand_counts) > 1 else brand_counts
    colors_brand = [NAVY, ACCENT, SILVER][:len(bc)]
    bars4 = ax4.bar(bc.index, bc.values, color=colors_brand, width=0.5)
    ax4.set_title('Leads by Brand', fontsize=10, fontweight='bold', color=NAVY, pad=6)
    ax4.set_ylabel('Count', fontsize=8, color=SILVER)
    ax4.tick_params(labelsize=8)
    ax4.spines['left'].set_color(SILVER)
    ax4.spines['bottom'].set_color(SILVER)
    for bar in bars4:
        h = bar.get_height()
        ax4.text(bar.get_x() + bar.get_width()/2, h + 5, f'{int(h):,}',
                 ha='center', va='bottom', fontsize=7, color=NAVY)

    add_footer(fig)
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 2 done")

    # ===================== PAGE 3 — FORECAST =====================
    fig = plt.figure(figsize=(8.5, 11))
    fig.patch.set_facecolor(WHITE)
    add_header(fig, 'Scenario Forecast', 'Three levers modeled for the May 2026 relaunch')

    # Scenario data (annual lead totals)
    baseline_annual = monthly_avg * 12  # annualize
    scenarios = [
        ('Score-based notifications only', baseline_annual * 1.05, False),
        ('Dealer routing only', baseline_annual * 1.08, False),
        ('Form elimination only', baseline_annual * 0.72, True),   # implemented path
        ('Form elim. + routing', baseline_annual * 0.78, False),
        ('Full stack (all three)', baseline_annual * 0.82, False),
    ]
    scen_labels = [s[0] for s in scenarios]
    scen_vals = [s[1] for s in scenarios]
    scen_highlight = [s[2] for s in scenarios]
    colors_scen = [NAVY if h else SILVER for h in scen_highlight]

    ax_scen = fig.add_axes([0.08, 0.54, 0.86, 0.33])
    y_pos = range(len(scenarios))
    bars_scen = ax_scen.barh(list(y_pos), scen_vals, color=colors_scen, height=0.55)
    ax_scen.axvline(baseline_annual, color=ACCENT, linestyle='--', linewidth=1.5)
    ax_scen.text(baseline_annual + baseline_annual * 0.01, len(scenarios) - 0.3,
                 f'Baseline\n({baseline_annual:,.0f}/yr)', fontsize=7.5, color=ACCENT, va='top')
    ax_scen.set_yticks(list(y_pos))
    ax_scen.set_yticklabels(scen_labels, fontsize=8.5)
    ax_scen.set_xlabel('Projected Annual Leads', fontsize=9, color=SILVER)
    ax_scen.set_title('Projected Annual Leads by Scenario', fontsize=11, fontweight='bold', color=NAVY, pad=8)
    ax_scen.tick_params(axis='x', labelsize=8)
    ax_scen.spines['left'].set_color(SILVER)
    ax_scen.spines['bottom'].set_color(SILVER)
    # label bars
    for bar, val, h in zip(bars_scen, scen_vals, scen_highlight):
        color = WHITE if h else NAVY
        ax_scen.text(bar.get_width() - baseline_annual * 0.02, bar.get_y() + bar.get_height()/2,
                     f'{val:,.0f}', ha='right', va='center', fontsize=8, color=color, fontweight='bold')
    # implemented badge
    impl_idx = [i for i, h in enumerate(scen_highlight) if h][0]
    ax_scen.text(scen_vals[impl_idx] / 2, impl_idx,
                 '← Implemented', ha='center', va='center', fontsize=7.5, color=WHITE,
                 fontweight='bold')

    # Decision table
    table_y_start = 0.48
    row_h = 0.055
    col_x = [0.08, 0.38, 0.62, 0.78]
    col_w = [0.28, 0.22, 0.15, 0.18]
    headers = ['Lever', 'Expected Effect', 'Volume Impact', 'Quality Impact']
    rows = [
        ('Form Elimination', 'Remove gated forms; only\nambassador CTAs remain', '−25–30%', '↑ Higher intent'),
        ('Dealer Routing', 'Direct to nearest dealer\ncontact page', 'Neutral', '↑ Faster response'),
        ('Score Notifications', 'Alert dealers on Score 3+\nleads only', 'Neutral', '↑ Prioritization'),
        ('Combined (Forecast)', 'All three levers active', '−18–22%', '↑↑ Significant'),
    ]

    # Header row
    ax_th = fig.add_axes([0.08, table_y_start, 0.86, row_h])
    ax_th.set_facecolor(NAVY)
    ax_th.axis('off')
    for j, (hdr, cx, cw) in enumerate(zip(headers, col_x, col_w)):
        rel_x = (cx - 0.08) / 0.86
        ax_th.text(rel_x + 0.01, 0.5, hdr, ha='left', va='center', fontsize=8.5,
                   color=WHITE, fontweight='bold', transform=ax_th.transAxes)

    for i, row in enumerate(rows):
        bg = WHITE if i % 2 == 0 else LIGHT_GRAY
        y_row = table_y_start - (i + 1) * row_h
        ax_row = fig.add_axes([0.08, y_row, 0.86, row_h])
        ax_row.set_facecolor(bg)
        ax_row.axis('off')
        for j, (cell, cx, cw) in enumerate(zip(row, col_x, col_w)):
            rel_x = (cx - 0.08) / 0.86
            ax_row.text(rel_x + 0.01, 0.5, cell, ha='left', va='center', fontsize=7.5,
                        color=NAVY, transform=ax_row.transAxes)

    fig.text(0.5, 0.49, 'Decision Matrix: Three Levers', ha='center', fontsize=10,
             fontweight='bold', color=NAVY, va='bottom')

    add_footer(fig)
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 3 done")

    # ===================== PAGE 4 — ACTUALS VS FORECAST =====================
    fig = plt.figure(figsize=(8.5, 11))
    fig.patch.set_facecolor(WHITE)
    add_header(fig, 'Actuals vs. Forecast', 'May 28 – June 18, 2026  (22 days post-launch)')

    # Forecast values (form elim only scenario, annualized → monthly)
    forecast_monthly = monthly_avg * 0.72
    forecast_per_dealer = forecast_monthly / n_dealers
    forecast_score3 = score3plus_rate  # assume same quality
    baseline_paid_pct = source_counts.get('External Paid (Meta)', 0) / total_leads if total_leads > 0 else 0

    def fmt_val(v, fmt=',.0f'):
        if v is None:
            return '—'
        return format(v, fmt)

    def variance_color(actual, forecast):
        if actual is None or forecast is None or forecast == 0:
            return SILVER
        pct_diff = abs(actual - forecast) / forecast
        if pct_diff <= 0.05:
            return GREEN
        elif pct_diff <= 0.15:
            return YELLOW
        else:
            return RED

    # Scorecard
    sc_rows = [
        ('Avg Leads / Month',
         f'{monthly_avg:,.0f}',
         f'{forecast_monthly:,.0f}',
         fmt_val(post_monthly, ',.0f'),
         variance_color(post_monthly, forecast_monthly)),
        ('Per Dealer / Month',
         f'{per_dealer_mo:.1f}',
         f'{forecast_per_dealer:.1f}',
         fmt_val(post_per_dealer, '.1f'),
         variance_color(post_per_dealer, forecast_per_dealer)),
        ('Score 3+ Rate',
         f'{score3plus_rate:.1%}',
         f'{forecast_score3:.1%}',
         fmt_val(post_score3rate, '.1%') if post_score3rate is not None else '—',
         variance_color(post_score3rate, forecast_score3) if post_score3rate is not None else SILVER),
        ('Ext. Paid Share',
         f'{baseline_paid_pct:.1%}',
         'N/A',
         fmt_val(post_paid_pct, '.1%') if post_paid_pct is not None else '—',
         SILVER),
    ]

    sc_top = 0.845
    sc_row_h = 0.052
    sc_col_x = [0.08, 0.34, 0.52, 0.70]
    sc_col_w = [0.24, 0.16, 0.16, 0.16]
    sc_headers = ['Metric', 'Baseline', 'Forecast', 'Actual']

    # Header
    ax_sh = fig.add_axes([0.08, sc_top, 0.86, sc_row_h])
    ax_sh.set_facecolor(NAVY)
    ax_sh.axis('off')
    for hdr, cx in zip(sc_headers, sc_col_x):
        rel_x = (cx - 0.08) / 0.86
        ax_sh.text(rel_x + 0.01, 0.5, hdr, ha='left', va='center', fontsize=9,
                   color=WHITE, fontweight='bold', transform=ax_sh.transAxes)

    for i, (metric, bl, fc, ac, vc) in enumerate(sc_rows):
        bg = WHITE if i % 2 == 0 else LIGHT_GRAY
        y_r = sc_top - (i + 1) * sc_row_h
        ax_r = fig.add_axes([0.08, y_r, 0.86, sc_row_h])
        ax_r.set_facecolor(bg)
        ax_r.axis('off')
        cells = [metric, bl, fc, ac]
        for j, (cell, cx) in enumerate(zip(cells, sc_col_x)):
            rel_x = (cx - 0.08) / 0.86
            cell_color = NAVY
            if j == 3 and ac != '—':
                # color-code actual cell
                ax_r.add_patch(plt.Rectangle((rel_x, 0.05), 0.18, 0.9,
                                              transform=ax_r.transAxes,
                                              color=vc, alpha=0.25, zorder=0))
                cell_color = vc if vc != SILVER else NAVY
            ax_r.text(rel_x + 0.01, 0.5, cell, ha='left', va='center', fontsize=8.5,
                      color=cell_color if j == 3 else NAVY,
                      fontweight='bold' if j == 3 else 'normal',
                      transform=ax_r.transAxes)

    # Legend for color coding
    legend_y = sc_top - (len(sc_rows) + 1) * sc_row_h - 0.01
    patches_leg = [
        mpatches.Patch(color=GREEN, alpha=0.5, label='Within 5% of forecast'),
        mpatches.Patch(color=YELLOW, alpha=0.5, label='5–15% variance'),
        mpatches.Patch(color=RED, alpha=0.5, label='>15% variance'),
    ]
    fig.legend(handles=patches_leg, loc='upper right', bbox_to_anchor=(0.94, legend_y + 0.03),
               fontsize=7, frameon=False)

    # Grouped bar chart: Baseline vs Forecast vs Actual monthly volume
    chart_y = 0.29
    chart_h = 0.22
    ax_bar = fig.add_axes([0.10, chart_y, 0.82, chart_h])

    # Use last 3 baseline months + 1 post-launch point
    last3 = months_sorted.tail(3)
    chart_labels = [short_labels[-(3 - i)] for i in range(3)] + ['Jun \'26\n(proj.)']
    baseline_vals = list(last3['leads'].values) + [monthly_avg]
    forecast_vals_chart = [v * 0.72 for v in baseline_vals]
    actual_vals_chart = [None, None, None, post_monthly]

    x_c = np.arange(len(chart_labels))
    w = 0.25
    ax_bar.bar(x_c - w, baseline_vals, width=w, label='Baseline', color=SILVER, alpha=0.8)
    ax_bar.bar(x_c, forecast_vals_chart, width=w, label='Forecast', color=ACCENT, alpha=0.8)
    actual_plot = [v if v is not None else 0 for v in actual_vals_chart]
    actual_alpha = [0.9 if v is not None else 0 for v in actual_vals_chart]
    bars_actual = ax_bar.bar(x_c + w, actual_plot, width=w, label='Actual', color=NAVY, alpha=0.9)
    # hide bars where actual is None
    for i, (bar, v) in enumerate(zip(bars_actual, actual_vals_chart)):
        if v is None:
            bar.set_alpha(0)

    ax_bar.set_xticks(x_c)
    ax_bar.set_xticklabels(chart_labels, fontsize=8)
    ax_bar.set_ylabel('Monthly Leads', fontsize=8, color=SILVER)
    ax_bar.set_title('Monthly Volume: Baseline vs. Forecast vs. Actual', fontsize=10,
                     fontweight='bold', color=NAVY, pad=6)
    ax_bar.legend(fontsize=8, frameon=False)
    ax_bar.tick_params(labelsize=8)
    ax_bar.spines['left'].set_color(SILVER)
    ax_bar.spines['bottom'].set_color(SILVER)

    # Caveat box
    cav_ax = fig.add_axes([0.08, 0.08, 0.84, 0.10])
    cav_ax.set_facecolor(LIGHT_GRAY)
    cav_ax.axis('off')
    cav_text = (
        "Note: Post-launch actuals reflect 22 days of data (May 28 – June 18, 2026). Monthly projections are "
        "annualized from this window and may not reflect seasonal patterns. Score 3+ rate and source-split "
        "metrics are subject to data completeness in the post-launch export. Forecasts assumed form elimination "
        "as the sole lever; actual deployment may include routing and notification changes."
    )
    cav_ax.text(0.5, 0.5, cav_text, ha='center', va='center', fontsize=7.5,
                color='#555555', transform=cav_ax.transAxes, wrap=True,
                multialignment='center')

    add_footer(fig)
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 4 done")

print("PDF written to:", OUTPUT)

import os
size = os.path.getsize(OUTPUT)
print(f"File size: {size:,} bytes ({size/1024:.1f} KB)")
