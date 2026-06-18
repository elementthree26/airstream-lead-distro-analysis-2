#!/usr/bin/env python3
"""
Airstream Dealer Lead Distribution Analysis Report
Generates a 4-page PDF using matplotlib PdfPages
"""

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

# ─── BRAND COLORS ────────────────────────────────────────────────────────────
NAVY   = '#1B3A6B'
SILVER = '#A8A9AD'
WHITE  = '#FFFFFF'
LGRAY  = '#F4F5F6'
DGRAY  = '#6B6B6B'
GREEN  = '#2E7D32'
YELLOW = '#F9A825'
RED    = '#C62828'

# ─── PAGE SIZE ────────────────────────────────────────────────────────────────
W, H = 8.5, 11.0   # letter

# ─── FILE PATHS ──────────────────────────────────────────────────────────────
BASE = '/root/.claude/uploads/272124da-ce12-5c3a-a827-6ea78f998512/'
FILES = {
    'apr_jun_2025':  (BASE + 'c386145b-040120256302025_leads.xlsx',         'xlsx'),
    'jul_sep_2025':  (BASE + '2a173753-7120259302025.xlsx',                  'xlsx'),
    'oct_dec_2025':  (BASE + '7cc26c1f-101202512312025_Leads.xlsx',         'xlsx'),
    'jan_feb_2026':  (BASE + '7b43883a-139_LeadsExport6_18_2026.xlsx',      'xlsx'),
    'mar_apr_2026':  (BASE + 'ebdaf3dd-3120265262026_Leads.xlsx',           'xlsx'),
    'post_launch':   (BASE + '417453fe-5282661826_Leads_Post_Launch.csv',   'csv'),
}

OUT = '/home/user/airstream-lead-distro-analysis-2/airstream_lead_distribution_report.pdf'

# ─── LOAD DATA ───────────────────────────────────────────────────────────────
def load_file(path, ftype):
    if ftype == 'xlsx':
        return pd.read_excel(path)
    else:
        return pd.read_csv(path)

print("Loading data files…")
dfs = {}
for key, (path, ftype) in FILES.items():
    print(f"  {key}: {path.split('/')[-1]}")
    dfs[key] = load_file(path, ftype)

# ─── PARSE DATES & COMBINE BASELINE ─────────────────────────────────────────
def parse_lead_date(df):
    df = df.copy()
    df['Lead Date'] = pd.to_datetime(df['Lead Date'], errors='coerce')
    return df

# Combine all xlsx baseline files
baseline_parts = []
for key in ['apr_jun_2025', 'jul_sep_2025', 'oct_dec_2025', 'jan_feb_2026', 'mar_apr_2026']:
    df = parse_lead_date(dfs[key])
    baseline_parts.append(df)

raw_baseline = pd.concat(baseline_parts, ignore_index=True)

# Remove Align program rows
if 'Program' in raw_baseline.columns:
    before = len(raw_baseline)
    raw_baseline = raw_baseline[raw_baseline['Program'].fillna('') != 'Align']
    print(f"  Removed {before - len(raw_baseline)} Align rows from baseline")

# Filter to Apr 1 2025 – Apr 30 2026
BL_START = pd.Timestamp('2025-04-01')
BL_END   = pd.Timestamp('2026-04-30')
baseline = raw_baseline[(raw_baseline['Lead Date'] >= BL_START) &
                         (raw_baseline['Lead Date'] <= BL_END)].copy()

# Drop duplicates by Lead Date + email + dealer to avoid double-counting from overlapping exports
dedup_cols = ['Lead Date', 'Email', 'Dealer']
available_dedup = [c for c in dedup_cols if c in baseline.columns]
baseline = baseline.drop_duplicates(subset=available_dedup)
print(f"  Baseline rows after dedup & date filter: {len(baseline)}")

# ─── POST-LAUNCH ─────────────────────────────────────────────────────────────
post = parse_lead_date(dfs['post_launch'])
if 'Program' in post.columns:
    post = post[post['Program'].fillna('') != 'Align']
PL_START = pd.Timestamp('2026-05-28')
PL_END   = pd.Timestamp('2026-06-18')
post = post[(post['Lead Date'] >= PL_START) & (post['Lead Date'] <= PL_END)].copy()
print(f"  Post-launch rows: {len(post)}")

# ─── SOURCE CLASSIFICATION ───────────────────────────────────────────────────
def classify_source(df):
    src = df['Lead Source'].fillna('').str.lower()
    mask_paid = src.str.contains('paid social|meta|facebook', regex=True)
    df = df.copy()
    df['source_class'] = np.where(mask_paid, 'External Paid (Meta)', 'Website-Captured')
    return df

baseline = classify_source(baseline)
post = classify_source(post)

# ─── MONTHLY AGGREGATION ─────────────────────────────────────────────────────
baseline['YearMonth'] = baseline['Lead Date'].dt.to_period('M')
monthly = baseline.groupby('YearMonth').size().reset_index(name='leads')
monthly['YearMonth_dt'] = monthly['YearMonth'].dt.to_timestamp()
monthly = monthly.sort_values('YearMonth_dt')

# ─── BASELINE KPIs ───────────────────────────────────────────────────────────
DEALERS = 83
total_bl  = len(baseline)
months_bl = 13  # Apr 2025 – Apr 2026
monthly_avg_bl = total_bl / months_bl
per_dealer_bl  = monthly_avg_bl / DEALERS

score3_bl = baseline[baseline['Lead Score'] >= 3]
score3_rate_bl = len(score3_bl) / total_bl if total_bl > 0 else 0

# Brand split
if 'Brand' in baseline.columns:
    brand_counts = baseline['Brand'].value_counts()
    tt_count = brand_counts.get('TT', 0)
    tc_count = brand_counts.get('TC', 0)
    # also check for full names
    if tt_count == 0:
        tt_count = sum(v for k, v in brand_counts.items() if 'touring' in str(k).lower() or k == 'TT')
    tt_rate = len(baseline[(baseline['Brand'].isin(['TT'])) & (baseline['Lead Score'] >= 3)]) / max(1, tt_count)
    tc_rate = len(baseline[(baseline['Brand'].isin(['TC'])) & (baseline['Lead Score'] >= 3)]) / max(1, tc_count)
    tt_pct = tt_count / total_bl * 100
    tc_pct = tc_count / total_bl * 100
else:
    tt_count = tc_count = 0
    tt_pct = tc_pct = 0
    tt_rate = tc_rate = 0

# Score distribution
score_dist = baseline['Lead Score'].value_counts().sort_index()
all_scores = range(1, 6)
score_counts = [score_dist.get(s, 0) for s in all_scores]
score_pcts   = [c / total_bl * 100 for c in score_counts]

# Source split
src_split = baseline['source_class'].value_counts()
paid_count = src_split.get('External Paid (Meta)', 0)
web_count  = src_split.get('Website-Captured', 0)

# ─── POST-LAUNCH KPIs ────────────────────────────────────────────────────────
total_pl = len(post)
DAYS_PL  = 22
daily_rate_pl   = total_pl / DAYS_PL
monthly_rate_pl = daily_rate_pl * 30
per_dealer_pl   = monthly_rate_pl / DEALERS

score3_pl = post[post['Lead Score'] >= 3]
score3_rate_pl = len(score3_pl) / total_pl if total_pl > 0 else 0

src_split_pl = post['source_class'].value_counts()
paid_pct_pl = src_split_pl.get('External Paid (Meta)', 0) / max(1, total_pl) * 100
paid_pct_bl = paid_count / max(1, total_bl) * 100

print(f"\n── BASELINE ──")
print(f"  Total leads:        {total_bl:,}")
print(f"  Monthly avg:        {monthly_avg_bl:,.0f}")
print(f"  Per dealer/mo:      {per_dealer_bl:.1f}")
print(f"  Score 3+ rate:      {score3_rate_bl:.1%}")
print(f"\n── POST-LAUNCH ──")
print(f"  Total leads (22d):  {total_pl:,}")
print(f"  Daily rate:         {daily_rate_pl:.1f}")
print(f"  Implied monthly:    {monthly_rate_pl:,.0f}")
print(f"  Score 3+ rate:      {score3_rate_pl:.1%}")

# ─── HELPERS ─────────────────────────────────────────────────────────────────
def add_header(fig, title="AIRSTREAM"):
    """Add AIRSTREAM wordmark and thin silver rule at top of page."""
    ax = fig.add_axes([0.06, 0.94, 0.88, 0.04])
    ax.axis('off')
    ax.text(0.0, 0.9, title, transform=ax.transAxes,
            fontsize=14, fontweight='bold', color=NAVY,
            fontfamily='sans-serif', va='top', ha='left',
            fontvariant='small-caps', letterspacing=4)
    # Silver rule
    ax.axhline(y=0.05, xmin=0, xmax=1, color=SILVER, linewidth=1.2)

def kpi_box(ax, label, value, subtitle=None):
    """Draw a KPI box with navy bg, white text."""
    ax.set_facecolor(NAVY)
    ax.axis('off')
    ax.text(0.5, 0.62, str(value), transform=ax.transAxes,
            fontsize=22, fontweight='bold', color=WHITE,
            ha='center', va='center')
    ax.text(0.5, 0.25, label, transform=ax.transAxes,
            fontsize=8, color=SILVER, ha='center', va='center',
            wrap=True)
    if subtitle:
        ax.text(0.5, 0.08, subtitle, transform=ax.transAxes,
                fontsize=7, color=SILVER, ha='center', va='center')

def variance_color(pct_change):
    """Color code variance."""
    a = abs(pct_change)
    if a <= 5:   return GREEN
    if a <= 15:  return YELLOW
    return RED

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 1 — COVER
# ─────────────────────────────────────────────────────────────────────────────
def make_page1(pdf):
    fig = plt.figure(figsize=(W, H), facecolor=WHITE)
    add_header(fig)

    # Main title block
    ax_title = fig.add_axes([0.06, 0.60, 0.88, 0.30])
    ax_title.set_facecolor(WHITE)
    ax_title.axis('off')

    ax_title.text(0.0, 0.95, 'Dealer Lead Distribution',
                  transform=ax_title.transAxes,
                  fontsize=36, fontweight='bold', color=NAVY,
                  va='top', ha='left')
    ax_title.text(0.0, 0.62, 'Pre-Launch Forecast vs. Post-Launch Actuals',
                  transform=ax_title.transAxes,
                  fontsize=16, color=DGRAY, va='top', ha='left')
    # Thin divider
    ax_title.axhline(y=0.48, xmin=0, xmax=0.55, color=SILVER, linewidth=1)

    # Three info cards
    card_data = [
        ('Baseline Window',     'April 2025 – April 2026'),
        ('Website Launch',      'May 27, 2026'),
        ('Post-Launch Window',  'May 28 – June 18, 2026\n(22 days)'),
    ]
    card_width  = 0.26
    card_gap    = 0.04
    card_left   = 0.06
    card_top    = 0.55
    card_height = 0.10

    for i, (lbl, val) in enumerate(card_data):
        x = card_left + i * (card_width + card_gap)
        ax_c = fig.add_axes([x, card_top, card_width, card_height])
        ax_c.set_facecolor(LGRAY)
        ax_c.axis('off')
        ax_c.add_patch(FancyBboxPatch((0.03, 0.05), 0.94, 0.90,
                                      boxstyle="round,pad=0.02",
                                      facecolor=LGRAY, edgecolor=SILVER,
                                      linewidth=0.8,
                                      transform=ax_c.transAxes, clip_on=False))
        ax_c.text(0.50, 0.78, lbl, transform=ax_c.transAxes,
                  fontsize=8, color=DGRAY, ha='center', va='top', fontweight='bold')
        ax_c.text(0.50, 0.38, val, transform=ax_c.transAxes,
                  fontsize=10, color=NAVY, ha='center', va='center',
                  fontweight='bold', multialignment='center')

    # Framing paragraph
    ax_para = fig.add_axes([0.06, 0.37, 0.88, 0.15])
    ax_para.axis('off')
    para = (
        "In advance of the May 2026 website relaunch, Element Three modeled how three changes — "
        "form elimination, dealer contact routing, and score-based notifications — would affect "
        "the volume and quality of leads delivered to Airstream's 83-dealer network. This report "
        "compares those forecasts against the first 22 days of actual post-launch data."
    )
    ax_para.text(0.0, 0.95, para, transform=ax_para.transAxes,
                 fontsize=10.5, color=DGRAY, va='top', ha='left',
                 wrap=True, multialignment='left',
                 bbox=dict(facecolor='none', edgecolor='none'),
                 linespacing=1.6,
                 # Use a fixed width approach:
                 )

    # Footer
    ax_foot = fig.add_axes([0.06, 0.04, 0.88, 0.03])
    ax_foot.axis('off')
    ax_foot.text(0.0, 0.5, 'Element Three  |  June 2026',
                 transform=ax_foot.transAxes,
                 fontsize=8, color=SILVER, va='center', ha='left')
    ax_foot.axhline(y=0.9, xmin=0, xmax=1, color=SILVER, linewidth=0.6)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 2 — BASELINE
# ─────────────────────────────────────────────────────────────────────────────
def make_page2(pdf):
    fig = plt.figure(figsize=(W, H), facecolor=WHITE)
    add_header(fig)

    # Section title
    ax_t = fig.add_axes([0.06, 0.885, 0.88, 0.045])
    ax_t.axis('off')
    ax_t.text(0.0, 0.95, 'The Pipeline Before Launch',
              transform=ax_t.transAxes, fontsize=18, fontweight='bold',
              color=NAVY, va='top')
    ax_t.text(0.0, 0.30, 'April 2025 – April 2026',
              transform=ax_t.transAxes, fontsize=10, color=DGRAY, va='top')

    # ── KPI BOXES ──
    kpi_data = [
        (f"{total_bl:,}",        "Total Leads",            "baseline period"),
        (f"{monthly_avg_bl:,.0f}", "Monthly Average",       "leads per month"),
        (f"{per_dealer_bl:.1f}", "Avg Per Dealer / Month", f"across {DEALERS} dealers"),
        (f"{score3_rate_bl:.0%}", "Score 3+ Rate",         "lead quality"),
    ]
    kpi_w = 0.19; kpi_gap = 0.025; kpi_left = 0.06
    for i, (val, lbl, sub) in enumerate(kpi_data):
        x = kpi_left + i * (kpi_w + kpi_gap)
        ax_k = fig.add_axes([x, 0.800, kpi_w, 0.075])
        kpi_box(ax_k, lbl, val, sub)

    # ── CHART A — Monthly Volume ──
    ax_a = fig.add_axes([0.06, 0.545, 0.54, 0.235])
    months_labels = [str(m) for m in monthly['YearMonth']]
    x_pos = np.arange(len(months_labels))
    bars = ax_a.bar(x_pos, monthly['leads'], color=NAVY, width=0.7, zorder=2)
    ax_a.axhline(monthly_avg_bl, color=SILVER, linestyle='--', linewidth=1.2,
                 label=f'Avg: {monthly_avg_bl:,.0f}', zorder=3)
    ax_a.set_xticks(x_pos)
    ax_a.set_xticklabels([m[-5:] for m in months_labels], rotation=45, ha='right', fontsize=7)
    ax_a.set_ylabel('Leads', fontsize=8, color=DGRAY)
    ax_a.set_title('A  —  Monthly Lead Volume', fontsize=9, fontweight='bold',
                   color=NAVY, loc='left', pad=4)
    ax_a.legend(fontsize=7, frameon=False)
    ax_a.spines['top'].set_visible(False)
    ax_a.spines['right'].set_visible(False)
    ax_a.tick_params(colors=DGRAY, labelsize=7)
    ax_a.yaxis.set_tick_params(labelsize=7)
    ax_a.set_facecolor(WHITE)
    for spine in ['left', 'bottom']:
        ax_a.spines[spine].set_color(SILVER)

    # ── CHART B — Source Donut ──
    ax_b = fig.add_axes([0.62, 0.560, 0.32, 0.215])
    donut_vals = [paid_count, web_count]
    donut_labels = ['External\nPaid (Meta)', 'Website-\nCaptured']
    donut_colors = [NAVY, SILVER]
    wedges, texts, autotexts = ax_b.pie(
        donut_vals, labels=donut_labels, colors=donut_colors,
        autopct='%1.0f%%', startangle=90,
        pctdistance=0.72, labeldistance=1.12,
        wedgeprops=dict(width=0.45, edgecolor=WHITE, linewidth=1.5)
    )
    for t in texts:   t.set_fontsize(7); t.set_color(DGRAY)
    for t in autotexts: t.set_fontsize(7.5); t.set_color(WHITE); t.set_fontweight('bold')
    ax_b.set_title('B  —  Source Split', fontsize=9, fontweight='bold',
                   color=NAVY, loc='left', pad=4)

    # ── CHART C — Score Distribution ──
    ax_c = fig.add_axes([0.06, 0.285, 0.88, 0.230])
    score_labels = [f'Score {s}' for s in all_scores]
    colors_score = [NAVY if s >= 3 else SILVER for s in all_scores]
    y_pos = np.arange(len(score_labels))
    hbars = ax_c.barh(y_pos, score_counts, color=colors_score, height=0.55, zorder=2)
    for i, (cnt, pct) in enumerate(zip(score_counts, score_pcts)):
        ax_c.text(cnt + max(score_counts) * 0.01, i,
                  f'{cnt:,}  ({pct:.0f}%)', va='center', fontsize=8, color=DGRAY)
    ax_c.set_yticks(y_pos)
    ax_c.set_yticklabels(score_labels, fontsize=8, color=DGRAY)
    ax_c.set_xlabel('Leads', fontsize=8, color=DGRAY)
    ax_c.set_title('C  —  Score Distribution  (navy = Score 3+)', fontsize=9,
                   fontweight='bold', color=NAVY, loc='left', pad=4)
    ax_c.spines['top'].set_visible(False)
    ax_c.spines['right'].set_visible(False)
    for spine in ['left', 'bottom']:
        ax_c.spines[spine].set_color(SILVER)
    ax_c.tick_params(colors=DGRAY, labelsize=7)
    ax_c.set_facecolor(WHITE)
    ax_c.set_xlim(0, max(score_counts) * 1.25)

    # Brand split note
    ax_brand = fig.add_axes([0.06, 0.215, 0.88, 0.060])
    ax_brand.set_facecolor(LGRAY)
    ax_brand.axis('off')
    brand_text = (
        f"Brand split:  TT {tt_pct:.0f}% ({tt_count:,} leads)  |  TC {tc_pct:.0f}% ({tc_count:,} leads)\n"
        f"Score 3+ rate:  TT {tt_rate:.0%}  vs.  TC {tc_rate:.0%}"
    )
    ax_brand.text(0.5, 0.5, brand_text, transform=ax_brand.transAxes,
                  fontsize=8.5, color=NAVY, ha='center', va='center',
                  multialignment='center')

    # Footer
    ax_foot = fig.add_axes([0.06, 0.04, 0.88, 0.02])
    ax_foot.axis('off')
    ax_foot.axhline(y=0.9, color=SILVER, linewidth=0.6)
    ax_foot.text(0.0, 0.2, 'Element Three  |  June 2026', fontsize=7, color=SILVER)
    ax_foot.text(1.0, 0.2, 'Page 2', fontsize=7, color=SILVER, ha='right')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 3 — FORECAST
# ─────────────────────────────────────────────────────────────────────────────
SCENARIOS = [
    ('Current (form submit)',                   19200),
    ('Form elimination only',                   17891),
    ('Score change (all, incl 0→1)',            17254),
    ('Elim + 50% checkbox',                     14715),
    ('Elim + 30% checkbox',                     13444),
    ('Elim + 2.4% checkbox (config proxy)',     11691),
    ('Elim + 30% CB + Score 2+ threshold',       5208),
    ('Elim + 30% CB + Score 3+ threshold',       1263),
]

DECISIONS = [
    ('D1 — Distribution trigger',
     'Hybrid: hand-raisers immediate,\nothers score-gated',
     'Brochure + quote requests\nroute on submission'),
    ('D2 — Brochure checkbox',
     'No checkbox\n(implicit consent)',
     'Brochure request = hand-raiser,\nroutes immediately'),
    ('D3A — New lead 0→1 counts?',
     'Yes',
     'Score-change model is\nvolume-neutral (~-10%)'),
    ('D3B — Score threshold',
     'All changes at launch',
     'Start near current volume;\ntune later'),
]

def make_page3(pdf):
    fig = plt.figure(figsize=(W, H), facecolor=WHITE)
    add_header(fig)

    # Section title
    ax_t = fig.add_axes([0.06, 0.885, 0.88, 0.045])
    ax_t.axis('off')
    ax_t.text(0.0, 0.95, 'Pre-Launch Forecast',
              transform=ax_t.transAxes, fontsize=18, fontweight='bold',
              color=NAVY, va='top')
    ax_t.text(0.0, 0.30, 'What the Model Predicted',
              transform=ax_t.transAxes, fontsize=10, color=DGRAY, va='top')

    # ── Scenario bar chart ──
    ax_s = fig.add_axes([0.06, 0.545, 0.88, 0.320])
    labels  = [s[0] for s in SCENARIOS]
    values  = [s[1] for s in SCENARIOS]
    y_pos   = np.arange(len(labels))
    colors  = [NAVY if lbl == 'Form elimination only' else SILVER for lbl in labels]
    hbars   = ax_s.barh(y_pos, values, color=colors, height=0.55, zorder=2)
    ax_s.axvline(x=19200, color=NAVY, linestyle='--', linewidth=1.0,
                 label='Baseline: 19,200/mo', zorder=3)
    for i, v in enumerate(values):
        ax_s.text(v + 150, i, f'{v:,}', va='center', fontsize=8, color=DGRAY)
    ax_s.set_yticks(y_pos)
    ax_s.set_yticklabels(labels, fontsize=8, color=DGRAY)
    ax_s.set_xlabel('Leads per Month', fontsize=8, color=DGRAY)
    ax_s.set_title('Scenario Modeling — Monthly Lead Volume', fontsize=10,
                   fontweight='bold', color=NAVY, loc='left', pad=6)
    ax_s.legend(fontsize=8, frameon=False)
    ax_s.spines['top'].set_visible(False)
    ax_s.spines['right'].set_visible(False)
    for sp in ['left', 'bottom']:
        ax_s.spines[sp].set_color(SILVER)
    ax_s.tick_params(colors=DGRAY)
    ax_s.set_facecolor(WHITE)
    ax_s.set_xlim(0, 23000)
    # Highlight note
    ax_s.text(17891 + 150, 6.38, '← Selected approach', fontsize=7.5,
              color=NAVY, va='bottom', style='italic')

    # ── Decision table ──
    ax_dt = fig.add_axes([0.06, 0.130, 0.88, 0.390])
    ax_dt.axis('off')
    ax_dt.set_title('Decision Summary', fontsize=10, fontweight='bold',
                    color=NAVY, loc='left', pad=6)

    # Table header
    col_headers = ['Decision', 'Choice Made', 'Effect']
    col_x = [0.0, 0.35, 0.68]
    col_widths = [0.34, 0.32, 0.32]

    # Header row bg
    ax_dt.add_patch(FancyBboxPatch((0, 0.84), 1.0, 0.14,
                                   boxstyle="square,pad=0",
                                   facecolor=NAVY, edgecolor='none',
                                   transform=ax_dt.transAxes, clip_on=False))
    for j, (hdr, x) in enumerate(zip(col_headers, col_x)):
        ax_dt.text(x + 0.01, 0.905, hdr, transform=ax_dt.transAxes,
                   fontsize=8.5, color=WHITE, fontweight='bold', va='center')

    row_colors = [WHITE, LGRAY, WHITE, LGRAY]
    row_height = 0.18
    for i, (dec, choice, effect) in enumerate(DECISIONS):
        y_top = 0.84 - (i + 1) * row_height
        ax_dt.add_patch(FancyBboxPatch((0, y_top), 1.0, row_height,
                                       boxstyle="square,pad=0",
                                       facecolor=row_colors[i], edgecolor=SILVER,
                                       linewidth=0.4,
                                       transform=ax_dt.transAxes, clip_on=False))
        y_center = y_top + row_height / 2
        row_data = [dec, choice, effect]
        for j, (txt, x) in enumerate(zip(row_data, col_x)):
            ax_dt.text(x + 0.01, y_center, txt, transform=ax_dt.transAxes,
                       fontsize=7.5, color=NAVY if j == 0 else DGRAY,
                       fontweight='bold' if j == 0 else 'normal',
                       va='center', multialignment='left')

    # Footer
    ax_foot = fig.add_axes([0.06, 0.04, 0.88, 0.02])
    ax_foot.axis('off')
    ax_foot.axhline(y=0.9, color=SILVER, linewidth=0.6)
    ax_foot.text(0.0, 0.2, 'Element Three  |  June 2026', fontsize=7, color=SILVER)
    ax_foot.text(1.0, 0.2, 'Page 3', fontsize=7, color=SILVER, ha='right')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE 4 — ACTUALS vs. FORECAST
# ─────────────────────────────────────────────────────────────────────────────
FORECAST_MONTHLY = 17891  # form elimination only scenario

def make_page4(pdf):
    fig = plt.figure(figsize=(W, H), facecolor=WHITE)
    add_header(fig)

    # Section title
    ax_t = fig.add_axes([0.06, 0.885, 0.88, 0.045])
    ax_t.axis('off')
    ax_t.text(0.0, 0.95, 'Post-Launch Actuals',
              transform=ax_t.transAxes, fontsize=18, fontweight='bold',
              color=NAVY, va='top')
    ax_t.text(0.0, 0.30, 'May 28 – June 18, 2026  (22 days)',
              transform=ax_t.transAxes, fontsize=10, color=DGRAY, va='top')

    # ── Comparison scorecard ──
    scorecard_rows = [
        ('Avg Leads / Month',    f'{monthly_avg_bl:,.0f}',   f'{FORECAST_MONTHLY:,}',   f'{monthly_rate_pl:,.0f}'),
        ('Per Dealer / Month',   f'{per_dealer_bl:.1f}',     f'{FORECAST_MONTHLY/DEALERS:.1f}', f'{per_dealer_pl:.1f}'),
        ('Score 3+ Rate',        f'{score3_rate_bl:.0%}',    '—',                        f'{score3_rate_pl:.0%}'),
        ('External Paid %',      f'{paid_pct_bl:.0f}%',      '—',                        f'{paid_pct_pl:.0f}%'),
    ]

    # Header
    ax_sc = fig.add_axes([0.06, 0.720, 0.88, 0.155])
    ax_sc.axis('off')
    ax_sc.set_title('Comparison Scorecard', fontsize=10, fontweight='bold',
                    color=NAVY, loc='left', pad=6)

    col_labels = ['Metric', 'Baseline\n(Apr 25–Apr 26)', 'Forecast\n(Form Elim.)', 'Actual\n(22-day)']
    col_x = [0.0, 0.30, 0.55, 0.75]

    # Header bg
    ax_sc.add_patch(FancyBboxPatch((0, 0.74), 1.0, 0.24,
                                   boxstyle="square,pad=0",
                                   facecolor=NAVY, edgecolor='none',
                                   transform=ax_sc.transAxes))
    for j, (hdr, x) in enumerate(zip(col_labels, col_x)):
        ax_sc.text(x + 0.005, 0.85, hdr, transform=ax_sc.transAxes,
                   fontsize=7.5, color=WHITE, fontweight='bold', va='center',
                   multialignment='center')

    row_h = 0.185
    for i, (metric, bl_v, fc_v, act_v) in enumerate(scorecard_rows):
        y = 0.74 - (i + 1) * row_h
        bg = LGRAY if i % 2 == 0 else WHITE
        ax_sc.add_patch(FancyBboxPatch((0, y), 1.0, row_h,
                                       boxstyle="square,pad=0",
                                       facecolor=bg, edgecolor=SILVER, linewidth=0.3,
                                       transform=ax_sc.transAxes))
        yc = y + row_h / 2
        ax_sc.text(col_x[0] + 0.005, yc, metric, transform=ax_sc.transAxes,
                   fontsize=8, color=NAVY, fontweight='bold', va='center')
        ax_sc.text(col_x[1] + 0.005, yc, bl_v, transform=ax_sc.transAxes,
                   fontsize=8, color=DGRAY, va='center')
        ax_sc.text(col_x[2] + 0.005, yc, fc_v, transform=ax_sc.transAxes,
                   fontsize=8, color=DGRAY, va='center')

        # Color-code actual vs forecast
        try:
            act_num = float(act_v.replace('%','').replace(',',''))
            fc_num  = float(fc_v.replace('%','').replace(',','').replace('—','0'))
            if fc_v != '—' and fc_num != 0:
                pct_diff = (act_num - fc_num) / fc_num * 100
                vc = variance_color(pct_diff)
                sign = '+' if pct_diff > 0 else ''
                var_str = f'{sign}{pct_diff:.0f}%'
            else:
                vc = DGRAY; var_str = ''
        except:
            vc = DGRAY; var_str = ''

        ax_sc.text(col_x[3] + 0.005, yc, act_v, transform=ax_sc.transAxes,
                   fontsize=8.5, color=vc, fontweight='bold', va='center')
        if var_str:
            ax_sc.text(col_x[3] + 0.085, yc, var_str, transform=ax_sc.transAxes,
                       fontsize=7, color=vc, va='center', style='italic')

    # ── Side-by-side bar chart ──
    ax_bar = fig.add_axes([0.06, 0.440, 0.60, 0.255])
    categories = ['Baseline\nAvg/Month', 'Forecast\n(Form Elim.)', 'Actual\n(Implied/Mo)']
    values_bar = [monthly_avg_bl, FORECAST_MONTHLY, monthly_rate_pl]
    colors_bar = [SILVER, NAVY, '#4A7DBF']
    x_pos = np.arange(len(categories))
    b = ax_bar.bar(x_pos, values_bar, color=colors_bar, width=0.55, zorder=2)
    for xi, v in zip(x_pos, values_bar):
        ax_bar.text(xi, v + 100, f'{v:,.0f}', ha='center', fontsize=8.5,
                    fontweight='bold', color=NAVY)
    ax_bar.set_xticks(x_pos)
    ax_bar.set_xticklabels(categories, fontsize=8, color=DGRAY)
    ax_bar.set_ylabel('Leads per Month', fontsize=8, color=DGRAY)
    ax_bar.set_title('Volume Comparison', fontsize=10, fontweight='bold',
                     color=NAVY, loc='left', pad=4)
    ax_bar.spines['top'].set_visible(False)
    ax_bar.spines['right'].set_visible(False)
    for sp in ['left', 'bottom']:
        ax_bar.spines[sp].set_color(SILVER)
    ax_bar.tick_params(colors=DGRAY)
    ax_bar.set_facecolor(WHITE)
    ax_bar.set_ylim(0, max(values_bar) * 1.20)

    # ── Callout box ──
    ax_call = fig.add_axes([0.68, 0.440, 0.26, 0.255])
    ax_call.set_facecolor(LGRAY)
    ax_call.axis('off')

    pct_vs_fc = (monthly_rate_pl - FORECAST_MONTHLY) / FORECAST_MONTHLY * 100
    pct_vs_bl = (monthly_rate_pl - monthly_avg_bl) / monthly_avg_bl * 100
    sign_fc = '+' if pct_vs_fc > 0 else ''
    sign_bl = '+' if pct_vs_bl > 0 else ''

    callout_txt = (
        f"Actuals vs. Forecast\n\n"
        f"{sign_fc}{pct_vs_fc:.0f}% vs. form-elim\nforecast\n\n"
        f"{sign_bl}{pct_vs_bl:.0f}% vs. baseline\naverage\n\n"
        f"Post-launch daily rate:\n{daily_rate_pl:.0f} leads/day\n\n"
        f"Implied monthly:\n{monthly_rate_pl:,.0f}"
    )
    ax_call.text(0.5, 0.95, callout_txt, transform=ax_call.transAxes,
                 fontsize=8, color=NAVY, ha='center', va='top',
                 multialignment='center', linespacing=1.5)

    # ── Caveat box ──
    ax_cav = fig.add_axes([0.06, 0.195, 0.88, 0.220])
    ax_cav.set_facecolor(LGRAY)
    ax_cav.axis('off')
    ax_cav.add_patch(FancyBboxPatch((0.005, 0.04), 0.990, 0.92,
                                    boxstyle="round,pad=0.02",
                                    facecolor=LGRAY, edgecolor=SILVER,
                                    linewidth=0.8,
                                    transform=ax_cav.transAxes))
    ax_cav.text(0.5, 0.92, 'Early Read — Statistical Caveats',
                transform=ax_cav.transAxes, fontsize=9, fontweight='bold',
                color=NAVY, ha='center', va='top')
    caveat = (
        "The 22-day post-launch window is shorter than the 60–90 day minimum recommended for "
        "statistical confidence. Seasonality (June is historically a peak month, ~28K leads), "
        "pipeline lag, and score-progression timing all affect this comparison. "
        "Treat actuals as directional only until a full quarter of post-launch data is available."
    )
    ax_cav.text(0.5, 0.65, caveat, transform=ax_cav.transAxes,
                fontsize=8.5, color=DGRAY, ha='center', va='top',
                multialignment='center', linespacing=1.6)

    # Footer
    ax_foot = fig.add_axes([0.06, 0.04, 0.88, 0.02])
    ax_foot.axis('off')
    ax_foot.axhline(y=0.9, color=SILVER, linewidth=0.6)
    ax_foot.text(0.0, 0.2, 'Element Three  |  June 2026', fontsize=7, color=SILVER)
    ax_foot.text(1.0, 0.2, 'Page 4', fontsize=7, color=SILVER, ha='right')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

# ─────────────────────────────────────────────────────────────────────────────
# GENERATE PDF
# ─────────────────────────────────────────────────────────────────────────────
print(f"\nGenerating PDF → {OUT}")
with PdfPages(OUT) as pdf:
    make_page1(pdf)
    print("  Page 1 done (Cover)")
    make_page2(pdf)
    print("  Page 2 done (Baseline)")
    make_page3(pdf)
    print("  Page 3 done (Forecast)")
    make_page4(pdf)
    print("  Page 4 done (Actuals)")

print(f"\nDONE: saved to {OUT}")
print(f"\n── SUMMARY ──────────────────────────────────────")
print(f"  Total baseline leads:       {total_bl:,}")
print(f"  Monthly avg (baseline):     {monthly_avg_bl:,.0f}")
print(f"  Score 3+ rate (baseline):   {score3_rate_bl:.1%}")
print(f"  Post-launch daily rate:     {daily_rate_pl:.1f} leads/day")
print(f"  Implied monthly rate:       {monthly_rate_pl:,.0f}")
print(f"  Score 3+ rate (post-launch):{score3_rate_pl:.1%}")
