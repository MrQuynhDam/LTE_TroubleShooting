"""
Ericsson LTE RAN Performance & Automated Root-Cause Analysis (RCA) Cockpit.

Hybrid Telemetry Cockpit (12-Column Grid) with High-Density Dark OLED (#0B0F19)
aesthetic, Multi-Tier Collapsible Sidebar Filters, Cross-Filtering Pareto Charts,
Slide-Over Diagnostic Evidence Drawer, One-Click AMOS/Moshell Terminal Copy,
and Executive HTML-to-PDF Report Exporter.
"""

from __future__ import annotations

import base64
import html
import json
from pathlib import Path
from typing import List, Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import streamlit.components.v1 as components

from pdf_generator import export_report_to_pdf
from rca_engine import AnalysisResult, analyze_counter_dump, compute_rop_series

# ============================================================================
# 1. PAGE CONFIG & DESIGN SYSTEM CSS (HIGH-DENSITY DARK OLED #0B0F19)
# ============================================================================

st.set_page_config(
    page_title="Ericsson LTE RAN Telemetry & RCA Cockpit",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_DATASET_PATH = Path(__file__).resolve().parent / "Data" / "Counters.xlsx"

TRIAGE_PRESETS = [
    "All Degraded",
    "Drop Call Spike (DCR > 1%)",
    "Handover / RACH Issue (HOSR < 95%)",
    "Accessibility Failure (CSSR < 97%)",
    "IRAT Ping-Pong Loop (> 5%)",
    "Radio Quality / CQI Degraded",
    "All Evaluated Cells",
]


def inject_cockpit_css() -> None:
    """Apply High-Density Dark OLED (#0B0F19) NOC Design System & Tabular Monospace."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');

        :root {
            --bg-oled: #0B0F19;
            --bg-card: #151D2E;
            --bg-card-hover: #1B253B;
            --border-hairline: #1E293B;
            --border-accent: rgba(56, 189, 248, 0.38);
            --text-primary: #F1F5F9;
            --text-secondary: #94A3B8;
            --text-muted: #64748B;
            --crit-hex: #FF4D4F;
            --crit-bg: rgba(255, 77, 79, 0.10);
            --crit-bdr: rgba(255, 77, 79, 0.40);
            --deg-hex: #FAAD14;
            --deg-bg: rgba(250, 173, 20, 0.10);
            --deg-bdr: rgba(250, 173, 20, 0.40);
            --norm-hex: #52C41A;
            --norm-bg: rgba(82, 196, 26, 0.10);
            --norm-bdr: rgba(82, 196, 26, 0.40);
            --cyan-hex: #38BDF8;
        }

        /* Base App Surface */
        .stApp, [data-testid="stAppViewContainer"] {
            background-color: var(--bg-oled) !important;
            color: var(--text-primary) !important;
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
        }
        [data-testid="stHeader"] {
            background: rgba(11, 15, 25, 0.90) !important;
            backdrop-filter: blur(8px);
            border-bottom: 1px solid rgba(30, 41, 59, 0.65);
        }
        [data-testid="stMainBlockContainer"], .block-container {
            padding-top: 3.95rem !important;
            padding-bottom: 2.5rem !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
            max-width: 100% !important;
        }

        /* Sidebar Styling */
        [data-testid="stSidebar"] {
            background-color: #0E1422 !important;
            border-right: 1px solid var(--border-hairline) !important;
        }
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            font-family: 'Space Grotesk', sans-serif !important;
            letter-spacing: -0.02em;
        }

        /* Typography Classes */
        h1, h2, h3, .noc-heading {
            font-family: 'Space Grotesk', sans-serif !important;
            letter-spacing: -0.025em;
            color: var(--text-primary);
        }
        .mono-num {
            font-family: 'JetBrains Mono', 'Roboto Mono', monospace !important;
            font-variant-numeric: tabular-nums;
        }

        /* Top Telemetry Header */
        .telemetry-bar {
            background: linear-gradient(135deg, #111827 0%, #151D2E 65%, #0F172A 100%);
            border: 1px solid var(--border-hairline);
            border-left: 3px solid var(--cyan-hex);
            border-radius: 10px;
            padding: 12px 18px;
            margin-bottom: 12px;
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
        }
        .telemetry-brand {
            display: flex;
            align-items: center;
            gap: 14px;
        }
        .telemetry-title {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 1.18rem;
            font-weight: 700;
            color: #F8FAFC;
            margin: 0;
            line-height: 1.2;
        }
        .telemetry-sub {
            font-size: 0.78rem;
            color: var(--text-secondary);
            margin-top: 3px;
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: center;
        }
        .fresh-pill {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 3px 10px;
            border-radius: 5px;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.72rem;
            font-weight: 600;
            background: rgba(82, 196, 26, 0.12);
            color: #52C41A;
            border: 1px solid rgba(82, 196, 26, 0.38);
        }
        .pulse-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: #52C41A;
            box-shadow: 0 0 8px #52C41A;
        }
        .meta-chip {
            background: rgba(15, 23, 42, 0.85);
            border: 1px solid var(--border-hairline);
            border-radius: 5px;
            padding: 2px 8px;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.73rem;
            color: #CBD5E1;
        }

        /* Tier 1 Executive KPI Cards */
        .kpi-card {
            background: var(--bg-card);
            border: 1px solid var(--border-hairline);
            border-radius: 10px;
            padding: 14px 16px;
            position: relative;
            overflow: hidden;
            transition: border-color 0.18s ease, transform 0.18s ease, background 0.18s ease;
            min-height: 112px;
        }
        .kpi-card:hover {
            background: var(--bg-card-hover);
            border-color: var(--border-accent);
            transform: translateY(-1px);
        }
        .kpi-card::before {
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 2px;
            background: var(--cyan-hex);
        }
        .kpi-card.card-crit::before { background: var(--crit-hex); }
        .kpi-card.card-deg::before { background: var(--deg-hex); }
        .kpi-card.card-norm::before { background: var(--norm-hex); }

        .kpi-label {
            font-size: 0.73rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            color: var(--text-secondary);
            margin-bottom: 6px;
        }
        .kpi-value-row {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 8px;
        }
        .kpi-value {
            font-family: 'JetBrains Mono', monospace;
            font-variant-numeric: tabular-nums;
            font-size: 1.68rem;
            font-weight: 700;
            color: #F8FAFC;
            line-height: 1.15;
        }
        .kpi-foot {
            margin-top: 8px;
            font-size: 0.75rem;
            color: var(--text-secondary);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 6px;
        }

        /* Semantic Status Badges */
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 2px 8px;
            border-radius: 5px;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.71rem;
            font-weight: 700;
            letter-spacing: 0.04em;
        }
        .badge-critical {
            color: var(--crit-hex);
            background: var(--crit-bg);
            border: 1px solid var(--crit-bdr);
        }
        .badge-degraded {
            color: var(--deg-hex);
            background: var(--deg-bg);
            border: 1px solid var(--deg-bdr);
        }
        .badge-normal {
            color: var(--norm-hex);
            background: var(--norm-bg);
            border: 1px solid var(--norm-bdr);
        }
        .badge-low {
            color: #94A3B8;
            background: rgba(148, 163, 184, 0.10);
            border: 1px solid rgba(148, 163, 184, 0.30);
        }

        /* Panel Containers */
        .panel-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 8px 12px;
            background: var(--bg-card);
            border: 1px solid var(--border-hairline);
            border-radius: 8px;
            margin-bottom: 8px;
        }
        .panel-title {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 0.92rem;
            font-weight: 600;
            color: #F8FAFC;
            letter-spacing: -0.01em;
        }
        .panel-hint {
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.71rem;
            color: var(--cyan-hex);
        }

        /* True Off-Canvas Fixed Slide-Over Inspector Drawer */
        @keyframes slideOverFromRight {
            from { opacity: 0; transform: translateX(100%); }
            to   { opacity: 1; transform: translateX(0); }
        }
        .st-key-slide_over_drawer {
            position: fixed !important;
            top: 3.75rem !important;
            right: 0 !important;
            bottom: 0 !important;
            width: min(540px, 94vw) !important;
            z-index: 999990 !important;
            background: linear-gradient(180deg, #111827 0%, #0B0F19 100%) !important;
            border-left: 2px solid var(--cyan-hex) !important;
            box-shadow: -24px 0 64px rgba(0, 0, 0, 0.85) !important;
            padding: 14px 18px 28px 18px !important;
            overflow-y: auto !important;
            animation: slideOverFromRight 0.26s cubic-bezier(0.16, 1, 0.3, 1) !important;
        }
        .inspector-drawer {
            background: rgba(21, 29, 46, 0.72);
            border: 1px solid var(--border-hairline);
            border-left: 3px solid var(--cyan-hex);
            border-radius: 10px;
            padding: 14px 15px;
        }
        .inspector-drawer.drawer-CRITICAL { border-left-color: var(--crit-hex); }
        .inspector-drawer.drawer-DEGRADED { border-left-color: var(--deg-hex); }
        .inspector-drawer.drawer-NORMAL   { border-left-color: var(--norm-hex); }

        .drawer-top {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            border-bottom: 1px solid var(--border-hairline);
            padding-bottom: 10px;
            margin-bottom: 10px;
        }
        .drawer-cell-id {
            font-family: 'JetBrains Mono', monospace;
            font-size: 1.05rem;
            font-weight: 700;
            color: #F8FAFC;
        }
        .drawer-enb-id {
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.76rem;
            color: var(--text-secondary);
            margin-top: 2px;
        }
        .rca-callout {
            background: rgba(11, 15, 25, 0.75);
            border: 1px solid var(--border-hairline);
            border-radius: 7px;
            padding: 10px 12px;
            margin-bottom: 10px;
            font-size: 0.81rem;
            line-height: 1.45;
        }
        .mini-kpi-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 6px;
            margin-bottom: 10px;
        }
        .mini-kpi {
            background: #0B0F19;
            border: 1px solid var(--border-hairline);
            border-radius: 6px;
            padding: 6px 6px;
            text-align: center;
        }
        .mini-kpi .mk-lbl {
            font-size: 0.65rem;
            color: var(--text-secondary);
            text-transform: uppercase;
        }
        .mini-kpi .mk-val {
            font-family: 'JetBrains Mono', monospace;
            font-variant-numeric: tabular-nums;
            font-size: 0.84rem;
            font-weight: 700;
            color: #F8FAFC;
            margin-top: 2px;
        }

        /* Diagnostic Evidence Table */
        table.ev-grid {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.74rem;
            margin-bottom: 10px;
            background: #0B0F19;
            border-radius: 6px;
            overflow: hidden;
        }
        table.ev-grid th {
            background: #151D2E;
            color: #94A3B8;
            text-align: left;
            padding: 5px 8px;
            font-weight: 600;
            border-bottom: 1px solid var(--border-hairline);
        }
        table.ev-grid td {
            padding: 4.5px 8px;
            border-bottom: 1px solid rgba(30, 41, 59, 0.65);
            color: #E2E8F0;
        }
        table.ev-grid td.mono {
            font-family: 'JetBrains Mono', monospace;
            font-variant-numeric: tabular-nums;
        }

        /* Pure Black OLED AMOS Terminal */
        .amos-terminal-wrap {
            background: #000000;
            border: 1px solid #1E293B;
            border-radius: 8px;
            padding: 10px 12px;
            margin-top: 6px;
        }
        .amos-term-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #1E293B;
            padding-bottom: 6px;
            margin-bottom: 8px;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.71rem;
            color: #94A3B8;
        }
        .amos-dots span {
            display: inline-block;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            margin-right: 4px;
        }
        .amos-code-box {
            margin: 0;
            background: #000000;
            color: #38BDF8;
            font-family: 'JetBrains Mono', 'Roboto Mono', Consolas, monospace;
            font-size: 0.73rem;
            line-height: 1.45;
            max-height: 240px;
            overflow-y: auto;
            padding: 2px 0;
        }
        .amos-line {
            white-space: pre-wrap;
            word-break: break-word;
            font-family: 'JetBrains Mono', 'Roboto Mono', Consolas, monospace;
        }
        .amos-cmt { color: #6EE7B7; }
        .amos-set { color: #FBBF24; font-weight: 600; }
        .amos-hdr { color: #F8FAFC; font-weight: 700; }

        /* Crisp Action Buttons */
        .stButton > button, .stDownloadButton > button {
            border-radius: 7px !important;
            border: 1px solid #26354D !important;
            background-color: #151D2E !important;
            color: #F1F5F9 !important;
            font-family: 'Inter', sans-serif !important;
            font-weight: 600 !important;
            font-size: 0.80rem !important;
            transition: all 0.16s ease !important;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            border-color: #38BDF8 !important;
            background-color: #1B2740 !important;
            color: #FFFFFF !important;
        }
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
            background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
            border-color: #38BDF8 !important;
            color: #FFFFFF !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# 2. CACHED DATA INGESTION & CLIPBOARD / AUTO-DOWNLOAD WIDGETS
# ============================================================================

def render_html(markup: str) -> None:
    """
    Strip leading line indentation & empty lines before st.markdown so CommonMark
    never misinterprets indented HTML tags as raw <pre><code> blocks.
    """
    compact = "\n".join(line.strip() for line in markup.splitlines() if line.strip())
    st.markdown(compact, unsafe_allow_html=True)


def format_amos_terminal_html(amos_script: str) -> str:
    """Format AMOS/Moshell script lines into syntax-colored HTML without blank lines."""
    out_lines: List[str] = []
    for raw_line in amos_script.splitlines():
        s = raw_line.strip()
        if not s:
            out_lines.append('<div class="amos-line">&nbsp;</div>')
        elif s.startswith("#"):
            out_lines.append(f'<div class="amos-line amos-cmt">{html.escape(raw_line)}</div>')
        elif s.startswith("set "):
            out_lines.append(f'<div class="amos-line amos-set">{html.escape(raw_line)}</div>')
        elif s.startswith(("moshell ", "confb", "gs+", "gs-", "lt ", "cvms ")):
            out_lines.append(f'<div class="amos-line amos-hdr">{html.escape(raw_line)}</div>')
        else:
            out_lines.append(f'<div class="amos-line">{html.escape(raw_line)}</div>')
    return "".join(out_lines)


@st.cache_data(show_spinner=False)
def load_cached_dump(file_bytes: bytes, file_name: str) -> AnalysisResult:
    """Cache parsed KPIs, vector pivots, and RCA outputs across UI interactions."""
    return analyze_counter_dump(file_bytes, filename=file_name)


@st.cache_data(show_spinner=False)
def load_default_dump(path_str: str, mtime: float) -> AnalysisResult:
    """Load workspace Data/Counters.xlsx with mtime cache invalidation."""
    _ = mtime
    p = Path(path_str)
    return analyze_counter_dump(p, filename=p.name)


def _render_iframe(markup: str, height: int) -> None:
    """Embed HTML/JS widget using st.iframe (Streamlit >= 1.56) or components.html fallback."""
    if hasattr(st, "iframe"):
        st.iframe(markup, height=max(1, height))
    else:
        components.html(markup, height=height)


def render_copy_button(text_to_copy: str, label: str, key_id: str, full_width: bool = True) -> None:
    """Render a tactile dark-OLED One-Click Copy button backed by browser Clipboard API."""
    payload = json.dumps(text_to_copy)
    width_css = "100%" if full_width else "auto"
    _render_iframe(
        f"""
        <div style="margin:0;padding:0;">
          <button id="btn_{key_id}" onclick="copyCmd_{key_id}()" style="
            width:{width_css};
            background:linear-gradient(135deg,#0F172A 0%,#1E293B 100%);
            color:#38BDF8;
            border:1px solid rgba(56,189,248,0.45);
            border-radius:6px;
            padding:7px 12px;
            font-family:'JetBrains Mono',Consolas,monospace;
            font-size:12px;
            font-weight:600;
            cursor:pointer;
            display:flex;
            align-items:center;
            justify-content:center;
            gap:8px;
            transition:all 0.15s ease;">
            <span>⌘</span> <span id="lbl_{key_id}">{html.escape(label)}</span>
          </button>
        </div>
        <script>
        function copyCmd_{key_id}() {{
            const txt = {payload};
            const lbl = document.getElementById("lbl_{key_id}");
            const btn = document.getElementById("btn_{key_id}");
            const done = () => {{
                lbl.innerText = "✓ COPIED TO CLIPBOARD";
                btn.style.borderColor = "#52C41A";
                btn.style.color = "#52C41A";
                setTimeout(() => {{
                    lbl.innerText = {json.dumps(label)};
                    btn.style.borderColor = "rgba(56,189,248,0.45)";
                    btn.style.color = "#38BDF8";
                }}, 2000);
            }};
            if (navigator.clipboard && navigator.clipboard.writeText) {{
                navigator.clipboard.writeText(txt).then(done).catch(() => fallbackCopy(txt, done));
            }} else {{
                fallbackCopy(txt, done);
            }}
        }}
        function fallbackCopy(text, cb) {{
            const ta = document.createElement("textarea");
            ta.value = text;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
            cb();
        }}
        </script>
        """,
        height=40,
    )


def trigger_pdf_autodownload(pdf_bytes: bytes, download_name: str) -> None:
    """Trigger immediate browser download of generated PDF bytes."""
    b64 = base64.b64encode(pdf_bytes).decode("ascii")
    _render_iframe(
        f"""
        <a id="auto_pdf_dl" href="data:application/pdf;base64,{b64}" download="{html.escape(download_name)}"></a>
        <script>document.getElementById('auto_pdf_dl').click();</script>
        """,
        height=1,
    )


# ============================================================================
# 3. PLOTLY HIGH-DENSITY DARK OLED CHARTS
# ============================================================================

PLOTLY_FONT_MONO = "JetBrains Mono, Roboto Mono, Consolas, monospace"
PLOTLY_FONT_SANS = "Inter, Space Grotesk, sans-serif"


def build_rop_trend_chart(ts_df: pd.DataFrame, scope_label: str) -> go.Figure:
    """Dual-axis ROP Time-Series Trendline: CSSR (%) vs DCR (%)."""
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    if ts_df.empty:
        fig.update_layout(
            paper_bgcolor="#151D2E",
            plot_bgcolor="#0B0F19",
            height=295,
            annotations=[dict(text="No ROP time-series records in selection", showarrow=False, font=dict(color="#94A3B8"))],
        )
        return fig

    fig.add_trace(
        go.Scatter(
            x=ts_df["ROP_Time"],
            y=ts_df["CSSR"],
            name="CSSR (%)",
            mode="lines+markers",
            line=dict(color="#38BDF8", width=2.4),
            marker=dict(size=5, color="#38BDF8"),
            hovertemplate="<b>%{x|%Y-%m-%d %H:%M}</b><br>CSSR: <b>%{y:.2f}%</b><extra></extra>",
        ),
        secondary_y=False,
    )

    dcr_ovr = ts_df["DCR_Overall"] if "DCR_Overall" in ts_df.columns else ts_df["DCR"]
    fig.add_trace(
        go.Bar(
            x=ts_df["ROP_Time"],
            y=dcr_ovr,
            name="Overall DCR (%)",
            marker=dict(
                color=np.where(dcr_ovr > 2.5, "rgba(255,77,79,0.58)", np.where(dcr_ovr > 1.0, "rgba(250,173,20,0.52)", "rgba(82,196,26,0.40)")),
                line=dict(width=0),
            ),
            hovertemplate="<b>%{x|%Y-%m-%d %H:%M}</b><br>Overall DCR: <b>%{y:.2f}%</b><extra></extra>",
        ),
        secondary_y=True,
    )
    if "DCR_eNB" in ts_df.columns:
        fig.add_trace(
            go.Scatter(
                x=ts_df["ROP_Time"],
                y=ts_df["DCR_eNB"],
                name="eNB DCR (%)",
                mode="lines",
                line=dict(color="#FAAD14", width=1.6, dash="dot"),
                hovertemplate="<b>%{x|%Y-%m-%d %H:%M}</b><br>eNB DCR: <b>%{y:.2f}%</b><extra></extra>",
            ),
            secondary_y=True,
        )

    # Baseline Reference Lines
    fig.add_hline(
        y=98.5,
        line_dash="dot",
        line_color="#52C41A",
        line_width=1.2,
        annotation_text="CSSR Baseline 98.5%",
        annotation_font=dict(size=10, color="#52C41A", family=PLOTLY_FONT_MONO),
        annotation_position="bottom right",
        secondary_y=False,
    )
    fig.add_hline(
        y=1.0,
        line_dash="dash",
        line_color="#FAAD14",
        line_width=1.1,
        secondary_y=True,
    )

    c_min = max(0.0, float(ts_df["CSSR"].min()) - 0.6)
    fig.update_layout(
        title=dict(
            text=f"ROP Telemetry Trendline — {scope_label}",
            font=dict(family="Space Grotesk, sans-serif", size=13, color="#F8FAFC"),
            x=0.01,
            y=0.96,
        ),
        paper_bgcolor="#151D2E",
        plot_bgcolor="#0B0F19",
        font=dict(family=PLOTLY_FONT_MONO, size=11, color="#94A3B8"),
        height=300,
        margin=dict(l=48, r=48, t=38, b=34),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=0.99, font=dict(size=10.5)),
        hovermode="x unified",
        bargap=0.28,
    )
    fig.update_xaxes(showgrid=True, gridcolor="#1E293B", zeroline=False)
    fig.update_yaxes(
        title_text="CSSR (%)",
        range=[min(96.5, c_min), 100.15],
        showgrid=True,
        gridcolor="#1E293B",
        secondary_y=False,
    )
    fig.update_yaxes(
        title_text="DCR (%)",
        showgrid=False,
        rangemode="tozero",
        secondary_y=True,
    )
    return fig


def build_cqi_stacked_chart(df: pd.DataFrame) -> go.Figure:
    """Stacked Bar Chart of CQI Modulation Quality Bands across Worst/Filtered Cells."""
    fig = go.Figure()
    sub = df[df["CQI_Samples"] > 0].copy()
    if sub.empty:
        fig.update_layout(
            paper_bgcolor="#151D2E",
            plot_bgcolor="#0B0F19",
            height=300,
            annotations=[dict(text="No CQI vector samples available in current filter", showarrow=False, font=dict(color="#94A3B8"))],
        )
        return fig

    # Show up to 14 cells sorted by highest Low-CQI tail for maximum diagnostic signal
    sub = sub.sort_values("Low_CQI_Ratio", ascending=False).head(14)

    bands = [
        ("QPSK • Low CQI (0–6)", "CQI_QPSK_Pct", "#FF4D4F"),
        ("16-QAM • Fair (7–9)", "CQI_16QAM_Pct", "#FAAD14"),
        ("64-QAM • Good (10–12)", "CQI_64QAM_Pct", "#38BDF8"),
        ("256-QAM • Peak (13–15)", "CQI_256QAM_Pct", "#52C41A"),
    ]
    for label, col, color in bands:
        fig.add_trace(
            go.Bar(
                x=sub["EUtranCellFDD"],
                y=sub[col],
                name=label,
                marker_color=color,
                hovertemplate=f"<b>%{{x}}</b><br>{label}: <b>%{{y:.1f}}%</b><extra></extra>",
            )
        )

    fig.update_layout(
        barmode="stack",
        title=dict(
            text="Radio Quality — CQI Modulation Band Distribution (Worst Low-CQI Cells)",
            font=dict(family="Space Grotesk, sans-serif", size=13, color="#F8FAFC"),
            x=0.01,
            y=0.96,
        ),
        paper_bgcolor="#151D2E",
        plot_bgcolor="#0B0F19",
        font=dict(family=PLOTLY_FONT_MONO, size=10.5, color="#94A3B8"),
        height=300,
        margin=dict(l=46, r=18, t=38, b=68),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=0.99, font=dict(size=10)),
        yaxis=dict(title="Share (%)", range=[0, 100], gridcolor="#1E293B"),
        xaxis=dict(tickangle=-28, showgrid=False),
    )
    return fig


def build_prb_spectrum_chart(cell_name: str, prb_values: List[float]) -> go.Figure:
    """PRB 1..100 Uplink Interference Spectrum Profile (PIM / Narrowband Detector)."""
    prb_idx = list(range(1, len(prb_values) + 1))
    arr = np.array(prb_values, dtype=float)
    bar_colors = np.where(arr > -108.0, "#FF4D4F", np.where(arr > -112.0, "#FAAD14", "#38BDF8"))

    fig = go.Figure(
        go.Bar(
            x=prb_idx,
            y=arr + 125.0,  # Offset from -125 dBm floor so bars rise upward cleanly
            base=-125.0,
            marker_color=bar_colors,
            customdata=arr,
            hovertemplate="<b>PRB %{x}</b><br>UL Noise: <b>%{customdata:.2f} dBm</b><extra></extra>",
        )
    )
    fig.add_hline(
        y=-110.0,
        line_dash="dash",
        line_color="#FF4D4F",
        line_width=1.2,
        annotation_text="Interference Limit (-110 dBm)",
        annotation_font=dict(size=10, color="#FF4D4F", family=PLOTLY_FONT_MONO),
    )
    fig.update_layout(
        title=dict(
            text=f"UL PRB Interference Spectrum (PRB 1–{len(prb_values)}) — {cell_name}",
            font=dict(family="Space Grotesk, sans-serif", size=13, color="#F8FAFC"),
            x=0.01,
            y=0.96,
        ),
        paper_bgcolor="#151D2E",
        plot_bgcolor="#0B0F19",
        font=dict(family=PLOTLY_FONT_MONO, size=10.5, color="#94A3B8"),
        height=300,
        margin=dict(l=52, r=18, t=38, b=36),
        xaxis=dict(title="Physical Resource Block (PRB Index)", gridcolor="#1E293B"),
        yaxis=dict(title="dBm / PRB", range=[-123, max(-102.0, float(arr.max()) + 2.0)], gridcolor="#1E293B"),
    )
    return fig


def build_pareto_worst_chart(
    df: pd.DataFrame, rank_metric: str, active_cross_cell: Optional[str]
) -> go.Figure:
    """Horizontal Pareto Bar Chart of Top 10 Worst Cells supporting click Cross-Filtering."""
    metric_cfg = {
        "Overall DCR (%)": ("DCR_Overall", "DCR_Denom", 30, "%", True),
        "eNB DCR (%)": ("DCR_eNB", "DCR_Denom", 30, "%", True),
        "Drops / Session Hr": ("DCR_Norm", "pmSessionTimeUe", 3600, "/hr", True),
        "Lowest HOSR (%)": ("HOSR", "HO_Denom", 20, "%", False),
        "Low CQI (<7) Ratio (%)": ("Low_CQI_Ratio", "CQI_Samples", 100, "%", True),
        "RACH OOC Ratio (%)": ("RACH_OOC_Ratio", "pmRaAttCbra", 100, "%", True),
        "IRAT Ping-Pong (%)": ("IRAT_PingPong_Rate", "IRAT_Denom", 1, "%", True),
        "Lowest CSSR (%)": ("CSSR", "CSSR_Denom", 30, "%", False),
        "VoLTE DCR (%)": ("VoLTE_DCR", "VoLTE_DCR_Denom", 10, "%", True),
        "Lowest VoLTE CSSR (%)": ("VoLTE_CSSR", "VoLTE_CSSR_Denom", 10, "%", False),
    }
    col, denom_col, min_d, unit, desc = metric_cfg[rank_metric]

    valid_sub = df[df[denom_col] >= min_d].copy()
    if valid_sub.empty:
        valid_sub = df.copy()

    top10 = valid_sub.sort_values(col, ascending=not desc).head(10).iloc[::-1]

    fig = go.Figure()
    if top10.empty:
        fig.update_layout(paper_bgcolor="#151D2E", plot_bgcolor="#0B0F19", height=342)
        return fig

    colors_list = []
    line_widths = []
    for _, r in top10.iterrows():
        c_name = str(r["EUtranCellFDD"])
        is_sel = active_cross_cell == c_name
        if r["Status"] == "CRITICAL":
            base_c = "#FF4D4F" if (active_cross_cell is None or is_sel) else "rgba(255,77,79,0.35)"
        elif r["Status"] == "DEGRADED":
            base_c = "#FAAD14" if (active_cross_cell is None or is_sel) else "rgba(250,173,20,0.35)"
        else:
            base_c = "#38BDF8" if (active_cross_cell is None or is_sel) else "rgba(56,189,248,0.35)"
        colors_list.append(base_c)
        line_widths.append(2.0 if is_sel else 0.0)

    fig.add_trace(
        go.Bar(
            x=top10[col],
            y=top10["EUtranCellFDD"],
            customdata=np.stack([top10["EUtranCellFDD"], top10["eNodeB_Name"], top10["Primary_Root_Cause"]], axis=-1),
            orientation="h",
            text=[f"{v:.2f}{unit}" for v in top10[col]],
            textposition="auto",
            textfont=dict(family=PLOTLY_FONT_MONO, size=10.5, color="#F8FAFC"),
            marker=dict(
                color=colors_list,
                line=dict(color="#FFFFFF", width=line_widths),
                cornerradius=4,
            ),
            hovertemplate=(
                "<b>%{customdata[0]}</b> (%{customdata[1]})<br>"
                f"{rank_metric}: <b>%{{x:.2f}}{unit}</b><br>"
                "RCA: %{customdata[2]}<br>"
                "<i>Click bar to cross-filter &amp; inspect</i><extra></extra>"
            ),
        )
    )

    fig.update_layout(
        title=dict(
            text=f"Top 10 Worst Cells — {rank_metric}",
            font=dict(family="Space Grotesk, sans-serif", size=13, color="#F8FAFC"),
            x=0.01,
            y=0.97,
        ),
        paper_bgcolor="#151D2E",
        plot_bgcolor="#0B0F19",
        font=dict(family=PLOTLY_FONT_MONO, size=10.5, color="#CBD5E1"),
        height=300,
        margin=dict(l=142, r=18, t=36, b=30),
        xaxis=dict(showgrid=True, gridcolor="#1E293B", zeroline=False),
        yaxis=dict(showgrid=False, tickfont=dict(family=PLOTLY_FONT_MONO, size=10.5, color="#F1F5F9")),
        clickmode="event+select",
    )
    return fig


# ============================================================================
# 4. SESSION STATE MANAGEMENT & FILTER ENGINE
# ============================================================================

def init_session_state() -> None:
    defaults = {
        "triage_preset": "All Degraded",
        "sel_enodebs": [],
        "sel_cells": [],
        "sel_rca_cats": [],
        "cross_filter_cell": None,
        "inspected_cell": None,
        "drawer_open": False,
        "last_grid_sel": None,
        "last_pareto_sel": None,
        "pareto_metric": "Overall DCR (%)",
        "pdf_cache_bytes": None,
        "pdf_cache_name": "",
        "pdf_cache_engine": "",
        "filter_gen": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_all_filters() -> None:
    """Restore all sidebar presets, multi-selects, and chart cross-filters to default."""
    st.session_state["triage_preset"] = "All Degraded"
    st.session_state["sel_enodebs"] = []
    st.session_state["sel_cells"] = []
    st.session_state["sel_rca_cats"] = []
    st.session_state["cross_filter_cell"] = None
    st.session_state["filter_gen"] = st.session_state.get("filter_gen", 0) + 1


def apply_filters(cell_df: pd.DataFrame) -> pd.DataFrame:
    """Apply Quick Triage Preset + Granular Multi-Selects."""
    df = cell_df.copy()
    preset = st.session_state["triage_preset"]

    if preset == "All Degraded":
        df = df[df["Status"].isin(["CRITICAL", "DEGRADED"])]
    elif preset == "Drop Call Spike (DCR > 1%)":
        df = df[
            (df["DCR_Denom"] >= 30)
            & ((df["DCR_Overall"] > 1.0) | ((df["pmSessionTimeUe"] >= 3600) & (df["DCR_Norm"] > 10.0)))
        ]
    elif preset == "Handover / RACH Issue (HOSR < 95%)":
        df = df[
            ((df["HO_Denom"] >= 20) & (df["HOSR"] < 95.0))
            | ((df["CSSR_Denom"] >= 30) & (df["RACH_OOC_Ratio"] > 5.0))
        ]
    elif preset == "Accessibility Failure (CSSR < 97%)":
        df = df[(df["CSSR_Denom"] >= 30) & (df["CSSR"] < 97.0)]
    elif preset == "IRAT Ping-Pong Loop (> 5%)":
        df = df[(df["IRAT_Denom"] >= 10) & (df["IRAT_PingPong_Rate"] > 5.0)]
    elif preset == "Radio Quality / CQI Degraded":
        df = df[(df["CQI_Samples"] >= 100) & ((df["Avg_CQI"] < 7.5) | (df["Low_CQI_Ratio"] > 30.0) | (df["UL_Intf_dBm"] > -110.0))]

    if st.session_state["sel_enodebs"]:
        df = df[df["eNodeB_Name"].isin(st.session_state["sel_enodebs"])]
    if st.session_state["sel_cells"]:
        df = df[df["EUtranCellFDD"].isin(st.session_state["sel_cells"])]
    if st.session_state["sel_rca_cats"]:
        df = df[df["RCA_Category"].isin(st.session_state["sel_rca_cats"])]

    return df.reset_index(drop=True)


def status_badge_html(status: str) -> str:
    st_u = str(status).upper()
    if st_u == "CRITICAL":
        return '<span class="status-badge badge-critical">● CRITICAL</span>'
    if st_u == "DEGRADED":
        return '<span class="status-badge badge-degraded">▲ DEGRADED</span>'
    if st_u == "NORMAL":
        return '<span class="status-badge badge-normal">✓ NORMAL</span>'
    return '<span class="status-badge badge-low">◌ LOW SAMPLE</span>'


# ============================================================================
# 5. MAIN COCKPIT APPLICATION
# ============================================================================

def main() -> None:
    inject_cockpit_css()
    init_session_state()

    # ------------------------------------------------------------------
    # COLLAPSIBLE SIDEBAR: DATA INGESTION & MULTI-TIER FILTERING
    # ------------------------------------------------------------------
    with st.sidebar:
        render_html(
            """
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
              <div style="width:34px;height:34px;border-radius:8px;background:#151D2E;border:1px solid #38BDF8;display:flex;align-items:center;justify-content:center;font-weight:700;color:#38BDF8;font-family:'JetBrains Mono',monospace;">E///</div>
              <div>
                <div style="font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:0.98rem;color:#F8FAFC;">RAN RCA COCKPIT</div>
                <div style="font-size:0.70rem;color:#94A3B8;font-family:'JetBrains Mono',monospace;">ENM / AMOS Telemetry v6.0</div>
              </div>
            </div>
            """
        )

        st.markdown("##### 📥 PM Counter Ingestion")
        uploaded_file = st.file_uploader(
            "Drag & Drop ENM / AMOS Dump (.xlsx, .xls, .csv)",
            type=["xlsx", "xls", "csv"],
            help="Supports multi-sheet ENIQ/ENM dumps (CSSR_DCR, CqiDistr, InterferencePwrPrb) and metadata header rows.",
        )

        # Load uploaded file or auto-ingest default Data/Counters.xlsx
        try:
            if uploaded_file is not None:
                with st.spinner("Executing Vectorized ETL & RCA Decision Trees..."):
                    result = load_cached_dump(uploaded_file.getvalue(), uploaded_file.name)
            elif DEFAULT_DATASET_PATH.exists():
                with st.spinner("Ingesting Data/Counters.xlsx..."):
                    result = load_default_dump(
                        str(DEFAULT_DATASET_PATH), DEFAULT_DATASET_PATH.stat().st_mtime
                    )
            else:
                st.warning("Please upload an Ericsson ENM/AMOS CSV or Excel counter dump to begin.")
                st.stop()
        except Exception as exc:
            st.error(f"Data Ingestion Error: {exc}")
            st.stop()

        full_df = result.cell_df
        gen_key = st.session_state["filter_gen"]

        st.markdown("---")
        st.markdown("##### ⚡ Quick Triage Presets")
        selected_preset = st.radio(
            "Diagnostic Preset",
            options=TRIAGE_PRESETS,
            index=TRIAGE_PRESETS.index(st.session_state["triage_preset"]),
            key=f"radio_preset_{gen_key}",
            label_visibility="collapsed",
        )
        if selected_preset != st.session_state["triage_preset"]:
            st.session_state["triage_preset"] = selected_preset
            st.session_state["cross_filter_cell"] = None
            st.rerun()

        st.markdown("##### 🎯 Granular Multi-Select")
        all_enbs = sorted(full_df["eNodeB_Name"].unique().tolist())
        sel_enbs = st.multiselect(
            "eNodeB_Name (ERBS Id)",
            options=all_enbs,
            default=st.session_state["sel_enodebs"],
            key=f"ms_enb_{gen_key}",
            placeholder="All eNodeBs...",
        )
        st.session_state["sel_enodebs"] = sel_enbs

        avail_cells_df = full_df[full_df["eNodeB_Name"].isin(sel_enbs)] if sel_enbs else full_df
        all_cells = sorted(avail_cells_df["EUtranCellFDD"].unique().tolist())
        sel_cells = st.multiselect(
            "EUtranCellFDD",
            options=all_cells,
            default=[c for c in st.session_state["sel_cells"] if c in all_cells],
            key=f"ms_cell_{gen_key}",
            placeholder="Search EUtranCellFDD...",
        )
        st.session_state["sel_cells"] = sel_cells

        present_cats = [c for c in full_df["RCA_Category"].unique().tolist() if c]
        sel_cats = st.multiselect(
            "Root Cause Group (RCA)",
            options=sorted(present_cats),
            default=[c for c in st.session_state["sel_rca_cats"] if c in present_cats],
            key=f"ms_rca_{gen_key}",
            placeholder="Filter by RCA Cause...",
        )
        st.session_state["sel_rca_cats"] = sel_cats

        st.markdown("")
        if st.button("↺ Reset All Filters", width="stretch"):
            reset_all_filters()
            st.rerun()

    # Apply sidebar filters
    sidebar_filtered_df = apply_filters(full_df)

    # Apply Pareto Cross-Filter on top for the Master Table if active
    cross_cell = st.session_state.get("cross_filter_cell")
    if cross_cell and cross_cell in full_df["EUtranCellFDD"].values:
        grid_df = full_df[full_df["EUtranCellFDD"] == cross_cell].reset_index(drop=True)
    else:
        grid_df = sidebar_filtered_df

    # Ensure inspected_cell always points to a valid cell
    if (
        st.session_state["inspected_cell"] is None
        or st.session_state["inspected_cell"] not in full_df["EUtranCellFDD"].values
    ):
        default_pool = grid_df if not grid_df.empty else full_df
        st.session_state["inspected_cell"] = str(default_pool.iloc[0]["EUtranCellFDD"])

    # Support URL query param ?inspect=1 or ?inspect=<EUtranCellFDD>
    qp_inspect = st.query_params.get("inspect")
    if qp_inspect and not st.session_state.get("_qp_applied"):
        st.session_state["_qp_applied"] = True
        if qp_inspect in full_df["EUtranCellFDD"].values:
            st.session_state["inspected_cell"] = str(qp_inspect)
        st.session_state["drawer_open"] = True

    # ------------------------------------------------------------------
    # TOP TELEMETRY HEADER
    # ------------------------------------------------------------------
    hdr_left, hdr_right = st.columns([9.2, 2.8], gap="small")
    with hdr_left:
        sheets_str = " + ".join(result.sheets_loaded)
        render_html(
            f"""
            <div class="telemetry-bar">
              <div class="telemetry-brand">
                <svg width="38" height="38" viewBox="0 0 44 44" fill="none">
                  <rect width="44" height="44" rx="9" fill="#0B0F19" stroke="#38BDF8" stroke-width="1.5"/>
                  <path d="M11 14L33 9.5M11 24.5L33 20M11 35L33 30.5" stroke="#38BDF8" stroke-width="4" stroke-linecap="round"/>
                  <circle cx="35" cy="9" r="3.2" fill="#FF4D4F"/>
                </svg>
                <div>
                  <div class="telemetry-title">ERICSSON LTE RAN TELEMETRY &amp; AUTOMATED RCA COCKPIT</div>
                  <div class="telemetry-sub">
                    <span class="meta-chip">📂 {html.escape(result.source_name)}</span>
                    <span class="meta-chip">🕒 ROP: {html.escape(result.data_start_time)} &rarr; {html.escape(result.data_end_time)}</span>
                    <span class="meta-chip">📑 Sheets: {html.escape(sheets_str)}</span>
                  </div>
                </div>
              </div>
              <div class="fresh-pill">
                <span class="pulse-dot"></span>
                <span>DATA FRESHNESS: VERIFIED &bull; {result.total_rop_rows:,} ROPs</span>
              </div>
            </div>
            """
        )

    with hdr_right:
        if st.button("📄 Export Report to PDF", type="primary", width="stretch"):
            with st.spinner("Compiling Branded HTML → PDF..."):
                pdf_target = grid_df if not grid_df.empty else full_df
                pdf_bytes, engine_used = export_report_to_pdf(result, pdf_target)
                dl_name = f"Ericsson_RAN_RCA_{Path(result.source_name).stem}.pdf"
                st.session_state["pdf_cache_bytes"] = pdf_bytes
                st.session_state["pdf_cache_name"] = dl_name
                st.session_state["pdf_cache_engine"] = engine_used
                trigger_pdf_autodownload(pdf_bytes, dl_name)

        if st.session_state.get("pdf_cache_bytes"):
            st.download_button(
                label=f"⬇ Save PDF ({st.session_state['pdf_cache_engine']})",
                data=st.session_state["pdf_cache_bytes"],
                file_name=st.session_state["pdf_cache_name"],
                mime="application/pdf",
                width="stretch",
            )

    # ------------------------------------------------------------------
    # TIER 1: EXECUTIVE METRIC BAR (4 COMPACT NOC CARDS)
    # ------------------------------------------------------------------
    total_cells = len(full_df)
    total_enbs = full_df["eNodeB_Name"].nunique()
    low_sample_cnt = int((full_df["Status"] == "LOW_SAMPLE").sum())

    worst_mask = full_df["Status"].isin(["CRITICAL", "DEGRADED"])
    worst_cnt = int(worst_mask.sum())
    crit_cnt = int((full_df["Status"] == "CRITICAL").sum())
    deg_cnt = int((full_df["Status"] == "DEGRADED").sum())
    worst_pct = (worst_cnt / total_cells * 100.0) if total_cells > 0 else 0.0

    valid_cssr_cells = full_df[full_df["CSSR_Denom"] >= 30]
    avg_cssr = (
        float(valid_cssr_cells["CSSR"].mean())
        if not valid_cssr_cells.empty
        else float(full_df["CSSR"].mean())
    )
    cssr_delta = avg_cssr - 98.5

    valid_dcr_cells = full_df[full_df["DCR_Denom"] >= 30]
    avg_dcr_ovr = (
        float(valid_dcr_cells["DCR_Overall"].mean())
        if not valid_dcr_cells.empty
        else float(full_df["DCR_Overall"].mean())
    )
    avg_dcr_enb = (
        float(valid_dcr_cells["DCR_eNB"].mean())
        if not valid_dcr_cells.empty
        else float(full_df["DCR_eNB"].mean())
    )
    avg_dcr_norm = float(full_df["DCR_Norm"].mean())

    c1, c2, c3, c4 = st.columns(4, gap="small")

    with c1:
        render_html(
            f"""
            <div class="kpi-card">
              <div class="kpi-label">Total Evaluated Cells</div>
              <div class="kpi-value-row">
                <div class="kpi-value">{total_cells:,}</div>
                <span class="status-badge badge-normal">{total_enbs} eNodeBs</span>
              </div>
              <div class="kpi-foot">
                <span>Active Filter: <strong class="mono-num" style="color:#F8FAFC;">{len(grid_df)}</strong> cells</span>
                <span class="mono-num">Low-Vol Excl: {low_sample_cnt}</span>
              </div>
            </div>
            """
        )

    with c2:
        w_cls = "card-crit" if crit_cnt > 0 else ("card-deg" if deg_cnt > 0 else "card-norm")
        w_badge = (
            f'<span class="status-badge badge-critical">{worst_pct:.1f}% OF RAN</span>'
            if worst_cnt > 0
            else '<span class="status-badge badge-normal">0.0% CLEAR</span>'
        )
        render_html(
            f"""
            <div class="kpi-card {w_cls}">
              <div class="kpi-label">Worst Cells Detected</div>
              <div class="kpi-value-row">
                <div class="kpi-value" style="color:{'#FF4D4F' if crit_cnt > 0 else '#F8FAFC'};">{worst_cnt:,}</div>
                {w_badge}
              </div>
              <div class="kpi-foot">
                <span class="mono-num" style="color:#FF4D4F;">● CRIT: {crit_cnt}</span>
                <span class="mono-num" style="color:#FAAD14;">▲ DEG: {deg_cnt}</span>
              </div>
            </div>
            """
        )

    with c3:
        cssr_card_cls = "card-norm" if avg_cssr >= 98.5 else ("card-deg" if avg_cssr >= 97.0 else "card-crit")
        delta_badge_cls = "badge-normal" if cssr_delta >= 0 else ("badge-degraded" if avg_cssr >= 97.0 else "badge-critical")
        render_html(
            f"""
            <div class="kpi-card {cssr_card_cls}">
              <div class="kpi-label">System Accessibility (Avg CSSR)</div>
              <div class="kpi-value-row">
                <div class="kpi-value">{avg_cssr:.2f}%</div>
                <span class="status-badge {delta_badge_cls}">{cssr_delta:+.2f}% vs 98.5%</span>
              </div>
              <div class="kpi-foot">
                <span>RRC &times; S1-C &times; Init E-RAB</span>
                <span class="mono-num">HOSR: {full_df['HOSR'].mean():.2f}%</span>
              </div>
            </div>
            """
        )

    with c4:
        dcr_breach = avg_dcr_ovr > 1.0
        dcr_card_cls = "card-crit" if avg_dcr_ovr > 2.5 else ("card-deg" if dcr_breach else "card-norm")
        dcr_badge = (
            '<span class="status-badge badge-critical">⚠ ALERT &gt; 1.0%</span>'
            if dcr_breach
            else '<span class="status-badge badge-normal">✓ &le; 1.0% OK</span>'
        )
        render_html(
            f"""
            <div class="kpi-card {dcr_card_cls}">
              <div class="kpi-label">System Retainability (Overall DCR)</div>
              <div class="kpi-value-row">
                <div class="kpi-value" style="color:{'#FF4D4F' if avg_dcr_ovr > 2.5 else ('#FAAD14' if dcr_breach else '#52C41A')};">{avg_dcr_ovr:.2f}%</div>
                {dcr_badge}
              </div>
              <div class="kpi-foot">
                <span class="mono-num">eNB DCR: {avg_dcr_enb:.2f}%</span>
                <span class="mono-num">{avg_dcr_norm:.1f} drops/hr</span>
              </div>
            </div>
            """
        )

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # ------------------------------------------------------------------
    # TIER 2: ANALYTICS & DIAGNOSTICS GRID (7 : 5 RATIO)
    # ------------------------------------------------------------------
    col_analytics, col_pareto = st.columns([7, 5], gap="small")

    with col_analytics:
        render_html(
            """
            <div class="panel-header">
              <span class="panel-title">📈 Radio Quality &amp; ROP Telemetry Diagnostics</span>
              <span class="panel-hint">3GPP Multi-Dimensional View</span>
            </div>
            """
        )
        tab_names = ["🕒 ROP Trendline (CSSR vs DCR)", "📊 CQI Distribution Stacked Bar"]
        if result.has_prb_spectrum:
            tab_names.append("📡 UL PRB Interference Spectrum")

        tabs = st.tabs(tab_names)

        with tabs[0]:
            if cross_cell:
                ts_df = compute_rop_series(result.rop_trend_df, [cross_cell])
                scope_lbl = f"Cell {cross_cell}"
            elif len(sidebar_filtered_df) < len(full_df) and not sidebar_filtered_df.empty:
                ts_df = compute_rop_series(
                    result.rop_trend_df, sidebar_filtered_df["EUtranCellFDD"].tolist()
                )
                scope_lbl = f"Filtered Cohort ({len(sidebar_filtered_df)} Cells)"
            else:
                ts_df = compute_rop_series(result.rop_trend_df, None)
                scope_lbl = f"Network Aggregate ({total_cells} Cells)"

            st.plotly_chart(
                build_rop_trend_chart(ts_df, scope_lbl),
                width="stretch",
                config={"displayModeBar": False},
            )

        with tabs[1]:
            cqi_source = grid_df if not grid_df.empty else full_df
            st.plotly_chart(
                build_cqi_stacked_chart(cqi_source),
                width="stretch",
                config={"displayModeBar": False},
            )

        if result.has_prb_spectrum and len(tabs) > 2:
            with tabs[2]:
                target_prb_cell = st.session_state["inspected_cell"]
                prb_vals = result.prb_cell_map.get(target_prb_cell, [])
                if prb_vals:
                    st.plotly_chart(
                        build_prb_spectrum_chart(target_prb_cell, prb_vals),
                        width="stretch",
                        config={"displayModeBar": False},
                    )
                else:
                    st.info("Select a cell with PRB interference measurements to inspect PRB 1–100 noise floor.")

    with col_pareto:
        render_html(
            """
            <div class="panel-header">
              <span class="panel-title">🔥 Top 10 Worst Cells Pareto</span>
              <span class="panel-hint">Click bar to Cross-Filter &amp; Slide-Over</span>
            </div>
            """
        )
        p_col1, p_col2 = st.columns([7.5, 4.5], gap="small")
        with p_col1:
            pareto_options = [
                "Overall DCR (%)",
                "eNB DCR (%)",
                "Drops / Session Hr",
                "Lowest HOSR (%)",
                "Low CQI (<7) Ratio (%)",
                "RACH OOC Ratio (%)",
                "Lowest CSSR (%)",
                "IRAT Ping-Pong (%)",
            ]
            if result.has_volte_counters:
                pareto_options.extend(["VoLTE DCR (%)", "Lowest VoLTE CSSR (%)"])
            if st.session_state["pareto_metric"] not in pareto_options:
                st.session_state["pareto_metric"] = pareto_options[0]
            rank_by = st.selectbox(
                "Rank Pareto By",
                options=pareto_options,
                index=pareto_options.index(st.session_state["pareto_metric"]),
                label_visibility="collapsed",
            )
            st.session_state["pareto_metric"] = rank_by
        with p_col2:
            if cross_cell:
                if st.button(f"✕ Clear: {cross_cell[:13]}", width="stretch"):
                    st.session_state["cross_filter_cell"] = None
                    st.rerun()
            else:
                st.markdown(
                    "<div style='padding-top:7px;font-size:0.72rem;color:#64748B;text-align:right;' class='mono-num'>Cross-Filter: Ready</div>",
                    unsafe_allow_html=True,
                )

        pareto_pool = sidebar_filtered_df if not sidebar_filtered_df.empty else full_df
        pareto_fig = build_pareto_worst_chart(pareto_pool, rank_by, cross_cell)
        pareto_event = st.plotly_chart(
            pareto_fig,
            width="stretch",
            key=f"pareto_chart_{gen_key}_{rank_by}",
            on_select="rerun",
            selection_mode="points",
            config={"displayModeBar": False},
        )

        # Handle Pareto Bar Click -> Cross-Filter Master Grid & Open Slide-Over Drawer
        sel_points = (
            pareto_event.get("selection", {}).get("points", [])
            if isinstance(pareto_event, dict)
            else getattr(getattr(pareto_event, "selection", None), "points", [])
        )
        pareto_sig = (gen_key, rank_by, sel_points[0].get("y") if sel_points else None)
        if pareto_sig != st.session_state.get("last_pareto_sel"):
            st.session_state["last_pareto_sel"] = pareto_sig
            if sel_points and sel_points[0].get("y"):
                clicked_y = str(sel_points[0].get("y"))
                st.session_state["cross_filter_cell"] = clicked_y
                st.session_state["inspected_cell"] = clicked_y
                st.session_state["drawer_open"] = True
                st.rerun()

    render_html("<div style='height:6px;'></div>")

    # ------------------------------------------------------------------
    # TIER 3: DATA-DENSE MASTER GRID & OFF-CANVAS SLIDE-OVER INSPECTOR
    # ------------------------------------------------------------------
    tb1, tb2, tb3, tb4 = st.columns([4.6, 3.2, 2.4, 1.8], gap="small")
    with tb1:
        filter_tag = (
            f"Cross-Filtered: {cross_cell}"
            if cross_cell
            else f"{st.session_state['triage_preset']} ({len(grid_df)} cells)"
        )
        render_html(
            f"""
            <div class="panel-header" style="margin-bottom:4px;">
              <span class="panel-title">🗂 Worst Cells Master Telemetry Grid</span>
              <span class="panel-hint">{html.escape(filter_tag)}</span>
            </div>
            """
        )

    with tb2:
        # Quick [Diagnose] Cell Selector synced with row click & Pareto click
        selector_pool = (
            grid_df["EUtranCellFDD"].tolist()
            if not grid_df.empty
            else full_df["EUtranCellFDD"].tolist()
        )
        if st.session_state["inspected_cell"] not in selector_pool and selector_pool:
            selector_idx = 0
        else:
            selector_idx = (
                selector_pool.index(st.session_state["inspected_cell"])
                if selector_pool
                else 0
            )

        chosen_diag = st.selectbox(
            "[Diagnose] Target Cell",
            options=selector_pool if selector_pool else ["NONE"],
            index=selector_idx,
            format_func=lambda x: f"🔍 [Diagnose] {x}",
            label_visibility="collapsed",
        )
        if chosen_diag != "NONE" and chosen_diag != st.session_state["inspected_cell"]:
            st.session_state["inspected_cell"] = chosen_diag
            st.session_state["drawer_open"] = True
            st.rerun()

    with tb3:
        # Consolidated "Copy All Commands" for currently filtered degraded cells
        degraded_in_view = (
            grid_df[grid_df["Status"].isin(["CRITICAL", "DEGRADED"])]
            if not grid_df.empty
            else full_df[full_df["Status"].isin(["CRITICAL", "DEGRADED"])]
        )
        batch_amos = "\n\n".join(degraded_in_view.head(25)["AMOS_Script"].tolist())
        if not batch_amos:
            batch_amos = "# No degraded cells requiring AMOS remediation in current filter."
        render_copy_button(
            batch_amos,
            f"Copy All AMOS ({min(25, len(degraded_in_view))})",
            key_id="batch_amos_top",
        )

    with tb4:
        drawer_lbl = "⇥ Close Slide-Over" if st.session_state["drawer_open"] else "⇤ Slide-Over Inspector"
        if st.button(drawer_lbl, width="stretch"):
            st.session_state["drawer_open"] = not st.session_state["drawer_open"]
            st.rerun()

    # Full-Width Master Interactive Grid
    if grid_df.empty:
        st.info(
            "No cells match the active filter criteria (e.g., 0 cells breached this specific threshold in the current dataset). "
            "Switch Quick Triage Preset or click 'Reset All Filters' in the sidebar."
        )
    else:
        status_icons = {
            "CRITICAL": "🔴 CRITICAL",
            "DEGRADED": "🟠 DEGRADED",
            "NORMAL": "🟢 NORMAL",
            "LOW_SAMPLE": "⚪ LOW_SAMPLE",
        }
        grid_cols_data = {
            "Action": ["[Diagnose] ▸"] * len(grid_df),
            "Status": grid_df["Status"].map(status_icons).fillna(grid_df["Status"]),
            "eNodeB_Name": grid_df["eNodeB_Name"],
            "EUtranCellFDD": grid_df["EUtranCellFDD"],
            "CSSR (%)": grid_df["CSSR"].round(2),
            "Overall DCR (%)": grid_df["DCR_Overall"].round(2),
            "eNB DCR (%)": grid_df["DCR_eNB"].round(2),
            "HOSR (%)": grid_df["HOSR"].round(2),
            "Avg CQI": grid_df["Avg_CQI"].round(2),
            "UL Intf (dBm)": grid_df["UL_Intf_dBm"].round(1),
            "IRAT PP (%)": grid_df["IRAT_PingPong_Rate"].round(2),
        }
        if result.has_volte_counters:
            grid_cols_data["VoLTE CSSR (%)"] = grid_df["VoLTE_CSSR"].round(2)
            grid_cols_data["VoLTE DCR (%)"] = grid_df["VoLTE_DCR"].round(2)
        grid_cols_data["RCA Type"] = grid_df["RCA_Category"]
        grid_cols_data["Root Cause Diagnosis"] = grid_df["Primary_Root_Cause"]
        grid_cols_data["Confidence"] = grid_df["Confidence_Score"]
        display_df = pd.DataFrame(grid_cols_data)

        grid_event = st.dataframe(
            display_df,
            width="stretch",
            height=470,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key=f"master_grid_{gen_key}_{cross_cell}_{st.session_state['triage_preset']}",
            column_config={
                "Action": st.column_config.TextColumn("Inspect", width="small"),
                "Status": st.column_config.TextColumn("Severity", width="small"),
                "eNodeB_Name": st.column_config.TextColumn("eNodeB_Name", width="medium"),
                "EUtranCellFDD": st.column_config.TextColumn("EUtranCellFDD", width="medium"),
                "CSSR (%)": st.column_config.NumberColumn("CSSR %", format="%.2f%%"),
                "Overall DCR (%)": st.column_config.NumberColumn("Overall DCR %", format="%.2f%%"),
                "eNB DCR (%)": st.column_config.NumberColumn("eNB DCR %", format="%.2f%%"),
                "HOSR (%)": st.column_config.NumberColumn("HOSR %", format="%.2f%%"),
                "Avg CQI": st.column_config.NumberColumn("Avg CQI", format="%.2f"),
                "UL Intf (dBm)": st.column_config.NumberColumn("UL Noise", format="%.1f dBm"),
                "IRAT PP (%)": st.column_config.NumberColumn("IRAT PP", format="%.2f%%"),
                "VoLTE CSSR (%)": st.column_config.NumberColumn("VoLTE CSSR", format="%.2f%%"),
                "VoLTE DCR (%)": st.column_config.NumberColumn("VoLTE DCR", format="%.2f%%"),
                "RCA Type": st.column_config.TextColumn("RCA Category", width="medium"),
                "Root Cause Diagnosis": st.column_config.TextColumn("Primary Root Cause", width="large"),
                "Confidence": st.column_config.ProgressColumn(
                    "Confidence",
                    format="%d%%",
                    min_value=0,
                    max_value=100,
                ),
            },
        )

        sel_rows = (
            grid_event.get("selection", {}).get("rows", [])
            if isinstance(grid_event, dict)
            else getattr(getattr(grid_event, "selection", None), "rows", [])
        )
        grid_sig = (gen_key, cross_cell, st.session_state["triage_preset"], tuple(sel_rows))
        if grid_sig != st.session_state.get("last_grid_sel"):
            st.session_state["last_grid_sel"] = grid_sig
            if sel_rows:
                row_idx = sel_rows[0]
                if 0 <= row_idx < len(grid_df):
                    picked_cell = str(grid_df.iloc[row_idx]["EUtranCellFDD"])
                    st.session_state["inspected_cell"] = picked_cell
                    st.session_state["drawer_open"] = True
                    st.rerun()

    # ------------------------------------------------------------------
    # OFF-CANVAS FIXED SLIDE-OVER DIAGNOSTIC & AMOS TERMINAL DRAWER
    # ------------------------------------------------------------------
    if st.session_state["drawer_open"]:
        with st.container(key="slide_over_drawer"):
            active_cell_name = st.session_state["inspected_cell"]
            cell_matches = full_df[full_df["EUtranCellFDD"] == active_cell_name]
            cell_row = cell_matches.iloc[0] if not cell_matches.empty else full_df.iloc[0]

            top_d1, top_d2 = st.columns([8.2, 3.8], gap="small")
            with top_d1:
                render_html(
                    """
                    <div style="padding-top:5px;font-family:'JetBrains Mono',monospace;font-size:0.72rem;letter-spacing:0.08em;color:#38BDF8;font-weight:700;">
                      ⚡ OFF-CANVAS CELL INSPECTOR &bull; RCA EVIDENCE
                    </div>
                    """
                )
            with top_d2:
                if st.button("✕ Close Drawer", key="btn_close_slide_over", width="stretch"):
                    st.session_state["drawer_open"] = False
                    st.rerun()

            st_val = str(cell_row["Status"])
            conf_val = int(cell_row["Confidence_Score"])
            cqi_str = f"{cell_row['Avg_CQI']:.2f}" if pd.notna(cell_row["Avg_CQI"]) else "N/A"
            intf_str = f"{cell_row['UL_Intf_dBm']:.1f}" if pd.notna(cell_row["UL_Intf_dBm"]) else "N/A"

            ev_rows_html = "".join(
                f"<tr>"
                f"<td class='mono'>{html.escape(str(ev['Counter / Indicator']))}</td>"
                f"<td class='mono' style='color:#38BDF8;font-weight:600;'>{html.escape(str(ev['Value']))}</td>"
                f"<td style='color:#94A3B8;'>{html.escape(str(ev['Meaning']))}</td>"
                f"</tr>"
                for ev in cell_row["Diagnostic_Evidence"]
            )

            secondary_html = (
                f"<div style='margin-top:5px;color:#FAAD14;font-size:0.75rem;'>"
                f"<strong>Secondary Triggers:</strong> {html.escape(str(cell_row['Secondary_Causes']))}</div>"
                if cell_row["Secondary_Causes"] not in ("None", "")
                else ""
            )

            amos_text = str(cell_row["AMOS_Script"])
            amos_lines_html = format_amos_terminal_html(amos_text)

            render_html(
                f"""
                <div class="inspector-drawer drawer-{html.escape(st_val)}">
                  <div class="drawer-top">
                    <div>
                      <div class="drawer-cell-id">{html.escape(str(cell_row['EUtranCellFDD']))}</div>
                      <div class="drawer-enb-id">eNodeB: {html.escape(str(cell_row['eNodeB_Name']))} &bull; {html.escape(str(cell_row['RCA_Category']))}</div>
                    </div>
                    <div style="text-align:right;">
                      {status_badge_html(st_val)}
                      <div class="mono-num" style="font-size:0.72rem;color:#94A3B8;margin-top:5px;">
                        Confidence: <strong style="color:#F8FAFC;">{conf_val}%</strong>
                      </div>
                    </div>
                  </div>
                  <div class="mini-kpi-grid">
                    <div class="mini-kpi">
                      <div class="mk-lbl">CSSR</div>
                      <div class="mk-val" style="color:{'#FF4D4F' if cell_row['CSSR']<97 else '#52C41A'};">{cell_row['CSSR']:.2f}%</div>
                    </div>
                    <div class="mini-kpi">
                      <div class="mk-lbl">Overall DCR</div>
                      <div class="mk-val" style="color:{'#FF4D4F' if cell_row['DCR_Overall']>2.5 else ('#FAAD14' if cell_row['DCR_Overall']>1.0 else '#52C41A')};">{cell_row['DCR_Overall']:.2f}%</div>
                    </div>
                    <div class="mini-kpi">
                      <div class="mk-lbl">eNB DCR</div>
                      <div class="mk-val" style="color:{'#FF4D4F' if cell_row['DCR_eNB']>2.5 else ('#FAAD14' if cell_row['DCR_eNB']>1.0 else '#52C41A')};">{cell_row['DCR_eNB']:.2f}%</div>
                    </div>
                    <div class="mini-kpi">
                      <div class="mk-lbl">HOSR</div>
                      <div class="mk-val" style="color:{'#FF4D4F' if cell_row['HOSR']<90 else ('#FAAD14' if cell_row['HOSR']<95 else '#F8FAFC')};">{cell_row['HOSR']:.2f}%</div>
                    </div>
                    <div class="mini-kpi">
                      <div class="mk-lbl">Avg CQI</div>
                      <div class="mk-val">{cqi_str}</div>
                    </div>
                    <div class="mini-kpi">
                      <div class="mk-lbl">UL Noise</div>
                      <div class="mk-val" style="color:{'#FF4D4F' if pd.notna(cell_row['UL_Intf_dBm']) and cell_row['UL_Intf_dBm']>-110 else '#38BDF8'};">{intf_str} dBm</div>
                    </div>
                  </div>
                  <div class="rca-callout">
                    <div style="color:#F8FAFC;font-weight:600;">🎯 {html.escape(str(cell_row['Primary_Root_Cause']))}</div>
                    <div style="color:#94A3B8;margin-top:4px;">💡 <strong>Remediation:</strong> {html.escape(str(cell_row['Recommendation']))}</div>
                    {secondary_html}
                  </div>
                  <div style="font-size:0.70rem;font-weight:700;letter-spacing:0.06em;color:#94A3B8;margin-bottom:5px;">
                    DIAGNOSTIC PM COUNTER EVIDENCE
                  </div>
                  <table class="ev-grid">
                    <thead>
                      <tr><th>PM Counter / Vector</th><th>Measured</th><th>Operational Diagnostic</th></tr>
                    </thead>
                    <tbody>{ev_rows_html}</tbody>
                  </table>
                  <div class="amos-terminal-wrap">
                    <div class="amos-term-bar">
                      <div class="amos-dots">
                        <span style="background:#FF5F56;"></span>
                        <span style="background:#FFBD2E;"></span>
                        <span style="background:#27C93F;"></span>
                        <strong style="color:#E2E8F0;margin-left:4px;">moshell@{html.escape(str(cell_row['eNodeB_Name']))}</strong>
                      </div>
                      <span>AMOS / MOM CLI</span>
                    </div>
                    <div class="amos-code-box">{amos_lines_html}</div>
                  </div>
                </div>
                """
            )

            render_copy_button(
                amos_text,
                f"One-Click Copy AMOS Command ({cell_row['EUtranCellFDD']})",
                key_id=f"copy_single_{active_cell_name.replace('-', '_')}",
            )


if __name__ == "__main__":
    main()
