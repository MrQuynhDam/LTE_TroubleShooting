"""
Executive HTML-to-PDF Report Exporter for Ericsson LTE RAN Troubleshooting Cockpit.

Coordinates `html-to-pdf` print-CSS standards (Headless Chromium `--headless=new`
with exact color reproduction, WeasyPrint/Playwright support, and pure-Python
ReportLab fallback) so color-coded CRITICAL/DEGRADED/NORMAL telemetry badges and
AMOS/Moshell terminal remediation blocks always export cleanly to PDF.
"""

from __future__ import annotations

import html
import io
import os
import shutil
import subprocess
import uuid
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

from rca_engine import AnalysisResult


def _find_chromium_binary() -> Optional[str]:
    """Locate Chrome, Chromium, or Edge binary across Windows/Linux/macOS."""
    candidates = [
        shutil.which("google-chrome"),
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        shutil.which("chrome"),
        shutil.which("msedge"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def _status_badge_html(status: str) -> str:
    st_upper = str(status).upper()
    if st_upper == "CRITICAL":
        return '<span class="badge badge-critical">CRITICAL</span>'
    if st_upper == "DEGRADED":
        return '<span class="badge badge-degraded">DEGRADED</span>'
    if st_upper == "NORMAL":
        return '<span class="badge badge-normal">NORMAL</span>'
    return '<span class="badge badge-low">LOW SAMPLE</span>'


def build_html_report(
    result: AnalysisResult,
    filtered_df: pd.DataFrame,
    max_deep_dive: int = 12,
) -> str:
    """
    Construct a self-contained, print-engineered HTML5 report preserving NOC
    semantic status badges, tabular monospace KPI columns, and AMOS CLI blocks.
    """
    full_df = result.cell_df
    total_cells = len(full_df)
    low_sample_cnt = int((full_df["Status"] == "LOW_SAMPLE").sum())
    worst_df = full_df[full_df["Status"].isin(["CRITICAL", "DEGRADED"])]
    worst_cnt = len(worst_df)
    crit_cnt = int((full_df["Status"] == "CRITICAL").sum())
    deg_cnt = int((full_df["Status"] == "DEGRADED").sum())
    worst_pct = (worst_cnt / total_cells * 100.0) if total_cells > 0 else 0.0

    valid_cssr_df = full_df[full_df["CSSR_Denom"] >= 30]
    avg_cssr = float(valid_cssr_df["CSSR"].mean()) if not valid_cssr_df.empty else float(full_df["CSSR"].mean())
    cssr_delta = avg_cssr - 98.5

    valid_dcr_df = full_df[full_df["DCR_Denom"] >= 30]
    avg_dcr_ovr = (
        float(valid_dcr_df["DCR_Overall"].mean())
        if not valid_dcr_df.empty
        else float(full_df["DCR_Overall"].mean())
    )
    avg_dcr_enb = (
        float(valid_dcr_df["DCR_eNB"].mean())
        if not valid_dcr_df.empty
        else float(full_df["DCR_eNB"].mean())
    )
    avg_hosr = float(full_df["HOSR"].mean()) if not full_df.empty else 100.0
    avg_cqi = float(full_df["Avg_CQI"].dropna().mean()) if full_df["Avg_CQI"].notna().any() else 0.0

    # Target table rows (prioritize filtered view if non-empty, else worst cells)
    report_df = filtered_df.copy() if not filtered_df.empty else worst_df.copy()
    if report_df.empty:
        report_df = full_df.head(25).copy()

    # Root cause distribution summary
    rca_counts = (
        worst_df["RCA_Category"].value_counts()
        if not worst_df.empty
        else report_df["RCA_Category"].value_counts()
    )
    rca_pills_html = "".join(
        f'<div class="rca-pill"><span class="rca-name">{html.escape(str(cat))}</span>'
        f'<span class="rca-val">{int(cnt)} cells</span></div>'
        for cat, cnt in rca_counts.items()
    )

    # Matrix table rows (up to 45 worst/filtered cells for clean executive pagination)
    matrix_rows: List[str] = []
    for _, r in report_df.head(45).iterrows():
        cqi_txt = f"{r['Avg_CQI']:.2f}" if pd.notna(r.get("Avg_CQI")) else "N/A"
        intf_txt = f"{r['UL_Intf_dBm']:.1f}" if pd.notna(r.get("UL_Intf_dBm")) else "N/A"
        conf_txt = f"{int(r['Confidence_Score'])}%" if int(r.get("Confidence_Score", 0)) > 0 else "N/A"
        dcr_o = float(r.get("DCR_Overall", r.get("DCR", 0.0)))
        dcr_e = float(r.get("DCR_eNB", 0.0))
        dcr_cls_cell = (
            "txt-crit"
            if dcr_o > 2.5 and r["DCR_Denom"] >= 30
            else ("txt-deg" if dcr_o > 1.0 and r["DCR_Denom"] >= 30 else "")
        )
        matrix_rows.append(
            f"""
            <tr>
              <td>{_status_badge_html(r['Status'])}</td>
              <td class="mono">{html.escape(str(r['eNodeB_Name']))}</td>
              <td class="mono fw-600">{html.escape(str(r['EUtranCellFDD']))}</td>
              <td class="mono num">{r['CSSR']:.2f}%</td>
              <td class="mono num {dcr_cls_cell}">{dcr_o:.2f}%</td>
              <td class="mono num">{dcr_e:.2f}%</td>
              <td class="mono num">{r['HOSR']:.2f}%</td>
              <td class="mono num">{cqi_txt}</td>
              <td class="mono num">{intf_txt}</td>
              <td>{html.escape(str(r['RCA_Category']))}</td>
              <td>{html.escape(str(r['Primary_Root_Cause']))}</td>
              <td class="mono num">{conf_txt}</td>
            </tr>
            """
        )

    # Deep-dive diagnostic & AMOS blocks for Top Degraded Cells in current selection
    deep_candidates = report_df[report_df["Status"].isin(["CRITICAL", "DEGRADED"])].head(max_deep_dive)
    if deep_candidates.empty:
        deep_candidates = report_df.head(min(5, len(report_df)))

    deep_cards: List[str] = []
    for idx, (_, r) in enumerate(deep_candidates.iterrows(), start=1):
        ev_list = r.get("Diagnostic_Evidence", [])
        ev_rows = "".join(
            f"<tr><td class='mono'>{html.escape(str(item['Counter / Indicator']))}</td>"
            f"<td class='mono fw-600'>{html.escape(str(item['Value']))}</td>"
            f"<td>{html.escape(str(item['Meaning']))}</td></tr>"
            for item in ev_list[:11]
        )
        amos_escaped = html.escape(str(r.get("AMOS_Script", "")))
        deep_cards.append(
            f"""
            <div class="deep-card">
              <div class="deep-header">
                <div>
                  <span class="deep-idx">#{idx:02d}</span>
                  <span class="mono deep-cell">{html.escape(str(r['EUtranCellFDD']))}</span>
                  <span class="mono deep-enb">({html.escape(str(r['eNodeB_Name']))})</span>
                </div>
                <div>
                  {_status_badge_html(r['Status'])}
                  <span class="conf-chip mono">Confidence: {int(r['Confidence_Score'])}%</span>
                </div>
              </div>
              <div class="deep-meta">
                <strong>Primary Root Cause:</strong> {html.escape(str(r['Primary_Root_Cause']))}<br/>
                <strong>Action Recommendation:</strong> {html.escape(str(r['Recommendation']))}
              </div>
              <div class="deep-grid">
                <div>
                  <div class="sub-label">DIAGNOSTIC EVIDENCE &amp; PM COUNTERS</div>
                  <table class="ev-table">
                    <thead><tr><th>PM Counter / Metric</th><th>Measured</th><th>Engineering Note</th></tr></thead>
                    <tbody>{ev_rows}</tbody>
                  </table>
                </div>
                <div>
                  <div class="sub-label">AUTOMATED AMOS / MOSHELL CLI PROTOCOL</div>
                  <pre class="amos-box">{amos_escaped}</pre>
                </div>
              </div>
            </div>
            """
        )

    doc_id = f"RAN-RCA-{pd.Timestamp.now().strftime('%Y%m%d-%H%M')}"
    cssr_cls = "txt-ok" if cssr_delta >= 0 else "txt-deg"
    dcr_cls = "txt-crit" if avg_dcr_ovr > 2.5 else ("txt-deg" if avg_dcr_ovr > 1.0 else "txt-ok")

    return f"""<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <title>Ericsson LTE RAN RCA Executive Report - {html.escape(result.source_name)}</title>
  <style>
    @page {{
      size: A4 landscape;
      margin: 11mm 12mm 11mm 12mm;
    }}
    html, body {{
      margin: 0;
      padding: 0;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
      font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif;
      color: #0F172A;
      background: #FFFFFF;
      font-size: 9pt;
      line-height: 1.42;
    }}
    .mono {{
      font-family: 'JetBrains Mono', 'Roboto Mono', Consolas, 'Courier New', monospace;
      font-variant-numeric: tabular-nums;
    }}
    .fw-600 {{ font-weight: 600; }}
    .num {{ text-align: right; }}
    .txt-crit {{ color: #DC2626; font-weight: 700; }}
    .txt-deg {{ color: #D97706; font-weight: 700; }}
    .txt-ok {{ color: #16A34A; font-weight: 700; }}

    /* Branded NOC Dark Header */
    .report-header {{
      background: linear-gradient(135deg, #0B0F19 0%, #151D2E 65%, #1E293B 100%);
      color: #F8FAFC;
      padding: 16px 22px;
      border-radius: 8px;
      border-bottom: 3px solid #38BDF8;
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 14px;
    }}
    .brand-left {{
      display: flex;
      align-items: center;
      gap: 14px;
    }}
    .brand-title h1 {{
      margin: 0;
      font-size: 15.5pt;
      letter-spacing: -0.02em;
      color: #FFFFFF;
    }}
    .brand-title p {{
      margin: 3px 0 0 0;
      font-size: 8.5pt;
      color: #94A3B8;
    }}
    .brand-meta {{
      text-align: right;
      font-size: 8pt;
      color: #CBD5E1;
      line-height: 1.55;
    }}

    /* Executive KPI Strip */
    .kpi-strip {{
      display: table;
      width: 100%;
      table-layout: fixed;
      border-spacing: 8px 0;
      margin: 0 -8px 14px -8px;
    }}
    .kpi-card {{
      display: table-cell;
      background: #0F172A;
      color: #F8FAFC;
      border: 1px solid #1E293B;
      border-top: 3px solid #38BDF8;
      border-radius: 6px;
      padding: 10px 14px;
      vertical-align: top;
    }}
    .kpi-card.crit-border {{ border-top-color: #FF4D4F; }}
    .kpi-card.deg-border {{ border-top-color: #FAAD14; }}
    .kpi-card.ok-border {{ border-top-color: #52C41A; }}
    .kpi-label {{
      font-size: 7.5pt;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: #94A3B8;
      margin-bottom: 4px;
    }}
    .kpi-val {{
      font-size: 16pt;
      font-weight: 700;
      color: #FFFFFF;
    }}
    .kpi-sub {{
      font-size: 7.8pt;
      color: #CBD5E1;
      margin-top: 3px;
    }}

    /* Status Badges matching Dark OLED Cockpit */
    .badge {{
      display: inline-block;
      padding: 2px 7px;
      border-radius: 4px;
      font-size: 7.2pt;
      font-weight: 700;
      letter-spacing: 0.04em;
      font-family: Consolas, monospace;
    }}
    .badge-critical {{
      color: #FF4D4F;
      background: rgba(255, 77, 79, 0.14);
      border: 1px solid rgba(255, 77, 79, 0.55);
    }}
    .badge-degraded {{
      color: #D97706;
      background: rgba(250, 173, 20, 0.16);
      border: 1px solid rgba(250, 173, 20, 0.60);
    }}
    .badge-normal {{
      color: #15803D;
      background: rgba(82, 196, 26, 0.15);
      border: 1px solid rgba(82, 196, 26, 0.55);
    }}
    .badge-low {{
      color: #475569;
      background: #F1F5F9;
      border: 1px solid #CBD5E1;
    }}

    .section-title {{
      font-size: 11pt;
      font-weight: 700;
      color: #0B0F19;
      border-left: 4px solid #0284C7;
      padding-left: 8px;
      margin: 14px 0 8px 0;
    }}

    .rca-bar {{
      margin-bottom: 12px;
    }}
    .rca-pill {{
      display: inline-block;
      background: #F8FAFC;
      border: 1px solid #CBD5E1;
      border-radius: 4px;
      padding: 3px 9px;
      margin: 0 6px 6px 0;
      font-size: 8pt;
    }}
    .rca-name {{ font-weight: 600; color: #1E293B; margin-right: 6px; }}
    .rca-val {{ font-family: Consolas, monospace; color: #0284C7; font-weight: 700; }}

    /* Master Table */
    table.master-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 7.8pt;
      margin-bottom: 16px;
    }}
    table.master-table th {{
      background: #0B0F19;
      color: #E2E8F0;
      text-align: left;
      padding: 6px 6px;
      font-weight: 600;
      border: 1px solid #1E293B;
    }}
    table.master-table td {{
      padding: 5px 6px;
      border: 1px solid #E2E8F0;
      vertical-align: middle;
    }}
    table.master-table tr:nth-child(even) td {{
      background: #F8FAFC;
    }}
    tr {{
      break-inside: avoid;
      page-break-inside: avoid;
    }}

    .page-break {{
      break-before: page;
      page-break-before: always;
    }}

    /* Deep Dive Cards */
    .deep-card {{
      border: 1px solid #CBD5E1;
      border-left: 4px solid #0B0F19;
      border-radius: 6px;
      padding: 10px 12px;
      margin-bottom: 12px;
      background: #FFFFFF;
      break-inside: avoid;
      page-break-inside: avoid;
    }}
    .deep-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 1px solid #E2E8F0;
      padding-bottom: 6px;
      margin-bottom: 6px;
    }}
    .deep-idx {{
      background: #0B0F19;
      color: #38BDF8;
      font-family: Consolas, monospace;
      font-weight: 700;
      padding: 2px 6px;
      border-radius: 4px;
      margin-right: 6px;
      font-size: 8pt;
    }}
    .deep-cell {{ font-size: 10pt; font-weight: 700; color: #0F172A; }}
    .deep-enb {{ font-size: 8.5pt; color: #64748B; margin-left: 4px; }}
    .conf-chip {{
      background: #0F172A;
      color: #38BDF8;
      padding: 2px 7px;
      border-radius: 4px;
      font-size: 7.5pt;
      margin-left: 6px;
    }}
    .deep-meta {{
      font-size: 8.2pt;
      color: #1E293B;
      background: #F8FAFC;
      padding: 6px 8px;
      border-radius: 4px;
      margin-bottom: 8px;
    }}
    .deep-grid {{
      display: table;
      width: 100%;
      table-layout: fixed;
    }}
    .deep-grid > div {{
      display: table-cell;
      vertical-align: top;
      width: 50%;
    }}
    .deep-grid > div:first-child {{
      padding-right: 8px;
    }}
    .deep-grid > div:last-child {{
      padding-left: 8px;
    }}
    .sub-label {{
      font-size: 7.2pt;
      font-weight: 700;
      letter-spacing: 0.06em;
      color: #475569;
      margin-bottom: 4px;
    }}
    table.ev-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 7.5pt;
    }}
    table.ev-table th {{
      background: #E2E8F0;
      color: #0F172A;
      text-align: left;
      padding: 4px 5px;
      border: 1px solid #CBD5E1;
    }}
    table.ev-table td {{
      padding: 3.5px 5px;
      border: 1px solid #E2E8F0;
    }}
    pre.amos-box {{
      margin: 0;
      background: #05080F;
      color: #38BDF8;
      border: 1px solid #1E293B;
      border-radius: 5px;
      padding: 7px 9px;
      font-family: 'JetBrains Mono', Consolas, monospace;
      font-size: 6.9pt;
      line-height: 1.35;
      white-space: pre-wrap;
      word-break: break-all;
    }}
    .footer-note {{
      margin-top: 10px;
      text-align: right;
      font-size: 7.5pt;
      color: #64748B;
      border-top: 1px solid #E2E8F0;
      padding-top: 6px;
    }}
  </style>
</head>
<body>

  <header class="report-header">
    <div class="brand-left">
      <svg width="42" height="42" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect width="44" height="44" rx="9" fill="#0F172A" stroke="#38BDF8" stroke-width="1.5"/>
        <path d="M11 14L33 9.5M11 24.5L33 20M11 35L33 30.5" stroke="#38BDF8" stroke-width="4" stroke-linecap="round"/>
        <circle cx="35" cy="9" r="3" fill="#FF4D4F"/>
      </svg>
      <div class="brand-title">
        <h1>ERICSSON LTE RAN PERFORMANCE &amp; RCA DIAGNOSTIC REPORT</h1>
        <p>Autonomous PM Counter Telemetry • Statistical Sample Gate • AMOS/Moshell Remediation</p>
      </div>
    </div>
    <div class="brand-meta mono">
      <div><strong>DOC ID:</strong> {html.escape(doc_id)}</div>
      <div><strong>SOURCE:</strong> {html.escape(result.source_name)}</div>
      <div><strong>ROP WINDOW:</strong> {html.escape(result.data_start_time)} &rarr; {html.escape(result.data_end_time)}</div>
      <div><strong>GENERATED:</strong> {html.escape(result.ingestion_timestamp)}</div>
    </div>
  </header>

  <section class="kpi-strip">
    <div class="kpi-card">
      <div class="kpi-label">Total Evaluated Cells</div>
      <div class="kpi-val mono">{total_cells:,}</div>
      <div class="kpi-sub">ROP Rows: {result.total_rop_rows:,} &bull; Low-Sample Excluded: {low_sample_cnt}</div>
    </div>
    <div class="kpi-card {'crit-border' if crit_cnt > 0 else 'deg-border'}">
      <div class="kpi-label">Worst Cells Detected</div>
      <div class="kpi-val mono" style="color:#FF4D4F;">{worst_cnt:,} <span style="font-size:10pt;color:#FAAD14;">({worst_pct:.1f}%)</span></div>
      <div class="kpi-sub">Critical: {crit_cnt} &bull; Degraded: {deg_cnt}</div>
    </div>
    <div class="kpi-card {'ok-border' if avg_cssr >= 98.5 else 'deg-border'}">
      <div class="kpi-label">System Accessibility (Avg CSSR)</div>
      <div class="kpi-val mono">{avg_cssr:.2f}%</div>
      <div class="kpi-sub">Baseline 98.50% (<span class="{cssr_cls}">{cssr_delta:+.2f}%</span>) &bull; HOSR: {avg_hosr:.2f}%</div>
    </div>
    <div class="kpi-card {'crit-border' if avg_dcr_ovr > 1.0 else 'ok-border'}">
      <div class="kpi-label">System Retainability (Overall DCR)</div>
      <div class="kpi-val mono" style="color:{'#FF4D4F' if avg_dcr_ovr > 1.0 else '#52C41A'};">{avg_dcr_ovr:.2f}%</div>
      <div class="kpi-sub">eNB DCR: {avg_dcr_enb:.2f}% &bull; Avg CQI: {avg_cqi:.2f}</div>
    </div>
  </section>

  <div class="section-title">1. Root Cause Taxonomy Distribution</div>
  <div class="rca-bar">
    {rca_pills_html if rca_pills_html else '<span class="rca-pill">No degraded root causes triggered</span>'}
  </div>

  <div class="section-title">2. Worst Cells &amp; Filtered Telemetry Matrix (Top {min(45, len(report_df))} Cells)</div>
  <table class="master-table">
    <thead>
      <tr>
        <th style="width:72px;">Status</th>
        <th style="width:105px;">eNodeB_Name</th>
        <th style="width:120px;">EUtranCellFDD</th>
        <th class="num" style="width:52px;">CSSR</th>
        <th class="num" style="width:56px;">Ovr DCR</th>
        <th class="num" style="width:56px;">eNB DCR</th>
        <th class="num" style="width:52px;">HOSR</th>
        <th class="num" style="width:44px;">CQI</th>
        <th class="num" style="width:58px;">UL dBm</th>
        <th style="width:128px;">RCA Category</th>
        <th>Primary Root Cause Diagnosis</th>
        <th class="num" style="width:45px;">Conf.</th>
      </tr>
    </thead>
    <tbody>
      {"".join(matrix_rows)}
    </tbody>
  </table>

  <div class="page-break"></div>
  <div class="section-title">3. Deep-Dive Cell Diagnostics &amp; Automated AMOS Remediation Scripts</div>
  {"".join(deep_cards)}

  <div class="footer-note mono">
    Antigravity Principal RAN Performance &amp; Optimization Suite &bull; 3GPP / Ericsson MOM Deterministic RCA Engine v6.0
  </div>

</body>
</html>
"""


def _render_via_chromium(html_content: str, browser_bin: str) -> Optional[bytes]:
    """Compile HTML to PDF via Headless Chromium/Edge (`--headless=new`)."""
    # Use isolated project runtime folder instead of shared /tmp
    state_dir = Path(__file__).resolve().parent / ".state" / "pdf_spool"
    state_dir.mkdir(parents=True, exist_ok=True)
    token = uuid.uuid4().hex[:12]
    html_file = state_dir / f"report_{token}.html"
    pdf_file = state_dir / f"report_{token}.pdf"

    try:
        html_file.write_text(html_content, encoding="utf-8")
        file_url = html_file.resolve().as_uri()
        cmd = [
            browser_bin,
            "--headless=new",
            "--no-sandbox",
            "--disable-gpu",
            "--no-pdf-header-footer",
            "--run-all-compositor-stages-before-draw",
            "--virtual-time-budget=4000",
            f"--print-to-pdf={str(pdf_file.resolve())}",
            file_url,
        ]
        subprocess.run(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=35,
            check=False,
        )
        if pdf_file.exists() and pdf_file.stat().st_size > 1000:
            return pdf_file.read_bytes()
    except Exception:
        return None
    finally:
        for f in (html_file, pdf_file):
            try:
                if f.exists():
                    f.unlink()
            except OSError:
                pass
    return None


def _render_via_reportlab(
    result: AnalysisResult, filtered_df: pd.DataFrame, max_deep_dive: int = 12
) -> bytes:
    """Pure-Python ReportLab fallback preserving status badge colors and AMOS blocks."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        PageBreak,
        Paragraph,
        Preformatted,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Ericsson LTE RAN RCA Report - {result.source_name}",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "NOC_Title",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=15,
        textColor=colors.HexColor("#F8FAFC"),
        spaceAfter=2,
    )
    sub_style = ParagraphStyle(
        "NOC_Sub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        textColor=colors.HexColor("#94A3B8"),
    )
    sec_style = ParagraphStyle(
        "SecHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        textColor=colors.HexColor("#0B0F19"),
        spaceBefore=10,
        spaceAfter=6,
    )
    cell_txt = ParagraphStyle(
        "CellTxt",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9.5,
    )
    amos_style = ParagraphStyle(
        "AmosCode",
        fontName="Courier",
        fontSize=7,
        leading=9,
        textColor=colors.HexColor("#38BDF8"),
        backColor=colors.HexColor("#0B0F19"),
        borderPadding=6,
    )

    story = []

    # Header Banner Table
    hdr_tbl = Table(
        [
            [
                Paragraph("ERICSSON LTE RAN PERFORMANCE & RCA EXECUTIVE REPORT", title_style),
                Paragraph(
                    f"<b>Source:</b> {html.escape(result.source_name)}<br/>"
                    f"<b>Window:</b> {result.data_start_time} to {result.data_end_time}",
                    sub_style,
                ),
            ]
        ],
        colWidths=[185 * mm, 88 * mm],
    )
    hdr_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0B0F19")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                ("LINEBELOW", (0, 0), (-1, -1), 2, colors.HexColor("#38BDF8")),
            ]
        )
    )
    story.append(hdr_tbl)
    story.append(Spacer(1, 5 * mm))

    full_df = result.cell_df
    worst_df = full_df[full_df["Status"].isin(["CRITICAL", "DEGRADED"])]
    report_df = filtered_df if not filtered_df.empty else worst_df
    if report_df.empty:
        report_df = full_df.head(25)

    story.append(Paragraph("1. Worst Cells Root-Cause Matrix", sec_style))

    tbl_data = [
        [
            "Status", "eNodeB_Name", "EUtranCellFDD", "CSSR", "Ovr DCR",
            "eNB DCR", "HOSR", "CQI", "UL dBm", "RCA Category", "Primary Root Cause", "Conf",
        ]
    ]
    row_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B0F19")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#F8FAFC")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.2),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    for idx, (_, r) in enumerate(report_df.head(40).iterrows(), start=1):
        st_val = str(r["Status"])
        dcr_o = float(r.get("DCR_Overall", r.get("DCR", 0.0)))
        dcr_e = float(r.get("DCR_eNB", 0.0))
        tbl_data.append(
            [
                st_val,
                str(r["eNodeB_Name"]),
                str(r["EUtranCellFDD"]),
                f"{r['CSSR']:.2f}%",
                f"{dcr_o:.2f}%",
                f"{dcr_e:.2f}%",
                f"{r['HOSR']:.2f}%",
                f"{r['Avg_CQI']:.1f}" if pd.notna(r.get("Avg_CQI")) else "N/A",
                f"{r['UL_Intf_dBm']:.1f}" if pd.notna(r.get("UL_Intf_dBm")) else "N/A",
                Paragraph(html.escape(str(r["RCA_Category"])), cell_txt),
                Paragraph(html.escape(str(r["Primary_Root_Cause"])), cell_txt),
                f"{int(r['Confidence_Score'])}%",
            ]
        )
        if st_val == "CRITICAL":
            row_styles.append(("BACKGROUND", (0, idx), (0, idx), colors.HexColor("#FEE2E2")))
            row_styles.append(("TEXTCOLOR", (0, idx), (0, idx), colors.HexColor("#DC2626")))
        elif st_val == "DEGRADED":
            row_styles.append(("BACKGROUND", (0, idx), (0, idx), colors.HexColor("#FEF3C7")))
            row_styles.append(("TEXTCOLOR", (0, idx), (0, idx), colors.HexColor("#D97706")))
        elif st_val == "NORMAL":
            row_styles.append(("BACKGROUND", (0, idx), (0, idx), colors.HexColor("#DCFCE7")))
            row_styles.append(("TEXTCOLOR", (0, idx), (0, idx), colors.HexColor("#15803D")))

    col_w = [17 * mm, 26 * mm, 30 * mm, 13 * mm, 13 * mm, 13 * mm, 13 * mm, 11 * mm, 14 * mm, 33 * mm, 78 * mm, 12 * mm]
    m_tbl = Table(tbl_data, colWidths=col_w, repeatRows=1)
    m_tbl.setStyle(TableStyle(row_styles))
    story.append(m_tbl)

    story.append(PageBreak())
    story.append(Paragraph("2. Deep-Dive Diagnostics & AMOS Remediation Scripts", sec_style))

    deep_df = report_df[report_df["Status"].isin(["CRITICAL", "DEGRADED"])].head(max_deep_dive)
    for _, r in deep_df.iterrows():
        story.append(
            Paragraph(
                f"<b>{html.escape(str(r['EUtranCellFDD']))}</b> ({html.escape(str(r['eNodeB_Name']))}) "
                f"— [{r['Status']}] {html.escape(str(r['Primary_Root_Cause']))}",
                 styles["Heading4"],
            )
        )
        story.append(Preformatted(str(r.get("AMOS_Script", "")), amos_style))
        story.append(Spacer(1, 3 * mm))

    doc.build(story)
    return buf.getvalue()


def export_report_to_pdf(
    result: AnalysisResult,
    filtered_df: pd.DataFrame,
    max_deep_dive: int = 12,
) -> Tuple[bytes, str]:
    """
    Generate production-grade PDF report from HTML/CSS template.
    Uses Headless Chromium/Edge first (per `html-to-pdf` skill), then WeasyPrint,
    then ReportLab fallback.
    """
    html_doc = build_html_report(result, filtered_df, max_deep_dive=max_deep_dive)

    # 1. Headless Chromium / Edge (`--headless=new`)
    browser_bin = _find_chromium_binary()
    if browser_bin:
        pdf_bytes = _render_via_chromium(html_doc, browser_bin)
        if pdf_bytes:
            return pdf_bytes, f"Headless Chromium ({Path(browser_bin).name})"

    # 2. WeasyPrint Python Engine (if available)
    try:
        import weasyprint  # type: ignore

        pdf_bytes = weasyprint.HTML(string=html_doc).write_pdf()
        if pdf_bytes:
            return pdf_bytes, "WeasyPrint"
    except Exception:
        pass

    # 3. ReportLab Pure-Python Fallback
    pdf_bytes = _render_via_reportlab(result, filtered_df, max_deep_dive=max_deep_dive)
    return pdf_bytes, "ReportLab Platypus"
