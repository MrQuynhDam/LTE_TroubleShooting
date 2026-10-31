"""
Ericsson LTE RAN Performance Troubleshooting & Root-Cause Prediction Engine (v6.0.0).

Implements ETL ingestion (CSV / multi-sheet Excel with metadata headers),
exact 3GPP/Ericsson MOM composite KPI calculations (CSSR, 3GPP Overall DCR,
eNB DCR, Session-Normalized Drop Rate, HOSR, CQI/HomCQI/MIMO/PRB Interference,
3G/4G IRAT Ping-Pong, and 23.Q3+ VoLTE QCI=1/QCI=5 KPIs), statistical sample-size
validation (Rule of Minimum Denominator), deterministic multi-domain RCA
decision trees (Trees 4.1 - 4.7), and automated AMOS/Moshell CLI remediation scripts.
"""

from __future__ import annotations

import csv
import io
import re
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, BinaryIO, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ============================================================================
# 1. CONSTANTS, BASELINES & ERICSSON MOM COUNTER TAXONOMY (SKILL v6.0.0)
# ============================================================================

# Ericsson 16-bin PDF midpoints (dBm/PRB) for pmRadioRecInterferencePwr[0..15]
INTF_PDF_MIDPOINTS_DBM = np.array(
    [
        -121.5, -120.5, -119.5, -118.5, -117.5, -116.5, -115.5, -114.5,
        -113.5, -112.5, -110.0, -106.0, -102.0, -98.0, -94.0, -90.0,
    ],
    dtype=np.float64,
)
INTF_PDF_MIDPOINTS_MW = np.power(10.0, INTF_PDF_MIDPOINTS_DBM / 10.0)

KPI_BASELINES = {
    "CSSR": {"target": 98.5, "degraded": 97.0, "critical": 95.0, "min_denom": 30},
    "DCR": {"target": 0.5, "degraded": 1.0, "critical": 2.5, "min_denom": 30},
    "DCR_NORM": {"target": 5.0, "degraded": 10.0, "critical": 25.0, "min_denom": 3600},
    "HOSR": {"target": 98.0, "degraded": 95.0, "critical": 90.0, "min_denom": 20},
    "CQI_AVG": {"target": 9.0, "degraded": 7.5, "critical": 6.0, "min_denom": 100},
    "CQI_LOW": {"target": 15.0, "degraded": 30.0, "critical": 50.0, "min_denom": 100},
    "IRAT_PP": {"target": 2.0, "degraded": 5.0, "critical": 10.0, "min_denom": 10},
    "VOLTE_CSSR": {"target": 99.0, "degraded": 97.5, "critical": 95.0, "min_denom": 10},
    "VOLTE_DCR": {"target": 0.3, "degraded": 0.8, "critical": 2.0, "min_denom": 10},
}

RCA_CATEGORIES = [
    "UL Interference / PIM",
    "Hardware / Baseband / RRU-RET",
    "License Congestion",
    "RACH Overshooting (OOC)",
    "Event B2/A2 & Mobility Tuning",
    "PRB / CCE / GBR Congestion",
    "RF Weak Coverage / UL Power",
    "Transport / Core (S1/MME)",
    "VoLTE QCI=1 Bearer / IMS",
    "Normal / Healthy",
    "Low Traffic Sample (Excluded)",
]

# Column normalization aliases mapping raw OSS/ENM/ENIQ headers -> canonical names
COLUMN_ALIASES: Dict[str, str] = {
    # Identifiers
    "ossid": "OSS_Id",
    "oss_id": "OSS_Id",
    "erbsid": "eNodeB_Name",
    "erbs_id": "eNodeB_Name",
    "erbs": "eNodeB_Name",
    "enodeb_name": "eNodeB_Name",
    "enodebname": "eNodeB_Name",
    "enodeb": "eNodeB_Name",
    "managedelement": "eNodeB_Name",
    "nodeid": "eNodeB_Name",
    "enb_name": "eNodeB_Name",
    "eutrancellid": "EUtranCellFDD",
    "eutrancell_id": "EUtranCellFDD",
    "eutrancellfdd": "EUtranCellFDD",
    "eutrancelltdd": "EUtranCellFDD",
    "cell_name": "EUtranCellFDD",
    "cellname": "EUtranCellFDD",
    "eutrancell": "EUtranCellFDD",
    "cellid": "Cell_Id",
    "cell_id": "Cell_Id",
    "earfcndl": "earfcndl",
    "earfcn": "earfcndl",
    # Time & Vector Index
    "date_time": "ROP_Time",
    "datetime": "ROP_Time",
    "datetime_id": "ROP_Time",
    "datetimeid": "ROP_Time",
    "rop_time": "ROP_Time",
    "roptime": "ROP_Time",
    "timestamp": "ROP_Time",
    "date": "Date",
    "date_id": "Date",
    "hour": "Hour",
    "hour_id": "Hour",
    "min": "Minute",
    "min_id": "Minute",
    "minute": "Minute",
    "vectorindex": "Vector_Index",
    "vector_index": "Vector_Index",
    "dcvector_index": "Vector_Index",
}

# Core scalar counters used across KPI & RCA formulas (v6.0.0 complete set)
SCALAR_COUNTERS: List[str] = [
    # 2.1 & 3.1 RACH & RRC Accessibility
    "pmRaAttCbra",
    "pmRaFailCbraMsg1DiscOoc",
    "pmRaFailCbraMsg1DiscSched",
    "pmRaFailCbraMsg2Disc",
    "pmRrcConnEstabAtt",
    "pmRrcConnEstabAttReatt",
    "pmRrcConnEstabSucc",
    "pmRrcConnEstabFailMpOverload",
    "pmRrcConnEstabFailLic",
    "pmRrcConnEstabFailDuIntens",
    "pmRrcConnEstabFailMISigQCong",
    "pmRrcConnEstabFailMmeOvlMos",
    "pmRrcConnEstabFailMmeOvlMod",
    # S1-C & Initial E-RAB
    "pmS1SigConnEstabAtt",
    "pmS1SigConnEstabSucc",
    "pmS1SigConnEstabFailMmeOvlMos",
    "pmErabEstabAttInit",
    "pmErabEstabSuccInit",
    "pmErabEstabFailInitLic",
    "pmErabEstabFailGbrDlEnb",
    "pmErabEstabFailGbrUlEnb",
    # 2.2 & 3.2 Retainability (3GPP Full Denominator + Active Breakdown)
    "pmErabRelAbnormalEnb",
    "pmErabRelNormalEnb",
    "pmErabRelMme",
    "pmErabRelNormalEnbAct",
    "pmErabRelAbnormalEnbAct",
    "pmErabRelAbnormalMmeAct",
    "pmErabRelAbnormalEnbActUeLost",
    "pmErabRelAbnormalEnbUeLostOos",
    "pmErabRelAbnormalEnbActHo",
    "pmErabRelAbnormalEnbActTnFail",
    "pmErabRelAbnormalEnbActPe",
    "pmSessionTimeUe",
    # 2.3 & 3.3 Mobility (IntraF, InterF, IRAT, 1xRTT SRVCC)
    "pmCellHoPrepAttLteIntraF",
    "pmCellHoPrepSuccLteIntraF",
    "pmCellHoExeAttLteIntraF",
    "pmCellHoExeSuccLteIntraF",
    "pmCellHoPrepAttLteInterF",
    "pmCellHoPrepSuccLteInterF",
    "pmCellHoExeAttLteInterF",
    "pmCellHoExeSuccLteInterF",
    "pmHoPrepAtt",
    "pmHoPrepSucc",
    "pmHoExeAtt",
    "pmHoExeSucc",
    "pmHoPrepAtt1xRttSrvcc",
    "pmHoPrepSucc1xRttSrvcc",
    "pmHoPrepRejInLicConnUsers",
    "pmHoPrepRejOutUlThres",
    # 2.5 & 3.5 IRAT 3G/4G Ping-Pong & Mobility Parameters
    "pmBadCovEvalReport",
    "pmBadCovSearchEvalReport",
    "pmB2BestCellEvalReport",
    "pmUeCtxtRelSCWcdma",
    "pmUeCtxtRelCsfbWcdma",
    "UtranCellRelation.pmHoPrepAtt",
    "UtranCellRelation.pmHoPrepSucc",
    "UtranCellRelation.pmHoExeAtt",
    "UtranCellRelation.pmHoExeSucc",
    "pmRapid3gTo4gReentries",
    "a2Threshold0Rsdp",
    "b2Threshold1Rsdp",
    "b2Threshold2RsdpUtran",
    "b2Threshold2EcNoUtran",
    "timeToTriggerB2",
    "hysteresisB2",
    # 2.6, 2.7 & 3.6 VoLTE QCI=1 (23.Q3+ Added Bearers & Legacy Fallbacks) + QCI=5
    "pmErabEstabAttAddedQci[1]",
    "pmErabEstabSuccAddedQci[1]",
    "pmErabEstabAttAddedHoOngoingQci[1]",
    "pmErabEstabAttAddedCsfbQci[1]",
    "pmErabEstabFailAddedErabIdInUseQci[1]",
    "pmErabEstabAttQci[1]",
    "pmErabEstabSuccQci[1]",
    "pmErabEstabAttInitQci[5]",
    "pmErabEstabSuccInitQci[5]",
    "pmErabRelAbnormalEnbActQci[1]",
    "pmErabRelAbnormalMmeActQci[1]",
    "pmErabRelAbnormalEnbQci[1]",
    "pmErabRelNormalEnbQci[1]",
    "pmErabRelMmeQci[1]",
    "pmErabRelNormalEnbActQci[1]",
    "pmErabRelAbnormalEnbActUeLostQci[1]",
    "pmErabRelAbnormalEnbActHoQci[1]",
    "pmErabRelAbnormalEnbActPeQci[1]",
    "pmSessionTimeDrbQci[1]",
    "pmRlcPduRetransUlQci[1]",
    "pmRlcPduTxUlQci[1]",
    "pmRlcPduRetransDlQci[1]",
    "pmHoPrepAttSrvcc",
    "pmHoExeSuccSrvcc",
]

PARAMETER_COLS = (
    "a2Threshold0Rsdp",
    "b2Threshold1Rsdp",
    "b2Threshold2RsdpUtran",
    "b2Threshold2EcNoUtran",
    "timeToTriggerB2",
    "hysteresisB2",
)

COUNTER_LOWER_MAP: Dict[str, str] = {c.lower(): c for c in SCALAR_COUNTERS}
COUNTER_LOWER_MAP.update(
    {
        # Legacy / variant OSS exports
        "pms1sigconnestabfailmmeovlmosce": "pmS1SigConnEstabFailMmeOvlMos",
        "utrancellrelation_pmhoprepatt": "UtranCellRelation.pmHoPrepAtt",
        "utrancellrelation_pmhoprepsucc": "UtranCellRelation.pmHoPrepSucc",
        "utrancellrelation_pmhoexeatt": "UtranCellRelation.pmHoExeAtt",
        "utrancellrelation_pmhoexesucc": "UtranCellRelation.pmHoExeSucc",
        "rapid_3g_to_4g_reentries": "pmRapid3gTo4gReentries",
        "rapid3gto4greentries": "pmRapid3gTo4gReentries",
        "pmsessiontimeueqci[1]": "pmSessionTimeDrbQci[1]",
        "pmsessiontimeueqci_1": "pmSessionTimeDrbQci[1]",
        "pmsessiontimedrbqci_1": "pmSessionTimeDrbQci[1]",
    }
)
# Register underscore/parenthesis QCI variants automatically
for _qci_c in [c for c in SCALAR_COUNTERS if "[" in c]:
    _base, _idx_part = _qci_c.split("[", 1)
    _idx = _idx_part.rstrip("]")
    COUNTER_LOWER_MAP[f"{_base.lower()}_{_idx}"] = _qci_c
    COUNTER_LOWER_MAP[f"{_base.lower()}({_idx})"] = _qci_c


@dataclass
class AnalysisResult:
    """Container for complete ETL, KPI, Time-Series, and RCA outputs."""

    source_name: str
    ingestion_timestamp: str
    data_start_time: str
    data_end_time: str
    total_rop_rows: int
    sheets_loaded: List[str]
    cell_df: pd.DataFrame
    rop_trend_df: pd.DataFrame
    cqi_dist_df: pd.DataFrame
    prb_cell_map: Dict[str, List[float]] = field(default_factory=dict)
    has_irat_counters: bool = False
    has_volte_counters: bool = False
    has_prb_spectrum: bool = False
    has_3gpp_dcr_denom: bool = True


# ============================================================================
# 2. ROBUST HEADER DETECTION & COLUMN NORMALIZATION
# ============================================================================

HEADER_KEYWORDS = {
    "oss id", "oss_id", "erbs id", "erbs", "enodeb_name", "enodeb",
    "managedelement", "eutrancell id", "eutrancellfdd", "eutrancelltdd",
    "cell_name", "cell id", "date", "hour", "rop_time", "date_time",
    "vector index",
}


def _detect_and_promote_header(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Scan first 20 rows of a raw headerless DataFrame to locate the real header row
    (handling ENIQ/ENM metadata titles and leading blank rows/columns).
    """
    if raw_df.empty:
        return raw_df

    max_scan = min(20, len(raw_df))
    best_idx = 0
    best_score = -1

    for idx in range(max_scan):
        row_vals = raw_df.iloc[idx].dropna().astype(str).str.strip()
        if row_vals.empty:
            continue
        score = 0
        for val in row_vals:
            v_low = val.lower()
            if v_low in HEADER_KEYWORDS or v_low.startswith("pm"):
                score += 2
            elif any(k in v_low for k in ("cell", "erbs", "enodeb", "date", "time")):
                score += 1
        if score > best_score:
            best_score = score
            best_idx = idx

    headers = [
        str(col).strip() if pd.notna(col) and str(col).strip() != "" else f"_drop_{i}"
        for i, col in enumerate(raw_df.iloc[best_idx])
    ]
    df = raw_df.iloc[best_idx + 1 :].copy()
    df.columns = headers

    keep_cols = [
        c for c in df.columns
        if not c.startswith("_drop_") and not c.lower().startswith("unnamed:")
    ]
    df = df[keep_cols].dropna(how="all").reset_index(drop=True)
    return df


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Standardize OSS/ENM/ENIQ column names into canonical Ericsson MOM schema."""
    df_out = df.copy()
    renamed: Dict[str, str] = {}

    for col in df_out.columns:
        c_strip = str(col).strip()
        c_key = re.sub(r"[\s\-]+", "_", c_strip.lower())
        c_compact = c_key.replace("_", "")

        if c_compact in COLUMN_ALIASES:
            renamed[col] = COLUMN_ALIASES[c_compact]
        elif c_key in COLUMN_ALIASES:
            renamed[col] = COLUMN_ALIASES[c_key]
        elif c_strip.lower() in COUNTER_LOWER_MAP:
            renamed[col] = COUNTER_LOWER_MAP[c_strip.lower()]
        elif c_key in COUNTER_LOWER_MAP:
            renamed[col] = COUNTER_LOWER_MAP[c_key]
        else:
            # Normalize QCI vector notation e.g. pmErabEstabAttAddedQci_qci1 / pmErabEstabAttAddedQci [1] -> [1]
            m_qci = re.match(r"^(pm[a-z0-9]+?qci)[_\[(\s]*(?:qci)?(\d+)[\])]?$", c_strip, re.IGNORECASE)
            # Normalize CQI/Rank/Interference vector notation e.g. pmRadioUeRepCqiDistr_0 -> [0]
            m_vec = re.match(r"^(pmradio\w+)[_\[(\s](\d+)[\])]?$", c_strip, re.IGNORECASE)
            if m_qci and f"{m_qci.group(1).lower()}[{m_qci.group(2)}]" in COUNTER_LOWER_MAP:
                renamed[col] = COUNTER_LOWER_MAP[f"{m_qci.group(1).lower()}[{m_qci.group(2)}]"]
            elif m_vec:
                base_raw, idx_str = m_vec.group(1).lower(), m_vec.group(2)
                canonical_base = {
                    "pmradiouerepcqidistr": "pmRadioUeRepCqiDistr",
                    "pmradiouereprankdistr": "pmRadioUeRepRankDistr",
                    "pmradiohomuerepcqidistr": "pmRadioHomUeRepCqiDistr",
                    "pmradiorecinterferencepwr": "pmRadioRecInterferencePwr",
                }.get(base_raw, m_vec.group(1))
                renamed[col] = f"{canonical_base}[{idx_str}]"
            else:
                m_prb = re.match(r"^pmradiorecinterferencepwrprb(\d+)$", c_compact)
                if m_prb:
                    renamed[col] = f"pmRadioRecInterferencePwrPrb{int(m_prb.group(1))}"
                elif c_strip.lower() == "pmradiouerepcqidistr":
                    renamed[col] = "pmRadioUeRepCqiDistr"
                elif c_strip.lower() == "pmradiouereprankdistr":
                    renamed[col] = "pmRadioUeRepRankDistr"
                elif c_strip.lower() == "pmradiohomuerepcqidistr":
                    renamed[col] = "pmRadioHomUeRepCqiDistr"
                elif c_strip.lower() == "pmradiorecinterferencepwr":
                    renamed[col] = "pmRadioRecInterferencePwr"
                else:
                    renamed[col] = c_strip

    df_out = df_out.rename(columns=renamed)
    if df_out.columns.duplicated().any():
        df_out = df_out.loc[:, ~df_out.columns.duplicated(keep="first")]
    return df_out


def _build_rop_timestamp(df: pd.DataFrame) -> pd.Series:
    """Construct unified datetime ROP_Time series from Date+Hour or ROP_Time."""
    if "ROP_Time" in df.columns:
        ts = pd.to_datetime(df["ROP_Time"], errors="coerce")
        if ts.notna().any():
            return ts

    if "Date" in df.columns:
        raw_date = df["Date"]
        if pd.api.types.is_numeric_dtype(raw_date):
            base_dt = pd.to_datetime(raw_date, unit="D", origin="1899-12-30", errors="coerce")
        else:
            num_try = pd.to_numeric(raw_date, errors="coerce")
            if num_try.notna().mean() > 0.8 and num_try.median() > 30000:
                base_dt = pd.to_datetime(num_try, unit="D", origin="1899-12-30", errors="coerce")
            else:
                base_dt = pd.to_datetime(raw_date, errors="coerce")

        if "Hour" in df.columns:
            hr = pd.to_numeric(df["Hour"], errors="coerce").fillna(0)
            base_dt = base_dt + pd.to_timedelta(hr, unit="h")
        if "Minute" in df.columns:
            mn = pd.to_numeric(df["Minute"], errors="coerce").fillna(0)
            base_dt = base_dt + pd.to_timedelta(mn, unit="m")
        return base_dt

    return pd.Series(pd.Timestamp("2026-10-28 00:00:00"), index=df.index)


# ============================================================================
# 3. MULTI-SHEET / CSV INGESTION & VECTOR PIVOTING PIPELINE
# ============================================================================

def _classify_sheet(df: pd.DataFrame) -> str:
    """Classify normalized DataFrame into scalar KPI, long vector CQI, long QCI, or PRB spectrum."""
    cols = set(df.columns)
    prb_cols = [c for c in cols if c.startswith("pmRadioRecInterferencePwrPrb")]
    non_prb_scalars = [c for c in cols if c in SCALAR_COUNTERS]
    if len(prb_cols) >= 10 and len(non_prb_scalars) == 0:
        return "prb_spectrum"
    if "Vector_Index" in cols and (
        "pmRadioUeRepCqiDistr" in cols or "pmRadioRecInterferencePwr" in cols
    ):
        return "cqi_vector_long"
    if "Vector_Index" in cols and any(c.endswith("Qci") for c in cols):
        return "qci_vector_long"
    if any(c.startswith("pm") or c.startswith("UtranCellRelation.") for c in cols):
        return "kpi_main"
    return "unknown"


def _pivot_qci_long_sheet(df_qci: pd.DataFrame) -> pd.DataFrame:
    """Pivot long-format QCI vector sheets (Vector_Index = QCI 1..9) into canonical `[1]` / `[5]` columns."""
    df = df_qci.copy()
    df["ROP_Time"] = _build_rop_timestamp(df)
    df["Vector_Index"] = pd.to_numeric(df["Vector_Index"], errors="coerce")
    df = df[df["Vector_Index"].isin([1, 5])].copy()
    if df.empty or "EUtranCellFDD" not in df.columns:
        return pd.DataFrame()
    df["Vector_Index"] = df["Vector_Index"].astype(int)

    qci_cols = [c for c in df.columns if c.startswith("pm") and c.endswith("Qci")]
    for c in qci_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    id_cols = [c for c in ("eNodeB_Name", "EUtranCellFDD", "ROP_Time") if c in df.columns]
    piv = df.pivot_table(index=id_cols, columns="Vector_Index", values=qci_cols, aggfunc="sum", fill_value=0.0)
    piv.columns = [f"{col_name}[{qci_idx}]" for col_name, qci_idx in piv.columns]
    return piv.reset_index()


def _process_cqi_long_sheet(df_cqi: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Vectorized transformation of long-format CqiDistr sheet (Vector_Index 0..15)
    into:
      1) Per-Cell aggregated CQI bins [0..15], 256-QAM HomCQI, MIMO Rank 1 ratio, and UL Interference (dBm)
      2) Per-ROP_Time Network CQI trendline metrics
    """
    df = df_cqi.copy()
    df["ROP_Time"] = _build_rop_timestamp(df)
    if "eNodeB_Name" not in df.columns and "EUtranCellFDD" in df.columns:
        df["eNodeB_Name"] = df["EUtranCellFDD"].astype(str).str.slice(0, 14)

    df["Vector_Index"] = pd.to_numeric(df["Vector_Index"], errors="coerce")
    df = df[df["Vector_Index"].isin(range(16))].copy()
    df["Vector_Index"] = df["Vector_Index"].astype(int)

    for col in (
        "pmRadioUeRepCqiDistr",
        "pmRadioHomUeRepCqiDistr",
        "pmRadioUeRepRankDistr",
        "pmRadioRecInterferencePwr",
    ):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        else:
            df[col] = 0.0

    # --- A. Cell-level CQI Bins (0..15), 256-QAM HomCQI, Rank 1..4, and Interference PDF ---
    cqi_cell_piv = (
        df.pivot_table(
            index="EUtranCellFDD",
            columns="Vector_Index",
            values="pmRadioUeRepCqiDistr",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reindex(columns=range(16), fill_value=0.0)
    )

    cqi_tot = cqi_cell_piv.sum(axis=1)
    cqi_weighted = cqi_cell_piv.values @ np.arange(16, dtype=np.float64)
    avg_cqi = np.where(cqi_tot > 0, cqi_weighted / np.maximum(cqi_tot, 1.0), np.nan)
    low_cqi_pct = np.where(
        cqi_tot > 0, cqi_cell_piv[list(range(7))].sum(axis=1) / np.maximum(cqi_tot, 1.0) * 100.0, 0.0
    )

    # 256-QAM High-Order Modulation CQI array (pmRadioHomUeRepCqiDistr[0..15])
    hom_cell_piv = (
        df.pivot_table(
            index="EUtranCellFDD",
            columns="Vector_Index",
            values="pmRadioHomUeRepCqiDistr",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reindex(columns=range(16), fill_value=0.0)
    )
    hom_tot = hom_cell_piv.reindex(cqi_cell_piv.index).fillna(0.0).sum(axis=1)
    hom_high_pct = np.where(
        hom_tot > 0,
        hom_cell_piv.reindex(cqi_cell_piv.index).fillna(0.0)[list(range(12, 16))].sum(axis=1)
        / np.maximum(hom_tot, 1.0)
        * 100.0,
        np.nan,
    )

    # MIMO Rank 1 Ratio (bins 0..3 correspond to Rank 1..4)
    rank_cell_piv = (
        df.pivot_table(
            index="EUtranCellFDD",
            columns="Vector_Index",
            values="pmRadioUeRepRankDistr",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reindex(columns=range(4), fill_value=0.0)
    )
    rank_tot = rank_cell_piv.sum(axis=1)
    rank1_pct = np.where(rank_tot > 0, rank_cell_piv[0] / np.maximum(rank_tot, 1.0) * 100.0, np.nan)

    # UL Interference Power (dBm/PRB) from 16-bin PDF
    intf_cell_piv = (
        df.pivot_table(
            index="EUtranCellFDD",
            columns="Vector_Index",
            values="pmRadioRecInterferencePwr",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reindex(columns=range(16), fill_value=0.0)
    )
    intf_tot = intf_cell_piv.sum(axis=1)
    intf_mw = np.where(
        intf_tot > 0,
        (intf_cell_piv.values @ INTF_PDF_MIDPOINTS_MW) / np.maximum(intf_tot, 1.0),
        np.nan,
    )
    ul_intf_dbm = np.where(
        intf_mw > 0,
        10.0 * np.log10(np.maximum(np.nan_to_num(intf_mw, nan=1e-30), 1e-30)),
        np.nan,
    )

    cell_radio_data: Dict[str, Any] = {
        "EUtranCellFDD": cqi_cell_piv.index,
        "CQI_Samples": cqi_tot.values,
        "Avg_CQI": avg_cqi,
        "Low_CQI_Ratio": low_cqi_pct,
        "Hom_256QAM_Ratio": hom_high_pct,
        "Rank1_Ratio": rank1_pct,
        "UL_Intf_dBm": ul_intf_dbm,
        "Intf_Samples": intf_tot.values,
    }
    for b in range(16):
        cell_radio_data[f"CQI_Bin_{b}"] = cqi_cell_piv[b].values
    cell_radio_df = pd.DataFrame(cell_radio_data)

    # --- B. ROP-level Network CQI for Trendlines ---
    rop_cqi_piv = (
        df.pivot_table(
            index="ROP_Time",
            columns="Vector_Index",
            values="pmRadioUeRepCqiDistr",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reindex(columns=range(16), fill_value=0.0)
    )
    rop_tot = rop_cqi_piv.sum(axis=1)
    rop_avg_cqi = np.where(
        rop_tot > 0, (rop_cqi_piv.values @ np.arange(16, dtype=np.float64)) / np.maximum(rop_tot, 1.0), np.nan
    )
    rop_low_cqi = np.where(
        rop_tot > 0, rop_cqi_piv[list(range(7))].sum(axis=1) / np.maximum(rop_tot, 1.0) * 100.0, 0.0
    )
    rop_radio_df = pd.DataFrame(
        {
            "ROP_Time": rop_cqi_piv.index,
            "Net_Avg_CQI": rop_avg_cqi,
            "Net_Low_CQI_Ratio": rop_low_cqi,
        }
    )

    return cell_radio_df, rop_radio_df


def _process_prb_spectrum_sheet(
    df_prb: pd.DataFrame, cell_intf_samples: Optional[pd.Series] = None
) -> Tuple[pd.DataFrame, Dict[str, List[float]]]:
    """
    Convert pmRadioRecInterferencePwrPrb0..100 into calibrated per-PRB dBm profiles
    and extract PIM / Narrowband crest metrics per cell.
    """
    df = df_prb.copy()
    prb_cols_map: Dict[int, str] = {}
    for c in df.columns:
        m = re.match(r"^pmRadioRecInterferencePwrPrb(\d+)$", c)
        if m:
            prb_cols_map[int(m.group(1))] = c

    if not prb_cols_map or "EUtranCellFDD" not in df.columns:
        return pd.DataFrame(), {}

    sorted_prbs = sorted(prb_cols_map.keys())
    ordered_cols = [prb_cols_map[p] for p in sorted_prbs]
    for c in ordered_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    rop_counts = df.groupby("EUtranCellFDD").size()
    prb_sum = df.groupby("EUtranCellFDD")[ordered_cols].sum()

    # 1 hour ROP = 3,600,000 ms -> 90,000 samples of 40ms
    if cell_intf_samples is not None and not cell_intf_samples.empty:
        aligned_samples = cell_intf_samples.reindex(prb_sum.index).fillna(0.0)
        n_40ms = np.where(
            aligned_samples > 0,
            aligned_samples / 40.0,
            rop_counts.reindex(prb_sum.index).values * 90000.0,
        )
    else:
        n_40ms = rop_counts.reindex(prb_sum.index).values * 90000.0

    # Ericsson unit: 2^-44 mW per 40ms sample
    mw_matrix = (prb_sum.values * (2.0 ** -44)) / np.maximum(n_40ms[:, None], 1.0)
    dbm_matrix = np.where(mw_matrix > 0, 10.0 * np.log10(np.maximum(mw_matrix, 1e-30)), -121.0)

    prb_max = np.nanmax(dbm_matrix, axis=1)
    prb_med = np.nanmedian(dbm_matrix, axis=1)
    prb_crest = prb_max - prb_med

    peak_texts: List[str] = []
    prb_cell_map: Dict[str, List[float]] = {}

    for row_idx, cell_name in enumerate(prb_sum.index):
        row_dbm = dbm_matrix[row_idx]
        prb_cell_map[str(cell_name)] = np.round(row_dbm, 2).tolist()
        top3_idx = np.argsort(row_dbm)[::-1][:3]
        parts = [f"PRB{sorted_prbs[i]} ({row_dbm[i]:.1f} dBm)" for i in top3_idx]
        peak_texts.append(", ".join(parts))

    prb_summary = pd.DataFrame(
        {
            "EUtranCellFDD": prb_sum.index,
            "PRB_Max_dBm": np.round(prb_max, 2),
            "PRB_Median_dBm": np.round(prb_med, 2),
            "PRB_Crest_dB": np.round(prb_crest, 2),
            "Peak_PRBs_Text": peak_texts,
        }
    )
    return prb_summary, prb_cell_map


# ============================================================================
# 4. COMPOSITE KPI COMPUTATION & DETERMINISTIC DECISION TREES (v6.0.0)
# ============================================================================

def _diagnose_cell_row(
    row: pd.Series, has_irat: bool, has_volte: bool
) -> Dict[str, Any]:
    """
    Execute Step 0 Statistical Sample Gate (Section 2.8) + Sequential Decision
    Trees (Sections 4.1 - 4.7) for a single aggregated cell row.
    """
    enb = str(row.get("eNodeB_Name", "UNKNOWN_ENB"))
    cell = str(row.get("EUtranCellFDD", "UNKNOWN_CELL"))

    cssr = float(row.get("CSSR", 100.0))
    cssr_denom = float(row.get("CSSR_Denom", 0.0))
    rrc_sr = float(row.get("RRC_SR", 100.0))
    s1_sr = float(row.get("S1_SR", 100.0))
    erab_sr = float(row.get("ERAB_SR", 100.0))

    dcr_overall = float(row.get("DCR_Overall", 0.0))
    dcr_enb = float(row.get("DCR_eNB", 0.0))
    dcr_denom = float(row.get("DCR_Denom", 0.0))
    dcr_norm = float(row.get("DCR_Norm", 0.0))
    session_time = float(row.get("pmSessionTimeUe", 0.0))

    hosr = float(row.get("HOSR", 100.0))
    ho_denom = float(row.get("HO_Denom", 0.0))
    ho_prep_sr = float(row.get("HO_Prep_SR", 100.0))
    ho_exe_sr = float(row.get("HO_Exe_SR", 100.0))

    avg_cqi = float(row.get("Avg_CQI", 10.0)) if pd.notna(row.get("Avg_CQI")) else 10.0
    low_cqi = float(row.get("Low_CQI_Ratio", 0.0)) if pd.notna(row.get("Low_CQI_Ratio")) else 0.0
    cqi_samples = float(row.get("CQI_Samples", 0.0))
    subband_delta = float(row.get("Subband_CQI_Delta", 0.0)) if pd.notna(row.get("Subband_CQI_Delta")) else 0.0

    ul_intf = float(row.get("UL_Intf_dBm", -118.0)) if pd.notna(row.get("UL_Intf_dBm")) else -118.0
    rank1_pct = float(row.get("Rank1_Ratio", 40.0)) if pd.notna(row.get("Rank1_Ratio")) else 40.0
    prb_crest = float(row.get("PRB_Crest_dB", 0.0)) if pd.notna(row.get("PRB_Crest_dB")) else 0.0
    prb_max_dbm = float(row.get("PRB_Max_dBm", ul_intf)) if pd.notna(row.get("PRB_Max_dBm")) else ul_intf
    peak_prbs_txt = str(row.get("Peak_PRBs_Text", "N/A"))

    irat_pp = float(row.get("IRAT_PingPong_Rate", 0.0))
    irat_denom = float(row.get("IRAT_Denom", 0.0))

    volte_bearer_sr = float(row.get("VoLTE_Bearer_SR", 100.0))
    volte_cssr = float(row.get("VoLTE_CSSR", 100.0))
    volte_cssr_denom = float(row.get("VoLTE_CSSR_Denom", 0.0))
    volte_qci5_cssr = float(row.get("CSSR_Default_Bearer_QCI5", cssr))
    volte_dcr = float(row.get("VoLTE_DCR", 0.0))
    volte_dcr_denom = float(row.get("VoLTE_DCR_Denom", 0.0))
    volte_dcr_norm = float(row.get("VoLTE_DCR_Norm", 0.0))

    # Step 0: Statistical Significance & Minimum Sample Thresholds (Section 2.8)
    valid_cssr = cssr_denom >= KPI_BASELINES["CSSR"]["min_denom"]
    valid_dcr = dcr_denom >= KPI_BASELINES["DCR"]["min_denom"]
    valid_dcr_norm = session_time >= KPI_BASELINES["DCR_NORM"]["min_denom"]
    valid_ho = ho_denom >= KPI_BASELINES["HOSR"]["min_denom"]
    valid_cqi = cqi_samples >= KPI_BASELINES["CQI_AVG"]["min_denom"]
    valid_irat = has_irat and (irat_denom >= KPI_BASELINES["IRAT_PP"]["min_denom"])
    valid_vcssr = has_volte and (volte_cssr_denom >= KPI_BASELINES["VOLTE_CSSR"]["min_denom"])
    valid_vdcr = has_volte and (volte_dcr_denom >= KPI_BASELINES["VOLTE_DCR"]["min_denom"])

    findings: List[Dict[str, Any]] = []
    low_sample_notes: List[str] = []

    if not valid_dcr and dcr_overall > 1.0:
        low_sample_notes.append(
            f"Cell {cell} excluded from degradation list due to insufficient sample size "
            f"(Sample = {int(dcr_denom)} < 30 releases)."
        )
    if not valid_cssr and cssr < 97.0:
        low_sample_notes.append(
            f"Cell {cell} excluded from degradation list due to insufficient sample size "
            f"(Sample = {int(cssr_denom)} < 30 attempts)."
        )
    if not valid_ho and hosr < 95.0 and ho_denom > 0:
        low_sample_notes.append(
            f"Cell {cell} excluded from degradation list due to insufficient sample size "
            f"(Sample = {int(ho_denom)} < 20 HO attempts)."
        )

    # ------------------------------------------------------------------
    # TREE 4.1: BAD CSSR (Accessibility < 97.0% + Standalone RACH OOC > 5%)
    # ------------------------------------------------------------------
    ra_att = float(row.get("pmRaAttCbra", 0.0))
    ra_ooc = float(row.get("pmRaFailCbraMsg1DiscOoc", 0.0))
    ra_sched = float(row.get("pmRaFailCbraMsg1DiscSched", 0.0))
    ooc_ratio = (ra_ooc / ra_att * 100.0) if ra_att > 0 else 0.0

    if valid_cssr and cssr < 97.0:
        sev = 3 if cssr < 95.0 else 2
        rrc_lic = float(row.get("pmRrcConnEstabFailLic", 0.0))
        rrc_mp = float(row.get("pmRrcConnEstabFailMpOverload", 0.0))
        rrc_du = float(row.get("pmRrcConnEstabFailDuIntens", 0.0))
        mme_ovl = (
            float(row.get("pmS1SigConnEstabFailMmeOvlMos", 0.0))
            + float(row.get("pmRrcConnEstabFailMmeOvlMos", 0.0))
            + float(row.get("pmRrcConnEstabFailMmeOvlMod", 0.0))
        )
        erab_lic = float(row.get("pmErabEstabFailInitLic", 0.0))
        gbr_fail = float(row.get("pmErabEstabFailGbrDlEnb", 0.0)) + float(
            row.get("pmErabEstabFailGbrUlEnb", 0.0)
        )

        if rrc_sr < 97.0:
            if rrc_lic > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "License Congestion",
                    "cause": "License Limitation - Connected Users License Exceeded",
                    "confidence": 96,
                    "recommendation": "Upgrade Connected Users License or activate Dynamic Load Control (DLC).",
                    "scenario": "B_LIC",
                })
            elif rrc_mp > 0 or rrc_du > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "Hardware / Baseband / RRU-RET",
                    "cause": "Baseband / MP CPU Hardware Overload",
                    "confidence": 93,
                    "recommendation": "Offload traffic via IFLB, expand Baseband hardware, or adjust cell max user limit.",
                    "scenario": "B_HW",
                })
            elif ooc_ratio > 5.0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "RACH Overshooting (OOC)",
                    "cause": f"Overshooting Cell / Timing Advance Out of Range ({ooc_ratio:.1f}% OOC)",
                    "confidence": 94,
                    "recommendation": "Increase mechanical/electrical tilt, decrease pMax, or expand EUtranCellFDD.cellRange.",
                    "scenario": "B_OOC",
                })
            elif ul_intf > -110.0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "UL Interference / PIM",
                    "cause": f"Uplink Interference ({ul_intf:.1f} dBm) Blocking RACH Msg3 Detection",
                    "confidence": 92,
                    "recommendation": "Perform PIM analysis, inspect external interference sources, or adjust PRACH power boosting.",
                    "scenario": "A_PIM",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "PRB / CCE / GBR Congestion",
                    "cause": "Downlink Control Channel (PDCCH/CCE) Congestion or Weak CCE Coverage",
                    "confidence": 87,
                    "recommendation": "Optimize PDCCH power boost, enable Dynamic CCE allocation, or check CCE aggregation levels.",
                    "scenario": "B_CCE",
                })
        elif s1_sr < 98.0:
            if mme_ovl > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "Transport / Core (S1/MME)",
                    "cause": "Core Network MME Signaling Overload",
                    "confidence": 94,
                    "recommendation": "Engage Core Network team to rebalance MME Pool capacity.",
                    "scenario": "CORE_MME",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "Transport / Core (S1/MME)",
                    "cause": "S1 Transport Bounded Latency / SCTP Transmission Drops",
                    "confidence": 88,
                    "recommendation": "Inspect S1 IP transport link quality, DSCP mapping, and router interface drops.",
                    "scenario": "TN_S1",
                })
        else:
            if erab_lic > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "License Congestion",
                    "cause": "License Limitation - E-RAB / RLC UM License Shortage",
                    "confidence": 95,
                    "recommendation": "Install additional E-RAB / Radio Bearer license capacity.",
                    "scenario": "B_LIC",
                })
            elif gbr_fail > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "PRB / CCE / GBR Congestion",
                    "cause": "GBR Resource Admission Failure (DL/UL Bandwidth Congestion)",
                    "confidence": 92,
                    "recommendation": "Adjust GBR CAC parameters, audit VoLTE ARP priority thresholds.",
                    "scenario": "E_VOLTE",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "Transport / Core (S1/MME)",
                    "cause": "Radio Bearer Allocation Failure during Initial Context Setup",
                    "confidence": 85,
                    "recommendation": "Audit QCI profile configuration and verify S1-U GTP-U bearer binding.",
                    "scenario": "TN_S1",
                })
    elif valid_cssr and ra_att >= 500 and ooc_ratio > 5.0:
        findings.append({
            "severity": 2,
            "domain": "RACH",
            "category": "RACH Overshooting (OOC)",
            "cause": f"Overshooting Cell / Timing Advance Out of Range ({ooc_ratio:.1f}% Msg1 OOC Discards)",
            "confidence": 92,
            "recommendation": "Increase RET electrical down-tilt, decrease pMax, or align EUtranCellFDD.cellRange.",
            "scenario": "B_OOC",
        })

    # ------------------------------------------------------------------
    # TREE 4.2: HIGH DCR (3GPP Overall DCR > 1.0% OR Normalized > 10.0)
    # ------------------------------------------------------------------
    abn_enb = float(row.get("pmErabRelAbnormalEnbAct", 0.0))
    abn_mme = float(row.get("pmErabRelAbnormalMmeAct", 0.0))
    tot_act_drops = abn_enb + abn_mme

    abn_ue_lost = float(row.get("pmErabRelAbnormalEnbActUeLost", 0.0))
    abn_ho = float(row.get("pmErabRelAbnormalEnbActHo", 0.0))
    abn_tn = float(row.get("pmErabRelAbnormalEnbActTnFail", 0.0))
    abn_pe = float(row.get("pmErabRelAbnormalEnbActPe", 0.0))

    dcr_triggered = valid_dcr and (
        dcr_overall > 1.0 or (valid_dcr_norm and dcr_norm > 10.0)
    )
    if dcr_triggered:
        sev = (
            3
            if dcr_overall > 2.5 or (valid_dcr_norm and dcr_norm > 25.0)
            else 2
        )
        # v6.0.0 Tree 4.2: all breakdown shares divide by (pmErabRelAbnormalEnbAct + pmErabRelAbnormalMmeAct)
        ue_lost_share = (abn_ue_lost / tot_act_drops) if tot_act_drops > 0 else 0.0
        ho_drop_share = (abn_ho / tot_act_drops) if tot_act_drops > 0 else 0.0
        tn_drop_share = (abn_tn / tot_act_drops) if tot_act_drops > 0 else 0.0
        mme_drop_share = (abn_mme / tot_act_drops) if tot_act_drops > 0 else 0.0
        pe_drop_share = (abn_pe / tot_act_drops) if tot_act_drops > 0 else 0.0

        if ue_lost_share > 0.50:
            if ul_intf > -108.0 or (ul_intf > -112.0 and prb_crest >= 8.0 and prb_max_dbm > -104.0):
                if prb_crest >= 6.0:
                    findings.append({
                        "severity": sev,
                        "domain": "Overall DCR",
                        "category": "UL Interference / PIM",
                        "cause": f"Passive Intermodulation (PIM) or Narrowband External Interference (Crest +{prb_crest:.1f} dB)",
                        "confidence": 96,
                        "recommendation": "Schedule PIM sweep testing on antenna jumpers and TMA units; check frequency co-siting.",
                        "scenario": "A_PIM",
                    })
                else:
                    findings.append({
                        "severity": sev,
                        "domain": "Overall DCR",
                        "category": "UL Interference / PIM",
                        "cause": f"Broadband Uplink External Interference ({ul_intf:.1f} dBm/PRB)",
                        "confidence": 95,
                        "recommendation": "Deploy Anritsu/spectrum analyzer for external jammer/repeater detection.",
                        "scenario": "A_PIM",
                    })
            elif avg_cqi < 7.0 and low_cqi > 40.0:
                findings.append({
                    "severity": sev,
                    "domain": "Overall DCR",
                    "category": "RF Weak Coverage / UL Power",
                    "cause": f"Severe Downlink Weak Coverage / Coverage Hole (Avg CQI {avg_cqi:.1f})",
                    "confidence": 93,
                    "recommendation": "Reduce down-tilt, increase RS Tx Power (crsGain), or deploy small cell/repeater.",
                    "scenario": "A_COV_RET",
                })
            else:
                conf = min(94, int(84 + min(10, abn_enb / 30.0)))
                findings.append({
                    "severity": sev,
                    "domain": "Overall DCR",
                    "category": "RF Weak Coverage / UL Power",
                    "cause": f"Uplink Power Control Bottleneck / UE Transmit Power Limited (UE Lost {ue_lost_share*100:.0f}%)",
                    "confidence": conf,
                    "recommendation": "Optimize alpha and p0NominalPusch parameters in EUtranCellFDD.",
                    "scenario": "C_ULPC",
                })
        elif ho_drop_share > 0.30:
            if hosr < 90.0:
                findings.append({
                    "severity": sev,
                    "domain": "Overall DCR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": "Handover Execution Failure / Missing Neighbor Relations",
                    "confidence": 94,
                    "recommendation": "Audit ANR settings, check X2 interface availability, and review a3Offset/timeToTrigger.",
                    "scenario": "C_HO",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "Overall DCR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": "Late Handover / Ping-Pong Handover in High Mobility Zone",
                    "confidence": 90,
                    "recommendation": "Increase hysteresisA3, decrease timeToTrigger, and enable Speed Dependent Mobility.",
                    "scenario": "C_HO",
                })
        elif tn_drop_share > 0.20:
            findings.append({
                "severity": sev,
                "domain": "Overall DCR",
                "category": "Transport / Core (S1/MME)",
                "cause": "Backhaul Transport Network Instability / S1 Link Flapping",
                "confidence": 95,
                "recommendation": "Check Ethernet port errors (EthernetPort), S1-U packet loss (pmS1uPktLostDlQci), and MW alignment.",
                "scenario": "TN_S1",
            })
        elif mme_drop_share > 0.30:
            findings.append({
                "severity": sev,
                "domain": "Overall DCR",
                "category": "Transport / Core (S1/MME)",
                "cause": f"Core Network MME Initiated Session Drop / EPC Bearer Release ({mme_drop_share*100:.0f}% MME Active Drops)",
                "confidence": 92,
                "recommendation": "Escalate to EPC team to inspect MME S1AP cause codes and SGW GTP-U path supervision timeouts.",
                "scenario": "CORE_MME",
            })
        elif pe_drop_share > 0.15:
            findings.append({
                "severity": sev,
                "domain": "Overall DCR",
                "category": "PRB / CCE / GBR Congestion",
                "cause": "Pre-emption Drop due to High Priority ARP Bearer Admission",
                "confidence": 93,
                "recommendation": "Increase cell PRB bandwidth, activate Multi-Carrier Load Management, or review ARP pre-emption settings.",
                "scenario": "B_CCE",
            })
        else:
            findings.append({
                "severity": sev,
                "domain": "Overall DCR",
                "category": "RF Weak Coverage / UL Power",
                "cause": "Composite E-RAB Abnormal Release (RF Fading & Signaling Timeout)",
                "confidence": 83,
                "recommendation": "Audit t310/n310 RLF timers and cell-edge UL SINR profile.",
                "scenario": "C_ULPC",
            })

    # ------------------------------------------------------------------
    # TREE 4.3: HANDOVER FAILURE (HOSR < 95.0%)
    # ------------------------------------------------------------------
    if valid_ho and hosr < 95.0:
        sev = 3 if hosr < 90.0 else 2
        ho_lic_rej = float(row.get("pmHoPrepRejInLicConnUsers", 0.0))
        ho_ul_rej = float(row.get("pmHoPrepRejOutUlThres", 0.0))

        if ho_prep_sr < 95.0:
            if ho_lic_rej > 0:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "License Congestion",
                    "cause": "Target Cell License Rejection - Connected Users Exceeded",
                    "confidence": 95,
                    "recommendation": "Expand Connected Users License on target eNodeB.",
                    "scenario": "B_LIC",
                })
            elif ho_ul_rej > 0:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "UL Interference / PIM",
                    "cause": "Target Cell Bad Uplink Quality Rejection",
                    "confidence": 91,
                    "recommendation": "Clear target cell UL interference, optimize target cell tilt.",
                    "scenario": "A_PIM",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "PRB / CCE / GBR Congestion",
                    "cause": "Target Cell Capacity / CCE / PRB Congestion",
                    "confidence": 88,
                    "recommendation": "Enable Inter-Frequency Load Balancing (IFLB) to shift traffic before HO.",
                    "scenario": "C_HO",
                })
        else:
            if abn_ho > 50:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": f"Physical Layer Handover Execution Collapse ({hosr:.2f}% HOSR, Severe Cell-Edge Interference)",
                    "confidence": 93,
                    "recommendation": "Adjust a3Offset (e.g. from 3dB to 2dB) for earlier HO triggering; check PCI collision/confusion.",
                    "scenario": "C_HO",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": f"Missing External Cell Definition / ANR Blocked Neighbor ({hosr:.2f}% HOSR)",
                    "confidence": 90,
                    "recommendation": "Check AnrFunction.anrState, force manual relation creation on EUtranCellRelation.",
                    "scenario": "C_HO",
                })

    # ------------------------------------------------------------------
    # TREE 4.4: LOW CQI / RADIO INTERFERENCE DEGRADATION
    # ------------------------------------------------------------------
    if valid_cqi and (avg_cqi < 7.5 or low_cqi > 30.0 or ul_intf > -110.0):
        sev = 3 if (avg_cqi < 6.0 or low_cqi > 50.0 or ul_intf > -105.0) else 2
        if ul_intf > -108.0:
            pim_detail = f", PIM Crest +{prb_crest:.1f} dB" if prb_crest >= 6.0 else ""
            findings.append({
                "severity": sev,
                "domain": "CQI / RF",
                "category": "UL Interference / PIM",
                "cause": f"High Radio Noise Floor / External Interference or PIM ({ul_intf:.1f} dBm/PRB{pim_detail})",
                "confidence": 95,
                "recommendation": "Execute PRB-level interference profile scan, check TMA/RRU hardware status.",
                "scenario": "A_PIM",
            })
        elif ul_intf > -110.0 and avg_cqi >= 7.5 and low_cqi <= 30.0:
            findings.append({
                "severity": 2,
                "domain": "CQI / RF",
                "category": "UL Interference / PIM",
                "cause": f"Elevated Uplink Noise Floor ({ul_intf:.1f} dBm/PRB > -110 dBm Threshold)",
                "confidence": 90,
                "recommendation": "Monitor PRB interference spectrum and inspect RRU/RET jumper shielding.",
                "scenario": "A_PIM",
            })
        elif rank1_pct > 75.0:
            findings.append({
                "severity": sev,
                "domain": "CQI / RF",
                "category": "Hardware / Baseband / RRU-RET",
                "cause": f"Lack of Dominant Server / Severe Pilot Pollution or RRU Branch Imbalance (Rank-1 {rank1_pct:.1f}%)",
                "confidence": 90,
                "recommendation": "Optimize cell dominance by adjusting antenna tilts; reduce overlap between co-channel cells.",
                "scenario": "A_COV_RET",
            })
        elif subband_delta > 4.0:
            findings.append({
                "severity": sev,
                "domain": "CQI / RF",
                "category": "UL Interference / PIM",
                "cause": f"Frequency-Selective Subband Interference (Subband CQI Delta {subband_delta:.1f})",
                "confidence": 89,
                "recommendation": "Shift EARFCN frequency, enable Adaptive Frequency Selective Scheduling (FSS).",
                "scenario": "A_PIM",
            })
        else:
            findings.append({
                "severity": sev,
                "domain": "CQI / RF",
                "category": "RF Weak Coverage / UL Power",
                "cause": f"General Low Signal-to-Noise Ratio / High Low-CQI Tail (Avg CQI {avg_cqi:.2f}, Low-CQI {low_cqi:.1f}%)",
                "confidence": 88,
                "recommendation": "Increase CRS Tx Power, audit feeder/optical SFP attenuation, enable Ericsson Lean Carrier.",
                "scenario": "A_COV_RET",
            })

    # ------------------------------------------------------------------
    # TREE 4.5: 3G/4G IRAT PING-PONG LOOP (v6.0.0)
    # ------------------------------------------------------------------
    rrc_succ = float(row.get("pmRrcConnEstabSucc", 0.0))
    irat_rel_rate = (irat_denom / rrc_succ) if rrc_succ > 0 else 0.0
    if valid_irat and (irat_pp > 5.0 or irat_rel_rate > 0.25):
        sev = 3 if irat_pp > 10.0 else 2
        bad_cov_rep = float(row.get("pmBadCovEvalReport", 0.0))
        csfb_rel = float(row.get("pmUeCtxtRelCsfbWcdma", 0.0))
        sc_wcdma = float(row.get("pmUeCtxtRelSCWcdma", 0.0))
        b2_t1 = float(row["b2Threshold1Rsdp"]) if pd.notna(row.get("b2Threshold1Rsdp")) else -108.0
        b2_t2 = float(row["b2Threshold2RsdpUtran"]) if pd.notna(row.get("b2Threshold2RsdpUtran")) else -98.0
        b2_ecno = float(row["b2Threshold2EcNoUtran"]) if pd.notna(row.get("b2Threshold2EcNoUtran")) else -12.0
        ttt_b2 = float(row["timeToTriggerB2"]) if pd.notna(row.get("timeToTriggerB2")) else 640.0
        hyst_b2 = float(row["hysteresisB2"]) if pd.notna(row.get("hysteresisB2")) else 30.0

        if rrc_succ > 0 and (bad_cov_rep / rrc_succ) > 0.40 and b2_t1 > -110.0:
            cause_txt = "Aggressive Event A2 / B2 Threshold Triggering (4G Threshold Too High)"
            rec_txt = "Lower a2Threshold0Rsdp (to -115 dBm) and b2Threshold1Rsdp (to -116 dBm) so UE stays on 4G longer."
        elif b2_t2 < -100.0 or b2_ecno < -14.0:
            cause_txt = "Loose 3G Target Quality Threshold (B2 Threshold 2 Too Low)"
            rec_txt = "Increase b2Threshold2RsdpUtran (to -95 dBm) or b2Threshold2EcNoUtran (to -11 dB)."
        elif ttt_b2 < 320.0 or hyst_b2 < 20.0:
            cause_txt = "Insufficient Time-to-Trigger / Hysteresis Margin on Event B2"
            rec_txt = "Increase timeToTrigger on ReportConfigB2 to 640 ms and set hysteresis to 30 (3.0 dB)."
        elif sc_wcdma > 0 and csfb_rel > 0.50 * sc_wcdma:
            cause_txt = "CSFB Voice Return Oscillation (Post-Call Fast Return Ping-Pong)"
            rec_txt = "Enable CSFB Fast Return to LTE on RNC/3G side and verify VoLTE QCI=1 stability."
        else:
            cause_txt = f"Inverted Cell Reselection Priorities between SIB6 (4G) and SIB19 (3G) ({irat_pp:.1f}%)"
            rec_txt = "Audit UtranFreqRelation.cellReselectionPriority (4G side) and 3G SIB19 LTE priority."

        findings.append({
            "severity": sev,
            "domain": "IRAT Ping-Pong",
            "category": "Event B2/A2 & Mobility Tuning",
            "cause": cause_txt,
            "confidence": 94,
            "recommendation": rec_txt,
            "scenario": "D_IRAT",
        })

    # ------------------------------------------------------------------
    # TREE 4.6: VoLTE BAD CSSR (v6.0.0)
    # ------------------------------------------------------------------
    if valid_vcssr and (volte_cssr < 97.5 or volte_bearer_sr < 98.0):
        sev = 3 if volte_cssr < 95.0 else 2
        gbr_dl = float(row.get("pmErabEstabFailGbrDlEnb", 0.0))
        gbr_ul = float(row.get("pmErabEstabFailGbrUlEnb", 0.0))
        erab_lic = float(row.get("pmErabEstabFailInitLic", 0.0))

        if gbr_dl > 0 or gbr_ul > 0:
            v_cause = "VoLTE Call Admission Control (CAC) GBR Bandwidth Exhaustion"
            v_cat = "PRB / CCE / GBR Congestion"
            v_rec = "Increase qci1GbrDlLimit and qci1GbrUlLimit in EUtranCellFDD; activate VoLTE GBR CAC optimization."
        elif erab_lic > 0:
            v_cause = "Missing Feature License for VoLTE / GBR Bearers"
            v_cat = "License Congestion"
            v_rec = "Verify license key installation for VoLTE (FAJ 121 0858) and GBR Bearer support."
        elif volte_qci5_cssr < 98.0:
            v_cause = "SIP IMS Signaling Bearer (QCI=5) Setup Failure"
            v_cat = "VoLTE QCI=1 Bearer / IMS"
            v_rec = "Inspect S1-U GTP-U bearer connectivity to P-CSCF; check QCI=5 DSCP mapping (EF/AF41)."
        else:
            v_cause = f"PUCCH SR / CQI Periodicity Resource Congestion for QCI=1 (VoLTE CSSR {volte_cssr:.2f}%)"
            v_cat = "VoLTE QCI=1 Bearer / IMS"
            v_rec = "Increase noOfPucchCqiUsers and optimize commonCqiPeriodicity to 40 ms."

        findings.append({
            "severity": sev,
            "domain": "VoLTE CSSR",
            "category": v_cat,
            "cause": v_cause,
            "confidence": 93,
            "recommendation": v_rec,
            "scenario": "E_VOLTE",
        })

    # ------------------------------------------------------------------
    # TREE 4.7: VoLTE HIGH DCR (v6.0.0)
    # ------------------------------------------------------------------
    if valid_vdcr and (volte_dcr > 0.8 or volte_dcr_norm > 3.0):
        sev = 3 if volte_dcr > 2.0 else 2
        v_abn_enb = float(row.get("pmErabRelAbnormalEnbActQci[1]", 0.0))
        v_abn_mme = float(row.get("pmErabRelAbnormalMmeActQci[1]", 0.0))
        v_tot_abn = v_abn_enb + v_abn_mme
        v_ue_lost = float(row.get("pmErabRelAbnormalEnbActUeLostQci[1]", 0.0))
        v_ho = float(row.get("pmErabRelAbnormalEnbActHoQci[1]", 0.0))
        v_pe = float(row.get("pmErabRelAbnormalEnbActPeQci[1]", 0.0))
        rlc_retx_ul = float(row.get("pmRlcPduRetransUlQci[1]", 0.0))
        rlc_tx_ul = float(row.get("pmRlcPduTxUlQci[1]", 0.0))
        srvcc_att = float(row.get("pmHoPrepAttSrvcc", 0.0))
        srvcc_succ = float(row.get("pmHoExeSuccSrvcc", 0.0))

        if v_pe > 0:
            findings.append({
                "severity": 3,
                "domain": "VoLTE DCR",
                "category": "VoLTE QCI=1 Bearer / IMS",
                "cause": "CRITICAL MISCONFIGURATION: VoLTE QCI=1 Bearer Pre-empted",
                "confidence": 98,
                "recommendation": "Set QciProfilePredefined=qci1 preEmptionCapability = NOT_PREEMPTABLE and preEmptionVulnerability = DISABLED.",
                "scenario": "E_VOLTE",
            })
        elif v_abn_enb > 0 and (v_ue_lost / v_abn_enb) > 0.50:
            rlc_ratio = (rlc_retx_ul / rlc_tx_ul) if rlc_tx_ul > 0 else 0.0
            if rlc_ratio > 0.10 or ul_intf > -108.0:
                findings.append({
                    "severity": sev,
                    "domain": "VoLTE DCR",
                    "category": "UL Interference / PIM",
                    "cause": "Uplink PIM / Radio Noise Floor Destroying VoLTE RTP Packets",
                    "confidence": 95,
                    "recommendation": "Perform PIM sweep on TMA/jumpers; enable ROHC (Robust Header Compression) and TTI Bundling.",
                    "scenario": "E_VOLTE",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "VoLTE DCR",
                    "category": "RF Weak Coverage / UL Power",
                    "cause": "VoLTE Weak Coverage Hole at Cell Edge",
                    "confidence": 91,
                    "recommendation": "Optimize cell tilt; enable Coverage-Triggered SRVCC to 2G/3G before 4G link collapse.",
                    "scenario": "E_VOLTE",
                })
        elif v_abn_enb > 0 and (v_ho / v_abn_enb) > 0.30:
            if srvcc_att > 0 and (srvcc_succ / srvcc_att) < 0.90:
                findings.append({
                    "severity": sev,
                    "domain": "VoLTE DCR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": "SRVCC Handover Execution Collapse to 2G/3G WCDMA",
                    "confidence": 94,
                    "recommendation": "Verify srvccHoAllowed on UtranCellRelation; check MSC Server Sv interface signaling.",
                    "scenario": "E_VOLTE",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "VoLTE DCR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": "Intra/Inter-Frequency LTE VoLTE Handover Execution Failure",
                    "confidence": 90,
                    "recommendation": "Enable cioQci1Enabled on EUtranFreqRelation for QCI=1 specific offset.",
                    "scenario": "E_VOLTE",
                })
        elif v_tot_abn > 0 and (v_abn_mme / v_tot_abn) > 0.40:
            findings.append({
                "severity": sev,
                "domain": "VoLTE DCR",
                "category": "Transport / Core (S1/MME)",
                "cause": "IMS Core Session Supervision Timeout / SIP BYE Abort",
                "confidence": 92,
                "recommendation": "Escalate to IMS team to inspect P-CSCF RTP inactivity timer and SGW GTP-U path failure.",
                "scenario": "CORE_MME",
            })
        else:
            findings.append({
                "severity": sev,
                "domain": "VoLTE DCR",
                "category": "VoLTE QCI=1 Bearer / IMS",
                "cause": f"VoLTE Dedicated QCI=1 Voice Drop Degradation ({volte_dcr:.2f}%)",
                "confidence": 88,
                "recommendation": "Enable ROHC & TTI Bundling on EUtranCellFDD and audit QCI=1 scheduling priority.",
                "scenario": "E_VOLTE",
            })

    # ------------------------------------------------------------------
    # CONSOLIDATE PRIMARY ROOT CAUSE, DIAGNOSTIC EVIDENCE & AMOS SCRIPT
    # ------------------------------------------------------------------
    ue_lost_pct = (abn_ue_lost / tot_act_drops * 100.0) if tot_act_drops > 0 else 0.0
    mme_act_pct = (abn_mme / tot_act_drops * 100.0) if tot_act_drops > 0 else 0.0

    evidence_items = [
        {
            "Counter / Indicator": "pmRrcConnEstabAtt / Succ",
            "Value": f"{int(row.get('pmRrcConnEstabAtt', 0)):,} / {int(row.get('pmRrcConnEstabSucc', 0)):,}",
            "Meaning": f"RRC SR: {rrc_sr:.2f}% | S1 SR: {s1_sr:.2f}% | E-RAB SR: {erab_sr:.2f}%",
        },
        {
            "Counter / Indicator": "pmRaFailCbraMsg1DiscOoc",
            "Value": f"{int(ra_ooc):,} ({ooc_ratio:.2f}%)",
            "Meaning": f"RACH Out-of-Coverage (Sched Disc: {int(ra_sched):,})",
        },
        {
            "Counter / Indicator": "3GPP E-RAB Total Releases",
            "Value": f"{int(dcr_denom):,} rel",
            "Meaning": f"NormEnb: {int(row.get('pmErabRelNormalEnb', 0)):,} | Mme: {int(row.get('pmErabRelMme', 0)):,} | AbnEnb: {int(row.get('pmErabRelAbnormalEnb', 0)):,}",
        },
        {
            "Counter / Indicator": "Active Drops (eNB + MME)",
            "Value": f"{int(abn_enb):,} + {int(abn_mme):,} = {int(tot_act_drops):,}",
            "Meaning": f"Overall DCR: {dcr_overall:.2f}% | eNB DCR: {dcr_enb:.2f}% ({dcr_norm:.1f}/hr)",
        },
        {
            "Counter / Indicator": "pmErabRelAbnormalEnbActUeLost",
            "Value": f"{int(abn_ue_lost):,} ({ue_lost_pct:.1f}% of drops)",
            "Meaning": "RF Link Failure / RLC Unack UE Lost",
        },
        {
            "Counter / Indicator": "pmErabRelAbnormalMmeAct",
            "Value": f"{int(abn_mme):,} ({mme_act_pct:.1f}% of drops)",
            "Meaning": f"Core MME Active Drops | HO Drops: {int(abn_ho):,}",
        },
        {
            "Counter / Indicator": "Handover Prep / Exe SR",
            "Value": f"{ho_prep_sr:.2f}% / {ho_exe_sr:.2f}%",
            "Meaning": f"Composite HOSR: {hosr:.2f}% ({int(ho_denom):,} Prep Att)",
        },
        {
            "Counter / Indicator": "pmRadioRecInterferencePwr",
            "Value": f"{ul_intf:.2f} dBm/PRB",
            "Meaning": "UL Noise Floor (> -110 dBm = Degraded, > -108 = Severe)",
        },
        {
            "Counter / Indicator": "PRB Interference Crest",
            "Value": f"+{prb_crest:.1f} dB ({peak_prbs_txt})",
            "Meaning": "PIM / Narrowband Spectral Imbalance",
        },
        {
            "Counter / Indicator": "CQI & MIMO Rank-1 Profile",
            "Value": f"CQI {avg_cqi:.2f} | Low {low_cqi:.1f}% | R1 {rank1_pct:.1f}%",
            "Meaning": "DL Radio SINR & Spatial Multiplexing Rank",
        },
    ]

    if not findings:
        if low_sample_notes or (not valid_cssr and not valid_dcr):
            excl_msg = (
                low_sample_notes[0]
                if low_sample_notes
                else f"Cell {cell} excluded from degradation list due to insufficient sample size."
            )
            return {
                "Status": "LOW_SAMPLE",
                "Severity_Rank": 0,
                "RCA_Category": "Low Traffic Sample (Excluded)",
                "Primary_Root_Cause": excl_msg,
                "Secondary_Causes": "None (Statistical Sample Gate Filtered)",
                "Affected_Domain": "LOW_TRAFFIC_SAMPLE",
                "Confidence_Score": 0,
                "Recommendation": "No action required — monitor across wider ROP aggregation window.",
                "AMOS_Script": f"# Cell {cell} ({enb}) skipped: LOW_TRAFFIC_SAMPLE (Rule of Minimum Denominator).",
                "Diagnostic_Evidence": evidence_items,
            }

        return {
            "Status": "NORMAL",
            "Severity_Rank": 1,
            "RCA_Category": "Normal / Healthy",
            "Primary_Root_Cause": "All KPIs within 3GPP / Ericsson Target Baselines",
            "Secondary_Causes": "None",
            "Affected_Domain": "HEALTHY",
            "Confidence_Score": 98,
            "Recommendation": "Maintain baseline configuration; routine PM monitoring.",
            "AMOS_Script": f"# No remediation required for healthy cell {cell} on {enb}\nst EUtranCellFDD={cell}",
            "Diagnostic_Evidence": evidence_items,
        }

    findings.sort(key=lambda x: (x["severity"], x["confidence"]), reverse=True)
    primary = findings[0]
    secondary_list = [f"{f['domain']}: {f['cause']}" for f in findings[1:]]
    status_str = "CRITICAL" if primary["severity"] >= 3 else "DEGRADED"

    scenarios = list(dict.fromkeys(f["scenario"] for f in findings))
    amos_script = generate_amos_script(enb, cell, scenarios, primary["cause"])

    return {
        "Status": status_str,
        "Severity_Rank": primary["severity"] + 1,  # 4=CRITICAL, 3=DEGRADED
        "RCA_Category": primary["category"],
        "Primary_Root_Cause": primary["cause"],
        "Secondary_Causes": " | ".join(secondary_list) if secondary_list else "None",
        "Affected_Domain": primary["domain"],
        "Confidence_Score": int(primary["confidence"]),
        "Recommendation": primary["recommendation"],
        "AMOS_Script": amos_script,
        "Diagnostic_Evidence": evidence_items,
    }


def generate_amos_script(
    enodeb: str, cell: str, scenarios: List[str], root_cause: str
) -> str:
    """Build copy-pasteable Ericsson AMOS / Moshell CLI script per Section 5.1 - 5.5."""
    lines = [
        "# ===================================================================",
        "# ERICSSON AMOS / MOSHELL REMEDIATION PROTOCOL (v6.0)",
        f"# Target Node : {enodeb}   |   Cell : {cell}",
        f"# Diagnosis   : {root_cause}",
        "# ===================================================================",
        f"moshell {enodeb}",
        "lt all",
        "confb+",
        "gs+",
        "",
        "# --- Step 1: Health, Alarm & Operational State Verification ---",
        f"st EUtranCellFDD={cell}",
        "alt",
    ]

    for sc in scenarios:
        if sc == "A_PIM":
            lines.extend([
                "",
                "# --- Scenario A: Low CQI / UL Interference & PIM Diagnostic Protocol ---",
                f"pget EUtranCellFDD={cell} pmRadioRecInterferencePwr$",
                f"pget EUtranCellFDD={cell} pmRadioRecInterferencePwrPrb",
                f"pget EUtranCellFDD={cell} pmRadioUeRepCqiDistr|pmRadioUeRepRankDistr",
                "cabx",
                "lhsh BXP_0 vsi",
                "set ENodeBFunction=1,EricssonLeanCarrier=1 leanCarrierEnabled true",
            ])
        elif sc == "A_COV_RET":
            lines.extend([
                "",
                "# --- Scenario A: Hardware RRU/RET Tilt & DL Coverage Protocol ---",
                "get RetSubUnit electricalAntennaTilt|iuantAntennaModelNumber|operationalState",
                "hget RfPort vswrSupervisionSensitivity|vswrValue",
                f"pget EUtranCellFDD={cell} pmRadioUeRepCqiDistr|pmRadioUeRepRankDistr",
                f"get EUtranCellFDD={cell} crsGain|pdschTypeBGain",
                "set ENodeBFunction=1,EricssonLeanCarrier=1 leanCarrierEnabled true",
            ])
        elif sc == "B_OOC":
            lines.extend([
                "",
                "# --- Scenario B: Bad CSSR / RACH Out-of-Coverage Protocol ---",
                f"pget EUtranCellFDD={cell} pmRaAttCbra|pmRaFailCbraMsg1DiscOoc",
                f"get EUtranCellFDD={cell} cellRange|rachRootSequence",
                "get RetSubUnit electricalAntennaTilt",
                f"set EUtranCellFDD={cell} cellRange 15",
            ])
        elif sc == "B_LIC":
            lines.extend([
                "",
                "# --- Scenario B: Connected Users / E-RAB License Congestion Protocol ---",
                f"pget EUtranCellFDD={cell} pmRrcConnEstabFailLic|pmErabEstabFailInitLic",
                "invl",
                "get CapacityState grantedCapacityLevel|operationalCapacityLevel",
                "set LoadBalancingFunction=1 lbCeiling 80",
            ])
        elif sc == "B_HW":
            lines.extend([
                "",
                "# --- Scenario B: Baseband MP/DU CPU Overload Mitigation ---",
                "cabx",
                f"pget EUtranCellFDD={cell} pmRrcConnEstabFailMpOverload|pmRrcConnEstabFailDuIntens",
                f"set EUtranCellFDD={cell} pdcchCceOptimizationEnabled true",
            ])
        elif sc == "B_CCE":
            lines.extend([
                "",
                "# --- Scenario B: PDCCH / CCE Congestion Optimization ---",
                f"pget EUtranCellFDD={cell} pmRaFailCbraMsg1DiscSched|pmRaFailCbraMsg2Disc",
                f"set EUtranCellFDD={cell} pdcchCceOptimizationEnabled true",
                f"set EUtranCellFDD={cell} pdcchPowerBoostMax 3",
            ])
        elif sc == "C_ULPC":
            lines.extend([
                "",
                "# --- Scenario C: High DCR & Uplink Power Control Optimization ---",
                f"pget EUtranCellFDD={cell} pmErabRelAbnormalEnbAct|pmErabRelAbnormalEnbActUeLost",
                f"get EUtranCellFDD={cell} p0NominalPusch|alpha|t310|n310",
                f"set EUtranCellFDD={cell} p0NominalPusch -100",
                f"set EUtranCellFDD={cell} alpha 8",
                f"set EUtranCellFDD={cell} ulSrsEnable true",
            ])
        elif sc == "C_HO":
            lines.extend([
                "",
                "# --- Scenario C: Handover Execution Failure & ANR Neighbor Protocol ---",
                f"pget EUtranCellFDD={cell} pmErabRelAbnormalEnbActHo|pmCellHoExeAttLteIntraF",
                f"get EUtranCellFDD={cell},EUtranCellRelation= hoAllowed|isHoAllowed|cellIndividualOffsetEUtran",
                "get AnrFunction anrState",
                f"set EUtranCellFDD={cell},EUtranCellRelation= hoAllowed true",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigEUtraBestCell=1 a3Offset 20",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigEUtraBestCell=1 timeToTrigger 80",
            ])
        elif sc == "D_IRAT":
            lines.extend([
                "",
                "# --- Scenario D: 3G/4G IRAT Ping-Pong Mitigation Protocol ---",
                f"get EUtranCellFDD={cell},UeMeasControl=1 a2Threshold0Rsdp",
                f"get EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2=1 b2Threshold1Rsdp|b2Threshold2RsdpUtran|timeToTrigger|hysteresis",
                f"get EUtranCellFDD={cell},UtranFreqRelation= cellReselectionPriority|threshXLow|qRxLevMin",
                f"set EUtranCellFDD={cell},UeMeasControl=1 a2Threshold0Rsdp -115",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2=1 b2Threshold1Rsdp -116",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2=1 b2Threshold2RsdpUtran -95",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2=1 timeToTrigger 640",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2=1 hysteresis 30",
                f"set EUtranCellFDD={cell},UtranFreqRelation= cellReselectionPriority 3",
            ])
        elif sc == "E_VOLTE":
            lines.extend([
                "",
                "# --- Scenario E: VoLTE QCI=1 Optimization Protocol ---",
                "get QciProfilePredefined=qci1 preEmption|schedulingPolicy",
                "set QciProfilePredefined=qci1 preEmptionCapability NOT_PREEMPTABLE",
                "set QciProfilePredefined=qci1 preEmptionVulnerability DISABLED",
                "set QciProfilePredefined=qci1 schedulingPolicy PRIORITY_BASED",
                f"set EUtranCellFDD={cell} qci1GbrDlLimit 20",
                f"set EUtranCellFDD={cell} qci1GbrUlLimit 20",
                f"set EUtranCellFDD={cell} rohcEnabled true",
                f"set EUtranCellFDD={cell} ttiBundlingEnabled true",
                f"set EUtranCellFDD={cell},UtranCellRelation= srvccHoAllowed true",
            ])
        elif sc in ("TN_S1", "CORE_MME"):
            lines.extend([
                "",
                "# --- S1 Transport & EPC MME Core Supervision Check ---",
                "st TermPointToMme",
                "get Sctp errorCounter|assocMaxRtx",
                f"pget EUtranCellFDD={cell} pmErabRelAbnormalMmeAct|pmErabRelAbnormalEnbActTnFail",
            ])

    lines.extend([
        "",
        "confb-",
        "gs-",
        "cvms Post_Antigravity_RCA_v6_Optimization",
    ])
    return "\n".join(lines)


# ============================================================================
# 5. END-TO-END ORCHESTRATOR (INGEST -> MERGE -> KPI -> RCA)
# ============================================================================

def analyze_counter_dump(
    file_source: Union[str, Path, BinaryIO, bytes],
    filename: Optional[str] = None,
) -> AnalysisResult:
    """
    Ingest CSV or Excel (.xlsx, .xls) Ericsson PM counter dump, execute
    vectorized 3GPP KPI aggregations, time-series synthesis, and v6.0.0 RCA trees.
    """
    ingestion_ts = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

    if isinstance(file_source, (str, Path)):
        p = Path(file_source)
        source_name = filename or p.name
        suffix = p.suffix.lower()
        raw_bytes = p.read_bytes()
    elif isinstance(file_source, bytes):
        raw_bytes = file_source
        source_name = filename or "Uploaded_Counters.xlsx"
        suffix = Path(source_name).suffix.lower()
    else:
        file_source.seek(0)
        raw_bytes = file_source.read()
        source_name = filename or getattr(file_source, "name", "Uploaded_Counters.xlsx")
        suffix = Path(source_name).suffix.lower()

    sheets_loaded: List[str] = []
    kpi_frames: List[pd.DataFrame] = []
    cqi_long_frames: List[pd.DataFrame] = []
    qci_long_frames: List[pd.DataFrame] = []
    prb_frames: List[pd.DataFrame] = []

    if suffix == ".csv":
        try:
            text_data = raw_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            text_data = raw_bytes.decode("latin1")

        preview_lines = text_data.splitlines()[:25]
        header_line_idx = 0
        best_line_score = -1
        for idx, row_tokens in enumerate(csv.reader(preview_lines)):
            cleaned = [t.strip().lower() for t in row_tokens if t and t.strip()]
            if not cleaned:
                continue
            score = sum(
                2 if (tok in HEADER_KEYWORDS or tok.startswith("pm")) else (
                    1 if any(k in tok for k in ("cell", "erbs", "enodeb", "date", "time")) else 0
                )
                for tok in cleaned
            )
            if score > best_line_score:
                best_line_score = score
                header_line_idx = idx

        raw_df = pd.read_csv(
            io.StringIO(text_data),
            skiprows=header_line_idx,
            header=None,
            low_memory=False,
            on_bad_lines="skip",
        )
        norm_df = normalize_columns(_detect_and_promote_header(raw_df))
        sheets_loaded.append("CSV_Dump")
        kind = _classify_sheet(norm_df)
        if kind == "cqi_vector_long":
            cqi_long_frames.append(norm_df)
        elif kind == "qci_vector_long":
            qci_long_frames.append(norm_df)
        elif kind == "prb_spectrum":
            prb_frames.append(norm_df)
        else:
            kpi_frames.append(norm_df)
    else:
        engine = "xlrd" if suffix == ".xls" else "openpyxl"
        xl = pd.ExcelFile(io.BytesIO(raw_bytes), engine=engine)
        for s_name in xl.sheet_names:
            raw_df = xl.parse(sheet_name=s_name, header=None)
            if raw_df.empty:
                continue
            norm_df = normalize_columns(_detect_and_promote_header(raw_df))
            if norm_df.empty:
                continue
            kind = _classify_sheet(norm_df)
            if kind == "kpi_main":
                kpi_frames.append(norm_df)
                sheets_loaded.append(s_name)
            elif kind == "cqi_vector_long":
                cqi_long_frames.append(norm_df)
                sheets_loaded.append(s_name)
            elif kind == "qci_vector_long":
                qci_long_frames.append(norm_df)
                sheets_loaded.append(s_name)
            elif kind == "prb_spectrum":
                prb_frames.append(norm_df)
                sheets_loaded.append(s_name)

    if qci_long_frames:
        piv_qci = _pivot_qci_long_sheet(pd.concat(qci_long_frames, ignore_index=True))
        if not piv_qci.empty:
            kpi_frames.append(piv_qci)

    if not kpi_frames and not cqi_long_frames:
        raise ValueError(
            "Không tìm thấy bảng counter Ericsson hợp lệ trong tệp tải lên. "
            "Vui lòng kiểm tra định dạng cột (EUtranCellFDD / ERBS Id / pm*)."
        )

    # Combine or outer-merge main KPI tables on (eNodeB_Name, EUtranCellFDD, ROP_Time)
    if len(kpi_frames) == 1:
        df_main = kpi_frames[0].copy()
    elif len(kpi_frames) > 1:
        # Prepare each frame with ROP_Time so multi-tab scalar sheets merge cleanly without duplicating rows
        prepared_frames: List[pd.DataFrame] = []
        for frm in kpi_frames:
            f_cp = frm.copy()
            if "EUtranCellFDD" not in f_cp.columns:
                continue
            f_cp["EUtranCellFDD"] = f_cp["EUtranCellFDD"].astype(str).str.strip()
            if "eNodeB_Name" not in f_cp.columns:
                f_cp["eNodeB_Name"] = f_cp["EUtranCellFDD"].str.replace(r"\d{2}(-[A-Z0-9]+)?$", r"\1", regex=True)
            f_cp["eNodeB_Name"] = f_cp["eNodeB_Name"].astype(str).str.strip()
            f_cp["ROP_Time"] = _build_rop_timestamp(f_cp)
            prepared_frames.append(f_cp)

        # Distinguish separate PM domain sheets (Jaccard < 0.70) vs same-schema time/region tabs (>= 0.70)
        col_sets = [set(f.columns) & set(SCALAR_COUNTERS) for f in prepared_frames]
        union_01 = col_sets[0].union(col_sets[1]) if len(col_sets) >= 2 else set()
        jaccard_01 = (len(col_sets[0].intersection(col_sets[1])) / len(union_01)) if union_01 else 1.0

        if len(prepared_frames) >= 2 and jaccard_01 < 0.70:
            # Aggregate first frame to (EUtranCellFDD, ROP_Time) first
            base_f = prepared_frames[0]
            base_nums = [c for c in base_f.columns if c in SCALAR_COUNTERS]
            for c in base_nums:
                base_f[c] = pd.to_numeric(base_f[c], errors="coerce").fillna(0.0)
            base_agg = {c: ("mean" if c in PARAMETER_COLS else "sum") for c in base_nums}
            if base_agg:
                df_main = base_f.groupby(["eNodeB_Name", "EUtranCellFDD", "ROP_Time"], as_index=False).agg(base_agg)
            else:
                df_main = base_f.drop_duplicates(subset=["EUtranCellFDD", "ROP_Time"]).copy()
            if "earfcndl" in base_f.columns:
                ear_map = base_f.groupby("EUtranCellFDD")["earfcndl"].first()
                df_main["earfcndl"] = df_main["EUtranCellFDD"].map(ear_map)

            for extra_f in prepared_frames[1:]:
                already_have = set(df_main.columns) & set(SCALAR_COUNTERS)
                num_extra = [c for c in extra_f.columns if c in SCALAR_COUNTERS and c not in already_have]
                if not num_extra:
                    continue
                for c in num_extra:
                    extra_f[c] = pd.to_numeric(extra_f[c], errors="coerce").fillna(0.0)
                extra_agg = {c: ("mean" if c in PARAMETER_COLS else "sum") for c in num_extra}
                grp_extra = extra_f.groupby(["EUtranCellFDD", "ROP_Time"], as_index=False).agg(extra_agg)
                df_main = df_main.merge(grp_extra, on=["EUtranCellFDD", "ROP_Time"], how="outer")
                if "eNodeB_Name" in extra_f.columns:
                    enb_lookup = extra_f.groupby("EUtranCellFDD")["eNodeB_Name"].first()
                    df_main["eNodeB_Name"] = df_main["eNodeB_Name"].fillna(df_main["EUtranCellFDD"].map(enb_lookup))
        else:
            df_main = pd.concat(prepared_frames, ignore_index=True)
    else:
        df_main = (
            cqi_long_frames[0][
                [c for c in ("eNodeB_Name", "EUtranCellFDD", "Date", "Hour", "ROP_Time") if c in cqi_long_frames[0].columns]
            ]
            .drop_duplicates()
            .copy()
        )

    if "EUtranCellFDD" not in df_main.columns:
        raise ValueError("Thiếu cột định danh Cell (`EUtranCellFDD` / `EUtranCell Id` / `Cell_Name`).")
    if "eNodeB_Name" not in df_main.columns:
        df_main["eNodeB_Name"] = df_main["EUtranCellFDD"].astype(str).str.replace(r"\d{2}(-[A-Z0-9]+)?$", r"\1", regex=True)

    df_main["eNodeB_Name"] = df_main["eNodeB_Name"].astype(str).str.strip()
    df_main["EUtranCellFDD"] = df_main["EUtranCellFDD"].astype(str).str.strip()
    df_main["ROP_Time"] = _build_rop_timestamp(df_main)

    total_rop_rows = len(df_main)
    valid_times = df_main["ROP_Time"].dropna()
    data_start = valid_times.min().strftime("%Y-%m-%d %H:%M") if not valid_times.empty else "N/A"
    data_end = valid_times.max().strftime("%Y-%m-%d %H:%M") if not valid_times.empty else "N/A"

    existing_cols = set(df_main.columns)
    has_irat = bool(
        existing_cols.intersection(
            {"pmUeCtxtRelSCWcdma", "UtranCellRelation.pmHoExeSucc", "pmRapid3gTo4gReentries", "pmB2BestCellEvalReport"}
        )
    )
    has_volte = bool(
        existing_cols.intersection(
            {
                "pmErabEstabAttAddedQci[1]",
                "pmErabEstabAttQci[1]",
                "pmErabRelAbnormalEnbActQci[1]",
                "pmErabRelAbnormalEnbQci[1]",
            }
        )
    )

    present_counters = [c for c in SCALAR_COUNTERS if c in df_main.columns]
    for c in present_counters:
        df_main[c] = pd.to_numeric(df_main[c], errors="coerce").fillna(0.0)

    has_3gpp_dcr_denom = any(
        c in present_counters and df_main[c].sum() > 0
        for c in ("pmErabRelAbnormalEnb", "pmErabRelNormalEnb", "pmErabRelMme")
    )

    # If PRB spectrum columns were exported inside the main wide sheet, queue df_main for PRB processing too
    if not prb_frames and len([c for c in df_main.columns if c.startswith("pmRadioRecInterferencePwrPrb")]) >= 10:
        prb_frames.append(df_main)

    wide_cqi_cols = [f"pmRadioUeRepCqiDistr[{i}]" for i in range(16)]
    has_wide_cqi = all(c in df_main.columns for c in wide_cqi_cols)
    if has_wide_cqi:
        for c in wide_cqi_cols:
            df_main[c] = pd.to_numeric(df_main[c], errors="coerce").fillna(0.0)

    wide_rank_cols = [f"pmRadioUeRepRankDistr[{i}]" for i in range(4)]
    has_wide_rank = all(c in df_main.columns for c in wide_rank_cols)
    if has_wide_rank:
        for c in wide_rank_cols:
            df_main[c] = pd.to_numeric(df_main[c], errors="coerce").fillna(0.0)

    wide_intf_cols = [f"pmRadioRecInterferencePwr[{i}]" for i in range(16)]
    has_wide_intf = all(c in df_main.columns for c in wide_intf_cols)
    if has_wide_intf:
        for c in wide_intf_cols:
            df_main[c] = pd.to_numeric(df_main[c], errors="coerce").fillna(0.0)

    has_scalar_intf = (
        "pmRadioRecInterferencePwr" in df_main.columns
        and not cqi_long_frames
        and not has_wide_intf
    )
    if has_scalar_intf:
        df_main["pmRadioRecInterferencePwr"] = pd.to_numeric(
            df_main["pmRadioRecInterferencePwr"], errors="coerce"
        )

    # --- Aggregate scalar counters per Cell ---
    agg_dict: Dict[str, Any] = {c: "sum" for c in present_counters}
    for param_col in PARAMETER_COLS:
        if param_col in agg_dict:
            agg_dict[param_col] = "mean"
    if has_wide_cqi:
        for c in wide_cqi_cols:
            agg_dict[c] = "sum"
    if has_wide_rank:
        for c in wide_rank_cols:
            agg_dict[c] = "sum"
    if has_wide_intf:
        for c in wide_intf_cols:
            agg_dict[c] = "sum"
    if has_scalar_intf:
        agg_dict["pmRadioRecInterferencePwr"] = "mean"

    meta_cols = ["eNodeB_Name", "EUtranCellFDD"]
    cell_df = (
        df_main.groupby(meta_cols, observed=False, as_index=False).agg(agg_dict)
        if agg_dict
        else df_main[meta_cols].drop_duplicates().reset_index(drop=True)
    )

    if "earfcndl" in df_main.columns:
        earfcn_map = df_main.groupby("EUtranCellFDD")["earfcndl"].first()
        cell_df["earfcndl"] = cell_df["EUtranCellFDD"].map(earfcn_map)

    missing_scalars = {
        c: (np.nan if c in PARAMETER_COLS else 0.0)
        for c in SCALAR_COUNTERS
        if c not in cell_df.columns
    }
    if missing_scalars:
        cell_df = pd.concat(
            [cell_df, pd.DataFrame(missing_scalars, index=cell_df.index)], axis=1
        )

    # ========================================================================
    # COMPUTE COMPOSITE 3GPP KPIs PER CELL (SKILL v6.0.0 SECTION 2.1 - 2.8)
    # ========================================================================

    # 2.1 Accessibility (CSSR %)
    rrc_stat_denom = np.maximum(
        cell_df["pmRrcConnEstabAtt"] - cell_df["pmRrcConnEstabAttReatt"], 0.0
    )
    rrc_mme_ovl = (
        cell_df["pmRrcConnEstabFailMmeOvlMos"] + cell_df["pmRrcConnEstabFailMmeOvlMod"]
    )
    rrc_formula_denom = np.maximum(rrc_stat_denom - rrc_mme_ovl, 0.0)
    rrc_sr = np.where(
        rrc_formula_denom > 0,
        np.clip(cell_df["pmRrcConnEstabSucc"] / rrc_formula_denom, 0.0, 1.0),
        1.0,
    )

    s1_formula_denom = np.maximum(
        cell_df["pmS1SigConnEstabAtt"] - cell_df["pmS1SigConnEstabFailMmeOvlMos"], 0.0
    )
    s1_sr = np.where(
        s1_formula_denom > 0,
        np.clip(cell_df["pmS1SigConnEstabSucc"] / s1_formula_denom, 0.0, 1.0),
        1.0,
    )

    erab_sr = np.where(
        cell_df["pmErabEstabAttInit"] > 0,
        np.clip(cell_df["pmErabEstabSuccInit"] / cell_df["pmErabEstabAttInit"], 0.0, 1.0),
        1.0,
    )

    cell_df["CSSR_Denom"] = rrc_stat_denom
    cell_df["RRC_SR"] = rrc_sr * 100.0
    cell_df["S1_SR"] = s1_sr * 100.0
    cell_df["ERAB_SR"] = erab_sr * 100.0
    cell_df["CSSR"] = rrc_sr * s1_sr * erab_sr * 100.0

    # 2.2 Retainability (Overall DCR %, eNB DCR %, Session-Time Normalized Drop Rate)
    dcr_denom_3gpp = (
        cell_df["pmErabRelAbnormalEnb"]
        + cell_df["pmErabRelNormalEnb"]
        + cell_df["pmErabRelMme"]
    )
    dcr_denom_act = cell_df["pmErabRelAbnormalEnbAct"] + cell_df["pmErabRelNormalEnbAct"]
    dcr_denom = np.where(has_3gpp_dcr_denom, dcr_denom_3gpp, dcr_denom_act)

    tot_active_drops = cell_df["pmErabRelAbnormalEnbAct"] + cell_df["pmErabRelAbnormalMmeAct"]
    cell_df["DCR_Denom"] = dcr_denom
    cell_df["DCR_Overall"] = np.where(
        dcr_denom > 0, tot_active_drops / np.maximum(dcr_denom, 1.0) * 100.0, 0.0
    )
    cell_df["DCR_eNB"] = np.where(
        dcr_denom > 0, cell_df["pmErabRelAbnormalEnbAct"] / np.maximum(dcr_denom, 1.0) * 100.0, 0.0
    )
    cell_df["DCR_eNB_Act"] = np.where(
        dcr_denom_act > 0,
        cell_df["pmErabRelAbnormalEnbAct"] / np.maximum(dcr_denom_act, 1.0) * 100.0,
        0.0,
    )
    # Primary DCR alias points to 3GPP Overall Retainability Percentage Lost (Section 2.2.1 & 4.2)
    cell_df["DCR"] = cell_df["DCR_Overall"]

    # Session-Time Normalized Drop Rate (Drops per Active Session Hour: 1 hr = 3,600 s;
    # auto-scales if pmSessionTimeUe is exported in milliseconds)
    med_sec_per_rel = float(
        (cell_df["pmSessionTimeUe"] / np.maximum(dcr_denom, 1.0)).median()
    ) if len(cell_df) > 0 else 2.5
    norm_scale = 36_000_000.0 if med_sec_per_rel > 1000.0 else 3_600.0
    cell_df["DCR_Norm"] = np.where(
        cell_df["pmSessionTimeUe"] > 0,
        tot_active_drops / np.maximum(cell_df["pmSessionTimeUe"], 1.0) * norm_scale,
        0.0,
    )
    cell_df["DCR_Norm_10kHr"] = np.where(
        cell_df["pmSessionTimeUe"] > 0,
        tot_active_drops / np.maximum(cell_df["pmSessionTimeUe"], 1.0) * 36_000_000.0,
        0.0,
    )

    # 2.3 Mobility (HOSR % across IntraF, InterF, IRAT, 1xRTT SRVCC)
    ir_prep_att = np.where(
        cell_df["pmHoPrepAtt"] > 0,
        cell_df["pmHoPrepAtt"],
        cell_df["UtranCellRelation.pmHoPrepAtt"],
    )
    ir_prep_succ = np.where(
        cell_df["pmHoPrepSucc"] > 0,
        cell_df["pmHoPrepSucc"],
        cell_df["UtranCellRelation.pmHoPrepSucc"],
    )
    ir_exe_att = np.where(
        cell_df["pmHoExeAtt"] > 0,
        cell_df["pmHoExeAtt"],
        cell_df["UtranCellRelation.pmHoExeAtt"],
    )
    ir_exe_succ = np.where(
        cell_df["pmHoExeSucc"] > 0,
        cell_df["pmHoExeSucc"],
        cell_df["UtranCellRelation.pmHoExeSucc"],
    )

    ho_prep_succ = (
        cell_df["pmCellHoPrepSuccLteIntraF"]
        + cell_df["pmCellHoPrepSuccLteInterF"]
        + ir_prep_succ
        + cell_df["pmHoPrepSucc1xRttSrvcc"]
    )
    ho_prep_att = (
        cell_df["pmCellHoPrepAttLteIntraF"]
        + cell_df["pmCellHoPrepAttLteInterF"]
        + ir_prep_att
        + cell_df["pmHoPrepAtt1xRttSrvcc"]
        + cell_df["pmHoPrepRejOutUlThres"]
    )
    ho_exe_succ = (
        cell_df["pmCellHoExeSuccLteIntraF"]
        + cell_df["pmCellHoExeSuccLteInterF"]
        + ir_exe_succ
    )
    ho_exe_att = (
        cell_df["pmCellHoExeAttLteIntraF"]
        + cell_df["pmCellHoExeAttLteInterF"]
        + ir_exe_att
    )

    ho_prep_rate = np.where(ho_prep_att > 0, np.clip(ho_prep_succ / np.maximum(ho_prep_att, 1.0), 0.0, 1.0), 1.0)
    ho_exe_rate = np.where(ho_exe_att > 0, np.clip(ho_exe_succ / np.maximum(ho_exe_att, 1.0), 0.0, 1.0), 1.0)
    cell_df["HO_Denom"] = ho_prep_att
    cell_df["HO_Prep_SR"] = ho_prep_rate * 100.0
    cell_df["HO_Exe_SR"] = ho_exe_rate * 100.0
    cell_df["HOSR"] = ho_prep_rate * ho_exe_rate * 100.0

    cell_df["RACH_OOC_Ratio"] = np.where(
        cell_df["pmRaAttCbra"] > 0,
        cell_df["pmRaFailCbraMsg1DiscOoc"] / np.maximum(cell_df["pmRaAttCbra"], 1.0) * 100.0,
        0.0,
    )

    # 2.5 3G/4G IRAT Ping-Pong Oscillation Rate
    utran_ho_succ = np.where(
        cell_df["UtranCellRelation.pmHoExeSucc"] > 0,
        cell_df["UtranCellRelation.pmHoExeSucc"],
        cell_df["pmHoExeSucc"],
    )
    irat_denom = cell_df["pmUeCtxtRelSCWcdma"] + utran_ho_succ
    cell_df["IRAT_Denom"] = irat_denom
    if "pmRapid3gTo4gReentries" in present_counters:
        cell_df["IRAT_PingPong_Rate"] = np.where(
            irat_denom > 0,
            cell_df["pmRapid3gTo4gReentries"] / np.maximum(irat_denom, 1.0) * 100.0,
            0.0,
        )
    elif has_irat:
        cell_df["IRAT_PingPong_Rate"] = np.where(
            cell_df["pmRrcConnEstabSucc"] > 0,
            irat_denom / np.maximum(cell_df["pmRrcConnEstabSucc"], 1.0) * 100.0,
            0.0,
        )
    else:
        cell_df["IRAT_PingPong_Rate"] = 0.0

    # 2.6 VoLTE Accessibility (23.Q3+ Added QCI=1 + Legacy QCI=1 Fallback)
    has_added_qci1 = "pmErabEstabAttAddedQci[1]" in present_counters and cell_df["pmErabEstabAttAddedQci[1]"].sum() > 0
    if has_added_qci1:
        v_raw_att = cell_df["pmErabEstabAttAddedQci[1]"]
        v_net_att = np.maximum(
            v_raw_att
            - cell_df["pmErabEstabAttAddedHoOngoingQci[1]"]
            - cell_df["pmErabEstabAttAddedCsfbQci[1]"]
            - cell_df["pmErabEstabFailAddedErabIdInUseQci[1]"],
            0.0,
        )
        v_bearer_sr = np.where(
            v_net_att > 0,
            np.clip(cell_df["pmErabEstabSuccAddedQci[1]"] / np.maximum(v_net_att, 1.0), 0.0, 1.0) * 100.0,
            100.0,
        )
        cell_df["VoLTE_CSSR_Denom"] = v_raw_att
    else:
        v_raw_att = cell_df["pmErabEstabAttQci[1]"]
        v_bearer_sr = np.where(
            v_raw_att > 0,
            np.clip(cell_df["pmErabEstabSuccQci[1]"] / np.maximum(v_raw_att, 1.0), 0.0, 1.0) * 100.0,
            100.0,
        )
        cell_df["VoLTE_CSSR_Denom"] = v_raw_att

    if "pmErabEstabAttInitQci[5]" in present_counters and cell_df["pmErabEstabAttInitQci[5]"].sum() > 0:
        qci5_erab_sr = np.where(
            cell_df["pmErabEstabAttInitQci[5]"] > 0,
            np.clip(cell_df["pmErabEstabSuccInitQci[5]"] / np.maximum(cell_df["pmErabEstabAttInitQci[5]"], 1.0), 0.0, 1.0),
            1.0,
        )
        cssr_qci5 = rrc_sr * s1_sr * qci5_erab_sr * 100.0
    else:
        cssr_qci5 = cell_df["CSSR"].values

    cell_df["CSSR_Default_Bearer_QCI5"] = cssr_qci5
    cell_df["VoLTE_Bearer_SR"] = v_bearer_sr
    cell_df["VoLTE_CSSR"] = (cssr_qci5 / 100.0) * v_bearer_sr

    # 2.7 VoLTE Retainability (3GPP VoLTE DCR % & Normalized Voice Drops)
    has_3gpp_vdcr = any(
        c in present_counters and cell_df[c].sum() > 0
        for c in ("pmErabRelAbnormalEnbQci[1]", "pmErabRelNormalEnbQci[1]", "pmErabRelMmeQci[1]")
    )
    vdcr_denom_3gpp = (
        cell_df["pmErabRelAbnormalEnbQci[1]"]
        + cell_df["pmErabRelNormalEnbQci[1]"]
        + cell_df["pmErabRelMmeQci[1]"]
    )
    vdcr_denom_act = (
        cell_df["pmErabRelAbnormalEnbActQci[1]"] + cell_df["pmErabRelNormalEnbActQci[1]"]
    )
    vdcr_denom = np.where(has_3gpp_vdcr, vdcr_denom_3gpp, vdcr_denom_act)
    v_tot_drops = cell_df["pmErabRelAbnormalEnbActQci[1]"] + cell_df["pmErabRelAbnormalMmeActQci[1]"]

    cell_df["VoLTE_DCR_Denom"] = vdcr_denom
    cell_df["VoLTE_DCR"] = np.where(
        vdcr_denom > 0, v_tot_drops / np.maximum(vdcr_denom, 1.0) * 100.0, 0.0
    )
    cell_df["VoLTE_DCR_Norm"] = np.where(
        cell_df["pmSessionTimeDrbQci[1]"] > 0,
        v_tot_drops / np.maximum(cell_df["pmSessionTimeDrbQci[1]"], 1.0) * 3_600.0,
        0.0,
    )

    # --- Merge CQI / Interference Long Sheet or Wide Columns ---
    rop_radio_df = pd.DataFrame()
    if cqi_long_frames:
        df_cqi_all = pd.concat(cqi_long_frames, ignore_index=True)
        cell_radio_df, rop_radio_df = _process_cqi_long_sheet(df_cqi_all)
        cell_df = cell_df.merge(cell_radio_df, on="EUtranCellFDD", how="left")
    elif has_wide_cqi:
        cqi_mat = cell_df[wide_cqi_cols].values
        cqi_tot = cqi_mat.sum(axis=1)
        if has_wide_rank:
            r_mat = cell_df[wide_rank_cols].values
            r_tot = r_mat.sum(axis=1)
            w_rank1 = np.where(r_tot > 0, r_mat[:, 0] / np.maximum(r_tot, 1.0) * 100.0, 40.0)
        else:
            w_rank1 = 40.0

        if has_wide_intf:
            i_mat = cell_df[wide_intf_cols].values
            i_tot = i_mat.sum(axis=1)
            i_mw = np.where(i_tot > 0, (i_mat @ INTF_PDF_MIDPOINTS_MW) / np.maximum(i_tot, 1.0), np.nan)
            w_intf_dbm = np.where(i_mw > 0, 10.0 * np.log10(np.maximum(np.nan_to_num(i_mw, nan=1e-30), 1e-30)), -118.0)
            w_intf_samples = i_tot
        else:
            w_intf_dbm = cell_df["pmRadioRecInterferencePwr"] if has_scalar_intf else -118.0
            w_intf_samples = 0.0

        wide_extra: Dict[str, Any] = {
            "CQI_Samples": cqi_tot,
            "Avg_CQI": np.where(
                cqi_tot > 0, (cqi_mat @ np.arange(16, dtype=np.float64)) / np.maximum(cqi_tot, 1.0), np.nan
            ),
            "Low_CQI_Ratio": np.where(
                cqi_tot > 0, cqi_mat[:, :7].sum(axis=1) / np.maximum(cqi_tot, 1.0) * 100.0, 0.0
            ),
            "Hom_256QAM_Ratio": np.nan,
            "Rank1_Ratio": w_rank1,
            "UL_Intf_dBm": w_intf_dbm,
            "Intf_Samples": w_intf_samples,
        }
        for b in range(16):
            wide_extra[f"CQI_Bin_{b}"] = cqi_mat[:, b]
        cell_df = pd.concat([cell_df, pd.DataFrame(wide_extra, index=cell_df.index)], axis=1)
    else:
        fallback_radio: Dict[str, Any] = {
            "CQI_Samples": 0.0,
            "Avg_CQI": np.nan,
            "Low_CQI_Ratio": 0.0,
            "Hom_256QAM_Ratio": np.nan,
            "Rank1_Ratio": np.nan,
            "UL_Intf_dBm": cell_df["pmRadioRecInterferencePwr"] if has_scalar_intf else -118.0,
            "Intf_Samples": 0.0,
        }
        for b in range(16):
            fallback_radio[f"CQI_Bin_{b}"] = 0.0
        cell_df = pd.concat(
            [cell_df.copy(), pd.DataFrame(fallback_radio, index=cell_df.index)], axis=1
        )

    # CQI Modulation Band Shares per cell
    cqi_tot_safe = np.where(cell_df["CQI_Samples"] > 0, cell_df["CQI_Samples"], np.nan)
    cell_df["CQI_QPSK_Pct"] = (
        cell_df[[f"CQI_Bin_{i}" for i in range(0, 7)]].sum(axis=1) / cqi_tot_safe * 100.0
    ).fillna(0.0)
    cell_df["CQI_16QAM_Pct"] = (
        cell_df[[f"CQI_Bin_{i}" for i in range(7, 10)]].sum(axis=1) / cqi_tot_safe * 100.0
    ).fillna(0.0)
    cell_df["CQI_64QAM_Pct"] = (
        cell_df[[f"CQI_Bin_{i}" for i in range(10, 13)]].sum(axis=1) / cqi_tot_safe * 100.0
    ).fillna(0.0)
    cell_df["CQI_256QAM_Pct"] = (
        cell_df[[f"CQI_Bin_{i}" for i in range(13, 16)]].sum(axis=1) / cqi_tot_safe * 100.0
    ).fillna(0.0)

    # --- Merge PRB Spectrum Sheet ---
    prb_cell_map: Dict[str, List[float]] = {}
    has_prb = False
    if prb_frames:
        df_prb_all = pd.concat(prb_frames, ignore_index=True)
        intf_samples_series = cell_df.set_index("EUtranCellFDD")["Intf_Samples"]
        prb_summary, prb_cell_map = _process_prb_spectrum_sheet(df_prb_all, intf_samples_series)
        if not prb_summary.empty:
            has_prb = True
            cell_df = cell_df.merge(prb_summary, on="EUtranCellFDD", how="left")

    if not has_prb:
        cell_df["PRB_Max_dBm"] = cell_df["UL_Intf_dBm"]
        cell_df["PRB_Median_dBm"] = cell_df["UL_Intf_dBm"]
        cell_df["PRB_Crest_dB"] = 0.0
        cell_df["Peak_PRBs_Text"] = "N/A"

    # --- Execute Decision Trees across all evaluated cells ---
    rca_records = [
        _diagnose_cell_row(row, has_irat=has_irat, has_volte=has_volte)
        for _, row in cell_df.iterrows()
    ]
    rca_df = pd.DataFrame(rca_records)
    cell_df = pd.concat([cell_df.reset_index(drop=True), rca_df], axis=1)

    # Sort Master Grid: CRITICAL -> DEGRADED -> NORMAL -> LOW_SAMPLE, then DCR_Overall desc, CSSR asc
    cell_df = cell_df.sort_values(
        by=["Severity_Rank", "DCR_Overall", "CSSR"],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    # --- Build ROP Time-Series Trend DataFrame ---
    rop_req_cols = (
        "pmRrcConnEstabAtt", "pmRrcConnEstabAttReatt", "pmRrcConnEstabSucc",
        "pmRrcConnEstabFailMmeOvlMos", "pmRrcConnEstabFailMmeOvlMod",
        "pmS1SigConnEstabAtt", "pmS1SigConnEstabSucc", "pmS1SigConnEstabFailMmeOvlMos",
        "pmErabEstabAttInit", "pmErabEstabSuccInit",
        "pmErabRelAbnormalEnb", "pmErabRelNormalEnb", "pmErabRelMme",
        "pmErabRelNormalEnbAct", "pmErabRelAbnormalEnbAct",
        "pmErabRelAbnormalMmeAct", "pmSessionTimeUe",
    )
    rop_num_cols = [c for c in rop_req_cols if c in df_main.columns]

    if rop_num_cols:
        rop_grp = (
            df_main.groupby(["ROP_Time", "EUtranCellFDD"], observed=False, as_index=False)[rop_num_cols]
            .sum()
        )
        missing_rop = {c: 0.0 for c in rop_req_cols if c not in rop_grp.columns}
        if missing_rop:
            rop_grp = pd.concat([rop_grp, pd.DataFrame(missing_rop, index=rop_grp.index)], axis=1)

        r_denom = np.maximum(
            rop_grp["pmRrcConnEstabAtt"]
            - rop_grp["pmRrcConnEstabAttReatt"]
            - rop_grp["pmRrcConnEstabFailMmeOvlMos"]
            - rop_grp["pmRrcConnEstabFailMmeOvlMod"],
            0.0,
        )
        r_sr = np.where(r_denom > 0, np.clip(rop_grp["pmRrcConnEstabSucc"] / np.maximum(r_denom, 1.0), 0.0, 1.0), 1.0)
        s_denom = np.maximum(
            rop_grp["pmS1SigConnEstabAtt"] - rop_grp["pmS1SigConnEstabFailMmeOvlMos"], 0.0
        )
        s_sr = np.where(s_denom > 0, np.clip(rop_grp["pmS1SigConnEstabSucc"] / np.maximum(s_denom, 1.0), 0.0, 1.0), 1.0)
        e_sr = np.where(
            rop_grp["pmErabEstabAttInit"] > 0,
            np.clip(rop_grp["pmErabEstabSuccInit"] / np.maximum(rop_grp["pmErabEstabAttInit"], 1.0), 0.0, 1.0),
            1.0,
        )
        rop_grp["CSSR"] = r_sr * s_sr * e_sr * 100.0

        d_denom_3gpp = (
            rop_grp["pmErabRelAbnormalEnb"]
            + rop_grp["pmErabRelNormalEnb"]
            + rop_grp["pmErabRelMme"]
        )
        d_denom_act = rop_grp["pmErabRelAbnormalEnbAct"] + rop_grp["pmErabRelNormalEnbAct"]
        d_denom = np.where(has_3gpp_dcr_denom, d_denom_3gpp, d_denom_act)
        rop_drops = rop_grp["pmErabRelAbnormalEnbAct"] + rop_grp["pmErabRelAbnormalMmeAct"]

        rop_grp["DCR_Overall"] = np.where(d_denom > 0, rop_drops / np.maximum(d_denom, 1.0) * 100.0, 0.0)
        rop_grp["DCR_eNB"] = np.where(
            d_denom > 0, rop_grp["pmErabRelAbnormalEnbAct"] / np.maximum(d_denom, 1.0) * 100.0, 0.0
        )
        rop_grp["DCR"] = rop_grp["DCR_Overall"]
        rop_grp["DCR_Norm"] = np.where(
            rop_grp["pmSessionTimeUe"] > 0,
            rop_drops / np.maximum(rop_grp["pmSessionTimeUe"], 1.0) * norm_scale,
            0.0,
        )
        rop_trend_df = rop_grp.sort_values("ROP_Time").reset_index(drop=True)
    else:
        rop_trend_df = pd.DataFrame(
            columns=["ROP_Time", "EUtranCellFDD", "CSSR", "DCR", "DCR_Overall", "DCR_eNB", "DCR_Norm"]
        )

    if not rop_radio_df.empty and not rop_trend_df.empty:
        rop_trend_df = rop_trend_df.merge(rop_radio_df, on="ROP_Time", how="left")

    cqi_bin_cols = ["EUtranCellFDD", "eNodeB_Name"] + [f"CQI_Bin_{b}" for b in range(16)]
    cqi_dist_df = cell_df[cqi_bin_cols].copy()

    return AnalysisResult(
        source_name=source_name,
        ingestion_timestamp=ingestion_ts,
        data_start_time=data_start,
        data_end_time=data_end,
        total_rop_rows=total_rop_rows,
        sheets_loaded=sheets_loaded,
        cell_df=cell_df,
        rop_trend_df=rop_trend_df,
        cqi_dist_df=cqi_dist_df,
        prb_cell_map=prb_cell_map,
        has_irat_counters=has_irat,
        has_volte_counters=has_volte,
        has_prb_spectrum=has_prb,
        has_3gpp_dcr_denom=has_3gpp_dcr_denom,
    )


def compute_rop_series(rop_trend_df: pd.DataFrame, cells: Optional[List[str]] = None) -> pd.DataFrame:
    """Aggregate ROP time-series KPIs (CSSR, 3GPP Overall DCR, eNB DCR) across filtered subset of cells."""
    empty_cols = ["ROP_Time", "CSSR", "DCR", "DCR_Overall", "DCR_eNB", "DCR_Norm"]
    if rop_trend_df.empty:
        return pd.DataFrame(columns=empty_cols)

    sub = rop_trend_df[rop_trend_df["EUtranCellFDD"].isin(cells)] if cells is not None else rop_trend_df
    if sub.empty:
        return pd.DataFrame(columns=empty_cols)

    req_cols = (
        "pmRrcConnEstabAtt", "pmRrcConnEstabAttReatt", "pmRrcConnEstabSucc",
        "pmRrcConnEstabFailMmeOvlMos", "pmRrcConnEstabFailMmeOvlMod",
        "pmS1SigConnEstabAtt", "pmS1SigConnEstabSucc", "pmS1SigConnEstabFailMmeOvlMos",
        "pmErabEstabAttInit", "pmErabEstabSuccInit",
        "pmErabRelAbnormalEnb", "pmErabRelNormalEnb", "pmErabRelMme",
        "pmErabRelNormalEnbAct", "pmErabRelAbnormalEnbAct",
        "pmErabRelAbnormalMmeAct", "pmSessionTimeUe",
    )
    sum_cols = [c for c in req_cols if c in sub.columns]
    ts = sub.groupby("ROP_Time", as_index=False)[sum_cols].sum().sort_values("ROP_Time")
    for c in req_cols:
        if c not in ts.columns:
            ts[c] = 0.0

    r_den = np.maximum(
        ts["pmRrcConnEstabAtt"]
        - ts["pmRrcConnEstabAttReatt"]
        - ts["pmRrcConnEstabFailMmeOvlMos"]
        - ts["pmRrcConnEstabFailMmeOvlMod"],
        0.0,
    )
    r_sr = np.where(r_den > 0, np.clip(ts["pmRrcConnEstabSucc"] / np.maximum(r_den, 1.0), 0.0, 1.0), 1.0)
    s_den = np.maximum(ts["pmS1SigConnEstabAtt"] - ts["pmS1SigConnEstabFailMmeOvlMos"], 0.0)
    s_sr = np.where(s_den > 0, np.clip(ts["pmS1SigConnEstabSucc"] / np.maximum(s_den, 1.0), 0.0, 1.0), 1.0)
    e_sr = np.where(
        ts["pmErabEstabAttInit"] > 0,
        np.clip(ts["pmErabEstabSuccInit"] / np.maximum(ts["pmErabEstabAttInit"], 1.0), 0.0, 1.0),
        1.0,
    )
    ts["CSSR"] = r_sr * s_sr * e_sr * 100.0

    d_den_3gpp = ts["pmErabRelAbnormalEnb"] + ts["pmErabRelNormalEnb"] + ts["pmErabRelMme"]
    d_den_act = ts["pmErabRelAbnormalEnbAct"] + ts["pmErabRelNormalEnbAct"]
    use_3gpp = float(d_den_3gpp.sum()) > 0
    d_den = d_den_3gpp if use_3gpp else d_den_act
    tot_drops = ts["pmErabRelAbnormalEnbAct"] + ts["pmErabRelAbnormalMmeAct"]

    ts["DCR_Overall"] = np.where(d_den > 0, tot_drops / np.maximum(d_den, 1.0) * 100.0, 0.0)
    ts["DCR_eNB"] = np.where(
        d_den > 0, ts["pmErabRelAbnormalEnbAct"] / np.maximum(d_den, 1.0) * 100.0, 0.0
    )
    ts["DCR"] = ts["DCR_Overall"]
    ts["DCR_Norm"] = np.where(
        ts["pmSessionTimeUe"] > 0,
        tot_drops / np.maximum(ts["pmSessionTimeUe"], 1.0) * 3_600.0,
        0.0,
    )
    return ts
