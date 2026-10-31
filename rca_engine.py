"""
Ericsson LTE RAN Performance Troubleshooting & Root-Cause Prediction Engine.

Implements ETL ingestion (CSV / multi-sheet Excel with metadata headers),
3GPP/Ericsson MOM composite KPI calculations, statistical sample-size
validation (Minimum Denominator Rule), deterministic multi-domain RCA
decision trees, and automated AMOS/Moshell CLI remediation script generation.
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
# 1. CONSTANTS, BASELINES & ERICSSON MOM COUNTER TAXONOMY
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
    # Time
    "date_time": "ROP_Time",
    "datetime": "ROP_Time",
    "rop_time": "ROP_Time",
    "roptime": "ROP_Time",
    "timestamp": "ROP_Time",
    "date": "Date",
    "date_id": "Date",
    "hour": "Hour",
    "hour_id": "Hour",
    "vectorindex": "Vector_Index",
    "vector_index": "Vector_Index",
    "dcvector_index": "Vector_Index",
}

# Core scalar counters used across KPI & RCA formulas
SCALAR_COUNTERS: List[str] = [
    # RACH & RRC Accessibility
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
    # S1 & Initial E-RAB
    "pmS1SigConnEstabAtt",
    "pmS1SigConnEstabSucc",
    "pmS1SigConnEstabFailMmeOvlMosCe",
    "pmErabEstabAttInit",
    "pmErabEstabSuccInit",
    "pmErabEstabFailInitLic",
    "pmErabEstabFailGbrDlEnb",
    "pmErabEstabFailGbrUlEnb",
    # Retainability (DCR)
    "pmErabRelNormalEnbAct",
    "pmErabRelAbnormalEnbAct",
    "pmErabRelAbnormalEnbActUeLost",
    "pmErabRelAbnormalEnbUeLostOos",
    "pmErabRelAbnormalEnbActHo",
    "pmErabRelAbnormalEnbActTnFail",
    "pmErabRelAbnormalMmeAct",
    "pmErabRelAbnormalEnbActPe",
    "pmSessionTimeUe",
    # Mobility (Handover)
    "pmCellHoPrepAttLteIntraF",
    "pmCellHoPrepSuccLteIntraF",
    "pmCellHoExeAttLteIntraF",
    "pmCellHoExeSuccLteIntraF",
    "pmCellHoPrepAttLteInterF",
    "pmCellHoPrepSuccLteInterF",
    "pmCellHoExeAttLteInterF",
    "pmCellHoExeSuccLteInterF",
    "pmHoPrepRejInLicConnUsers",
    "pmHoPrepRejOutUlThres",
    # IRAT 3G/4G Ping-Pong & Mobility
    "pmBadCovEvalReport",
    "pmBadCovSearchEvalReport",
    "pmB2BestCellEvalReport",
    "pmUeCtxtRelSCWcdma",
    "pmUeCtxtRelCsfbWcdma",
    "UtranCellRelation.pmHoPrepAtt",
    "UtranCellRelation.pmHoPrepSucc",
    "UtranCellRelation.pmHoExeSucc",
    "pmRapid3gTo4gReentries",
    "b2Threshold1Rsdp",
    "b2Threshold2RsdpUtran",
    "b2Threshold2EcNoUtran",
    "timeToTriggerB2",
    "hysteresisB2",
    # VoLTE QCI=1
    "pmErabEstabAttQci[1]",
    "pmErabEstabSuccQci[1]",
    "pmErabRelNormalEnbActQci[1]",
    "pmErabRelAbnormalEnbActQci[1]",
    "pmErabRelAbnormalEnbActUeLostQci[1]",
    "pmErabRelAbnormalEnbActHoQci[1]",
    "pmErabRelAbnormalMmeActQci[1]",
    "pmErabRelAbnormalEnbActPeQci[1]",
    "pmSessionTimeUeQci[1]",
    "pmRlcPduRetransUlQci[1]",
    "pmRlcPduRetransDlQci[1]",
    "pmHoPrepAttSrvcc",
    "pmHoExeSuccSrvcc",
]

COUNTER_LOWER_MAP: Dict[str, str] = {c.lower(): c for c in SCALAR_COUNTERS}
COUNTER_LOWER_MAP.update(
    {
        "utrancellrelation_pmhoprepatt": "UtranCellRelation.pmHoPrepAtt",
        "utrancellrelation_pmhoprepsucc": "UtranCellRelation.pmHoPrepSucc",
        "utrancellrelation_pmhoexesucc": "UtranCellRelation.pmHoExeSucc",
        "rapid_3g_to_4g_reentries": "pmRapid3gTo4gReentries",
        "rapid3gto4greentries": "pmRapid3gTo4gReentries",
        "pmerabestabattqci_1": "pmErabEstabAttQci[1]",
        "pmerabestabsuccqci_1": "pmErabEstabSuccQci[1]",
        "pmerabrelnormalenbactqci_1": "pmErabRelNormalEnbActQci[1]",
        "pmerabrelabnormalenbactqci_1": "pmErabRelAbnormalEnbActQci[1]",
        "pmerabrelabnormalenbactuelostqci_1": "pmErabRelAbnormalEnbActUeLostQci[1]",
        "pmerabrelabnormalenbacthoqci_1": "pmErabRelAbnormalEnbActHoQci[1]",
        "pmerabrelabnormalmmeactqci_1": "pmErabRelAbnormalMmeActQci[1]",
        "pmerabrelabnormalenbactpeqci_1": "pmErabRelAbnormalEnbActPeQci[1]",
        "pmsessiontimeueqci_1": "pmSessionTimeUeQci[1]",
    }
)


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

    # Remove empty placeholder columns and fully blank rows
    keep_cols = [
        c for c in df.columns
        if not c.startswith("_drop_") and not c.lower().startswith("unnamed:")
    ]
    df = df[keep_cols].dropna(how="all").reset_index(drop=True)
    return df


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Standardize OSS/ENM/ENIQ column names into canonical MOM schema."""
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
            # Normalize vector notation e.g. pmRadioUeRepCqiDistr_0 -> pmRadioUeRepCqiDistr[0]
            m_vec = re.match(r"^(pmradio\w+distr)[_\[](\d+)[\]]?$", c_strip, re.IGNORECASE)
            if m_vec:
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
    # Deduplicate columns if both ERBS Id and ManagedElement existed
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
        # Handle Excel serial dates if numeric
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
            return base_dt + pd.to_timedelta(hr, unit="h")
        return base_dt

    return pd.Series(pd.Timestamp("2026-10-28 00:00:00"), index=df.index)


# ============================================================================
# 3. MULTI-SHEET / CSV INGESTION & VECTOR PIVOTING PIPELINE
# ============================================================================

def _classify_sheet(df: pd.DataFrame) -> str:
    """Classify normalized DataFrame into scalar KPI, long vector CQI, or PRB spectrum."""
    cols = set(df.columns)
    prb_cols = [c for c in cols if c.startswith("pmRadioRecInterferencePwrPrb")]
    if len(prb_cols) >= 10:
        return "prb_spectrum"
    if "Vector_Index" in cols and (
        "pmRadioUeRepCqiDistr" in cols or "pmRadioRecInterferencePwr" in cols
    ):
        return "cqi_vector_long"
    if any(c.startswith("pm") for c in cols):
        return "kpi_main"
    return "unknown"


def _process_cqi_long_sheet(df_cqi: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Vectorized transformation of long-format CqiDistr sheet (Vector_Index 0..15)
    into:
      1) Per-(Cell, ROP_Time) radio quality metrics
      2) Per-Cell aggregated CQI bin counts [0..15], Rank 1 ratio, and UL Interference (dBm)
    """
    df = df_cqi.copy()
    df["ROP_Time"] = _build_rop_timestamp(df)
    if "eNodeB_Name" not in df.columns and "EUtranCellFDD" in df.columns:
        df["eNodeB_Name"] = df["EUtranCellFDD"].astype(str).str.slice(0, 14)

    df["Vector_Index"] = pd.to_numeric(df["Vector_Index"], errors="coerce")
    df = df[df["Vector_Index"].isin(range(16))].copy()
    df["Vector_Index"] = df["Vector_Index"].astype(int)

    for col in ("pmRadioUeRepCqiDistr", "pmRadioUeRepRankDistr", "pmRadioRecInterferencePwr"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        else:
            df[col] = 0.0

    # --- A. Cell-level CQI Bins (0..15), Rank 1..4, and Interference PDF ---
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
    avg_cqi = np.where(cqi_tot > 0, cqi_weighted / cqi_tot, np.nan)
    low_cqi_pct = np.where(
        cqi_tot > 0, cqi_cell_piv[list(range(7))].sum(axis=1) / cqi_tot * 100.0, 0.0
    )

    # Rank 1 Ratio (bins 0..3 correspond to Rank 1..4)
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
    rank1_pct = np.where(rank_tot > 0, rank_cell_piv[0] / rank_tot * 100.0, np.nan)

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
        intf_tot > 0, (intf_cell_piv.values @ INTF_PDF_MIDPOINTS_MW) / np.maximum(intf_tot, 1.0), np.nan
    )
    ul_intf_dbm = np.where(intf_mw > 0, 10.0 * np.log10(np.maximum(np.nan_to_num(intf_mw, nan=1e-30), 1e-30)), np.nan)

    cell_radio_df = pd.DataFrame(
        {
            "EUtranCellFDD": cqi_cell_piv.index,
            "CQI_Samples": cqi_tot.values,
            "Avg_CQI": avg_cqi,
            "Low_CQI_Ratio": low_cqi_pct,
            "Rank1_Ratio": rank1_pct,
            "UL_Intf_dBm": ul_intf_dbm,
            "Intf_Samples": intf_tot.values,
        }
    )
    for b in range(16):
        cell_radio_df[f"CQI_Bin_{b}"] = cqi_cell_piv[b].values

    # --- B. ROP-level Network CQI & Interference for Trendlines ---
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
        rop_tot > 0, (rop_cqi_piv.values @ np.arange(16, dtype=np.float64)) / rop_tot, np.nan
    )
    rop_low_cqi = np.where(
        rop_tot > 0, rop_cqi_piv[list(range(7))].sum(axis=1) / rop_tot * 100.0, 0.0
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
    Convert pmRadioRecInterferencePwrPrb1..100 into calibrated per-PRB dBm profiles
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

    # Determine 40ms measurement sample count per cell
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
# 4. COMPOSITE KPI COMPUTATION & DETERMINISTIC DECISION TREES
# ============================================================================

def _diagnose_cell_row(
    row: pd.Series, has_irat: bool, has_volte: bool
) -> Dict[str, Any]:
    """
    Execute Pre-Execution Sample Gate (4.0) + Sequential Decision Trees (4.1 - 4.7)
    for a single aggregated cell row, returning RCA classification, confidence,
    evidence items, and ready-to-run AMOS/Moshell script.
    """
    enb = str(row.get("eNodeB_Name", "UNKNOWN_ENB"))
    cell = str(row.get("EUtranCellFDD", "UNKNOWN_CELL"))

    cssr = float(row.get("CSSR", 100.0))
    cssr_denom = float(row.get("CSSR_Denom", 0.0))
    rrc_sr = float(row.get("RRC_SR", 100.0))
    s1_sr = float(row.get("S1_SR", 100.0))
    erab_sr = float(row.get("ERAB_SR", 100.0))

    dcr = float(row.get("DCR", 0.0))
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

    ul_intf = float(row.get("UL_Intf_dBm", -118.0)) if pd.notna(row.get("UL_Intf_dBm")) else -118.0
    rank1_pct = float(row.get("Rank1_Ratio", 40.0)) if pd.notna(row.get("Rank1_Ratio")) else 40.0
    prb_crest = float(row.get("PRB_Crest_dB", 0.0)) if pd.notna(row.get("PRB_Crest_dB")) else 0.0
    prb_max_dbm = float(row.get("PRB_Max_dBm", ul_intf)) if pd.notna(row.get("PRB_Max_dBm")) else ul_intf
    peak_prbs_txt = str(row.get("Peak_PRBs_Text", "N/A"))

    irat_pp = float(row.get("IRAT_PingPong_Rate", 0.0))
    irat_denom = float(row.get("IRAT_Denom", 0.0))

    volte_cssr = float(row.get("VoLTE_CSSR", 100.0))
    volte_cssr_denom = float(row.get("VoLTE_CSSR_Denom", 0.0))
    volte_dcr = float(row.get("VoLTE_DCR", 0.0))
    volte_dcr_denom = float(row.get("VoLTE_DCR_Denom", 0.0))

    # Validity gates (Section 2.8 & 4.0)
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

    # Track low-sample exclusions where raw metric looked degraded
    if not valid_dcr and dcr > 1.0:
        low_sample_notes.append(
            f"Cell {cell} excluded from DCR alarm due to insufficient sample size "
            f"(Releases = {int(dcr_denom)} < 30)."
        )
    if not valid_cssr and cssr < 97.0:
        low_sample_notes.append(
            f"Cell {cell} excluded from CSSR alarm due to insufficient sample size "
            f"(Attempts = {int(cssr_denom)} < 30)."
        )
    if not valid_ho and hosr < 95.0 and ho_denom > 0:
        low_sample_notes.append(
            f"Cell {cell} excluded from HOSR alarm due to insufficient sample size "
            f"(HO Prep Att = {int(ho_denom)} < 20)."
        )

    # ------------------------------------------------------------------
    # TREE 4.1: BAD CSSR (Accessibility) - Trigger at < 98.5% warning or < 97.0% degraded
    # ------------------------------------------------------------------
    ra_att = float(row.get("pmRaAttCbra", 0.0))
    ra_ooc = float(row.get("pmRaFailCbraMsg1DiscOoc", 0.0))
    ra_sched = float(row.get("pmRaFailCbraMsg1DiscSched", 0.0))
    ra_msg2 = float(row.get("pmRaFailCbraMsg2Disc", 0.0))
    ooc_ratio = (ra_ooc / ra_att * 100.0) if ra_att > 0 else 0.0

    if valid_cssr and (cssr < 97.0 or (cssr < 98.5 and ooc_ratio > 5.0)):
        sev = 3 if cssr < 95.0 else 2
        rrc_lic = float(row.get("pmRrcConnEstabFailLic", 0.0))
        rrc_mp = float(row.get("pmRrcConnEstabFailMpOverload", 0.0))
        rrc_du = float(row.get("pmRrcConnEstabFailDuIntens", 0.0))
        mme_ovl = float(row.get("pmRrcConnEstabFailMmeOvlMos", 0.0)) + float(
            row.get("pmS1SigConnEstabFailMmeOvlMosCe", 0.0)
        )
        erab_lic = float(row.get("pmErabEstabFailInitLic", 0.0))
        gbr_fail = float(row.get("pmErabEstabFailGbrDlEnb", 0.0)) + float(
            row.get("pmErabEstabFailGbrUlEnb", 0.0)
        )

        if rrc_sr < 97.0 or ooc_ratio > 5.0:
            if rrc_lic > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "License Congestion",
                    "cause": "License Limitation - Connected Users License Exceeded",
                    "confidence": 96,
                    "recommendation": "Upgrade Connected Users License capacity or activate Dynamic Load Control (DLC).",
                    "scenario": "B_LIC",
                })
            elif rrc_mp > 0 or rrc_du > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "Hardware / Baseband / RRU-RET",
                    "cause": "Baseband / MP CPU Hardware Overload",
                    "confidence": 93,
                    "recommendation": "Offload traffic via IFLB, audit Baseband DU intensity, or expand Baseband capacity.",
                    "scenario": "B_HW",
                })
            elif ooc_ratio > 5.0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "RACH Overshooting (OOC)",
                    "cause": f"Overshooting Cell / RACH Timing Advance Out-of-Coverage ({ooc_ratio:.1f}%)",
                    "confidence": 94,
                    "recommendation": "Increase RET electrical down-tilt, reduce crsGain/pMax, or align EUtranCellFDD.cellRange.",
                    "scenario": "B_OOC",
                })
            elif ul_intf > -110.0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "UL Interference / PIM",
                    "cause": f"Uplink Interference ({ul_intf:.1f} dBm) Blocking RACH Msg3 Detection",
                    "confidence": 92,
                    "recommendation": "Perform PIM spectrum sweep, inspect external UL jammer, or boost PRACH power.",
                    "scenario": "A_PIM",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "PRB / CCE / GBR Congestion",
                    "cause": "Downlink Control Channel (PDCCH/CCE) Congestion or Weak CCE Coverage",
                    "confidence": 86,
                    "recommendation": "Enable Dynamic CCE Optimization (pdcchCceOptimizationEnabled) and check Msg2/Scheduler discards.",
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
                    "recommendation": "Engage EPC Core team to rebalance MME Pool relative capacity.",
                    "scenario": "CORE_MME",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "Transport / Core (S1/MME)",
                    "cause": "S1 Transport Bounded Latency / SCTP Transmission Drops",
                    "confidence": 88,
                    "recommendation": "Inspect S1 IP backhaul link quality, DSCP marking, and SCTP retransmissions.",
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
                    "recommendation": "Install additional E-RAB / Radio Bearer software license capacity.",
                    "scenario": "B_LIC",
                })
            elif gbr_fail > 0:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "PRB / CCE / GBR Congestion",
                    "cause": "GBR Resource Admission Failure (DL/UL Bandwidth Congestion)",
                    "confidence": 92,
                    "recommendation": "Adjust GBR Call Admission Control limits and audit ARP thresholds.",
                    "scenario": "E_VOLTE",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "CSSR",
                    "category": "Transport / Core (S1/MME)",
                    "cause": "Radio Bearer Allocation Failure during Initial Context Setup",
                    "confidence": 84,
                    "recommendation": "Audit QCI profile mapping and verify S1-U GTP-U tunnel binding.",
                    "scenario": "TN_S1",
                })

    # Also detect standalone severe RACH Overshooting (> 5% of CBRA preambles) even if CSSR >= 98.5%
    elif valid_cssr and ra_att >= 500 and ooc_ratio > 5.0:
        findings.append({
            "severity": 2,
            "domain": "RACH",
            "category": "RACH Overshooting (OOC)",
            "cause": f"RACH Out-of-Coverage Preamble Surge ({ooc_ratio:.1f}% OOC Discards)",
            "confidence": 91,
            "recommendation": "Audit RET electrical tilt for overshooting or expand EUtranCellFDD.cellRange.",
            "scenario": "B_OOC",
        })

    # ------------------------------------------------------------------
    # TREE 4.2: HIGH DCR (Retainability)
    # ------------------------------------------------------------------
    abn_enb = float(row.get("pmErabRelAbnormalEnbAct", 0.0))
    abn_ue_lost = float(row.get("pmErabRelAbnormalEnbActUeLost", 0.0))
    abn_ho = float(row.get("pmErabRelAbnormalEnbActHo", 0.0))
    abn_tn = float(row.get("pmErabRelAbnormalEnbActTnFail", 0.0))
    abn_mme = float(row.get("pmErabRelAbnormalMmeAct", 0.0))
    abn_pe = float(row.get("pmErabRelAbnormalEnbActPe", 0.0))

    # Enforce Section 4.0 Pre-Execution Gate: skip Tree 4.2 whenever DCR_Denom < 30
    dcr_triggered = valid_dcr and (dcr > 1.0 or (valid_dcr_norm and dcr_norm > 10.0))
    if dcr_triggered:
        sev = 3 if (valid_dcr and dcr > 2.5) or (valid_dcr_norm and dcr_norm > 25.0) else 2
        ue_lost_share = (abn_ue_lost / abn_enb) if abn_enb > 0 else 0.0
        ho_drop_share = (abn_ho / abn_enb) if abn_enb > 0 else 0.0
        tn_drop_share = (abn_tn / abn_enb) if abn_enb > 0 else 0.0
        pe_drop_share = (abn_pe / abn_enb) if abn_enb > 0 else 0.0
        mme_drop_share = (
            abn_mme / (abn_enb + abn_mme) if (abn_enb + abn_mme) > 0 else 0.0
        )

        if ue_lost_share > 0.50:
            if ul_intf > -108.0 or (ul_intf > -112.0 and prb_crest >= 8.0 and prb_max_dbm > -104.0):
                if prb_crest >= 6.0:
                    findings.append({
                        "severity": sev,
                        "domain": "DCR",
                        "category": "UL Interference / PIM",
                        "cause": f"Passive Intermodulation (PIM) / Narrowband UL Interference (Crest +{prb_crest:.1f} dB)",
                        "confidence": 96,
                        "recommendation": "Schedule PIM sweep on RRU antenna jumpers/RET ports; inspect noisy edge PRBs.",
                        "scenario": "A_PIM",
                    })
                else:
                    findings.append({
                        "severity": sev,
                        "domain": "DCR",
                        "category": "UL Interference / PIM",
                        "cause": f"Broadband Uplink External Interference ({ul_intf:.1f} dBm/PRB)",
                        "confidence": 95,
                        "recommendation": "Deploy spectrum analyzer for external repeater/jammer hunting.",
                        "scenario": "A_PIM",
                    })
            elif avg_cqi < 7.0 and low_cqi > 40.0:
                findings.append({
                    "severity": sev,
                    "domain": "DCR",
                    "category": "Hardware / Baseband / RRU-RET",
                    "cause": f"Severe DL Weak Coverage / RRU-RET Tilt Degradation (Avg CQI {avg_cqi:.1f})",
                    "confidence": 93,
                    "recommendation": "Verify RET actuator calibration & RRU VSWR alarms; optimize down-tilt / crsGain.",
                    "scenario": "A_COV_RET",
                })
            elif ooc_ratio > 3.5:
                findings.append({
                    "severity": sev,
                    "domain": "DCR",
                    "category": "RACH Overshooting (OOC)",
                    "cause": f"Overshooting Cell Edge Radio Link Failure (UE Lost {ue_lost_share*100:.0f}% + RACH OOC)",
                    "confidence": 91,
                    "recommendation": "Tighten RET electrical down-tilt to contain cell footprint and prevent distant UE drops.",
                    "scenario": "B_OOC",
                })
            elif rank1_pct > 75.0:
                findings.append({
                    "severity": sev,
                    "domain": "DCR",
                    "category": "Hardware / Baseband / RRU-RET",
                    "cause": f"RRU TX/RX Branch Imbalance or Severe Pilot Pollution (Rank-1 {rank1_pct:.1f}%)",
                    "confidence": 89,
                    "recommendation": "Check RRU RF port VSWR / RET tilt alignment and reduce co-channel cell overlap.",
                    "scenario": "A_COV_RET",
                })
            else:
                conf = min(94, int(82 + min(12, abn_enb / 15.0)))
                findings.append({
                    "severity": sev,
                    "domain": "DCR",
                    "category": "RF Weak Coverage / UL Power",
                    "cause": f"UL Power Control Bottleneck / Cell-Edge UE Lost ({ue_lost_share*100:.0f}% RLC Max Retries)",
                    "confidence": conf,
                    "recommendation": "Optimize p0NominalPusch & alpha in EUtranCellFDD; audit cell-edge A2/A3 handover margins.",
                    "scenario": "C_ULPC",
                })
        elif ho_drop_share > 0.30:
            if hosr < 90.0:
                findings.append({
                    "severity": sev,
                    "domain": "DCR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": "Handover Execution Failure / Missing Neighbor Relations",
                    "confidence": 94,
                    "recommendation": "Audit ANR state, X2 link status, and advance a3Offset / timeToTrigger.",
                    "scenario": "C_HO",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "DCR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": "Late Handover / Ping-Pong Handover in High Mobility Zone",
                    "confidence": 90,
                    "recommendation": "Tune hysteresisA3, decrease timeToTrigger, and enable Speed Dependent Mobility.",
                    "scenario": "C_HO",
                })
        elif tn_drop_share > 0.20:
            findings.append({
                "severity": sev,
                "domain": "DCR",
                "category": "Transport / Core (S1/MME)",
                "cause": "Backhaul Transport Network Instability / S1-X2 Link Flapping",
                "confidence": 95,
                "recommendation": "Check EthernetPort CRC errors, S1-U packet loss, and microwave/fiber link stability.",
                "scenario": "TN_S1",
            })
        elif pe_drop_share > 0.15:
            findings.append({
                "severity": sev,
                "domain": "DCR",
                "category": "PRB / CCE / GBR Congestion",
                "cause": "Pre-emption Drop due to High Priority ARP Bearer Admission",
                "confidence": 93,
                "recommendation": "Activate Inter-Frequency Load Balancing (IFLB) and audit ARP pre-emption flags.",
                "scenario": "B_CCE",
            })
        elif mme_drop_share > 0.30:
            findings.append({
                "severity": sev,
                "domain": "DCR",
                "category": "Transport / Core (S1/MME)",
                "cause": f"Core Network MME Initiated Session Drop ({mme_drop_share*100:.0f}% MME Releases)",
                "confidence": 88,
                "recommendation": "Inspect MME S1AP release cause codes and SGW GTP-U path supervision timers.",
                "scenario": "CORE_MME",
            })
        else:
            findings.append({
                "severity": sev,
                "domain": "DCR",
                "category": "RF Weak Coverage / UL Power",
                "cause": "Composite E-RAB Abnormal Release (RF Fading & Signaling Timeout)",
                "confidence": 82,
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
                    "recommendation": "Expand Connected Users License on target neighbor eNodeBs.",
                    "scenario": "B_LIC",
                })
            elif ho_ul_rej > 0:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "UL Interference / PIM",
                    "cause": "Target Cell Bad Uplink Quality Rejection",
                    "confidence": 91,
                    "recommendation": "Clear target cell UL interference and audit neighbor UL RSSI.",
                    "scenario": "A_PIM",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "PRB / CCE / GBR Congestion",
                    "cause": "Target Cell Capacity / CCE / PRB Admission Congestion",
                    "confidence": 88,
                    "recommendation": "Enable Inter-Frequency Load Balancing (IFLB) to shift load prior to handover.",
                    "scenario": "C_HO",
                })
        else:
            if abn_ho > 50:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": f"Physical Layer HO Execution Collapse ({hosr:.1f}% HOSR, PCI/Edge Interference)",
                    "confidence": 93,
                    "recommendation": "Reduce a3Offset (e.g., 30 -> 20) for earlier HO trigger; check PCI mod-3 collision.",
                    "scenario": "C_HO",
                })
            else:
                findings.append({
                    "severity": sev,
                    "domain": "HOSR",
                    "category": "Event B2/A2 & Mobility Tuning",
                    "cause": f"Neighbor Handover Execution Degradation ({hosr:.1f}% HOSR / Blocked Relation)",
                    "confidence": 89,
                    "recommendation": "Audit AnrFunction.anrState, X2 neighbor definitions, and EUtranCellRelation.hoAllowed.",
                    "scenario": "C_HO",
                })

    # ------------------------------------------------------------------
    # TREE 4.4: LOW CQI / RADIO QUALITY DEGRADATION
    # ------------------------------------------------------------------
    if valid_cqi and (avg_cqi < 7.5 or low_cqi > 30.0 or ul_intf > -108.0):
        sev = 3 if (avg_cqi < 6.0 or low_cqi > 50.0 or ul_intf > -105.0) else 2
        if ul_intf > -108.0:
            findings.append({
                "severity": sev,
                "domain": "CQI / RF",
                "category": "UL Interference / PIM",
                "cause": f"High Radio Noise Floor / PIM Interference ({ul_intf:.1f} dBm/PRB)",
                "confidence": 95,
                "recommendation": "Execute PRB-level interference scan; inspect TMA/RRU connectors and RET hardware.",
                "scenario": "A_PIM",
            })
        elif rank1_pct > 75.0:
            findings.append({
                "severity": sev,
                "domain": "CQI / RF",
                "category": "Hardware / Baseband / RRU-RET",
                "cause": f"Lack of Dominant Server / RRU MIMO Branch Fault (Rank-1 {rank1_pct:.1f}%)",
                "confidence": 90,
                "recommendation": "Inspect RRU antenna port VSWR, RET tilt settings, and co-channel pilot pollution.",
                "scenario": "A_COV_RET",
            })
        else:
            findings.append({
                "severity": sev,
                "domain": "CQI / RF",
                "category": "RF Weak Coverage / UL Power",
                "cause": f"Low DL SINR / High Low-CQI Tail (Avg CQI {avg_cqi:.1f}, Low-CQI {low_cqi:.1f}%)",
                "confidence": 87,
                "recommendation": "Audit optical SFP/feeder loss, RET tilt, and enable Ericsson Lean Carrier.",
                "scenario": "A_COV_RET",
            })

    # ------------------------------------------------------------------
    # TREE 4.5: 3G/4G IRAT PING-PONG LOOP
    # ------------------------------------------------------------------
    rrc_succ = float(row.get("pmRrcConnEstabSucc", 0.0))
    irat_rel_rate = (irat_denom / rrc_succ) if rrc_succ > 0 else 0.0
    if valid_irat and (irat_pp > 5.0 or irat_rel_rate > 0.25):
        sev = 3 if irat_pp > 10.0 else 2
        bad_cov_rep = float(row.get("pmBadCovEvalReport", 0.0))
        csfb_rel = float(row.get("pmUeCtxtRelCsfbWcdma", 0.0))
        sc_wcdma = float(row.get("pmUeCtxtRelSCWcdma", 0.0))
        b2_t1 = float(row.get("b2Threshold1Rsdp", -108.0))
        b2_t2 = float(row.get("b2Threshold2RsdpUtran", -98.0))

        if rrc_succ > 0 and (bad_cov_rep / rrc_succ) > 0.40 and b2_t1 > -110.0:
            cause_txt = "Aggressive Event A2 / B2 Threshold Triggering (4G Threshold Too High)"
        elif b2_t2 < -100.0:
            cause_txt = "Loose 3G Target Quality Threshold (Event B2 Threshold 2 Too Low)"
        elif sc_wcdma > 0 and csfb_rel > 0.50 * sc_wcdma:
            cause_txt = "CSFB Voice Return Oscillation (Post-Call Fast Return Ping-Pong)"
        else:
            cause_txt = f"3G/4G IRAT Ping-Pong Loop ({irat_pp:.1f}%) / Reselection Priority Mismatch"

        findings.append({
            "severity": sev,
            "domain": "IRAT",
            "category": "Event B2/A2 & Mobility Tuning",
            "cause": cause_txt,
            "confidence": 94,
            "recommendation": "Lower a2Threshold0Rsdp (-115 dBm) & b2Threshold1Rsdp (-116 dBm); tighten b2Threshold2RsdpUtran (-95 dBm).",
            "scenario": "D_IRAT",
        })

    # ------------------------------------------------------------------
    # TREES 4.6 & 4.7: VoLTE QCI=1 CSSR & DCR
    # ------------------------------------------------------------------
    if valid_vcssr and volte_cssr < 97.5:
        sev = 3 if volte_cssr < 95.0 else 2
        findings.append({
            "severity": sev,
            "domain": "VoLTE CSSR",
            "category": "PRB / CCE / GBR Congestion",
            "cause": f"VoLTE QCI=1 Bearer Setup Degradation ({volte_cssr:.2f}%)",
            "confidence": 92,
            "recommendation": "Expand qci1GbrDlLimit/qci1GbrUlLimit and set QciProfilePredefined=qci1 to PRIORITY_BASED.",
            "scenario": "E_VOLTE",
        })

    if valid_vdcr and volte_dcr > 0.8:
        sev = 3 if volte_dcr > 2.0 else 2
        findings.append({
            "severity": sev,
            "domain": "VoLTE DCR",
            "category": "Event B2/A2 & Mobility Tuning",
            "cause": f"VoLTE QCI=1 Voice Drop Spike ({volte_dcr:.2f}%)",
            "confidence": 91,
            "recommendation": "Enable ROHC & TTI Bundling on EUtranCellFDD and verify SRVCC neighbor relation.",
            "scenario": "E_VOLTE",
        })

    # ------------------------------------------------------------------
    # CONSOLIDATE PRIMARY ROOT CAUSE, STATUS & AMOS REMEDIATION SCRIPT
    # ------------------------------------------------------------------
    evidence_items = [
        {"Counter / Indicator": "pmRrcConnEstabAtt / Succ", "Value": f"{int(row.get('pmRrcConnEstabAtt', 0)):,} / {int(row.get('pmRrcConnEstabSucc', 0)):,}", "Meaning": f"RRC SR: {rrc_sr:.2f}%"},
        {"Counter / Indicator": "pmRaFailCbraMsg1DiscOoc", "Value": f"{int(ra_ooc):,} ({ooc_ratio:.2f}%)", "Meaning": "RACH Out-of-Coverage Preambles"},
        {"Counter / Indicator": "pmRaFailCbraMsg1DiscSched", "Value": f"{int(ra_sched):,}", "Meaning": "Baseband Scheduler Msg1 Discards"},
        {"Counter / Indicator": "pmErabRelAbnormalEnbAct", "Value": f"{int(abn_enb):,} / {int(dcr_denom):,} rel", "Meaning": f"eNB DCR: {dcr:.2f}% ({dcr_norm:.1f}/10k hr)"},
        {"Counter / Indicator": "pmErabRelAbnormalEnbActUeLost", "Value": f"{int(abn_ue_lost):,} ({(abn_ue_lost/abn_enb*100 if abn_enb>0 else 0):.1f}%)", "Meaning": "RF Link / RLC Unack UE Lost"},
        {"Counter / Indicator": "pmErabRelAbnormalEnbActHo", "Value": f"{int(abn_ho):,}", "Meaning": "Abnormal Drops During Handover"},
        {"Counter / Indicator": "pmErabRelAbnormalMmeAct", "Value": f"{int(abn_mme):,}", "Meaning": "Core MME Triggered Abnormal Releases"},
        {"Counter / Indicator": "pmRadioRecInterferencePwr", "Value": f"{ul_intf:.2f} dBm/PRB", "Meaning": "UL Noise Floor (> -110 dBm = Degraded)"},
        {"Counter / Indicator": "PRB Interference Crest", "Value": f"+{prb_crest:.1f} dB ({peak_prbs_txt})", "Meaning": "PIM / Narrowband Spectral Imbalance"},
        {"Counter / Indicator": "CQI & MIMO Rank-1 Profile", "Value": f"CQI {avg_cqi:.2f} | Low {low_cqi:.1f}% | R1 {rank1_pct:.1f}%", "Meaning": "DL Radio SINR & Spatial Rank"},
    ]

    if not findings:
        # Check if excluded due to low traffic sample
        if low_sample_notes or (not valid_cssr and not valid_dcr):
            excl_msg = low_sample_notes[0] if low_sample_notes else (
                f"Cell {cell} excluded from degradation list due to insufficient sample size."
            )
            return {
                "Status": "LOW_SAMPLE",
                "Severity_Rank": 0,
                "RCA_Category": "Low Traffic Sample (Excluded)",
                "Primary_Root_Cause": excl_msg,
                "Secondary_Causes": "None (Statistical Sample Gate Filtered)",
                "Affected_Domain": "LOW_VOL",
                "Confidence_Score": 0,
                "Recommendation": "No action required — monitor across wider ROP aggregation window.",
                "AMOS_Script": f"# Cell {cell} ({enb}) skipped: LOW_TRAFFIC_SAMPLE (Minimum Denominator Gate).",
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

    # Sort findings by severity desc, then confidence desc
    findings.sort(key=lambda x: (x["severity"], x["confidence"]), reverse=True)
    primary = findings[0]
    secondary_list = [f"{f['domain']}: {f['cause']}" for f in findings[1:]]
    status_str = "CRITICAL" if primary["severity"] >= 3 else "DEGRADED"

    # Collect unique remediation scenarios triggered on this cell
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
    """Build copy-pasteable Ericsson AMOS / Moshell CLI script tailored to the cell's RCA."""
    lines = [
        f"# ===================================================================",
        f"# ERICSSON AMOS / MOSHELL REMEDIATION PROTOCOL",
        f"# Target Node : {enodeb}   |   Cell : {cell}",
        f"# Diagnosis   : {root_cause}",
        f"# ===================================================================",
        f"moshell {enodeb}",
        f"lt all",
        f"confb+",
        f"gs+",
        f"",
        f"# --- Step 1: Health, Alarm & Operational State Verification ---",
        f"st EUtranCellFDD={cell}",
        f"alt",
    ]

    for sc in scenarios:
        if sc == "A_PIM":
            lines.extend([
                "",
                "# --- UL Interference & PIM Spectrum Audit (Scenario A) ---",
                f"pget EUtranCellFDD={cell} pmRadioRecInterferencePwr$",
                f"pget EUtranCellFDD={cell} pmRadioRecInterferencePwrPrb",
                "cabx",
                "lhsh BXP_0 vsi",
                "set ENodeBFunction=1,EricssonLeanCarrier=1 leanCarrierEnabled true",
            ])
        elif sc == "A_COV_RET":
            lines.extend([
                "",
                "# --- Hardware RRU/RET Tilt & DL Coverage Audit (Scenario A) ---",
                "get RetSubUnit electricalAntennaTilt|iuantAntennaModelNumber|operationalState",
                "hget RfPort vswrSupervisionSensitivity|vswrValue",
                f"pget EUtranCellFDD={cell} pmRadioUeRepCqiDistr|pmRadioUeRepRankDistr",
                f"get EUtranCellFDD={cell} crsGain|pdschTypeBGain",
                "set ENodeBFunction=1,EricssonLeanCarrier=1 leanCarrierEnabled true",
            ])
        elif sc == "B_OOC":
            lines.extend([
                "",
                "# --- RACH Out-of-Coverage / Overshooting Remediation (Scenario B) ---",
                f"pget EUtranCellFDD={cell} pmRaAttCbra|pmRaFailCbraMsg1DiscOoc",
                f"get EUtranCellFDD={cell} cellRange|rachRootSequence",
                "get RetSubUnit electricalAntennaTilt",
                f"set EUtranCellFDD={cell} cellRange 15",
            ])
        elif sc == "B_LIC":
            lines.extend([
                "",
                "# --- Connected Users / E-RAB License Audit (Scenario B) ---",
                f"pget EUtranCellFDD={cell} pmRrcConnEstabFailLic|pmErabEstabFailInitLic",
                "invl",
                "get CapacityState grantedCapacityLevel|operationalCapacityLevel",
                "set LoadBalancingFunction=1 lbCeiling 80",
            ])
        elif sc == "B_HW":
            lines.extend([
                "",
                "# --- Baseband MP/DU CPU Overload Mitigation (Scenario B) ---",
                "cabx",
                f"pget EUtranCellFDD={cell} pmRrcConnEstabFailMpOverload|pmRrcConnEstabFailDuIntens",
                f"set EUtranCellFDD={cell} pdcchCceOptimizationEnabled true",
            ])
        elif sc == "B_CCE":
            lines.extend([
                "",
                "# --- PDCCH / CCE & Scheduler Optimization (Scenario B) ---",
                f"pget EUtranCellFDD={cell} pmRaFailCbraMsg1DiscSched|pmRaFailCbraMsg2Disc",
                f"set EUtranCellFDD={cell} pdcchCceOptimizationEnabled true",
                f"set EUtranCellFDD={cell} pdcchPowerBoostMax 3",
            ])
        elif sc == "C_ULPC":
            lines.extend([
                "",
                "# --- Retainability & Uplink Power Control Optimization (Scenario C) ---",
                f"pget EUtranCellFDD={cell} pmErabRelAbnormalEnbAct|pmErabRelAbnormalEnbActUeLost",
                f"get EUtranCellFDD={cell} p0NominalPusch|alpha|t310|n310",
                f"set EUtranCellFDD={cell} p0NominalPusch -100",
                f"set EUtranCellFDD={cell} alpha 8",
                f"set EUtranCellFDD={cell} ulSrsEnable true",
            ])
        elif sc == "C_HO":
            lines.extend([
                "",
                "# --- Handover Execution & Neighbor Relation Tuning (Scenario C) ---",
                f"pget EUtranCellFDD={cell} pmErabRelAbnormalEnbActHo|pmCellHoExeAttLteIntraF",
                f"get EUtranCellFDD={cell},EUtranCellRelation= isHoAllowed|cellIndividualOffsetEUtran",
                "get AnrFunction anrState",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigEUtraBestCell=1 a3Offset 20",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigEUtraBestCell=1 timeToTriggerA3 80",
            ])
        elif sc == "D_IRAT":
            lines.extend([
                "",
                "# --- 3G/4G IRAT Event A2/B2 Ping-Pong Mitigation (Scenario D) ---",
                f"get EUtranCellFDD={cell},UeMeasControl=1 a2Threshold0Rsdp",
                f"get EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold1Rsdp|b2Threshold2RsdpUtran|timeToTriggerB2|hysteresisB2",
                f"set EUtranCellFDD={cell},UeMeasControl=1 a2Threshold0Rsdp -115",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold1Rsdp -116",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold2RsdpUtran -95",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2Utra=1 timeToTriggerB2 640",
                f"set EUtranCellFDD={cell},UeMeasControl=1,ReportConfigB2Utra=1 hysteresisB2 30",
                f"set EUtranCellFDD={cell},UtranFreqRelation= cellReselectionPriority 3",
            ])
        elif sc == "E_VOLTE":
            lines.extend([
                "",
                "# --- VoLTE QCI=1 CAC, ROHC & TTI Bundling Protocol (Scenario E) ---",
                "get QciProfilePredefined=qci1 preEmption|schedulingPolicy",
                "set QciProfilePredefined=qci1 preEmptionCapability NOT_PREEMPTABLE",
                "set QciProfilePredefined=qci1 preEmptionVulnerability DISABLED",
                "set QciProfilePredefined=qci1 schedulingPolicy PRIORITY_BASED",
                f"set EUtranCellFDD={cell} qci1GbrDlLimit 20",
                f"set EUtranCellFDD={cell} qci1GbrUlLimit 20",
                f"set EUtranCellFDD={cell} rohcEnabled true",
                f"set EUtranCellFDD={cell} ttiBundlingEnabled true",
            ])
        elif sc in ("TN_S1", "CORE_MME"):
            lines.extend([
                "",
                "# --- S1 Transport & MME Core Supervision Check ---",
                "st TermPointToMme",
                "get Sctp errorCounter|assocMaxRtx",
                "pmr -m 1 | grep -iE 'S1|TnFail|Mme'",
            ])

    lines.extend([
        "",
        "confb-",
        "gs-",
        "cvms Post_Antigravity_RCA_Optimization",
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
    vectorized KPI aggregations, time-series synthesis, and full RCA decision trees.
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
    prb_frames: List[pd.DataFrame] = []

    if suffix == ".csv":
        try:
            text_data = raw_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            text_data = raw_bytes.decode("latin1")

        # Locate header row using csv.reader so ragged metadata title rows never trigger ParserError
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
            elif kind == "prb_spectrum":
                prb_frames.append(norm_df)
                sheets_loaded.append(s_name)

    if not kpi_frames and not cqi_long_frames:
        raise ValueError(
            "Không tìm thấy bảng counter Ericsson hợp lệ trong tệp tải lên. "
            "Vui lòng kiểm tra định dạng cột (EUtranCellFDD / ERBS Id / pm*)."
        )

    # Combine main KPI tables
    if kpi_frames:
        df_main = pd.concat(kpi_frames, ignore_index=True)
    else:
        # Synthesize cell scaffold if only CQI sheet was uploaded
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

    # Detect optional IRAT & VoLTE counter presence before zero-filling missing columns
    existing_cols = set(df_main.columns)
    has_irat = bool(
        existing_cols.intersection(
            {"pmUeCtxtRelSCWcdma", "UtranCellRelation.pmHoExeSucc", "pmRapid3gTo4gReentries", "pmB2BestCellEvalReport"}
        )
    )
    has_volte = bool(
        existing_cols.intersection({"pmErabEstabAttQci[1]", "pmErabRelAbnormalEnbActQci[1]"})
    )

    present_counters = [c for c in SCALAR_COUNTERS if c in df_main.columns]
    for c in present_counters:
        df_main[c] = pd.to_numeric(df_main[c], errors="coerce").fillna(0.0)

    # Check if wide-format CQI columns exist directly in df_main
    wide_cqi_cols = [f"pmRadioUeRepCqiDistr[{i}]" for i in range(16)]
    has_wide_cqi = all(c in df_main.columns for c in wide_cqi_cols)
    if has_wide_cqi:
        for c in wide_cqi_cols:
            df_main[c] = pd.to_numeric(df_main[c], errors="coerce").fillna(0.0)

    # Check if scalar UL interference (in dBm) exists directly in df_main
    has_scalar_intf = "pmRadioRecInterferencePwr" in df_main.columns and not cqi_long_frames
    if has_scalar_intf:
        df_main["pmRadioRecInterferencePwr"] = pd.to_numeric(
            df_main["pmRadioRecInterferencePwr"], errors="coerce"
        )

    # --- Aggregate scalar counters per Cell ---
    agg_dict: Dict[str, Any] = {c: "sum" for c in present_counters}
    for threshold_col in ("b2Threshold1Rsdp", "b2Threshold2RsdpUtran", "b2Threshold2EcNoUtran", "timeToTriggerB2", "hysteresisB2"):
        if threshold_col in agg_dict:
            agg_dict[threshold_col] = "mean"
    if has_wide_cqi:
        for c in wide_cqi_cols:
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

    # Fill missing scalar counters with 0.0 in one batch to avoid DataFrame fragmentation
    missing_scalars = {c: 0.0 for c in SCALAR_COUNTERS if c not in cell_df.columns}
    if missing_scalars:
        cell_df = pd.concat(
            [cell_df, pd.DataFrame(missing_scalars, index=cell_df.index)], axis=1
        )

    # --- Compute Composite KPIs per Cell (Section 2.1 - 2.7) ---
    rrc_denom = np.maximum(cell_df["pmRrcConnEstabAtt"] - cell_df["pmRrcConnEstabAttReatt"], 0.0)
    rrc_sr = np.where(rrc_denom > 0, np.clip(cell_df["pmRrcConnEstabSucc"] / rrc_denom, 0.0, 1.0), 1.0)
    s1_sr = np.where(
        cell_df["pmS1SigConnEstabAtt"] > 0,
        np.clip(cell_df["pmS1SigConnEstabSucc"] / cell_df["pmS1SigConnEstabAtt"], 0.0, 1.0),
        1.0,
    )
    erab_sr = np.where(
        cell_df["pmErabEstabAttInit"] > 0,
        np.clip(cell_df["pmErabEstabSuccInit"] / cell_df["pmErabEstabAttInit"], 0.0, 1.0),
        1.0,
    )

    cell_df["CSSR_Denom"] = rrc_denom
    cell_df["RRC_SR"] = rrc_sr * 100.0
    cell_df["S1_SR"] = s1_sr * 100.0
    cell_df["ERAB_SR"] = erab_sr * 100.0
    cell_df["CSSR"] = rrc_sr * s1_sr * erab_sr * 100.0

    dcr_denom = cell_df["pmErabRelAbnormalEnbAct"] + cell_df["pmErabRelNormalEnbAct"]
    cell_df["DCR_Denom"] = dcr_denom
    cell_df["DCR"] = np.where(
        dcr_denom > 0, cell_df["pmErabRelAbnormalEnbAct"] / dcr_denom * 100.0, 0.0
    )
    cell_df["DCR_Norm"] = np.where(
        cell_df["pmSessionTimeUe"] > 0,
        (cell_df["pmErabRelAbnormalEnbAct"] + cell_df["pmErabRelAbnormalMmeAct"])
        / cell_df["pmSessionTimeUe"]
        * 36_000_000.0,
        0.0,
    )

    ho_prep_att = cell_df["pmCellHoPrepAttLteIntraF"] + cell_df["pmCellHoPrepAttLteInterF"]
    ho_prep_succ = cell_df["pmCellHoPrepSuccLteIntraF"] + cell_df["pmCellHoPrepSuccLteInterF"]
    ho_exe_att = cell_df["pmCellHoExeAttLteIntraF"] + cell_df["pmCellHoExeAttLteInterF"]
    ho_exe_succ = cell_df["pmCellHoExeSuccLteIntraF"] + cell_df["pmCellHoExeSuccLteInterF"]

    ho_prep_rate = np.where(ho_prep_att > 0, np.clip(ho_prep_succ / ho_prep_att, 0.0, 1.0), 1.0)
    ho_exe_rate = np.where(ho_exe_att > 0, np.clip(ho_exe_succ / ho_exe_att, 0.0, 1.0), 1.0)
    cell_df["HO_Denom"] = ho_prep_att
    cell_df["HO_Prep_SR"] = ho_prep_rate * 100.0
    cell_df["HO_Exe_SR"] = ho_exe_rate * 100.0
    cell_df["HOSR"] = ho_prep_rate * ho_exe_rate * 100.0

    cell_df["RACH_OOC_Ratio"] = np.where(
        cell_df["pmRaAttCbra"] > 0,
        cell_df["pmRaFailCbraMsg1DiscOoc"] / cell_df["pmRaAttCbra"] * 100.0,
        0.0,
    )

    # IRAT Ping-Pong Rate
    irat_denom = cell_df["pmUeCtxtRelSCWcdma"] + cell_df["UtranCellRelation.pmHoExeSucc"]
    cell_df["IRAT_Denom"] = irat_denom
    if "pmRapid3gTo4gReentries" in present_counters:
        cell_df["IRAT_PingPong_Rate"] = np.where(
            irat_denom > 0, cell_df["pmRapid3gTo4gReentries"] / irat_denom * 100.0, 0.0
        )
    elif has_irat:
        # Estimate oscillation pressure from B2/CSFB/SCWcdma ratio if explicit reentry counter absent
        cell_df["IRAT_PingPong_Rate"] = np.where(
            cell_df["pmRrcConnEstabSucc"] > 0,
            irat_denom / cell_df["pmRrcConnEstabSucc"] * 100.0,
            0.0,
        )
    else:
        cell_df["IRAT_PingPong_Rate"] = 0.0

    # VoLTE CSSR & DCR
    cell_df["VoLTE_CSSR_Denom"] = cell_df["pmErabEstabAttQci[1]"]
    cell_df["VoLTE_CSSR"] = np.where(
        cell_df["pmErabEstabAttQci[1]"] > 0,
        cell_df["pmErabEstabSuccQci[1]"] / cell_df["pmErabEstabAttQci[1]"] * 100.0,
        100.0,
    )
    vdcr_denom = cell_df["pmErabRelAbnormalEnbActQci[1]"] + cell_df["pmErabRelNormalEnbActQci[1]"]
    cell_df["VoLTE_DCR_Denom"] = vdcr_denom
    cell_df["VoLTE_DCR"] = np.where(
        vdcr_denom > 0, cell_df["pmErabRelAbnormalEnbActQci[1]"] / vdcr_denom * 100.0, 0.0
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
        cell_df["CQI_Samples"] = cqi_tot
        cell_df["Avg_CQI"] = np.where(
            cqi_tot > 0, (cqi_mat @ np.arange(16, dtype=np.float64)) / cqi_tot, np.nan
        )
        cell_df["Low_CQI_Ratio"] = np.where(
            cqi_tot > 0, cqi_mat[:, :7].sum(axis=1) / cqi_tot * 100.0, 0.0
        )
        for b in range(16):
            cell_df[f"CQI_Bin_{b}"] = cqi_mat[:, b]
        cell_df["Rank1_Ratio"] = 40.0
        cell_df["UL_Intf_dBm"] = (
            cell_df["pmRadioRecInterferencePwr"] if has_scalar_intf else -118.0
        )
        cell_df["Intf_Samples"] = 0.0
    else:
        fallback_radio = {"CQI_Samples": 0.0, "Avg_CQI": np.nan, "Low_CQI_Ratio": 0.0}
        for b in range(16):
            fallback_radio[f"CQI_Bin_{b}"] = 0.0
        fallback_radio["Rank1_Ratio"] = np.nan
        fallback_radio["UL_Intf_dBm"] = (
            cell_df["pmRadioRecInterferencePwr"] if has_scalar_intf else -118.0
        )
        fallback_radio["Intf_Samples"] = 0.0
        cell_df = pd.concat(
            [cell_df.copy(), pd.DataFrame(fallback_radio, index=cell_df.index)], axis=1
        )

    # CQI Modulation Band Shares per cell (for Stacked Bar Visualization)
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

    # Sort Master Grid: CRITICAL -> DEGRADED -> NORMAL -> LOW_SAMPLE, then DCR desc, CSSR asc
    cell_df = cell_df.sort_values(
        by=["Severity_Rank", "DCR", "CSSR"],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    # --- Build ROP Time-Series Trend DataFrame ---
    rop_num_cols = [
        c
        for c in (
            "pmRrcConnEstabAtt", "pmRrcConnEstabAttReatt", "pmRrcConnEstabSucc",
            "pmS1SigConnEstabAtt", "pmS1SigConnEstabSucc",
            "pmErabEstabAttInit", "pmErabEstabSuccInit",
            "pmErabRelNormalEnbAct", "pmErabRelAbnormalEnbAct",
            "pmErabRelAbnormalMmeAct", "pmSessionTimeUe",
            "pmCellHoPrepAttLteIntraF", "pmCellHoPrepSuccLteIntraF",
            "pmCellHoExeAttLteIntraF", "pmCellHoExeSuccLteIntraF",
            "pmCellHoPrepAttLteInterF", "pmCellHoPrepSuccLteInterF",
            "pmCellHoExeAttLteInterF", "pmCellHoExeSuccLteInterF",
        )
        if c in df_main.columns
    ]

    if rop_num_cols:
        rop_grp = (
            df_main.groupby(["ROP_Time", "EUtranCellFDD"], observed=False, as_index=False)[rop_num_cols]
            .sum()
        )
        for col_req in (
            "pmRrcConnEstabAtt", "pmRrcConnEstabAttReatt", "pmRrcConnEstabSucc",
            "pmS1SigConnEstabAtt", "pmS1SigConnEstabSucc",
            "pmErabEstabAttInit", "pmErabEstabSuccInit",
            "pmErabRelNormalEnbAct", "pmErabRelAbnormalEnbAct",
            "pmErabRelAbnormalMmeAct", "pmSessionTimeUe",
        ):
            if col_req not in rop_grp.columns:
                rop_grp[col_req] = 0.0

        r_denom = np.maximum(rop_grp["pmRrcConnEstabAtt"] - rop_grp["pmRrcConnEstabAttReatt"], 0.0)
        r_sr = np.where(r_denom > 0, np.clip(rop_grp["pmRrcConnEstabSucc"] / r_denom, 0.0, 1.0), 1.0)
        s_sr = np.where(
            rop_grp["pmS1SigConnEstabAtt"] > 0,
            np.clip(rop_grp["pmS1SigConnEstabSucc"] / rop_grp["pmS1SigConnEstabAtt"], 0.0, 1.0),
            1.0,
        )
        e_sr = np.where(
            rop_grp["pmErabEstabAttInit"] > 0,
            np.clip(rop_grp["pmErabEstabSuccInit"] / rop_grp["pmErabEstabAttInit"], 0.0, 1.0),
            1.0,
        )
        rop_grp["CSSR"] = r_sr * s_sr * e_sr * 100.0
        d_denom = rop_grp["pmErabRelAbnormalEnbAct"] + rop_grp["pmErabRelNormalEnbAct"]
        rop_grp["DCR"] = np.where(
            d_denom > 0, rop_grp["pmErabRelAbnormalEnbAct"] / d_denom * 100.0, 0.0
        )
        rop_grp["DCR_Norm"] = np.where(
            rop_grp["pmSessionTimeUe"] > 0,
            (rop_grp["pmErabRelAbnormalEnbAct"] + rop_grp["pmErabRelAbnormalMmeAct"])
            / rop_grp["pmSessionTimeUe"]
            * 36_000_000.0,
            0.0,
        )
        rop_trend_df = rop_grp.sort_values("ROP_Time").reset_index(drop=True)
    else:
        rop_trend_df = pd.DataFrame(columns=["ROP_Time", "EUtranCellFDD", "CSSR", "DCR", "DCR_Norm"])

    if not rop_radio_df.empty and not rop_trend_df.empty:
        rop_trend_df = rop_trend_df.merge(rop_radio_df, on="ROP_Time", how="left")

    # Build CQI bin distribution table (0..15) per cell for flexible filtering
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
    )


def compute_rop_series(rop_trend_df: pd.DataFrame, cells: Optional[List[str]] = None) -> pd.DataFrame:
    """Aggregate ROP time-series KPIs across filtered subset of cells."""
    if rop_trend_df.empty:
        return pd.DataFrame(columns=["ROP_Time", "CSSR", "DCR", "DCR_Norm"])

    sub = rop_trend_df[rop_trend_df["EUtranCellFDD"].isin(cells)] if cells is not None else rop_trend_df
    if sub.empty:
        return pd.DataFrame(columns=["ROP_Time", "CSSR", "DCR", "DCR_Norm"])

    sum_cols = [
        c
        for c in (
            "pmRrcConnEstabAtt", "pmRrcConnEstabAttReatt", "pmRrcConnEstabSucc",
            "pmS1SigConnEstabAtt", "pmS1SigConnEstabSucc",
            "pmErabEstabAttInit", "pmErabEstabSuccInit",
            "pmErabRelNormalEnbAct", "pmErabRelAbnormalEnbAct",
            "pmErabRelAbnormalMmeAct", "pmSessionTimeUe",
        )
        if c in sub.columns
    ]
    ts = sub.groupby("ROP_Time", as_index=False)[sum_cols].sum().sort_values("ROP_Time")

    r_den = np.maximum(ts["pmRrcConnEstabAtt"] - ts["pmRrcConnEstabAttReatt"], 0.0)
    r_sr = np.where(r_den > 0, np.clip(ts["pmRrcConnEstabSucc"] / r_den, 0.0, 1.0), 1.0)
    s_sr = np.where(
        ts["pmS1SigConnEstabAtt"] > 0,
        np.clip(ts["pmS1SigConnEstabSucc"] / ts["pmS1SigConnEstabAtt"], 0.0, 1.0),
        1.0,
    )
    e_sr = np.where(
        ts["pmErabEstabAttInit"] > 0,
        np.clip(ts["pmErabEstabSuccInit"] / ts["pmErabEstabAttInit"], 0.0, 1.0),
        1.0,
    )
    ts["CSSR"] = r_sr * s_sr * e_sr * 100.0
    d_den = ts["pmErabRelAbnormalEnbAct"] + ts["pmErabRelNormalEnbAct"]
    ts["DCR"] = np.where(d_den > 0, ts["pmErabRelAbnormalEnbAct"] / d_den * 100.0, 0.0)
    ts["DCR_Norm"] = np.where(
        ts["pmSessionTimeUe"] > 0,
        (ts["pmErabRelAbnormalEnbAct"] + ts["pmErabRelAbnormalMmeAct"])
        / ts["pmSessionTimeUe"]
        * 36_000_000.0,
        0.0,
    )
    return ts
