"""
Unified Quality & Client Rejection Analytics Suite
Combines Internal Quality Audit ('RE-Audit_Data'), External Client Rejections
('External Rejection  Zibber Dump' & 'Other Clients rejection Dump'), and
Designer Productivity & Spec-wise Data into one web app.

Tabs:
1. Quality Audit Report (Daily/Range summary, 3-day recurring streak alerts, editable takeaways, .docx/.pdf/.xlsx/.csv/JPG export)
2. Client Rejection Report (Daily/Range summary, 3-day recurring streak alerts, editable takeaways, .docx/.pdf/.xlsx/.csv/JPG export)
3. Client Attention Radar (Daily, Weekly & Monthly comparison with editable takeaways & exports)
4. Analytics & Charts (Interactive executive charts with editable takeaways & exports)
5. Period Compare (Side-by-side Period A vs Period B comparison with editable takeaways & exports)
6. Group Data (Designer Productivity, Active Time, Expected vs Total ETA, ATI & Rejection % by Team/Designation)
7. Spec-wise Data (Spec ID breakdown with Target Time, Designer ATI, Reject Files, Efficiency % & Top/Low Specs)

Usage:
    python quality_analytics_suite_v2_3.py

Needs: pandas, openpyxl, python-docx
"""
import base64
import io
import json
import os
import re
import secrets
import sys
import threading
import time
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pandas as pd

try:
    import docx
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Inches, Pt, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ===================================================================== SETTINGS
QA_SHEET = "RE-Audit_Data"
QA_ONLINE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1cjapBN7BMu8FcMZoaM80A5vd-U6ilyhcc-gytwtYAD8/export?format=csv&gid=807106360"
)
QA_CLIENTS = {
    67484: "HomeVisuals",
    67676: "Ogulo GmbH",
}
ESCALATE_ABOVE = 10.0
STRONG_BELOW = 10.0
STREAK_DAYS = 3
ACTION_LINES_ESCALATE = [
    "If this happens again for two consecutive days, we will escalate it to the Performance Team.",
    "If the issue is not too concerning, we will share the update on Saturday.",
]
ACTION_LINE_OK = "No escalation needed; all clients are within the acceptable defect rate."

CR_ZIBBER_SHEET = "External Rejection  Zibber Dump"
CR_OTHER_SHEET = "Other Clients rejection Dump"
CR_ONLINE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1LpVQf2apePI6Ehh1dWzSVNMJp2-dLosrSo3mI2OjHnQ/export?format=xlsx"
)
FIXED_CLIENTS = ["Ogulo-67676", "70023"]
ZIBBER_REASON_ALIASES = ["rejected reason", "reason 1", "reason"]
OTHER_REASON_ALIASES = ["reject categroy", "reject category", "rejection category", "category", "issues"]
JOB_COL_ALIASES = ["rejected job id", "original job id", "job id", "order id", "job_id", "order_id", "job", "order"]
DES_COL_ALIASES = ["designer name", "designer", "editor", "designer_name", "editor name", "done by"]
QC_COL_ALIASES = ["previous qc", "previous qc name", "qc name", "qc", "reviewer", "qa", "qa name", "qc_name", "checked by", "qc by", "audited by"]
CORRECTION_QC_ALIASES = ["correction qc", "correction qc name", "correction_qc"]

PROD_SHEET = "Data"
PROD_ONLINE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1tXAnsbvW-JmgN0BVEu4-1UCSVts74j28a-qsZAbnZGM/export?format=csv&gid=0"
)
DESIGNATION_TARGETS_SEC = {
    "Probationary Designer": 4 * 3600 + 30 * 60,   # 4:30:00
    "Junior Designer": 5 * 3600 + 15 * 60,         # 5:15:00
    "Designer": 5 * 3600 + 45 * 60,                # 5:45:00
    "Senior Designer": 6 * 3600 + 30 * 60,         # 6:30:00
}
DEFAULT_TARGET_SEC = 5 * 3600 + 15 * 60


def app_folder():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def norm(s):
    return str(s).strip().lower()


def find_col(df, candidates):
    for candidate in candidates:
        for col in df.columns:
            if norm(col) == candidate:
                return col
    return None


def pct(err, total):
    return (err / total * 100) if total else 0.0


def span_label(span):
    start, end = span
    return f"{start:%d/%m/%Y}" if start == end else f"{start:%d/%m/%Y} to {end:%d/%m/%Y}"


def sec_to_hms(sec):
    try:
        s = int(round(float(sec)))
    except (ValueError, TypeError):
        return "0:00:00"
    neg = s < 0
    s = abs(s)
    h = s // 3600
    m = (s % 3600) // 60
    rem = s % 60
    return f"{'-' if neg else ''}{h}:{m:02d}:{rem:02d}"


def parse_duration_sec(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        v = float(val)
        # If stored as Excel fraction of day (< 2.0 and float), convert if needed, else assume seconds if large
        if 0 < v < 1.0:
            return v * 86400.0
        return v
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "-", "n/a"):
        return 0.0
    neg = s.startswith("-")
    s_clean = s.lstrip("+-").strip()
    parts = s_clean.split(":")
    try:
        if len(parts) == 3:
            sec = float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
            return -sec if neg else sec
        if len(parts) == 2:
            sec = float(parts[0]) * 60 + float(parts[1])
            return -sec if neg else sec
        return (-1 if neg else 1) * float(s_clean.replace(",", ""))
    except ValueError:
        return 0.0


# ===================================================================== QA DATA
def qa_client_name(spec):
    try:
        spec_int = int(spec)
        return QA_CLIENTS.get(spec_int, f"Spec {spec_int}")
    except (ValueError, TypeError):
        return str(spec)


def clean_comment(text, categories):
    if isinstance(text, str) and text.strip():
        t = re.sub(r"\s+", " ", text.strip())
        return t[0].upper() + t[1:].rstrip(".") + "."
    if isinstance(categories, str) and categories.strip() and categories != "No Issues":
        return categories + "."
    return "Clean batch."


def _num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")


def _parse_qa_dates(df):
    col = df["Date"]
    if pd.api.types.is_datetime64_any_dtype(col):
        return pd.to_datetime(col).dt.normalize()
    t = col.astype(str).str.strip().str.split(" ").str[0]
    iso = pd.to_datetime(t, format="%Y-%m-%d", errors="coerce")
    mdy = pd.to_datetime(t, format="%m/%d/%Y", errors="coerce")
    dmy = pd.to_datetime(t, format="%d/%m/%Y", errors="coerce")
    label = df["Month"].astype(str).str.strip() if "Month" in df.columns else pd.Series("", index=df.index)
    mdy_ok = mdy.dt.strftime("%B %Y") == label
    dmy_ok = dmy.dt.strftime("%B %Y") == label
    out = iso.copy()
    out = out.fillna(mdy.where(mdy_ok))
    out = out.fillna(dmy.where(dmy_ok))
    out = out.fillna(mdy).fillna(dmy)
    return out.dt.normalize()


def prepare_qa(df):
    df = df.copy()
    df["Date"] = _parse_qa_dates(df)
    for c in ("Spec ID", "Order ID", "Total Image Count", "Total Image Count(Error)"):
        df[c] = _num(df[c]) if c in df.columns else float("nan")
    df = df[df["Total Image Count"].notna() & df["Date"].notna() & df["Spec ID"].notna()].copy()
    df["err"] = df["Total Image Count(Error)"].fillna(0).astype(int)
    df["imgs"] = df["Total Image Count"].astype(int)
    df["auditor"] = df["Audited by"].astype(str).str.strip().str.lower() if "Audited by" in df.columns else "unknown"
    df["_client_name"] = df["Spec ID"].apply(qa_client_name)
    return df


def load_qa_online(url):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        raw = r.read()
    head = raw[:300].lstrip().lower()
    if head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        raise RuntimeError("Google asked for sign-in on QA sheet.")
    return prepare_qa(pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, na_values=[""]))


# ===================================================================== PRODUCTIVITY (GROUP & SPEC-WISE) DATA
def format_sub_spec(spec_val):
    """Normalize sub-spec padding to 2 digits (e.g., '60345_001' or '60345_1' -> '60345_01')."""
    s = str(spec_val or "").strip()
    if not s or s.lower() in ("nan", "none", "unknown spec"):
        return "Unknown Spec"
    m = re.match(r"^([^_]+)_0*(\d+)$", s)
    if m:
        base, sub = m.groups()
        return f"{base.strip()}_{int(sub):02d}"
    return s


def extract_base_spec(spec_val):
    """Extract base Spec ID before '_' (e.g., '60345_01', '60345_02', '60345_001' -> '60345')."""
    s = str(spec_val or "").strip()
    if not s or s.lower() in ("nan", "none", "unknown spec"):
        return "Unknown Spec"
    m = re.match(r"^([^_]+)", s)
    return m.group(1).strip() if m else s


def get_sorted_designer_group(df, start, end, spec_filter="ALL"):
    """
    Filters Date between start & end and Spec matching spec_filter (e.g. '60345_01', '60345_02', or base '60345'),
    returning a sorted list of unique Designer Names.
    """
    sub = df[(df["_date"] >= start) & (df["_date"] <= end)].copy()
    if spec_filter and spec_filter != "ALL":
        target = format_sub_spec(spec_filter)
        if "_" in target:
            sub = sub[sub["_spec"] == target]
        else:
            prefix = f"{target}_"
            sub = sub[
                (sub["_base_spec"] == target)
                | (sub["_spec"].astype(str).str.strip().str.startswith(prefix))
            ]
    designers = sub["_designer"].dropna().astype(str).str.strip()
    designers = designers[~designers.str.lower().isin(["", "nan", "none", "unknown"])]
    return sorted(designers.unique().tolist())


def prepare_prod(df):
    df = df.copy()
    col_map = {norm(c): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n in col_map:
                return col_map[n]
        return None

    c_date = pick("date")
    c_spec = pick("spec id", "spec_id", "specid", "spec")
    c_des = pick("designername", "designer name", "designer_name", "designer")
    c_team = pick("teamname", "team name", "team_name", "team")
    c_desig = pick("designation", "role", "grade", "level")
    c_files = pick("filecount", "file count", "file_count", "image count", "images")
    c_qc_rej = pick("qcrejection", "qc rejection", "qc_rejection", "reject file", "reject files")
    c_cust_rej = pick("customerrejectioncount", "customer rejection count", "customer_rejection_count", "client rejection")
    c_tot_time = pick("totaltime", "total time", "total_time", "total eta")
    c_spec_time = pick("spectime", "spec time", "spec_time", "expected eta")
    c_eff = pick("efficiency", "efficiency %", "eff")
    c_act_time = pick("actual_time", "actual time", "actualtime")
    c_active = pick("active_time", "active time", "activetime", "total active time")
    c_passive = pick("passive_time", "passive time", "passivetime")
    c_pause = pick("pause_time", "pause time", "pausetime")
    c_ati = pick("ati", "designer ati")

    if not c_date:
        raise RuntimeError("Date column not found in Productivity sheet.")

    df["_date"] = pd.to_datetime(df[c_date], errors="coerce").dt.normalize()
    df = df[df["_date"].notna()].copy()

    def s_col(c, default=""):
        if not c:
            return pd.Series(default, index=df.index)
        s = df[c].fillna(default).astype(str).str.strip()
        return s.where(~s.str.lower().isin(["", "nan", "none"]), default)

    df["_spec"] = s_col(c_spec, "Unknown Spec").apply(format_sub_spec)
    df["_base_spec"] = df["_spec"].apply(extract_base_spec)
    df["_designer"] = s_col(c_des, "Unknown").str.replace(r"_\d+$", "", regex=True).str.strip()
    df["_team"] = s_col(c_team, "Unassigned Team")
    df["_designation"] = s_col(c_desig, "Designer")
    df = df[~df["_designer"].str.lower().isin(["unknown", ""])].copy()

    df["_files"] = _num(df[c_files]).fillna(0) if c_files else 0.0
    df["_qc_rej"] = _num(df[c_qc_rej]).fillna(0) if c_qc_rej else 0.0
    df["_cust_rej"] = _num(df[c_cust_rej]).fillna(0) if c_cust_rej else 0.0
    df["_rej"] = df["_qc_rej"]

    df["_active_sec"] = df[c_active].apply(parse_duration_sec) if c_active else 0.0
    df["_actual_sec"] = df[c_act_time].apply(parse_duration_sec) if c_act_time else df["_active_sec"]
    df["_tot_time_sec"] = df[c_tot_time].apply(parse_duration_sec) if c_tot_time else df["_active_sec"]
    df["_spec_time_sec"] = df[c_spec_time].apply(parse_duration_sec) if c_spec_time else 0.0
    df["_passive_sec"] = df[c_passive].apply(parse_duration_sec) if c_passive else 0.0
    df["_pause_sec"] = df[c_pause].apply(parse_duration_sec) if c_pause else 0.0

    # Fallback active time if zero
    zero_act = df["_active_sec"] <= 0
    if zero_act.any():
        df.loc[zero_act, "_active_sec"] = df.loc[zero_act, "_actual_sec"].where(
            df.loc[zero_act, "_actual_sec"] > 0, df.loc[zero_act, "_tot_time_sec"]
        )

    if c_eff:
        raw_eff = df[c_eff].astype(str).str.replace("%", "", regex=False).str.strip()
        df["_eff_raw"] = pd.to_numeric(raw_eff, errors="coerce")
        # If stored as decimal (e.g. 1.21 -> 121%)
        dec_mask = df["_eff_raw"].notna() & (df["_eff_raw"] > 0) & (df["_eff_raw"] <= 4.0)
        df.loc[dec_mask, "_eff_raw"] = df.loc[dec_mask, "_eff_raw"] * 100.0
    else:
        df["_eff_raw"] = float("nan")

    if c_ati:
        df["_ati_sec_raw"] = df[c_ati].apply(parse_duration_sec)
    else:
        df["_ati_sec_raw"] = 0.0

    return df


def load_prod_online(url):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        raw = r.read()
    head = raw[:300].lstrip().lower()
    if head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        raise RuntimeError("Google asked for sign-in on Productivity sheet.")
    return prepare_prod(pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, na_values=[""]))


def build_group_data(df, span, f_team="ALL", f_desig="ALL", f_designer="ALL", f_spec="ALL"):
    start, end = span
    sub = df[(df["_date"] >= start) & (df["_date"] <= end)].copy()
    if sub.empty:
        return None

    teams_opt = sorted(sub["_team"].dropna().unique().tolist())
    desigs_opt = sorted(sub["_designation"].dropna().unique().tolist())
    designers_opt = sorted(sub["_designer"].dropna().unique().tolist())
    specs_opt = sorted(sub["_base_spec"].dropna().unique().tolist())

    if f_team and f_team != "ALL":
        sub = sub[sub["_team"] == f_team]
    if f_desig and f_desig != "ALL":
        sub = sub[sub["_designation"] == f_desig]
    if f_designer and f_designer != "ALL":
        sub = sub[sub["_designer"] == f_designer]
    if f_spec and f_spec != "ALL":
        prefix = f"{str(f_spec).strip()}_"
        sub = sub[
            (sub["_base_spec"] == str(f_spec).strip())
            | (sub["_spec"].astype(str).str.strip().str.startswith(prefix))
        ]

    if sub.empty:
        return {
            "empty_filter": True,
            "span_label": span_label(span),
            "options": {"teams": teams_opt, "designations": desigs_opt, "designers": designers_opt, "specs": specs_opt},
        }

    num_days = max(1, int(sub["_date"].nunique()))
    rows = []
    for des_name, g in sub.groupby("_designer"):
        team = g["_team"].mode().iloc[0] if not g["_team"].mode().empty else g["_team"].iloc[0]
        desig = g["_designation"].mode().iloc[0] if not g["_designation"].mode().empty else g["_designation"].iloc[0]
        active_sec = float(g["_active_sec"].sum())
        total_eta_sec = float(g["_tot_time_sec"].sum()) if g["_tot_time_sec"].sum() > 0 else active_sec
        base_target = DESIGNATION_TARGETS_SEC.get(desig, DEFAULT_TARGET_SEC)
        exp_eta_sec = float(g["_spec_time_sec"].sum()) if g["_spec_time_sec"].sum() > 0 else float(base_target * g["_date"].nunique())
        files = int(round(g["_files"].sum()))
        rej = int(round(g["_rej"].sum()))
        qc_rej = int(round(g["_qc_rej"].sum()))
        cust_rej = int(round(g["_cust_rej"].sum()))
        ati_sec = (active_sec / files) if files > 0 else 0.0
        rej_pct = (rej / files * 100.0) if files > 0 else 0.0
        eff_pct = (exp_eta_sec / total_eta_sec * 100.0) if total_eta_sec > 0 else 0.0

        rows.append({
            "designer": des_name,
            "team": team,
            "designation": desig,
            "active_sec": active_sec,
            "active_hms": sec_to_hms(active_sec),
            "exp_eta_sec": exp_eta_sec,
            "exp_eta_hms": sec_to_hms(exp_eta_sec),
            "total_eta_sec": total_eta_sec,
            "total_eta_hms": sec_to_hms(total_eta_sec),
            "files": files,
            "ati_sec": ati_sec,
            "ati_hms": sec_to_hms(ati_sec),
            "rej": rej,
            "qc_rej": qc_rej,
            "cust_rej": cust_rej,
            "rej_pct": round(rej_pct, 2),
            "eff_pct": round(eff_pct, 2),
            "active_over_target": active_sec < exp_eta_sec,
            "eta_over_target": total_eta_sec < exp_eta_sec,
            "ati_high": ati_sec > 45,
        })

    rows.sort(key=lambda r: (r["designer"].lower(), r["team"].lower()))

    tot_files = sum(r["files"] for r in rows)
    tot_rej = sum(r["rej"] for r in rows)
    tot_active = sum(r["active_sec"] for r in rows)
    tot_exp_eta = sum(r["exp_eta_sec"] for r in rows)
    tot_eta = sum(r["total_eta_sec"] for r in rows)
    avg_ati_sec = (tot_active / tot_files) if tot_files > 0 else 0.0
    tot_rej_pct = (tot_rej / tot_files * 100.0) if tot_files > 0 else 0.0

    # Team summary breakdown (mirrors top-right summary box in screenshot)
    team_summary = []
    for tm, tg in sub.groupby("_team"):
        t_files = int(round(tg["_files"].sum()))
        t_rej = int(round(tg["_rej"].sum()))
        t_act = float(tg["_active_sec"].sum())
        t_ati = (t_act / t_files) if t_files > 0 else 0.0
        team_summary.append({
            "team": tm,
            "designers": int(tg["_designer"].nunique()),
            "files": t_files,
            "rej": t_rej,
            "rej_pct": round((t_rej / t_files * 100.0) if t_files > 0 else 0.0, 2),
            "active_hms": sec_to_hms(t_act),
            "avg_ati_hms": sec_to_hms(t_ati),
        })
    team_summary.sort(key=lambda x: -x["files"])

    # Designation targets reference
    desig_targets = [
        {"designation": k, "target_hms": sec_to_hms(v)}
        for k, v in DESIGNATION_TARGETS_SEC.items()
    ]

    # Auto takeaways
    top_vol = max(rows, key=lambda x: x["files"]) if rows else None
    high_ati = [r for r in rows if r["ati_sec"] > 45]
    rej_rows = sorted([r for r in rows if r["rej"] > 0], key=lambda x: (-x["rej"], -x["rej_pct"]))
    takeaways = [
        [
            "Group Output & Velocity",
            f"Processed {tot_files:,} files across {len(rows)} designers ({num_days} day{'s' if num_days!=1 else ''}) with Group Average ATI of {sec_to_hms(avg_ati_sec)}."
            + (f" Top producer: {top_vol['designer']} ({top_vol['files']:,} files)." if top_vol else "")
        ],
        [
            "Rejection Rate & Quality",
            f"Total Rejections: {tot_rej:,} files ({tot_rej_pct:.2f}% rejection rate). "
            + (f"Highest rejections: {', '.join(r['designer'] + ' (' + str(r['rej']) + ')' for r in rej_rows[:3])}." if rej_rows else "Zero rejections recorded in this window.")
        ],
        [
            "ATI & Target Time Compliance",
            f"{len(high_ati)} designer(s) exceeded 0:00:45 ATI threshold"
            + (f" ({', '.join(r['designer'] + ' @ ' + r['ati_hms'] for r in high_ati[:4])})." if high_ati else " — all designers within target ATI pace.")
        ],
    ]

    return {
        "span_label": span_label(span),
        "options": {"teams": teams_opt, "designations": desigs_opt, "designers": designers_opt, "specs": specs_opt},
        "kpis": {
            "designers": len(rows),
            "total_files": tot_files,
            "avg_ati_hms": sec_to_hms(avg_ati_sec),
            "total_rej": tot_rej,
            "rej_pct": round(tot_rej_pct, 2),
            "total_active_hms": sec_to_hms(tot_active),
            "total_exp_eta_hms": sec_to_hms(tot_exp_eta),
            "total_eta_hms": sec_to_hms(tot_eta),
        },
        "rows": rows,
        "team_summary": team_summary,
        "desig_targets": desig_targets,
        "takeaways": takeaways,
    }


def build_specwise_data(df, span, f_spec="ALL", f_team="ALL", f_desig="ALL", f_designer="ALL"):
    start, end = span
    sub = df[(df["_date"] >= start) & (df["_date"] <= end)].copy()
    if sub.empty:
        return None

    # Dropdown options list individual formatted sub-specs (e.g. '60345_01', '60345_02', '60345_03')
    specs_opt = sorted(sub["_spec"].dropna().unique().tolist())
    teams_opt = sorted(sub["_team"].dropna().unique().tolist())
    desigs_opt = sorted(sub["_designation"].dropna().unique().tolist())
    designers_opt = sorted(sub["_designer"].dropna().unique().tolist())

    # Build Top Specs & Low Specs panels across the full date window grouped by individual sub-spec (_01, _02, _03, etc.)
    spec_agg = []
    for sp, sg in sub.groupby("_spec"):
        s_files = int(round(sg["_files"].sum()))
        s_act = float(sg["_active_sec"].sum())
        s_rej = int(round(sg["_rej"].sum()))
        s_ati = (s_act / s_files) if s_files > 0 else 0.0
        s_exp = float(sg["_spec_time_sec"].sum())
        s_tot = float(sg["_tot_time_sec"].sum()) if sg["_tot_time_sec"].sum() > 0 else s_act
        s_eff = (s_exp / s_tot * 100.0) if (s_tot > 0 and s_exp > 0) else 0.0
        spec_agg.append({
            "spec": sp,
            "files": s_files,
            "rej": s_rej,
            "ati_sec": s_ati,
            "ati_hms": sec_to_hms(s_ati),
            "eff_pct": round(s_eff, 2),
        })
    top_specs = sorted(spec_agg, key=lambda x: (-x["files"], x["ati_sec"]))[:15]
    low_specs = sorted(spec_agg, key=lambda x: (x["files"], -x["ati_sec"]))[:15]

    # Filter by selected individual sub-spec (e.g. '60345_01', '60345_02', '60345_03')
    if f_spec and f_spec != "ALL":
        target_spec = format_sub_spec(f_spec)
        if "_" in target_spec:
            sub = sub[sub["_spec"] == target_spec]
        else:
            prefix = f"{target_spec}_"
            sub = sub[
                (sub["_base_spec"] == target_spec)
                | (sub["_spec"].astype(str).str.strip().str.startswith(prefix))
            ]
    if f_team and f_team != "ALL":
        sub = sub[sub["_team"] == f_team]
    if f_desig and f_desig != "ALL":
        sub = sub[sub["_designation"] == f_desig]
    if f_designer and f_designer != "ALL":
        sub = sub[sub["_designer"] == f_designer]

    if sub.empty:
        return {
            "empty_filter": True,
            "span_label": span_label(span),
            "options": {"specs": specs_opt, "teams": teams_opt, "designations": desigs_opt, "designers": designers_opt},
            "top_specs": top_specs,
            "low_specs": low_specs,
        }

    # Sorted unique designer group for the selected sub-spec
    designer_group = get_sorted_designer_group(sub, start, end, spec_filter=f_spec)

    # Target ATI for selected spec (derived from SpecTime / FileCount or default 0:00:45)
    tot_spec_time = float(sub["_spec_time_sec"].sum())
    tot_files_all = float(sub["_files"].sum())
    target_ati_sec = (tot_spec_time / tot_files_all) if (tot_spec_time > 0 and tot_files_all > 0) else 45.0

    rows = []
    for des_name in designer_group:
        g = sub[sub["_designer"] == des_name]
        if g.empty:
            continue
        team = g["_team"].mode().iloc[0] if not g["_team"].mode().empty else g["_team"].iloc[0]
        desig = g["_designation"].mode().iloc[0] if not g["_designation"].mode().empty else g["_designation"].iloc[0]
        files = int(round(g["_files"].sum()))
        active_sec = float(g["_active_sec"].sum())
        tot_sec = float(g["_tot_time_sec"].sum()) if g["_tot_time_sec"].sum() > 0 else active_sec
        spec_sec = float(g["_spec_time_sec"].sum())
        rej = int(round(g["_rej"].sum()))
        ati_sec = (active_sec / files) if files > 0 else 0.0

        if spec_sec > 0 and tot_sec > 0:
            eff_pct = (spec_sec / tot_sec) * 100.0
        elif g["_eff_raw"].notna().any():
            eff_pct = float(g["_eff_raw"].dropna().mean())
        elif ati_sec > 0:
            eff_pct = (target_ati_sec / ati_sec) * 100.0
        else:
            eff_pct = 0.0

        rows.append({
            "designer": des_name,
            "team": team,
            "designation": desig,
            "files": files,
            "ati_sec": ati_sec,
            "ati_hms": sec_to_hms(ati_sec),
            "rej": rej,
            "eff_pct": round(eff_pct, 2),
            "low_eff": eff_pct < 100.0,
            "high_ati": ati_sec > target_ati_sec,
        })

    tot_files = sum(r["files"] for r in rows)
    tot_rej = sum(r["rej"] for r in rows)
    tot_active = float(sub["_active_sec"].sum())
    avg_ati_sec = (tot_active / tot_files) if tot_files > 0 else 0.0
    avg_eff_pct = (sum(r["eff_pct"] * r["files"] for r in rows) / tot_files) if tot_files > 0 else (
        sum(r["eff_pct"] for r in rows) / len(rows) if rows else 0.0
    )

    best_eff = max(rows, key=lambda x: x["eff_pct"]) if rows else None
    low_eff_list = sorted([r for r in rows if r["eff_pct"] < 100.0], key=lambda x: x["eff_pct"])
    rej_list = sorted([r for r in rows if r["rej"] > 0], key=lambda x: -x["rej"])

    spec_label = (
        format_sub_spec(f_spec)
        if (f_spec and f_spec != "ALL")
        else f"All Specs ({len(specs_opt)})"
    )
    takeaways = [
        [
            "Spec Efficiency & Pace",
            f"Spec '{spec_label}' recorded {tot_files:,} files at {sec_to_hms(avg_ati_sec)} average ATI (Target: {sec_to_hms(target_ati_sec)}) and {avg_eff_pct:.2f}% weighted efficiency."
            + (f" Top efficiency: {best_eff['designer']} ({best_eff['eff_pct']:.2f}%)." if best_eff else "")
        ],
        [
            "Spec Rejections",
            f"Total Reject Files: {tot_rej:,}. "
            + (f"Impacted designers: {', '.join(r['designer'] + ' (' + str(r['rej']) + ')' for r in rej_list[:4])}." if rej_list else "Zero rejected files on this spec.")
        ],
        [
            "Below-Target Efficiency (<100%)",
            f"{len(low_eff_list)} designer(s) below 100% spec efficiency"
            + (f" ({', '.join(r['designer'] + ' @ ' + str(r['eff_pct']) + '%' for r in low_eff_list[:4])})." if low_eff_list else " — all designers met or exceeded spec target time.")
        ],
    ]

    return {
        "span_label": span_label(span),
        "selected_spec": spec_label,
        "designer_group": designer_group,
        "options": {"specs": specs_opt, "teams": teams_opt, "designations": desigs_opt, "designers": designers_opt},
        "kpis": {
            "spec_name": spec_label,
            "target_ati_hms": sec_to_hms(target_ati_sec),
            "avg_ati_hms": sec_to_hms(avg_ati_sec),
            "total_files": tot_files,
            "total_rej": tot_rej,
            "avg_eff_pct": round(avg_eff_pct, 2),
            "designers": len(rows),
        },
        "rows": rows,
        "top_specs": top_specs,
        "low_specs": low_specs,
        "takeaways": takeaways,
    }


def _split_people(raw_val):
    invalid = {"", "nan", "none", "unassigned", "unassigned client", "unknown", "n/a", "-"}
    s = str(raw_val or "").strip()
    if s.lower() in invalid:
        return []
    parts = [p.strip() for p in re.split(r"[,/&]+", s) if p.strip()]
    return [p for p in parts if p.lower() not in invalid]


def _find_consecutive_runs(present_days, global_days, min_len=STREAK_DAYS):
    runs = []
    cur = []
    for d in global_days:
        if d in present_days:
            cur.append(d)
        else:
            if len(cur) >= min_len:
                runs.append(cur)
            cur = []
    if len(cur) >= min_len:
        runs.append(cur)
    return runs


def _compute_streak_alerts(records, all_days, start=None, end=None, f_client="ALL", f_designer="ALL", f_qc="ALL", f_reason="ALL"):
    if not all_days or not records:
        return []
    df_rec = pd.DataFrame(records)
    df_rec = df_rec[df_rec["issue"].astype(str).str.strip() != ""].copy()
    if df_rec.empty:
        return []

    win_start = start if start is not None else all_days[0]
    win_end = end if end is not None else all_days[-1]

    dims = [("Designer", "designer"), ("QC", "qc"), ("Client", "client")]
    invalid = {"", "nan", "none", "unassigned", "unassigned client", "unknown", "n/a", "-"}

    for col in ("client", "designer", "qc", "issue"):
        df_rec[col] = df_rec[col].astype(str).str.strip()
        df_rec[col + "_key"] = df_rec[col].str.lower()

    if f_client and f_client != "ALL":
        fc = f_client.strip().lower()
        df_rec = df_rec[
            (df_rec["client_key"] == fc)
            | (df_rec.get("client_base_key", df_rec["client_key"]) == fc)
        ]
    if f_designer and f_designer != "ALL":
        df_rec = df_rec[df_rec["designer_key"] == f_designer.strip().lower()]
    if f_qc and f_qc != "ALL":
        df_rec = df_rec[df_rec["qc_key"] == f_qc.strip().lower()]
    if f_reason and f_reason != "ALL":
        df_rec = df_rec[df_rec["issue_key"] == f_reason.strip().lower()]

    if df_rec.empty:
        return []

    alerts = []
    for dim_label, col in dims:
        key_col = col + "_key"
        valid_df = df_rec[~df_rec[key_col].isin(invalid)].copy()
        if valid_df.empty:
            continue

        for ent_key, ent_g in valid_df.groupby(key_col):
            ent_disp = ent_g[col].mode().iloc[0] if not ent_g[col].mode().empty else ent_g[col].iloc[0]
            covered_runs = []

            for issue_key, g in ent_g.groupby("issue_key"):
                issue_disp = g["issue"].mode().iloc[0] if not g["issue"].mode().empty else g["issue"].iloc[0]
                present = set(g["date"])
                for run in _find_consecutive_runs(present, all_days, STREAK_DAYS):
                    if run[-1] < win_start or run[0] > win_end:
                        continue
                    covered_runs.append(set(run))
                    since, until = run[0], run[-1]
                    sub = g[(g["date"] >= since) & (g["date"] <= until)]
                    uniq_orders = sub["order_uid"].nunique() if "order_uid" in sub.columns else len(sub)
                    other_parts = []
                    for o_lbl, o_col in dims:
                        if o_lbl == dim_label:
                            continue
                        o_sub = sub[~sub[o_col + "_key"].isin(invalid)]
                        if not o_sub.empty:
                            counts = o_sub.groupby(o_col)["order_uid"].nunique() if "order_uid" in o_sub.columns else o_sub[o_col].value_counts()
                            counts = counts.sort_values(ascending=False)
                            other_parts.append(f"{o_lbl}s: " + ", ".join(f"{k} ({int(v)})" for k, v in counts.items()))
                    range_txt = f"{since:%d/%m/%Y} - {until:%d/%m/%Y}"
                    alerts.append({
                        "dimension": dim_label,
                        "entity": str(ent_disp),
                        "issue": str(issue_disp),
                        "days": int(len(run)),
                        "orders": int(uniq_orders),
                        "since": f"{since:%d/%m/%Y}",
                        "until": f"{until:%d/%m/%Y}",
                        "details": " | ".join(other_parts),
                        "message": (
                            f"ACTION NEEDED - {dim_label} {ent_disp}: '{issue_disp}' occurred {len(run)} consecutive days "
                            f"({uniq_orders} orders, {range_txt})"
                            + (f" | {' | '.join(other_parts)}" if other_parts else "")
                        ),
                    })

            all_present = set(ent_g["date"])
            for run in _find_consecutive_runs(all_present, all_days, STREAK_DAYS):
                if run[-1] < win_start or run[0] > win_end:
                    continue
                run_set = set(run)
                if any(run_set.issubset(cr) for cr in covered_runs):
                    continue
                since, until = run[0], run[-1]
                sub = ent_g[(ent_g["date"] >= since) & (ent_g["date"] <= until)]
                uniq_orders = sub["order_uid"].nunique() if "order_uid" in sub.columns else len(sub)
                top_issues = sub["issue"].value_counts()
                issue_summary = ", ".join(f"{k} ({int(v)})" for k, v in top_issues.items())
                other_parts = []
                for o_lbl, o_col in dims:
                    if o_lbl == dim_label:
                        continue
                    o_sub = sub[~sub[o_col + "_key"].isin(invalid)]
                    if not o_sub.empty:
                        counts = o_sub.groupby(o_col)["order_uid"].nunique() if "order_uid" in o_sub.columns else o_sub[o_col].value_counts()
                        counts = counts.sort_values(ascending=False)
                        other_parts.append(f"{o_lbl}s: " + ", ".join(f"{k} ({int(v)})" for k, v in counts.items()))
                range_txt = f"{since:%d/%m/%Y} - {until:%d/%m/%Y}"
                alerts.append({
                    "dimension": dim_label,
                    "entity": str(ent_disp),
                    "issue": f"Multiple Issues ({issue_summary})",
                    "days": int(len(run)),
                    "orders": int(uniq_orders),
                    "since": f"{since:%d/%m/%Y}",
                    "until": f"{until:%d/%m/%Y}",
                    "details": " | ".join(other_parts),
                    "message": (
                        f"ACTION NEEDED - {dim_label} {ent_disp}: mistakes occurred {len(run)} consecutive days "
                        f"[{issue_summary}] ({uniq_orders} orders, {range_txt})"
                        + (f" | {' | '.join(other_parts)}" if other_parts else "")
                    ),
                })

    dim_rank = {"Designer": 0, "QC": 1, "Client": 2}
    alerts.sort(key=lambda a: (dim_rank.get(a["dimension"], 9), -a["days"], -a["orders"], a["entity"].lower()))
    return alerts


def qa_streak_alerts(df, end, auditor=None, start=None, f_client="ALL", f_designer="ALL", f_qc="ALL", f_reason="ALL"):
    base = df[df["Date"] <= end].copy()
    if auditor and auditor.strip().lower() not in ("", "all"):
        aud_sub = base[base["auditor"] == auditor.strip().lower()]
        if not aud_sub.empty:
            base = aud_sub
    days = sorted(base["Date"].unique())
    if not days:
        return []

    records = []
    for idx, r in base[base["err"] > 0].iterrows():
        spec_int = int(r["Spec ID"])
        base_client = qa_client_name(spec_int)
        client_disp = f"{base_client} ({spec_int})"
        designers = _split_people(r.get("Designer Name", "")) or [""]
        qcs = _split_people(r.get("QC Name", "")) or [""]
        cats = [c.strip() for c in str(r.get("Rejection Category", "")).split(",") if c.strip() and c.strip().lower() not in ("nan", "no issues")]
        if not cats:
            cats = ["Quality Error"]
        order_uid = f"{r['Date']:%Y-%m-%d}_{r.get('Order ID', idx)}"
        for cat in cats:
            for des in designers:
                for qc in qcs:
                    records.append({
                        "date": r["Date"],
                        "order_uid": order_uid,
                        "client": client_disp,
                        "client_base_key": base_client.strip().lower(),
                        "designer": des,
                        "qc": qc,
                        "issue": cat,
                    })
    return _compute_streak_alerts(
        records, days, start=start, end=end,
        f_client=f_client, f_designer=f_designer, f_qc=f_qc, f_reason=f_reason
    )


def cr_streak_alerts(zibber, other, end, start=None, f_client="ALL", f_designer="ALL", f_qc="ALL", f_reason="ALL"):
    frames = []
    for x in (zibber, other):
        if x is not None and not x.empty:
            frames.append(x[x["_date"] <= end][["_date", "_client_name", "_norm_des", "_norm_qc", "_norm_rsn"]].copy())
    if not frames:
        return []
    base = pd.concat(frames, ignore_index=True)
    days = sorted(base["_date"].dropna().unique())
    if not days:
        return []

    records = []
    for idx, r in base.iterrows():
        rsn = str(r.get("_norm_rsn", "")).strip()
        if not rsn or rsn.lower() in ("nan", "none", "no reason"):
            rsn = "Client Rejection"
        designers = _split_people(r.get("_norm_des", "")) or [""]
        qcs = _split_people(r.get("_norm_qc", "")) or [""]
        order_uid = f"{r['_date']:%Y-%m-%d}_{idx}"
        cname = str(r.get("_client_name", "")).strip()
        for des in designers:
            for qc in qcs:
                records.append({
                    "date": r["_date"],
                    "order_uid": order_uid,
                    "client": cname,
                    "client_base_key": cname.lower(),
                    "designer": des,
                    "qc": qc,
                    "issue": rsn,
                })
    return _compute_streak_alerts(
        records, days, start=start, end=end,
        f_client=f_client, f_designer=f_designer, f_qc=f_qc, f_reason=f_reason
    )


def streak_flags(df, end, auditor=None, start=None):
    alerts = qa_streak_alerts(df, end, auditor=auditor, start=start)
    out = []
    grouped = {}
    for a in alerts:
        grouped.setdefault((a["dimension"], a["entity"]), []).append(a)
    for (dim_lbl, ent), items in grouped.items():
        out.append(
            f"ACTION NEEDED - {dim_lbl} {ent}: mistakes {STREAK_DAYS}+ days in a row. "
            f"Please decide what action to take."
        )
        for a in items:
            out.append(
                f"   - {a['issue']} ({a['days']} days, {a['orders']} orders, {a['since']} to {a['until']})"
                + (f" | {a['details']}" if a["details"] else "")
            )
    return out


def build_qa_takeaways(df, span, auditor):
    start, end = span
    d = df[(df["Date"] >= start) & (df["Date"] <= end) & (df["auditor"] == auditor.lower())]
    if d.empty:
        return []
    T, E = int(d["imgs"].sum()), int(d["err"].sum())
    spec_order = list(dict.fromkeys(d["Spec ID"].astype(int)))
    stats = {spec: (int(d[d["Spec ID"].astype(int) == spec]["imgs"].sum()), int(d[d["Spec ID"].astype(int) == spec]["err"].sum())) for spec in spec_order}
    rows = []

    if E > 0:
        top_spec = max(stats, key=lambda s: stats[s][1])
        st_imgs, st_err = stats[top_spec]
        ed_series = d[(d["Spec ID"].astype(int) == top_spec) & (d["err"] > 0)].groupby("Designer Name")["err"].sum().sort_values(ascending=False)
        eds = ", ".join(str(k) for k in ed_series.index if str(k).strip() and str(k).lower() != "nan")
        ed_txt = f" Editors involved: {eds}." if eds else ""
        rows.append([
            "Primary Defect Source",
            f"{pct(st_err, E):.0f}% of errors ({st_err}/{E}) originated from {qa_client_name(top_spec)} ({pct(st_err, st_imgs):.2f}% defect rate across {st_imgs} images).{ed_txt}"
        ])
        qc_series = d[d["err"] > 0].groupby("QC Name")["err"].sum().sort_values(ascending=False)
        qc_parts = [f"{k} leaked {int(v)} error{'s' if int(v)!=1 else ''}" for k, v in qc_series.items() if int(v) > 0 and str(k).strip() and str(k).lower() != "nan"]
        if qc_parts:
            rows.append([
                "QC Accountability",
                f"{E} defective image{'s' if E!=1 else ''} passed through QC review — {' and '.join(qc_parts)}."
            ])
        worst = max(pct(e, t) for t, e in stats.values()) if stats else 0
        if worst > ESCALATE_ABOVE:
            rows.append([
                "Immediate Action Taken",
                "Already informed QV regarding the identified issues for immediate fix. Escalate to Performance Team if repeated for 2 consecutive days."
            ])
        else:
            rows.append([
                "Immediate Action Taken",
                "Already informed QV regarding the identified issues for immediate fix."
            ])
    else:
        rows.append([
            "Primary Defect Source",
            f"0 errors recorded across {T} audited images (0.00% defect rate). All batches passed cleanly."
        ])
        rows.append([
            "QC Accountability",
            "100% QC pass accuracy — zero defective images leaked."
        ])
        rows.append([
            "Immediate Action Taken",
            "No corrective action required; all clients operating within quality thresholds."
        ])

    alerts = qa_streak_alerts(df, end, auditor=auditor, start=start)
    if alerts:
        rows.append([
            f"{STREAK_DAYS}-Day Recurring Issue Reminders",
            " | ".join(f"{a['dimension']} {a['entity']}: {a['issue']} ({a['days']}d)" for a in alerts)
        ])
    return rows


def build_qa_report_text(df, span, auditor):
    start, end = span
    multi = start != end
    d = df[(df["Date"] >= start) & (df["Date"] <= end) & (df["auditor"] == auditor.lower())]
    if d.empty:
        return None

    lines = []
    add = lines.append
    add(f"📊 QUALITY AUDIT SUMMARY — {auditor.upper()} ({span_label(span)})")
    add("1. TOTAL IMAGES AUDITED")
    spec_order = list(dict.fromkeys(d["Spec ID"].astype(int)))
    stats = {}
    for spec in spec_order:
        s = d[d["Spec ID"].astype(int) == spec]
        stats[spec] = (int(s["imgs"].sum()), int(s["err"].sum()))
        t, e = stats[spec]
        add(f"{qa_client_name(spec)} ({spec}): {t} images | {e} errors ({pct(e, t):.2f}% defect rate)")
    T, E = int(d["imgs"].sum()), int(d["err"].sum())
    add(f"TOTAL: {T} images | {E} errors ({pct(E, T):.2f}% defect rate)")
    add("")

    add("2. ISSUES IDENTIFIED")
    for _, r in d.iterrows():
        pre = f"{r['Date']:%d/%m} | " if multi else ""
        add(
            f"{pre}{int(r['Order ID'])} ({r['err']} errors): {clean_comment(r['Comments'], r['Rejection Category'])}"
            if r["err"]
            else f"{pre}{int(r['Order ID'])} (0 errors): Clean batch."
        )
    add("")

    add("3. ACTIONS TAKEN")
    worst = max(pct(e, t) for t, e in stats.values()) if stats else 0
    if worst > ESCALATE_ABOVE:
        lines.extend(ACTION_LINES_ESCALATE)
    else:
        add(ACTION_LINE_OK)
    lines.extend(streak_flags(df, end, auditor=auditor, start=start))
    add("")

    add("4. ADDITIONAL FINDINGS")
    if E:
        top_spec = max(stats, key=lambda s: stats[s][1])
        ed_series = d[(d["Spec ID"].astype(int) == top_spec) & (d["err"] > 0)].groupby("Designer Name")["err"].sum().sort_values(ascending=False)
        eds = ", ".join(str(k) for k in ed_series.index if str(k).strip() and str(k).lower() != "nan")
        add(
            f"{pct(stats[top_spec][1], E):.2f}% of errors ({stats[top_spec][1]}/{E}) came from {qa_client_name(top_spec)} ({top_spec})"
            + (f" under editor(s) {eds}." if eds else ".")
        )
        qc_series = d[d["err"] > 0].groupby("QC Name")["err"].sum().sort_values(ascending=False)
        qc_parts = [f"{k} leaked {int(v)} errors" for k, v in qc_series.items() if int(v) > 0 and str(k).strip() and str(k).lower() != "nan"]
        if qc_parts:
            add(f"QC Leak: {E} defective images passed review — {', '.join(qc_parts)}.")
        fault = d["Fault"].astype(str).str.strip().str.upper() if "Fault" in d.columns else pd.Series("", index=d.index)
        prod = int(d.loc[fault == "PRODUCTION FAULT", "err"].sum())
        cust = int(d.loc[fault == "CUSTOMER FAULT", "err"].sum())
        if prod or cust:
            add(f"Fault Split: Production Fault {prod} errors ({pct(prod, E):.0f}%) | Customer Fault {cust} errors ({pct(cust, E):.0f}%).")
    for spec in spec_order:
        t, e = stats[spec]
        if pct(e, t) <= STRONG_BELOW:
            add(f"{qa_client_name(spec)} performed strongly with only a {pct(e, t):.2f}% defect rate.")
    return "\n".join(lines)


def extract_qa_orders(df):
    orders = []
    fault_col = "Fault" if "Fault" in df.columns else None
    cat_col = "Rejection Category" if "Rejection Category" in df.columns else None

    for _, r in df.iterrows():
        raw_oid = r.get("Order ID")
        order_id = str(int(raw_oid)) if pd.notna(raw_oid) else "N/A"
        raw_spec = r.get("Spec ID")
        spec_id = int(raw_spec) if pd.notna(raw_spec) else 0
        client = qa_client_name(spec_id)
        raw_date = r.get("Date")
        date_str = str(raw_date.date()) if pd.notna(raw_date) else ""

        vol = int(r["imgs"])
        err = int(r["err"])

        des = str(r.get("Designer Name", "")).strip()
        if not des or des.lower() == "nan":
            des = "Unassigned"
        qc = str(r.get("QC Name", "")).strip()
        if not qc or qc.lower() == "nan":
            qc = "Unassigned"

        raw_cat = r.get(cat_col) if cat_col else None
        cats = str(raw_cat).strip() if pd.notna(raw_cat) else ""
        if not cats or cats.lower() == "nan":
            cats = "No Issues"

        comments = clean_comment(r.get("Comments"), cats)
        raw_fault = r.get(fault_col) if fault_col else None
        fault_val = str(raw_fault).strip().upper() if pd.notna(raw_fault) else ""
        is_customer = ("CUSTOMER" in fault_val) or ("CUSTOMER" in cats.upper())

        orders.append({
            "source": "QA",
            "date": date_str,
            "order_id": order_id,
            "spec_id": spec_id,
            "client": client,
            "volume": vol,
            "errors": err,
            "categories": cats,
            "comments": comments,
            "designer": des,
            "qc": qc,
            "fault": fault_val,
            "is_customer": is_customer,
            "auditor": str(r.get("auditor", "")),
        })
    return orders


# ===================================================================== CLIENT REJECTION DATA
def prepare_cr(raw, is_zibber=False):
    hdr = None
    for i in range(min(10, len(raw))):
        row_vals = [norm(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
        if "rejected job id" in row_vals or "job id" in row_vals:
            hdr = i
    if hdr is None:
        raise RuntimeError("Header row ('Rejected Job ID') not found in Client Rejection sheet.")
    df = raw.iloc[hdr + 1:].copy()
    cols, seen = [], {}
    for v in raw.iloc[hdr].tolist():
        name = norm(v) if pd.notna(v) else "_blank"
        seen[name] = seen.get(name, 0) + 1
        cols.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    df.columns = cols
    df = df.loc[:, ~df.columns.duplicated()].copy()

    job_col = find_col(df, JOB_COL_ALIASES)
    if not job_col:
        raise RuntimeError("Job ID column not found in Client Rejection sheet.")

    df["_date"] = pd.to_datetime(df["date"], errors="coerce").dt.normalize() if "date" in df.columns else pd.NaT
    df["_vol"] = pd.to_numeric(df["volume"], errors="coerce").fillna(0) if "volume" in df.columns else 0
    df["_is_zibber"] = is_zibber
    df = df[df[job_col].notna() & df["_date"].notna()].copy()

    if is_zibber:
        df["_client_name"] = "Zibber Client"
    else:
        if "client name" in df.columns:
            c = df["client name"].astype(str).str.strip()
            df["_client_name"] = c.where(~df["client name"].isna() & ~c.str.lower().isin(["", "nan"]), "Unassigned client")
        else:
            df["_client_name"] = "Unassigned client"

    des_col = find_col(df, DES_COL_ALIASES)
    qc_col = find_col(df, QC_COL_ALIASES)
    rsn_col = find_col(df, ZIBBER_REASON_ALIASES if is_zibber else OTHER_REASON_ALIASES) or find_col(df, OTHER_REASON_ALIASES + ZIBBER_REASON_ALIASES)

    def _clean_series(col_name, invalid_vals=("", "nan", "none")):
        if not col_name or col_name not in df.columns:
            return pd.Series("", index=df.index)
        s = df[col_name].fillna("").astype(str).str.strip()
        return s.where(~s.str.lower().isin(invalid_vals), "")

    df["_norm_des"] = _clean_series(des_col)
    df["_norm_qc"] = _clean_series(qc_col)
    df["_norm_rsn"] = _clean_series(rsn_col, invalid_vals=("", "nan", "none", "no reason"))
    return df


def read_cr_book(source):
    book = pd.read_excel(source, sheet_name=[CR_ZIBBER_SHEET, CR_OTHER_SHEET], header=None)
    return prepare_cr(book[CR_ZIBBER_SHEET], is_zibber=True), prepare_cr(book[CR_OTHER_SHEET], is_zibber=False)


def load_cr_online(url):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        raw = r.read()
    if not raw.startswith(b"PK"):
        raise RuntimeError("Google asked for sign-in on Client Rejection sheet.")
    return read_cr_book(io.BytesIO(raw))


def delay_info(rows):
    if "delay reason" not in rows.columns:
        return 0, 0, []
    txt = rows["delay reason"].astype(str).str.strip()
    mask = rows["delay reason"].notna() & ~txt.str.lower().isin(["", "nan", "delay reason", "none"])
    reasons = []
    for r in txt[mask]:
        if r not in reasons:
            reasons.append(r)
    return int(mask.sum()), int(rows.loc[mask, "_vol"].sum()), reasons


def fault_line(rows, key, label):
    if "reason" not in rows.columns:
        return f"{label} - 0 orders / 0 images"
    sub = rows[rows["reason"].astype(str).str.upper().str.contains(key, na=False)]
    return f"{label} - {len(sub)} orders / {int(sub['_vol'].sum())} images"


def cr_block_lines(rows):
    d_orders, d_images, reasons = delay_info(rows)
    delayed = "0" if d_orders == 0 else f"{d_orders} ({d_images} images)"
    return [
        f"rejections Order - {len(rows)}",
        f"rejections imags - {int(rows['_vol'].sum())}",
        fault_line(rows, "CUSTOMER", "Customer fault"),
        fault_line(rows, "PRODUCTION", "Production fault"),
        f"Total delayed orders of images - {delayed}",
        f"Delay reason - {', '.join(reasons) if reasons else 'None'}",
    ]


def build_cr_report_text(zibber, other, span):
    start, end = span
    z = zibber[(zibber["_date"] >= start) & (zibber["_date"] <= end)]
    o = other[(other["_date"] >= start) & (other["_date"] <= end)].copy()
    if z.empty and o.empty:
        return None

    found = list(dict.fromkeys(o["_client_name"]))
    clients = [c for c in FIXED_CLIENTS if c in found] + [c for c in found if c not in FIXED_CLIENTS]

    date_str = f"{start:%d %b %Y}" if start == end else f"{start:%d %b %Y} to {end:%d %b %Y}"
    lines = [f"Date: {date_str}"]
    if len(z):
        lines += ["", "Zibber Client reject jobs", ""] + cr_block_lines(z)
    if clients:
        lines += ["", "Rest of RE client-rejected jobs"]
        for c in clients:
            lines += ["", c] + cr_block_lines(o[o["_client_name"] == c])

    alerts = cr_streak_alerts(zibber, other, end, start=start)
    if alerts:
        lines += ["", f"3-DAY RECURRING REJECTION REMINDERS ({STREAK_DAYS}+ CONSECUTIVE DAYS)"]
        grouped = {}
        for a in alerts:
            grouped.setdefault((a["dimension"], a["entity"]), []).append(a)
        for (dim_lbl, ent), items in grouped.items():
            lines.append(f"ACTION NEEDED - {dim_lbl} {ent}: mistakes {STREAK_DAYS}+ days in a row. Please decide what action to take.")
            for a in items:
                lines.append(
                    f"   - {a['issue']} ({a['days']} days, {a['orders']} orders, {a['since']} to {a['until']})"
                    + (f" | {a['details']}" if a["details"] else "")
                )
    return "\n".join(lines)


def extract_cr_orders(df, client_override=None):
    job_col = find_col(df, JOB_COL_ALIASES)
    corr_qc_col = find_col(df, CORRECTION_QC_ALIASES)

    orders = []
    for _, r in df.iterrows():
        job_val = r[job_col] if job_col and job_col in r else None
        job_id = str(job_val).strip() if pd.notna(job_val) and str(job_val).strip() not in ["", "nan"] else "N/A"
        raw_date = r.get("_date")
        date_str = str(raw_date.date()) if pd.notna(raw_date) else ""
        vol = int(r["_vol"])

        client = client_override if client_override else r.get("_client_name", "Unassigned client")
        reason = r.get("_norm_rsn") or "No reason"
        des = r.get("_norm_des") or "Unassigned"
        qc = r.get("_norm_qc") or "Unassigned"
        corr_qc = str(r[corr_qc_col]).strip() if corr_qc_col and pd.notna(r[corr_qc_col]) and str(r[corr_qc_col]).strip() not in ["", "nan"] else ""

        delay_r = (
            str(r["delay reason"]).strip()
            if "delay reason" in r and pd.notna(r["delay reason"]) and str(r["delay reason"]).strip().lower() not in ["", "nan", "delay reason", "none"]
            else None
        )
        is_customer = (
            "CUSTOMER" in str(r.get("reason", "")).upper()
            or "CUSTOMER" in str(r.get("correction qc", "")).upper()
            or "CUSTOMER" in reason.upper()
        )

        orders.append({
            "source": "CR",
            "date": date_str,
            "client": client,
            "job_id": job_id,
            "volume": vol,
            "reason": reason,
            "designer": des,
            "qc": qc,
            "correction_qc": corr_qc,
            "delay": delay_r,
            "is_customer": is_customer,
        })
    return orders


# ===================================================================== CLIENT ATTENTION RADAR
def normalize_client_key(name):
    s = str(name).strip()
    m = re.search(r"(\d{4,6})", s)
    if m:
        code = int(m.group(1))
        if code in QA_CLIENTS:
            return f"{QA_CLIENTS[code]} ({code})"
        return f"Client {code}"
    if "zibber" in s.lower():
        return "Zibber Client"
    if "ogulo" in s.lower():
        return "Ogulo GmbH (67676)"
    if "homevisual" in s.lower():
        return "HomeVisuals (67484)"
    return s


def build_client_attention(qa_df, cr_zib, cr_oth, start_str=None, end_str=None, f_auditor="ALL"):
    auditors = []
    if qa_df is not None and not qa_df.empty and "auditor" in qa_df.columns:
        auditors = sorted([a for a in qa_df["auditor"].dropna().unique().tolist() if a and str(a).lower() not in ("nan", "unknown", "")])
        if f_auditor and f_auditor != "ALL":
            qa_df = qa_df[qa_df["auditor"] == f_auditor.lower()]
    all_dates = []
    if qa_df is not None and not qa_df.empty:
        all_dates.extend(qa_df["Date"].dropna().tolist())
    if cr_zib is not None and not cr_zib.empty:
        all_dates.extend(cr_zib["_date"].dropna().tolist())
    if cr_oth is not None and not cr_oth.empty:
        all_dates.extend(cr_oth["_date"].dropna().tolist())

    today = pd.Timestamp.today().normalize()
    past = [d for d in all_dates if d <= today]
    anchor = max(past) if past else (max(all_dates) if all_dates else today)

    distinct_past = sorted(set(past))
    prev_day = distinct_past[-2] if len(distinct_past) >= 2 else anchor - pd.Timedelta(days=1)

    windows = {
        "DAILY": {
            "label": f"Daily ({anchor:%d %b %Y} vs {prev_day:%d %b %Y})",
            "a": (anchor, anchor),
            "b": (prev_day, prev_day),
        },
        "WEEKLY": {
            "label": f"Weekly (Last 7d: {anchor - pd.Timedelta(days=6):%d %b}–{anchor:%d %b} vs Prior 7d)",
            "a": (anchor - pd.Timedelta(days=6), anchor),
            "b": (anchor - pd.Timedelta(days=13), anchor - pd.Timedelta(days=7)),
        },
        "MONTHLY": {
            "label": f"Monthly ({anchor:%B %Y} MTD vs Previous Month)",
            "a": (anchor.replace(day=1), anchor),
            "b": ((anchor.replace(day=1) - pd.Timedelta(days=1)).replace(day=1), anchor.replace(day=1) - pd.Timedelta(days=1)),
        },
    }

    if start_str and end_str:
        c_start = pd.Timestamp(start_str).normalize()
        c_end = pd.Timestamp(end_str).normalize()
        if c_end < c_start:
            c_start, c_end = c_end, c_start
        span_days = (c_end - c_start).days + 1
        p_end = c_start - pd.Timedelta(days=1)
        p_start = p_end - pd.Timedelta(days=span_days - 1)
        windows["CUSTOM"] = {
            "label": f"Custom Range ({span_label((c_start, c_end))} vs Prior {span_days}d: {span_label((p_start, p_end))})",
            "a": (c_start, c_end),
            "b": (p_start, p_end),
        }

    cr_frames = [
        x[[c for c in ["_date", "_vol", "_client_name", "_norm_des", "_norm_qc", "_norm_rsn", "reason"] if c in x.columns]].copy()
        for x in (cr_zib, cr_oth) if x is not None and not x.empty
    ]
    cr_all = pd.concat(cr_frames, ignore_index=True) if cr_frames else pd.DataFrame()

    result = {"anchor": str(anchor.date()), "prev_day": str(prev_day.date()), "auditors": auditors, "windows": {}}

    for wkey, winfo in windows.items():
        a1, a2 = winfo["a"]
        b1, b2 = winfo["b"]
        clients = {}

        def ensure(cname):
            if cname not in clients:
                clients[cname] = {
                    "client": cname,
                    "qa_a_imgs": 0, "qa_a_err": 0, "qa_a_ord": 0,
                    "qa_b_imgs": 0, "qa_b_err": 0,
                    "cr_a_ord": 0, "cr_a_imgs": 0, "cr_a_prod": 0, "cr_a_cust": 0,
                    "cr_b_ord": 0, "cr_b_imgs": 0,
                    "qa_reasons": {},
                    "qa_designers": {},
                    "qa_qcs": {},
                    "cr_reasons": {},
                    "cr_designers": {},
                    "cr_qcs": {},
                }
            return clients[cname]

        if qa_df is not None and not qa_df.empty:
            qa_a = qa_df[(qa_df["Date"] >= a1) & (qa_df["Date"] <= a2)]
            qa_b = qa_df[(qa_df["Date"] >= b1) & (qa_df["Date"] <= b2)]
            for _, r in qa_a.iterrows():
                c = ensure(normalize_client_key(r["_client_name"]))
                c["qa_a_imgs"] += int(r["imgs"])
                c["qa_a_err"] += int(r["err"])
                c["qa_a_ord"] += 1
                if int(r["err"]) > 0:
                    des = str(r.get("Designer Name", "")).strip()
                    if des and des.lower() not in ("", "nan", "none"):
                        c["qa_designers"][des] = c["qa_designers"].get(des, 0) + int(r["err"])
                    qc = str(r.get("QC Name", "")).strip()
                    if qc and qc.lower() not in ("", "nan", "none"):
                        c["qa_qcs"][qc] = c["qa_qcs"].get(qc, 0) + int(r["err"])
                    for cat in str(r.get("Rejection Category", "")).split(","):
                        cat = cat.strip()
                        if cat and cat.lower() not in ("", "nan", "no issues"):
                            c["qa_reasons"][cat] = c["qa_reasons"].get(cat, 0) + 1
            for _, r in qa_b.iterrows():
                c = ensure(normalize_client_key(r["_client_name"]))
                c["qa_b_imgs"] += int(r["imgs"])
                c["qa_b_err"] += int(r["err"])

        if not cr_all.empty:
            cr_a = cr_all[(cr_all["_date"] >= a1) & (cr_all["_date"] <= a2)]
            cr_b = cr_all[(cr_all["_date"] >= b1) & (cr_all["_date"] <= b2)]
            for _, r in cr_a.iterrows():
                c = ensure(normalize_client_key(r["_client_name"]))
                c["cr_a_ord"] += 1
                c["cr_a_imgs"] += int(r["_vol"])
                rsn_upper = str(r.get("reason", "")).upper()
                if "CUSTOMER" in rsn_upper:
                    c["cr_a_cust"] += 1
                else:
                    c["cr_a_prod"] += 1
                if r["_norm_des"]:
                    c["cr_designers"][r["_norm_des"]] = c["cr_designers"].get(r["_norm_des"], 0) + 1
                if r["_norm_qc"]:
                    c["cr_qcs"][r["_norm_qc"]] = c["cr_qcs"].get(r["_norm_qc"], 0) + 1
                if r["_norm_rsn"]:
                    c["cr_reasons"][r["_norm_rsn"]] = c["cr_reasons"].get(r["_norm_rsn"], 0) + 1
            for _, r in cr_b.iterrows():
                c = ensure(normalize_client_key(r["_client_name"]))
                c["cr_b_ord"] += 1
                c["cr_b_imgs"] += int(r["_vol"])

        ranked = []
        for cname, c in clients.items():
            qa_a_rate = round(pct(c["qa_a_err"], c["qa_a_imgs"]), 2)
            qa_b_rate = round(pct(c["qa_b_err"], c["qa_b_imgs"]), 2)
            qa_rate_diff = round(qa_a_rate - qa_b_rate, 2)
            cr_ord_diff = c["cr_a_ord"] - c["cr_b_ord"]

            score = 0
            flags = []
            if qa_a_rate > ESCALATE_ABOVE:
                score += 45 + (qa_a_rate - ESCALATE_ABOVE) * 2
                flags.append(f"QA defect rate {qa_a_rate:.1f}% (> {ESCALATE_ABOVE}%)")
            elif qa_a_rate > 5.0:
                score += 18
                flags.append(f"Elevated QA defect rate ({qa_a_rate:.1f}%)")

            if qa_rate_diff >= 2.5 and c["qa_a_imgs"] > 0:
                score += 20 + qa_rate_diff
                flags.append(f"QA rate worsened +{qa_rate_diff:.1f}% vs prior period")

            if c["cr_a_ord"] >= 5:
                score += 35 + c["cr_a_ord"] * 2
                flags.append(f"High external rejections ({c['cr_a_ord']} orders / {c['cr_a_imgs']} imgs)")
            elif c["cr_a_ord"] >= 2:
                score += 18 + c["cr_a_ord"] * 2
                flags.append(f"{c['cr_a_ord']} external rejection orders")
            elif c["cr_a_ord"] == 1:
                score += 8
                flags.append("1 external rejection order")

            if cr_ord_diff > 0:
                score += cr_ord_diff * 5
                flags.append(f"+{cr_ord_diff} more rejected orders than prior period")

            if score >= 40:
                status_level = "CRITICAL"
            elif score >= 15:
                status_level = "WATCH"
            else:
                status_level = "HEALTHY"
                if not flags:
                    flags.append("Within acceptable quality thresholds")

            fmt_top = lambda d: [f"{k} ({v})" for k, v in sorted(d.items(), key=lambda x: -x[1])[:3]]

            ranked.append({
                "client": cname,
                "status": status_level,
                "score": round(score, 1),
                "flags": flags,
                "qa_a_imgs": c["qa_a_imgs"],
                "qa_a_err": c["qa_a_err"],
                "qa_a_rate": qa_a_rate,
                "qa_b_rate": qa_b_rate,
                "qa_rate_diff": qa_rate_diff,
                "qa_reasons": fmt_top(c["qa_reasons"]),
                "qa_designers": fmt_top(c["qa_designers"]),
                "qa_qcs": fmt_top(c["qa_qcs"]),
                "cr_a_ord": c["cr_a_ord"],
                "cr_a_imgs": c["cr_a_imgs"],
                "cr_a_prod": c["cr_a_prod"],
                "cr_a_cust": c["cr_a_cust"],
                "cr_b_ord": c["cr_b_ord"],
                "cr_ord_diff": cr_ord_diff,
                "cr_reasons": fmt_top(c["cr_reasons"]),
                "cr_designers": fmt_top(c["cr_designers"]),
                "cr_qcs": fmt_top(c["cr_qcs"]),
            })

        ranked.sort(key=lambda x: (-x["score"], -x["cr_a_ord"], -x["qa_a_rate"]))
        result["windows"][wkey] = {
            "label": winfo["label"],
            "a_span": span_label((a1, a2)),
            "b_span": span_label((b1, b2)),
            "clients": ranked,
            "critical_count": sum(1 for x in ranked if x["status"] == "CRITICAL"),
            "watch_count": sum(1 for x in ranked if x["status"] == "WATCH"),
            "healthy_count": sum(1 for x in ranked if x["status"] == "HEALTHY"),
        }
    return result


# ===================================================================== ANALYTICS & CHARTS PAYLOAD
def build_charts_data(qa_df, cr_zib, cr_oth, days_back=30, start_str=None, end_str=None, f_client="ALL", f_designer="ALL", f_qc="ALL", f_reason="ALL", f_auditor="ALL"):
    today = pd.Timestamp.today().normalize()
    all_dates = []
    if qa_df is not None and not qa_df.empty:
        all_dates.extend(qa_df["Date"].dropna().tolist())
    if cr_zib is not None and not cr_zib.empty:
        all_dates.extend(cr_zib["_date"].dropna().tolist())
    if cr_oth is not None and not cr_oth.empty:
        all_dates.extend(cr_oth["_date"].dropna().tolist())

    past = [d for d in all_dates if d <= today]
    default_end = max(past) if past else (max(all_dates) if all_dates else today)
    min_date = min(all_dates) if all_dates else default_end

    if start_str and end_str:
        start = pd.Timestamp(start_str).normalize()
        end = pd.Timestamp(end_str).normalize()
        if end < start:
            start, end = end, start
    else:
        end = default_end
        start = (end - pd.Timedelta(days=int(days_back) - 1)) if int(days_back) > 0 else min_date

    qa_sub = qa_df[(qa_df["Date"] >= start) & (qa_df["Date"] <= end)].copy() if (qa_df is not None and not qa_df.empty) else pd.DataFrame()

    cr_frames = []
    for x in (cr_zib, cr_oth):
        if x is not None and not x.empty:
            cols_keep = [c for c in ["_date", "_vol", "_client_name", "_norm_des", "_norm_qc", "_norm_rsn", "reason"] if c in x.columns]
            cr_frames.append(x[cols_keep].copy())
    cr_all = pd.concat(cr_frames, ignore_index=True) if cr_frames else pd.DataFrame()
    cr_sub = cr_all[(cr_all["_date"] >= start) & (cr_all["_date"] <= end)].copy() if not cr_all.empty else pd.DataFrame()

    opt_clients, opt_designers, opt_qcs, opt_reasons, opt_auditors = set(), set(), set(), set(), set()

    if not qa_sub.empty:
        qa_sub["_norm_client"] = qa_sub["_client_name"].apply(normalize_client_key)
        qa_sub["_norm_des"] = qa_sub["Designer Name"].fillna("").astype(str).str.strip() if "Designer Name" in qa_sub.columns else ""
        qa_sub["_norm_qc"] = qa_sub["QC Name"].fillna("").astype(str).str.strip() if "QC Name" in qa_sub.columns else ""
        qa_sub["_norm_rsn_list"] = qa_sub["Rejection Category"].fillna("").astype(str).apply(
            lambda s: [x.strip() for x in s.split(",") if x.strip() and x.strip().lower() not in ("nan", "no issues")]
        ) if "Rejection Category" in qa_sub.columns else [[] for _ in range(len(qa_sub))]

        if "auditor" in qa_sub.columns:
            opt_auditors.update(str(x) for x in qa_sub["auditor"].dropna().unique() if str(x).strip() and str(x).lower() not in ("nan", "unknown"))
        opt_clients.update(str(x) for x in qa_sub["_norm_client"].dropna().unique() if str(x).strip() and str(x).lower() != "nan")
        opt_designers.update(str(x) for x in qa_sub["_norm_des"].dropna().unique() if str(x).strip() and str(x).lower() not in ("nan", "none"))
        opt_qcs.update(str(x) for x in qa_sub["_norm_qc"].dropna().unique() if str(x).strip() and str(x).lower() not in ("nan", "none"))
        for lst in qa_sub["_norm_rsn_list"]:
            opt_reasons.update(lst)

    if not cr_sub.empty:
        cr_sub["_norm_client"] = cr_sub["_client_name"].apply(normalize_client_key)
        opt_clients.update(str(x) for x in cr_sub["_norm_client"].dropna().unique() if str(x).strip() and str(x).lower() != "nan")
        opt_designers.update(str(x) for x in cr_sub["_norm_des"].dropna().unique() if str(x).strip() and str(x).lower() not in ("nan", "none"))
        opt_qcs.update(str(x) for x in cr_sub["_norm_qc"].dropna().unique() if str(x).strip() and str(x).lower() not in ("nan", "none"))
        opt_reasons.update(str(x) for x in cr_sub["_norm_rsn"].dropna().unique() if str(x).strip() and str(x).lower() not in ("nan", "none"))

    if not qa_sub.empty:
        if f_auditor and f_auditor != "ALL" and "auditor" in qa_sub.columns:
            qa_sub = qa_sub[qa_sub["auditor"] == f_auditor.lower()]
        if f_client and f_client != "ALL":
            qa_sub = qa_sub[qa_sub["_norm_client"] == f_client]
        if f_designer and f_designer != "ALL":
            qa_sub = qa_sub[qa_sub["_norm_des"] == f_designer]
        if f_qc and f_qc != "ALL":
            qa_sub = qa_sub[qa_sub["_norm_qc"] == f_qc]
        if f_reason and f_reason != "ALL":
            qa_sub = qa_sub[qa_sub["_norm_rsn_list"].apply(lambda lst: f_reason in lst)]

    if not cr_sub.empty:
        if f_client and f_client != "ALL":
            cr_sub = cr_sub[cr_sub["_norm_client"] == f_client]
        if f_designer and f_designer != "ALL":
            cr_sub = cr_sub[cr_sub["_norm_des"] == f_designer]
        if f_qc and f_qc != "ALL":
            cr_sub = cr_sub[cr_sub["_norm_qc"] == f_qc]
        if f_reason and f_reason != "ALL":
            cr_sub = cr_sub[cr_sub["_norm_rsn"] == f_reason]

    date_keys = sorted(set(
        ([str(d.date()) for d in qa_sub["Date"].dropna().unique()] if not qa_sub.empty else []) +
        ([str(d.date()) for d in cr_sub["_date"].dropna().unique()] if not cr_sub.empty else [])
    ))
    trend = []
    for dk in date_keys[-45:]:
        dt = pd.Timestamp(dk)
        q_day = qa_sub[qa_sub["Date"] == dt] if not qa_sub.empty else pd.DataFrame()
        c_day = cr_sub[cr_sub["_date"] == dt] if not cr_sub.empty else pd.DataFrame()
        q_imgs = int(q_day["imgs"].sum()) if not q_day.empty else 0
        q_err = int(q_day["err"].sum()) if not q_day.empty else 0
        c_ord = int(len(c_day)) if not c_day.empty else 0
        c_imgs = int(c_day["_vol"].sum()) if not c_day.empty else 0
        trend.append({
            "date": dk[5:],
            "full_date": dk,
            "qa_imgs": q_imgs,
            "qa_err": q_err,
            "qa_rate": round(pct(q_err, q_imgs), 2),
            "cr_ord": c_ord,
            "cr_imgs": c_imgs,
        })

    client_map = {}
    if not qa_sub.empty:
        for _, r in qa_sub.iterrows():
            cn = r["_norm_client"]
            client_map.setdefault(cn, {"client": cn, "qa_imgs": 0, "qa_err": 0, "cr_ord": 0, "cr_imgs": 0})
            client_map[cn]["qa_imgs"] += int(r["imgs"])
            client_map[cn]["qa_err"] += int(r["err"])
    if not cr_sub.empty:
        for _, r in cr_sub.iterrows():
            cn = r["_norm_client"]
            client_map.setdefault(cn, {"client": cn, "qa_imgs": 0, "qa_err": 0, "cr_ord": 0, "cr_imgs": 0})
            client_map[cn]["cr_ord"] += 1
            client_map[cn]["cr_imgs"] += int(r["_vol"])
    clients_chart = []
    for cn, v in client_map.items():
        v["qa_rate"] = round(pct(v["qa_err"], v["qa_imgs"]), 2)
        clients_chart.append(v)
    clients_chart.sort(key=lambda x: (-x["cr_ord"], -x["qa_err"], -x["qa_imgs"]))

    qa_prod, qa_cust = 0, 0
    if not qa_sub.empty and "Fault" in qa_sub.columns:
        f = qa_sub["Fault"].astype(str).str.strip().str.upper()
        qa_prod = int(qa_sub.loc[f == "PRODUCTION FAULT", "err"].sum())
        qa_cust = int(qa_sub.loc[f == "CUSTOMER FAULT", "err"].sum())

    cr_prod, cr_cust = 0, 0
    if not cr_sub.empty and "reason" in cr_sub.columns:
        f = cr_sub["reason"].astype(str).str.upper()
        cr_cust = int(f.str.contains("CUSTOMER", na=False).sum())
        cr_prod = int(len(cr_sub) - cr_cust)

    des_map = {}
    rsn_map = {}
    if not qa_sub.empty:
        for _, r in qa_sub[qa_sub["err"] > 0].iterrows():
            d = str(r["_norm_des"]).strip()
            if d and d.lower() not in ("nan", "none"):
                des_map.setdefault(d, {"name": d, "qa_err": 0, "cr_ord": 0})
                des_map[d]["qa_err"] += int(r["err"])
            for cat in r["_norm_rsn_list"]:
                rsn_map[cat] = rsn_map.get(cat, 0) + 1
    if not cr_sub.empty:
        for _, r in cr_sub.iterrows():
            d = str(r["_norm_des"]).strip()
            if d and d.lower() not in ("nan", "none"):
                des_map.setdefault(d, {"name": d, "qa_err": 0, "cr_ord": 0})
                des_map[d]["cr_ord"] += 1
            rtxt = str(r["_norm_rsn"]).strip()
            if rtxt and rtxt.lower() not in ("nan", "none"):
                rsn_map[rtxt] = rsn_map.get(rtxt, 0) + 1

    top_designers = sorted(des_map.values(), key=lambda x: -(x["qa_err"] + x["cr_ord"] * 3))[:10]
    top_reasons = [{"reason": k, "count": v} for k, v in sorted(rsn_map.items(), key=lambda x: -x[1])[:10]]

    tot_qa_imgs = int(qa_sub["imgs"].sum()) if not qa_sub.empty else 0
    tot_qa_err = int(qa_sub["err"].sum()) if not qa_sub.empty else 0
    tot_cr_ord = int(len(cr_sub)) if not cr_sub.empty else 0
    tot_cr_imgs = int(cr_sub["_vol"].sum()) if not cr_sub.empty else 0

    return {
        "span_label": span_label((start, end)),
        "start_date": str(start.date()),
        "end_date": str(end.date()),
        "kpis": {
            "qa_imgs": tot_qa_imgs,
            "qa_err": tot_qa_err,
            "qa_rate": round(pct(tot_qa_err, tot_qa_imgs), 2),
            "cr_ord": tot_cr_ord,
            "cr_imgs": tot_cr_imgs,
        },
        "options": {
            "clients": sorted(opt_clients),
            "designers": sorted(opt_designers),
            "qcs": sorted(opt_qcs),
            "reasons": sorted(opt_reasons),
            "auditors": sorted(opt_auditors),
        },
        "trend": trend,
        "clients": clients_chart[:12],
        "faults": {"qa_prod": qa_prod, "qa_cust": qa_cust, "cr_prod": cr_prod, "cr_cust": cr_cust},
        "top_designers": top_designers,
        "top_reasons": top_reasons,
    }


# ===================================================================== DOCX EXPORT HELPERS
DOCX_THEME = {
    "primary_hex": "1B2A4A",
    "primary_rgb": (27, 42, 74),
    "accent_hex": "2C5282",
    "accent_rgb": (44, 82, 130),
    "ink_rgb": (30, 41, 59),
    "zebra_hex": "F8FAFC",
    "callout_hex": "F1F5F9",
    "border_hex": "CBD5E1",
    "bad_rgb": (153, 27, 27),
    "good_rgb": (22, 101, 52),
}


def _set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(
        f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'
    ))


def _set_table_borders(table, border_color="CBD5E1"):
    table._tbl.tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{DOCX_THEME["primary_hex"]}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:left w:val="none"/><w:right w:val="none"/><w:insideV w:val="none"/></w:tblBorders>'
    ))


def _add_section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text.upper())
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(*DOCX_THEME["accent_rgb"])
    p._p.get_or_add_pPr().append(parse_xml(
        f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="3" color="{DOCX_THEME["border_hex"]}"/></w:pBdr>'
    ))


def _add_doc_banner(doc, title, subtitle):
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = banner.rows[0].cells[0]
    _set_cell_margins(cell, top=160, bottom=160, left=200, right=200)
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{DOCX_THEME["primary_hex"]}"/>'))
    p_title = cell.paragraphs[0]
    p_title.paragraph_format.space_after = Pt(3)
    r_title = p_title.add_run(title)
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(255, 255, 255)
    p_sub = cell.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(0)
    r_sub = p_sub.add_run(subtitle)
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = RGBColor(203, 213, 225)


def _style_table(table, headers, rows, right_align_cols=None, highlight_cols=None):
    right_align_cols = right_align_cols or set()
    highlight_cols = highlight_cols or {}
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(table, DOCX_THEME["border_hex"])
    hdr_cells = table.rows[0].cells
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for i, text in enumerate(headers):
        hdr_cells[i].text = ""
        _set_cell_margins(hdr_cells[i], top=110, bottom=110, left=120, right=120)
        hdr_cells[i]._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{DOCX_THEME["primary_hex"]}"/>'))
        p = hdr_cells[i].paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        if i in right_align_cols:
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(str(text))
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, row_data in enumerate(rows):
        row = table.add_row()
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        bg = DOCX_THEME["zebra_hex"] if r_idx % 2 == 1 else "FFFFFF"
        for i, val in enumerate(row_data):
            row.cells[i].text = ""
            _set_cell_margins(row.cells[i], top=85, bottom=85, left=120, right=120)
            row.cells[i]._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
            p = row.cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if i in right_align_cols:
                p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
            val_str = str(val)
            run = p.add_run(val_str)
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(*DOCX_THEME["ink_rgb"])
            if i in highlight_cols:
                mode = highlight_cols[i]
                if mode == "delta":
                    if val_str.startswith("+"):
                        run.bold = True
                        run.font.color.rgb = RGBColor(*DOCX_THEME["bad_rgb"])
                    elif val_str.startswith("-"):
                        run.bold = True
                        run.font.color.rgb = RGBColor(*DOCX_THEME["good_rgb"])


def _configure_doc_base(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(9.5)
    normal.font.color.rgb = RGBColor(*DOCX_THEME["ink_rgb"])
    for sec in doc.sections:
        sec.top_margin = Inches(0.55)
        sec.bottom_margin = Inches(0.55)
        sec.left_margin = Inches(0.55)
        sec.right_margin = Inches(0.55)


def _style_takeaways_table(table, rows):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(table, DOCX_THEME["border_hex"])
    hdr_cells = table.rows[0].cells
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    headers = ["Category", "Details"]
    widths = [Inches(2.1), Inches(5.3)]
    for i, text in enumerate(headers):
        hdr_cells[i].text = ""
        hdr_cells[i].width = widths[i]
        _set_cell_margins(hdr_cells[i], top=110, bottom=110, left=140, right=140)
        hdr_cells[i]._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{DOCX_THEME["primary_hex"]}"/>'))
        p = hdr_cells[i].paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, (cat, det) in enumerate(rows):
        row = table.add_row()
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        bg = DOCX_THEME["zebra_hex"] if r_idx % 2 == 1 else "FFFFFF"
        for i, val in enumerate([cat, det]):
            cell = row.cells[i]
            cell.text = ""
            cell.width = widths[i]
            _set_cell_margins(cell, top=105, bottom=105, left=140, right=140)
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(str(val))
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            if i == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(*DOCX_THEME["primary_rgb"])
            else:
                run.font.color.rgb = RGBColor(*DOCX_THEME["ink_rgb"])


def _add_client_attention_cards(doc, cards):
    for c in cards:
        status = str(c.get("status", "HEALTHY")).upper()
        if status == "CRITICAL":
            border_hex = "B91C1C"
            hdr_bg = "FDECEC"
            badge_bg = "B91C1C"
            badge_txt = "ACTION NEEDED"
        elif status == "WATCH":
            border_hex = "B45309"
            hdr_bg = "FFF4E0"
            badge_bg = "B45309"
            badge_txt = "WATCHLIST"
        else:
            border_hex = "15803D"
            hdr_bg = "E7F6EC"
            badge_bg = "15803D"
            badge_txt = "HEALTHY"

        tbl = doc.add_table(rows=2, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl._tbl.tblPr.append(parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'<w:insideV w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
            f'</w:tblBorders>'
        ))
        for r in tbl.rows:
            r._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        h_cell = tbl.rows[0].cells[0].merge(tbl.rows[0].cells[1])
        h_cell.text = ""
        _set_cell_margins(h_cell, top=95, bottom=95, left=140, right=140)
        h_cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hdr_bg}"/>'))
        p_h = h_cell.paragraphs[0]
        p_h.paragraph_format.space_after = Pt(2)

        r_badge = p_h.add_run(f" [{badge_txt}] ")
        r_badge.bold = True
        r_badge.font.size = Pt(8.5)
        r_badge.font.color.rgb = RGBColor.from_string(badge_bg)

        r_name = p_h.add_run(f"  {c.get('client', '')}   ")
        r_name.bold = True
        r_name.font.size = Pt(11.5)
        r_name.font.color.rgb = RGBColor(*DOCX_THEME["primary_rgb"])

        r_sc = p_h.add_run(f"Priority Score: {c.get('score', 0)}")
        r_sc.font.size = Pt(9)
        r_sc.font.color.rgb = RGBColor(100, 116, 139)

        flags = c.get("flags") or []
        if flags:
            p_fl = h_cell.add_paragraph()
            p_fl.paragraph_format.space_after = Pt(0)
            r_fl = p_fl.add_run(" • ".join(str(f) for f in flags))
            r_fl.font.size = Pt(8.5)
            r_fl.font.color.rgb = RGBColor(71, 85, 105)

        c_qa = tbl.rows[1].cells[0]
        c_qa.text = ""
        c_qa.width = Inches(3.7)
        _set_cell_margins(c_qa, top=100, bottom=100, left=140, right=140)
        c_qa._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))

        p = c_qa.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("INTERNAL QA AUDIT BREAKDOWN")
        r.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(*DOCX_THEME["accent_rgb"])

        qa_rate = float(c.get("qa_a_rate", 0))
        qa_b_rate = float(c.get("qa_b_rate", 0))
        qa_diff = round(qa_rate - qa_b_rate, 2)
        p = c_qa.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{qa_rate:.2f}%")
        r.bold = True
        r.font.size = Pt(14)
        r.font.color.rgb = RGBColor(*DOCX_THEME["bad_rgb"]) if qa_rate > ESCALATE_ABOVE else RGBColor(*DOCX_THEME["primary_rgb"])
        if abs(qa_diff) >= 0.01:
            sign = "+" if qa_diff > 0 else ""
            rd = p.add_run(f"   ({sign}{qa_diff:.2f}%)")
            rd.bold = True
            rd.font.size = Pt(9)
            rd.font.color.rgb = RGBColor(*DOCX_THEME["bad_rgb"]) if qa_diff > 0 else RGBColor(*DOCX_THEME["good_rgb"])

        p = c_qa.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(f"{int(c.get('qa_a_err', 0)):,} errors / {int(c.get('qa_a_imgs', 0)):,} imgs (Prior: {qa_b_rate:.2f}%)")
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(100, 116, 139)

        for lbl, key in [("Reasons: ", "qa_reasons"), ("Designers: ", "qa_designers"), ("QC Names: ", "qa_qcs")]:
            p = c_qa.add_paragraph()
            p.paragraph_format.space_after = Pt(1.5)
            rl = p.add_run(lbl)
            rl.bold = True
            rl.font.size = Pt(8.5)
            rl.font.color.rgb = RGBColor(*DOCX_THEME["ink_rgb"])
            vals = c.get(key) or []
            rv = p.add_run(", ".join(vals) if vals else "None")
            rv.font.size = Pt(8.5)
            rv.font.color.rgb = RGBColor(71, 85, 105) if vals else RGBColor(148, 163, 184)

        c_cr = tbl.rows[1].cells[1]
        c_cr.text = ""
        c_cr.width = Inches(3.7)
        _set_cell_margins(c_cr, top=100, bottom=100, left=140, right=140)
        c_cr._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))

        p = c_cr.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("EXTERNAL CLIENT REJECTIONS BREAKDOWN")
        r.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(180, 83, 9)

        cr_ord = int(c.get("cr_a_ord", 0))
        cr_b_ord = int(c.get("cr_b_ord", 0))
        cr_diff = cr_ord - cr_b_ord
        p = c_cr.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{cr_ord:,} orders")
        r.bold = True
        r.font.size = Pt(14)
        r.font.color.rgb = RGBColor(*DOCX_THEME["primary_rgb"])
        if cr_diff != 0:
            sign = "+" if cr_diff > 0 else ""
            rd = p.add_run(f"   ({sign}{cr_diff:,})")
            rd.bold = True
            rd.font.size = Pt(9)
            rd.font.color.rgb = RGBColor(*DOCX_THEME["bad_rgb"]) if cr_diff > 0 else RGBColor(*DOCX_THEME["good_rgb"])

        p = c_cr.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(
            f"{int(c.get('cr_a_imgs', 0)):,} imgs · Prod: {int(c.get('cr_a_prod', 0))} / Cust: {int(c.get('cr_a_cust', 0))} (Prior: {cr_b_ord} ord)"
        )
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(100, 116, 139)

        for lbl, key in [("Reasons: ", "cr_reasons"), ("Designers: ", "cr_designers"), ("QC Names: ", "cr_qcs")]:
            p = c_cr.add_paragraph()
            p.paragraph_format.space_after = Pt(1.5)
            rl = p.add_run(lbl)
            rl.bold = True
            rl.font.size = Pt(8.5)
            rl.font.color.rgb = RGBColor(*DOCX_THEME["ink_rgb"])
            vals = c.get(key) or []
            rv = p.add_run(", ".join(vals) if vals else "None")
            rv.font.size = Pt(8.5)
            rv.font.color.rgb = RGBColor(71, 85, 105) if vals else RGBColor(148, 163, 184)

        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(6)


def generate_generic_docx(title, subtitle, kpi_headers, kpi_row, summary_text, table_title, table_headers, table_rows, right_cols=None, takeaways=None, attn_cards=None):
    if not HAS_DOCX:
        raise RuntimeError("python-docx is not installed. Run: pip install python-docx")
    doc = docx.Document()
    _configure_doc_base(doc)
    _add_doc_banner(doc, title, subtitle)

    sec_num = 1
    if kpi_headers and kpi_row:
        _add_section_heading(doc, f"{sec_num}. Executive Summary Metrics")
        sec_num += 1
        t_kpi = doc.add_table(rows=1, cols=len(kpi_headers))
        t_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(t_kpi, DOCX_THEME["border_hex"])
        for i, htxt in enumerate(kpi_headers):
            c = t_kpi.rows[0].cells[i]
            c.text = ""
            _set_cell_margins(c, top=90, bottom=90, left=120, right=120)
            c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{DOCX_THEME["primary_hex"]}"/>'))
            p = c.paragraphs[0]
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(htxt).upper())
            r.bold = True
            r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(255, 255, 255)
        vrow = t_kpi.add_row()
        for i, val in enumerate(kpi_row):
            c = vrow.cells[i]
            c.text = ""
            _set_cell_margins(c, top=110, bottom=110, left=120, right=120)
            c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{DOCX_THEME["callout_hex"]}"/>'))
            p = c.paragraphs[0]
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(val))
            r.bold = True
            r.font.size = Pt(13)
            r.font.color.rgb = RGBColor(*DOCX_THEME["primary_rgb"])

    if attn_cards:
        _add_section_heading(doc, f"{sec_num}. Client Attention Breakdown Cards ({len(attn_cards)})")
        sec_num += 1
        _add_client_attention_cards(doc, attn_cards)

    if table_rows and not attn_cards:
        _add_section_heading(doc, f"{sec_num}. {table_title}")
        sec_num += 1
        t_ord = doc.add_table(rows=1, cols=len(table_headers))
        _style_table(t_ord, table_headers, table_rows, right_align_cols=set(right_cols or []))

    if takeaways:
        _add_section_heading(doc, f"{sec_num}. Key Takeaways & Action Plan")
        sec_num += 1
        t_tk = doc.add_table(rows=1, cols=2)
        _style_takeaways_table(t_tk, takeaways)

    if summary_text:
        _add_section_heading(doc, f"{sec_num}. Audit Summary & Notes")
        box = doc.add_table(rows=1, cols=1)
        box.alignment = WD_TABLE_ALIGNMENT.CENTER
        bcell = box.rows[0].cells[0]
        _set_cell_margins(bcell, top=120, bottom=120, left=160, right=160)
        bcell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{DOCX_THEME["callout_hex"]}"/>'))
        bcell._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="18" w:space="0" w:color="{DOCX_THEME["accent_hex"]}"/>'
            f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>'
        ))
        first_p = True
        for line in summary_text.splitlines():
            clean_line = re.sub(r"^[^\w\d]+", "", line).strip() if line.strip().startswith("📊") else line.strip()
            if not clean_line:
                continue
            p = bcell.paragraphs[0] if first_p else bcell.add_paragraph()
            first_p = False
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(clean_line)
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(*DOCX_THEME["ink_rgb"])

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ===================================================================== SERVER & STATE
STATE = {
    "qa_df": None, "qa_source": "", "qa_error": None,
    "cr_zib": None, "cr_oth": None, "cr_source": "", "cr_error": None,
    "prod_df": None, "prod_source": "", "prod_error": None,
    "loading": True,
}
HEARTBEAT = {"last": 0.0, "seen": False}


def _load_all_state():
    STATE.update(loading=True, qa_error=None, cr_error=None, prod_error=None)
    folder = app_folder()
    xlsx_files = sorted((p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")), key=lambda p: p.stat().st_mtime, reverse=True)

    qa_df, qa_src = None, ""
    if QA_ONLINE_URL:
        try:
            qa_df, qa_src = load_qa_online(QA_ONLINE_URL), "QA: Google Sheet"
        except Exception as e:
            STATE["qa_error"] = str(e)
    if qa_df is None:
        for f in xlsx_files:
            try:
                xl = pd.ExcelFile(f)
                if QA_SHEET in xl.sheet_names:
                    qa_df = prepare_qa(pd.read_excel(xl, sheet_name=QA_SHEET))
                    qa_src = f"QA: {f.name}"
                    STATE["qa_error"] = None
                    break
            except Exception:
                pass
    STATE["qa_df"] = qa_df
    STATE["qa_source"] = qa_src or "QA: Not loaded"

    cr_z, cr_o, cr_src = None, None, ""
    if CR_ONLINE_URL:
        try:
            (cr_z, cr_o), cr_src = load_cr_online(CR_ONLINE_URL), "CR: Google Sheet"
        except Exception as e:
            STATE["cr_error"] = str(e)
    if cr_z is None:
        for f in xlsx_files:
            try:
                xl = pd.ExcelFile(f)
                if CR_ZIBBER_SHEET in xl.sheet_names and CR_OTHER_SHEET in xl.sheet_names:
                    cr_z, cr_o = read_cr_book(f)
                    cr_src = f"CR: {f.name}"
                    STATE["cr_error"] = None
                    break
            except Exception:
                pass
    STATE["cr_zib"] = cr_z
    STATE["cr_oth"] = cr_o
    STATE["cr_source"] = cr_src or "CR: Not loaded"

    prod_df, prod_src = None, ""
    if PROD_ONLINE_URL:
        try:
            prod_df, prod_src = load_prod_online(PROD_ONLINE_URL), "Prod: Google Sheet"
        except Exception as e:
            STATE["prod_error"] = str(e)
    if prod_df is None:
        for f in xlsx_files:
            try:
                xl = pd.ExcelFile(f)
                if PROD_SHEET in xl.sheet_names:
                    prod_df = prepare_prod(pd.read_excel(xl, sheet_name=PROD_SHEET))
                    prod_src = f"Prod: {f.name}"
                    STATE["prod_error"] = None
                    break
            except Exception:
                pass
    STATE["prod_df"] = prod_df
    STATE["prod_source"] = prod_src or "Prod: Not loaded"
    STATE["loading"] = False


def serve():
    token = secrets.token_urlsafe(16)
    HEARTBEAT["last"] = time.time()
    threading.Thread(target=_load_all_state, daemon=True).start()

    def status():
        qa_df = STATE["qa_df"]
        cr_z, cr_o = STATE["cr_zib"], STATE["cr_oth"]
        prod_df = STATE["prod_df"]
        today = pd.Timestamp.today().normalize()
        out = {
            "loading": STATE["loading"],
            "qa_loaded": qa_df is not None,
            "qa_source": STATE["qa_source"],
            "qa_rows": len(qa_df) if qa_df is not None else 0,
            "qa_error": STATE["qa_error"],
            "cr_loaded": cr_z is not None and cr_o is not None,
            "cr_source": STATE["cr_source"],
            "cr_rows": (len(cr_z) + len(cr_o)) if (cr_z is not None and cr_o is not None) else 0,
            "cr_error": STATE["cr_error"],
            "prod_loaded": prod_df is not None,
            "prod_source": STATE["prod_source"],
            "prod_rows": len(prod_df) if prod_df is not None else 0,
            "prod_error": STATE["prod_error"],
            "has_docx": HAS_DOCX,
        }
        if qa_df is not None and not qa_df.empty:
            names = qa_df["auditor"].value_counts().index.tolist()
            aud_list = []
            for n in names:
                if not n or n == "nan":
                    continue
                sub_dates = sorted(qa_df[(qa_df["auditor"] == n) & (qa_df["Date"] <= today)]["Date"].unique())
                latest = str(sub_dates[-1].date()) if len(sub_dates) >= 1 else None
                prev_latest = str(sub_dates[-2].date()) if len(sub_dates) >= 2 else latest
                aud_list.append({"name": n, "latest": latest, "prev_latest": prev_latest})
            out["auditors"] = aud_list
            out["default_auditor"] = "sadik" if "sadik" in names else (names[0] if names else "")

        if cr_z is not None and cr_o is not None:
            all_dates = sorted(pd.concat([cr_z["_date"], cr_o["_date"]]).dropna().unique())
            past_dates = [d for d in all_dates if d <= today]
            out["cr_latest"] = str(past_dates[-1].date()) if len(past_dates) >= 1 else str(today.date())
            out["cr_prev_latest"] = str(past_dates[-2].date()) if len(past_dates) >= 2 else out["cr_latest"]

        if prod_df is not None and not prod_df.empty:
            p_dates = sorted(prod_df["_date"].dropna().unique())
            past_p = [d for d in p_dates if d <= today]
            out["prod_latest"] = str(past_p[-1].date()) if len(past_p) >= 1 else (str(p_dates[-1].date()) if p_dates else str(today.date()))
            out["prod_specs"] = sorted(prod_df["_spec"].dropna().unique().tolist())
        return out

    def _cr_snapshot_for_span(start, end):
        zibber, other = STATE["cr_zib"], STATE["cr_oth"]
        if zibber is None or other is None:
            return {"orders": 0, "images": 0, "prod": 0, "cust": 0, "delayed": 0, "orders_list": []}
        z = zibber[(zibber["_date"] >= start) & (zibber["_date"] <= end)]
        o = other[(other["_date"] >= start) & (other["_date"] <= end)]
        all_rows = pd.concat([z, o], ignore_index=True)
        olist = extract_cr_orders(all_rows)
        return {
            "orders": len(olist),
            "images": sum(x["volume"] for x in olist),
            "prod": sum(1 for x in olist if not x["is_customer"]),
            "cust": sum(1 for x in olist if x["is_customer"]),
            "delayed": sum(1 for x in olist if x["delay"]),
            "orders_list": olist,
        }

    def _qa_snapshot_for_span(start, end, auditor=None):
        df = STATE["qa_df"]
        if df is None:
            return {"orders": 0, "images": 0, "errors": 0, "rate": 0.0, "clean": 0, "orders_list": []}
        d = df[(df["Date"] >= start) & (df["Date"] <= end)]
        if auditor and auditor.strip().lower() not in ("", "all"):
            d = d[d["auditor"] == auditor.strip().lower()]
        olist = extract_qa_orders(d)
        imgs = sum(x["volume"] for x in olist)
        errs = sum(x["errors"] for x in olist)
        return {
            "orders": len(olist),
            "images": imgs,
            "errors": errs,
            "rate": round(pct(errs, imgs), 2),
            "clean": sum(1 for x in olist if x["errors"] == 0),
            "orders_list": olist,
        }

    def api_qa_report(body):
        df = STATE["qa_df"]
        if df is None:
            return {"error": STATE["qa_error"] or "QA Audit data not available."}
        auditor = (body.get("auditor") or "").strip()
        try:
            start = pd.Timestamp(body["start"]).normalize()
            end = pd.Timestamp(body.get("end") or body["start"]).normalize()
        except Exception:
            return {"error": "Invalid date."}
        if end < start:
            return {"error": "End date cannot be before start date."}
        span = (start, end)
        text = build_qa_report_text(df, span, auditor)
        if text is None:
            return {"empty": True, "error": f"No QA audits found for {auditor} on {span_label(span)}."}

        d = df[(df["Date"] >= start) & (df["Date"] <= end) & (df["auditor"] == auditor.lower())]
        T, E = int(d["imgs"].sum()), int(d["err"].sum())
        clients = []
        for spec in dict.fromkeys(d["Spec ID"].astype(int)):
            s = d[d["Spec ID"].astype(int) == spec]
            clients.append({"name": qa_client_name(spec), "orders": len(s)})

        f_client = body.get("client", "ALL")
        f_designer = body.get("designer", "ALL")
        f_qc = body.get("qc", "ALL")
        f_reason = body.get("reason", "ALL")
        alerts = qa_streak_alerts(
            df, end, auditor=auditor, start=start,
            f_client=f_client, f_designer=f_designer, f_qc=f_qc, f_reason=f_reason
        )
        return {
            "text": text,
            "label": span_label(span),
            "filename": f"QA_Report_{start:%d%m%Y}.txt",
            "summary": {
                "images": T, "errors": E, "rate": round(pct(E, T), 2),
                "orders": len(d), "clean": int((d["err"] == 0).sum()), "days": int(d["Date"].nunique()),
                "clients": clients,
                "orders_list": extract_qa_orders(d),
                "takeaways": build_qa_takeaways(df, span, auditor),
                "escalate_above": ESCALATE_ABOVE,
                "streak_days": STREAK_DAYS,
                "streak_alerts": alerts,
                "alerts": [a["message"] for a in alerts],
                "paired_cr": _cr_snapshot_for_span(start, end),
            },
        }

    def api_cr_report(body):
        zibber, other = STATE["cr_zib"], STATE["cr_oth"]
        if zibber is None or other is None:
            return {"error": STATE["cr_error"] or "Client Rejection data not available."}
        try:
            start = pd.Timestamp(body["start"]).normalize()
            end = pd.Timestamp(body.get("end") or body["start"]).normalize()
        except Exception:
            return {"error": "Invalid date."}
        if end < start:
            return {"error": "End date cannot be before start date."}
        span = (start, end)
        text = build_cr_report_text(zibber, other, span)
        if text is None:
            return {"empty": True, "error": f"No client rejections found on {span_label(span)}."}

        z = zibber[(zibber["_date"] >= start) & (zibber["_date"] <= end)]
        o = other[(other["_date"] >= start) & (other["_date"] <= end)]
        all_rows = pd.concat([z, o], ignore_index=True)
        orders_list = extract_cr_orders(all_rows)
        clients = sorted(list(set(x["client"] for x in orders_list)))
        f_client = body.get("client", "ALL")
        f_designer = body.get("designer", "ALL")
        f_qc = body.get("qc", "ALL")
        f_reason = body.get("reason", "ALL")
        alerts = cr_streak_alerts(
            zibber, other, end, start=start,
            f_client=f_client, f_designer=f_designer, f_qc=f_qc, f_reason=f_reason
        )

        return {
            "text": text,
            "label": span_label(span),
            "filename": f"CR_Report_{start:%d%m%Y}.txt",
            "summary": {
                "days": int(all_rows["_date"].nunique()),
                "clients": clients,
                "orders_list": orders_list,
                "streak_days": STREAK_DAYS,
                "streak_alerts": alerts,
                "paired_qa": _qa_snapshot_for_span(start, end),
            },
        }

    def api_group_data(body):
        df = STATE["prod_df"]
        if df is None:
            return {"error": STATE["prod_error"] or "Productivity data sheet not available."}
        try:
            start = pd.Timestamp(body["start"]).normalize()
            end = pd.Timestamp(body.get("end") or body["start"]).normalize()
        except Exception:
            return {"error": "Invalid date."}
        if end < start:
            return {"error": "End date cannot be before start date."}
        res = build_group_data(
            df, (start, end),
            f_team=body.get("team", "ALL"),
            f_desig=body.get("designation", "ALL"),
            f_designer=body.get("designer", "ALL"),
            f_spec=body.get("spec", "ALL"),
        )
        if res is None:
            return {"empty": True, "error": f"No designer productivity records found on {span_label((start, end))}."}
        return res

    def api_specwise_data(body):
        df = STATE["prod_df"]
        if df is None:
            return {"error": STATE["prod_error"] or "Productivity data sheet not available."}
        try:
            start = pd.Timestamp(body["start"]).normalize()
            end = pd.Timestamp(body.get("end") or body["start"]).normalize()
        except Exception:
            return {"error": "Invalid date."}
        if end < start:
            return {"error": "End date cannot be before start date."}
        res = build_specwise_data(
            df, (start, end),
            f_spec=body.get("spec", "ALL"),
            f_team=body.get("team", "ALL"),
            f_desig=body.get("designation", "ALL"),
            f_designer=body.get("designer", "ALL"),
        )
        if res is None:
            return {"empty": True, "error": f"No spec-wise productivity records found on {span_label((start, end))}."}
        return res

    def api_attention(body):
        b = body or {}
        return build_client_attention(
            STATE["qa_df"],
            STATE["cr_zib"],
            STATE["cr_oth"],
            start_str=b.get("start"),
            end_str=b.get("end"),
            f_auditor=b.get("auditor", "ALL"),
        )

    def api_charts(body):
        return build_charts_data(
            STATE["qa_df"],
            STATE["cr_zib"],
            STATE["cr_oth"],
            days_back=int(body.get("days", 30)),
            start_str=body.get("start"),
            end_str=body.get("end"),
            f_client=body.get("client", "ALL"),
            f_designer=body.get("designer", "ALL"),
            f_qc=body.get("qc", "ALL"),
            f_reason=body.get("reason", "ALL"),
            f_auditor=body.get("auditor", "ALL"),
        )

    def api_compare(body):
        mode = (body.get("mode") or "QA").upper()
        try:
            a1 = pd.Timestamp(body["a_start"]).normalize()
            a2 = pd.Timestamp(body.get("a_end") or body["a_start"]).normalize()
            b1 = pd.Timestamp(body["b_start"]).normalize()
            b2 = pd.Timestamp(body.get("b_end") or body["b_start"]).normalize()
        except Exception:
            return {"error": "Select valid dates for both periods."}
        if a2 < a1 or b2 < b1:
            return {"error": "End date cannot be earlier than start date."}

        if mode == "QA":
            df = STATE["qa_df"]
            if df is None:
                return {"error": "QA data not loaded."}
            aud = (body.get("auditor") or "ALL").strip().lower()
            base = df if aud in ("", "all") else df[df["auditor"] == aud]
            dfa = base[(base["Date"] >= a1) & (base["Date"] <= a2)]
            dfb = base[(base["Date"] >= b1) & (base["Date"] <= b2)]
            return {
                "mode": "QA",
                "a_label": span_label((a1, a2)),
                "b_label": span_label((b1, b2)),
                "orders_a": extract_qa_orders(dfa),
                "orders_b": extract_qa_orders(dfb),
                "escalate_above": ESCALATE_ABOVE,
                "paired_cr_a": _cr_snapshot_for_span(a1, a2),
                "paired_cr_b": _cr_snapshot_for_span(b1, b2),
            }
        else:
            zib, oth = STATE["cr_zib"], STATE["cr_oth"]
            if zib is None or oth is None:
                return {"error": "Client Rejection data not loaded."}
            base = pd.concat([zib, oth], ignore_index=True)
            dfa = base[(base["_date"] >= a1) & (base["_date"] <= a2)]
            dfb = base[(base["_date"] >= b1) & (base["_date"] <= b2)]
            return {
                "mode": "CR",
                "a_label": span_label((a1, a2)),
                "b_label": span_label((b1, b2)),
                "orders_a": extract_cr_orders(dfa),
                "orders_b": extract_cr_orders(dfb),
                "paired_qa_a": _qa_snapshot_for_span(a1, a2),
                "paired_qa_b": _qa_snapshot_for_span(b1, b2),
            }

    def api_export_docx(body):
        try:
            raw = generate_generic_docx(
                body.get("title", "Quality Report"),
                body.get("subtitle", ""),
                body.get("kpi_headers", []),
                body.get("kpi_row", []),
                body.get("summary_text", ""),
                body.get("table_title", "Details"),
                body.get("table_headers", []),
                body.get("table_rows", []),
                body.get("right_cols", []),
                body.get("takeaways", []),
                body.get("attn_cards", []),
            )
            return {"b64": base64.b64encode(raw).decode("ascii")}
        except Exception as e:
            return {"error": str(e)}

    def api_export_xlsx(body):
        try:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                k_hdr = body.get("kpi_headers", [])
                k_row = body.get("kpi_row", [])
                if k_hdr and k_row:
                    pd.DataFrame([k_row], columns=k_hdr).to_excel(writer, sheet_name="Executive KPIs", index=False)
                tk = body.get("takeaways", [])
                if tk:
                    pd.DataFrame(tk, columns=["Category", "Details"]).to_excel(writer, sheet_name="Key Takeaways & Action Plan", index=False)
                t_hdr = body.get("table_headers", [])
                t_rows = body.get("table_rows", [])
                if t_hdr and t_rows:
                    pd.DataFrame(t_rows, columns=t_hdr).to_excel(writer, sheet_name="Detailed Data", index=False)
            return {"b64": base64.b64encode(buf.getvalue()).decode("ascii")}
        except Exception as e:
            return {"error": str(e)}

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, obj, code=200, ctype="application/json"):
            data = obj.encode("utf-8") if isinstance(obj, str) else json.dumps(obj).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype + "; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(data)

        def _ok(self):
            return self.headers.get("X-Token") == token

        def do_GET(self):
            if self.path.split("?")[0] == "/":
                self._send(HTML_PAGE.replace("__TOKEN__", token), ctype="text/html")
            elif not self._ok():
                self._send({"error": "forbidden"}, 403)
            elif self.path == "/ping":
                HEARTBEAT.update(last=time.time(), seen=True)
                self._send({"ok": True})
            elif self.path == "/api/status":
                self._send(status())
            else:
                self._send({"error": "not found"}, 404)

        def do_POST(self):
            if not self._ok():
                return self._send({"error": "forbidden"}, 403)
            n = int(self.headers.get("Content-Length") or 0)
            try:
                body = json.loads(self.rfile.read(n) or b"{}")
            except Exception:
                body = {}
            routes = {
                "/api/qa_report": api_qa_report,
                "/api/cr_report": api_cr_report,
                "/api/group_data": api_group_data,
                "/api/specwise_data": api_specwise_data,
                "/api/attention": api_attention,
                "/api/charts": api_charts,
                "/api/compare": api_compare,
                "/api/export_docx": api_export_docx,
                "/api/export_xlsx": api_export_xlsx,
            }
            if self.path in routes:
                try:
                    self._send(routes[self.path](body))
                except Exception as e:
                    self._send({"error": str(e)})
            elif self.path == "/api/reload":
                threading.Thread(target=_load_all_state, daemon=True).start()
                self._send({"ok": True})
            elif self.path == "/api/quit":
                self._send({"ok": True})
                threading.Timer(0.4, lambda: os._exit(0)).start()
            else:
                self._send({"error": "not found"}, 404)

    port = int(os.environ.get("PORT", 8080))
    is_cloud = "PORT" in os.environ
    host = "0.0.0.0" if is_cloud else "127.0.0.1"
    server = ThreadingHTTPServer((host, port), H)
    url = f"http://{host}:{server.server_address[1]}/"

    if not is_cloud:
        def watchdog():
            while True:
                time.sleep(3)
                idle = time.time() - HEARTBEAT["last"]
                if (HEARTBEAT["seen"] and idle > 25) or (not HEARTBEAT["seen"] and idle > 180):
                    os._exit(0)

        threading.Thread(target=watchdog, daemon=True).start()
        try:
            webbrowser.open(f"http://127.0.0.1:{port}/")
        except Exception:
            pass

    print(f"Unified Quality & Client Analytics Suite running at {url}", flush=True)
    server.serve_forever()


HTML_PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Quality &amp; Client Analytics Suite</title>
<style>
:root{--bg:#f3f5f9;--surface:#fff;--ink:#0f172a;--muted:#64748b;--line:#e2e8f0;--brand:#1e40af;--brand-ink:#fff;
--brand-soft:#e8eefc;--bad:#b91c1c;--bad-soft:#fdecec;--good:#15803d;--good-soft:#e7f6ec;--warn:#b45309;--warn-soft:#fff4e0;
--shadow:0 1px 2px rgba(15,23,42,.06),0 4px 14px rgba(15,23,42,.05)}
@media (prefers-color-scheme:dark){:root{--bg:#0b1220;--surface:#121b2e;--ink:#e8edf6;--muted:#93a1b8;--line:#243049;
--brand:#6d9bff;--brand-ink:#0b1220;--brand-soft:#17254a;--bad:#ff8a8a;--bad-soft:#3a1717;--good:#5fd38a;--good-soft:#122a1c;
--warn:#f5b45c;--warn-soft:#33260f;--shadow:none}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 "Segoe UI",system-ui,-apple-system,Roboto,sans-serif}
header{background:#0f1d3d;color:#fff;padding:13px 26px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.logo{width:34px;height:34px;border-radius:9px;background:linear-gradient(135deg,#2f5be0,#0ea5e9);display:grid;place-items:center;font-weight:800;letter-spacing:.5px}
header h1{font-size:16.5px;margin:0;font-weight:700}
header .sub{font-size:11.5px;opacity:.72}
.grow{flex:1}
.badge{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.1);border-radius:99px;padding:4px 11px;font-size:11.5px}
.dot{width:7px;height:7px;border-radius:50%;background:#7ee2a0}.dot.off{background:#f5b45c}.dot.err{background:#ff8a8a}
.hbtn{background:rgba(255,255,255,.12);color:#fff;border:0;border-radius:8px;padding:6px 11px;font:inherit;font-size:12px;cursor:pointer}
.hbtn:hover{background:rgba(255,255,255,.22)}
.hbtn.accent{background:linear-gradient(135deg,#2563eb,#0ea5e9);font-weight:600}
.nav-tabs{background:#16274f;padding:0 26px;display:flex;gap:4px;border-top:1px solid rgba(255,255,255,.08);overflow-x:auto}
.nav-tab{background:transparent;border:0;color:rgba(255,255,255,.72);padding:11px 16px;font:inherit;font-size:13px;font-weight:600;cursor:pointer;border-bottom:3px solid transparent;white-space:nowrap}
.nav-tab:hover{color:#fff}
.nav-tab.on{color:#fff;border-bottom-color:#6d9bff;background:rgba(255,255,255,.08)}
main{max-width:1340px;margin:0 auto;padding:20px 24px 44px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);padding:18px 20px}
.controls{display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end}
.field{display:flex;flex-direction:column;gap:5px}
label,.lbl{font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
input,select{font:inherit;color:var(--ink);background:var(--surface);border:1px solid var(--line);border-radius:99px;padding:7px 14px;min-width:140px}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:99px;overflow:hidden}
.seg button{border:0;background:var(--surface);color:var(--muted);padding:7px 13px;font:inherit;font-size:13px;cursor:pointer}
.seg button.on{background:var(--brand-soft);color:var(--brand);font-weight:600}
.btn{border:0;border-radius:99px;padding:8px 18px;font:inherit;font-weight:600;cursor:pointer}
.btn.primary{background:var(--brand);color:var(--brand-ink)}.btn.primary:hover{filter:brightness(1.08)}
.btn.ghost{background:var(--brand-soft);color:var(--brand)}.btn.ghost:hover{filter:brightness(.96)}
.btn.sm{padding:5px 12px;font-size:12px}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;align-items:center}
.chip{border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:99px;padding:5px 13px;font:inherit;font-size:12.5px;cursor:pointer}
.chip:hover,.chip.on{border-color:var(--brand);color:var(--brand);background:var(--brand-soft);font-weight:600}
.filter-bar{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end;margin-bottom:14px;padding:14px 16px}
.filter-bar .field{flex:1;min-width:155px}
.filter-bar select{width:100%;min-width:0}
.hide{display:none!important}
.title{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:18px 2px 12px}
.title h2{margin:0;font-size:18px}.title span{color:var(--muted)}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.kpis5{display:grid;grid-template-columns:repeat(5,1fr);gap:14px}
.kpi .k{font-size:11px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.5px}
.kpi .v{font-size:24px;font-weight:700;margin-top:4px;line-height:1.15;display:flex;align-items:baseline;gap:8px}
.kpi .s{font-size:12px;color:var(--muted);margin-top:4px}
.delta{font-size:11.5px;font-weight:700;padding:2px 8px;border-radius:99px}
.delta.up-bad{background:var(--bad-soft);color:var(--bad)}
.delta.down-good{background:var(--good-soft);color:var(--good)}
.delta.neutral{background:var(--bg);color:var(--muted)}
.bad{color:var(--bad)}.good{color:var(--good)}.warnc{color:var(--warn)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-top:14px}
.prod-layout{display:grid;grid-template-columns:270px 1fr;gap:14px;margin-top:14px;align-items:start}
.vs-grid{display:grid;grid-template-columns:1fr auto 1fr;gap:14px;align-items:center;margin-top:14px}
.vs-box{background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:12px 16px}
.vs-badge{font-weight:800;font-size:13px;color:var(--brand);background:var(--brand-soft);padding:6px 12px;border-radius:99px}
h3{margin:0 0 12px;font-size:12.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;display:flex;justify-content:space-between;align-items:center}
table{width:100%;border-collapse:collapse}
th{font-size:11px;text-align:left;color:var(--muted);font-weight:700;padding:0 8px 8px 0;border-bottom:1px solid var(--line);text-transform:uppercase;white-space:nowrap}
td{padding:8px 8px 8px 0;border-bottom:1px solid var(--line);vertical-align:middle;font-size:13px}
tr:last-child td{border-bottom:0}
.num{text-align:right;font-variant-numeric:tabular-nums}th.num{text-align:right}
.cell-good{background:var(--good-soft);color:var(--good);font-weight:700;padding:3px 8px;border-radius:6px;display:inline-block}
.cell-bad{background:var(--bad-soft);color:var(--bad);font-weight:700;padding:3px 8px;border-radius:6px;display:inline-block}
.cell-warn{background:var(--warn-soft);color:var(--warn);font-weight:700;padding:3px 8px;border-radius:6px;display:inline-block}
.pill{display:inline-block;border-radius:99px;padding:2px 9px;font-size:11px;font-weight:700;letter-spacing:.3px}
.pill.bad{background:var(--bad-soft);color:var(--bad)}
.pill.good{background:var(--good-soft);color:var(--good)}
.pill.warn{background:var(--warn-soft);color:var(--warn)}
.team-pill{display:inline-block;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700}
.team-blue{background:#dbeafe;color:#1e40af}
.team-red{background:#fee2e2;color:#991b1b}
.team-green{background:#dcfce7;color:#166534}
.team-other{background:var(--brand-soft);color:var(--brand)}
.attn-card{border-left:4px solid var(--line);transition:transform .12s}
.attn-card.CRITICAL{border-left-color:var(--bad);background:linear-gradient(to right,var(--bad-soft),var(--surface) 22%)}
.attn-card.WATCH{border-left-color:var(--warn);background:linear-gradient(to right,var(--warn-soft),var(--surface) 22%)}
.attn-card.HEALTHY{border-left-color:var(--good)}
.streak-banner{border-left:5px solid var(--bad);background:linear-gradient(to right,var(--bad-soft),var(--surface) 35%);margin-bottom:14px}
.streak-item{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:7px 10px;border-radius:8px;background:var(--surface);border:1px solid var(--line);margin-top:6px;font-size:13px}
pre{margin:0;white-space:pre-wrap;word-break:break-word;font:13px/1.55 Consolas,"Cascadia Mono","DejaVu Sans Mono",monospace;
background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:16px 18px;max-height:480px;overflow:auto}
.note{color:var(--muted);font-size:12px}
.empty,.loading{text-align:center;padding:42px 20px;color:var(--muted)}
.spin{width:26px;height:26px;border:3px solid var(--line);border-top-color:var(--brand);border-radius:50%;animation:s 1s linear infinite;margin:0 auto 12px}
@keyframes s{to{transform:rotate(360deg)}}
.chart-box{width:100%;overflow-x:auto}
svg.chart{width:100%;height:auto;display:block;font-family:inherit}
.bar-row{display:flex;align-items:center;gap:10px;margin-bottom:9px}
.bar-lbl{width:165px;font-size:12.5px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar-track{flex:1;height:12px;background:var(--bg);border-radius:99px;overflow:hidden;border:1px solid var(--line)}
.bar-fill{height:100%;border-radius:99px;background:var(--brand)}
.bar-fill.bad{background:var(--bad)}
.bar-fill.warn{background:var(--warn)}
.bar-val{width:70px;text-align:right;font-size:12px;font-weight:700;font-variant-numeric:tabular-nums}
.exp-wrap{position:relative;display:inline-flex;gap:8px;align-items:center}
.exp-menu{position:absolute;right:0;top:calc(100% + 6px);min-width:230px;background:var(--surface);border:1px solid var(--line);border-radius:12px;box-shadow:0 10px 28px rgba(15,23,42,.18);padding:6px;z-index:50;display:none}
.exp-menu.open{display:block}
.exp-item{width:100%;text-align:left;background:transparent;border:0;color:var(--ink);padding:8px 12px;border-radius:8px;font:inherit;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:8px}
.exp-item:hover{background:var(--brand-soft);color:var(--brand)}
.tk-table th{background:#1b2a4a;color:#fff;padding:9px 12px;font-size:12px;text-transform:none;letter-spacing:0}
.tk-table td{padding:10px 12px;font-size:13px;border-bottom:1px solid var(--line);vertical-align:top}
.tk-table td:first-child{font-weight:700;color:var(--brand);width:220px;white-space:nowrap}
.tk-inp,.tk-txt{width:100%;font:inherit;font-size:13px;color:var(--ink);background:var(--bg);border:1px solid var(--brand);border-radius:8px;padding:6px 10px}
.tk-inp{font-weight:700;color:var(--brand)}
.tk-txt{min-height:48px;resize:vertical;line-height:1.4}
.del-row-btn{background:var(--bad-soft);color:var(--bad);border:0;border-radius:6px;padding:4px 8px;font-size:11.5px;font-weight:700;cursor:pointer}
.modal-bg{position:fixed;inset:0;background:rgba(15,23,42,.72);backdrop-filter:blur(4px);z-index:120;display:none;align-items:center;justify-content:center;padding:20px}
.modal-bg.open{display:flex}
.modal-card{background:var(--surface);border:1px solid var(--line);border-radius:16px;max-width:960px;width:100%;max-height:92vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 24px 60px rgba(0,0,0,.4)}
.modal-hdr{padding:14px 20px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:10px;background:var(--bg)}
.modal-body{padding:18px;overflow:auto;display:grid;place-items:center;background:#0b1120}
.modal-body img{max-width:100%;height:auto;border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,.5)}
#toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%) translateY(20px);background:#0f1d3d;color:#fff;padding:10px 18px;border-radius:99px;opacity:0;transition:.25s;pointer-events:none;font-weight:600;z-index:150}
#toast.show{opacity:1;transform:translateX(-50%)}
@media (max-width:980px){.kpis,.kpis5,.grid3,.grid2,.vs-grid,.prod-layout{grid-template-columns:1fr}header,.nav-tabs{padding:12px 16px}main{padding:16px}}
</style></head><body>

<header>
  <div class="logo">QA+</div>
  <div><h1>Quality &amp; Client Analytics Suite</h1><div class="sub">Internal QA Audits · External Rejections · Client Attention Radar · Group &amp; Spec-wise Productivity</div></div>
  <div class="grow"></div>
  <span class="badge"><i class="dot off" id="dotQa"></i><span id="srcQa">Loading QA…</span></span>
  <span class="badge"><i class="dot off" id="dotCr"></i><span id="srcCr">Loading CR…</span></span>
  <span class="badge"><i class="dot off" id="dotProd"></i><span id="srcProd">Loading Prod…</span></span>
  <button class="hbtn accent" onclick="exportActiveTabJpg()">🖼️ Quick Share JPG</button>
  <button class="hbtn" id="reload">⟳ Refresh</button>
  <button class="hbtn" id="quit">Close</button>
</header>

<nav class="nav-tabs">
  <button class="nav-tab on" id="tabQA" onclick="switchTab('QA')">1. Quality Audit Report</button>
  <button class="nav-tab" id="tabCR" onclick="switchTab('CR')">2. Client Rejection Report</button>
  <button class="nav-tab" id="tabATTN" onclick="switchTab('ATTN')">3. Client Attention</button>
  <button class="nav-tab" id="tabCHARTS" onclick="switchTab('CHARTS')">4. Charts &amp; Analytics</button>
  <button class="nav-tab" id="tabCOMP" onclick="switchTab('COMP')">5. Period Compare</button>
  <button class="nav-tab" id="tabGROUP" onclick="switchTab('GROUP')">6. Group Data</button>
  <button class="nav-tab" id="tabSPEC" onclick="switchTab('SPEC')">7. Spec-wise Data</button>
</nav>

<main>
  <!-- ==================== TAB 1: QUALITY AUDIT REPORT ==================== -->
  <div id="viewQA">
    <section class="card">
      <div class="controls">
        <div class="field"><label for="qaAud">Auditor</label><select id="qaAud"></select></div>
        <div class="field"><span class="lbl">Period</span>
          <div class="seg"><button id="qaDay" class="on" onclick="setQaRange(false)">Single day</button><button id="qaRange" onclick="setQaRange(true)">Date range</button></div>
        </div>
        <div class="field"><label for="qaFrom" id="qaFromLbl">Date</label><input type="date" id="qaFrom"></div>
        <div class="field hide" id="qaToWrap"><label for="qaTo">To</label><input type="date" id="qaTo"></div>
        <button class="btn primary" onclick="runQaReport()">Generate QA Report</button>
      </div>
      <div class="chips"><span class="lbl">Quick pick:</span>
        <button class="chip" onclick="qaQuick('latest')">Latest audit day</button>
        <button class="chip" onclick="qaQuick('last7')">Last 7 days</button>
        <button class="chip" onclick="qaQuick('month')">This month</button>
      </div>
    </section>
    <div id="qaResult"></div>
  </div>

  <!-- ==================== TAB 2: CLIENT REJECTION REPORT ==================== -->
  <div id="viewCR" class="hide">
    <section class="card">
      <div class="controls">
        <div class="field"><span class="lbl">Period</span>
          <div class="seg"><button id="crDay" class="on" onclick="setCrRange(false)">Single day</button><button id="crRange" onclick="setCrRange(true)">Date range</button></div>
        </div>
        <div class="field"><label for="crFrom" id="crFromLbl">Date</label><input type="date" id="crFrom"></div>
        <div class="field hide" id="crToWrap"><label for="crTo">To</label><input type="date" id="crTo"></div>
        <button class="btn primary" onclick="runCrReport()">Generate Rejection Report</button>
      </div>
      <div class="chips"><span class="lbl">Quick pick:</span>
        <button class="chip" onclick="crQuick('latest')">Latest rejection day</button>
        <button class="chip" onclick="crQuick('last7')">Last 7 days</button>
        <button class="chip" onclick="crQuick('month')">This month</button>
      </div>
    </section>
    <div id="crResult"></div>
  </div>

  <!-- ==================== TAB 3: CLIENT ATTENTION RADAR ==================== -->
  <div id="viewATTN" class="hide">
    <section class="card">
      <div class="controls" style="justify-content:space-between;border-bottom:1px solid var(--line);padding-bottom:14px">
        <div>
          <div class="lbl" style="margin-bottom:6px">Client Attention Horizon (Combined QA Audits + Client Rejections)</div>
          <div class="seg">
            <button id="attnDAILY" class="on" onclick="setAttnWindow('DAILY')">Daily Compare</button>
            <button id="attnWEEKLY" onclick="setAttnWindow('WEEKLY')">Weekly Compare (7d vs 7d)</button>
            <button id="attnMONTHLY" onclick="setAttnWindow('MONTHLY')">Monthly Compare (MTD vs Last Month)</button>
            <button id="attnCUSTOM" class="hide" onclick="setAttnWindow('CUSTOM')">Custom Range</button>
          </div>
        </div>
        <div class="controls" style="gap:10px">
          <div class="field"><label for="attnFrom">From Date</label><input type="date" id="attnFrom"></div>
          <div class="field"><label for="attnTo">To Date</label><input type="date" id="attnTo"></div>
          <button class="btn primary" onclick="applyAttnDateRange()">Apply Range</button>
          <button class="btn ghost" onclick="resetAttnRange()">Reset</button>
        </div>
      </div>
      <div class="controls" style="justify-content:space-between;margin-top:12px">
        <div class="chips" style="margin:0">
          <span class="lbl">Status Filter:</span>
          <button class="chip on" id="attnF_ALL" onclick="setAttnFilter('ALL')">All Clients</button>
          <button class="chip" id="attnF_CRITICAL" onclick="setAttnFilter('CRITICAL')">Needs Immediate Attention</button>
          <button class="chip" id="attnF_WATCH" onclick="setAttnFilter('WATCH')">Watchlist</button>
          <button class="chip" id="attnF_HEALTHY" onclick="setAttnFilter('HEALTHY')">Healthy</button>
        </div>
        <div class="controls" style="gap:10px">
          <div class="field" style="min-width:165px">
            <select id="attnAudSel" onchange="onAttnAuditorChange()" aria-label="Filter by Auditor">
              <option value="ALL">All Auditors</option>
            </select>
          </div>
          <div class="field" style="min-width:200px">
            <select id="attnClientSel" onchange="renderAttention()" aria-label="Filter by Client">
              <option value="ALL">All Clients (Dropdown)</option>
            </select>
          </div>
        </div>
      </div>
    </section>
    <div id="attnResult"></div>
  </div>

  <!-- ==================== TAB 4: CHARTS & ANALYTICS ==================== -->
  <div id="viewCHARTS" class="hide">
    <section class="card">
      <div class="controls" style="justify-content:space-between;border-bottom:1px solid var(--line);padding-bottom:14px">
        <div>
          <div class="lbl" style="margin-bottom:6px">Quick Window</div>
          <div class="seg">
            <button id="ch7" onclick="setChartPreset(7)">7 Days</button>
            <button id="ch30" class="on" onclick="setChartPreset(30)">30 Days</button>
            <button id="ch90" onclick="setChartPreset(90)">90 Days</button>
            <button id="ch0" onclick="setChartPreset(0)">All Time</button>
          </div>
        </div>
        <div class="controls" style="gap:10px">
          <div class="field"><label for="chFrom">From Date</label><input type="date" id="chFrom" onchange="onChartDateChange()"></div>
          <div class="field"><label for="chTo">To Date</label><input type="date" id="chTo" onchange="onChartDateChange()"></div>
          <button class="btn primary" onclick="loadCharts()">Apply</button>
          <button class="btn ghost" onclick="resetChartFilters()">Reset Filters</button>
        </div>
      </div>
      <div class="controls" style="margin-top:14px">
        <div class="field" style="flex:1;min-width:150px">
          <label for="chAud">Auditor</label>
          <select id="chAud" onchange="loadCharts()"><option value="ALL">All Auditors</option></select>
        </div>
        <div class="field" style="flex:1;min-width:170px">
          <label for="chClient">Client</label>
          <select id="chClient" onchange="loadCharts()"><option value="ALL">All Clients</option></select>
        </div>
        <div class="field" style="flex:1;min-width:170px">
          <label for="chDes">Designer</label>
          <select id="chDes" onchange="loadCharts()"><option value="ALL">All Designers</option></select>
        </div>
        <div class="field" style="flex:1;min-width:170px">
          <label for="chQc">QC Name</label>
          <select id="chQc" onchange="loadCharts()"><option value="ALL">All QCs</option></select>
        </div>
        <div class="field" style="flex:1;min-width:190px">
          <label for="chRsn">Reason / Category</label>
          <select id="chRsn" onchange="loadCharts()"><option value="ALL">All Reasons</option></select>
        </div>
      </div>
      <div class="note" id="chartSpanNote" style="margin-top:10px">Visualizing combined QA defect trends and external client rejections</div>
    </section>
    <div id="chartsResult"></div>
  </div>

  <!-- ==================== TAB 5: PERIOD COMPARE ==================== -->
  <div id="viewCOMP" class="hide">
    <section class="card">
      <div class="controls" style="justify-content:space-between;border-bottom:1px solid var(--line);padding-bottom:14px">
        <div class="seg">
          <button id="cmQA" class="on" onclick="setCompMode('QA')">Internal QA Audits</button>
          <button id="cmCR" onclick="setCompMode('CR')">External Client Rejections</button>
        </div>
        <div class="chips" style="margin:0">
          <span class="lbl">Preset:</span>
          <button class="chip on" id="cpDAY" onclick="applyCompPreset('DAY')">Day vs Prev Day</button>
          <button class="chip" id="cpWEEK" onclick="applyCompPreset('WEEK')">This Week vs Last Week</button>
          <button class="chip" id="cpMONTH" onclick="applyCompPreset('MONTH')">This Month vs Last Month</button>
        </div>
      </div>
      <div class="vs-grid">
        <div class="vs-box">
          <div class="lbl" style="color:var(--brand);margin-bottom:8px">Period A (Current)</div>
          <div class="controls">
            <div class="field"><label for="caFrom">From</label><input type="date" id="caFrom"></div>
            <div class="field"><label for="caTo">To</label><input type="date" id="caTo"></div>
          </div>
        </div>
        <div style="text-align:center"><span class="vs-badge">VS</span></div>
        <div class="vs-box">
          <div class="lbl" style="color:var(--muted);margin-bottom:8px">Period B (Previous)</div>
          <div class="controls">
            <div class="field"><label for="cbFrom">From</label><input type="date" id="cbFrom"></div>
            <div class="field"><label for="cbTo">To</label><input type="date" id="cbTo"></div>
          </div>
        </div>
      </div>
      <div class="controls" style="margin-top:14px;border-top:1px solid var(--line);padding-top:14px">
        <div class="field" id="compAudWrap" style="flex:1;min-width:155px">
          <label for="compAud">Auditor</label>
          <select id="compAud" onchange="runCompare()"><option value="ALL">All Auditors</option></select>
        </div>
        <div class="field" style="flex:1;min-width:170px">
          <label for="compClientSel">Client</label>
          <select id="compClientSel" onchange="compClient=this.value;renderCompare()"><option value="ALL">All Clients</option></select>
        </div>
        <div class="field" style="flex:1;min-width:170px">
          <label for="compDesSel">Designer</label>
          <select id="compDesSel" onchange="compDes=this.value;renderCompare()"><option value="ALL">All Designers</option></select>
        </div>
        <div class="field" style="flex:1;min-width:170px">
          <label for="compQcSel">QC Name</label>
          <select id="compQcSel" onchange="compQc=this.value;renderCompare()"><option value="ALL">All QCs</option></select>
        </div>
        <div style="display:flex;gap:8px">
          <button class="btn ghost" onclick="resetCompFilters()">Reset Filters</button>
          <button class="btn primary" onclick="runCompare()">Run Comparison</button>
        </div>
      </div>
    </section>
    <div id="compResult"></div>
  </div>

  <!-- ==================== TAB 6: GROUP DATA ==================== -->
  <div id="viewGROUP" class="hide">
    <section class="card">
      <div class="controls">
        <div class="field"><span class="lbl">Period</span>
          <div class="seg"><button id="grpDay" class="on" onclick="setGrpRange(false)">Single day</button><button id="grpRange" onclick="setGrpRange(true)">Date range</button></div>
        </div>
        <div class="field"><label for="grpFrom" id="grpFromLbl">From</label><input type="date" id="grpFrom"></div>
        <div class="field hide" id="grpToWrap"><label for="grpTo">To</label><input type="date" id="grpTo"></div>
        <div class="field"><label for="grpTeamSel">Team</label><select id="grpTeamSel" onchange="runGroupData(false)"><option value="ALL">All Teams</option></select></div>
        <div class="field"><label for="grpDesigSel">Designation</label><select id="grpDesigSel" onchange="runGroupData(false)"><option value="ALL">All Designations</option></select></div>
        <div class="field"><label for="grpDesSel">Designer</label><select id="grpDesSel" onchange="runGroupData(false)"><option value="ALL">All Designers</option></select></div>
        <div class="field"><label for="grpSpecSel">Spec ID</label><select id="grpSpecSel" onchange="runGroupData(false)"><option value="ALL">All Specs</option></select></div>
        <button class="btn primary" onclick="runGroupData(false)">Generate Group Data</button>
        <button class="btn ghost" onclick="resetGroupFilters()">Reset</button>
      </div>
      <div class="chips"><span class="lbl">Quick pick:</span>
        <button class="chip" onclick="grpQuick('latest')">Latest day</button>
        <button class="chip" onclick="grpQuick('last7')">Last 7 days</button>
        <button class="chip" onclick="grpQuick('month')">This month</button>
      </div>
    </section>
    <div id="groupResult"></div>
  </div>

  <!-- ==================== TAB 7: SPEC-WISE DATA ==================== -->
  <div id="viewSPEC" class="hide">
    <section class="card">
      <div class="controls">
        <div class="field"><label for="specIdSel">Select Spec ID</label><select id="specIdSel" onchange="runSpecwiseData(false)"><option value="ALL">All Specs</option></select></div>
        <div class="field"><span class="lbl">Period</span>
          <div class="seg"><button id="specDay" class="on" onclick="setSpecRange(false)">Single day</button><button id="specRange" onclick="setSpecRange(true)">Date range</button></div>
        </div>
        <div class="field"><label for="specFrom" id="specFromLbl">From</label><input type="date" id="specFrom"></div>
        <div class="field hide" id="specToWrap"><label for="specTo">To</label><input type="date" id="specTo"></div>
        <div class="field"><label for="specTeamSel">Team</label><select id="specTeamSel" onchange="runSpecwiseData(false)"><option value="ALL">All Teams</option></select></div>
        <div class="field"><label for="specDesigSel">Designation</label><select id="specDesigSel" onchange="runSpecwiseData(false)"><option value="ALL">All Designations</option></select></div>
        <div class="field"><label for="specDesSel">Designer</label><select id="specDesSel" onchange="runSpecwiseData(false)"><option value="ALL">All Designers</option></select></div>
        <button class="btn primary" onclick="runSpecwiseData(false)">Generate Spec Report</button>
        <button class="btn ghost" onclick="resetSpecFilters()">Reset</button>
      </div>
      <div class="chips"><span class="lbl">Quick pick:</span>
        <button class="chip" onclick="specQuick('latest')">Latest day</button>
        <button class="chip" onclick="specQuick('last7')">Last 7 days</button>
        <button class="chip" onclick="specQuick('month')">This month</button>
      </div>
    </section>
    <div id="specResult"></div>
  </div>
</main>

<!-- QUICK SHARE JPG PREVIEW MODAL -->
<div class="modal-bg" id="jpgModal" onclick="if(event.target===this) closeJpgModal()">
  <div class="modal-card">
    <div class="modal-hdr">
      <strong style="font-size:15px">🖼️ Quick Share Executive Card (1-Page JPG)</strong>
      <div class="grow"></div>
      <button class="btn ghost" onclick="copyJpgToClipboard()">📋 Copy Image</button>
      <a class="btn primary" id="jpgDownloadLink" download="Quick_Share_Update.jpg" style="text-decoration:none">⬇ Save JPG</a>
      <button class="btn ghost" onclick="closeJpgModal()">✕</button>
    </div>
    <div class="modal-body">
      <img id="jpgPreviewImg" alt="Quick Share Executive Report">
    </div>
  </div>
</div>

<div id="toast"></div>

<script>
const T="__TOKEN__";
const $=s=>document.querySelector(s);
let ST=null, qaRange=false, crRange=false, grpRange=false, specRange=false, ACTIVE_TAB="QA";
let CUR_QA=null, qaClient="ALL", qaDes="ALL", qaQc="ALL", qaRsn="ALL", qaFault="ALL";
let CUR_CR=null, crClient="ALL", crDes="ALL", crQc="ALL", crRsn="ALL", crFault="ALL";
let CUR_ATTN=null, attnWin="DAILY", attnFilter="ALL";
let CUR_CHARTS=null, chartDays=30;
let CUR_COMP=null, compMode="QA", compClient="ALL", compDes="ALL", compQc="ALL", compGroupBy="client";
let CUR_GROUP=null, grpSearch="";
let CUR_SPEC=null, specSearch="", specSideMode="TOP";
let LAST_JPG_BLOB=null;

// Custom takeaways & summary overrides per tab
const CUSTOM_TK={QA:null, CR:null, ATTN:null, CHARTS:null, COMP:null, GROUP:null, SPEC:null};
const EDIT_TK_MODE={QA:false, CR:false, ATTN:false, CHARTS:false, COMP:false, GROUP:false, SPEC:false};
const CUSTOM_SUM_TXT={qaSumBox:null, crSumBox:null};

async function api(p,body){const r=await fetch(p,{method:body?"POST":"GET",headers:{"X-Token":T,"Content-Type":"application/json"},body:body?JSON.stringify(body):undefined});return r.json();}
const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const iso=d=>d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
const nf=n=>Number(n).toLocaleString("en-US");
function toast(m){const t=$("#toast");t.textContent=m;t.classList.add("show");clearTimeout(toast.h);toast.h=setTimeout(()=>t.classList.remove("show"),2200);}

function downloadBlobB64(b64, filename, mime){
  const bin=atob(b64), bytes=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i);
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([bytes],{type:mime}));
  a.download=filename; a.click();
}
function downloadDocx(b64, filename){
  downloadBlobB64(b64, filename, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
}
function toggleExpMenu(id, ev){
  if(ev) ev.stopPropagation();
  document.querySelectorAll(".exp-menu").forEach(m=>{ if(m.id!==id) m.classList.remove("open"); });
  const el=document.getElementById(id);
  if(el) el.classList.toggle("open");
}
document.addEventListener("click", ()=>document.querySelectorAll(".exp-menu").forEach(m=>m.classList.remove("open")));

function expDropdownHtml(tabKey){
  return `<div class="exp-wrap">
    <button class="btn ghost" onclick="runTabExport('${tabKey}','JPG')">🖼️ Quick Share JPG</button>
    <button class="btn primary" onclick="toggleExpMenu('exp_${tabKey}', event)">Download Report ▾</button>
    <div class="exp-menu" id="exp_${tabKey}">
      <button class="exp-item" onclick="runTabExport('${tabKey}','JPG')">🖼️ Quick Share Card (.jpg)</button>
      <button class="exp-item" onclick="runTabExport('${tabKey}','DOCX')">📄 Word Document (.docx)</button>
      <button class="exp-item" onclick="runTabExport('${tabKey}','PDF')">📕 PDF Document (.pdf)</button>
      <button class="exp-item" onclick="runTabExport('${tabKey}','XLSX')">📊 Excel Workbook (.xlsx)</button>
      <button class="exp-item" onclick="runTabExport('${tabKey}','CSV')">📑 CSV Spreadsheet (.csv)</button>
    </div>
  </div>`;
}

function exportActiveTabJpg(){
  runTabExport(ACTIVE_TAB, "JPG");
}

function teamBadgeHtml(team){
  const t=String(team||"");
  const tl=t.toLowerCase();
  let cls="team-other";
  if(tl.includes("blue")) cls="team-blue";
  else if(tl.includes("red")) cls="team-red";
  else if(tl.includes("green")) cls="team-green";
  return `<span class="team-pill ${cls}">${esc(t)}</span>`;
}

function renderStreakAlertsCard(alerts, streakDays=3, contextLabel="Audit"){
  if(!alerts || !alerts.length) return "";
  return `<div class="card streak-banner">
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:6px">
      <span class="pill bad">⚠ ${streakDays}+ DAY RECURRING ISSUE REMINDER</span>
      <strong style="font-size:14.5px">${alerts.length} Repeated Mistake Alert${alerts.length===1?"":"s"} (${esc(contextLabel)} — Designer, QC &amp; Client)</strong>
      <span class="note">Mistakes occurred for ${streakDays}+ consecutive days</span>
    </div>
    ${alerts.map(a=>`<div class="streak-item">
      <span class="pill ${a.dimension==='Client'?'bad':(a.dimension==='QC'?'warn':'good')}">${esc(a.dimension.toUpperCase())}</span>
      <strong>${esc(a.entity)}</strong>
      <span>—</span>
      <span class="bad"><b>${esc(a.issue)}</b></span>
      <span class="pill bad">${a.days} consecutive days (${a.orders} orders · ${esc(a.since)}${a.until && a.until!==a.since ? ' to '+esc(a.until) : ''})</span>
      ${a.details ? `<span class="note">${esc(a.details)}</span>` : ""}
    </div>`).join("")}
  </div>`;
}

// ==================== EDITABLE KEY TAKEAWAYS & ACTION PLAN ====================
function getEffectiveTakeaways(tabKey, defaultRows){
  if(CUSTOM_TK[tabKey] && Array.isArray(CUSTOM_TK[tabKey])){
    return CUSTOM_TK[tabKey].map(r=>[r[0], r[1]]);
  }
  return (defaultRows||[]).map(r=>[r[0], r[1]]);
}

function rerenderActiveTab(tabKey){
  if(tabKey==="QA") renderQaReport();
  else if(tabKey==="CR") renderCrReport();
  else if(tabKey==="ATTN") renderAttention();
  else if(tabKey==="CHARTS") renderCharts();
  else if(tabKey==="COMP") renderCompare();
  else if(tabKey==="GROUP") renderGroupData();
  else if(tabKey==="SPEC") renderSpecwiseData();
}

function toggleEditTakeaways(tabKey, defaultRows){
  if(!EDIT_TK_MODE[tabKey]){
    if(!CUSTOM_TK[tabKey]){
      CUSTOM_TK[tabKey]=(defaultRows||[]).map(r=>[r[0], r[1]]);
    }
    EDIT_TK_MODE[tabKey]=true;
  } else {
    saveTakeawaysFromDom(tabKey);
    EDIT_TK_MODE[tabKey]=false;
    toast("Key Takeaways updated ✓");
  }
  rerenderActiveTab(tabKey);
}

function saveTakeawaysFromDom(tabKey){
  const tbody=document.getElementById("tk_tbody_"+tabKey);
  if(!tbody) return;
  const rows=[];
  tbody.querySelectorAll("tr").forEach(tr=>{
    const catInp=tr.querySelector(".tk-inp");
    const detTxt=tr.querySelector(".tk-txt");
    if(catInp && detTxt){
      const c=catInp.value.trim();
      const d=detTxt.value.trim();
      if(c || d) rows.push([c || "Note", d]);
    }
  });
  CUSTOM_TK[tabKey]=rows;
}

function addTakeawayRow(tabKey){
  saveTakeawaysFromDom(tabKey);
  if(!CUSTOM_TK[tabKey]) CUSTOM_TK[tabKey]=[];
  CUSTOM_TK[tabKey].push(["Action Note", "Enter details here..."]);
  EDIT_TK_MODE[tabKey]=true;
  rerenderActiveTab(tabKey);
}

function deleteTakeawayRow(tabKey, idx){
  saveTakeawaysFromDom(tabKey);
  if(CUSTOM_TK[tabKey]){
    CUSTOM_TK[tabKey].splice(idx, 1);
  }
  rerenderActiveTab(tabKey);
}

function resetTakeaways(tabKey){
  CUSTOM_TK[tabKey]=null;
  EDIT_TK_MODE[tabKey]=false;
  toast("Reset to default takeaways");
  rerenderActiveTab(tabKey);
}

function renderTakeawaysCard(tabKey, defaultTakeaways, title="Key Takeaways & Action Plan"){
  const rows=getEffectiveTakeaways(tabKey, defaultTakeaways);
  const isEditing=!!EDIT_TK_MODE[tabKey];
  const isCustomized=Array.isArray(CUSTOM_TK[tabKey]);
  if(!rows.length && !isEditing) return "";

  const btnGroup = isEditing
    ? `<div style="display:flex;gap:6px;align-items:center">
        <button class="btn ghost sm" onclick="addTakeawayRow('${tabKey}')">+ Add Row</button>
        <button class="btn ghost sm" onclick="resetTakeaways('${tabKey}')">Reset</button>
        <button class="btn primary sm" onclick='toggleEditTakeaways("${tabKey}", ${JSON.stringify(defaultTakeaways||[])})'>✓ Save Changes</button>
      </div>`
    : `<div style="display:flex;gap:6px;align-items:center">
        ${isCustomized ? `<span class="pill warn">Edited</span><button class="btn ghost sm" onclick="resetTakeaways('${tabKey}')">Reset</button>` : ""}
        <button class="btn ghost sm" onclick='toggleEditTakeaways("${tabKey}", ${JSON.stringify(defaultTakeaways||[])})'>✏️ Edit</button>
      </div>`;

  return `<div class="card" style="margin-top:14px">
    <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:10px;flex-wrap:wrap">
      <h3 style="margin:0">${esc(title)}</h3>
      ${btnGroup}
    </div>
    <div style="overflow-x:auto;border:1px solid var(--line);border-radius:10px">
      <table class="tk-table">
        <thead><tr><th>Category</th><th>Details</th>${isEditing?'<th style="width:56px;text-align:center">Del</th>':''}</tr></thead>
        <tbody id="tk_tbody_${tabKey}">
          ${rows.map((r, idx)=> isEditing
            ? `<tr>
                <td><input class="tk-inp" value="${esc(r[0])}"></td>
                <td><textarea class="tk-txt">${esc(r[1])}</textarea></td>
                <td style="text-align:center"><button class="del-row-btn" onclick="deleteTakeawayRow('${tabKey}', ${idx})">✕</button></td>
              </tr>`
            : `<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td></tr>`
          ).join("")}
        </tbody>
      </table>
    </div>
  </div>`;
}

// ==================== EDITABLE TEXT SUMMARY BOX ====================
function renderMinimizedSummaryCard(boxId, defaultText){
  const text = CUSTOM_SUM_TXT[boxId] !== null ? CUSTOM_SUM_TXT[boxId] : defaultText;
  if(!text) return "";
  window[boxId+"_TXT"]=text;
  const lines=text.split("\n").filter(Boolean);
  const preview=lines.slice(0,2).join(" · ");
  const isEdited = CUSTOM_SUM_TXT[boxId] !== null;
  return `<div class="card" style="margin-top:14px;padding:13px 18px">
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <h3 style="margin:0">Text Summary ${isEdited?'<span class="pill warn" style="margin-left:6px">Edited</span>':''}</h3>
      <span class="note" style="flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(preview)}</span>
      <button class="btn ghost sm" onclick="toggleEditSummary('${boxId}')">✏️ Edit Text</button>
      <button class="btn ghost sm" onclick="copyText(window.${boxId}_TXT)">Copy Text</button>
      <button class="btn ghost sm" id="${boxId}_btn" onclick="toggleSummaryBox('${boxId}')">Expand ▾</button>
    </div>
    <pre id="${boxId}_pre" class="hide" style="margin-top:12px">${esc(text)}</pre>
    <div id="${boxId}_editWrap" class="hide" style="margin-top:12px">
      <textarea id="${boxId}_ta" style="width:100%;min-height:220px;font:13px/1.55 Consolas,monospace;padding:12px;border-radius:10px;border:1px solid var(--brand);background:var(--bg);color:var(--ink)">${esc(text)}</textarea>
      <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:8px">
        <button class="btn ghost sm" onclick="resetSummaryBox('${boxId}')">Reset Original</button>
        <button class="btn primary sm" onclick="saveSummaryBox('${boxId}')">✓ Save Text</button>
      </div>
    </div>
  </div>`;
}
function toggleSummaryBox(boxId){
  const pre=document.getElementById(boxId+"_pre");
  const btn=document.getElementById(boxId+"_btn");
  if(!pre) return;
  const hidden=pre.classList.toggle("hide");
  if(btn) btn.textContent=hidden?"Expand ▾":"Minimize ▴";
}
function toggleEditSummary(boxId){
  const wrap=document.getElementById(boxId+"_editWrap");
  const pre=document.getElementById(boxId+"_pre");
  if(!wrap) return;
  const opening=wrap.classList.contains("hide");
  wrap.classList.toggle("hide");
  if(opening && pre) pre.classList.add("hide");
}
function saveSummaryBox(boxId){
  const ta=document.getElementById(boxId+"_ta");
  if(!ta) return;
  CUSTOM_SUM_TXT[boxId]=ta.value;
  toast("Summary text saved ✓");
  rerenderActiveTab(boxId==="qaSumBox"?"QA":"CR");
}
function resetSummaryBox(boxId){
  CUSTOM_SUM_TXT[boxId]=null;
  toast("Restored original summary text");
  rerenderActiveTab(boxId==="qaSumBox"?"QA":"CR");
}

// ==================== ONE-PAGE QUICK SHARE JPG RENDERER ====================
function wrapCanvasText(ctx, text, maxWidth){
  const words=String(text||"").split(/\s+/);
  const lines=[];
  let cur="";
  for(const w of words){
    const test=cur ? cur+" "+w : w;
    if(ctx.measureText(test).width > maxWidth && cur){
      lines.push(cur);
      cur=w;
    } else {
      cur=test;
    }
  }
  if(cur) lines.push(cur);
  return lines.length ? lines : [""];
}

function roundRectPath(ctx, x, y, w, h, r){
  ctx.beginPath();
  ctx.moveTo(x+r, y);
  ctx.lineTo(x+w-r, y);
  ctx.quadraticCurveTo(x+w, y, x+w, y+r);
  ctx.lineTo(x+w, y+h-r);
  ctx.quadraticCurveTo(x+w, y+h, x+w-r, y+h);
  ctx.lineTo(x+r, y+h);
  ctx.quadraticCurveTo(x, y+h, x, y+h-r);
  ctx.lineTo(x, y+r);
  ctx.quadraticCurveTo(x, y, x+r, y);
  ctx.closePath();
}

function generateQuickShareJpg(payload){
  const W=1160;
  const pad=34;
  const innerW=W - pad*2;
  const dpr=2;

  const probe=document.createElement("canvas");
  const pctx=probe.getContext("2d");

  const takeaways=(payload.takeaways||[]).slice(0,5);
  const alerts=(payload.streak_alerts||[]).slice(0,5);
  const rows=(payload.table_rows||[]).slice(0,10);
  const headers=(payload.table_headers||[]).slice(0,7);
  const combinedBar=payload.combined_strip || [];

  let H=124;
  H+=24 + 102;
  if(combinedBar.length){
    H+=18 + 86;
  }
  if(alerts.length){
    H+=22 + 38 + alerts.length*30 + 14;
  }
  if(takeaways.length){
    pctx.font="500 13.5px 'Segoe UI', system-ui, sans-serif";
    let tkH=44;
    takeaways.forEach(t=>{
      const lines=wrapCanvasText(pctx, t[1], innerW - 250);
      tkH += Math.max(32, lines.length*19 + 14);
    });
    H+=22 + tkH;
  }
  if(rows.length && headers.length){
    H+=22 + 40 + 34 + rows.length*32 + 12;
  }
  H+=46;

  const canvas=document.createElement("canvas");
  canvas.width=W*dpr;
  canvas.height=H*dpr;
  const ctx=canvas.getContext("2d");
  ctx.scale(dpr, dpr);

  ctx.fillStyle="#f1f5f9";
  ctx.fillRect(0,0,W,H);

  const grad=ctx.createLinearGradient(0,0,W,124);
  grad.addColorStop(0,"#0f1d3d");
  grad.addColorStop(1,"#1e3a8a");
  ctx.fillStyle=grad;
  ctx.fillRect(0,0,W,124);

  roundRectPath(ctx, pad, 28, 46, 46, 11);
  ctx.fillStyle="#2563eb";
  ctx.fill();
  ctx.fillStyle="#ffffff";
  ctx.font="800 17px 'Segoe UI', system-ui, sans-serif";
  ctx.fillText("QA+", pad+7, 57);

  ctx.fillStyle="#ffffff";
  ctx.font="700 23px 'Segoe UI', system-ui, sans-serif";
  ctx.fillText(payload.title || "Quality & Rejection Quick Share", pad+62, 52);

  ctx.fillStyle="#cbd5e1";
  ctx.font="500 13.5px 'Segoe UI', system-ui, sans-serif";
  ctx.fillText(payload.subtitle || "", pad+62, 76);

  const stampText=payload.badge_text || "EXECUTIVE SNAPSHOT";
  ctx.font="700 11.5px 'Segoe UI', system-ui, sans-serif";
  const stW=ctx.measureText(stampText).width + 26;
  roundRectPath(ctx, W-pad-stW, 34, stW, 30, 15);
  ctx.fillStyle="rgba(255,255,255,0.14)";
  ctx.fill();
  ctx.fillStyle="#93c5fd";
  ctx.fillText(stampText, W-pad-stW+13, 53);

  let y=146;

  const kHdrs=(payload.kpi_headers||[]).slice(0,4);
  const kVals=(payload.kpi_row||[]).slice(0,4);
  const kSubs=(payload.kpi_subs||[]);
  if(kHdrs.length){
    const gap=14;
    const cardW=(innerW - gap*(kHdrs.length-1))/kHdrs.length;
    kHdrs.forEach((kh, i)=>{
      const cx=pad + i*(cardW+gap);
      roundRectPath(ctx, cx, y, cardW, 96, 12);
      ctx.fillStyle="#ffffff";
      ctx.fill();
      ctx.strokeStyle="#cbd5e1";
      ctx.lineWidth=1;
      ctx.stroke();

      ctx.fillStyle="#64748b";
      ctx.font="700 11px 'Segoe UI', system-ui, sans-serif";
      ctx.fillText(String(kh).toUpperCase(), cx+16, y+24);

      const valStr=String(kVals[i]??"0");
      const isBad=/err|rej/i.test(kh) && parseFloat(valStr.replace(/[^0-9.]/g,""))>0;
      const isRate=/rate/i.test(kh);
      const rateNum=parseFloat(valStr.replace(/[^0-9.]/g,""))||0;
      ctx.fillStyle=isRate ? (rateNum>10 ? "#b91c1c" : "#15803d") : (isBad ? "#b91c1c" : "#0f172a");
      ctx.font="800 26px 'Segoe UI', system-ui, sans-serif";
      ctx.fillText(valStr, cx+16, y+58);

      if(kSubs[i]){
        ctx.fillStyle="#64748b";
        ctx.font="500 11.5px 'Segoe UI', system-ui, sans-serif";
        ctx.fillText(String(kSubs[i]), cx+16, y+80);
      }
    });
    y+=114;
  }

  if(combinedBar.length){
    roundRectPath(ctx, pad, y, innerW, 82, 12);
    ctx.fillStyle="#e8eefc";
    ctx.fill();
    ctx.strokeStyle="#bfdbfe";
    ctx.stroke();

    ctx.fillStyle="#1e40af";
    ctx.font="800 11px 'Segoe UI', system-ui, sans-serif";
    ctx.fillText(payload.combined_title || "COMBINED AUDIT & CLIENT REJECTION CROSS-CHECK (SAME PERIOD)", pad+16, y+22);

    const colW=(innerW-32)/combinedBar.length;
    combinedBar.forEach((item, idx)=>{
      const bx=pad+16 + idx*colW;
      ctx.fillStyle="#475569";
      ctx.font="600 11.5px 'Segoe UI', system-ui, sans-serif";
      ctx.fillText(item.label, bx, y+44);
      ctx.fillStyle=item.bad ? "#b91c1c" : "#0f172a";
      ctx.font="800 17px 'Segoe UI', system-ui, sans-serif";
      ctx.fillText(item.value, bx, y+67);
    });
    y+=100;
  }

  if(alerts.length){
    const boxH=38 + alerts.length*30 + 12;
    roundRectPath(ctx, pad, y, innerW, boxH, 12);
    ctx.fillStyle="#fef2f2";
    ctx.fill();
    ctx.strokeStyle="#fecaca";
    ctx.stroke();

    ctx.fillStyle="#b91c1c";
    ctx.fillRect(pad, y+8, 5, boxH-16);

    ctx.fillStyle="#991b1b";
    ctx.font="800 12px 'Segoe UI', system-ui, sans-serif";
    ctx.fillText(`⚠ 3+ DAY RECURRING ISSUE REMINDERS (${alerts.length} ACTIVE STREAK ALERT${alerts.length>1?"S":""})`, pad+18, y+25);

    let ay=y+48;
    alerts.forEach(a=>{
      const pillTxt=String(a.dimension||"").toUpperCase();
      ctx.font="700 10.5px 'Segoe UI', system-ui, sans-serif";
      const pw=ctx.measureText(pillTxt).width + 16;
      roundRectPath(ctx, pad+18, ay-13, pw, 19, 9);
      ctx.fillStyle=pillTxt==="CLIENT"?"#fee2e2":(pillTxt==="QC"?"#fef3c7":"#dcfce7");
      ctx.fill();
      ctx.fillStyle=pillTxt==="CLIENT"?"#991b1b":(pillTxt==="QC"?"#92400e":"#166534");
      ctx.fillText(pillTxt, pad+26, ay);

      ctx.fillStyle="#0f172a";
      ctx.font="700 12.5px 'Segoe UI', system-ui, sans-serif";
      const line=`${a.entity}: ${a.issue} — ${a.days} consecutive days (${a.orders} orders · ${a.since} to ${a.until})`;
      const trimmed=line.length>112 ? line.slice(0,109)+"..." : line;
      ctx.fillText(trimmed, pad+26+pw, ay);
      ay+=30;
    });
    y+=boxH+18;
  }

  if(takeaways.length){
    ctx.font="500 13.5px 'Segoe UI', system-ui, sans-serif";
    const measured=takeaways.map(t=>{
      const lines=wrapCanvasText(ctx, t[1], innerW - 255);
      return {cat:t[0], lines, h:Math.max(32, lines.length*19 + 14)};
    });
    const boxH=42 + measured.reduce((s,m)=>s+m.h, 0);
    roundRectPath(ctx, pad, y, innerW, boxH, 12);
    ctx.fillStyle="#ffffff";
    ctx.fill();
    ctx.strokeStyle="#cbd5e1";
    ctx.stroke();

    ctx.fillStyle="#1b2a4a";
    ctx.fillRect(pad, y, innerW, 36);
    ctx.fillStyle="#ffffff";
    ctx.font="700 12.5px 'Segoe UI', system-ui, sans-serif";
    ctx.fillText("KEY TAKEAWAYS & ACTION PLAN", pad+16, y+23);

    let ty=y+36;
    measured.forEach((m, idx)=>{
      if(idx%2===1){
        ctx.fillStyle="#f8fafc";
        ctx.fillRect(pad+1, ty, innerW-2, m.h);
      }
      ctx.strokeStyle="#e2e8f0";
      ctx.beginPath(); ctx.moveTo(pad, ty+m.h); ctx.lineTo(pad+innerW, ty+m.h); ctx.stroke();

      ctx.fillStyle="#1e40af";
      ctx.font="700 12.5px 'Segoe UI', system-ui, sans-serif";
      ctx.fillText(String(m.cat), pad+16, ty+21);

      ctx.fillStyle="#1e293b";
      ctx.font="500 13px 'Segoe UI', system-ui, sans-serif";
      m.lines.forEach((ln, li)=>{
        ctx.fillText(ln, pad+235, ty+21 + li*19);
      });
      ty+=m.h;
    });
    y+=boxH+18;
  }

  if(rows.length && headers.length){
    const boxH=38 + 32 + rows.length*31 + 10;
    roundRectPath(ctx, pad, y, innerW, boxH, 12);
    ctx.fillStyle="#ffffff";
    ctx.fill();
    ctx.strokeStyle="#cbd5e1";
    ctx.stroke();

    ctx.fillStyle="#1b2a4a";
    ctx.fillRect(pad, y, innerW, 36);
    ctx.fillStyle="#ffffff";
    ctx.font="700 12.5px 'Segoe UI', system-ui, sans-serif";
    ctx.fillText(String(payload.table_title || "DETAILS").toUpperCase(), pad+16, y+23);

    const colW=(innerW-32)/headers.length;
    let ry=y+36;
    ctx.fillStyle="#f1f5f9";
    ctx.fillRect(pad+1, ry, innerW-2, 30);
    ctx.fillStyle="#475569";
    ctx.font="700 11px 'Segoe UI', system-ui, sans-serif";
    headers.forEach((hd, ci)=>{
      ctx.fillText(String(hd).toUpperCase(), pad+16 + ci*colW, ry+19);
    });
    ry+=30;

    rows.forEach((r, ri)=>{
      if(ri%2===1){
        ctx.fillStyle="#f8fafc";
        ctx.fillRect(pad+1, ry, innerW-2, 31);
      }
      ctx.strokeStyle="#e2e8f0";
      ctx.beginPath(); ctx.moveTo(pad, ry+31); ctx.lineTo(pad+innerW, ry+31); ctx.stroke();

      headers.forEach((_, ci)=>{
        const cell=String(r[ci]??"");
        const maxChars=Math.max(10, Math.floor(colW/7.5));
        const disp=cell.length>maxChars ? cell.slice(0, maxChars-2)+"…" : cell;
        ctx.fillStyle=ci===0 ? "#0f172a" : "#334155";
        ctx.font=ci===0 ? "700 12px 'Segoe UI', system-ui, sans-serif" : "500 12px 'Segoe UI', system-ui, sans-serif";
        ctx.fillText(disp, pad+16 + ci*colW, ry+20);
      });
      ry+=31;
    });
    y+=boxH+14;
  }

  ctx.fillStyle="#64748b";
  ctx.font="500 11.5px 'Segoe UI', system-ui, sans-serif";
  ctx.fillText(`Generated by Quality & Client Analytics Suite · ${new Date().toLocaleString()}`, pad, H-16);
  const totalShown=rows.length;
  const totalAll=(payload.table_rows||[]).length;
  if(totalAll > totalShown){
    const moreTxt=`Showing top ${totalShown} of ${totalAll} rows (full table in .docx / .xlsx)`;
    const mw=ctx.measureText(moreTxt).width;
    ctx.fillText(moreTxt, W-pad-mw, H-16);
  }

  const jpgUrl=canvas.toDataURL("image/jpeg", 0.94);
  $("#jpgPreviewImg").src=jpgUrl;
  const dl=$("#jpgDownloadLink");
  dl.href=jpgUrl;
  dl.download=(payload.file_base || "Quick_Share_Report") + ".jpg";

  canvas.toBlob(b=>{ LAST_JPG_BLOB=b; }, "image/png");
  $("#jpgModal").classList.add("open");
}

function closeJpgModal(){
  $("#jpgModal").classList.remove("open");
}

async function copyJpgToClipboard(){
  if(!LAST_JPG_BLOB){ toast("Image not ready"); return; }
  try{
    await navigator.clipboard.write([new ClipboardItem({"image/png": LAST_JPG_BLOB})]);
    toast("Image copied! Paste directly into Slack / Teams / WhatsApp ✓");
  }catch(e){
    toast("Clipboard copy blocked by browser — use Save JPG");
  }
}

function openPrintPdf(payload){
  const w=window.open("","_blank");
  if(!w){ toast("Allow popups to save PDF"); return; }
  const kpiHtml=(payload.kpi_headers||[]).map((h,i)=>`<div class="kpi"><div class="kl">${esc(h)}</div><div class="kv">${esc((payload.kpi_row||[])[i]??"")}</div></div>`).join("");
  const tkHtml=(payload.takeaways&&payload.takeaways.length)?`<h2>Key Takeaways &amp; Action Plan</h2>
    <table><thead><tr><th style="width:28%">Category</th><th>Details</th></tr></thead>
    <tbody>${payload.takeaways.map(r=>`<tr><td><b>${esc(r[0])}</b></td><td>${esc(r[1])}</td></tr>`).join("")}</tbody></table>`:"";
  const cardsHtml=(payload.attn_cards&&payload.attn_cards.length)?`<h2>Client Attention Breakdown Cards (${payload.attn_cards.length})</h2>` + payload.attn_cards.map(c=>{
    const st=c.status||"HEALTHY";
    const bCol=st==="CRITICAL"?"#b91c1c":(st==="WATCH"?"#b45309":"#15803d");
    const hBg=st==="CRITICAL"?"#fdecec":(st==="WATCH"?"#fff4e0":"#e7f6ec");
    const bLbl=st==="CRITICAL"?"ACTION NEEDED":(st==="WATCH"?"WATCHLIST":"HEALTHY");
    return `<div style="border:1px solid #cbd5e1;border-left:5px solid ${bCol};border-radius:8px;margin-bottom:12px;overflow:hidden;page-break-inside:avoid">
      <div style="background:${hBg};padding:8px 12px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #cbd5e1">
        <div><span style="background:${bCol};color:#fff;font-size:9px;font-weight:700;padding:2px 7px;border-radius:99px;margin-right:8px">${bLbl}</span><b>${esc(c.client)}</b> <span style="color:#64748b;font-size:10px">Priority Score: ${c.score}</span></div>
        <div style="font-size:10px;color:#475569">${(c.flags||[]).map(f=>esc(f)).join(" · ")}</div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;padding:10px 12px;background:#f8fafc">
        <div style="background:#fff;border:1px solid #e2e8f0;border-radius:6px;padding:8px 10px">
          <div style="font-size:9px;font-weight:700;color:#2c5282">INTERNAL QA AUDIT BREAKDOWN</div>
          <div style="font-size:16px;font-weight:700;color:#1b2a4a;margin:2px 0">${Number(c.qa_a_rate||0).toFixed(2)}%</div>
          <div style="font-size:10px;color:#64748b;margin-bottom:6px">${nf(c.qa_a_err||0)} errors / ${nf(c.qa_a_imgs||0)} imgs (Prior: ${Number(c.qa_b_rate||0).toFixed(2)}%)</div>
          <div style="font-size:10.5px;border-top:1px solid #e2e8f0;padding-top:5px;line-height:1.45">
            <div><b>Reasons:</b> ${(c.qa_reasons&&c.qa_reasons.length)?esc(c.qa_reasons.join(", ")):"None"}</div>
            <div><b>Designers:</b> ${(c.qa_designers&&c.qa_designers.length)?esc(c.qa_designers.join(", ")):"None"}</div>
            <div><b>QC Names:</b> ${(c.qa_qcs&&c.qa_qcs.length)?esc(c.qa_qcs.join(", ")):"None"}</div>
          </div>
        </div>
        <div style="background:#fff;border:1px solid #e2e8f0;border-radius:6px;padding:8px 10px">
          <div style="font-size:9px;font-weight:700;color:#b45309">EXTERNAL CLIENT REJECTIONS BREAKDOWN</div>
          <div style="font-size:16px;font-weight:700;color:#1b2a4a;margin:2px 0">${nf(c.cr_a_ord||0)} orders</div>
          <div style="font-size:10px;color:#64748b;margin-bottom:6px">${nf(c.cr_a_imgs||0)} imgs · Prod: ${c.cr_a_prod||0} / Cust: ${c.cr_a_cust||0} (Prior: ${c.cr_b_ord||0} ord)</div>
          <div style="font-size:10.5px;border-top:1px solid #e2e8f0;padding-top:5px;line-height:1.45">
            <div><b>Reasons:</b> ${(c.cr_reasons&&c.cr_reasons.length)?esc(c.cr_reasons.join(", ")):"None"}</div>
            <div><b>Designers:</b> ${(c.cr_designers&&c.cr_designers.length)?esc(c.cr_designers.join(", ")):"None"}</div>
            <div><b>QC Names:</b> ${(c.cr_qcs&&c.cr_qcs.length)?esc(c.cr_qcs.join(", ")):"None"}</div>
          </div>
        </div>
      </div>
    </div>`;
  }).join(""):"";
  const tblHtml=(!cardsHtml && payload.table_rows&&payload.table_rows.length)?`<h2>${esc(payload.table_title||"Details")}</h2>
    <table><thead><tr>${(payload.table_headers||[]).map(h=>`<th>${esc(h)}</th>`).join("")}</tr></thead>
    <tbody>${payload.table_rows.map(r=>`<tr>${r.map(c=>`<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</tbody></table>`:"";
  const sumHtml=payload.summary_text?`<h2>Audit Summary &amp; Notes</h2><pre>${esc(payload.summary_text)}</pre>`:"";

  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(payload.title)}</title>
  <style>
    body{font-family:Calibri,"Segoe UI",Arial,sans-serif;color:#1e293b;margin:28px;font-size:12px}
    .banner{background:#1b2a4a;color:#fff;padding:18px 22px;border-radius:8px;margin-bottom:18px}
    .banner h1{margin:0 0 4px;font-size:20px}
    .banner .sub{color:#cbd5e1;font-size:12px}
    h2{color:#2c5282;font-size:13px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1.5px solid #cbd5e1;padding-bottom:4px;margin:18px 0 10px}
    .kpis{display:grid;grid-template-columns:repeat(${Math.max(1,(payload.kpi_headers||[]).length)},1fr);gap:10px}
    .kpi{background:#f1f5f9;border:1px solid #cbd5e1;border-radius:6px;padding:10px;text-align:center}
    .kl{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase}
    .kv{font-size:17px;font-weight:700;color:#1b2a4a;margin-top:3px}
    table{width:100%;border-collapse:collapse;margin-top:6px}
    th{background:#1b2a4a;color:#fff;text-align:left;padding:8px 10px;font-size:11px}
    td{padding:8px 10px;border-bottom:1px solid #cbd5e1;font-size:11.5px;vertical-align:top}
    tr:nth-child(even) td{background:#f8fafc}
    pre{background:#f1f5f9;border-left:4px solid #2c5282;padding:12px 14px;white-space:pre-wrap;font-family:Consolas,monospace;font-size:11px;line-height:1.45}
    @media print{body{margin:12mm}@page{size:A4;margin:12mm}}
  </style></head><body>
    <div class="banner"><h1>${esc(payload.title)}</h1><div class="sub">${esc(payload.subtitle||"")}</div></div>
    ${kpiHtml?`<h2>Executive Summary Metrics</h2><div class="kpis">${kpiHtml}</div>`:""}
    ${cardsHtml}${tkHtml}${tblHtml}${sumHtml}
    <script>window.onload=function(){setTimeout(function(){window.print();},250);};<\/script>
  </body></html>`);
  w.document.close();
}

function downloadCsv(payload, filename){
  const rows=[];
  if(payload.table_headers && payload.table_headers.length) rows.push(payload.table_headers);
  (payload.table_rows||[]).forEach(r=>rows.push(r));
  const csv=rows.map(r=>r.map(v=>`"${String(v??"").replace(/"/g,'""')}"`).join(",")).join("\r\n");
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8;"}));
  a.download=filename; a.click();
}

async function runTabExport(tabKey, fmt){
  if(EDIT_TK_MODE[tabKey]){
    saveTakeawaysFromDom(tabKey);
  }
  const payload=getTabExportPayload(tabKey);
  if(!payload){ toast("Generate report data first"); return; }
  const baseName=payload.file_base || ("Quality_Report_"+tabKey);
  if(fmt==="JPG"){
    generateQuickShareJpg(payload);
    return;
  }
  if(fmt==="PDF"){
    openPrintPdf(payload);
    return;
  }
  if(fmt==="CSV"){
    downloadCsv(payload, baseName+".csv");
    toast("CSV downloaded ✓");
    return;
  }
  if(fmt==="XLSX"){
    toast("Generating Excel .xlsx…");
    const res=await api("/api/export_xlsx", payload);
    if(res.b64) downloadBlobB64(res.b64, baseName+".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    else if(res.error) toast(res.error);
    return;
  }
  toast("Generating Word .docx…");
  const res=await api("/api/export_docx", payload);
  if(res.b64) downloadDocx(res.b64, baseName+".docx");
  else if(res.error) toast(res.error);
}

function deltaBadge(curr, prev, invertGood=true, isPct=false){
  const diff=curr-prev;
  if(Math.abs(diff)<0.005) return '<span class="delta neutral">0</span>';
  const sign=diff>0?"+":"";
  const valStr=isPct?`${sign}${diff.toFixed(2)}%`:`${sign}${nf(diff)}`;
  const isGood=invertGood?diff<0:diff>0;
  return `<span class="delta ${isGood?'down-good':'up-bad'}">${valStr} ${diff>0?'↑':'↓'}</span>`;
}

function switchTab(tab){
  ACTIVE_TAB=tab;
  ["QA","CR","ATTN","CHARTS","COMP","GROUP","SPEC"].forEach(t=>{
    $("#tab"+t).classList.toggle("on", t===tab);
    $("#view"+t).classList.toggle("hide", t!==tab);
  });
  if(tab==="CR" && !CUR_CR) runCrReport();
  if(tab==="ATTN" && !CUR_ATTN) loadAttention();
  if(tab==="CHARTS" && !CUR_CHARTS) loadCharts(chartDays);
  if(tab==="COMP" && !CUR_COMP) applyCompPreset("DAY");
  if(tab==="GROUP" && !CUR_GROUP) runGroupData(true);
  if(tab==="SPEC" && !CUR_SPEC) runSpecwiseData(true);
}

// ==================== INIT ====================
async function init(){
  ST=await api("/api/status");
  if(ST.loading){
    $("#qaResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Loading QA, Client Rejection &amp; Productivity spreadsheets…</div>';
    setTimeout(init,900); return;
  }
  $("#dotQa").className=ST.qa_loaded?"dot":"dot err";
  $("#srcQa").textContent=ST.qa_source+(ST.qa_loaded?` (${nf(ST.qa_rows)})`:"");
  $("#dotCr").className=ST.cr_loaded?"dot":"dot err";
  $("#srcCr").textContent=ST.cr_source+(ST.cr_loaded?` (${nf(ST.cr_rows)})`:"");
  $("#dotProd").className=ST.prod_loaded?"dot":"dot err";
  $("#srcProd").textContent=ST.prod_source+(ST.prod_loaded?` (${nf(ST.prod_rows)})`:"");

  if(ST.auditors && ST.auditors.length){
    const cur=$("#qaAud").value;
    $("#qaAud").innerHTML=ST.auditors.map(a=>`<option value="${esc(a.name)}">${esc(a.name)}</option>`).join("");
    $("#qaAud").value=cur||ST.default_auditor;
    if($("#compAud")){
      const curCompA=$("#compAud").value||"ALL";
      $("#compAud").innerHTML='<option value="ALL">All Auditors</option>' + ST.auditors.map(a=>`<option value="${esc(a.name)}">${esc(a.name.toUpperCase())}</option>`).join("");
      $("#compAud").value=curCompA;
    }
  }
  if(!$("#qaFrom").value){
    const meta=(ST.auditors||[]).find(x=>x.name===$("#qaAud").value);
    $("#qaFrom").value=(meta&&meta.latest)||iso(new Date());
    runQaReport();
  }
  if(!$("#crFrom").value){
    $("#crFrom").value=ST.cr_latest||iso(new Date());
  }
  if(!$("#grpFrom").value){
    $("#grpFrom").value=ST.prod_latest||iso(new Date());
  }
  if(!$("#specFrom").value){
    $("#specFrom").value=ST.prod_latest||iso(new Date());
  }
}

// ==================== TAB 1: QA REPORT ====================
function setQaRange(r){qaRange=r;$("#qaDay").classList.toggle("on",!r);$("#qaRange").classList.toggle("on",r);$("#qaToWrap").classList.toggle("hide",!r);$("#qaFromLbl").textContent=r?"From":"Date";if(!r)$("#qaTo").value="";}
function qaQuick(q){
  const meta=(ST.auditors||[]).find(x=>x.name===$("#qaAud").value);
  const lat=(meta&&meta.latest)||iso(new Date()), d=new Date(lat+"T00:00:00");
  if(q==="latest"){setQaRange(false);$("#qaFrom").value=lat;}
  else if(q==="last7"){const s=new Date(d);s.setDate(s.getDate()-6);setQaRange(true);$("#qaFrom").value=iso(s);$("#qaTo").value=lat;}
  else{const s=new Date(d.getFullYear(),d.getMonth(),1);setQaRange(true);$("#qaFrom").value=iso(s);$("#qaTo").value=lat;}
  runQaReport();
}
async function runQaReport(resetFilters=true){
  const f=$("#qaFrom").value, t=qaRange?$("#qaTo").value:"";
  if(!f) return;
  if(resetFilters){ qaClient="ALL"; qaDes="ALL"; qaQc="ALL"; qaRsn="ALL"; qaFault="ALL"; CUSTOM_TK.QA=null; EDIT_TK_MODE.QA=false; CUSTOM_SUM_TXT.qaSumBox=null; }
  $("#qaResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Building QA report…</div>';
  CUR_QA=await api("/api/qa_report",{
    auditor:$("#qaAud").value, start:f, end:t||f,
    client:qaClient, designer:qaDes, qc:qaQc, reason:qaRsn
  });
  renderQaReport();
}
function applyQaDropdownFilters(){
  qaClient=$("#qaClientSel") ? $("#qaClientSel").value : "ALL";
  qaDes=$("#qaDesSel") ? $("#qaDesSel").value : "ALL";
  qaQc=$("#qaQcSel") ? $("#qaQcSel").value : "ALL";
  qaRsn=$("#qaRsnSel") ? $("#qaRsnSel").value : "ALL";
  runQaReport(false);
}
function resetQaFilters(){
  qaClient="ALL"; qaDes="ALL"; qaQc="ALL"; qaRsn="ALL"; qaFault="ALL";
  runQaReport(true);
}
function getFilteredQaOrders(){
  if(!CUR_QA || CUR_QA.error) return [];
  let orders=(CUR_QA.summary && CUR_QA.summary.orders_list) || [];
  if(qaClient!=="ALL") orders=orders.filter(o=>o.client===qaClient);
  if(qaDes!=="ALL") orders=orders.filter(o=>o.designer===qaDes);
  if(qaQc!=="ALL") orders=orders.filter(o=>o.qc===qaQc);
  if(qaRsn!=="ALL") orders=orders.filter(o=>String(o.categories||"").split(",").map(x=>x.trim()).includes(qaRsn));
  if(qaFault==="PROD") orders=orders.filter(o=>!o.is_customer);
  if(qaFault==="CUST") orders=orders.filter(o=>o.is_customer);
  return orders;
}
function renderQaReport(){
  const el=$("#qaResult"), r=CUR_QA;
  if(!r) return;
  if(r.error){el.innerHTML=`<div class="card empty" style="margin-top:16px"><b>No QA Report</b><br>${esc(r.error)}</div>`;return;}
  const s=r.summary;
  const allOrders=s.orders_list||[];
  const clientsOpt=Array.from(new Set(allOrders.map(o=>o.client).filter(Boolean))).sort();
  const desOpt=Array.from(new Set(allOrders.map(o=>o.designer).filter(v=>v && v!=="Unassigned"))).sort();
  const qcOpt=Array.from(new Set(allOrders.map(o=>o.qc).filter(v=>v && v!=="Unassigned"))).sort();
  const rsnSet=new Set();
  allOrders.forEach(o=>String(o.categories||"").split(",").forEach(c=>{ const v=c.trim(); if(v && v!=="No Issues") rsnSet.add(v); }));
  const rsnOpt=Array.from(rsnSet).sort();

  const orders=getFilteredQaOrders();
  const imgs=orders.reduce((a,o)=>a+o.volume,0), errs=orders.reduce((a,o)=>a+o.errors,0);
  const clean=orders.filter(o=>o.errors===0).length, rate=imgs?(errs/imgs*100):0;

  const activeFilters=[];
  if(qaClient!=="ALL") activeFilters.push(`Client: ${qaClient}`);
  if(qaDes!=="ALL") activeFilters.push(`Designer: ${qaDes}`);
  if(qaQc!=="ALL") activeFilters.push(`QC: ${qaQc}`);
  if(qaRsn!=="ALL") activeFilters.push(`Reason: ${qaRsn}`);

  let h=`<div class="title"><h2>QA Audit — ${esc($("#qaAud").value)}</h2><span>${esc(r.label)}${activeFilters.length ? " · <b>"+esc(activeFilters.join(" | "))+"</b>" : ""}</span><div class="grow"></div>
    ${expDropdownHtml("QA")}</div>`;
  h+=`<div class="card filter-bar">
    <div class="field"><label for="qaClientSel">Client</label>
      <select id="qaClientSel" onchange="applyQaDropdownFilters()">
        <option value="ALL">All Clients (${clientsOpt.length})</option>
        ${clientsOpt.map(c=>`<option value="${esc(c)}" ${qaClient===c?'selected':''}>${esc(c)}</option>`).join("")}
      </select>
    </div>
    <div class="field"><label for="qaDesSel">Designer</label>
      <select id="qaDesSel" onchange="applyQaDropdownFilters()">
        <option value="ALL">All Designers (${desOpt.length})</option>
        ${desOpt.map(d=>`<option value="${esc(d)}" ${qaDes===d?'selected':''}>${esc(d)}</option>`).join("")}
      </select>
    </div>
    <div class="field"><label for="qaQcSel">QC Name</label>
      <select id="qaQcSel" onchange="applyQaDropdownFilters()">
        <option value="ALL">All QCs (${qcOpt.length})</option>
        ${qcOpt.map(q=>`<option value="${esc(q)}" ${qaQc===q?'selected':''}>${esc(q)}</option>`).join("")}
      </select>
    </div>
    <div class="field"><label for="qaRsnSel">Reason / Category</label>
      <select id="qaRsnSel" onchange="applyQaDropdownFilters()">
        <option value="ALL">All Reasons (${rsnOpt.length})</option>
        ${rsnOpt.map(x=>`<option value="${esc(x)}" ${qaRsn===x?'selected':''}>${esc(x)}</option>`).join("")}
      </select>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      <div class="seg">
        <button class="${qaFault==='ALL'?'on':''}" onclick="qaFault='ALL';renderQaReport()">All</button>
        <button class="${qaFault==='PROD'?'on':''}" onclick="qaFault='PROD';renderQaReport()">Prod</button>
        <button class="${qaFault==='CUST'?'on':''}" onclick="qaFault='CUST';renderQaReport()">Cust</button>
      </div>
      <button class="btn ghost" onclick="resetQaFilters()">Reset</button>
    </div>
  </div>`;
  h+=renderStreakAlertsCard(s.streak_alerts, s.streak_days||3, "QA Audit");
  h+=`<div class="kpis">
    <div class="card kpi"><div class="k">Images Audited</div><div class="v">${nf(imgs)}</div><div class="s">${nf(orders.length)} orders</div></div>
    <div class="card kpi"><div class="k">Errors Found</div><div class="v bad">${nf(errs)}</div><div class="s">${nf(orders.length-clean)} orders with issues</div></div>
    <div class="card kpi"><div class="k">Defect Rate</div><div class="v ${rate>s.escalate_above?'bad':'good'}">${rate.toFixed(2)}%</div><div class="s">threshold ${s.escalate_above}%</div></div>
    <div class="card kpi"><div class="k">Clean Orders</div><div class="v good">${nf(clean)}</div><div class="s">${orders.length?Math.round(clean/orders.length*100):0}% clean</div></div>
  </div>`;
  h+=renderTakeawaysCard("QA", s.takeaways, "Key Takeaways & Action Plan");
  h+=`<div class="card" style="margin-top:14px"><h3>Job Audit Details (${orders.length})</h3><div style="max-height:340px;overflow:auto"><table>
    <thead><tr><th>Client</th><th>Order ID</th><th class="num">Images</th><th class="num">Errors</th><th>Category</th><th>Designer</th><th>QC</th></tr></thead>
    <tbody>${orders.map(o=>`<tr><td><b>${esc(o.client)}</b></td><td>${esc(o.order_id)}</td><td class="num">${nf(o.volume)}</td><td class="num ${o.errors?'bad':''}">${nf(o.errors)}</td><td>${esc(o.categories)}</td><td>${esc(o.designer)}</td><td>${esc(o.qc)}</td></tr>`).join("")}</tbody>
  </table></div></div>`;
  h+=renderMinimizedSummaryCard("qaSumBox", r.text);
  el.innerHTML=h;
}

// ==================== TAB 2: CLIENT REJECTION REPORT ====================
function setCrRange(r){crRange=r;$("#crDay").classList.toggle("on",!r);$("#crRange").classList.toggle("on",r);$("#crToWrap").classList.toggle("hide",!r);$("#crFromLbl").textContent=r?"From":"Date";if(!r)$("#crTo").value="";}
function crQuick(q){
  const lat=(ST&&ST.cr_latest)||iso(new Date()), d=new Date(lat+"T00:00:00");
  if(q==="latest"){setCrRange(false);$("#crFrom").value=lat;}
  else if(q==="last7"){const s=new Date(d);s.setDate(s.getDate()-6);setCrRange(true);$("#crFrom").value=iso(s);$("#crTo").value=lat;}
  else{const s=new Date(d.getFullYear(),d.getMonth(),1);setCrRange(true);$("#crFrom").value=iso(s);$("#crTo").value=lat;}
  runCrReport();
}
async function runCrReport(resetFilters=true){
  const f=$("#crFrom").value, t=crRange?$("#crTo").value:"";
  if(!f) return;
  if(resetFilters){ crClient="ALL"; crDes="ALL"; crQc="ALL"; crRsn="ALL"; crFault="ALL"; CUSTOM_TK.CR=null; EDIT_TK_MODE.CR=false; CUSTOM_SUM_TXT.crSumBox=null; }
  $("#crResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Building Client Rejection report…</div>';
  CUR_CR=await api("/api/cr_report",{
    start:f, end:t||f,
    client:crClient, designer:crDes, qc:crQc, reason:crRsn
  });
  renderCrReport();
}
function applyCrDropdownFilters(){
  crClient=$("#crClientSel") ? $("#crClientSel").value : "ALL";
  crDes=$("#crDesSel") ? $("#crDesSel").value : "ALL";
  crQc=$("#crQcSel") ? $("#crQcSel").value : "ALL";
  crRsn=$("#crRsnSel") ? $("#crRsnSel").value : "ALL";
  runCrReport(false);
}
function resetCrFilters(){
  crClient="ALL"; crDes="ALL"; crQc="ALL"; crRsn="ALL"; crFault="ALL";
  runCrReport(true);
}
function getFilteredCrOrders(){
  if(!CUR_CR || CUR_CR.error) return [];
  let orders=(CUR_CR.summary && CUR_CR.summary.orders_list) || [];
  if(crClient!=="ALL") orders=orders.filter(o=>o.client===crClient);
  if(crDes!=="ALL") orders=orders.filter(o=>o.designer===crDes);
  if(crQc!=="ALL") orders=orders.filter(o=>o.qc===crQc);
  if(crRsn!=="ALL") orders=orders.filter(o=>o.reason===crRsn);
  if(crFault==="PROD") orders=orders.filter(o=>!o.is_customer);
  if(crFault==="CUST") orders=orders.filter(o=>o.is_customer);
  return orders;
}
function renderCrReport(){
  const el=$("#crResult"), r=CUR_CR;
  if(!r) return;
  if(r.error){el.innerHTML=`<div class="card empty" style="margin-top:16px"><b>No Client Rejections</b><br>${esc(r.error)}</div>`;return;}
  const s=r.summary;
  const allOrders=s.orders_list||[];
  const clientsOpt=Array.from(new Set(allOrders.map(o=>o.client).filter(Boolean))).sort();
  const desOpt=Array.from(new Set(allOrders.map(o=>o.designer).filter(v=>v && v!=="Unassigned"))).sort();
  const qcOpt=Array.from(new Set(allOrders.map(o=>o.qc).filter(v=>v && v!=="Unassigned"))).sort();
  const rsnOpt=Array.from(new Set(allOrders.map(o=>o.reason).filter(v=>v && v!=="No reason"))).sort();

  const orders=getFilteredCrOrders();
  const imgs=orders.reduce((a,o)=>a+o.volume,0);
  const prod=orders.filter(o=>!o.is_customer).length, cust=orders.filter(o=>o.is_customer).length;
  const delayed=orders.filter(o=>o.delay).length;
  const crTakeaways=buildCrTakeaways(orders, s.streak_alerts, s.streak_days||3);

  const activeFilters=[];
  if(crClient!=="ALL") activeFilters.push(`Client: ${crClient}`);
  if(crDes!=="ALL") activeFilters.push(`Designer: ${crDes}`);
  if(crQc!=="ALL") activeFilters.push(`QC: ${crQc}`);
  if(crRsn!=="ALL") activeFilters.push(`Reason: ${crRsn}`);

  let h=`<div class="title"><h2>Client Rejection Breakdown</h2><span>${esc(r.label)}${activeFilters.length ? " · <b>"+esc(activeFilters.join(" | "))+"</b>" : ""}</span><div class="grow"></div>
    ${expDropdownHtml("CR")}</div>`;
  h+=`<div class="card filter-bar">
    <div class="field"><label for="crClientSel">Client</label>
      <select id="crClientSel" onchange="applyCrDropdownFilters()">
        <option value="ALL">All Clients (${clientsOpt.length})</option>
        ${clientsOpt.map(c=>`<option value="${esc(c)}" ${crClient===c?'selected':''}>${esc(c)}</option>`).join("")}
      </select>
    </div>
    <div class="field"><label for="crDesSel">Designer</label>
      <select id="crDesSel" onchange="applyCrDropdownFilters()">
        <option value="ALL">All Designers (${desOpt.length})</option>
        ${desOpt.map(d=>`<option value="${esc(d)}" ${crDes===d?'selected':''}>${esc(d)}</option>`).join("")}
      </select>
    </div>
    <div class="field"><label for="crQcSel">QC Name</label>
      <select id="crQcSel" onchange="applyCrDropdownFilters()">
        <option value="ALL">All QCs (${qcOpt.length})</option>
        ${qcOpt.map(q=>`<option value="${esc(q)}" ${crQc===q?'selected':''}>${esc(q)}</option>`).join("")}
      </select>
    </div>
    <div class="field"><label for="crRsnSel">Reason</label>
      <select id="crRsnSel" onchange="applyCrDropdownFilters()">
        <option value="ALL">All Reasons (${rsnOpt.length})</option>
        ${rsnOpt.map(x=>`<option value="${esc(x)}" ${crRsn===x?'selected':''}>${esc(x)}</option>`).join("")}
      </select>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      <div class="seg">
        <button class="${crFault==='ALL'?'on':''}" onclick="crFault='ALL';renderCrReport()">All</button>
        <button class="${crFault==='PROD'?'on':''}" onclick="crFault='PROD';renderCrReport()">Prod</button>
        <button class="${crFault==='CUST'?'on':''}" onclick="crFault='CUST';renderCrReport()">Cust</button>
      </div>
      <button class="btn ghost" onclick="resetCrFilters()">Reset</button>
    </div>
  </div>`;
  h+=renderStreakAlertsCard(s.streak_alerts, s.streak_days||3, "Client Rejections");
  h+=`<div class="kpis">
    <div class="card kpi"><div class="k">Rejected Orders</div><div class="v bad">${nf(orders.length)}</div><div class="s">${nf(imgs)} total images</div></div>
    <div class="card kpi"><div class="k">Production Faults</div><div class="v warnc">${nf(prod)}</div><div class="s">internal fault orders</div></div>
    <div class="card kpi"><div class="k">Customer Faults</div><div class="v">${nf(cust)}</div><div class="s">client fault orders</div></div>
    <div class="card kpi"><div class="k">Delayed Orders</div><div class="v">${nf(delayed)}</div><div class="s">orders with delay reason</div></div>
  </div>`;
  h+=renderTakeawaysCard("CR", crTakeaways, "Key Takeaways & Action Plan");
  h+=`<div class="card" style="margin-top:14px"><h3>Rejected Job Details (${orders.length})</h3><div style="max-height:340px;overflow:auto"><table>
    <thead><tr><th>Client</th><th>Job ID</th><th class="num">Images</th><th>Fault &amp; Reason</th><th>Designer</th><th>QC Name</th></tr></thead>
    <tbody>${orders.map(o=>`<tr><td><b>${esc(o.client)}</b></td><td>${esc(o.job_id)}</td><td class="num">${nf(o.volume)}</td><td><span class="pill ${o.is_customer?'warn':'bad'}">${o.is_customer?'CUSTOMER':'PROD'}</span> ${esc(o.reason)}</td><td>${esc(o.designer)}</td><td>${esc(o.qc)}</td></tr>`).join("")}</tbody>
  </table></div></div>`;
  h+=renderMinimizedSummaryCard("crSumBox", r.text);
  el.innerHTML=h;
}
function buildCrTakeaways(orders, streakAlerts=[], streakDays=3){
  if(!orders||!orders.length) return [];
  const byC={}, byQc={}, byRsn={};
  orders.forEach(o=>{
    byC[o.client]=(byC[o.client]||0)+1;
    if(o.qc && o.qc!=="Unassigned") byQc[o.qc]=(byQc[o.qc]||0)+1;
    if(o.reason && o.reason!=="No reason") byRsn[o.reason]=(byRsn[o.reason]||0)+1;
  });
  const topC=Object.entries(byC).sort((a,b)=>b[1]-a[1])[0];
  const topQcs=Object.entries(byQc).sort((a,b)=>b[1]-a[1]).slice(0,3).map(([k,v])=>`${k} (${v})`);
  const topRsns=Object.entries(byRsn).sort((a,b)=>b[1]-a[1]).slice(0,3).map(([k,v])=>`${k} (${v})`);
  const rows=[
    ["Primary Rejection Source", `${topC[0]} accounted for ${Math.round(topC[1]/orders.length*100)}% of rejections (${topC[1]}/${orders.length} orders). Top reasons: ${topRsns.join(", ")||"N/A"}.`],
    ["QC Accountability", topQcs.length ? `Rejected jobs passed through QC reviewers: ${topQcs.join(", ")}.` : "No specific QC reviewer tagged on rejected orders."],
    ["Immediate Action Plan", "Review recurring rejection categories with production leads and enforce pre-delivery checklist on affected client specs."]
  ];
  if(streakAlerts && streakAlerts.length){
    rows.push([
      `${streakDays}-Day Recurring Issue Reminders`,
      streakAlerts.map(a=>`${a.dimension} ${a.entity}: ${a.issue} (${a.days}d)`).join(" | ")
    ]);
  }
  return rows;
}

// ==================== TAB 3: CLIENT ATTENTION RADAR ====================
let attnCustomActive=false;
async function loadAttention(customPayload=null){
  $("#attnResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Analyzing client risk across periods…</div>';
  CUSTOM_TK.ATTN=null; EDIT_TK_MODE.ATTN=false;
  const aud=$("#attnAudSel") ? $("#attnAudSel").value : "ALL";
  const reqPayload=Object.assign({auditor: aud}, customPayload||{});
  CUR_ATTN=await api("/api/attention", reqPayload);
  if(CUR_ATTN && CUR_ATTN.auditors && $("#attnAudSel")){
    const curA=$("#attnAudSel").value||"ALL";
    $("#attnAudSel").innerHTML='<option value="ALL">All Auditors</option>' + CUR_ATTN.auditors.map(a=>`<option value="${esc(a)}">${esc(a.toUpperCase())}</option>`).join("");
    $("#attnAudSel").value=CUR_ATTN.auditors.includes(curA)?curA:"ALL";
  }
  if(CUR_ATTN && CUR_ATTN.anchor && !attnCustomActive){
    if(!$("#attnFrom").value) $("#attnFrom").value=CUR_ATTN.anchor;
    if(!$("#attnTo").value) $("#attnTo").value=CUR_ATTN.anchor;
  }
  renderAttention();
}
async function onAttnAuditorChange(){
  if(attnCustomActive){
    const s=$("#attnFrom").value, e=$("#attnTo").value||s;
    await loadAttention({start:s, end:e});
    if(CUR_ATTN && CUR_ATTN.windows && CUR_ATTN.windows.CUSTOM) setAttnWindow("CUSTOM");
  } else {
    await loadAttention();
  }
}
async function applyAttnDateRange(){
  const s=$("#attnFrom").value, e=$("#attnTo").value||s;
  if(!s){ toast("Select a From Date first"); return; }
  if(!$("#attnTo").value) $("#attnTo").value=s;
  attnCustomActive=true;
  await loadAttention({start:s, end:e});
  if(CUR_ATTN && CUR_ATTN.windows && CUR_ATTN.windows.CUSTOM){
    $("#attnCUSTOM").classList.remove("hide");
    setAttnWindow("CUSTOM");
  }
}
async function resetAttnRange(){
  attnCustomActive=false;
  $("#attnCUSTOM").classList.add("hide");
  if($("#attnClientSel")) $("#attnClientSel").value="ALL";
  if($("#attnAudSel")) $("#attnAudSel").value="ALL";
  attnFilter="ALL";
  ["ALL","CRITICAL","WATCH","HEALTHY"].forEach(k=>$("#attnF_"+k).classList.toggle("on",k==="ALL"));
  await loadAttention();
  if(CUR_ATTN && CUR_ATTN.anchor){
    $("#attnFrom").value=CUR_ATTN.anchor;
    $("#attnTo").value=CUR_ATTN.anchor;
  }
  setAttnWindow("DAILY");
}
function setAttnWindow(w){
  attnWin=w;
  CUSTOM_TK.ATTN=null; EDIT_TK_MODE.ATTN=false;
  ["DAILY","WEEKLY","MONTHLY","CUSTOM"].forEach(k=>{
    const btn=$("#attn"+k);
    if(btn) btn.classList.toggle("on",k===w);
  });
  renderAttention();
}
function setAttnFilter(f){
  attnFilter=f;
  CUSTOM_TK.ATTN=null; EDIT_TK_MODE.ATTN=false;
  ["ALL","CRITICAL","WATCH","HEALTHY"].forEach(k=>{
    const btn=$("#attnF_"+k);
    if(btn) btn.classList.toggle("on",k===f);
  });
  renderAttention();
}
function buildAttnTakeaways(filtered){
  const crit=filtered.filter(c=>c.status==="CRITICAL");
  const watch=filtered.filter(c=>c.status==="WATCH");
  return [
    ["Immediate Escalation Clients", crit.length ? crit.map(c=>`${c.client} (${c.flags.join("; ")})`).join(" | ") : "None — no clients in critical threshold."],
    ["Watchlist Monitoring", watch.length ? watch.map(c=>`${c.client} (${c.flags.join("; ")})`).join(" | ") : "No watchlist clients."],
    ["Recommended Action Plan", "Prioritize immediate QA spot-checks and designer/QC coaching on critical clients above 10% defect rate or spiking external rejections."]
  ];
}
function renderAttention(){
  const el=$("#attnResult");
  if(!CUR_ATTN||!CUR_ATTN.windows) return;
  const win=CUR_ATTN.windows[attnWin] || CUR_ATTN.windows.DAILY;
  const allClients=win.clients||[];

  $("#attnF_ALL").textContent=`All Clients (${allClients.length})`;
  $("#attnF_CRITICAL").textContent=`Needs Immediate Attention (${win.critical_count||0})`;
  $("#attnF_WATCH").textContent=`Watchlist (${win.watch_count||0})`;
  if($("#attnF_HEALTHY")) $("#attnF_HEALTHY").textContent=`Healthy (${win.healthy_count||0})`;

  const cSel=$("#attnClientSel");
  if(cSel){
    const prevC=cSel.value||"ALL";
    const names=allClients.map(c=>c.client).sort();
    cSel.innerHTML=`<option value="ALL">All Clients (${names.length})</option>` + names.map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join("");
    cSel.value=names.includes(prevC)?prevC:"ALL";
  }

  let list=allClients;
  if(attnFilter==="CRITICAL") list=list.filter(c=>c.status==="CRITICAL");
  else if(attnFilter==="WATCH") list=list.filter(c=>c.status==="WATCH");
  else if(attnFilter==="HEALTHY") list=list.filter(c=>c.status==="HEALTHY");

  if(cSel && cSel.value!=="ALL"){
    list=list.filter(c=>c.client===cSel.value);
  }

  let h=`<div class="title"><h2>Client Attention Radar</h2><span>${esc(win.label)}</span><div class="grow"></div>
    ${expDropdownHtml("ATTN")}</div>`;

  h+=`<div class="kpis" style="grid-template-columns:repeat(3,1fr);margin-bottom:14px">
    <div class="card kpi"><div class="k">Needs Immediate Attention</div><div class="v bad">${win.critical_count}</div><div class="s">High defect rate or rejection spike</div></div>
    <div class="card kpi"><div class="k">Watchlist Clients</div><div class="v warnc">${win.watch_count}</div><div class="s">Moderate issues or upward trend</div></div>
    <div class="card kpi"><div class="k">Healthy / Stable Clients</div><div class="v good">${win.healthy_count}</div><div class="s">Operating within normal thresholds</div></div>
  </div>`;

  h+=renderTakeawaysCard("ATTN", buildAttnTakeaways(list), "Key Takeaways & Action Plan");

  if(!list.length){
    h+=`<div class="card empty" style="margin-top:14px">No clients match this filter for the selected period.</div>`;
    el.innerHTML=h; return;
  }

  h+=`<div style="display:flex;flex-direction:column;gap:12px;margin-top:14px">`;
  list.forEach(c=>{
    const badgeCls=c.status==="CRITICAL"?"bad":(c.status==="WATCH"?"warn":"good");
    const statusLbl=c.status==="CRITICAL"?"ACTION NEEDED":(c.status==="WATCH"?"WATCHLIST":"HEALTHY");
    h+=`<div class="card attn-card ${c.status}">
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px">
        <span class="pill ${badgeCls}">${statusLbl}</span>
        <strong style="font-size:16px">${esc(c.client)}</strong>
        <span class="note">Priority Score: <b>${c.score}</b></span>
        <div class="grow"></div>
        <span class="note">${c.flags.map(f=>esc(f)).join(" · ")}</span>
      </div>
      <div class="grid2" style="margin-top:8px">
        <div style="background:var(--bg);padding:12px 15px;border-radius:10px;border:1px solid var(--line)">
          <div class="lbl" style="color:var(--brand)">Internal QA Audit Breakdown</div>
          <div style="font-size:20px;font-weight:700;margin-top:4px">${c.qa_a_rate.toFixed(2)}% ${deltaBadge(c.qa_a_rate,c.qa_b_rate,true,true)}</div>
          <div class="note">${nf(c.qa_a_err)} errors / ${nf(c.qa_a_imgs)} imgs (Prior: ${c.qa_b_rate.toFixed(2)}%)</div>
          <div style="margin-top:9px;padding-top:8px;border-top:1px solid var(--line);font-size:12.5px;line-height:1.5">
            <div><span class="note">Reasons:</span> <b>${c.qa_reasons.length ? esc(c.qa_reasons.join(", ")) : '<span class="note">None</span>'}</b></div>
            <div><span class="note">Designers:</span> <b>${c.qa_designers.length ? esc(c.qa_designers.join(", ")) : '<span class="note">None</span>'}</b></div>
            <div><span class="note">QC Names:</span> <b>${c.qa_qcs.length ? esc(c.qa_qcs.join(", ")) : '<span class="note">None</span>'}</b></div>
          </div>
        </div>
        <div style="background:var(--bg);padding:12px 15px;border-radius:10px;border:1px solid var(--line)">
          <div class="lbl" style="color:var(--warn)">External Client Rejections Breakdown</div>
          <div style="font-size:20px;font-weight:700;margin-top:4px">${nf(c.cr_a_ord)} orders ${deltaBadge(c.cr_a_ord,c.cr_b_ord,true,false)}</div>
          <div class="note">${nf(c.cr_a_imgs)} imgs · Prod: ${c.cr_a_prod} / Cust: ${c.cr_a_cust} (Prior: ${c.cr_b_ord} ord)</div>
          <div style="margin-top:9px;padding-top:8px;border-top:1px solid var(--line);font-size:12.5px;line-height:1.5">
            <div><span class="note">Reasons:</span> <b>${c.cr_reasons.length ? esc(c.cr_reasons.join(", ")) : '<span class="note">None</span>'}</b></div>
            <div><span class="note">Designers:</span> <b>${c.cr_designers.length ? esc(c.cr_designers.join(", ")) : '<span class="note">None</span>'}</b></div>
            <div><span class="note">QC Names:</span> <b>${c.cr_qcs.length ? esc(c.cr_qcs.join(", ")) : '<span class="note">None</span>'}</b></div>
          </div>
        </div>
      </div>
    </div>`;
  });
  h+=`</div>`;
  el.innerHTML=h;
}

// ==================== TAB 4: INTERACTIVE CHARTS ====================
let chartUseCustomDates=false;
function setChartPreset(days){
  chartDays=days;
  chartUseCustomDates=false;
  $("#chFrom").value="";
  $("#chTo").value="";
  [7,30,90,0].forEach(d=>$("#ch"+d).classList.toggle("on",d===days));
  loadCharts(days);
}
function onChartDateChange(){
  chartUseCustomDates=true;
  [7,30,90,0].forEach(d=>$("#ch"+d).classList.remove("on"));
}
function resetChartFilters(){
  $("#chAud").value="ALL";
  $("#chClient").value="ALL";
  $("#chDes").value="ALL";
  $("#chQc").value="ALL";
  $("#chRsn").value="ALL";
  setChartPreset(30);
}
function populateSelect(selId, items, allLabel){
  const el=$(selId);
  if(!el) return;
  const cur=el.value||"ALL";
  el.innerHTML=`<option value="ALL">${allLabel} (${items.length})</option>` + items.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join("");
  el.value=items.includes(cur)?cur:"ALL";
}
async function loadCharts(daysOverride){
  if(typeof daysOverride==="number") chartDays=daysOverride;
  CUSTOM_TK.CHARTS=null; EDIT_TK_MODE.CHARTS=false;
  $("#chartsResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Rendering analytics charts…</div>';
  const payload={
    days: chartDays,
    auditor: $("#chAud").value || "ALL",
    client: $("#chClient").value || "ALL",
    designer: $("#chDes").value || "ALL",
    qc: $("#chQc").value || "ALL",
    reason: $("#chRsn").value || "ALL",
  };
  if(chartUseCustomDates && $("#chFrom").value){
    payload.start=$("#chFrom").value;
    payload.end=$("#chTo").value || $("#chFrom").value;
  }
  CUR_CHARTS=await api("/api/charts", payload);
  if(CUR_CHARTS && CUR_CHARTS.options){
    populateSelect("#chAud", CUR_CHARTS.options.auditors||[], "All Auditors");
    populateSelect("#chClient", CUR_CHARTS.options.clients||[], "All Clients");
    populateSelect("#chDes", CUR_CHARTS.options.designers||[], "All Designers");
    populateSelect("#chQc", CUR_CHARTS.options.qcs||[], "All QCs");
    populateSelect("#chRsn", CUR_CHARTS.options.reasons||[], "All Reasons");
  }
  if(CUR_CHARTS && !chartUseCustomDates){
    if(CUR_CHARTS.start_date) $("#chFrom").value=CUR_CHARTS.start_date;
    if(CUR_CHARTS.end_date) $("#chTo").value=CUR_CHARTS.end_date;
  }
  renderCharts();
}
function buildChartsTakeaways(d){
  const k=d.kpis||{};
  const topDes=(d.top_designers||[]).slice(0,3).map(x=>`${x.name} (${x.qa_err} QA err, ${x.cr_ord} CR ord)`).join(", ");
  const topRsn=(d.top_reasons||[]).slice(0,3).map(x=>`${x.reason} (${x.count})`).join(", ");
  return [
    ["Period Defect & Rejection Summary", `Audited ${nf(k.qa_imgs||0)} images with ${nf(k.qa_err||0)} errors (${k.qa_rate||0}% defect rate). External client rejections totaled ${nf(k.cr_ord||0)} orders (${nf(k.cr_imgs||0)} images).`],
    ["Top Error Contributors", topDes ? `Highest combined defect & rejection volume: ${topDes}.` : "No designer concentration detected in this window."],
    ["Dominant Rejection Drivers", topRsn ? `Most frequent categories: ${topRsn}. Target these issues in upcoming QC calibration.` : "No dominant rejection categories recorded."]
  ];
}
function renderCharts(){
  const d=CUR_CHARTS; if(!d) return;
  const activeChips=[];
  if($("#chAud").value!=="ALL") activeChips.push(`Auditor: ${$("#chAud").value.toUpperCase()}`);
  if($("#chClient").value!=="ALL") activeChips.push(`Client: ${$("#chClient").value}`);
  if($("#chDes").value!=="ALL") activeChips.push(`Designer: ${$("#chDes").value}`);
  if($("#chQc").value!=="ALL") activeChips.push(`QC: ${$("#chQc").value}`);
  if($("#chRsn").value!=="ALL") activeChips.push(`Reason: ${$("#chRsn").value}`);
  $("#chartSpanNote").innerHTML=`Showing data for <b>${esc(d.span_label)}</b>` + (activeChips.length ? ` · Filtered by <b>${esc(activeChips.join(" | "))}</b>` : "");

  const k=d.kpis||{};
  const trend=d.trend||[];
  const clients=d.clients||[];
  const f=d.faults||{qa_prod:0,qa_cust:0,cr_prod:0,cr_cust:0};

  const W=640, H=210, padL=36, padR=14, padT=16, padB=28;
  const maxRate=Math.max(15, ...trend.map(t=>t.qa_rate));
  const stepX=trend.length>1?(W-padL-padR)/(trend.length-1):0;
  const pts=trend.map((t,i)=>{
    const x=padL + i*stepX;
    const y=padT + (H-padT-padB)*(1 - t.qa_rate/maxRate);
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  }).join(" ");
  const threshY=padT + (H-padT-padB)*(1 - 10/maxRate);

  const qaSvg=`<svg class="chart" viewBox="0 0 ${W} ${H}">
    <line x1="${padL}" y1="${threshY}" x2="${W-padR}" y2="${threshY}" stroke="#b91c1c" stroke-dasharray="4 4" stroke-width="1.2"/>
    <text x="${W-padR-4}" y="${threshY-4}" text-anchor="end" fill="#b91c1c" font-size="10" font-weight="700">10% Escalation Threshold</text>
    ${trend.length>1?`<polyline fill="none" stroke="#2563eb" stroke-width="2.5" points="${pts}"/>`:""}
    ${trend.map((t,i)=>{
      const x=padL+i*stepX, y=padT+(H-padT-padB)*(1-t.qa_rate/maxRate);
      const showLbl=trend.length<=14 || i%Math.ceil(trend.length/10)===0;
      return `<circle cx="${x}" cy="${y}" r="3.5" fill="${t.qa_rate>10?'#b91c1c':'#2563eb'}"><title>${t.full_date}: ${t.qa_rate}% (${t.qa_err}/${t.qa_imgs})</title></circle>
      ${showLbl?`<text x="${x}" y="${H-8}" text-anchor="middle" fill="#64748b" font-size="9.5">${t.date}</text>`:""}`;
    }).join("")}
  </svg>`;

  const maxCr=Math.max(4, ...trend.map(t=>t.cr_ord));
  const barW=trend.length?Math.max(6, Math.min(24, (W-padL-padR)/trend.length - 4)):14;
  const crSvg=`<svg class="chart" viewBox="0 0 ${W} ${H}">
    ${trend.map((t,i)=>{
      const slot=(W-padL-padR)/Math.max(1,trend.length);
      const x=padL + i*slot + (slot-barW)/2;
      const bh=(H-padT-padB)*(t.cr_ord/maxCr);
      const y=H-padB-bh;
      const showLbl=trend.length<=14 || i%Math.ceil(trend.length/10)===0;
      return `<rect x="${x}" y="${y}" width="${barW}" height="${bh}" rx="3" fill="${t.cr_ord>=5?'#b91c1c':'#f59e0b'}"><title>${t.full_date}: ${t.cr_ord} orders (${t.cr_imgs} imgs)</title></rect>
      ${showLbl?`<text x="${x+barW/2}" y="${H-8}" text-anchor="middle" fill="#64748b" font-size="9.5">${t.date}</text>`:""}`;
    }).join("")}
  </svg>`;

  const maxClientScore=Math.max(1, ...clients.map(c=>c.cr_ord*3 + c.qa_err));
  const clientBars=clients.map(c=>{
    const pctW=Math.min(100, Math.round(((c.cr_ord*3 + c.qa_err)/maxClientScore)*100));
    const cls=c.qa_rate>10||c.cr_ord>=5?"bad":(c.qa_rate>5||c.cr_ord>=2?"warn":"");
    return `<div class="bar-row"><div class="bar-lbl" title="${esc(c.client)}">${esc(c.client)}</div>
      <div class="bar-track"><div class="bar-fill ${cls}" style="width:${pctW}%"></div></div>
      <div class="bar-val">${c.cr_ord} CR · ${c.qa_rate}%</div></div>`;
  }).join("") || '<div class="note">No client volume in this range.</div>';

  const totFault=(f.qa_prod+f.qa_cust+f.cr_prod+f.cr_cust)||1;
  const faultHtml=`
    <div class="bar-row"><div class="bar-lbl">QA Production Fault</div><div class="bar-track"><div class="bar-fill bad" style="width:${Math.round(f.qa_prod/totFault*100)}%"></div></div><div class="bar-val">${nf(f.qa_prod)} imgs</div></div>
    <div class="bar-row"><div class="bar-lbl">QA Customer Fault</div><div class="bar-track"><div class="bar-fill" style="width:${Math.round(f.qa_cust/totFault*100)}%"></div></div><div class="bar-val">${nf(f.qa_cust)} imgs</div></div>
    <div class="bar-row"><div class="bar-lbl">CR Production Fault</div><div class="bar-track"><div class="bar-fill warn" style="width:${Math.round(f.cr_prod/totFault*100)}%"></div></div><div class="bar-val">${nf(f.cr_prod)} ord</div></div>
    <div class="bar-row"><div class="bar-lbl">CR Customer Fault</div><div class="bar-track"><div class="bar-fill" style="width:${Math.round(f.cr_cust/totFault*100)}%"></div></div><div class="bar-val">${nf(f.cr_cust)} ord</div></div>
  `;

  const maxDes=Math.max(1, ...(d.top_designers||[]).map(x=>x.qa_err + x.cr_ord*3));
  const desHtml=(d.top_designers||[]).map(x=>{
    const w=Math.round(((x.qa_err+x.cr_ord*3)/maxDes)*100);
    return `<div class="bar-row"><div class="bar-lbl">${esc(x.name)}</div><div class="bar-track"><div class="bar-fill bad" style="width:${w}%"></div></div><div class="bar-val">${x.qa_err} err · ${x.cr_ord} CR</div></div>`;
  }).join("") || '<div class="note">No designer errors recorded.</div>';

  const maxRsn=Math.max(1, ...(d.top_reasons||[]).map(x=>x.count));
  const rsnHtml=(d.top_reasons||[]).map(x=>{
    const w=Math.round((x.count/maxRsn)*100);
    return `<div class="bar-row"><div class="bar-lbl" title="${esc(x.reason)}">${esc(x.reason)}</div><div class="bar-track"><div class="bar-fill warn" style="width:${w}%"></div></div><div class="bar-val">${nf(x.count)}</div></div>`;
  }).join("") || '<div class="note">No rejection reasons recorded.</div>';

  $("#chartsResult").innerHTML=`
    <div class="title"><h2>Analytics &amp; Trends</h2><span>${esc(d.span_label)}</span><div class="grow"></div>
      ${expDropdownHtml("CHARTS")}</div>
    <div class="kpis" style="margin-bottom:14px">
      <div class="card kpi"><div class="k">Filtered QA Images</div><div class="v">${nf(k.qa_imgs||0)}</div><div class="s">Total audited images</div></div>
      <div class="card kpi"><div class="k">Filtered QA Errors</div><div class="v bad">${nf(k.qa_err||0)}</div><div class="s">${(k.qa_rate||0).toFixed(2)}% defect rate</div></div>
      <div class="card kpi"><div class="k">Filtered CR Orders</div><div class="v warnc">${nf(k.cr_ord||0)}</div><div class="s">External client rejections</div></div>
      <div class="card kpi"><div class="k">Filtered CR Images</div><div class="v">${nf(k.cr_imgs||0)}</div><div class="s">Total rejected volume</div></div>
    </div>
    ${renderTakeawaysCard("CHARTS", buildChartsTakeaways(d), "Key Takeaways & Action Plan")}
    <div class="grid2">
      <div class="card"><h3>Daily QA Defect Rate (%) Trend</h3><div class="chart-box">${trend.length?qaSvg:'<div class="note">No QA data in range.</div>'}</div></div>
      <div class="card"><h3>Daily Client Rejection Orders Trend</h3><div class="chart-box">${trend.length?crSvg:'<div class="note">No CR data in range.</div>'}</div></div>
    </div>
    <div class="grid2">
      <div class="card"><h3>Client Defect &amp; Rejection Impact</h3>${clientBars}</div>
      <div class="card"><h3>Production vs Customer Fault Attribution</h3>${faultHtml}</div>
    </div>
    <div class="grid2">
      <div class="card"><h3>Top Designers by Combined Errors &amp; Rejections</h3>${desHtml}</div>
      <div class="card"><h3>Top Rejection &amp; Defect Categories</h3>${rsnHtml}</div>
    </div>
  `;
}

// ==================== TAB 5: PERIOD COMPARE ====================
function setCompMode(m){
  compMode=m;
  CUSTOM_TK.COMP=null; EDIT_TK_MODE.COMP=false;
  $("#cmQA").classList.toggle("on",m==="QA");
  $("#cmCR").classList.toggle("on",m==="CR");
  if($("#compAudWrap")) $("#compAudWrap").classList.toggle("hide", m!=="QA");
  runCompare();
}
function resetCompFilters(){
  if($("#compAud")) $("#compAud").value="ALL";
  compClient="ALL"; compDes="ALL"; compQc="ALL";
  if($("#compClientSel")) $("#compClientSel").value="ALL";
  if($("#compDesSel")) $("#compDesSel").value="ALL";
  if($("#compQcSel")) $("#compQcSel").value="ALL";
  runCompare();
}
function applyCompPreset(p){
  ["DAY","WEEK","MONTH"].forEach(x=>$("#cp"+x).classList.toggle("on",x===p));
  const meta=(ST&&ST.auditors||[]).find(x=>x.name===$("#qaAud").value);
  const latStr=compMode==="QA"?((meta&&meta.latest)||iso(new Date())):((ST&&ST.cr_latest)||iso(new Date()));
  const d=new Date(latStr+"T00:00:00");
  if(p==="DAY"){
    const prev=new Date(d); prev.setDate(prev.getDate()-1);
    $("#caFrom").value=latStr; $("#caTo").value=latStr;
    $("#cbFrom").value=iso(prev); $("#cbTo").value=iso(prev);
  }else if(p==="WEEK"){
    const aStart=new Date(d); aStart.setDate(aStart.getDate()-6);
    const bEnd=new Date(aStart); bEnd.setDate(bEnd.getDate()-1);
    const bStart=new Date(bEnd); bStart.setDate(bStart.getDate()-6);
    $("#caFrom").value=iso(aStart); $("#caTo").value=latStr;
    $("#cbFrom").value=iso(bStart); $("#cbTo").value=iso(bEnd);
  }else{
    const aStart=new Date(d.getFullYear(),d.getMonth(),1);
    const bEnd=new Date(d.getFullYear(),d.getMonth(),0);
    const bStart=new Date(bEnd.getFullYear(),bEnd.getMonth(),1);
    $("#caFrom").value=iso(aStart); $("#caTo").value=latStr;
    $("#cbFrom").value=iso(bStart); $("#cbTo").value=iso(bEnd);
  }
  runCompare();
}
async function runCompare(){
  const a1=$("#caFrom").value, a2=$("#caTo").value||a1, b1=$("#cbFrom").value, b2=$("#cbTo").value||b1;
  if(!a1||!b1) return;
  CUSTOM_TK.COMP=null; EDIT_TK_MODE.COMP=false;
  $("#compResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Comparing periods…</div>';
  CUR_COMP=await api("/api/compare",{mode:compMode, auditor:($("#compAud")?$("#compAud").value:"ALL"), a_start:a1, a_end:a2, b_start:b1, b_end:b2});
  if(CUR_COMP && !CUR_COMP.error){
    const all=[...(CUR_COMP.orders_a||[]), ...(CUR_COMP.orders_b||[])];
    const clients=Array.from(new Set(all.map(o=>o.client).filter(Boolean))).sort();
    const des=Array.from(new Set(all.map(o=>o.designer).filter(v=>v && v!=="Unassigned"))).sort();
    const qcs=Array.from(new Set(all.map(o=>o.qc).filter(v=>v && v!=="Unassigned"))).sort();
    populateSelect("#compClientSel", clients, "All Clients");
    populateSelect("#compDesSel", des, "All Designers");
    populateSelect("#compQcSel", qcs, "All QCs");
    compClient=$("#compClientSel").value;
    compDes=$("#compDesSel").value;
    compQc=$("#compQcSel").value;
  }
  renderCompare();
}
function filterCompOrders(list){
  return (list||[]).filter(o=>{
    if(compClient!=="ALL" && o.client!==compClient) return false;
    if(compDes!=="ALL" && o.designer!==compDes) return false;
    if(compQc!=="ALL" && o.qc!==compQc) return false;
    return true;
  });
}
function aggCompRows(oa, ob, mode, keyField){
  const map={};
  const ensure=k=>{ if(!map[k]) map[k]={name:k, a_imgs:0, a_err:0, a_ord:0, b_imgs:0, b_err:0, b_ord:0}; return map[k]; };
  oa.forEach(o=>{
    const k=o[keyField]||"Unassigned";
    const r=ensure(k);
    r.a_imgs+=o.volume; r.a_err+=(o.errors||0); r.a_ord+=1;
  });
  ob.forEach(o=>{
    const k=o[keyField]||"Unassigned";
    const r=ensure(k);
    r.b_imgs+=o.volume; r.b_err+=(o.errors||0); r.b_ord+=1;
  });
  return Object.values(map).map(r=>{
    r.a_rate=r.a_imgs?(r.a_err/r.a_imgs*100):0;
    r.b_rate=r.b_imgs?(r.b_err/r.b_imgs*100):0;
    r.rate_diff=r.a_rate - r.b_rate;
    r.ord_diff=r.a_ord - r.b_ord;
    return r;
  }).sort((x,y)=> mode==="QA" ? (y.a_err - x.a_err || y.a_imgs - x.a_imgs) : (y.a_ord - x.a_ord || y.a_imgs - x.a_imgs));
}
function buildCompareTakeaways(c, oa, ob, rows){
  if(c.mode==="QA"){
    const aImgs=oa.reduce((s,o)=>s+o.volume,0), aErr=oa.reduce((s,o)=>s+o.errors,0);
    const bImgs=ob.reduce((s,o)=>s+o.volume,0), bErr=ob.reduce((s,o)=>s+o.errors,0);
    const aRate=aImgs?(aErr/aImgs*100):0, bRate=bImgs?(bErr/bImgs*100):0;
    const diff=(aRate-bRate).toFixed(2);
    const worsened=rows.filter(r=>r.rate_diff>0.5).slice(0,3).map(r=>`${r.name} (+${r.rate_diff.toFixed(2)}%)`).join(", ");
    return [
      ["Period Defect Rate Shift", `Period A recorded ${aRate.toFixed(2)}% (${nf(aErr)}/${nf(aImgs)}) vs Period B ${bRate.toFixed(2)}% (${nf(bErr)}/${nf(bImgs)}), a net change of ${Number(diff)>0?'+':''}${diff}%.`],
      ["Entities Requiring Attention", worsened ? `Defect rate increased for: ${worsened}.` : "No significant defect rate regression across entities."],
      ["Recommended Action Plan", "Focus immediate feedback on entities showing positive defect delta and reinforce QC gate checks."]
    ];
  } else {
    const diff=oa.length - ob.length;
    const worsened=rows.filter(r=>r.ord_diff>0).slice(0,3).map(r=>`${r.name} (+${r.ord_diff} ord)`).join(", ");
    return [
      ["Period Rejection Volume Shift", `Period A recorded ${nf(oa.length)} rejected orders vs Period B ${nf(ob.length)} orders (${diff>0?'+':''}${nf(diff)} orders).`],
      ["Entities with Rising Rejections", worsened ? `Increased rejections observed in: ${worsened}.` : "Rejection volume remained flat or improved across all entities."],
      ["Recommended Action Plan", "Audit recurring rejection drivers on entities with rising order counts and align QC verification standards."]
    ];
  }
}
function renderCompare(){
  const el=$("#compResult"), c=CUR_COMP;
  if(!c) return;
  if(c.error){el.innerHTML=`<div class="card empty" style="margin-top:16px">${esc(c.error)}</div>`;return;}

  const oa=filterCompOrders(c.orders_a);
  const ob=filterCompOrders(c.orders_b);
  const rows=aggCompRows(oa, ob, c.mode, compGroupBy);
  const grpTitle=compGroupBy==="client"?"Client":(compGroupBy==="designer"?"Designer":"QC Name");

  let h=`<div class="title"><h2>${c.mode==="QA"?"QA Audit":"Client Rejection"} Comparison</h2>
    <span><b>Period A:</b> ${esc(c.a_label)} &nbsp;vs&nbsp; <b>Period B:</b> ${esc(c.b_label)}</span><div class="grow"></div>
    ${expDropdownHtml("COMP")}</div>`;

  if(c.mode==="QA"){
    const aImgs=oa.reduce((s,o)=>s+o.volume,0), aErr=oa.reduce((s,o)=>s+o.errors,0);
    const bImgs=ob.reduce((s,o)=>s+o.volume,0), bErr=ob.reduce((s,o)=>s+o.errors,0);
    const aRate=aImgs?(aErr/aImgs*100):0, bRate=bImgs?(bErr/bImgs*100):0;
    const aClean=oa.filter(o=>o.errors===0).length, bClean=ob.filter(o=>o.errors===0).length;

    h+=`<div class="kpis">
      <div class="card kpi"><div class="k">Defect Rate (A vs B)</div><div class="v ${aRate>c.escalate_above?'bad':'good'}">${aRate.toFixed(2)}% ${deltaBadge(aRate,bRate,true,true)}</div><div class="s">Prev: ${bRate.toFixed(2)}%</div></div>
      <div class="card kpi"><div class="k">Total Errors (A vs B)</div><div class="v bad">${nf(aErr)} ${deltaBadge(aErr,bErr,true,false)}</div><div class="s">Prev: ${nf(bErr)} errors</div></div>
      <div class="card kpi"><div class="k">Images Audited (A vs B)</div><div class="v">${nf(aImgs)} ${deltaBadge(aImgs,bImgs,false,false)}</div><div class="s">Prev: ${nf(bImgs)} images</div></div>
      <div class="card kpi"><div class="k">Clean Orders (A vs B)</div><div class="v good">${nf(aClean)} ${deltaBadge(aClean,bClean,false,false)}</div><div class="s">Prev: ${nf(bClean)} clean</div></div>
    </div>`;
    h+=renderTakeawaysCard("COMP", buildCompareTakeaways(c, oa, ob, rows), "Key Takeaways & Action Plan");
    h+=`<div class="card" style="margin-top:14px">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px">
        <h3 style="margin:0">${grpTitle} Comparison (${rows.length})</h3>
        <div class="seg">
          <button class="${compGroupBy==='client'?'on':''}" onclick="compGroupBy='client';renderCompare()">By Client</button>
          <button class="${compGroupBy==='designer'?'on':''}" onclick="compGroupBy='designer';renderCompare()">By Designer</button>
          <button class="${compGroupBy==='qc'?'on':''}" onclick="compGroupBy='qc';renderCompare()">By QC</button>
        </div>
      </div>
      <table>
      <thead><tr><th>${grpTitle}</th><th class="num">A Imgs</th><th class="num">A Err</th><th class="num">A Rate</th><th class="num">B Rate</th><th class="num">Rate Change</th></tr></thead>
      <tbody>${rows.map(r=>`<tr><td><b>${esc(r.name)}</b></td><td class="num">${nf(r.a_imgs)}</td><td class="num ${r.a_err?'bad':''}">${nf(r.a_err)}</td>
        <td class="num"><b>${r.a_rate.toFixed(2)}%</b></td><td class="num">${r.b_rate.toFixed(2)}%</td><td class="num">${deltaBadge(r.a_rate,r.b_rate,true,true)}</td></tr>`).join("")}</tbody>
    </table></div>`;
  }else{
    const aImgs=oa.reduce((s,o)=>s+o.volume,0), bImgs=ob.reduce((s,o)=>s+o.volume,0);
    const aProd=oa.filter(o=>!o.is_customer).length, bProd=ob.filter(o=>!o.is_customer).length;
    const aCust=oa.filter(o=>o.is_customer).length, bCust=ob.filter(o=>o.is_customer).length;

    h+=`<div class="kpis">
      <div class="card kpi"><div class="k">Rejected Orders (A vs B)</div><div class="v bad">${nf(oa.length)} ${deltaBadge(oa.length,ob.length,true,false)}</div><div class="s">Prev: ${nf(ob.length)} orders</div></div>
      <div class="card kpi"><div class="k">Rejected Images (A vs B)</div><div class="v">${nf(aImgs)} ${deltaBadge(aImgs,bImgs,true,false)}</div><div class="s">Prev: ${nf(bImgs)} images</div></div>
      <div class="card kpi"><div class="k">Production Faults (A vs B)</div><div class="v warnc">${nf(aProd)} ${deltaBadge(aProd,bProd,true,false)}</div><div class="s">Prev: ${nf(bProd)} orders</div></div>
      <div class="card kpi"><div class="k">Customer Faults (A vs B)</div><div class="v">${nf(aCust)} ${deltaBadge(aCust,bCust,true,false)}</div><div class="s">Prev: ${nf(bCust)} orders</div></div>
    </div>`;
    h+=renderTakeawaysCard("COMP", buildCompareTakeaways(c, oa, ob, rows), "Key Takeaways & Action Plan");
    h+=`<div class="card" style="margin-top:14px">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px">
        <h3 style="margin:0">${grpTitle} Rejection Comparison (${rows.length})</h3>
        <div class="seg">
          <button class="${compGroupBy==='client'?'on':''}" onclick="compGroupBy='client';renderCompare()">By Client</button>
          <button class="${compGroupBy==='designer'?'on':''}" onclick="compGroupBy='designer';renderCompare()">By Designer</button>
          <button class="${compGroupBy==='qc'?'on':''}" onclick="compGroupBy='qc';renderCompare()">By QC</button>
        </div>
      </div>
      <table>
      <thead><tr><th>${grpTitle}</th><th class="num">A Orders</th><th class="num">A Imgs</th><th class="num">B Orders</th><th class="num">B Imgs</th><th class="num">Order Change</th></tr></thead>
      <tbody>${rows.map(r=>`<tr><td><b>${esc(r.name)}</b></td><td class="num"><b>${nf(r.a_ord)}</b></td><td class="num">${nf(r.a_imgs)}</td>
        <td class="num">${nf(r.b_ord)}</td><td class="num">${nf(r.b_imgs)}</td><td class="num">${deltaBadge(r.a_ord,r.b_ord,true,false)}</td></tr>`).join("")}</tbody>
    </table></div>`;
  }
  el.innerHTML=h;
}

// ==================== TAB 6: GROUP DATA ====================
function setGrpRange(r){
  grpRange=r;
  $("#grpDay").classList.toggle("on",!r);
  $("#grpRange").classList.toggle("on",r);
  $("#grpToWrap").classList.toggle("hide",!r);
  $("#grpFromLbl").textContent=r?"From":"Date";
  if(!r) $("#grpTo").value="";
}
function grpQuick(q){
  const lat=(ST&&ST.prod_latest)||iso(new Date()), d=new Date(lat+"T00:00:00");
  if(q==="latest"){setGrpRange(false);$("#grpFrom").value=lat;}
  else if(q==="last7"){const s=new Date(d);s.setDate(s.getDate()-6);setGrpRange(true);$("#grpFrom").value=iso(s);$("#grpTo").value=lat;}
  else{const s=new Date(d.getFullYear(),d.getMonth(),1);setGrpRange(true);$("#grpFrom").value=iso(s);$("#grpTo").value=lat;}
  runGroupData(true);
}
function resetGroupFilters(){
  $("#grpTeamSel").value="ALL";
  $("#grpDesigSel").value="ALL";
  $("#grpDesSel").value="ALL";
  $("#grpSpecSel").value="ALL";
  grpSearch="";
  runGroupData(true);
}
async function runGroupData(resetFilters=false){
  const f=$("#grpFrom").value, t=grpRange?$("#grpTo").value:"";
  if(!f) return;
  if(resetFilters){
    CUSTOM_TK.GROUP=null; EDIT_TK_MODE.GROUP=false;
    if($("#grpTeamSel")) $("#grpTeamSel").value="ALL";
    if($("#grpDesigSel")) $("#grpDesigSel").value="ALL";
    if($("#grpDesSel")) $("#grpDesSel").value="ALL";
    if($("#grpSpecSel")) $("#grpSpecSel").value="ALL";
  }
  $("#groupResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Loading Group Productivity Data…</div>';
  CUR_GROUP=await api("/api/group_data",{
    start:f, end:t||f,
    team:$("#grpTeamSel").value||"ALL",
    designation:$("#grpDesigSel").value||"ALL",
    designer:$("#grpDesSel").value||"ALL",
    spec:$("#grpSpecSel").value||"ALL",
  });
  if(CUR_GROUP && CUR_GROUP.options){
    populateSelect("#grpTeamSel", CUR_GROUP.options.teams||[], "All Teams");
    populateSelect("#grpDesigSel", CUR_GROUP.options.designations||[], "All Designations");
    populateSelect("#grpDesSel", CUR_GROUP.options.designers||[], "All Designers");
    populateSelect("#grpSpecSel", CUR_GROUP.options.specs||[], "All Specs");
  }
  renderGroupData();
}
function renderGroupData(){
  const el=$("#groupResult"), g=CUR_GROUP;
  if(!g) return;
  if(g.error){
    el.innerHTML=`<div class="card empty" style="margin-top:16px"><b>No Group Data</b><br>${esc(g.error)}</div>`;
    return;
  }
  if(g.empty_filter){
    el.innerHTML=`<div class="card empty" style="margin-top:16px"><b>No matching records</b><br>Try clearing one or more filters for ${esc(g.span_label)}.</div>`;
    return;
  }
  const k=g.kpis;
  const rows=(g.rows||[]).filter(r=>{
    if(!grpSearch) return true;
    const q=grpSearch.toLowerCase();
    return r.designer.toLowerCase().includes(q) || r.team.toLowerCase().includes(q) || r.designation.toLowerCase().includes(q);
  });

  let h=`<div class="title"><h2>Group Data — Designer Productivity</h2><span>${esc(g.span_label)}</span><div class="grow"></div>
    ${expDropdownHtml("GROUP")}</div>`;

  // Top KPI strip matching screenshot header cards (Total Files, Group Average ATI, Total Rejection, Rejection %)
  h+=`<div class="kpis5">
    <div class="card kpi" style="border-top:3px solid #f59e0b"><div class="k">Total Image / File Count</div><div class="v">${nf(k.total_files)}</div><div class="s">${nf(k.designers)} active designers</div></div>
    <div class="card kpi" style="border-top:3px solid #0ea5e9"><div class="k">Group Average ATI</div><div class="v">${esc(k.avg_ati_hms)}</div><div class="s">Active time per image</div></div>
    <div class="card kpi" style="border-top:3px solid #2563eb"><div class="k">Total Active Time</div><div class="v">${esc(k.total_active_hms)}</div><div class="s">Expected: ${esc(k.total_exp_eta_hms)}</div></div>
    <div class="card kpi" style="border-top:3px solid #b91c1c"><div class="k">Total Rejection</div><div class="v bad">${nf(k.total_rej)}</div><div class="s">QC rejected files only</div></div>
    <div class="card kpi" style="border-top:3px solid #15803d"><div class="k">Rejection %</div><div class="v ${k.rej_pct>2?'bad':'good'}">${k.rej_pct.toFixed(2)}%</div><div class="s">Overall group rejection rate</div></div>
  </div>`;

  h+=renderTakeawaysCard("GROUP", g.takeaways, "Key Takeaways & Group Summary");

  h+=`<div class="prod-layout">
    <div style="display:flex;flex-direction:column;gap:14px">
      <div class="card">
        <h3>Designation Target Time</h3>
        <table>
          <thead><tr><th>Designation</th><th class="num">Target Time</th></tr></thead>
          <tbody>${(g.desig_targets||[]).map(t=>`<tr><td><b>${esc(t.designation)}</b></td><td class="num"><b>${esc(t.target_hms)}</b></td></tr>`).join("")}</tbody>
        </table>
      </div>
      <div class="card">
        <h3>Team Breakdown</h3>
        <table>
          <thead><tr><th>Team</th><th class="num">Files</th><th class="num">Avg ATI</th><th class="num">Rej %</th></tr></thead>
          <tbody>${(g.team_summary||[]).map(tm=>`<tr>
            <td>${teamBadgeHtml(tm.team)}<div class="note">${tm.designers} designers</div></td>
            <td class="num"><b>${nf(tm.files)}</b></td>
            <td class="num">${esc(tm.avg_ati_hms)}</td>
            <td class="num ${tm.rej>0?'bad':''}">${tm.rej_pct.toFixed(2)}%</td>
          </tr>`).join("")}</tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:12px;flex-wrap:wrap">
        <h3 style="margin:0">Designer Productivity Table (${rows.length})</h3>
        <input type="text" placeholder="Search designer / team…" value="${esc(grpSearch)}" oninput="grpSearch=this.value;renderGroupData()" style="min-width:210px;padding:6px 12px;font-size:12.5px">
      </div>
      <div style="max-height:540px;overflow:auto">
        <table>
          <thead>
            <tr>
              <th>Designer Name</th>
              <th>Team</th>
              <th>Designation</th>
              <th class="num">Total Active Time</th>
              <th class="num">Expected ETA</th>
              <th class="num">Total ETA</th>
              <th class="num">File Count</th>
              <th class="num">Designer ATI</th>
              <th class="num">Reject File</th>
              <th class="num">Rejection %</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r=>`<tr>
              <td><b>${esc(r.designer)}</b></td>
              <td>${teamBadgeHtml(r.team)}</td>
              <td>${esc(r.designation)}</td>
              <td class="num"><b>${esc(r.active_hms)}</b></td>
              <td class="num"><span class="${r.active_over_target?'cell-bad':'cell-good'}">${esc(r.exp_eta_hms)}</span></td>
              <td class="num"><span class="${r.eta_over_target?'cell-bad':'cell-good'}">${esc(r.total_eta_hms)}</span></td>
              <td class="num"><b>${nf(r.files)}</b></td>
              <td class="num"><span class="${r.ati_high?'cell-bad':'cell-good'}">${esc(r.ati_hms)}</span></td>
              <td class="num ${r.rej>0?'bad':''}"><b>${nf(r.rej)}</b></td>
              <td class="num ${r.rej_pct>0?'bad':''}">${r.rej_pct.toFixed(2)}%</td>
            </tr>`).join("")}
          </tbody>
        </table>
      </div>
    </div>
  </div>`;

  el.innerHTML=h;
}

// ==================== TAB 7: SPEC-WISE DATA ====================
function setSpecRange(r){
  specRange=r;
  $("#specDay").classList.toggle("on",!r);
  $("#specRange").classList.toggle("on",r);
  $("#specToWrap").classList.toggle("hide",!r);
  $("#specFromLbl").textContent=r?"From":"Date";
  if(!r) $("#specTo").value="";
}
function specQuick(q){
  const lat=(ST&&ST.prod_latest)||iso(new Date()), d=new Date(lat+"T00:00:00");
  if(q==="latest"){setSpecRange(false);$("#specFrom").value=lat;}
  else if(q==="last7"){const s=new Date(d);s.setDate(s.getDate()-6);setSpecRange(true);$("#specFrom").value=iso(s);$("#specTo").value=lat;}
  else{const s=new Date(d.getFullYear(),d.getMonth(),1);setSpecRange(true);$("#specFrom").value=iso(s);$("#specTo").value=lat;}
  runSpecwiseData(true);
}
function resetSpecFilters(){
  $("#specIdSel").value="ALL";
  $("#specTeamSel").value="ALL";
  $("#specDesigSel").value="ALL";
  $("#specDesSel").value="ALL";
  specSearch="";
  runSpecwiseData(true);
}
function selectSpecFromSidebar(sp){
  if($("#specIdSel")){
    $("#specIdSel").value=sp;
    runSpecwiseData(false);
  }
}
async function runSpecwiseData(resetFilters=false){
  const f=$("#specFrom").value, t=specRange?$("#specTo").value:"";
  if(!f) return;
  if(resetFilters){
    CUSTOM_TK.SPEC=null; EDIT_TK_MODE.SPEC=false;
    if($("#specIdSel")) $("#specIdSel").value="ALL";
    if($("#specTeamSel")) $("#specTeamSel").value="ALL";
    if($("#specDesigSel")) $("#specDesigSel").value="ALL";
    if($("#specDesSel")) $("#specDesSel").value="ALL";
  }
  $("#specResult").innerHTML='<div class="card loading" style="margin-top:16px"><div class="spin"></div>Loading Spec-wise Productivity Data…</div>';
  CUR_SPEC=await api("/api/specwise_data",{
    start:f, end:t||f,
    spec:$("#specIdSel").value||"ALL",
    team:$("#specTeamSel").value||"ALL",
    designation:$("#specDesigSel").value||"ALL",
    designer:$("#specDesSel").value||"ALL",
  });
  if(CUR_SPEC && CUR_SPEC.options){
    populateSelect("#specIdSel", CUR_SPEC.options.specs||[], "All Specs");
    populateSelect("#specTeamSel", CUR_SPEC.options.teams||[], "All Teams");
    populateSelect("#specDesigSel", CUR_SPEC.options.designations||[], "All Designations");
    populateSelect("#specDesSel", CUR_SPEC.options.designers||[], "All Designers");
  }
  renderSpecwiseData();
}
function renderSpecwiseData(){
  const el=$("#specResult"), sp=CUR_SPEC;
  if(!sp) return;
  if(sp.error){
    el.innerHTML=`<div class="card empty" style="margin-top:16px"><b>No Spec-wise Data</b><br>${esc(sp.error)}</div>`;
    return;
  }
  if(sp.empty_filter){
    el.innerHTML=`<div class="card empty" style="margin-top:16px"><b>No matching records</b><br>Try selecting another Spec ID or clearing filters for ${esc(sp.span_label)}.</div>`;
    return;
  }
  const k=sp.kpis;
  const rows=(sp.rows||[]).filter(r=>{
    if(!specSearch) return true;
    const q=specSearch.toLowerCase();
    return r.designer.toLowerCase().includes(q) || r.team.toLowerCase().includes(q) || r.designation.toLowerCase().includes(q);
  });
  const sideList=specSideMode==="TOP" ? (sp.top_specs||[]) : (sp.low_specs||[]);

  let h=`<div class="title"><h2>Spec-wise Report — ${esc(sp.selected_spec)}</h2><span>${esc(sp.span_label)}</span><div class="grow"></div>
    ${expDropdownHtml("SPEC")}</div>`;

  // Top KPI strip matching screenshot 2 header cards (Selected Spec, Target Time, Average ATI, Total File Count, Total Reject File, Avg Efficiency)
  h+=`<div class="kpis5">
    <div class="card kpi" style="border-top:3px solid #f59e0b"><div class="k">Selected Spec ID</div><div class="v" style="font-size:20px">${esc(k.spec_name)}</div><div class="s">${nf(k.designers)} designers worked</div></div>
    <div class="card kpi" style="border-top:3px solid #0ea5e9"><div class="k">Target Time / Avg ATI</div><div class="v">${esc(k.avg_ati_hms)}</div><div class="s">Spec Target: ${esc(k.target_ati_hms)}</div></div>
    <div class="card kpi" style="border-top:3px solid #2563eb"><div class="k">Total File Count</div><div class="v">${nf(k.total_files)}</div><div class="s">Images processed on spec</div></div>
    <div class="card kpi" style="border-top:3px solid #b91c1c"><div class="k">Total Reject File</div><div class="v bad">${nf(k.total_rej)}</div><div class="s">QC rejected files only</div></div>
    <div class="card kpi" style="border-top:3px solid #15803d"><div class="k">Weighted Efficiency</div><div class="v ${k.avg_eff_pct<100?'warnc':'good'}">${k.avg_eff_pct.toFixed(2)}%</div><div class="s">SpecTime vs Actual ETA</div></div>
  </div>`;

  h+=renderTakeawaysCard("SPEC", sp.takeaways, "Key Takeaways & Spec Efficiency Summary");

  h+=`<div class="prod-layout">
    <div class="card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
        <h3 style="margin:0">${specSideMode==="TOP"?"Top Specs":"Low Specs"}</h3>
        <div class="seg">
          <button class="${specSideMode==='TOP'?'on':''}" style="padding:4px 10px;font-size:11.5px" onclick="specSideMode='TOP';renderSpecwiseData()">Top</button>
          <button class="${specSideMode==='LOW'?'on':''}" style="padding:4px 10px;font-size:11.5px" onclick="specSideMode='LOW';renderSpecwiseData()">Low</button>
        </div>
      </div>
      <div class="note" style="margin-bottom:8px">Click any Spec ID to filter table</div>
      <div style="max-height:500px;overflow:auto">
        <table>
          <thead><tr><th>Spec ID</th><th class="num">ATI</th><th class="num">Images</th></tr></thead>
          <tbody>
            ${sideList.map(item=>`<tr style="cursor:pointer" onclick="selectSpecFromSidebar('${esc(item.spec)}')">
              <td><b style="color:var(--brand)">${esc(item.spec)}</b></td>
              <td class="num">${esc(item.ati_hms)}</td>
              <td class="num"><b>${nf(item.files)}</b></td>
            </tr>`).join("")}
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:12px;flex-wrap:wrap">
        <h3 style="margin:0">Spec-wise Designer Breakdown (${rows.length})</h3>
        <input type="text" placeholder="Search designer / team…" value="${esc(specSearch)}" oninput="specSearch=this.value;renderSpecwiseData()" style="min-width:210px;padding:6px 12px;font-size:12.5px">
      </div>
      <div style="max-height:540px;overflow:auto">
        <table>
          <thead>
            <tr>
              <th>Designer Name</th>
              <th>Team</th>
              <th>Designation</th>
              <th class="num">File Count</th>
              <th class="num">Designer ATI</th>
              <th class="num">Reject File</th>
              <th class="num">Efficiency</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r=>`<tr>
              <td><b>${esc(r.designer)}</b></td>
              <td>${teamBadgeHtml(r.team)}</td>
              <td>${esc(r.designation)}</td>
              <td class="num"><b>${nf(r.files)}</b></td>
              <td class="num"><span class="${r.high_ati?'cell-bad':'cell-good'}">${esc(r.ati_hms)}</span></td>
              <td class="num ${r.rej>0?'bad':''}"><b>${nf(r.rej)}</b></td>
              <td class="num"><span class="${r.low_eff?'cell-warn':'cell-good'}">${r.eff_pct.toFixed(2)}%</span></td>
            </tr>`).join("")}
          </tbody>
        </table>
      </div>
    </div>
  </div>`;

  el.innerHTML=h;
}

// ==================== UNIFIED EXPORT PAYLOADS ====================
function getTabExportPayload(tabKey){
  if(tabKey==="QA" && CUR_QA && !CUR_QA.error){
    const s=CUR_QA.summary;
    const orders=getFilteredQaOrders();
    const imgs=orders.reduce((a,o)=>a+o.volume,0), errs=orders.reduce((a,o)=>a+o.errors,0);
    const clean=orders.filter(o=>o.errors===0).length, rate=imgs?(errs/imgs*100):0;
    const pCr=s.paired_cr||{orders:0,images:0,prod:0,cust:0};
    const sumTxt = CUSTOM_SUM_TXT.qaSumBox !== null ? CUSTOM_SUM_TXT.qaSumBox : CUR_QA.text;
    return {
      file_base: "QA_Audit_Report",
      badge_text: "INTERNAL QA AUDIT",
      title: `Quality Audit Report — ${$("#qaAud").value.toUpperCase()}`,
      subtitle: `${CUR_QA.label}  |  Generated ${new Date().toLocaleDateString()}`,
      kpi_headers: ["Images Audited", "Errors Found", "Defect Rate", "Clean Orders"],
      kpi_row: [nf(imgs), nf(errs), `${rate.toFixed(2)}%`, nf(clean)],
      kpi_subs: [`${nf(orders.length)} orders`, `${nf(orders.length-clean)} with issues`, `Target ≤ ${s.escalate_above}%`, `${orders.length?Math.round(clean/orders.length*100):0}% clean`],
      combined_title: "EXTERNAL CLIENT REJECTIONS IN SAME PERIOD (CROSS-CHECK)",
      combined_strip: [
        {label:"Client Rejected Orders", value:`${nf(pCr.orders)} orders`, bad:pCr.orders>0},
        {label:"Client Rejected Volume", value:`${nf(pCr.images)} images`, bad:false},
        {label:"CR Production Faults", value:`${nf(pCr.prod)} orders`, bad:pCr.prod>0},
        {label:"CR Customer Faults", value:`${nf(pCr.cust)} orders`, bad:false},
      ],
      streak_alerts: s.streak_alerts || [],
      takeaways: getEffectiveTakeaways("QA", s.takeaways),
      summary_text: sumTxt,
      table_title: `Job Audit Details (${orders.length})`,
      table_headers: ["Client", "Order ID", "Images", "Errors", "Category", "Designer", "QC"],
      table_rows: orders.map(o=>[o.client, o.order_id, nf(o.volume), nf(o.errors), o.categories, o.designer, o.qc]),
      right_cols: [2, 3],
    };
  }
  if(tabKey==="CR" && CUR_CR && !CUR_CR.error){
    const s=CUR_CR.summary;
    const orders=getFilteredCrOrders();
    const imgs=orders.reduce((a,o)=>a+o.volume,0);
    const prod=orders.filter(o=>!o.is_customer).length, cust=orders.filter(o=>o.is_customer).length;
    const delayed=orders.filter(o=>o.delay).length;
    const pQa=s.paired_qa||{orders:0,images:0,errors:0,rate:0};
    const sumTxt = CUSTOM_SUM_TXT.crSumBox !== null ? CUSTOM_SUM_TXT.crSumBox : CUR_CR.text;
    return {
      file_base: "Client_Rejection_Report",
      badge_text: "CLIENT REJECTION REPORT",
      title: "External Client Rejection Report",
      subtitle: `${CUR_CR.label}  |  Generated ${new Date().toLocaleDateString()}`,
      kpi_headers: ["Rejected Orders", "Rejected Images", "Production Faults", "Customer Faults"],
      kpi_row: [nf(orders.length), nf(imgs), nf(prod), nf(cust)],
      kpi_subs: [`${nf(delayed)} delayed`, "Total rejected volume", "Internal fault orders", "Client fault orders"],
      combined_title: "INTERNAL QA AUDITS IN SAME PERIOD (CROSS-CHECK)",
      combined_strip: [
        {label:"QA Images Audited", value:`${nf(pQa.images)} images`, bad:false},
        {label:"QA Errors Found", value:`${nf(pQa.errors)} errors`, bad:pQa.errors>0},
        {label:"QA Defect Rate", value:`${Number(pQa.rate||0).toFixed(2)}%`, bad:(pQa.rate||0)>10},
        {label:"QA Orders Audited", value:`${nf(pQa.orders)} orders`, bad:false},
      ],
      streak_alerts: s.streak_alerts || [],
      takeaways: getEffectiveTakeaways("CR", buildCrTakeaways(orders, s.streak_alerts, s.streak_days||3)),
      summary_text: sumTxt,
      table_title: `Rejected Job Details (${orders.length})`,
      table_headers: ["Client", "Job ID", "Images", "Fault Type", "Reason", "Designer", "QC Name"],
      table_rows: orders.map(o=>[o.client, o.job_id, nf(o.volume), o.is_customer?"Customer":"Production", o.reason, o.designer, o.qc]),
      right_cols: [2],
    };
  }
  if(tabKey==="ATTN" && CUR_ATTN){
    const win=CUR_ATTN.windows[attnWin] || CUR_ATTN.windows.DAILY;
    let list=win.clients||[];
    if(attnFilter==="CRITICAL") list=list.filter(c=>c.status==="CRITICAL");
    else if(attnFilter==="WATCH") list=list.filter(c=>c.status==="WATCH");
    else if(attnFilter==="HEALTHY") list=list.filter(c=>c.status==="HEALTHY");
    if($("#attnClientSel") && $("#attnClientSel").value!=="ALL") list=list.filter(c=>c.client===$("#attnClientSel").value);
    return {
      file_base: `Client_Attention_${attnWin}`,
      badge_text: "CLIENT ATTENTION RADAR",
      title: `Client Attention Radar — ${attnWin}`,
      subtitle: `${win.label}  |  Generated ${new Date().toLocaleDateString()}`,
      kpi_headers: ["Immediate Attention", "Watchlist Clients", "Healthy Clients", "Total Evaluated"],
      kpi_row: [nf(win.critical_count), nf(win.watch_count), nf(win.healthy_count), nf((win.clients||[]).length)],
      kpi_subs: ["Critical priority", "Monitor closely", "Within thresholds", "Combined QA + CR"],
      takeaways: getEffectiveTakeaways("ATTN", buildAttnTakeaways(list)),
      attn_cards: list,
      summary_text: "",
      table_title: `Client Risk Breakdown (${list.length})`,
      table_headers: ["Client", "Status", "Score", "QA Rate", "QA Prior", "CR Orders", "CR Prior", "Key Drivers"],
      table_rows: list.map(c=>[c.client, c.status, c.score, `${c.qa_a_rate.toFixed(2)}%`, `${c.qa_b_rate.toFixed(2)}%`, nf(c.cr_a_ord), nf(c.cr_b_ord), c.flags.join("; ")]),
      right_cols: [2, 3, 4, 5, 6],
    };
  }
  if(tabKey==="CHARTS" && CUR_CHARTS){
    const k=CUR_CHARTS.kpis||{};
    return {
      file_base: "Quality_Analytics_Summary",
      badge_text: "ANALYTICS & TRENDS",
      title: "Quality & Rejection Analytics Summary",
      subtitle: `${CUR_CHARTS.span_label}  |  Generated ${new Date().toLocaleDateString()}`,
      kpi_headers: ["QA Images", "QA Errors", "QA Defect Rate", "CR Orders"],
      kpi_row: [nf(k.qa_imgs||0), nf(k.qa_err||0), `${(k.qa_rate||0).toFixed(2)}%`, nf(k.cr_ord||0)],
      kpi_subs: ["Audited volume", "Internal defects", "Combined rate", `${nf(k.cr_imgs||0)} rejected imgs`],
      takeaways: getEffectiveTakeaways("CHARTS", buildChartsTakeaways(CUR_CHARTS)),
      summary_text: "",
      table_title: "Client Defect & Rejection Impact",
      table_headers: ["Client", "QA Images", "QA Errors", "QA Rate", "CR Orders", "CR Images"],
      table_rows: (CUR_CHARTS.clients||[]).map(c=>[c.client, nf(c.qa_imgs), nf(c.qa_err), `${c.qa_rate.toFixed(2)}%`, nf(c.cr_ord), nf(c.cr_imgs)]),
      right_cols: [1, 2, 3, 4, 5],
    };
  }
  if(tabKey==="COMP" && CUR_COMP && !CUR_COMP.error){
    const oa=filterCompOrders(CUR_COMP.orders_a), ob=filterCompOrders(CUR_COMP.orders_b);
    const rows=aggCompRows(oa, ob, CUR_COMP.mode, compGroupBy);
    const grpTitle=compGroupBy==="client"?"Client":(compGroupBy==="designer"?"Designer":"QC Name");
    const pCrA=CUR_COMP.paired_cr_a, pCrB=CUR_COMP.paired_cr_b;
    const pQaA=CUR_COMP.paired_qa_a, pQaB=CUR_COMP.paired_qa_b;
    if(CUR_COMP.mode==="QA"){
      const aImgs=oa.reduce((s,o)=>s+o.volume,0), aErr=oa.reduce((s,o)=>s+o.errors,0);
      const bImgs=ob.reduce((s,o)=>s+o.volume,0), bErr=ob.reduce((s,o)=>s+o.errors,0);
      const aRate=aImgs?(aErr/aImgs*100):0, bRate=bImgs?(bErr/bImgs*100):0;
      return {
        file_base: "QA_Period_Comparison",
        badge_text: "QA PERIOD COMPARE",
        title: "QA Audit Period Comparison",
        subtitle: `Period A: ${CUR_COMP.a_label}  vs  Period B: ${CUR_COMP.b_label}`,
        kpi_headers: ["Period A Rate", "Period B Rate", "Period A Errors", "Period A Images"],
        kpi_row: [`${aRate.toFixed(2)}%`, `${bRate.toFixed(2)}%`, nf(aErr), nf(aImgs)],
        kpi_subs: [`Delta: ${(aRate-bRate)>=0?'+':''}${(aRate-bRate).toFixed(2)}%`, `Prev: ${nf(bErr)} err`, `Prev: ${nf(bErr)} err`, `Prev: ${nf(bImgs)} imgs`],
        combined_title: "CLIENT REJECTIONS COMPARISON (SAME PERIODS A vs B)",
        combined_strip: pCrA ? [
          {label:"Period A Client Rejections", value:`${nf(pCrA.orders)} ord (${nf(pCrA.images)} imgs)`, bad:pCrA.orders>0},
          {label:"Period B Client Rejections", value:`${nf(pCrB.orders)} ord (${nf(pCrB.images)} imgs)`, bad:false},
          {label:"CR Order Delta (A vs B)", value:`${pCrA.orders-pCrB.orders>=0?'+':''}${nf(pCrA.orders-pCrB.orders)} orders`, bad:(pCrA.orders-pCrB.orders)>0},
          {label:"Period A Prod / Cust Faults", value:`${pCrA.prod} Prod / ${pCrA.cust} Cust`, bad:pCrA.prod>0},
        ] : [],
        takeaways: getEffectiveTakeaways("COMP", buildCompareTakeaways(CUR_COMP, oa, ob, rows)),
        summary_text: "",
        table_title: `${grpTitle} Comparison (${rows.length})`,
        table_headers: [grpTitle, "A Imgs", "A Err", "A Rate", "B Rate", "Change"],
        table_rows: rows.map(r=>[r.name, nf(r.a_imgs), nf(r.a_err), `${r.a_rate.toFixed(2)}%`, `${r.b_rate.toFixed(2)}%`, `${r.rate_diff>=0?'+':''}${r.rate_diff.toFixed(2)}%`]),
        right_cols: [1, 2, 3, 4, 5],
      };
    } else {
      const aImgs=oa.reduce((s,o)=>s+o.volume,0), bImgs=ob.reduce((s,o)=>s+o.volume,0);
      return {
        file_base: "CR_Period_Comparison",
        badge_text: "CR PERIOD COMPARE",
        title: "Client Rejection Period Comparison",
        subtitle: `Period A: ${CUR_COMP.a_label}  vs  Period B: ${CUR_COMP.b_label}`,
        kpi_headers: ["Period A Orders", "Period B Orders", "Period A Images", "Period B Images"],
        kpi_row: [nf(oa.length), nf(ob.length), nf(aImgs), nf(bImgs)],
        kpi_subs: [`Delta: ${oa.length-ob.length>=0?'+':''}${nf(oa.length-ob.length)}`, "Previous period", `Delta: ${aImgs-bImgs>=0?'+':''}${nf(aImgs-bImgs)}`, "Previous volume"],
        combined_title: "INTERNAL QA AUDITS COMPARISON (SAME PERIODS A vs B)",
        combined_strip: pQaA ? [
          {label:"Period A QA Defect Rate", value:`${Number(pQaA.rate||0).toFixed(2)}% (${nf(pQaA.errors)} err)`, bad:(pQaA.rate||0)>10},
          {label:"Period B QA Defect Rate", value:`${Number(pQaB.rate||0).toFixed(2)}% (${nf(pQaB.errors)} err)`, bad:false},
          {label:"Period A QA Images", value:`${nf(pQaA.images)} audited`, bad:false},
          {label:"Period B QA Images", value:`${nf(pQaB.images)} audited`, bad:false},
        ] : [],
        takeaways: getEffectiveTakeaways("COMP", buildCompareTakeaways(CUR_COMP, oa, ob, rows)),
        summary_text: "",
        table_title: `${grpTitle} Rejection Comparison (${rows.length})`,
        table_headers: [grpTitle, "A Orders", "A Imgs", "B Orders", "B Imgs", "Order Change"],
        table_rows: rows.map(r=>[r.name, nf(r.a_ord), nf(r.a_imgs), nf(r.b_ord), nf(r.b_imgs), `${r.ord_diff>=0?'+':''}${nf(r.ord_diff)}`]),
        right_cols: [1, 2, 3, 4, 5],
      };
    }
  }
  if(tabKey==="GROUP" && CUR_GROUP && !CUR_GROUP.error && !CUR_GROUP.empty_filter){
    const k=CUR_GROUP.kpis;
    return {
      file_base: "Group_Productivity_Data",
      badge_text: "GROUP PRODUCTIVITY",
      title: "Group Data — Designer Productivity",
      subtitle: `${CUR_GROUP.span_label}  |  Generated ${new Date().toLocaleDateString()}`,
      kpi_headers: ["Total Files", "Group Avg ATI", "Total Rejections", "Rejection %"],
      kpi_row: [nf(k.total_files), k.avg_ati_hms, nf(k.total_rej), `${k.rej_pct.toFixed(2)}%`],
      kpi_subs: [`${nf(k.designers)} designers`, `Active: ${k.total_active_hms}`, `Expected: ${k.total_exp_eta_hms}`, `Total ETA: ${k.total_eta_hms}`],
      takeaways: getEffectiveTakeaways("GROUP", CUR_GROUP.takeaways),
      summary_text: "",
      table_title: `Designer Productivity (${(CUR_GROUP.rows||[]).length})`,
      table_headers: ["Designer Name", "Team", "Designation", "Active Time", "Expected ETA", "Total ETA", "File Count", "Designer ATI", "Reject File", "Rejection %"],
      table_rows: (CUR_GROUP.rows||[]).map(r=>[r.designer, r.team, r.designation, r.active_hms, r.exp_eta_hms, r.total_eta_hms, nf(r.files), r.ati_hms, nf(r.rej), `${r.rej_pct.toFixed(2)}%`]),
      right_cols: [3, 4, 5, 6, 7, 8, 9],
    };
  }
  if(tabKey==="SPEC" && CUR_SPEC && !CUR_SPEC.error && !CUR_SPEC.empty_filter){
    const k=CUR_SPEC.kpis;
    return {
      file_base: "Specwise_Productivity_Data",
      badge_text: "SPEC-WISE REPORT",
      title: `Spec-wise Report — ${CUR_SPEC.selected_spec}`,
      subtitle: `${CUR_SPEC.span_label}  |  Target Time: ${k.target_ati_hms}  |  Generated ${new Date().toLocaleDateString()}`,
      kpi_headers: ["Total Files", "Average ATI", "Total Reject File", "Efficiency %"],
      kpi_row: [nf(k.total_files), k.avg_ati_hms, nf(k.total_rej), `${k.avg_eff_pct.toFixed(2)}%`],
      kpi_subs: [`Spec: ${k.spec_name}`, `Target: ${k.target_ati_hms}`, `${nf(k.designers)} designers`, "Weighted efficiency"],
      takeaways: getEffectiveTakeaways("SPEC", CUR_SPEC.takeaways),
      summary_text: "",
      table_title: `Spec-wise Designer Breakdown (${(CUR_SPEC.rows||[]).length})`,
      table_headers: ["Designer Name", "Team", "Designation", "File Count", "Designer ATI", "Reject File", "Efficiency"],
      table_rows: (CUR_SPEC.rows||[]).map(r=>[r.designer, r.team, r.designation, nf(r.files), r.ati_hms, nf(r.rej), `${r.eff_pct.toFixed(2)}%`]),
      right_cols: [3, 4, 5, 6],
    };
  }
  return null;
}

function copyText(text){
  if(navigator.clipboard) navigator.clipboard.writeText(text).then(()=>toast("Copied to clipboard"));
}

// ==================== GLOBAL EVENTS ====================
$("#qaAud").onchange=()=>{
  const m=(ST.auditors||[]).find(x=>x.name===$("#qaAud").value);
  if(m&&m.latest&&!qaRange) $("#qaFrom").value=m.latest;
  runQaReport();
};
$("#reload").onclick=async()=>{
  toast("Reloading all sheets…");
  CUR_QA=CUR_CR=CUR_ATTN=CUR_CHARTS=CUR_COMP=CUR_GROUP=CUR_SPEC=null;
  await api("/api/reload",{});
  setTimeout(init,800);
};
$("#quit").onclick=async()=>{
  await api("/api/quit",{});
  document.body.innerHTML='<div style="display:grid;place-items:center;height:100vh;font:16px system-ui;color:#64748b">Quality &amp; Client Analytics Suite closed. You may close this tab.</div>';
};

setInterval(()=>fetch("/ping",{headers:{"X-Token":T}}).catch(()=>{}), 4000);
// Heartbeat watchdog handles shutdown cleanly without killing server on page refresh
init();
</script>
</body></html>
"""

if __name__ == "__main__":
    serve()
