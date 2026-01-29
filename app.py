import json
import requests
import pandas as pd
import streamlit as st
from io import BytesIO
from msal import ConfidentialClientApplication
from datetime import datetime
import altair as alt
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF

st.set_page_config(page_title="Resource Dashboard", layout="wide")
st.markdown(
    """
<style>
/* Full width layout */
.block-container {
    max-width: 100% !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
}
/* Dashboard Header */
.dashboard-header {
    background: linear-gradient(90deg, #0b1d3a, #1d3557);
    padding: 18px 26px;
    border-radius: 16px;
    display: grid;
    grid-template-columns: 1fr auto 1fr;
    align-items: center;
}
/* Title */
.dashboard-title {
    grid-column: 2;
    text-align: black;
    font-size: 32px;
    font-weight: 900;
    color: #f8f9fa;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.6);
    letter-spacing: 0.4px;
}
/* Refresh button wrapper */
.refresh-wrapper {
    grid-column: 3;
    display: flex;
    justify-content: flex-end;
    align-items : center;
}
/* Refresh button */
div.stButton > button {
    background: linear-gradient(90deg,#43aa8b,#4d908e);
    color: white;
    font-weight: 700;
    border-radius: 10px;
    padding: 0.45rem 1.2rem;
    border: none;
    /* Make it taller and bigger */
    font-size: 28px; /* close to title size */
    padding: 0.8rem 1.6rem; /* more vertical padding */
    height: 80px; /* force height if needed */
    line-height: 1.2;
}
div.stButton > button:hover {
    background: linear-gradient(90deg,#2d6a4f,#40916c);
    transform: scale(1.05);
}
/* Tabs styling (unchanged) */
div[role="tablist"] { gap: 40px !important; }
div[role="tablist"] button[role="tab"] {
    font-weight: 900 !important;
    font-size: 42px !important;
    color: #000 !important;
}
div[role="tablist"] button[role="tab"][aria-selected="true"] {
    border-bottom: 6px solid #e63946 !important;
}
</style>
""",
    unsafe_allow_html=True,
)
# # ================= HEADER =================
# st.markdown("""
# <div class="dashboard-header">
#     <div></div>
#     <div class="dashboard-title">📊 Resource Dashboard</div>
#     <div class="refresh-wrapper">
# """, unsafe_allow_html=True)
# ======================================================
# LOAD CREDENTIALS
# ======================================================
with open("cred.json", "r") as f:
    creds = json.load(f)
TENANT_ID = creds["TENANT_ID"]
CLIENT_ID = creds["CLIENT_ID"]
CLIENT_SECRET = creds["CLIENT_SECRET"]
# ======================================================
# SHAREPOINT CONFIG
# ======================================================
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]
SHAREPOINT_HOSTNAME = "cognizantonline.sharepoint.com"
SITE_PATH = "/sites/ResourceDashboard"
FILE_PATH = "Tracker.xlsx"


# ======================================================
# AUTH FUNCTIONS
# ======================================================
def get_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    return app.acquire_token_for_client(scopes=SCOPES)["access_token"]


def get_site_id(token):
    url = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    return r.json()["id"]


def get_excel_bytes():
    token = get_token()
    site_id = get_site_id(token)
    drives = requests.get(
        f"{GRAPH_BASE}/sites/{site_id}/drives",
        headers={"Authorization": f"Bearer {token}"},
    ).json()["value"]
    drive_id = drives[0]["id"]
    r = requests.get(
        f"{GRAPH_BASE}/drives/{drive_id}/root:/{FILE_PATH}:/content",
        headers={"Authorization": f"Bearer {token}"},
        allow_redirects=False,
    )
    return requests.get(r.headers["Location"]).content


# ======================================================
# LOAD EXCEL (CACHED) with manual refresh
# ======================================================
@st.cache_data(ttl=300)  # auto refresh every 5 mins
def load_data():
    xls = pd.ExcelFile(BytesIO(get_excel_bytes()))
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    return df, datetime.now()


# ================= HEADER =================
header_left, header_right = st.columns([12, 1])  # title wide, button narrow
with header_left:
    st.markdown(
        """
        <div class="dashboard-header">
        <div class='dashboard-title'> Resource Dashboard</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with header_right:
    if st.button("🔄", key="refresh_button"):  # icon-only keeps it compact
        st.cache_data.clear()
# Load data (fresh if cache was cleared)
df, refreshed_time = load_data()
if df.empty:
    st.stop()
st.caption(f"🕒 Last refreshed at: {refreshed_time.strftime('%d-%b-%Y %I:%M:%S %p')}")
# ======================================================
# PRESERVE ORIGINAL DATA (for display and saving)
# ======================================================
# Store original data with original case and column names
df_original = df.copy()
original_column_names = df.columns.tolist()  # Store original column names


# ======================================================
# NORMALIZE COLUMN NAMES (for internal processing only)
# ======================================================
def clean_col(c):
    return str(c).lower().strip().replace(" ", "_").replace("/", "")


# Create mapping from cleaned to original names
column_name_mapping = {clean_col(c): c for c in original_column_names}
# Store in session state for persistence across reruns
if "column_name_mapping" not in st.session_state:
    st.session_state["column_name_mapping"] = column_name_mapping
if "original_column_names" not in st.session_state:
    st.session_state["original_column_names"] = original_column_names
# Use session state version
column_name_mapping = st.session_state["column_name_mapping"]
original_column_names = st.session_state["original_column_names"]
df.columns = [clean_col(c) for c in df.columns]
df_original.columns = [clean_col(c) for c in df_original.columns]


# ======================================================
# NORMALIZE STRING VALUES FOR COMPARISON ONLY
# - Create normalized copy for filtering/comparison
# - Keep original for display and saving
# ======================================================
def normalize_for_comparison(series):
    """Normalize series for case-insensitive comparison"""
    if series.dtype == "object":
        s = series.astype(str).str.strip()
        s = s.replace({"nan": "", "none": "", "NAN": "", "NONE": ""})
        return s.str.upper()
    return series


# Create normalized version for filtering/comparison
df_normalized = df.copy()
for _col in df_normalized.columns:
    if df_normalized[_col].dtype == "object":
        df_normalized[_col] = normalize_for_comparison(df_normalized[_col])


# ======================================================
# COLUMN MAPPING (robust lookup)
# ======================================================
def find_col(*variants):
    variants_norm = [
        v.lower().replace(" ", "_").replace("/", "").replace("-", "") for v in variants
    ]
    for v in variants_norm:
        if v in df.columns:
            return v
    # try substring match
    for col in df.columns:
        for v in variants_norm:
            if v in col:
                return col
    return None


def ensure_col(primary_name, *variants):
    col = find_col(*variants)
    if col is None:
        fake = f"__missing_{primary_name}__"
        df[fake] = ""
        st.warning(
            f"Expected column for `{primary_name}` not found. Created empty column `{fake}` to avoid errors."
        )
        return fake
    return col


# Keep SO_COL for SO type
SO_COL = ensure_col("type", "type", "so_type", "so type", "so#")
# Add a dedicated Priority column mapping
PRIORITY_COL = ensure_col(
    "priority",
    "priority",
    "priorty",
    "priority(crotical/non critical)",
    "priority(crotical/noncritical)",
)
REQ_COL = ensure_col("reqtype", "reqtype", "req_type", "req type", "request_type")
STATUS_COL = ensure_col("sourcingstatus", "sourcingstatus", "status", "sourcing status")
MFR_COL = ensure_col("mfr", "mfr")
ESC_COL = ensure_col(
    "escalated_need_attention",
    "escalated/need attention",
    "escalatedneed_attention",
    "escalated_need_attention",
    "escalated",
    "need_attentionescalated",
    "need_attention",
)
# Add Loss column mapping
LOSS_COL = ensure_col(
    "loss", "loss", "loss_y_n", "loss (y/n)", "rev_loss", "revenue_loss"
)
# Add AIA / NON AIA column mapping
AIA_COL = ensure_col(
    "ia_non_ai", "ia/non ai", "ia_non_ai", "ai/non ai", "aia/non aia", "aia", "non ai"
)
# Demand-type column (Hiring / Internal / Proactive) — detect variants from noisy header names
DEMAND_COL = ensure_col(
    "demand_flag",
    "hiring",
    "demand",
    "hiring/can be",
    "hiring/can be fulfill",
    "hiring/can be fulfillled internally/procative sos flagged are valud demands",
    "hiring/can be fulfillled internally",
    "hiring/can be fulfilled internally",
)
# Needed by column (month names)
NEEDED_BY_COL = ensure_col(
    "needed_by", "needed by", "neededby", "needed-by", "needed_by"
)


# ======================================================
# FUNCTION TO FORMAT LARGE NUMBERS
# ======================================================
def format_large_number(num):
    """Format large numbers for better readability (e.g., 2000 -> 2,000 or 2K)"""
    if pd.isna(num):
        return num
    try:
        # Convert to float first to handle string numbers
        num_float = float(num)
        # If it's an integer, format with commas
        if num_float == int(num_float):
            return f"{int(num_float):,}"
        else:
            # For decimals, use 2 decimal places with commas
            return f"{num_float:,.2f}"
    except (ValueError, TypeError):
        return num


# ======================================================
# FUNCTION TO PARSE MONTH AND COUNT NEEDED BY
# ======================================================
def parse_month_name(month_str):
    """Parse month name to month number (1-12). Returns None if invalid."""
    if pd.isna(month_str) or month_str == "":
        return None
    month_str = str(month_str).strip().upper()
    # Common month name patterns
    month_map = {
        "JAN": 1,
        "JANUARY": 1,
        "FEB": 2,
        "FEBRUARY": 2,
        "MAR": 3,
        "MARCH": 3,
        "APR": 4,
        "APRIL": 4,
        "MAY": 5,
        "JUN": 6,
        "JUNE": 6,
        "JUL": 7,
        "JULY": 7,
        "AUG": 8,
        "AUGUST": 8,
        "SEP": 9,
        "SEPT": 9,
        "SEPTEMBER": 9,
        "OCT": 10,
        "OCTOBER": 10,
        "NOV": 11,
        "NOVEMBER": 11,
        "DEC": 12,
        "DECEMBER": 12,
    }
    # Try direct match
    if month_str in month_map:
        return month_map[month_str]
    # Try substring match (e.g., "Jan" in "January")
    for key, val in month_map.items():
        if key in month_str or month_str in key:
            return val
    # Try numeric (1-12)
    try:
        month_num = int(month_str)
        if 1 <= month_num <= 12:
            return month_num
    except:
        pass
    return None


def count_needed_by(df_normalized, status_col, status_val):
    """Count items where needed_by month is <= current month for given status."""
    if NEEDED_BY_COL not in df_normalized.columns:
        return 0
    current_month = datetime.now().month
    status_mask = df_normalized[status_col] == status_val
    needed_by_series = df_normalized.loc[status_mask, NEEDED_BY_COL]
    count = 0
    for month_val in needed_by_series:
        month_num = parse_month_name(month_val)
        if month_num is not None and month_num <= current_month:
            count += 1
    return count


# ======================================================
# NORMALIZE STATUS FOR FILTERING (use normalized version)
# ======================================================
def normalize_status(val: str) -> str:
    val = str(val).upper().strip()
    if "OPEN" in val:
        return "OPEN"
    if "CLOSE" in val:
        return "CLOSED"
    if "CANCEL" in val:
        return "CANCEL"
    return val  # fallback to original


# Apply status normalization to normalized dataframe ONLY (preserve original values)
if STATUS_COL in df_normalized.columns:
    df_normalized[STATUS_COL] = df_normalized[STATUS_COL].apply(normalize_status)


    # DO NOT modify df_original - keep original values as-is
# ======================================================
# SUMMARY TABLE FUNCTION
# ======================================================
# Will be calculated from master_df
def build_summary(data, status_label, total_count=None):
    """Build a one-row summary table using tolerant substring matching so small spelling/format variants don't break counts.
    Args:
        data: Filtered dataframe for the specific status
        status_label: Label for the status (e.g., "Open", "Closed")
        total_count: Total count across all statuses (defaults to len of master_df from session state)
    """

    def _contains(series, patterns):
        # returns boolean Series (same length as series) checking if any pattern exists (case-insensitive)
        if series.name not in data.columns:
            return pd.Series([False] * len(series))
        s = series.fillna("").astype(str).str.upper()
        pat = "|".join(patterns)
        return s.str.contains(pat, na=False)

    # Use provided total_count or get from session state
    if total_count is None:
        total_count = len(st.session_state.get("master_df", data))
    total = total_count
    status_count = len(data)
    # Prefer the demand-type column (if present) for Hiring/Internal/Proactive logic, otherwise fall back to SO_COL
    demand_col = (
        DEMAND_COL
        if "DEMAND_COL" in globals() and DEMAND_COL in data.columns
        else SO_COL
    )
    # 'Can be internally fulfilled' may appear as 'Can be fulfillled internally' etc. Match on 'FULFIL' substring
    cbif_mask = (
        _contains(data[demand_col], ["FULFIL"])
        if demand_col in data.columns
        else pd.Series([False] * len(data))
    )
    new_mask = (
        _contains(data[REQ_COL], ["NEW"])
        if REQ_COL in data.columns
        else pd.Series([False] * len(data))
    )
    repl_mask = (
        _contains(data[REQ_COL], ["REPLAC"])
        if REQ_COL in data.columns
        else pd.Series([False] * len(data))
    )
    can_be_new = int((cbif_mask & new_mask).sum())
    can_be_repl = int((cbif_mask & repl_mask).sum())
    hiring = (
        int(_contains(data[demand_col], ["HIRING"]).sum())
        if demand_col in data.columns
        else 0
    )
    proactive = (
        int(_contains(data[demand_col], ["PROACT"]).sum())
        if demand_col in data.columns
        else 0
    )
    summary_df = pd.DataFrame(
        [[total, status_count, can_be_new, can_be_repl, hiring, proactive]],
        columns=pd.MultiIndex.from_tuples(
            [
                ("Total SO", ""),  # no None, just empty string
                (status_label, ""),  # same here
                ("Can Be Internally Fulfilled", "New"),
                ("Can Be Internally Fulfilled", "Replacement"),
                ("Hiring", ""),
                ("Proactive", ""),
            ]
        ),
    )
    return (
        summary_df.style.hide(axis="index")
        .set_properties(**{"text-align": "center", "color": "#000000"})
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("text-align", "center"),
                        ("color", "#000000"),
                        ("font-weight", "700"),
                    ],
                },
                {
                    "selector": "td",
                    "props": [("text-align", "center"), ("color", "#000000")],
                },
            ]
        )
    )


# ======================================================
# DISPLAY: Tabs (Summary/Details + Visuals)
# ======================================================
st.markdown(
    """
<style>
/* Tabs container box */
div[data-baseweb="tab-list"] {
    background: linear-gradient(90deg, #edf2fb, #e9ecef);
    padding: 12px 18px;
    border-radius: 14px;
    width: 100%;
    justify-content: center;
    margin-bottom: 24px;
    box-shadow: 0 6px 14px rgba(0,0,0,0.08);
}
/* Individual tab */
button[data-baseweb="tab"] {
    font-size: 22px !important;
    font-weight: 800 !important;
    padding: 10px 28px !important;
    margin: 0 16px !important;
    border-radius: 12px !important;
    background: transparent !important;
    color: #343a40 !important;
}
/* Selected tab */
button[data-baseweb="tab"][aria-selected="true"] {
    background: linear-gradient(90deg, #4361ee, #4895ef) !important;
    color: white !important;
    box-shadow: 0 6px 14px rgba(67,97,238,0.35);
}
/* Remove default underline */
button[data-baseweb="tab"]::after {
    display: none !important;
}
</style>
""",
    unsafe_allow_html=True,
)


def save_df_to_sharepoint(df):
    token = get_token()
    site_id = get_site_id(token)
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    # Build Excel in memory
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    # Upload (overwrite)
    upload_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{FILE_PATH}:/content"
    resp = requests.put(upload_url, headers=headers, data=buffer)
    if resp.status_code not in [200, 201]:
        raise Exception(resp.text)


# ==================================================
# Initialize master dataframe (preserve original case)
# ==================================================
if "master_df" not in st.session_state:
    st.session_state["master_df"] = df_original.copy()
    st.session_state["master_df_normalized"] = df_normalized.copy()
master_df = st.session_state["master_df"]
master_df_normalized = st.session_state["master_df_normalized"]


# ==================================================
# Helper function to get normalized version for comparison
# ==================================================
def get_normalized_df(df):
    """Get normalized version of dataframe for case-insensitive comparison"""
    df_norm = df.copy()
    for col in df_norm.columns:
        if df_norm[col].dtype == "object":
            df_norm[col] = normalize_for_comparison(df_norm[col])
    # Normalize status
    if STATUS_COL in df_norm.columns:
        df_norm[STATUS_COL] = df_norm[STATUS_COL].apply(normalize_status)
    return df_norm


# ==================================================
# Derive filtered subsets for tab1 & tab2
# ==================================================
def get_status_dfs(df):
    """Get filtered dataframes by status using case-insensitive comparison"""
    df_norm = get_normalized_df(df)
    open_df = df[df_norm[STATUS_COL] == "OPEN"].copy()
    closed_df = df[df_norm[STATUS_COL] == "CLOSED"].copy()
    cancel_df = df[df_norm[STATUS_COL] == "CANCEL"].copy()
    return open_df, closed_df, cancel_df


open_df, closed_df, cancel_df = get_status_dfs(master_df)
# ==================================================
# Preserve active tab state
# ==================================================
if "active_tab" not in st.session_state:
    st.session_state["active_tab"] = "📊 SO Summary"
nav_tabs = st.tabs(["📊 SO Summary", "📈 Visual Insights", "➕ Manage Resource"])
with nav_tabs[0]:
    # ---- Your old: with tab1: content goes here ----
    st.session_state["active_tab"] = "📊 SO Summary"
    master_df = st.session_state["master_df"]
    open_df, closed_df, cancel_df = get_status_dfs(master_df)
    df = master_df
    # ==================================================
    # PREMIUM STYLES
    # ==================================================
    st.markdown(
        """
        <style>
        .kpi-big {
            padding: 28px 24px;
            border-radius: 16px;
            font-weight: 800;
            box-shadow: 0 8px 20px rgba(0,0,0,0.2);
            color: white;
            text-align: center;
            height: 180px !important;
            min-height: 180px !important;
            max-height: 180px !important;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            font-size: 22px;
            letter-spacing: 0.5px;
        }
        .kpi-big .total {
            font-size: 28px;
            margin-bottom: 12px;
            font-weight: 800;
        }
        .kpi-big .bottom-row {
            display: flex;
            justify-content: space-between;
            padding: 0 20px;
            margin-bottom: 10px;
            gap: 40px;
            width: 100%;
        }
        .kpi-big .bottom-row div {
            font-size: 18px;
            font-weight: 700;
        }
        .kpi-mini {
            padding: 14px;
            border-radius: 12px;
            text-align: center;
            font-weight: 800;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            color: white;
        }
        .kpi-mini-val {
            font-size: 56px;
            margin-top: 12px;
            font-weight: 900;
            line-height: 1.2;
        }
        /* Clickable KPI styling */
        .kpi-clickable {
            cursor: pointer;
            transition: all 0.3s ease;
            position: relative;
        }
        .kpi-clickable:hover {
            transform: translateY(-4px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.15) !important;
        }
        .kpi-clickable.selected {
            border: 3px solid #ffd700;
            box-shadow: 0 0 20px rgba(255,215,0,0.5) !important;
        }
        /* Table container styling */
        .kpi-table-container {
            margin-top: 24px;
            padding: 20px;
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        .kpi-table-header {
            background: linear-gradient(90deg, #1d3557, #457b9d);
            padding: 12px 20px;
            border-radius: 8px;
            color: white;
            font-size: 20px;
            font-weight: 700;
            margin-bottom: 16px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    # ==================================================
    # Initialize KPI selection state
    # ==================================================
    if "selected_kpi" not in st.session_state:
        st.session_state["selected_kpi"] = None
    if "selected_kpi_label" not in st.session_state:
        st.session_state["selected_kpi_label"] = None

    # ==================================================
    # Helper function to filter data based on KPI
    # ==================================================
    def get_kpi_filtered_data(kpi_type, df_norm, df_orig):
        """Filter data based on selected KPI"""
        if kpi_type is None:
            return None, None
        filtered_norm = df_norm.copy()
        filtered_orig = df_orig.copy()
        mask = None
        # Handle summary KPIs
        if kpi_type == "summary_total":
            return filtered_orig.copy(), filtered_norm.copy()
        elif kpi_type == "summary_open":
            mask = filtered_norm[STATUS_COL] == "OPEN"
        elif kpi_type == "summary_closed":
            mask = filtered_norm[STATUS_COL] == "CLOSED"
        elif kpi_type == "summary_cancel":
            mask = filtered_norm[STATUS_COL] == "CANCEL"
        else:
            # Parse KPI type: format is "status|filter_type|value"
            parts = kpi_type.split("|")
            if len(parts) < 2:
                return None, None
            status = parts[0]
            filter_type = parts[1]
            value = parts[2] if len(parts) > 2 else None
            # Filter by status first
            if status == "OPEN":
                mask = filtered_norm[STATUS_COL] == "OPEN"
            elif status == "CLOSED":
                mask = filtered_norm[STATUS_COL] == "CLOSED"
            elif status == "CANCEL":
                mask = filtered_norm[STATUS_COL] == "CANCEL"
            elif status == "ALL":
                mask = pd.Series([True] * len(filtered_norm), index=filtered_norm.index)
            else:
                return None, None
            # Apply additional filter based on filter_type
            if filter_type == "need_attention":
                mask = mask & (filtered_norm[ESC_COL].str.contains("NEED", na=False))
            elif filter_type == "escalations":
                mask = (
                    mask
                    & (filtered_norm[ESC_COL].notna())
                    & (~filtered_norm[ESC_COL].str.contains("NEED", na=False))
                )
            elif filter_type == "critical":
                mask = mask & (filtered_norm[PRIORITY_COL] == "CRITICAL")
            elif filter_type == "non_critical":
                mask = mask & (filtered_norm[PRIORITY_COL] == "NON CRITICAL")
            elif filter_type == "mfr_no":
                mask = mask & (filtered_norm[MFR_COL].str.upper().isin(["N", "NO"]))
            elif filter_type == "mfr_yes":
                mask = mask & (filtered_norm[MFR_COL].str.upper().isin(["Y", "YES"]))
            elif filter_type == "fulfilled":
                mask = mask & (
                    filtered_norm[DEMAND_COL].str.contains(
                        "FULFILLLED INTERNALLY", na=False
                    )
                )
                if value == "new":
                    mask = mask & (filtered_norm[REQ_COL] == "NEW")
                elif value == "replacement":
                    mask = mask & (filtered_norm[REQ_COL] == "REPLACEMENT")
            elif filter_type == "hiring":
                mask = mask & (filtered_norm[DEMAND_COL] == "HIRING")
                if value == "new":
                    mask = mask & (filtered_norm[REQ_COL] == "NEW")
                elif value == "replacement":
                    mask = mask & (filtered_norm[REQ_COL] == "REPLACEMENT")
            elif filter_type == "proactive":
                mask = mask & (filtered_norm[DEMAND_COL] == "PROACTIVE")
                if value == "new":
                    mask = mask & (filtered_norm[REQ_COL] == "NEW")
                elif value == "replacement":
                    mask = mask & (filtered_norm[REQ_COL] == "REPLACEMENT")
            elif filter_type == "revenue_loss":
                mask = mask & (filtered_norm[LOSS_COL].str.upper() == "Y")
            elif filter_type == "so_count":
                mask = (
                    mask
                    & (filtered_norm[SO_COL].notna())
                    & (filtered_norm[SO_COL].astype(str).str.strip() != "")
                )
            elif filter_type == "needed_by":
                # Filter by needed_by month <= current month
                if NEEDED_BY_COL in filtered_norm.columns:
                    current_month = datetime.now().month
                    needed_by_mask = pd.Series(
                        [False] * len(filtered_norm), index=filtered_norm.index
                    )
                    for idx in filtered_norm.index:
                        month_val = filtered_norm.loc[idx, NEEDED_BY_COL]
                        month_num = parse_month_name(month_val)
                        if month_num is not None and month_num <= current_month:
                            needed_by_mask.loc[idx] = True
                    mask = mask & needed_by_mask
            # For status-only filters, mask is already set
        if mask is None:
            return None, None
        return filtered_orig[mask].copy(), filtered_norm[mask].copy()

    # ==================================================
    # Helper function to display colorful table below clicked KPI
    # ==================================================
    def display_kpi_table(
        selected_kpi,
        selected_label,
        df_normalized,
        master_df,
        column_name_mapping,
        original_column_names,
        clean_col,
    ):
        """Display colorful table for selected KPI"""
        filtered_df, _ = get_kpi_filtered_data(selected_kpi, df_normalized, master_df)
        if filtered_df is not None and len(filtered_df) > 0:
            col_header, col_clear = st.columns([10, 1])
            with col_header:
                st.markdown(
                    f"""
                    <div class="kpi-table-header" style="margin-bottom: 10px; margin-top: 20px;">
                        📋 {selected_label} - {len(filtered_df)} record(s) found
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with col_clear:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button(
                    "✕ Clear",
                    help="Clear KPI selection",
                    key=f"clear_{selected_kpi}_{st.session_state.get('active_tab','tab')}",
                    use_container_width=True,
                ):
                    st.session_state["selected_kpi"] = None
                    st.session_state["selected_kpi_label"] = None
                    ()
            # Display filtered table with original column names
            display_df = filtered_df.copy()
            # Exclude sno and needed_by columns
            excluded_cols = ["sno", "needed_by", "needed by", "neededby", "needed-by"]
            excluded_cols_normalized = [clean_col(c) for c in excluded_cols]
            display_df = display_df.loc[
                :, ~display_df.columns.str.lower().isin(excluded_cols_normalized)
            ]
            # Restore original column names for display
            display_columns = {}
            for col in display_df.columns:
                if col in column_name_mapping:
                    display_columns[col] = column_name_mapping[col]
                else:
                    for orig_name in original_column_names:
                        if clean_col(orig_name) == col:
                            display_columns[col] = orig_name
                            break
            display_df_renamed = display_df.rename(columns=display_columns)
            display_df_renamed = display_df_renamed.replace(r"^\s*$", pd.NA, regex=True)
            display_df_renamed = display_df_renamed.dropna(how="all")

            # Colorful table styling with no empty rows
            def style_colorful_table(df):
                def format_cell(x):
                    if pd.isna(x):
                        return x
                    # Try to format as number if it's numeric
                    try:
                        num = float(x)
                        return format_large_number(num)
                    except (ValueError, TypeError):
                        return x

                return df.style.format(format_cell).set_table_styles(
                    [
                        {
                            "selector": "thead th",
                            "props": [
                                (
                                    "background",
                                    "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
                                ),
                                ("color", "white"),
                                ("font-weight", "bold"),
                                ("font-size", "14px"),
                                ("text-align", "center"),
                                ("padding", "12px 8px"),
                                ("border", "none"),
                            ],
                        },
                        {
                            "selector": "tbody tr:nth-child(odd)",
                            "props": [
                                ("background-color", "#f8f9fa"),
                            ],
                        },
                        {
                            "selector": "tbody tr:nth-child(even)",
                            "props": [
                                ("background-color", "#ffffff"),
                            ],
                        },
                        {
                            "selector": "tbody tr:hover",
                            "props": [
                                ("background-color", "#e3f2fd"),
                                ("transition", "background-color 0.2s"),
                            ],
                        },
                        {
                            "selector": "tbody td",
                            "props": [
                                ("padding", "10px 8px"),
                                ("text-align", "center"),
                                ("border-bottom", "1px solid #e0e0e0"),
                                ("font-size", "13px"),
                            ],
                        },
                        {
                            "selector": "tbody tr:last-child td",
                            "props": [("border-bottom", "none")],
                        },
                    ]
                )

            # Calculate exact height needed (no extra rows)
            num_rows = len(display_df_renamed)
            header_height = 50
            row_height = 45
            exact_height = header_height + (num_rows * row_height) + 10  # 10px padding
            st.dataframe(
                style_colorful_table(display_df_renamed),
                use_container_width=True,
                hide_index=True,
            )

    # ==================================================
    # ==================================================
    # HEADER
    # ==================================================
    st.markdown(
        """
        <div style="
            background: linear-gradient(90deg,#1d3557,#457b9d);
            padding:14px;
            border-radius:12px;
            color:white;
            font-size:28px;
            font-weight:900;
            margin-bottom:10px;">
            📈 Visual Insights — Key Metrics & Trends
        </div>
        """,
        unsafe_allow_html=True,
    )
    # ==================================================
    # KPI STRIP (VISUAL) - CLICKABLE
    # ==================================================
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_key = "summary_total"
        is_selected = st.session_state.get("selected_kpi") == kpi_key
        if st.button(
            f"📊\n**Total SOs**\n\n### {len(df)}",
            key=f"btn_{kpi_key}",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Total SOs"
            ()
    with k2:
        kpi_key = "summary_open"
        is_selected = st.session_state.get("selected_kpi") == kpi_key
        if st.button(
            f"🟢\n**Open SOs**\n\n### {len(open_df)}",
            key=f"btn_{kpi_key}",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open SOs"
            ()
    with k3:
        kpi_key = "summary_closed"
        is_selected = st.session_state.get("selected_kpi") == kpi_key
        if st.button(
            f"🔵\n**Closed SOs**\n\n### {len(closed_df)}",
            key=f"btn_{kpi_key}",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed SOs"
            ()
    with k4:
        kpi_key = "summary_cancel"
        is_selected = st.session_state.get("selected_kpi") == kpi_key
        if st.button(
            f"⚪\n**Cancelled SOs**\n\n### {len(cancel_df)}",
            key=f"btn_{kpi_key}",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled SOs"
            ()
    # Display table below Summary KPIs if a Summary KPI is clicked
    if st.session_state.get("selected_kpi") and st.session_state.get(
        "selected_kpi", ""
    ).startswith("summary_"):
        selected_kpi = st.session_state["selected_kpi"]
        selected_label = st.session_state.get("selected_kpi_label", "Summary KPI")
        display_kpi_table(
            selected_kpi,
            selected_label,
            df_normalized,
            master_df,
            column_name_mapping,
            original_column_names,
            clean_col,
        )
    # Apply custom styling to KPI buttons
    st.markdown(
        """
        <style>
        div[data-testid*="btn_summary_"] button {
            color: white !important;
            border: none !important;
            border-radius: 20px !important;
            padding: 28px 20px !important;
            font-size: 20px !important;
            font-weight: 800 !important;
            box-shadow: 0 10px 25px rgba(0,0,0,0.15), 0 4px 10px rgba(0,0,0,0.1) !important;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1) !important;
            height: 180px !important;
            min-height: 180px !important;
            max-height: 180px !important;
            line-height: 1.4 !important;
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            align-items: center !important;
            position: relative !important;
            overflow: hidden !important;
        }
        div[data-testid*="btn_summary_"] button::before {
            content: '' !important;
            position: absolute !important;
            top: 0 !important;
            left: -100% !important;
            width: 100% !important;
            height: 100% !important;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent) !important;
            transition: left 0.5s !important;
        }
        div[data-testid*="btn_summary_"] button:hover::before {
            left: 100% !important;
        }
        div[data-testid*="btn_summary_"] button:hover {
            transform: translateY(-8px) scale(1.02) !important;
            box-shadow: 0 15px 35px rgba(0,0,0,0.25), 0 5px 15px rgba(0,0,0,0.15) !important;
        }
        div[data-testid*="btn_summary_"] button h3 {
            font-size: 52px !important;
            font-weight: 900 !important;
            margin: 10px 0 !important;
            line-height: 1.2 !important;
        }
        /* Vibrant Colorful Gradients for Summary KPIs - Enhanced */
        div[data-testid*="btn_summary_total"] button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_summary_open"] button {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_summary_closed"] button {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_summary_cancel"] button {
            background: linear-gradient(135deg, #6b7280 0%, #4b5563 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        /* Style for Open Requirements KPI buttons - Equal Size & Colorful - Enhanced */
        div[data-testid*="btn_open_"] button {
            color: white !important;
            border: none !important;
            border-radius: 20px !important;
            padding: 28px 20px !important;
            font-size: 18px !important;
            font-weight: 800 !important;
            box-shadow: 0 10px 25px rgba(0,0,0,0.15), 0 4px 10px rgba(0,0,0,0.1) !important;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1) !important;
            height: 180px !important;
            min-height: 180px !important;
            max-height: 180px !important;
            white-space: normal !important;
            line-height: 1.4 !important;
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            align-items: center !important;
            position: relative !important;
            overflow: hidden !important;
        }
        div[data-testid*="btn_open_"] button::before {
            content: '' !important;
            position: absolute !important;
            top: 0 !important;
            left: -100% !important;
            width: 100% !important;
            height: 100% !important;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent) !important;
            transition: left 0.5s !important;
        }
        div[data-testid*="btn_open_"] button:hover::before {
            left: 100% !important;
        }
        div[data-testid*="btn_open_"] button:hover {
            transform: translateY(-8px) scale(1.02) !important;
            box-shadow: 0 15px 35px rgba(0,0,0,0.25), 0 5px 15px rgba(0,0,0,0.15) !important;
        }
        div[data-testid*="btn_open_"] button h3 {
            font-size: 52px !important;
            font-weight: 900 !important;
            margin: 8px 0 !important;
            line-height: 1.2 !important;
        }
        /* Vibrant Colorful Gradients for Open KPIs - Enhanced */
        div[data-testid*="btn_open_need_attn"] button {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_escalations"] button {
            background: linear-gradient(135deg, #f97316 0%, #ea580c 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_critical"] button {
            background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_noncritical"] button {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_mfr_no"] button {
            background: linear-gradient(135deg, #f87171 0%, #ef4444 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_mfr_yes"] button {
            background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_fulfilled"] button {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_hiring"] button {
            background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_proactive"] button {
            background: linear-gradient(135deg, #06b6d4 0%, #0891b2 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_loss"] button {
            background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_so_count"] button {
            background: linear-gradient(135deg, #14b8a6 0%, #0d9488 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_open_needed_by"] button {
            background: linear-gradient(135deg, #ec4899 0%, #db2777 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        /* Style for Closed and Cancelled Requirements KPI buttons - Equal Size & Colorful - Enhanced */
        div[data-testid*="btn_closed_"] button, div[data-testid*="btn_cancel_"] button {
            color: white !important;
            border: none !important;
            border-radius: 20px !important;
            padding: 28px 20px !important;
            font-size: 18px !important;
            font-weight: 800 !important;
            box-shadow: 0 10px 25px rgba(0,0,0,0.15), 0 4px 10px rgba(0,0,0,0.1) !important;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1) !important;
            height: 180px !important;
            min-height: 180px !important;
            max-height: 180px !important;
            white-space: normal !important;
            line-height: 1.4 !important;
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            align-items: center !important;
            position: relative !important;
            overflow: hidden !important;
        }
        div[data-testid*="btn_closed_"] button::before, div[data-testid*="btn_cancel_"] button::before {
            content: '' !important;
            position: absolute !important;
            top: 0 !important;
            left: -100% !important;
            width: 100% !important;
            height: 100% !important;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent) !important;
            transition: left 0.5s !important;
        }
        div[data-testid*="btn_closed_"] button:hover::before, div[data-testid*="btn_cancel_"] button:hover::before {
            left: 100% !important;
        }
        div[data-testid*="btn_closed_"] button:hover, div[data-testid*="btn_cancel_"] button:hover {
            transform: translateY(-8px) scale(1.02) !important;
            box-shadow: 0 15px 35px rgba(0,0,0,0.25), 0 5px 15px rgba(0,0,0,0.15) !important;
        }
        div[data-testid*="btn_closed_"] button h3, div[data-testid*="btn_cancel_"] button h3 {
            font-size: 52px !important;
            font-weight: 900 !important;
            margin: 8px 0 !important;
            line-height: 1.2 !important;
        }
        /* Vibrant Colorful Gradients for Closed KPIs - Enhanced */
        div[data-testid*="btn_closed_need_attn"] button {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_escalations"] button {
            background: linear-gradient(135deg, #f97316 0%, #ea580c 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_critical"] button {
            background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_noncritical"] button {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_mfr_no"] button {
            background: linear-gradient(135deg, #f87171 0%, #ef4444 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_mfr_yes"] button {
            background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_fulfilled"] button {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_hiring"] button {
            background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_proactive"] button {
            background: linear-gradient(135deg, #06b6d4 0%, #0891b2 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_loss"] button {
            background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_so_count"] button {
            background: linear-gradient(135deg, #14b8a6 0%, #0d9488 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_closed_needed_by"] button {
            background: linear-gradient(135deg, #ec4899 0%, #db2777 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        /* Vibrant Colorful Gradients for Cancelled KPIs - Enhanced */
        div[data-testid*="btn_cancel_need_attn"] button {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_escalations"] button {
            background: linear-gradient(135deg, #f97316 0%, #ea580c 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_critical"] button {
            background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_noncritical"] button {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_mfr_no"] button {
            background: linear-gradient(135deg, #f87171 0%, #ef4444 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_mfr_yes"] button {
            background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_fulfilled"] button {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_hiring"] button {
            background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_proactive"] button {
            background: linear-gradient(135deg, #06b6d4 0%, #0891b2 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_loss"] button {
            background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_so_count"] button {
            background: linear-gradient(135deg, #14b8a6 0%, #0d9488 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        div[data-testid*="btn_cancel_needed_by"] button {
            background: linear-gradient(135deg, #ec4899 0%, #db2777 100%) !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
        }
        /* Professional spacing for KPI grid - Enhanced */
        .element-container:has(button[data-testid*="btn_open_"]),
        .element-container:has(button[data-testid*="btn_closed_"]),
        .element-container:has(button[data-testid*="btn_cancel_"]) {
            padding: 12px !important;
        }
        /* Improve column spacing for professional layout - Enhanced */
        [data-testid="column"] {
            gap: 20px !important;
        }
        /* Better spacing between KPI rows - no extra margins */
        .stButton {
            margin-bottom: 0 !important;
        }
        /* Remove extra spacing after KPI sections */
        div[data-testid="stVerticalBlock"] > div:has(button[data-testid*="btn_"]) {
            margin-bottom: 0 !important;
        }
        /* Professional typography improvements - Enhanced */
        div[data-testid*="btn_"] button {
            letter-spacing: 0.5px !important;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2), 0 1px 2px rgba(0,0,0,0.15) !important;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Oxygen', 'Ubuntu', 'Cantarell', sans-serif !important;
        }
        /* Enhanced shadow for depth and professionalism */
        .kpi-big {
            box-shadow: 0 8px 20px rgba(0,0,0,0.2) !important;
        }
        /* Smooth transitions for all interactive elements */
        div[data-testid*="btn_"] button,
        .kpi-big {
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }
        /* Ensure tables appear directly below KPIs with minimal spacing */
        .element-container:has([data-testid="stDataFrame"]) {
            margin-top: 10px !important;
        }
        /* Colorful table styling */
        .stDataFrame {
            border-radius: 12px !important;
            overflow: hidden !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
        }
        .stDataFrame thead th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
            font-weight: 700 !important;
            text-align: center !important;
            padding: 12px 8px !important;
            border: none !important;
        }
        .stDataFrame tbody tr:nth-child(odd) {
            background-color: #f8f9fa !important;
        }
        .stDataFrame tbody tr:nth-child(even) {
            background-color: #ffffff !important;
        }
        .stDataFrame tbody tr:hover {
            background-color: #e3f2fd !important;
            transition: background-color 0.2s !important;
        }
        .stDataFrame tbody td {
            padding: 10px 8px !important;
            text-align: center !important;
            border-bottom: 1px solid #e0e0e0 !important;
        }
        .stDataFrame tbody tr:last-child td {
            border-bottom: none !important;
        }
        /* Remove empty rows - limit table height and hide empty rows */
        div[data-testid="stDataFrame"] {
            max-height: none !important;
        }
        /* Hide empty table rows */
        .stDataFrame tbody tr:empty,
        .stDataFrame tbody tr:has(td:empty:only-child) {
            display: none !important;
        }
        /* Ensure table container doesn't show extra space */
        div[data-testid="stDataFrame"] > div {
            overflow: visible !important;
        }
        /* Remove padding that creates empty row appearance */
        .stDataFrame tbody {
            border-bottom: none !important;
        }
        /* Colorful Clear Button Styling - Target Streamlit button structure */
        div[data-testid*="baseButton-clear_"] button {
            background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%) !important;
            color: white !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
            border-radius: 12px !important;
            padding: 10px 20px !important;
            font-weight: 700 !important;
            font-size: 14px !important;
            box-shadow: 0 4px 12px rgba(244, 63, 94, 0.3) !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
            text-shadow: 0 1px 2px rgba(0,0,0,0.2) !important;
        }
        div[data-testid*="baseButton-clear_"] button:hover {
            background: linear-gradient(135deg, #e11d48 0%, #be185d 100%) !important;
            transform: translateY(-2px) scale(1.05) !important;
            box-shadow: 0 6px 16px rgba(244, 63, 94, 0.4) !important;
        }
        /* Colorful Refresh Button Styling */
        div[data-testid*="baseButton-refresh_button"] button {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            color: white !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
            border-radius: 12px !important;
            font-weight: 700 !important;
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3) !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }
        div[data-testid*="baseButton-refresh_button"] button:hover {
            background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
            transform: rotate(180deg) scale(1.1) !important;
            box-shadow: 0 6px 16px rgba(16, 185, 129, 0.4) !important;
        }
        /* Style buttons that contain Clear text - using a more general approach */
        .stButton button {
            position: relative;
        }
        /* Target buttons in columns next to table headers (where Clear button appears) */
        [data-testid="column"]:has([data-testid*="baseButton-clear_"]) button {
            background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%) !important;
            color: white !important;
            border: 2px solid rgba(255,255,255,0.2) !important;
            border-radius: 12px !important;
            font-weight: 700 !important;
            box-shadow: 0 4px 12px rgba(244, 63, 94, 0.3) !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }
        [data-testid="column"]:has([data-testid*="baseButton-clear_"]) button:hover {
            background: linear-gradient(135deg, #e11d48 0%, #be185d 100%) !important;
            transform: translateY(-2px) scale(1.05) !important;
            box-shadow: 0 6px 16px rgba(244, 63, 94, 0.4) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div style="
            background: linear-gradient(90deg,#1d3557,#457b9d);
            padding:14px;
            border-radius:12px;
            color:white;
            font-size:28px;
            font-weight:900;
            margin-bottom:20px;
            margin-top:20px;">
            📈 Open Requirements Data
        </div>
        """,
        unsafe_allow_html=True,
    )
    # ================================================== # KPI GRID (2 rows × 6 cols) # ==================================================
    # Count how many rows have "need attention" in esc column
    # need_attention_count = (df[ESC_COL] == "Need Attention").sum()
    # Count Need Attention only for rows with specific STATUS
    need_attention_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    need_attention_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    need_attention_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    # Total escalations (non-empty values in ESC_COL)
    total_escalations = len(open_df)
    # Escalations excluding "need attention"
    other_escalations_open = total_escalations - need_attention_open
    # Count Critical and Non-Critical
    critical_open_count = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[PRIORITY_COL] == "CRITICAL")
    ].shape[0]
    # Count Non-Critical only for OPEN status
    non_critical_open_count = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[PRIORITY_COL] == "NON CRITICAL")
    ].shape[0]
    # Count YES in MFR_COL only for OPEN status
    mfr_yes_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[MFR_COL].str.upper().isin(["Y", "YES"]))
    ].shape[0]
    # Count NO in MFR_COL only for OPEN status
    mfr_no_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[MFR_COL].str.upper().isin(["N", "NO"]))
    ].shape[0]
    # 1. Count "Can be fulfilled internally" when status is OPEN
    fulfilled_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
    ].shape[0]
    # 2. Count "Can be fulfilled internally" when status is OPEN and REQ_COL = New
    fulfilled_open_new = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Can be fulfilled internally" when status is OPEN and REQ_COL = Replacement
    fulfilled_open_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # 1. Count "Hiring" when status is OPEN
    hiring_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN") & (df_normalized[DEMAND_COL] == "HIRING")
    ].shape[0]
    # 2. Count "Hiring" when status is OPEN and REQ_COL = New
    hiring_open_new = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL] == "HIRING")
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Hiring" when status is OPEN and REQ_COL = Replacement
    hiring_open_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL] == "HIRING")
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # 1. Count "Proactive" when status is OPEN
    proactive_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
    ].shape[0]
    # 2. Count "Proactive" when status is OPEN and REQ_COL = New
    proactive_open_new = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Proactive" when status is OPEN and REQ_COL = Replacement
    proactive_open_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # Count YES in LOSS_COL only for OPEN status
    loss_yes_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[LOSS_COL].str.upper() == "Y")
    ].shape[0]
    so_count = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[SO_COL].notna())
        & (df_normalized[SO_COL].astype(str).str.strip() != "")
    ].shape[0]
    # Count Needed by (month <= current month) for OPEN status
    needed_by_open = count_needed_by(df_normalized, STATUS_COL, "OPEN")
    # # First row - Clickable KPIs
    row1 = st.columns(6)
    with row1[0]:
        kpi_key = "OPEN|need_attention|"
        if st.button(
            f"Needed Attention\n### {need_attention_open}",
            key=f"btn_open_need_attn",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Needed Attention"
            ()
    with row1[1]:
        kpi_key = "OPEN|escalations|"
        if st.button(
            f"Escalations Count\n### {other_escalations_open}",
            key=f"btn_open_escalations",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Escalations"
            ()
    with row1[2]:
        kpi_key = "OPEN|critical|"
        if st.button(
            f"Critical Count\n### {critical_open_count}",
            key=f"btn_open_critical",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Critical"
            ()
    with row1[3]:
        kpi_key = "OPEN|non_critical|"
        if st.button(
            f"Non-Critical Count\n### {non_critical_open_count}",
            key=f"btn_open_noncritical",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Non-Critical"
            ()
    with row1[4]:
        kpi_key = "OPEN|mfr_no|"
        if st.button(
            f"MFR-NO\n### {mfr_no_open}",
            key=f"btn_open_mfr_no",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - MFR NO"
            ()
    with row1[5]:
        kpi_key = "OPEN|mfr_yes|"
        if st.button(
            f"MFR-YES\n### {mfr_yes_open}",
            key=f"btn_open_mfr_yes",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - MFR YES"
            ()
    # Second row - Clickable KPIs (no extra spacing)
    row2 = st.columns(6)
    with row2[0]:
        kpi_key = "OPEN|fulfilled|"
        if st.button(
            f"Fulfilled: {fulfilled_open}\nNew: {fulfilled_open_new} | Rep: {fulfilled_open_replacement}",
            key=f"btn_open_fulfilled",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = (
                "Open - Can Be Fulfilled Internally"
            )
            ()
    with row2[1]:
        kpi_key = "OPEN|hiring|"
        if st.button(
            f"Hiring: {hiring_open}\nNew: {hiring_open_new} | Rep: {hiring_open_replacement}",
            key=f"btn_open_hiring",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Hiring"
            ()
    with row2[2]:
        kpi_key = "OPEN|proactive|"
        if st.button(
            f"Proactive: {proactive_open}\nNew: {proactive_open_new} | Rep: {proactive_open_replacement}",
            key=f"btn_open_proactive",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Proactive"
            ()
    with row2[3]:
        kpi_key = "OPEN|revenue_loss|"
        if st.button(
            f"Revenue Loss\n### {loss_yes_open}",
            key=f"btn_open_loss",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Revenue Loss"
            ()
    with row2[4]:
        kpi_key = "OPEN|so_count|"
        if st.button(
            f"SO Count\n### {so_count}",
            key=f"btn_open_so_count",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - SO Count"
            ()
    with row2[5]:
        kpi_key = "OPEN|needed_by|"
        if st.button(
            f"Needed by\n### {needed_by_open}",
            key=f"btn_open_needed_by",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Open - Needed by"
            ()
    # Display table below Open Requirements KPIs if an Open KPI is clicked
    if st.session_state.get("selected_kpi") and st.session_state.get(
        "selected_kpi", ""
    ).startswith("OPEN|"):
        selected_kpi = st.session_state["selected_kpi"]
        selected_label = st.session_state.get("selected_kpi_label", "Open KPI")
        display_kpi_table(
            selected_kpi,
            selected_label,
            df_normalized,
            master_df,
            column_name_mapping,
            original_column_names,
            clean_col,
        )

    # CLosed Requirements
    # ================================================== # KPI GRID (2 rows × 6 cols) # ==================================================
    # Count how many rows have "need attention" in esc column
    # need_attention_count = (df[ESC_COL] == "Need Attention").sum()
    # Count Need Attention only for rows with specific STATUS
    st.markdown(
        """
        <div style="
            background: linear-gradient(90deg,#1d3557,#457b9d);
            padding:14px;
            border-radius:12px;
            color:white;
            font-size:28px;
            font-weight:900;
            margin-bottom:20px;
            margin-top:20px;">
            📈 Closed Requirements Data
        </div>
        """,
        unsafe_allow_html=True,
    )
    need_attention_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    # Correct way to count Need Attention when status is CLOSED (case-insensitive)
    need_attention_closed = df_normalized[
        ((df_normalized[STATUS_COL].str.upper() == "CLOSED"))
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    need_attention_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    # Total escalations (non-empty values in ESC_COL)
    total_escalations_closed = len(closed_df)
    # Escalations excluding "need attention"
    other_escalations_closed = total_escalations_closed - need_attention_closed
    # Count Critical and Non-Critical
    critical_closed_count = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[PRIORITY_COL] == "CRITICAL")
    ].shape[0]
    # Count Non-Critical only for closed status
    non_critical_closed_count = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[PRIORITY_COL] == "NON CRITICAL")
    ].shape[0]
    # Count YES in MFR_COL only for closed status
    mfr_yes_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[MFR_COL].str.upper().isin(["Y", "YES"]))
    ].shape[0]
    # Count NO in MFR_COL only for closed status
    mfr_no_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[MFR_COL].str.upper().isin(["N", "NO"]))
    ].shape[0]
    # 1. Count "Can be fulfilled internally" when status is closed
    fulfilled_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
    ].shape[0]
    # 2. Count "Can be fulfilled internally" when status is OPEN and REQ_COL = New
    fulfilled_closed_new = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Can be fulfilled internally" when status is OPEN and REQ_COL = Replacement
    fulfilled_closed_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # 1. Count "Hiring" when status is CLOSED
    hiring_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL] == "HIRING")
    ].shape[0]
    # 2. Count "Hiring" when status is OPEN and REQ_COL = New
    hiring_closed_new = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL] == "HIRING")
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Hiring" when status is OPEN and REQ_COL = Replacement
    hiring_closed_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL] == "HIRING")
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # 1. Count "Proactive" when status is OPEN
    proactive_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
    ].shape[0]
    # 2. Count "Proactive" when status is OPEN and REQ_COL = New
    proactive_closed_new = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Proactive" when status is OPEN and REQ_COL = Replacement
    proactive_closed_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # Count YES in LOSS_COL only for OPEN status
    loss_yes_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[LOSS_COL].str.upper() == "Y")
    ].shape[0]
    so_count_closed = df_normalized[
        (df_normalized[STATUS_COL] == "CLOSED")
        & (df_normalized[SO_COL].notna())
        & (df_normalized[SO_COL].astype(str).str.strip() != "")
    ].shape[0]
    # Count Needed by (month <= current month) for CLOSED status
    needed_by_closed = count_needed_by(df_normalized, STATUS_COL, "CLOSED")
    # # First row - Clickable KPIs for Closed
    row1 = st.columns(6)
    with row1[0]:
        kpi_key = "CLOSED|need_attention|"
        if st.button(
            f"Needed Attention\n### {need_attention_closed}",
            key=f"btn_closed_need_attn",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Needed Attention"
            ()
    with row1[1]:
        kpi_key = "CLOSED|escalations|"
        if st.button(
            f"Escalations Count\n### {other_escalations_closed}",
            key=f"btn_closed_escalations",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Escalations"
            ()
    with row1[2]:
        kpi_key = "CLOSED|critical|"
        if st.button(
            f"Critical Count\n### {critical_closed_count}",
            key=f"btn_closed_critical",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Critical"
            ()
    with row1[3]:
        kpi_key = "CLOSED|non_critical|"
        if st.button(
            f"Non-Critical Count\n### {non_critical_closed_count}",
            key=f"btn_closed_noncritical",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Non-Critical"
            ()
    with row1[4]:
        kpi_key = "CLOSED|mfr_no|"
        if st.button(
            f"MFR-NO\n### {mfr_no_closed}",
            key=f"btn_closed_mfr_no",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - MFR NO"
            ()
    with row1[5]:
        kpi_key = "CLOSED|mfr_yes|"
        if st.button(
            f"MFR-YES\n### {mfr_yes_closed}",
            key=f"btn_closed_mfr_yes",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - MFR YES"
            ()
    # Second row - Clickable KPIs for Closed
    row2 = st.columns(6)
    with row2[0]:
        kpi_key = "CLOSED|fulfilled|"
        if st.button(
            f"Fulfilled: {fulfilled_closed}\nNew: {fulfilled_closed_new} | Rep: {fulfilled_closed_replacement}",
            key=f"btn_closed_fulfilled",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = (
                "Closed - Can Be Fulfilled Internally"
            )
            ()
    with row2[1]:
        kpi_key = "CLOSED|hiring|"
        if st.button(
            f"Hiring: {hiring_closed}\nNew: {hiring_closed_new} | Rep: {hiring_closed_replacement}",
            key=f"btn_closed_hiring",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Hiring"
            ()
    with row2[2]:
        kpi_key = "CLOSED|proactive|"
        if st.button(
            f"Proactive: {proactive_closed}\nNew: {proactive_closed_new} | Rep: {proactive_closed_replacement}",
            key=f"btn_closed_proactive",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Proactive"
            ()
    with row2[3]:
        kpi_key = "CLOSED|revenue_loss|"
        if st.button(
            f"Revenue Loss\n### {loss_yes_closed}",
            key=f"btn_closed_loss",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Revenue Loss"
            ()
    with row2[4]:
        kpi_key = "CLOSED|so_count|"
        if st.button(
            f"SO Count\n### {so_count_closed}",
            key=f"btn_closed_so_count",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - SO Count"
            ()
    with row2[5]:
        kpi_key = "CLOSED|needed_by|"
        if st.button(
            f"Needed by\n### {needed_by_closed}",
            key=f"btn_closed_needed_by",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Closed - Needed by"
            ()
    # Display table below Closed Requirements KPIs if a Closed KPI is clicked
    if st.session_state.get("selected_kpi") and st.session_state.get(
        "selected_kpi", ""
    ).startswith("CLOSED|"):
        selected_kpi = st.session_state["selected_kpi"]
        selected_label = st.session_state.get("selected_kpi_label", "Closed KPI")
        display_kpi_table(
            selected_kpi,
            selected_label,
            df_normalized,
            master_df,
            column_name_mapping,
            original_column_names,
            clean_col,
        )
    # Cancelled Requirements
    # ================================================== # KPI GRID (2 rows × 6 cols) # ==================================================
    # Count how many rows have "need attention" in esc column
    # need_attention_count = (df[ESC_COL] == "Need Attention").sum()
    # Count Need Attention only for rows with specific STATUS
    st.markdown(
        """
        <div style="
            background: linear-gradient(90deg,#1d3557,#457b9d);
            padding:14px;
            border-radius:12px;
            color:white;
            font-size:28px;
            font-weight:900;
            margin-bottom:20px;
            margin-top:20px;">
            📈 Cancelled Requirements Data
        </div>
        """,
        unsafe_allow_html=True,
    )
    need_attention_open = df_normalized[
        (df_normalized[STATUS_COL] == "OPEN")
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    # Correct way to count Need Attention when status is CLOSED (case-insensitive)
    need_attention_closed = df_normalized[
        ((df_normalized[STATUS_COL].str.upper() == "CLOSED"))
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    need_attention_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[ESC_COL].str.contains("NEED", na=False))
    ].shape[0]
    # Total escalations (non-empty values in ESC_COL)
    total_escalations_cancel = len(cancel_df)
    # Escalations excluding "need attention"
    other_escalations_cancel = total_escalations_cancel - need_attention_cancel
    # Count Critical and Non-Critical
    critical_cancel_count = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[PRIORITY_COL] == "CRITICAL")
    ].shape[0]
    # Count Non-Critical only for closed status
    non_critical_cancel_count = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[PRIORITY_COL] == "NON CRITICAL")
    ].shape[0]
    # Count YES in MFR_COL only for closed status
    mfr_yes_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[MFR_COL].str.upper().isin(["Y", "YES"]))
    ].shape[0]
    # Count NO in MFR_COL only for closed status
    mfr_no_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[MFR_COL].str.upper().isin(["N", "NO"]))
    ].shape[0]
    # 1. Count "Can be fulfilled internally" when status is closed
    fulfilled_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
    ].shape[0]
    # 2. Count "Can be fulfilled internally" when status is OPEN and REQ_COL = New
    fulfilled_cancel_new = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Can be fulfilled internally" when status is OPEN and REQ_COL = Replacement
    fulfilled_cancel_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL].str.contains("FULFILLLED INTERNALLY", na=False))
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # 1. Count "Hiring" when status is CLOSED
    hiring_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL] == "HIRING")
    ].shape[0]
    # 2. Count "Hiring" when status is OPEN and REQ_COL = New
    hiring_cancel_new = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL] == "HIRING")
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Hiring" when status is OPEN and REQ_COL = Replacement
    hiring_cancel_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL] == "HIRING")
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # 1. Count "Proactive" when status is OPEN
    proactive_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
    ].shape[0]
    # 2. Count "Proactive" when status is OPEN and REQ_COL = New
    proactive_cancel_new = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
        & (df_normalized[REQ_COL] == "NEW")
    ].shape[0]
    # 3. Count "Proactive" when status is OPEN and REQ_COL = Replacement
    proactive_cancel_replacement = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[DEMAND_COL] == "PROACTIVE")
        & (df_normalized[REQ_COL] == "REPLACEMENT")
    ].shape[0]
    # Count YES in LOSS_COL only for OPEN status
    loss_yes_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[LOSS_COL].str.upper() == "Y")
    ].shape[0]
    so_count_cancel = df_normalized[
        (df_normalized[STATUS_COL] == "CANCEL")
        & (df_normalized[SO_COL].notna())
        & (df_normalized[SO_COL].astype(str).str.strip() != "")
    ].shape[0]
    # Count Needed by (month <= current month) for CANCEL status
    needed_by_cancel = count_needed_by(df_normalized, STATUS_COL, "CANCEL")
    # # First row - Clickable KPIs for Cancelled
    row1 = st.columns(6)
    with row1[0]:
        kpi_key = "CANCEL|need_attention|"
        if st.button(
            f"Needed Attention\n### {need_attention_cancel}",
            key=f"btn_cancel_need_attn",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Needed Attention"
            ()
    with row1[1]:
        kpi_key = "CANCEL|escalations|"
        if st.button(
            f"Escalations Count\n### {other_escalations_cancel}",
            key=f"btn_cancel_escalations",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Escalations"
            ()
    with row1[2]:
        kpi_key = "CANCEL|critical|"
        if st.button(
            f"Critical Count\n### {critical_cancel_count}",
            key=f"btn_cancel_critical",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Critical"
            ()
    with row1[3]:
        kpi_key = "CANCEL|non_critical|"
        if st.button(
            f"Non-Critical Count\n### {non_critical_cancel_count}",
            key=f"btn_cancel_noncritical",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Non-Critical"
            ()
    with row1[4]:
        kpi_key = "CANCEL|mfr_no|"
        if st.button(
            f"MFR-NO\n### {mfr_no_cancel}",
            key=f"btn_cancel_mfr_no",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - MFR NO"
            ()
    with row1[5]:
        kpi_key = "CANCEL|mfr_yes|"
        if st.button(
            f"MFR-YES\n### {mfr_yes_cancel}",
            key=f"btn_cancel_mfr_yes",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - MFR YES"
            ()
    # Second row - Clickable KPIs for Cancelled
    row2 = st.columns(6)
    with row2[0]:
        kpi_key = "CANCEL|fulfilled|"
        if st.button(
            f"Fulfilled: {fulfilled_cancel}\nNew: {fulfilled_cancel_new} | Rep: {fulfilled_cancel_replacement}",
            key=f"btn_cancel_fulfilled",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = (
                "Cancelled - Can Be Fulfilled Internally"
            )
            ()
    with row2[1]:
        kpi_key = "CANCEL|hiring|"
        if st.button(
            f"Hiring: {hiring_cancel}\nNew: {hiring_cancel_new} | Rep: {hiring_cancel_replacement}",
            key=f"btn_cancel_hiring",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Hiring"
            ()
    with row2[2]:
        kpi_key = "CANCEL|proactive|"
        if st.button(
            f"Proactive: {proactive_cancel}\nNew: {proactive_cancel_new} | Rep: {proactive_cancel_replacement}",
            key=f"btn_cancel_proactive",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Proactive"
            ()
    with row2[3]:
        kpi_key = "CANCEL|revenue_loss|"
        if st.button(
            f"Revenue Loss\n### {loss_yes_cancel}",
            key=f"btn_cancel_loss",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Revenue Loss"
            ()
    with row2[4]:
        kpi_key = "CANCEL|so_count|"
        if st.button(
            f"SO Count\n### {so_count_cancel}",
            key=f"btn_cancel_so_count",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - SO Count"
            ()
    with row2[5]:
        kpi_key = "CANCEL|needed_by|"
        if st.button(
            f"Needed by\n### {needed_by_cancel}",
            key=f"btn_cancel_needed_by",
            use_container_width=True,
        ):
            st.session_state["selected_kpi"] = kpi_key
            st.session_state["selected_kpi_label"] = "Cancelled - Needed by"
            ()
    # Display table below Cancelled Requirements KPIs if a Cancelled KPI is clicked
    if st.session_state.get("selected_kpi") and st.session_state.get(
        "selected_kpi", ""
    ).startswith("CANCEL|"):
        selected_kpi = st.session_state["selected_kpi"]
        selected_label = st.session_state.get("selected_kpi_label", "Cancelled KPI")
        display_kpi_table(
            selected_kpi,
            selected_label,
            df_normalized,
            master_df,
            column_name_mapping,
            original_column_names,
            clean_col,
        )
    # ==================================================
    # PREMIUM STYLES (TAB 1)
    # ==================================================
    st.markdown(
        """
        <style>
        /* KPI card styling */
        .kpi-card {
            padding: 18px;
            border-radius: 14px;
            text-align: center;
            font-weight: 900;
            color: white;
            box-shadow: 0 6px 14px rgba(0,0,0,0.12);
        }
        .kpi-card-value {
            font-size: 42px;
            margin-top: 6px;
        }
        .section-header {
            background: linear-gradient(90deg,#1d3557,#457b9d);
            padding: 14px;
            border-radius: 12px;
            color: white;
            font-size: 26px;
            font-weight: 900;
            margin-bottom: 12px;
        }
        /* Card heading base */
        .card-heading {
            padding: 10px 14px;
            border-radius: 10px;
            font-size: 22px;
            font-weight: 900;
            margin-bottom: 10px;
            text-align: center;
        }
        /* Individual colors */
        .open-box {
            background: linear-gradient(90deg, #d8f3dc, #b7e4c7);
            color: #1b4332;
        }
        .closed-box {
            background: linear-gradient(90deg, #dbeafe, #bfdbfe);
            color: #1e3a8a;
        }
        .cancel-box {
            background: linear-gradient(90deg, #f1f5f9, #e5e7eb);
            color: #374151;
        }
        /* Center headers and values in dataframes */
        .stDataFrame th {
            text-align: center !important;
            font-weight: bold !important;
            background-color: #f8f9fa !important;
        }
        .stDataFrame td {
            text-align: center !important;
            vertical-align: middle !important;
            padding: 6px !important;
        }
        /* Optional: zebra striping for readability */
        .stDataFrame tbody tr:nth-child(odd) td {
            background-color: #ffffff !important;
        }
        .stDataFrame tbody tr:nth-child(even) td {
            background-color: #f7f7f7 !important;
        }
        /* Optional: hover effect */
        .stDataFrame tbody tr:hover td {
            background-color: #e3f2fd !important;
        }
        </style>
    """,
        unsafe_allow_html=True,
    )
    # Spacer
    st.markdown("<br>", unsafe_allow_html=True)
    # Second row: Cancelled SOs
    with st.container(border=True):
        st.markdown(
            '<div class="card-heading open-box">🟢 OPEN SOs</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(
            build_summary(open_df, "Open", len(master_df)).set_table_styles(
                [
                    {"selector": "th", "props": [("background-color", "#e8f5e9")]},
                    {"selector": "td", "props": [("background-color", "#f1f8e9")]},
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )
    # First row: Open and Closed SOs
    c1, c2 = st.columns(2)
    with c2:
        with st.container(border=True):
            st.markdown(
                '<div class="card-heading cancel-box">⚪ CANCELLED SOs</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(
                build_summary(cancel_df, "Cancel", len(master_df)).set_table_styles(
                    [
                        {"selector": "th", "props": [("background-color", "#eeeeee")]},
                        {"selector": "td", "props": [("background-color", "#fafafa")]},
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
    with c1:
        with st.container(border=True):
            st.markdown(
                '<div class="card-heading closed-box">🔵 CLOSED SOs</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(
                build_summary(closed_df, "Closed", len(master_df)).set_table_styles(
                    [
                        {"selector": "th", "props": [("background-color", "#e3f2fd")]},
                        {"selector": "td", "props": [("background-color", "#f5faff")]},
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
    # ==================================================
    # DETAILS HEADER
    # ==================================================
    st.markdown(
        """
        <div class="section-header">
            📄 SO Details & Attention Tracking
        </div>
        """,
        unsafe_allow_html=True,
    )
    # ==================================================
    # FILTER PANEL
    # ==================================================
    with st.container(border=True):
        f1, f2, f3 = st.columns([2, 2, 3])
        with f1:
            table_choice = st.selectbox(
                "SO Status", ["All", "Open SOs", "Closed SOs", "Cancelled SOs"]
            )
        if table_choice == "All":
            base_df = master_df.copy()
        else:
            base_df = (
                open_df
                if table_choice == "Open SOs"
                else closed_df if table_choice == "Closed SOs" else cancel_df
            )
        with f2:
            # Case-insensitive priority filtering with normalized unique values
            if PRIORITY_COL in base_df.columns:
                # Get unique priorities (case-insensitive, space-tolerant)
                priority_series = base_df[PRIORITY_COL].dropna().astype(str)
                # Normalize for grouping
                priority_normalized = (
                    priority_series.str.strip()
                    .str.upper()
                    .str.replace(r"\s+", " ", regex=True)
                )
                # Get unique normalized values and their first original representation
                unique_norm = priority_normalized.unique()
                priority_map = {}
                for norm_val in unique_norm:
                    orig_val = priority_series[priority_normalized == norm_val].iloc[0]
                    priority_map[orig_val] = norm_val
                priority_options = ["ALL"] + sorted(
                    priority_map.keys(), key=lambda x: str(x).upper()
                )
                selected_priority = st.selectbox("Business Priority", priority_options)
            else:
                selected_priority = "ALL"
        if selected_priority != "ALL":
            # Case-insensitive, space-tolerant comparison
            base_df_norm = get_normalized_df(base_df)
            selected_norm = " ".join(str(selected_priority).strip().upper().split())
            base_df_norm_priority = (
                base_df_norm[PRIORITY_COL]
                .astype(str)
                .str.strip()
                .str.upper()
                .str.replace(r"\s+", " ", regex=True)
            )
            base_df = base_df[base_df_norm_priority == selected_norm]
        with f3:
            search_text = st.text_input("🔍 Global Search (case-insensitive)")
    # ==================================================
    # SEARCH (case-insensitive, space-tolerant)
    # ==================================================
    if search_text:
        # Normalize search text: remove extra spaces, case-insensitive
        search_normalized = " ".join(str(search_text).strip().upper().split())

        # Search across all columns with case-insensitive, space-tolerant matching
        def search_match(row):
            row_str = " ".join(
                [str(val).strip().upper() for val in row.values if pd.notna(val)]
            )
            row_str = " ".join(row_str.split())  # Normalize spaces
            return search_normalized in row_str

        base_df = base_df[base_df.apply(search_match, axis=1)]
    st.caption(f"📌 Showing **{len(base_df)}** records")
    # ==================================================
    # FINAL TABLE
    # ==================================================
    # Create a copy
    base_df_display = base_df.copy()
    base_df_display.head()
    # Exclude sno and needed_by columns from display
    excluded_cols = ["sno", "needed_by", "needed by", "neededby", "needed-by"]
    excluded_cols_normalized = [clean_col(c) for c in excluded_cols]
    base_df_display = base_df_display.loc[
        :, ~base_df_display.columns.str.lower().isin(excluded_cols_normalized)
    ]
    # Restore original column names for display
    display_columns = {}
    for col in base_df_display.columns:
        if col in column_name_mapping:
            display_columns[col] = column_name_mapping[col]
        else:
            # Try to find original name by reverse lookup
            for orig_name in original_column_names:
                if clean_col(orig_name) == col:
                    display_columns[col] = orig_name
                    break
    base_df_display_renamed = base_df_display.rename(columns=display_columns)
    base_df_display_renamed = base_df_display_renamed.head(15)

    # Colorful table styling function
    def style_colorful_table(df):
        def format_cell(x):
            if pd.isna(x):
                return x
            # Try to format as number if it's numeric
            try:
                num = float(x)
                return format_large_number(num)
            except (ValueError, TypeError):
                return x

        return df.style.format(format_cell).set_table_styles(
            [
                {
                    "selector": "thead th",
                    "props": [
                        (
                            "background",
                            "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
                        ),
                        ("color", "white"),
                        ("font-weight", "bold"),
                        ("font-size", "14px"),
                        ("text-align", "center"),
                        ("padding", "12px 8px"),
                        ("border", "none"),
                    ],
                },
                {
                    "selector": "tbody tr:nth-child(odd)",
                    "props": [
                        ("background-color", "#f8f9fa"),
                    ],
                },
                {
                    "selector": "tbody tr:nth-child(even)",
                    "props": [
                        ("background-color", "#ffffff"),
                    ],
                },
                {
                    "selector": "tbody tr:hover",
                    "props": [
                        ("background-color", "#e3f2fd"),
                        ("transition", "background-color 0.2s"),
                    ],
                },
                {
                    "selector": "tbody td",
                    "props": [
                        ("padding", "10px 8px"),
                        ("text-align", "center"),
                        ("border-bottom", "1px solid #e0e0e0"),
                        ("font-size", "13px"),
                    ],
                },
                {
                    "selector": "tbody tr:last-child td",
                    "props": [("border-bottom", "none")],
                },
            ]
        )

    # Calculate exact height needed (no extra rows)
    num_rows = len(base_df_display_renamed)
    header_height = 50
    row_height = 45
    exact_height = min(header_height + (num_rows * row_height) + 10, 800)  # Max 800px
    # Display dataframe with colorful styling, no empty rows
    df15 = base_df_display_renamed.head(15)

    st.dataframe(
        style_colorful_table(df15),
        use_container_width=True,
        hide_index=True,
        height=38 + 35 * len(df15),
    )


   
    # from reportlab.lib.pagesizes import letter
    # from reportlab.pdfgen import canvas
    # import io

    # # Function to generate PDF in memory
    # def create_kpi_pdf(total, open_so, closed_so, cancelled_so):
    #     buffer = io.BytesIO()
    #     c = canvas.Canvas(buffer, pagesize=letter)
    #     width, height = letter

    #     c.setFont("Helvetica-Bold", 16)
    #     c.drawString(200, height - 50, "Sales Order KPI Report")

    #     c.setFont("Helvetica", 12)
    #     c.drawString(100, height - 100, f"📊 Total SOs: {total}")
    #     c.drawString(100, height - 130, f"🟢 Open SOs: {open_so}")
    #     c.drawString(100, height - 160, f"🔵 Closed SOs: {closed_so}")
    #     c.drawString(100, height - 190, f"⚪ Cancelled SOs: {cancelled_so}")

    #     c.showPage()
    #     c.save()
    #     buffer.seek(0)
    #     return buffer

    # # Pull KPI values directly from your existing dataframes
    # total_sos = len(df)
    # open_sos = len(open_df)
    # closed_sos = len(closed_df)
    # cancelled_sos = len(cancel_df)

    # # Create PDF
    # pdf_file = create_kpi_pdf(total_sos, open_sos, closed_sos, cancelled_sos)

    # # Streamlit download button
    # st.download_button(
    #     label="📥 Download KPI Report (PDF)",
    #     data=pdf_file,
    #     file_name="SO_KPI_Report.pdf",
    #     mime="application/pdf",
    # )


    #Run 2 - without styles
    # import streamlit as st
    # from reportlab.lib.pagesizes import letter
    # from reportlab.pdfgen import canvas
    # import io

    # # Function to generate PDF in memory
    # def create_kpi_pdf(total, open_so, closed_so, cancelled_so,
    #                 need_attention, escalations, critical, non_critical,
    #                 mfr_no, mfr_yes,
    #                 fulfilled, fulfilled_new, fulfilled_rep,
    #                 hiring, hiring_new, hiring_rep,
    #                 proactive, proactive_new, proactive_rep,
    #                 revenue_loss, so_count, needed_by):
    #     buffer = io.BytesIO()
    #     c = canvas.Canvas(buffer, pagesize=letter)
    #     width, height = letter

    #     # Title
    #     c.setFont("Helvetica-Bold", 16)
    #     c.drawString(180, height - 50, "Sales Order KPI Report")

    #     # Summary KPIs
    #     c.setFont("Helvetica-Bold", 14)
    #     c.drawString(100, height - 100, "Summary KPIs")
    #     c.setFont("Helvetica", 12)
    #     c.drawString(120, height - 120, f"📊 Total SOs: {total}")
    #     c.drawString(120, height - 140, f"🟢 Open SOs: {open_so}")
    #     c.drawString(120, height - 160, f"🔵 Closed SOs: {closed_so}")
    #     c.drawString(120, height - 180, f"⚪ Cancelled SOs: {cancelled_so}")

    #     # Open Requirements KPIs
    #     c.setFont("Helvetica-Bold", 14)
    #     c.drawString(100, height - 220, "Open Requirements KPIs")
    #     c.setFont("Helvetica", 12)
    #     y = height - 240
    #     line_gap = 20

    #     c.drawString(120, y, f"Needed Attention: {need_attention}"); y -= line_gap
    #     c.drawString(120, y, f"Escalations Count: {escalations}"); y -= line_gap
    #     c.drawString(120, y, f"Critical Count: {critical}"); y -= line_gap
    #     c.drawString(120, y, f"Non-Critical Count: {non_critical}"); y -= line_gap
    #     c.drawString(120, y, f"MFR-NO: {mfr_no}"); y -= line_gap
    #     c.drawString(120, y, f"MFR-YES: {mfr_yes}"); y -= line_gap
    #     c.drawString(120, y, f"Fulfilled: {fulfilled} (New: {fulfilled_new} | Rep: {fulfilled_rep})"); y -= line_gap
    #     c.drawString(120, y, f"Hiring: {hiring} (New: {hiring_new} | Rep: {hiring_rep})"); y -= line_gap
    #     c.drawString(120, y, f"Proactive: {proactive} (New: {proactive_new} | Rep: {proactive_rep})"); y -= line_gap
    #     c.drawString(120, y, f"Revenue Loss: {revenue_loss}"); y -= line_gap
    #     c.drawString(120, y, f"SO Count: {so_count}"); y -= line_gap
    #     c.drawString(120, y, f"Needed By: {needed_by}"); y -= line_gap

    #     c.showPage()
    #     c.save()
    #     buffer.seek(0)
    #     return buffer

    # # Pull KPI values directly from your app variables
    # pdf_file = create_kpi_pdf(
    #     len(df), len(open_df), len(closed_df), len(cancel_df),
    #     need_attention_open, other_escalations_open, critical_open_count, non_critical_open_count,
    #     mfr_no_open, mfr_yes_open,
    #     fulfilled_open, fulfilled_open_new, fulfilled_open_replacement,
    #     hiring_open, hiring_open_new, hiring_open_replacement,
    #     proactive_open, proactive_open_new, proactive_open_replacement,
    #     loss_yes_open, so_count, needed_by_open
    # )

    # # Streamlit download button
    # st.download_button(
    #     label="📥 Download KPI Report (PDF)",
    #     data=pdf_file,
    #     file_name="SO_KPI_Report.pdf",
    #     mime="application/pdf",
    # )

   
   
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    import io

    def create_kpi_pdf(total, open_so, closed_so, cancelled_so,
                    need_attention, escalations, critical, non_critical,
                    mfr_no, mfr_yes,
                    fulfilled, fulfilled_new, fulfilled_rep,
                    hiring, hiring_new, hiring_rep,
                    proactive, proactive_new, proactive_rep,
                    revenue_loss, so_count, needed_by):

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter,
                                leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
        elements = []
        styles = getSampleStyleSheet()

        # Custom centered styles
        title_style = ParagraphStyle(
            name="CenteredTitle",
            parent=styles["Title"],
            alignment=1,  # CENTER
            textColor=colors.HexColor("#1F4B99")
        )
        section_style = ParagraphStyle(
            name="CenteredSection",
            parent=styles["Heading2"],
            alignment=1,  # CENTER
            textColor=colors.HexColor("#1F4B99")
        )

        # Report title
        elements.append(Paragraph("📊 Sales Order KPI Report", title_style))
        elements.append(Spacer(1, 12))

        # Summary KPIs table (with header row)
        summary_data = [
            ["Metric", "Value"],
            ["Total SOs", f"{total}"],
            ["Open SOs", f"{open_so}"],
            ["Closed SOs", f"{closed_so}"],
            ["Cancelled SOs", f"{cancelled_so}"],
        ]
        summary_table = Table(summary_data, colWidths=[240, 240])
        summary_table.setStyle(TableStyle([
            # Header styling
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4F81BD")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,0), 'MIDDLE'),

            # Body styling
            ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.beige]),
            ('FONTNAME', (0,1), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (1,1), (1,-1), 'Helvetica'),
            ('ALIGN', (0,1), (0,-1), 'LEFT'),
            ('ALIGN', (1,1), (1,-1), 'RIGHT'),
            ('VALIGN', (0,1), (-1,-1), 'MIDDLE'),

            # Grid and padding
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))

        elements.append(Paragraph("Summary KPIs", section_style))
        elements.append(Spacer(1, 6))
        elements.append(summary_table)
        elements.append(Spacer(1, 18))

        # Open Requirements KPIs table (with header row)
        open_data = [
            ["Metric", "Value"],
            ["Needed Attention", f"{need_attention}"],
            ["Escalations Count", f"{escalations}"],
            ["Critical Count", f"{critical}"],
            ["Non-Critical Count", f"{non_critical}"],
            ["MFR-NO", f"{mfr_no}"],
            ["MFR-YES", f"{mfr_yes}"],
            ["Fulfilled", f"{fulfilled} (New: {fulfilled_new} | Rep: {fulfilled_rep})"],
            ["Hiring", f"{hiring} (New: {hiring_new} | Rep: {hiring_rep})"],
            ["Proactive", f"{proactive} (New: {proactive_new} | Rep: {proactive_rep})"],
            ["Revenue Loss", f"{revenue_loss}"],
            ["SO Count", f"{so_count}"],
            ["Needed By", f"{needed_by}"],
        ]
        open_table = Table(open_data, colWidths=[240, 240])
        open_table.setStyle(TableStyle([
            # Header styling
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1F4B99")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,0), 'MIDDLE'),

            # Body styling
            ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.beige]),
            ('FONTNAME', (0,1), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (1,1), (1,-1), 'Helvetica'),
            ('ALIGN', (0,1), (0,-1), 'LEFT'),
            ('ALIGN', (1,1), (1,-1), 'RIGHT'),
            ('VALIGN', (0,1), (-1,-1), 'MIDDLE'),

            # Grid and padding
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))

        elements.append(Paragraph("Open Requirements KPIs", section_style))
        elements.append(Spacer(1, 6))
        elements.append(open_table)

        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        return buffer

    # Pull KPI values directly from your app variables
    pdf_file = create_kpi_pdf(
        len(df), len(open_df), len(closed_df), len(cancel_df),
        need_attention_open, other_escalations_open, critical_open_count, non_critical_open_count,
        mfr_no_open, mfr_yes_open,
        fulfilled_open, fulfilled_open_new, fulfilled_open_replacement,
        hiring_open, hiring_open_new, hiring_open_replacement,
        proactive_open, proactive_open_new, proactive_open_replacement,
        loss_yes_open, so_count, needed_by_open
    )

    # Streamlit download button
    st.download_button(
        label="📥 Download KPI Report (PDF)",
        data=pdf_file,
        file_name="SO_KPI_Report.pdf",
        mime="application/pdf",
    )



   
   

   



with nav_tabs[1]:
    # ---- Your old: with tab2: content goes here ----
    st.session_state["active_tab"] = "📈 Visual Insights"
    # Refresh data from master_df
    master_df = st.session_state["master_df"]
    open_df, closed_df, cancel_df = get_status_dfs(master_df)
    df = master_df
    # ==================================================
    # PREMIUM STYLES
    # ==================================================
    st.markdown(
        """
        <style>
        .kpi-mini {
            padding: 14px;
            border-radius: 12px;
            text-align: center;
            font-weight: 800;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            color: white;
        }
        .kpi-mini-val {
            font-size: 56px;
            margin-top: 12px;
            font-weight: 900;
            line-height: 1.2;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    # ==================================================
    # HEADER
    # ==================================================
    st.markdown(
        """
        <div style="
            background: linear-gradient(90deg,#1d3557,#457b9d);
            padding:14px;
            border-radius:12px;
            color:white;
            font-size:28px;
            font-weight:900;
            margin-bottom:10px;">
            📈 Visual Insights — Key Metrics & Trends
        </div>
        """,
        unsafe_allow_html=True,
    )
    # ==================================================
    # KPI STRIP (VISUAL)
    # ==================================================
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            f"""
        <div class="kpi-mini" style="background:#457b9d">
                Total SOs
        <div class="kpi-mini-val">{len(df)}</div>
        </div>
            """,
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f"""
        <div class="kpi-mini" style="background:#2a9d8f">
                Open SOs
        <div class="kpi-mini-val">{len(open_df)}</div>
        </div>
            """,
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f"""
        <div class="kpi-mini" style="background:#4361ee">
                Closed SOs
        <div class="kpi-mini-val">{len(closed_df)}</div>
        </div>
            """,
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f"""
        <div class="kpi-mini" style="background:#6c757d">
                Cancelled SOs
        <div class="kpi-mini-val">{len(cancel_df)}</div>
        </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("---")
    # ==================================================
    # COLUMN MAPPING
    # ==================================================
    SKILL_COL = ensure_col("skill", "skill")
    LOCATION_COL = ensure_col("location", "location")
    PRIORITY_COL = ensure_col("priority", "priority", "priorty")
    ACT_PIPE_COL = ensure_col("actual_pipeline", "actualpipeline", "actual pipeline")
    DATE_COL = ensure_col("requirement_received_date", "requirement received date")
    TOWER_COL = ensure_col("tower", "tower")
    if DATE_COL in df.columns:
        df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
        st.session_state["master_df"][DATE_COL] = df[DATE_COL]
    data = df.copy()
    # Find quantity column (case-insensitive)
    QUANTITY_COL = None
    for col in data.columns:
        if "quantity" in col.lower() or "qty" in col.lower():
            QUANTITY_COL = col
            break

    def normalize_value_for_counting(val):
        """Normalize value for case-insensitive and space-tolerant counting"""
        if pd.isna(val) or val == "":
            return None
        val_str = str(val).strip()
        if val_str.upper() in ["NAN", "NONE", ""]:
            return None
        # Normalize: uppercase, remove extra spaces
        return " ".join(val_str.upper().split())

    def nonempty_counts(c, top=10, nan_label="NaN"):
        if c not in data.columns:
            return pd.DataFrame(columns=[c, "count"])
        s_original = (
            data[c]
            .replace("", nan_label)
            .replace(["NAN", "NONE", None, np.nan], nan_label)
        )
        s_normalized = s_original.apply(
            lambda x: nan_label if x == nan_label else normalize_value_for_counting(x)
        )
        if QUANTITY_COL and QUANTITY_COL in data.columns:
            quantities = pd.to_numeric(data[QUANTITY_COL], errors="coerce").fillna(1)
        else:
            quantities = pd.Series(1, index=data.index)
        # ❌ DO NOT DROP NaN
        grouped = pd.DataFrame(
            {"normalized": s_normalized, "original": s_original, "quantity": quantities}
        )
        quantity_sums = (
            grouped.groupby("normalized")["quantity"]
            .sum()
            .sort_values(ascending=False)
            .head(top)
        )
        result_data = []
        for norm_val in quantity_sums.index:
            orig_values = grouped[grouped["normalized"] == norm_val]["original"]
            display_val = orig_values.value_counts().index[0]
            result_data.append({c: display_val, "count": int(quantity_sums[norm_val])})
        return pd.DataFrame(result_data)

    # ==================================================
    # SUB-TABS
    # ==================================================
    tabs = st.tabs(
        [
            "🧠 Skills & Revloss",
            "📌 Request & Priority",
            "🏗 Tower & Escalation",  # was Tab 4 → now Tab 3
            "📊 Demand & Pipeline",  # was Tab 3 → now Tab 4
            "🌍 MFR & Location",
        ]
    )
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1, st.container(border=True):
            st.markdown("#### Skills")
            skills = nonempty_counts(SKILL_COL, top=None)
            if not skills.empty:
                # Format count column with commas
                skills_formatted = skills.copy()
                if "count" in skills_formatted.columns:
                    skills_formatted["count"] = skills_formatted["count"].apply(
                        format_large_number
                    )
                st.dataframe(
                    skills_formatted.style.hide(axis="index"),
                    use_container_width=True,
                    hide_index="true",
                )
        # ===== Rev Loss Split =====
        with c2, st.container(border=True):
            st.markdown("#### Rev Loss")
            rev_loss = nonempty_counts(LOSS_COL, 10)
            if not rev_loss.empty:
                chart_loss = (
                    alt.Chart(rev_loss)
                    .mark_bar()
                    .encode(
                        x=alt.X("count:Q", title="Count", axis=alt.Axis(format=",s")),
                        y=alt.Y(f"{LOSS_COL}:N", sort="-x", title="Loss Status"),
                        color=alt.Color(
                            f"{LOSS_COL}:N",
                            scale=alt.Scale(scheme="set2"),  # 🌈 colourful palette
                        ),
                        tooltip=[
                            LOSS_COL,
                            alt.Tooltip("count:Q", format=",", title="Count"),
                        ],
                    )
                    .properties(height=300)
                )
                st.altair_chart(chart_loss, use_container_width=True)
    # ==================================================
    # TAB 2 — Request & Priority
    # ==================================================
    with tabs[1]:
        c1, c2 = st.columns(2)
        # ===== Request Type Distribution =====
        with c1, st.container(border=True):
            st.markdown("#### Request Type Distribution")
            req = nonempty_counts(REQ_COL, 10)
            if not req.empty:
                chart_req = (
                    alt.Chart(req)
                    .mark_bar()
                    .encode(
                        x=alt.X(f"{REQ_COL}:N", sort="-y", title="Request Type"),
                        y=alt.Y("count:Q", title="Count", axis=alt.Axis(format=",s")),
                        color=alt.Color(
                            f"{REQ_COL}:N", scale=alt.Scale(scheme="tableau10")
                        ),
                        tooltip=[
                            REQ_COL,
                            alt.Tooltip("count:Q", format=",", title="Count"),
                        ],
                    )
                )
                st.altair_chart(chart_req, use_container_width=True)
            else:
                st.info("No data available for Request Type")
        # ===== Priority Distribution =====
        with c2, st.container(border=True):
            st.markdown("#### Priority Distribution")
            pri = nonempty_counts(PRIORITY_COL, 10)
            if not pri.empty:
                chart_pri = (
                    alt.Chart(pri)
                    .mark_bar()
                    .encode(
                        x=alt.X(f"{PRIORITY_COL}:N", sort="-y", title="Priority"),
                        y=alt.Y("count:Q", title="Count", axis=alt.Axis(format=",s")),
                        color=alt.Color(
                            f"{PRIORITY_COL}:N", scale=alt.Scale(scheme="dark2")
                        ),
                        tooltip=[
                            PRIORITY_COL,
                            alt.Tooltip("count:Q", format=",", title="Count"),
                        ],
                    )
                )
                st.altair_chart(chart_pri, use_container_width=True)
            else:
                st.info("No data available for Priority")
    # ==================================================
    # TAB 3 — Tower & Escalation
    # ==================================================
    with tabs[2]:
        c1, c2 = st.columns(2)
        # ===== Tower Distribution =====
        with c1, st.container(border=True):
            st.markdown("#### Tower Distribution")
            tower = nonempty_counts(TOWER_COL, 20)
            if not tower.empty:
                chart_tower = (
                    alt.Chart(tower)
                    .mark_bar()
                    .encode(
                        x=alt.X("count:Q", title="Count", axis=alt.Axis(format=",s")),
                        y=alt.Y(f"{TOWER_COL}:N", sort="-x", title="Tower"),
                        color=alt.Color(
                            f"{TOWER_COL}:N",
                            scale=alt.Scale(scheme="tableau10"),  # 🔥 bold colors
                        ),
                        tooltip=[
                            TOWER_COL,
                            alt.Tooltip("count:Q", format=",", title="Count"),
                        ],
                    )
                    .properties(height=400, width=500)  # ⬆️ bigger chart
                )
                st.altair_chart(chart_tower, use_container_width=True)
        # ===== Need Attention / Escalation =====
        with c2, st.container(border=True):
            st.markdown("#### Need Attention / Escalation")
            esc = nonempty_counts(ESC_COL, 10)
            if not esc.empty:
                chart_esc = (
                    alt.Chart(esc)
                    .mark_bar()
                    .encode(
                        x=alt.X("count:Q", title="Count", axis=alt.Axis(format=",s")),
                        y=alt.Y(f"{ESC_COL}:N", sort="-x", title="Escalation Status"),
                        color=alt.Color(
                            f"{ESC_COL}:N",
                            scale=alt.Scale(scheme="dark2"),  # 🔥 deep contrast
                        ),
                        tooltip=[
                            ESC_COL,
                            alt.Tooltip("count:Q", format=",", title="Count"),
                        ],
                    )
                    .properties(height=400, width=500)  # ⬆️ bigger chart
                )
                st.altair_chart(chart_esc, use_container_width=True)
    # ==================================================
    # TAB 4 — Tower & Escalation
    # ==================================================
    with tabs[3]:
        DEMAND_COL = ensure_col("demand_flag", "hiring", "hiring/can be fulfill")
        open_mask = data[STATUS_COL] == "OPEN"

        def bucket(v):
            v = str(v).upper()
            if "HIRING" in v:
                return "Hiring"
            if "PROACT" in v:
                return "Proactive"
            if "FULFIL" in v or "INTERN" in v:
                return "Internal"
            return "Other"

        demand_open = (
            data.loc[open_mask, DEMAND_COL]
            .fillna("")
            .apply(bucket)
            .value_counts()
            .rename_axis("demand")
            .reset_index(name="count")
        )
        c1, c2 = st.columns(2)
        # ===== Open Demand Mix =====
        with c1, st.container(border=True):
            st.markdown("#### Demand Distribution")

            def bucket(v):
                v = str(v).upper()
                if "HIRING" in v:
                    return "Hiring"
                if "PROACT" in v:
                    return "Proactive"
                if "FULFIL" in v or "INTERN" in v:
                    return "Internal"
                return "Other"

            # Use full dataset, not just open SOs
            demand_all = (
                data[DEMAND_COL]
                .fillna("")
                .apply(bucket)
                .value_counts()
                .rename_axis("demand")
                .reset_index(name="count")
            )
            if not demand_all.empty:
                st.altair_chart(
                    alt.Chart(demand_all)
                    .mark_bar()
                    .encode(
                        x=alt.X("demand:N", sort="-y", title="Demand Type"),
                        y=alt.Y("count:Q", title="Count", axis=alt.Axis(format=",s")),
                        color=alt.Color(
                            "demand:N",
                            scale=alt.Scale(scheme="tableau10"),  # 🔥 bold colors
                        ),
                        tooltip=[
                            "demand",
                            alt.Tooltip("count:Q", format=",", title="Count"),
                        ],
                    )
                    .properties(height=400),
                    use_container_width=True,
                )
        # ===== Actual Pipeline =====
        with c2, st.container(border=True):
            st.markdown("#### Actual Pipeline")
            ap = nonempty_counts(ACT_PIPE_COL, 10)
            if not ap.empty:
                st.altair_chart(
                    alt.Chart(ap)
                    .mark_bar()
                    .encode(
                        x=alt.X(f"{ACT_PIPE_COL}:N", sort="-y", title="Pipeline Type"),
                        y=alt.Y("count:Q", title="Count", axis=alt.Axis(format=",s")),
                        color=alt.Color(
                            f"{ACT_PIPE_COL}:N",
                            scale=alt.Scale(scheme="dark2"),  # 🔥 deep colors
                        ),
                        tooltip=[
                            ACT_PIPE_COL,
                            alt.Tooltip("count:Q", format=",", title="Count"),
                        ],
                    ),
                    use_container_width=True,
                )
    # ==================================================
    # TAB 5 — MFR & Location
    # ==================================================
    with tabs[4]:
        c1, c2 = st.columns(2)
        # ===== MFR Split =====
        with c1, st.container(border=True):
            st.markdown("#### MFR Split")
            mfr = nonempty_counts(MFR_COL, 10)
            if not mfr.empty:
                st.altair_chart(
                    alt.Chart(mfr)
                    .mark_arc(innerRadius=45)
                    .encode(theta="count:Q", color=f"{MFR_COL}:N"),
                    use_container_width=True,
                )
        # ===== Top Locations =====
        with c2, st.container(border=True):
            st.markdown("#### Top Locations")
            loc = nonempty_counts(LOCATION_COL, 10)
            if not loc.empty:
                st.altair_chart(
                    alt.Chart(loc)
                    .mark_arc(innerRadius=45)
                    .encode(theta="count:Q", color=f"{LOCATION_COL}:N"),
                    use_container_width=True,
                )
with nav_tabs[2]:
    # ---- Your old: with tab3: content goes here ----
    st.session_state["active_tab"] = "➕ Manage Resource"
    st.markdown(
        """
        <style>
        .manage-header {
            background: linear-gradient(90deg, #1d3557, #457b9d);
            padding: 16px 24px;
            border-radius: 12px;
            color: white;
            font-size: 24px;
            font-weight: 900;
            margin-bottom: 20px;
            text-align: center;
        }
        .stDataEditor th {
            background-color: #f8f9fa !important;
            font-weight: 700 !important;
            text-align: center !important;
        }
        .stDataEditor td {
            text-align: center !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="manage-header">➕ Manage Resource Data</div>',
        unsafe_allow_html=True,
    )
    # ==================================================
    # INIT MANAGE DF (ONE TIME ONLY)
    # ==================================================
    master_df = st.session_state["master_df"]
    if "manage_base_df" not in st.session_state:
        base = master_df.copy()
        base.insert(0, "DELETE", False)
        st.session_state["manage_base_df"] = base
    # ==================================================
    # INSTRUCTIONS
    # ==================================================
    st.info(
        "💡 **Instructions**: Edit cells freely, add rows using ➕, "
        "mark rows using DELETE. Changes apply only after Save/Delete."
    )
    # ==================================================
    # DATA EDITOR (BUFFERED EDITS)
    # ==================================================
    edited_df = st.data_editor(
        st.session_state["manage_base_df"],
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="manage_editor",
        column_config={
            "DELETE": st.column_config.CheckboxColumn(
                "DELETE",
                help="Mark rows for deletion",
                default=False,
            )
        },
    )
    # Buffer edits (do not commit yet)
    st.session_state["manage_pending_df"] = edited_df
    # ==================================================
    # ACTION BUTTONS
    # ==================================================
    col1, col2 = st.columns([1, 4])
    # ---------- DELETE ----------
    with col1:
        if st.button("🗑️ Delete Marked Rows", use_container_width=True):
            rows_to_delete = st.session_state["manage_pending_df"][
                st.session_state["manage_pending_df"]["DELETE"]
            ]
            if rows_to_delete.empty:
                st.warning("⚠️ No rows marked for deletion.")
            else:
                cleaned_df = (
                    st.session_state["manage_pending_df"]
                    .loc[~st.session_state["manage_pending_df"]["DELETE"]]
                    .drop(columns="DELETE")
                    .reset_index(drop=True)
                )
                st.session_state["master_df"] = cleaned_df
                st.session_state["master_df_normalized"] = get_normalized_df(cleaned_df)
                refreshed = cleaned_df.copy()
                refreshed.insert(0, "DELETE", False)
                st.session_state["manage_base_df"] = refreshed
                st.success(f"✅ Deleted {len(rows_to_delete)} row(s).")
    # ---------- SAVE ----------
    with col2:
        if st.button(
            "💾 Save All Changes to Excel", type="primary", use_container_width=True
        ):
            try:
                with st.spinner("Saving to SharePoint Excel..."):
                    final_df = (
                        st.session_state["manage_pending_df"]
                        .drop(columns="DELETE")
                        .fillna("")
                    )
                    save_df_to_sharepoint(final_df)
                st.session_state["master_df"] = final_df
                st.session_state["master_df_normalized"] = get_normalized_df(final_df)
                refreshed = final_df.copy()
                refreshed.insert(0, "DELETE", False)
                st.session_state["manage_base_df"] = refreshed
                st.cache_data.clear()
                st.success("✅ Excel updated successfully!")
            except Exception as e:
                st.error(f"❌ Failed to save changes: {e}")
