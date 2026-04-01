import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from io import StringIO, BytesIO
import requests
import os
import openpyxl

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Zoom Meeting Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Dark Theme Colors ─────────────────────────────────────────────────────────
BG_DARK      = "#0F1117"
BG_CARD      = "#1A1D27"
BG_CARD2     = "#1E2130"
BORDER       = "#2A2D3E"
TEXT_PRIMARY = "#E8EAF6"
TEXT_MUTED   = "#7B8DB0"
ACCENT_BLUE  = "#4D9FFF"
ACCENT_GREEN = "#43D9AD"
ACCENT_AMBER = "#FFB938"
ACCENT_RED   = "#FF6B6B"
ACCENT_PURP  = "#B197FC"
ACCENT_TEAL  = "#38D9C0"
ACCENT_PINK  = "#FF79C6"
ACCENT_CYAN  = "#8BE9FD"

PIE_DISCLAIMER = [ACCENT_GREEN, ACCENT_AMBER, ACCENT_RED]
PIE_GUEST      = [ACCENT_BLUE,  ACCENT_PURP]
PIE_WAITING    = [ACCENT_GREEN, ACCENT_RED]
BAR_BLUES      = ["#1565C0","#1976D2","#2196F3","#42A5F5","#90CAF9","#BBDEFB"]
PLATFORM_COLORS = [ACCENT_BLUE, ACCENT_GREEN, ACCENT_AMBER, ACCENT_PURP,
                   ACCENT_TEAL, ACCENT_PINK, ACCENT_CYAN, ACCENT_RED]

# ── Plotly dark layout template ───────────────────────────────────────────────
DARK_LAYOUT = dict(
    paper_bgcolor=BG_CARD,
    plot_bgcolor=BG_CARD,
    font=dict(color=TEXT_PRIMARY, size=12, family="Inter, sans-serif"),
    title_font=dict(size=14, color=TEXT_PRIMARY),
    xaxis=dict(
        gridcolor=BORDER, linecolor=BORDER,
        tickcolor=TEXT_MUTED, tickfont=dict(color=TEXT_MUTED),
        title_font=dict(color=TEXT_MUTED),
    ),
    yaxis=dict(
        gridcolor=BORDER, linecolor=BORDER,
        tickcolor=TEXT_MUTED, tickfont=dict(color=TEXT_MUTED),
        title_font=dict(color=TEXT_MUTED),
    ),
    margin=dict(t=50, b=35, l=40, r=20),
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');

  .stApp, .stApp > div {{
      background-color: {BG_DARK} !important;
      font-family: 'Inter', sans-serif;
  }}
  .block-container {{
      padding: 1.5rem 2rem 3rem 2rem;
      background-color: {BG_DARK};
  }}
  section[data-testid="stSidebar"] {{
      background-color: #13151F !important;
      border-right: 1px solid {BORDER};
  }}
  section[data-testid="stSidebar"] * {{
      color: {TEXT_PRIMARY} !important;
  }}
  section[data-testid="stSidebar"] .stFileUploader label,
  section[data-testid="stSidebar"] .stTextInput label,
  section[data-testid="stSidebar"] .stRadio label {{
      color: {TEXT_MUTED} !important;
  }}
  h1, h2, h3 {{ color: {TEXT_PRIMARY} !important; }}
  p, span, label {{ color: {TEXT_MUTED}; }}
  .banner {{
      background: linear-gradient(135deg, #1A2980, #26D0CE22);
      border: 1px solid {ACCENT_BLUE}44;
      border-radius: 14px;
      padding: 1.3rem 1.8rem;
      color: white;
      margin-bottom: 1.4rem;
  }}
  .banner .b-title {{
      font-size: 1.45rem; font-weight: 800;
      color: {TEXT_PRIMARY}; margin-bottom: 8px;
  }}
  .banner .b-tag {{
      display: inline-block;
      background: rgba(77,159,255,0.18);
      border: 1px solid {ACCENT_BLUE}55;
      border-radius: 20px;
      padding: 3px 14px;
      margin: 3px 5px 3px 0;
      font-size: 0.8rem;
      font-weight: 600;
      color: {ACCENT_BLUE};
  }}
  .banner .b-sub {{
      font-size: 0.85rem;
      color: {TEXT_MUTED};
      margin-top: 8px;
  }}
  .kpi-wrap {{ display: flex; gap: 14px; margin-bottom: 1.4rem; flex-wrap: wrap; }}
  .kpi-card {{
      background: {BG_CARD};
      border: 1px solid {BORDER};
      border-radius: 12px;
      padding: 1.1rem 1.4rem;
      flex: 1; min-width: 150px;
      position: relative;
      overflow: hidden;
  }}
  .kpi-card::before {{
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 3px;
  }}
  .kpi-card.blue::before   {{ background: {ACCENT_BLUE}; }}
  .kpi-card.green::before  {{ background: {ACCENT_GREEN}; }}
  .kpi-card.amber::before  {{ background: {ACCENT_AMBER}; }}
  .kpi-card.purple::before {{ background: {ACCENT_PURP}; }}
  .kpi-card.teal::before   {{ background: {ACCENT_TEAL}; }}
  .kpi-card.pink::before   {{ background: {ACCENT_PINK}; }}
  .kpi-card h4 {{
      margin: 0 0 6px 0;
      font-size: 0.7rem;
      text-transform: uppercase;
      letter-spacing: .08em;
      color: {TEXT_MUTED};
      font-weight: 700;
  }}
  .kpi-card .kv {{
      font-size: 1.85rem;
      font-weight: 800;
      color: {TEXT_PRIMARY};
      line-height: 1.1;
  }}
  .kpi-card .ksub {{
      font-size: 0.74rem;
      color: {TEXT_MUTED};
      margin-top: 4px;
  }}
  .sec-title {{
      font-size: 0.92rem;
      font-weight: 700;
      color: {TEXT_MUTED};
      margin: 1.5rem 0 0.7rem 0;
      padding-bottom: 0.35rem;
      border-bottom: 1px solid {BORDER};
      text-transform: uppercase;
      letter-spacing: .06em;
  }}
  .stats-card {{
      background: {BG_CARD};
      border: 1px solid {BORDER};
      border-radius: 12px;
      padding: 1rem 1.3rem;
      margin-bottom: 0.5rem;
  }}
  .stats-card table {{
      width: 100%;
      border-collapse: collapse;
  }}
  .stats-card th {{
      font-size: 0.72rem;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: {TEXT_MUTED};
      padding: 0 8px 10px 8px;
      font-weight: 700;
      text-align: left;
  }}
  .stats-card th:last-child {{ text-align: right; }}
  .stats-card td {{
      padding: 7px 8px;
      font-size: 0.86rem;
      color: {TEXT_PRIMARY};
      border-top: 1px solid {BORDER};
  }}
  .stats-card td:last-child {{
      text-align: right;
      color: {ACCENT_BLUE};
      font-weight: 600;
  }}
  div[data-testid="stPlotlyChart"] {{
      background: {BG_CARD};
      border: 1px solid {BORDER};
      border-radius: 12px;
      padding: 0.3rem;
      margin-bottom: 0.5rem;
  }}
  div[data-testid="stDataFrame"] {{
      background: {BG_CARD} !important;
      border: 1px solid {BORDER} !important;
      border-radius: 12px;
  }}
  .stTextInput input {{
      background: {BG_CARD2} !important;
      border: 1px solid {BORDER} !important;
      color: {TEXT_PRIMARY} !important;
      border-radius: 8px;
  }}
  .stDownloadButton button {{
      background: {ACCENT_BLUE}22 !important;
      color: {ACCENT_BLUE} !important;
      border: 1px solid {ACCENT_BLUE}55 !important;
      border-radius: 8px;
      font-weight: 600;
  }}
  .stDownloadButton button:hover {{
      background: {ACCENT_BLUE}44 !important;
  }}
  .format-badge {{
      display: inline-block;
      padding: 2px 10px;
      border-radius: 20px;
      font-size: 0.75rem;
      font-weight: 700;
      margin-left: 8px;
  }}
  .format-flat {{
      background: {ACCENT_GREEN}22;
      border: 1px solid {ACCENT_GREEN}55;
      color: {ACCENT_GREEN};
  }}
  .format-split {{
      background: {ACCENT_AMBER}22;
      border: 1px solid {ACCENT_AMBER}55;
      color: {ACCENT_AMBER};
  }}
  .tab-header {{
      background: {BG_CARD};
      border: 1px solid {BORDER};
      border-radius: 10px;
      padding: 0.6rem 1rem;
      margin-bottom: 1rem;
      font-size: 0.85rem;
      color: {TEXT_MUTED};
  }}
  .file-pill {{
      display: inline-block;
      background: {ACCENT_BLUE}18;
      border: 1px solid {ACCENT_BLUE}44;
      border-radius: 20px;
      padding: 2px 12px;
      margin: 2px 3px;
      font-size: 0.75rem;
      color: {ACCENT_BLUE};
      font-weight: 600;
  }}
  .rank-table {{
      background: {BG_CARD};
      border: 1px solid {BORDER};
      border-radius: 12px;
      padding: 1rem 1.3rem;
      margin-bottom: 0.5rem;
  }}
  .rank-table table {{ width: 100%; border-collapse: collapse; }}
  .rank-table th {{
      font-size: 0.72rem; text-transform: uppercase; letter-spacing: .06em;
      color: {TEXT_MUTED}; padding: 0 8px 10px 8px; font-weight: 700; text-align: left;
  }}
  .rank-table td {{
      padding: 8px 8px; font-size: 0.87rem; color: {TEXT_PRIMARY};
      border-top: 1px solid {BORDER};
  }}
  .rank-table .rank-num {{
      font-weight: 800; color: {ACCENT_AMBER}; min-width: 24px; display: inline-block;
  }}
  .rank-table .bar-bg {{
      background: {BORDER}; border-radius: 4px; height: 6px; width: 100%; margin-top: 4px;
  }}
  .rank-table .bar-fill {{
      background: {ACCENT_BLUE}; border-radius: 4px; height: 6px;
  }}
  .error-pill {{
      background: {ACCENT_RED}18;
      border: 1px solid {ACCENT_RED}44;
      border-radius: 8px;
      padding: 4px 12px;
      font-size: 0.8rem;
      color: {ACCENT_RED};
  }}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS (sama seperti versi sebelumnya + tambahan xlsx)
# ══════════════════════════════════════════════════════════════════════════════

def parse_dt(series):
    for fmt in [
        "%m/%d/%Y %I:%M:%S %p",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M",
    ]:
        try:
            return pd.to_datetime(series, format=fmt)
        except Exception:
            continue
    return pd.to_datetime(series, infer_datetime_format=True, errors="coerce")


def detect_sep(raw_text):
    first_line = raw_text.splitlines()[0] if raw_text.splitlines() else ""
    return ";" if first_line.count(";") > first_line.count(",") else ","


def xlsx_to_csv_string(file_bytes):
    """Konversi file XLSX ke string CSV menggunakan openpyxl."""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(",".join(
            "" if v is None else str(v).replace(",", ";")
            for v in row
        ))
    return "\n".join(rows)


def load_csv(file_or_stringio, filename="data.csv"):
    """
    Auto-detect dan load CSV/XLSX Zoom.
    Mendukung FORMAT FLAT dan FORMAT SPLIT.
    """
    # ── Baca raw ──────────────────────────────────────────────────────────────
    if hasattr(file_or_stringio, "read"):
        # Cek apakah XLSX
        fname = getattr(file_or_stringio, "name", filename)
        if fname.lower().endswith(".xlsx"):
            raw_bytes = file_or_stringio.read()
            raw = xlsx_to_csv_string(raw_bytes)
        else:
            raw = file_or_stringio.read()
            if isinstance(raw, bytes):
                raw = raw.decode("utf-8-sig")
        try:
            file_or_stringio.seek(0)
        except Exception:
            pass
    else:
        raw = file_or_stringio

    sep    = detect_sep(raw)
    df_try = pd.read_csv(StringIO(raw), sep=sep)

    PESERTA_MARKERS = [
        "Nama (nama asli)", "Name (Original Name)", "Name (original name)",
        "Waktu bergabung",  "Join Time", "Join time",
        "Waktu keluar",     "Leave Time", "Leave time",
    ]

    if any(mk in list(df_try.columns) for mk in PESERTA_MARKERS):
        return df_try, "flat"

    cols_lower = {c.lower() for c in df_try.columns}
    markers_lower = {mk.lower() for mk in PESERTA_MARKERS}
    if cols_lower & markers_lower:
        return df_try, "flat"

    lines          = raw.splitlines()
    split_line_idx = None
    for i, line in enumerate(lines):
        cells = [c.strip().strip('"').lower() for c in line.split(sep)]
        if any(mk.lower() in cells for mk in PESERTA_MARKERS):
            split_line_idx = i
            break

    if split_line_idx is None:
        return df_try, "flat"

    meeting_row = df_try.iloc[0] if len(df_try) > 0 else pd.Series(dtype=object)

    peserta_raw = "\n".join(lines[split_line_idx:])
    df_peserta  = pd.read_csv(StringIO(peserta_raw), sep=sep)
    df_peserta.dropna(how="all", inplace=True)
    df_peserta.reset_index(drop=True, inplace=True)

    DUR_PESERTA_NAMES = ["Durasi (menit)", "Duration (Minutes)", "Duration (minutes)"]
    for dur_name in DUR_PESERTA_NAMES:
        if dur_name in df_peserta.columns:
            df_peserta.rename(columns={dur_name: dur_name + ".1"}, inplace=True)
            break

    MEETING_COL_MAP = {
        "Topik":              "Topik",
        "Topic":              "Topik",
        "Host":               "Nama host",
        "Nama host":          "Nama host",
        "Host Name":          "Nama host",
        "Host name":          "Nama host",
        "Waktu mulai":        "Waktu mulai",
        "Start Time":         "Waktu mulai",
        "Start time":         "Waktu mulai",
        "Waktu berakhir":     "Waktu berakhir",
        "End Time":           "Waktu berakhir",
        "End time":           "Waktu berakhir",
        "Peserta":            "Peserta",
        "Participants":       "Peserta",
        "Durasi (menit)":     "Durasi (menit)",
        "Duration (Minutes)": "Durasi (menit)",
        "Duration (minutes)": "Durasi (menit)",
        "ID":                 "ID",
    }
    for src_col, dst_col in MEETING_COL_MAP.items():
        if src_col in df_try.columns:
            val = meeting_row.get(src_col, None)
            if val is not None and str(val).strip() not in ("", "nan"):
                df_peserta[dst_col] = val

    return df_peserta, "split"


def detect_cols(df):
    candidates = {
        "topic":        ["Topik", "Topic"],
        "host":         ["Nama host", "Host Name", "Host name", "Host"],
        "start":        ["Waktu mulai", "Start Time", "Start time"],
        "end":          ["Waktu berakhir", "End Time", "End time"],
        "participants": ["Peserta", "Participants"],
        "total_min":    ["Total menit peserta", "Total Participant Minutes", "Total participant minutes"],
        "duration_m":   ["Durasi (menit)", "Duration (Minutes)", "Duration (minutes)"],
        "name":         ["Nama (nama asli)", "Name (Original Name)", "Name (original name)", "Name"],
        "email":        ["Email"],
        "join":         ["Waktu bergabung", "Join Time", "Join time"],
        "leave":        ["Waktu keluar", "Leave Time", "Leave time"],
        "duration_p":   ["Durasi (menit).1", "Duration (Minutes).1", "Duration (minutes).1"],
        "guest":        ["Tamu", "Guest"],
        "waiting":      ["Di ruang tunggu", "In Waiting Room", "In waiting room"],
        "disclaimer":   [
            "Respons penafian rekaman",
            "Recording Disclaimer Response",
            "Recording disclaimer response",
        ],
    }
    col_lower_map = {col.lower(): col for col in df.columns}
    found = {}
    for key, opts in candidates.items():
        matched = False
        for col in opts:
            if col in df.columns:
                found[key] = col
                matched = True
                break
        if not matched:
            for col in opts:
                if col.lower() in col_lower_map:
                    found[key] = col_lower_map[col.lower()]
                    break
    return found


# ── Daftar nama lengkap pemateri ─────────────────────────────────────────────
PEMATERI_LENGKAP = [
    "Akhmad Syahrul Mubarok, M.Pd.",
    "Alfikri Oktavian Yudhistira S.IP.,M.M",
    "Annisa Nurjannah S.Pi.",
    "Ayuni Kemala Safira, S.Pd., M.Si",
    "Diah Rachmawati, S.T., M.Sc",
    "Dicky Adra Pratama, A.Md., S.IP., M.A.",
    "dr Margareth A. T. Alfares",
    "dr. Yama Sirly Putri., M.Med.Sc., Sp.PK., FISQua",
    "Dyla Aliffah Saffitri, S. Pd",
    "Imron Ahmadi",
    "Indah Rachmawati, M.Pd.",
    "Putri Aulia Sudjana S.Pd.",
    "Putri Rizki Ilahi",
    "Rahmatullah S.E., M.H.",
    "Rifania Nurizqia, S.Pd.",
    "Romi Kurniawan M.Pd.",
    "Siti Hanipah S.Hum",
    "Wahyabiyantara Permana Adi S.PWK., M.Sc",
    "Zahro Rokhmawati M.Pd.",
]

def _tokenize(s: str) -> set:
    """Pecah string jadi set token huruf kecil, buang gelar/tanda baca."""
    import re
    s = s.lower()
    # hapus gelar umum dan karakter non-alfanumerik
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    tokens = {t for t in s.split() if len(t) > 2}
    # buang token yang merupakan singkatan gelar umum
    gelar = {"mpd","msi","msc","sip","spi","spd","shum","se","mh","mm",
             "mmed","sppk","fisqua","amd","ma","st","dr","drs","sst","skm"}
    return tokens - gelar


def resolve_pemateri(nama_pendek: str) -> str:
    """
    Cocokkan nama pendek dari filename ke nama lengkap di PEMATERI_LENGKAP.
    Strategi:
      1. Substring case-insensitive langsung
      2. Token overlap — pilih nama lengkap dengan token overlap terbanyak
    Kembalikan nama lengkap jika ditemukan, atau nama_pendek asli jika tidak.
    """
    key = nama_pendek.strip().lower()
    if not key or key == "-":
        return nama_pendek

    # Strategi 1: substring langsung
    for full in PEMATERI_LENGKAP:
        if key in full.lower():
            return full

    # Strategi 2: cek apakah nama lengkap mengandung kata pertama nama pendek
    # (berguna untuk "Indah" → "Indah Rachmawati, M.Pd.")
    key_tokens = _tokenize(key)
    if not key_tokens:
        return nama_pendek

    best_match = None
    best_score = 0
    for full in PEMATERI_LENGKAP:
        full_tokens = _tokenize(full)
        overlap = len(key_tokens & full_tokens)
        if overlap > best_score:
            best_score = overlap
            best_match = full

    # Minimal 1 token harus cocok untuk dianggap match
    if best_score >= 1 and best_match:
        return best_match

    return nama_pendek   # fallback: kembalikan apa adanya


def normalize_platform(raw: str) -> str:
    """
    Ambil kata pertama saja, lalu jadikan UPPERCASE agar
    'JADIASN', 'jadiasn', 'JadiASN' → 'JADIASN' (konsisten).
    """
    word = raw.strip().split()[0] if raw.strip() else raw
    return word.strip().upper()


def parse_filename(filename):
    name  = os.path.splitext(filename)[0]
    name  = name.replace(" ", "_")
    parts = [p for p in name.split("_") if p.strip()]
    raw_platform = parts[0] if len(parts) > 0 else "-"
    platform     = normalize_platform(raw_platform)
    raw_pemateri = parts[-1].strip().replace("-", " ") if len(parts) > 2 else "-"
    pemateri     = resolve_pemateri(raw_pemateri)
    materi       = " ".join(p.strip() for p in parts[1:-1]).replace("-", " ") \
                   if len(parts) > 2 else (parts[1].strip() if len(parts) > 1 else "-")
    return platform, materi, pemateri


def dark_chart(fig, height=320, title=""):
    layout = dict(DARK_LAYOUT)
    layout["height"] = height
    if title:
        layout["title"] = dict(text=title, font=dict(size=14, color=TEXT_PRIMARY))
    fig.update_layout(**layout)
    return fig


def make_pie(df, col, title, colors):
    d = df[col].value_counts().reset_index()
    d.columns = ["Label", "Jumlah"]
    fig = px.pie(d, names="Label", values="Jumlah",
                 color_discrete_sequence=colors, hole=0.38)
    fig.update_traces(
        textinfo="percent+label",
        textfont=dict(size=13, color=TEXT_PRIMARY),
        pull=[0.04] * len(d),
        marker=dict(line=dict(color=BG_DARK, width=2)),
    )
    fig.update_layout(
        paper_bgcolor=BG_CARD, plot_bgcolor=BG_CARD,
        font=dict(color=TEXT_PRIMARY, size=12),
        title=dict(text=title, font=dict(size=14, color=TEXT_PRIMARY)),
        showlegend=True,
        legend=dict(
            font=dict(color=TEXT_MUTED, size=11),
            bgcolor="rgba(0,0,0,0)",
            orientation="h", yanchor="bottom", y=-0.2,
            xanchor="center", x=0.5,
        ),
        height=310,
        margin=dict(t=50, b=60, l=10, r=10),
    )
    return fig


def kpi(col, accent_cls, icon, label, value, sub):
    col.markdown(f"""
    <div class="kpi-card {accent_cls}">
      <h4>{icon} {label}</h4>
      <div class="kv">{value}</div>
      <div class="ksub">{sub}</div>
    </div>""", unsafe_allow_html=True)


def stats_table(container, rows):
    rows_html = "".join(f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in rows)
    container.markdown(f"""
    <div class="stats-card">
      <table>
        <tr><th>Statistik</th><th>Nilai</th></tr>
        {rows_html}
      </table>
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# BATCH LOADING: baca semua file dan buat summary dataframe
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def load_all_files(files_data):
    """
    files_data: list of (filename, bytes_content)
    Return: list of dict dengan ringkasan tiap meeting
    """
    records = []
    errors  = []

    for fname, fbytes in files_data:
        try:
            if fname.lower().endswith(".xlsx"):
                raw_str = xlsx_to_csv_string(fbytes)
                df_raw, fmt = load_csv(StringIO(raw_str), filename=fname)
            else:
                raw_str = fbytes.decode("utf-8-sig")
                df_raw, fmt = load_csv(StringIO(raw_str), filename=fname)

            m = detect_cols(df_raw)
            platform, materi, pemateri = parse_filename(fname)

            # Konversi kolom datetime
            for key in ["join", "leave", "start", "end"]:
                if key in m:
                    df_raw[m[key]] = parse_dt(df_raw[m[key]])
            if "duration_p" in m:
                df_raw[m["duration_p"]] = pd.to_numeric(df_raw[m["duration_p"]], errors="coerce")
            if "duration_m" in m:
                df_raw[m["duration_m"]] = pd.to_numeric(df_raw[m["duration_m"]], errors="coerce")

            # Hitung statistik ringkasan
            topic = df_raw[m["topic"]].iloc[0] if "topic" in m else fname
            host  = df_raw[m["host"]].iloc[0]  if "host"  in m else "-"

            n_peserta = int(df_raw[m["participants"]].iloc[0]) if "participants" in m and \
                pd.notna(df_raw[m["participants"]].iloc[0]) else len(df_raw)

            avg_dur = round(df_raw[m["duration_p"]].mean(), 1) if "duration_p" in m else None
            med_dur = round(df_raw[m["duration_p"]].median(), 1) if "duration_p" in m else None

            dur_meeting = None
            if "duration_m" in m:
                _dm = df_raw[m["duration_m"]].iloc[0]
                if pd.notna(_dm):
                    dur_meeting = float(_dm)

            comp_rate = None
            if "duration_p" in m and dur_meeting:
                n_comp = (df_raw[m["duration_p"]] >= dur_meeting * 0.8).sum()
                comp_rate = round(n_comp / len(df_raw) * 100, 1) if len(df_raw) > 0 else 0

            start_dt = df_raw[m["start"]].iloc[0] if "start" in m and pd.notna(df_raw[m["start"]].iloc[0]) else None
            end_dt   = df_raw[m["end"]].iloc[0]   if "end"   in m and pd.notna(df_raw[m["end"]].iloc[0])   else None

            records.append({
                "filename":    fname,
                "platform":    platform,
                "materi":      materi,
                "pemateri":    pemateri,
                "topic":       topic,
                "host":        host,
                "n_peserta":   n_peserta,
                "avg_dur":     avg_dur,
                "med_dur":     med_dur,
                "dur_meeting": dur_meeting,
                "comp_rate":   comp_rate,
                "start_dt":    start_dt,
                "end_dt":      end_dt,
                "format":      fmt,
                "_df":         df_raw,
                "_m":          m,
            })
        except Exception as e:
            errors.append((fname, str(e)))

    return records, errors


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📂 Sumber Data")
    mode = st.radio("Pilih sumber:", [
        "Single File (CSV/XLSX)",
        "Multi File — Batch Dashboard",
        "Link URL / Google Sheet",
    ])

    df_raw      = None
    filename    = "data_zoom.csv"
    format_type = "flat"
    batch_records = None

    # ── Single file ──────────────────────────────────────────────────────────
    if mode == "Single File (CSV/XLSX)":
        f = st.file_uploader("Upload CSV/XLSX laporan Zoom", type=["csv", "xlsx"])
        if f:
            try:
                if f.name.lower().endswith(".xlsx"):
                    raw_bytes = f.read()
                    raw_str   = xlsx_to_csv_string(raw_bytes)
                    df_raw, format_type = load_csv(StringIO(raw_str), filename=f.name)
                else:
                    df_raw, format_type = load_csv(f)
                filename = f.name
                fmt_label = "FLAT" if format_type == "flat" else "SPLIT"
                fmt_color = "format-flat" if format_type == "flat" else "format-split"
                st.success(f"✅ {len(df_raw)} baris dimuat")
                st.markdown(
                    f'Format terdeteksi: <span class="format-badge {fmt_color}">{fmt_label}</span>',
                    unsafe_allow_html=True
                )
            except Exception as e:
                st.error(f"Gagal: {e}")

    # ── Multi file ───────────────────────────────────────────────────────────
    elif mode == "Multi File — Batch Dashboard":
        files = st.file_uploader(
            "Upload semua CSV/XLSX Zoom (pilih banyak)",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
        )
        if files:
            with st.spinner(f"Memuat {len(files)} file..."):
                files_data = []
                for f in files:
                    content = f.read()
                    try:
                        f.seek(0)
                    except Exception:
                        pass
                    files_data.append((f.name, content if isinstance(content, bytes) else content.encode()))

                batch_records, batch_errors = load_all_files(files_data)

            st.success(f"✅ {len(batch_records)} file berhasil dimuat")
            if batch_errors:
                st.warning(f"⚠️ {len(batch_errors)} file gagal:")
                for fn, err in batch_errors:
                    st.caption(f"• {fn}: {err}")

    # ── URL ──────────────────────────────────────────────────────────────────
    else:
        url = st.text_input("URL CSV / Google Sheet:", placeholder="https://...")
        if st.button("Muat Data") and url:
            with st.spinner("Mengambil data..."):
                try:
                    if "docs.google.com/spreadsheets" in url:
                        base = url.split("/edit")[0].split("/pub")[0]
                        url  = base + "/export?format=csv"
                    r = requests.get(url, timeout=30)
                    r.raise_for_status()
                    df_raw, format_type = load_csv(StringIO(r.text))
                    fmt_label = "FLAT" if format_type == "flat" else "SPLIT"
                    fmt_color = "format-flat" if format_type == "flat" else "format-split"
                    st.success(f"✅ {len(df_raw)} baris dimuat")
                    st.markdown(
                        f'Format terdeteksi: <span class="format-badge {fmt_color}">{fmt_label}</span>',
                        unsafe_allow_html=True
                    )
                except Exception as e:
                    st.error(f"Gagal: {e}")

    st.markdown("---")
    st.markdown("**Format yang didukung**")
    st.caption(
        "- Laporan peserta Zoom (ID / EN)\n"
        "- Format FLAT & SPLIT\n"
        "- CSV (koma / titik koma)\n"
        "- Excel (.xlsx)\n"
        "- Google Sheet (share link)"
    )
    st.markdown("**Nama file:**")
    st.code("NamaPlatform_Materi_Pemateri.csv", language=None)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN — routing berdasarkan mode
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("# 📊 Zoom Meeting Analyzer")
st.caption("Dashboard analisis laporan peserta Zoom — Single & Batch Multi-Meeting.")

# ─────────────────────────────────────────────────────────────────────────────
# MODE BATCH DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
if mode == "Multi File — Batch Dashboard":
    if not batch_records:
        st.info("👈 Upload beberapa file CSV/XLSX di sidebar untuk memulai Batch Dashboard.")
        st.stop()

    df_sum = pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")}
                            for r in batch_records])

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab_overview, tab_platform, tab_pemateri, tab_meeting, tab_peserta = st.tabs([
        "🌐 Overview",
        "🏢 Per Platform",
        "🎓 Per Pemateri",
        "📋 Per Meeting",
        "👥 Gabungan Peserta",
    ])

    # ════════════════════════════════════════════════════════════════════════
    # TAB 1 — OVERVIEW
    # ════════════════════════════════════════════════════════════════════════
    with tab_overview:
        st.markdown('<div class="sec-title">📊 Ringkasan Keseluruhan</div>', unsafe_allow_html=True)

        total_meetings  = len(df_sum)
        total_peserta   = int(df_sum["n_peserta"].sum())
        avg_peserta     = round(df_sum["n_peserta"].mean(), 1)
        avg_comp        = round(df_sum["comp_rate"].dropna().mean(), 1) if df_sum["comp_rate"].notna().any() else 0
        avg_dur_all     = round(df_sum["avg_dur"].dropna().mean(), 1) if df_sum["avg_dur"].notna().any() else 0
        n_platform      = df_sum["platform"].nunique()
        n_pemateri      = df_sum["pemateri"].nunique()

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        kpi(k1, "blue",   "📹", "Total Meeting",       total_meetings,    "file dianalisis")
        kpi(k2, "green",  "👥", "Total Peserta",        f"{total_peserta:,}","kumulatif semua sesi")
        kpi(k3, "amber",  "📈", "Rata-rata Peserta",    avg_peserta,        "per meeting")
        kpi(k4, "purple", "✅", "Avg Completion Rate",  f"{avg_comp}%",     "kehadiran ≥80%")
        kpi(k5, "teal",   "⏱️", "Avg Durasi Peserta",   f"{avg_dur_all} mnt","rata-rata semua sesi")
        kpi(k6, "pink",   "🏢", "Platform / Pemateri",  f"{n_platform} / {n_pemateri}", "unik")

        # Bar: peserta per meeting
        st.markdown('<div class="sec-title">👥 Jumlah Peserta per Meeting</div>', unsafe_allow_html=True)
        df_bar = df_sum.sort_values("n_peserta", ascending=True)
        label_list = [f"{r['platform']} — {r['materi'][:25]}" for _, r in df_bar.iterrows()]
        fig_bar = go.Figure(go.Bar(
            x=df_bar["n_peserta"], y=label_list,
            orientation="h",
            marker_color=[PLATFORM_COLORS[i % len(PLATFORM_COLORS)] for i in range(len(df_bar))],
            text=df_bar["n_peserta"], textposition="outside",
            textfont=dict(color=TEXT_PRIMARY, size=11),
            marker_line_color=BG_DARK, marker_line_width=1,
        ))
        dark_chart(fig_bar, height=max(350, len(df_bar) * 26), title="Jumlah Peserta per Meeting")
        fig_bar.update_layout(
            xaxis_title="Jumlah Peserta",
            margin=dict(t=50, b=35, l=280, r=60),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # Scatter: peserta vs completion rate
        if df_sum["comp_rate"].notna().any():
            st.markdown('<div class="sec-title">📉 Peserta vs Completion Rate</div>', unsafe_allow_html=True)
            df_sc = df_sum.dropna(subset=["comp_rate"])
            fig_sc = px.scatter(
                df_sc, x="n_peserta", y="comp_rate",
                color="platform",
                size="avg_dur",
                hover_name="materi",
                hover_data={"pemateri": True, "n_peserta": True, "comp_rate": True},
                color_discrete_sequence=PLATFORM_COLORS,
                labels={"n_peserta": "Jumlah Peserta", "comp_rate": "Completion Rate (%)", "platform": "Platform"},
            )
            fig_sc.update_layout(**DARK_LAYOUT)
            fig_sc.update_layout(height=380, title=dict(text="Scatter: Peserta vs Completion Rate"))
            fig_sc.update_traces(marker=dict(line=dict(color=BG_DARK, width=1)))
            st.plotly_chart(fig_sc, use_container_width=True)

        # Tabel ringkasan
        st.markdown('<div class="sec-title">📋 Tabel Ringkasan Semua Meeting</div>', unsafe_allow_html=True)
        tbl_sum = df_sum[["platform","materi","pemateri","n_peserta","avg_dur","comp_rate","dur_meeting"]].copy()
        tbl_sum.columns = ["Platform","Materi","Pemateri","Peserta","Avg Durasi (mnt)","Completion (%)","Dur Meeting (mnt)"]
        st.dataframe(tbl_sum, use_container_width=True, height=360)

        st.download_button(
            "⬇️ Download Ringkasan CSV",
            data=tbl_sum.to_csv(index=False).encode("utf-8"),
            file_name="zoom_batch_summary.csv",
            mime="text/csv",
        )

    # ════════════════════════════════════════════════════════════════════════
    # TAB 2 — PER PLATFORM
    # ════════════════════════════════════════════════════════════════════════
    with tab_platform:
        st.markdown('<div class="sec-title">🏢 Analisis per Platform</div>', unsafe_allow_html=True)

        plat_grp = df_sum.groupby("platform").agg(
            total_meeting  = ("filename",    "count"),
            total_peserta  = ("n_peserta",   "sum"),
            avg_peserta    = ("n_peserta",   "mean"),
            avg_dur        = ("avg_dur",     "mean"),
            avg_comp       = ("comp_rate",   "mean"),
        ).reset_index()
        plat_grp = plat_grp.sort_values("total_peserta", ascending=False)

        # KPI per platform
        for _, row in plat_grp.iterrows():
            pa, pb, pc, pd_ = st.columns(4)
            pa.metric("Platform", row["platform"])
            pb.metric("Total Meeting", int(row["total_meeting"]))
            pc.metric("Total Peserta", f"{int(row['total_peserta']):,}")
            pd_.metric("Avg Completion", f"{row['avg_comp']:.1f}%" if pd.notna(row["avg_comp"]) else "N/A")

        st.markdown("---")
        c1, c2 = st.columns(2)
        with c1:
            fig_p1 = go.Figure(go.Bar(
                x=plat_grp["platform"], y=plat_grp["total_peserta"],
                marker_color=PLATFORM_COLORS[:len(plat_grp)],
                text=plat_grp["total_peserta"], textposition="outside",
                textfont=dict(color=TEXT_PRIMARY),
                marker_line_color=BG_DARK, marker_line_width=1,
            ))
            dark_chart(fig_p1, height=320, title="Total Peserta per Platform")
            fig_p1.update_layout(xaxis_title="Platform", yaxis_title="Peserta")
            st.plotly_chart(fig_p1, use_container_width=True)

        with c2:
            fig_p2 = go.Figure(go.Bar(
                x=plat_grp["platform"], y=plat_grp["total_meeting"],
                marker_color=PLATFORM_COLORS[:len(plat_grp)],
                text=plat_grp["total_meeting"], textposition="outside",
                textfont=dict(color=TEXT_PRIMARY),
                marker_line_color=BG_DARK, marker_line_width=1,
            ))
            dark_chart(fig_p2, height=320, title="Jumlah Meeting per Platform")
            fig_p2.update_layout(xaxis_title="Platform", yaxis_title="Jumlah Meeting")
            st.plotly_chart(fig_p2, use_container_width=True)

        if plat_grp["avg_comp"].notna().any():
            fig_comp = go.Figure(go.Bar(
                x=plat_grp["platform"],
                y=plat_grp["avg_comp"].fillna(0),
                marker_color=[ACCENT_GREEN if v >= 70 else ACCENT_AMBER if v >= 50 else ACCENT_RED
                               for v in plat_grp["avg_comp"].fillna(0)],
                text=[f"{v:.1f}%" for v in plat_grp["avg_comp"].fillna(0)],
                textposition="outside",
                textfont=dict(color=TEXT_PRIMARY),
                marker_line_color=BG_DARK, marker_line_width=1,
            ))
            dark_chart(fig_comp, height=300, title="Rata-rata Completion Rate per Platform (%)")
            fig_comp.update_layout(xaxis_title="Platform", yaxis_title="Avg Completion Rate (%)")
            st.plotly_chart(fig_comp, use_container_width=True)

    # ════════════════════════════════════════════════════════════════════════
    # TAB 3 — PER PEMATERI
    # ════════════════════════════════════════════════════════════════════════
    with tab_pemateri:
        st.markdown('<div class="sec-title">🎓 Analisis per Pemateri</div>', unsafe_allow_html=True)

        pm_grp = df_sum.groupby("pemateri").agg(
            total_meeting  = ("filename",    "count"),
            total_peserta  = ("n_peserta",   "sum"),
            avg_peserta    = ("n_peserta",   "mean"),
            avg_dur        = ("avg_dur",     "mean"),
            avg_comp       = ("comp_rate",   "mean"),
        ).reset_index().sort_values("avg_peserta", ascending=False).reset_index(drop=True)

        # ── Tabel ranking pakai st.dataframe (aman, tidak ada HTML sanitasi) ──
        st.markdown('<div class="sec-title">🏅 Ranking Pemateri</div>', unsafe_allow_html=True)
        tbl_pm = pm_grp.copy()
        tbl_pm.index = tbl_pm.index + 1
        tbl_pm.index.name = "#"
        tbl_pm["avg_peserta"] = tbl_pm["avg_peserta"].round(1)
        tbl_pm["avg_comp"]    = tbl_pm["avg_comp"].round(1)
        tbl_pm["avg_dur"]     = tbl_pm["avg_dur"].round(1)
        tbl_pm.rename(columns={
            "pemateri":      "Pemateri",
            "total_meeting": "Jumlah Sesi",
            "total_peserta": "Total Peserta",
            "avg_peserta":   "Avg Peserta",
            "avg_comp":      "Avg Completion (%)",
            "avg_dur":       "Avg Durasi (mnt)",
        }, inplace=True)
        st.dataframe(tbl_pm, use_container_width=True, height=min(600, 40 + len(tbl_pm) * 36))

        # ── Chart 1: Avg Peserta per Pemateri ────────────────────────────────
        st.markdown('<div class="sec-title">📊 Chart Perbandingan Pemateri</div>', unsafe_allow_html=True)
        top_n = pm_grp.sort_values("avg_peserta", ascending=True).tail(20)
        fig_pm1 = go.Figure(go.Bar(
            x=top_n["avg_peserta"],
            y=top_n["pemateri"],
            orientation="h",
            marker_color=ACCENT_BLUE,
            text=[f"{v:.0f}" for v in top_n["avg_peserta"]],
            textposition="outside",
            textfont=dict(color=TEXT_PRIMARY, size=11),
            marker_line_color=BG_DARK, marker_line_width=1,
        ))
        dark_chart(fig_pm1, height=max(360, len(top_n) * 30), title="Avg Peserta per Pemateri (Top 20)")
        fig_pm1.update_layout(
            xaxis_title="Avg Peserta",
            margin=dict(t=50, b=35, l=180, r=80),
        )
        st.plotly_chart(fig_pm1, use_container_width=True)

        # ── Chart 2: Total Peserta per Pemateri ──────────────────────────────
        top_tot = pm_grp.sort_values("total_peserta", ascending=True).tail(20)
        fig_pm3 = go.Figure(go.Bar(
            x=top_tot["total_peserta"],
            y=top_tot["pemateri"],
            orientation="h",
            marker_color=ACCENT_TEAL,
            text=[f"{int(v)}" for v in top_tot["total_peserta"]],
            textposition="outside",
            textfont=dict(color=TEXT_PRIMARY, size=11),
            marker_line_color=BG_DARK, marker_line_width=1,
        ))
        dark_chart(fig_pm3, height=max(360, len(top_tot) * 30), title="Total Peserta Kumulatif per Pemateri (Top 20)")
        fig_pm3.update_layout(
            xaxis_title="Total Peserta",
            margin=dict(t=50, b=35, l=180, r=80),
        )
        st.plotly_chart(fig_pm3, use_container_width=True)

        # ── Chart 3: Avg Completion Rate ─────────────────────────────────────
        pm_comp = pm_grp.dropna(subset=["avg_comp"])
        if len(pm_comp) > 0:
            top_c = pm_comp.sort_values("avg_comp", ascending=True).tail(20)
            fig_pm2 = go.Figure(go.Bar(
                x=top_c["avg_comp"],
                y=top_c["pemateri"],
                orientation="h",
                marker_color=[
                    ACCENT_GREEN if v >= 70 else ACCENT_AMBER if v >= 50 else ACCENT_RED
                    for v in top_c["avg_comp"]
                ],
                text=[f"{v:.1f}%" for v in top_c["avg_comp"]],
                textposition="outside",
                textfont=dict(color=TEXT_PRIMARY, size=11),
                marker_line_color=BG_DARK, marker_line_width=1,
            ))
            dark_chart(fig_pm2, height=max(360, len(top_c) * 30), title="Avg Completion Rate per Pemateri (Top 20)")
            fig_pm2.update_layout(
                xaxis_title="Avg Completion Rate (%)",
                margin=dict(t=50, b=35, l=180, r=80),
            )
            st.plotly_chart(fig_pm2, use_container_width=True)

    # ════════════════════════════════════════════════════════════════════════
    # TAB 4 — PER MEETING (pilih dari dropdown)
    # ════════════════════════════════════════════════════════════════════════
    with tab_meeting:
        st.markdown('<div class="sec-title">📋 Detail per Meeting</div>', unsafe_allow_html=True)

        meeting_options = [f"{r['platform']} | {r['materi']} | {r['pemateri']}"
                           for r in batch_records]
        selected_idx = st.selectbox("Pilih Meeting:", range(len(meeting_options)),
                                    format_func=lambda i: meeting_options[i])

        rec = batch_records[selected_idx]
        df  = rec["_df"]
        m   = rec["_m"]

        # Banner
        topic     = df[m["topic"]].iloc[0] if "topic" in m else rec["materi"]
        host      = df[m["host"]].iloc[0]  if "host"  in m else "-"
        start_val = df[m["start"]].iloc[0] if "start" in m else None
        end_val   = df[m["end"]].iloc[0]   if "end"   in m else None
        start_str = start_val.strftime("%d %b %Y, %H:%M") if (start_val is not None and pd.notna(start_val)) else "-"
        end_str   = end_val.strftime("%H:%M")              if (end_val   is not None and pd.notna(end_val))   else "-"

        st.markdown(f"""
        <div class="banner">
          <div class="b-title">📹 {topic}</div>
          <div>
            <span class="b-tag">🏢 {rec['platform']}</span>
            <span class="b-tag">📚 {rec['materi'][:30]}</span>
            <span class="b-tag">🎓 {rec['pemateri']}</span>
          </div>
          <div class="b-sub">👤 Host: <b style="color:#E8EAF6">{host}</b> &nbsp;|&nbsp; 🗓️ {start_str} → {end_str}</div>
        </div>
        """, unsafe_allow_html=True)

        # KPI
        avg_dur = round(df[m["duration_p"]].mean(), 1) if "duration_p" in m else 0
        comp_rate = f"{rec['comp_rate']}%" if rec["comp_rate"] is not None else "N/A"
        comp_sub  = f"dari {rec['n_peserta']} peserta"

        if start_val is not None and end_val is not None and pd.notna(start_val) and pd.notna(end_val):
            total_sec = max(int((end_val - start_val).total_seconds()), 0)
            jam, menit, detik = total_sec // 3600, (total_sec % 3600) // 60, total_sec % 60
            durasi_str = f"{jam:02d}:{menit:02d}:{detik:02d}"
            durasi_sub = f"{jam} jam {menit} menit"
        elif rec["dur_meeting"]:
            total_sec = int(rec["dur_meeting"] * 60)
            jam, menit, detik = total_sec // 3600, (total_sec % 3600) // 60, total_sec % 60
            durasi_str = f"{jam:02d}:{menit:02d}:{detik:02d}"
            durasi_sub = f"{jam} jam {menit} menit"
        else:
            durasi_str, durasi_sub = "--:--:--", "data tidak tersedia"

        k1, k2, k3, k4 = st.columns(4)
        kpi(k1, "blue",   "👥", "Total Peserta",       rec["n_peserta"],   "orang terdaftar")
        kpi(k2, "green",  "⏱️", "Durasi Meeting",       durasi_str,         durasi_sub)
        kpi(k3, "amber",  "📈", "Rata-rata Durasi",     f"{avg_dur} mnt",   "per peserta")
        kpi(k4, "purple", "✅", "Completion Rate ≥80%", comp_rate,          comp_sub)

        # Distribusi durasi
        if "duration_p" in m:
            dur = df[m["duration_p"]].dropna()
            if len(dur) > 0:
                c1, c2 = st.columns(2)
                with c1:
                    bins = [0, 10, 20, 30, 45, 60, 9999]
                    lbls = ["< 10m", "10-20m", "20-30m", "30-45m", "45-60m", "60m+"]
                    grp  = pd.cut(dur, bins=bins, labels=lbls).value_counts().sort_index().reset_index()
                    grp.columns = ["Rentang", "Jumlah"]
                    fig = go.Figure(go.Bar(
                        x=grp["Rentang"], y=grp["Jumlah"],
                        text=grp["Jumlah"], textposition="outside",
                        textfont=dict(color=TEXT_PRIMARY, size=12),
                        marker_color=BAR_BLUES[:len(grp)],
                        marker_line_color=BG_DARK, marker_line_width=1.5,
                    ))
                    dark_chart(fig, height=300, title="Distribusi Durasi Peserta")
                    fig.update_layout(xaxis_title="Rentang", yaxis_title="Peserta")
                    st.plotly_chart(fig, use_container_width=True)

                with c2:
                    fig_box = go.Figure()
                    fig_box.add_trace(go.Box(
                        y=dur, name="Durasi",
                        marker_color=ACCENT_BLUE, line_color=ACCENT_BLUE,
                        fillcolor="rgba(77,159,255,0.18)",
                        boxmean="sd", notched=True, whiskerwidth=0.6,
                    ))
                    dark_chart(fig_box, height=300, title="Box Plot Durasi")
                    fig_box.update_layout(showlegend=False, yaxis_title="Menit")
                    st.plotly_chart(fig_box, use_container_width=True)

        # Tabel peserta
        st.markdown('<div class="sec-title">📋 Data Peserta</div>', unsafe_allow_html=True)
        display_keys = ["name","email","join","leave","duration_p","guest","waiting","disclaimer"]
        rename_map   = {
            "name": "Nama", "email": "Email",
            "join": "Waktu Bergabung", "leave": "Waktu Keluar",
            "duration_p": "Durasi (menit)", "guest": "Tamu",
            "waiting": "Ruang Tunggu", "disclaimer": "Respons Rekaman",
        }
        cols_to_show = [m[k] for k in display_keys if k in m]
        tbl = df[cols_to_show].copy()
        tbl.rename(columns={m[k]: rename_map[k] for k in display_keys if k in m}, inplace=True)
        for c in ["Waktu Bergabung", "Waktu Keluar"]:
            if c in tbl.columns:
                tbl[c] = pd.to_datetime(tbl[c], errors="coerce").dt.strftime("%H:%M:%S")
        st.dataframe(tbl, use_container_width=True, height=320)

    # ════════════════════════════════════════════════════════════════════════
    # TAB 5 — GABUNGAN PESERTA (semua meeting digabung)
    # ════════════════════════════════════════════════════════════════════════
    with tab_peserta:
        st.markdown('<div class="sec-title">👥 Data Gabungan Semua Peserta</div>', unsafe_allow_html=True)

        all_rows = []
        for rec in batch_records:
            df_i = rec["_df"].copy()
            m_i  = rec["_m"]
            if "name" not in m_i:
                continue
            row_df = pd.DataFrame()
            row_df["Nama"]     = df_i[m_i["name"]]
            row_df["Email"]    = df_i[m_i["email"]]    if "email"      in m_i else None
            row_df["Durasi"]   = pd.to_numeric(df_i[m_i["duration_p"]], errors="coerce") if "duration_p" in m_i else None
            row_df["Platform"] = rec["platform"]
            row_df["Materi"]   = rec["materi"]
            row_df["Pemateri"] = rec["pemateri"]
            all_rows.append(row_df)

        if all_rows:
            df_all = pd.concat(all_rows, ignore_index=True)

            # Frekuensi kehadiran per peserta
            st.markdown('<div class="sec-title">🔢 Frekuensi Kehadiran (Top 30)</div>', unsafe_allow_html=True)
            freq = df_all.groupby("Nama").agg(
                Sesi_Dihadiri = ("Materi", "count"),
                Avg_Durasi    = ("Durasi", "mean"),
                Platform      = ("Platform", lambda x: ", ".join(sorted(set(x)))),
            ).reset_index().sort_values("Sesi_Dihadiri", ascending=False).head(30)

            fig_freq = go.Figure(go.Bar(
                x=freq["Nama"], y=freq["Sesi_Dihadiri"],
                marker_color=ACCENT_GREEN,
                text=freq["Sesi_Dihadiri"], textposition="outside",
                textfont=dict(color=TEXT_PRIMARY, size=10),
                marker_line_color=BG_DARK, marker_line_width=1,
            ))
            dark_chart(fig_freq, height=350, title="Top 30 Peserta — Frekuensi Kehadiran")
            fig_freq.update_layout(xaxis_title="Nama", yaxis_title="Jumlah Sesi", xaxis_tickangle=-45)
            st.plotly_chart(fig_freq, use_container_width=True)

            # Tabel gabungan
            st.markdown('<div class="sec-title">📋 Tabel Semua Peserta (Gabungan)</div>', unsafe_allow_html=True)
            search2 = st.text_input("🔍 Cari nama / email:", "", key="search_gabungan")
            tbl_all = df_all.copy()
            if search2:
                mask = tbl_all.apply(
                    lambda col: col.astype(str).str.contains(search2, case=False, na=False)
                ).any(axis=1)
                tbl_all = tbl_all[mask]
            tbl_all["Durasi"] = tbl_all["Durasi"].apply(
                lambda v: f"{v:.0f} mnt" if pd.notna(v) else "-"
            )
            st.dataframe(tbl_all, use_container_width=True, height=400)

            st.download_button(
                "⬇️ Download Data Gabungan CSV",
                data=df_all.to_csv(index=False).encode("utf-8"),
                file_name="zoom_all_participants.csv",
                mime="text/csv",
            )
        else:
            st.warning("Tidak ada data peserta yang dapat digabung (kolom nama tidak ditemukan).")

    st.stop()  # ← tidak lanjut ke blok single-file


# ─────────────────────────────────────────────────────────────────────────────
# MODE SINGLE FILE (kode asli + perbaikan kecil)
# ─────────────────────────────────────────────────────────────────────────────
if df_raw is None:
    st.info("👈 Upload file CSV/XLSX atau masukkan link URL di sidebar untuk memulai.")
    st.stop()

m  = detect_cols(df_raw)
df = df_raw.copy()

for key in ["join", "leave", "start", "end"]:
    if key in m:
        df[m[key]] = parse_dt(df[m[key]])

if "duration_p" in m:
    df[m["duration_p"]] = pd.to_numeric(df[m["duration_p"]], errors="coerce")
if "duration_m" in m:
    df[m["duration_m"]] = pd.to_numeric(df[m["duration_m"]], errors="coerce")
if "participants" in m:
    df[m["participants"]] = pd.to_numeric(df[m["participants"]], errors="coerce")

platform, materi, pemateri = parse_filename(filename)

# ── Banner ────────────────────────────────────────────────────────────────────
topic     = df[m["topic"]].iloc[0] if "topic" in m else "Meeting"
host      = df[m["host"]].iloc[0]  if "host"  in m else "-"
start_val = df[m["start"]].iloc[0] if "start" in m else None
end_val   = df[m["end"]].iloc[0]   if "end"   in m else None

start_str = start_val.strftime("%d %b %Y, %H:%M") if (start_val is not None and pd.notna(start_val)) else "-"
end_str   = end_val.strftime("%H:%M")              if (end_val   is not None and pd.notna(end_val))   else "-"

st.markdown(f"""
<div class="banner">
  <div class="b-title">📹 {topic}</div>
  <div>
    <span class="b-tag">🏢 {platform}</span>
    <span class="b-tag">📚 {materi}</span>
    <span class="b-tag">🎓 {pemateri}</span>
  </div>
  <div class="b-sub">👤 Host: <b style="color:#E8EAF6">{host}</b> &nbsp;|&nbsp; 🗓️ {start_str} → {end_str}</div>
</div>
""", unsafe_allow_html=True)

# ── KPI ────────────────────────────────────────────────────────────────────────
if "participants" in m:
    _p_val  = df[m["participants"]].iloc[0]
    total_p = int(_p_val) if pd.notna(_p_val) else len(df)
else:
    total_p = len(df)

avg_dur = round(df[m["duration_p"]].mean(), 1) if "duration_p" in m else 0

if "duration_p" in m and "duration_m" in m:
    _dur_m = df[m["duration_m"]].iloc[0]
    if pd.notna(_dur_m):
        meeting_dur = float(_dur_m)
        n_complete  = (df[m["duration_p"]] >= meeting_dur * 0.8).sum()
        comp_rate   = f"{round(n_complete / len(df) * 100, 1)}%"
        comp_sub    = f"sebanyak {int(n_complete)} peserta hadir penuh"
    else:
        comp_rate = "N/A"; comp_sub = "data tidak tersedia"
else:
    comp_rate = "N/A"; comp_sub = "data tidak tersedia"

if "start" in m and "end" in m:
    s_dt = df[m["start"]].iloc[0]; e_dt = df[m["end"]].iloc[0]
    if pd.notna(s_dt) and pd.notna(e_dt):
        total_sec = max(int((e_dt - s_dt).total_seconds()), 0)
        jam, menit, detik = total_sec // 3600, (total_sec % 3600) // 60, total_sec % 60
        durasi_str = f"{jam:02d}:{menit:02d}:{detik:02d}"; durasi_sub = f"{jam} jam {menit} menit"
    else:
        durasi_str, durasi_sub = "--:--:--", "data tidak tersedia"
elif "duration_m" in m:
    _dm = df[m["duration_m"]].iloc[0]
    if pd.notna(_dm):
        total_sec = int(float(_dm) * 60)
        jam, menit, detik = total_sec // 3600, (total_sec % 3600) // 60, total_sec % 60
        durasi_str = f"{jam:02d}:{menit:02d}:{detik:02d}"; durasi_sub = f"{jam} jam {menit} menit"
    else:
        durasi_str, durasi_sub = "--:--:--", "data tidak tersedia"
else:
    durasi_str, durasi_sub = "--:--:--", "data tidak tersedia"

k1, k2, k3, k4 = st.columns(4)
kpi(k1, "blue",   "👥", "Total Peserta",       total_p,          "orang terdaftar")
kpi(k2, "green",  "⏱️", "Durasi Meeting",       durasi_str,       durasi_sub)
kpi(k3, "amber",  "📈", "Rata-rata Durasi",     f"{avg_dur} mnt", "per peserta")
kpi(k4, "purple", "✅", "Completion Rate ≥80%", comp_rate,        comp_sub)

# ── Statistik Deskriptif ──────────────────────────────────────────────────────
if "duration_p" in m:
    st.markdown('<div class="sec-title">📐 Statistik Deskriptif Durasi Kehadiran (menit)</div>',
                unsafe_allow_html=True)
    dur = df[m["duration_p"]].dropna()
    if len(dur) > 0:
        all_stats = [
            ("Jumlah Peserta",   f"{int(dur.count())}"),
            ("Minimum",          f"{dur.min():.0f} mnt"),
            ("Maksimum",         f"{dur.max():.0f} mnt"),
            ("Rata-rata (Mean)", f"{dur.mean():.1f} mnt"),
            ("Median (Q2)",      f"{dur.median():.1f} mnt"),
            ("Modus",            f"{dur.mode().iloc[0]:.0f} mnt"),
            ("Std. Deviasi",     f"{dur.std():.1f} mnt"),
            ("Q1 (25%)",         f"{dur.quantile(0.25):.1f} mnt"),
            ("Q3 (75%)",         f"{dur.quantile(0.75):.1f} mnt"),
            ("IQR",              f"{dur.quantile(0.75) - dur.quantile(0.25):.1f} mnt"),
            ("Skewness",         f"{dur.skew():.3f}"),
            ("Kurtosis",         f"{dur.kurt():.3f}"),
        ]
        half = len(all_stats) // 2
        cs1, cs2 = st.columns(2)
        stats_table(cs1, all_stats[:half])
        stats_table(cs2, all_stats[half:])

        cb, cv = st.columns(2)
        with cb:
            fig_box = go.Figure()
            fig_box.add_trace(go.Box(
                y=dur, name="Durasi", marker_color=ACCENT_BLUE, line_color=ACCENT_BLUE,
                fillcolor="rgba(77,159,255,0.18)", boxmean="sd", notched=True, whiskerwidth=0.6,
            ))
            dark_chart(fig_box, height=300, title="Box Plot Durasi")
            fig_box.update_layout(showlegend=False, yaxis_title="Menit")
            st.plotly_chart(fig_box, use_container_width=True)
        with cv:
            fig_vio = go.Figure()
            fig_vio.add_trace(go.Violin(
                y=dur, name="Durasi", box_visible=True, meanline_visible=True,
                fillcolor="rgba(177,151,252,0.25)", line_color=ACCENT_PURP, marker_color=ACCENT_PURP,
            ))
            dark_chart(fig_vio, height=300, title="Violin Plot Durasi")
            fig_vio.update_layout(showlegend=False, yaxis_title="Menit")
            st.plotly_chart(fig_vio, use_container_width=True)
    else:
        st.warning("Kolom durasi peserta ditemukan namun tidak ada data numerik yang valid.")

# ── Distribusi Kehadiran ──────────────────────────────────────────────────────
st.markdown('<div class="sec-title">⏳ Distribusi Kehadiran</div>', unsafe_allow_html=True)
col_a, col_b = st.columns(2)

with col_a:
    if "duration_p" in m:
        dur = df[m["duration_p"]].dropna()
        if len(dur) > 0:
            bins = [0, 10, 20, 30, 45, 60, 9999]
            lbls = ["< 10m", "10-20m", "20-30m", "30-45m", "45-60m", "60m+"]
            grp  = pd.cut(dur, bins=bins, labels=lbls).value_counts().sort_index().reset_index()
            grp.columns = ["Rentang", "Jumlah"]
            fig = go.Figure(go.Bar(
                x=grp["Rentang"], y=grp["Jumlah"],
                text=grp["Jumlah"], textposition="outside",
                textfont=dict(color=TEXT_PRIMARY, size=12),
                marker_color=BAR_BLUES[:len(grp)],
                marker_line_color=BG_DARK, marker_line_width=1.5,
            ))
            dark_chart(fig, height=320, title="Distribusi Durasi Peserta")
            fig.update_layout(xaxis_title="Rentang Waktu", yaxis_title="Jumlah Peserta")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Data durasi tidak tersedia.")
    else:
        st.warning("Kolom durasi peserta tidak ditemukan.")

with col_b:
    if "join" in m:
        jt = df[m["join"]].dropna()
        if len(jt) > 0:
            hrs = jt.dt.hour + jt.dt.minute / 60
            fig2 = go.Figure(go.Histogram(
                x=hrs, nbinsx=30, marker_color=ACCENT_TEAL,
                marker_line_color=BG_DARK, marker_line_width=1, opacity=0.85,
            ))
            dark_chart(fig2, height=320, title="Sebaran Waktu Bergabung")
            fig2.update_layout(xaxis_title="Jam bergabung", yaxis_title="Jumlah Peserta")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.warning("Data waktu bergabung tidak tersedia.")
    else:
        st.warning("Kolom waktu bergabung tidak ditemukan.")

# ── Pie Charts ────────────────────────────────────────────────────────────────
st.markdown('<div class="sec-title">📋 Analisis Status Peserta</div>', unsafe_allow_html=True)
col_c, col_d, col_e = st.columns(3)
for container, key, title, colors in [
    (col_c, "disclaimer", "Respons Rekaman", PIE_DISCLAIMER),
    (col_d, "guest",      "Status Tamu",     PIE_GUEST),
    (col_e, "waiting",    "Ruang Tunggu",    PIE_WAITING),
]:
    with container:
        if key in m:
            st.plotly_chart(make_pie(df, m[key], title, colors), use_container_width=True)
        else:
            st.warning(f"Kolom '{title}' tidak ditemukan.")

# ── Timeline Gantt ────────────────────────────────────────────────────────────
if "join" in m and "leave" in m and "name" in m:
    st.markdown('<div class="sec-title">📅 Timeline Kehadiran Peserta (Top 40)</div>',
                unsafe_allow_html=True)
    sub = df[[m["name"], m["join"], m["leave"]] +
              ([m["duration_p"]] if "duration_p" in m else [])
             ].dropna(subset=[m["name"], m["join"], m["leave"]]).copy()
    sub = sub.sort_values(m["join"]).head(40).reset_index(drop=True)
    sub[m["join"]]  = pd.to_datetime(sub[m["join"]],  errors="coerce")
    sub[m["leave"]] = pd.to_datetime(sub[m["leave"]], errors="coerce")
    sub = sub.dropna(subset=[m["join"], m["leave"]])

    if len(sub) > 0:
        base_time         = sub[m["join"]].min()
        sub["_start_min"] = (sub[m["join"]]  - base_time).dt.total_seconds() / 60
        sub["_end_min"]   = (sub[m["leave"]] - base_time).dt.total_seconds() / 60
        sub["_dur_min"]   = (sub["_end_min"] - sub["_start_min"]).clip(lower=1)

        ref_col = sub[m["duration_p"]] if ("duration_p" in m and m["duration_p"] in sub.columns) else sub["_dur_min"]
        ref_col = pd.to_numeric(ref_col, errors="coerce")
        q33 = ref_col.quantile(0.33); q66 = ref_col.quantile(0.66)
        q33_int = int(round(q33));    q66_int = int(round(q66))

        label_pendek  = f"Durasi Pendek (< {q33_int} menit)"
        label_sedang  = f"Durasi Sedang ({q33_int}–{q66_int} menit)"
        label_panjang = f"Durasi Panjang (> {q66_int} menit)"

        def dur_color(d):
            if   pd.isna(d): return ACCENT_BLUE
            elif d <= q33:   return ACCENT_RED
            elif d <= q66:   return ACCENT_AMBER
            else:            return ACCENT_GREEN

        sub["_color"] = ref_col.apply(dur_color)
        fig6 = go.Figure()
        color_groups = {
            ACCENT_RED:   (label_pendek,  []),
            ACCENT_AMBER: (label_sedang,  []),
            ACCENT_GREEN: (label_panjang, []),
            ACCENT_BLUE:  ("Durasi",      []),
        }
        for _, row in sub.iterrows():
            color_groups[row["_color"]][1].append(row)

        for color, (legend_label, rows) in color_groups.items():
            if not rows: continue
            names      = [str(r[m["name"]])[:30] for r in rows]
            starts     = [r["_start_min"]          for r in rows]
            durations  = [r["_dur_min"]             for r in rows]
            customdata = [
                [r[m["join"]].strftime("%H:%M"), r[m["leave"]].strftime("%H:%M"), round(r["_dur_min"])]
                for r in rows
            ]
            fig6.add_trace(go.Bar(
                x=durations, y=names, base=starts, orientation="h",
                name=legend_label, marker_color=color,
                marker_line_color=BG_DARK, marker_line_width=0.8,
                customdata=customdata,
                hovertemplate=(
                    "<b>%{y}</b><br>Bergabung: %{customdata[0]}<br>"
                    "Keluar: %{customdata[1]}<br>Durasi: %{customdata[2]} menit<extra></extra>"
                ),
            ))

        fig6.update_layout(**DARK_LAYOUT)
        fig6.update_layout(
            height=max(400, len(sub) * 22),
            title=dict(text="Timeline Kehadiran Peserta", font=dict(size=14, color=TEXT_PRIMARY)),
            xaxis_title="Menit sejak peserta pertama bergabung",
            barmode="overlay", margin=dict(t=50, b=30, l=210, r=20), bargap=0.28,
            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1,
                        font=dict(color=TEXT_MUTED, size=11), bgcolor="rgba(0,0,0,0)"),
        )
        fig6.update_yaxes(autorange="reversed", gridcolor=BORDER, tickfont=dict(color=TEXT_MUTED, size=10))
        st.plotly_chart(fig6, use_container_width=True)

# ── Data Table ────────────────────────────────────────────────────────────────
st.markdown('<div class="sec-title">📋 Data Peserta Lengkap</div>', unsafe_allow_html=True)
display_keys = ["name","email","join","leave","duration_p","guest","waiting","disclaimer"]
rename_map   = {
    "name": "Nama", "email": "Email", "join": "Waktu Bergabung", "leave": "Waktu Keluar",
    "duration_p": "Durasi (menit)", "guest": "Tamu", "waiting": "Ruang Tunggu",
    "disclaimer": "Respons Rekaman",
}
cols_to_show = [m[k] for k in display_keys if k in m]
tbl = df[cols_to_show].copy()
tbl.rename(columns={m[k]: rename_map[k] for k in display_keys if k in m}, inplace=True)
for c in ["Waktu Bergabung", "Waktu Keluar"]:
    if c in tbl.columns:
        tbl[c] = pd.to_datetime(tbl[c], errors="coerce").dt.strftime("%H:%M:%S")

search = st.text_input("🔍 Cari nama / email:", "")
if search:
    mask = tbl.apply(
        lambda col: col.astype(str).str.contains(search, case=False, na=False)
    ).any(axis=1)
    tbl = tbl[mask]

st.dataframe(tbl, use_container_width=True, height=380)
st.download_button(
    "⬇️ Download Data sebagai CSV",
    data=tbl.to_csv(index=False).encode("utf-8"),
    file_name="zoom_analysis_export.csv",
    mime="text/csv",
)

st.markdown("---")
st.caption("Zoom Meeting Analyzer v2 · Dibuat dengan Streamlit & Plotly")