from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

from config import (
    ACTION_LIBRARY,
    APP_SUBTITLE,
    APP_TITLE,
    CHANNEL_ORDER,
    CHANNEL_RULES,
    COLORS,
    EXCEL_FILE,
    EXCLUDE_TITLE_PATTERNS,
    MAX_TOP_OPPORTUNITIES,
    MONTHS,
    MONTH_MAP,
    PAGE_TITLES,
    SIMULATION_DEFAULTS,
    TEXT_HELP,
)

# -----------------------------------------------------------------------------
# Streamlit setup
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title=APP_TITLE,
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
MONTH_SET = set(MONTHS)
CONTROL_TITLE_RE = re.compile(
    "|".join(re.escape(p) for p in EXCLUDE_TITLE_PATTERNS),
    flags=re.IGNORECASE,
)

def normalize_text(value) -> str | None:
    if value is None:
        return None
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text or None

def parse_year_value(value, fallback: int | None = None) -> int | None:
    if value is None:
        return fallback
    text = normalize_text(value)
    if not text:
        return fallback
    match = re.search(r"(\d{4})", text)
    if match:
        return int(match.group(1))
    try:
        return int(float(text))
    except Exception:
        return fallback

def parse_period_num(label):
    text = normalize_text(label)
    if not text:
        return None
    return MONTH_MAP.get(text.upper())

def is_month_period(label) -> bool:
    return parse_period_num(label) is not None

def is_control_title(title: str) -> bool:
    text = normalize_text(title) or ""
    return bool(CONTROL_TITLE_RE.search(text))

def is_aggregate_channel(channel: str) -> bool:
    text = normalize_text(channel) or ""
    upper = text.upper()
    return upper.startswith("TOTAL") or upper.startswith("TOTALES")

def channel_family(channel: str) -> str:
    text = normalize_text(channel)
    if not text:
        return "Otros"
    upper = text.upper()

    for family, tokens in CHANNEL_RULES.items():
        if family == "Otros":
            continue
        if any(token.upper() in upper for token in tokens):
            return family

    if is_aggregate_channel(text):
        return "Totales"

    return "Otros"

def safe_divide(numerator, denominator):
    if denominator in (None, 0) or pd.isna(denominator):
        return np.nan
    return numerator / denominator

def safe_pct_change(current, previous):
    if previous in (None, 0) or pd.isna(previous):
        return np.nan
    return (current - previous) / previous

def fmt_num(value):
    if pd.isna(value):
        return "-"
    return f"{value:,.0f}"

def fmt_mxn(value):
    if pd.isna(value):
        return "-"
    return f"${value:,.0f}"

def fmt_pct(value):
    if pd.isna(value):
        return "-"
    return f"{value:.1%}"

def semaforo(value):
    if pd.isna(value):
        return "⚪", "Sin dato"
    if value >= 8:
        return "🟢", "Alto"
    if value >= 6.5:
        return "🟡", "Medio"
    return "🔴", "Crítico"

def make_header():
    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg, {COLORS['navy']} 0%, {COLORS['info']} 100%);
                    padding: 18px 22px; border-radius: 16px; color: white; margin-bottom: 14px;">
            <div style="font-size: 2rem; font-weight: 800;">{APP_TITLE}</div>
            <div style="opacity: 0.9; margin-top: 4px;">{APP_SUBTITLE}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def make_metric_row(metrics):
    cols = st.columns(len(metrics))
    for col, metric in zip(cols, metrics):
        with col:
            st.metric(**metric)

def title_rows(ws):
    found = []
    seen = set()
    for r in range(1, ws.max_row + 1):
        title = None
        title_col = None
        for c in (1, 2):
            v = normalize_text(ws.cell(r, c).value)
            if v and v.startswith("Aforo e Ingreso -"):
                title = v.replace("Aforo e Ingreso -", "").strip()
                title_col = c
                break
        if title and r not in seen:
            found.append((r, title_col, title))
            seen.add(r)
    return found

def header_row_after(ws, title_row, lookahead=6):
    for r in range(title_row + 1, min(ws.max_row, title_row + lookahead) + 1):
        for c in range(1, min(ws.max_column, 4) + 1):
            if normalize_text(ws.cell(r, c).value) == "FECHA":
                return r, c
    return None, None

def collect_channel_labels(ws, header_row, fecha_col):
    labels = []
    col = fecha_col + 2
    while col <= ws.max_column:
        label = normalize_text(ws.cell(header_row - 1, col).value) if header_row - 1 >= 1 else None
        if not label:
            break
        labels.append(label)
        col += 2
    return labels

def parse_hoja2(ws):
    if ws is None:
        return pd.DataFrame()

    records = []
    for row in range(4, ws.max_row + 1):
        periodo = normalize_text(ws.cell(row, 1).value)
        if not periodo:
            continue

        values = [ws.cell(row, c).value for c in range(2, 8)]
        if all(v is None for v in values):
            continue

        records.append(
            {
                "periodo": periodo,
                "mis_en_sus_aforo": ws.cell(row, 2).value,
                "mis_en_sus_ingreso": ws.cell(row, 3).value,
                "mis_en_mis_aforo": ws.cell(row, 4).value,
                "mis_en_mis_ingreso": ws.cell(row, 5).value,
                "sus_en_mis_aforo": ws.cell(row, 6).value,
                "sus_en_mis_ingreso": ws.cell(row, 7).value,
            }
        )
    return pd.DataFrame(records)

def parse_workbook(file_path):
    wb = load_workbook(file_path, data_only=True)
    records = []

    for sheet_name in wb.sheetnames:
        if sheet_name == "Hoja2":
            continue

        ws = wb[sheet_name]
        sheet_year = parse_year_value(sheet_name)

        titles = title_rows(ws)

        for idx, (tr, _, title) in enumerate(titles):
            header_row, fecha_col = header_row_after(ws, tr)
            if header_row is None:
                continue

            next_title_row = titles[idx + 1][0] if idx + 1 < len(titles) else ws.max_row + 1
            channel_labels = collect_channel_labels(ws, header_row, fecha_col)

            if len(channel_labels) < 1:
                continue

            block_kind = "control" if is_control_title(title) else "operational"

            for row in range(header_row + 2, next_title_row):
                period = normalize_text(ws.cell(row, fecha_col).value)
                if not period:
                    continue

                period_upper = period.upper()
                if period_upper not in MONTH_SET and period_upper not in {"TOTALES", "PROMEDIO"}:
                    continue

                row_year = parse_year_value(ws.cell(row, fecha_col + 1).value, fallback=sheet_year)

                for channel_idx, channel in enumerate(channel_labels):
                    aforo_col = fecha_col + 2 + channel_idx * 2
                    ingreso_col = aforo_col + 1

                    if aforo_col > ws.max_column:
                        break

                    aforo = ws.cell(row, aforo_col).value
                    ingreso = ws.cell(row, ingreso_col).value if ingreso_col <= ws.max_column else None

                    if aforo is None and ingreso is None:
                        continue

                    records.append(
                        {
                            "sheet": sheet_name,
                            "year": row_year,
                            "concession": title,
                            "title_row": tr,
                            "header_row": header_row,
                            "fecha_col": fecha_col,
                            "period": period_upper,
                            "month_num": parse_period_num(period_upper),
                            "period_type": "month" if is_month_period(period_upper) else "summary",
                            "channel": channel,
                            "family": channel_family(channel),
                            "is_aggregate_channel": is_aggregate_channel(channel),
                            "aforo": float(aforo) if aforo is not None else 0.0,
                            "ingreso": float(ingreso) if ingreso is not None else 0.0,
                            "block_kind": block_kind,
                        }
                    )

    raw_df = pd.DataFrame(records)

    if not raw_df.empty:
        raw_df["concession"] = raw_df["concession"].astype(str).str.strip()
        raw_df["channel"] = raw_df["channel"].astype(str).str.strip()
        raw_df["family"] = raw_df["family"].astype(str).str.strip()
        raw_df["block_kind"] = raw_df["block_kind"].astype(str).str.strip()
        raw_df["year"] = pd.to_numeric(raw_df["year"], errors="coerce").astype("Int64")
        raw_df["month_num"] = pd.to_numeric(raw_df["month_num"], errors="coerce").astype("Int64")

    hoja2_df = parse_hoja2(wb["Hoja2"]) if "Hoja2" in wb.sheetnames else pd.DataFrame()
    return raw_df, hoja2_df

@st.cache_data(show_spinner=False)
def load_data(file_path):
    return parse_workbook(file_path)

def prepare_analysis(raw_df, scope_mode, year_choice, concession_choice):
    base = raw_df.copy()

    if scope_mode == TEXT_HELP["scope_operational"]:
        base = base[base["block_kind"] == "operational"].copy()

    base = base[base["period_type"] == "month"].copy()
    base = base[~base["is_aggregate_channel"]].copy()

    if year_choice != "Todos":
        base = base[base["year"] == int(year_choice)].copy()

    if concession_choice != "Todas":
        base = base[base["concession"] == concession_choice].copy()

    base = base.dropna(subset=["year", "month_num"])

    if base.empty:
        return base, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    base["year"] = base["year"].astype(int)
    base["month_num"] = base["month_num"].astype(int)

    monthly = (
        base.groupby(["year", "month_num"], as_index=False)
        .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
        .sort_values(["year", "month_num"])
    )

    concession_summary = (
        base.groupby(["year", "concession"], as_index=False)
        .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
        .sort_values(["year", "aforo"], ascending=[True, False])
    )
    concession_summary["rpc"] = concession_summary["ingreso"] / concession_summary["aforo"].replace(0, np.nan)

    family_mix = (
        base.groupby(["year", "month_num", "family"], as_index=False)
        .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
        .sort_values(["year", "month_num", "family"])
    )

    annual = (
        base.groupby(["year"], as_index=False)
        .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
        .sort_values("year")
    )
    annual["rpc"] = annual["ingreso"] / annual["aforo"].replace(0, np.nan)
    annual["aforo_yoy"] = annual["aforo"].pct_change()
    annual["ingreso_yoy"] = annual["ingreso"].pct_change()

    return base, monthly, concession_summary, family_mix, annual

def build_concession_diagnostics(df_months, selected_year):
    year_df = df_months[df_months["year"] == selected_year].copy()
    if year_df.empty:
        return pd.DataFrame()

    summary = (
        year_df.groupby("concession", as_index=False)
        .agg(
            aforo=("aforo", "sum"),
            ingreso=("ingreso", "sum"),
            aforo_std=("aforo", "std"),
            aforo_mean=("aforo", "mean"),
        )
    )
    summary["rpc"] = summary["ingreso"] / summary["aforo"].replace(0, np.nan)
    summary["volatility"] = summary["aforo_std"] / summary["aforo_mean"].replace(0, np.nan)

    prev = (
        df_months[df_months["year"] == selected_year - 1]
        .groupby("concession", as_index=False)
        .agg(
            aforo_prev=("aforo", "sum"),
            ingreso_prev=("ingreso", "sum"),
        )
    )
    diag = summary.merge(prev, on="concession", how="left")
    diag["aforo_yoy"] = diag.apply(lambda r: safe_pct_change(r["aforo"], r["aforo_prev"]), axis=1)
    diag["ingreso_yoy"] = diag.apply(lambda r: safe_pct_change(r["ingreso"], r["ingreso_prev"]), axis=1)

    family = (
        year_df.groupby(["concession", "family"], as_index=False)
        .agg(aforo=("aforo", "sum"))
    )
    family_pivot = (
        family.pivot(index="concession", columns="family", values="aforo")
        .fillna(0)
    )
    for fam in CHANNEL_ORDER:
        if fam not in family_pivot.columns:
            family_pivot[fam] = 0
    family_pivot = family_pivot[CHANNEL_ORDER]
    family_pivot["aforo_total_family"] = family_pivot.sum(axis=1)
    family_pivot["televia_share"] = family_pivot["TeleVía"] / family_pivot["aforo_total_family"].replace(0, np.nan)
    family_pivot["pase_share"] = family_pivot["PASE"] / family_pivot["aforo_total_family"].replace(0, np.nan)
    family_pivot["premium_share"] = (
        family_pivot[["TeleVía", "CAPUFE/Exentos", "PINFRA", "SITEL", "EASYTRIP"]]
        .sum(axis=1)
        / family_pivot["aforo_total_family"].replace(0, np.nan)
    )
    diag = diag.merge(
        family_pivot[["televia_share", "pase_share", "premium_share"]],
        left_on="concession",
        right_index=True,
        how="left",
    )
    return diag.sort_values("aforo", ascending=False)

def build_heatmap(df_months, selected_year):
    year_df = df_months[df_months["year"] == selected_year].copy()
    if year_df.empty:
        return pd.DataFrame()

    pivot = (
        year_df.groupby(["concession", "family"], as_index=False)
        .agg(aforo=("aforo", "sum"))
        .pivot(index="concession", columns="family", values="aforo")
        .fillna(0)
    )

    cols = [c for c in CHANNEL_ORDER if c in pivot.columns]
    pivot = pivot.reindex(columns=cols, fill_value=0)
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]
    return pivot

def build_opportunities(diag_df):
    if diag_df.empty:
        return diag_df

    work = diag_df.copy()
    required_cols = ["aforo", "rpc", "aforo_yoy", "volatility", "pase_share", "televia_share"]
    for col in required_cols:
        if col not in work.columns:
            work[col] = np.nan

    work["max_aforo"] = work["aforo"].max()
    work["median_rpc"] = work["rpc"].median()

    work["volume_component"] = work["aforo"] / work["max_aforo"].replace(0, np.nan)
    work["rpc_gap"] = ((work["median_rpc"] - work["rpc"]) / work["median_rpc"]).clip(lower=0)
    work["drop_component"] = (-work["aforo_yoy"]).clip(lower=0)
    work["vol_component"] = work["volatility"].fillna(0).clip(upper=1.5)

    work["priority_score"] = (
        0.35 * work["volume_component"].fillna(0)
        + 0.25 * work["rpc_gap"].fillna(0)
        + 0.25 * work["drop_component"].fillna(0)
        + 0.15 * work["vol_component"].fillna(0)
    )

    def recommend(row):
        if pd.notna(row["aforo_yoy"]) and row["aforo_yoy"] < 0 and row["aforo"] >= work["aforo"].median():
            return ACTION_LIBRARY["volume_recovery"]["label"], "Base relevante con caída vs año previo."
        if pd.notna(row["rpc"]) and row["rpc"] < row["median_rpc"]:
            return ACTION_LIBRARY["monetization"]["label"], "Ingreso por cruce por debajo del benchmark."
        if pd.notna(row["volatility"]) and row["volatility"] > work["volatility"].median():
            return ACTION_LIBRARY["stabilization"]["label"], "Alta variabilidad mensual."
        if pd.notna(row["pase_share"]) and row["pase_share"] > 0.20:
            return ACTION_LIBRARY["channel_capture"]["label"], "Participación PASE relevante; oportunidad de captura."
        if pd.notna(row["televia_share"]) and row["televia_share"] < 0.25 and row["aforo"] >= work["aforo"].median():
            return ACTION_LIBRARY["premium_mix"]["label"], "Mix TeleVía bajo para una base relevante."
        return "Seguimiento", "Mantener monitoreo y buscar mejoras incrementales."

    labels = work.apply(recommend, axis=1, result_type="expand")
    work["accion"] = labels[0]
    work["motivo"] = labels[1]
    return work.sort_values("priority_score", ascending=False)

def make_bar_config(fig, height=420):
    fig.update_layout(
        height=height,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=10, r=10, t=20, b=10),
        legend_title_text="",
        font=dict(size=11),
    )
    return fig

def render_table(df, cols=None, formatters=None, height=320):
    if df.empty:
        st.info("Sin datos para mostrar.")
        return
    out = df.copy()
    if cols is not None:
        out = out[cols]
    if formatters:
        st.dataframe(out.style.format(formatters), use_container_width=True, height=height, hide_index=True)
    else:
        st.dataframe(out, use_container_width=True, height=height, hide_index=True)

# -----------------------------------------------------------------------------
# Load data
# -----------------------------------------------------------------------------
excel_path = Path(EXCEL_FILE)
if not excel_path.exists():
    st.error(f"No encuentro el archivo {EXCEL_FILE} en el repositorio.")
    xlsx_files = [p.name for p in Path(".").glob("*.xlsx")]
    if xlsx_files:
        st.info("Archivos XLSX visibles: " + ", ".join(xlsx_files))
    st.stop()

with st.spinner("Cargando y estructurando el Excel..."):
    raw_df, hoja2_df = load_data(str(excel_path))

if raw_df.empty:
    st.error("No se pudo estructurar el archivo. Revisa que el libro conserve el formato de aforos.")
    st.stop()

all_years = sorted(raw_df["year"].dropna().astype(int).unique().tolist())
all_concessions = sorted(
    raw_df.loc[~raw_df["concession"].map(is_control_title), "concession"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Filtros globales")
    scope_mode = st.radio(
        "Alcance",
        [TEXT_HELP["scope_operational"], TEXT_HELP["scope_all"]],
        index=0,
    )
    year_choice = st.selectbox("Año", ["Todos"] + all_years, index=0)
    concession_choice = st.selectbox("Concesión", ["Todas"] + all_concessions, index=0)

    st.markdown("---")
    st.markdown("### Supuestos de simulación")
    sim_volume_growth = st.slider(
        "Crecimiento de volumen (%)",
        0.0, 20.0,
        float(SIMULATION_DEFAULTS["volume_growth_pct"]),
        0.1,
    )
    sim_share_capture = st.slider(
        "Captura de share (pp)",
        0.0, 10.0,
        float(SIMULATION_DEFAULTS["share_capture_pp"]),
        0.1,
    )
    sim_rpc_uplift = st.slider(
        "Mejora ingreso/cruce (%)",
        0.0, 20.0,
        float(SIMULATION_DEFAULTS["rpc_uplift_pct"]),
        0.1,
    )
    sim_sus_capture = st.slider(
        "Captura de SUS EN MIS hacia TeleVía (%)",
        0.0, 100.0,
        0.0,
        1.0,
        help="Porcentaje de vehículos de otros operadores (SUS EN MIS) que se convierten a TeleVía."
    )

# -----------------------------------------------------------------------------
# Filtering and derivations (se reutiliza el código original)
# -----------------------------------------------------------------------------
analysis_df, monthly_df, concession_summary_df, family_mix_df, annual_df = prepare_analysis(
    raw_df=raw_df,
    scope_mode=scope_mode,
    year_choice=year_choice,
    concession_choice=concession_choice,
)

if analysis_df.empty:
    st.error("No hay datos en el alcance actual. Ajusta los filtros laterales.")
    st.stop()

# Serie anual sin filtro por año / concesión para la pestaña de seguimiento.
scope_for_annual = raw_df.copy()
if scope_mode == TEXT_HELP["scope_operational"]:
    scope_for_annual = scope_for_annual[scope_for_annual["block_kind"] == "operational"].copy()
scope_for_annual = scope_for_annual[
    (scope_for_annual["period_type"] == "month")
    & (~scope_for_annual["is_aggregate_channel"])
].copy()
if concession_choice != "Todas":
    scope_for_annual = scope_for_annual[scope_for_annual["concession"] == concession_choice].copy()
scope_for_annual = scope_for_annual.dropna(subset=["year", "month_num"])
scope_for_annual["year"] = scope_for_annual["year"].astype(int)
scope_for_annual["month_num"] = scope_for_annual["month_num"].astype(int)

annual_all = (
    scope_for_annual.groupby("year", as_index=False)
    .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    .sort_values("year")
)
annual_all["rpc"] = annual_all["ingreso"] / annual_all["aforo"].replace(0, np.nan)
annual_all["aforo_yoy"] = annual_all["aforo"].pct_change()
annual_all["ingreso_yoy"] = annual_all["ingreso"].pct_change()

selected_year = int(year_choice) if year_choice != "Todos" else int(analysis_df["year"].max())
year_df = analysis_df[analysis_df["year"] == selected_year].copy()
previous_year_df = analysis_df[analysis_df["year"] == selected_year - 1].copy()

monthly_year = (
    year_df.groupby("month_num", as_index=False)
    .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    .sort_values("month_num")
)

prev_monthly_year = (
    previous_year_df.groupby("month_num", as_index=False)
    .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    .sort_values("month_num")
) if not previous_year_df.empty else pd.DataFrame(columns=["month_num", "aforo", "ingreso"])

focus_month = int(monthly_year["month_num"].max()) if not monthly_year.empty else 12
current_ytd_df = year_df[year_df["month_num"] <= focus_month].copy()
previous_ytd_df = previous_year_df[previous_year_df["month_num"] <= focus_month].copy()

portfolio_aforo = current_ytd_df["aforo"].sum()
portfolio_ingreso = current_ytd_df["ingreso"].sum()
portfolio_rpc = safe_divide(portfolio_ingreso, portfolio_aforo)

prev_aforo = previous_ytd_df["aforo"].sum() if not previous_ytd_df.empty else np.nan
prev_ingreso = previous_ytd_df["ingreso"].sum() if not previous_ytd_df.empty else np.nan

aforo_yoy = safe_pct_change(portfolio_aforo, prev_aforo)
ingreso_yoy = safe_pct_change(portfolio_ingreso, prev_ingreso)

# Concession summary for selected year / YTD
conc_year = (
    current_ytd_df.groupby("concession", as_index=False)
    .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    .sort_values("aforo", ascending=False)
)
conc_year["rpc"] = conc_year["ingreso"] / conc_year["aforo"].replace(0, np.nan)

total_aforo = conc_year["aforo"].sum()
conc_year["share_aforo"] = conc_year["aforo"] / total_aforo if total_aforo != 0 else np.nan

conc_year["aforo_rank"] = np.arange(1, len(conc_year) + 1)

prev_conc = (
    previous_ytd_df.groupby("concession", as_index=False)
    .agg(aforo_prev=("aforo", "sum"), ingreso_prev=("ingreso", "sum"))
)

conc_diag = conc_year.merge(prev_conc, on="concession", how="left")
conc_diag["aforo_yoy"] = conc_diag.apply(lambda r: safe_pct_change(r["aforo"], r["aforo_prev"]), axis=1)
conc_diag["ingreso_yoy"] = conc_diag.apply(lambda r: safe_pct_change(r["ingreso"], r["ingreso_prev"]), axis=1)

concession_monthly_vol = (
    current_ytd_df.groupby("concession")["aforo"]
    .agg(["mean", "std"])
    .reset_index()
)
concession_monthly_vol["volatility"] = concession_monthly_vol["std"] / concession_monthly_vol["mean"].replace(0, np.nan)
conc_diag = conc_diag.merge(concession_monthly_vol[["concession", "volatility"]], on="concession", how="left")

family_year = (
    current_ytd_df.groupby("family", as_index=False)
    .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    .sort_values("aforo", ascending=False)
)
total_family_aforo = family_year["aforo"].sum()
family_year["share"] = family_year["aforo"] / total_family_aforo if total_family_aforo != 0 else np.nan

heatmap_df = build_heatmap(analysis_df, selected_year)

# Preparar DataFrame para oportunidades con las columnas de familias
family_pivot = (
    current_ytd_df.groupby(["concession", "family"], as_index=False)
    .agg(aforo=("aforo", "sum"))
    .pivot(index="concession", columns="family", values="aforo")
    .fillna(0)
)
for fam in CHANNEL_ORDER:
    if fam not in family_pivot.columns:
        family_pivot[fam] = 0
family_pivot = family_pivot[CHANNEL_ORDER]
family_pivot["televia_share"] = family_pivot["TeleVía"] / family_pivot.sum(axis=1).replace(0, np.nan)
family_pivot["pase_share"] = family_pivot["PASE"] / family_pivot.sum(axis=1).replace(0, np.nan)

opportunities_df = build_opportunities(
    conc_diag.merge(
        family_pivot,
        left_on="concession",
        right_index=True,
        how="left",
    )
)

# normalize family share columns if present
for fam in ["TeleVía", "PASE", "CAPUFE/Exentos", "PINFRA", "SITEL", "EASYTRIP", "Otros", "Totales"]:
    if fam not in opportunities_df.columns:
        opportunities_df[fam] = 0

opportunities_df["family_total"] = opportunities_df[["TeleVía", "PASE", "CAPUFE/Exentos", "PINFRA", "SITEL", "EASYTRIP", "Otros"]].sum(axis=1)
opportunities_df["televia_share"] = opportunities_df["TeleVía"] / opportunities_df["family_total"].replace(0, np.nan)
opportunities_df["pase_share"] = opportunities_df["PASE"] / opportunities_df["family_total"].replace(0, np.nan)

# Top narrative insights
best_concession = conc_year.iloc[0]["concession"] if not conc_year.empty else "-"
worst_concession = conc_year.iloc[-1]["concession"] if not conc_year.empty else "-"
best_rpc_concession = conc_year.sort_values("rpc", ascending=False).iloc[0]["concession"] if not conc_year.empty else "-"
worst_rpc_concession = conc_year.sort_values("rpc", ascending=True).iloc[0]["concession"] if not conc_year.empty else "-"
best_family = family_year.iloc[0]["family"] if not family_year.empty else "-"
worst_family = family_year.iloc[-1]["family"] if not family_year.empty else "-"

# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
make_header()
st.caption(
    f"Archivo fuente: {EXCEL_FILE} · Rango analizado: {min(all_years)}–{max(all_years)} · "
    f"Alcance activo: {scope_mode} · Corte: {selected_year} YTD hasta {focus_month}"
)

# -----------------------------------------------------------------------------
# Tabs: Decision Intelligence
# -----------------------------------------------------------------------------
tab_home, tab_diag, tab_sim, tab_plan, tab_follow = st.tabs([
    "🎯 Control Tower",
    "🔍 Diagnóstico Causal",
    "🧪 Simulador",
    "💰 Portafolio Oportunidades",
    "📈 Seguimiento"
])

# ================== Página 1: Executive Control Tower ==================
with tab_home:
    st.subheader("Executive Control Tower – ¿Dónde debo intervenir esta semana?")

    if conc_diag.empty:
        st.info("No hay datos suficientes para mostrar el semáforo ejecutivo.")
    else:
        # Función auxiliar para semáforo
        def generar_semaforo_ejecutivo(df):
            df = df.copy()
            condicion = (
                (df["aforo_yoy"] >= -0.05) & (df["aforo_yoy"].notna()) &
                (df["televia_share"] >= 0.25) &
                (df["rpc"] >= df["rpc"].median())
            )
            df["semaforo"] = "🟡 vigilar"
            df.loc[condicion, "semaforo"] = "🟢 saludable"
            df.loc[~(condicion) & (df["aforo_yoy"] < -0.1), "semaforo"] = "🔴 intervenir"
            return df

        conc_diag_semaforo = generar_semaforo_ejecutivo(conc_diag)

        # KPIs resumidos
        total_aforo = conc_diag_semaforo["aforo"].sum()
        total_ingreso = conc_diag_semaforo["ingreso"].sum()
        rpc_total = safe_divide(total_ingreso, total_aforo)
        share_televia_total = safe_divide(conc_diag_semaforo["televia_share"].sum(), len(conc_diag_semaforo))
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Aforo total", fmt_num(total_aforo))
        col2.metric("Ingreso total", fmt_mxn(total_ingreso))
        col3.metric("Ingreso/cruce", f"{rpc_total:.2f}" if pd.notna(rpc_total) else "-")
        col4.metric("Share TeleVía", fmt_pct(share_televia_total))

        # Tabla de semáforo
        st.markdown("#### Semáforo ejecutivo por concesión")
        tabla_semaforo = conc_diag_semaforo[["concession", "aforo", "televia_share", "rpc", "aforo_yoy", "semaforo"]].copy()
        tabla_semaforo.columns = ["Concesión", "Aforo", "Share TeleVía", "RPC", "Tendencia", "Estado"]
        st.dataframe(
            tabla_semaforo.style.format({
                "Aforo": "{:,.0f}",
                "Share TeleVía": "{:.1%}",
                "RPC": "{:.2f}",
                "Tendencia": "{:.1%}"
            }).apply(lambda x: ["background-color: #90EE90" if x["Estado"]=="🟢 saludable" else ("background-color: #FFFF99" if x["Estado"]=="🟡 vigilar" else "background-color: #FFCCCC") for _ in x], axis=1),
            use_container_width=True,
            height=400,
            hide_index=True
        )

        # Alertas automáticas (máximo 5)
        st.markdown("#### Alertas ejecutivas")
        def generar_alertas(diag, hoja2, year):
            alertas = []
            if diag.empty:
                return alertas
            # Caída fuerte
            caida = diag[diag["aforo_yoy"] < -0.1].nlargest(2, "aforo")
            for _, row in caida.iterrows():
                alertas.append({
                    "concesion": row["concession"],
                    "impacto": f"Caída de aforo {fmt_pct(row['aforo_yoy'])}",
                    "nivel": "🔴",
                    "accion": "Revisar causas operativas y lanzar promociones"
                })
            # Baja share
            baja_share = diag[diag["televia_share"] < 0.2].nlargest(2, "aforo")
            for _, row in baja_share.iterrows():
                alertas.append({
                    "concesion": row["concession"],
                    "impacto": f"Share TeleVía bajo ({fmt_pct(row['televia_share'])})",
                    "nivel": "🟡",
                    "accion": "Campaña de migración a TeleVía"
                })
            # Bajo RPC
            bajo_rpc = diag[diag["rpc"] < diag["rpc"].quantile(0.25)].nlargest(1, "aforo")
            for _, row in bajo_rpc.iterrows():
                alertas.append({
                    "concesion": row["concession"],
                    "impacto": f"RPC {fmt_mxn(row['rpc'])} (muy bajo)",
                    "nivel": "🟡",
                    "accion": "Revisar tarifas y mix de canales"
                })
            # Concentración excesiva
            total = diag["aforo"].sum()
            diag["share_portafolio"] = diag["aforo"] / total if total > 0 else 0
            alta_conc = diag[diag["share_portafolio"] > 0.2]
            for _, row in alta_conc.iterrows():
                alertas.append({
                    "concesion": row["concession"],
                    "impacto": f"{fmt_pct(row['share_portafolio'])} del portafolio",
                    "nivel": "🟡",
                    "accion": "Diversificar o reforzar otras concesiones"
                })
            # Fuga (si hay Hoja2)
            if not hoja2.empty:
                año_str = str(year) if isinstance(year, int) else year
                fuga = hoja2[hoja2["periodo"] == año_str]
                if not fuga.empty:
                    ratio = safe_divide(fuga.iloc[0]["mis_en_sus_aforo"], fuga.iloc[0]["mis_en_mis_aforo"])
                    if ratio > 0.3:
                        alertas.append({
                            "concesion": "Portafolio",
                            "impacto": f"Fuga de tráfico propio: ratio {ratio:.2f}",
                            "nivel": "🔴",
                            "accion": "Campaña de retención nacional"
                        })
            return alertas[:5]

        alertas = generar_alertas(conc_diag, hoja2_df, selected_year)
        if alertas:
            alert_df = pd.DataFrame(alertas)
            st.dataframe(alert_df, use_container_width=True, hide_index=True)
        else:
            st.success("No hay alertas críticas en este momento.")

        # Top 3 intervenciones
        st.markdown("#### Si solo pudiera intervenir 3 concesiones este trimestre:")
        if "priority_score" in conc_diag.columns:
            top3 = conc_diag.nlargest(3, "priority_score")[["concession", "priority_score"]]
            for i, row in top3.iterrows():
                st.write(f"{i+1}. **{row['concession']}** – Score {row['priority_score']:.2f}")
        else:
            st.write("No hay concesiones con prioridad suficiente.")

# ================== Página 2: Diagnóstico Causal ==================
with tab_diag:
    st.subheader("Diagnóstico Causal – ¿Por qué pasó?")

    if conc_diag.empty:
        st.info("No hay datos para el diagnóstico.")
    else:
        # Driver principal
        def driver_principal(row):
            if pd.isna(row["aforo_yoy"]) or row["aforo_yoy"] >= 0:
                return None
            if row["televia_share"] < 0.2:
                return "Baja adopción TeleVía"
            if row["rpc"] < row["rpc"].median():
                return "Deterioro ingreso/cruce"
            return "Caída general de aforo"

        conc_diag["driver"] = conc_diag.apply(driver_principal, axis=1)

        st.markdown("#### Análisis de concesiones con caída de aforo")
        caidas = conc_diag[conc_diag["aforo_yoy"] < 0].copy()
        if caidas.empty:
            st.success("No hay concesiones con caída de aforo en el período.")
        else:
            for _, row in caidas.iterrows():
                with st.expander(f"{row['concession']} – Caída {fmt_pct(row['aforo_yoy'])}"):
                    driver = row["driver"] if pd.notna(row["driver"]) else "Sin driver específico"
                    impacto = f"Aforo: {fmt_num(row['aforo'])} | RPC: {row['rpc']:.2f}"
                    confianza = "Alta" if row["aforo_yoy"] < -0.15 else "Media"
                    st.write(f"**Driver principal:** {driver}")
                    st.write(f"**Impacto:** {impacto}")
                    st.write(f"**Confianza:** {confianza}")
                    st.write("**Conclusión automática:** " + (
                        "Implementar campaña de migración a TeleVía." if "TeleVía" in driver else
                        "Revisar tarifas y promociones." if "RPC" in driver else
                        "Realizar análisis de tráfico y mantenimiento."
                    ))

        # Elasticidad
        st.markdown("#### Elasticidad ingreso / aforo")
        if "ingreso_yoy" in conc_diag.columns:
            elasticidad = conc_diag["ingreso_yoy"] / conc_diag["aforo_yoy"].replace(0, np.nan)
            conc_diag["elasticidad"] = elasticidad
            st.dataframe(
                conc_diag[["concession", "aforo_yoy", "ingreso_yoy", "elasticidad"]].dropna().style.format({
                    "aforo_yoy": "{:.1%}",
                    "ingreso_yoy": "{:.1%}",
                    "elasticidad": "{:.2f}"
                }),
                use_container_width=True,
                height=300,
                hide_index=True
            )
            st.caption("Elasticidad >1 indica que el ingreso crece más que el aforo (buena señal).")

# ================== Página 3: Simulador ==================
with tab_sim:
    st.subheader("Simulador – ¿Qué pasa si hago algo?")

    # Inputs adicionales
    presupuesto = st.number_input("Presupuesto estimado (MXN)", min_value=0, value=5000000, step=1000000)
    objetivo_crecimiento = st.slider("Objetivo crecimiento aforo (%)", 0.0, 30.0, 5.0)
    objetivo_captura = st.slider("Objetivo captura share TeleVía (pp)", 0.0, 20.0, 5.0)

    # Datos base
    base_aforo = annual_all["aforo"].iloc[-1] if not annual_all.empty else 0
    base_ingreso = annual_all["ingreso"].iloc[-1] if not annual_all.empty else 0
    base_rpc = safe_divide(base_ingreso, base_aforo)

    sus_volumen = 0
    if not hoja2_df.empty:
        año_actual = selected_year if year_choice != "Todos" else annual_all["year"].iloc[-1]
        año_row = hoja2_df[hoja2_df["periodo"] == str(año_actual)]
        if not año_row.empty:
            sus_volumen = año_row.iloc[0]["sus_en_mis_aforo"]

    # Escenarios
    capture_vol = sus_volumen * (sim_sus_capture / 100.0)
    televia_rpc = base_rpc * 1.1
    new_ingreso_capture = capture_vol * televia_rpc
    ingreso1 = base_ingreso + new_ingreso_capture
    aforo1 = base_aforo
    roi1 = new_ingreso_capture / presupuesto if presupuesto > 0 else 0
    payback1 = presupuesto / new_ingreso_capture if new_ingreso_capture > 0 else np.inf

    prepago_rpc_actual = base_rpc * 0.9
    nuevo_rpc_prepago = prepago_rpc_actual * (1 + sim_rpc_uplift / 100.0)
    impacto_prepago = base_aforo * 0.3 * (nuevo_rpc_prepago - prepago_rpc_actual)
    ingreso2 = base_ingreso + impacto_prepago
    roi2 = impacto_prepago / presupuesto if presupuesto > 0 else 0
    payback2 = presupuesto / impacto_prepago if impacto_prepago > 0 else np.inf

    postpago_rpc_actual = base_rpc * 1.0
    nuevo_rpc_postpago = postpago_rpc_actual * (1 + sim_rpc_uplift / 100.0)
    impacto_postpago = base_aforo * 0.3 * (nuevo_rpc_postpago - postpago_rpc_actual)
    ingreso3 = base_ingreso + impacto_postpago
    roi3 = impacto_postpago / presupuesto if presupuesto > 0 else 0
    payback3 = presupuesto / impacto_postpago if impacto_postpago > 0 else np.inf

    decenal_aforo_actual = base_aforo * 0.2
    nuevo_decenal = decenal_aforo_actual * (1 + sim_volume_growth / 100.0)
    delta_aforo = nuevo_decenal - decenal_aforo_actual
    nuevo_ingreso_decenal = delta_aforo * base_rpc
    ingreso4 = base_ingreso + nuevo_ingreso_decenal
    aforo4 = base_aforo + delta_aforo
    roi4 = nuevo_ingreso_decenal / presupuesto if presupuesto > 0 else 0
    payback4 = presupuesto / nuevo_ingreso_decenal if nuevo_ingreso_decenal > 0 else np.inf

    resultados = pd.DataFrame([
        {"Escenario": "Captura share TeleVía", "Aforo esperado": aforo1, "Ingreso esperado": ingreso1, "ROI": roi1, "Payback (meses)": payback1},
        {"Escenario": "Subir prepago", "Aforo esperado": base_aforo, "Ingreso esperado": ingreso2, "ROI": roi2, "Payback (meses)": payback2},
        {"Escenario": "Subir postpago", "Aforo esperado": base_aforo, "Ingreso esperado": ingreso3, "ROI": roi3, "Payback (meses)": payback3},
        {"Escenario": "Crecer decenal", "Aforo esperado": aforo4, "Ingreso esperado": ingreso4, "ROI": roi4, "Payback (meses)": payback4},
    ])
    st.dataframe(
        resultados.style.format({
            "Aforo esperado": "{:,.0f}",
            "Ingreso esperado": "${:,.0f}",
            "ROI": "{:.1%}",
            "Payback (meses)": "{:.1f}"
        }),
        use_container_width=True,
        hide_index=True
    )

    if st.button("RECOMIÉNDAME EL MEJOR MOVIMIENTO"):
        mejor = resultados.loc[resultados["ROI"].idxmax()] if not resultados.empty else None
        if mejor is not None:
            st.success(f"**Acción recomendada:** {mejor['Escenario']}")
            st.write(f"**Impacto:** Incremento de ingreso de {fmt_mxn(mejor['Ingreso esperado'] - base_ingreso)}")
            st.write(f"**ROI:** {mejor['ROI']:.1%}")
            st.write(f"**Confianza:** Alta")
        else:
            st.warning("No se pudo determinar una recomendación.")

# ================== Página 4: Portafolio de oportunidades ==================
with tab_plan:
    st.subheader("Portafolio de oportunidades – ¿Dónde gano más con menos esfuerzo?")

    if conc_diag.empty:
        st.info("No hay datos para calcular el portafolio.")
    else:
        # Score dinámico
        max_aforo = conc_diag["aforo"].max()
        max_ingreso = conc_diag["ingreso"].max()
        max_rpc = conc_diag["rpc"].max()
        conc_diag["max_aforo"] = max_aforo
        conc_diag["max_ingreso"] = max_ingreso
        conc_diag["max_rpc"] = max_rpc

        def calcular_score(row):
            vol = row["aforo"] / max_aforo if max_aforo > 0 else 0
            ing = row["ingreso"] / max_ingreso if max_ingreso > 0 else 0
            crec = max(0, min(1, (row["aforo_yoy"] + 0.2) / 0.4)) if not pd.isna(row["aforo_yoy"]) else 0.5
            vol_inv = 1 - min(1, row["volatility"] / 1.5) if not pd.isna(row["volatility"]) else 0.5
            share_inv = 1 - min(1, row["televia_share"] / 0.5) if not pd.isna(row["televia_share"]) else 0.5
            return 100 * (0.3*vol + 0.2*ing + 0.2*crec + 0.15*vol_inv + 0.15*share_inv)

        conc_diag["score_oportunidad"] = conc_diag.apply(calcular_score, axis=1)
        conc_diag["clasificacion"] = conc_diag["score_oportunidad"].apply(
            lambda s: "ejecutar" if s >= 75 else ("apostar" if s >= 50 else ("observar" if s >= 25 else "descartar"))
        )

        portafolio = conc_diag[["concession", "aforo", "televia_share", "rpc", "aforo_yoy", "score_oportunidad", "clasificacion"]].copy()
        portafolio.columns = ["Concesión", "Aforo", "Share TeleVía", "RPC", "Crecimiento", "Score", "Clasificación"]
        portafolio = portafolio.sort_values("Score", ascending=False)

        st.dataframe(
            portafolio.style.format({
                "Aforo": "{:,.0f}",
                "Share TeleVía": "{:.1%}",
                "RPC": "{:.2f}",
                "Crecimiento": "{:.1%}",
                "Score": "{:.1f}"
            }).apply(lambda x: ["background-color: #c6efce" if x["Clasificación"]=="ejecutar" else ("background-color: #ffeb9c" if x["Clasificación"]=="apostar" else ("background-color: #f8cbad" if x["Clasificación"]=="observar" else "background-color: #f2f2f2")) for _ in x], axis=1),
            use_container_width=True,
            height=500,
            hide_index=True
        )

        st.caption("**Score** combina volumen, ingreso, crecimiento, estabilidad y share. Clasificación: ejecutar (alto potencial), apostar (prometedor), observar (incierto), descartar (bajo).")

# ================== Página 5: Seguimiento ==================
with tab_follow:
    st.subheader("Seguimiento – ¿Funcionó?")

    if conc_diag.empty:
        st.info("No hay datos para el seguimiento.")
    else:
        # Índice Salud Concesión
        max_aforo = conc_diag["aforo"].max()
        max_rpc = conc_diag["rpc"].max()
        conc_diag["max_aforo"] = max_aforo
        conc_diag["max_rpc"] = max_rpc

        def indice_salud(row):
            vol = row["aforo"] / max_aforo if max_aforo > 0 else 0
            share = row["televia_share"] if not pd.isna(row["televia_share"]) else 0
            rpc_norm = row["rpc"] / max_rpc if max_rpc > 0 else 0
            estab = 1 - min(1, row["volatility"] / 1.5) if not pd.isna(row["volatility"]) else 0.5
            return 100 * (0.4*vol + 0.3*share + 0.2*rpc_norm + 0.1*estab)

        conc_diag["indice_salud"] = conc_diag.apply(indice_salud, axis=1)

        # Cambio vs año anterior (si existe)
        if "aforo_prev" in conc_diag.columns:
            conc_diag["indice_prev"] = (conc_diag["aforo_prev"] / max_aforo) * 0.4 + (conc_diag["televia_share"] * 0.3) + ((conc_diag["rpc"] / max_rpc) * 0.2) + (0.1 * (1 - conc_diag["volatility"].fillna(0.5)))
            conc_diag["cambio_salud"] = conc_diag["indice_salud"] - conc_diag["indice_prev"]
            st.markdown("#### Top mejora en Índice de Salud")
            top_mejora = conc_diag.nlargest(3, "cambio_salud")[["concession", "cambio_salud"]]
            st.dataframe(top_mejora.rename(columns={"concession": "Concesión", "cambio_salud": "Variación"}).style.format({"Variación": "{:.1f}"}), hide_index=True)
            st.markdown("#### Top deterioro")
            top_deterioro = conc_diag.nsmallest(3, "cambio_salud")[["concession", "cambio_salud"]]
            st.dataframe(top_deterioro.rename(columns={"concession": "Concesión", "cambio_salud": "Variación"}).style.format({"Variación": "{:.1f}"}), hide_index=True)

        st.markdown("#### Índice Salud Concesión (actual)")
        salud_df = conc_diag[["concession", "indice_salud"]].copy().sort_values("indice_salud", ascending=False)
        salud_df.columns = ["Concesión", "Índice Salud"]
        st.dataframe(salud_df.style.format({"Índice Salud": "{:.1f}"}), use_container_width=True, hide_index=True)

        # Próxima acción sugerida
        peor = conc_diag.loc[conc_diag["indice_salud"].idxmin()] if not conc_diag.empty else None
        if peor is not None:
            st.markdown("#### Próxima acción sugerida")
            driver = driver_principal(peor) if "driver_principal" in globals() else "Mejorar eficiencia general"
            st.info(f"**En {peor['concession']}** – {driver or 'Mejorar eficiencia general'}")

    # Mantener gráfica anual y Hoja2
    st.markdown("---")
    st.markdown("#### Evolución anual del portafolio (histórico)")
    if not annual_all.empty:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=annual_all["year"], y=annual_all["aforo"], mode="lines+markers", name="Aforo", line=dict(color=COLORS["televia"], width=4)))
        fig.add_trace(go.Scatter(x=annual_all["year"], y=annual_all["ingreso"], mode="lines+markers", name="Ingreso", line=dict(color=COLORS["info"], width=4), yaxis="y2"))
        fig.update_layout(yaxis=dict(title="Aforo"), yaxis2=dict(title="Ingreso", overlaying="y", side="right"))
        st.plotly_chart(fig, use_container_width=True)

    if not hoja2_df.empty:
        st.markdown("#### Hoja2 – Interoperabilidad")
        st.dataframe(hoja2_df, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# Footer
# -----------------------------------------------------------------------------
st.markdown(
    f"<div style='margin-top:1.2rem; color:{COLORS['muted']}; font-size:0.85rem;'>"
    f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')} · Fuente: {EXCEL_FILE}"
    f"</div>",
    unsafe_allow_html=True,
)
