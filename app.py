# app.py
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
    CHANNEL_GROUPS,
    CHANNEL_ORDER,
    COLORS,
    EXCEL_FILE,
    EXCLUDE_TITLE_PATTERNS,
    MAX_TOP_OPPORTUNITIES,
    MONTHS,
    MONTH_MAP,
    PAGE_TITLES,
    SIMULATION_DEFAULTS,
)

st.set_page_config(
    page_title=APP_TITLE,
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
MONTH_SET = set(MONTHS)


def norm_text(value):
    if value is None:
        return None
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def is_valid_period_label(label):
    if label is None:
        return False
    u = norm_text(label).upper()
    return u in MONTH_SET or u in {"TOTALES", "PROMEDIO"}


def period_to_num(label):
    u = norm_text(label).upper()
    if u in MONTH_MAP:
        return MONTH_MAP[u]
    if u == "TOTALES":
        return 13
    if u == "PROMEDIO":
        return 14
    return None


def channel_group(channel):
    ch = norm_text(channel)
    if not ch:
        return "Otros"
    ch_u = ch.upper()

    for group_name, tokens in CHANNEL_GROUPS.items():
        if group_name == "Otros":
            continue
        if any(tok.upper() in ch_u for tok in tokens):
            return group_name

    if "TOTAL" in ch_u:
        return "Totales"

    return "Otros"


def infer_block_kind(labels):
    labels_u = [norm_text(x).upper() for x in labels if norm_text(x)]
    if any(("PREPAGO" in x) or ("POST-PAGO" in x) or ("POSPAGO" in x) or ("DECENAL" in x) for x in labels_u):
        return "main"
    if any(("TELEVIA" in x) or ("MIS EN MIS" in x) or ("MIS EN SUS" in x) or ("SUS EN MIS" in x) for x in labels_u):
        return "market"
    return "other"


def find_title_above(ws, header_row):
    for rr in range(max(1, header_row - 6), header_row):
        for cc in range(1, ws.max_column + 1):
            v = ws.cell(rr, cc).value
            if isinstance(v, str):
                t = norm_text(v)
                if t and t.startswith("Aforo e Ingreso -"):
                    return t.replace("Aforo e Ingreso -", "").strip()
    return None


def collect_block_labels(ws, header_row, fecha_col):
    labels = []
    col = fecha_col + 2
    while col <= ws.max_column:
        label = norm_text(ws.cell(header_row - 1, col).value) if header_row - 1 >= 1 else None
        if not label:
            break
        labels.append(label)
        col += 2
    return labels


def parse_hoja2(ws):
    if ws is None:
        return pd.DataFrame()

    records = []
    for r in range(4, ws.max_row + 1):
        period = norm_text(ws.cell(r, 1).value)
        if not period:
            continue

        values = [ws.cell(r, c).value for c in range(2, 8)]
        if all(v is None for v in values):
            continue

        records.append(
            {
                "periodo": period,
                "mis_en_sus_aforo": ws.cell(r, 2).value,
                "mis_en_sus_ingreso": ws.cell(r, 3).value,
                "mis_en_mis_aforo": ws.cell(r, 4).value,
                "mis_en_mis_ingreso": ws.cell(r, 5).value,
                "sus_en_mis_aforo": ws.cell(r, 6).value,
                "sus_en_mis_ingreso": ws.cell(r, 7).value,
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
        year_match = re.search(r"(\d{4})", sheet_name)
        sheet_year = int(year_match.group(1)) if year_match else None

        fecha_cells = []
        for rr in range(1, ws.max_row + 1):
            for cc in range(1, ws.max_column + 1):
                if norm_text(ws.cell(rr, cc).value) == "FECHA":
                    fecha_cells.append((rr, cc))

        for header_row, fecha_col in fecha_cells:
            labels = collect_block_labels(ws, header_row, fecha_col)
            if not labels:
                continue

            concession = find_title_above(ws, header_row) or sheet_name.replace("Aforo e Ingreso ", "").strip()
            block_kind = infer_block_kind(labels)

            for rr in range(header_row + 2, min(ws.max_row, header_row + 18) + 1):
                period = norm_text(ws.cell(rr, fecha_col).value)
                if not is_valid_period_label(period):
                    continue

                year_val = ws.cell(rr, fecha_col + 1).value
                if not isinstance(year_val, int):
                    year_val = sheet_year

                for idx, channel in enumerate(labels):
                    aforo_col = fecha_col + 2 + idx * 2
                    ingreso_col = aforo_col + 1
                    if aforo_col > ws.max_column:
                        break

                    aforo = ws.cell(rr, aforo_col).value
                    ingreso = ws.cell(rr, ingreso_col).value

                    if aforo is None and ingreso is None:
                        continue

                    records.append(
                        {
                            "sheet": sheet_name,
                            "year": int(year_val),
                            "concession": concession,
                            "header_row": header_row,
                            "fecha_col": fecha_col,
                            "period": norm_text(period).upper(),
                            "period_num": period_to_num(period),
                            "period_type": "month" if norm_text(period).upper() in MONTH_SET else "summary",
                            "channel": channel,
                            "channel_group": channel_group(channel),
                            "aforo": float(aforo) if isinstance(aforo, (int, float, np.integer, np.floating)) and not pd.isna(aforo) else 0.0,
                            "ingreso": float(ingreso) if isinstance(ingreso, (int, float, np.integer, np.floating)) and not pd.isna(ingreso) else 0.0,
                            "block_kind": block_kind,
                        }
                    )

    raw_df = pd.DataFrame(records)

    if not raw_df.empty:
        raw_df["concession"] = raw_df["concession"].astype(str).str.strip()
        raw_df["channel"] = raw_df["channel"].astype(str).str.strip()
        raw_df["channel_group"] = raw_df["channel_group"].astype(str).str.strip()

    hoja2_df = parse_hoja2(wb["Hoja2"]) if "Hoja2" in wb.sheetnames else pd.DataFrame()
    return raw_df, hoja2_df


@st.cache_data(show_spinner=False)
def load_data(file_path):
    return parse_workbook(file_path)


def safe_pct_change(current, previous):
    if previous in (None, 0) or pd.isna(previous):
        return np.nan
    return (current - previous) / previous


def fmt_mxn(value):
    if pd.isna(value):
        return "-"
    return f"${value:,.0f}"


def fmt_num(value):
    if pd.isna(value):
        return "-"
    return f"{value:,.0f}"


def make_header():
    st.markdown(
        f"<div style='background: linear-gradient(135deg, {COLORS['navy']} 0%, {COLORS['blue']} 100%);"
        f"padding: 16px 20px; border-radius: 14px; color: white; margin-bottom: 12px;'>"
        f"<h1 style='margin:0; font-size: 2rem;'>{APP_TITLE}</h1>"
        f"<div style='opacity:0.86; margin-top: 4px;'>{APP_SUBTITLE}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )


def make_metric_row(metrics):
    cols = st.columns(len(metrics))
    for col, metric in zip(cols, metrics):
        with col:
            st.metric(**metric)


def build_year_month_company(df_months):
    return (
        df_months.groupby(["year", "month_num"], as_index=False)
        .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
        .sort_values(["year", "month_num"])
    )


def build_concession_summary(df_months):
    conc = (
        df_months.groupby(["year", "concession"], as_index=False)
        .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    )
    conc["rpc"] = conc["ingreso"] / conc["aforo"].replace(0, np.nan)

    vol = (
        df_months.groupby(["year", "concession"])["aforo"]
        .agg(["mean", "std"])
        .reset_index()
    )
    vol["volatility"] = vol["std"] / vol["mean"].replace(0, np.nan)
    return conc.merge(vol[["year", "concession", "volatility"]], on=["year", "concession"], how="left")


def build_channel_mix(df_months):
    return (
        df_months[df_months["channel_group"] != "Totales"]
        .groupby(["year", "month_num", "channel_group"], as_index=False)
        .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    )


def build_heatmap(df_months, year_value):
    sub = df_months[(df_months["year"] == year_value) & (df_months["channel_group"] != "Totales")].copy()
    if sub.empty:
        return pd.DataFrame()

    pivot = (
        sub.groupby(["concession", "channel_group"], as_index=False)
        .agg(aforo=("aforo", "sum"))
        .pivot(index="concession", columns="channel_group", values="aforo")
        .fillna(0)
    )
    available_cols = [c for c in CHANNEL_ORDER if c in pivot.columns]
    pivot = pivot.reindex(columns=available_cols, fill_value=0)
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]
    return pivot


def score_opportunity(row, medians):
    volume_component = row["aforo"] / medians["aforo_median"] if medians["aforo_median"] else 0
    rpc_component = medians["rpc_median"] / row["rpc"] if pd.notna(row["rpc"]) and row["rpc"] not in (0, np.nan) else 1.5
    yoy_component = max(0, -row["aforo_yoy"]) if pd.notna(row["aforo_yoy"]) else 0
    vol_component = row["volatility"] if pd.notna(row["volatility"]) else 0
    return 0.35 * volume_component + 0.25 * rpc_component + 0.25 * yoy_component + 0.15 * vol_component


def recommendation_for_row(row, medians):
    prepago_share = row.get("prepago_share", 0)
    postpago_share = row.get("postpago_share", 0)
    decenal_share = row.get("decenal_share", 0)

    if pd.notna(row.get("aforo_yoy")) and row["aforo_yoy"] < 0 and row["aforo"] >= medians["aforo_median"]:
        return ACTION_LIBRARY["volume_recovery"]["label"], "Volumen relevante con caída vs año previo."
    if pd.notna(row.get("rpc")) and row["rpc"] < medians["rpc_median"]:
        return ACTION_LIBRARY["monetization"]["label"], "Ingreso por cruce por debajo del benchmark."
    if pd.notna(row.get("volatility")) and row["volatility"] > medians["vol_median"]:
        return ACTION_LIBRARY["stabilization"]["label"], "Alta variabilidad mensual."
    if prepago_share < 0.15 and row["aforo"] >= medians["aforo_median"]:
        return ACTION_LIBRARY["channel_capture"]["label"], "Baja captura de prepago sobre una base relevante."
    if decenal_share > 0.20:
        return ACTION_LIBRARY["premium_mix"]["label"], "Mix con espacio para monetización premium."
    if postpago_share < 0.15:
        return ACTION_LIBRARY["channel_capture"]["label"], "Post-pago con oportunidad de profundizar captura."
    return "Seguimiento", "Mantener monitoreo y buscar mejoras incrementales."


# -----------------------------------------------------------------------------
# Load workbook
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

# -----------------------------------------------------------------------------
# Sidebar filters
# -----------------------------------------------------------------------------
all_years = sorted(raw_df["year"].dropna().astype(int).unique().tolist())
all_concessions = sorted(raw_df["concession"].dropna().astype(str).unique().tolist())

with st.sidebar:
    st.markdown("### Filtros globales")
    scope_mode = st.radio("Alcance", ["Concesiones operativas", "Todos los bloques"], index=0)
    year_choice = st.selectbox("Año", ["Todos"] + all_years, index=0)
    concession_choice = st.selectbox("Concesión", ["Todas"] + all_concessions, index=0)

    st.markdown("---")
    st.markdown("### Supuestos de simulación")
    sim_volume_growth = st.slider("Crecimiento de volumen (%)", 0.0, 20.0, SIMULATION_DEFAULTS["volume_growth_pct"], 0.1)
    sim_share_capture = st.slider("Captura de share (pp)", 0.0, 10.0, SIMULATION_DEFAULTS["share_capture_pp"], 0.1)
    sim_rpc_uplift = st.slider("Mejora ingreso/cruce (%)", 0.0, 20.0, SIMULATION_DEFAULTS["rpc_uplift_pct"], 0.1)

# -----------------------------------------------------------------------------
# Scope selection
# -----------------------------------------------------------------------------
scope_df = raw_df.copy()

if scope_mode == "Concesiones operativas":
    scope_df = scope_df[~scope_df["concession"].str.contains("|".join(EXCLUDE_TITLE_PATTERNS), case=False, regex=True, na=False)].copy()

if year_choice != "Todos":
    scope_df = scope_df[scope_df["year"] == int(year_choice)].copy()

if concession_choice != "Todas":
    scope_df = scope_df[scope_df["concession"] == concession_choice].copy()

months_df = scope_df[(scope_df["period_type"] == "month") & (scope_df["channel_group"] != "Totales")].copy()
if months_df.empty:
    st.error("No hay datos en el alcance actual. Ajusta los filtros laterales.")
    st.stop()

annual_company = build_year_month_company(months_df)
annual_concessions = build_concession_summary(months_df)
channel_mix = build_channel_mix(months_df)

focus_year = int(year_choice) if year_choice != "Todos" else int(months_df["year"].max())
focus_year_df = months_df[months_df["year"] == focus_year].copy()
previous_year_df = months_df[months_df["year"] == focus_year - 1].copy()

company_focus = annual_company[annual_company["year"] == focus_year].copy()
focus_month = int(company_focus.loc[company_focus["aforo"] > 0, "month_num"].max()) if not company_focus.empty else 12
company_current = focus_year_df[focus_year_df["month_num"] <= focus_month].copy()
company_prev = previous_year_df[previous_year_df["month_num"] <= focus_month].copy() if not previous_year_df.empty else pd.DataFrame()

# Derived metrics
total_aforo = company_current["aforo"].sum()
total_ingreso = company_current["ingreso"].sum()
rpc = total_ingreso / total_aforo if total_aforo else np.nan

prev_aforo = company_prev["aforo"].sum() if not company_prev.empty else np.nan
prev_ingreso = company_prev["ingreso"].sum() if not company_prev.empty else np.nan
aforo_yoy = safe_pct_change(total_aforo, prev_aforo)
ingreso_yoy = safe_pct_change(total_ingreso, prev_ingreso)

focus_concession_summary = annual_concessions[annual_concessions["year"] == focus_year].copy()
if not focus_concession_summary.empty:
    total_focus_aforo = focus_concession_summary["aforo"].sum()
    focus_concession_summary["aforo_share"] = focus_concession_summary["aforo"] / total_focus_aforo if total_focus_aforo else np.nan

prev_conc = annual_concessions[annual_concessions["year"] == focus_year - 1][["concession", "aforo", "ingreso"]].rename(
    columns={"aforo": "aforo_prev", "ingreso": "ingreso_prev"}
)
opp = focus_concession_summary.merge(prev_conc, on="concession", how="left") if not focus_concession_summary.empty else pd.DataFrame()

if not opp.empty:
    opp["aforo_yoy"] = opp.apply(lambda r: safe_pct_change(r["aforo"], r["aforo_prev"]), axis=1)
    opp["ingreso_yoy"] = opp.apply(lambda r: safe_pct_change(r["ingreso"], r["ingreso_prev"]), axis=1)

    prepago = (
        focus_year_df[focus_year_df["channel_group"] == "Prepago"]
        .groupby("concession", as_index=False)
        .agg(aforo=("aforo", "sum"))
        .rename(columns={"aforo": "prepago_aforo"})
    )
    postpago = (
        focus_year_df[focus_year_df["channel_group"] == "Post-pago"]
        .groupby("concession", as_index=False)
        .agg(aforo=("aforo", "sum"))
        .rename(columns={"aforo": "postpago_aforo"})
    )
    decenal = (
        focus_year_df[focus_year_df["channel_group"] == "Decenal"]
        .groupby("concession", as_index=False)
        .agg(aforo=("aforo", "sum"))
        .rename(columns={"aforo": "decenal_aforo"})
    )

    opp = opp.merge(prepago, on="concession", how="left").merge(postpago, on="concession", how="left").merge(decenal, on="concession", how="left")
    for col in ["prepago_aforo", "postpago_aforo", "decenal_aforo"]:
        opp[col] = opp[col].fillna(0)

    opp["prepago_share"] = opp["prepago_aforo"] / opp["aforo"].replace(0, np.nan)
    opp["postpago_share"] = opp["postpago_aforo"] / opp["aforo"].replace(0, np.nan)
    opp["decenal_share"] = opp["decenal_aforo"] / opp["aforo"].replace(0, np.nan)

    medians = {
        "aforo_median": float(opp["aforo"].median()) if not opp.empty else 0,
        "rpc_median": float(opp["rpc"].median()) if not opp.empty else 0,
        "vol_median": float(opp["volatility"].median()) if "volatility" in opp else 0,
    }

    opp["priority_score"] = opp.apply(lambda r: score_opportunity(r, medians), axis=1)
    recs = opp.apply(lambda r: recommendation_for_row(r, medians), axis=1, result_type="expand")
    opp["accion"] = recs[0]
    opp["motivo"] = recs[1]
    opp = opp.sort_values("priority_score", ascending=False)

annual_portfolio = (
    months_df.groupby("year", as_index=False)
    .agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
    .sort_values("year")
)
annual_portfolio["rpc"] = annual_portfolio["ingreso"] / annual_portfolio["aforo"].replace(0, np.nan)
annual_portfolio["aforo_yoy"] = annual_portfolio["aforo"].pct_change()
annual_portfolio["ingreso_yoy"] = annual_portfolio["ingreso"].pct_change()

# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
make_header()
st.caption(
    f"Archivo fuente: {EXCEL_FILE} · Rango analizado: {min(all_years)}–{max(all_years)} · "
    f"Alcance activo: {'operativo' if scope_mode == 'Concesiones operativas' else 'todos los bloques'}"
)

# -----------------------------------------------------------------------------
# Tabs
# -----------------------------------------------------------------------------
tab_home, tab_diag, tab_sim, tab_plan, tab_follow = st.tabs([
    PAGE_TITLES["home"],
    PAGE_TITLES["diagnostico"],
    PAGE_TITLES["simulacion"],
    PAGE_TITLES["plan"],
    PAGE_TITLES["seguimiento"],
])

# -----------------------------------------------------------------------------
# 1) Executive Summary
# -----------------------------------------------------------------------------
with tab_home:
    st.subheader("Executive Summary")

    top_concession_name = "-"
    top_concession_aforo = np.nan
    best_rpc_concession = "-"
    worst_rpc_concession = "-"

    if not focus_concession_summary.empty:
        top_row = focus_concession_summary.sort_values("aforo", ascending=False).iloc[0]
        top_concession_name = top_row["concession"]
        top_concession_aforo = top_row["aforo"]

        best_rpc_concession = focus_concession_summary.sort_values("rpc", ascending=False).iloc[0]["concession"]
        worst_rpc_concession = focus_concession_summary.sort_values("rpc", ascending=True).iloc[0]["concession"]

    make_metric_row([
        {"label": "Aforo total", "value": fmt_num(total_aforo), "delta": f"{aforo_yoy*100:+.1f}%" if pd.notna(aforo_yoy) else "n/a"},
        {"label": "Ingreso total", "value": fmt_mxn(total_ingreso), "delta": f"{ingreso_yoy*100:+.1f}%" if pd.notna(ingreso_yoy) else "n/a"},
        {"label": "Ingreso / cruce", "value": f"{rpc:.2f}" if pd.notna(rpc) else "-", "delta": f"{focus_year} YTD" if focus_month < 12 else f"{focus_year}"},
        {"label": "Mejor concesión", "value": top_concession_name, "delta": fmt_num(top_concession_aforo)},
    ])

    st.markdown("#### Lectura ejecutiva")
    insight_cols = st.columns(4)
    with insight_cols[0]:
        st.info(f"**Mayor volumen:** {top_concession_name}")
    with insight_cols[1]:
        st.warning(f"**RPC más alto:** {best_rpc_concession}")
    with insight_cols[2]:
        st.error(f"**RPC más bajo:** {worst_rpc_concession}")
    with insight_cols[3]:
        if not hoja2_df.empty:
            st.success(f"**Hoja2 disponible:** {len(hoja2_df)} periodos")
        else:
            st.success("**Hoja2 disponible**")

    c1, c2 = st.columns([1.2, 0.8])

    with c1:
        st.markdown("##### Tendencia mensual de aforo")
        trend = company_focus.groupby("month_num", as_index=False).agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum"))
        prev_trend = previous_year_df.groupby("month_num", as_index=False).agg(aforo=("aforo", "sum"), ingreso=("ingreso", "sum")) if not previous_year_df.empty else pd.DataFrame(columns=["month_num", "aforo", "ingreso"])

        fig = go.Figure()
        if not prev_trend.empty:
            fig.add_trace(go.Scatter(
                x=prev_trend["month_num"],
                y=prev_trend["aforo"],
                mode="lines+markers",
                name=str(focus_year - 1),
                line=dict(color=COLORS["muted"], width=3),
            ))
        fig.add_trace(go.Scatter(
            x=trend["month_num"],
            y=trend["aforo"],
            mode="lines+markers",
            name=str(focus_year),
            line=dict(color=COLORS["televia"], width=4),
        ))
        fig.update_layout(
            height=360,
            xaxis=dict(title="", tickmode="array", tickvals=list(range(1, 13)), ticktext=MONTHS),
            yaxis=dict(title="Aforo", gridcolor=COLORS["grid"]),
            plot_bgcolor="white",
            paper_bgcolor="white",
            legend_title_text="",
            margin=dict(l=10, r=10, t=20, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown("##### Mix de canales")
        mix_focus = channel_mix[channel_mix["year"] == focus_year].copy()
        mix_focus = mix_focus[mix_focus["channel_group"] != "Totales"]
        mix_focus["month_name"] = mix_focus["month_num"].map(lambda x: MONTHS[int(x)-1] if int(x) in range(1, 13) else str(x))
        fig = px.bar(
            mix_focus,
            x="month_name",
            y="aforo",
            color="channel_group",
            barmode="stack",
            category_orders={"month_name": MONTHS},
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig.update_layout(
            height=360,
            xaxis_title="",
            yaxis_title="Aforo",
            legend_title_text="Canal",
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=10, r=10, t=20, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### Top concesiones por volumen")
    if not focus_concession_summary.empty:
        top_tbl = focus_concession_summary.sort_values("aforo", ascending=False).head(10).copy()
        top_tbl["rpc"] = top_tbl["rpc"].round(2)
        st.dataframe(
            top_tbl[["concession", "aforo", "ingreso", "rpc"]].rename(columns={
                "concession": "Concesión",
                "aforo": "Aforo",
                "ingreso": "Ingreso",
                "rpc": "Ingreso / cruce",
            }),
            use_container_width=True,
            hide_index=True,
        )

    if not hoja2_df.empty:
        st.markdown("##### Hoja2 · resumen complementario")
        st.dataframe(hoja2_df, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# 2) Diagnostic
# -----------------------------------------------------------------------------
with tab_diag:
    st.subheader("Diagnóstico")

    diag_year = st.selectbox("Año para diagnóstico", all_years, index=len(all_years)-1, key="diag_year")
    diag_df = months_df[months_df["year"] == int(diag_year)].copy()

    diag_conc = build_concession_summary(diag_df).sort_values("aforo", ascending=False)

    if diag_conc.empty:
        st.info("No hay datos para el año seleccionado.")
    else:
        previous_diag = annual_concessions[annual_concessions["year"] == int(diag_year) - 1][["concession", "aforo", "ingreso"]].rename(
            columns={"aforo": "aforo_prev", "ingreso": "ingreso_prev"}
        )
        diag_conc = diag_conc.merge(previous_diag, on="concession", how="left")
        diag_conc["aforo_yoy"] = diag_conc.apply(lambda r: safe_pct_change(r["aforo"], r["aforo_prev"]), axis=1)
        diag_conc["ingreso_yoy"] = diag_conc.apply(lambda r: safe_pct_change(r["ingreso"], r["ingreso_prev"]), axis=1)

        prepago_by_conc = diag_df[diag_df["channel_group"] == "Prepago"].groupby("concession")["aforo"].sum()
        postpago_by_conc = diag_df[diag_df["channel_group"] == "Post-pago"].groupby("concession")["aforo"].sum()
        decenal_by_conc = diag_df[diag_df["channel_group"] == "Decenal"].groupby("concession")["aforo"].sum()

        diag_conc["prepago_aforo"] = diag_conc["concession"].map(prepago_by_conc).fillna(0)
        diag_conc["postpago_aforo"] = diag_conc["concession"].map(postpago_by_conc).fillna(0)
        diag_conc["decenal_aforo"] = diag_conc["concession"].map(decenal_by_conc).fillna(0)

        diag_conc["prepago_share"] = diag_conc["prepago_aforo"] / diag_conc["aforo"].replace(0, np.nan)
        diag_conc["postpago_share"] = diag_conc["postpago_aforo"] / diag_conc["aforo"].replace(0, np.nan)
        diag_conc["decenal_share"] = diag_conc["decenal_aforo"] / diag_conc["aforo"].replace(0, np.nan)

        diag_summary = {
            "aforo_max": diag_conc["aforo"].max(),
            "aforo_min": diag_conc["aforo"].min(),
            "aforo_median": diag_conc["aforo"].median(),
            "rpc_median": diag_conc["rpc"].median(),
            "vol_median": diag_conc["volatility"].median(),
        }

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Concesiones", f"{diag_conc['concession'].nunique():,}")
        c2.metric("Aforo max", fmt_num(diag_summary["aforo_max"]))
        c3.metric("Aforo min", fmt_num(diag_summary["aforo_min"]))
        c4.metric("RPC mediano", f"{diag_summary['rpc_median']:.2f}")

        diag_cols = st.columns([1.05, 0.95])
        with diag_cols[0]:
            st.markdown("##### Ranking de aforo")
            fig = px.bar(
                diag_conc.head(15),
                x="concession",
                y="aforo",
                color="rpc",
                color_continuous_scale="RdYlGn",
                text_auto=".0f",
            )
            fig.update_layout(
                height=420,
                xaxis_title="",
                yaxis_title="Aforo",
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=10, r=10, t=20, b=10),
            )
            st.plotly_chart(fig, use_container_width=True)

        with diag_cols[1]:
            st.markdown("##### Variación vs año anterior")
            diag_conc["aforo_yoy_pct"] = diag_conc["aforo_yoy"] * 100
            fig = px.bar(
                diag_conc.sort_values("aforo_yoy_pct").tail(15),
                x="aforo_yoy_pct",
                y="concession",
                orientation="h",
                color="aforo_yoy_pct",
                color_continuous_scale="RdYlGn",
                text_auto=".1f",
            )
            fig.update_layout(
                height=420,
                xaxis_title="% YoY",
                yaxis_title="",
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=10, r=10, t=20, b=10),
            )
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### Heatmap de mix por concesión")
        heatmap = build_heatmap(diag_df, int(diag_year))
        if not heatmap.empty:
            heatmap_pct = heatmap.div(heatmap.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
            fig = go.Figure(
                data=go.Heatmap(
                    z=heatmap_pct.values,
                    x=heatmap_pct.columns.tolist(),
                    y=heatmap_pct.index.tolist(),
                    colorscale="RdYlGn",
                    zmin=0,
                    zmax=max(float(heatmap_pct.values.max()), 1e-6),
                    colorbar=dict(title="Mix %"),
                    hovertemplate="Concesión=%{y}<br>Canal=%{x}<br>Mix=%{z:.1%}<extra></extra>",
                )
            )
            fig.update_layout(height=max(420, 24 * len(heatmap_pct) + 160), plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### Tabla diagnóstica")
        show_cols = ["concession", "aforo", "ingreso", "rpc", "volatility", "aforo_yoy", "ingreso_yoy", "prepago_share", "postpago_share", "decenal_share"]
        display_diag = diag_conc[show_cols].copy()
        display_diag.columns = [
            "Concesión", "Aforo", "Ingreso", "Ingreso/cruce", "Volatilidad",
            "YoY aforo", "YoY ingreso", "Prepago %", "Post-pago %", "Decenal %",
        ]
        st.dataframe(
            display_diag.style.format({
                "Aforo": "{:,.0f}",
                "Ingreso": "${:,.0f}",
                "Ingreso/cruce": "{:,.2f}",
                "Volatilidad": "{:.2f}",
                "YoY aforo": "{:.1%}",
                "YoY ingreso": "{:.1%}",
                "Prepago %": "{:.1%}",
                "Post-pago %": "{:.1%}",
                "Decenal %": "{:.1%}",
            }),
            use_container_width=True,
            hide_index=True,
        )

# -----------------------------------------------------------------------------
# 3) Simulation
# -----------------------------------------------------------------------------
with tab_sim:
    st.subheader("Simulación")

    sim_scope_options = ["Portafolio completo"] + all_concessions
    sim_target = st.selectbox("Base de simulación", sim_scope_options, index=0, key="sim_target")

    if sim_target == "Portafolio completo":
        sim_base = company_current.copy()
        sim_label = f"{focus_year} YTD"
    else:
        sim_base = company_current[company_current["concession"] == sim_target].copy()
        sim_label = sim_target

    if sim_base.empty:
        st.info("No hay datos suficientes para simular esa base.")
    else:
        base_aforo = sim_base["aforo"].sum()
        base_ingreso = sim_base["ingreso"].sum()
        base_rpc = base_ingreso / base_aforo if base_aforo else np.nan

        colA, colB, colC = st.columns(3)
        with colA:
            st.markdown("##### Supuesto 1")
            st.write(f"Crecimiento de volumen: **{sim_volume_growth:.1f}%**")
        with colB:
            st.markdown("##### Supuesto 2")
            st.write(f"Captura de share: **{sim_share_capture:.1f} pp**")
        with colC:
            st.markdown("##### Supuesto 3")
            st.write(f"Mejora ingreso/cruce: **{sim_rpc_uplift:.1f}%**")

        extra_aforo = base_aforo * (sim_volume_growth / 100.0)
        capture_aforo = base_aforo * (sim_share_capture / 100.0)
        scenario_aforo = base_aforo + extra_aforo + capture_aforo
        scenario_rpc = base_rpc * (1 + sim_rpc_uplift / 100.0) if pd.notna(base_rpc) else np.nan
        scenario_ingreso = scenario_aforo * scenario_rpc if pd.notna(scenario_rpc) else np.nan

        delta_aforo = scenario_aforo - base_aforo
        delta_ingreso = scenario_ingreso - base_ingreso if pd.notna(scenario_ingreso) else np.nan

        kpi_cols = st.columns(4)
        kpi_cols[0].metric("Base aforo", fmt_num(base_aforo))
        kpi_cols[1].metric("Escenario aforo", fmt_num(scenario_aforo), f"+{sim_volume_growth:.1f}% + {sim_share_capture:.1f}pp")
        kpi_cols[2].metric("Base ingreso", fmt_mxn(base_ingreso))
        kpi_cols[3].metric("Escenario ingreso", fmt_mxn(scenario_ingreso), f"+{sim_rpc_uplift:.1f}% RPC")

        st.markdown("##### Resumen del escenario")
        scen_df = pd.DataFrame([
            {"Métrica": "Base", "Aforo": base_aforo, "Ingreso": base_ingreso, "Ingreso/cruce": base_rpc},
            {"Métrica": "Escenario", "Aforo": scenario_aforo, "Ingreso": scenario_ingreso, "Ingreso/cruce": scenario_rpc},
            {"Métrica": "Delta", "Aforo": delta_aforo, "Ingreso": delta_ingreso, "Ingreso/cruce": scenario_rpc - base_rpc if pd.notna(scenario_rpc) else np.nan},
        ])
        st.dataframe(
            scen_df.style.format({
                "Aforo": "{:,.0f}",
                "Ingreso": "${:,.0f}",
                "Ingreso/cruce": "{:,.2f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        fig = go.Figure()
        fig.add_bar(name="Base", x=["Aforo", "Ingreso"], y=[base_aforo, base_ingreso], marker_color=COLORS["muted"])
        fig.add_bar(name="Escenario", x=["Aforo", "Ingreso"], y=[scenario_aforo, scenario_ingreso], marker_color=COLORS["televia"])
        fig.update_layout(
            barmode="group",
            height=420,
            plot_bgcolor="white",
            paper_bgcolor="white",
            yaxis=dict(gridcolor=COLORS["grid"]),
            legend_title_text="",
            margin=dict(l=10, r=10, t=20, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### Lectura de palancas")
        st.write(f"- **{sim_target}** puede sumar **{fmt_num(delta_aforo)} cruces** bajo el supuesto actual.")
        st.write(f"- El ingreso estimado subiría en **{fmt_mxn(delta_ingreso)}** si el ingreso/cruce mejora al ritmo supuesto.")
        st.write("- Usa este bloque para probar escenarios antes de pedir presupuestos o compromisos comerciales.")

# -----------------------------------------------------------------------------
# 4) Action plan
# -----------------------------------------------------------------------------
with tab_plan:
    st.subheader("Plan de acción")

    if opp.empty:
        st.info("No hay datos suficientes para construir un plan de acción.")
    else:
        top_opp = opp.head(MAX_TOP_OPPORTUNITIES).copy()
        top_opp["priority_score"] = top_opp["priority_score"].round(2)

        st.markdown("##### Oportunidades prioritarias")
        action_table = top_opp[["concession", "aforo", "rpc", "aforo_yoy", "volatility", "priority_score", "accion", "motivo"]].copy()
        action_table.columns = ["Concesión", "Aforo", "Ingreso/cruce", "YoY aforo", "Volatilidad", "Prioridad", "Acción", "Motivo"]
        st.dataframe(
            action_table.style.format({
                "Aforo": "{:,.0f}",
                "Ingreso/cruce": "{:,.2f}",
                "YoY aforo": "{:.1%}",
                "Volatilidad": "{:.2f}",
                "Prioridad": "{:.2f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        c1, c2 = st.columns([1.15, 0.85])

        with c1:
            st.markdown("##### Priorización")
            fig = px.bar(
                top_opp.sort_values("priority_score", ascending=True),
                x="priority_score",
                y="concession",
                orientation="h",
                color="priority_score",
                color_continuous_scale="RdYlGn_r",
                text_auto=".2f",
            )
            fig.update_layout(
                height=max(420, 28 * len(top_opp) + 120),
                xaxis_title="Prioridad",
                yaxis_title="",
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=10, r=10, t=20, b=10),
            )
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            st.markdown("##### Playbook sugerido")
            st.markdown(
                """
                <div style='background:white;border:1px solid #E2E8F0;border-radius:14px;padding:14px 16px;'>
                <ul style='margin:0;padding-left:18px;'>
                    <li><b>Recuperación de volumen</b>: donde la concesión cae pero sigue siendo relevante.</li>
                    <li><b>Monetización</b>: cuando el ingreso/cruce está por debajo del benchmark.</li>
                    <li><b>Estabilización</b>: cuando la variabilidad mensual es alta.</li>
                    <li><b>Captura de canal</b>: cuando hay poco peso de prepago/post-pago/decenal.</li>
                    <li><b>Mix premium</b>: cuando decenal o canales de alto valor pueden crecer.</li>
                </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.write("")
            st.caption("Ajusta la lectura con los sliders de simulación en la barra lateral.")

# -----------------------------------------------------------------------------
# 5) Follow-up
# -----------------------------------------------------------------------------
with tab_follow:
    st.subheader("Seguimiento")

    left, right = st.columns([1.1, 0.9])

    with left:
        st.markdown("##### Evolución anual del portafolio")
        annual_plot = annual_portfolio.copy()

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=annual_plot["year"],
            y=annual_plot["aforo"],
            mode="lines+markers",
            name="Aforo",
            line=dict(color=COLORS["televia"], width=4),
        ))
        fig.add_trace(go.Scatter(
            x=annual_plot["year"],
            y=annual_plot["ingreso"],
            mode="lines+markers",
            name="Ingreso",
            line=dict(color=COLORS["blue"], width=4),
            yaxis="y2",
        ))
        fig.update_layout(
            height=420,
            xaxis=dict(title="Año"),
            yaxis=dict(title="Aforo", gridcolor=COLORS["grid"]),
            yaxis2=dict(title="Ingreso", overlaying="y", side="right", showgrid=False),
            plot_bgcolor="white",
            paper_bgcolor="white",
            legend_title_text="",
            margin=dict(l=10, r=10, t=20, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.markdown("##### KPI anual")
        annual_display = annual_portfolio.copy()
        annual_display["aforo_yoy"] = annual_display["aforo"].pct_change()
        annual_display["ingreso_yoy"] = annual_display["ingreso"].pct_change()
        annual_display["rpc"] = annual_display["rpc"].round(2)

        st.dataframe(
            annual_display.rename(columns={
                "year": "Año",
                "aforo": "Aforo",
                "ingreso": "Ingreso",
                "rpc": "Ingreso/cruce",
                "aforo_yoy": "YoY aforo",
                "ingreso_yoy": "YoY ingreso",
            }).style.format({
                "Aforo": "{:,.0f}",
                "Ingreso": "${:,.0f}",
                "Ingreso/cruce": "{:,.2f}",
                "YoY aforo": "{:.1%}",
                "YoY ingreso": "{:.1%}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("##### Lectura para seguimiento")
        if len(annual_display) >= 2:
            last = annual_display.iloc[-1]
            prev = annual_display.iloc[-2]
            st.metric("Variación aforo último año", f"{safe_pct_change(last['aforo'], prev['aforo'])*100:+.1f}%")
            st.metric("Variación ingreso último año", f"{safe_pct_change(last['ingreso'], prev['ingreso'])*100:+.1f}%")

    st.markdown("##### Próximo paso")
    st.info(
        "Cuando agregues otro Excel con la misma estructura, el flujo ideal es duplicarlo en el repositorio "
        "y volver a desplegar para comparar trimestres o años completos sin cambiar la lógica analítica."
    )

# -----------------------------------------------------------------------------
# Footer
# -----------------------------------------------------------------------------
st.markdown(
    f"<div style='margin-top:1.2rem; color:{COLORS['muted']}; font-size:0.85rem;'>"
    f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')} · Fuente: {EXCEL_FILE}"
    f"</div>",
    unsafe_allow_html=True,
)
