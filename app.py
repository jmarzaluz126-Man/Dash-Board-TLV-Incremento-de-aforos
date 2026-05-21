import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
import re

# ===== CONFIGURACIÓN =====
EXCEL_FILE = "dashboard_televia.xlsx"
MONTHS = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
          "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
MONTH_MAP = {m: i+1 for i, m in enumerate(MONTHS)}

COLORS = {
    'televia': '#2E7D32',
    'pase': '#999999',
    'info': '#1976D2',
    'warning': '#FF9800',
    'danger': '#D32F2F',
    'muted': '#CCCCCC',
    'grid': '#E8E8E8',
}

# ===== STREAMLIT CONFIG =====
st.set_page_config(
    page_title="Dashboard TeleVía 2026",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ===== HELPERS =====
def normalize_text(value):
    if value is None:
        return None
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text or None

def safe_divide(num, denom):
    if denom in (None, 0) or pd.isna(denom):
        return np.nan
    return num / denom

def fmt_num(value):
    if pd.isna(value) or value is None:
        return "-"
    return f"{value:,.0f}"

def fmt_mxn(value):
    if pd.isna(value) or value is None:
        return "-"
    return f"${value:,.0f}"

def fmt_pct(value):
    if pd.isna(value) or value is None:
        return "-"
    return f"{value:.1%}"

# ===== PARSER DEL EXCEL =====
@st.cache_data(show_spinner=False)
def load_excel_data(file_path):
    wb = load_workbook(file_path, data_only=True)
    all_records = []
    
    for sheet_name in wb.sheetnames:
        if sheet_name == "Hoja2":
            continue
        
        ws = wb[sheet_name]
        sheet_year = None
        
        match = re.search(r'(\d{4})', sheet_name)
        if match:
            sheet_year = int(match.group(1))
        
        if sheet_year is None:
            continue
        
        concession = normalize_text(ws.cell(3, 2).value)
        if not concession:
            continue
        
        concession = str(ws.cell(3, 2).value).replace("Aforo e Ingreso -", "").strip()
        
        channels = {}
        col = 4
        while col <= ws.max_column:
            channel_name = normalize_text(ws.cell(5, col).value)
            if not channel_name:
                break
            channels[channel_name] = col
            col += 2
        
        for row in range(8, ws.max_row + 1):
            period = normalize_text(ws.cell(row, 2).value)
            if not period:
                continue
            
            period_upper = period.upper()
            
            if period_upper not in MONTHS and period_upper != "TOTALES":
                continue
            
            month_num = MONTH_MAP.get(period_upper, 13) if period_upper != "TOTALES" else 13
            
            for channel_name, start_col in channels.items():
                aforo_col = start_col
                ingreso_col = start_col + 1
                
                aforo = ws.cell(row, aforo_col).value
                ingreso = ws.cell(row, ingreso_col).value
                
                if aforo is None and ingreso is None:
                    continue
                
                aforo = float(aforo) if aforo is not None else 0.0
                ingreso = float(ingreso) if ingreso is not None else 0.0
                
                family = "Otros"
                if "PREPAGO" in channel_name.upper():
                    family = "TeleVía"
                elif "DECENAL" in channel_name.upper():
                    family = "TeleVía"
                elif "I+D" in channel_name.upper():
                    family = "PASE"
                elif "CAPUFE" in channel_name.upper():
                    family = "Exentos"
                elif "PINFRA" in channel_name.upper():
                    family = "Premium"
                
                all_records.append({
                    'sheet': sheet_name,
                    'year': sheet_year,
                    'concession': concession,
                    'period': period_upper,
                    'month_num': month_num,
                    'channel': channel_name,
                    'family': family,
                    'aforo': aforo,
                    'ingreso': ingreso,
                })
    
    df = pd.DataFrame(all_records)
    
    if not df.empty:
        df['year'] = df['year'].astype(int)
        df['month_num'] = df['month_num'].astype(int)
    
    return df

# ===== CARGAR DATOS =====
excel_path = Path(EXCEL_FILE)
if not excel_path.exists():
    st.error(f"❌ No encuentro {EXCEL_FILE}")
    st.stop()

try:
    with st.spinner("⏳ Cargando Excel..."):
        raw_df = load_excel_data(str(excel_path))
except Exception as e:
    st.error(f"❌ Error: {str(e)}")
    st.stop()

if raw_df.empty:
    st.error("❌ Sin datos en el Excel")
    st.stop()

all_years = sorted(raw_df['year'].dropna().astype(int).unique().tolist())
all_concessions = sorted(raw_df['concession'].dropna().unique().tolist())

# ===== HEADER =====
st.markdown(f"""
<div style="background: linear-gradient(135deg, {COLORS['televia']} 0%, {COLORS['info']} 100%);
            padding: 18px 22px; border-radius: 16px; color: white; margin-bottom: 14px;">
    <div style="font-size: 2rem; font-weight: 800;">📊 Dashboard TeleVía 2026</div>
    <div style="opacity: 0.9; margin-top: 4px;">Prepago, Post-pago, Decenal vs PASE (I+D)</div>
</div>
""", unsafe_allow_html=True)

st.caption(f"Archivo: {EXCEL_FILE} | Años: {min(all_years)}–{max(all_years)} | Concesiones: {len(all_concessions)}")

# ===== SIDEBAR =====
with st.sidebar:
    st.markdown("### 🎛️ Filtros")
    year_choice = st.selectbox("Año", all_years, index=len(all_years)-1)
    concession_choice = st.selectbox("Concesión", ["TODAS"] + all_concessions, index=0)

# ===== FILTRAR =====
filtered_df = raw_df[
    (raw_df['year'] == year_choice) & 
    (raw_df['period'] != 'TOTALES')
].copy()

if concession_choice != "TODAS":
    filtered_df = filtered_df[filtered_df['concession'] == concession_choice].copy()

if filtered_df.empty:
    st.warning("⚠️ Sin datos")
    st.stop()

# ===== AGREGACIONES =====
monthly_family = (
    filtered_df.groupby(['month_num', 'family'], as_index=False)
    .agg(aforo=('aforo', 'sum'), ingreso=('ingreso', 'sum'))
    .sort_values(['month_num', 'family'])
)

annual_conc = (
    filtered_df.groupby('concession', as_index=False)
    .agg(aforo=('aforo', 'sum'), ingreso=('ingreso', 'sum'))
    .sort_values('aforo', ascending=False)
)
annual_conc['rpc'] = annual_conc['ingreso'] / annual_conc['aforo'].replace(0, np.nan)

total_aforo = filtered_df['aforo'].sum()
total_ingreso = filtered_df['ingreso'].sum()
total_rpc = safe_divide(total_ingreso, total_aforo)

# ===== TABS =====
tab_home, tab_diag, tab_conc = st.tabs(["🏠 Home", "🔍 Diagnóstico", "📊 Concesiones"])

# ===== TAB 1 =====
with tab_home:
    st.subheader("Executive Summary")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Aforo Total", fmt_num(total_aforo))
    col2.metric("Ingresos Total", fmt_mxn(total_ingreso))
    col3.metric("Ingreso/Cruce", f"{total_rpc:.2f}" if pd.notna(total_rpc) else "-")
    col4.metric("Concesiones", len(annual_conc))
    
    left, right = st.columns([1.2, 0.8])
    
    with left:
        st.markdown("#### Aforo mensual por familia")
        monthly_pivot = monthly_family.pivot_table(
            index='month_num', 
            columns='family', 
            values='aforo', 
            fill_value=0
        )
        
        fig = go.Figure()
        for family in monthly_pivot.columns:
            color = COLORS['televia'] if family == 'TeleVía' else COLORS['pase']
            fig.add_trace(go.Scatter(
                x=monthly_pivot.index,
                y=monthly_pivot[family],
                mode='lines+markers',
                name=family,
                line=dict(width=3, color=color),
            ))
        
        fig.update_layout(
            xaxis=dict(title="Mes", tickvals=list(range(1, 13)), ticktext=MONTHS),
            yaxis=dict(title="Aforo", gridcolor=COLORS['grid']),
            hovermode='x unified',
            height=380,
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with right:
        st.markdown("#### Mix actual")
        family_total = filtered_df.groupby('family')['aforo'].sum()
        
        fig = px.pie(
            names=family_total.index,
            values=family_total.values,
            hole=0.5,
        )
        fig.update_layout(height=380)
        st.plotly_chart(fig, use_container_width=True)

# ===== TAB 2 =====
with tab_diag:
    st.subheader("Diagnóstico por Concesión")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Aforo máximo", fmt_num(annual_conc['aforo'].max()))
    col2.metric("RPC promedio", f"{annual_conc['rpc'].mean():.2f}")
    col3.metric("RPC máximo", f"{annual_conc['rpc'].max():.2f}")
    
    st.markdown("#### Top 10 Concesiones")
    top10 = annual_conc.head(10).copy()
    
    col1, col2 = st.columns([1.1, 0.9])
    
    with col1:
        fig = px.bar(
            top10,
            x='concession',
            y='aforo',
            color='rpc',
            color_continuous_scale='RdYlGn',
            text_auto='.0f',
            title="Aforo vs Ingreso/Cruce",
        )
        fig.update_layout(height=400, xaxis_title="", yaxis_title="Aforo")
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.dataframe(
            top10[['concession', 'aforo', 'ingreso', 'rpc']].style.format({
                'aforo': '{:,.0f}',
                'ingreso': '${:,.0f}',
                'rpc': '{:.2f}',
            }),
            use_container_width=True,
            height=400,
            hide_index=True,
        )

# ===== TAB 3 =====
with tab_conc:
    st.subheader("Análisis Detallado")
    
    if concession_choice == "TODAS":
        st.info("Selecciona concesión en sidebar")
    else:
        conc_data = filtered_df[filtered_df['concession'] == concession_choice].copy()
        
        if conc_data.empty:
            st.warning("Sin datos")
        else:
            conc_aforo = conc_data['aforo'].sum()
            conc_ingreso = conc_data['ingreso'].sum()
            conc_rpc = safe_divide(conc_ingreso, conc_aforo)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Aforo", fmt_num(conc_aforo))
            col2.metric("Ingresos", fmt_mxn(conc_ingreso))
            col3.metric("Ingreso/Cruce", f"{conc_rpc:.2f}" if pd.notna(conc_rpc) else "-")
            
            channel_summary = conc_data.groupby('channel').agg(
                aforo=('aforo', 'sum'),
                ingreso=('ingreso', 'sum'),
            ).sort_values('aforo', ascending=False)
            channel_summary['rpc'] = safe_divide(channel_summary['ingreso'], channel_summary['aforo'])
            
            st.markdown("#### Desglose por canal")
            st.dataframe(
                channel_summary.style.format({
                    'aforo': '{:,.0f}',
                    'ingreso': '${:,.0f}',
                    'rpc': '{:.2f}',
                }),
                use_container_width=True,
                height=250,
            )

# ===== FOOTER =====
st.markdown(f"""
<div style='margin-top: 2rem; padding-top: 1rem; border-top: 1px solid {COLORS['grid']}; 
color: {COLORS['muted']}; font-size: 0.85rem;'>
    {datetime.now().strftime('%d/%m/%Y %H:%M')} | {EXCEL_FILE}
</div>
""", unsafe_allow_html=True)
