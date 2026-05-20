import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

# Configuración de página
st.set_page_config(
    page_title="Estrategia TeleVía 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos personalizados
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
    }
    .kpi-card {
        background: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
        border-top: 4px solid #2E7D32;
    }
    .metric-value {
        font-size: 32px;
        font-weight: bold;
        color: #1a1a1a;
        margin: 10px 0;
    }
    .metric-label {
        font-size: 14px;
        color: #666;
        margin-bottom: 5px;
    }
    .metric-change {
        font-size: 13px;
        color: #2E7D32;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
col1, col2 = st.columns([1, 1])
with col1:
    st.title("📊 Estrategia TeleVía 2026")
    st.subheader("Crecimiento de Aforos + Captura de Mercado vs PASE")

with col2:
    # Selector de idioma
    idioma = st.radio("Idioma:", ["Español", "English"], horizontal=True)
    
# ===== NAVEGACIÓN =====
st.markdown("---")
tabs = st.tabs([
    "🏠 Home",
    "🔍 Diagnóstico",
    "🎯 Estrategia",
    "📈 Ejecución",
    "⚠️ Riesgos"
])

# ===== TAB 1: HOME =====
with tabs[0]:
    st.markdown("### Resumen Ejecutivo")
    
    # KPIs principales en 4 columnas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="Aforo 2025",
            value="187.5M",
            delta="cruces",
            delta_color="off"
        )
        st.caption("Meta 2026: 197.2M (+5.2%)")
    
    with col2:
        st.metric(
            label="Ingresos 2025",
            value="$18.2B",
            delta="MXN",
            delta_color="off"
        )
        st.caption("Meta 2026: $19.5B (+7.3%)")
    
    with col3:
        st.metric(
            label="Share TeleVía",
            value="27.3%",
            delta="2026: 35%",
            delta_color="normal"
        )
        st.caption("+7.7pp captura PASE")
    
    with col4:
        st.metric(
            label="Share PASE",
            value="72.7%",
            delta="2026: 65%",
            delta_color="inverse"
        )
        st.caption("Desplazamiento objetivo")
    
    # Gráfico central: Donut interactivo
    st.markdown("### Mercado: TeleVía vs PASE")
    
    col_left, col_right = st.columns([2, 1])
    
    with col_left:
        # Donut actual
        fig_donut_actual = go.Figure(data=[go.Pie(
            labels=['PASE (I+D)', 'TeleVía'],
            values=[72.7, 27.3],
            hole=0.5,
            marker=dict(colors=['#999', '#2E7D32']),
            textposition='inside',
            textinfo='label+percent'
        )])
        fig_donut_actual.update_layout(
            title="2025 (Hoy)",
            height=400,
            showlegend=True
        )
        st.plotly_chart(fig_donut_actual, use_container_width=True)
    
    with col_right:
        # Donut meta
        fig_donut_meta = go.Figure(data=[go.Pie(
            labels=['PASE (I+D)', 'TeleVía'],
            values=[65, 35],
            hole=0.5,
            marker=dict(colors=['#ccc', '#1B5E20']),
            textposition='inside',
            textinfo='label+percent'
        )])
        fig_donut_meta.update_layout(
            title="2026 (Meta)",
            height=400,
            showlegend=True
        )
        st.plotly_chart(fig_donut_meta, use_container_width=True)
    
    st.info("💡 **Oportunidad:** Pasar de 27.3% a 35% de share = +7.7pp captura del mercado PASE")
    
    # Elasticidad
    st.markdown("### ¿Por qué crecer TeleVía?")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("""
        **Elasticidad de Ingresos (MXN/cruce):**
        - PASE (I+D): **$100.55**
        - Prepago TeleVía: $70.42
        - Post-pago TeleVía: $69.57
        - **Decenal TeleVía: $143.75** ⭐
        
        **Insight:** Decenal genera MÁS que PASE, pero tiene baja penetración (2%).
        """)
    
    with col2:
        # Gráfico barras elasticidad
        elasticidad_data = {
            'Método': ['PASE\n(I+D)', 'Prepago\nTeleVía', 'Post-pago\nTeleVía', 'Decenal\nTeleVía'],
            'MXN/cruce': [100.55, 70.42, 69.57, 143.75],
            'color': ['#999', '#90EE90', '#90EE90', '#1B5E20']
        }
        
        fig_elasticidad = go.Figure(data=[go.Bar(
            x=elasticidad_data['Método'],
            y=elasticidad_data['MXN/cruce'],
            marker=dict(color=elasticidad_data['color']),
            text=[f"${v:.0f}" for v in elasticidad_data['MXN/cruce']],
            textposition='outside'
        )])
        fig_elasticidad.update_layout(
            title="Elasticidad por Método de Pago",
            yaxis_title="MXN por cruce",
            xaxis_title="",
            height=300,
            showlegend=False
        )
        st.plotly_chart(fig_elasticidad, use_container_width=True)
    
    # Presupuesto e inversión
    st.markdown("### Inversión vs Retorno")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col1:
        st.metric("Inversión Total", "$9.55M", "5.1% del valor generado")
    
    with col2:
        st.metric("Retorno Esperado", "$940M", "Año 1")
    
    with col3:
        st.metric("ROI", "98x", "Ratio retorno/inversión")
    
    st.success("✓ **Estrategia de bajo riesgo, alto retorno:**\n$9.55M de inversión generan $940M de ingresos adicionales")
    
    # Acciones rápidas
    st.markdown("---")
    st.markdown("### Próximos Pasos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        **Semana 1-2**
        - Validar supuestos
        - Ajustar presupuesto
        - Aprobación Finanzas
        """)
    
    with col2:
        st.markdown("""
        **Semana 3-4**
        - Aprobación Dir. Comercial
        - Kick-off Gerentes
        - Capacitación
        """)
    
    with col3:
        st.markdown("""
        **Junio 2026**
        - Lanzamiento Fase 1
        - "Regresa a TeleVía"
        - Piloto 3 plazas
        """)
    
    # Footer
    st.markdown("---")
    st.caption(f"Dashboard v1.0 | Actualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Autor: Manuel Cuéllar")

# ===== TAB 2: DIAGNÓSTICO =====
with tabs[1]:
    st.markdown("### 🔍 Análisis Actual: ¿Dónde Estamos?")
    st.info("Sección en desarrollo - Incluirá: Share por concesión, penetración por método, matriz de oportunidad, datos Hoja2")

# ===== TAB 3: ESTRATEGIA =====
with tabs[2]:
    st.markdown("### 🎯 Estrategia 2026: ¿Qué Hacemos?")
    st.info("Sección en desarrollo - Incluirá: 3 Pilares (Volumen, Captura, Elasticidad), Timeline, Presupuesto desglosado")

# ===== TAB 4: EJECUCIÓN =====
with tabs[3]:
    st.markdown("### 📈 Ejecución: ¿Cómo lo Medimos?")
    st.info("Sección en desarrollo - Incluirá: KPI Real-time, Tracking por concesión, vs PASE, Proyecciones Q1-Q4")

# ===== TAB 5: RIESGOS =====
with tabs[4]:
    st.markdown("### ⚠️ Riesgos y Mitigaciones")
    st.info("Sección en desarrollo - Incluirá: Matriz de riesgos, Probabilidad/Impacto, Planes B, Contingencias")
