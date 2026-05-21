# config.py

# -----------------------------------------------------------------------------
# App metadata
# -----------------------------------------------------------------------------
APP_TITLE = "Dashboard TeleVía - Incremento de MS"
APP_SUBTITLE = "Plataforma de Decision Intelligence para crecimiento de market share"

# -----------------------------------------------------------------------------
# File paths
# -----------------------------------------------------------------------------
EXCEL_FILE = "dashboard_televia.xlsx"   # Cambia si tu archivo tiene otro nombre

# -----------------------------------------------------------------------------
# Page titles (usadas en st.tabs)
# -----------------------------------------------------------------------------
PAGE_TITLES = {
    "home": "🎯 Control Tower",
    "diagnostico": "🔍 Diagnóstico Causal",
    "simulacion": "🧪 Simulador",
    "plan": "💰 Portafolio Oportunidades",
    "seguimiento": "📈 Seguimiento",
}

# -----------------------------------------------------------------------------
# Color palette (estilo corporativo)
# -----------------------------------------------------------------------------
COLORS = {
    "navy": "#0A2B4E",
    "televia": "#1E88E5",
    "pase": "#43A047",
    "info": "#00ACC1",
    "warning": "#FFB300",
    "danger": "#E53935",
    "purple": "#8E24AA",
    "muted": "#78909C",
    "grid": "#E0E0E0",
}

# -----------------------------------------------------------------------------
# Month mapping (para parseo de periodos)
# -----------------------------------------------------------------------------
MONTHS = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

MONTH_MAP = {mes: i+1 for i, mes in enumerate(MONTHS)}

# -----------------------------------------------------------------------------
# Reglas de clasificación de canales en familias
# -----------------------------------------------------------------------------
CHANNEL_RULES = {
    "TeleVía": ["PREPAGO", "POST-PAGO", "POSTPAGO", "DECENAL"],
    "PASE": ["PASE", "AUSUR", "CHAMAPA"],
    "CAPUFE/Exentos": ["CAPUFE", "EXENTOS"],
    "PINFRA": ["PINFRA"],
    "SITEL": ["SITEL"],
    "EASYTRIP": ["EASYTRIP"],
    "Otros": ["I+D", "AUDI", "OTRAS"],   # canales que no encajan directamente
}
# Nota: "Totales" se asigna automáticamente a agregados

# Orden de familias para visualización
CHANNEL_ORDER = ["TeleVía", "PASE", "CAPUFE/Exentos", "PINFRA", "SITEL", "EASYTRIP", "Otros", "Totales"]

# -----------------------------------------------------------------------------
# Exclusión de títulos de control (bloques que no son concesiones operativas)
# -----------------------------------------------------------------------------
EXCLUDE_TITLE_PATTERNS = [
    "ACUMULADO",
    "INTEROPERABILIDAD",
    "Trafico Total",
    "Trafico Facturable",
    "AUSUR",
    "CHAMAPA",
    "CAPUFE-CHAMAPA",
    "CEPSICE",
    "EASYTRIP",
    "SITEL",
    "PINFRA",
    "META",
    "COTESA",
    "VET",
    "METLAPIL",
]

# -----------------------------------------------------------------------------
# Textos de ayuda (para el radio de alcance)
# -----------------------------------------------------------------------------
TEXT_HELP = {
    "scope_operational": "Solo concesiones operativas",
    "scope_all": "Incluir bloques de control",
}

# -----------------------------------------------------------------------------
# Acciones recomendadas (usadas en el plan de acción)
# -----------------------------------------------------------------------------
ACTION_LIBRARY = {
    "volume_recovery": {"label": "Recuperación de volumen"},
    "monetization": {"label": "Monetización"},
    "stabilization": {"label": "Estabilización"},
    "channel_capture": {"label": "Captura de canal"},
    "premium_mix": {"label": "Mix premium"},
}

# -----------------------------------------------------------------------------
# Valores por defecto para la simulación
# -----------------------------------------------------------------------------
SIMULATION_DEFAULTS = {
    "volume_growth_pct": 5.0,
    "share_capture_pp": 5.0,
    "rpc_uplift_pct": 5.0,
}

# -----------------------------------------------------------------------------
# Límites para la tabla de oportunidades
# -----------------------------------------------------------------------------
MAX_TOP_OPPORTUNITIES = 10
