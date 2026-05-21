# config.py
# Configuración centralizada del dashboard de aforos TeleVía

EXCEL_FILE = "dashboard_televia.xlsx"

APP_TITLE = "TeleVía | Decision Dashboard"
APP_SUBTITLE = "Executive Summary · Diagnóstico · Simulación · Plan de Acción · Seguimiento"

PAGE_TITLES = {
    "home": "1. Executive Summary",
    "diagnostico": "2. Diagnóstico",
    "simulacion": "3. Simulación",
    "plan": "4. Plan de Acción",
    "seguimiento": "5. Seguimiento",
}

MONTHS = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

MONTH_MAP = {m: i + 1 for i, m in enumerate(MONTHS)}

# Paleta corporativa limpia
COLORS = {
    "navy": "#1B2A4A",
    "blue": "#2563EB",
    "televia": "#2E7D32",
    "televia_light": "#E8F5E9",
    "pase": "#8A8A8A",
    "pase_light": "#F3F4F6",
    "info": "#1976D2",
    "success": "#16A34A",
    "warning": "#D97706",
    "danger": "#DC2626",
    "muted": "#64748B",
    "bg": "#F8FAFC",
    "card": "#FFFFFF",
    "border": "#E2E8F0",
    "grid": "#E5E7EB",
}

# Umbrales de lectura ejecutiva
THRESHOLDS = {
    "good": 8.0,
    "watch": 6.5,
}

# Patrones para filtrar bloques de soporte / acumulados
# Por defecto el dashboard trabaja con concesiones operativas.
EXCLUDE_TITLE_PATTERNS = [
    r"ACUMULADO",
    r"INTEROPERABILIDAD",
    r"MIS EN MIS",
    r"MIS EN SUS",
    r"SUS EN MIS",
    r"TOTALES MIS",
    r"TOTAL DE FACTURACION",
    r"TOTAL DE MIS EN MIS Y MIS EN SUS",
    r"TRAFICO TOTAL",
    r"TOTAL\b",
]

# Grupos de canales para análisis resumido
CHANNEL_GROUPS = {
    "Prepago": ["PREPAGO"],
    "Post-pago": ["POST-PAGO", "POSPAGO"],
    "Decenal": ["DECENAL"],
    "TeleVía": ["TELEVIA"],
    "I+D": ["I+D", "OTRAS I+D"],
    "CAPUFE": ["CAPUFE", "OTRAS CAPUFE", "CAPUFE-BANOBRAS", "CAPUFE-BANOBRAS (VET)", "CAPUFE-CD VALLES", "CAPUFE-CHAMAPA"],
    "PINFRA": ["PINFRA", "OTRAS PINFRA", "PINFRA (OSIPASS)"],
    "SITEL": ["SITEL", "OTRAS SITEL"],
    "EASYTRIP": ["EASYTRIP"],
    "Exentos/AUDI": ['EXENTOS', '"AUDI"'],
    "Totales": ["TOTALES", "TOTAL", "TOTALES MIS EN SUS", "TOTALES MIS EN MIS Y SUS EN MIS"],
    "Otros": [],
}

CHANNEL_ORDER = [
    "Prepago", "Post-pago", "Decenal", "TeleVía", "I+D",
    "CAPUFE", "PINFRA", "SITEL", "EASYTRIP", "Exentos/AUDI", "Totales", "Otros"
]

# Supuestos de simulación (solo defaults; el usuario puede moverlos)
SIMULATION_DEFAULTS = {
    "volume_growth_pct": 5.0,
    "share_capture_pp": 2.0,
    "rpc_uplift_pct": 3.0,
    "target_year": 2026,
}

# Librería básica de acciones para el plan de acción
ACTION_LIBRARY = {
    "volume_recovery": {
        "label": "Recuperación de volumen",
        "hint": "Concesión con volumen relevante y caída / estancamiento.",
    },
    "monetization": {
        "label": "Monetización / pricing",
        "hint": "Ingresos por cruce por debajo del benchmark del portafolio.",
    },
    "stabilization": {
        "label": "Estabilización operativa",
        "hint": "Alta volatilidad o dispersión en el desempeño mensual.",
    },
    "channel_capture": {
        "label": "Captura de canal",
        "hint": "Aumentar participación de un canal con potencial.",
    },
    "premium_mix": {
        "label": "Mejorar mix premium",
        "hint": "Subir el peso de canales de mayor ingreso por cruce.",
    },
}

MAX_TOP_OPPORTUNITIES = 10
