APP_TITLE = "TeleVía | Decision Dashboard"
APP_SUBTITLE = "Aforo, ingreso, captura de mercado y simulación de crecimiento"

EXCEL_FILE = "dashboard_televia.xlsx"

PAGE_TITLES = {
    "home": "1. Executive Summary",
    "diagnostico": "2. Diagnóstico",
    "simulacion": "3. Simulación",
    "plan": "4. Plan de acción",
    "seguimiento": "5. Seguimiento",
}

MONTHS = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
]

MONTH_MAP = {m: i + 1 for i, m in enumerate(MONTHS)}

COLORS = {
    "televia": "#2E7D32",
    "televia_light": "#A5D6A7",
    "pase": "#757575",
    "warning": "#F59E0B",
    "danger": "#D32F2F",
    "info": "#1976D2",
    "success": "#16A34A",
    "navy": "#1B2A4A",
    "bg": "#F8FAFC",
    "grid": "#E2E8F0",
    "muted": "#64748B",
    "white": "#FFFFFF",
    "purple": "#7C3AED",
}

# Grupos de canales para análisis y visualización.
# La función de clasificación en app.py usa estas reglas en orden.
CHANNEL_RULES = {
    "TeleVía": [
        "TELEVIA",
        "PREPAGO",
        "POST-PAGO",
        "POSPAGO",
        "DECENAL",
        "I+D",
        "MIS EN MIS",
        "MIS EN SUS",
        "SUS EN MIS",
    ],
    "PASE": [
        "AUSUR (PASE)",
        "AUSUR",
        "CHAMAPA Y OTRAS (PASE)",
        "CHAMAPA Y OTRAS",
        "PASE",
    ],
    "CAPUFE/Exentos": [
        "CAPUFE-BANOBRAS",
        "CAPUFE-CD VALLES",
        "CAPUFE-CHAMAPA",
        "CAPUFE",
        "CEP-SICE",
        "EXENTOS",
        "AUDI",
        "OTRAS CAPUFE",
        "OTRAS I+D",
    ],
    "PINFRA": [
        "PINFRA (OSIPASS)",
        "PINFRA",
        "OTRAS PINFRA",
    ],
    "SITEL": [
        "SITEL",
        "OTRAS SITEL",
    ],
    "EASYTRIP": [
        "EASYTRIP",
    ],
    "Otros": [],
}

CHANNEL_ORDER = [
    "TeleVía",
    "PASE",
    "CAPUFE/Exentos",
    "PINFRA",
    "SITEL",
    "EASYTRIP",
    "Otros",
    "Totales",
]

# Títulos que no representan una concesión operativa sino bloques de control/totalización.
EXCLUDE_TITLE_PATTERNS = [
    "ACUMULADO",
    "INTEROPERABILIDAD",
    "TOTALES",
    "TOTAL DE FACTURACION",
    "Trafico Total",
    "Trafico Facturable",
]

SIMULATION_DEFAULTS = {
    "volume_growth_pct": 5.0,
    "share_capture_pp": 1.5,
    "rpc_uplift_pct": 3.0,
}

MAX_TOP_OPPORTUNITIES = 12

ACTION_LIBRARY = {
    "volume_recovery": {
        "label": "Recuperación de volumen",
        "description": "La base es relevante, pero el aforo cayó vs. el año previo.",
    },
    "monetization": {
        "label": "Monetización",
        "description": "El ingreso por cruce está por debajo del benchmark.",
    },
    "stabilization": {
        "label": "Estabilización",
        "description": "La variabilidad mensual es alta y debe controlarse.",
    },
    "channel_capture": {
        "label": "Captura de canal",
        "description": "Hay oportunidad para ganar participación en canales clave.",
    },
    "premium_mix": {
        "label": "Mix premium",
        "description": "El mix puede migrar hacia canales de mayor valor.",
    },
    "benchmark": {
        "label": "Replicar benchmark",
        "description": "Hay una concesión líder que puede servir como referencia.",
    },
}

TEXT_HELP = {
    "scope_operational": "Solo concesiones operativas",
    "scope_all": "Incluir bloques de control",
    "ytd_note": "Se usa corte YTD hasta el último mes disponible del año seleccionado.",
}
