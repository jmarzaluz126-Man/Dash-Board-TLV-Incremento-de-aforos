# Configuración centralizada del Dashboard TeleVía 2026

# ===== COLORES =====
COLORS = {
    'televia': '#2E7D32',      # Verde TeleVía
    'televia_light': '#90EE90', # Verde claro
    'pase': '#999999',          # Gris PASE
    'pase_light': '#cccccc',    # Gris claro
    'warning': '#FF9800',       # Naranja riesgos
    'danger': '#D32F2F',        # Rojo crítico
    'info': '#1976D2',          # Azul info
    'success': '#388E3C'        # Verde éxito
}

# ===== DATOS 2025 ACTUAL =====
DATA_2025 = {
    'aforo_total': 187_500_000,           # cruces
    'ingresos_total': 18_221_337_147,     # MXN
    'share_televia': 0.273,               # 27.3%
    'share_pase': 0.727,                  # 72.7%
    'televia_cruces': 51_100_250,         # (27.3%)
    'pase_cruces': 136_399_750,           # (72.7%)
}

# ===== METAS 2026 =====
DATA_2026_META = {
    'aforo_total': 197_200_000,           # +5.2%
    'ingresos_total': 19_547_000_000,     # +7.3%
    'share_televia': 0.35,                # 35% (meta)
    'share_pase': 0.65,                   # 65% (meta)
    'televia_cruces': 69_020_000,         # 35% de 197M
    'pase_cruces': 128_180_000,           # 65% de 197M
}

# ===== ELASTICIDAD (MXN/cruce) =====
ELASTICIDAD = {
    'pase': 100.55,
    'prepago_televia': 70.42,
    'postpago_televia': 69.57,
    'decenal_televia': 143.75,
}

# ===== LOS 3 PILARES =====
PILARES = {
    'pilar_1': {
        'nombre': 'Crecer Aforo Total',
        'objetivo': '+9.7M cruces (+5.2%)',
        'inversión': 4_500_000,
        'ingresos_adicionales': 240_000_000,
        'meta_cruces': 4_600_000,
        'acciones': [
            'B2B Corporativo: +180K cruces/mes',
            'Retail Frecuente: +120K cruces/mes',
            'Rutas nuevas: +80K cruces/mes'
        ]
    },
    'pilar_2': {
        'nombre': 'Capturar Volumen PASE',
        'objetivo': '+7.7pp share (27.3% → 35%)',
        'inversión': 2_600_000,
        'ingresos_adicionales': 350_000_000,
        'meta_cruces': 5_000_000,
        'acciones': [
            'Prepago: Descuento 15%, meta +2M cruces',
            'Post-pago: Integración caja, meta +1.5M',
            'Decenal: Premium corporate, meta +500K'
        ]
    },
    'pilar_3': {
        'nombre': 'Optimizar Ingresos/Cruce',
        'objetivo': '+8.2% elasticidad ($87.84 → $95)',
        'inversión': 950_000,
        'ingresos_adicionales': 350_000_000,
        'meta_mxn': 350_000_000,
        'acciones': [
            'Pricing dinámico: +$80M',
            'Bundling servicios: +$120M',
            'Decenal premium: +$150M'
        ]
    }
}

# ===== PRESUPUESTO CONSOLIDADO =====
PRESUPUESTO = {
    'pilar_1': 4_500_000,
    'pilar_2': 2_600_000,
    'pilar_3': 950_000,
    'soporte': 1_500_000,
    'total': 9_550_000,
}

RETORNO_ESPERADO = {
    'pilar_1': 240_000_000,
    'pilar_2': 350_000_000,
    'pilar_3': 350_000_000,
    'total': 940_000_000,
}

ROI = RETORNO_ESPERADO['total'] / PRESUPUESTO['total']  # 98.3x

# ===== ROADMAP TRIMESTRAL =====
ROADMAP = {
    'q1': {
        'trimestre': 'Q1 (Ene–Mar)',
        'hitos': [
            'Lanzar "Regresa a TeleVía" (Prepago, 3 plazas)',
            'Auditoría grupos C (Lepsa, Cobi, Gana)',
            'Implementar pricing dinámico piloto',
        ],
        'metas': {
            'aforo': '+1.2M cruces',
            'share_televia': '+2pp',
            'inversión': 2_000_000
        }
    },
    'q2': {
        'trimestre': 'Q2 (Abr–Jun)',
        'hitos': [
            'Expandir "Regresa" a todas las concesiones',
            'Lanzar Post-pago integrado',
            'Posicionar Decenal premium',
        ],
        'metas': {
            'aforo': '+2.5M cruces',
            'share_televia': '+4pp',
            'usuarios_decenal': '+150K',
            'inversión': 2_500_000
        }
    },
    'q3': {
        'trimestre': 'Q3 (Jul–Sep)',
        'hitos': [
            'Escalar B2B: 80 empresas',
            'Lanzar app lealtad TeleVía',
            'Campaña Black Friday',
        ],
        'metas': {
            'aforo': '+3.0M cruces',
            'share_televia': '+6pp',
            'empresas_b2b': 80,
            'inversión': 2_500_000
        }
    },
    'q4': {
        'trimestre': 'Q4 (Oct–Dic)',
        'hitos': [
            'Navidad: bundling (tag + viaje + estacionamiento)',
            'Consolidación retención usuarios',
            'Preparación 2027',
        ],
        'metas': {
            'aforo': '+2.0M cruces (acumulado +9.7M)',
            'share_televia': '+8pp (meta 35%)',
            'retención': '95% usuarios nuevos',
            'inversión': 2_000_000
        }
    }
}

# ===== CONCESIONES SEGMENTADAS =====
CONCESIONES = {
    'tier_1': {
        'nombre': 'MOTOROLAS (Consolidación)',
        'concesiones': ['CALH', 'IRAPUATO', 'META', 'CONMEX'],
        'accion': 'Retención + optimización precio',
        'inversión': 1_000_000,
    },
    'tier_2': {
        'nombre': 'OPORTUNIDAD MÁXIMA (Crecimiento)',
        'concesiones': ['AUNORTE', 'VIADUCTO', 'SUPERVIA', 'ACCSA', 'LEPSA'],
        'accion': 'Captura agresiva PASE, convenios B2B',
        'inversión': 4_200_000,
    },
    'tier_3': {
        'nombre': 'CRÍTICO (Intervención)',
        'concesiones': ['GANA', 'COBI'],
        'accion': 'Auditoría + recuperación urgente',
        'inversión': 800_000,
    }
}

# ===== RIESGOS =====
RIESGOS = [
    {
        'riesgo': 'Falla sistema TeleVía',
        'probabilidad': '15%',
        'impacto': 'Alto (pierden 2-3pp share)',
        'mitigacion': 'Redundancia, SLA > 99.5%'
    },
    {
        'riesgo': 'PASE baja tarifa 20%',
        'probabilidad': '35%',
        'impacto': 'Alto (presión margen)',
        'mitigacion': 'Valor (bundling), no precio puro'
    },
    {
        'riesgo': 'Baja adopción post-pago caja',
        'probabilidad': '25%',
        'impacto': 'Medio (1pp share menos)',
        'mitigacion': 'Simplificar, training gerentes'
    },
    {
        'riesgo': 'Usuarios no migran de PASE',
        'probabilidad': '20%',
        'impacto': 'Medio (crecimiento más lento)',
        'mitigacion': 'Mayor descuento, referral'
    }
]

# ===== TEXTOS MULTIIDIOMA =====
TEXTOS = {
    'español': {
        'titulo': 'Estrategia TeleVía 2026',
        'subtitulo': 'Crecimiento de Aforos + Captura de Mercado vs PASE',
        'home': 'Home',
        'diagnostico': 'Diagnóstico',
        'estrategia': 'Estrategia',
        'ejecucion': 'Ejecución',
        'riesgos': 'Riesgos',
    },
    'english': {
        'titulo': 'TeleVía Strategy 2026',
        'subtitulo': 'Growth in Volumes + Market Capture vs PASE',
        'home': 'Home',
        'diagnostico': 'Diagnosis',
        'estrategia': 'Strategy',
        'ejecucion': 'Execution',
        'riesgos': 'Risks',
    }
}
