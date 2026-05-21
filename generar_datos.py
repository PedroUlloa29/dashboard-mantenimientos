#!/usr/bin/env python3
"""
CNEL EP — Generador de datos para dashboard de mantenimientos planificados
Uso: python generar_datos.py [archivo.xlsx]

Genera data.json en el mismo directorio que este script.
Requiere: pip install pandas openpyxl
"""

import sys
import json
import pandas as pd
from datetime import datetime
from pathlib import Path

EXCEL_DEFAULT = "Mantenimientos_planificados_CNEL_EP_3031_mayo_y_67junio_SEGUIMIENTO.xlsx"

def fmt_date(v):
    try:
        if pd.isna(v): return None
        return pd.Timestamp(v).strftime('%Y-%m-%d')
    except:
        return None

def fmt_time(v):
    try:
        if pd.isna(v): return None
        t = pd.Timestamp(v)
        return t.strftime('%H:%M')
    except:
        return None

def ss(v):
    """Safe string"""
    try:
        if pd.isna(v): return ''
    except:
        pass
    return str(v).strip() if v is not None else ''

def sf(v):
    """Safe float"""
    try:
        if pd.isna(v): return 0.0
        return float(v)
    except:
        return 0.0

def si(v):
    """Safe int"""
    try:
        if pd.isna(v): return 0
        return int(float(v))
    except:
        return 0

def is_transelectric(r):
    """Detecta si el trabajo está coordinado con Transelectric"""
    fields_to_check = [
        'Unnamed: 22',
        'PROYECTO O DESCRIPCIÓN DE LOS TRABAJOS',
        'Justificación del horario de ejecución',
    ]
    for field in fields_to_check:
        val = str(r.get(field, '') or '').lower()
        if 'transelectric' in val or 'transmisor' in val:
            return True
    return False

def normalize_tipo(tipo):
    """Normaliza el tipo de trabajo a valores canónicos"""
    t = ss(tipo).lower()
    if 'ingreso' in t or 'carga' in t:
        return 'Ingreso Nuevas Cargas'
    return ss(tipo)

def main():
    # Determinar archivo fuente
    excel_path = sys.argv[1] if len(sys.argv) > 1 else EXCEL_DEFAULT
    excel_path = Path(excel_path)

    if not excel_path.exists():
        print(f"ERROR: No se encontró el archivo '{excel_path}'")
        print(f"Uso: python generar_datos.py <archivo.xlsx>")
        sys.exit(1)

    print(f"Leyendo: {excel_path}")
    xl = pd.ExcelFile(excel_path)
    df = pd.read_excel(xl, sheet_name='Hoja1', parse_dates=['Fecha de Inicio', 'Fecha de Fin'])
    df.columns = [c.strip().replace('\n', ' ') for c in df.columns]

    print(f"  → {len(df)} filas encontradas")

    trabajos = []
    for i, row in df.iterrows():
        r = row.to_dict()
        t = {
            "id":                  i + 1,
            "empresa":             ss(r.get('Empresa')),
            "descripcion":         ss(r.get('PROYECTO O DESCRIPCIÓN DE LOS TRABAJOS')),
            "linea_subtransmision":ss(r.get('LINEA DE SUBTRANSMISIÓN')),
            "subestacion":         ss(r.get('S/E')),
            "alimentador":         ss(r.get('ALIMENTADOR')),
            "provincia":           ss(r.get('PROVINCIA')),
            "canton":              ss(r.get('CANTON')),
            "codigo_gis":          ss(r.get('CODIGO GIS')),
            "tipo_trabajo":        normalize_tipo(r.get('Desconexiones Programadas')),
            "mw_desconectados":    sf(r.get('MW (Desconectados)')),
            "fecha_inicio":        fmt_date(r.get('Fecha de Inicio')),
            "hora_inicio":         fmt_time(r.get('Hora inicio')),
            "fecha_fin":           fmt_date(r.get('Fecha de Fin')),
            "hora_fin":            fmt_time(r.get('Hora fin')),
            "valoracion":          sf(r.get('$ Valoración de los trabajos (materiales + Mano de obra)')),
            "presupuesto":         sf(r.get('Presupuesto requerido $')),
            "justificacion":       ss(r.get('Justificación del horario de ejecución')),
            "etapa":               ss(r.get('ETAPA FUNCIONAL')),
            "clientes_afectados":  si(r.get('Clientes Afectados')),
            "clientes_totales":    si(r.get('Clientes Totales')),
            "energia_kwh":         sf(r.get('Energía kW-h (por desconexión)}')),
            "usd_energia":         sf(r.get('USD energia (por desconexión)')),
            "obs_adicionales":     ss(r.get('Unnamed: 22')),
            "es_transelectric":    is_transelectric(r),
            # Campos de seguimiento (valores iniciales vacíos)
            "estado":              "pendiente",
            "completado_por":      "",
            "completado_fecha":    "",
            "obs_ejecucion":       "",
            "trabajos_adicionales":"",
            "validado_director":   False,
        }
        trabajos.append(t)

    # Resumen
    transelectric = [t for t in trabajos if t['es_transelectric']]
    fechas = sorted([t['fecha_inicio'] for t in trabajos if t['fecha_inicio']])

    print(f"  → Transelectric: {len(transelectric)} trabajos")
    if fechas:
        print(f"  → Rango de fechas: {fechas[0]} → {fechas[-1]}")

    empresas = sorted(set(t['empresa'] for t in trabajos))
    print(f"  → Unidades de Negocio ({len(empresas)}): {', '.join(empresas)}")

    # Guardar JSON
    output = {
        "trabajos":  trabajos,
        "generado":  datetime.now().isoformat(),
        "fuente":    str(excel_path.name),
    }

    output_path = Path(__file__).parent / 'data.json'
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✓ data.json generado: {len(trabajos)} trabajos → {output_path}")
    print("\nEstructura del directorio para desplegar en Netlify/GitHub:")
    print("  /")
    print("  ├── index.html")
    print("  └── data.json")
    print("\nNOTA: Los cambios de seguimiento se guardan en el navegador (localStorage).")
    print("      Para persistir los datos de seguimiento, use la función de exportación.")

if __name__ == '__main__':
    main()
