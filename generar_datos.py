#!/usr/bin/env python3
"""
CNEL EP — Generador de datos para dashboard de mantenimientos planificados

Uso simple (recomendado):
    python generar_datos.py
    → Busca automáticamente el único .xlsx en la misma carpeta

Uso con nombre específico:
    python generar_datos.py nombre_archivo.xlsx

Requiere: pip install pandas openpyxl

ID ESTABLE: cada trabajo tiene un ID único basado en su contenido
(empresa + descripción + subestación + fecha), NO en el número de fila.
Aunque agreguen nuevos trabajos en cualquier posición del Excel,
los trabajos existentes conservan su ID y el seguimiento no se pierde.
"""

import sys
import json
import hashlib
import pandas as pd
from datetime import datetime
from pathlib import Path

# ── helpers ───────────────────────────────────────────────────────────────────

def fmt_date(v):
    try:
        if pd.isna(v): return None
        return pd.Timestamp(v).strftime('%Y-%m-%d')
    except:
        return None

def fmt_time(v):
    try:
        if pd.isna(v): return None
        import datetime as dt_mod
        if isinstance(v, dt_mod.time):
            return f"{v.hour:02d}:{v.minute:02d}"
        if isinstance(v, dt_mod.datetime):
            return f"{v.hour:02d}:{v.minute:02d}"
        s = str(v).strip()
        if ':' in s and len(s) >= 4:
            parts = s.split(':')
            return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
        return pd.Timestamp(v).strftime('%H:%M')
    except:
        return None

def ss(v):
    try:
        if pd.isna(v): return ''
    except:
        pass
    return str(v).strip() if v is not None else ''

def sf(v):
    try:
        if pd.isna(v): return 0.0
        return float(v)
    except:
        return 0.0

def si(v):
    try:
        if pd.isna(v): return 0
        return int(float(v))
    except:
        return 0

def generar_id(empresa, descripcion, subestacion, alimentador, fecha_inicio, hora_inicio=''):
    """
    ID de 8 caracteres basado en el contenido del trabajo.
    Incluye hora de inicio para diferenciar trabajos del mismo día con misma descripción.
    Estable: no cambia si agregan filas nuevas en cualquier posición del Excel.
    """
    clave = f"{empresa}|{descripcion[:80]}|{subestacion}|{alimentador}|{fecha_inicio}|{hora_inicio}"
    return hashlib.md5(clave.encode('utf-8')).hexdigest()[:8].upper()

def is_transelectric(r):
    for field in ['Unnamed: 22',
                  'PROYECTO O DESCRIPCIÓN DE LOS TRABAJOS',
                  'Justificación del horario de ejecución']:
        val = str(r.get(field, '') or '').lower()
        if 'transelectric' in val or 'transmisor' in val:
            return True
    return False

def normalize_tipo(tipo):
    t = ss(tipo).lower()
    if 'ingreso' in t or 'carga' in t or 'extensi' in t:
        return 'Ingreso Nuevas Cargas'
    if 'prevent' in t: return 'Trabajos Preventivos'
    if 'correct' in t: return 'Trabajos Correctivos'
    return ss(tipo)

def encontrar_excel(arg):
    """
    Encuentra el archivo Excel a procesar:
    1. Si se pasó como argumento, usa ese
    2. Si no, busca el único .xlsx en la carpeta del script
    3. Si hay varios, muestra cuáles son y pide elegir
    """
    carpeta = Path(__file__).parent

    if arg:
        p = Path(arg)
        # Si es solo nombre de archivo, buscarlo en la carpeta del script
        if not p.is_absolute():
            p = carpeta / p
        if not p.exists():
            print(f"ERROR: No se encontró '{p}'")
            sys.exit(1)
        return p

    # Buscar automáticamente — ignorar archivos temporales de Excel (~$)
    candidatos = [f for f in carpeta.glob('*.xlsx')
                  if not f.name.startswith('~$')]

    if len(candidatos) == 1:
        print(f"Excel encontrado automáticamente: {candidatos[0].name}")
        return candidatos[0]
    elif len(candidatos) == 0:
        print("ERROR: No se encontró ningún archivo .xlsx en la carpeta.")
        print(f"Carpeta buscada: {carpeta}")
        print("Copie el archivo Excel aquí y vuelva a correr el script.")
        sys.exit(1)
    else:
        print(f"Se encontraron {len(candidatos)} archivos Excel en la carpeta:")
        for i, f in enumerate(candidatos, 1):
            print(f"  {i}. {f.name}")
        print("\nEspecifique cuál usar:")
        print(f"  python generar_datos.py \"nombre_del_archivo.xlsx\"")
        sys.exit(1)

# ── main ──────────────────────────────────────────────────────────────────────

def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    excel_path = encontrar_excel(arg)

    print(f"\nLeyendo: {excel_path.name}")
    xl = pd.ExcelFile(excel_path)

    # Buscar la hoja correcta sin importar el nombre
    df = None
    for sheet in xl.sheet_names:
        for hdr in [4, 0, 1, 2, 3]:
            try:
                temp = pd.read_excel(xl, sheet_name=sheet, header=hdr)
                temp.columns = [str(c).strip().replace('\n', ' ') for c in temp.columns]
                if 'Empresa' in temp.columns:
                    df = temp
                    print(f"  → Hoja: '{sheet}' (fila encabezado: {hdr})")
                    break
            except Exception:
                continue
        if df is not None:
            break

    if df is None:
        print('ERROR: No se encontro ninguna hoja con la columna Empresa.')
        print(f'Hojas disponibles: {xl.sheet_names}')
        sys.exit(1)

    # Limpiar filas sin empresa
    df = df[df['Empresa'].notna() & (df['Empresa'].astype(str).str.strip() != 'nan')].reset_index(drop=True)

    # Parsear fechas
    for col in ['Fecha de Inicio', 'Fecha de Fin']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    print(f'  → {len(df)} filas encontradas')

    # Cargar seguimiento existente para preservarlo
    output_path = Path(__file__).parent / 'data.json'
    seguimiento_previo = {}
    if output_path.exists():
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                data_prev = json.load(f)
            for t in data_prev.get('trabajos', []):
                if t.get('estado') != 'pendiente' or t.get('completado_por') or t.get('obs_ejecucion'):
                    seguimiento_previo[t['id']] = {
                        'estado':              t.get('estado', 'pendiente'),
                        'completado_por':      t.get('completado_por', ''),
                        'completado_fecha':    t.get('completado_fecha', ''),
                        'obs_ejecucion':       t.get('obs_ejecucion', ''),
                        'trabajos_adicionales':t.get('trabajos_adicionales', ''),
                        'validado_director':   t.get('validado_director', False),
                    }
            if seguimiento_previo:
                print(f"  → Seguimiento previo encontrado: {len(seguimiento_previo)} trabajo(s) con datos")
        except Exception as e:
            print(f"  → Aviso: no se pudo leer seguimiento previo ({e})")

    trabajos = []
    ids_vistos = {}  # para detectar duplicados

    for i, row in df.iterrows():
        r = row.to_dict()

        empresa     = ss(r.get('Empresa'))
        descripcion = ss(r.get('PROYECTO O DESCRIPCIÓN DE LOS TRABAJOS'))
        subestacion = ss(r.get('S/E'))
        alimentador = ss(r.get('ALIMENTADOR'))
        fecha_ini   = fmt_date(r.get('Fecha de Inicio'))

        # Saltar filas vacías
        if not empresa and not descripcion:
            continue

        # Generar ID estable
        hora_ini = fmt_time(r.get('Hora inicio'))
        tid = generar_id(empresa, descripcion, subestacion, alimentador, fecha_ini or '', hora_ini or '')

        # Manejar duplicados (muy raro, pero posible)
        if tid in ids_vistos:
            tid = tid + str(ids_vistos[tid])
        ids_vistos[tid] = ids_vistos.get(tid, 0) + 1

        # Recuperar seguimiento previo si existe
        seg = seguimiento_previo.get(tid, {})

        t = {
            "id":                  tid,
            "empresa":             empresa,
            "descripcion":         descripcion,
            "linea_subtransmision":ss(r.get('LINEA DE SUBTRANSMISIÓN')),
            "subestacion":         subestacion,
            "alimentador":         alimentador,
            "provincia":           ss(r.get('PROVINCIA')),
            "canton":              ss(r.get('CANTON')),
            "codigo_gis":          ss(r.get('CODIGO GIS')),
            "tipo_trabajo":        normalize_tipo(r.get('Desconexiones Programadas')),
            "mw_desconectados":    sf(r.get('MW (Desconectados)')),
            "fecha_inicio":        fecha_ini,
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
            # Seguimiento — recuperado del JSON previo o vacío si es nuevo
            "estado":                 seg.get('estado', 'pendiente'),
            "completado_por":         seg.get('completado_por', ''),
            "fecha_hora_inicio_real":  seg.get('fecha_hora_inicio_real', ''),
            "fecha_hora_fin_real":     seg.get('fecha_hora_fin_real', ''),
            "obs_ejecucion":          seg.get('obs_ejecucion', ''),
            "trabajos_adicionales":   seg.get('trabajos_adicionales', ''),
            "validado_director":      seg.get('validado_director', False),
        }
        trabajos.append(t)

    # ── Resumen ──
    nuevos = [t for t in trabajos if t['id'] not in seguimiento_previo]
    trans  = [t for t in trabajos if t['es_transelectric']]
    fechas = sorted([t['fecha_inicio'] for t in trabajos if t['fecha_inicio']])
    empresas = sorted(set(t['empresa'] for t in trabajos))

    print(f"\n  → Total trabajos:    {len(trabajos)}")
    print(f"  → Nuevos (sin seg.): {len(nuevos)}")
    print(f"  → Transelectric:     {len(trans)}")
    if fechas:
        print(f"  → Fechas:            {fechas[0]} → {fechas[-1]}")
    print(f"  → Unidades ({len(empresas)}):    {', '.join(e.replace('CNEL ','') for e in empresas)}")

    # ── Guardar ──
    output = {
        "trabajos": trabajos,
        "generado": datetime.now().isoformat(),
        "fuente":   excel_path.name,
        "total":    len(trabajos),
    }
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✓ data.json actualizado → {output_path}")
    if seguimiento_previo:
        recuperados = len([t for t in trabajos if t['id'] in seguimiento_previo])
        print(f"✓ Seguimiento preservado: {recuperados} trabajo(s) con datos existentes")
    print(f"\nPróximo paso: GitHub Desktop → Commit → Push")
    print(f"Netlify se actualizará automáticamente en ~1 minuto.")

if __name__ == '__main__':
    main()
