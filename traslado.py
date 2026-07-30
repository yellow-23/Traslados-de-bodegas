"""Armado del traslado a partir de los VIN pedidos, contra la hoja del Excel."""

import os
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_CONFIG = os.path.join(BASE_DIR, "traslado_config.json")
CARPETA_SALIDA_DEFECTO = os.path.join(BASE_DIR, "traslados_salida")

HOJA_MAESTRA = "Inventario Vehiculos Global"
COL_VIN = "VIN_Chasis"
COL_BODEGA = "BODEGA"
COL_VIN_BUSQUEDA = "A"

# Se aplica solo en memoria: no se escribe ninguna hoja con el resultado.
FILTROS = [
    ("TIPO", "Vehiculo Nuevo"),
    ("empresa", "HIPERMARC S.A."),
    ("VIGENTE", "1"),
    ("EMBARQUE_ANULADO", "0"),
]


def normalizar_vin(serie):
    """NaN -> "" y no "NAN": un VIN vacio no debe hacer match con nada."""
    return serie.fillna("").astype(str).str.strip().str.upper()


# VIGENTE y EMBARQUE_ANULADO no siempre llegan como texto: desde el Excel
# vienen como float o bool. Comparar el astype(str) crudo contra "1" dejaba el
# filtro en cero filas sin avisar, y ahi ningun VIN se encuentra.
EQUIVALENTES = {"true": "1", "false": "0", "1.0": "1", "0.0": "0"}


def como_texto(serie):
    texto = serie.fillna("").astype(str).str.strip()
    return texto.str.lower().map(EQUIVALENTES).fillna(texto)


def filtrar(df):
    filtrado = df
    for columna, valor in FILTROS:
        if columna not in filtrado.columns:
            print(f"  (no viene la columna {columna}, se omite ese filtro)")
            continue
        filtrado = filtrado[como_texto(filtrado[columna]) == valor]
    return filtrado.copy()


def obtener_inventario(ruta_excel):
    """Lee la maestra del Excel tal cual esta. Devuelve (df completa, df filtrada).

    El Excel nunca se modifica: la hoja se toma como viene y se filtra en memoria.
    """
    print(f"Leyendo '{HOJA_MAESTRA}' del archivo...")
    df = pd.read_excel(ruta_excel, sheet_name=HOJA_MAESTRA, dtype={COL_VIN: str})
    print(f"  {len(df):,} filas")

    df[COL_VIN] = normalizar_vin(df[COL_VIN])
    df_filtrado = filtrar(df)
    detalle = " / ".join(f"{c}={v}" for c, v in FILTROS)
    print(f"  {len(df_filtrado):,} filas tras el filtro ({detalle})")

    return df, df_filtrado


def preparar_maestra(df_filtrado):
    """Devuelve (conflictos por VIN, indice por VIN).

    La vista repite filas identicas del mismo VIN, asi que solo se marca como
    conflicto cuando trae mas de un NRO_STOCK o mas de una bodega distintos.
    """
    indice = df_filtrado.drop_duplicates(subset=COL_VIN, keep="first").set_index(COL_VIN)

    conflictos = {}
    for columna, etiqueta in (("NRO_STOCK", "NRO_STOCK"), (COL_BODEGA, "bodega")):
        if columna not in df_filtrado.columns:
            continue
        valores = df_filtrado.groupby(COL_VIN)[columna].unique()
        for vin, distintos in valores[valores.map(len) > 1].items():
            texto = ", ".join(str(v).strip() for v in distintos)
            conflictos.setdefault(vin, []).append(f"tiene mas de un {etiqueta}: {texto}")

    return conflictos, indice


def resolver_bodegas(df, prefijo):
    if COL_BODEGA not in df.columns or not prefijo:
        return []
    bodegas = df[COL_BODEGA].dropna().astype(str).str.strip().unique()
    return sorted(b for b in bodegas if b.startswith(prefijo))


def leer_vins(ruta_excel, hoja):
    """Columna A de una hoja de pegado. Devuelve (VIN en orden, repetidos)."""
    df = pd.read_excel(ruta_excel, sheet_name=hoja, usecols=COL_VIN_BUSQUEDA, dtype=str)
    df.columns = [COL_VIN]
    df[COL_VIN] = normalizar_vin(df[COL_VIN])

    vistos, vins, repetidos = set(), [], []
    for vin in df[COL_VIN]:
        if not vin:
            continue
        if vin in vistos:
            repetidos.append(vin)
            continue
        vistos.add(vin)
        vins.append(vin)
    return vins, repetidos


def armar_filas(vins, origen, destino, conflictos, indice):
    """Una fila por VIN pedido, en el orden del correo.

    No se descarta ninguno: lo que no cuadra se anota en 'observacion' y el
    reporte lo pinta en rojo.
    """
    filas = []

    for vin in vins:
        observaciones = []

        if vin in indice.index:
            fila = indice.loc[vin]
            bodega_actual = str(fila[COL_BODEGA]).strip()

            observaciones.extend(conflictos.get(vin, []))
            if origen and not bodega_actual.startswith(origen):
                observaciones.append(f"esta en {bodega_actual}, no en {origen}")

            datos = {
                "nro_stock": fila.get("NRO_STOCK"),
                "empresa": fila.get("empresa"),
                "marca": fila.get("MARCA"),
                "modelo": fila.get("MODELO"),
                "color": fila.get("COLOR"),
                "etapa": fila.get("ETAPA"),
                "embarque": fila.get("EMBARQUE"),
                "bodega_actual": bodega_actual,
            }
        else:
            observaciones.append("no esta en la base o no pasa el filtro")
            datos = {
                "nro_stock": "", "empresa": "", "marca": "", "modelo": "",
                "color": "", "etapa": "", "embarque": "", "bodega_actual": "",
            }

        filas.append({
            "vin": vin,
            **datos,
            "bodega_destino": destino,
            "concatena": f"'{vin}',",
            "observacion": "; ".join(observaciones),
        })

    return filas


# Los mismos tonos del estilo "Incorrecto" de Excel.
RELLENO_ROJO = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
LETRA_ROJA = Font(color="9C0006")

ANCHOS = {
    "vin": 20, "nro_stock": 11, "empresa": 16, "marca": 12, "modelo": 24,
    "color": 14, "etapa": 7, "embarque": 14, "bodega_actual": 30,
    "bodega_destino": 30, "concatena": 22, "observacion": 55,
}


def guardar_salidas(carpeta_salida, origen, destino, filas):
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    fecha = datetime.now().strftime("%Y%m%d")
    sufijo = f"{fecha}_{origen}_{destino}" if destino else f"{fecha}_{origen}"

    ruta_txt = carpeta_salida / f"vins_traslado_{sufijo}.txt"
    ruta_txt.write_text(",\n".join(f"'{f['vin']}'" for f in filas), encoding="utf-8")

    ruta_reporte = carpeta_salida / f"reporte_traslado_{sufijo}.xlsx"
    df = pd.DataFrame(filas)
    with pd.ExcelWriter(ruta_reporte, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Traslado", index=False)

        hoja = writer.sheets["Traslado"]
        hoja.freeze_panes = "A2"
        for i, columna in enumerate(df.columns, start=1):
            hoja.column_dimensions[get_column_letter(i)].width = ANCHOS.get(columna, 14)

        for i, fila in enumerate(filas, start=2):
            if fila["observacion"]:
                for celda in hoja[i]:
                    celda.fill = RELLENO_ROJO
                    celda.font = LETRA_ROJA

    return ruta_txt, ruta_reporte


def procesar_grupo(grupo, ruta_excel, carpeta_salida, df, conflictos, indice):
    """Una hoja de pegado con su par de bodegas. Devuelve el resumen."""
    hoja, origen, destino = grupo["hoja"], grupo["origen"], grupo["destino"]

    print()
    print(f"--- {hoja}: {origen} -> {destino or '(sin destino)'} ---")

    # Se buscan en la base completa y no en la filtrada: una bodega de destino
    # vacia igual tiene que reconocerse.
    matches_origen = resolver_bodegas(df, origen)
    matches_destino = resolver_bodegas(df, destino)

    if matches_origen:
        print(f"  Bodega origen ({origen}): {', '.join(matches_origen)}")
    else:
        print(f"  ATENCION: ninguna bodega empieza con '{origen}'")

    if destino:
        if matches_destino:
            print(f"  Bodega destino ({destino}): {', '.join(matches_destino)}")
        else:
            print(f"  ATENCION: ninguna bodega empieza con '{destino}'")

    destino_resuelto = matches_destino[0] if len(matches_destino) == 1 else destino

    vins, repetidos = leer_vins(ruta_excel, hoja)
    if repetidos:
        print(f"  (se ignoraron {len(repetidos)} VIN repetidos en la hoja)")
    print(f"  VIN a evaluar: {len(vins)}")

    filas = armar_filas(vins, origen, destino_resuelto, conflictos, indice)
    con_observacion = [f for f in filas if f["observacion"]]

    for f in con_observacion:
        print(f"  [REVISAR] {f['vin']}: {f['observacion']}")
    print(f"  total={len(filas)} con observacion={len(con_observacion)}")

    resumen = {
        "hoja": hoja, "origen": origen, "destino": destino,
        "total": len(filas), "observaciones": len(con_observacion),
        "txt": None, "reporte": None, "error": False,
    }

    if not filas:
        print("  la hoja no tiene VIN, no se genera archivo")
        return resumen

    ruta_txt, ruta_reporte = guardar_salidas(carpeta_salida, origen, destino, filas)
    print(f"  txt:     {ruta_txt.name}")
    print(f"  reporte: {ruta_reporte.name}")

    resumen["txt"] = ruta_txt
    resumen["reporte"] = ruta_reporte
    return resumen
