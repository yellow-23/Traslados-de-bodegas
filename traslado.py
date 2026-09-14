"""Armado del traslado a partir de los VIN pedidos, contra la hoja del Excel."""

import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_CONFIG = os.path.join(BASE_DIR, "traslado_config.json")
CARPETA_SALIDA_DEFECTO = os.path.join(BASE_DIR, "traslados_salida")

HOJA_MAESTRA = "Inventario Vehiculos Global"
COL_VIN = "VIN_Chasis"
COL_BODEGA = "BODEGA"
COL_EXCLUIR = "excluir_guia"

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


def parse_vins_pegados(texto):
    """Solo la columna de VIN, pegada tal cual (una por linea; si pegan mas
    columnas por error, se toma nomas la primera). No se adivina nada del
    correo: origen y destino de este grupo se preguntan aparte.

    Devuelve (vins en orden, repetidos).
    """
    vistos, vins, repetidos = set(), [], []
    for linea in texto.splitlines():
        celda = linea.split("\t")[0].strip()
        vin = normalizar_vin(pd.Series([celda])).iloc[0]
        if not vin or vin == "VIN":
            continue  # linea vacia o el encabezado de columna
        if vin in vistos:
            repetidos.append(vin)
            continue
        vistos.add(vin)
        vins.append(vin)
    return vins, repetidos


# Como decir en criollo que un VIN no paso cada filtro, en vez de mostrar el
# nombre crudo de la columna.
MENSAJES_FILTRO = {
    "VIGENTE": lambda actual: "no esta vigente",
    "EMBARQUE_ANULADO": lambda actual: "embarque anulado",
    "TIPO": lambda actual: f"tipo '{actual or '(vacio)'}' (no es Vehiculo Nuevo)",
    "empresa": lambda actual: f"empresa '{actual or '(vacio)'}' (no es HIPERMARC S.A.)",
}


def diagnosticar_ausente(df, vin):
    """Un VIN que no quedo en la maestra filtrada: dice si de plano no esta
    en el Excel, o cual filtro (TIPO/empresa/VIGENTE/EMBARQUE_ANULADO) lo saca."""
    filas = df[df[COL_VIN] == vin]
    if filas.empty:
        return "no esta en la base"

    razones = set()
    for _, fila in filas.iterrows():
        for columna, valor in FILTROS:
            if columna not in df.columns:
                continue
            actual = como_texto(pd.Series([fila[columna]])).iloc[0]
            if actual != valor:
                mensaje = MENSAJES_FILTRO.get(columna)
                razones.add(mensaje(actual) if mensaje else f"{columna}={actual or '(vacio)'}, se pide {valor}")
                break
        else:
            razones.add("cumple los filtros pero no quedo en la maestra (VIN duplicado?)")

    return "; ".join(sorted(razones))


def armar_filas(vins, origen, destino, conflictos, indice, df=None):
    """Una fila por VIN pedido, en el orden en que se pegaron.

    No se descarta ninguno: lo que no cuadra se anota en 'observacion' y el
    reporte lo pinta en rojo. Si se pasa 'df' (la maestra sin filtrar), un VIN
    ausente dice ademas por que: no existe, o que filtro lo descarta.
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
                "vigente": fila.get("VIGENTE"),
                "nota_de_venta": fila.get("NOTA_DE_VENTA"),
                "factura": fila.get("FACTURA"),
                "reservado": fila.get("esta_reservado"),
            }
        else:
            if df is not None:
                observaciones.append(diagnosticar_ausente(df, vin))
            else:
                observaciones.append("no esta en la base o no pasa el filtro")
            datos = {
                "nro_stock": "", "empresa": "", "marca": "", "modelo": "",
                "color": "", "etapa": "", "embarque": "", "bodega_actual": "",
                "vigente": "", "nota_de_venta": "", "factura": "", "reservado": "",
            }

        filas.append({
            "grupo": nombre_hoja(origen, destino),
            "vin": vin,
            **datos,
            "bodega_destino": destino,
            "concatena": f"'{vin}',",
            "observacion": "; ".join(observaciones),
            COL_EXCLUIR: "",
        })

    return filas


# Los mismos tonos del estilo "Incorrecto" de Excel.
RELLENO_ROJO = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
LETRA_ROJA = Font(color="9C0006")

ANCHOS = {
    "grupo": 16, "vin": 20, "nro_stock": 11, "empresa": 16, "marca": 12, "modelo": 24,
    "color": 14, "etapa": 7, "embarque": 14, "bodega_actual": 30,
    "vigente": 9, "nota_de_venta": 14, "factura": 12, "reservado": 11,
    "bodega_destino": 30, "concatena": 22, "observacion": 55, COL_EXCLUIR: 16,
}

# Los mismos tonos del estilo "Neutro" de Excel, para llamar la atencion sobre
# la columna que se llena a mano.
RELLENO_AMARILLO = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

# Bandas para distinguir a simple vista los grupos del correo (visible,
# restringido, etc.) dentro de la misma hoja. Se van rotando en el orden en
# que aparece cada grupo.
PALETA_GRUPOS = ["DDEBF7", "FCE4D6", "E2EFDA", "E4DFEC"]

HOJA_REPORTE = "Traslado"


def nombre_hoja(origen, destino):
    """'30100_31256': como se identifica cada grupo (columna 'grupo' del
    reporte y nombre de archivo del txt borrador)."""
    return f"{origen}_{destino}" if destino else origen


def guardar_txt_borrador(carpeta_salida, grupos_con_filas, fecha=None):
    """Un solo .txt con todos los VIN pedidos (todos los grupos, en el orden
    del correo), antes de marcar los que van con guia."""
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    fecha = fecha or datetime.now().strftime("%Y%m%d")

    ruta_txt = carpeta_salida / f"vins_traslado_{fecha}.txt"
    vins = [f["vin"] for _, filas in grupos_con_filas for f in filas]
    ruta_txt.write_text(",\n".join(f"'{v}'" for v in vins), encoding="utf-8")
    return ruta_txt


def guardar_reporte(carpeta_salida, grupos_con_filas, fecha=None):
    """Un solo Excel, una sola hoja, con los VIN de todos los grupos del
    correo (visible, restringido, etc.) uno tras otro en ese orden. Cada
    grupo se distingue por una banda de color, y hay una columna para
    marcar a mano los VIN que van con guia de despacho y no por el sistema."""
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    fecha = fecha or datetime.now().strftime("%Y%m%d")
    ruta_reporte = carpeta_salida / f"reporte_traslado_{fecha}.xlsx"

    filas = [f for _, filas_grupo in grupos_con_filas for f in filas_grupo]
    if not filas:
        raise ValueError("no hay filas que guardar")

    df = pd.DataFrame(filas)
    columnas = list(df.columns)
    col_excluir_idx = columnas.index(COL_EXCLUIR) + 1

    colores_por_grupo = {}
    for f in filas:
        if f["grupo"] not in colores_por_grupo:
            color = PALETA_GRUPOS[len(colores_por_grupo) % len(PALETA_GRUPOS)]
            colores_por_grupo[f["grupo"]] = PatternFill(
                start_color=color, end_color=color, fill_type="solid"
            )

    with pd.ExcelWriter(ruta_reporte, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=HOJA_REPORTE, index=False)

        hoja = writer.sheets[HOJA_REPORTE]
        hoja.freeze_panes = "A2"
        for i, columna in enumerate(columnas, start=1):
            hoja.column_dimensions[get_column_letter(i)].width = ANCHOS.get(columna, 14)
        hoja.cell(row=1, column=col_excluir_idx).fill = RELLENO_AMARILLO

        for i, fila in enumerate(filas, start=2):
            relleno_grupo = colores_por_grupo[fila["grupo"]]
            for celda in hoja[i]:
                celda.fill = relleno_grupo
            if fila["observacion"]:
                for celda in hoja[i]:
                    celda.fill = RELLENO_ROJO
                    celda.font = LETRA_ROJA
            hoja.cell(row=i, column=col_excluir_idx).fill = RELLENO_AMARILLO

    return ruta_reporte


def generar_txt_final(ruta_reporte, carpeta_salida=None):
    """Segundo paso: lee el reporte ya revisado (con las X en excluir_guia) y
    escribe UN solo .txt final con los VIN de todos los grupos (visible,
    restringido, etc.), en el mismo orden en que vinieron en el correo, sin
    los VIN marcados para guia de despacho.

    Devuelve (ruta_txt, detalle), donde detalle es una lista de
    (nombre_grupo, incluidos, excluidos) por cada grupo del reporte.
    """
    ruta_reporte = Path(ruta_reporte)
    carpeta_salida = Path(carpeta_salida) if carpeta_salida else ruta_reporte.parent
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    fecha = datetime.now().strftime("%Y%m%d")
    df = pd.read_excel(ruta_reporte, sheet_name=HOJA_REPORTE, dtype=str)
    if "vin" not in df.columns or COL_EXCLUIR not in df.columns or "grupo" not in df.columns:
        raise ValueError(f"la hoja '{HOJA_REPORTE}' no es un reporte de traslado valido")

    marcados = df[COL_EXCLUIR].fillna("").str.strip() != ""
    incluidos = df.loc[~marcados]

    detalle = []
    for grupo, sub in df.groupby("grupo", sort=False):
        excluidos_grupo = (sub[COL_EXCLUIR].fillna("").str.strip() != "").sum()
        detalle.append((grupo, len(sub) - int(excluidos_grupo), int(excluidos_grupo)))

    ruta_txt = carpeta_salida / f"vins_traslado_{fecha}_FINAL.txt"
    ruta_txt.write_text(",\n".join(f"'{v}'" for v in incluidos["vin"]), encoding="utf-8")

    return ruta_txt, detalle


def procesar_grupo(grupo, df, conflictos, indice):
    """Una tabla del correo (un deposito de destino, con su origen). Devuelve
    (filas, resumen); el llamador junta las filas de todos los grupos en un
    solo reporte y un .txt por grupo."""
    origen, destino, vins = grupo["origen"], grupo["destino"], grupo["vins"]

    print()
    print(f"--- {origen} -> {destino or '(sin destino)'} ---")

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

    if grupo.get("repetidos"):
        print(f"  (se ignoraron {len(grupo['repetidos'])} VIN repetidos en lo pegado)")
    print(f"  VIN a evaluar: {len(vins)}")

    filas = armar_filas(vins, origen, destino_resuelto, conflictos, indice, df=df)
    con_observacion = [f for f in filas if f["observacion"]]

    for f in con_observacion:
        print(f"  [REVISAR] {f['vin']}: {f['observacion']}")
    print(f"  total={len(filas)} con observacion={len(con_observacion)}")

    resumen = {
        "origen": origen, "destino": destino,
        "total": len(filas), "observaciones": len(con_observacion),
        "txt": None, "error": False,
    }
    return filas, resumen
