import json
import sys
from datetime import datetime
from pathlib import Path

import traslado as t


def pedir(mensaje, default=None):
    texto = f"{mensaje} [{default}]: " if default else f"{mensaje}: "
    respuesta = input(texto).strip()
    return respuesta if respuesta else default


def pedir_obligatorio(mensaje, default=None):
    valor = pedir(mensaje, default)
    while not valor:
        valor = pedir(f"{mensaje} - obligatorio")
    return valor


def confirmar(mensaje, default_si=True):
    respuesta = input(f"{mensaje} (s/n) [{'s' if default_si else 'n'}]: ").strip().lower()
    if not respuesta:
        return default_si
    return respuesta in ("s", "si", "sí")


def cargar_config():
    ruta = Path(t.RUTA_CONFIG)
    if ruta.exists():
        try:
            return json.loads(ruta.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            return {}
    return {}


def guardar_config(ruta_excel, carpeta_salida, grupos_bodegas):
    datos = {
        "ruta_excel": str(ruta_excel),
        "carpeta_salida": str(carpeta_salida),
        "grupos": grupos_bodegas,
    }
    try:
        Path(t.RUTA_CONFIG).write_text(
            json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        print(f"(no se pudo guardar la configuracion: {e})")


def elegir_ruta(titulo, carpeta=False, tipos=None):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        if carpeta:
            ruta = filedialog.askdirectory(title=titulo)
        else:
            ruta = filedialog.askopenfilename(title=titulo, filetypes=tipos or [("Todos", "*.*")])
        root.destroy()
        return ruta or None
    except Exception as e:
        print(f"(no se pudo abrir el selector: {e})")
        return None


NUM_GRUPOS_DEFECTO = 2  # casi siempre vienen visible + restringido, pero a veces llega una sola tabla
ETIQUETAS_GRUPO = {0: " (normalmente: visible)", 1: " (normalmente: restringido)"}


def pedir_grupos_bodegas(guardada_grupos):
    """Cuantas tablas vienen esta vez y la bodega origen/destino de cada una.
    No se adivina nada del correo: esto se pregunta siempre, igual que antes."""
    n_defecto = len(guardada_grupos) if guardada_grupos else NUM_GRUPOS_DEFECTO

    print()
    n_texto = pedir("¿Cuantas tablas de VIN vienen en este correo?", str(n_defecto))
    try:
        n = max(1, int(n_texto))
    except (TypeError, ValueError):
        n = n_defecto

    grupos_bodegas = []
    for i in range(n):
        anterior = guardada_grupos[i] if i < len(guardada_grupos) else {}
        print(f"\nTabla de VIN {i + 1} de {n}{ETIQUETAS_GRUPO.get(i, '')}:")
        origen = pedir_obligatorio("  bodega origen", anterior.get("origen"))
        destino = pedir("  bodega destino", anterior.get("destino", ""))
        grupos_bodegas.append({"origen": origen, "destino": destino or ""})

    return grupos_bodegas


def pedir_parametros():
    print("=" * 62)
    print(" TRASLADO DE BODEGA")
    print("=" * 62)

    guardada = cargar_config()
    ruta_excel = carpeta_salida = None

    previa = guardada.get("ruta_excel")
    if previa and Path(previa).exists():
        print("\nConfiguracion de la ultima vez:")
        print(f"  Excel:             {previa}")
        print(f"  Carpeta de salida: {guardada.get('carpeta_salida', '')}")
        if confirmar("¿Usar la misma?"):
            ruta_excel = previa
            carpeta_salida = guardada.get("carpeta_salida")

    if not ruta_excel:
        print("\nSelecciona el archivo Excel de inventario en la ventana que se va a abrir...")
        ruta_excel = elegir_ruta(
            "Selecciona el Excel de inventario",
            tipos=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")],
        )
        if not ruta_excel:
            ruta_excel = pedir_obligatorio("Escribe la ruta del Excel a mano")
        if not Path(ruta_excel).exists():
            raise FileNotFoundError(f"No existe el archivo: {ruta_excel}")

    if not carpeta_salida:
        print("Selecciona la carpeta donde se guardan los resultados...")
        carpeta_salida = elegir_ruta("Carpeta de salida", carpeta=True)
        if not carpeta_salida:
            carpeta_salida = pedir("Escribe la carpeta de salida", t.CARPETA_SALIDA_DEFECTO)

    print(f"\nSe usa la hoja '{t.HOJA_MAESTRA}'")

    grupos_bodegas = pedir_grupos_bodegas(guardada.get("grupos", []))
    guardar_config(ruta_excel, carpeta_salida, grupos_bodegas)

    return {
        "ruta_excel": ruta_excel,
        "carpeta_salida": Path(carpeta_salida),
        "grupos_bodegas": grupos_bodegas,
    }


def pedir_vins_pegados(numero, origen, destino):
    print()
    print(f"--- Grupo {numero}: {origen} -> {destino or '(sin destino)'} ---")
    print("Pega SOLO la columna de VIN (chasis) de esta tabla, una por linea.")
    print("Cuando termines, escribe FIN en una linea aparte y presiona Enter.")
    print()

    lineas = []
    while True:
        try:
            linea = input()
        except EOFError:
            break
        if linea.strip().upper() == "FIN":
            break
        lineas.append(linea)

    return "\n".join(lineas)


def resumen_final(resumenes, ruta_txt=None, ruta_reporte=None):
    print()
    print("=" * 62)
    print(" RESUMEN")
    print("=" * 62)

    total = observaciones = 0
    for r in resumenes:
        estado = "FALLO" if r["error"] else ""
        print(
            f"  {r['origen']} -> {r['destino'] or '-':<8} "
            f"total={r['total']:<4} con observacion={r['observaciones']} {estado}"
        )
        total += r["total"]
        observaciones += r["observaciones"]
    print(f"  {'TOTAL':<19} total={total} con observacion={observaciones}")

    if ruta_txt or ruta_reporte:
        print(f"\n  Archivos generados en:\n    {(ruta_txt or ruta_reporte).parent}")
    if ruta_txt:
        print(f"    {ruta_txt.name}  (borrador, todos los VIN)")
    if ruta_reporte:
        print(f"    {ruta_reporte.name}  (una sola hoja, todos los grupos)")

    if observaciones:
        print(f"\n  {observaciones} VIN quedaron marcados en rojo en el reporte.")
        print("  Van igual en el txt y en el concatenar; revisalos antes de cargar.")

    print()
    print("  Si algun VIN va con guia de despacho (no por el sistema), marca")
    print("  una X en la columna 'excluir_guia' del reporte, guarda el Excel y")
    print("  corre generar_txt_final.bat para obtener el .txt definitivo.")

    # Una observacion es lo que se viene a detectar, no una falla: solo se sale
    # con error si un grupo no se pudo procesar.
    return 1 if any(r["error"] for r in resumenes) else 0


def main():
    try:
        parametros = pedir_parametros()
    except (OSError, ValueError) as e:
        print(f"\nERROR: {e}")
        return 1

    print()
    print("=" * 62)
    print(" 1/3  INVENTARIO")
    print("=" * 62)
    try:
        df, df_filtrado = t.obtener_inventario(parametros["ruta_excel"])
    except Exception as e:
        print(f"\nERROR al obtener el inventario: {e}")
        print("No se proceso ningun traslado: sin el inventario no se puede decidir.")
        return 1

    conflictos, indice = t.preparar_maestra(df_filtrado)
    if conflictos:
        print(f"  VIN con datos ambiguos en la maestra: {len(conflictos)}")

    print()
    print("=" * 62)
    print(f" 2/3  VIN ({len(parametros['grupos_bodegas'])} grupo(s))")
    print("=" * 62)

    carpeta_salida = parametros["carpeta_salida"]
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    fecha = datetime.now().strftime("%Y%m%d")

    grupos = []
    for i, gb in enumerate(parametros["grupos_bodegas"], start=1):
        texto = pedir_vins_pegados(i, gb["origen"], gb["destino"])
        vins, repetidos = t.parse_vins_pegados(texto)
        print(f"  VIN pegados: {len(vins)}")

        grupos.append({
            "origen": gb["origen"], "destino": gb["destino"],
            "vins": vins, "repetidos": repetidos,
        })

    if not any(g["vins"] for g in grupos):
        print("\n  No se pego ningun VIN en ningun grupo.")
        return 1

    print()
    print("=" * 62)
    print(f" 3/3  TRASLADOS ({len(grupos)})")
    print("=" * 62)

    resumenes = []
    grupos_con_filas = []
    for grupo in grupos:
        try:
            filas, resumen = t.procesar_grupo(grupo, df, conflictos, indice)
            grupos_con_filas.append((grupo, filas))
            resumenes.append(resumen)
        except Exception as e:
            # Que falle un grupo no puede dejar sin procesar a los demas.
            print(f"  ERROR procesando {grupo['origen']} -> {grupo['destino']}: {e}")
            resumenes.append({
                "origen": grupo["origen"], "destino": grupo["destino"],
                "total": 0, "observaciones": 0, "error": True,
            })

    ruta_txt = ruta_reporte = None
    if any(filas for _, filas in grupos_con_filas):
        ruta_txt = t.guardar_txt_borrador(carpeta_salida, grupos_con_filas, fecha=fecha)
        ruta_reporte = t.guardar_reporte(carpeta_salida, grupos_con_filas, fecha=fecha)

    return resumen_final(resumenes, ruta_txt, ruta_reporte)


if __name__ == "__main__":
    sys.exit(main())
