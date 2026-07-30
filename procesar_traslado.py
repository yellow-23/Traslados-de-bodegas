import json
import sys
from pathlib import Path

import openpyxl

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


def guardar_config(ruta_excel, carpeta_salida, grupos):
    datos = {
        "ruta_excel": str(ruta_excel),
        "carpeta_salida": str(carpeta_salida),
        "grupos": grupos,
    }
    try:
        Path(t.RUTA_CONFIG).write_text(
            json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        print(f"(no se pudo guardar la configuracion: {e})")


def elegir_ruta(titulo, carpeta=False):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        if carpeta:
            ruta = filedialog.askdirectory(title=titulo)
        else:
            ruta = filedialog.askopenfilename(
                title=titulo,
                filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")],
            )
        root.destroy()
        return ruta or None
    except Exception as e:
        print(f"(no se pudo abrir el selector: {e})")
        return None


def pedir_grupos(ruta_excel, grupos_previos):
    """Un grupo por cada tabla del correo: hoja pegada + par de bodegas."""
    libro = openpyxl.load_workbook(ruta_excel, read_only=True)
    try:
        hojas = list(libro.sheetnames)
    finally:
        libro.close()

    print()
    print(f"Hojas donde pegar los VIN: {', '.join(h for h in hojas if h != t.HOJA_MAESTRA)}")

    previos = {g["hoja"]: g for g in grupos_previos}
    default = ",".join(g["hoja"] for g in grupos_previos) or None

    while True:
        texto = pedir_obligatorio(
            "Hojas a procesar (una por tabla del correo, separadas por coma)", default
        )
        pedidas = [h.strip() for h in texto.split(",") if h.strip()]
        faltan = [h for h in pedidas if h not in hojas]
        if faltan:
            print(f"  No existen en el archivo: {', '.join(faltan)}")
            continue
        if pedidas:
            break

    grupos = []
    for hoja in pedidas:
        anterior = previos.get(hoja, {})
        print(f"  {hoja}:")
        origen = pedir_obligatorio("    bodega origen (A)", anterior.get("origen"))
        destino = pedir("    bodega destino (B)", anterior.get("destino", ""))
        grupos.append({"hoja": hoja, "origen": origen, "destino": destino or ""})

    return grupos


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
        print("\nSelecciona el archivo Excel en la ventana que se va a abrir...")
        ruta_excel = elegir_ruta("Selecciona el Excel de inventario")
        if not ruta_excel:
            ruta_excel = pedir_obligatorio("Escribe la ruta del Excel a mano")
        if not Path(ruta_excel).exists():
            raise FileNotFoundError(f"No existe el archivo: {ruta_excel}")

    if not carpeta_salida:
        print("Selecciona la carpeta donde se guardan los resultados...")
        carpeta_salida = elegir_ruta("Carpeta de salida", carpeta=True)
        if not carpeta_salida:
            carpeta_salida = pedir("Escribe la carpeta de salida", t.CARPETA_SALIDA_DEFECTO)

    print(f"\nSe usa la hoja '{t.HOJA_MAESTRA}' del Excel tal cual esta.")

    grupos = pedir_grupos(ruta_excel, guardada.get("grupos", []))
    guardar_config(ruta_excel, carpeta_salida, grupos)

    return {
        "ruta_excel": ruta_excel,
        "carpeta_salida": Path(carpeta_salida),
        "grupos": grupos,
    }


def resumen_final(resumenes):
    print()
    print("=" * 62)
    print(" RESUMEN")
    print("=" * 62)

    total = observaciones = 0
    for r in resumenes:
        estado = "FALLO" if r["error"] else ""
        print(
            f"  {r['hoja']:<10} {r['origen']} -> {r['destino'] or '-':<8} "
            f"total={r['total']:<4} con observacion={r['observaciones']} {estado}"
        )
        total += r["total"]
        observaciones += r["observaciones"]
    print(f"  {'TOTAL':<10} {'':<14} total={total} con observacion={observaciones}")

    archivos = [r["txt"] for r in resumenes if r["txt"]]
    if archivos:
        print(f"\n  Archivos generados en:\n    {archivos[0].parent}")

    if observaciones:
        print(f"\n  {observaciones} VIN quedaron marcados en rojo en el reporte.")
        print("  Van igual en el txt y en el concatenar; revisalos antes de cargar.")

    # Una observacion es lo que se viene a detectar, no una falla: solo se sale
    # con error si una hoja no se pudo procesar.
    return 1 if any(r["error"] for r in resumenes) else 0


def main():
    try:
        parametros = pedir_parametros()
    except (OSError, ValueError) as e:
        print(f"\nERROR: {e}")
        return 1

    print()
    print("=" * 62)
    print(" 1/2  INVENTARIO")
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
    print(f" 2/2  TRASLADOS ({len(parametros['grupos'])})")
    print("=" * 62)

    resumenes = []
    for grupo in parametros["grupos"]:
        try:
            resumenes.append(t.procesar_grupo(
                grupo, parametros["ruta_excel"], parametros["carpeta_salida"],
                df, conflictos, indice,
            ))
        except Exception as e:
            # Que falle una hoja no puede dejar sin procesar a las demas.
            print(f"  ERROR procesando {grupo['hoja']}: {e}")
            resumenes.append({
                "hoja": grupo["hoja"], "origen": grupo["origen"],
                "destino": grupo["destino"], "total": 0, "observaciones": 0,
                "txt": None, "reporte": None, "error": True,
            })

    return resumen_final(resumenes)


if __name__ == "__main__":
    sys.exit(main())
