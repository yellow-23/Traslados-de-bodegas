"""Paso 2: lee el reporte_traslado_FECHA.xlsx ya revisado (con las X puestas
en la columna 'excluir_guia' para los VIN que van con guia de despacho) y
genera UN solo .txt final, combinando todos los grupos, sin esos VIN."""

import sys
from pathlib import Path

import traslado as t


def elegir_ruta(titulo, tipos):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ruta = filedialog.askopenfilename(title=titulo, filetypes=tipos)
        root.destroy()
        return ruta or None
    except Exception as e:
        print(f"(no se pudo abrir el selector: {e})")
        return None


def main():
    print("=" * 62)
    print(" GENERAR TXT FINAL (sin los VIN con guia de despacho)")
    print("=" * 62)

    carpeta_salida_defecto = Path(t.CARPETA_SALIDA_DEFECTO)
    ruta_reporte = None
    if carpeta_salida_defecto.exists():
        reportes = sorted(carpeta_salida_defecto.glob("reporte_traslado_*.xlsx"))
        if reportes:
            ultimo = reportes[-1]
            respuesta = input(f"¿Usar el ultimo reporte generado? ({ultimo.name}) (s/n) [s]: ").strip().lower()
            if respuesta in ("", "s", "si", "sí"):
                ruta_reporte = ultimo

    if not ruta_reporte:
        print("Selecciona el reporte_traslado_FECHA.xlsx ya revisado...")
        elegido = elegir_ruta(
            "Selecciona el reporte de traslado",
            tipos=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if not elegido:
            elegido = input("Escribe la ruta del reporte: ").strip()
        ruta_reporte = Path(elegido)

    if not ruta_reporte.exists():
        print(f"\nERROR: no existe el archivo: {ruta_reporte}")
        return 1

    try:
        ruta_txt, detalle = t.generar_txt_final(ruta_reporte)
    except ValueError as e:
        print(f"\nERROR: {e}")
        return 1

    print()
    total_incluidos = total_excluidos = 0
    for grupo, incluidos, excluidos in detalle:
        print(f"  {grupo}: {incluidos} VIN incluidos, {excluidos} excluidos (guia de despacho)")
        total_incluidos += incluidos
        total_excluidos += excluidos

    print(f"\n  TOTAL: {total_incluidos} VIN en el txt final, {total_excluidos} excluidos")
    print(f"  -> {ruta_txt}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
