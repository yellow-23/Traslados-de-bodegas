import argparse
import re
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

URL_OMBU = "http://oyg.cidef.cl/oyg/cidef/"
PAUSA = 0.6             # segundos entre unidades
ESPERA = 15             # segundos maximos esperando que Ombu responda
FALLOS_MAX = 3          # fallas seguidas antes de detenerse
TIPO_COMPROBANTE = "TRASLADO MASIVO DE VEHICULOS"
TIPO_PRODUCTO = "5"     # vehiculo Nuevo


class OmbuError(Exception):
    pass


def codigo(texto):
    #'31250-FERRY TRUCK' -> '31250' 
    return re.split(r"[-\s]", str(texto).strip(), maxsplit=1)[0]



class VigilaAvisos:
    # clase que lee alerts de ombu y avisos.
    def __init__(self):
        self.mensajes = []
        self._fin = threading.Event()

    def iniciar(self):
        threading.Thread(target=self._correr, daemon=True).start()

    def detener(self):
        self._fin.set()

    def _correr(self):
        
        vistos = set()
        import win32gui
        while not self._fin.is_set():
            try:
                win32gui.EnumWindows(lambda h, _: self._revisar(h, vistos), None)
            except Exception:
                pass
            time.sleep(0.2)

    def _revisar(self, hwnd, vistos):
        import win32con, win32gui
        if hwnd in vistos or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd) != "#32770":
            return True
        textos, botones = [], []

        def hijo(h, _):
            clase = win32gui.GetClassName(h)
            if clase == "Static" and win32gui.GetWindowText(h).strip():
                textos.append(win32gui.GetWindowText(h).strip())
            elif clase == "Button":
                botones.append(h)
            return True

        win32gui.EnumChildWindows(hwnd, hijo, None)
        vistos.add(hwnd)
        texto = " ".join(textos) or win32gui.GetWindowText(hwnd)
        self.mensajes.append(texto)
        if botones:
            win32gui.PostMessage(botones[0], win32con.BM_CLICK, 0, 0)
        return True

# internet explorer : encuentra la pantalla y manejar sus campos

def esperar(condicion, seg=ESPERA, paso=0.25):
    fin = time.time() + seg
    while time.time() < fin:
        if condicion():
            return True
        time.sleep(paso)
    return False


def buscar_doc(doc, parte, prof=0):
    # busca, emtre los marcos de la pagina, el documento cuya direcc contiene parte
    try:
        if parte.lower() in str(doc.location.href).lower():
            return doc
        for i in range(doc.frames.length):
            r = buscar_doc(doc.frames.item(i).document, parte, prof + 1) if prof < 6 else None
            if r is not None:
                return r
    except Exception:
        pass
    return None


def buscar_pantalla():
    # busca pantalla de transferencias de stock
    import win32com.client
    for w in win32com.client.Dispatch("Shell.Application").Windows():
        try:
            if "oyg.cidef.cl" in str(w.LocationURL).lower():
                doc = buscar_doc(w.Document, "TransferStk.asp")
                if doc is not None:
                    return doc
        except Exception:
            continue
    return None


def abrir_ie():
    import win32com.client
    print("Abriendo Internet Explorer en Ombu...")
    ie = win32com.client.DispatchEx("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate(URL_OMBU)


def filas_grilla(doc):
    return doc.getElementById("tgriTransfStock").getElementsByTagName("TR").length


def texto_combo(doc, id_):
    sel = doc.getElementById(id_)
    return "" if sel.selectedIndex < 0 else str(sel.options.item(sel.selectedIndex).text)


def elegir(doc, id_, coincide, espera=12):
    # elige la primera opcion del combo que coincida con la funcion coincide(valor, texto)
    sel = doc.getElementById(id_)
    fin = time.time() + espera
    while True:
        for i in range(sel.options.length):
            op = sel.options.item(i)
            if coincide(str(op.value), str(op.text)):
                if sel.selectedIndex != i:
                    sel.selectedIndex = i
                    sel.fireEvent("onchange")
                    time.sleep(1.5)          # Ombu recarga datos tras cada cambio
                return True
        if time.time() > fin:
            return False
        time.sleep(0.5)


def preparar_pantalla(doc, cod_o, cod_d):
    pasos = [
        ("Tipo Comprobante", "cmbTipoComprobante", lambda v, t: TIPO_COMPROBANTE in t.upper()),
        ("Bodega Origen", "cmbDepositoOrigen", lambda v, t: codigo(t) == cod_o),
        ("Bodega Destino", "cmbDepositoDestino", lambda v, t: codigo(t) == cod_d),
        ("Tipo Producto", "cmbTipoProducto", lambda v, t: v == TIPO_PRODUCTO),
    ]
    for nombre, id_, coincide in pasos:
        if not elegir(doc, id_, coincide):
            raise OmbuError("No encontre la opcion para %s en la pantalla de Ombu" % nombre)
        print("  %-17s %s" % (nombre + ":", texto_combo(doc, id_)))


def poner_tipo_producto(doc):
    # Ombu vacia Tipo Producto despues de cada Agregar o Cancelar; hay que volver a elegirlo. 
    sel = doc.getElementById("cmbTipoProducto")
    if str(sel.value) != TIPO_PRODUCTO:
        sel.value = TIPO_PRODUCTO
        sel.fireEvent("onchange")
        time.sleep(1.5)
    return str(sel.value) == TIPO_PRODUCTO


# Agregar las unidades, una por una
def agregar_unidades(doc, unidades, vig):
    # unidades: lista de nro_stock. Devuelve (agregadas, [(nro, motivo)]). No aprieta Finalizar.    
    campo = doc.getElementById("txtNroStock")
    articulo = doc.getElementById("txtCodArticulo")
    agregar = doc.getElementById("cmdAgregar")
    cancelar = doc.getElementById("cmdCancelarBoton")
    agregadas, fallidas, seguidos = [], [], 0

    for nro in unidades:
        if seguidos >= FALLOS_MAX:
            print("\n%d fallas seguidas; me detengo para no forzar el sistema." % seguidos)
            break
        if str(articulo.value).strip() or str(campo.value).strip():
            cancelar.click()
        n_av, n_fil = len(vig.mensajes), filas_grilla(doc)
        motivo = None

        if not poner_tipo_producto(doc):
            motivo = "No pude elegir Tipo Producto"
        else:
            campo.value = nro
            campo.fireEvent("onchange")      # equivale a apretar Tab
            cargo = lambda: len(vig.mensajes) > n_av or str(articulo.value).strip() != ""
            if not esperar(cargo, seg=4):    # si no reacciona, se llama a la rutina de Ombu directamente
                campo.focus()
                doc.parentWindow.execScript("txtNroStock_change", "VBScript")
            esperar(cargo)
            if len(vig.mensajes) > n_av:
                motivo = " / ".join(vig.mensajes[n_av:])
            elif not str(articulo.value).strip():
                motivo = "Ombu no cargo la unidad (sin aviso)"

        if motivo is None:
            agregar.click()
            esperar(lambda: filas_grilla(doc) > n_fil or len(vig.mensajes) > n_av)
            if len(vig.mensajes) > n_av:
                motivo = " / ".join(vig.mensajes[n_av:])
            elif filas_grilla(doc) <= n_fil:
                motivo = "No se agrego a la grilla"

        if motivo is None:
            agregadas.append(nro)
            seguidos = 0
            print("  %-8s agregada (%d en la grilla)" % (nro, len(agregadas)))
        else:
            fallidas.append((nro, motivo))
            seguidos += 1
            print("  %-8s NO -> %s" % (nro, motivo))
        time.sleep(PAUSA)
    return agregadas, fallidas


def cargar_grupo(cod_o, cod_d, unidades, abrir):
    # Espera la pantalla de traslados, la prepara y agrega las unidades
    if abrir:
        abrir_ie()
    print("\nEntra a Stock > Transferencias (Nuevo). Espero hasta 30 minutos...")
    doc, fin = None, time.time() + 1800
    while doc is None and time.time() < fin:
        doc = buscar_pantalla()
        if doc is None:
            time.sleep(2)
    if doc is None:
        raise OmbuError("No llegaste a la pantalla de traslados; no agregue nada.")
    print("Pantalla detectada.")
    time.sleep(2)
    preparar_pantalla(doc, cod_o, cod_d)
    input("\nListo. Revisa la pantalla y presiona Enter para agregar %d unidades... " % len(unidades))

    vig = VigilaAvisos()
    vig.iniciar()
    try:
        agregadas, fallidas = agregar_unidades(doc, unidades, vig)
    finally:
        vig.detener()
    print("\nAgregadas: %d | no agregadas: %d. Revisa la grilla y aprieta Finalizar tu." %
          (len(agregadas), len(fallidas)))
    return agregadas, fallidas


def ruta_reporte(nombre):
    # Acepta la ruta completa, o solo el nombre (con o sin .xlsx) del reporte en traslados_salida.
    carpeta = Path(__file__).resolve().parent / "traslados_salida"
    for base in (Path(nombre), carpeta / nombre):
        for ruta in (base, base.with_name(base.name + ".xlsx")):
            if ruta.is_file():
                return ruta
    raise OmbuError("No encuentro el reporte '%s' (busque tambien en %s)" % (nombre, carpeta))


def probar_avisos():
    # Prueba solo la vigilancia de avisos: provoca un aviso en Ombu y mira que se lea y se cierre.
    print("Vigilando avisos de Ombu durante 2 minutos.")
    print("En Ombu (pantalla Transferencias) escribe un Nro. Stock que no exista, ej. 99999999, y presiona Tab.")
    vig = VigilaAvisos()
    vig.iniciar()
    vistos = 0
    try:
        fin = time.time() + 120
        while time.time() < fin:
            while vistos < len(vig.mensajes):
                print("  AVISO LEIDO Y CERRADO:", vig.mensajes[vistos])
                vistos += 1
            time.sleep(0.3)
    except KeyboardInterrupt:
        pass
    vig.detener()
    print("Fin. Avisos leidos: %d" % len(vig.mensajes))


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--reporte", help="reporte_traslado_FECHA.xlsx")
    ap.add_argument("--grupo", help="procesar solo este grupo (por defecto todos los del Excel)")
    ap.add_argument("--max", type=int, help="maximo de unidades por grupo (para probar)")
    ap.add_argument("--abrir-ie", action="store_true", help="abrir Internet Explorer en Ombu")
    ap.add_argument("--vigilaavisos", action="store_true",
                    help="solo prueba la lectura de avisos de Ombu (2 minutos); no toca ninguna pantalla")
    args = ap.parse_args()

    if args.vigilaavisos:
        probar_avisos()
        return 0
    if not args.reporte:
        ap.error("indica --reporte (o --vigilaavisos para probar los avisos)")

    import pandas as pd
    ruta = ruta_reporte(args.reporte)
    df = pd.concat(pd.read_excel(ruta, sheet_name=None, dtype=str).values()).fillna("")
    for c in ("grupo", "nro_stock"):
        if c not in df.columns:
            raise OmbuError("El reporte no trae la columna '%s'" % c)

    grupos, pendientes = {}, []
    for fila in df.to_dict("records"):
        if args.grupo and fila["grupo"] != args.grupo:
            continue
        if "_" not in fila["grupo"]:
            pendientes.append({**fila, "motivo": "la fila no tiene grupo (formato ORIGEN_DESTINO)"})
            continue
        o, d = (codigo(x) for x in fila["grupo"].split("_", 1))
        if str(fila.get("excluir_guia", "")).strip():
            pendientes.append({**fila, "motivo": "marcada excluir_guia (va con guia de despacho)"})
        elif codigo(fila.get("bodega_actual", "")) == d:
            pendientes.append({**fila, "motivo": "el Excel ya la muestra en el destino"})
        else:
            grupos.setdefault((o, d), []).append(fila)

    if not grupos:
        print("No hay unidades para procesar.")
        return 0
    print("Grupos del reporte:")
    for (o, d), filas in grupos.items():
        print("  %s -> %s : %d unidades" % (o, d, len(filas)))

    total = len(grupos)
    for n, ((o, d), filas) in enumerate(grupos.items(), 1):
        filas = filas[:args.max] if args.max else filas
        por_nro = {str(f["nro_stock"]).strip(): f for f in filas}
        print("\n=== Grupo %d de %d: %s -> %s (%d unidades) ===" % (n, total, o, d, len(filas)))
        try:
            agregadas, fallidas = cargar_grupo(o, d, list(por_nro), args.abrir_ie and n == 1)
        except OmbuError as e:
            print("\nDETENIDO: %s" % e)
            agregadas, fallidas = [], []
        hechas = set(agregadas) | {x for x, _ in fallidas}
        pendientes += [{**por_nro[x], "motivo": m} for x, m in fallidas]
        pendientes += [{**f, "motivo": "no se alcanzo a procesar"} for x, f in por_nro.items() if x not in hechas]
        if n < total:
            if input("\nFinaliza este traslado y deja abierta una pantalla Transferencias (Nuevo) nueva.\n"
                     "Enter para el siguiente grupo, N para terminar: ").strip().lower() == "n":
                for (_, _), resto in list(grupos.items())[n:]:
                    pendientes += [{**f, "motivo": "grupo no procesado"} for f in resto]
                break

    if pendientes:
        cols = ["vin", "nro_stock", "empresa", "bodega_actual", "bodega_destino", "vigente", "motivo"]
        dfp = pd.DataFrame(pendientes)
        for c in cols:
            if c not in dfp.columns:
                dfp[c] = ""
        salida = Path(__file__).resolve().parent / "traslados_salida"
        salida.mkdir(exist_ok=True)
        ruta = salida / ("pendientes_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S"))
        dfp[cols].to_excel(ruta, index=False)
        print("Pendientes guardados en:", ruta)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except OmbuError as e:
        print("ERROR:", e)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nInterrumpido por el usuario.")
        sys.exit(130)
