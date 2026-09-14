TRASLADO DE BODEGA
==================

QUE HACE
--------
A partir del cuerpo del correo (tal cual llega, con una tabla por deposito:
visible, restringido, etc.), detecta cada tabla y su origen/destino, cruza
los VIN contra la hoja "Inventario Vehiculos Global" del Excel, y genera:

  - reporte_traslado_FECHA.xlsx   un Excel con una hoja por deposito del
                                   correo, cada VIN con su detalle y una
                                   columna "excluir_guia" para marcar a mano
                                   los que van con guia de despacho.
  - vins_traslado_FECHA_ORIGEN_DESTINO.txt   (borrador, todos los VIN)

En el reporte, las filas en rojo son las que hay que revisar antes de cargar:
VIN que no aparecen en la base, o que estan en una bodega distinta a la de
origen esperada. Igual salen en el .txt, no se descartan.


COMO SE USA (proceso normal, en 2 pasos)
-----------------------------------------
1. Instalar Python desde https://www.python.org/downloads/
   IMPORTANTE: marcar la casilla "Add Python to PATH" al instalar.
   (Solo la primera vez, en un equipo nuevo.)

PASO 1 - cruzar y armar el reporte
2. Doble clic en  traslado_solo_excel.bat
3. Responde lo que pregunta:
   - que Excel de inventario usar (se abre un selector de archivos)
   - en que carpeta dejar los resultados
4. Cuando pida el correo, copia el cuerpo completo (las tablas de visible Y
   de restringido, con sus encabezados "Deposito ... (CODIGO-NOMBRE):") y
   pegalo directo en esa misma ventana (clic derecho > Pegar, o Ctrl+V).
   Al final escribe FIN en una linea aparte y Enter. No hace falta crear
   ningun archivo ni pegar nada en Excel.
   El script detecta solas las tablas del correo, su origen (columna
   DEPOSITO) y su destino (encabezado "Deposito ... (CODIGO-...)"), procesa
   todas las tablas en la misma corrida, y deja una copia de lo pegado en
   correo_FECHA.txt por si hay que revisarlo despues.

PASO 2 - sacar del txt los VIN que van con guia de despacho
5. Abrir el reporte_traslado_FECHA.xlsx que se genero. Un VIN que no pasa
   por el sistema (porque se le hace guia de despacho y llega directo por
   base) se marca con una X en la columna amarilla "excluir_guia", en la
   hoja que corresponda. Guardar el Excel.
6. Doble clic en  generar_txt_final.bat
   Genera UN solo archivo vins_traslado_FECHA_FINAL.txt con los VIN de
   todas las hojas del reporte (visible, restringido, etc.) juntos, en el
   mismo orden en que vinieron en el correo, sin los VIN marcados. Ese es
   el que se usa para pasar a PATIO por base.

Si ningun VIN de la tanda va con guia, se puede usar directamente el .txt
"borrador" del paso 1 (trae todos).


IMPORTANTE
----------
El script NO se conecta a la base: usa la hoja "Inventario Vehiculos Global"
tal cual viene en el Excel. Si esa hoja esta desactualizada, van a aparecer
VIN marcados en rojo como "no esta en la base" que en realidad si existen.
Pedir el Excel actualizado antes de procesar.

Dejar todos los archivos juntos en la misma carpeta.


ARCHIVOS
--------
traslado_solo_excel.bat   Paso 1: cruzar el correo contra la base (doble clic)
generar_txt_final.bat     Paso 2: generar el txt final sin los VIN con guia
procesar_traslado.py      las preguntas y el resumen final del paso 1
generar_txt_final.py      lee el reporte revisado y escribe el txt final
traslado.py               el proceso (parseo del correo, cruce, archivos)
requirements.txt          dependencias que instala solo
traslado_config.json      lo que elegiste la ultima vez
