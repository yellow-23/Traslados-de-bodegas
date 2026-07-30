TRASLADO DE BODEGA
==================

QUE HACE
--------
Toma los VIN que vienen en el correo, los cruza contra la hoja
"Inventario Vehiculos Global" del Excel y genera, por cada tabla del correo:

  - vins_traslado_FECHA_ORIGEN_DESTINO.txt      (para pegar en el sistema)
  - reporte_traslado_FECHA_ORIGEN_DESTINO.xlsx  (detalle de cada VIN)

En el reporte, las filas en rojo son las que hay que revisar antes de cargar:
VIN que no aparecen en la base, o que estan en una bodega distinta a la de
origen. Igual salen en el .txt, no se descartan.


COMO SE USA
-----------
1. Instalar Python desde https://www.python.org/downloads/
   IMPORTANTE: marcar la casilla "Add Python to PATH" al instalar.
   (Solo la primera vez, en un equipo nuevo.)

2. Pegar los VIN en la columna A de la hoja que corresponda del Excel
   (Hoja2, Hoja3, etc. - una hoja por cada tabla del correo).
   Guardar y cerrar el Excel.

3. Doble clic en  traslado_solo_excel.bat

4. Responde lo que pregunta:
   - que Excel usar (se abre un selector de archivos)
   - en que carpeta dejar los resultados
   - que hojas procesar y el par de bodegas de cada una (origen -> destino)

   La primera vez instala solo lo que necesita y demora un poco.
   Las siguientes veces recuerda lo que elegiste: si esta bien, das Enter.


IMPORTANTE
----------
El script NO se conecta a la base: usa la hoja "Inventario Vehiculos Global"
tal cual viene en el Excel. Si esa hoja esta desactualizada, van a aparecer
VIN marcados en rojo como "no esta en la base" que en realidad si existen.
Pedir el Excel actualizado antes de procesar.

Dejar todos los archivos juntos en la misma carpeta.


ARCHIVOS
--------
traslado_solo_excel.bat   lo que se ejecuta (doble clic)
procesar_traslado.py      las preguntas y el resumen final
traslado.py               el proceso (cruce, observaciones, archivos de salida)
requirements.txt          dependencias que instala solo
traslado_config.json      lo que elegiste la ultima vez
