# Conciliación de pallets ESPI

App de Streamlit con tres modos (elige en la barra lateral):

## 🚚 Pallets y embarques (NUEVO — prioridad de Sherlyn, base de datos compartida)

A diferencia de los otros dos modos (que trabajan por sesión/Excel), este
módulo usa una **base de datos compartida real** (`espi.db`, SQLite): todos
los que entren a la app —desde cualquier computadora que la tenga
desplegada— ven y editan los mismos datos, en tiempo real.

- **Productores**: nombre, lote, teléfono, empresa, huerta.
- **Pallets**: # de pallet, productor, variedad, calibre (con sugerencia
  automática de cajas según el calibre — editable, ej. 225 estándar, 240 si
  lleva doble línea, 40 si es 6x8, 75 si es 6x6), orgánico, estado (en
  cuarto frío / cargando / en viaje / entregado).
  - **Deshacer un pallet**: lo divide en 2 o más pallets nuevos (para cuando
    se rompe o se reparte entre productores/variedades).
  - **Armar (fusionar) pallets**: junta 2 o más pallets en uno nuevo, sumando
    cajas (para cuando se mixtean por no completarse).
- **Embarques**: crear un viaje (fecha, chofer, teléfono, placas, destino,
  contacto en destino), asignarle pallets disponibles, cambiar su estado
  (armando/cargando/en viaje/entregado — se refleja también en los pallets
  que lleva), y ver el **resumen por productor** (cuántos pallets y cajas le
  tocan de ese viaje) — que es justo el reporte que se le entrega a cada
  productor.

**Importante sobre esta base de datos:** vive en un archivo `espi.db` junto
al código. Si se despliega en Streamlit Community Cloud (gratis), ese
archivo se puede borrar cuando la app se reinicia por inactividad o al subir
un cambio de código.

Por eso el módulo trae un panel de **"💾 Respaldo"** arriba de todo (se abre
al hacer clic): descarga un Excel con todo lo capturado (productores,
pallets, embarques) cuando quieras, y si un día abres la app y ves todo
vacío, sube ahí el último respaldo para restaurarlo completo — incluyendo a
qué embarque va cada pallet. Conviene bajar un respaldo seguido (ej. al
terminar cada turno) para no perder trabajo. Cuando quieran una solución que
no dependa de acordarse de bajar el respaldo, la alternativa es migrar a una
base en la nube persistente (ej. Supabase) — el acceso a datos está aislado
en `db.py` justo para que ese cambio no obligue a rehacer el resto de la
app.

## 📝 Captura de monitoreos

Reemplaza el llenado manual del Excel de LIQUIDACIÓN para el registro de
rezaga/merma/jugo por corte. Vas capturando:

1. **Productores** (nombre + número de lote).
2. **Monitoreos** por productor (fecha, folio, kilos recibidos/empacados,
   merma, jugo) — cada corte que le haces a un productor.
3. **Pallets** de cada monitoreo (# pallet, calibre, cajas) en una tabla tipo
   Excel donde puedes escribir o pegar filas.

Como cada pallet se captura *dentro* de su monitoreo, la app ya sabe en qué
monitoreo está cada uno — no hay que buscarlo. Al subir el archivo RCF
(que sigue generando el otro sistema/persona) en la barra lateral, la pestaña
**"Resumen y conciliación"** cruza automáticamente lo capturado contra el RCF
y muestra, por productor y por calibre: pallets encontrados, con diferencia,
no encontrados (con su número) y "sobrantes" (pallets del RCF que
pertenecen a ese lote/productor pero que aún no capturaste en ningún
monitoreo).

**Guardar tu trabajo:** no hay base de datos — la captura vive en la sesión
del navegador. Usa el botón **"💾 Descargar captura"** de la barra lateral
para bajar un Excel con todo lo capturado, y **"📂 Continuar una captura
guardada"** para volver a subirlo la próxima vez y seguir donde quedaste (o
para que otra persona continúe la captura). Al final también puedes
descargar un **"reporte de conciliación"** en Excel, ya con colores y el
detalle por productor, para archivar o imprimir.

## 📂 Reconciliar Excel existente (modo de compatibilidad)

El modo original: sube un Excel de LIQUIDACIÓN ya armado con el formato
viejo de bloques por día (como los que ya tenías) y el archivo RCF, y la app
hace el mismo cruce pero escribiendo el resultado de vuelta en una copia de
ese Excel (marcas de color + tabla de conciliación al final de cada hoja).
Útil para procesar temporadas o archivos anteriores sin tener que
recapturar todo.

## Cómo correrla

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cómo funciona el cruce contra RCF (aplica a ambos modos)

1. **LOTE + # pallet como llave principal de búsqueda**: un mismo pallet
   físico puede repartirse entre dos lotes distintos en el RCF (pallets
   mixtos), así que nunca se busca solo por número de pallet.
2. **Nombre del productor como respaldo**: un mismo número de lote puede
   estar compartido por varios productores (pasa con el lote 4 y el lote 8
   en los archivos de ejemplo), así que el cruce siempre compara también el
   nombre del productor, tolerando variantes/errores de escritura ("RAFAEL
   B." vs "RAFAEL BALDERRAMA", "MAICO FELIX" vs "MAICO FELIZ").
3. **Cajas/calibre no capturados no cuentan como diferencia**: si no
   registras cajas por pallet, el pallet se marca como encontrado con solo
   que el número exista en el RCF para ese lote/productor.
4. **Deduplicación del RCF**: el archivo RCF trae una pestaña maestra que
   repite todo lo que ya está en las pestañas individuales de cada
   manifiesto — se deduplica automáticamente para no inflar los conteos.

## Archivos

- `app.py` — router: elige entre los tres modos (este es el que se configura
  como "main file" al hacer el deploy).
- `mode_pallets.py` + `db.py` — módulo de pallets/embarques con base de
  datos compartida (SQLite).
- `mode_captura.py` + `capture_core.py` — modo de captura de monitoreos por
  sesión (interfaz y lógica).
- `mode_excel.py` + `core.py` — modo de compatibilidad con Excel existente
  (interfaz y lógica).
- `requirements.txt` — dependencias.

## Qué revisar en la próxima temporada / con más pestañas de RCF

- Si aparecen pestañas de manifiesto en el RCF con encabezados muy distintos
  a "FECHA / # PALLET / LOTE / CAJAS / CALIBRE / PRODUCTOR / MANIFIESTO", la
  app las reporta como "omitidas" en vez de fallar — conviene revisar esa
  lista después de cada corrida.
- El umbral de similitud de nombres de productor está en `core.py`
  (`PRODUCTOR_MATCH_THRESHOLD = 0.45`) — si ves cruces raros entre
  productores con nombres parecidos, se puede subir ese número.
