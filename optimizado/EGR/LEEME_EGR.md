# Formato EGR_FL_HYCITE: instalación y guía del operador

Etapa 2 del despacho: recibe los pedidos ya validados por PEDIDOS HCE, confirma el destino (PRO / GYE / UIO / GPS),
imprime las etiquetas, sigue el empaque y exporta TMS, TRAMACO y DESPACHOS.

> **Decisión:** por ahora se mantienen **dos archivos**, cada uno con su panel. EGR tiene un modelo de datos y tablas
> dinámicas que no se pueden mover a PEDIDOS HCE sin rehacerlos, y hacerlo ahora pondría en riesgo lo que hoy
> funciona. El paso 7 de PEDIDOS HCE ("Enviar a EGR") une los dos. La unificación queda planificada en
> [`docs/ANALISIS_EGR_Y_DECISION.md`](../../docs/ANALISIS_EGR_Y_DECISION.md).

## 1. Instalación (una sola vez, con una COPIA del archivo)

1. Guarda una copia del EGR. Abre el editor de VBA con **Alt + F11**.
2. **Módulo7**: bórralo (clic derecho › Quitar Módulo7 › No exportar). Sus macros `TABLAS`, `TABLA_APIS` y
   `ActualizarTodo` están ahora en `modEGR`, con los mismos nombres. El botón ACTUALIZAR de EMPAQUETADO sigue
   funcionando.
3. Inserta › Módulo, nómbralo **modEGR** y pega [`modEGR.bas`](modEGR.bas).
4. Inserta › Módulo, nómbralo **modZebra** y pega [`modZebra.bas`](modZebra.bas).
5. Inserta › Módulo, nómbralo **modVentanas** y pega [`modVentanas.bas`](modVentanas.bas) (es el mismo de PEDIDOS HCE).
6. Inserta › UserForm, nómbralo **frmEGR** y pega [`frmEGR.frm`](frmEGR.frm) en su código.
7. Inserta otro UserForm, nómbralo **frmEtiquetas** y pega [`frmEtiquetas.frm`](frmEtiquetas.frm): es la vista previa
   de etiquetas.
8. En **ThisWorkbook** pega [`ThisWorkbook.cls`](ThisWorkbook.cls).
9. Ejecuta **Depuración › Compilar VBAProject**. Si aparece un error, envía la captura con la línea marcada.
10. Guarda, cierra y vuelve a abrir. En **Complementos** aparece **Panel EGR (despacho)**.
11. En el panel, pulsa **Reparar fórmulas (una vez)**. Corrige los `#REF!` y crea la hoja **REGLAS_DESTINO**.
12. En **PEDIDOS HCE** vuelve a pegar su **Módulo1**. Ahora "7 Enviar a EGR" también escribe en DATOS!AI:AO el gestor,
    el destino, el trayecto, el tipo de entrega, la zona peligrosa y los gestores Q/R de la etapa 1. El panel de EGR
    los muestra.

### Paso manual recomendado: EMPAQUETADO solo a la hoja (sin modelo de datos)

La consulta EMPAQUETADO carga al **modelo de datos** y a la hoja al mismo tiempo. Eso es lo que provoca el "Error de
automatización" que cierra Excel (sección 5). El **filtro por fecha no cambia**: lo hace Power Query dentro de la
consulta (paso `Filas filtradas` = `Date.IsInCurrentDay([FECHA])`). Por eso, aunque el Google Sheets tenga todo el año,
a la hoja solo llegan las órdenes de hoy. Solo cambia **dónde se carga**.

1. Guarda y haz una copia del archivo.
2. **Datos › Consultas y conexiones**. En el panel de la derecha, clic derecho en **EMPAQUETADO** › **Cargar en...**
3. Elige **Tabla** y **Hoja de cálculo existente** = `=EMPAQUETADO!$B$1`.
4. **Desmarca "Agregar estos datos al modelo de datos"** y acepta. Si Excel avisa que se eliminará el modelo de datos
   o la tabla anterior, acepta: la tabla se vuelve a crear en el mismo lugar.
5. Excel crea la tabla solo con **B:E** (FECHA, # ORDEN, TIIPO CAJA, # CONTENEDORA). Pulsa **Alt + F8**, elige
   **RestaurarColumnasEmpaquetado** y **Ejecutar**. La macro le devuelve el nombre `EMPAQUETADO` y vuelve a poner las
   columnas calculadas F:K (contenedora, bultos, peso caja, vol. caja, vol. ítems, %).
6. Pulsa **Actualizar datos** en el panel. Con el Google Sheets sin órdenes de hoy, la macro avisa y pregunta si sigue
   con lo demás, sin cerrar Excel.

**Comprobar el filtro por fecha:** Datos › Consultas y conexiones › doble clic en EMPAQUETADO. En **PASOS APLICADOS**
debe estar `Filas filtradas`. Al seleccionarlo, la fórmula dice:

```
= Table.SelectRows(#"Tipo cambiado1", each Date.IsInCurrentDay([FECHA]))
```

Para traer otra fecha (por ejemplo, ayer y hoy), cambia esa línea por:

```
= Table.SelectRows(#"Tipo cambiado1", each [FECHA] >= Date.AddDays(Date.From(DateTime.LocalNow()), -1))
```

Las tablas dinámicas no dependen del modelo: leen la hoja EMPAQUETADO (C:K).

## 2. Panel EGR y orden del día

Al abrir el panel (Complementos › **Panel EGR (despacho)**) se ven **todos los pedidos** de DATOS, igual que en PEDIDOS HCE:

| Columna | Qué muestra |
|---|---|
| SEÑAL | `X FUERA TMS` fuera de cobertura TMS · `!! ZONA` zona peligrosa · `> CAMBIO` cambio de destino sugerido · `@ REIMPRIMIR` etiqueta impresa con otro destino · `! SECTOR` verificar sector · `* CONFIRMADO` destino confirmado · `OK` |
| PROVINCIA / CANTÓN / PARROQUIA | Cobertura TMS (DATOS O:Q) |
| DEST. | Destino actual (DATOS!A), el que usan TMS, TRAMACO, DESPACHOS y etiquetas |
| SUGER. | Destino que proponen las reglas. **Solo aparece cuando es distinto** del actual |
| COURIER / TRAYECTO / ZONA | Courier (DATOS!Y), trayecto TRAMACO y zona peligrosa de PEDIDOS HCE |
| ETIQUETA | PENDIENTE / IMPRESA / REIMPRIMIR |
| EMPAQUE | SIN PICKING / PICKING x/y / PICKEADO / EMPACADO n cajas |

Arriba de la lista están el **filtro rápido** (cambios sugeridos, fuera de cobertura TMS, zonas, etiquetas pendientes o
por reimprimir, sin empacar, destino PRO/GYE/UIO/GPS…) y el **buscador**. Al hacer clic en un pedido, el detalle
muestra la dirección, la cobertura TMS, el destino actual, el sugerido **y por qué**, el gestor en cobertura, el gestor
sugerido, el trayecto, la zona, la etiqueta y el empaque.

| Paso | Qué hacer |
|---|---|
| 0 | PEDIDOS HCE: pasos 0 a 6, y después **7 Enviar a EGR** (con este archivo abierto) |
| 1 | **Revisar cobertura TMS** y **Ver cambios sugeridos**. El operario decide: **Aplicar sugerencia** (seleccionados), **Aplicar todas las sugerencias**, **Asignar a mano** PRO/GYE/UIO/GPS, o dejarlo como está |
| 2 | **Etiquetas**: elige los pedidos en la lista (el filtro ayuda) y pulsa **Imprimir seleccionadas**, o **Imprimir todas**. Se abre la **vista previa**: allí se marcan o desmarcan etiquetas y se confirma la impresora |
| — | Bodega empaqueta y llena el Google Sheets |
| 3 | **Avance empaque**: picking, cajas, peso y volumen % por pedido (datos de las tablas dinámicas), más SKU o cajas sin datos de costo |
| 4 | **Exportar**: elige la **hoja** (TRAMACO, TMS, DESPACHOS o LAS TRES) y el **formato** (CSV, XLSX, PDF). Debajo se ve cuántas filas con datos tiene cada hoja |

**Guía del flujo (como en PEDIDOS HCE):**

- el panel resalta en naranja el **siguiente paso**, con `>> <<`, y marca como `(hecho)` los terminados;
- el cuadro amarillo de la izquierda explica qué falta. Por ejemplo: "SIGUIENTE: decidir los destinos. 5 pedidos
  tienen un destino sugerido distinto…"

**Actualizar datos (items, empaque, tablas)** es un botón de **operación**: se usa en cualquier momento para traer
ITEMS API, ITEMS DEPOT, EMPAQUETADO y las tablas dinámicas. No es el inicio del proceso.

**Si se reasigna un destino después de imprimir o exportar:**

1. el pedido aparece como `@ REIMPRIMIR`;
2. filtra **REIMPRIMIR ETIQUETA** y vuelve a imprimir;
3. exporta otra vez los reportes.

Todo queda en el **registro** del panel y en la hoja oculta **LOG_EGR**: fecha, usuario, acción y error.

## 3. Reglas de destino (PRO / GYE / UIO / GPS)

El destino de DATOS!A alimenta TMS, TRAMACO, DESPACHOS y las etiquetas. Por eso las reglas **solo proponen**, en la
columna SUGER. del panel, y el operario confirma o deja el destino como está.

Las reglas viven en la hoja **REGLAS_DESTINO**, que el supervisor edita sin tocar código:

| Columna | Uso |
|---|---|
| ACTIVA | SI / NO |
| PRIORIDAD | Gana la regla activa de **menor** prioridad que coincida |
| PROVINCIA, CANTON, PARROQUIA | Vacío o `*` = cualquiera. Varias opciones separadas por `;` |
| TEXTO | Palabras buscadas en la dirección y en la parroquia. `KM>=15` = kilómetro 15 en adelante |
| EXCEPTO PARROQUIA | Parroquias excluidas de la regla |
| DESTINO | PRO, GYE, UIO o GPS |
| MOTIVO | Texto que ve el operario |

La comparación no distingue tildes ni paréntesis: `ELOY ALFARO (DURAN / ...)` = `ELOY ALFARO`.

Reglas iniciales (las que pediste):

| Prioridad | Regla | Destino |
|---|---|---|
| 10 | Guayas / Daule / parroquia Daule | PRO |
| 20 | Guayas / Daule / resto de parroquias | GYE |
| 30 | Guayas / Guayaquil con COOP o COOPERATIVA en la dirección | PRO |
| 40 | Guayas / Guayaquil con GUASMO, BAUTISTA AGUIRRE, BASTION POPULAR, SALITRE, FORTIN DE LA FLOR, PALESTINA, SIMON BOLIVAR, MAPASINGUE o KM ≥ 15 | PRO |
| 90 | Guayas / Guayaquil (resto) | GYE |
| 95 | Guayas, cualquier otro cantón. Cubre toda tu lista: Durán, Milagro, Playas, Salitre, Naranjal, Yaguachi, Bucay, Balzar, El Empalme, Nobol… | PRO |
| 100 | Azuay: Gualaceo, Paute, Chordeleg, Camilo Ponce Enríquez, Pucará, Sígsig | PRO |
| 110 | Azuay / Cuenca rural: Molleturo, Chaucha, Sayausí, Llacao, Paccha, Nulti, Quingeo, Cumbe, Irquis, Santa Ana | PRO |
| 900 / 910 / 999 | Base: Pichincha = UIO, Galápagos = GPS, resto = PRO | — |

**Ejemplo de cambio:** si mañana GYE ya cubre Durán, agrega la fila `SI | 50 | GUAYAS | DURAN | | | | GYE | Durán lo
cubre GYE`. Para UIO pasa lo mismo, por ejemplo `PICHINCHA | RUMIÑAHUI | ... | PRO`.

**Cómo se guarda:** el destino confirmado se escribe en DATOS!AF, junto con el número de pedido (AG) y el motivo y el
usuario (AH). La fórmula de A lo usa solo si AG coincide con el pedido de esa fila. Así, al enviar los pedidos del día
siguiente, un destino viejo nunca se aplica a otro pedido.

## 4. Etiquetas Zebra ZD230 (203 dpi, 10 × 5 cm)

- Es **una etiqueta por pedido**, igual al modelo: código de barras Code 128 con el pedido y el número debajo; destino
  grande (GYE / UIO / PRO) arriba a la derecha y la parroquia debajo; nombres y apellidos abajo a la izquierda.
- Solo hay dos botones: **Imprimir seleccionadas** e **Imprimir todas**. Los dos abren la **vista previa**
  (frmEtiquetas):
  - lista de pedidos, marcados para imprimir, que se pueden desmarcar;
  - dibujo de la etiqueta del pedido elegido;
  - impresora.
- La impresora se elige **una vez** y queda guardada; la próxima vez ya aparece seleccionada.
- **Cierre de Excel al elegir la impresora:** venía del gancho de la rueda del mouse, que seguía activo mientras se
  abría el cuadro de diálogo. Ahora el gancho se quita antes de cualquier acción, y la lista de impresoras se lee de
  Windows de forma liviana.
- Se envía **ZPL directo** a la impresora, sin fuentes de código de barras.
- Al imprimir, en la hoja ETIQUETAS se escribe:
  - F (STATUS) = **OK**;
  - M = destino impreso;
  - N = fecha.

  Si después cambia el destino, el pedido pasa a **REIMPRIMIR**. ETIQUETAS ZEBRA queda sin uso.
- **Guardar ZPL (prueba)** crea un archivo `.zpl`. Se puede ver en labelary.com.
- Si la etiqueta sale corrida, ajusta `ETIQ_OFFSET_X` / `ETIQ_OFFSET_Y` (8 puntos = 1 mm) en la hoja oculta CONFIG_EGR.

## 5. "Error de automatización" al pulsar ACTUALIZAR (corregido)

**Causa:** la consulta EMPAQUETADO carga al **modelo de datos**. Cuando el Google Sheets no tiene órdenes del día (hoja
vacía para hoy) o no responde, refrescarla deja al modelo con error. Excel lanza "Error de automatización" y se cierra
con todos los libros abiertos. VBA no puede atrapar ese cierre.

**Ahora `ActualizarTodo`:**

1. ofrece **guardar los otros libros abiertos** y guarda un **respaldo** en `RESPALDOS_EGR`;
2. **antes** de tocar EMPAQUETADO, descarga el Google Sheets y comprueba que responde y que tiene órdenes con fecha de
   hoy;
3. si no hay datos de hoy o no hay conexión, **no refresca EMPAQUETADO** y **pregunta**: "¿Sigo con las demás
   conexiones (ITEMS API, ITEMS DEPOT) y las tablas dinámicas?";
4. refresca las consultas ODBC **una a la vez**. Si una falla, pregunta si sigue;
5. refresca cada caché de tablas dinámicas **una vez**, anota todo en el registro, **libera la pantalla** y recarga el
   panel.

Para quitar el riesgo del todo, haz también el paso manual de la sección 1: EMPAQUETADO solo a tabla, sin modelo de
datos.

## 6. Correcciones de fórmulas (botón "Reparar fórmulas")

| Dónde | Antes | Después |
|---|---|---|
| DATOS!Y (COURIER), 499 filas | `INDEX(#REF!, ...)` → vacío en los 152 pedidos | PRO = TRAMACO; UIO/GYE = ITSANET o LAAR COURIER según COBERTURA (misma regla que PEDIDOS HCE) |
| DATOS!A (DESTINO), 499 filas | Solo por provincia | Destino confirmado (AF) o, si no hay, la misma regla por provincia |
| EMPAQUETADO!F, 145 filas | `COUNTIF(#REF!, ...)` | `COUNTIF(ITEMS_API[DOC_EXT], ...)`: "validar" = pedido del día sin contenedora; "sin pedido" = no es del día |
| DESPACHOS!P, 500 filas | `CHOOSE({2,1}, DATOS!A:A, DATOS!B:B)`, muy lento | `INDEX/MATCH` sobre DATOS!A2:B500 |
| COD POS I:J, 2 606 celdas | `XLOOKUP` a un libro de otra PC en la red → `#REF!` | Vacías (ninguna hoja las usa) |
| Nombres definidos | 12 nombres con `#REF!` o apuntando a otra copia del archivo | Eliminados |
| Vínculos externos | Carpeta de red y la propia copia del archivo | Rotos (quedan como valores) |
| DATA CODIGO Y CAJAS | 185 reglas de formato "duplicados", varias de columna completa | 1 regla (B2:B1000) |

Las demás fórmulas de cada columna se revisaron: son iguales en todas sus filas y dan resultado.

Quedan 3 valores `#N/A` que no son errores de fórmula:

- EMPAQUETADO H, I, K: un tipo de caja que no está en DATA_CAJAS;
- TABLAS DINAMICAS J, K, Q: el mismo caso.

"Avance empaque" los informa en el registro. Se corrigen agregando la caja en DATA CODIGO Y CAJAS.

## 7. Exportación

- Se elige la **hoja** (TRAMACO, TMS, DESPACHOS o LAS TRES) y el **formato** (CSV, XLSX o PDF) en listas desplegables.
- Se exportan **solo las filas con datos de esa hoja**, con **las mismas columnas y en el mismo orden**, porque son
  plantillas de carga. Las hojas no se modifican.
  - **TRAMACO** lleva solo los pedidos PRO: toma las filas con nombre (C) y pedido (AD). En los datos actuales son 70.
  - **TMS** lleva **todos** los pedidos del día (152), porque así está armada la hoja TMS: una fila por cada pedido de
    DATOS.
  - **DESPACHOS** lleva todos los pedidos (152).

  Antes de exportar, el panel muestra cuántas filas tiene cada hoja.
- Las columnas de texto conservan el 0 inicial (teléfonos, códigos).
- Antes de exportar se revisan errores (`#N/A`), celdas con VALIDAR o Verificar y campos obligatorios vacíos. Todo
  queda en el registro y se pregunta si se exporta igual.
- **CSV:** UTF-8, separado por coma. Si el sistema que lo recibe pide punto y coma, pon `SI` en
  `CSV_SEPARADOR_PUNTOYCOMA` (hoja oculta CONFIG_EGR).
- **Carpeta:** `EXPORTES\aaaa-mm-dd`, junto al archivo. Otra ruta se configura en `CARPETA_EXPORTES`.
- **Correo:** destinatarios en `CORREO_PARA`. Se crea un **borrador** en Outlook; nunca se envía solo.

## 8. Qué verificar en la prueba

1. Compilar sin errores: módulos modEGR, modZebra y modVentanas; formularios frmEGR y frmEtiquetas.
2. **Reparar fórmulas**: DATOS!Y deja de estar vacío.
3. **Actualizar datos** con el Google Sheets sin órdenes de hoy. Debe preguntar si sigue con las demás conexiones,
   **sin cerrar Excel**.
4. Abrir el panel: aparecen los 152 pedidos. Con el filtro **CAMBIO SUGERIDO** aparecen 5, de GYE a PRO:
   - 2 de Samborondón / La Puntilla: confirma si Samborondón debe seguir en GYE. Si es así, agrega una regla
     `GUAYAS | SAMBORONDON | GYE` con prioridad 50;
   - 1 de El Triunfo;
   - 1 de Durán / El Recreo;
   - 1 de Guayaquil / Ximena porque la dirección dice GUASMO. Confirma que Guasmo va por PRO.
5. **Imprimir seleccionadas** con 1 pedido: se ve la vista previa. Elige la Zebra e imprime. Vuelve a imprimir: la
   impresora ya aparece elegida.
6. **Exportar** TRAMACO en CSV: el archivo debe tener las mismas filas que la hoja TRAMACO.
