# Formato EGR_FL_HYCITE: instalación y guía del operador

Etapa 2 del despacho: recibe los pedidos ya validados por PEDIDOS HCE, confirma el destino con reglas, imprime las
etiquetas, controla picking y empaque y exporta TMS, TRAMACO y DESPACHOS.

> **Decisión:** por ahora se mantienen **dos archivos**, cada uno con su panel. EGR tiene un modelo de datos y tablas
> dinámicas que no se pueden mover a PEDIDOS HCE sin rehacerlos, y en una hora de trabajo eso pondría en riesgo lo
> que hoy funciona. El paso 7 de PEDIDOS HCE ("Enviar a EGR") sigue siendo la unión entre los dos. La unificación queda
> planificada en [`docs/ANALISIS_EGR_Y_DECISION.md`](../../docs/ANALISIS_EGR_Y_DECISION.md).

## 1. Instalación (una sola vez, con una COPIA del archivo)

1. Guarda una copia del EGR. Abre el editor de VBA con **Alt + F11**.
2. **Módulo7**: bórralo (clic derecho › Quitar Módulo7 › No exportar). Sus macros `TABLAS`, `TABLA_APIS` y
   `ActualizarTodo` están ahora en `modEGR`, con los mismos nombres. El botón ACTUALIZAR de EMPAQUETADO sigue
   funcionando.
3. Inserta › Módulo, nómbralo **modEGR** y pega [`modEGR.bas`](modEGR.bas).
4. Inserta › Módulo, nómbralo **modZebra** y pega [`modZebra.bas`](modZebra.bas).
5. Inserta › Módulo, nómbralo **modVentanas** y pega [`modVentanas.bas`](modVentanas.bas) (es el mismo de PEDIDOS HCE).
6. Inserta › UserForm, nómbralo **frmEGR** y pega [`frmEGR.frm`](frmEGR.frm) en su código.
7. En **ThisWorkbook** pega [`ThisWorkbook.cls`](ThisWorkbook.cls).
8. Ejecuta **Depuración › Compilar VBAProject**. Si aparece un error, envía la captura con la línea marcada.
9. Guarda, cierra y vuelve a abrir. En **Complementos** aparece **Panel EGR (despacho)**.
10. En el panel, pulsa **Reparar fórmulas (una vez)**. Corrige los `#REF!` y crea la hoja **REGLAS_DESTINO**.

### Paso manual recomendado: sacar EMPAQUETADO del modelo de datos

La consulta EMPAQUETADO carga al **modelo de datos** y a la hoja al mismo tiempo. Refrescar el modelo con la macro
anterior fue la causa del mensaje "Error de automatización" y del cierre de Excel (sección 5).

Para quitar el riesgo del todo:

1. Datos › Consultas y conexiones. Clic derecho en **EMPAQUETADO** › **Cargar en...**
2. Elige **Tabla**, en la hoja EMPAQUETADO (celda B1), y **desmarca "Agregar estos datos al modelo de datos"**.
3. Revisa que las columnas F:K de la tabla (contenedora, bultos, peso, volumen, %) sigan con sus fórmulas.

Las tablas dinámicas no dependen del modelo: leen la hoja EMPAQUETADO.

## 2. Orden del día

| Paso | Dónde | Qué hacer |
|---|---|---|
| 0 | PEDIDOS HCE | Pasos 0 a 6, y después **7 Enviar a EGR** (con este archivo abierto) |
| 1 | Panel EGR | **Actualizar todo**: guarda un respaldo y trae ITEMS API, ITEMS DEPOT, EMPAQUETADO y las tablas dinámicas |
| 2 | Panel EGR | **Revisar cobertura TMS**: segunda comprobación, sin revalidar. Muestra solo lo que está fuera de la cobertura TMS, sin código postal o sin teléfono |
| 3 | Panel EGR | **Proponer destinos**: lista los pedidos que cambian de destino por reglas. Revisa y pulsa **Aplicar seleccionados** o **Aplicar todos** |
| 4 | Panel EGR | **Etiquetas › Imprimir pendientes**: una etiqueta por pedido en la Zebra |
| — | Bodega | Empaquetar y llenar el Google Sheets de empaquetado |
| 1 | Panel EGR | **Actualizar todo** otra vez, para traer bultos, pesos y contenedoras |
| 5 | Panel EGR | **Avance picking / empaque**: estado de cada pedido y SKU o cajas sin datos de costo |
| 6 | Panel EGR | **Exportar**: CSV, XLSX o PDF de TMS, TRAMACO y DESPACHOS. Con **Exportar + correo** se abre un correo de Outlook con los adjuntos |

Todo queda en el **registro** del panel y en la hoja oculta **LOG_EGR**: fecha, usuario, acción y error.

## 3. Reglas de destino (PRO / GYE / UIO / GPS)

El destino de DATOS!A alimenta TMS, TRAMACO, DESPACHOS y las etiquetas. Por eso las reglas **solo proponen**: el
operador ve la lista y confirma.

Las reglas viven en la hoja **REGLAS_DESTINO**, que el supervisor edita sin tocar código:

| Columna | Uso |
|---|---|
| ACTIVA | SI / NO |
| PRIORIDAD | Gana la regla activa de **menor** prioridad que coincida |
| PROVINCIA, CANTON, PARROQUIA | Vacío o `*` = cualquiera. Varias opciones separadas por `;` |
| TEXTO | Palabras buscadas en la dirección y en la parroquia. `KM>=15` = kilómetro 15 en adelante |
| EXCEPTO PARROQUIA | Parroquias excluidas de la regla |
| DESTINO | PRO, GYE, UIO o GPS |
| MOTIVO | Texto que ve el operador |

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

**Cómo se guarda:** el destino confirmado se escribe en DATOS!AF, junto con el número de pedido (AG) y el motivo (AH).
La fórmula de A lo usa solo si AG coincide con el pedido de esa fila. Así, al enviar los pedidos del día siguiente, un
destino viejo nunca se aplica a otro pedido.

## 4. Etiquetas Zebra ZD230 (203 dpi, 10 × 5 cm)

- Es **una etiqueta por pedido**, igual al modelo: código de barras Code 128 con el pedido y el número debajo; destino
  grande (GYE / UIO / PRO) arriba a la derecha y la parroquia debajo; nombres y apellidos abajo a la izquierda.
- Se envía **ZPL directo** a la impresora. No hace falta ninguna fuente de código de barras ni ajustar la página en
  Excel.
- La hoja de control es **ETIQUETAS**: al imprimir, la columna F (STATUS) pasa a **OK**, y "Imprimir pendientes" no las
  repite. **ETIQUETAS ZEBRA** (una fila por bulto) queda sin uso. No se borra, por si se necesitan etiquetas por
  bulto.
- **Elegir impresora** guarda el nombre de la Zebra una sola vez.
- **Guardar ZPL (prueba)** crea un archivo `.zpl`. Se puede ver en labelary.com antes de imprimir.
- Si la etiqueta sale corrida, ajusta `ETIQ_OFFSET_X` / `ETIQ_OFFSET_Y` (8 puntos = 1 mm) en la hoja oculta CONFIG_EGR.
- Los pedidos sin destino, por estar fuera de la cobertura TMS, no se imprimen y quedan en el registro.

## 5. "Error de automatización" al pulsar ACTUALIZAR (corregido)

**Causa:** la macro anterior recorría **todas** las conexiones del libro. Entre ellas refrescaba directamente el modelo
de datos (`ThisWorkbookDataModel` y `ModelConnection_DatosExternos_6`) y la consulta EMPAQUETADO, que también lo carga.
Todo se hacía seguido y sin control de errores. Si fallaba la red o Google Sheets, Excel lanzaba "Error de automatización" y
se cerraba, con todos los libros abiertos.

**Ahora `ActualizarTodo`:**

1. ofrece **guardar los otros libros abiertos**;
2. guarda un **respaldo** en `RESPALDOS_EGR` (se conservan los últimos 15);
3. refresca **una consulta a la vez**: primero las ODBC y después EMPAQUETADO por su consulta y su tabla, **nunca el
   modelo directamente**;
4. refresca cada caché de tablas dinámicas **una vez**;
5. si algo falla, lo **anota** (qué consulta y por qué), sigue con el resto y **libera la pantalla**.

## 6. Correcciones de fórmulas (botón "Reparar fórmulas")

| Dónde | Antes | Después |
|---|---|---|
| DATOS!Y (COURIER), 499 filas | `INDEX(#REF!, ...)` → vacío en los 152 pedidos | PRO = TRAMACO; UIO/GYE = ITSANET o LAAR COURIER según COBERTURA (misma regla que PEDIDOS HCE) |
| DATOS!A (DESTINO), 499 filas | Solo por provincia | Destino confirmado por reglas (AF) o, si no hay, la misma regla por provincia |
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

El paso 5 los informa ("tipo de caja sin medidas"). Se corrigen agregando la caja en DATA CODIGO Y CAJAS.

## 7. Exportación

- Solo se exportan las filas con datos, con **las mismas columnas y en el mismo orden**. Las hojas no se modifican.
- Las columnas de texto conservan el 0 inicial (teléfonos, códigos).
- Antes de exportar se revisan errores (`#N/A`), celdas con VALIDAR o Verificar y campos obligatorios vacíos. Todo
  queda en el registro y se pregunta si se exporta igual.
- **CSV:** UTF-8, separado por coma. Si el sistema que lo recibe pide punto y coma, pon `SI` en
  `CSV_SEPARADOR_PUNTOYCOMA` (hoja oculta CONFIG_EGR).
- **Carpeta:** `EXPORTES\aaaa-mm-dd`, junto al archivo. Otra ruta se configura en `CARPETA_EXPORTES`.
- **Correo:** destinatarios en `CORREO_PARA`. Se crea un **borrador** en Outlook; nunca se envía solo.

## 8. Qué verificar en la prueba

1. Compilar sin errores.
2. **Reparar fórmulas**: DATOS!Y deja de estar vacío y el PANEL deja de marcar "Sin courier asignado".
3. **Actualizar todo** con la red desconectada: debe mostrar el error en el registro **sin cerrar Excel**.
4. **Proponer destinos** con los pedidos actuales. Con las reglas iniciales cambian 5 pedidos, de GYE a PRO:
   - 2 de Samborondón / La Puntilla: confirma si Samborondón debe seguir en GYE. Si es así, agrega una regla
     `GUAYAS | SAMBORONDON | GYE` con prioridad 50;
   - 1 de El Triunfo;
   - 1 de Durán / El Recreo;
   - 1 de Guayaquil / Ximena porque la dirección dice GUASMO. Confirma que Guasmo va por PRO: en tu lista aparecía
     pegado a "EL EMPALME".
5. **Guardar ZPL (prueba)** y revisarlo en labelary.com (10 × 5 cm, 8 dpmm). Después imprimir una etiqueta.
6. **Exportar** en CSV y abrirlo en el bloc de notas para comprobar el separador y los ceros iniciales.
