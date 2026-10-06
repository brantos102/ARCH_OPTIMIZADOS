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
7. Inserta otro UserForm, nómbralo **frmEtiquetas** y pega [`frmEtiquetas.frm`](frmEtiquetas.frm): lista filtrable y
   vista previa de etiquetas. Inserta un tercero, nómbralo **frmReglas** y pega [`frmReglas.frm`](frmReglas.frm): editor de
   las reglas de destino.
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

### EMPAQUETADO trae miles de filas vacías (37 109 filas con la fecha de hoy)

**Causa:** en el Google Sheets los operadores llenan la FECHA hacia abajo en miles de filas antes de registrar las
cajas. La consulta anterior filtraba por fecha, pero dejaba pasar esas filas sin número de orden, así que la tabla, las
fórmulas F:K y las tablas dinámicas procesaban 37 mil filas.

**Solución:** reemplaza el texto de la consulta por [`EMPAQUETADO.m`](EMPAQUETADO.m). La consulta nueva:

1. descarta primero las filas **sin "# ORDEN"**;
2. filtra la fecha de hoy, aceptando dd/mm/aaaa o mm/dd/aaaa;
3. deja la orden como texto limpio;
4. descarta un `# ORDEN` que no sea un número de pedido (8 dígitos o más, solo dígitos): un `5465443` tecleado
   por error entraba a la hoja y salía como *sin pedido* / *Verificar*;
5. **ordena por FECHA, # ORDEN y # CONTENEDORA**, así la hoja se presenta en orden y el primer pedido deja de
   quedar al final;
6. pone un tope de seguridad de 2 000 cajas.

Pasos:

1. Datos › Consultas y conexiones › clic derecho en **EMPAQUETADO** › **Editar**.
2. En Power Query: **Inicio › Editor avanzado**. **Copia la URL** de la línea `Origen = Csv.Document(Web.Contents("...")`.
3. Borra todo, pega el contenido de `EMPAQUETADO.m` y reemplaza `PEGA_AQUI_LA_URL` por la URL copiada (entre comillas).
4. **Listo** › **Cerrar y cargar**. La tabla queda solo con las cajas de hoy (decenas o cientos de filas).
5. Pulsa **Actualizar datos** y **guarda** el archivo, para que Excel libere el espacio de las filas que ya no existen.

No hace falta pedir a los operadores que dejen de llenar la fecha: la consulta ya ignora las filas vacías. Si quieres
que el Google Sheets también sea más liviano, la fecha se puede llenar solo cuando se registra la orden.

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

Arriba de la lista hay tres herramientas de búsqueda, que se combinan entre sí:

0. **Ordenar**: clic en cualquier encabezado de la lista ordena por esa columna; otro clic invierte el orden
   (la flecha ▲▼ marca cuál). Ordena solo lo que está en pantalla: no toca los datos ni la hoja DATOS. Al cambiar de
   vista (Pedidos / Empaque / Cobertura) el orden se quita, porque las columnas son otras.
1. **Filtro rápido**: cambios sugeridos, fuera de cobertura TMS, zonas, etiquetas pendientes o por reimprimir, sin
   empacar, destino PRO/GYE/UIO/GPS, difiere de PEDIDOS HCE…
2. **Tres filtros por columna**: se elige la columna (pedido, destinatario, provincia, cantón, parroquia, destino,
   sugerido, courier, trayecto, zona, etiqueta, empaque o fila) y se escribe el texto. Por ejemplo, GUAYAS + GUAYAQUIL +
   etiqueta PENDIENTE. **Quitar filtros** los vacía todos.
3. **Búsqueda libre** sobre todas las columnas. Al hacer clic en un pedido, el detalle
muestra la dirección, la cobertura TMS, el destino actual, el sugerido **y por qué**, el gestor en cobertura, el gestor
sugerido, el trayecto, la zona, la etiqueta y el empaque.

| Paso | Qué hacer |
|---|---|
| 0 | PEDIDOS HCE: pasos 0 a 6, y después **7 Enviar a EGR** (con este archivo abierto) |
| 1 | **Revisar cobertura TMS** y **Ver cambios sugeridos**. El operario decide: **Aplicar sugerencia** (seleccionados), **Aplicar todas las sugerencias**, **Asignar a mano** PRO/GYE/UIO/GPS, o dejarlo como está |
| 2 | **IMPRIMIR ETIQUETAS**: se abre la lista completa, se filtra (estado, destino, búsqueda) y se imprime lo que quede en la lista |
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

### Novedades del panel

- **Sugerencias:** además de las reglas de REGLAS_DESTINO, el sistema aplica la lógica de PEDIDOS HCE. Un pedido UIO o
  GYE cuya cobertura TMS no tiene ITSANET ni LAAR se sugiere como **PRO**; el motivo aparece en el detalle. El filtro
  **DIFIERE DE PEDIDOS HCE** muestra los pedidos cuyo destino no coincide con el que calculó la etapa 1. La opción
  manual (*Asignar a mano*) se mantiene.
- **Ir a hoja:** una lista desplegable con las hojas necesarias y el botón **Ir a hoja**, que la muestra aunque esté
  oculta. Las hojas son DATOS, TMS, TRAMACO, DESPACHOS, ETIQUETAS, EMPAQUETADO, TABLAS DINAMICAS, ITEMS APIS,
  ITEMS DEPOT, COBERTURAS Y TARIFAS, DATA CODIGO Y CAJAS, REGLAS_DESTINO y PANEL.
- **Cobertura:** la vista **Cobertura** muestra COBERTURAS Y TARIFAS completa (gestor, gestor sugerido, trayecto, tipo,
  días y código postal), con buscador, para validarla como en PEDIDOS HCE.
- **Productos (códigos)** y **Cajas:** abren el formulario de datos de Excel sobre DATA CODIGO Y CAJAS. Con *Nuevo* se
  agrega un registro, con *Criterios* se busca, y los datos se editan directamente. Los códigos o cajas nuevos se usan
  de inmediato en pesos, volúmenes y costos.
- **Estabilidad con los dos archivos abiertos:**
  - el gancho de la rueda del mouse **se quita solo** cuando el mouse sale del formulario, en PEDIDOS HCE y en EGR
    (el mismo modVentanas en ambos);
  - también se quita antes de cualquier botón.

  Así no queda activo durante mensajes, macros largas ni al cambiar de libro. **Pega el modVentanas nuevo en los dos
  archivos.**

### Operar sin el panel: botones en Complementos

La barra **Despacho EGR HYCITE** (pestaña Complementos) tiene **todas** las funciones del panel, sin abrir formularios.
Están en el orden del día, numeradas:

| Botón | Qué hace |
|---|---|
| Panel EGR | Abre el panel (opcional) |
| **¿Qué sigue?** | Cuenta el estado de hoy (fuera de cobertura, cambios sugeridos, etiquetas, empaque) y dice cuál es el siguiente paso. Si no sabes por dónde seguir, empieza aquí |
| **0 Limpiar día** | Borra la validación anterior de DATOS y el estado de las etiquetas. Pregunta si también borra los pedidos de ayer. Guarda un respaldo antes |
| 1 Revisar cobertura TMS | Lista los pedidos fuera de cobertura, sin código postal o sin teléfono, y va al primero |
| 2 Ver cambios sugeridos | Muestra qué pedidos cambian de destino y por qué, sin aplicar nada |
| 2b Aplicar destinos sugeridos | Confirma todos los cambios de una vez |
| 2c Asignar destino a un pedido | Pide el número de pedido y el destino nuevo: PRO, GYE, UIO o GPS |
| 3 IMPRIMIR ETIQUETAS | Abre la lista completa, se filtra y se imprime lo filtrado |
| 3b Imprimir pendientes | Imprime directamente las pendientes y las que cambiaron de destino |
| 3c Reimprimir pedido(s) | Pide los números de pedido |
| 4 Avance empaque | Resumen de picking y empaque |
| Trazabilidad | Arma la hoja TRAZABILIDAD y ofrece guardar una copia en la carpeta de exportes |
| Datos anteriores | Abre HISTORICO_EMPAQUE: las cajas de los días anteriores, con filtro por fecha |
| Respaldo del día (.xlsb) | Guarda EGR_DIA_aaaa-mm-dd.xlsb en RESPALDOS_EGR, ya congelado |
| Congelar como histórico | Convierte la COPIA abierta en un histórico de solo consulta |
| 5 Exportar reportes | Pide la hoja (1 TRAMACO, 2 TMS, 3 DESPACHOS, 4 las tres) y el formato (1 CSV, 2 XLSX, 3 PDF) |
| Actualizar datos | Trae ITEMS API, ITEMS DEPOT, EMPAQUETADO y las tablas dinámicas |
| **Reglas de destino** | Abre el editor de REGLAS_DESTINO |
| Productos · Cajas | Alta y edición del maestro de productos y de cajas |
| Impresora | Elige la Zebra |
| **Ir a hoja** | Lista desplegable con las hojas del día: elige y te lleva, aunque esté oculta |
| Reparar fórmulas · Desbloquear | Igual que en el panel |

### Editor de reglas de destino

**Reglas de destino**, en el panel o en Complementos, abre una ventana con las reglas en orden de prioridad:

- la lista muestra si está activa, su prioridad, provincia, cantón, destino y motivo;
- la ficha de la derecha edita cualquier campo, y **Guardar** la escribe en la hoja y la reordena por prioridad;
- **Nueva regla** y **Eliminar** dan de alta y de baja, con confirmación;
- la casilla **Activa** desactiva una regla sin borrarla;
- **Probar con un pedido**: escribes provincia, cantón, parroquia y dirección, y muestra qué regla gana y qué destino
  propone. Sirve para comprobar una regla antes de dejarla activa.

Cada alta, cambio o baja queda en el registro LOG_EGR.

### Impresora Zebra compartida ("ZDesigner ZD230-203dpi ZPL en emphycite")

Windows muestra las impresoras compartidas como `IMPRESORA en SERVIDOR`, pero su nombre real es
`\\SERVIDOR\IMPRESORA`. El envío prueba en este orden:

1. el nombre tal como aparece;
2. `\\emphycite\ZDesigner ZD230-203dpi ZPL`;
3. como último recurso, copia el ZPL directo al recurso compartido (`copy /b`).

Cada intento y su código de error de Windows quedan en el registro, así se ve exactamente por qué falló.

Si falla todo, en la PC donde está conectada la Zebra (emphycite): Panel de control › Dispositivos e impresoras ›
Zebra › Propiedades de impresora › **Compartir**. Anota el **nombre del recurso compartido** y escríbelo en CONFIG_EGR
(`IMPRESORA_ZEBRA`) como `\\emphycite\NOMBRE_COMPARTIDO`.

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

- Es **una etiqueta por pedido**, con el formato de la muestra impresa, repartido en bandas que **no se pisan**
  (800 × 400 puntos a 203 ppp):

  | Banda (puntos) | Contenido |
  |---|---|
  | 26 – 138 | código de barras Code 128 |
  | 140 – 196 | número del pedido, centrado bajo las barras |
  | 186 – 282 | **destino en grande** a la derecha (UIO, GYE, PRO, GPS) |
  | 216 – 322 | **destinatario** a la izquierda, en uno o dos renglones |
  | 326 – 362 | **parroquia** a la derecha |

- **Nombres encimados (corregido).** El nombre se repartía con `^FB ancho,2 líneas`. Cuando no entraba en esos
  dos renglones, la impresora escribía el sobrante **encima del último renglón**, y salían dos textos
  superpuestos. Ahora el nombre se parte **en el código**, sin cortar palabras, y cada renglón se imprime en su
  propia posición. Además la letra se achica sola (48 → 42 → 36 → 30 → 26 puntos) hasta que el nombre entra; si
  ni con la más chica entra, se recorta, pero **nunca se escribe encima**. La parroquia se recorta a 23
  caracteres por la misma razón.
- Hay **un solo botón**: **IMPRIMIR ETIQUETAS** (en el panel, en la pestaña Complementos y con Alt + F8 ›
  `MenuEtiquetas`). No hay que seleccionar nada antes: la ventana abre con **todos** los pedidos con destino y se
  imprime **exactamente lo que quede en la lista**.
- Dentro de la ventana:
  - **filtro por estado**: pendientes + por reimprimir (así abre), pendientes, por reimprimir, impresas o todas;
  - **destinos con casillas**: PRO, GYE, UIO y GPS se **combinan**. Para un lote mezclado se dejan marcados los
    que toque (p. ej. GYE + UIO y PRO fuera) y la lista queda solo con esos. Abre con los cuatro marcados;
  - **búsqueda** por pedido, destinatario, parroquia o destino, mientras se escribe;
  - **Quitar filtros** vuelve a la lista completa con los cuatro destinos marcados;
  - la lista sale **agrupada por destino y, dentro de cada destino, por número de pedido**, así el lote se imprime
    ordenado y es fácil separarlo después;
  - **selección opcional**: sin seleccionar nada se imprime **toda la lista**; si se seleccionan filas
    (Ctrl + clic o Shift + clic, o **Seleccionar todo** / **Quitar selección**) se imprime **solo lo
    seleccionado**. El contador y el botón lo dicen en todo momento: *Imprimir las 23 de la lista* o
    *Imprimir las 5 seleccionada(s)*;
  - vista previa de la etiqueta del pedido en el que se hace clic;
  - antes de imprimir se confirma con los filtros aplicados y si va la lista entera o solo lo seleccionado.
- Sigue habiendo **un solo botón para abrir**: no hay que marcar nada antes en el panel.
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
  Si sale clara o muy quemada, sube o baja `ETIQ_OSCURIDAD` (0 a 30; por defecto 12).

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
| DATOS!R (BULTOS) | `VLOOKUP` a la dinámica "Suma de BULTOS" | `COUNTIF` de las contenedoras del pedido en la hoja EMPAQUETADO; mínimo 1 |
| EMPAQUETADO!G (BULTOS) | `1` en cada fila | Total de contenedoras del pedido |
| EMPAQUETADO!J (VOL. ITEMS) | Volumen de la contenedora o, si no estaba, el de TODO el pedido | Solo el volumen de esa contenedora; si no lo tiene, `Verificar` |

Las demás fórmulas de cada columna se revisaron: son iguales en todas sus filas y dan resultado.

Quedan 3 valores `#N/A` que no son errores de fórmula:

- EMPAQUETADO H, I, K: un tipo de caja que no está en DATA_CAJAS;
- TABLAS DINAMICAS J, K, Q: el mismo caso.

"Avance empaque" los informa en el registro. Se corrigen agregando la caja en DATA CODIGO Y CAJAS.

### Cantidades, pesos y costos duplicados en ITEMS DEPOT

**Síntoma:** un pedido de 4 productos en 2 cajas aparecía con 8 filas y su peso y su costo salían al doble.

**Causa:** la consulta `Estado` unía `VIEW_TIEMPO_EMPAQUETADO` fila a fila. Esa vista devuelve una fila por
contenedora y **no sabe qué producto va en cuál**, así que cada producto se repetía **con la cantidad completa** en
cada caja. El `SELECT DISTINCT` no lo evitaba, porque las filas sí eran distintas: cambiaba la contenedora.

**Las dos reglas que hay que cumplir a la vez:**

1. cada ítem sale con **la contenedora en la que realmente se empacó** — de ahí salen el peso y el volumen por caja,
   y el **box density** (volumen de los ítems de la caja ÷ volumen de la caja = % de ocupación);
2. el **número de cajas** del pedido es el **total de contenedoras**, sin importar cuántos ítems lleve cada una.

**Solución:** reemplaza el SQL con [`SQL_ITEMS_DEPOT.sql`](SQL_ITEMS_DEPOT.sql).

- La contenedora de cada ítem sale de **`PICKING.NRO_UCEMPAQUETADO`**, que es la caja en la que se empacó esa línea
  de picking. El picking se agrupa por **documento + producto + contenedora**, así que:
  - un producto que fue en una sola caja da **una fila**;
  - un producto repartido en dos cajas da **dos filas**, cada una con la cantidad que fue en esa caja. Sumadas dan la
    cantidad real: el peso, el volumen y el costo del pedido no se duplican.
- Si una línea todavía no está empacada, la contenedora queda **vacía**; y si el pedido tiene **una sola**
  contenedora, se le asigna esa (no hay ambigüedad posible). Con dos o más, no se adivina: queda vacía hasta que se
  empaque.
- La consulta devuelve **exactamente las 7 columnas de siempre**, con los mismos nombres y en el mismo orden. En
  ITEMS DEPOT no se agrega ninguna columna: sus columnas calculadas H:K (PEDIDO, PESO, VOLUMEN, PRECIO) siguen en su
  sitio. El total de cajas del pedido se resuelve en EMPAQUETADO (ver más abajo), no aquí.
- **Dos intentos que se descartaron**, por si vuelven a aparecer: poner la misma contenedora (`MIN`) en todas las
  filas del pedido, y agregar una columna con la lista de contenedoras. Las dos rompen el box density, porque el
  volumen del pedido entero se le carga a una sola caja.

Ejemplo con el pedido 102344955 (4 productos, 2 cajas):

| | Antes | Ahora |
|---|---|---|
| Filas | 8 | 4, 5 o 6 según cómo se repartieron los productos entre las 2 cajas |
| Contenedora de la fila | las 2, repetidas en cada producto | la caja en la que fue ese producto |
| Cantidad por fila | repetida en las 2 cajas | la que fue en esa caja |
| Peso del pedido | 29,70 kg | 14,85 kg |
| CAJAS | — | 2 en todas sus filas |

**Antes de aplicarla, ejecuta el bloque 1 de [`SQL_VERIFICACION.sql`](SQL_VERIFICACION.sql)** en SSMS. Compara, por
pedido, las cajas que ve el picking con las que ve `VIEW_TIEMPO_EMPAQUETADO`: deben coincidir. Si en pedidos ya
empacados el picking devuelve 0, **no la apliques** y avísame: habría que sacar la contenedora de cada ítem de otra
columna. Los bloques 2 y 3 revisan el reparto ítem por caja, y el 4 confirma que el total de unidades no cambió.

Para aplicarla:

1. Datos › Consultas y conexiones › clic derecho en **Estado** › **Editar**.
2. En PASOS APLICADOS, pulsa el engranaje del paso **Origen**.
3. Borra el SQL y pega **todo** el archivo, tal cual. Aceptar › **Cerrar y cargar**.
4. Vuelve a Excel y pulsa **Reparar fórmulas**: ahí se aplica la fórmula nueva de BULTOS (ver abajo).

> El archivo no lleva comentarios, `WITH`, `ORDER BY` ni punto y coma final, y no anida tablas derivadas: con
> cualquiera de esas cosas el servidor devolvía *"Sintaxis incorrecta cerca de 'd'"* (error 102).

**Solo el día de hoy.** El filtro de fecha es:

```sql
AND syd_AD.FECHA_CREACION >= CAST(GETDATE() AS DATE)
AND syd_AD.FECHA_CREACION <  DATEADD(DAY, 1, CAST(GETDATE() AS DATE))
```

Antes era `DATEADD(DAY, -1, ...)`, o sea **desde ayer**, y por eso en ITEMS DEPOT aparecían mezclados los pedidos de
ayer y los de hoy. Ahora la hoja trae únicamente los del día, que es lo que se despacha y lo que se factura.

> Consecuencia a tener presente: un pedido **creado ayer** que se despache hoy ya **no** aparece en ITEMS DEPOT. Si
> alguna vez hace falta arrastrar los de ayer, se vuelve a poner `DATEADD(DAY, -1, CAST(GETDATE() AS DATE))` en la
> primera línea del filtro y se quita la segunda.

**Las fórmulas que calculan el box density no se tocan.** Siguen agrupando ITEMS DEPOT por contenedora, que es
justamente lo que vuelve a funcionar bien al devolver cada ítem con su caja:

| Dónde | Qué calcula | Estado |
|---|---|---|
| `EMPAQUETADO` H (PESO CAJA), I (VOL. CAJA) | peso y volumen del tipo de caja, desde DATA_CAJAS | ahora dicen `Verificar` si el tipo de caja no está en DATA_CAJAS |
| `EMPAQUETADO` J (VOL. ÍTEMS) | volumen de los ítems **de esa contenedora**, desde TABLAS DINAMICAS | vuelve a ser correcto |
| `EMPAQUETADO` K (PORCENTAJE) | VOL. ÍTEMS ÷ VOL. CAJA = **% de ocupación** | vuelve a ser correcto |
| `ITEMS DEPOT` PESO, VOLUMEN, PRECIO | unitario × cantidad de la fila | correcto: la cantidad ya no se repite |

### Las columnas F:K de EMPAQUETADO se vuelven a escribir en cada actualización

**Síntoma:** después de actualizar, PESO CAJA y VOL. CAJA salían **vacías**, PORCENTAJE daba **#¡DIV/0!** y más
abajo, fuera de la tabla, quedaban colgados los valores del día anterior.

**Causa:** la consulta devuelve muchas menos filas que antes (de ~145 cajas a 6). Al encogerse la tabla, Excel
conserva las fórmulas solo de las columnas que son *columnas calculadas* de verdad; las que en algún momento
quedaron como fórmulas sueltas se pierden. Con `VOL. CAJA` vacía, `PORCENTAJE` divide entre cero.

**Solución:** `ActualizarTodo` vuelve a escribir **F:K** al terminar de refrescar, siempre, dentro de un
envoltorio que ante cualquier fallo solo anota en el registro. También se puede forzar desde **Reparar fórmulas**
o con Alt + F8 › `RestaurarColumnasEmpaquetado`.

Además las fórmulas ya no devuelven errores, sino el aviso que el operario entiende:

| Columna | Antes | Ahora |
|---|---|---|
| H (PESO CAJA) | `#N/A` si el tipo de caja no está en DATA_CAJAS | `Verificar` |
| I (VOL. CAJA) | `#N/A` | `Verificar` |
| K (PORCENTAJE) | `#¡DIV/0!` si VOL. CAJA está vacía o en 0 | `Verificar` |

> Un tipo de caja nuevo, como `CAJA ORIGINAL WF1400`, sale en `Verificar` hasta que se agregue en
> **DATA CODIGO Y CAJAS** (botón **Cajas** del panel). *Avance empaque* los lista en el registro.

### "Verificar" en EMPAQUETADO: las dos fuentes tienen que mirar el mismo día

`VOL. ITEMS` se calcula agrupando ITEMS DEPOT **por contenedora**. Si una contenedora está en EMPAQUETADO pero
sus ítems no están en ITEMS DEPOT, no hay volumen que sumar y sale **Verificar**. Eso pasa cuando las dos
consultas miran ventanas de fecha distintas: las filas de ayer en EMPAQUETADO contra un ITEMS DEPOT que ya solo
trae hoy.

Por eso **las dos consultas filtran ahora el día de hoy**:

| Consulta | Dónde está el filtro |
|---|---|
| `Estado` (ITEMS DEPOT) | `SQL_ITEMS_DEPOT.sql`: `FECHA_CREACION >= CAST(GETDATE() AS DATE)` y `< GETDATE()+1` |
| `EMPAQUETADO` | `EMPAQUETADO.m`, paso `Hoy`: `[FECHA] = Date.From(DateTime.LocalNow())` |

> Si algún día hace falta trabajar con dos días, hay que cambiar **las dos**, nunca una sola. Con una sola, las
> contenedoras del día que sobra se quedan sin ítems y vuelven los *Verificar*.

Las filas de ayer que quedaron en *Verificar* no se pueden recalcular hacia atrás: su volumen por contenedora ya
no está en ITEMS DEPOT. Lo que sí queda registrado del día de ayer es la hoja **TRAZABILIDAD**, si se generó
antes de cerrar el día; por eso conviene generarla al terminar cada jornada.

### Número de cajas (BULTOS) y box density

**Regla:** una contenedora = una caja, sin importar cuántos ítems lleve.

Como en **ITEMS DEPOT no se puede agregar una columna**, el total de cajas se resuelve en la hoja EMPAQUETADO, que ya
tiene **una fila por contenedora**:

| Columna | Qué hace ahora |
|---|---|
| `EMPAQUETADO!G` (BULTOS) | **1 por contenedora**. El total del pedido sale de **sumar**: lo hace la tabla dinámica *Suma de BULTOS*, y `DATOS!R` lo cuenta aparte |
| `DATOS!R` (BULTOS) | El mismo `COUNTIF` sobre EMPAQUETADO, con mínimo 1. De ahí pasa solo a **DESPACHOS!L**, **TRAMACO!W** y a la etiqueta |
| `EMPAQUETADO!J` (VOL. ITEMS) | Solo el volumen de los ítems **de esa contenedora** (dinámica por contenedora). Si esa contenedora no tiene ítems en ITEMS DEPOT, dice **Verificar** |
| `EMPAQUETADO!K` (PORCENTAJE) | Sin cambios: J ÷ I, con el tope de 100 % y el piso de 30 % de siempre |
| `EMPAQUETADO!H`, `I` | Sin cambios: peso y volumen del tipo de caja, desde DATA_CAJAS |

**BULTOS va 1 por contenedora, nunca el total del pedido.** Se probó poner el total en cada fila y **duplicaba los
bultos**: la dinámica *Suma de BULTOS* suma esa columna, así que un pedido de 2 cajas salía con **4**. La regla es:

- `EMPAQUETADO!G` = **1**, el bulto de esa contenedora;
- el **total del pedido** se obtiene sumando: lo hace la dinámica *Suma de BULTOS*, y `DATOS!R` lo cuenta aparte
  con `COUNTIF` sobre las contenedoras del pedido, sin depender de esa dinámica.

> Si alguna vez hay que mostrar el total del pedido en la hoja EMPAQUETADO, tiene que ir en una **columna nueva**,
> nunca en BULTOS: cualquier suma de BULTOS por pedido se multiplicaría por el número de cajas.

**Por qué `VOL. ITEMS` ya no se respalda con el volumen del pedido.** El respaldo anterior era
`SUMIF(por contenedora)` y, si fallaba, `SUMIF(por pedido)`. Cuando ITEMS DEPOT tenía todos los ítems colgados de la
primera contenedora, la segunda y la tercera caían al respaldo y se llevaban **el volumen de todo el pedido**: de ahí
salían los `100 %` falsos de la hoja, mezclados con los `Verificar`. Para facturar es peor un 100 % inventado que un
`Verificar`, así que el respaldo se quitó.

Con la consulta nueva, cada contenedora tiene sus propios ítems en ITEMS DEPOT, la dinámica por contenedora los ve y
los `Verificar` deberían desaparecer salvo en un caso legítimo: la contenedora todavía no tiene ítems empacados.

**El código lee la tabla `Estado` por nombre de encabezado**, no por posición, así que si algún día la consulta
cambia de columnas, el panel y el avance de empaque no se descuadran.

## 7. Trazabilidad de pedidos

Para presentar la trazabilidad completa del día hay una hoja **TRAZABILIDAD**, con tres bloques:

| Bloque | Qué contiene |
|---|---|
| 1. Estado final de cada pedido | Una fila por pedido con la hoja DATOS tal como quedó: pedido, destinatario, provincia/cantón/parroquia, parroquia TMS y código postal, dirección, teléfono, **destino final**, **origen del destino** (regla base por provincia / confirmado por regla / confirmado a mano / sin destino), **destino que daría la regla base**, sugerencia vigente sin aplicar y su motivo, motivo con **quién y cuándo**, courier, trayecto, zona peligrosa, cobertura TMS, diagnóstico, bultos, peso, etiqueta (estado, destino impreso y fecha), empaque, cajas y unidades confirmadas/solicitadas |
| 2. Cambios de destino confirmados | Solo los pedidos que se apartaron de la regla base: **destino sin confirmar → destino final**, con el motivo, el usuario y la fecha |
| 3. Registro de acciones de hoy | Las líneas de LOG_EGR del día: revisiones de cobertura, destinos, etiquetas impresas, exportes y errores |

**Cuándo se genera:**

- **Sola**, cada vez que se aplican destinos sugeridos (desde el panel o desde *2b Aplicar destinos sugeridos*). Si por
  lo que sea fallara, solo queda anotado en el registro: **nunca interrumpe** lo que se estaba haciendo.
- **A pedido**, con el botón **Trazabilidad de pedidos** del panel o con *Trazabilidad* en la pestaña Complementos.
  Ahí pregunta si además guarda una copia **.xlsx** en la carpeta de exportes
  (`EXPORTES\aaaa-mm-dd\TRAZABILIDAD_aaaammdd_hhmm.xlsx`), que es la que se adjunta o se imprime.

La hoja se rehace entera cada vez y es **solo lectura** sobre las hojas de trabajo: no cambia un dato, una fórmula ni
una conexión. Si la borras, se vuelve a crear sola la próxima vez.

> El "destino que daría la regla base" se calcula con la misma regla de la fórmula de `DATOS!A` (Pichincha = UIO,
> Guayas = GYE, Galápagos = GPS, el resto = PRO). Así el **antes → después** de cada cambio queda sin depender de que
> alguien se acuerde de anotarlo.

## 8. Datos de días anteriores y respaldo del día

Las consultas traen **solo el día de hoy**, así que al actualizar se pierde lo de ayer. Hay dos mecanismos para
conservarlo, y los dos funcionan solos:

### Hoja HISTORICO_EMPAQUE

Guarda las cajas de cada día: FECHA, # ORDEN, TIPO CAJA, CONTENEDORA, BULTOS, PESO CAJA, VOL. CAJA, VOL. ITEMS,
PORCENTAJE y cuándo se guardó.

- Se guarda **sola** en tres momentos: al **empezar** a actualizar (lo que hay todavía es el día anterior, y el
  refresco lo iba a borrar), al **terminar** de actualizar (ya con el día de hoy y sus fórmulas) y al **exportar**
  los reportes.
- Una contenedora se guarda **una sola vez por día**: si se vuelve a actualizar, esa fila se refresca con los datos
  nuevos en lugar de duplicarse.
- Se abre con **Datos anteriores** (panel o pestaña Complementos): deja la hoja con filtro puesto y dice cuántas
  cajas y cuántos días hay. Se filtra por FECHA para revisar o corregir un día pasado.

### Respaldo del día en binario (.xlsb)

- `RESPALDOS_EGR\EGR_DIA_aaaa-mm-dd.xlsb`: el archivo **completo** del día, en formato binario (pesa bastante
  menos que el .xlsm y abre más rápido). **Un archivo por día**: durante la jornada se va reescribiendo, así que
  siempre refleja el estado final.
- Se genera **solo** al terminar de exportar los reportes, que es el cierre del proceso del día. Se puede forzar
  con **Respaldo del día** (panel o Complementos).
- Se apaga poniendo `NO` en `RESPALDO_DIARIO` (hoja oculta CONFIG_EGR).
- Para consultar un día pasado completo —pedidos, destinos, etiquetas, empaque— se abre ese .xlsb. Para solo el
  empaque, basta la hoja HISTORICO_EMPAQUE.

> No se confunda con los respaldos `EGR_aaaammdd_hhmmss_motivo.xlsm`, que son instantáneas antes de un proceso
> delicado (actualizar, reparar, limpiar día) y de las que se conservan las 15 últimas.

### Corregir los "validar" de EMPAQUETADO, sin actualizar nada

`validar` en **EMPAQUETADO!F** (NRO. DE CONTENEDORA) significa: *el pedido existe hoy, pero no se le encontró
contenedora*. `sin pedido` significa que ni siquiera está en los pedidos del día.

**De dónde salen.** La fórmula de F usa la columna E (# CONTENEDORA, la que llena bodega en el Google Sheets) y,
si está vacía, busca la contenedora del pedido en `Estado` con `XLOOKUP`. Eso falla en dos casos:

- el pedido **no está en `Estado`** (p. ej. es de otro día, y `Estado` solo trae hoy) → `validar`;
- el pedido **tiene varias cajas**: `XLOOKUP` devuelve siempre **la primera**, así que todas las filas del pedido
  recibían la misma contenedora y el volumen se iba entero a una caja.

**La corrección** (menú **Corregir contenedoras (validar)**) trabaja **solo con lo que ya está en el libro**: la
tabla `Estado` de ITEMS DEPOT y la propia hoja EMPAQUETADO. **No actualiza ninguna consulta, no abre ninguna
conexión y no trae datos nuevos.** Lo único que recalcula es la hoja, en local.

A cada fila sin contenedora le asigna **una contenedora distinta del mismo pedido** y la escribe en la **columna
E**, que la fórmula respeta por encima de la búsqueda (`IF($E2<>"",$E2,…)`). Ventajas: no se cambia ninguna
fórmula, y **para deshacer basta borrar lo escrito en E**.

Antes de aplicar muestra el reparto y lo clasifica:

| Caso | Qué significa |
|---|---|
| **Exactas** | el pedido tiene una sola caja y una sola contenedora: no hay ninguna duda |
| **Repartidas por orden** | el pedido tiene varias cajas; se asignan las contenedoras en orden ascendente |
| **Sin contenedora en Estado** | no hay de dónde sacarla; la fila se queda como está y se lista en el registro |

> **Lo que hay que mirar antes de facturar:** cuando un pedido tiene varias cajas, **ningún dato del libro dice qué
> contenedora corresponde a qué tipo de caja**. El reparto por orden es una suposición razonable, no un hecho. El
> total del pedido (peso, volumen, costo) sale bien igual, porque son las mismas contenedoras; lo que puede quedar
> cruzado es el % de ocupación de cada caja por separado. Esas filas salen listadas en el aviso y en el registro.

Antes de escribir deja un respaldo en `RESPALDOS_EGR`. Funciona también en un archivo ya congelado: si F ya no es
fórmula, escribe el valor directamente.

### Congelar un archivo como histórico

Un histórico es una **foto del día**: debe poder consultarse sin riesgo de que se actualice ni de que alguien lo
cambie sin querer. Un simple "Guardar como" **no** sirve: la copia se lleva las 5 conexiones vivas (EMPAQUETADO,
Estado, ITEMS API y las dos del modelo de datos) y todas las fórmulas, así que basta abrirla y pulsar algo para que
se contamine con los datos de otro día.

Congelar hace cinco cosas:

| Paso | Qué hace | Para qué |
|---|---|---|
| 1 | Fórmulas → **valores**, en todas las hojas | nada se recalcula al abrirlo |
| 2 | Tablas de consulta → **desconectadas** del origen (`Unlink`), conservando los datos | la tabla sigue ahí, pero ya no apunta a DEPOT ni al Sheets |
| 3 | **Se eliminan las conexiones** (ODBC DEPOTUIO, Google Sheets, modelo de datos) | no hay nada que actualizar |
| 4 | Tablas dinámicas: guardan sus datos y **no se refrescan al abrir** | los números siguen a la vista |
| 5 | Marca `ARCHIVO_HISTORICO = SI` y `FECHA_HISTORICO` en CONFIG_EGR | el propio código se bloquea |

Con esa marca, el archivo avisa en la pestaña Complementos (**ARCHIVO HISTÓRICO dd/mm/aaaa – solo consulta**) y
quedan **bloqueados**: Actualizar datos, Reparar fórmulas, Limpiar día, los cambios de destino y el respaldo del
día. Lo que sí funciona: mirar, filtrar, ordenar, la trazabilidad y los exportes.

**Hay dos caminos:**

1. **El respaldo diario sale ya congelado.** `EGR_DIA_aaaa-mm-dd.xlsb` se genera al exportar los reportes y pasa por
   los cinco pasos antes de guardarse. No hay que hacer nada más: ese archivo **es** el histórico del día.
2. **Congelar una copia a mano** (p. ej. un `BACKUP.xlsm` que ya tengas): se abre **esa copia**, pestaña
   **Complementos › Congelar como histórico**, se escribe la fecha del día y se confirma escribiendo `HISTORICO`.
   Al terminar, **guardar** (mejor como `.xlsb`, pesa bastante menos).

> **Protección:** la macro se niega a correr si el nombre del archivo contiene *Formato EGR*, para que nunca se
> congele el de producción por error. Aun así, **ejecútala siempre sobre una copia**: no se puede deshacer.

## 9. Exportación

- Se elige la **hoja** (TRAMACO, TMS, DESPACHOS o LAS TRES) y el **formato** (CSV, XLSX o PDF) en listas desplegables.
- Se exportan **solo las filas con datos de esa hoja**, con **las mismas columnas y en el mismo orden**, porque son
  plantillas de carga. Las hojas no se modifican.
  - **TRAMACO** lleva solo los pedidos PRO: toma las filas con nombre (C) y pedido (AD). En los datos actuales son 70.
    En el archivo que sale, la columna auxiliar **FILA_DATOS** (el número de fila, que al courier no le sirve) se
    reemplaza por **ZONA**: la zona peligrosa que viene de PEDIDOS HCE (DATOS!AM). Mismo número de columnas, mismo
    orden, y **la hoja TRAMACO no se modifica**: el cambio es solo en el archivo exportado.
  - **TMS** lleva **todos** los pedidos del día (152), porque así está armada la hoja TMS: una fila por cada pedido de
    DATOS. En el archivo que sale se ponen los **encabezados oficiales de la interfaz de TMS**
    (`EMPRESA;FECHA_INTERFAZ;…`), porque la hoja los tiene abreviados (`CP_RTTE` en vez de `CODIGO_POSTAL_RTTE`,
    `DNI(DESTINATARIO)` en vez de `DNI_DEST`, etc.). Las **42 columnas de la hoja están en el mismo orden que la
    plantilla**: solo cambia el nombre. Se agregan además las 4 columnas finales de la plantilla
    (`NRO_TRACKING_EXPRESO`, `NRO_TRACKING_REPRESENTANTE`, `DEVOLUCION`, `NRO_DEVOLUCION`), que van vacías, para
    completar las 46. **La hoja TMS no se modifica.**
    Si el archivo oficial cambiara de columnas, la lista está en `CabecerasTMS` (modEGR) y si la hoja llegara a
    tener más columnas que la plantilla, el código no toca nada y lo avisa en el registro.

    **El CSV de TMS se escribe a mano, no con "Guardar como".** `SaveAs` deja que Excel decida el separador, el
    formato de fecha y el de número según la **configuración regional de la PC** y el formato de cada celda. Por eso
    el archivo que exportaba el sistema no se podía subir. Comparado con el que TMS sí aceptó:

    | Campo | Salía | Debe ser |
    |---|---|---|
    | Separador | `,` — y además entrecomillaba las direcciones que llevan coma | `;` |
    | FECHA_INTERFAZ / FECHA_COMPRA | `10/6/2026` (mm/dd) | `6/10/2026` (dd/mm) |
    | VALOR_TOTAL_FACTURA | `$4.70 ` (con símbolo y espacio) | `4.7` |
    | TELEFONO_MOVIL / FIJO | `0968032761` | `968032761` |

    Ahora el archivo lo escribe el código: `;` como separador, **sin comillas**, fechas `dd/mm/aaaa`, números con
    punto decimal y sin símbolo, teléfonos solo dígitos y sin el 0 inicial, UTF-8 con BOM. Sale **igual en
    cualquier PC**, sin depender de la configuración regional ni de `CSV_SEPARADOR_PUNTOYCOMA`.
    Dentro de un campo nunca viaja un `;`, una comilla ni un salto de línea: se reemplazan por un espacio.

    > Las fechas solo se convierten si la celda tiene una **fecha de verdad**. Si TMS!B o TMS!Z tuvieran texto, no
    > se adivina si `10/6` es 10 de junio o 6 de octubre: se deja tal cual y queda avisado en el registro.

    **Esto solo afecta a TMS en CSV.** TRAMACO y DESPACHOS se exportan exactamente como antes.
  - **DESPACHOS** lleva todos los pedidos (152).

  Antes de exportar, el panel muestra cuántas filas tiene cada hoja.
- Las columnas de texto conservan el 0 inicial (teléfonos, códigos).
- Antes de exportar se revisan errores (`#N/A`), celdas con VALIDAR o Verificar y campos obligatorios vacíos. Todo
  queda en el registro y se pregunta si se exporta igual.
- **CSV:** UTF-8, separado por coma. Si el sistema que lo recibe pide punto y coma, pon `SI` en
  `CSV_SEPARADOR_PUNTOYCOMA` (hoja oculta CONFIG_EGR).
- **Carpeta:** `EXPORTES\aaaa-mm-dd`, junto al archivo. Otra ruta se configura en `CARPETA_EXPORTES`.
- **Correo:** destinatarios en `CORREO_PARA`. Se crea un **borrador** en Outlook; nunca se envía solo.

## 10. Qué verificar en la prueba

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
5. **IMPRIMIR ETIQUETAS**: la ventana abre con la lista llena. Filtra por destino GYE, comprueba que el botón dice
   cuántas van e imprime. Vuelve a abrirla: la impresora ya aparece elegida.
5b. **Clic en un encabezado de la lista** del panel: ordena por esa columna y aparece la flecha; otro clic invierte.
5c. **Trazabilidad de pedidos**: se arma la hoja con los tres bloques y, si aceptas, deja el .xlsx en EXPORTES.
6. **Exportar** TRAMACO en CSV: el archivo debe tener las mismas filas que la hoja TRAMACO.
