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
4. pone un tope de seguridad de 2 000 cajas.

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

- Es **una etiqueta por pedido**, con el formato de la muestra impresa:
  - código de barras Code 128 arriba a la izquierda, con el **número centrado debajo**, sin espacios;
  - **destino en grande** a la derecha (UIO, GYE, PRO o GPS) y la **parroquia** debajo;
  - **nombre completo del destinatario** abajo a la izquierda, en una sola línea.
- Hay **un solo botón**: **IMPRIMIR ETIQUETAS** (en el panel, en la pestaña Complementos y con Alt + F8 ›
  `MenuEtiquetas`). No hay que seleccionar nada antes: la ventana abre con **todos** los pedidos con destino y se
  imprime **exactamente lo que quede en la lista**.
- Dentro de la ventana:
  - **filtro por estado**: pendientes + por reimprimir (así abre), pendientes, por reimprimir, impresas o todas;
  - **filtro por destino**: PRO, GYE, UIO, GPS o todos;
  - **búsqueda** por pedido, destinatario, parroquia o destino, mientras se escribe;
  - **Quitar filtros** vuelve a la lista completa;
  - vista previa de la etiqueta del pedido en el que se hace clic;
  - el botón dice cuántas van: *Imprimir las 23 etiqueta(s) de la lista*, y antes de imprimir se confirma con el
    detalle de los filtros aplicados.
- Ya no hay casillas que marcar ni "imprimir seleccionadas": era el paso donde la ventana abría sin nada marcado.
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

**Las fórmulas que calculan el box density no se tocan.** Siguen agrupando ITEMS DEPOT por contenedora, que es
justamente lo que vuelve a funcionar bien al devolver cada ítem con su caja:

| Dónde | Qué calcula | Estado |
|---|---|---|
| `EMPAQUETADO` H (PESO CAJA), I (VOL. CAJA) | peso y volumen del tipo de caja, desde DATA_CAJAS | sin cambios |
| `EMPAQUETADO` J (VOL. ÍTEMS) | volumen de los ítems **de esa contenedora**, desde TABLAS DINAMICAS | vuelve a ser correcto |
| `EMPAQUETADO` K (PORCENTAJE) | VOL. ÍTEMS ÷ VOL. CAJA = **% de ocupación** | vuelve a ser correcto |
| `ITEMS DEPOT` PESO, VOLUMEN, PRECIO | unitario × cantidad de la fila | correcto: la cantidad ya no se repite |

### Número de cajas (BULTOS) y box density

**Regla:** una contenedora = una caja, sin importar cuántos ítems lleve.

Como en **ITEMS DEPOT no se puede agregar una columna**, el total de cajas se resuelve en la hoja EMPAQUETADO, que ya
tiene **una fila por contenedora**:

| Columna | Qué hace ahora |
|---|---|
| `EMPAQUETADO!G` (BULTOS) | `COUNTIF` de las filas con el mismo **# ORDEN** = **total de contenedoras del pedido**. Antes era `1` fijo, que era el bulto de esa fila, no el del pedido |
| `DATOS!R` (BULTOS) | El mismo `COUNTIF` sobre EMPAQUETADO, con mínimo 1. De ahí pasa solo a **DESPACHOS!L**, **TRAMACO!W** y a la etiqueta |
| `EMPAQUETADO!J` (VOL. ITEMS) | Solo el volumen de los ítems **de esa contenedora** (dinámica por contenedora). Si esa contenedora no tiene ítems en ITEMS DEPOT, dice **Verificar** |
| `EMPAQUETADO!K` (PORCENTAJE) | Sin cambios: J ÷ I, con el tope de 100 % y el piso de 30 % de siempre |
| `EMPAQUETADO!H`, `I` | Sin cambios: peso y volumen del tipo de caja, desde DATA_CAJAS |

**Por qué `DATOS!R` dejó de usar la tabla dinámica.** Antes hacía `VLOOKUP` a la dinámica *Suma de BULTOS*
(`TABLAS DINAMICAS` H:K). Con `G = 1` esa suma daba el número de cajas, pero ahora que `G` trae el total del pedido
en cada una de sus filas, la suma daría el total **al cuadrado** (3 cajas → 9). Por eso `DATOS!R` cuenta las
contenedoras directamente y ya no depende de esa dinámica.

> La dinámica *Suma de BULTOS* queda solo informativa y mostrará ese número al cuadrado. Si molesta, se arregla en
> dos clics: clic derecho sobre ella › **Configuración de campo de valor** › **Cuenta**. Ninguna fórmula la usa.

**Por qué `VOL. ITEMS` ya no se respalda con el volumen del pedido.** El respaldo anterior era
`SUMIF(por contenedora)` y, si fallaba, `SUMIF(por pedido)`. Cuando ITEMS DEPOT tenía todos los ítems colgados de la
primera contenedora, la segunda y la tercera caían al respaldo y se llevaban **el volumen de todo el pedido**: de ahí
salían los `100 %` falsos de la hoja, mezclados con los `Verificar`. Para facturar es peor un 100 % inventado que un
`Verificar`, así que el respaldo se quitó.

Con la consulta nueva, cada contenedora tiene sus propios ítems en ITEMS DEPOT, la dinámica por contenedora los ve y
los `Verificar` deberían desaparecer salvo en un caso legítimo: la contenedora todavía no tiene ítems empacados.

**El código lee la tabla `Estado` por nombre de encabezado**, no por posición, así que si algún día la consulta
cambia de columnas, el panel y el avance de empaque no se descuadran.

## 7. Exportación

- Se elige la **hoja** (TRAMACO, TMS, DESPACHOS o LAS TRES) y el **formato** (CSV, XLSX o PDF) en listas desplegables.
- Se exportan **solo las filas con datos de esa hoja**, con **las mismas columnas y en el mismo orden**, porque son
  plantillas de carga. Las hojas no se modifican.
  - **TRAMACO** lleva solo los pedidos PRO: toma las filas con nombre (C) y pedido (AD). En los datos actuales son 70.
    En el archivo que sale, la columna auxiliar **FILA_DATOS** (el número de fila, que al courier no le sirve) se
    reemplaza por **ZONA**: la zona peligrosa que viene de PEDIDOS HCE (DATOS!AM). Mismo número de columnas, mismo
    orden, y **la hoja TRAMACO no se modifica**: el cambio es solo en el archivo exportado.
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
