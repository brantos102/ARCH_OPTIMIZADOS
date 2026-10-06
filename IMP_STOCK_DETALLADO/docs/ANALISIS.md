# Análisis del VBA original (`VistaPreviaEImpresion`)

Archivo analizado: `IMP_STOCK_DETALLADO.xlsm` (17 MB, macros en `vbaProject.bin`).

## 1. Estructura real del libro

| Hoja | Estado | Contenido |
|---|---|---|
| `IMPRESION VACIA` | oculta | formato de planilla en blanco |
| `PLANILLA_IMPRESION` | oculta | planilla de ingreso/traslado |
| `B_COD` | visible | base de códigos (autofiltro `D2:AP800`) |
| `B_POS` | visible | base de posiciones (autofiltro `D2:AI71`) |
| `ETQ` | visible | **plantilla de la etiqueta**, área de impresión `A1:J17` |
| `IMPRIMIR` | oculta | hoja de trabajo con la misma información, otro orden de columnas |
| `Consolidado` | visible | **origen real**: tabla `Consolidado` `A1:AF130181` (130.180 filas) |

Datos medidos en el archivo entregado:

- La tabla `Consolidado` tiene **autofiltro activo** sobre la columna 2 (`cliente_id = "VENTURA"`):
  **120.638 filas ocultas** y **9.542 filas visibles**.
- `ETQ`: papel personalizado **100 x 50 mm**, 203 ppp, orientación vertical, márgenes 0,
  blanco y negro, ajustado a 1 página. Driver guardado en el libro:
  **`ZDesigner ZD421-203dpi ZPL`**.
- Las celdas `A6` y `A15` usan la fuente **`Free 3 of 9 Extended`, 80 pt** (Code 39):
  por eso el valor se escribe entre asteriscos (`*CODIGO*`).

## 2. Defectos encontrados en la macro original

| # | Defecto | Consecuencia |
|---|---|---|
| 1 | `For Each celda In Selection.Rows` recorre **todas** las filas del rango, incluidas las ocultas por el autofiltro | Al seleccionar un bloque filtrado se imprimen etiquetas de filas que el filtro había descartado. **Este es el problema central del pedido.** |
| 2 | La primera etiqueta sólo se manda a `PrintPreview`; al responder "Sí" el bucle sigue con la **segunda** fila | **La primera etiqueta nunca se imprime** |
| 3 | Un `wsETQ.PrintOut` por fila | 1 trabajo de impresión por etiqueta. Con 100-200 etiquetas son 100-200 trabajos: lento, y la cola del driver ZPL puede mezclar o descartar trabajos |
| 4 | `Dim filasProcesadas As Integer` | Desborda (error 6) por encima de 32.767 |
| 5 | `If IsError(valor) Or valor = ""` en `ValidarDato` | VBA evalúa **los dos** lados de `Or`: si la celda tiene `#N/A` o `#REF!`, la comparación lanza error 13 (`Type mismatch`) y la macro muere |
| 6 | `Select Case Selection.Worksheet.Name` se ejecuta **antes** de validar `TypeName(Selection) = "Range"` | Si hay un gráfico o un objeto seleccionado: error 438 |
| 7 | El mensaje habla de la hoja `STOCK DETALLADO UIO`, que ya no existe | Confunde al operador |
| 8 | `wsETQ.Activate` dentro del bucle, sin `ScreenUpdating = False` | Repintado y repaginación por cada etiqueta |
| 9 | El mapeo de columnas es **posicional**, pero las dos hojas de origen no tienen el mismo orden: en `Consolidado` la col. 7 es `nro_despacho` y la 8 `nro_partida`; en `IMPRIMIR` están **invertidas** | Imprimiendo desde `IMPRIMIR`, el recuadro `RESPONSABLE` y el de `H5` salen cruzados |
| 10 | No hay forma de pedir más de una copia por fila | Había que repetir la selección a mano |
| 11 | El código de barras se arma con el valor crudo | Code 39 (`Free 3 of 9 Extended`) no codifica minúsculas ni símbolos: el código sale impreso pero **no se lee con el lector** |
| 12 | Se escribe `.Value` directo | Una descripción con salto de línea, una fecha o un número en notación científica se renderizan mal en la etiqueta |
| 13 | No se restauran hoja activa, selección ni estado de Excel | El operador queda en `ETQ` y pierde su selección |

## 3. Decisiones de la nueva implementación

1. **Un solo trabajo de impresión por lote.** Se crea una hoja temporal `ETQ_LOTE` que es una
   copia de `ETQ` (hereda tamaño de papel, márgenes, fuentes, combinaciones y escala), se
   replica el bloque de 17 filas una vez por etiqueta, se coloca un salto de página al inicio
   de cada bloque y se imprime todo junto. Por defecto se agrupa en trabajos de
   **100 etiquetas** (`PAGINAS_POR_TRABAJO`), que es el tamaño que mejor tolera la cola del
   driver ZPL de la ZD421.
2. **Réplica por duplicación (1, 2, 4, 8, ...)**: armar 100 etiquetas cuesta 7 operaciones de
   copiado, no 100.
3. **Filtro respetado siempre**: las filas se obtienen con `SpecialCells(xlCellTypeVisible)`
   recorrido por bloques de 20.000 filas, para no agotar `SpecialCells` cuando hay más de
   120.000 filas ocultas (que es el caso real de este archivo).
4. **Mapeo mixto**: los campos se buscan por **nombre de encabezado** (`producto_id`,
   `descripcion`, `nro_serie`, `nro_partida`, ...) y, si no aparece el encabezado, se usa la
   columna fija de la macro original. Así `Consolidado` imprime exactamente igual que antes y
   `IMPRIMIR` deja de cruzar los dos campos del defecto 9.
   La única excepción es `A2` (cliente/marca), que se mantiene en la **columna 2 fija** porque
   en `Consolidado` es el nombre del cliente y en `IMPRIMIR` es el campo `ABC`: en las dos
   hojas la columna 2 es el texto que debe ir en la etiqueta.
5. **Verificación de paginación**: antes de imprimir se compara `PageSetup.Pages.Count` con el
   número de etiquetas del lote. Si no coinciden se reintenta con escala fija del 28 % y, si
   aún así no cuadra, se avisa en pantalla en vez de gastar etiquetas.
6. **La plantilla `ETQ` no se modifica** durante el proceso (sólo, al final y de forma
   opcional, se le carga la primera etiqueta del lote para mantener la costumbre anterior).
   La hoja `ETQ_LOTE` se elimina al terminar, incluso si hay un error.
