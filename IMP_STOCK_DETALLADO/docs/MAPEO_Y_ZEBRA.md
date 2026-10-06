# Mapeo de la etiqueta y configuración de la Zebra ZD421

## 1. La plantilla `ETQ`

Área de impresión `A1:J17` = **una etiqueta**. Celdas combinadas y su contenido:

```
 A2:F4   cliente / marca                       H3:J4   categoría lógica (DISPONIBLE...)
 A5:B5   "CODIGO:"                             H5:J7   nro_despacho
 A6:G8   *CODIGO*  <- código de barras (Code 39, fuente Free 3 of 9 Extended 80 pt)
 A9:G9   código legible                        H8      "QTY"      J8   "ESTADO"
 A10:D10 "DESCRIPCION:"                        H10:H11 cantidad   I10:I11 unidad   J10:J11 estado
 A11:G13 descripción                           H13:J13 "RESPONSABLE"
 A14:B14 "SERIE:"                              H14:J14 nro_partida  <- responsable
 A15:G16 *SERIE*   <- código de barras
 A17:G17 serie legible
```

## 2. Mapeo celda ← columna de origen

| Celda `ETQ` | Encabezado buscado | Columna fija (respaldo) | Notas |
|---|---|---|---|
| `A2`  | — (posición fija) | **2** (`B`) | `Consolidado`: `cliente_id`; `IMPRIMIR`: `ABC` |
| `H3`  | `categoria_logica` | 9 (`I`) | |
| `H5`  | `nro_despacho` | 7 (`G`) | |
| `A6`  | `producto_id` | 3 (`C`) | **código de barras** `*VALOR*`, en mayúsculas |
| `A9`  | `producto_id` | 3 (`C`) | texto legible |
| `H10` | `cantidad` | 13 (`M`) | QTY, se conserva como número |
| `I10` | `unidad_medida` | 12 (`L`) | |
| `J10` | `estado_mercaderia` | 10 (`J`) | |
| `A11` | `descripcion` | 4 (`D`) | se eliminan saltos de línea |
| `H14` | `nro_partida` | 8 (`H`) | recuadro RESPONSABLE |
| `A15` | `nro_serie` | 5 (`E`) | **código de barras** `*VALOR*` |
| `A17` | `nro_serie` | 5 (`E`) | texto legible |

El mapeo es idéntico al de la macro original cuando se imprime desde `Consolidado`.
Desde `IMPRIMIR`, la búsqueda por encabezado corrige el cruce entre `nro_partida` y
`nro_despacho` que tenía la versión anterior.

Para ver el mapeo que se está aplicando en la hoja activa: macro
**`EtiquetasDiagnostico`**.

## 3. Código de barras (Code 39)

La fuente `Free 3 of 9 Extended` sólo codifica:

```
0-9   A-Z   -   .   espacio   $   /   +   %
```

Por eso el módulo pasa el valor a **mayúsculas** antes de envolverlo en asteriscos.
Si una fila trae otros caracteres (minúsculas acentuadas, `#`, `(`, `_`, `,`, ...), la
etiqueta se imprime igual pero el lector **no** va a poder leer ese código: al terminar el
lote aparece un aviso con el número de fila y los caracteres problemáticos.

También se avisa cuando el código supera los **20 caracteres**: en `A6:G8` el texto está con
"Reducir hasta ajustar", así que un código como `VWMKT-201902-SCC-ATRIL TABLET` (29
caracteres) se comprime tanto que las barras quedan por debajo de lo que lee un escáner a
203 ppp. En esos casos conviene imprimir la etiqueta sin barras o pasar ese campo a Code 128.

## 4. Impresora

El libro ya trae guardada la configuración de `ZDesigner ZD421-203dpi ZPL`:

| Parámetro | Valor |
|---|---|
| Papel | definido por el usuario, **100 x 50 mm** (3,94 x 1,97 pulg.) |
| Resolución | 203 x 203 ppp |
| Orientación | vertical |
| Márgenes | 0 en los cuatro lados |
| Color | blanco y negro |
| Escala | ajustar a 1 página de ancho |

El módulo comprueba que la impresora activa contenga el texto `ZD421`
(constante `IMPRESORA_CONTIENE`). Si no, intenta seleccionarla automáticamente (consulta
`Win32_Printer`) y, si tampoco la encuentra, pregunta antes de imprimir en otra impresora.
Al terminar deja la impresora activa como estaba.

### Recomendaciones del lado de la impresora

1. **Calibrar el medio** antes de un lote grande (botón de alimentación o utilidad de Zebra),
   para que el sensor de espacio detecte bien la etiqueta de 50 mm.
2. En el driver, `Propiedades de impresora > Preferencias`: tamaño **100 x 50 mm**,
   `Darkness` entre 15 y 20 para barras nítidas en papel térmico, velocidad 4 ips o menor.
3. Desactivar `Spool`/`Imprimir directamente en la impresora` sólo si la cola se traba:
   lo normal con lotes de 100 páginas es dejar el spooler activo.
4. Si al imprimir aparecen etiquetas en blanco intercaladas, es paginación: use la vista
   previa del paso 4 y revise con `EtiquetasDiagnostico` la escala medida.

## 5. Cómo se calcula la escala de impresión

La hoja `ETQ` imprime con "ajustar a 1 página": Excel calcula solo el porcentaje con el que
las columnas `A:J` y las 17 filas entran en la etiqueta de 100 x 50 mm. Ese ajuste **no se
puede trasladar tal cual** a una hoja con muchas etiquetas, porque se aplicaría al lote
completo en vez de a cada etiqueta.

Por eso el módulo **mide** la escala antes de imprimir: prueba porcentajes por búsqueda
binaria (10 % a 100 %) y, en cada uno, consulta a Excel dónde corta la página con la
impresora real. Se queda con el mayor porcentaje que cumple las dos condiciones de una
etiqueta:

- las columnas `A:J` entran a lo ancho (`VPageBreaks.Count = 0`);
- las 17 filas entran a lo alto (el primer corte horizontal cae en la fila 18 o más abajo).

El porcentaje elegido se informa al terminar el lote y en `EtiquetasDiagnostico`. Si la
medición no fuera posible en el equipo, el módulo avisa y pasa a **modo compatible**:
imprime etiqueta por etiqueta con el mismo "ajustar a 1 página" de `ETQ`, lo que da el
resultado correcto aunque más lento. Con `ESCALA_FIJA` se puede forzar un porcentaje
concreto y saltarse la medición.
