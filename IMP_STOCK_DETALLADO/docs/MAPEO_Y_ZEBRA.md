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

El módulo pasa el valor a **mayúsculas** y descarta lo que la fuente no sabe dibujar, para
que no queden barras basura en medio del código. El valor completo y sin tocar se imprime
igual en texto legible debajo (`A9` / `A17`) y queda registrado en `ETQ_LOG`.

### Ancho de la barra fina

Es lo que decide si el lector puede leer el código. En Code 39 cada carácter ocupa 16
módulos (15 + separación), y el área `A6:G8` mide **63,5 mm** impresos, o sea 508 puntos a
203 ppp:

| Caracteres | Módulos | Barra fina | Lectura |
|---|---|---|---|
| 9 | 176 | 2,9 puntos · 0,36 mm | holgada |
| 12 | 224 | 2,3 puntos · 0,28 mm | holgada |
| 14 | 256 | 2,0 puntos · 0,25 mm | correcta |
| 18 | 320 | 1,6 puntos · 0,20 mm | justa |
| 29 | 496 | 1,0 punto · 0,13 mm | al límite del cabezal |

Con el "Reducir hasta ajustar" de la celda, Excel encogía la fuente sin mirar nada de esto
y las barras terminaban pegadas unas a otras. Ahora el módulo **calcula el tamaño de fuente
de cada código**: lo agranda hasta ocupar todo el ancho disponible (sin pasarse del tamaño
original de la plantilla, para que los códigos cortos salgan igual que siempre) y lo ajusta
para que la barra fina caiga en un número **entero** de puntos de impresora, que es lo que
permite al lector distinguir barra fina de barra gruesa.

El ancho de carácter de la fuente se mide en tiempo de ejecución, así que el cálculo sigue
siendo válido si se cambia la fuente o el diseño de la etiqueta. `EtiquetasDiagnostico`
informa el ancho útil y hasta cuántos caracteres entran con barra de 2 puntos.

### Códigos largos: Code 128 dibujado

Para códigos muy largos el límite de Code 39 es físico: 29 caracteres necesitan 124 mm con
barra de 0,25 mm, más que la etiqueta entera. Por eso, cuando la barra fina quedaría por
debajo de **2 puntos de impresora**, el módulo deja de usar la fuente y **dibuja** el código
en **Code 128**:

| | Code 39 | Code 128 |
|---|---|---|
| Módulos por carácter | 16 | ~11 (2 dígitos por símbolo en el subconjunto C) |
| `VWMKT-201902-SCC-ATRIL TABLET` | 496 módulos | **343 módulos** (-31 %) |
| Barra fina en `A6:G8` | 1,02 puntos | 1,40 puntos |
| Barra fina a ancho completo | — | **2,02 puntos** |

Tres cosas hacen que esto sí se lea:

1. **Code 128 es un 30 % más corto** con exactamente el mismo dato.
2. Las barras se dibujan como rectángulos, cada una de un número **entero** de puntos de
   impresora: el cabezal no redondea nada y la fina se distingue de la gruesa.
3. Si aún hace falta ancho y el recuadro de la derecha de la etiqueta está vacío, el código
   se extiende hasta el borde (`EXTENDER_BARRAS`).

Lo que cambia es la simbología, no el contenido: el lector devuelve el mismo texto que está
impreso debajo del código. Cualquier lector de los últimos 20 años lee Code 128; si el suyo
está configurado para aceptar sólo Code 39, hay que habilitarlo (o poner
`BARRAS_DIBUJADAS = 0` para volver a Code 39 siempre).

El codificador de Code 128 se verificó decodificando **20.026 cadenas de ida y vuelta**,
sin un solo fallo, y su tabla de patrones coincide byte a byte con la de referencia.

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

## 5. Cómo se calcula la escala de impresión  (10 x 5 cm en cualquier impresora)

La hoja `ETQ` imprime con "ajustar a 1 página": Excel calcula solo el porcentaje con el que
las columnas `A:J` y las 17 filas entran en la etiqueta de 100 x 50 mm. Ese ajuste **no se
puede trasladar tal cual** a una hoja con muchas etiquetas, porque se aplicaría al lote
completo en vez de a cada etiqueta.

Y además la etiqueta mide siempre 10 x 5 cm, use la Zebra u otra impresora. Por eso la
escala se **calcula** para que el bloque `A1:J17` ocupe exactamente ese tamaño:

```
escala = el menor de    ancho etiqueta / ancho del bloque
                        alto  etiqueta / alto  del bloque
```

`Range.Width` y `Range.Height` devuelven puntos, así que el cálculo es exacto y no depende
del driver. Con la plantilla actual:

| | |
|---|---|
| Bloque `A1:J17` | 904,5 x 524,2 pt |
| Etiqueta 10 x 5 cm | 283,5 x 141,7 pt |
| Escala | **27 %** (limita el alto: 524,2 x 0,27 = 49,9 mm) |

Aparte se **mide el área imprimible real**: al 100 % se mira dónde corta Excel la página y
se suman los anchos de columna y los altos de fila que entraron. Si el papel configurado en
el driver es más chico que la etiqueta, el módulo avisa antes de imprimir en vez de sacar
etiquetas cortadas. El dato aparece en `EtiquetasDiagnostico`.

Con `ESCALA_FIJA` se puede forzar un porcentaje concreto y saltarse el cálculo.
