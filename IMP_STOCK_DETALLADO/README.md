# IMP_STOCK_DETALLADO — impresión de etiquetas en lote (Zebra ZD421)

Código VBA optimizado para la impresión de etiquetas de `IMP_STOCK_DETALLADO.xlsm`.
Sustituye a la macro `VistaPreviaEImpresion`, que imprimía etiqueta por etiqueta y no
respetaba el autofiltro de la hoja `Consolidado`.

| Antes | Ahora |
|---|---|
| Imprimía también las filas ocultas por el filtro | Sólo filas **visibles**; si hay selección, manda la selección |
| 1 trabajo de impresión por etiqueta (200 etiquetas = 200 trabajos, cola saturada) | **1 trabajo por cada 100 etiquetas**, una página = una etiqueta |
| La primera etiqueta se perdía en la vista previa | Se imprimen todas, exactamente una vez |
| 1 copia por fila, sin opción | **Copias a elegir**: número fijo o la columna `CANTIDAD` |
| La escala dependía del papel del driver | Escala calculada para **10 × 5 cm** en cualquier impresora |
| Excel se quedaba colgado con lotes grandes | Repaginación desactivada durante el armado + `DoEvents` |
| Se caía con celdas `#N/A`, descripciones con salto de línea, etc. | Valores saneados antes de imprimir |
| Código de barras tal cual (podía no ser legible) | Mayúsculas + aviso de caracteres no válidos en Code 39 |

## Archivos

```
src/modEtiquetas.bas              módulo del motor de impresión (para IMPORTAR)
src/Modulo1.bas                   reemplazo de Módulo1: sólo queda ValidarDato
src/*_para_pegar.bas              los mismos módulos para COPIAR Y PEGAR
src/_original/                    código tal como estaba en el .xlsm (respaldo)
docs/ANALISIS.md                  análisis del VBA original y defectos corregidos
docs/MAPEO_Y_ZEBRA.md             mapeo celda <- columna y configuración de la ZD421
```

> **Importar o pegar, pero no mezclar.** Los `.bas` empiezan con líneas `Attribute ...` que
> el editor de VBA sólo acepta al **importar** el archivo. Si copia y pega ese contenido,
> al compilar aparece **"Atributo no válido en Sub o Function"**: para pegar use los
> archivos `*_para_pegar.bas`.

## Instalación

1. Abrir `IMP_STOCK_DETALLADO.xlsm` y **guardar una copia de respaldo**.
2. `Alt + F11` para abrir el editor de VBA.
3. Si ya tenía una versión anterior: clic derecho sobre `modEtiquetas` > `Quitar modEtiquetas...` > *No*.
4. `Archivo > Importar archivo...` y elegir `src/modEtiquetas.bas`.
5. **Reemplazar `Módulo1`**: clic derecho > `Quitar Módulo1...` > *No*, y después
   `Archivo > Importar archivo...` con `src/Modulo1.bas`.
   Este paso es obligatorio: el `Módulo1` viejo llama a una macro que ya no existe y el
   proyecto no compilaría.
6. `Depuración > Compilar VBAProject` para verificar que no hay errores.
7. Guardar como `.xlsm`.

### Combinación de teclas

Al importar, `ImprimirEtiquetas` queda con `Ctrl + P`. Para cambiarla:
`Vista > Macros > ImprimirEtiquetas > Opciones...` y escribir la letra deseada.

## Uso

1. En `Consolidado`, aplicar los filtros de siempre.
2. **Seleccionar las filas** que se quieren etiquetar (una, dos, las que sean; con `Ctrl`
   se eligen salteadas).
3. Lanzar `ImprimirEtiquetas` con la combinación de teclas.
4. Responder **una sola pregunta**: cuántas copias de cada fila.

```
ETIQUETAS
--------------------------------------------
Filas a imprimir: 2   (seleccionadas)
Impresora: ZDesigner ZD421-203dpi ZPL
--------------------------------------------

Copias de CADA fila:

     1        una etiqueta por fila
     2, 3 ... ese número de copias de cada fila
     C        usar la columna CANTIDAD de cada fila
```

Aceptar e imprime. Nada más.

- **Sin filas seleccionadas**, antes pregunta cuántas filas del filtro imprimir
  (`TODAS`, `100`, `101-300`).
- A partir de **300 etiquetas** pide una confirmación extra.
- Al terminar informa cuántas etiquetas se enviaron, en cuántos trabajos, la escala
  aplicada y los avisos de código de barras.

No hay vista previa: la ventana de vista previa de Excel tiene su propio botón *Imprimir*
y usarlo duplicaba el lote. Para verificar antes de un lote grande, seleccione **una fila**
e imprima esa sola etiqueta.

### Macros

| Macro | Para qué |
|---|---|
| `ImprimirEtiquetas` | **la única entrada**: imprime las filas seleccionadas o las del filtro |
| `EtiquetasDiagnostico` | informe de verificación: no imprime |
| `ValidarDato` | función auxiliar heredada (`Módulo1`), por si alguna fórmula la usa |

## Ajustes (parte superior de `modEtiquetas.bas`)

| Constante | Valor | Para qué |
|---|---|---|
| `ANCHO_ETIQUETA_MM` / `ALTO_ETIQUETA_MM` | `100` / `50` | tamaño físico de la etiqueta. La escala se calcula para que el bloque `A1:J17` mida exactamente esto |
| `PAGINAS_POR_TRABAJO` | `100` | etiquetas por trabajo de impresión. Si la cola del driver se atasca, bajar a `50` |
| `TOPE_ETIQUETAS` | `10000` | tope de seguridad por corrida |
| `AVISO_DESDE` | `300` | desde cuántas etiquetas se pide confirmación extra |
| `ESCALA_FIJA` | `0` | `0` = calcular la escala; `1`-`100` = forzar ese porcentaje |
| `IMPRESORA_CONTIENE` | `"ZD421"` | texto que identifica a la Zebra |
| `ACTUALIZAR_ETQ` | `True` | deja cargada en `ETQ` la primera etiqueta del lote |
| `CONSERVAR_LOTE` | `False` | `True` = no borra `ETQ_LOTE` (para revisar el lote armado) |

## Rendimiento

La tabla `Consolidado` tiene 130.180 filas. Un lote de 100 etiquetas se arma con 7
operaciones de copiado y se envía en **1 solo trabajo de impresión**; el tiempo que tarda
después es el del driver rasterizando las páginas, no el de Excel.

Para que Excel no se quede "sin responder" durante el armado: la repaginación automática de
la hoja temporal se desactiva (`DisplayPageBreaks = False`, era la causa principal del
cuelgue), los formatos se aplican una vez antes de replicar los bloques, y hay `DoEvents`
en el armado y entre trabajos. `Esc` cancela de forma limpia.

## Si algo sale mal

| Síntoma | Causa / solución |
|---|---|
| `Atributo no válido en Sub o Function` | se pegó un `.bas` en vez de importarlo: use los `*_para_pegar.bas` |
| `Sub o Function no definida` al compilar | falta reemplazar `Módulo1` por el nuevo (paso 5 de la instalación) |
| "Debe estar en la hoja Consolidado o IMPRIMIR" | ejecutar la macro con esa hoja activa |
| "No hay filas visibles para imprimir" | el filtro no deja ninguna fila, o la selección está toda oculta |
| Las etiquetas salen **muy pequeñas** o descolocadas | ejecute `EtiquetasDiagnostico`: muestra el área imprimible real, la escala calculada y si un lote de 2 ocupa 2 páginas |
| Las etiquetas salen **cortadas** | el papel del driver es más chico que 100 × 50 mm: corríjalo en las preferencias de la impresora (tamaño definido por el usuario) |
| El código de barras no se lee | ver el aviso al final del lote: caracteres no válidos para Code 39 o código demasiado largo (`docs/MAPEO_Y_ZEBRA.md`) |
| Quedó una hoja `ETQ_LOTE` | se interrumpió el proceso: se puede borrar a mano, la siguiente corrida la borra sola |
