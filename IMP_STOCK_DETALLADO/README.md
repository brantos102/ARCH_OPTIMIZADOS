# IMP_STOCK_DETALLADO — impresión de etiquetas en lote (Zebra ZD421)

Código VBA optimizado para la impresión de etiquetas de `IMP_STOCK_DETALLADO.xlsm`.
Sustituye a la macro `VistaPreviaEImpresion`, que imprimía etiqueta por etiqueta y no
respetaba el autofiltro de la hoja `Consolidado`.

| Antes | Ahora |
|---|---|
| Imprimía también las filas ocultas por el filtro | Sólo filas **visibles** del filtro |
| 1 trabajo de impresión por etiqueta (200 etiquetas = 200 trabajos) | **1 trabajo por cada 100 etiquetas** |
| La primera etiqueta se perdía (quedaba sólo en la vista previa) | Se imprimen todas |
| 1 copia por fila, sin opción | **Copias a elegir**: número fijo o la columna `CANTIDAD` |
| Se caía con celdas `#N/A`, descripciones con salto de línea, etc. | Valores saneados antes de imprimir |
| Código de barras tal cual (podía no ser legible) | Mayúsculas + aviso de caracteres no válidos en Code 39 |

## Archivos

```
src/modEtiquetas.bas            módulo nuevo: motor de impresión por lotes
src/Modulo1.bas                 reemplazo de Módulo1, para IMPORTAR (conserva el atajo)
src/Modulo1_para_pegar.bas      el mismo código, para COPIAR Y PEGAR (sin líneas Attribute)
src/_original/                  código tal como estaba en el .xlsm (respaldo)
docs/ANALISIS.md                análisis del VBA original y defectos corregidos
docs/MAPEO_Y_ZEBRA.md           mapeo celda <- columna y configuración de la ZD421
```

> **Importar o pegar, pero no mezclar.** Los archivos `.bas` empiezan con líneas
> `Attribute ...` que el editor de VBA sólo acepta al **importar** el archivo. Si copia y
> pega ese contenido dentro de un módulo, al compilar aparece
> **"Atributo no válido en Sub o Function"**. Para pegar use
> `src/Modulo1_para_pegar.bas`, que no las lleva.

## Instalación (5 minutos)

1. Abrir `IMP_STOCK_DETALLADO.xlsm` y **guardar una copia de respaldo** del archivo.
2. `Alt + F11` para abrir el editor de VBA.
3. **Importar el módulo nuevo**: `Archivo > Importar archivo...` y elegir
   `src/modEtiquetas.bas`.
4. **Reemplazar `Módulo1`**:
   - En el panel de proyecto, clic derecho sobre `Módulo1` > `Quitar Módulo1...` >
     *No* (no hace falta exportarlo, ya está en `src/_original/`).
   - `Archivo > Importar archivo...` y elegir `src/Modulo1.bas`.
   - Así se conserva el atajo de teclado que ya tenía la macro.

   > Si prefiere no quitar el módulo: abra `Módulo1`, borre todo su contenido y pegue el de
   > **`src/Modulo1_para_pegar.bas`** (es el mismo código sin las líneas `Attribute`, que el
   > editor no acepta pegadas). En ese caso vuelva a asignar el atajo en
   > `Vista > Macros > VistaPreviaEImpresion > Opciones...`, tecla `P`.
5. `Ctrl + G` (ventana Inmediato) y `Depuración > Compilar VBAProject` para verificar que
   compila sin errores.
6. Guardar el libro como `.xlsm`.

No hay que tocar las hojas `ETQ`, `Consolidado` ni `IMPRIMIR`: el módulo trabaja sobre una
hoja temporal (`ETQ_LOTE`) que se elimina al terminar.

## Uso

1. En la hoja **`Consolidado`**, aplicar los filtros de siempre (cliente, posición, estado...).
2. Opcional: seleccionar las filas concretas que se quieren etiquetar.
3. Ejecutar la macro (el atajo de siempre, o `Vista > Macros > ImprimirEtiquetasZebra`).
   La primera vez conviene probar con `ImprimirUnaEtiquetaPrueba`: imprime una sola
   etiqueta para confirmar tamaño, posición y lectura del código de barras.
4. Responder los 4 pasos:

| Paso | Cuándo aparece | Respuestas |
|---|---|---|
| 1 · Alcance | sólo si hay filas seleccionadas | `Sí` = imprimir **sólo esas N** (muestra el número) · `No` = todas las filas visibles del filtro |
| 2 · Filas a imprimir | **sólo** si va a imprimir todo el filtro, o si seleccionó más de 100 filas | `TODAS` · `200` (las primeras 200) · `101-300` (un rango) |
| 3 · Copias de cada etiqueta | siempre | `1`, `2`, `3`... (fijo para todas) · `C` (usar la columna `CANTIDAD`) |
| 4 · Confirmación | siempre | `Sí` = ver vista previa primero · `No` = imprimir ya · `Cancelar` |

> Si selecciona 2 filas y responde `Sí` en el paso 1, se imprimen **2 etiquetas**: el paso 2
> ni siquiera aparece. El filtro se respeta siempre, en los dos caminos.

Al terminar se informa cuántas etiquetas se enviaron, en cuántos trabajos, a qué impresora,
cuántas filas se omitieron (filas sin código ni serie) y los avisos de código de barras.

### Macros disponibles

| Macro | Para qué |
|---|---|
| `VistaPreviaEImpresion` | nombre y atajo de siempre; ahora llama al asistente completo |
| `ImprimirEtiquetasZebra` | asistente de 4 pasos |
| `ImprimirEtiquetasSeleccion` | directo: sólo las filas seleccionadas (visibles) |
| `ImprimirEtiquetasFiltroCompleto` | directo: todas las filas visibles del filtro |
| `ImprimirUnaEtiquetaPrueba` | imprime **una sola** etiqueta de prueba, para verificar tamaño antes de un lote |
| `EtiquetasDiagnostico` | informe: filas visibles, impresora, papel, **escala medida** y mapeo de columnas |
| `VistaPreviaEImpresion_Clasica` | respaldo: lógica original corregida, una etiqueta por trabajo |

## Ajustes (parte superior de `modEtiquetas.bas`)

| Constante | Valor | Para qué |
|---|---|---|
| `PAGINAS_POR_TRABAJO` | `100` | etiquetas por trabajo de impresión. Si la cola del driver se traba, bajar a `50` |
| `TOPE_ETIQUETAS` | `10000` | tope de seguridad por corrida |
| `AVISO_DESDE` | `300` | desde cuántas etiquetas se pide confirmación extra |
| `ANCHO_ETIQUETA_MM` / `ALTO_ETIQUETA_MM` | `100` / `50` | tamaño físico de la etiqueta (10 x 5 cm). La escala se calcula para que el bloque `A1:J17` mida exactamente esto, en cualquier impresora |
| `ESCALA_FIJA` | `0` | `0` = calcular la escala con el tamaño de etiqueta; `1`-`100` = forzar ese porcentaje |
| `IMPRESORA_CONTIENE` | `"ZD421"` | texto que identifica a la Zebra |
| `ACTUALIZAR_ETQ` | `True` | deja cargada en `ETQ` la primera etiqueta del lote |
| `CONSERVAR_LOTE` | `False` | `True` = no borra `ETQ_LOTE` (para revisar el lote armado) |

## Rendimiento medido sobre este archivo

La tabla `Consolidado` tiene 130.180 filas y, con el filtro `cliente_id = VENTURA`,
**9.542 filas visibles**. Un lote de 200 etiquetas se arma con 8 operaciones de copiado y se
envía en 2 trabajos de impresión (antes: 200 trabajos).

## Si algo sale mal

| Síntoma | Causa / solución |
|---|---|
| "Debe estar en la hoja Consolidado o IMPRIMIR" | ejecutar la macro con esa hoja activa |
| "No se encontraron filas visibles" | el filtro no deja ninguna fila, o la selección está toda oculta |
| Las etiquetas salen **muy pequeñas** o descolocadas | es la escala. Ejecute `EtiquetasDiagnostico`: muestra el área imprimible real, la escala calculada y si un lote de 2 ocupa 2 páginas |
| Las etiquetas salen **cortadas** | el papel del driver es más chico que 100 x 50 mm: corríjalo en las preferencias de la impresora (tamaño definido por el usuario) |
| Aviso de paginación | el papel de `ETQ` no es la etiqueta de 100 x 50 mm, o la impresora activa no es la ZD421: revisar la vista previa antes de continuar |
| Etiquetas en blanco intercaladas | misma causa anterior (una etiqueta ocupa más de una página) |
| El código de barras no se lee | ver el aviso al final del lote: caracteres no válidos para Code 39 o código demasiado largo (`docs/MAPEO_Y_ZEBRA.md`) |
| `Atributo no válido en Sub o Function` al compilar | se pegó un `.bas` en vez de importarlo: borre las líneas que empiezan con `Attribute` (o use `src/Modulo1_para_pegar.bas`) y reasigne el atajo en `Vista > Macros > Opciones...` |
| Quedó una hoja `ETQ_LOTE` | se interrumpió el proceso: se puede borrar a mano, la siguiente corrida la borra sola |
