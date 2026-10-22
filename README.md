# Cuadro de Mando SIMCARDS — hoja `CUADRO DE MANDO T` optimizada

Corrección y optimización del tablero de rotación y cobertura de inventario de
SIM cards. **No se adultera ningún dato de origen**: las hojas `PMP`,
`REFERENCIA`, `RESUMEN INVENTARIO T`, `RESUMEN DESPACHO` y `DESPACHOS 20 DIAS`
quedan idénticas celda por celda, y de `INVENTARIO` solo cambia una celda que
estaba corrupta.

## Archivos

| Ruta | Descripción |
|---|---|
| `data/CUADRO_DE_MANDO_SIMCARDS_ORIGINAL.xlsx` | Archivo recibido, sin modificar (respaldo) |
| `data/CUADRO_DE_MANDO_SIMCARDS_OPTIMIZADO.xlsx` | **Entregable** |
| `scripts/optimizar_cuadro.py` | Script que regenera el entregable desde el original |
| `docs/CAMBIOS.md` | Detalle de cada corrección y su impacto |

Regenerar (requiere `openpyxl`; es idempotente, siempre parte del original):

```bash
python3 scripts/optimizar_cuadro.py \
  data/CUADRO_DE_MANDO_SIMCARDS_ORIGINAL.xlsx \
  data/CUADRO_DE_MANDO_SIMCARDS_OPTIMIZADO.xlsx
```

## El error de fondo

La fila 31 sumaba `=SUM(H28:H30)` cuando el bloque MOVISTAR va de la fila 26 a
la 30, y la fila 44 sumaba `=SUM(H26:H43)`, invadiendo MOVISTAR completo más su
propio subtotal.

| Grupo | Stock antes | Stock ahora | Corrección |
|---|---:|---:|---:|
| KITE M2M | 93.292 | 93.292 | 0 |
| M2M LTE | 7.936 | 7.936 | 0 |
| **MOVISTAR** | 6.920 | **13.468** | **+6.548** |
| **TIGO** | 1.473.990 | **1.453.602** | **−20.388** |
| AKIMOVIL | 30.351 | 30.351 | 0 |
| MAXIPLUS | 16.187 | 16.187 | 0 |
| **Suma de grupos** | 1.628.676 | **1.614.836** | **−13.840** |

La suma coincide ahora exactamente con la hoja `INVENTARIO` (1.614.836) y con
`DESPACHOS 20 DIAS` (332.431).

## Estructura

Se mantienen los 6 grupos originales y se agrega un segundo nivel por familia
de producto, cada uno con sus cifras completas:

```
KITE M2M      ├─ SIM / HALFSIM KITE M2M ESTÁNDAR
              └─ HALFSIM KITE M2M DEVICE CONTROL
M2M LTE       └─ (familia única)
MOVISTAR      ├─ HALFSIM MOVISTAR (FÍSICA)
              └─ ESIM MOVISTAR (DIGITAL)
TIGO          ├─ HALFSIM TIGO (COMERCIAL)
              ├─ HALFSIM TIGO CHIP PREPAGO $3 (NUMERADA)
              ├─ ESIM TIGO DIGITAL
              └─ HALFSIM TIGO M2M
AKIMOVIL      ├─ HALFSIM AKIMOVIL (FÍSICA)
              ├─ HALFSIM AKIMOVIL NUMERADA
              └─ ESIM AKIMOVIL (DIGITAL)
MAXIPLUS      ├─ HALFSIM MAXIPLUS (FÍSICA)
              ├─ HALFSIM MAXIPLUS NUMERADA
              └─ ESIM MAXIPLUS (DIGITAL)
```

Tres niveles de banda: familia (azul claro), grupo (azul fuerte, el de siempre)
y TOTAL GENERAL (azul oscuro).

## Columnas nuevas

| Col. | Nombre | Para qué |
|---|---|---|
| `A` | **CLAVE** | Familia de cada artículo. Alimenta los `SUMIFS` de todas las bandas. Puede ocultarse. |
| `R` | **ESTADO** | `SIN ROTACIÓN` / `CRÍTICO` / `ALERTA` / `OK`. Separa el stock inmovilizado del sano: antes ambos salían en verde. |
| `S` | **ALERTA DE PRODUCCIÓN** | Avisa cuando la cobertura total está sana pero el producto terminado está en rojo. |

La columna `S` detecta el riesgo que el tablero no mostraba: **`HALFSIM TIGO`
tiene 435 días de cobertura total (verde) y 4 días de producto terminado**.
No es un problema de abastecimiento sino de capacidad de producción.

## Mantenimiento perpetuo

Los subtotales ya **no** usan rangos de filas:

```excel
=SUMIFS(H$13:H$73, $A$13:$A$73, "TIGO.CHIP3")   ' familia
=SUMIFS(H$13:H$73, $A$13:$A$73, "TIGO.*")       ' grupo
=SUMIFS(H$13:H$73, $A$13:$A$73, "?*")           ' TOTAL GENERAL
```

**Agregar un artículo:** inserte la fila dentro de su bloque, copie las
fórmulas de la fila de arriba y escriba la misma CLAVE en la columna A. Los
tres niveles se actualizan solos; es imposible que un bloque quede fuera o se
cuente dos veces.

**Agregar una familia:** use una clave `GRUPO.NOMBRE` y replique la banda de
subtotal. El total del grupo (`GRUPO.*`) ya la incluye.

**Cargar un período nuevo:** pegue los datos en `INVENTARIO` y
`DESPACHOS 20 DIAS` y revise las tres filas de **CONTROL DE CUADRATURA** bajo
la tabla. Si alguna no dice `OK`, hay artículos o canales que el cuadro no está
sumando.

## Divisor del promedio diario — atención

Antes era la constante `20`. Ahora:

| Celda | Contenido |
|---|---|
| `I7` | **AUTO** (o escriba un número para fijarlo) |
| `I8` | Días con despacho detectados automáticamente |
| `I9` | Divisor aplicado — celda nombrada `DIAS_PERIODO` |

En **AUTO** el divisor es el período real. Con los datos actuales son **17
días** (02/10 a 22/10/2026), no 20, así que **todos los días de stock bajan
alrededor de un 15 %**: es la cobertura correcta, antes estaba sobreestimada.

Para volver exactamente a las cifras publicadas antes, escriba `20` en `I7`.
Verificado: con `I7 = 20` las 41 filas reproducen el original **sin una sola
diferencia**.
