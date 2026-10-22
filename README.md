# Cuadro de Mando SIMCARDS — hoja `CUADRO DE MANDO T` optimizada

Corrección y optimización del tablero de rotación y cobertura de inventario de
SIM cards, conservando la lógica de negocio, los grupos existentes y el resto
del libro sin cambios.

## Archivos

| Ruta | Descripción |
|---|---|
| `data/CUADRO_DE_MANDO_SIMCARDS_ORIGINAL.xlsx` | Archivo recibido, sin modificar (respaldo) |
| `data/CUADRO_DE_MANDO_SIMCARDS_OPTIMIZADO.xlsx` | **Entregable** |
| `scripts/optimizar_cuadro.py` | Script que regenera el entregable desde el original |
| `docs/CAMBIOS.md` | Detalle de cada cambio y el impacto en cifras |

Regenerar:

```bash
python3 scripts/optimizar_cuadro.py \
  data/CUADRO_DE_MANDO_SIMCARDS_ORIGINAL.xlsx \
  data/CUADRO_DE_MANDO_SIMCARDS_OPTIMIZADO.xlsx
```

Requiere `openpyxl`. El script es idempotente: siempre parte del original.

## Qué se corrigió

El error principal estaba en los subtotales de la hoja. **La fila 31 sumaba
`=SUM(H28:H30)`** cuando el bloque MOVISTAR va de la fila 26 a la 30: dejaba
fuera dos artículos. Y **la fila 44 sumaba `=SUM(H26:H43)`**, invadiendo el
bloque MOVISTAR completo más su propio subtotal, duplicando esas cifras dentro
del total de TIGO.

| Grupo | Stock antes | Stock ahora | Corrección |
|---|---:|---:|---:|
| KITE M2M | 93.292 | 93.292 | 0 |
| M2M LTE | 7.936 | 7.936 | 0 |
| **MOVISTAR** | 6.920 | **13.468** | **+6.548** |
| **TIGO** | 1.473.990 | **1.453.602** | **−20.388** |
| AKIMOVIL | 30.351 | 30.351 | 0 |
| MAXIPLUS | 16.187 | 16.187 | 0 |
| **Suma de grupos** | 1.628.676 | **1.614.836** | **−13.840** |

La suma de los grupos ahora coincide exactamente con el total de la hoja
`INVENTARIO` (1.614.836) y con el total de `DESPACHOS 20 DIAS` (332.431).

## Cómo quedó estructurada

Se mantienen los 6 grupos (equipos) originales y se agrega un segundo nivel de
separación por familia de producto, con sus cifras propias:

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

Tres niveles de banda: subtotal de familia (azul claro), total de grupo (azul
fuerte, el de siempre) y TOTAL GENERAL (azul oscuro).

## Mantenimiento perpetuo

Los subtotales ya **no** usan rangos fijos. Cada banda es:

```excel
=SUMIFS(H$13:H$73, $A$13:$A$73, "TIGO.CHIP3")     ' subtotal de familia
=SUMIFS(H$13:H$73, $A$13:$A$73, "TIGO.*")         ' total de grupo
=SUMIFS(H$13:H$73, $A$13:$A$73, "?*")             ' TOTAL GENERAL
```

La columna **A (`CLAVE`)** marca a qué familia pertenece cada artículo. Es la
única columna que hay que mantener.

**Para agregar un artículo:** inserte la fila dentro de su bloque, copie las
fórmulas de la fila de arriba y escriba la misma CLAVE en la columna A. Los
tres niveles de total se actualizan solos y es imposible que un bloque quede
fuera o se cuente dos veces.

**Para agregar una familia nueva:** use una clave `GRUPO.NOMBRE` y agregue su
banda de subtotal con el mismo patrón. El total del grupo (`GRUPO.*`) ya la
incluye automáticamente.

**Al cargar un período nuevo:** pegue los datos en `INVENTARIO` y
`DESPACHOS 20 DIAS` y revise las dos filas de **CONTROL DE CUADRATURA** bajo la
tabla. Si alguna no dice `OK`, hay artículos en las hojas origen que no están
en el cuadro y se están perdiendo.

## Divisor del promedio diario

Antes estaba escrito como constante (`=H13/20`). Ahora apunta a la celda
nombrada **`DIAS_PERIODO`** (`I7`), que se deja en **20** para no alterar
ninguna cifra publicada.

La celda `I8` detecta automáticamente cuántos días distintos con despacho trae
la hoja origen y `E9` avisa si no coinciden. Con los datos actuales el período
real es de **17 días**, no 20: con el divisor en 20 la cobertura se sobreestima
alrededor de un 18 %. Cambiando `I7` a 17 todo el tablero se recalcula.
