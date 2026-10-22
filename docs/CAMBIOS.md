# Detalle de cambios — `CUADRO DE MANDO T`

## Garantía de integridad de los datos

Verificación celda por celda contra el original:

| Hoja | Celdas distintas |
|---|---|
| `PMP` | **0** |
| `REFERENCIA` | **0** |
| `RESUMEN INVENTARIO T` | **0** |
| `RESUMEN DESPACHO` | **0** |
| `DESPACHOS 20 DIAS` | **0** |
| `INVENTARIO` | **1** (el NBSP roto, ver punto 8) |

Las 3 tablas dinámicas, las 5 imágenes y las 7 hojas se conservan.

> Nota técnica: `openpyxl` serializa los flotantes con 15 dígitos
> significativos, lo que convierte `40,709999999999994` en
> `40,70999999999999` — un `double` distinto. El paso 12 del script restaura el
> texto original de los **849** valores numéricos afectados en las hojas no
> intervenidas, de modo que quedan bit a bit como estaban.

---

## 1. Subtotales mal rangeados (causa raíz)

| Fila | Fórmula original | Bloque real | Problema |
|---|---|---|---|
| 21 | `=SUM(H13:H20)` | 13–20 | correcta |
| 25 | `=SUM(H22:H24)` | 22–24 | correcta |
| **31** | `=SUM(H28:H30)` | **26–30** | omite las filas 26 y 27 |
| **44** | `=SUM(H26:H43)` | **32–43** | invade MOVISTAR y su subtotal |
| 53 | `=SUM(H45:H52)` | 45–52 | correcta |
| 59 | `=SUM(H54:H58)` | 54–58 | correcta |

En la fila 31, `P31` mostraba 6.920 contra 13.468 reales (−49 %) y `H31`
mostraba 0 despachos habiendo 2.

En la fila 44 sobraban 20.388 unidades: exactamente el bloque MOVISTAR (13.468)
más el subtotal de la fila 31 (6.920), contados dos veces.

**Ahora** ninguna banda depende de un rango de filas:

```excel
=SUMIFS(H$13:H$73, $A$13:$A$73, "MOVISTAR.HALFSIM")   ' familia
=SUMIFS(H$13:H$73, $A$13:$A$73, "MOVISTAR.*")         ' grupo
=SUMIFS(H$13:H$73, $A$13:$A$73, "?*")                 ' TOTAL GENERAL
```

El rango `$13:$73` se expande solo al insertar filas, y la pertenencia la
define la CLAVE de la columna A, no la posición.

## 2. Segundo nivel de separación

14 subtotales de familia dentro de los 6 grupos existentes, con todas sus
cifras. Ningún artículo cambió de grupo; los 41 siguen siendo los mismos. Se
reagrupó lo mínimo para que cada familia quede contigua:

- **MOVISTAR**: las dos ESIM (antes filas 26 y 30) quedaron juntas al final.
- **TIGO**: PORTABILIDAD y REPOSICIÓN subieron junto a HALFSIM TIGO; los cuatro
  CHIP PREPAGO $3 quedaron contiguos.
- **AKIMOVIL**: las numeradas se ordenaron 7D → 15D → 30D.

## 3. Días de stock: de texto a número

**Antes** las celdas devolvían el **texto** `">150"` / `">15"` cuando no había
rotación. Eso impedía ordenar, promediar o graficar las columnas `M`, `O` y
`Q`, y como en Excel un texto es mayor que cualquier número, el formato
condicional las pintaba de **verde**: las 25 referencias sin ningún movimiento
(144.271 unidades inmovilizadas) aparecían como saludables.

**Ahora**:

```excel
=IF($I13=0,"",L13/$I13)
```

La celda es un número real. El rótulo `">150"` se conserva con un formato
numérico condicional — `[>150]">150";#,##0` para Blanca y `[>15]">15";#,##0`
para Numerada — así que el aspecto es el mismo pero la columna ya es
calculable. Cuando no hay rotación la celda queda vacía (no hay denominador) y
se pinta de gris.

## 4. Columna `R` — ESTADO

`SIN ROTACIÓN` / `CRÍTICO` / `ALERTA` / `OK`, con los mismos umbrales de
siempre (Numerada: <8 / 8–14 / ≥15; Blanca: ≤90 / 91–150 / >150). Separa por
fin el stock inmovilizado del stock sano. En las bandas que mezclan Blanca y
Numerada queda vacía, porque los umbrales de los dos tipos son distintos y
clasificarlas sería engañoso.

## 5. Columna `S` — ALERTA DE PRODUCCIÓN

Se activa cuando la cobertura **total** está sana pero el **producto
terminado** está en rojo. Solo en filas de artículo, que es donde la acción
aplica. Con los datos actuales:

| Artículo | Días totales | Días terminado |
|---|---:|---:|
| `HALFSIM TIGO` | 435 | **4** |
| `ESIM TIGO DIGITAL` | 220 | **14** |
| `HALFSIM KITE M2M 3FF DEVICE CONTROL` | 102 | **26** |
| `HALFSIM KITE M2M 4FF DEVICE CONTROL` | 124 | **43** |

Los dos primeros tienen 741.743 y 139.986 unidades de materia prima esperando
conversión. El riesgo no es de abastecimiento sino de capacidad de producción,
y el tablero no lo mostraba en ninguna parte.

## 6. Divisor del promedio diario

`=H13/20` → `=IFERROR(H13/DIAS_PERIODO,0)`.

| Celda | Contenido |
|---|---|
| `I7` | `AUTO`, o un número para fijarlo |
| `I8` | Días con despacho detectados (cuenta de fechas distintas, con rango autoajustable) |
| `I9` | Divisor aplicado = `DIAS_PERIODO` |

El período cargado tiene **17 días** (02/10 a 22/10/2026), no 20. En `AUTO`
todos los días de stock bajan ~15 %, que es la cobertura correcta. Con `I7 = 20`
se reproducen exactamente las cifras anteriores.

El `IFERROR` evita además el `#¡DIV/0!` si la celda queda en blanco.

## 7. Fila TOTAL GENERAL y control de cuadratura

| Verificación | Cuadro | Hoja origen | Dif. |
|---|---:|---:|---:|
| Stock total vs. `INVENTARIO` | 1.614.836 | 1.614.836 | 0 OK |
| Despacho total vs. `DESPACHOS 20 DIAS` | 332.431 | 332.431 | 0 OK |
| Líneas de `DESPACHOS` cuyo CANAL no existe en `REFERENCIA` | — | — | 0 OK |

Es el control que cubre el riesgo silencioso del diseño original: si llega un
artículo nuevo a las hojas origen y nadie lo da de alta en el cuadro, el
tablero lo ignoraba sin avisar. Ahora la diferencia deja de ser 0.

## 8. Espacios no separables (NBSP)

`HALFSIM KITE M2M 4FF DEVICE CONTROL BLANCA` tenía caracteres `U+00A0` en lugar
de espacios, en el cuadro y en `INVENTARIO!B46`. Funcionaba de casualidad
porque el error estaba en las dos hojas; cualquier recarga o retipeo de una de
ellas habría roto el `SUMIFS` en silencio y la fila habría mostrado 0. Ambos
normalizados.

## 9. Formato condicional

- Reconstruido sobre valores numéricos, con los mismos colores y umbrales.
- Se cerró el hueco del valor exacto **15** en `Tipo = Numerada`: las reglas
  eran `<8`, `>=8 y <15` y `>15`, así que 15 no caía en ninguna.
- `SIN ROTACIÓN` tiene ahora su propio gris, en vez de verde.
- Fuente oscura en los semáforos: en las bandas el texto es blanco y sobre un
  relleno claro quedaba ilegible.
- Las tres reglas de la leyenda (`L7`, `L8`, `L9`) se conservan intactas.
- El umbral `>150` de las SIM Blanca **no** se tocó.

## 10. Subtotales que faltaban y formatos

Las bandas originales solo totalizaban `H`, `I`, `L`, `N` y `P`. Se agregaron
`F` (canal indirecto), `G` (canal directo), `J` (en proceso de liberación),
`K` (en producción) y los días de stock `M`, `O`, `Q`.

Formatos: cantidades `#,##0`; promedio diario `#,##0.00` (antes era entero y
los artículos de 0,05/día se leían como `0`); días de stock con el formato
condicional del punto 3.

---

## Verificación realizada

Recalculado con LibreOffice Calc:

- **0 errores de fórmula** (`#REF!`, `#N/A`, `#¡DIV/0!`, `#NAME?`).
- Con `I7 = 20`, las 41 filas de detalle devuelven **0 diferencias** contra el
  archivo original, en las 11 columnas calculadas.
- **0 celdas de días de stock siguen siendo texto**.
- Las tres filas de control de cuadratura dan **0 / OK**.
- Hojas de origen intactas (tabla al inicio de este documento).

---

## Puntos que quedan abiertos (requieren decisión del área)

No se aplicaron porque implican modificar datos de origen o criterios de
negocio:

1. **`FECHA` es texto, no fecha** en `DESPACHOS 20 DIAS` (1.371 filas,
   `'02/10/2026'`). Impide filtrar por rango. El conteo automático de días de
   `I8` ya funciona sobre texto, así que no bloquea nada hoy.
2. **`PMP` es texto, no número** (`'0.69'`), y la columna `VALORADO` es
   inconsistente: para la misma línea de 1 unidad a 0,69 aparecen 69, 138, 207
   o 1.380, y 13 filas de `HALFSIM MOVISTAR` con valorado 0. El total de
   224.897 USD no es confiable.
3. **Hojas residuales.** `RESUMEN INVENTARIO T`, `RESUMEN DESPACHO` y `PMP`
   pertenecen a otro catálogo (códigos SCL, productos Tuenti/Multisim) y
   ninguna fórmula las referencia. Son ~40 % del peso del archivo y muestran
   totales de inventario que no reconcilian con el cuadro.
4. **Akimovil y Maxiplus: 46.538 unidades con cero movimiento** en todo el
   período. O hay un canal que no se está capturando, o es obsolescencia que
   debería provisionarse.
5. **`STOCK EN PROCESO DE LIBERACIÓN` (columna J)** es cero constante y de
   carga manual en las 41 filas. Si no se usa, conviene retirarla.
