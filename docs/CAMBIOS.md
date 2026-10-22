# Detalle de cambios — `CUADRO DE MANDO T`

Todo lo que no se menciona aquí quedó **sin modificar**: hojas `PMP`,
`REFERENCIA`, `RESUMEN INVENTARIO T`, `RESUMEN DESPACHO`, `DESPACHOS 20 DIAS`,
las 3 tablas dinámicas, las 5 imágenes, el formato condicional de la leyenda y
los estilos del encabezado.

---

## 1. Subtotales mal rangeados (causa raíz)

**Antes**

| Fila | Fórmula | Bloque real | Problema |
|---|---|---|---|
| 21 | `=SUM(H13:H20)` | 13–20 | correcta |
| 25 | `=SUM(H22:H24)` | 22–24 | correcta |
| **31** | `=SUM(H28:H30)` | **26–30** | omite las filas 26 y 27 |
| **44** | `=SUM(H26:H43)` | **32–43** | invade MOVISTAR y su subtotal |
| 53 | `=SUM(H45:H52)` | 45–52 | correcta |
| 59 | `=SUM(H54:H58)` | 54–58 | correcta |

Efecto medido en la fila 31: `P31` mostraba 6.920 contra 13.468 reales (−49 %),
y `H31` mostraba 0 despachos cuando había 2.

Efecto en la fila 44: +20.388 unidades de más, que es exactamente el bloque
MOVISTAR (13.468) más el subtotal de la fila 31 (6.920) contados dos veces.

**Ahora** — ninguna banda depende de un rango de filas:

```excel
=SUMIFS(H$13:H$73, $A$13:$A$73, "MOVISTAR.HALFSIM")   ' subtotal de familia
=SUMIFS(H$13:H$73, $A$13:$A$73, "MOVISTAR.*")         ' total de grupo
=SUMIFS(H$13:H$73, $A$13:$A$73, "?*")                 ' TOTAL GENERAL
```

El rango `$13:$73` se expande solo al insertar filas dentro de la tabla, y la
pertenencia la define la CLAVE de la columna A, no la posición.

## 2. Segundo nivel de separación

Se agregaron 14 subtotales de familia dentro de los 6 grupos existentes, con
todas sus cifras (despacho por canal, promedio diario, stock por etapa y días
de stock). El orden de los artículos se conservó, salvo dos reagrupaciones
necesarias para que cada familia quede contigua:

- **MOVISTAR**: las dos ESIM (antes filas 26 y 30) quedaron juntas al final del
  bloque, después de las tres HALFSIM.
- **TIGO**: PORTABILIDAD y REPOSICIÓN subieron junto a HALFSIM TIGO; los cuatro
  CHIP PREPAGO $3 quedaron contiguos.
- **AKIMOVIL**: las numeradas se ordenaron 7D → 15D → 30D.

Ningún artículo cambió de grupo. Los 41 artículos siguen siendo los mismos.

## 3. Divisor del promedio diario parametrizado

`=H13/20` → `=IFERROR(H13/DIAS_PERIODO,0)`, con `DIAS_PERIODO` = celda `I7`.

Se deja en **20**, así que ninguna cifra actual cambia. Se agregó:

- `I8`: conteo automático de días distintos con despacho en la hoja origen.
- `E9`: aviso cuando `I8 ≠ I7`.

Con los datos cargados hoy, `I8` devuelve **17** (02/10 a 22/10/2026) y el aviso
se activa. El `IFERROR` además evita el `#¡DIV/0!` si alguien deja la celda en
blanco.

## 4. Fila TOTAL GENERAL y control de cuadratura

Nueva fila 74 con el total de los 41 artículos, y debajo dos verificaciones:

| Verificación | Cuadro | Hoja origen | Dif. |
|---|---:|---:|---:|
| Stock total vs. `INVENTARIO` | 1.614.836 | 1.614.836 | 0 OK |
| Despacho total vs. `DESPACHOS 20 DIAS` | 332.431 | 332.431 | 0 OK |

Es el control que detecta el riesgo silencioso del diseño original: si llega un
artículo nuevo a las hojas origen y nadie lo da de alta en el cuadro, el
tablero lo ignora sin avisar. Ahora la diferencia deja de ser 0.

## 5. Subtotales que faltaban

Las bandas originales solo totalizaban `H`, `I`, `L`, `N` y `P`. Se agregaron
`F` (canal indirecto), `G` (canal directo), `J` (en proceso de liberación) y
`K` (en producción), más los días de stock `M`, `O` y `Q` a nivel de banda.

Los días de stock de la banda usan la misma fórmula de negocio que el detalle y
se colorean con el mismo semáforo cuando la familia es homogénea en `Tipo`. En
las bandas que mezclan Blanca y Numerada (TOTAL TIGO, TOTAL AKIMOVIL, TOTAL
MAXIPLUS, Subtotal ESIM TIGO DIGITAL) la cifra se muestra sin color, porque los
umbrales de los dos tipos son distintos y pintarla sería engañoso.

## 6. Hueco en el formato condicional

Las reglas para `Tipo = Numerada` eran `<8`, `>=8 y <15`, y `>15`: el valor
exacto **15** no caía en ninguna y quedaba sin color. Corregido a `>=15`.

El umbral de las SIM Blanca (`>150`) **no** se tocó.

## 7. Espacios no separables (NBSP)

`HALFSIM KITE M2M 4FF DEVICE CONTROL BLANCA` tenía caracteres NBSP (`U+00A0`)
en lugar de espacios, tanto en el cuadro como en `INVENTARIO!B46`. Funcionaba
de casualidad porque el error estaba en las dos hojas; cualquier recarga o
retipeo de una de ellas habría roto el `SUMIFS` en silencio y la fila habría
mostrado 0. Ambos quedaron normalizados.

## 8. Formatos numéricos

- Cantidades: `#,##0` (separador de miles, antes sin formato en las bandas).
- Promedio diario: `#,##0.00`. Antes era entero, y artículos con 0,05/día se
  leían como `0`.
- Días de stock: `#,##0`.

## 9. Columna A — `CLAVE`

Nueva columna de control (ancho 15,5, texto gris 7 pt) con la clave de familia
de cada artículo. Es la que alimenta todos los `SUMIFS` de las bandas. Puede
ocultarse sin afectar los cálculos.

---

## Verificación realizada

El archivo se recalculó con un motor de hojas de cálculo real (LibreOffice
Calc) y se comparó celda por celda contra el original:

- **0 errores de fórmula** (`#REF!`, `#N/A`, `#¡DIV/0!`, `#NAME?`) en la hoja.
- Las 41 filas de detalle devuelven **exactamente los mismos valores** que el
  archivo original.
- Las dos filas de control de cuadratura dan **0**.
- Hojas `PMP` y `RESUMEN DESPACHO`: contenido idéntico.
- Hojas `REFERENCIA` y `RESUMEN INVENTARIO T`: contenido idéntico (solo se
  recortaron filas vacías al final).
- Hoja `INVENTARIO`: una sola celda modificada, la del NBSP.
- Hoja `DESPACHOS 20 DIAS`: 122 celdas de la columna `VALORADO` cambian en el
  dígito 15.º o 16.º por la reescritura del formato XLSX (p. ej.
  `40,709999999999994` → `40,70999999999999`). La suma de la columna es
  **idéntica**: diferencia exactamente 0,0.
- Tablas dinámicas (3), imágenes (5) y hojas (7): conservadas.

---

## Pendientes recomendados (no aplicados)

Quedan fuera de este entregable porque cambian cifras publicadas o criterios de
negocio, y son decisión del área:

1. **Poner `DIAS_PERIODO` en 17.** El período real tiene 17 días con despacho,
   no 20. Con 20, la cobertura se sobreestima ~18 %.
2. **`">150"` y `">15"` son texto, no números.** Impide promediar, ordenar o
   graficar esas columnas, y el formato condicional las pinta de verde: las 25
   referencias sin ninguna rotación en el período aparecen como "saludables"
   cuando en realidad son stock inmovilizado (144.271 unidades).
3. **Vigilar el stock terminado, no el total.** `HALFSIM TIGO` tiene 511 días de
   stock total (verde) pero solo **4,8 días de producto terminado**; `ESIM TIGO
   DIGITAL`, 258 contra **16,4**. El riesgo no es de abastecimiento sino de
   capacidad de producción, y la columna que lo muestra es la `O`, no la `Q`.
4. **Canales sin mapear.** `INDIRECTO` y `RETAIL` aparecen en
   `DESPACHOS 20 DIAS` (466 filas, 34 %) pero no existen en `REFERENCIA`. Hoy
   no falla porque `CANAL 2` viene pegado a mano.
5. **`FECHA` y `PMP` son texto**, no fecha ni número. Impide filtrar por rango
   y hace poco confiable la columna `VALORADO`.
6. **Hojas residuales.** `RESUMEN INVENTARIO T`, `RESUMEN DESPACHO` y `PMP`
   pertenecen a otro catálogo (códigos SCL, productos Tuenti/Multisim) y nadie
   las referencia. Son ~40 % del peso del archivo y muestran totales de
   inventario que no reconcilian con el cuadro.
