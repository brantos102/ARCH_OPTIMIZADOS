# Formato EGR_FL_HYCITE: análisis del estado actual y decisión de unificación

Versión analizada: `Formato EGR_FL_HYCITE.xlsm` recibido el 28/09/2026 (152 pedidos, 160 bultos cargados).
Complementa a [`ANALISIS_FLUJO_ACTUAL.md`](ANALISIS_FLUJO_ACTUAL.md), que describe la versión anterior.

## Contenido

1. [Decisión](#1-decisión)
2. [Inventario actual](#2-inventario-actual)
3. [Qué cambió desde el análisis anterior](#3-qué-cambió-desde-el-análisis-anterior)
4. [Hallazgos](#4-hallazgos)
5. [Por qué se cuelga Excel](#5-por-qué-se-cuelga-excel)
6. [Cómo se unifica sin romper nada](#6-cómo-se-unifica-sin-romper-nada)
7. [Funciones nuevas del archivo único](#7-funciones-nuevas-del-archivo-único)
8. [Decisiones pendientes](#8-decisiones-pendientes)

---

## 1. Decisión

**Sí se unifica, pero no copiando las hojas de EGR dentro de PEDIDOS HCE.**

| Opción | Resultado |
|---|---|
| Copiar las 16 hojas de EGR a PEDIDOS | ❌ Se descarta. Arrastra 64 000 fórmulas, 13 300 matriciales, 5 tablas dinámicas, el modelo de datos y 2 vínculos externos. Las consultas de Power Query y el modelo de datos **no se copian** con las hojas: se rompen. Es la vía más probable de que Excel se cuelgue o se cierre. |
| Seguir con dos archivos | ⚠️ Funciona, pero mantiene dos maestros de cobertura que ya no coinciden (ver H‑3), la copia manual entre libros y el límite de 254/500 filas. |
| **Archivo único construido sobre PEDIDOS HCE** | ✅ **Recomendada.** PEDIDOS HCE ya tiene el panel, los formularios, la cobertura con gestor/destino/trayecto, las zonas peligrosas y el historial. Las funciones de EGR se reconstruyen como **pasos del panel escritos en VBA que dejan valores**, no fórmulas vivas. |

Por qué es la opción óptima:

- **Rendimiento:** los datos se calculan una vez, al pulsar el paso, con diccionarios en memoria (segundos). Después las hojas solo muestran valores y Excel no recalcula miles de fórmulas al abrir, filtrar o desplazarse.
- **Un solo maestro de cobertura:** lo que se valida en la etapa 1 es exactamente lo que se usa para TMS, TRAMACO y etiquetas.
- **Sin límites fijos:** TMS (254), DATOS (500) y ETIQUETAS ZEBRA (800 bultos) hoy dependen de fórmulas precargadas. Con VBA el tamaño se ajusta a los pedidos del día.
- **Sin romper nada:** el EGR actual se sigue usando **en paralelo** hasta que el archivo único dé las mismas salidas. Una macro de comparación lo comprueba (fase 3, sección 6).

---

## 2. Inventario actual

Hay 16 hojas, unas 64 000 fórmulas (13 300 matriciales), 5 tablas dinámicas, 3 consultas, el modelo de datos y 2 vínculos externos.

| Hoja | Visible | Qué hace | Fórmulas | Estado |
|---|---|---|---|---|
| PANEL | Sí | Nueva. Indicadores del día (pedidos, bultos, peso, REVISAR, destinos) y los 5 pasos de producción | 9 | ✅ Útil; "Sin courier asignado" marca 151 porque DATOS!Y está roto (H‑1) |
| MENU | Oculta | Imágenes y 9 formas de navegación **sin macro asignada** | 0 | ⚪ Obsoleta |
| COBERTURAS Y TARIFAS | Oculta | Maestro de cobertura (1 297 filas). Tiene gestor (F), gestor sugerido TRAMACO (U), trayecto (V), código postal (X), clave `PROV_CANTON_PARROQUIA` (AB) y la tabla TARIFAS (AD:AG) | 2 594 | ⚠️ Difiere del maestro de PEDIDOS (H‑3). TARIFAS está vacía (H‑4) |
| DATOS | Sí | Hoja central: 499 filas precargadas con fórmulas en A, B y O:AE | ~9 500 | ⚠️ Ver H‑1, H‑2, H‑4, H‑5 |
| TMS | Oculta | Formato de carga al TMS, 42 columnas, filas 2‑255 | ~8 600 | ⚠️ Máximo 254 pedidos. Valor fijo 4,70 cuando la tarifa es 0 |
| TRAMACO | Oculta | Formato del portal TRAMACO. Solo pedidos PRO, vía DATOS!AC (contador N_PRO) | ~7 500 (la mayoría matriciales) | ✅ Ya no depende de Temp_Tramaco |
| DESPACHOS | Oculta | Reporte: unidades, ítems, valor, box density, transporte | ~8 000 | ⚠️ `CHOOSE({2,1},DATOS!A:A,DATOS!B:B)` sobre columnas completas, en 500 filas (H‑6) |
| ETIQUETAS | Sí | Lista de etiquetas (nombre y apellido partidos por espacios) | ~5 500 | ⚪ Duplica a ETIQUETAS ZEBRA |
| ETIQUETAS ZEBRA | Sí | Una fila por bulto (800 filas) + plantilla P3:U15 controlada por Q1 | ~9 600 matriciales | ⚠️ Impresión manual, etiqueta por etiqueta |
| ITEMS APIS | Sí | Consulta ODBC `ITEMS API`: líneas solicitadas del día + peso por SKU | 1 174 | ✅ |
| ITEMS DEPOT | Sí | Consulta ODBC `Estado`: picking/empaque confirmado + peso, volumen y precio | 2 556 | ⚠️ Duplica cantidades por contenedora (H‑7) |
| EMPAQUETADO | Sí | Consulta de Google Sheets publicado: caja por orden + cálculo de box density | 870 | ⚠️ `COUNTIF(#REF!...)` en F (H‑8); 1 bulto por fila fijo |
| TABLAS DINAMICAS | Sí | 4 tablas dinámicas (unidades/peso por pedido, empaque, volumen, box density) | 545 | ⚠️ 2 cachés leen `C1:K1048576` completo (H‑9) |
| Temp_Tramaco | Oculta | Vacía: "HOJA OBSOLETA" | 0 | ⚪ Eliminar |
| COD POS | Oculta | Códigos postales TMS; 2 606 `XLOOKUP` a un libro externo en una carpeta de red | 5 212 | 🔴 Vínculo externo (H‑10) |
| DATA CODIGO Y CAJAS | Oculta | Maestro de SKU (peso, volumen, precio) y de cajas | ~600 | ⚠️ 185 reglas de formato condicional repetidas (H‑11) |

### Consultas (sin cambios desde el análisis anterior)

| Consulta | Origen | Carga |
|---|---|---|
| `Estado` | ODBC `DEPOTUIO` (SQL Server), desde ayer 00:00 | Tabla `Estado` (ITEMS DEPOT) |
| `ITEMS API` | ODBC `DEPOTUIO`, solo hoy | Tabla `ITEMS_API` |
| `EMPAQUETADO` | CSV de Google Sheets publicado (URL pública) | Tabla EMPAQUETADO **y modelo de datos** |

### VBA (sin cambios)

Solo el Módulo7 tiene código:

- `TABLAS`: refresca 4 tablas dinámicas seleccionando hojas.
- `TABLA_APIS`: refresca `TablaDinámica3`, **que ya no existe** en ITEMS APIS. Ahí hay una tabla dinámica llamada `TablaDinámica1`, así que la macro da error.
- `ActualizarTodo`: botón "ACTUALIZAR" en EMPAQUETADO. No tiene manejo de errores: si falla el ODBC, ScreenUpdating queda apagado.

Los módulos 1‑6 y `UserForm1` están vacíos.

---

## 3. Qué cambió desde el análisis anterior

| Cambio | Efecto |
|---|---|
| Nueva hoja **PANEL** con indicadores y pasos | Bueno. En el archivo único pasa al formulario frmPanel (etapa 2) |
| COBERTURAS Y TARIFAS **oculta** y reordenada: gestor en F, sugerido en U, trayecto en V, CP en X, clave en AB, TARIFAS en AD:AG | La reorganización dejó **`#REF!` en DATOS!Y (COURIER)** |
| DATOS gana **V VALIDACION** (OK/REVISAR), **AD DIAGNOSTICO** y **AE SUGERENCIA** | Buena idea, pero duplica lo que ya hace PEDIDOS HCE con otro maestro |
| TRAMACO ya no usa Temp_Tramaco (usa DATOS!AC) | Correcto; Temp_Tramaco puede eliminarse |
| Se eliminó la hoja "Claude Log" | — |
| Consultas, tablas dinámicas y VBA | Sin cambios |

---

## 4. Hallazgos

### 🔴 Críticos (afectan lo que sale al TMS o al courier)

| # | Hallazgo | Evidencia | Corrección en el archivo único |
|---|---|---|---|
| H‑1 | **COURIER (DATOS!Y) roto**: `INDEX(#REF!, …)` dentro de `IFERROR` devuelve vacío sin avisar | 152 de 152 pedidos sin courier; PANEL muestra "Sin courier asignado: 151"; ETIQUETAS ZEBRA imprime courier vacío | El courier sale de la etapa 1 (`GESTOR_ASIGNADO`, DEPOT!V): ITSANET/LAAR en UIO/GYE, TRAMACO en PRO |
| H‑2 | **DESTINO (DATOS!A) por provincia**, no por cobertura: `PICHINCHA→UIO, GUAYAS→GYE, GALAPAGOS→GPS, resto→PRO` | Un pedido de Pichincha fuera del DMQ (p. ej. Cayambe) sale como UIO; uno de Guayas fuera de Guayaquil, como GYE | Usar `DESTINO` de la etapa 1 (DEPOT!W), que ya aplica la regla acordada |
| H‑3 | **Dos maestros de cobertura distintos** | 1 288 claves comunes; 6 solo en EGR y 13 solo en PEDIDOS (p. ej. `PUEBLO VIEJO` / `PUEBLOVIEJO`, `TONCHIGÜE` / `TONCHIGUE`) → un pedido **OK en PEDIDOS puede quedar VALIDAR en EGR** | Un solo maestro (COBERTURA de PEDIDOS) + la columna de código postal de EGR (X) |
| H‑4 | **TARIFAS vacía** → VALOR (DATOS!T) = 0 en todos los pedidos → TMS "VALOR TOTAL FACTURA" = 4,70 fijo | 152 de 152 con VALOR 0 | Definir qué va en ese campo (sección 8) |
| H‑5 | **Pedido sin empaquetar = 1 bulto** (`IFERROR(…,1)` en DATOS!R) | El pedido sale con 1 etiqueta aunque no esté empacado | Bloquear o marcar PENDIENTE y excluir de TMS y etiquetas |

### 🟠 Errores de fórmulas y de código

| # | Hallazgo |
|---|---|
| H‑6 | DESPACHOS!P usa `VLOOKUP(…, CHOOSE({2,1}, DATOS!$A:$A, DATOS!$B:$B), 2, 0)`: arma una matriz de 2 × 1 048 576 celdas **en cada una de las 500 filas** |
| H‑7 | SQL `Estado`: la unión con `VIEW_TIEMPO_EMPAQUETADO` repite las cantidades por cada contenedora (pedidos de 2+ cajas muestran unidades duplicadas). Ya reportado; sigue igual |
| H‑8 | EMPAQUETADO!F: `COUNTIF(#REF!, …)`; cuando la orden no está en `Estado` devuelve error y no "validar" |
| H‑9 | 2 cachés de tablas dinámicas leen `EMPAQUETADO!C1:K1048576` (columnas completas) en vez de la tabla |
| H‑10 | **COD POS**: 2 606 `XLOOKUP` a `CODIGOS POSTALES TMS ECU.xlsx` en una carpeta de red de otro usuario. Además hay un vínculo del libro **a sí mismo** (`[1]` = otra copia de Formato EGR en `FIN_28092026`), usado por los nombres `DATA_CAJAS` / `DATA_CODIGOS` de la hoja COBERTURAS |
| H‑11 | 185 reglas de formato condicional "valores duplicados" repetidas en DATA CODIGO Y CAJAS, varias sobre `B1:B1048576`; otras de columna completa en DATOS (U, O:Q, A, T) y COBERTURAS (D:E) |
| H‑12 | Nombres definidos rotos: `CIUDAD`, `CIUDAD_PARROQ`, `CIUDAD_RUTA`, `COD_POSTAL`, `PROVINCIA`, `RUTAS`, `PROV_CIUD`, `TARIFAS` (local) y `EGRESOS_CORRCAL` apuntan a `#REF!` |
| H‑13 | `TABLA_APIS` refresca una tabla dinámica que no existe; `ActualizarTodo` no restaura la pantalla si falla |
| H‑14 | Teléfonos guardados como número (`969164115`). TMS antepone "0" por fórmula; TRAMACO y ETIQUETAS ZEBRA **salen sin el 0** |
| H‑15 | `TODAY()` en TMS (B, Z) y DESPACHOS (A): si el archivo se abre al día siguiente, cambian las fechas de un despacho ya hecho |

### ⚪ Obsoleto

MENU (formas sin macro), Temp_Tramaco, ETIQUETAS (duplica ZEBRA), Módulos 1‑6 y UserForm1 vacíos, y nombres `#REF!`.

---

## 5. Por qué se cuelga Excel

No es la cantidad de hojas, sino **lo que se recalcula y se redibuja**.

| Causa | Tamaño | Cuándo pesa |
|---|---|---|
| Fórmulas precargadas en filas vacías (DATOS 500, TMS 255, TRAMACO 500, DESPACHOS 500, ETIQUETAS 500, ZEBRA 800) | ~64 000 fórmulas; cadena de cálculo de 1,7 MB | Cada cambio en DATOS o cada refresco de consulta |
| Fórmulas matriciales (CSE) en TRAMACO y ETIQUETAS ZEBRA | 13 300 | Igual; no se optimizan como las normales |
| `CHOOSE({2,1}, columnas completas)` en DESPACHOS | 500 × 2 M de celdas | Cualquier cambio en DATOS |
| `VLOOKUP`/`SUMIF` a columnas completas de TABLAS DINAMICAS | ~4 100 fórmulas | Al refrescar las tablas dinámicas |
| Formato condicional "duplicados" sobre columnas completas | 185 + 17 reglas | **Al redibujar**: desplazarse, filtrar, cambiar de hoja |
| Vínculos externos (carpeta de red + auto‑vínculo) | 2 606 `XLOOKUP` | Al abrir: Excel intenta ir a la red; si no responde, queda "No responde" |
| Modelo de datos + refresco en segundo plano | 1 consulta al modelo | Al refrescar: `ActualizarTodo` fuerza modo síncrono sin control de errores |

En el archivo único esto desaparece:

- no hay fórmulas precargadas, porque los pasos escriben valores;
- no hay tablas dinámicas, porque las sumas se hacen con diccionarios;
- no hay vínculos externos;
- no hay modelo de datos;
- no hay formato condicional de columna completa;
- los procesos corren con `ScreenUpdating`/`Calculation` controlados y **siempre restaurados**, como ya hacen los pasos de la etapa 1 (con el botón "Desbloquear" como respaldo).

---

## 6. Cómo se unifica sin romper nada

### Arquitectura del archivo único (base: PEDIDOS HCE)

```text
ETAPA 1 (ya hecha)          ETAPA 2 (nueva, desde el mismo panel)
DEPOT ──validar──► DEPOT V:AD ──► DESPACHO (ex DATOS, solo valores)
COBERTURA (+ CP)                    ▲   ▲
ZONAS PELIGROSAS         ITEMS_API ─┘   └─ ESTADO + EMPAQUETADO  (consultas, solo tabla)
CAMBIOS / LOG                          │
                     ┌─────────────────┼──────────────────┬───────────────┐
                     ▼                 ▼                  ▼               ▼
                   TMS             TRAMACO          ETIQUETAS ZEBRA   DESPACHOS (reporte)
                (valores)         (valores)       (una fila por bulto + impresión masiva)
```

### Fases

| Fase | Qué se hace | Riesgo para lo que hoy funciona |
|---|---|---|
| **0. Alivio inmediato del EGR actual** (opcional, mientras se construye) | Macro `RepararEGR` que se ejecuta una vez sobre una copia: corrige DATOS!Y y DESPACHOS!P, rompe el vínculo de COD POS (pega valores), quita el auto‑vínculo, borra nombres `#REF!` y el formato condicional de columna completa, corrige `TABLA_APIS` y el manejo de errores de `ActualizarTodo` | Bajo: no cambia la estructura ni las salidas |
| **1. Datos de EGR en PEDIDOS** | Hojas ocultas `SKU` y `CAJAS` (de DATA CODIGO Y CAJAS), código postal en COBERTURA y las 3 consultas (`Estado`, `ITEMS API`, `EMPAQUETADO`) cargadas solo a tabla, sin modelo de datos | Ninguno: la etapa 1 no las usa |
| **2. Módulo `modEgreso` + pasos 8‑13 en el panel** | Genera DESPACHO, TMS, TRAMACO, ETIQUETAS ZEBRA y DESPACHOS con los mismos formatos de columna que hoy | Ninguno: el EGR actual sigue en uso |
| **3. Marcha en paralelo** (3‑5 días) | Macro `CompararConEGR`: abre el EGR del día y compara TMS, TRAMACO y bultos/peso celda por celda. Las diferencias van al registro del panel | Ninguno |
| **4. Cambio** | Cuando 3 días seguidos salen iguales (o las diferencias son correcciones conocidas: courier, destino, cobertura), se deja de usar el EGR | — |

### Reglas de construcción (para que no se cuelgue)

- Cada paso lee las hojas **una vez** a matrices en memoria, calcula con `Scripting.Dictionary` y escribe **un solo bloque** de valores.
- `Application.Calculation = xlCalculationManual`, `ScreenUpdating = False`, `EnableEvents = False` durante el paso. Un `On Error` **siempre** los restaura y anota el error en el registro del panel.
- Las consultas se refrescan una por una con `BackgroundQuery = False` y con tiempo medido. Si una falla, se informa cuál es (ODBC `DEPOTUIO`, Google Sheets…) y el resto sigue.
- Sin tablas dinámicas, sin fórmulas matriciales, sin vínculos externos, sin modelo de datos.
- Formato condicional solo en el rango usado y creado por macro.
- Teléfonos como **texto** con el 0 inicial.
- Fecha del despacho **fija** al generar, no `TODAY()`.

---

## 7. Funciones nuevas del archivo único

| Paso del panel | Función |
|---|---|
| 8 Actualizar picking y empaque | Refresca `Estado`, `ITEMS API` y `EMPAQUETADO` con control de errores y avance en el registro |
| 9 Validar picking | Por pedido y SKU: solicitado (ITEMS API) vs. confirmado (Estado). Marca FALTANTE / SOBRANTE / SKU NO SOLICITADO / SIN PICKING |
| 10 Validar empaque | Pedidos sin caja → PENDIENTE (no pasan a TMS ni a etiquetas); cajas sin pedido; box density fuera de rango; peso sin SKU en el maestro |
| 11 Generar despacho | DESPACHO + TMS + TRAMACO + DESPACHOS con destino, courier y cobertura de la etapa 1, sin límite de filas |
| 12 Etiquetas | Vista en el panel; imprimir todas, un pedido o un rango de bultos en la Zebra |
| 13 Exportar y cerrar el día | Guarda TMS y TRAMACO como archivos listos para subir, archiva el día en el historial (CAMBIOS + despacho) y deja el libro limpio para mañana |

El panel también suma:

- vistas DESPACHO, PICKING y EMPAQUE, con los mismos filtros y señales de la etapa 1;
- indicadores de la hoja PANEL actual (pedidos, bultos, peso, REVISAR, destinos, pendientes de empaque).

---

## 8. Decisiones pendientes

Hay que resolverlas antes de la fase 2. Entre paréntesis, la opción recomendada.

1. **VALOR del TMS** ("VALOR TOTAL FACTURA"): ¿4,70 fijo, flete por tarifa (llenar TARIFAS) o valor del producto (DATA CODIGO: precio × cantidad)? *(Valor del producto: ya se calcula en ITEMS DEPOT!K)*
2. **Pedido sin empaquetar:** ¿bloquearlo o salir con 1 bulto como hoy? *(Bloquear y marcar PENDIENTE)*
3. **Hoja ETIQUETAS (lista):** ¿se usa, o solo ETIQUETAS ZEBRA? *(Solo ZEBRA)*
4. **Impresora de etiquetas:** modelo y tamaño (¿Zebra 10 × 15 cm?), para la impresión masiva.
5. **Origen del empaquetado:** ¿seguir con el Google Sheets publicado o leer `VIEW_TIEMPO_EMPAQUETADO` de DEPOT por ODBC? *(DEPOT: evita una URL pública con datos de clientes)*
6. **Corrección del SQL `Estado`** (cantidades duplicadas por contenedora): validarla con quien administra DEPOT.
7. **Fase 0:** ¿se quiere la macro de alivio para el EGR actual mientras se construye el archivo único? *(Sí, si hoy se cuelga en producción)*
