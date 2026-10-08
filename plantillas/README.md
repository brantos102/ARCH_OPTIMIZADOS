# Plantilla maestra de cotizaciones ITSA

Plantilla **en blanco** reconstruida a partir de `COT-2026-005_Impoventura_v1.pdf`,
respetando logo, retícula, colores, tipografía y pie de firma del Brandbook ITSA.
Es la base que la Gema debe usar para todo documento oficial que emita.

## Contenido

| Archivo | Para qué sirve |
|---|---|
| `cotizacion/plantilla_cotizacion.html` | **Fuente editable** de la maqueta. Referencia los logos y tipografías de `assets/`. Es el archivo que se toca para cambiar el diseño. |
| `cotizacion/PLANTILLA_COTIZACION_EN_BLANCO.html` | La misma maqueta en **un solo archivo** (logos y tipografías incrustados). Es la versión que se sube al Conocimiento de la Gema o se envía por correo. |
| `cotizacion/PLANTILLA_COTIZACION_EN_BLANCO.pdf` | Vista previa impresa de la plantilla vacía, para revisión visual. |
| `cotizacion/PLANTILLA_COTIZACION_EN_BLANCO.docx` | Misma plantilla en Word, para el equipo comercial y para llenado automático. |
| `cotizacion/ficha_datos.EN_BLANCO.json` | **Ficha de datos vacía**: el formulario que se llena por cliente. |
| `cotizacion/ficha_datos.EJEMPLO_COT-2026-005.json` | Ficha completa de ejemplo (reproduce la cotización original). |
| `cotizacion/ejemplo_salida/` | Resultado real de fusionar esa ficha con la plantilla (PDF y DOCX). |
| `herramientas/generar_cotizacion.py` | Motor de llenado: ficha JSON + plantilla → HTML, PDF y DOCX. |
| `herramientas/crear_plantilla_docx.py` | Reconstruye la plantilla Word desde cero si hay que cambiar el diseño. |
| `assets/img/` | Logo itsanet Ecuador, caja azul con logo blanco, logo FlexNet y patrón de flechas, extraídos del PDF original. |
| `assets/fonts/` | Schibsted Grotesk (400/500/600/700) en TTF y WOFF2. Los TTF se instalan en el equipo para que Word muestre la tipografía de marca. |

## Uso

```bash
# 1. Copiar la ficha vacía y llenarla con los datos del cliente
cp cotizacion/ficha_datos.EN_BLANCO.json mi_cotizacion.json

# 2. Generar los entregables (HTML + PDF + DOCX)
python3 herramientas/generar_cotizacion.py mi_cotizacion.json --salida ./salida
```

El script nombra los archivos según la convención `COT-{AAAA}-{NNN}_{Cliente}_v{N}`
y, al terminar, lista los campos que quedaron sin completar.

---

# Cómo se incrusta la información en cada reporte

La plantilla **no se edita nunca**: se copia y se rellenan sus marcadores. Hay
cuatro mecanismos, y entre ellos cubren todos los datos de una cotización.

## 1. Campos simples — `{{CAMPO}}`

Cada dato puntual es un marcador entre llaves dobles. El motor lo sustituye por
el valor correspondiente de la ficha JSON.

| Marcador | Origen en la ficha |
|---|---|
| `{{ETIQUETA_DOCUMENTO}}` | `documento.etiqueta` — la palabra del encabezado (COTIZACIÓN, PROPUESTA, INFORME) |
| `{{TITULO_DOCUMENTO}}` | `documento.titulo` |
| `{{CODIGO_DOCUMENTO}}` | `documento.codigo` → `COT-2026-005` |
| `{{VERSION}}` | `documento.version` → `v1` |
| `{{CIUDAD_EMISION}}` | `documento.ciudad` |
| `{{FECHA_EMISION_LARGA}}` | `documento.fecha_emision`, convertida a `7 de octubre de 2026` |
| `{{MONEDA}}` | `documento.moneda` |
| `{{VALIDEZ}}` | `documento.validez` |
| `{{CLIENTE}}` / `{{CLIENTE_MAYUSCULAS}}` | `cliente.nombre` |
| `{{ATENCION_A}}` | `cliente.atencion_a` |
| `{{CIUDADES_OPERACION}}` | `cliente.ciudades_operacion` (lista unida con comas) |
| `{{SERVICIO_COTIZADO}}` | `servicio_cotizado` |
| `{{GERENTE_COMERCIAL}}`, `{{FIRMANTE_*}}` | bloque `emisor` |
| `{{NOTA_FISCAL}}` | `totales.nota` |
| `{{PARRAFO_CIERRE}}` | `parrafo_cierre` |

## 2. Bloques repetibles — altas y bajas de servicios

Este es el mecanismo que permite **aumentar o reducir servicios** sin tocar el diseño.
En el HTML el bloque va entre marcadores de comentario; en el DOCX es la fila patrón
de la tabla:

```html
<!-- INICIO_ITEMS -->
<tr>
  <td>{{ITEM_SERVICIO}}</td><td>{{ITEM_CANTIDAD}}</td>
  <td>{{ITEM_TARIFA}}</td><td class="num">{{ITEM_TOTAL}}</td>
</tr>
<!-- FIN_ITEMS -->
```

El motor clona ese bloque **una vez por elemento del arreglo `items`**. Tres servicios
producen tres filas; ocho producen ocho; quitar uno de la ficha quita su fila. Lo mismo
ocurre con `INICIO_CONDICIONES / FIN_CONDICIONES` y el arreglo `condiciones`.

```json
"items": [
  { "servicio": "Recepción (19 al 23 Oct)", "cantidad": "5 días (40 horas)",
    "tarifa": "$25 / hora (L-V)", "total": 2000.00 }
]
```

Reglas al modificar servicios:

- **Agregar** un servicio = añadir un objeto al arreglo `items`.
- **Quitar** un servicio = eliminar su objeto.
- **Cambiar** cantidad o tarifa = editar sus campos y recalcular `total`.
- Cualquiera de las tres operaciones **obliga a subir la versión** (`v1` → `v2`),
  porque cambia el alcance o el precio de una oferta ya emitida.

## 3. Totales calculados

El motor no confía en un total escrito a mano: suma los `total` de los ítems.

- `totales.mostrar: false` → la fila de totales desaparece por completo.
- `totales.tasa_iva: 0` → muestra `Total estimado` y la nota *Valores no incluyen IVA*.
- `totales.tasa_iva: 15` → calcula el IVA, muestra `Total (IVA 15 % incluido)` y ajusta la nota.

## 4. Campos ausentes — `[PENDIENTE: …]`

Si un dato falta en la ficha, el motor **no lo inventa**: escribe
`[PENDIENTE: cliente.atencion_a]` en el documento y lo lista en consola al terminar.
Así ninguna cotización sale con un dato inventado, y se ve a simple vista qué falta.

## Flujo completo

```
Ficha JSON (datos del cliente)
        │
        ├─ campos simples      →  reemplazo directo
        ├─ items[]             →  una fila por servicio   (altas y bajas)
        ├─ condiciones[]       →  una viñeta por condición
        └─ totales             →  suma, IVA y nota fiscal
        │
        ▼
Plantilla en blanco (maqueta de marca, nunca se modifica)
        │
        ▼
COT-2026-005_Impoventura_v1  →  .pdf  ·  .docx  ·  .html
```

Para regenerar una cotización actualizada basta con editar la ficha JSON y volver a
ejecutar el motor: la maqueta, el logo y el pie de firma se mantienen idénticos.

## Nota sobre la tipografía

Schibsted Grotesk se incrusta en el HTML, así que el PDF siempre sale correcto.
Para que **Word** la muestre, hay que instalar los `.ttf` de `assets/fonts/` en el
equipo (clic derecho → Instalar). Sin instalarla, Word aplica el fallback Arial y el
documento sigue siendo legible, pero pierde fidelidad de marca.
