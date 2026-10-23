# Indicadores de servicio — Activos Fijos

Tablero web que calcula los indicadores del contrato de Activos Fijos a partir de los
libros de Excel del proveedor. **El archivo se procesa dentro del navegador**: no se
sube a ningún servidor, así que el master de 60 MB o más no tiene límite de carga.

## Qué resuelve

El master mensual no se podía subir a ninguna herramienta por su peso. Aquí no se sube:
el navegador lo lee localmente con [SheetJS](https://sheetjs.com), extrae solo las columnas
que alimentan los indicadores y descarta el resto. Lo único que sale del equipo es el
`indicadores.json` que tú decidas publicar, de unos 18 KB.

## Uso

1. Abre la página publicada (ver *Publicación* más abajo).
2. **Elegir archivos** o arrastra el master, o bien los archivos separados por proceso.
   Acepta `.xlsx`, `.xlsb`, `.xlsm` y varios archivos a la vez.
3. El ciclo se detecta solo. El contrato reporta **del 23 de un mes al 22 del siguiente**,
   y el selector de arriba permite cambiarlo si hay más de uno en los datos.
4. Para conservar el histórico, pulsa **Descargar indicadores.json** y reemplaza
   `data/historico.json` en este repositorio. La página carga ese archivo al abrirse,
   de modo que los ciclos anteriores quedan visibles sin volver a subir nada.

Un master de 43 MB con 532.000 filas tarda unos 25 segundos.

## Qué calcula

Las ocho actividades del contrato, cada una con su **estado de medición**:

| Estado | Significado |
|---|---|
| Medible | El origen trae todos los campos que exige la fórmula del KPI |
| Parcial | Se verifica la entrega del insumo, no su oportunidad |
| No medible | Falta un campo sin el cual el KPI no puede calcularse |

Cuando una actividad no es medible, el tablero muestra un **indicador de control**: una
medición auxiliar que el dato sí admite (unicidad de series, completitud de horas,
cobertura de ubicación). No sustituye al KPI contractual y la página lo dice.

**La detección es automática.** Si el proveedor agrega la columna que falta —por ejemplo
una fecha de despacho en el alistamiento—, el motor la reconoce, la fila pasa a «Medible»
y empieza a calcular el SLA real sin tocar una línea de código. Los nombres de columna que
busca están en `assets/parser.worker.js` (`RE_DISPONIBLE`, `RE_DESPACHO`, `RE_FRANJA`,
`RE_CONTEO`, `RE_ETIQUETA`).

## Cómo reconoce las pestañas

Por **contenido**, no por nombre, para que siga funcionando cuando el nombre cambie de mes:

| Marca en la cabecera | Se trata como |
|---|---|
| `doc_ext` | Alistamiento (picking y despacho) |
| `SERIE` + `FECHA INICIAL` | Recepción |
| `ALMACEN_DESC` | Almacenaje — stock serializado |
| `DESCRIPCION` + `CANTIDAD` | Almacenaje — cantidades |
| `Cantidad de Ubicación` | Almacenaje — ubicaciones |

Las pestañas que no coinciden con ningún formato se omiten y se listan como aviso.
La cabecera se busca en las primeras doce filas, así que los membretes no estorban.

En las hojas de almacenaje con dos bloques lado a lado (mes anterior y mes actual),
el motor toma el **último bloque** como cierre del ciclo y usa el anterior para calcular
ingresos, salidas y permanencias.

## Publicación en GitHub Pages

En **Settings → Pages** del repositorio, elige *Deploy from a branch*, rama `main`
y carpeta `/ (root)`. En un par de minutos queda en:

```
https://brantos102.github.io/ARCH_OPTIMIZADOS/
```

No hay build ni dependencias que instalar: son archivos estáticos.

## Estructura

```
index.html                  página
assets/app.css              estilos, claro y oscuro
assets/app.js               interfaz, tablas y gráficos
assets/parser.worker.js     extracción y cálculo (Web Worker)
data/historico.json         ciclos publicados
```

SheetJS se carga desde jsDelivr, fijado a la versión 0.18.5. Es la única dependencia
externa; el resto es código propio sin framework.

## Privacidad

Los libros de origen contienen series, IMEI, ICC y precios. Nada de eso sale del navegador:
el `indicadores.json` publicable solo lleva conteos, porcentajes y nombres de pestaña.
Aun así, revisa el JSON antes de subirlo si el repositorio es público.
