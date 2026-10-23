# Indicadores de servicio — Activos Fijos

Tablero web que calcula los indicadores del contrato de Activos Fijos a partir de la
proforma mensual del proveedor. **El libro se procesa dentro del navegador**: no se sube
a ningún servidor, así que el master de 60 MB o más no tiene límite de carga.

👉 **https://brantos102.github.io/ARCH_OPTIMIZADOS/**

## Rutina mensual

El ciclo de reporte corre **del 24 de un mes al 23 del siguiente**, tal como lo nombra la
proforma (*«DEL 24 DE AGOSTO AL 23 DE SEPTIEMBRE»*).

1. Abre la página.
2. **Elegir archivos** o arrastra la proforma del mes. Acepta `.xlsx`, `.xlsb`, `.xlsm`,
   y varios archivos a la vez si prefieres cargarlos por proceso.
3. Espera. Un libro de 43 MB con 530.000 filas tarda unos 25 segundos.
4. Revisa la tabla **Pestañas leídas y qué aporta cada una**, al final de la página.
   Cada pestaña declara cuánto sumó a cada total. Si algo aparece aportando lo que no
   debe, el número de arriba está inflado y hay que avisar antes de reportar.
5. Comprueba que el ciclo del encabezado sea el que esperas.
6. **Descargar indicadores.json** y reemplaza `data/historico.json` en este repositorio.
   La página carga ese archivo al abrirse, así los ciclos anteriores quedan a la vista
   sin volver a cargar el libro.

## Qué calcula

Las ocho actividades del contrato, cada una con su **estado de medición**:

| Estado | Significado |
|---|---|
| Medible | El origen trae todos los campos que exige la fórmula del KPI |
| Parcial | Se verifica la entrega del insumo, no su oportunidad |
| No medible | Falta un campo sin el cual el KPI no puede calcularse |

Con los archivos de septiembre 2026: **0 medibles, 2 parciales, 6 no medibles**, y 10 de
13 campos críticos en 0%. La causa de fondo es que los archivos registran un solo
instante de tiempo y todos los SLA del contrato necesitan dos.

Cuando una actividad no es medible se muestra un **indicador de control**: una medición
auxiliar que el dato sí admite (unicidad de series, completitud de horas, cobertura de
ubicación). No sustituye al KPI contractual y la página lo dice.

**La detección es automática.** Si el proveedor agrega la columna que falta —por ejemplo
una fecha de despacho—, el motor la reconoce, la actividad pasa a «Medible» y empieza a
calcular el SLA real sin tocar una línea de código. Los patrones que busca están en
`assets/parser.worker.js`: `RE_DISPONIBLE`, `RE_DESPACHO`, `RE_FRANJA`, `RE_CONTEO`,
`RE_ETIQUETA`.

## Qué pestañas lee

Las reconoce por **contenido**, no por nombre, para que sigan entrando cuando el nombre
cambie de mes. Cada formato exige más de una columna: una firma de una sola columna
genérica dejaba entrar pestañas ajenas y sus cifras se sumaban a los totales en silencio.

| Firma en la cabecera | Se trata como |
|---|---|
| `doc_ext` + `FEC_CREAC` | Alistamiento (picking y despacho) |
| `SERIE` + `FECHA INICIAL` + 4 de las 6 del formulario | Recepción |
| `ALMACEN_DESC` + `SERIE` | Almacenaje — stock serializado |
| `DESCRIPCION` + `CANTIDAD`, sin más columnas | Almacenaje — cantidades |
| `Cantidad de Ubicación` | Almacenaje — ubicaciones |

La cabecera se busca en las primeras doce filas, así que los membretes no estorban.
En las hojas con bloques apilados o lado a lado (mes anterior y mes actual) toma el
**último** como cierre del ciclo y usa el anterior para ingresos, salidas y permanencias.

### Qué deja fuera, y por qué

La proforma trae 25 pestañas; el tablero lee las de los cuatro procesos de Activos Fijos.
Las demás (`LOGISTICA NACIONAL`, `VALOR AGREGADO`, `SEGURO`, `MATERIAL TIENDA`,
`L. INVERSA&TRANSPORTE`, `PEDIDOS CON TURNO`, `HORAS*`, `REPROCESOS`) son otras líneas de
facturación y se omiten a propósito. Cada omisión se avisa en pantalla.

Dos de las omitidas sí tocan el alcance y quedan pendientes: `RECEPCION ACCESORIO` y
`CONTROL DE CALIDAD` —el contrato incluye el control de calidad dentro del SLA de
recepción—. Para incorporarlas hace falta su estructura: pulsa **Descargar diagnóstico**,
que exporta los nombres de columna, el tipo y cuántas filas los traen poblados de *todas*
las pestañas, sin un solo valor. Son unos pocos KB y permite extender el lector sin mover
el libro.

## Publicación

Se publica con GitHub Pages desde la rama configurada en **Settings → Pages**. No hay
build ni dependencias: son archivos estáticos. SheetJS se carga desde jsDelivr, fijado a
la versión 0.18.5; es la única dependencia externa.

```
index.html                  página
assets/app.css              estilos, claro y oscuro
assets/app.js               interfaz, tablas y gráficos
assets/parser.worker.js     extracción y cálculo (Web Worker)
data/historico.json         ciclos publicados
```

Para revisar sin publicar: `python3 -m http.server 8000` en esta carpeta y abrir
http://localhost:8000. Abrir `index.html` con doble clic **no** funciona, porque los
navegadores bloquean los Web Workers bajo `file://`; la página lo detecta y lo avisa.

## Qué queda publicado

El repositorio es público, así que la página y `data/historico.json` los puede abrir
cualquiera. Ese JSON lleva volúmenes del ciclo, el reparto por cliente, los nombres de
las pestañas y el estado de medición de cada actividad. **No lleva** series, IMEI, ICC ni
precios: esos se quedan en el navegador de quien carga el libro y no salen de ahí.

Si en algún momento esas cifras no deben ser públicas, basta con quitar
`data/historico.json`: la página sigue funcionando, cada quien carga su archivo, y no se
publica ningún número.
