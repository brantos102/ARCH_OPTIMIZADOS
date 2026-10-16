# Fuentes originales (línea base)

Código extraído, sin modificar, de los dos libros tal como se recibieron el 25/09/2026. Los datos que contienen
corresponden al proceso del 24/09/2026. Sirve como referencia y para comparar con el archivo único que se construirá.

| Carpeta | Contenido |
|---|---|
| `PEDIDOS_HCE/vba/` | `Modulo1.bas` (813 líneas), `frmValidar.frm` (código del formulario), `ThisWorkbook.cls` |
| `PEDIDOS_HCE/powerquery/Section1.m` | Consultas Power Query `Depot` y `TMS` |
| `PEDIDOS_HCE/sql/` | SQL nativo de cada consulta ODBC, con los saltos de línea decodificados |
| `PEDIDOS_HCE/formulas.md` | Inventario de fórmulas por hoja y columna |
| `Formato_EGR_FL_HYCITE/vba/Modulo7.bas` | Única macro con código (`TABLAS`, `TABLA_APIS`, `ActualizarTodo`); el resto de módulos está vacío |
| `Formato_EGR_FL_HYCITE/powerquery/Section1.m` | Consultas `EMPAQUETADO`, `Estado`, `ITEMS API` |
| `Formato_EGR_FL_HYCITE/sql/` | SQL nativo de `Estado` e `ITEMS API` |
| `Formato_EGR_FL_HYCITE/formulas.md` | Inventario de fórmulas de las 17 hojas |

Notas:

- Los archivos de VBA están en UTF‑8 para leerlos en GitHub. Son una extracción de texto: para importarlos en el editor
  de VBA hay que guardarlos en ANSI (Windows‑1252) y, en `.frm`, el formulario necesita su `.frx`.
- En la copia de `Formato_EGR_FL_HYCITE/powerquery/Section1.m` se **ocultó** la URL publicada de Google Sheets
  (`<ID_PUBLICADO_OCULTO>`), porque da acceso de lectura a cualquiera que la tenga.
- Los `.xlsm` **no** se versionan: contienen datos personales de clientes (nombres, direcciones, teléfonos y correos).

Análisis completo: [`docs/ANALISIS_FLUJO_ACTUAL.md`](../docs/ANALISIS_FLUJO_ACTUAL.md).
