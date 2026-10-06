// Consulta EMPAQUETADO (Power Query) - trae SOLO las cajas de HOY, ya ordenadas.
// Reemplaza TODO el texto de la consulta en: Datos > Consultas y conexiones > clic derecho EMPAQUETADO >
// Editar > Inicio > Editor avanzado. Cambia PEGA_AQUI_LA_URL por la URL que ya tiene la consulta actual
// (la línea "Origen = Csv.Document(Web.Contents(...". Cópiala antes de borrar).
//
// Qué resuelve:
//  1. En el Google Sheets los operadores arrastran la FECHA de hoy sobre miles de filas vacías. La consulta
//     original las dejaba pasar (36 000 filas en la tabla). Aquí se descartan ANTES de cualquier otra cosa.
//  2. Un "# ORDEN" tecleado por error (p. ej. 5465443, de 7 dígitos) entraba a la hoja y salía como
//     "sin pedido" / "Verificar". Ahora solo pasan los números de pedido de 8 dígitos o más.
//  3. La hoja salía desordenada y el primer pedido quedaba al final. Ahora se ordena por FECHA y # ORDEN,
//     así se presenta y se revisa en orden.
//  4. Solo HOY, igual que la consulta Estado (ITEMS DEPOT). Las dos tienen que mirar el mismo día: si una
//     trae ayer y la otra no, las contenedoras de ayer quedan sin items y VOL. ITEMS sale "Verificar".
let
    Origen = Csv.Document(Web.Contents("PEGA_AQUI_LA_URL"), [Delimiter = ",", Columns = 26, Encoding = 65001, QuoteStyle = QuoteStyle.None]),
    Encabezados = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),
    Columnas = Table.SelectColumns(Encabezados, {"FECHA", "# ORDEN", "TIIPO CAJA", "# CONTENEDORA"}, MissingField.UseNull),
    // 1) la orden como texto limpio, sin espacios ni decimales
    Orden = Table.TransformColumns(Columnas, {{"# ORDEN", each if _ = null then "" else Text.BeforeDelimiter(Text.Trim(Text.From(_)), "."), type text}}),
    // 2) solo filas con un número de pedido de verdad: 8 dígitos o más, sin letras
    ConOrden = Table.SelectRows(Orden, each Text.Length([#"# ORDEN"]) >= 8 and Text.Select([#"# ORDEN"], {"0".."9"}) = [#"# ORDEN"]),
    // 3) fecha: primero formato Ecuador (dd/mm/aaaa); si no, formato EE. UU.; si no se entiende, queda vacía
    Fechas = Table.TransformColumns(ConOrden, {{"FECHA", each try Date.From(_, "es-EC") otherwise try Date.From(_, "en-US") otherwise null, type date}}),
    // 4) solo HOY
    Hoy = Table.SelectRows(Fechas, each [FECHA] = Date.From(DateTime.LocalNow())),
    // 5) tipos: la contenedora como número entero, el tipo de caja como texto limpio
    Contenedora = Table.TransformColumns(Hoy, {{"# CONTENEDORA", each try Int64.From(_) otherwise null, Int64.Type}}),
    Caja = Table.TransformColumns(Contenedora, {{"TIIPO CAJA", each if _ = null then null else Text.Trim(Text.From(_)), type text}}),
    // 6) ordenado: por fecha, por pedido y por contenedora
    Ordenado = Table.Sort(Caja, {{"FECHA", Order.Ascending}, {"# ORDEN", Order.Ascending}, {"# CONTENEDORA", Order.Ascending}}),
    // 7) seguridad: nunca más de 2 000 cajas en un día (un día normal tiene hasta ~300 pedidos)
    Limite = Table.FirstN(Ordenado, 2000)
in
    Limite
