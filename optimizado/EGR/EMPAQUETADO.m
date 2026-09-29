// Consulta EMPAQUETADO (Power Query) - trae SOLO las filas de HOY que tienen número de orden.
// Reemplaza TODO el texto de la consulta en: Datos > Consultas y conexiones > clic derecho EMPAQUETADO >
// Editar > Inicio > Editor avanzado. Cambia PEGA_AQUI_LA_URL por la URL que ya tiene la consulta actual
// (la línea "Origen = Csv.Document(Web.Contents(...". Cópiala antes de borrar).
//
// Por qué: en el Google Sheets los operadores llenan la FECHA de hoy en miles de filas vacías. La consulta
// anterior filtraba por fecha pero dejaba pasar esas filas (37 109 filas en la tabla). Ahora se descartan
// primero las filas sin "# ORDEN" y después se filtra la fecha de hoy.
let
    Origen = Csv.Document(Web.Contents("PEGA_AQUI_LA_URL"), [Delimiter = ",", Columns = 26, Encoding = 65001, QuoteStyle = QuoteStyle.None]),
    Encabezados = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),
    Columnas = Table.SelectColumns(Encabezados, {"FECHA", "# ORDEN", "TIIPO CAJA", "# CONTENEDORA"}, MissingField.UseNull),
    // 1) solo filas con número de orden (descarta las miles de filas que solo tienen fecha)
    ConOrden = Table.SelectRows(Columnas, each [#"# ORDEN"] <> null and Text.Trim(Text.From([#"# ORDEN"])) <> ""),
    // 2) fecha: primero formato Ecuador (dd/mm/aaaa); si no, formato EE. UU.; si no se entiende, queda vacía
    Fechas = Table.TransformColumns(ConOrden, {{"FECHA", each try Date.From(_, "es-EC") otherwise try Date.From(_, "en-US") otherwise null, type date}}),
    // 3) solo HOY
    Hoy = Table.SelectRows(Fechas, each [FECHA] = Date.From(DateTime.LocalNow())),
    // 4) tipos: la orden como texto sin espacios ni decimales; la contenedora como número entero
    Orden = Table.TransformColumns(Hoy, {{"# ORDEN", each Text.BeforeDelimiter(Text.Trim(Text.From(_)), "."), type text}}),
    Contenedora = Table.TransformColumns(Orden, {{"# CONTENEDORA", each try Int64.From(_) otherwise null, Int64.Type}}),
    Caja = Table.TransformColumns(Contenedora, {{"TIIPO CAJA", each if _ = null then null else Text.Trim(Text.From(_)), type text}}),
    // 5) seguridad: nunca más de 2 000 cajas en un día (un día normal tiene hasta ~300 pedidos)
    Limite = Table.FirstN(Caja, 2000)
in
    Limite
