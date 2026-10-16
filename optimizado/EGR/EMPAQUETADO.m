// Consulta EMPAQUETADO (Power Query) - trae SOLO las cajas de HOY, ya ordenadas.
// Reemplaza TODO el texto de la consulta en: Datos > Consultas y conexiones > clic derecho EMPAQUETADO >
// Editar > Inicio > Editor avanzado. Cambia PEGA_AQUI_LA_URL por la URL que ya tiene la consulta actual
// (la línea "Origen = Csv.Document(Web.Contents(...". Cópiala antes de borrar).
//
// Qué resuelve, comparado con la consulta que está en producción:
//
//  1. PEDIDOS CON SUFIJO (102384109_A). La consulta actual hace, en "Tipo cambiado":
//         {"# ORDEN", Int64.Type}
//     "102384109_A" no es un número entero, así que esa celda se convierte en ERROR. El paso
//     siguiente la pasa a texto, pero ya es tarde: el error no se deshace. Por eso la fila está
//     en el Google Sheets y no llega a la hoja.
//     Aquí el # ORDEN se trata como TEXTO DE PRINCIPIO A FIN y nunca se convierte a número.
//     REGLA: no volver a poner un tipo numérico en "# ORDEN". Es lo que rompe los sufijos.
//
//  2. FILAS VACÍAS. En el Sheets los operadores arrastran la FECHA de hoy sobre miles de filas
//     sin pedido. La consulta actual las deja pasar (36 000 filas en la tabla). Aquí se
//     descartan por el # ORDEN antes que nada.
//
//  3. FILAS QUE NO SON PEDIDOS (un nombre, "F11", una celda con un comentario). La regla mira
//     el número BASE, lo que va antes del "_": solo dígitos y 6 o más. Entra todo lo que sea
//     un pedido, con o sin sufijo, y se quedan fuera los textos.
//
//  4. FECHA. La consulta actual usa {"FECHA", type date}, que depende del idioma del libro, y
//     si falla deja un error en la celda. Aquí se prueba dd/mm/aaaa, luego el formato de EE. UU.,
//     y si no se entiende queda vacía, sin romper la fila.
//
//  5. FECHA: DE NORMAL, SOLO HOY. El -0 de abajo es "desde hoy", igual que antes. La diferencia
//     con la consulta vieja es que ese número lo reescribe el botón "Días" del panel: tras un
//     feriado se pone 3 o 7 y entran también las cajas de los días pendientes, sin editar nada.
//     El botón cambia el número en las TRES consultas a la vez (esta, Estado e ITEMS API), que
//     es como tienen que ir: si esta mirara más días que Estado, esas contenedoras se quedarían
//     sin items y VOL. ITEMS saldría "Verificar".
//
//  6. ORDEN Y TOPE. La hoja salía desordenada y el primer pedido quedaba al final. Ahora sale el
//     día más reciente arriba y, dentro de cada día, por pedido y contenedora, con tope de cajas.
//
// Las 4 columnas que entrega (FECHA, # ORDEN, TIIPO CAJA, # CONTENEDORA) y su orden son los
// mismos que ahora, así que las columnas calculadas F:K de la hoja no se tocan.
// Lo que se quita respecto a la consulta actual son las conversiones de tipo de las columnas que
// de todas formas se eliminan (VOLUMEN M3, PESO, Largo, Ancho, Alto...): convertirlas a Int64
// para borrarlas después solo gasta tiempo y puede generar errores.
let
    Origen = Csv.Document(Web.Contents("PEGA_AQUI_LA_URL"), [Delimiter = ",", Columns = 26, Encoding = 65001, QuoteStyle = QuoteStyle.None]),
    Encabezados = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),
    Columnas = Table.SelectColumns(Encabezados, {"FECHA", "# ORDEN", "TIIPO CAJA", "# CONTENEDORA"}, MissingField.UseNull),
    // 1) la orden como texto limpio, sin espacios ni decimales
    Orden = Table.TransformColumns(Columnas, {{"# ORDEN", each if _ = null then "" else Text.BeforeDelimiter(Text.Trim(Text.From(_)), "."), type text}}),
    // 2) solo filas con un número de pedido. La regla es la mínima que no pierde ninguno:
    //    el número BASE (lo que va antes del "_") tiene que ser SOLO DÍGITOS y de 6 o más.
    //    Después del "_" puede ir cualquier cosa (_A, _B, _AB...).
    //       102384109, 102384109_A, 9824569_A, 123456  -> entran
    //       vacío, "Adriana Morillo", "F11", "12345"   -> no entran
    //    Se prefiere dejar entrar algo de más antes que perder un pedido: un número corto
    //    tecleado por error (5465443) entra y se ve en la hoja como "sin pedido", que es
    //    visible y se corrige; un pedido perdido no se ve y no se despacha.
    EsPedido = (t as text) as logical =>
        let
            base = if Text.Contains(t, "_") then Text.BeforeDelimiter(t, "_") else t
        in
            Text.Length(base) >= 6 and Text.Select(base, {"0".."9"}) = base,
    ConOrden = Table.SelectRows(Orden, each EsPedido([#"# ORDEN"])),
    // 3) fecha: primero formato Ecuador (dd/mm/aaaa); si no, formato EE. UU.; si no se entiende, queda vacía
    Fechas = Table.TransformColumns(ConOrden, {{"FECHA", each try Date.From(_, "es-EC") otherwise try Date.From(_, "en-US") otherwise null, type date}}),
    // 4) fecha: -0 = solo hoy. El botón "Días" del panel reescribe ese número cuando tras un
    //    feriado hay que ver también los días anteriores. No hay que editar esta consulta.
    Hoy = Table.SelectRows(Fechas, each [FECHA] >= Date.AddDays(Date.From(DateTime.LocalNow()), -0)),
    // 5) tipos: la contenedora como número entero, el tipo de caja como texto limpio
    Contenedora = Table.TransformColumns(Hoy, {{"# CONTENEDORA", each try Int64.From(_) otherwise null, Int64.Type}}),
    Caja = Table.TransformColumns(Contenedora, {{"TIIPO CAJA", each if _ = null then null else Text.Trim(Text.From(_)), type text}}),
    // 6) ordenado: el día más reciente arriba y, dentro de cada día, por pedido y contenedora
    Ordenado = Table.Sort(Caja, {{"FECHA", Order.Descending}, {"# ORDEN", Order.Ascending}, {"# CONTENEDORA", Order.Ascending}}),
    // 7) seguridad: tope de cajas (un día normal tiene hasta ~300 pedidos; la ventana son varios días)
    Limite = Table.FirstN(Ordenado, 5000)
in
    Limite
