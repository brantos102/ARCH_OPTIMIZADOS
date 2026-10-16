section Section1;

shared Depot = let
    Origen = Odbc.Query("dsn=DEPOTUIO", "SELECT DISTINCT#(lf)    -- 1. Selección de campos solicitados (Sintaxis nativa T-SQL)#(lf)    SUBSTRING(syd.cliente_id, 1, 6) AS EMPRESA,#(lf)    syd.CUSTOMS_1 AS DIRECCION_DEST,#(lf)    syd.doc_ext AS NRO_REFERENCIA,#(lf)    syd.doc_ext AS NRO_REFERENCIA_VIAJE,#(lf)    syd.info_adicional_3 AS DESTINATARIO,#(lf)    syd.info_adicional_6 AS PROVINCIA_DEST,#(lf)    syd.CUSTOMS_2 AS LOCALIDAD_DEST,#(lf)    syd.CUSTOMS_3 AS CP_DEST,#(lf)    REPLACE('0' + syd.info_adicional_5, '-', '') AS TELEFONO_MOVIL,#(lf)    REPLACE('0' + syd.info_adicional_4, '-', '') AS Telefono_FIJO,#(lf)    syd.NRO_DESPACHO_IMPORTACION AS Email#(lf)FROM sys_int_documento syd#(lf)#(lf)-- 2. Eliminamos LTRIM/RTRIM para permitir que SQL Server use los índices existentes#(lf)INNER JOIN cliente cl #(lf)    ON syd.cliente_id = cl.cliente_id#(lf)#(lf)INNER JOIN sys_int_documento_adicional syd_ad #(lf)    ON syd.doc_ext = syd_ad.doc_ext #(lf)    AND syd.cliente_id = syd_ad.cliente_id#(lf)#(lf)WHERE syd.tipo_documento_id = 'E04'#(lf)  AND syd.cliente_id = 'HYCITE2'#(lf)  #(lf)  -- 3. Filtro SARGable en SQL Server (Permite Index Seek en lugar de Scan)#(lf)  -- Equivale a: fecha_creacion sea de hace 7 días o más reciente#(lf)  AND syd_ad.fecha_creacion >= DATEADD(day, -1, GETDATE())#(lf)  #(lf)  -- 4. Exclusión directa sin funciones anidadas#(lf)  AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2');")
in
    Origen;

shared TMS = let
    Origen = Odbc.Query("dsn=TMS1", "SELECT #(lf)    idreferencia,#(lf)    fecha_interfaz #(lf)FROM #(lf)    pedido#(lf)WHERE #(lf)    idempresa = '1'#(lf)    AND fecha_interfaz >= CURDATE() - INTERVAL 7 DAY;")
in
    Origen;
