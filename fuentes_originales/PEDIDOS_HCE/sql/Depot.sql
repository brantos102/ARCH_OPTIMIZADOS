-- Consulta Power Query: Depot
-- Origen: Odbc.Query("dsn=DEPOTUIO")
-- Texto SQL nativo tal como está en el libro (solo se decodificaron los saltos de línea #(lf)).

SELECT DISTINCT
    -- 1. Selección de campos solicitados (Sintaxis nativa T-SQL)
    SUBSTRING(syd.cliente_id, 1, 6) AS EMPRESA,
    syd.CUSTOMS_1 AS DIRECCION_DEST,
    syd.doc_ext AS NRO_REFERENCIA,
    syd.doc_ext AS NRO_REFERENCIA_VIAJE,
    syd.info_adicional_3 AS DESTINATARIO,
    syd.info_adicional_6 AS PROVINCIA_DEST,
    syd.CUSTOMS_2 AS LOCALIDAD_DEST,
    syd.CUSTOMS_3 AS CP_DEST,
    REPLACE('0' + syd.info_adicional_5, '-', '') AS TELEFONO_MOVIL,
    REPLACE('0' + syd.info_adicional_4, '-', '') AS Telefono_FIJO,
    syd.NRO_DESPACHO_IMPORTACION AS Email
FROM sys_int_documento syd

-- 2. Eliminamos LTRIM/RTRIM para permitir que SQL Server use los índices existentes
INNER JOIN cliente cl 
    ON syd.cliente_id = cl.cliente_id

INNER JOIN sys_int_documento_adicional syd_ad 
    ON syd.doc_ext = syd_ad.doc_ext 
    AND syd.cliente_id = syd_ad.cliente_id

WHERE syd.tipo_documento_id = 'E04'
  AND syd.cliente_id = 'HYCITE2'
  
  -- 3. Filtro SARGable en SQL Server (Permite Index Seek en lugar de Scan)
  -- Equivale a: fecha_creacion sea de hace 7 días o más reciente
  AND syd_ad.fecha_creacion >= DATEADD(day, -1, GETDATE())
  
  -- 4. Exclusión directa sin funciones anidadas
  AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2');
