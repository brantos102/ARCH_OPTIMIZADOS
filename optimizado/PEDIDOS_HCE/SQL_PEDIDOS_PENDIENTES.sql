SELECT DISTINCT
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
FROM sys_int_documento syd WITH (NOLOCK)
WHERE syd.tipo_documento_id = 'E04'
  AND syd.cliente_id = 'HYCITE2'
  AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2')
  AND EXISTS (
        SELECT 1
        FROM sys_int_documento_adicional ad WITH (NOLOCK)
        WHERE ad.doc_ext = syd.doc_ext
          AND ad.cliente_id = syd.cliente_id
          AND ad.fecha_creacion >= DATEADD(DAY, -7, CAST(GETDATE() AS DATE))
          AND ad.fecha_creacion <  DATEADD(DAY, 1, CAST(GETDATE() AS DATE)))
  AND NOT EXISTS (
        SELECT 1
        FROM sys_int_det_documento dd WITH (NOLOCK)
        INNER JOIN picking pk WITH (NOLOCK)
                ON pk.documento_id = dd.documento_id
        WHERE dd.doc_ext = syd.doc_ext
          AND dd.cliente_id = syd.cliente_id);
