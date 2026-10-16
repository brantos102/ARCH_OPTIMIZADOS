-- Consulta Power Query: Estado
-- Origen: Odbc.Query("dsn=DEPOTUIO")
-- Texto SQL nativo tal como está en el libro (solo se decodificaron los saltos de línea #(lf)).

SELECT DISTINCT
    CAST(syd_AD.FECHA_CREACION AS DATE) AS FECHA_ALTA, 
    DOC.STATUS,
    CAST(SYD.DOC_EXT AS VARCHAR(50)) AS DOC_EXT,
    p_aux.PRODUCTO_ID,
    ISNULL(p_aux.qty_proc, 0) AS Cantidad_Procesada,
    ISNULL(p_aux.qty_conf, 0) AS Cantidad_Confirmada,
    CAST(VTE.NRO_CONTENEDORA_EMPAQUE AS VARCHAR(50)) AS NRO_CONTENEDORA_EMPAQUE
FROM SYS_INT_DOCUMENTO syd (NOLOCK)
LEFT JOIN CLIENTE CL (NOLOCK) 
    ON CL.CLIENTE_ID = SYD.CLIENTE_ID
LEFT JOIN SYS_INT_DET_DOCUMENTO syded (NOLOCK) 
    ON (syd.doc_ext = syded.doc_ext AND syd.cliente_id = syded.cliente_id)
LEFT JOIN SYS_INT_DOCUMENTO_ADICIONAL syd_ad (NOLOCK) 
    ON (syd.doc_ext = syd_ad.doc_ext AND syd.cliente_id = syd_ad.cliente_id)
LEFT JOIN TIPO_COMPROBANTE tc (NOLOCK) 
    ON (tc.TIPO_COMPROBANTE_ID = syd.TIPO_DOCUMENTO_ID)
LEFT JOIN SUCURSAL SUC (NOLOCK) 
    ON (suc.CLIENTE_ID = syd.CLIENTE_ID AND SUC.SUCURSAL_ID = SYD.AGENTE_ID)
LEFT JOIN VIEW_TIEMPO_EMPAQUETADO VTE (NOLOCK) 
    ON VTE.PEDIDO = SYD.DOC_EXT AND VTE.CLIENTE_ID = syd.CLIENTE_ID
LEFT JOIN DOCUMENTO doc (NOLOCK) 
    ON (doc.DOCUMENTO_ID = syded.DOCUMENTO_ID)
LEFT JOIN (
    SELECT 
        p.DOCUMENTO_ID,
        p.PICKING_ID,
        p.PRODUCTO_ID,
        p.DESCRIPCION,
        p.CANTIDAD AS CANTIDAD_SOLICITADA,
        ISNULL(p.CANT_CONFIRMADA, 0) AS CANT_CONFIRMADA,
        p.POSICION_COD,
        p.PROP1,
        P.NRO_LOTE,
        P.NRO_PARTIDA,
        P.NRO_SERIE,
        SUM(p.CANTIDAD) AS qty_proc,
        SUM(p.CANT_CONFIRMADA) AS qty_conf,
        MAX(CASE WHEN p.fin_picking = 2 THEN P.FECHA_FIN ELSE NULL END) AS fecha_fin_picking,
        MIN(p.FECHA_UCEMPAQUETADO) AS fecha_min_empaq,
        MAX(p.FECHA_UCEMPAQUETADO) AS fecha_max_empaq,
        SUM(CASE WHEN p.nro_ucempaquetado IS NULL THEN P.CANT_CONFIRMADA ELSE 0 END) AS qty_pendiente_empaq,
        SUM(CASE WHEN p.nro_ucempaquetado IS NOT NULL THEN P.CANT_CONFIRMADA ELSE 0 END) AS qty_ya_empaq
    FROM picking p (NOLOCK)
    GROUP BY 
        p.DOCUMENTO_ID,
        p.PRODUCTO_ID,
        p.DESCRIPCION,
        p.CANTIDAD,
        p.CANT_CONFIRMADA,
        p.POSICION_COD,
        p.PROP1,
        P.NRO_LOTE,
        P.NRO_PARTIDA,
        P.NRO_SERIE,
        P.POSICION_COD,
        p.PICKING_ID
) p_aux ON (p_aux.DOCUMENTO_ID = doc.DOCUMENTO_ID)
WHERE (UPPER(tc.TIPO_OPERACION_ID) = 'EGR' OR UPPER(doc.TIPO_OPERACION_ID) = 'EGR')
  AND SYD.cliente_id IN ('HYCITE2')

  -- 🔄 FILTRO DINÁMICO: Desde ayer a las 00:00:00 en adelante
  AND syd_AD.FECHA_CREACION >= DATEADD(DAY, -1, CAST(GETDATE() AS DATE));
