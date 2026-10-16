-- Consulta Power Query: "ITEMS API"
-- Origen: Odbc.Query("dsn=DEPOTUIO")
-- Texto SQL nativo tal como está en el libro (solo se decodificaron los saltos de línea #(lf)).

SELECT 
    CLIENTE_ID,
    DOC_EXT,
    ESTADO_GT,
    FECHA_ESTADO_GT,
    NRO_LINEA,
    PRODUCTO_ID,
    CANTIDAD_SOLICITADA,
    UNIDAD_ID,
    DOCUMENTO_ID
FROM SYS_INT_DET_DOCUMENTO WITH (NOLOCK)
WHERE CLIENTE_ID = 'HYCITE2'
  -- 🔄 Filtra automáticamente todo lo generado el día de HOY
  AND FECHA_ESTADO_GT >= CAST(GETDATE() AS DATE)
  AND FECHA_ESTADO_GT <  CAST(DATEADD(DAY, 1, GETDATE()) AS DATE)
ORDER BY FECHA_ESTADO_GT DESC, DOC_EXT, NRO_LINEA;
