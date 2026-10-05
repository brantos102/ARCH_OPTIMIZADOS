/* Comprobaciones para ejecutar en SSMS, NO en Power Query.
   Ejecuta cada bloque por separado (selecciona el bloque y pulsa F5). */


/* 1. El pedido del ejemplo: deben salir 4 filas, una por producto.
      Antes salian 8: los mismos 4 productos repetidos por cada contenedora. */

SELECT DISTINCT
    CAST(SYD.DOC_EXT AS VARCHAR(50)) AS DOC_EXT,
    p_aux.PRODUCTO_ID,
    ISNULL(p_aux.qty_conf, 0) AS Cantidad_Confirmada,
    (SELECT MIN(CAST(VTE.NRO_CONTENEDORA_EMPAQUE AS VARCHAR(50)))
       FROM VIEW_TIEMPO_EMPAQUETADO VTE (NOLOCK)
      WHERE VTE.PEDIDO = SYD.DOC_EXT
        AND VTE.CLIENTE_ID = SYD.CLIENTE_ID) AS NRO_CONTENEDORA_EMPAQUE
FROM SYS_INT_DOCUMENTO syd (NOLOCK)
LEFT JOIN SYS_INT_DET_DOCUMENTO syded (NOLOCK)
    ON (syd.doc_ext = syded.doc_ext AND syd.cliente_id = syded.cliente_id)
LEFT JOIN DOCUMENTO doc (NOLOCK)
    ON (doc.DOCUMENTO_ID = syded.DOCUMENTO_ID)
LEFT JOIN (
    SELECT
        p.DOCUMENTO_ID,
        p.PRODUCTO_ID,
        SUM(ISNULL(p.CANT_CONFIRMADA, 0)) AS qty_conf
    FROM picking p (NOLOCK)
    GROUP BY
        p.DOCUMENTO_ID,
        p.PRODUCTO_ID
) p_aux ON (p_aux.DOCUMENTO_ID = doc.DOCUMENTO_ID)
WHERE SYD.cliente_id IN ('HYCITE2')
  AND SYD.DOC_EXT = '102344955'


/* 2. Cuantas cajas tiene cada pedido.
      Un pedido con 2 cajas era el que duplicaba sus productos. */

SELECT
    CAST(PEDIDO AS VARCHAR(50)) AS PEDIDO,
    COUNT(DISTINCT NRO_CONTENEDORA_EMPAQUE) AS CAJAS
FROM VIEW_TIEMPO_EMPAQUETADO (NOLOCK)
WHERE CLIENTE_ID = 'HYCITE2'
GROUP BY CAST(PEDIDO AS VARCHAR(50))
HAVING COUNT(DISTINCT NRO_CONTENEDORA_EMPAQUE) > 1


/* 3. Columnas de VIEW_TIEMPO_EMPAQUETADO.
      Si tiene una columna de producto, se puede saber en que caja viajo cada
      producto en lugar de devolver solo una caja por pedido. */

SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'VIEW_TIEMPO_EMPAQUETADO'
ORDER BY ORDINAL_POSITION
