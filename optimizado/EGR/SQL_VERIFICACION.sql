/* Comprobaciones para ejecutar en SSMS, NO en Power Query.
   Ejecuta un bloque a la vez: selecciónalo y pulsa F5.
   El bloque 1 es el que decide si la consulta nueva sirve: hazlo primero. */


/* ============================================================================
   1. ¿La contenedora del picking es la misma que la de VIEW_TIEMPO_EMPAQUETADO?

   La consulta nueva toma la caja de PICKING.NRO_UCEMPAQUETADO, que es la caja
   en la que se empacó cada producto. Esta comprobación muestra, por pedido,
   las cajas que ve cada origen. Deben coincidir.

   Si la columna UC sale vacía en todos los pedidos ya empacados, avísame:
   habría que tomar la caja de otra columna.
   ============================================================================ */

SELECT
    CAST(SYD.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    (SELECT COUNT(DISTINCT p.NRO_UCEMPAQUETADO)
       FROM picking p (NOLOCK)
       JOIN SYS_INT_DET_DOCUMENTO s2 (NOLOCK) ON s2.DOCUMENTO_ID = p.DOCUMENTO_ID
      WHERE s2.DOC_EXT = SYD.DOC_EXT
        AND s2.CLIENTE_ID = SYD.CLIENTE_ID
        AND p.NRO_UCEMPAQUETADO IS NOT NULL) AS CAJAS_SEGUN_PICKING,
    (SELECT COUNT(DISTINCT VTE.NRO_CONTENEDORA_EMPAQUE)
       FROM VIEW_TIEMPO_EMPAQUETADO VTE (NOLOCK)
      WHERE VTE.PEDIDO = SYD.DOC_EXT
        AND VTE.CLIENTE_ID = SYD.CLIENTE_ID
        AND VTE.NRO_CONTENEDORA_EMPAQUE IS NOT NULL) AS CAJAS_SEGUN_EMPAQUETADO
FROM SYS_INT_DOCUMENTO syd (NOLOCK)
LEFT JOIN SYS_INT_DOCUMENTO_ADICIONAL syd_ad (NOLOCK)
    ON (syd.doc_ext = syd_ad.doc_ext AND syd.cliente_id = syd_ad.cliente_id)
WHERE SYD.cliente_id IN ('HYCITE2')
  AND syd_AD.FECHA_CREACION >= DATEADD(DAY, -1, CAST(GETDATE() AS DATE))
ORDER BY 3 DESC


/* ============================================================================
   2. El pedido del ejemplo, producto por producto y caja por caja.

   El pedido 102344955 tiene 4 productos en 2 cajas. Antes salían 8 filas con
   la cantidad REPETIDA en cada caja (peso al doble). Ahora cada fila lleva la
   cantidad que fue en esa caja, así que el total vuelve a ser el real.
   ============================================================================ */

SELECT
    CAST(SYD.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    p_aux.PRODUCTO_ID,
    p_aux.NRO_CONTENEDORA_EMPAQUE AS CAJA,
    p_aux.qty_conf AS CONFIRMADA
FROM SYS_INT_DOCUMENTO syd (NOLOCK)
LEFT JOIN SYS_INT_DET_DOCUMENTO syded (NOLOCK)
    ON (syd.doc_ext = syded.doc_ext AND syd.cliente_id = syded.cliente_id)
LEFT JOIN DOCUMENTO doc (NOLOCK)
    ON (doc.DOCUMENTO_ID = syded.DOCUMENTO_ID)
LEFT JOIN (
    SELECT
        p.DOCUMENTO_ID,
        p.PRODUCTO_ID,
        CAST(p.NRO_UCEMPAQUETADO AS VARCHAR(50)) AS NRO_CONTENEDORA_EMPAQUE,
        SUM(ISNULL(p.CANT_CONFIRMADA, 0)) AS qty_conf
    FROM picking p (NOLOCK)
    GROUP BY
        p.DOCUMENTO_ID,
        p.PRODUCTO_ID,
        CAST(p.NRO_UCEMPAQUETADO AS VARCHAR(50))
) p_aux ON (p_aux.DOCUMENTO_ID = doc.DOCUMENTO_ID)
WHERE SYD.cliente_id IN ('HYCITE2')
  AND SYD.DOC_EXT = '102344955'


/* ============================================================================
   3. Total de unidades por pedido, para facturación.

   Este total NO debe cambiar al aplicar la consulta nueva: es el mismo
   picking, solo que ahora repartido por caja en lugar de repetido.
   ============================================================================ */

SELECT
    CAST(s2.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    COUNT(DISTINCT p.PRODUCTO_ID) AS PRODUCTOS,
    SUM(ISNULL(p.CANT_CONFIRMADA, 0)) AS UNIDADES_CONFIRMADAS,
    COUNT(DISTINCT p.NRO_UCEMPAQUETADO) AS CAJAS
FROM picking p (NOLOCK)
JOIN SYS_INT_DET_DOCUMENTO s2 (NOLOCK)
    ON s2.DOCUMENTO_ID = p.DOCUMENTO_ID
JOIN SYS_INT_DOCUMENTO_ADICIONAL ad (NOLOCK)
    ON (ad.doc_ext = s2.doc_ext AND ad.cliente_id = s2.cliente_id)
WHERE s2.CLIENTE_ID = 'HYCITE2'
  AND ad.FECHA_CREACION >= DATEADD(DAY, -1, CAST(GETDATE() AS DATE))
GROUP BY CAST(s2.DOC_EXT AS VARCHAR(50))
ORDER BY 4 DESC


/* ============================================================================
   4. Columnas de VIEW_TIEMPO_EMPAQUETADO, por si hace falta otra fuente
      para la caja (peso de la caja, fecha de empaque, etc.).
   ============================================================================ */

SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'VIEW_TIEMPO_EMPAQUETADO'
ORDER BY ORDINAL_POSITION
