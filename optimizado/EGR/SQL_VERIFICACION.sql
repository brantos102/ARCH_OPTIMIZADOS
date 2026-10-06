/* Comprobaciones para ejecutar en SSMS, NO en Power Query.
   Ejecuta un bloque a la vez: selecciónalo y pulsa F5.

   Regla del negocio: cada NRO_CONTENEDORA_EMPAQUE es UNA caja del pedido,
   sin importar cuántos items lleve dentro. Un pedido puede tener varias.
   Las cantidades (y por tanto el peso y el costo) se cuentan UNA sola vez
   por pedido + producto: nunca se repiten por caja. */


/* ============================================================================
   1. El pedido del ejemplo: items y cajas.

   102344955 tiene 4 productos en 2 cajas. Debe devolver 4 filas (una por
   producto), con CAJAS = 2 en todas y CONTENEDORAS con las dos contenedoras.
   Si salen 8 filas, la consulta nueva no se aplicó.
   ============================================================================ */

SELECT
    CAST(SYD.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    p_aux.PRODUCTO_ID,
    ISNULL(p_aux.qty_conf, 0) AS CONFIRMADA,
    (SELECT COUNT(DISTINCT V2.NRO_CONTENEDORA_EMPAQUE)
       FROM VIEW_TIEMPO_EMPAQUETADO V2 (NOLOCK)
      WHERE V2.PEDIDO = SYD.DOC_EXT
        AND V2.CLIENTE_ID = SYD.CLIENTE_ID
        AND V2.NRO_CONTENEDORA_EMPAQUE IS NOT NULL) AS CAJAS
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


/* ============================================================================
   2. Cajas por pedido del día: esto es lo que debe salir en BULTOS.

   La columna CAJAS de la consulta Estado alimenta DATOS!R (BULTOS), y de ahí
   pasa a DESPACHOS, TRAMACO y a la etiqueta. Compara estos números con los
   que muestra el panel (vista EMPAQUE, columna CAJAS).
   ============================================================================ */

SELECT
    CAST(VTE.PEDIDO AS VARCHAR(50)) AS PEDIDO,
    COUNT(DISTINCT VTE.NRO_CONTENEDORA_EMPAQUE) AS CAJAS
FROM VIEW_TIEMPO_EMPAQUETADO VTE (NOLOCK)
JOIN SYS_INT_DOCUMENTO_ADICIONAL ad (NOLOCK)
    ON (ad.doc_ext = VTE.PEDIDO AND ad.cliente_id = VTE.CLIENTE_ID)
WHERE VTE.CLIENTE_ID = 'HYCITE2'
  AND VTE.NRO_CONTENEDORA_EMPAQUE IS NOT NULL
  AND ad.FECHA_CREACION >= DATEADD(DAY, -1, CAST(GETDATE() AS DATE))
GROUP BY CAST(VTE.PEDIDO AS VARCHAR(50))
ORDER BY 2 DESC


/* ============================================================================
   3. Total de unidades por pedido, para facturación.

   Este total NO debe cambiar con la consulta nueva: es el mismo picking.
   Si cambia, algo está duplicando filas.
   ============================================================================ */

SELECT
    CAST(s2.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    COUNT(DISTINCT p.PRODUCTO_ID) AS PRODUCTOS,
    SUM(ISNULL(p.CANT_CONFIRMADA, 0)) AS UNIDADES_CONFIRMADAS
FROM picking p (NOLOCK)
JOIN SYS_INT_DET_DOCUMENTO s2 (NOLOCK)
    ON s2.DOCUMENTO_ID = p.DOCUMENTO_ID
JOIN SYS_INT_DOCUMENTO_ADICIONAL ad (NOLOCK)
    ON (ad.doc_ext = s2.doc_ext AND ad.cliente_id = s2.cliente_id)
WHERE s2.CLIENTE_ID = 'HYCITE2'
  AND ad.FECHA_CREACION >= DATEADD(DAY, -1, CAST(GETDATE() AS DATE))
GROUP BY CAST(s2.DOC_EXT AS VARCHAR(50))
ORDER BY 3 DESC


/* ============================================================================
   4. Pedidos del día que todavía no tienen ninguna contenedora.

   Son los que aún no se empacan: en el panel salen como PICKEADO o SIN PICKING
   y en BULTOS quedan en 1 hasta que se empaquen.
   ============================================================================ */

SELECT
    CAST(SYD.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    syd_ad.FECHA_CREACION
FROM SYS_INT_DOCUMENTO syd (NOLOCK)
LEFT JOIN SYS_INT_DOCUMENTO_ADICIONAL syd_ad (NOLOCK)
    ON (syd.doc_ext = syd_ad.doc_ext AND syd.cliente_id = syd_ad.cliente_id)
WHERE SYD.cliente_id IN ('HYCITE2')
  AND syd_AD.FECHA_CREACION >= DATEADD(DAY, -1, CAST(GETDATE() AS DATE))
  AND NOT EXISTS (SELECT 1
                    FROM VIEW_TIEMPO_EMPAQUETADO V (NOLOCK)
                   WHERE V.PEDIDO = SYD.DOC_EXT
                     AND V.CLIENTE_ID = SYD.CLIENTE_ID
                     AND V.NRO_CONTENEDORA_EMPAQUE IS NOT NULL)


/* ============================================================================
   5. Columnas de VIEW_TIEMPO_EMPAQUETADO, por si hace falta otro dato
      de la caja (peso real de la caja, fecha de empaque, etc.).
   ============================================================================ */

SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'VIEW_TIEMPO_EMPAQUETADO'
ORDER BY ORDINAL_POSITION
