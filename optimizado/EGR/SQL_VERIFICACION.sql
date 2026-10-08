/* Comprobaciones para ejecutar en SSMS, NO en Power Query.
   Ejecuta un bloque a la vez: selecciónalo y pulsa F5.

   Reglas del negocio que se están comprobando:
   - cada contenedora es UNA caja, sin importar cuántos ítems lleve;
   - cada ítem va en la contenedora en la que REALMENTE se empacó: eso es lo que
     permite calcular el box density (volumen de los ítems de esa caja / volumen
     de la caja);
   - el número de cajas del pedido es el TOTAL de contenedoras, y ese total se calcula
     en la hoja EMPAQUETADO, no en ITEMS DEPOT (ahi no se puede agregar una columna);
   - las cantidades no se repiten: el total de unidades, el peso y el costo del
     pedido tienen que seguir siendo los reales.

   Todos los bloques filtran SOLO EL DIA DE HOY, igual que la consulta.

   EL BLOQUE 1 ES EL QUE DECIDE SI LA CONSULTA NUEVA SIRVE. HAZLO PRIMERO. */


/* ============================================================================
   1. ¿La contenedora del picking es la misma que la de VIEW_TIEMPO_EMPAQUETADO?

   La consulta nueva toma la contenedora de cada ítem de PICKING.NRO_UCEMPAQUETADO,
   que es la caja en la que se empacó esa línea. El total de cajas lo sigue dando
   VIEW_TIEMPO_EMPAQUETADO. Las dos columnas deben dar el mismo número.

   - Si coinciden: aplica la consulta.
   - Si CAJAS_SEGUN_PICKING sale 0 en pedidos ya empacados, NO la apliques y avísame:
     habría que sacar la contenedora de cada ítem de otra columna.
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
  AND syd_AD.FECHA_CREACION >= CAST(GETDATE() AS DATE)
  AND syd_AD.FECHA_CREACION < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))
ORDER BY 3 DESC


/* ============================================================================
   2. Un pedido de varias cajas, ítem por ítem.

   Cambia el número de pedido por uno que en el bloque 1 tenga 2 o 3 cajas.
   Cada ítem debe salir con SU contenedora, no con la misma en todas las filas.
   Si un producto se repartió en dos cajas, salen dos filas con la cantidad que
   fue en cada una: sumadas dan la cantidad real del pedido.
   ============================================================================ */

SELECT
    CAST(s2.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    p.PRODUCTO_ID,
    CAST(p.NRO_UCEMPAQUETADO AS VARCHAR(50)) AS CONTENEDORA,
    SUM(ISNULL(p.CANT_CONFIRMADA, 0)) AS CONFIRMADA
FROM picking p (NOLOCK)
JOIN SYS_INT_DET_DOCUMENTO s2 (NOLOCK)
    ON s2.DOCUMENTO_ID = p.DOCUMENTO_ID
WHERE s2.CLIENTE_ID = 'HYCITE2'
  AND s2.DOC_EXT = '102344955'
GROUP BY
    CAST(s2.DOC_EXT AS VARCHAR(50)),
    p.PRODUCTO_ID,
    CAST(p.NRO_UCEMPAQUETADO AS VARCHAR(50))
ORDER BY 3, 2


/* ============================================================================
   3. Box density: unidades por contenedora.

   Esta es la base del porcentaje de ocupación. Cada contenedora debe salir una
   sola vez y con los ítems que realmente lleva. Si una contenedora apareciera
   con TODOS los ítems del pedido, la consulta está mal aplicada.
   ============================================================================ */

SELECT
    CAST(s2.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    CAST(p.NRO_UCEMPAQUETADO AS VARCHAR(50)) AS CONTENEDORA,
    COUNT(DISTINCT p.PRODUCTO_ID) AS PRODUCTOS,
    SUM(ISNULL(p.CANT_CONFIRMADA, 0)) AS UNIDADES
FROM picking p (NOLOCK)
JOIN SYS_INT_DET_DOCUMENTO s2 (NOLOCK)
    ON s2.DOCUMENTO_ID = p.DOCUMENTO_ID
JOIN SYS_INT_DOCUMENTO_ADICIONAL ad (NOLOCK)
    ON (ad.doc_ext = s2.doc_ext AND ad.cliente_id = s2.cliente_id)
WHERE s2.CLIENTE_ID = 'HYCITE2'
  AND p.NRO_UCEMPAQUETADO IS NOT NULL
  AND ad.FECHA_CREACION >= CAST(GETDATE() AS DATE)
  AND ad.FECHA_CREACION < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))
GROUP BY
    CAST(s2.DOC_EXT AS VARCHAR(50)),
    CAST(p.NRO_UCEMPAQUETADO AS VARCHAR(50))
ORDER BY 1, 2


/* ============================================================================
   4. Total de unidades por pedido, para facturación.

   Este total NO debe cambiar con la consulta nueva: es el mismo picking, solo
   que ahora repartido por contenedora en lugar de repetido. Compáralo contra
   el mismo bloque ejecutado antes de aplicar la consulta.
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
  AND ad.FECHA_CREACION >= CAST(GETDATE() AS DATE)
  AND ad.FECHA_CREACION < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))
GROUP BY CAST(s2.DOC_EXT AS VARCHAR(50))
ORDER BY 3 DESC


/* ============================================================================
   5. Ítems ya confirmados que todavía no tienen contenedora.

   Son los que se pickearon pero aún no se empacan. En ITEMS DEPOT salen con la
   contenedora vacía (o con la única del pedido, si el pedido tiene una sola).
   ============================================================================ */

SELECT
    CAST(s2.DOC_EXT AS VARCHAR(50)) AS PEDIDO,
    COUNT(*) AS LINEAS_SIN_CONTENEDORA,
    SUM(ISNULL(p.CANT_CONFIRMADA, 0)) AS UNIDADES
FROM picking p (NOLOCK)
JOIN SYS_INT_DET_DOCUMENTO s2 (NOLOCK)
    ON s2.DOCUMENTO_ID = p.DOCUMENTO_ID
JOIN SYS_INT_DOCUMENTO_ADICIONAL ad (NOLOCK)
    ON (ad.doc_ext = s2.doc_ext AND ad.cliente_id = s2.cliente_id)
WHERE s2.CLIENTE_ID = 'HYCITE2'
  AND p.NRO_UCEMPAQUETADO IS NULL
  AND ad.FECHA_CREACION >= CAST(GETDATE() AS DATE)
  AND ad.FECHA_CREACION < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))
GROUP BY CAST(s2.DOC_EXT AS VARCHAR(50))
ORDER BY 2 DESC


/* ============================================================================
   6. Columnas de VIEW_TIEMPO_EMPAQUETADO, por si hace falta otro dato
      de la caja (peso real de la caja, fecha de empaque, tipo de caja).

   Si esta vista tuviera PRODUCTO_ID, se podría sacar de aquí la contenedora de
   cada ítem en lugar de PICKING.NRO_UCEMPAQUETADO. Si la ves, avísame.
   ============================================================================ */

SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'VIEW_TIEMPO_EMPAQUETADO'
ORDER BY ORDINAL_POSITION
