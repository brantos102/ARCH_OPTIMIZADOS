/* Comprobaciones para ejecutar en SSMS, NO en Power Query.
   Ejecuta un bloque a la vez: selecciónalo y pulsa F5.

   ESTA CONSULTA ES LA PUERTA DE ENTRADA DEL DIA. Si esconde un pedido, ese pedido no
   se valida, no se le asigna cobertura y no se despacha. Por eso NO se aplica el SQL
   nuevo sin antes correr el bloque 1 y el bloque 2 y mirar los numeros. */


/* ============================================================================
   1. ¿Cuantos pedidos trae cada version, HOY?

   ACTUAL  = lo que trae la consulta de produccion ahora mismo
   VENTANA = la misma, pero con dias completos en vez de "hace 24 horas"
   NUEVA   = la ventana, quitando los que YA tienen picking

   Lo normal es que NUEVA sea menor que VENTANA (los ya pickeados se caen) y
   mayor o igual que ACTUAL (recupera los pendientes de dias anteriores).
   ============================================================================ */

SELECT
    (SELECT COUNT(DISTINCT syd.doc_ext)
       FROM sys_int_documento syd WITH (NOLOCK)
       INNER JOIN sys_int_documento_adicional ad WITH (NOLOCK)
               ON ad.doc_ext = syd.doc_ext AND ad.cliente_id = syd.cliente_id
      WHERE syd.tipo_documento_id = 'E04'
        AND syd.cliente_id = 'HYCITE2'
        AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2')
        AND ad.fecha_creacion >= DATEADD(DAY, -1, GETDATE())) AS ACTUAL,
    (SELECT COUNT(DISTINCT syd.doc_ext)
       FROM sys_int_documento syd WITH (NOLOCK)
      WHERE syd.tipo_documento_id = 'E04'
        AND syd.cliente_id = 'HYCITE2'
        AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2')
        AND EXISTS (SELECT 1 FROM sys_int_documento_adicional ad WITH (NOLOCK)
                     WHERE ad.doc_ext = syd.doc_ext AND ad.cliente_id = syd.cliente_id
                       AND ad.fecha_creacion >= DATEADD(DAY, -7, CAST(GETDATE() AS DATE))
                       AND ad.fecha_creacion <  DATEADD(DAY, 1, CAST(GETDATE() AS DATE)))) AS VENTANA,
    (SELECT COUNT(DISTINCT syd.doc_ext)
       FROM sys_int_documento syd WITH (NOLOCK)
      WHERE syd.tipo_documento_id = 'E04'
        AND syd.cliente_id = 'HYCITE2'
        AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2')
        AND EXISTS (SELECT 1 FROM sys_int_documento_adicional ad WITH (NOLOCK)
                     WHERE ad.doc_ext = syd.doc_ext AND ad.cliente_id = syd.cliente_id
                       AND ad.fecha_creacion >= DATEADD(DAY, -7, CAST(GETDATE() AS DATE))
                       AND ad.fecha_creacion <  DATEADD(DAY, 1, CAST(GETDATE() AS DATE)))
        AND NOT EXISTS (SELECT 1 FROM sys_int_det_documento dd WITH (NOLOCK)
                        INNER JOIN picking pk WITH (NOLOCK) ON pk.documento_id = dd.documento_id
                         WHERE dd.doc_ext = syd.doc_ext AND dd.cliente_id = syd.cliente_id)) AS NUEVA


/* ============================================================================
   2. EL BLOQUE QUE HAY QUE MIRAR CON LUPA.

   Lista los pedidos que la version NUEVA esconderia por tener picking.
   Revisa uno por uno: todos deben ser pedidos que YA se procesaron.
   Si aparece aqui un pedido que todavia hay que despachar, NO apliques el
   filtro NOT EXISTS y avisame.
   ============================================================================ */

SELECT
    CAST(syd.doc_ext AS VARCHAR(50)) AS PEDIDO,
    syd.info_adicional_3 AS DESTINATARIO,
    syd.info_adicional_6 AS PROVINCIA,
    (SELECT MIN(CAST(ad.fecha_creacion AS DATE))
       FROM sys_int_documento_adicional ad WITH (NOLOCK)
      WHERE ad.doc_ext = syd.doc_ext AND ad.cliente_id = syd.cliente_id) AS FECHA_CREACION,
    (SELECT COUNT(*)
       FROM sys_int_det_documento dd WITH (NOLOCK)
       INNER JOIN picking pk WITH (NOLOCK) ON pk.documento_id = dd.documento_id
      WHERE dd.doc_ext = syd.doc_ext AND dd.cliente_id = syd.cliente_id) AS LINEAS_DE_PICKING,
    (SELECT SUM(ISNULL(pk.CANT_CONFIRMADA, 0))
       FROM sys_int_det_documento dd WITH (NOLOCK)
       INNER JOIN picking pk WITH (NOLOCK) ON pk.documento_id = dd.documento_id
      WHERE dd.doc_ext = syd.doc_ext AND dd.cliente_id = syd.cliente_id) AS UNIDADES_CONFIRMADAS
FROM sys_int_documento syd WITH (NOLOCK)
WHERE syd.tipo_documento_id = 'E04'
  AND syd.cliente_id = 'HYCITE2'
  AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2')
  AND EXISTS (SELECT 1 FROM sys_int_documento_adicional ad WITH (NOLOCK)
               WHERE ad.doc_ext = syd.doc_ext AND ad.cliente_id = syd.cliente_id
                 AND ad.fecha_creacion >= DATEADD(DAY, -7, CAST(GETDATE() AS DATE))
                 AND ad.fecha_creacion <  DATEADD(DAY, 1, CAST(GETDATE() AS DATE)))
  AND EXISTS (SELECT 1 FROM sys_int_det_documento dd WITH (NOLOCK)
              INNER JOIN picking pk WITH (NOLOCK) ON pk.documento_id = dd.documento_id
               WHERE dd.doc_ext = syd.doc_ext AND dd.cliente_id = syd.cliente_id)
ORDER BY 4, 1


/* ============================================================================
   3. Los que la version NUEVA recupera: pendientes de dias anteriores que la
      consulta actual no mostraba (el caso del viernes feriado).
   ============================================================================ */

SELECT
    CAST(syd.doc_ext AS VARCHAR(50)) AS PEDIDO,
    syd.info_adicional_3 AS DESTINATARIO,
    syd.info_adicional_6 AS PROVINCIA,
    (SELECT MIN(CAST(ad.fecha_creacion AS DATE))
       FROM sys_int_documento_adicional ad WITH (NOLOCK)
      WHERE ad.doc_ext = syd.doc_ext AND ad.cliente_id = syd.cliente_id) AS FECHA_CREACION
FROM sys_int_documento syd WITH (NOLOCK)
WHERE syd.tipo_documento_id = 'E04'
  AND syd.cliente_id = 'HYCITE2'
  AND (syd.agente_id IS NULL OR syd.agente_id <> 'HYCITE2')
  AND EXISTS (SELECT 1 FROM sys_int_documento_adicional ad WITH (NOLOCK)
               WHERE ad.doc_ext = syd.doc_ext AND ad.cliente_id = syd.cliente_id
                 AND ad.fecha_creacion >= DATEADD(DAY, -7, CAST(GETDATE() AS DATE))
                 AND ad.fecha_creacion <  DATEADD(DAY, -1, GETDATE()))
  AND NOT EXISTS (SELECT 1 FROM sys_int_det_documento dd WITH (NOLOCK)
                  INNER JOIN picking pk WITH (NOLOCK) ON pk.documento_id = dd.documento_id
                   WHERE dd.doc_ext = syd.doc_ext AND dd.cliente_id = syd.cliente_id)
ORDER BY 4, 1


/* ============================================================================
   4. Para entender los estados: que valores de ESTADO_GT y de STATUS tienen los
      pedidos del rango, y cuantos hay de cada uno.

   Si de aqui sale un estado que signifique "ya despachado" o "anulado", se puede
   afinar el filtro por estado en vez de por picking. Mandame el resultado.
   ============================================================================ */

SELECT
    dd.ESTADO_GT,
    doc.STATUS,
    COUNT(DISTINCT dd.doc_ext) AS PEDIDOS
FROM sys_int_det_documento dd WITH (NOLOCK)
LEFT JOIN documento doc WITH (NOLOCK)
       ON doc.documento_id = dd.documento_id
INNER JOIN sys_int_documento_adicional ad WITH (NOLOCK)
        ON ad.doc_ext = dd.doc_ext AND ad.cliente_id = dd.cliente_id
WHERE dd.cliente_id = 'HYCITE2'
  AND ad.fecha_creacion >= DATEADD(DAY, -7, CAST(GETDATE() AS DATE))
  AND ad.fecha_creacion <  DATEADD(DAY, 1, CAST(GETDATE() AS DATE))
GROUP BY dd.ESTADO_GT, doc.STATUS
ORDER BY 3 DESC
