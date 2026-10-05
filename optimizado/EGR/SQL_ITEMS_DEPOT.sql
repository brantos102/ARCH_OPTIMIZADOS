/* =====================================================================================
   Consulta "Estado" -> hoja ITEMS DEPOT del Formato EGR_FL_HYCITE
   VERSION CORREGIDA: una sola fila por PEDIDO + PRODUCTO.

   PROBLEMA DE LA VERSION ANTERIOR
   -------------------------------
   El pedido 102344955 tiene 4 productos (CO4840, CO4911, CO9771, SP0098), pero en
   ITEMS DEPOT aparecian 8 filas: los mismos 4 productos repetidos con las
   contenedoras 1559332 y 1559342.

   La causa es esta linea:

       LEFT JOIN VIEW_TIEMPO_EMPAQUETADO VTE ON VTE.PEDIDO = SYD.DOC_EXT ...

   VIEW_TIEMPO_EMPAQUETADO devuelve UNA FILA POR CONTENEDORA. Al unirla fila a fila,
   cada linea de picking se multiplica por el numero de cajas del pedido:

       4 productos x 2 contenedoras = 8 filas

   El SELECT DISTINCT no lo corrige, porque las filas SI son distintas: cambia la
   columna NRO_CONTENEDORA_EMPAQUE.

   Impacto en costos: las columnas calculadas PESO, VOLUMEN y PRECIO multiplican el
   valor unitario por Cantidad_Confirmada en CADA fila. Con las filas duplicadas, el
   peso y el costo del pedido salian al doble (14,85 kg reales -> 29,70 kg).
   Un pedido de 3 cajas los habria triplicado.

   QUE CAMBIA
   ----------
   1. Las contenedoras se agregan ANTES de unirlas (subconsulta "emp"), asi que ya no
      multiplican filas. Se devuelve una contenedora por pedido, que es justamente lo
      que lee la hoja EMPAQUETADO con XLOOKUP.
   2. El picking se agrupa por DOCUMENTO + PRODUCTO, asi que un producto tomado de dos
      posiciones o dos lotes suma, en vez de repetirse.
   3. El GROUP BY final garantiza una fila por PEDIDO + PRODUCTO, pase lo que pase.
   4. Se quitan las uniones que no se usaban (CLIENTE y SUCURSAL).

   Las 7 columnas salen con el MISMO nombre y en el MISMO orden, asi que la tabla de
   Excel, sus columnas calculadas (PEDIDO, PESO, VOLUMEN, PRECIO) y las tablas
   dinamicas siguen funcionando sin tocar nada.

   COMO APLICARLA
   --------------
   1. Datos > Consultas y conexiones > clic derecho en "Estado" > Editar.
   2. En Power Query, en PASOS APLICADOS, pulsa el engranaje del paso "Origen".
   3. Borra el SQL del cuadro y pega TODO lo que esta debajo de la linea de guiones.
   4. Aceptar > Cerrar y cargar.
   5. Comprueba con la consulta de verificacion que esta al final de este archivo.
   ===================================================================================== */

-- ---------------------------------------------------------------------------------------

WITH cab AS (
    /* Cabeceras del cliente con fecha de creacion, una fila por pedido */
    SELECT
        syd.CLIENTE_ID,
        CAST(syd.DOC_EXT AS VARCHAR(50))          AS DOC_EXT,
        syd.TIPO_DOCUMENTO_ID,
        CAST(MIN(syd_ad.FECHA_CREACION) AS DATE)  AS FECHA_ALTA
    FROM SYS_INT_DOCUMENTO syd WITH (NOLOCK)
    INNER JOIN SYS_INT_DOCUMENTO_ADICIONAL syd_ad WITH (NOLOCK)
            ON syd_ad.DOC_EXT   = syd.DOC_EXT
           AND syd_ad.CLIENTE_ID = syd.CLIENTE_ID
    WHERE syd.CLIENTE_ID IN ('HYCITE2')
      -- FILTRO DINAMICO: desde ayer a las 00:00 en adelante
      AND syd_ad.FECHA_CREACION >= DATEADD(DAY, -1, CAST(GETDATE() AS DATE))
    GROUP BY
        syd.CLIENTE_ID,
        CAST(syd.DOC_EXT AS VARCHAR(50)),
        syd.TIPO_DOCUMENTO_ID
),
docs AS (
    /* Documentos de egreso de cada pedido, sin repetir por linea de detalle */
    SELECT DISTINCT
        c.CLIENTE_ID,
        c.DOC_EXT,
        c.FECHA_ALTA,
        d.DOCUMENTO_ID,
        d.STATUS
    FROM cab c
    INNER JOIN SYS_INT_DET_DOCUMENTO sdd WITH (NOLOCK)
            ON sdd.DOC_EXT    = c.DOC_EXT
           AND sdd.CLIENTE_ID = c.CLIENTE_ID
    INNER JOIN DOCUMENTO d WITH (NOLOCK)
            ON d.DOCUMENTO_ID = sdd.DOCUMENTO_ID
    LEFT JOIN TIPO_COMPROBANTE tc WITH (NOLOCK)
           ON tc.TIPO_COMPROBANTE_ID = c.TIPO_DOCUMENTO_ID
    WHERE UPPER(ISNULL(tc.TIPO_OPERACION_ID, '')) = 'EGR'
       OR UPPER(ISNULL(d.TIPO_OPERACION_ID,  '')) = 'EGR'
),
pick AS (
    /* Picking agrupado por documento y producto: suma lotes, series y posiciones */
    SELECT
        p.DOCUMENTO_ID,
        p.PRODUCTO_ID,
        SUM(p.CANTIDAD)                     AS CANT_PROC,
        SUM(ISNULL(p.CANT_CONFIRMADA, 0))   AS CANT_CONF
    FROM PICKING p WITH (NOLOCK)
    GROUP BY
        p.DOCUMENTO_ID,
        p.PRODUCTO_ID
),
emp AS (
    /* Contenedoras agregadas por pedido: AQUI se corta la duplicacion.
       Antes esta vista se unia fila a fila y multiplicaba cada producto por caja. */
    SELECT
        vte.CLIENTE_ID,
        CAST(vte.PEDIDO AS VARCHAR(50))                              AS PEDIDO,
        MIN(CAST(vte.NRO_CONTENEDORA_EMPAQUE AS VARCHAR(50)))        AS NRO_CONTENEDORA_EMPAQUE
    FROM VIEW_TIEMPO_EMPAQUETADO vte WITH (NOLOCK)
    WHERE vte.NRO_CONTENEDORA_EMPAQUE IS NOT NULL
    GROUP BY
        vte.CLIENTE_ID,
        CAST(vte.PEDIDO AS VARCHAR(50))
)
SELECT
    MIN(d.FECHA_ALTA)                       AS FECHA_ALTA,
    MAX(d.STATUS)                           AS STATUS,
    d.DOC_EXT                               AS DOC_EXT,
    pk.PRODUCTO_ID                          AS PRODUCTO_ID,
    ISNULL(SUM(pk.CANT_PROC), 0)            AS Cantidad_Procesada,
    ISNULL(SUM(pk.CANT_CONF), 0)            AS Cantidad_Confirmada,
    MIN(e.NRO_CONTENEDORA_EMPAQUE)          AS NRO_CONTENEDORA_EMPAQUE
FROM docs d
LEFT JOIN pick pk
       ON pk.DOCUMENTO_ID = d.DOCUMENTO_ID
LEFT JOIN emp e
       ON e.PEDIDO     = d.DOC_EXT
      AND e.CLIENTE_ID = d.CLIENTE_ID
GROUP BY
    d.DOC_EXT,
    pk.PRODUCTO_ID
ORDER BY
    d.DOC_EXT,
    pk.PRODUCTO_ID;


/* =====================================================================================
   VERIFICACION (ejecutar aparte, no forma parte de la consulta de Excel)

   1) No debe devolver ninguna fila: si devuelve alguna, ese pedido+producto sigue
      repetido.
   ===================================================================================== */
/*
WITH q AS (
    -- pegar aqui la consulta corregida completa, sin el ORDER BY
)
SELECT DOC_EXT, PRODUCTO_ID, COUNT(*) AS VECES
FROM q
GROUP BY DOC_EXT, PRODUCTO_ID
HAVING COUNT(*) > 1;
*/

/* 2) Control del pedido del ejemplo: deben salir 4 filas y 14,85 kg en total
      (0,45 + 7,00 + 4,50 + 2,90), no 8 filas ni 29,70 kg.

SELECT DOC_EXT, PRODUCTO_ID, Cantidad_Confirmada, NRO_CONTENEDORA_EMPAQUE
FROM ( ... consulta corregida ... ) x
WHERE DOC_EXT = '102344955';
*/

/* 3) Cuantas cajas tiene realmente cada pedido (informativo: la consulta de Excel
      devuelve solo una contenedora por pedido, que es lo que usa la hoja EMPAQUETADO).

SELECT
    CAST(PEDIDO AS VARCHAR(50))                  AS PEDIDO,
    COUNT(DISTINCT NRO_CONTENEDORA_EMPAQUE)      AS CAJAS
FROM VIEW_TIEMPO_EMPAQUETADO WITH (NOLOCK)
WHERE CLIENTE_ID = 'HYCITE2'
GROUP BY CAST(PEDIDO AS VARCHAR(50))
ORDER BY CAJAS DESC;
*/
