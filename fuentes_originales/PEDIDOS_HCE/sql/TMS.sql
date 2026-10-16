-- Consulta Power Query: TMS
-- Origen: Odbc.Query("dsn=TMS1")
-- Texto SQL nativo tal como está en el libro (solo se decodificaron los saltos de línea #(lf)).

SELECT 
    idreferencia,
    fecha_interfaz 
FROM 
    pedido
WHERE 
    idempresa = '1'
    AND fecha_interfaz >= CURDATE() - INTERVAL 7 DAY;
