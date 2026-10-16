# Inventario de fórmulas — PEDIDOS HCE.xlsm

Generado automáticamente desde el archivo recibido (datos al 24/09/2026). Para cada columna con fórmulas se muestra
la fórmula de la primera celda, el número de celdas que la usan y cuántas variantes distintas existen
(comparadas en forma relativa: 1 variante = la misma fórmula copiada en toda la columna). No incluye datos de clientes.

## Hoja `COBERTURA` (visible, rango A1:Q1303, 0 fórmulas)

- Tabla **Tabla45** `A1:Q1303`
- Formato condicional: 1 reglas

Sin fórmulas en celdas.

## Hoja `Claude Log` (oculta, rango A1:F9, 0 fórmulas)


Sin fórmulas en celdas.

## Hoja `DEPOT` (visible, rango A1:V136, 135 fórmulas)

- Tabla **Depot** `A1:N136` (tabla de consulta Power Query)
  - columna calculada `ESTAD EN TMS`: `=IF(VLOOKUP(Depot[[#This Row],[NRO_REFERENCIA]],TMS[idreferencia],1,0)=Depot[[#This Row],[NRO_REFERENCIA]],"Cargado en TMS","Pendiente Cargar TMS")`
- Formato condicional: 6 reglas

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| M | ESTAD EN TMS | 2–136 | 135 | 1 | `M2`: `=IF(VLOOKUP(Depot[[#This Row],[NRO_REFERENCIA]],TMS[idreferencia],1,0)=Depot[[#This Row],[NRO_REFERENCIA]],"Cargado en TMS","Pendiente Cargar TMS")` |

## Hoja `CAMBIOS` (visible, rango A1:K63, 0 fórmulas)


Sin fórmulas en celdas.

## Hoja `TMS` (oculta, rango A1:B986, 0 fórmulas)

- Tabla **TMS** `A1:B986` (tabla de consulta Power Query)

Sin fórmulas en celdas.

