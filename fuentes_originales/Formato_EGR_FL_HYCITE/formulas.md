# Inventario de fórmulas — Formato EGR_FL_HYCITE.xlsm

Generado automáticamente desde el archivo recibido (datos al 24/09/2026). Para cada columna con fórmulas se muestra
la fórmula de la primera celda, el número de celdas que la usan y cuántas variantes distintas existen
(comparadas en forma relativa: 1 variante = la misma fórmula copiada en toda la columna). No incluye datos de clientes.

## Hoja `Claude Log` (visible, rango A1:H7, 3 fórmulas)


| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| H |  | 1–3 | 3 (matriz dinámica) | 3 | `H1`: `=IFERROR(_xlfn.XLOOKUP("AZUAY",'COBERTURAS Y TARIFAS'!B2:B10,'COBERTURAS Y TARIFAS'!C2:C10),"NO")` |

## Hoja `PANEL` (visible, rango A1:C21, 9 fórmulas)


| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| C |  | 5–13 | 9 | 9 | `C5`: `=COUNTIF(DATOS!$B$2:$B$500,">0")` |

## Hoja `MENU` (oculta, rango A1:G13, 0 fórmulas)


Sin fórmulas en celdas.

## Hoja `DATOS` (visible, rango A1:AD500, 7984 fórmulas)

- Formato condicional: 17 reglas

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| A | DESTINO | 2–500 | 499 | 1 | `A2`: `=IF($O2="","",IF($O2="VALIDAR","",IF($O2="PICHINCHA","UIO",IF($O2="GUAYAS","GYE",IF($O2="GALAPAGOS","GPS","PRO")))))` |
| B | PEDIDOS | 2–500 | 499 | 1 | `B2`: `=IF($E2="","",IFERROR(VALUE($E2),$E2))` |
| O | PROVINCIA | 2–500 | 499 | 1 | `O2`: `=IF($H2="","",IFERROR(INDEX('COBERTURAS Y TARIFAS'!$X$2:$X$1804,MATCH(_xlfn.TEXTJOIN("_",TRUE,$H2,$I2,$J2),COBERT_KEY,0)),"VALIDAR"))` |
| P | CANTON | 2–500 | 499 | 1 | `P2`: `=IF($H2="","",IFERROR(INDEX('COBERTURAS Y TARIFAS'!$Y$2:$Y$1804,MATCH(_xlfn.TEXTJOIN("_",TRUE,$H2,$I2,$J2),COBERT_KEY,0)),"VALIDAR"))` |
| Q | PARROQUIA / CIUDAD | 2–500 | 499 | 1 | `Q2`: `=IF($H2="","",IFERROR(INDEX('COBERTURAS Y TARIFAS'!$Z$2:$Z$1804,MATCH(_xlfn.TEXTJOIN("_",TRUE,$H2,$I2,$J2),COBERT_KEY,0)),"VALIDAR"))` |
| R | BULTOS | 2–500 | 499 | 1 | `R2`: `=IFERROR(IF($B2="","",VLOOKUP(TEXT($B2,"0"),'TABLAS DINAMICAS'!$H:$K,2,FALSE)),1)` |
| S | PESOS | 2–500 | 499 | 1 | `S2`: `=IF($B2="","",IFERROR(IF(VLOOKUP(TEXT($B2,"0"),'TABLAS DINAMICAS'!$H:$K,4,FALSE)>0,ROUND(VLOOKUP(TEXT($B2,"0"),'TABLAS DINAMICAS'!$H:$K,4,FALSE),2),NA()),IF(SUMIFS(Estado[PESO],Estado[DOC_EXT],$B2)>0,ROUND(SUMIFS(Estado[PESO],Estado[DOC_EXT],$B2),2),"-")))` |
| T | VALOR | 2–500 | 499 | 1 | `T2`: `=IF($B2="","",IFERROR(VLOOKUP($O2,TARIFAS,3,FALSE)+ROUND(VLOOKUP($O2,TARIFAS,4,FALSE)*IF(NOT(ISNUMBER($S2)),0,IF($S2<=10,0,$S2-10)),2),0))` |
| V | VALIDACION | 2–500 | 499 | 1 | `V2`: `=IF($B2="","",IF(OR($O2="VALIDAR",$P2="VALIDAR",$Q2="VALIDAR"),"REVISAR","OK"))` |
| W | RUTA | 2–500 | 499 | 1 | `W2`: `=IF($H2="","",_xlfn.IFNA(IF(INDEX(RUTA_COB,MATCH(_xlfn.TEXTJOIN("_",TRUE,$H2,$I2,$J2),COBERT_KEY,0))=0,"",INDEX(RUTA_COB,MATCH(_xlfn.TEXTJOIN("_",TRUE,$H2,$I2,$J2),COBERT_KEY,0))),""))` |
| X | ACUM_BULTOS | 2–500 | 499 | 1 | `X2`: `=IF($B2="","",N(X1)+MAX(1,IF(ISNUMBER($R2),$R2,1)))` |
| Y | COURIER | 2–500 | 499 | 1 | `Y2`: `=IF($H2="","",IFERROR(INDEX('COBERTURAS Y TARIFAS'!$E$2:$E$1804,MATCH(_xlfn.TEXTJOIN("_",TRUE,$H2,$I2,$J2),COBERT_KEY,0)),""))` |
| Z | PARR SIN () | 2–500 | 499 | 1 | `Z2`: `=IF(AND(ISNUMBER(SEARCH("(",Q2)),RIGHT(Q2,1)=")"),LEFT(Q2,SEARCH("(",Q2)-2),Q2)` |
| AA | CONCA | 2–500 | 499 | 1 | `AA2`: `=+CONCATENATE(O2,P2,Z2)` |
| AB | PARR TMS | 2–500 | 499 | 1 | `AB2`: `=IFERROR(VLOOKUP($AA2,'COD POS'!$D:$E,2,0),"")` |
| AC | N_PRO | 2–500 | 499 | 1 | `AC2`: `=IF($B2="","",IF($A2="PRO",COUNTIF($A$2:$A2,"PRO"),""))` |

## Hoja `TMS` (oculta, rango A1:AP255, 6858 fórmulas)


| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| A | EMPRESA | 2–255 | 254 | 1 | `A2`: `=IF($J2="","","HYCITE")` |
| B | FECHA_INTERFAZ | 2–255 | 254 | 1 | `B2`: `=IF($J2="","",TODAY())` |
| C | REMITENTE | 2–255 | 254 | 1 | `C2`: `=IF($J2="","","HYCITE QUITO")` |
| D | DIRECCION_RTTE | 2–255 | 254 | 1 | `D2`: `=IF($J2="","","FLEX NET CALLE 28 DE JUNIO S/N Y GARCIA MORENO ENTRADA LLANO GRANDE")` |
| E | PISO_RTTE | 2–255 | 254 | 1 | `E2`: `=IF($J2="","","1")` |
| F | DEPTO_RTTE | 2–255 | 254 | 1 | `F2`: `=IF($J2="","","1")` |
| G | CP_RTTE | 2–255 | 254 | 1 | `G2`: `=IF($J2="","","CALDERON (QUI)")` |
| H | LOCALIDAD_RTTE | 2–255 | 254 | 1 | `H2`: `=IF($J2="","","QUITO")` |
| I | PROVINCIA_RTTE | 2–255 | 254 | 1 | `I2`: `=IF($J2="","","PICHINCHA")` |
| J | DESTINATARIO | 2–255 | 254 | 1 | `J2`: `=IF(DATOS!$A2="","",IF(DATOS!$A2="ALEXIS",CONCATENATE(DATOS!$G2," / RUTA - ",DATOS!$W2),DATOS!$G2))` |
| K | DIRECCION_DEST | 2–255 | 254 | 1 | `K2`: `=IF(DATOS!$D2="","",DATOS!$D2)` |
| L | PISO_DEST | 2–255 | 254 | 1 | `L2`: `=IF($J2="","","1")` |
| M | DEPTO_DEST | 2–255 | 254 | 1 | `M2`: `=IF($J2="","","1")` |
| N | CP_DEST | 2–255 | 254 | 1 | `N2`: `=IF(DATOS!$AB2="","",DATOS!$AB2)` |
| O | LOCALIDAD_DEST | 2–255 | 254 | 1 | `O2`: `=IF(DATOS!$P2="","",DATOS!$P2)` |
| P | PROVINCIA_DEST | 2–255 | 254 | 1 | `P2`: `=IF(DATOS!$O2="","",DATOS!$O2)` |
| Q | NRO_REFERENCIA | 2–255 | 254 | 1 | `Q2`: `=IF(DATOS!$B2="","",DATOS!$B2)` |
| R | NRO_FACTURA | 2–255 | 254 | 1 | `R2`: `=IF(DATOS!$B2="","",DATOS!$B2)` |
| S | VALOR TOTAL FACTURA | 2–255 | 254 | 1 | `S2`: `=IF($R2="","",IF(DATOS!$T2=0,4.7,DATOS!$T2))` |
| W | TELEFONO_MOVIL | 2–255 | 254 | 1 | `W2`: `=IF(DATOS!$K2="","",CONCATENATE(0,SUBSTITUTE(DATOS!$K2,"-","")))` |
| X | Telefono_FIJO | 2–255 | 254 | 1 | `X2`: `=IF(DATOS!$L2="","",CONCATENATE(0,SUBSTITUTE(DATOS!$L2,"-","")))` |
| Y | Email | 2–255 | 254 | 1 | `Y2`: `=IF(DATOS!$M2="","servicioalcliente.ec@royalprestige.com",DATOS!$M2)` |
| Z | FECHA_COMPRA | 2–255 | 254 | 1 | `Z2`: `=IF($J2="","",TODAY()+1)` |
| AC | CANTIDAD | 2–255 | 254 | 1 | `AC2`: `=IF(DATOS!$R2="","",DATOS!$R2)` |
| AD | PESO_KG | 2–255 | 254 | 1 | `AD2`: `=IF(DATOS!$S2="","",DATOS!$S2)` |
| AH | OPERACION | 2–255 | 254 | 1 | `AH2`: `=IF($J2="","","DESPACHO")` |
| AI | TIPO_SERVICIO | 2–255 | 254 | 1 | `AI2`: `=IF($J2="","","NORMAL")` |

## Hoja `TRAMACO` (oculta, rango A1:AG501, 8000 fórmulas)


| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| C | APELLIDOS O RAZON SOCIAL | 2–501 | 500 (matriz dinámica) | 1 | `C2`: `=IF($AG2="","",INDEX(DATOS!$G$2:$G$500,$AG2))` |
| D | NOMBRES | 2–501 | 500 | 1 | `D2`: `=IF($C2="","",".")` |
| F | PROVINCIA | 2–501 | 500 (matriz dinámica) | 1 | `F2`: `=IF($AG2="","",INDEX(DATOS!$O$2:$O$500,$AG2))` |
| G | CIUDAD | 2–501 | 500 (matriz dinámica) | 1 | `G2`: `=IF($AG2="","",INDEX(DATOS!$P$2:$P$500,$AG2))` |
| H | LACALIDAD | 2–501 | 500 (matriz dinámica) | 1 | `H2`: `=IF($AG2="","",INDEX(DATOS!$Q$2:$Q$500,$AG2))` |
| I | CALLE PRIMARIA | 2–501 | 500 (matriz dinámica) | 1 | `I2`: `=IF($AG2="","",INDEX(DATOS!$D$2:$D$500,$AG2))` |
| J | NUMERO | 2–501 | 500 | 1 | `J2`: `=IF($C2="","",".")` |
| K | CALLE SECUNDARIA | 2–501 | 500 | 1 | `K2`: `=IF($C2="","",".")` |
| L | REFERENCIA | 2–501 | 500 | 1 | `L2`: `=IF($C2="","",".")` |
| N | TELEFONO 1 | 2–501 | 500 (matriz dinámica) | 1 | `N2`: `=IF($AG2="","",INDEX(DATOS!$K$2:$K$500,$AG2))` |
| Q | PRODUCTO | 2–501 | 500 | 1 | `Q2`: `=IF($C2="","","CARGA LIVIANA")` |
| R | PESO | 2–501 | 500 (matriz dinámica) | 1 | `R2`: `=IF($AG2="","",INDEX(DATOS!$S$2:$S$500,$AG2))` |
| W | NUMERO CAJAS | 2–501 | 500 (matriz dinámica) | 1 | `W2`: `=IF($AG2="","",IF(INDEX(DATOS!$R$2:$R$500,$AG2)="-",1,INDEX(DATOS!$R$2:$R$500,$AG2)))` |
| AB | DESCRIPCION CONTENIDO | 2–501 | 500 | 1 | `AB2`: `=IF($C2="","","TAPAS Y OLLAS")` |
| AD | OBSERVACIONES | 2–501 | 500 (matriz dinámica) | 1 | `AD2`: `=IF($AG2="","",INDEX(DATOS!$B$2:$B$500,$AG2))` |
| AG | FILA_DATOS | 2–501 | 500 | 1 | `AG2`: `=IFERROR(MATCH(ROW()-1,DATOS!$AC$2:$AC$500,0),"")` |

## Hoja `DESPACHOS` (oculta, rango A1:Q501, 8000 fórmulas)

- Formato condicional: 1 reglas

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| A | Fecha del Pedido | 2–501 | 500 | 2 | `A2`: `=IF($B2="","",TODAY()*1)` |
| B | No. Pedido HCE | 2–501 | 500 | 1 | `B2`: `=IF(DATOS!$B2="","",DATOS!$B2)` |
| C | Nombre del Cliente | 2–501 | 500 | 1 | `C2`: `=IF(DATOS!$G2="","",DATOS!$G2)` |
| D | Teléfono | 2–501 | 500 | 1 | `D2`: `=IF(DATOS!$K2="","",DATOS!$K2)` |
| F | Dirección | 2–501 | 500 | 1 | `F2`: `=IF(DATOS!$D2="","",DATOS!$D2)` |
| G | Provincia | 2–501 | 500 | 1 | `G2`: `=IF(DATOS!$O2="","",DATOS!$O2)` |
| H | Ciudad | 2–501 | 500 | 1 | `H2`: `=IF(DATOS!$P2="","",DATOS!$P2)` |
| I | 1° Localidad | 2–501 | 500 | 1 | `I2`: `=IF(DATOS!$Q2="","",DATOS!$Q2)` |
| J | Total Unidades | 2–501 | 500 | 1 | `J2`: `=IF($B2="","",IFERROR(VLOOKUP($B2,'TABLAS DINAMICAS'!$B:$D,2,FALSE),"-"))` |
| K | Total Items | 2–501 | 500 | 1 | `K2`: `=IFERROR(IF($B2="","",VLOOKUP($B2,'TABLAS DINAMICAS'!$B:$F,5,FALSE)),"-")` |
| L | Total Bultos | 2–501 | 500 | 1 | `L2`: `=IF(DATOS!$R2="","",DATOS!$R2)` |
| M | Total Pesos KG | 2–501 | 500 | 1 | `M2`: `=IF(DATOS!$S2="","",DATOS!$S2)` |
| N | Valor del Producto | 2–501 | 500 | 1 | `N2`: `=IF($B2="","",IFERROR(ROUND(VLOOKUP($B2,'TABLAS DINAMICAS'!$B:$F,4,FALSE),2),"-"))` |
| O | BOX DENSITY | 2–501 | 500 | 1 | `O2`: `=IF($B2="","",IFERROR(ROUND(SUMIF('TABLAS DINAMICAS'!$P:$P,$B2,'TABLAS DINAMICAS'!$Q:$Q),2),"-"))` |
| P | Transporte | 2–501 | 500 | 1 | `P2`: `=IF($B2="","",IF(VLOOKUP($B2,CHOOSE({2,1},DATOS!$A:$A,DATOS!$B:$B),2,0)="PRO","TRAMACO","FLEXNET"))` |
| Q | No. Guia Transportadora | 2–501 | 500 | 1 | `Q2`: `=IF(DATOS!$U2="","",DATOS!$U2)` |

## Hoja `ETIQUETAS` (visible, rango A1:L501, 5500 fórmulas)

- Formato condicional: 1 reglas
- Validación de datos `F2:F501`: list "-,OK"

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| A | # PEDIDOS | 2–501 | 500 | 1 | `A2`: `=IF(DATOS!$B2="","",DATOS!$B2)` |
| B | NOMBRES | 2–501 | 500 | 1 | `B2`: `=IFERROR(IF($G2="","",CONCATENATE(IF($H2="-","",LEFT($G2,$H2-1))&" "&IF($I2="-","",MID($G2,$H2+1,$I2-$H2-1)))),"")` |
| C | APELLIDOS | 2–501 | 500 | 1 | `C2`: `=IFERROR(IF($G2="","",IF($J2="-","",CONCATENATE(IF($J2="-","",MID($G2,$I2+1,$J2-$I2-1))&" "&IF($K2="-","",MID($G2,$J2+1,$K2-$J2-1))&" "&IF($K2="-","",RIGHT($G2,LEN($G2)-$K2))))),"")` |
| D | DESTINOS | 2–501 | 500 | 1 | `D2`: `=IF(DATOS!$A2="","",DATOS!$A2)` |
| E | RUTAS | 2–501 | 500 | 1 | `E2`: `=IF(DATOS!$A2="ALEXIS",DATOS!$W2,"")` |
| G | COMPLEMENTOS | 2–501 | 500 | 1 | `G2`: `=CONCATENATE(IF(DATOS!$G2="","",DATOS!$G2)&" ")` |
| H |  | 2–501 | 500 | 1 | `H2`: `=IF($A2="","",IFERROR(FIND(" ",$G2),"-"))` |
| I |  | 2–501 | 500 | 1 | `I2`: `=IF($A2="","",IFERROR(FIND(" ",$G2,$H2+1),"-"))` |
| J |  | 2–501 | 500 | 1 | `J2`: `=IF($A2="","",IFERROR(FIND(" ",$G2,$I2+1),"-"))` |
| K |  | 2–501 | 500 | 1 | `K2`: `=IF($A2="","",IFERROR(FIND(" ",$G2,$J2+1),"-"))` |
| L |  | 2–501 | 500 | 1 | `L2`: `=IF($A2="","",IFERROR(FIND(" ",$G2,$K2+1),"-"))` |

## Hoja `ETIQUETAS ZEBRA` (visible, rango A1:U800, 10396 fórmulas)


| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| A | # | 2–800 | 799 | 1 | `A2`: `=IF(ROW()-1>MAX(DATOS!$X$2:$X$500),"",ROW()-1)` |
| B | PEDIDO | 2–800 | 799 (matriz dinámica) | 1 | `B2`: `=IF($A2="","",INDEX(DATOS!$B$2:$B$500,$N2))` |
| C | DESTINATARIO | 2–800 | 799 (matriz dinámica) | 1 | `C2`: `=IF($A2="","",INDEX(DATOS!$G$2:$G$500,$N2))` |
| D | DIRECCION | 2–800 | 799 (matriz dinámica) | 1 | `D2`: `=IF($A2="","",INDEX(DATOS!$D$2:$D$500,$N2))` |
| E | PROVINCIA | 2–800 | 799 (matriz dinámica) | 1 | `E2`: `=IF($A2="","",INDEX(DATOS!$O$2:$O$500,$N2))` |
| F | CANTON | 2–800 | 799 (matriz dinámica) | 1 | `F2`: `=IF($A2="","",INDEX(DATOS!$P$2:$P$500,$N2))` |
| G | PARROQUIA | 2–800 | 799 (matriz dinámica) | 1 | `G2`: `=IF($A2="","",INDEX(DATOS!$Q$2:$Q$500,$N2))` |
| H | COURIER | 2–800 | 799 (matriz dinámica) | 1 | `H2`: `=IF($A2="","",INDEX(DATOS!$Y$2:$Y$500,$N2))` |
| I | BULTO | 2–800 | 799 (matriz dinámica) | 1 | `I2`: `=IF($A2="","",$A2-IF($N2=1,0,INDEX(DATOS!$X$2:$X$500,$N2-1)))` |
| J | DE | 2–800 | 799 (matriz dinámica) | 1 | `J2`: `=IF($A2="","",INDEX(DATOS!$R$2:$R$500,$N2))` |
| K | PESO KG | 2–800 | 799 (matriz dinámica) | 1 | `K2`: `=IF($A2="","",INDEX(DATOS!$S$2:$S$500,$N2))` |
| L | TELEFONO | 2–800 | 799 (matriz dinámica) | 1 | `L2`: `=IF($A2="","",INDEX(DATOS!$K$2:$K$500,$N2))` |
| N | FILA AUX | 2–800 | 799 | 1 | `N2`: `=IF($A2="","",COUNTIF(DATOS!$X$2:$X$500,"<"&$A2)+1)` |
| P | ETIQUETA # | 3–14 | 8 (matriz dinámica) | 8 | `P3`: `=IFERROR(INDEX($H$2:$H$800,$Q$1),"")` |
| R |  | 1–1 | 1 | 1 | `R1`: `="de "&MAX(DATOS!$X$2:$X$500)&" bultos"` |

## Hoja `COBERTURAS Y TARIFAS` (visible, rango A1:AF1804, 1303 fórmulas)

- Formato condicional: 5 reglas

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| AA | COBERT | 2–1304 | 1303 | 1 | `AA2`: `=IF($X2="","",_xlfn.TEXTJOIN("_",TRUE,$X2,$Y2,$Z2))` |

## Hoja `ITEMS APIS` (visible, rango A1:N581, 1160 fórmulas)

- Tabla **ITEMS_API** `A1:K581` (tabla de consulta Power Query)
  - columna calculada `PEDIDOS`: `=IF($B2="","",IFERROR(VALUE($B2),$B2))`
  - columna calculada `PESO`: `=IF($J2="","",VLOOKUP($F2,DATA_CODIGOS,8,FALSE)*$G2)`

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| J | PEDIDOS | 2–581 | 580 | 1 | `J2`: `=IF($B2="","",IFERROR(VALUE($B2),$B2))` |
| K | PESO | 2–581 | 580 | 1 | `K2`: `=IF($J2="","",VLOOKUP($F2,DATA_CODIGOS,8,FALSE)*$G2)` |

## Hoja `ITEMS DEPOT` (visible, rango A1:K1291, 5160 fórmulas)

- Tabla **Estado** `A1:K1291` (tabla de consulta Power Query)
  - columna calculada `PEDIDO`: `=IFERROR(IF($C2="","",VALUE($C2)),$C2)`
  - columna calculada `PESO`: `=_xlfn.IFNA(IF($C2="","",VLOOKUP($D2,DATA_CODIGOS,8,FALSE)*$F2),"")`
  - columna calculada `VOLUMEN`: `=_xlfn.IFNA(IF($C2="","",VLOOKUP($D2,DATA_CODIGOS,7,FALSE)*$F2),"")`
  - columna calculada `PRECIO`: `=IFERROR(IF($C2="","",VLOOKUP($D2,DATA_CODIGOS,11,FALSE)*$F2),"Verificar")`
- Formato condicional: 1 reglas

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| H | PEDIDO | 2–1291 | 1290 | 1 | `H2`: `=IFERROR(IF($C2="","",VALUE($C2)),$C2)` |
| I | PESO | 2–1291 | 1290 | 1 | `I2`: `=_xlfn.IFNA(IF($C2="","",VLOOKUP($D2,DATA_CODIGOS,8,FALSE)*$F2),"")` |
| J | VOLUMEN | 2–1291 | 1290 | 1 | `J2`: `=_xlfn.IFNA(IF($C2="","",VLOOKUP($D2,DATA_CODIGOS,7,FALSE)*$F2),"")` |
| K | PRECIO | 2–1291 | 1290 | 1 | `K2`: `=IFERROR(IF($C2="","",VLOOKUP($D2,DATA_CODIGOS,11,FALSE)*$F2),"Verificar")` |

## Hoja `EMPAQUETADO` (visible, rango A1:R117, 696 fórmulas)

- Tabla **EMPAQUETADO** `B1:K117` (tabla de consulta Power Query)
  - columna calculada `NRO. DE CONTENEDORA`: `=IF($C2="","",IF($E2<>"",$E2,_xlfn.XLOOKUP(TEXT($C2,"@"),Estado[DOC_EXT],Estado[NRO_CONTENEDORA_EMPAQUE],IF(COUNTIF(#REF!,TEXT($C2,"@"))>0,"validar","sin pedido"),0)))`
  - columna calculada `BULTOS`: `=IF($C2="","",1)`
  - columna calculada `PESO CAJA`: `=IF($C2="","",VLOOKUP($D2,DATA_CAJAS,5,FALSE)+0.1)`
  - columna calculada `VOL. CAJA`: `=IF($C2="","",VLOOKUP($D2,DATA_CAJAS,6,FALSE))`
  - columna calculada `VOL. ITEMS`: `=IF($C2="", "", IF(SUMIF('TABLAS DINAMICAS'!$M:$M, $F2, 'TABLAS DINAMICAS'!$N:$N)>0, SUMIF('TABLAS DINAMICAS'!$M:$M, $F2, 'TABLAS DINAMICAS'!$N:$N), IF(SUMIF('TABLAS DINAMICAS'!$A:$A, $C2, 'TABLAS DINAMICAS'!$E:$E)>0, SUMIF('TABLAS DINAMICAS'!$A:$A, $C2, 'TABLAS DINAMICAS'!$E:$E), "Verificar")))`
  - columna calculada `PORCENTAJE`: `=IF($C2="","",IF(OR($D2="F1",$D2="F5",$D2="F11",$D2="F10",$D2="SOBRE 1"),100%,IF($J2="Verificar","Verificar",IF($J2/$I2>100%,100%,IF($J2/$I2<30%,30%,$J2/$I2)))))`
- Formato condicional: 11 reglas

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| F | NRO. DE CONTENEDORA | 2–117 | 116 | 1 | `F2`: `=IF($C2="","",IF($E2<>"",$E2,_xlfn.XLOOKUP(TEXT($C2,"@"),Estado[DOC_EXT],Estado[NRO_CONTENEDORA_EMPAQUE],IF(COUNTIF(#REF!,TEXT($C2,"@"))>0,"validar","sin pedido"),0)))` |
| G | BULTOS | 2–117 | 116 | 1 | `G2`: `=IF($C2="","",1)` |
| H | PESO CAJA | 2–117 | 116 | 1 | `H2`: `=IF($C2="","",VLOOKUP($D2,DATA_CAJAS,5,FALSE)+0.1)` |
| I | VOL. CAJA | 2–117 | 116 | 1 | `I2`: `=IF($C2="","",VLOOKUP($D2,DATA_CAJAS,6,FALSE))` |
| J | VOL. ITEMS | 2–117 | 116 | 1 | `J2`: `=IF($C2="", "", IF(SUMIF('TABLAS DINAMICAS'!$M:$M, $F2, 'TABLAS DINAMICAS'!$N:$N)>0, SUMIF('TABLAS DINAMICAS'!$M:$M, $F2, 'TABLAS DINAMICAS'!$N:$N), IF(SUMIF('TABLAS DINAMICAS'!$A:$A, $C2, 'TABLAS DINAMICAS'!$E:$E)>0, SUMIF('TABLAS DINAMICAS'!$A:$A, $C2, 'TABLAS DINAMICAS'!$E:$E), "Verificar")))` |
| K | PORCENTAJE | 2–117 | 116 | 1 | `K2`: `=IF($C2="","",IF(OR($D2="F1",$D2="F5",$D2="F11",$D2="F10",$D2="SOBRE 1"),100%,IF($J2="Verificar","Verificar",IF($J2/$I2>100%,100%,IF($J2/$I2<30%,30%,$J2/$I2)))))` |

## Hoja `TABLAS DINAMICAS` (visible, rango A1:Q548, 545 fórmulas)


| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| K | TOTAL PESO DEL PEDIDO | 4–548 | 545 | 1 | `K4`: `=IF($H4="", "", SUMIF($B:$B, $H4, $D:$D) + $J4)` |

## Hoja `Temp_Tramaco` (oculta, rango A1:J501, 0 fórmulas)


Sin fórmulas en celdas.

## Hoja `COD POS` (oculta, rango A1:J1305, 7818 fórmulas)


| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| C |  | 3–1305 | 1303 | 1 | `C3`: `=IF(AND(ISNUMBER(SEARCH("(",E3)),RIGHT(E3,1)=")"),LEFT(E3,SEARCH("(",E3)-2),E3)` |
| D |  | 3–1305 | 1303 | 1 | `D3`: `=+CONCATENATE(A3,B3,C3)` |
| G |  | 3–1305 | 1303 | 1 | `G3`: `=+A3&B3` |
| H |  | 3–1305 | 1303 | 1 | `H3`: `=+_xlfn.TEXTJOIN("",TRUE,G3,E3)` |
| I |  | 3–1305 | 1303 | 1 | `I3`: `=+_xlfn.XLOOKUP(G3,[1]!Consulta2[TMS PRO-LOC],[1]!Consulta2[TMS PRO-LOC])` |
| J |  | 3–1305 | 1303 | 1 | `J3`: `=+_xlfn.XLOOKUP(H3,[1]!Consulta1[latitud],[1]!Consulta1[latitud])` |

## Hoja `DATA CODIGO Y CAJAS` (oculta, rango A1:AF517, 623 fórmulas)

- Formato condicional: 185 reglas
- Validación de datos `J2:J292 J295 J321:J323 J326:J330 J333:J362 J437`: list "-,PALLET ESTANDAR,PAPELERIA,ACCESORIOS"

| Col | Encabezado (fila 1) | Filas | Celdas | Variantes | Fórmula (primera celda) |
|---|---|---|---|---|---|
| H | VOL_M3 | 2–516 | 515 | 1 | `H2`: `=$E2*$F2*$G2` |
| T | Volumen | 2–36 | 35 | 2 | `T2`: `=+P2*Q2*R2` |
| U | Largo | 11–36 | 16 | 2 | `U11`: `=+P11/100` |
| V | Ancho | 11–36 | 16 | 2 | `V11`: `=+Q11/100` |
| W | Alto | 11–36 | 16 | 2 | `W11`: `=+R11/100` |
| Y | VOLUMEN | 2–36 | 25 | 2 | `Y2`: `=+U2*V2*W2` |

