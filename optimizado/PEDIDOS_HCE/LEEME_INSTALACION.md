# PEDIDOS HCE: versión final (VBA) - Etapa 1 del sistema de egresos

Esta versión mantiene **los mismos nombres de macros, las mismas columnas A:U de DEPOT y el mismo envío C:N a
`DATOS`** del archivo EGR. Todo lo nuevo se agrega aparte: columnas V:AD, un panel nuevo e historial en CAMBIOS.

## 1. Instalación (unos 10 minutos)

1. **Respalda** el archivo: *Guardar como* `PEDIDOS HCE - respaldo.xlsm`.
2. Abre el editor con `Alt+F11`.
3. Reemplaza el código existente:

   | Dónde | Qué hacer |
   |---|---|
   | **Módulo1** | Selecciona todo el código (`Ctrl+A`), bórralo y pega [`Modulo1.bas`](Modulo1.bas) |
   | **frmValidar** (clic derecho › *Ver código*) | Borra todo y pega [`frmValidar.frm`](frmValidar.frm). El diseño del formulario no se toca: los controles se crean solos |
   | **ThisWorkbook** | Borra todo y pega [`ThisWorkbook.cls`](ThisWorkbook.cls) |

4. **Crea el panel nuevo**: *Insertar › UserForm*. En Propiedades pon `(Name)` = **frmPanel**. Abre su código, borra lo
   que traiga y pega [`frmPanel.frm`](frmPanel.frm). **Este paso es obligatorio**, porque Módulo1 usa `frmPanel`.
5. Ejecuta *Depuración › Compilar VBAProject*. No debe aparecer ningún error; antes fallaba por `frmRevisar`.
6. Guarda, cierra y vuelve a abrir el libro. La pestaña **Complementos** muestra la barra nueva.

> No pegues líneas que empiecen con `Attribute`: dan error de sintaxis.
> Para volver atrás, basta con el respaldo del paso 1.

## 2. Flujo diario (barra en Complementos, o botón **Panel HYCITE**)

| Paso | Botón | Qué hace |
|---|---|---|
| 0 | **0 Actualizar datos** | Refresca las consultas DEPOT y TMS sin ir a *Datos*, y ofrece limpiar la validación anterior |
| 1 | **1 Limpiar** | Borra L y el panel N:AD |
| 2 | **2 Validar** | Valida y rellena N:U, la sigla L y las columnas nuevas V:AD. Al terminar pregunta: *"¿Revisar ahora los N casos pendientes?"*. **Sí** abre el validador solo con los REVISAR |
| 3 | **3 Revisar pendientes** | `frmValidar` recorre solo los REVISAR |
| 4 | **4 Aprobar lote** | Aprueba en bloque las sugerencias restantes (avisa cuántas son fuera de cobertura) |
| 5 | **5 Aplicar aprobados** | Copia la propuesta N:P a F:H y **agrega** los cambios al historial CAMBIOS |
| 6 | **6 Actualizar siglas** | Recalcula L y V:AB si se editó F:H a mano |
| 7 | **7 Enviar a EGR** | Mismo envío C:N a `DATOS`. Avisa si quedan pedidos REVISAR o correcciones sin aplicar |
| — | **Exportar CAMBIOS** | Guarda un `.xlsx` con los cambios de hoy o con todo el historial |
| — | **Panel HYCITE** | Interfaz principal: ver la sección 2.1 |

"Filtrar día" se quitó de la barra porque la consulta DEPOT no trae columna FECHA: nunca funcionó. La macro sigue en el
código.

### 2.1 Panel HYCITE (interfaz principal)

- **Pasos 0 a 7 en columna, a la izquierda**, en orden.
  - Al pasar el mouse sobre cada botón se explica qué hace.
  - El recuadro **SIGUIENTE PASO** indica qué toca hacer y resalta en naranja ese botón (`>> paso <<`). Lo calcula
    según el estado real: sin pedidos, sin validar, con REVISAR, correcciones sin aplicar o listo para enviar.
- **Funciona a la par con Excel**: la ventana no bloquea la hoja.
  - "Seguir fila en Excel" selecciona la fila del registro elegido.
  - Al volver del validador, la lista se actualiza sola.
- **Vistas**:
  - **PEDIDOS**.
  - **COBERTURA**: permite buscar y filtrar cualquier provincia, cantón o parroquia, con sigla, gestor Q, sugerido R y
    trayecto S.
  - **ZONAS PELIGROSAS**: barrio, sector y punto TMC.
- **Encabezados alineados** sobre cada columna. Hay 12 columnas visibles, calculadas para que no aparezca la barra
  horizontal; el resto de datos se ve en el **detalle**.
- **Filtros**:
  - **Filtro rápido**: con señal, REVISAR, APROBADO, OK, zona peligrosa, cobertura cercana, corregidos, sin aplicar,
    sin validar y destino PRO/UIO/GYE/GPS.
  - **Buscar**: busca en todas las columnas mientras se escribe.
  - **Filtro 1** y **Filtro 2**: eligen un campo (pedido, fila, provincia, cantón, parroquia, gestor, sugerido R…) y un
    texto.
  - **Orden** por cualquier campo, ascendente o descendente.
- **Columna SEÑAL** (en orden de prioridad):

  | Señal | Significado |
  |---|---|
  | `!! ZONA` | Zona peligrosa |
  | `! SECTOR` | Verificar sector |
  | `? REVISAR` | El pedido está en REVISAR |
  | `~ CERCANA` | Se asignó una cobertura cercana (cabecera o ciudad principal/secundaria) |
  | `+ APROBADO` | Aprobado por el operario o en lote |
  | `* CORREGIDO` | El sistema corrigió el dato del cliente |
  | `- SIN VALIDAR` | Todavía no se validó |
  | `OK` | Sin cambios |

- **Detalle** del registro seleccionado:
  - dirección;
  - **dato original del cliente → propuesta** y si ya se aplicó a F:H;
  - tipo de corrección;
  - gestor asignado frente a **gestor de cobertura (Q)** y **sugerido (R)**, indicando si coinciden;
  - trayecto (S), tipo de entrega y zona peligrosa con su punto TMC, **en rojo**.
- **Tamaño**: 1000 × 600 puntos. Si la pantalla es más chica, se reduce proporcionalmente (hasta el 60 %) para no
  salirse.
- Doble clic en un pedido lo abre en el validador; en COBERTURA o ZONAS, lleva a la fila de la hoja.

### 2.2 Columnas de DEPOT

| Columnas | Contenido |
|---|---|
| A:K | Datos del cliente |
| L | Sigla |
| N:U | Panel de validación (igual que antes) |
| V | GESTOR_ASIGNADO |
| W | DESTINO |
| X | TRAYECTO_TRAMACO |
| Y | TIPO_ENTREGA |
| Z | ZONA_PELIGROSA |
| AA | GESTOR_COBERTURA (Q) |
| AB | GESTOR_SUGERIDO (R) |
| AC | ORIGINAL_CLIENTE (se conserva aunque se aplique la corrección) |
| AD | TIPO_CORRECCION: SIN CAMBIO, CORREGIDO SISTEMA, COBERTURA CERCANA, CORRECCION DE ESCRITURA, DMQ → CALDERON (QUI), MANUAL OPERARIO… |

CAMBIOS registra además TIPO_CORRECCION en la columna Q.

## 3. Reglas implementadas

### 3.1 Gestor y destino (columnas V `GESTOR_ASIGNADO` y W `DESTINO`)

| Destino final (cantón) | DESTINO | Gestor |
|---|---|---|
| QUITO (Pichincha) | UIO | ITSANET o LAAR COURIER, según la columna *GESTOR DE ENTREGAS* de COBERTURA (si no dice ninguno, ITSANET) |
| GUAYAQUIL (Guayas) | GYE | igual que Quito |
| Galápagos | GPS | el gestor de COBERTURA (TRAMACO si está vacío) |
| Cualquier otro cantón | **PRO** | **GESTOR SUGERIDO A COBERTURA**; si está vacío, **TRAMACO** |

La columna X `TRAYECTO_TRAMACO` traduce el código de *TRAYECTO(TRAMACO)*: `CP` = CIUDAD PRINCIPAL,
`CS` = CIUDAD SECUNDARIA, `TE` = TRAYECTO ESPECIAL y `TD` = TRAYECTO DIFERENCIADO.

Resultado emulado con los 169 pedidos del archivo recibido: **PRO/TRAMACO 108**, **UIO/ITSANET 42**,
**GYE/ITSANET 19**.

La regla es **por cantón**, no por provincia. Samborondón, Daule o Rumiñahui salen **PRO**, aunque la fórmula actual de
EGR (`DATOS!A`) los marca GYE o UIO por provincia. Lo ajustamos cuando trabajemos el archivo EGR: la columna W ya tiene el
valor listo para enviarse.

### 3.2 Fuera de cobertura: sugerencia automática

Se aplica en este orden, y la primera que existe es la sugerida:

1. **Misma parroquia mal escrita** (diferencia de 1 o 2 letras; ej. CHUQUILPE → CHIGUILPE).
2. **Parroquia principal del cantón**: la cabecera cantonal, es decir, el mismo nombre del cantón.
3. **Ciudad principal** (`CP`), y luego **ciudad secundaria** (`CS`), del **mismo cantón**.
4. **Ciudad principal o secundaria mencionada en la dirección**. Es la cercanía que se puede deducir del texto.
5. **Ciudad principal**, y luego **secundaria**, de la **provincia**.
6. **Capital provincial**. Es el respaldo para las provincias sin `CP`/`CS` en COBERTURA: hoy CAÑAR, MORONA SANTIAGO,
   NAPO, PASTAZA, ORELLANA, SUCUMBÍOS, SANTA ELENA y GALÁPAGOS.
7. Como último recurso, la parroquia más parecida del cantón.

El pedido queda en **REVISAR** con la acción *"Fuera de cobertura: confirmar sugerida"*. La columna SUGERENCIAS muestra
`PARROQUIA [CANTÓN] (motivo)`. Ya no se aceptan parecidos lejanos: antes pasaba, por ejemplo, PUNZARA → QUINARA.

> **Cercanía geográfica real.** COBERTURA no tiene coordenadas, así que la cercanía se deduce del cantón y de la
> dirección. Si se agregan columnas `LATITUD` y `LONGITUD`, se puede sugerir por distancia real.

### 3.3 Regla Quito / DMQ (se mantiene)

Una dirección de Quito sin parroquia reconocible queda como **DISTRITO METROPOLITANO DE QUITO**, sigla
**CALDERON (QUI)** y acción *Confirmar Quito/DMQ*.

### 3.4 Cantón del cliente respetado (corrección de un error)

Si la parroquia existe en varios cantones, ahora se usa el cantón que escribió el cliente, o el sufijo de su parroquia
("(TUL)", "(MAN)", "(IBA)"). Si no se puede decidir, el pedido queda en **REVISAR: "Verificar cantón"**. Antes elegía el
primer cantón y lo marcaba OK; así pasaron TULCAN→MONTUFAR, MANTA→CHONE e IBARRA→COTACACHI.

### 3.5 Zonas peligrosas (columnas Y `TIPO_ENTREGA` y Z `ZONA_PELIGROSA`)

- Se lee la hoja **ZONAS PELIGROSAS**, detectando las columnas por su encabezado: PROVINCIA, CIUDAD, PARROQUIA,
  ZONA PELIGROSA, SECTOR, PUNTO DE ATENCION MAS CERCANO, su DIRECCIÓN y VALIDACION. Las filas con VALIDACION = "NO" se
  ignoran.
- La provincia admite nombres cortos ("SANTO DOMINGO") y la ciudad debe coincidir con el cantón del pedido o aparecer en
  la dirección.
- Un pedido cae en zona peligrosa en estos casos:

  | Caso | Condición | Resultado |
  |---|---|---|
  | Barrio | El barrio (ZONA) o un sector específico aparece en la dirección. Se aceptan prefijos como BARRIO, CDLA o COOP | Con destino PRO, `TIPO_ENTREGA` = **C.O.D - RETIRO OFICINA** + el punto indicado (ej. *TMC PORTOVIEJO*). Z muestra el punto y su dirección |
  | Parroquia completa | La zona es la propia parroquia y no tiene sector (ej. PUERTO LOPEZ, JOSE LUIS TAMAYO) | Igual que el barrio |
  | Sector de la parroquia | La zona es la parroquia pero con un sector direccional (SUR, NORTE…) | **POSIBLE ZONA PELIGROSA**, `TIPO_ENTREGA` = *VERIFICAR SECTOR*. No se asigna C.O.D, porque no se sabe si la dirección está en ese sector |

- Los sectores genéricos (NORTE, SUR, CENTRO, SUBURBIO o el nombre de un cantón) no se usan solos para detectar zonas;
  así se evitan falsos positivos.
- Se marca automáticamente, sin editar a mano. Si prefieres que además quede en REVISAR, cambia en Módulo1
  `ZONA_PELIGROSA_A_REVISAR = True`.
- Emulación con los 169 pedidos del archivo:
  - **5 zonas peligrosas**: JOSE LUIS TAMAYO ×2 (TMC LIBERTAD), CARIGÁN (TMC LOJA), ALPACHACA (TMC IBARRA) y el barrio
    4 DE NOVIEMBRE de Manta (TMC MANTA);
  - **6 avisos de "verificar sector"**: sector SUR de ANDRES DE VERA ×5 y de PICOAZA.

### 3.6 Otros cambios

- **CAMBIOS** ya no se borra: acumula el historial.
  - Columnas nuevas: ORIGEN (SISTEMA / OPERARIO / LOTE), USUARIO, SIGLA, GESTOR y DESTINO.
  - Botón **Exportar CAMBIOS** para sacar el día o el historial completo.
- **CAÑAR / CANAR**: la provincia se aplica y se envía como `CANAR`, que es como la espera el archivo EGR. Antes había que
  corregirla a mano en DATOS.
- **frmValidar**:
  - listas dependientes Provincia › Cantón › Parroquia, con todas las parroquias del cantón;
  - al elegir una sugerencia `PARROQUIA [CANTÓN]` también se ajusta el cantón;
  - botón **Sugerir cobertura cercana**;
  - vista previa de sigla, gestor, destino, trayecto y zona;
  - aviso si se intenta aplicar algo fuera de cobertura.
- **Enviar a EGR**: si hay dos libros EGR abiertos, prefiere el que no se llama "copia".
- La búsqueda por parecido de texto es más rápida: ya no llama a `Application.Min`.

## 4. Prueba recomendada antes de pasar a producción

1. En el respaldo, con los pedidos del día: **1 Limpiar** › **2 Validar**. Anota los totales OK y REVISAR.
2. Comprueba estos casos:
   - Un pedido de Quito debe dar UIO con ITSANET o LAAR.
   - Uno de Guayaquil debe dar GYE.
   - Uno de otra provincia debe dar PRO con TRAMACO.
   - Uno en zona peligrosa debe llevar C.O.D - RETIRO OFICINA TMC.
   - Un pedido con parroquia que no existe debe mostrar "Fuera de cobertura" y una sugerencia en el mismo cantón.
3. Responde **Sí** al aviso final: el validador debe abrirse solo con los REVISAR.
4. Ejecuta **5 Aplicar aprobados** y revisa que CAMBIOS **agregue** filas, sin borrar las anteriores.
5. Con EGR abierto, ejecuta **7 Enviar a EGR**. DATOS debe llenarse igual que antes (C:N).

## 5. Pendiente para la etapa del archivo EGR

- Usar W (DESTINO) y V (GESTOR) en lugar de la fórmula de `DATOS!A`, que decide por provincia. Esto requiere ampliar el
  envío a DATOS; se hará junto con los cambios del archivo EGR.
- Unificar el maestro de cobertura: la COBERTURA de este archivo pasa a ser la fuente única.
