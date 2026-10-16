# Correo de entrega v1.1 — listo para enviar

**Para:** Cristina Suasnavas (Ejecutiva de cuenta, HY CITE) · Jimmy Celi (Distribución) · Adriana Morillo (Operadora de sistemas)

**Asunto:** Entrega v1.1 del sistema de despacho + acuerdos de la reunión de hoy

**Adjuntos:** `Sistema_Despacho_HYCITE_v1.1.zip` (PEDIDOS HCE.xlsm y Formato EGR_FL_HYCITE.xlsm)

---

Buenas tardes, Cristina, Jimmy y Adriana:

Gracias por el tiempo de esta mañana. Les comparto la versión **v1.1** del sistema de despacho, que es la que quedó operativa hoy, y el resumen de lo que acordamos para la siguiente versión.

## Manual interactivo

Antes de los detalles, este es el enlace más útil del correo:

**👉 Manual del operador (guía interactiva):** https://claude.ai/artifact/CaCKUphWQKnBGwjXRoi5jc

Es una página web con el flujo completo, paso a paso: la lista de control del día, un simulador donde se puede escribir una dirección y ver qué destino sugiere el sistema y por qué, el detalle de las etiquetas y una sección de "si algo falla" con las soluciones que no requieren entrar al código. Adriana, está escrito pensando en la operación diaria; Jimmy, la parte de reglas y reportes es la que te interesa revisar.

## Qué se entrega hoy

El sistema son **dos archivos que trabajan encadenados**, cada uno con su panel y con los mismos botones en la pestaña Complementos, por si se prefiere trabajar sin abrir ventanas.

**1. PEDIDOS HCE — valida la dirección del cliente.** Trae los pedidos del día desde DEPOT y compara provincia, cantón y parroquia contra la cobertura. Cuando algo no coincide no lo corrige solo: propone una opción y explica el motivo, y el operador confirma uno por uno. Reconoce abreviaturas (FCO, PTO, STO), distingue parroquias con el mismo nombre en dos cantones, sugiere la cobertura más cercana cuando el destino no existe y marca las direcciones que caen en zonas peligrosas para cambiarlas a C.O.D con retiro en oficina. Cada corrección queda registrada en la hoja CAMBIOS con el dato original del cliente, el motivo, el usuario y la fecha.

**2. Formato EGR — decide el courier y despacha.** Recibe los pedidos ya validados y hace una segunda comprobación contra la cobertura del TMS. Aquí se define si el pedido sale **PRO, GYE, UIO o GPS**, se imprimen las etiquetas en la Zebra, se sigue el avance del empaque y se generan los archivos de TMS, TRAMACO y DESPACHOS.

**Cómo se enlazan.** El paso 7 de PEDIDOS HCE ("Enviar a EGR") copia los pedidos a la hoja DATOS y, además, el gestor, el destino, el trayecto, el tipo de entrega y la zona peligrosa que calculó la primera etapa. Así, quien despacha ve en la misma pantalla lo que decidió la validación y puede sustentar cualquier cambio.

**Cómo se presentan las validaciones.** Nada queda escondido en fórmulas. Los dos paneles muestran una fila por pedido con una señal visible: fuera de cobertura, zona peligrosa, cambio de destino sugerido, etiqueta por reimprimir, destino confirmado, o sin observaciones. Al hacer clic en un pedido, el detalle explica en texto por qué el sistema propone lo que propone. Las reglas de destino viven en una hoja editable (REGLAS_DESTINO), no en el código, así que se ajustan cuando cambie un acuerdo con el courier.

## Acuerdos de la reunión (hoy, 10:00 a 11:15)

| # | Punto | Estado |
|---|---|---|
| 1 | Revisar el reporte de cambios de PEDIDOS HCE, en particular las filas inconsistentes | Pendiente de revisión conjunta |
| 2 | Confirmar que la cobertura vigente es **COBERTURA HYCITE ENERO 2026.xlsm** y compararla con la que hoy usan los dos archivos | Pendiente de confirmación |
| 3 | Enviar la solicitud de alza de coberturas: provincias, cantones y parroquias que hoy no existen en el maestro | Pendiente de envío |
| 4 | Agregar al CSV de TRAMACO la columna **días de frecuencia**, junto con la disponibilidad del courier y la marca de fuera de cobertura, para la validación en distribución | Confirmado para la v1.2 |

Sobre el punto 4: la columna de días de frecuencia es la que permite distinguir un destino que realmente no tiene cobertura de uno que sí la tiene pero con frecuencia limitada. Esa información ya existe en la cobertura, así que es cuestión de llevarla al CSV.

## Qué necesito de cada uno

- **Cristina:** confirmar si COBERTURA HYCITE ENERO 2026 es la versión vigente. En cuanto la tenga, comparo las dos y les paso la lista de diferencias.
- **Jimmy:** revisar el formato del CSV de TRAMACO con la columna nueva, para que salga con el nombre y la posición que espera distribución.
- **Adriana:** probar el flujo completo con el manual y anotar cualquier paso que no quede claro; eso se corrige en el manual, no en la operación.

Con esos tres insumos genero la **v1.2** con los cuatro puntos incorporados.

Quedo atento a sus comentarios.

Saludos,

Brando Espinosa
Sistemas
