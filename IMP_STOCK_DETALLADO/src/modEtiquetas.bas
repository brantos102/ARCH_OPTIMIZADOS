Attribute VB_Name = "modEtiquetas"
Option Explicit

'==============================================================================================
' modEtiquetas  -  Impresión masiva de etiquetas  (IMP_STOCK_DETALLADO.xlsm)
'----------------------------------------------------------------------------------------------
' Plantilla ....: hoja "ETQ"  (bloque A1:J17, una etiqueta por página)
' Origen .......: hoja "Consolidado" (tabla con autofiltro) u hoja "IMPRIMIR"
' Impresora ....: Zebra ZD421 203 dpi driver ZPL - etiqueta 100 x 50 mm
'
' QUÉ RESUELVE (frente a la macro original):
'   1. RESPETA EL FILTRO: nunca imprime filas ocultas por el autofiltro.
'   2. LOTES GRANDES: arma una hoja temporal "ETQ_LOTE" con N copias de la plantilla,
'      una etiqueta por página, y la envía en UN solo trabajo de impresión por cada
'      bloque de PAGINAS_POR_TRABAJO etiquetas (en vez de un trabajo por etiqueta).
'   3. COPIAS: permite elegir cuántas copias se imprimen de cada fila, fijas para todas
'      o tomadas de la columna "cantidad".
'   4. RENDERIZADO: limpia saltos de línea, errores #N/A, fechas y números antes de
'      escribirlos, y verifica que el código de barras sea válido para Code 39.
'   5. NO ROMPE NADA: la plantilla ETQ y el resto del libro quedan intactos; la hoja
'      temporal se elimina al terminar y el estado de Excel se restaura siempre.
'
' PROCEDIMIENTOS PÚBLICOS:
'   ImprimirEtiquetasZebra ........... asistente de 4 pasos (entrada recomendada)
'   ImprimirEtiquetasSeleccion ....... solo las filas seleccionadas (visibles)
'   ImprimirEtiquetasFiltroCompleto .. todas las filas visibles del filtro actual
'   EtiquetasDiagnostico ............. informe de configuración, mapeo e impresora
'==============================================================================================


'=============================== CONFIGURACIÓN (editable) =====================================

' Hojas
Private Const HOJA_PLANTILLA     As String = "ETQ"
Private Const HOJA_LOTE          As String = "ETQ_LOTE"
Private Const HOJA_ORIGEN_1      As String = "Consolidado"
Private Const HOJA_ORIGEN_2      As String = "IMPRIMIR"

' Geometría de la plantilla: la etiqueta ocupa las filas 1 a 17 y las columnas A a J
Private Const FILAS_ETIQUETA     As Long = 17
Private Const COL_FINAL          As String = "J"

' Cuántas etiquetas se envían en cada trabajo de impresión.
' 100 es un buen equilibrio para la ZD421; si la cola de impresión se atasca, baje a 50.
Private Const PAGINAS_POR_TRABAJO As Long = 100

' Tope de filas visibles que se recorren al explorar la hoja de origen
Private Const TOPE_FILAS         As Long = 20000

' A partir de esta cantidad de etiquetas se pide una confirmación extra
Private Const AVISO_DESDE        As Long = 300

' Máximo de etiquetas que se aceptan en una sola corrida (seguridad)
Private Const TOPE_ETIQUETAS     As Long = 10000

' Escala de impresión:
'   0 = MEDIR la escala real con la paginación de Excel (recomendado). Se busca el
'       mayor porcentaje con el que las columnas A:J entran en el ancho de la etiqueta
'       y las 17 filas entran en el alto: exactamente lo mismo que hace la hoja ETQ
'       cuando se imprime una sola etiqueta.
'   1-100 = forzar ese porcentaje (sólo si la medición automática no funciona)
Private Const ESCALA_FIJA        As Long = 0

' Columna fija para el campo "cliente / marca" (celda A2 de la etiqueta).
' Se usa posición fija porque en "Consolidado" la columna B es el nombre del cliente
' y en "IMPRIMIR" la columna B es el campo "ABC"; ambas son el texto que va en A2.
Private Const COL_CLIENTE        As Long = 2

' Impresora
Private Const IMPRESORA_CONTIENE As String = "ZD421"   ' texto que identifica a la Zebra
Private Const BUSCAR_IMPRESORA   As Boolean = True     ' intentar seleccionarla automáticamente

' Compatibilidad / depuración
Private Const ACTUALIZAR_ETQ     As Boolean = True     ' deja la 1a etiqueta cargada en ETQ
Private Const CONSERVAR_LOTE     As Boolean = False    ' True = no borra la hoja ETQ_LOTE

'==============================================================================================


Private Type tCampo
    Celda As String          ' celda de la plantilla ETQ (esquina superior de la combinación)
    Encabezado As String     ' encabezado buscado en la fila 1 del origen ("" = posición fija)
    ColFija As Long          ' columna usada si no se encuentra el encabezado
    Col As Long              ' columna resuelta
    Barras As Boolean        ' True = se escribe como *CÓDIGO* para la fuente Code 39
    Numero As Boolean        ' True = se conserva el valor numérico
End Type

Private mScreenUpdating  As Boolean
Private mEnableEvents    As Boolean
Private mDisplayAlerts   As Boolean
Private mEstadoGuardado  As Boolean
Private mZoom            As Long
Private mAvisos          As String
Private mNumAvisos       As Long


'==============================================================================================
' ENTRADAS PÚBLICAS
'==============================================================================================
Public Sub ImprimirEtiquetasZebra()
    EjecutarLote 0
End Sub

Public Sub ImprimirEtiquetasSeleccion()
    EjecutarLote 1
End Sub

Public Sub ImprimirEtiquetasFiltroCompleto()
    EjecutarLote 2
End Sub

' Imprime UNA sola etiqueta de prueba (la primera fila visible de la selección o del
' filtro) para verificar tamaño y posición antes de lanzar un lote grande.
Public Sub ImprimirUnaEtiquetaPrueba()
    EjecutarLote 3
End Sub


'==============================================================================================
' MOTOR
'   modo 0 = preguntar alcance    1 = solo selección    2 = todas las filas visibles
'   modo 3 = una sola etiqueta de prueba
'==============================================================================================
Private Sub EjecutarLote(ByVal modo As Long)

    Dim wsOrigen As Worksheet, wsPlantilla As Worksheet, wsLote As Worksheet
    Dim campos() As tCampo
    Dim filas() As Long, copias() As Long
    Dim cola() As Long
    Dim nFilas As Long, nCola As Long
    Dim soloSeleccion As Boolean, usarCantidad As Boolean, vistaPrevia As Boolean
    Dim copiasFijas As Long, colCantidad As Long, colProducto As Long, colSerie As Long
    Dim desde As Long, hasta As Long, omitidas As Long
    Dim bloquesHoja As Long, k As Long, idx As Long
    Dim trabajos As Long, impresas As Long, nTrabajos As Long
    Dim i As Long, c As Long, b As Long
    Dim respuesta As VbMsgBoxResult
    Dim rngCol As Range
    Dim hojaAnterior As Worksheet, selAnterior As Range
    Dim impresoraOrig As String, impresoraUsada As String
    Dim txt As String, msg As String
    Dim t0 As Single
    Dim nError As Long, sError As String

    mAvisos = "": mNumAvisos = 0: mZoom = 0

    '--- 1. Contexto ---------------------------------------------------------------------
    If Not ValidarEntorno(wsOrigen, wsPlantilla) Then Exit Sub

    Set hojaAnterior = wsOrigen
    If TypeName(Selection) = "Range" Then Set selAnterior = Selection

    ResolverCampos wsOrigen, campos
    colProducto = ResolverColumna(wsOrigen, "producto_id", 3)
    colSerie = ResolverColumna(wsOrigen, "nro_serie", 5)
    colCantidad = ResolverColumna(wsOrigen, "cantidad", 13)

    '--- 2. Alcance ----------------------------------------------------------------------
    soloSeleccion = (modo = 1) Or (modo = 3 And TypeName(Selection) = "Range")

    If modo = 1 And TypeName(Selection) <> "Range" Then
        MsgBox "Seleccione primero una o varias filas en la hoja """ & wsOrigen.Name & """.", _
               vbExclamation, "Etiquetas"
        Exit Sub
    End If

    If modo = 0 Then
        If TypeName(Selection) = "Range" Then
            If Selection.Rows.Count > 1 Or Selection.Areas.Count > 1 Then
                respuesta = MsgBox("PASO 1 de 4 - Alcance" & vbCrLf & vbCrLf & _
                    "SÍ = imprimir solo las filas SELECCIONADAS (se omiten las ocultas por el filtro)" & vbCrLf & vbCrLf & _
                    "NO = imprimir TODAS las filas visibles del filtro actual de """ & wsOrigen.Name & """", _
                    vbYesNoCancel + vbQuestion, "Etiquetas")
                If respuesta = vbCancel Then Exit Sub
                soloSeleccion = (respuesta = vbYes)
            End If
        End If
    End If

    '--- 3. Filas visibles ---------------------------------------------------------------
    Application.StatusBar = "Explorando filas visibles..."
    Set rngCol = RangoCandidato(wsOrigen, colProducto, soloSeleccion)
    nFilas = RecolectarVisibles(rngCol, TOPE_FILAS, filas)

    If nFilas = 0 And modo = 3 And soloSeleccion Then
        soloSeleccion = False
        Set rngCol = RangoCandidato(wsOrigen, colProducto, False)
        nFilas = RecolectarVisibles(rngCol, TOPE_FILAS, filas)
    End If
    Application.StatusBar = False

    If nFilas = 0 Then
        MsgBox "No se encontraron filas visibles para imprimir." & vbCrLf & vbCrLf & _
               "Revise el filtro de la hoja """ & wsOrigen.Name & """ o la selección.", _
               vbExclamation, "Etiquetas"
        Exit Sub
    End If

    '--- 4. Cuántas filas ----------------------------------------------------------------
    txt = "TODAS"
    If modo = 3 Then txt = "1"
    If (modo <> 1 And modo <> 3) Or (modo = 1 And nFilas > PAGINAS_POR_TRABAJO) Then
        txt = InputBox("PASO 2 de 4 - Filas a imprimir" & vbCrLf & vbCrLf & _
            "Filas visibles disponibles: " & nFilas & IIf(nFilas = TOPE_FILAS, " (tope alcanzado)", "") & vbCrLf & vbCrLf & _
            "Escriba:" & vbCrLf & _
            "    TODAS      para las " & nFilas & " filas" & vbCrLf & _
            "    200        para las primeras 200" & vbCrLf & _
            "    101-300    para un rango (de la 101 a la 300)", _
            "Etiquetas - filas", "TODAS")
        If Len(Trim$(txt)) = 0 Then Exit Sub
    End If

    desde = 1: hasta = nFilas
    txt = UCase$(Replace(Trim$(txt), " ", ""))
    If txt <> "TODAS" Then
        If InStr(txt, "-") > 0 Then
            desde = CLng(Val(Split(txt, "-")(0)))
            hasta = CLng(Val(Split(txt, "-")(1)))
        Else
            hasta = CLng(Val(txt))
        End If
    End If
    If desde < 1 Then desde = 1
    If hasta > nFilas Then hasta = nFilas
    If hasta < desde Then
        MsgBox "El rango indicado no es válido.", vbExclamation, "Etiquetas"
        Exit Sub
    End If

    '--- 5. Copias por etiqueta ----------------------------------------------------------
    If modo = 3 Then
        copiasFijas = 1
        GoTo ArmarCola
    End If

    txt = InputBox("PASO 3 de 4 - Copias de cada etiqueta" & vbCrLf & vbCrLf & _
        "Cada fila es una etiqueta. Indique cuántas copias se imprimen de cada una:" & vbCrLf & vbCrLf & _
        "    1, 2, 3 ...   número fijo de copias para todas" & vbCrLf & _
        "    C             usar la columna CANTIDAD de cada fila", _
        "Etiquetas - copias", "1")
    If Len(Trim$(txt)) = 0 Then Exit Sub

    txt = UCase$(Trim$(txt))
    If Left$(txt, 1) = "C" Then
        usarCantidad = True
        copiasFijas = 1
    Else
        copiasFijas = CLng(Val(txt))
        If copiasFijas < 1 Then copiasFijas = 1
    End If

    '--- 6. Armado de la cola de etiquetas ----------------------------------------------
ArmarCola:
    ReDim copias(1 To nFilas)
    nCola = 0
    For i = desde To hasta
        c = 0
        If FilaImprimible(wsOrigen, filas(i), colProducto, colSerie) Then
            If usarCantidad Then
                c = CopiasDeCantidad(wsOrigen.Cells(filas(i), colCantidad).Value)
            Else
                c = copiasFijas
            End If
        End If
        If c <= 0 Then omitidas = omitidas + 1
        copias(i) = c
        nCola = nCola + c
    Next i

    If nCola = 0 Then
        MsgBox "Las filas elegidas no tienen código ni número de serie, no hay nada que imprimir.", _
               vbExclamation, "Etiquetas"
        Exit Sub
    End If

    If nCola > TOPE_ETIQUETAS Then
        MsgBox "El lote resultante es de " & nCola & " etiquetas y el tope de seguridad es " & _
               TOPE_ETIQUETAS & "." & vbCrLf & vbCrLf & _
               "Reduzca el rango de filas o las copias, o aumente TOPE_ETIQUETAS en el módulo.", _
               vbExclamation, "Etiquetas"
        Exit Sub
    End If

    ReDim cola(1 To nCola)
    idx = 0
    For i = desde To hasta
        For c = 1 To copias(i)
            idx = idx + 1
            cola(idx) = filas(i)
        Next c
    Next i

    '--- 7. Impresora --------------------------------------------------------------------
    impresoraOrig = ImpresoraActual()
    If Not AsegurarImpresora() Then Exit Sub
    impresoraUsada = ImpresoraActual()

    '--- 8. Confirmación -----------------------------------------------------------------
    trabajos = (nCola + PAGINAS_POR_TRABAJO - 1) \ PAGINAS_POR_TRABAJO

    msg = "PASO 4 de 4 - Confirmación" & vbCrLf & String$(46, "-") & vbCrLf & _
          "Hoja de origen .......: " & wsOrigen.Name & vbCrLf & _
          "Filas a imprimir .....: " & (hasta - desde + 1) & "   (de la " & desde & " a la " & hasta & ")" & vbCrLf & _
          "Copias por etiqueta ..: " & IIf(usarCantidad, "según columna CANTIDAD", CStr(copiasFijas)) & vbCrLf & _
          "ETIQUETAS TOTALES ....: " & nCola & vbCrLf & _
          "Trabajos de impresión : " & trabajos & " (hasta " & PAGINAS_POR_TRABAJO & " etiquetas cada uno)" & vbCrLf & _
          "Filas omitidas .......: " & omitidas & vbCrLf & _
          "Impresora ............: " & impresoraUsada & vbCrLf & String$(46, "-") & vbCrLf & vbCrLf & _
          "SÍ = ver la VISTA PREVIA antes de imprimir" & vbCrLf & _
          "NO = enviar a imprimir directamente" & vbCrLf & _
          "CANCELAR = salir sin imprimir"

    If nCola >= AVISO_DESDE Then
        msg = msg & vbCrLf & vbCrLf & "ATENCIÓN: va a imprimir " & nCola & " etiquetas."
    End If

    If modo = 3 Then
        msg = "ETIQUETA DE PRUEBA" & vbCrLf & String$(46, "-") & vbCrLf & _
              "Se imprimirá UNA etiqueta con los datos de la primera fila visible," & vbCrLf & _
              "para verificar tamaño, posición y lectura del código de barras." & vbCrLf & _
              "Impresora: " & impresoraUsada & vbCrLf & vbCrLf & _
              "SÍ = ver la vista previa primero" & vbCrLf & "NO = imprimir directamente"
    End If

    respuesta = MsgBox(msg, vbYesNoCancel + vbQuestion, "Etiquetas")
    If respuesta = vbCancel Then
        RestaurarImpresora impresoraOrig, impresoraUsada
        Exit Sub
    End If
    vistaPrevia = (respuesta = vbYes)

    '--- 9. Impresión --------------------------------------------------------------------
    t0 = Timer
    On Error GoTo Limpieza
    GuardarEstado

    Set wsLote = CrearHojaLote(wsPlantilla)
    bloquesHoja = MinL(nCola, PAGINAS_POR_TRABAJO)

    Application.StatusBar = "Preparando el lote de etiquetas..."
    PrepararBloques wsLote, bloquesHoja

    Application.StatusBar = "Calculando la escala de impresión..."
    mZoom = DeterminarZoom(wsLote)
    MarcarSaltos wsLote, bloquesHoja

    idx = 1
    Do While idx <= nCola
        k = MinL(bloquesHoja, nCola - idx + 1)

        LlenarBloques wsLote, wsOrigen, cola, idx, k, campos

        If mZoom > 0 Then
            AjustarPagina wsLote, k
            If Not VerificarPaginacion(wsLote, k) Then GoTo Limpieza
        Else
            ' no se pudo medir la escala: modo compatible, se imprime etiqueta por
            ' etiqueta con el mismo ajuste de página que usa la hoja ETQ
            AjustarPaginaEtiqueta wsLote, 1
        End If

        If vistaPrevia Then
            vistaPrevia = False
            Application.ScreenUpdating = True
            wsLote.Activate
            wsLote.PrintPreview
            Application.ScreenUpdating = False
            respuesta = MsgBox("¿Enviar a imprimir las " & nCola & " etiquetas?", _
                               vbYesNo + vbQuestion, "Etiquetas")
            If respuesta <> vbYes Then GoTo Limpieza
        End If

        If mZoom > 0 Then
            Application.StatusBar = "Imprimiendo etiquetas " & idx & " a " & (idx + k - 1) & " de " & nCola & "..."
            wsLote.PrintOut Copies:=1, Collate:=True
            nTrabajos = nTrabajos + 1
        Else
            For b = 1 To k
                AjustarPaginaEtiqueta wsLote, b
                Application.StatusBar = "Imprimiendo etiqueta " & (idx + b - 1) & " de " & nCola & "..."
                wsLote.PrintOut Copies:=1
                nTrabajos = nTrabajos + 1
            Next b
        End If

        impresas = impresas + k
        idx = idx + k
    Loop

Limpieza:
    nError = Err.Number
    sError = Err.Description
    On Error Resume Next

    If ACTUALIZAR_ETQ And impresas > 0 Then
        LlenarEtiquetaPlantilla wsPlantilla, wsOrigen, cola(1), campos
    End If

    If Not wsLote Is Nothing Then
        If CONSERVAR_LOTE Then
            AjustarPagina wsLote, bloquesHoja
        Else
            Application.DisplayAlerts = False
            wsLote.Delete
        End If
    End If

    RestaurarImpresora impresoraOrig, impresoraUsada
    RestaurarEstado

    If Not hojaAnterior Is Nothing Then hojaAnterior.Activate
    If Not selAnterior Is Nothing Then selAnterior.Select

    Application.StatusBar = False
    Err.Clear

    If nError <> 0 Then
        MsgBox "Se interrumpió el proceso." & vbCrLf & vbCrLf & _
               "Etiquetas enviadas antes del fallo: " & impresas & vbCrLf & vbCrLf & _
               "Error " & nError & ": " & sError, vbCritical, "Etiquetas"
    ElseIf impresas = 0 Then
        MsgBox "No se envió ninguna etiqueta a imprimir.", vbInformation, "Etiquetas"
    Else
        msg = "Impresión enviada correctamente" & vbCrLf & String$(42, "-") & vbCrLf & _
              "Etiquetas enviadas ...: " & impresas & vbCrLf & _
              "Trabajos de impresión : " & nTrabajos & vbCrLf & _
              "Impresora ............: " & impresoraUsada & vbCrLf & _
              "Filas omitidas .......: " & omitidas & vbCrLf & _
              "Escala de impresión ..: " & IIf(mZoom > 0, mZoom & " %", "ajuste de la hoja ETQ") & vbCrLf & _
              "Tiempo ...............: " & Format$(Timer - t0, "0.0") & " s"
        If mZoom = 0 Then
            msg = msg & vbCrLf & vbCrLf & _
                  "Nota: no se pudo medir la escala de impresión, así que se imprimió" & vbCrLf & _
                  "en modo compatible (un trabajo por etiqueta). El resultado es correcto," & vbCrLf & _
                  "pero más lento. Ejecute EtiquetasDiagnostico para ver el detalle."
        End If
        If mNumAvisos > 0 Then
            msg = msg & vbCrLf & vbCrLf & "Avisos de código de barras (" & mNumAvisos & "):" & mAvisos
            If mNumAvisos > 10 Then msg = msg & vbCrLf & "  ... y " & (mNumAvisos - 10) & " más."
        End If
        MsgBox msg, vbInformation, "Etiquetas"
    End If
End Sub


'==============================================================================================
' VALIDACIÓN DE CONTEXTO
'==============================================================================================
Private Function ValidarEntorno(ByRef wsOrigen As Worksheet, ByRef wsPlantilla As Worksheet) As Boolean

    Dim nombre As String

    On Error Resume Next
    Set wsPlantilla = ThisWorkbook.Worksheets(HOJA_PLANTILLA)
    On Error GoTo 0

    If wsPlantilla Is Nothing Then
        MsgBox "No se encontró la hoja de plantilla """ & HOJA_PLANTILLA & """.", vbCritical, "Etiquetas"
        Exit Function
    End If

    If ActiveWorkbook Is Nothing Then Exit Function
    If Not ActiveWorkbook Is ThisWorkbook Then
        MsgBox "Active el libro """ & ThisWorkbook.Name & """ antes de imprimir etiquetas.", _
               vbExclamation, "Etiquetas"
        Exit Function
    End If
    If ActiveSheet Is Nothing Then Exit Function

    nombre = ActiveSheet.Name
    Select Case UCase$(nombre)
        Case UCase$(HOJA_ORIGEN_1), UCase$(HOJA_ORIGEN_2)
            Set wsOrigen = ThisWorkbook.Worksheets(nombre)
        Case Else
            MsgBox "Debe estar en la hoja """ & HOJA_ORIGEN_1 & """ o """ & HOJA_ORIGEN_2 & """ " & _
                   "para imprimir etiquetas." & vbCrLf & vbCrLf & _
                   "Hoja activa: " & nombre, vbExclamation, "Etiquetas"
            Exit Function
    End Select

    ValidarEntorno = True
End Function


'==============================================================================================
' MAPEO DE CAMPOS  (celda de la etiqueta  <--  columna del origen)
'==============================================================================================
Private Sub ResolverCampos(ByVal ws As Worksheet, ByRef campos() As tCampo)

    Dim i As Long

    ReDim campos(1 To 12)
    '          celda   encabezado             col. fija  barras  número
    DefCampo campos(1), "A2", "", COL_CLIENTE, False, False               ' cliente / marca
    DefCampo campos(2), "H3", "categoria_logica", 9, False, False         ' DISPONIBLE ...
    DefCampo campos(3), "H5", "nro_despacho", 7, False, False
    DefCampo campos(4), "A6", "producto_id", 3, True, False               ' código de barras
    DefCampo campos(5), "A9", "producto_id", 3, False, False              ' código legible
    DefCampo campos(6), "H10", "cantidad", 13, False, True                ' QTY
    DefCampo campos(7), "I10", "unidad_medida", 12, False, False
    DefCampo campos(8), "J10", "estado_mercaderia", 10, False, False
    DefCampo campos(9), "A11", "descripcion", 4, False, False
    DefCampo campos(10), "H14", "nro_partida", 8, False, False            ' RESPONSABLE
    DefCampo campos(11), "A15", "nro_serie", 5, True, False               ' serie en barras
    DefCampo campos(12), "A17", "nro_serie", 5, False, False              ' serie legible

    For i = LBound(campos) To UBound(campos)
        campos(i).Col = ResolverColumna(ws, campos(i).Encabezado, campos(i).ColFija)
    Next i
End Sub

Private Sub DefCampo(ByRef c As tCampo, ByVal celda As String, ByVal encabezado As String, _
                     ByVal colFija As Long, ByVal barras As Boolean, ByVal numero As Boolean)
    c.Celda = celda
    c.Encabezado = encabezado
    c.ColFija = colFija
    c.Col = colFija
    c.Barras = barras
    c.Numero = numero
End Sub

' Busca el encabezado en la fila 1. Si no lo encuentra usa la columna fija de la macro original.
Private Function ResolverColumna(ByVal ws As Worksheet, ByVal encabezado As String, _
                                 ByVal colFija As Long) As Long
    Dim c As Long, v As Variant

    ResolverColumna = colFija
    If Len(encabezado) = 0 Then Exit Function

    For c = 1 To 80
        v = ws.Cells(1, c).Value
        If Not IsError(v) Then
            If LCase$(Trim$(CStr(v & ""))) = LCase$(encabezado) Then
                ResolverColumna = c
                Exit Function
            End If
        End If
    Next c
End Function


'==============================================================================================
' RECOLECCIÓN DE FILAS VISIBLES  (respeta el autofiltro)
'==============================================================================================
Private Function RangoCandidato(ByVal ws As Worksheet, ByVal colRef As Long, _
                                ByVal soloSeleccion As Boolean) As Range
    Dim ult As Long

    ult = UltimaFila(ws, colRef)
    If ult < 2 Then Exit Function

    If soloSeleccion Then
        If TypeName(Selection) <> "Range" Then Exit Function
        Set RangoCandidato = Application.Intersect(Selection.EntireRow, _
                                ws.Range(ws.Cells(2, colRef), ws.Cells(ult, colRef)))
    Else
        Set RangoCandidato = ws.Range(ws.Cells(2, colRef), ws.Cells(ult, colRef))
    End If
End Function

Private Function UltimaFila(ByVal ws As Worksheet, ByVal colRef As Long) As Long
    Dim lo As ListObject

    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        If Not lo.DataBodyRange Is Nothing Then
            UltimaFila = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
            Exit Function
        End If
    End If
    UltimaFila = ws.Cells(ws.Rows.Count, colRef).End(xlUp).Row
End Function

' Devuelve en filas() los números de fila VISIBLES del rango, hasta "tope".
' Trabaja por bloques para no agotar SpecialCells cuando hay miles de áreas ocultas.
Private Function RecolectarVisibles(ByVal rngCol As Range, ByVal tope As Long, _
                                    ByRef filas() As Long) As Long
    Const BLOQUE As Long = 20000

    Dim ar As Range, vis As Range, a As Range, ws As Worksheet
    Dim ini As Long, fin As Long, r1 As Long, r2 As Long, i As Long, n As Long

    ReDim filas(1 To tope)
    If rngCol Is Nothing Then Exit Function

    Set ws = rngCol.Worksheet

    For Each ar In rngCol.Areas
        ini = ar.Row
        fin = ar.Row + ar.Rows.Count - 1
        r1 = ini
        Do While r1 <= fin
            r2 = r1 + BLOQUE - 1
            If r2 > fin Then r2 = fin

            Set vis = Nothing
            On Error Resume Next
            Set vis = ws.Range(ws.Cells(r1, ar.Column), ws.Cells(r2, ar.Column)) _
                        .SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            If Not vis Is Nothing Then
                For Each a In vis.Areas
                    For i = 1 To a.Rows.Count
                        n = n + 1
                        filas(n) = a.Row + i - 1
                        If n >= tope Then GoTo Fin
                    Next i
                Next a
            End If

            r1 = r2 + 1
        Loop
    Next ar

Fin:
    RecolectarVisibles = n
End Function

Private Function FilaImprimible(ByVal ws As Worksheet, ByVal fila As Long, _
                                ByVal colProducto As Long, ByVal colSerie As Long) As Boolean
    Dim a As String, b As String

    a = TextoCelda(ws.Cells(fila, colProducto))
    b = TextoCelda(ws.Cells(fila, colSerie))
    FilaImprimible = (Len(a) > 0 Or Len(b) > 0)
End Function

Private Function CopiasDeCantidad(ByVal v As Variant) As Long
    Dim n As Double

    If IsError(v) Then Exit Function
    If Not EsNumero(v) Then
        If Len(Trim$(CStr(v & ""))) = 0 Then
            CopiasDeCantidad = 1            ' sin cantidad -> una etiqueta
            Exit Function
        End If
        If Not IsNumeric(v) Then
            CopiasDeCantidad = 1
            Exit Function
        End If
    End If

    n = CDbl(v)
    If n < 1 Then
        CopiasDeCantidad = 1
    ElseIf n > 1000 Then
        CopiasDeCantidad = 1000
    Else
        CopiasDeCantidad = CLng(Int(n))
    End If
End Function


'==============================================================================================
' HOJA DE LOTE
'==============================================================================================
Private Function CrearHojaLote(ByVal wsPlantilla As Worksheet) As Worksheet

    Dim ws As Worksheet

    EliminarHojaLote

    wsPlantilla.Copy After:=wsPlantilla
    Set ws = wsPlantilla.Parent.Sheets(wsPlantilla.Index + 1)

    On Error Resume Next
    ws.Name = HOJA_LOTE
    ws.Visible = xlSheetVisible
    ws.Unprotect
    On Error GoTo 0

    ' la etiqueta son solo las filas 1 a 17: se limpia todo lo que haya debajo
    ws.Rows((FILAS_ETIQUETA + 1) & ":" & (FILAS_ETIQUETA + 400)).Clear

    Set CrearHojaLote = ws
End Function

Private Sub EliminarHojaLote()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_LOTE)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ws.Delete
        On Error GoTo 0
        If Not mEstadoGuardado Then Application.DisplayAlerts = True
    End If
End Sub

' Replica el bloque de 17 filas hasta tener "n" etiquetas y coloca un salto de página
' al inicio de cada una. La réplica se hace duplicando lo ya copiado (1, 2, 4, 8 ...),
' por lo que son pocas operaciones de copiado aunque el lote sea grande.
Private Sub PrepararBloques(ByVal ws As Worksheet, ByVal n As Long)

    Dim hay As Long, nuevos As Long, i As Long

    Application.CutCopyMode = False
    hay = 1
    Do While hay < n
        nuevos = MinL(hay, n - hay)
        ws.Rows("1:" & (FILAS_ETIQUETA * nuevos)).Copy _
            Destination:=ws.Rows((FILAS_ETIQUETA * hay + 1) & ":" & (FILAS_ETIQUETA * (hay + nuevos)))
        hay = hay + nuevos
    Loop
    Application.CutCopyMode = False

    ' área de impresión de al menos dos etiquetas: hace falta para poder medir dónde
    ' corta Excel la página (ver DeterminarZoom)
    On Error Resume Next
    Application.PrintCommunication = False
    ws.PageSetup.PrintArea = "$A$1:$" & COL_FINAL & "$" & (FILAS_ETIQUETA * MaxL(n, 2))
    Application.PrintCommunication = True
    ws.ResetAllPageBreaks
    On Error GoTo 0
End Sub

' Coloca un salto de página al inicio de cada etiqueta.
Private Sub MarcarSaltos(ByVal ws As Worksheet, ByVal n As Long)
    Dim i As Long

    On Error Resume Next
    ws.ResetAllPageBreaks
    For i = 1 To n - 1
        ws.Rows(FILAS_ETIQUETA * i + 1).PageBreak = xlPageBreakManual
    Next i
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------------------------------
' ESCALA DE IMPRESIÓN
'
' La plantilla ETQ imprime "ajustar a 1 página": Excel calcula solo el porcentaje con el
' que las 10 columnas y las 17 filas entran en la etiqueta de 100 x 50 mm. Ese cálculo no
' se puede trasladar tal cual a una hoja de varias etiquetas (el ajuste se aplicaría al
' lote completo, no a cada etiqueta), así que aquí se MIDE: se prueban porcentajes y se
' mira dónde corta Excel la página con la impresora real.
'
' Se busca el mayor porcentaje que cumple las dos condiciones de una etiqueta:
'    - las columnas A:J entran en el ancho de la página  (sin saltos verticales)
'    - las 17 filas entran en el alto de la página       (el 1er salto va en la fila 18)
'----------------------------------------------------------------------------------------------
Private Function DeterminarZoom(ByVal ws As Worksheet) As Long

    Dim lo As Long, hi As Long, medio As Long, mejor As Long
    Dim filasPag As Long, anchoOK As Boolean
    Dim screenAnt As Boolean, saltosAnt As Boolean

    If ESCALA_FIJA > 0 Then
        DeterminarZoom = ESCALA_FIJA
        Exit Function
    End If

    screenAnt = Application.ScreenUpdating
    saltosAnt = ws.DisplayPageBreaks

    ' Excel sólo calcula los saltos automáticos con el dibujo de pantalla activo
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.DisplayPageBreaks = True
    On Error GoTo 0

    lo = 10: hi = 100
    Do While lo <= hi
        medio = (lo + hi) \ 2
        filasPag = FilasPorPagina(ws, medio, anchoOK)
        If anchoOK And filasPag >= FILAS_ETIQUETA Then
            mejor = medio
            lo = medio + 1
        Else
            hi = medio - 1
        End If
    Loop

    On Error Resume Next
    ws.DisplayPageBreaks = saltosAnt
    On Error GoTo 0
    Application.ScreenUpdating = screenAnt

    DeterminarZoom = mejor
End Function

' Aplica el zoom indicado y devuelve cuántas filas entran en la primera página.
' anchoOK queda en True si las columnas del área de impresión entran a lo ancho.
Private Function FilasPorPagina(ByVal ws As Worksheet, ByVal z As Long, _
                                ByRef anchoOK As Boolean) As Long
    anchoOK = False

    On Error Resume Next
    Application.PrintCommunication = False
    ws.PageSetup.Zoom = z
    Application.PrintCommunication = True
    ws.ResetAllPageBreaks

    anchoOK = (ws.VPageBreaks.Count = 0)

    If ws.HPageBreaks.Count = 0 Then
        FilasPorPagina = 1000000            ' no hay corte: entra todo el área
    Else
        FilasPorPagina = ws.HPageBreaks(1).Location.Row - 1
    End If
    On Error GoTo 0
End Function

Private Sub AjustarPagina(ByVal ws As Worksheet, ByVal k As Long)

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$A$1:$" & COL_FINAL & "$" & (FILAS_ETIQUETA * k)
        If mZoom > 0 Then
            .Zoom = mZoom                   ' escala medida: una etiqueta por página
        Else
            .Zoom = False                   ' no se pudo medir: ajuste heredado de ETQ
            .FitToPagesWide = 1
            .FitToPagesTall = k
        End If
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' Modo compatible: deja como área de impresión UNA sola etiqueta del lote, con el mismo
' "ajustar a 1 página" que usa la hoja ETQ. Se imprime igual que la macro original.
Private Sub AjustarPaginaEtiqueta(ByVal ws As Worksheet, ByVal b As Long)

    Dim f1 As Long, f2 As Long

    f1 = (b - 1) * FILAS_ETIQUETA + 1
    f2 = b * FILAS_ETIQUETA

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$A$" & f1 & ":$" & COL_FINAL & "$" & f2
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' Comprueba que el lote realmente se pagine a una etiqueta por página.
Private Function VerificarPaginacion(ByVal ws As Worksheet, ByVal k As Long) As Boolean

    Dim p As Long

    On Error Resume Next
    p = ws.PageSetup.Pages.Count
    On Error GoTo 0

    If p = 0 Or p = k Then
        VerificarPaginacion = True
        Exit Function
    End If

    Application.ScreenUpdating = True
    VerificarPaginacion = (MsgBox("Aviso de paginación" & vbCrLf & vbCrLf & _
        "Se esperaban " & k & " páginas (una por etiqueta) y Excel calcula " & p & "." & vbCrLf & _
        "Escala medida: " & mZoom & " %" & vbCrLf & vbCrLf & _
        "Revise en Diseño de página que el tamaño de papel de la hoja ETQ sea la etiqueta " & _
        "de 100 x 50 mm de la Zebra, y que la impresora activa sea la ZD421." & vbCrLf & vbCrLf & _
        "¿Desea continuar de todas formas?", vbYesNo + vbExclamation, "Etiquetas") = vbYes)
    Application.ScreenUpdating = False
End Function


'==============================================================================================
' ESCRITURA DE DATOS
'==============================================================================================
Private Sub LlenarBloques(ByVal wsLote As Worksheet, ByVal wsOrigen As Worksheet, _
                          ByRef cola() As Long, ByVal idx As Long, ByVal k As Long, _
                          ByRef campos() As tCampo)
    Dim b As Long, i As Long, off As Long, fila As Long

    For b = 1 To k
        fila = cola(idx + b - 1)
        off = (b - 1) * FILAS_ETIQUETA
        For i = LBound(campos) To UBound(campos)
            EscribirCampo wsLote.Range(campos(i).Celda).Offset(off, 0), _
                          wsOrigen.Cells(fila, campos(i).Col), campos(i), fila, True
        Next i
    Next b
End Sub

' Carga la primera etiqueta del lote en la plantilla ETQ (compatibilidad con la macro original)
Private Sub LlenarEtiquetaPlantilla(ByVal wsPlantilla As Worksheet, ByVal wsOrigen As Worksheet, _
                                    ByVal fila As Long, ByRef campos() As tCampo)
    Dim i As Long

    On Error Resume Next
    For i = LBound(campos) To UBound(campos)
        EscribirCampo wsPlantilla.Range(campos(i).Celda), _
                      wsOrigen.Cells(fila, campos(i).Col), campos(i), fila, False
    Next i
    On Error GoTo 0
End Sub

Private Sub EscribirCampo(ByVal destino As Range, ByVal origen As Range, ByRef campo As tCampo, _
                          ByVal fila As Long, ByVal forzarTexto As Boolean)
    Dim v As Variant
    Dim s As String

    v = origen.Value

    If campo.Numero Then
        If Not IsError(v) Then
            If EsNumero(v) Then
                destino.Value = v
                Exit Sub
            End If
        End If
    End If

    s = TextoCelda(origen)
    If campo.Barras Then s = CodigoBarras(s, fila)

    If forzarTexto Then destino.NumberFormat = "@"
    destino.Value = s
End Sub

' Convierte el contenido de una celda en el texto que debe salir impreso.
' Neutraliza errores, saltos de línea, notación científica y fechas.
Private Function TextoCelda(ByVal origen As Range) As String

    Dim v As Variant
    Dim s As String

    v = origen.Value
    If IsError(v) Then Exit Function

    Select Case VarType(v)
        Case vbEmpty, vbNull
            Exit Function
        Case vbDate
            s = Format$(v, "dd/mm/yyyy")
        Case vbBoolean
            s = IIf(CBool(v), "SI", "NO")
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            If EsFormatoFecha(origen) Then
                s = Format$(CDate(v), "dd/mm/yyyy")
            Else
                s = Format$(v, "0.##########")
            End If
        Case Else
            s = CStr(v)
    End Select

    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    TextoCelda = Trim$(s)
End Function

Private Function EsFormatoFecha(ByVal c As Range) As Boolean
    Dim nf As String

    On Error Resume Next
    nf = LCase$(c.NumberFormat)
    On Error GoTo 0
    EsFormatoFecha = (InStr(nf, "yy") > 0 Or InStr(nf, "aa") > 0 Or InStr(nf, "dd/") > 0)
End Function

Private Function EsNumero(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            EsNumero = True
    End Select
End Function

' Prepara el texto para la fuente "Free 3 of 9 Extended" (Code 39): mayúsculas y
' asteriscos de inicio/fin. Si hay caracteres que Code 39 no puede codificar, avisa.
Private Function CodigoBarras(ByVal s As String, ByVal fila As Long) As String

    Const VALIDOS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"

    Dim t As String, ch As String, malos As String
    Dim i As Long

    t = UCase$(Trim$(s))
    If Len(t) = 0 Then Exit Function

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If InStr(1, VALIDOS, ch, vbBinaryCompare) = 0 Then
            If InStr(1, malos, ch, vbBinaryCompare) = 0 Then malos = malos & ch
        End If
    Next i

    If Len(malos) > 0 Then
        AgregarAviso "fila " & fila & ": """ & malos & """ no se puede codificar en Code 39"
    ElseIf Len(t) > 20 Then
        AgregarAviso "fila " & fila & ": código de " & Len(t) & " caracteres, el ancho de barras queda muy comprimido"
    End If

    CodigoBarras = "*" & t & "*"
End Function

Private Sub AgregarAviso(ByVal s As String)
    mNumAvisos = mNumAvisos + 1
    If mNumAvisos <= 10 Then mAvisos = mAvisos & vbCrLf & "  - " & s
End Sub


'==============================================================================================
' IMPRESORA
'==============================================================================================
Private Function ImpresoraActual() As String
    On Error Resume Next
    ImpresoraActual = Application.ActivePrinter
    On Error GoTo 0
End Function

Private Function AsegurarImpresora() As Boolean

    Dim act As String

    AsegurarImpresora = True
    act = ImpresoraActual()

    If InStr(1, act, IMPRESORA_CONTIENE, vbTextCompare) > 0 Then Exit Function
    If BUSCAR_IMPRESORA Then
        If SeleccionarImpresoraZebra() Then Exit Function
    End If

    AsegurarImpresora = (MsgBox("La impresora activa es:" & vbCrLf & vbCrLf & "    " & act & vbCrLf & vbCrLf & _
        "No parece ser la Zebra " & IMPRESORA_CONTIENE & ". Si continúa, las etiquetas se " & _
        "enviarán a esa impresora." & vbCrLf & vbCrLf & _
        "¿Desea continuar?", vbYesNo + vbExclamation, "Etiquetas") = vbYes)
End Function

Private Function SeleccionarImpresoraZebra() As Boolean

    Dim wmi As Object, lista As Object, p As Object
    Dim nombre As String, puerto As String

    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then Exit Function
    Set lista = wmi.ExecQuery("SELECT Name, PortName FROM Win32_Printer")
    If lista Is Nothing Then Exit Function

    For Each p In lista
        nombre = CStr(p.Name & "")
        If InStr(1, nombre, IMPRESORA_CONTIENE, vbTextCompare) > 0 Then
            puerto = CStr(p.PortName & "")
            If ProbarImpresora(nombre & " en " & puerto) Then GoTo Exito
            If ProbarImpresora(nombre & " on " & puerto) Then GoTo Exito
            If ProbarImpresora(nombre) Then GoTo Exito
        End If
    Next p
    Exit Function

Exito:
    SeleccionarImpresoraZebra = True
End Function

Private Sub RestaurarImpresora(ByVal original As String, ByVal usada As String)
    If Len(original) = 0 Then Exit Sub
    If usada = original Then Exit Sub

    On Error Resume Next
    Application.ActivePrinter = original
    On Error GoTo 0
End Sub

Private Function ProbarImpresora(ByVal s As String) As Boolean
    On Error Resume Next
    Err.Clear
    Application.ActivePrinter = s

    ' algunas versiones de Excel no dan error: hay que confirmar el cambio
    If Err.Number = 0 Then
        ProbarImpresora = (InStr(1, Application.ActivePrinter, IMPRESORA_CONTIENE, vbTextCompare) > 0)
    End If
    Err.Clear
End Function


'==============================================================================================
' ESTADO DE EXCEL
'==============================================================================================
Private Sub GuardarEstado()
    If mEstadoGuardado Then Exit Sub

    mScreenUpdating = Application.ScreenUpdating
    mEnableEvents = Application.EnableEvents
    mDisplayAlerts = Application.DisplayAlerts
    mEstadoGuardado = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Private Sub RestaurarEstado()
    If Not mEstadoGuardado Then Exit Sub

    On Error Resume Next
    Application.DisplayAlerts = mDisplayAlerts
    Application.EnableEvents = mEnableEvents
    Application.ScreenUpdating = mScreenUpdating
    Application.CutCopyMode = False
    On Error GoTo 0

    mEstadoGuardado = False
End Sub

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function


'==============================================================================================
' DIAGNÓSTICO
'==============================================================================================
Public Sub EtiquetasDiagnostico()

    Dim wsOrigen As Worksheet, wsPlantilla As Worksheet, wsLote As Worksheet
    Dim campos() As tCampo
    Dim filas() As Long
    Dim nFilas As Long, i As Long, escala As Long, paginas As Long
    Dim msg As String

    If Not ValidarEntorno(wsOrigen, wsPlantilla) Then Exit Sub

    ResolverCampos wsOrigen, campos
    nFilas = RecolectarVisibles(RangoCandidato(wsOrigen, ResolverColumna(wsOrigen, "producto_id", 3), False), _
                                TOPE_FILAS, filas)

    ' se arma un lote de prueba de 2 etiquetas para medir la escala y la paginación reales
    On Error Resume Next
    GuardarEstado
    Set wsLote = CrearHojaLote(wsPlantilla)
    PrepararBloques wsLote, 2
    escala = DeterminarZoom(wsLote)
    mZoom = escala
    AjustarPagina wsLote, 2
    MarcarSaltos wsLote, 2
    paginas = wsLote.PageSetup.Pages.Count
    Application.DisplayAlerts = False
    wsLote.Delete
    RestaurarEstado
    wsOrigen.Activate
    On Error GoTo 0

    msg = "DIAGNÓSTICO DE ETIQUETAS" & vbCrLf & String$(46, "-") & vbCrLf & _
          "Hoja de origen .......: " & wsOrigen.Name & vbCrLf & _
          "Filas visibles .......: " & nFilas & vbCrLf & _
          "Plantilla ............: " & wsPlantilla.Name & " (" & FILAS_ETIQUETA & " filas x A:" & COL_FINAL & ")" & vbCrLf & _
          "Etiquetas por trabajo : " & PAGINAS_POR_TRABAJO & vbCrLf & _
          "Impresora activa .....: " & ImpresoraActual() & vbCrLf & _
          "Papel de la plantilla : " & TamanoPapel(wsPlantilla) & vbCrLf & _
          "Escala medida ........: " & IIf(escala > 0, escala & " %", "no se pudo medir") & vbCrLf & _
          "Páginas de un lote de 2: " & paginas & IIf(paginas = 2, "  (correcto)", "  <-- REVISAR") & vbCrLf & _
          String$(46, "-") & vbCrLf & _
          "MAPEO CELDA <- COLUMNA" & vbCrLf

    For i = LBound(campos) To UBound(campos)
        msg = msg & "  " & PadDer(campos(i).Celda, 5) & " <- col " & PadDer(CStr(campos(i).Col), 3) & _
              "  " & IIf(Len(campos(i).Encabezado) = 0, "(posición fija)", campos(i).Encabezado) & _
              IIf(campos(i).Barras, "   [barras]", "") & vbCrLf
    Next i

    MsgBox msg, vbInformation, "Etiquetas"
End Sub

Private Function TamanoPapel(ByVal ws As Worksheet) As String
    On Error Resume Next
    TamanoPapel = "PaperSize=" & ws.PageSetup.PaperSize & _
                  "  Orientación=" & ws.PageSetup.Orientation & _
                  "  Área=" & ws.PageSetup.PrintArea
    On Error GoTo 0
End Function

Private Function PadDer(ByVal s As String, ByVal n As Long) As String
    PadDer = Left$(s & Space$(n), n)
End Function
