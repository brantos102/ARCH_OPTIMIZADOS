Attribute VB_Name = "modEtiquetas"
Option Explicit

'==============================================================================================
' modEtiquetas  -  Impresión de etiquetas en lote   (IMP_STOCK_DETALLADO.xlsm)
'----------------------------------------------------------------------------------------------
' Plantilla ....: hoja "ETQ", bloque A1:J17 = una etiqueta
' Origen .......: hoja "Consolidado" (tabla con autofiltro) u hoja "IMPRIMIR"
' Etiqueta .....: 10 x 5 cm (100 x 50 mm) en cualquier impresora - Zebra ZD421 203 ppp ZPL
'
' UNA SOLA ENTRADA
'     ImprimirEtiquetas ....... imprime las filas seleccionadas (o las del filtro)
'     EtiquetasDiagnostico .... herramienta de verificación, no imprime
'
' La combinación de teclas se asigna en Vista > Macros > ImprimirEtiquetas > Opciones...
' (a propósito no se guarda en el código: una línea Attribute pegada en el editor da
'  "Error de sintaxis", y este módulo tiene que poder importarse Y pegarse)
'
' CÓMO FUNCIONA
'   1. Toma las filas VISIBLES (nunca las ocultas por el autofiltro): las seleccionadas
'      si hay selección, o las del filtro completo si no la hay.
'   2. Pregunta una sola cosa: cuántas copias de cada fila (1 = una etiqueta por fila).
'   3. Arma una hoja temporal "ETQ_LOTE" con una copia de la plantilla por etiqueta, una
'      etiqueta por página, y la manda en UN SOLO trabajo de impresión por cada
'      PAGINAS_POR_TRABAJO etiquetas. Nunca imprime hoja por hoja: eso satura la cola.
'   4. La escala se calcula para que la etiqueta mida 10 x 5 cm exactos.
'   5. Borra la hoja temporal y deja el libro como estaba.
'
' NO hay vista previa: la ventana de vista previa de Excel tiene su propio botón Imprimir,
' y usarlo duplicaba el lote (se imprimía desde la vista previa y otra vez desde la macro).
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

' TAMAÑO FÍSICO DE LA ETIQUETA, en milímetros (10 x 5 cm).
' La escala de impresión se calcula para que el bloque A1:J17 mida exactamente esto,
' en la Zebra o en cualquier otra impresora. Si cambia el rollo, se cambia aquí.
Private Const ANCHO_ETIQUETA_MM  As Double = 100
Private Const ALTO_ETIQUETA_MM   As Double = 50

' Etiquetas que van en cada trabajo de impresión (cada página = una etiqueta).
' Si la cola del driver llegara a atascarse, baje a 50.
Private Const PAGINAS_POR_TRABAJO As Long = 100

' Topes de seguridad
Private Const TOPE_FILAS         As Long = 20000   ' filas visibles que se recorren
Private Const TOPE_ETIQUETAS     As Long = 10000   ' etiquetas por corrida
Private Const AVISO_DESDE        As Long = 300     ' desde aquí se pide confirmación extra

' Escala: 0 = calcularla con el tamaño de etiqueta (recomendado); 1-100 = forzar ese %
Private Const ESCALA_FIJA        As Long = 0

' Columna fija del campo "cliente / marca" (celda A2). Se usa posición fija porque en
' "Consolidado" la columna B es el nombre del cliente y en "IMPRIMIR" es el campo "ABC":
' en las dos hojas la columna B es el texto que va en la etiqueta.
Private Const COL_CLIENTE        As Long = 2

' Impresora
Private Const IMPRESORA_CONTIENE As String = "ZD421"
Private Const BUSCAR_IMPRESORA   As Boolean = True

' Compatibilidad / depuración
Private Const ACTUALIZAR_ETQ     As Boolean = True    ' deja la 1a etiqueta cargada en ETQ
Private Const CONSERVAR_LOTE     As Boolean = False   ' True = no borra la hoja ETQ_LOTE

'==============================================================================================


Private Type tCampo
    Celda As String          ' celda de la plantilla (esquina superior de la combinación)
    Encabezado As String     ' encabezado buscado en la fila 1 del origen ("" = posición fija)
    ColFija As Long          ' columna usada si no aparece el encabezado
    Col As Long              ' columna resuelta
    Barras As Boolean        ' True = se escribe *CÓDIGO* para la fuente Code 39
    Numero As Boolean        ' True = se conserva el valor numérico
End Type

Private mScreenUpdating  As Boolean
Private mEnableEvents    As Boolean
Private mDisplayAlerts   As Boolean
Private mCancelKey       As XlEnableCancelKey
Private mEstadoGuardado  As Boolean
Private mZoom            As Long
Private mAnchoMm         As Double
Private mAltoMm          As Double
Private mAvisos          As String
Private mNumAvisos       As Long


'==============================================================================================
' ENTRADA ÚNICA
'==============================================================================================
Public Sub ImprimirEtiquetas()

    Dim wsOrigen As Worksheet, wsPlantilla As Worksheet, wsLote As Worksheet
    Dim campos() As tCampo
    Dim filas() As Long, copias() As Long, cola() As Long
    Dim nFilas As Long, nCola As Long
    Dim soloSeleccion As Boolean, usarCantidad As Boolean
    Dim copiasFijas As Long
    Dim colProducto As Long, colSerie As Long, colCantidad As Long
    Dim desde As Long, hasta As Long, omitidas As Long
    Dim bloquesHoja As Long, k As Long, idx As Long
    Dim impresas As Long, nTrabajos As Long
    Dim i As Long, c As Long
    Dim hojaAnterior As Worksheet, selAnterior As Range
    Dim impresoraOrig As String, impresoraUsada As String
    Dim txt As String, msg As String
    Dim t0 As Single
    Dim nError As Long, sError As String

    mAvisos = "": mNumAvisos = 0: mZoom = 0: mAnchoMm = 0: mAltoMm = 0

    '--- 1. Contexto ---------------------------------------------------------------------
    If Not ValidarEntorno(wsOrigen, wsPlantilla) Then Exit Sub

    Set hojaAnterior = wsOrigen
    If TypeName(Selection) = "Range" Then Set selAnterior = Selection

    ResolverCampos wsOrigen, campos
    colProducto = ResolverColumna(wsOrigen, "producto_id", 3)
    colSerie = ResolverColumna(wsOrigen, "nro_serie", 5)
    colCantidad = ResolverColumna(wsOrigen, "cantidad", 13)

    '--- 2. Qué filas: primero manda la selección ----------------------------------------
    If TypeName(Selection) = "Range" Then
        Application.StatusBar = "Revisando la selección..."
        nFilas = RecolectarVisibles(RangoCandidato(wsOrigen, colProducto, True), TOPE_FILAS, filas)
        soloSeleccion = (nFilas > 0)
        Application.StatusBar = False
    End If

    desde = 1
    hasta = nFilas

    If Not soloSeleccion Then
        Application.StatusBar = "Explorando las filas visibles del filtro..."
        nFilas = RecolectarVisibles(RangoCandidato(wsOrigen, colProducto, False), TOPE_FILAS, filas)
        Application.StatusBar = False

        If nFilas = 0 Then
            MsgBox "No hay filas visibles para imprimir." & vbCrLf & vbCrLf & _
                   "Seleccione las filas que quiere etiquetar, o revise el filtro de la hoja """ & _
                   wsOrigen.Name & """.", vbExclamation, "Etiquetas"
            Exit Sub
        End If

        txt = InputBox("No hay filas seleccionadas." & vbCrLf & vbCrLf & _
            "Filas visibles del filtro: " & nFilas & IIf(nFilas >= TOPE_FILAS, " o más", "") & vbCrLf & vbCrLf & _
            "¿Cuántas desea imprimir?" & vbCrLf & _
            "     TODAS       las " & nFilas & " filas" & vbCrLf & _
            "     100         las primeras 100" & vbCrLf & _
            "     101-300     un rango", _
            "Etiquetas - filas del filtro", "TODAS")
        If Len(Trim$(txt)) = 0 Then Exit Sub

        desde = 1
        hasta = nFilas
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
    End If

    '--- 3. Impresora --------------------------------------------------------------------
    impresoraOrig = ImpresoraActual()
    If Not AsegurarImpresora() Then Exit Sub
    impresoraUsada = ImpresoraActual()

    '--- 4. Única pregunta: copias de cada fila ------------------------------------------
    msg = "ETIQUETAS" & vbCrLf & String$(44, "-") & vbCrLf & _
          "Filas a imprimir: " & (hasta - desde + 1) & _
          IIf(soloSeleccion, "   (seleccionadas)", "   (del filtro)") & vbCrLf & _
          "Impresora: " & impresoraUsada & vbCrLf & _
          String$(44, "-") & vbCrLf & vbCrLf & _
          "Copias de CADA fila:" & vbCrLf & vbCrLf & _
          "     1        una etiqueta por fila" & vbCrLf & _
          "     2, 3 ... ese número de copias de cada fila" & vbCrLf & _
          "     C        usar la columna CANTIDAD de cada fila"

    txt = InputBox(msg, "Etiquetas - copias", "1")
    If Len(Trim$(txt)) = 0 Then Exit Sub

    txt = UCase$(Trim$(txt))
    If Left$(txt, 1) = "C" Then
        usarCantidad = True
        copiasFijas = 1
    Else
        copiasFijas = CLng(Val(txt))
        If copiasFijas < 1 Then copiasFijas = 1
    End If

    '--- 5. Cola de etiquetas ------------------------------------------------------------
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
        MsgBox "Las filas elegidas no tienen código ni número de serie: no hay nada que imprimir.", _
               vbExclamation, "Etiquetas"
        Exit Sub
    End If

    If nCola > TOPE_ETIQUETAS Then
        MsgBox "El lote sería de " & nCola & " etiquetas y el tope de seguridad es " & _
               TOPE_ETIQUETAS & "." & vbCrLf & vbCrLf & _
               "Reduzca las filas o las copias, o suba TOPE_ETIQUETAS en el módulo.", _
               vbExclamation, "Etiquetas"
        Exit Sub
    End If

    If nCola >= AVISO_DESDE Then
        If MsgBox("Va a imprimir " & nCola & " etiquetas en " & _
                  ((nCola + PAGINAS_POR_TRABAJO - 1) \ PAGINAS_POR_TRABAJO) & " trabajo(s) de impresión." & vbCrLf & vbCrLf & _
                  "Impresora: " & impresoraUsada & vbCrLf & vbCrLf & _
                  "¿Continuar?", vbYesNo + vbQuestion, "Etiquetas") <> vbYes Then
            RestaurarImpresora impresoraOrig, impresoraUsada
            Exit Sub
        End If
    End If

    ReDim cola(1 To nCola)
    idx = 0
    For i = desde To hasta
        For c = 1 To copias(i)
            idx = idx + 1
            cola(idx) = filas(i)
        Next c
    Next i

    '--- 6. Impresión --------------------------------------------------------------------
    t0 = Timer
    On Error GoTo Limpieza
    GuardarEstado

    Application.StatusBar = "Preparando la plantilla..."
    Set wsLote = CrearHojaLote(wsPlantilla)

    ' la escala y el área imprimible se miden sobre la hoja recién copiada (es chica y rápida)
    MedirAreaImprimible wsLote
    mZoom = DeterminarZoom(wsLote)
    If mZoom <= 0 Then
        Err.Raise vbObjectError + 1, , "No se pudo calcular la escala de impresión de la etiqueta."
    End If
    If Not AreaSuficiente() Then GoTo Limpieza

    PrepararFormatos wsLote, campos
    bloquesHoja = MinL(nCola, PAGINAS_POR_TRABAJO)
    ReplicarBloques wsLote, bloquesHoja
    MarcarSaltos wsLote, bloquesHoja

    idx = 1
    Do While idx <= nCola
        k = MinL(bloquesHoja, nCola - idx + 1)

        Application.StatusBar = "Preparando etiquetas " & idx & " a " & (idx + k - 1) & " de " & nCola & "..."
        LlenarBloques wsLote, wsOrigen, cola, idx, k, campos
        AjustarPagina wsLote, k

        Application.StatusBar = "Enviando a imprimir: etiquetas " & idx & " a " & (idx + k - 1) & _
                               " de " & nCola & "  (" & impresoraUsada & ")"
        DoEvents
        wsLote.PrintOut Copies:=1, Collate:=True
        DoEvents

        nTrabajos = nTrabajos + 1
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
        If Not CONSERVAR_LOTE Then
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
        msg = "Listo: " & impresas & " etiqueta(s) enviada(s) a imprimir." & vbCrLf & String$(44, "-") & vbCrLf & _
              "Impresora ....: " & impresoraUsada & vbCrLf & _
              "Trabajos .....: " & nTrabajos & "   (una página = una etiqueta)" & vbCrLf & _
              "Escala .......: " & mZoom & " %   para " & Format$(ANCHO_ETIQUETA_MM, "0") & " x " & _
                                    Format$(ALTO_ETIQUETA_MM, "0") & " mm" & vbCrLf & _
              "Filas omitidas: " & omitidas & vbCrLf & _
              "Tiempo .......: " & Format$(Timer - t0, "0.0") & " s"
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
' FILAS VISIBLES  (respeta el autofiltro)
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
' Trabaja por bloques para no agotar SpecialCells cuando hay decenas de miles de filas ocultas.
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
    FilaImprimible = (Len(TextoCelda(ws.Cells(fila, colProducto))) > 0) Or _
                     (Len(TextoCelda(ws.Cells(fila, colSerie))) > 0)
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
    ' sin esto Excel vuelve a paginar en CADA escritura: es lo que dejaba el libro colgado
    ws.DisplayPageBreaks = False
    On Error GoTo 0

    ' la etiqueta son sólo las filas 1 a 17: se limpia todo lo que haya debajo
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

' Formato de texto en las celdas de la etiqueta ANTES de replicar: así se copia a todos los
' bloques de una vez y el llenado posterior sólo escribe valores (mucho más rápido).
Private Sub PrepararFormatos(ByVal ws As Worksheet, ByRef campos() As tCampo)
    Dim i As Long

    On Error Resume Next
    For i = LBound(campos) To UBound(campos)
        If Not campos(i).Numero Then ws.Range(campos(i).Celda).NumberFormat = "@"
    Next i
    On Error GoTo 0
End Sub

' Replica el bloque de 17 filas hasta tener "n" etiquetas, duplicando lo ya copiado
' (1, 2, 4, 8 ...): para 100 etiquetas son 7 operaciones de copiado, no 100.
Private Sub ReplicarBloques(ByVal ws As Worksheet, ByVal n As Long)

    Dim hay As Long, nuevos As Long

    Application.CutCopyMode = False
    hay = 1
    Do While hay < n
        nuevos = MinL(hay, n - hay)
        ws.Rows("1:" & (FILAS_ETIQUETA * nuevos)).Copy _
            Destination:=ws.Rows((FILAS_ETIQUETA * hay + 1) & ":" & (FILAS_ETIQUETA * (hay + nuevos)))
        hay = hay + nuevos
        DoEvents
    Loop
    Application.CutCopyMode = False
End Sub

' Un salto de página al inicio de cada etiqueta: una etiqueta = una página.
Private Sub MarcarSaltos(ByVal ws As Worksheet, ByVal n As Long)
    Dim i As Long

    On Error Resume Next
    ws.ResetAllPageBreaks
    For i = 1 To n - 1
        ws.Rows(FILAS_ETIQUETA * i + 1).PageBreak = xlPageBreakManual
    Next i
    On Error GoTo 0
End Sub

Private Sub AjustarPagina(ByVal ws As Worksheet, ByVal k As Long)

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$A$1:$" & COL_FINAL & "$" & (FILAS_ETIQUETA * k)
        .Zoom = mZoom
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub


'----------------------------------------------------------------------------------------------
' ESCALA DE IMPRESIÓN
'
' La etiqueta mide siempre 10 x 5 cm, en la Zebra o en cualquier otra impresora. Por eso la
' escala no se hereda del "ajustar a 1 página" de la hoja ETQ (que depende del papel del
' driver y, aplicado a una hoja de varias etiquetas, se calcularía sobre el lote entero):
' se calcula para que el bloque A1:J17 ocupe exactamente la etiqueta.
'
'     escala = el menor de   ancho etiqueta / ancho del bloque
'                            alto  etiqueta / alto  del bloque
'
' Range.Width y Range.Height devuelven puntos, así que el cálculo es exacto.
'----------------------------------------------------------------------------------------------
Private Function DeterminarZoom(ByVal ws As Worksheet) As Long

    Dim bloque As Range
    Dim anchoPt As Double, altoPt As Double
    Dim escalaAncho As Double, escalaAlto As Double, escala As Double
    Dim z As Long

    If ESCALA_FIJA > 0 Then
        DeterminarZoom = ESCALA_FIJA
        Exit Function
    End If

    On Error Resume Next
    Set bloque = ws.Range("A1:" & COL_FINAL & FILAS_ETIQUETA)
    If bloque Is Nothing Then Exit Function
    anchoPt = bloque.Width
    altoPt = bloque.Height
    On Error GoTo 0

    If anchoPt <= 0 Or altoPt <= 0 Then Exit Function

    escalaAncho = MilimetrosAPuntos(ANCHO_ETIQUETA_MM) / anchoPt
    escalaAlto = MilimetrosAPuntos(ALTO_ETIQUETA_MM) / altoPt

    escala = escalaAncho
    If escalaAlto < escala Then escala = escalaAlto

    z = Int(escala * 100)              ' hacia abajo: nunca pasarse del borde
    If z < 10 Then z = 10              ' 10 % es el mínimo que admite Excel
    If z > 400 Then z = 400

    DeterminarZoom = z
End Function

Private Function MilimetrosAPuntos(ByVal mm As Double) As Double
    MilimetrosAPuntos = mm * 72# / 25.4
End Function

Private Function PuntosAMilimetros(ByVal pt As Double) As Double
    PuntosAMilimetros = pt * 25.4 / 72#
End Function


'----------------------------------------------------------------------------------------------
' ÁREA IMPRIMIBLE REAL
'
' Mide cuánto espacio da la impresora activa: al 100 % se mira dónde corta Excel la página y
' se suman los anchos de columna y los altos de fila que entraron. Sirve para avisar si el
' papel del driver es más chico que la etiqueta. Se hace sobre la hoja recién copiada, con
' pocas filas, para que sea instantáneo.
'----------------------------------------------------------------------------------------------
Private Sub MedirAreaImprimible(ByVal ws As Worksheet)

    Dim col As Long, fila As Long
    Dim screenAnt As Boolean

    mAnchoMm = 0
    mAltoMm = 0

    screenAnt = Application.ScreenUpdating
    Application.ScreenUpdating = True      ' Excel sólo calcula los saltos automáticos así

    On Error Resume Next
    ws.DisplayPageBreaks = True
    Application.PrintCommunication = False
    With ws.PageSetup
        .Zoom = 100
        .PrintArea = "$A$1:$Z$" & (FILAS_ETIQUETA * 2)
    End With
    Application.PrintCommunication = True
    ws.ResetAllPageBreaks

    If ws.VPageBreaks.Count > 0 Then
        col = ws.VPageBreaks(1).Location.Column - 1
        If col >= 1 Then mAnchoMm = PuntosAMilimetros(ws.Range(ws.Cells(1, 1), ws.Cells(1, col)).Width)
    End If

    If ws.HPageBreaks.Count > 0 Then
        fila = ws.HPageBreaks(1).Location.Row - 1
        If fila >= 1 Then mAltoMm = PuntosAMilimetros(ws.Range(ws.Cells(1, 1), ws.Cells(fila, 1)).Height)
    End If

    ws.DisplayPageBreaks = False
    On Error GoTo 0

    Application.ScreenUpdating = screenAnt
End Sub

Private Function AreaImprimibleTexto() As String
    If mAnchoMm <= 0 And mAltoMm <= 0 Then
        AreaImprimibleTexto = "no se pudo medir"
    Else
        AreaImprimibleTexto = Format$(mAnchoMm, "0.0") & " x " & Format$(mAltoMm, "0.0") & " mm"
    End If
End Function

' Avisa si la impresora da menos espacio que la etiqueta: saldría cortada.
Private Function AreaSuficiente() As Boolean

    Const HOLGURA As Double = 1.5      ' mm de tolerancia

    AreaSuficiente = True
    If mAnchoMm <= 0 Or mAltoMm <= 0 Then Exit Function
    If mAnchoMm >= ANCHO_ETIQUETA_MM - HOLGURA And mAltoMm >= ALTO_ETIQUETA_MM - HOLGURA Then Exit Function

    Application.ScreenUpdating = True
    AreaSuficiente = (MsgBox("El papel configurado en la impresora es más chico que la etiqueta." & vbCrLf & vbCrLf & _
        "Etiqueta necesaria ..: " & Format$(ANCHO_ETIQUETA_MM, "0") & " x " & Format$(ALTO_ETIQUETA_MM, "0") & " mm" & vbCrLf & _
        "Área imprimible .....: " & AreaImprimibleTexto() & vbCrLf & _
        "Impresora ...........: " & ImpresoraActual() & vbCrLf & vbCrLf & _
        "Si continúa, las etiquetas saldrán cortadas. Corrija el tamaño de papel en las" & vbCrLf & _
        "preferencias de la impresora (tamaño definido por el usuario: 100 x 50 mm)." & vbCrLf & vbCrLf & _
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
                          wsOrigen.Cells(fila, campos(i).Col), campos(i), fila
        Next i
        If b Mod 10 = 0 Then DoEvents      ' que Excel no se vea "sin responder"
    Next b
End Sub

' Carga la primera etiqueta del lote en la plantilla ETQ (como hacía la macro original)
Private Sub LlenarEtiquetaPlantilla(ByVal wsPlantilla As Worksheet, ByVal wsOrigen As Worksheet, _
                                    ByVal fila As Long, ByRef campos() As tCampo)
    Dim i As Long

    On Error Resume Next
    For i = LBound(campos) To UBound(campos)
        EscribirCampo wsPlantilla.Range(campos(i).Celda), _
                      wsOrigen.Cells(fila, campos(i).Col), campos(i), fila
    Next i
    On Error GoTo 0
End Sub

Private Sub EscribirCampo(ByVal destino As Range, ByVal origen As Range, ByRef campo As tCampo, _
                          ByVal fila As Long)
    Dim v As Variant
    Dim s As String

    If campo.Numero Then
        v = origen.Value
        If Not IsError(v) Then
            If EsNumero(v) Then
                destino.Value = v
                Exit Sub
            End If
        End If
    End If

    s = TextoCelda(origen)
    If campo.Barras Then s = CodigoBarras(s, fila)
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
' asteriscos de inicio/fin. Si hay caracteres que Code 39 no codifica, avisa.
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
        AgregarAviso "fila " & fila & ": código de " & Len(t) & " caracteres, las barras quedan muy comprimidas"
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

Private Sub RestaurarImpresora(ByVal original As String, ByVal usada As String)
    If Len(original) = 0 Then Exit Sub
    If usada = original Then Exit Sub

    On Error Resume Next
    Application.ActivePrinter = original
    On Error GoTo 0
End Sub


'==============================================================================================
' ESTADO DE EXCEL
'==============================================================================================
Private Sub GuardarEstado()
    If mEstadoGuardado Then Exit Sub

    mScreenUpdating = Application.ScreenUpdating
    mEnableEvents = Application.EnableEvents
    mDisplayAlerts = Application.DisplayAlerts
    mCancelKey = Application.EnableCancelKey
    mEstadoGuardado = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    ' si el operador pulsa Esc, se cancela limpio en vez de abrir el editor de VBA
    Application.EnableCancelKey = xlErrorHandler
End Sub

Private Sub RestaurarEstado()
    If Not mEstadoGuardado Then Exit Sub

    On Error Resume Next
    Application.EnableCancelKey = mCancelKey
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


'==============================================================================================
' DIAGNÓSTICO  (no imprime)
'==============================================================================================
Public Sub EtiquetasDiagnostico()

    Dim wsOrigen As Worksheet, wsPlantilla As Worksheet, wsLote As Worksheet
    Dim campos() As tCampo
    Dim filas() As Long
    Dim nFilas As Long, nSel As Long, i As Long
    Dim escala As Long, paginas As Long
    Dim msg As String

    If Not ValidarEntorno(wsOrigen, wsPlantilla) Then Exit Sub

    ResolverCampos wsOrigen, campos

    If TypeName(Selection) = "Range" Then
        nSel = RecolectarVisibles(RangoCandidato(wsOrigen, ResolverColumna(wsOrigen, "producto_id", 3), True), _
                                  TOPE_FILAS, filas)
    End If
    nFilas = RecolectarVisibles(RangoCandidato(wsOrigen, ResolverColumna(wsOrigen, "producto_id", 3), False), _
                                TOPE_FILAS, filas)

    ' lote de prueba de 2 etiquetas para medir escala y paginación reales
    On Error Resume Next
    GuardarEstado
    Set wsLote = CrearHojaLote(wsPlantilla)
    MedirAreaImprimible wsLote
    escala = DeterminarZoom(wsLote)
    mZoom = escala
    ReplicarBloques wsLote, 2
    MarcarSaltos wsLote, 2
    AjustarPagina wsLote, 2
    paginas = wsLote.PageSetup.Pages.Count
    Application.DisplayAlerts = False
    wsLote.Delete
    RestaurarEstado
    wsOrigen.Activate
    On Error GoTo 0

    msg = "DIAGNÓSTICO DE ETIQUETAS" & vbCrLf & String$(46, "-") & vbCrLf & _
          "Hoja de origen .........: " & wsOrigen.Name & vbCrLf & _
          "Filas seleccionadas ....: " & nSel & vbCrLf & _
          "Filas visibles del filtro: " & nFilas & IIf(nFilas >= TOPE_FILAS, " o más", "") & vbCrLf & _
          "Etiquetas por trabajo ..: " & PAGINAS_POR_TRABAJO & vbCrLf & _
          "Impresora activa .......: " & ImpresoraActual() & vbCrLf & _
          "Etiqueta configurada ...: " & Format$(ANCHO_ETIQUETA_MM, "0") & " x " & _
                                          Format$(ALTO_ETIQUETA_MM, "0") & " mm" & vbCrLf & _
          "Área imprimible real ...: " & AreaImprimibleTexto() & AvisoArea() & vbCrLf & _
          "Escala calculada .......: " & IIf(escala > 0, escala & " %", "no se pudo calcular") & vbCrLf & _
          "Páginas de un lote de 2 : " & paginas & IIf(paginas = 2, "   (correcto)", "   <-- REVISAR") & vbCrLf & _
          String$(46, "-") & vbCrLf & _
          "MAPEO CELDA <- COLUMNA" & vbCrLf

    For i = LBound(campos) To UBound(campos)
        msg = msg & "  " & PadDer(campos(i).Celda, 5) & " <- col " & PadDer(CStr(campos(i).Col), 3) & _
              "  " & IIf(Len(campos(i).Encabezado) = 0, "(posición fija)", campos(i).Encabezado) & _
              IIf(campos(i).Barras, "   [barras]", "") & vbCrLf
    Next i

    MsgBox msg, vbInformation, "Etiquetas"
End Sub

Private Function AvisoArea() As String
    If mAnchoMm <= 0 Or mAltoMm <= 0 Then Exit Function
    If mAnchoMm < ANCHO_ETIQUETA_MM - 1.5 Or mAltoMm < ALTO_ETIQUETA_MM - 1.5 Then
        AvisoArea = "   <-- MÁS CHICA QUE LA ETIQUETA"
    End If
End Function

Private Function PadDer(ByVal s As String, ByVal n As Long) As String
    PadDer = Left$(s & Space$(n), n)
End Function
