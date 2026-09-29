Option Explicit
' =====================================================================================
'  modEGR - Formato EGR_FL_HYCITE (etapa 2: recepción, destinos, avance y reportes)
'  Reemplaza al Módulo7 (conserva los nombres TABLAS, TABLA_APIS y ActualizarTodo, que usa
'  el botón ACTUALIZAR de EMPAQUETADO).
'   - ActualizarTodo SEGURO: respaldo previo, consulta por consulta, sin tocar el modelo
'     de datos directamente, con registro de errores (antes: "Error de automatización" y
'     Excel se cerraba).
'   - RepararFormulasEGR: corrige los #REF! (DATOS!Y, EMPAQUETADO!F, COD POS I:J,
'     nombres rotos y vínculos externos) y deja DATOS!A preparado para las reglas.
'   - Reglas de destino (hoja REGLAS_DESTINO, editable sin código): PRO / GYE / UIO / GPS.
'   - Revisión de cobertura TMS (segunda validación, sin revalidar lo de PEDIDOS HCE).
'   - Avance de picking y empaque, SKU/cajas sin datos de costo.
'   - Exportación de TMS, TRAMACO y DESPACHOS a CSV / XLSX / PDF sin cambiar sus formatos.
'   - Registro de acciones y errores (hoja oculta LOG_EGR + panel).
' =====================================================================================
Public Const HDAT As String = "DATOS"
Public Const HREG As String = "REGLAS_DESTINO"
Public Const HCFG As String = "CONFIG_EGR"
Public Const HLOGE As String = "LOG_EGR"
Public Const MAXF As Long = 500           ' última fila de fórmulas de DATOS

' columnas de DATOS
Public Const D_DEST As Long = 1, D_PED As Long = 2, D_DIR As Long = 4, D_NOM As Long = 7
Public Const D_H As Long = 8, D_I As Long = 9, D_J As Long = 10, D_TEL As Long = 11
Public Const D_PROV As Long = 15, D_CANT As Long = 16, D_PARR As Long = 17
Public Const D_BUL As Long = 18, D_PESO As Long = 19, D_VAL As Long = 22, D_COUR As Long = 25
Public Const D_PARRSP As Long = 26, D_CPTMS As Long = 28, D_DIAG As Long = 30, D_SUG As Long = 31
Public Const D_RDEST As Long = 32, D_RPED As Long = 33, D_RMOT As Long = 34   ' AF, AG, AH (reglas confirmadas)

Public gOcupadoE As Boolean
Public gLogE As Collection
Public gPropuestas As Collection      ' Array(fila, pedido, destinatario, prov, canton, parroquia, actual, nuevo, motivo)

' =====================================================================================
'  Utilidades
' =====================================================================================
Public Function TXE(v As Variant) As String
  If IsError(v) Then TXE = "" Else TXE = Trim$(CStr(v & ""))
End Function

Public Function NormE(ByVal s As String) As String
  Dim a, b, i As Long, t As String, o As String, ch As String, p1 As Long, p2 As Long
  a = Array("Á", "É", "Í", "Ó", "Ú", "Ä", "Ë", "Ï", "Ö", "Ü", "Ñ", "À", "È", "Ì", "Ò", "Ù")
  b = Array("A", "E", "I", "O", "U", "A", "E", "I", "O", "U", "N", "A", "E", "I", "O", "U")
  t = UCase$(Trim$(Replace(s, "-", " ")))
  For i = LBound(a) To UBound(a): t = Replace(t, a(i), b(i)): Next
  Do While InStr(t, "(") > 0
    p1 = InStr(t, "("): p2 = InStr(p1, t, ")")
    If p2 > p1 Then t = Left$(t, p1 - 1) & " " & Mid$(t, p2 + 1) Else t = Replace(t, "(", " ")
  Loop
  For i = 1 To Len(t)
    ch = Mid$(t, i, 1)
    If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then o = o & ch Else o = o & " "
  Next
  Do While InStr(o, "  ") > 0: o = Replace(o, "  ", " "): Loop
  o = " " & Trim$(o) & " "
  o = Replace(o, " PTO ", " PUERTO "): o = Replace(o, " FCO ", " FRANCISCO "): o = Replace(o, " GRAL ", " GENERAL ")
  NormE = Trim$(o)
End Function

Private Function Hoja(ByVal nombre As String) As Worksheet
  On Error Resume Next
  Set Hoja = ThisWorkbook.Worksheets(nombre)
  On Error GoTo 0
End Function

Public Function UltimaFilaDatos() As Long
  Dim ws As Worksheet, v, i As Long
  Set ws = ThisWorkbook.Worksheets(HDAT)
  v = ws.Range(ws.Cells(1, D_PED), ws.Cells(MAXF, D_PED)).Value
  For i = MAXF To 2 Step -1
    If Len(TXE(v(i, 1))) > 0 Then UltimaFilaDatos = i: Exit Function
  Next
  UltimaFilaDatos = 1
End Function

Public Sub Desbloquear()
  Application.ScreenUpdating = True: Application.EnableEvents = True
  Application.DisplayAlerts = True: Application.Calculation = xlCalculationAutomatic
  Application.Cursor = xlDefault: Application.StatusBar = False
  gOcupadoE = False
  LogE "Desbloquear: pantalla, eventos y cálculo restaurados"
End Sub

Private Sub Congelar()
  Application.ScreenUpdating = False: Application.EnableEvents = False
  Application.Calculation = xlCalculationManual: Application.Cursor = xlWait
End Sub

Private Sub Liberar()
  Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True
  Application.ScreenUpdating = True: Application.DisplayAlerts = True
  Application.Cursor = xlDefault: Application.StatusBar = False
End Sub

' ---------- configuración (hoja oculta CONFIG_EGR: clave en A, valor en B) ----------
Public Function Cfg(ByVal clave As String, Optional ByVal defecto As String = "") As String
  Dim ws As Worksheet, r As Long
  Set ws = HojaCfg()
  For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If UCase$(TXE(ws.Cells(r, 1).Value)) = UCase$(clave) Then Cfg = TXE(ws.Cells(r, 2).Value): Exit Function
  Next
  Cfg = defecto
End Function

Public Sub SetCfg(ByVal clave As String, ByVal valor As String)
  Dim ws As Worksheet, r As Long, lr As Long
  Set ws = HojaCfg()
  lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  For r = 2 To lr
    If UCase$(TXE(ws.Cells(r, 1).Value)) = UCase$(clave) Then ws.Cells(r, 2).Value = valor: Exit Sub
  Next
  ws.Cells(lr + 1, 1).Value = clave: ws.Cells(lr + 1, 2).NumberFormat = "@": ws.Cells(lr + 1, 2).Value = valor
End Sub

Private Function HojaCfg() As Worksheet
  Dim ws As Worksheet
  Set ws = Hoja(HCFG)
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = HCFG
    ws.Range("A1:C1").Value = Array("CLAVE", "VALOR", "DESCRIPCION")
    ws.Range("A2:C2").Value = Array("IMPRESORA_ZEBRA", "", "Nombre de la impresora Zebra ZD230 (se elige desde el panel)")
    ws.Range("A3:C3").Value = Array("CSV_SEPARADOR_PUNTOYCOMA", "NO", "SI = CSV con ';' (configuración regional). NO = CSV con ','")
    ws.Range("A4:C4").Value = Array("CARPETA_EXPORTES", "", "Vacío = carpeta EXPORTES junto a este archivo")
    ws.Range("A5:C5").Value = Array("CORREO_PARA", "", "Destinatarios del correo de reportes (separados por ;)")
    ws.Range("A6:C6").Value = Array("ETIQ_OFFSET_X", "0", "Corrimiento horizontal de la etiqueta en puntos (8 = 1 mm)")
    ws.Range("A7:C7").Value = Array("ETIQ_OFFSET_Y", "0", "Corrimiento vertical de la etiqueta en puntos (8 = 1 mm)")
    ws.Columns("A:C").AutoFit
    ws.Visible = xlSheetHidden
  End If
  Set HojaCfg = ws
End Function

' ---------- registro ----------
Public Sub LogE(ByVal msg As String, Optional ByVal nivel As String = "INFO")
  Dim linea As String, ws As Worksheet, r As Long
  linea = Format(Now, "hh:nn:ss") & "  " & IIf(nivel = "INFO", "", "[" & nivel & "] ") & msg
  If gLogE Is Nothing Then Set gLogE = New Collection
  gLogE.Add linea
  If gLogE.Count > 400 Then gLogE.Remove 1
  On Error Resume Next
  Dim i As Long
  For i = 0 To VBA.UserForms.Count - 1
    If VBA.UserForms(i).Name = "frmEGR" Then VBA.UserForms(i).AgregarLog linea
  Next
  Set ws = Hoja(HLOGE)
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = HLOGE
    ws.Range("A1:D1").Value = Array("FECHA", "USUARIO", "NIVEL", "MENSAJE")
    ws.Visible = xlSheetHidden
  End If
  r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
  If r > 20000 Then ws.Rows("2:5001").Delete: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
  ws.Cells(r, 1).Value = Now: ws.Cells(r, 2).Value = Application.UserName
  ws.Cells(r, 3).Value = nivel: ws.Cells(r, 4).Value = msg
End Sub

' ---------- respaldo antes de procesos delicados ----------
Public Function Respaldo(ByVal motivo As String) As Boolean
  Dim carp As String, f As String
  On Error GoTo fallo
  If Len(ThisWorkbook.Path) = 0 Then Exit Function
  carp = ThisWorkbook.Path & Application.PathSeparator & "RESPALDOS_EGR"
  If Len(Dir(carp, vbDirectory)) = 0 Then MkDir carp
  f = carp & Application.PathSeparator & "EGR_" & Format(Now, "yyyymmdd_hhnnss") & "_" & motivo & ".xlsm"
  ThisWorkbook.SaveCopyAs f
  LogE "Respaldo creado: " & f
  LimpiarRespaldos carp
  Respaldo = True
  Exit Function
fallo:
  LogE "No se pudo crear el respaldo: " & Err.Description, "AVISO"
End Function

Private Sub LimpiarRespaldos(ByVal carp As String)
  Dim f As String, col As New Collection, i As Long, j As Long, t As String, arr() As String
  On Error Resume Next
  f = Dir(carp & Application.PathSeparator & "EGR_*.xlsm")
  Do While Len(f) > 0: col.Add f: f = Dir(): Loop
  If col.Count <= 15 Then Exit Sub
  ReDim arr(1 To col.Count)
  For i = 1 To col.Count: arr(i) = col(i): Next
  For i = 1 To UBound(arr) - 1
    For j = i + 1 To UBound(arr)
      If arr(j) < arr(i) Then t = arr(i): arr(i) = arr(j): arr(j) = t
    Next
  Next
  For i = 1 To UBound(arr) - 15: Kill carp & Application.PathSeparator & arr(i): Next
End Sub

' =====================================================================================
'  1. ACTUALIZAR TODO (seguro)
' =====================================================================================
Sub ActualizarTodo()
  ' Botón de operación: se puede pulsar en cualquier momento. Trae ITEMS API, ITEMS DEPOT (ODBC),
  ' EMPAQUETADO (Google Sheets) y actualiza las tablas dinámicas. Si una fuente falla, pregunta si sigue.
  Dim ws As Worksheet, lo As ListObject, pc As PivotCache, t0 As Single, nOk As Long, nErr As Long
  Dim wb As Workbook, pend As String, motivo As String, nHoy As Long, empOk As Boolean
  If gOcupadoE Then MsgBox "Hay un proceso en curso.", vbInformation: Exit Sub
  RuedaDesactivar
  For Each wb In Application.Workbooks
    If Not wb Is ThisWorkbook And Not wb.Saved And Len(wb.Path) > 0 And Not wb.ReadOnly Then pend = pend & vbCrLf & "  - " & wb.Name
  Next
  If Len(pend) > 0 Then
    Select Case MsgBox("Hay libros abiertos sin guardar:" & pend & vbCrLf & vbCrLf & "¿Guardarlos antes de actualizar? (recomendado)", vbYesNoCancel + vbQuestion, "Actualizar datos")
      Case vbCancel: Exit Sub
      Case vbYes
        On Error Resume Next
        For Each wb In Application.Workbooks
          If Not wb Is ThisWorkbook And Not wb.Saved And Len(wb.Path) > 0 And Not wb.ReadOnly Then wb.Save
        Next
        On Error GoTo 0
    End Select
  End If
  gOcupadoE = True
  Respaldo "antes_actualizar"
  LogE "ACTUALIZAR: inicio"
  t0 = Timer
  ' 1) EMPAQUETADO: se comprueba ANTES Google Sheets. Refrescar la consulta con la hoja vacía o sin
  '    conexión es lo que cerraba Excel ("Error de automatización"), así que en ese caso NO se refresca.
  Application.StatusBar = "Comprobando Google Sheets (EMPAQUETADO)..."
  empOk = EmpaquetadoDisponible(motivo, nHoy)
  If empOk Then
    LogE "EMPAQUETADO: Google Sheets responde, " & nHoy & " fila(s) de hoy"
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
      For Each lo In ws.ListObjects
        If lo.SourceType = 4 Then RefrescarTablaModelo lo, nOk, nErr        ' 4 = xlSrcModel
      Next
    Next
    Application.ScreenUpdating = True
  Else
    nErr = nErr + 1
    LogE "EMPAQUETADO no se actualizó: " & motivo & ". Se conservan los datos anteriores.", "AVISO"
    If MsgBox("No se obtuvieron datos de EMPAQUETADO (Google Sheets):" & vbCrLf & "  " & motivo & vbCrLf & vbCrLf & _
              "Se conservan los datos anteriores de empaquetado." & vbCrLf & _
              "¿Sigo con las demás conexiones (ITEMS API, ITEMS DEPOT) y las tablas dinámicas?", vbYesNo + vbQuestion, "Actualizar datos") <> vbYes Then
      Call Liberar: gOcupadoE = False
      LogE "ACTUALIZAR: cancelado por el usuario"
      Exit Sub
    End If
  End If
  ' 2) consultas ODBC cargadas a tabla (Estado, ITEMS API), una por una
  Application.ScreenUpdating = False: Application.EnableEvents = False
  For Each ws In ThisWorkbook.Worksheets
    For Each lo In ws.ListObjects
      If lo.SourceType = xlSrcQuery And Not (Not empOk And InStr(1, lo.Name, "EMPAQUETADO", vbTextCompare) > 0) Then
        If Not RefrescarTabla(lo, nOk, nErr) Then
          Application.ScreenUpdating = True
          If MsgBox("No se pudo actualizar " & lo.Name & " (revisa la red / ODBC DEPOTUIO)." & vbCrLf & _
                    "¿Sigo con las demás conexiones y las tablas dinámicas?", vbYesNo + vbQuestion, "Actualizar datos") <> vbYes Then
            Call Liberar: gOcupadoE = False
            LogE "ACTUALIZAR: detenido por el usuario tras el error en " & lo.Name, "AVISO"
            Exit Sub
          End If
          Application.ScreenUpdating = False
        End If
      End If
    Next
  Next
  ' 3) tablas dinámicas: cada caché una vez
  For Each pc In ThisWorkbook.PivotCaches
    Application.StatusBar = "Actualizando tablas dinámicas..."
    Err.Clear
    On Error Resume Next
    pc.Refresh
    If Err.Number <> 0 Then
      nErr = nErr + 1: LogE "Tabla dinámica (caché " & pc.Index & "): " & Err.Description, "ERROR"
    Else
      nOk = nOk + 1
    End If
    On Error GoTo 0
    DoEvents
  Next
  Application.Calculate
  Liberar
  gOcupadoE = False
  LogE "ACTUALIZAR: terminado en " & Format(Timer - t0, "0.0") & " s. Correctos " & nOk & ", con aviso/error " & nErr, IIf(nErr > 0, "AVISO", "INFO")
  RefrescarPanel
  If nErr > 0 Then
    MsgBox "Actualización terminada con " & nErr & " aviso(s). El detalle está en el registro del panel." & vbCrLf & _
           "Hay respaldo en la carpeta RESPALDOS_EGR.", vbExclamation, "Actualizar datos"
  Else
    MsgBox "Datos y tablas dinámicas actualizados.", vbInformation, "Listo"
  End If
End Sub

' Comprueba que el Google Sheets de EMPAQUETADO responde y tiene filas con fecha de HOY
Private Function EmpaquetadoDisponible(ByRef motivo As String, ByRef nHoy As Long) As Boolean
  Dim q As Object, url As String, re As Object, m As Object, http As Object, body As String
  Dim lineas, i As Long, c As String, d As Date
  motivo = "": nHoy = 0
  On Error Resume Next
  Set q = ThisWorkbook.Queries("EMPAQUETADO")
  On Error GoTo 0
  If q Is Nothing Then motivo = "no existe la consulta EMPAQUETADO": Exit Function
  Set re = CreateObject("VBScript.RegExp"): re.Pattern = "Web\.Contents\(""([^""]+)"""
  If Not re.Test(q.Formula) Then EmpaquetadoDisponible = True: Exit Function     ' otra fuente: se refresca normal
  url = re.Execute(q.Formula)(0).SubMatches(0)
  On Error Resume Next
  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  http.setTimeouts 5000, 5000, 15000, 20000
  http.Open "GET", url, False
  http.send
  If Err.Number <> 0 Then                   ' con proxy corporativo: se usa la conexión de Windows/Office
    Err.Clear
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
  End If
  If Err.Number <> 0 Then GoTo fallo
  On Error GoTo fallo
  If http.Status <> 200 Then motivo = "Google Sheets respondió " & http.Status: Exit Function
  body = http.responseText
  lineas = Split(Replace(body, vbCr, ""), vbLf)
  For i = 1 To UBound(lineas)
    c = Split(lineas(i) & ",", ",")(0)
    If EsFechaHoy(Replace(c, """", "")) Then
      If Len(Trim$(Replace(Split(lineas(i) & ",,", ",")(1), """", ""))) > 0 Then nHoy = nHoy + 1
    End If
  Next
  If nHoy = 0 Then motivo = "la hoja de Google Sheets no tiene órdenes con fecha de hoy (" & Format(Date, "dd/mm/yyyy") & ")": Exit Function
  EmpaquetadoDisponible = True
  Exit Function
fallo:
  motivo = "sin conexión con Google Sheets (" & Err.Description & ")"
End Function

Private Function EsFechaHoy(ByVal s As String) As Boolean
  Dim p, a As Long, b As Long, c As Long
  s = Trim$(s): If Len(s) < 6 Then Exit Function
  If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
  s = Replace(s, "-", "/"): p = Split(s, "/")
  If UBound(p) <> 2 Then Exit Function
  a = Val(p(0)): b = Val(p(1)): c = Val(p(2))
  If c < 100 And a > 1000 Then            ' aaaa/mm/dd
    EsFechaHoy = (a = Year(Date) And b = Month(Date) And c = Day(Date)): Exit Function
  End If
  If c < 100 Then c = c + 2000
  If c <> Year(Date) Then Exit Function
  EsFechaHoy = (a = Day(Date) And b = Month(Date)) Or (a = Month(Date) And b = Day(Date))
End Function

Private Function RefrescarTabla(lo As ListObject, ByRef nOk As Long, ByRef nErr As Long) As Boolean
  Dim t As Single: t = Timer
  Application.StatusBar = "Actualizando " & lo.Name & "..."
  On Error Resume Next
  lo.QueryTable.BackgroundQuery = False
  Err.Clear
  lo.QueryTable.Refresh BackgroundQuery:=False
  If Err.Number <> 0 Then
    nErr = nErr + 1
    LogE "Consulta " & lo.Name & " (hoja " & lo.Parent.Name & "): " & Err.Description & " -> revisa la conexión ODBC DEPOTUIO / red", "ERROR"
  Else
    nOk = nOk + 1: RefrescarTabla = True
    LogE "Consulta " & lo.Name & ": " & lo.ListRows.Count & " filas (" & Format(Timer - t, "0.0") & " s)"
  End If
  On Error GoTo 0
  DoEvents
End Function

Private Sub RefrescarTablaModelo(lo As ListObject, ByRef nOk As Long, ByRef nErr As Long)
  Dim t As Single
  t = Timer
  Application.StatusBar = "Actualizando " & lo.Name & " (Google Sheets)..."
  On Error Resume Next
  Err.Clear
  lo.TableObject.Refresh          ' refresca su consulta en el modelo y la tabla (una sola vez)
  If Err.Number <> 0 Then
    nErr = nErr + 1
    LogE "Tabla " & lo.Name & ": " & Err.Description, "ERROR"
  Else
    nOk = nOk + 1
    LogE "Tabla " & lo.Name & ": " & lo.ListRows.Count & " filas (" & Format(Timer - t, "0.0") & " s)"
  End If
  On Error GoTo 0
  DoEvents
End Sub

' refresca el panel si está abierto
Public Sub RefrescarPanel()
  Dim i As Long
  On Error Resume Next
  For i = 0 To VBA.UserForms.Count - 1
    If VBA.UserForms(i).Name = "frmEGR" Then VBA.UserForms(i).Recargar
  Next
End Sub

' Después de cambiar la carga de EMPAQUETADO a "solo tabla" (sin modelo de datos), Excel crea la tabla de nuevo
' solo con B:E. Esta macro le devuelve el nombre EMPAQUETADO y las columnas calculadas F:K.
Sub RestaurarColumnasEmpaquetado()
  Dim ws As Worksheet, lo As ListObject, x As ListObject, nom, frm, i As Long, lc As ListColumn
  Set ws = ThisWorkbook.Worksheets("EMPAQUETADO")
  For Each x In ws.ListObjects
    On Error Resume Next
    If Not x.ListColumns("# ORDEN") Is Nothing Then If Err.Number = 0 Then Set lo = x
    Err.Clear
    On Error GoTo 0
    If Not lo Is Nothing Then Exit For
  Next
  If lo Is Nothing Then MsgBox "No hay en la hoja EMPAQUETADO una tabla con la columna '# ORDEN'.", vbExclamation: Exit Sub
  If lo.Range.Column <> 2 Then MsgBox "La tabla debe empezar en la columna B (celda B1). Está en " & lo.Range.Cells(1, 1).Address(False, False) & ".", vbExclamation: Exit Sub
  On Error Resume Next
  If lo.Name <> "EMPAQUETADO" Then lo.Name = "EMPAQUETADO"
  If Err.Number <> 0 Then LogE "EMPAQUETADO: no se pudo renombrar la tabla " & lo.Name & ": " & Err.Description, "AVISO": Err.Clear
  On Error GoTo 0
  nom = Array("NRO. DE CONTENEDORA", "BULTOS", "PESO CAJA", "VOL. CAJA", "VOL. ITEMS", "PORCENTAJE")
  frm = Array( _
    "=IF($C2="""","""",IF($E2<>"""",$E2,XLOOKUP(TEXT($C2,""@""),Estado[DOC_EXT],Estado[NRO_CONTENEDORA_EMPAQUE],IF(COUNTIF(ITEMS_API[DOC_EXT],TEXT($C2,""@""))>0,""validar"",""sin pedido""),0)))", _
    "=IF($C2="""","""",1)", _
    "=IF($C2="""","""",VLOOKUP($D2,DATA_CAJAS,5,FALSE)+0.1)", _
    "=IF($C2="""","""",VLOOKUP($D2,DATA_CAJAS,6,FALSE))", _
    "=IF($C2="""","""",IF(SUMIF('TABLAS DINAMICAS'!$M:$M,$F2,'TABLAS DINAMICAS'!$N:$N)>0,SUMIF('TABLAS DINAMICAS'!$M:$M,$F2,'TABLAS DINAMICAS'!$N:$N),IF(SUMIF('TABLAS DINAMICAS'!$A:$A,$C2,'TABLAS DINAMICAS'!$E:$E)>0,SUMIF('TABLAS DINAMICAS'!$A:$A,$C2,'TABLAS DINAMICAS'!$E:$E),""Verificar"")))", _
    "=IF($C2="""","""",IF(OR($D2=""F1"",$D2=""F5"",$D2=""F11"",$D2=""F10"",$D2=""SOBRE 1""),1,IF($J2=""Verificar"",""Verificar"",IF($J2/$I2>1,1,IF($J2/$I2<0.3,0.3,$J2/$I2)))))")
  For i = 0 To 5
    Set lc = Nothing
    On Error Resume Next
    Set lc = lo.ListColumns(nom(i))
    On Error GoTo 0
    If lc Is Nothing Then
      Set lc = lo.ListColumns.Add(Position:=5 + i)
      lc.Name = nom(i)
    End If
    If lo.ListRows.Count > 0 Then lc.DataBodyRange.Formula2 = frm(i)
  Next
  If lo.ListRows.Count > 0 Then lo.ListColumns("PORCENTAJE").DataBodyRange.NumberFormat = "0%"
  LogE "EMPAQUETADO: tabla lista (nombre EMPAQUETADO, columnas calculadas F:K restauradas)"
  MsgBox "Tabla EMPAQUETADO lista: nombre y columnas F:K restaurados.", vbInformation
End Sub

Sub TABLAS()
  Dim pt As PivotTable
  On Error Resume Next
  For Each pt In ThisWorkbook.Worksheets("TABLAS DINAMICAS").PivotTables
    pt.PivotCache.Refresh
    If Err.Number <> 0 Then LogE "Tabla dinámica " & pt.Name & ": " & Err.Description, "ERROR": Err.Clear
  Next
End Sub

Sub TABLA_APIS()      ' antes llamaba a "TablaDinámica3", que ya no existe
  Dim pt As PivotTable
  On Error Resume Next
  For Each pt In ThisWorkbook.Worksheets("ITEMS APIS").PivotTables
    pt.PivotCache.Refresh
    If Err.Number <> 0 Then LogE "Tabla dinámica " & pt.Name & ": " & Err.Description, "ERROR": Err.Clear
  Next
End Sub

' =====================================================================================
'  2. REPARAR FÓRMULAS (#REF!) - se ejecuta una vez
' =====================================================================================
Sub RepararFormulasEGR()
  Dim ws As Worksheet, nm As Name, i As Long, links, n As Long, lo As ListObject
  If MsgBox("Se corregirán las fórmulas con #REF! y se prepararán las reglas de destino:" & vbCrLf & vbCrLf & _
            "  - DATOS!A: destino con reglas confirmadas (AF:AH)" & vbCrLf & _
            "  - DATOS!Y: courier (antes INDEX(#REF!) = vacío en todos)" & vbCrLf & _
            "  - EMPAQUETADO!F: COUNTIF(#REF!) -> pedidos del día (ITEMS API)" & vbCrLf & _
            "  - DESPACHOS!P: búsqueda sin columnas completas" & vbCrLf & _
            "  - COD POS I:J (vínculo externo roto), nombres #REF!, vínculos externos" & vbCrLf & _
            "  - 185 reglas de formato duplicadas en DATA CODIGO Y CAJAS" & vbCrLf & vbCrLf & _
            "Antes se guarda un respaldo. ¿Continuar?", vbYesNo + vbQuestion, "Reparar fórmulas") <> vbYes Then Exit Sub
  Respaldo "antes_reparar"
  On Error GoTo fallo
  Congelar
  Set ws = ThisWorkbook.Worksheets(HDAT)
  ws.Range("AF1:AH1").Value = Array("DESTINO_REGLA", "PEDIDO_REGLA", "MOTIVO_REGLA")
  ws.Range("AF1:AH1").Font.Bold = True
  ws.Range("A2:A" & MAXF).Formula = "=IF(AND($AF2<>"""",$AG2=$B2),$AF2,IF($O2="""","""",IF($O2=""VALIDAR"","""",IF($O2=""PICHINCHA"",""UIO"",IF($O2=""GUAYAS"",""GYE"",IF($O2=""GALAPAGOS"",""GPS"",""PRO""))))))"
  LogE "REPARAR: DATOS!A2:A" & MAXF & " -> usa el destino confirmado por reglas (AF) si corresponde al mismo pedido (AG)"
  ws.Range("Y2:Y" & MAXF).Formula2 = "=IF(OR($B2="""",$A2=""""),"""",IF($A2=""PRO"",""TRAMACO"",IFERROR(IF(ISNUMBER(SEARCH(""ITSANET""," & _
      "INDEX('COBERTURAS Y TARIFAS'!$F$2:$F$1804,MATCH(TEXTJOIN(""_"",TRUE,$H2,$I2,$J2),COBERT_KEY,0)))),""ITSANET"",IF($A2=""GPS"",""TRAMACO"",""LAAR COURIER"")),""ITSANET"")))"
  LogE "REPARAR: DATOS!Y (COURIER): PRO = TRAMACO; UIO/GYE = ITSANET o LAAR según COBERTURA (misma regla que PEDIDOS HCE)"
  Set lo = Nothing
  On Error Resume Next
  Set lo = ThisWorkbook.Worksheets("EMPAQUETADO").ListObjects("EMPAQUETADO")
  If Not lo Is Nothing Then
    lo.ListColumns("NRO. DE CONTENEDORA").DataBodyRange.Formula2 = _
      "=IF($C2="""","""",IF($E2<>"""",$E2,XLOOKUP(TEXT($C2,""@""),Estado[DOC_EXT],Estado[NRO_CONTENEDORA_EMPAQUE],IF(COUNTIF(ITEMS_API[DOC_EXT],TEXT($C2,""@""))>0,""validar"",""sin pedido""),0)))"
    If Err.Number <> 0 Then LogE "REPARAR: EMPAQUETADO!F no se pudo corregir: " & Err.Description, "ERROR": Err.Clear Else LogE "REPARAR: EMPAQUETADO!F sin #REF! (validar = pedido del día sin contenedora)"
  End If
  On Error GoTo fallo
  ThisWorkbook.Worksheets("DESPACHOS").Range("P2:P501").Formula = _
    "=IF($B2="""","""",IF(IFERROR(INDEX(DATOS!$A$2:$A$500,MATCH($B2,DATOS!$B$2:$B$500,0)),"""")=""PRO"",""TRAMACO"",""FLEXNET""))"
  LogE "REPARAR: DESPACHOS!P sin CHOOSE sobre columnas completas"
  With ThisWorkbook.Worksheets("COD POS")
    .Range("I2:J" & .Cells(.Rows.Count, 1).End(xlUp).Row).ClearContents
  End With
  LogE "REPARAR: COD POS I:J (XLOOKUP a un libro de red inexistente) vaciadas; no las usa ninguna hoja"
  For i = ThisWorkbook.Names.Count To 1 Step -1
    Set nm = ThisWorkbook.Names(i)
    If InStr(nm.RefersTo, "#REF!") > 0 Or (InStr(1, nm.RefersTo, ".xls", vbTextCompare) > 0 And InStr(nm.RefersTo, "]") > 0) Then
      If Left$(nm.Name, 6) <> "_xlnm." Then LogE "REPARAR: nombre eliminado " & nm.Name & " = " & nm.RefersTo: nm.Delete: n = n + 1
    End If
  Next
  links = ThisWorkbook.LinkSources(xlExcelLinks)
  If Not IsEmpty(links) Then
    For i = LBound(links) To UBound(links)
      On Error Resume Next
      ThisWorkbook.BreakLink links(i), xlLinkTypeExcelLinks
      If Err.Number = 0 Then LogE "REPARAR: vínculo externo roto -> valores: " & links(i) Else LogE "REPARAR: vínculo " & links(i) & ": " & Err.Description, "AVISO"
      Err.Clear
      On Error GoTo fallo
    Next
  End If
  With ThisWorkbook.Worksheets("DATA CODIGO Y CAJAS")
    .Cells.FormatConditions.Delete
    With .Range("B2:B1000").FormatConditions.AddUniqueValues
      .DupeUnique = xlDuplicate
      .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
    End With
  End With
  LogE "REPARAR: DATA CODIGO Y CAJAS: 185 reglas de formato -> 1 (códigos duplicados en B2:B1000)"
  CrearHojaReglas
  Liberar
  LogE "REPARAR: terminado (" & n & " nombres rotos eliminados)"
  MsgBox "Fórmulas reparadas. Detalle en el registro del panel.", vbInformation
  Exit Sub
fallo:
  Liberar
  LogE "REPARAR: detenido por error: " & Err.Description, "ERROR"
  MsgBox "Error al reparar: " & Err.Description & vbCrLf & "Hay respaldo en RESPALDOS_EGR.", vbExclamation
End Sub

' =====================================================================================
'  3. REGLAS DE DESTINO (hoja REGLAS_DESTINO, editable por el supervisor)
'  Se evalúan por PRIORIDAD (menor primero); gana la primera regla que coincide.
'  Campos vacíos o "*" = cualquiera. Listas separadas por ";". Comparación sin tildes
'  ni paréntesis. TEXTO busca en la dirección y en la parroquia; "KM>=15" = kilómetro 15 o más.
' =====================================================================================
Public Sub CrearHojaReglas()
  Dim ws As Worksheet, r As Long
  Set ws = Hoja(HREG)
  If Not ws Is Nothing Then Exit Sub
  Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(HDAT))
  ws.Name = HREG
  ws.Range("A1:I1").Value = Array("ACTIVA", "PRIORIDAD", "PROVINCIA", "CANTON", "PARROQUIA", "TEXTO (DIRECCION O PARROQUIA)", "EXCEPTO PARROQUIA", "DESTINO", "MOTIVO (se muestra al operador)")
  r = 2
  AddRegla ws, r, 10, "GUAYAS", "DAULE", "DAULE", "", "", "PRO", "Daule cabecera se entrega por PRO (otro delivery)"
  AddRegla ws, r, 20, "GUAYAS", "DAULE", "", "", "", "GYE", "Resto del cantón Daule lo cubre GYE"
  AddRegla ws, r, 30, "GUAYAS", "GUAYAQUIL", "", "COOP;COOPERATIVA", "", "PRO", "Guayaquil: cooperativa en la dirección se entrega por PRO"
  AddRegla ws, r, 40, "GUAYAS", "GUAYAQUIL", "", "GUASMO;BAUTISTA AGUIRRE;BASTION POPULAR;SALITRE;FORTIN DE LA FLOR;PALESTINA;SIMON BOLIVAR;MAPASINGUE;KM>=15", "", "PRO", "Guayaquil: sector atendido por otro delivery (PRO)"
  AddRegla ws, r, 90, "GUAYAS", "GUAYAQUIL", "", "", "", "GYE", "Guayaquil urbano: GYE"
  AddRegla ws, r, 95, "GUAYAS", "", "", "", "", "PRO", "Guayas fuera de Guayaquil (Durán, Milagro, Playas, Salitre, Naranjal...): PRO"
  AddRegla ws, r, 100, "AZUAY", "GUALACEO;PAUTE;CHORDELEG;CAMILO PONCE ENRIQUEZ;PUCARA;SIGSIG", "", "", "", "PRO", "Azuay: cantón por PRO"
  AddRegla ws, r, 110, "AZUAY", "CUENCA", "MOLLETURO;CHAUCHA;SAYAUSI;LLACAO;PACCHA;NULTI;QUINGEO;CUMBE;IRQUIS;SANTA ANA", "", "", "PRO", "Azuay: parroquia rural de Cuenca por PRO"
  AddRegla ws, r, 900, "PICHINCHA", "", "", "", "", "UIO", "Regla base: Pichincha = UIO"
  AddRegla ws, r, 910, "GALAPAGOS", "", "", "", "", "GPS", "Regla base: Galápagos = GPS"
  AddRegla ws, r, 999, "*", "", "", "", "", "PRO", "Regla base: resto del país = PRO"
  With ws
    .Range("A1:I1").Font.Bold = True: .Range("A1:I1").Interior.Color = RGB(48, 84, 150): .Range("A1:I1").Font.Color = vbWhite
    .Columns("A:B").ColumnWidth = 10: .Columns("C:E").ColumnWidth = 22: .Columns("F").ColumnWidth = 60
    .Columns("G").ColumnWidth = 18: .Columns("H").ColumnWidth = 10: .Columns("I").ColumnWidth = 60
    .Range("A2:A200").Validation.Delete
    .Range("A2:A200").Validation.Add Type:=xlValidateList, Formula1:="SI" & Application.International(xlListSeparator) & "NO"
    .Range("H2:H200").Validation.Delete
    .Range("H2:H200").Validation.Add Type:=xlValidateList, Formula1:=Replace("PRO,GYE,UIO,GPS", ",", Application.International(xlListSeparator))
    .Range("K1").Value = "CÓMO SE USA"
    .Range("K2").Value = "1. Gana la regla ACTIVA de menor PRIORIDAD que coincide con el pedido."
    .Range("K3").Value = "2. Vacío o * = cualquiera. Varias opciones separadas por ; (punto y coma)."
    .Range("K4").Value = "3. TEXTO busca palabras en la dirección y en la parroquia. KM>=15 = kilómetro 15 en adelante."
    .Range("K5").Value = "4. Para cambiar un destino (ej.: GYE ya recibe Durán) edita o agrega una fila; no hace falta tocar el código."
    .Range("K6").Value = "5. Las reglas solo PROPONEN: el operador confirma los cambios en el panel (paso 3)."
    .Range("K1").Font.Bold = True
  End With
  LogE "REGLAS: hoja REGLAS_DESTINO creada con las reglas iniciales (Guayas, Daule, Azuay, bases)"
End Sub

Private Sub AddRegla(ws As Worksheet, ByRef r As Long, ByVal pr As Long, ByVal P As String, ByVal c As String, ByVal q As String, _
                     ByVal txt As String, ByVal exc As String, ByVal dest As String, ByVal mot As String)
  ws.Cells(r, 1).Value = "SI": ws.Cells(r, 2).Value = pr: ws.Cells(r, 3).Value = P: ws.Cells(r, 4).Value = c
  ws.Cells(r, 5).Value = q: ws.Cells(r, 6).Value = txt: ws.Cells(r, 7).Value = exc: ws.Cells(r, 8).Value = dest: ws.Cells(r, 9).Value = mot
  r = r + 1
End Sub

' Reglas activas ordenadas por prioridad -> matriz (n, 8): prov, cant, parr, texto, excepto, destino, motivo, prioridad
Public Function CargarReglas() As Variant
  Dim ws As Worksheet, lr As Long, v, i As Long, j As Long, n As Long, a() As Variant, t As Variant, k As Long
  CrearHojaReglas
  Set ws = ThisWorkbook.Worksheets(HREG)
  lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
  If lr < 2 Then CargarReglas = Empty: Exit Function
  v = ws.Range("A2:I" & lr).Value
  ReDim a(1 To UBound(v, 1), 1 To 8)
  For i = 1 To UBound(v, 1)
    If UCase$(TXE(v(i, 1))) <> "NO" And Len(TXE(v(i, 8))) > 0 Then
      n = n + 1
      a(n, 1) = ListaNorm(TXE(v(i, 3))): a(n, 2) = ListaNorm(TXE(v(i, 4))): a(n, 3) = ListaNorm(TXE(v(i, 5)))
      a(n, 4) = UCase$(TXE(v(i, 6))): a(n, 5) = ListaNorm(TXE(v(i, 7))): a(n, 6) = UCase$(TXE(v(i, 8)))
      a(n, 7) = TXE(v(i, 9)): a(n, 8) = Val(TXE(v(i, 2)))
    End If
  Next
  If n = 0 Then CargarReglas = Empty: Exit Function
  ' orden por prioridad (inserción)
  For i = 2 To n
    For j = i To 2 Step -1
      If a(j, 8) < a(j - 1, 8) Then
        For k = 1 To 8: t = a(j, k): a(j, k) = a(j - 1, k): a(j - 1, k) = t: Next
      Else
        Exit For
      End If
    Next
  Next
  Dim b() As Variant: ReDim b(1 To n, 1 To 8)
  For i = 1 To n: For k = 1 To 8: b(i, k) = a(i, k): Next: Next
  CargarReglas = b
End Function

' "A;B (X);*" -> "|A|B|" ; vacío o * -> ""
Private Function ListaNorm(ByVal s As String) As String
  Dim p, it, o As String
  s = Trim$(s)
  If Len(s) = 0 Or s = "*" Then Exit Function
  p = Split(s, ";")
  For Each it In p
    If Len(NormE(CStr(it))) > 0 Then o = o & "|" & NormE(CStr(it))
  Next
  If Len(o) > 0 Then ListaNorm = o & "|"
End Function

Private Function EnLista(ByVal lista As String, ByVal valor As String) As Boolean
  If Len(lista) = 0 Then EnLista = True Else EnLista = (InStr(lista, "|" & valor & "|") > 0)
End Function

Private Function TextoCoincide(ByVal lista As String, ByVal texto As String) As String
  Dim p, it, w As String, re As Object, m As Object, km As Long
  If Len(Trim$(lista)) = 0 Then TextoCoincide = "*": Exit Function
  p = Split(lista, ";")
  For Each it In p
    w = Trim$(CStr(it))
    If UCase$(Left$(w, 4)) = "KM>=" Then
      km = Val(Mid$(w, 5))
      Set re = CreateObject("VBScript.RegExp"): re.Global = True
      re.Pattern = "\b(KM|KILOMETRO)\s?(\d{1,3})\b"
      For Each m In re.Execute(texto)
        If Val(m.SubMatches(1)) >= km Then TextoCoincide = "KM " & m.SubMatches(1): Exit Function
      Next
    ElseIf Len(NormE(w)) > 0 Then
      If InStr(" " & texto & " ", " " & NormE(w) & " ") > 0 Then TextoCoincide = NormE(w): Exit Function
    End If
  Next
End Function

' Destino según reglas. motivo devuelve el texto de la regla aplicada.
Public Function DestinoPorReglas(reglas As Variant, ByVal prov As String, ByVal cant As String, ByVal parr As String, _
                                 ByVal direccion As String, ByRef motivo As String) As String
  Dim i As Long, P As String, c As String, q As String, texto As String, hit As String
  motivo = ""
  If Not IsArray(reglas) Then Exit Function
  P = NormE(prov): c = NormE(cant): q = NormE(parr)
  texto = NormE(direccion) & " " & q
  For i = 1 To UBound(reglas, 1)
    If EnLista(reglas(i, 1), P) And EnLista(reglas(i, 2), c) And EnLista(reglas(i, 3), q) Then
      If Len(reglas(i, 5)) = 0 Or Not EnLista(reglas(i, 5), q) Then
        hit = TextoCoincide(reglas(i, 4), texto)
        If Len(hit) > 0 Then
          DestinoPorReglas = reglas(i, 6)
          motivo = reglas(i, 7) & IIf(hit <> "*", " [" & hit & "]", "")
          Exit Function
        End If
      End If
    End If
  Next
End Function

' Paso 3: calcula y deja en gPropuestas los pedidos cuyo destino cambia por reglas
Public Function ProponerDestinos() As Long
  Dim ws As Worksheet, lr As Long, v, i As Long, reglas, nuevo As String, mot As String, act As String, nFuera As Long
  Set ws = ThisWorkbook.Worksheets(HDAT)
  Set gPropuestas = New Collection
  lr = UltimaFilaDatos()
  If lr < 2 Then LogE "DESTINOS: DATOS está vacío", "AVISO": Exit Function
  reglas = CargarReglas()
  If Not IsArray(reglas) Then LogE "DESTINOS: no hay reglas activas en REGLAS_DESTINO", "ERROR": Exit Function
  v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, D_RMOT)).Value
  For i = 2 To lr
    If Len(TXE(v(i, D_PED))) > 0 Then
      If TXE(v(i, D_PROV)) = "" Or UCase$(TXE(v(i, D_PROV))) = "VALIDAR" Then
        nFuera = nFuera + 1
      Else
        nuevo = DestinoPorReglas(reglas, TXE(v(i, D_PROV)), TXE(v(i, D_CANT)), TXE(v(i, D_PARR)), TXE(v(i, D_DIR)), mot)
        act = UCase$(TXE(v(i, D_DEST)))
        If Len(nuevo) > 0 And nuevo <> act Then
          gPropuestas.Add Array(i, TXE(v(i, D_PED)), TXE(v(i, D_NOM)), TXE(v(i, D_PROV)), TXE(v(i, D_CANT)), TXE(v(i, D_PARR)), act, nuevo, mot)
        End If
      End If
    End If
  Next
  LogE "DESTINOS: " & gPropuestas.Count & " pedido(s) cambian de destino por reglas" & IIf(nFuera > 0, "; " & nFuera & " sin cobertura TMS (VALIDAR) no se tocan", "")
  ProponerDestinos = gPropuestas.Count
End Function

' Aplica las propuestas indicadas (índices de gPropuestas; vacío = todas)
Public Sub AplicarDestinos(Optional sel As Collection)
  Dim ws As Worksheet, i As Long, p, n As Long, idx
  If gPropuestas Is Nothing Then Exit Sub
  Set ws = ThisWorkbook.Worksheets(HDAT)
  If sel Is Nothing Then
    Set sel = New Collection
    For i = 1 To gPropuestas.Count: sel.Add i: Next
  End If
  If ws.Range("A2").HasFormula Then
    If InStr(ws.Range("A2").Formula, "$AF2") = 0 Then
      MsgBox "Primero ejecuta 'Reparar fórmulas' (herramientas del panel): DATOS!A todavía no lee las reglas confirmadas.", vbExclamation
      Exit Sub
    End If
  End If
  For Each idx In sel
    p = gPropuestas(idx)
    ws.Cells(p(0), D_RDEST).Value = p(7)
    ws.Cells(p(0), D_RPED).Value = ws.Cells(p(0), D_PED).Value
    ws.Cells(p(0), D_RMOT).Value = p(8) & " | " & Application.UserName & " " & Format(Now, "yyyy-mm-dd hh:nn")
    LogE "DESTINO fila " & p(0) & " pedido " & p(1) & ": " & p(6) & " -> " & p(7) & " (" & p(8) & ")"
    n = n + 1
  Next
  Application.Calculate
  LogE "DESTINOS: " & n & " cambio(s) aplicados. TMS, TRAMACO, DESPACHOS y etiquetas ya usan el nuevo destino."
End Sub

' Confirma un destino (sugerido por reglas o elegido por el operador) para una fila de DATOS
Public Function ConfirmarDestino(ByVal fila As Long, ByVal destino As String, ByVal motivo As String) As Boolean
  Dim ws As Worksheet, ant As String
  Set ws = ThisWorkbook.Worksheets(HDAT)
  If ws.Range("A2").HasFormula Then
    If InStr(ws.Range("A2").Formula, "$AF2") = 0 Then
      MsgBox "Primero ejecuta 'Reparar fórmulas' (Operación del panel): DATOS!A todavía no lee los destinos confirmados.", vbExclamation
      Exit Function
    End If
  End If
  ant = TXE(ws.Cells(fila, D_DEST).Value)
  ws.Cells(fila, D_RDEST).Value = destino
  ws.Cells(fila, D_RPED).Value = ws.Cells(fila, D_PED).Value
  ws.Cells(fila, D_RMOT).Value = motivo & " | " & Application.UserName & " " & Format(Now, "yyyy-mm-dd hh:nn")
  LogE "DESTINO fila " & fila & " pedido " & TXE(ws.Cells(fila, D_PED).Value) & ": " & IIf(Len(ant) > 0, ant, "(vacío)") & " -> " & destino & " (" & motivo & ")"
  ConfirmarDestino = True
End Function

Public Sub QuitarConfirmacion(ByVal fila As Long)
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(HDAT)
  ws.Range(ws.Cells(fila, D_RDEST), ws.Cells(fila, D_RMOT)).ClearContents
  LogE "DESTINO fila " & fila & " pedido " & TXE(ws.Cells(fila, D_PED).Value) & ": se quitó la confirmación (vuelve a la regla base)"
End Sub

' COBERTURAS Y TARIFAS por clave PROVINCIA_CANTON_PARROQUIA -> Array(gestor F, gestor sugerido U, trayecto V, días T, CP X)
Public Function DictCobertura() As Object
  Dim ws As Worksheet, lr As Long, v, i As Long, d As Object, k As String
  Set d = CreateObject("Scripting.Dictionary")
  Set ws = ThisWorkbook.Worksheets("COBERTURAS Y TARIFAS")
  lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
  If lr >= 2 Then
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, 28)).Value
    For i = 2 To lr
      k = UCase$(TXE(v(i, 28)))
      If Len(k) > 0 And Not d.Exists(k) Then d(k) = Array(TXE(v(i, 6)), TXE(v(i, 21)), TXE(v(i, 22)), TXE(v(i, 20)), TXE(v(i, 24)))
    Next
  End If
  Set DictCobertura = d
End Function

Public Function ClaveTMS(ByVal h As String, ByVal i As String, ByVal j As String) As String
  Dim s As String
  If Len(h) > 0 Then s = h
  If Len(i) > 0 Then s = s & IIf(Len(s) > 0, "_", "") & i
  If Len(j) > 0 Then s = s & IIf(Len(s) > 0, "_", "") & j
  ClaveTMS = UCase$(s)
End Function

' Empaque y picking por pedido -> Array(estado, cajas, peso cajas, volumen %, unidades conf, unidades sol)
Public Function DictEmpaque() As Object
  Dim d As Object, sol As Object, conf As Object, cont1 As Object, cajas As Object, pesoC As Object, dPct As Object
  Dim lo As ListObject, a, i As Long, doc As String, k
  Set d = CreateObject("Scripting.Dictionary"): Set sol = CreateObject("Scripting.Dictionary")
  Set conf = CreateObject("Scripting.Dictionary"): Set cont1 = CreateObject("Scripting.Dictionary")
  Set cajas = CreateObject("Scripting.Dictionary"): Set pesoC = CreateObject("Scripting.Dictionary"): Set dPct = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  Set lo = ThisWorkbook.Worksheets("ITEMS APIS").ListObjects("ITEMS_API")
  If Not lo Is Nothing Then
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      For i = 1 To UBound(a, 1): doc = TXE(a(i, 2)): sol(doc) = sol(doc) + Val(TXE(a(i, 7))): Next
    End If
  End If
  Set lo = Nothing: Set lo = ThisWorkbook.Worksheets("ITEMS DEPOT").ListObjects("Estado")
  If Not lo Is Nothing Then
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      For i = 1 To UBound(a, 1)
        doc = TXE(a(i, 3))
        If Not cont1.Exists(doc) Then cont1(doc) = TXE(a(i, 7))
        If TXE(a(i, 7)) = cont1(doc) Then conf(doc) = conf(doc) + Val(TXE(a(i, 6)))
      Next
    End If
  End If
  Set lo = Nothing: Set lo = ThisWorkbook.Worksheets("EMPAQUETADO").ListObjects("EMPAQUETADO")
  If Not lo Is Nothing Then
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      For i = 1 To UBound(a, 1)
        doc = TXE(a(i, 2))
        If Len(doc) > 0 Then
          cajas(doc) = cajas(doc) + 1
          If Not IsError(a(i, 7)) Then pesoC(doc) = pesoC(doc) + Val(TXE(a(i, 7)))
          If Not IsError(a(i, 10)) Then If IsNumeric(a(i, 10)) Then dPct(doc) = a(i, 10)
        End If
      Next
    End If
  End If
  On Error GoTo 0
  For Each k In sol.Keys
    d(k) = Array("", 0, 0, "", conf(k), sol(k))
  Next
  For Each k In conf.Keys
    If Not d.Exists(k) Then d(k) = Array("", 0, 0, "", conf(k), sol(k))
  Next
  For Each k In cajas.Keys
    d(k) = Array("EMPACADO", cajas(k), pesoC(k), IIf(dPct.Exists(k), Format(dPct(k), "0%"), ""), conf(k), sol(k))
  Next
  Set DictEmpaque = d
End Function

' Estado de empaque legible para un pedido
Public Function EstadoEmpaque(dE As Object, ByVal ped As String) As String
  Dim x
  If Not dE.Exists(ped) Then EstadoEmpaque = "SIN PICKING": Exit Function
  x = dE(ped)
  If x(0) = "EMPACADO" Then
    EstadoEmpaque = "EMPACADO " & x(1) & " caja" & IIf(x(1) > 1, "s", "")
  ElseIf Val(x(4)) = 0 Then
    EstadoEmpaque = "SIN PICKING"
  ElseIf Val(x(4)) < Val(x(5)) Then
    EstadoEmpaque = "PICKING " & x(4) & "/" & x(5)
  Else
    EstadoEmpaque = "PICKEADO"
  End If
End Function

Public Sub QuitarDestinosConfirmados()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(HDAT)
  ws.Range(ws.Cells(2, D_RDEST), ws.Cells(MAXF, D_RMOT)).ClearContents
  Application.Calculate
  LogE "DESTINOS: se quitaron todos los destinos confirmados (vuelve la regla base por provincia)"
End Sub

' =====================================================================================
'  4. REVISIÓN DE COBERTURA TMS (segunda comprobación, no revalida lo de PEDIDOS HCE)
' =====================================================================================
Public Function RevisarCoberturaTMS(lista As Collection) As Long
  Dim ws As Worksheet, lr As Long, v, i As Long, n As Long, prob As String
  Set ws = ThisWorkbook.Worksheets(HDAT)
  Application.Calculate
  lr = UltimaFilaDatos()
  If lr < 2 Then Exit Function
  v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, D_SUG)).Value
  For i = 2 To lr
    If Len(TXE(v(i, D_PED))) > 0 Then
      prob = ""
      If UCase$(TXE(v(i, D_VAL))) = "REVISAR" Then prob = "Fuera de cobertura TMS: " & TXE(v(i, D_DIAG)) & IIf(Len(TXE(v(i, D_SUG))) > 0, " -> sugerida: " & TXE(v(i, D_SUG)), "")
      If Len(prob) = 0 And Len(TXE(v(i, D_CPTMS))) = 0 Then prob = "Sin código postal TMS (PARR TMS vacío)"
      If Len(prob) = 0 And Len(TXE(v(i, D_TEL))) = 0 Then prob = "Sin teléfono móvil"
      If Len(prob) > 0 Then
        n = n + 1
        lista.Add Array(i, TXE(v(i, D_PED)), TXE(v(i, D_NOM)), TXE(v(i, D_H)) & " / " & TXE(v(i, D_I)) & " / " & TXE(v(i, D_J)), prob)
        LogE "COBERTURA TMS fila " & i & " pedido " & TXE(v(i, D_PED)) & ": " & prob, "REVISAR"
      End If
    End If
  Next
  LogE "COBERTURA TMS: " & (lr - 1) & " pedidos revisados, " & n & " con observación"
  RevisarCoberturaTMS = n
End Function

' =====================================================================================
'  5. AVANCE DE PICKING Y EMPAQUE + COSTOS
' =====================================================================================
Public Function AvancePedidos(lista As Collection, ByRef resumen As String) As Long
  Dim ws As Worksheet, lr As Long, v, i As Long, sol As Object, conf As Object, cont1 As Object, emp As Object
  Dim lo As ListObject, a, k As String, doc As String, ped As String, est As String
  Dim nSin As Long, nPar As Long, nPick As Long, nEmp As Long, nTot As Long, skuSinPeso As Object, skuSinPrecio As Object, cajaSin As Object
  Set sol = CreateObject("Scripting.Dictionary"): Set conf = CreateObject("Scripting.Dictionary")
  Set cont1 = CreateObject("Scripting.Dictionary"): Set emp = CreateObject("Scripting.Dictionary")
  Set skuSinPeso = CreateObject("Scripting.Dictionary"): Set skuSinPrecio = CreateObject("Scripting.Dictionary")
  Set cajaSin = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  ' solicitado (ITEMS API)
  Set lo = ThisWorkbook.Worksheets("ITEMS APIS").ListObjects("ITEMS_API")
  If Not lo Is Nothing Then
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      For i = 1 To UBound(a, 1)
        doc = TXE(a(i, 2)): sol(doc) = sol(doc) + Val(TXE(a(i, 7)))
      Next
    End If
  End If
  ' confirmado (Estado): la consulta repite cada línea por contenedora -> se cuenta una sola contenedora por pedido
  Set lo = Nothing: Set lo = ThisWorkbook.Worksheets("ITEMS DEPOT").ListObjects("Estado")
  If Not lo Is Nothing Then
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      For i = 1 To UBound(a, 1)
        doc = TXE(a(i, 3))
        If Not cont1.Exists(doc) Then cont1(doc) = TXE(a(i, 7))
        If TXE(a(i, 7)) = cont1(doc) Then conf(doc) = conf(doc) + Val(TXE(a(i, 6)))
        If TXE(a(i, 9)) = "" And Len(TXE(a(i, 4))) > 0 Then skuSinPeso(TXE(a(i, 4))) = 1
        If UCase$(TXE(a(i, 11))) = "VERIFICAR" And Len(TXE(a(i, 4))) > 0 Then skuSinPrecio(TXE(a(i, 4))) = 1
      Next
    End If
  End If
  ' empacado (EMPAQUETADO)
  Set lo = Nothing: Set lo = ThisWorkbook.Worksheets("EMPAQUETADO").ListObjects("EMPAQUETADO")
  If Not lo Is Nothing Then
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      For i = 1 To UBound(a, 1)
        doc = TXE(a(i, 2)): If Len(doc) > 0 Then emp(doc) = emp(doc) + 1
        If IsError(a(i, 7)) Or IsError(a(i, 8)) Then cajaSin(TXE(a(i, 3))) = 1
      Next
    End If
  End If
  On Error GoTo 0
  Set ws = ThisWorkbook.Worksheets(HDAT)
  lr = UltimaFilaDatos()
  If lr >= 2 Then
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, D_NOM)).Value
    For i = 2 To lr
      ped = TXE(v(i, D_PED))
      If Len(ped) > 0 Then
        nTot = nTot + 1
        If emp.Exists(ped) Then
          est = "EMPACADO (" & emp(ped) & " caja/s)": nEmp = nEmp + 1
        ElseIf Not conf.Exists(ped) Then
          est = "SIN PICKING": nSin = nSin + 1
        ElseIf conf(ped) < sol(ped) Then
          est = "PICKING PARCIAL " & conf(ped) & "/" & sol(ped): nPar = nPar + 1
        Else
          est = "PICKEADO, FALTA EMPACAR": nPick = nPick + 1
        End If
        lista.Add Array(i, ped, TXE(v(i, D_NOM)), TXE(v(i, D_DEST)), est)
      End If
    Next
  End If
  resumen = "Pedidos " & nTot & " | Empacados " & nEmp & " (" & Pct(nEmp, nTot) & ") | Pickeados sin empacar " & nPick & _
            " | Picking parcial " & nPar & " | Sin picking " & nSin
  LogE "AVANCE: " & resumen
  If skuSinPeso.Count > 0 Then LogE "COSTOS: " & skuSinPeso.Count & " SKU sin peso/volumen en DATA CODIGO Y CAJAS: " & Left$(Join(skuSinPeso.Keys, ", "), 300), "AVISO"
  If skuSinPrecio.Count > 0 Then LogE "COSTOS: " & skuSinPrecio.Count & " SKU sin precio (Verificar): " & Left$(Join(skuSinPrecio.Keys, ", "), 300), "AVISO"
  If cajaSin.Count > 0 Then LogE "COSTOS: tipo de caja sin medidas en DATA_CAJAS: " & Join(cajaSin.Keys, ", "), "AVISO"
  AvancePedidos = nTot
End Function

Private Function Pct(ByVal a As Long, ByVal b As Long) As String
  If b = 0 Then Pct = "0 %" Else Pct = Format(a / b, "0%")
End Function

' =====================================================================================
'  6. EXPORTAR REPORTES (TMS, TRAMACO, DESPACHOS) - no modifica las hojas
' =====================================================================================
Public Function CarpetaExportes() As String
  Dim c As String
  c = Cfg("CARPETA_EXPORTES")
  If Len(c) = 0 Then c = ThisWorkbook.Path & Application.PathSeparator & "EXPORTES"
  If Len(Dir(c, vbDirectory)) = 0 Then MkDir c
  c = c & Application.PathSeparator & Format(Date, "yyyy-mm-dd")
  If Len(Dir(c, vbDirectory)) = 0 Then MkDir c
  CarpetaExportes = c
End Function

' columnas clave (la fila se exporta solo si TODAS tienen dato) y columnas obligatorias
Private Sub DefHoja(ByVal hojaN As String, ByRef claves As Variant, ByRef oblig As Variant)
  Select Case UCase$(hojaN)
    Case "TMS":      claves = Array(10, 17): oblig = Array(10, 11, 14, 15, 16, 17, 23)   ' J destinatario + Q referencia
    Case "TRAMACO":  claves = Array(3, 30): oblig = Array(3, 6, 7, 8, 9, 14, 18, 23)     ' C nombre + AD pedido (solo PRO)
    Case Else:       claves = Array(2, 3): oblig = Array(2, 3, 7, 8, 12, 13)             ' DESPACHOS B pedido + C nombre
  End Select
End Sub

' Filas de la hoja que realmente tienen datos (números de fila)
Public Function FilasExport(ByVal hojaN As String) As Collection
  Dim ws As Worksheet, claves, oblig, lr As Long, v, i As Long, c, ok As Boolean, col As New Collection, cmax As Long
  Set ws = ThisWorkbook.Worksheets(hojaN)
  DefHoja hojaN, claves, oblig
  For Each c In claves
    If ws.Cells(ws.Rows.Count, c).End(xlUp).Row > lr Then lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    If c > cmax Then cmax = c
  Next
  If lr >= 2 Then
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, cmax)).Value
    For i = 2 To lr
      ok = True
      For Each c In claves
        If IsError(v(i, c)) Then
          ok = False
        ElseIf Len(TXE(v(i, c))) = 0 Then
          ok = False
        End If
      Next
      If ok Then col.Add i
    Next
  End If
  Set FilasExport = col
End Function

' Revisa errores antes de exportar. Devuelve número de problemas (-1 si no hay filas)
Public Function ValidarHojaExport(ByVal hojaN As String, filas As Collection) As Long
  Dim ws As Worksheet, claves, oblig, lr As Long, lc As Long, v, r, j As Long, n As Long, c
  Set ws = ThisWorkbook.Worksheets(hojaN)
  DefHoja hojaN, claves, oblig
  If filas.Count = 0 Then LogE "EXPORTAR " & hojaN & ": no tiene filas con datos", "AVISO": ValidarHojaExport = -1: Exit Function
  lr = filas(filas.Count)
  lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
  v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
  For Each r In filas
    For j = 1 To lc
      If IsError(v(r, j)) Then
        n = n + 1: If n <= 40 Then LogE hojaN & " fila " & r & " col " & ws.Cells(1, j).Address(False, False) & " (" & TXE(v(1, j)) & "): valor con error", "ERROR"
      ElseIf UCase$(TXE(v(r, j))) = "VALIDAR" Or UCase$(TXE(v(r, j))) = "VERIFICAR" Then
        n = n + 1: If n <= 40 Then LogE hojaN & " fila " & r & " (" & TXE(v(1, j)) & "): dice " & TXE(v(r, j)), "ERROR"
      End If
    Next
    For Each c In oblig
      If c <= lc Then
        If Not IsError(v(r, c)) Then
          If Len(TXE(v(r, c))) = 0 Then n = n + 1: If n <= 40 Then LogE hojaN & " fila " & r & ": falta " & TXE(v(1, c)), "ERROR"
        End If
      End If
    Next
  Next
  LogE "EXPORTAR " & hojaN & ": " & filas.Count & " filas con datos revisadas, " & n & " problema(s)" & IIf(n > 40, " (se muestran 40)", ""), IIf(n > 0, "AVISO", "INFO")
  ValidarHojaExport = n
End Function

' formato: "CSV", "XLSX" o "PDF". Devuelve la ruta del archivo creado ("" si no se creó)
Public Function ExportarHoja(ByVal hojaN As String, ByVal formato As String, Optional ByVal preguntar As Boolean = True) As String
  Dim ws As Worksheet, lc As Long, j As Long, nProb As Long, filas As Collection, r, k As Long
  Dim wbN As Workbook, wsN As Worksheet, ruta As String, rutaBase As String, src, vv(), i As Long, esTxt As Boolean
  On Error GoTo fallo
  Set ws = ThisWorkbook.Worksheets(hojaN)
  Application.Calculate
  Set filas = FilasExport(hojaN)
  nProb = ValidarHojaExport(hojaN, filas)
  If nProb = -1 Then
    If preguntar Then MsgBox hojaN & " no tiene filas con datos.", vbExclamation
    Exit Function
  End If
  If nProb > 0 And preguntar Then
    If MsgBox(hojaN & ": " & filas.Count & " filas, con " & nProb & " problema(s) (detalle en el registro)." & vbCrLf & "¿Exportar de todas formas?", vbYesNo + vbExclamation, "Exportar " & hojaN) <> vbYes Then Exit Function
  End If
  lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
  src = ws.Range(ws.Cells(1, 1), ws.Cells(filas(filas.Count), lc)).Value
  ' encabezado + SOLO las filas con datos, mismas columnas y mismo orden
  ReDim vv(1 To filas.Count + 1, 1 To lc)
  For j = 1 To lc: vv(1, j) = src(1, j): Next
  k = 1
  For Each r In filas
    k = k + 1
    For j = 1 To lc
      If IsError(src(r, j)) Then vv(k, j) = "" Else vv(k, j) = src(r, j)
    Next
  Next
  Application.ScreenUpdating = False: Application.DisplayAlerts = False
  Set wbN = Workbooks.Add(xlWBATWorksheet)
  Set wsN = wbN.Worksheets(1)
  wsN.Name = Left$(hojaN, 31)
  For j = 1 To lc
    esTxt = False
    For i = 2 To k
      If VarType(vv(i, j)) = vbString Then If Len(vv(i, j)) > 0 Then esTxt = True: Exit For
    Next
    ' columnas de texto quedan como texto (conserva el 0 de teléfonos y códigos); el resto con su formato
    If esTxt Then
      wsN.Range(wsN.Cells(1, j), wsN.Cells(k, j)).NumberFormat = "@"
    Else
      If k >= 2 Then wsN.Range(wsN.Cells(2, j), wsN.Cells(k, j)).NumberFormat = ws.Cells(filas(1), j).NumberFormat
      wsN.Cells(1, j).NumberFormat = "@"
    End If
    wsN.Columns(j).ColumnWidth = ws.Columns(j).ColumnWidth
  Next
  wsN.Range(wsN.Cells(1, 1), wsN.Cells(k, lc)).Value = vv
  wsN.Rows(1).Font.Bold = True
  rutaBase = CarpetaExportes() & Application.PathSeparator & Replace(hojaN, " ", "_") & "_" & Format(Now, "yyyymmdd_hhnn")
  Select Case UCase$(formato)
    Case "CSV"
      ruta = rutaBase & ".csv"
      wbN.SaveAs Filename:=ruta, FileFormat:=62, Local:=(UCase$(Cfg("CSV_SEPARADOR_PUNTOYCOMA", "NO")) = "SI")   ' 62 = CSV UTF-8
    Case "XLSX"
      ruta = rutaBase & ".xlsx"
      wbN.SaveAs Filename:=ruta, FileFormat:=51
    Case "PDF"
      ruta = rutaBase & ".pdf"
      With wsN.PageSetup
        .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintTitleRows = "$1:$1": .CenterFooter = "Página &P de &N": .LeftHeader = hojaN & " - " & Format(Now, "dd/mm/yyyy hh:nn")
      End With
      wsN.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, Quality:=xlQualityStandard, OpenAfterPublish:=False
  End Select
  wbN.Close SaveChanges:=False
  Application.DisplayAlerts = True: Application.ScreenUpdating = True
  LogE "EXPORTAR: " & hojaN & " -> " & ruta & " (" & filas.Count & " filas)"
  ExportarHoja = ruta
  Exit Function
fallo:
  On Error Resume Next
  If Not wbN Is Nothing Then wbN.Close SaveChanges:=False
  Application.DisplayAlerts = True: Application.ScreenUpdating = True
  LogE "EXPORTAR " & hojaN & " (" & formato & "): " & Err.Description, "ERROR"
  If preguntar Then MsgBox "No se pudo exportar " & hojaN & ": " & Err.Description, vbExclamation
End Function

' Exporta las tres hojas en el formato elegido; opcionalmente arma el correo en Outlook
Public Sub ExportarReportes(ByVal formato As String, ByVal hojas As Variant, ByVal correo As Boolean)
  Dim h, rutas As New Collection, f As String, tot As Long
  For Each h In hojas
    f = ExportarHoja(CStr(h), formato, True)
    If Len(f) > 0 Then rutas.Add f
  Next
  If rutas.Count = 0 Then Exit Sub
  If correo Then
    CrearCorreo rutas
  Else
    On Error Resume Next
    Shell "explorer.exe """ & CarpetaExportes() & """", vbNormalFocus
  End If
End Sub

Private Sub CrearCorreo(rutas As Collection)
  Dim ol As Object, m As Object, r
  On Error GoTo sinOutlook
  Set ol = CreateObject("Outlook.Application")
  Set m = ol.CreateItem(0)
  m.To = Cfg("CORREO_PARA")
  m.Subject = "Despacho HYCITE " & Format(Date, "dd/mm/yyyy")
  m.Body = "Adjunto los reportes del despacho del " & Format(Date, "dd/mm/yyyy") & "." & vbCrLf & vbCrLf & "Saludos."
  For Each r In rutas: m.Attachments.Add CStr(r): Next
  m.Display
  LogE "CORREO: borrador creado en Outlook con " & rutas.Count & " adjunto(s)"
  Exit Sub
sinOutlook:
  LogE "CORREO: Outlook no disponible (" & Err.Description & "). Se abre la carpeta de exportes.", "AVISO"
  On Error Resume Next
  Shell "explorer.exe """ & CarpetaExportes() & """", vbNormalFocus
End Sub

' =====================================================================================
'  Menú y panel
' =====================================================================================
Sub AbrirPanelEGR()
  Dim i As Long
  For i = 0 To VBA.UserForms.Count - 1
    If VBA.UserForms(i).Name = "frmEGR" Then VBA.UserForms(i).Show vbModeless: Exit Sub
  Next
  frmEGR.Show vbModeless
End Sub

Sub CrearMenuEGR()
  Dim bar As CommandBar, b As CommandBarButton
  On Error Resume Next
  Application.CommandBars("Despacho EGR HYCITE").Delete
  On Error GoTo 0
  Set bar = Application.CommandBars.Add(Name:="Despacho EGR HYCITE", Position:=msoBarTop, Temporary:=True)
  Set b = bar.Controls.Add(Type:=msoControlButton)
  b.Caption = "Panel EGR (despacho)": b.OnAction = "'" & ThisWorkbook.Name & "'!AbrirPanelEGR"
  b.Style = msoButtonIconAndCaption: b.FaceId = 1087
  Set b = bar.Controls.Add(Type:=msoControlButton)
  b.Caption = "Desbloquear": b.OnAction = "'" & ThisWorkbook.Name & "'!Desbloquear"
  b.Style = msoButtonIconAndCaption: b.FaceId = 346: b.BeginGroup = True
  bar.Visible = True
End Sub

Sub BorrarMenuEGR()
  On Error Resume Next
  Application.CommandBars("Despacho EGR HYCITE").Delete
End Sub
