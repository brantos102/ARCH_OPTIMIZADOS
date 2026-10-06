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
Public Const HTRZ As String = "TRAZABILIDAD"
Public Const MAXF As Long = 500           ' última fila de fórmulas de DATOS

' columnas de DATOS
Public Const D_DEST As Long = 1, D_PED As Long = 2, D_DIR As Long = 4, D_NOM As Long = 7
Public Const D_H As Long = 8, D_I As Long = 9, D_J As Long = 10, D_TEL As Long = 11
Public Const D_PROV As Long = 15, D_CANT As Long = 16, D_PARR As Long = 17
Public Const D_BUL As Long = 18, D_PESO As Long = 19, D_VAL As Long = 22, D_COUR As Long = 25
Public Const D_PARRSP As Long = 26, D_CPTMS As Long = 28, D_DIAG As Long = 30, D_SUG As Long = 31
Public Const D_RDEST As Long = 32, D_RPED As Long = 33, D_RMOT As Long = 34   ' AF, AG, AH (reglas confirmadas)
Public Const D_ZONA As Long = 39                                              ' AM zona peligrosa (la escribe PEDIDOS HCE)

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
    ws.Range("A8:C8").Value = Array("ETIQ_OSCURIDAD", "12", "Oscuridad de la Zebra, 0 a 30. Súbela si la etiqueta sale clara")
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
    "=IF($C2="""","""",COUNTIF(EMPAQUETADO!$C:$C,$C2))", _
    "=IF($C2="""","""",VLOOKUP($D2,DATA_CAJAS,5,FALSE)+0.1)", _
    "=IF($C2="""","""",VLOOKUP($D2,DATA_CAJAS,6,FALSE))", _
    "=IF($C2="""","""",IF($F2="""",""Verificar"",IF(SUMIF('TABLAS DINAMICAS'!$M:$M,$F2,'TABLAS DINAMICAS'!$N:$N)>0,SUMIF('TABLAS DINAMICAS'!$M:$M,$F2,'TABLAS DINAMICAS'!$N:$N),""Verificar"")))", _
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
            "  - DATOS!R: bultos = contenedoras del pedido (1 contenedora = 1 caja)" & vbCrLf & _
            "  - EMPAQUETADO!G: bultos totales del pedido;  EMPAQUETADO!J: volumen solo de esa caja" & vbCrLf & _
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
  ' BULTOS (DATOS!R): la regla es una contenedora = una caja, sin importar cuántos items lleve.
  ' Se cuentan directamente las contenedoras del pedido en la hoja EMPAQUETADO (una fila por
  ' contenedora). No se usa la tabla dinámica "Suma de BULTOS": como ahora EMPAQUETADO!G trae el
  ' total del pedido en cada una de sus filas, esa suma daría el total al cuadrado.
  ws.Range("R2:R" & MAXF).Formula2 = _
    "=IF($B2="""","""",IFERROR(MAX(1,COUNTIF(EMPAQUETADO!$C:$C,TEXT($B2,""0""))),1))"
  LogE "REPARAR: DATOS!R (BULTOS) = contenedoras del pedido en la hoja EMPAQUETADO (1 contenedora = 1 caja). Si el pedido todavía no está empacado, queda en 1"
  Set lo = Nothing
  On Error Resume Next
  Set lo = ThisWorkbook.Worksheets("EMPAQUETADO").ListObjects("EMPAQUETADO")
  If Not lo Is Nothing Then
    lo.ListColumns("NRO. DE CONTENEDORA").DataBodyRange.Formula2 = _
      "=IF($C2="""","""",IF($E2<>"""",$E2,XLOOKUP(TEXT($C2,""@""),Estado[DOC_EXT],Estado[NRO_CONTENEDORA_EMPAQUE],IF(COUNTIF(ITEMS_API[DOC_EXT],TEXT($C2,""@""))>0,""validar"",""sin pedido""),0)))"
    If Err.Number <> 0 Then LogE "REPARAR: EMPAQUETADO!F no se pudo corregir: " & Err.Description, "ERROR": Err.Clear Else LogE "REPARAR: EMPAQUETADO!F sin #REF! (validar = pedido del día sin contenedora)"
    ' G (BULTOS): el total de cajas del PEDIDO, no 1 por fila. Así el dato queda en la hoja
    ' sin tener que agregar una columna a ITEMS DEPOT.
    lo.ListColumns("BULTOS").DataBodyRange.Formula2 = "=IF($C2="""","""",COUNTIF(EMPAQUETADO!$C:$C,$C2))"
    If Err.Number <> 0 Then LogE "REPARAR: EMPAQUETADO!G no se pudo corregir: " & Err.Description, "ERROR": Err.Clear Else LogE "REPARAR: EMPAQUETADO!G (BULTOS) = total de contenedoras del pedido (antes 1 por fila)"
    ' J (VOL. ITEMS): SOLO el volumen de los items de ESA contenedora. Se quita el respaldo
    ' por pedido, que le cargaba a una sola caja el volumen de todo el pedido y daba 100% falsos.
    lo.ListColumns("VOL. ITEMS").DataBodyRange.Formula2 = "=IF($C2="""","""",IF($F2="""",""Verificar"",IF(SUMIF('TABLAS DINAMICAS'!$M:$M,$F2,'TABLAS DINAMICAS'!$N:$N)>0,SUMIF('TABLAS DINAMICAS'!$M:$M,$F2,'TABLAS DINAMICAS'!$N:$N),""Verificar"")))"
    If Err.Number <> 0 Then LogE "REPARAR: EMPAQUETADO!J no se pudo corregir: " & Err.Description, "ERROR": Err.Clear Else LogE "REPARAR: EMPAQUETADO!J (VOL. ITEMS) solo suma los items de esa contenedora; si no los tiene dice Verificar en vez de usar el volumen de todo el pedido"
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
  Dim dC As Object, g As String, k As String
  Set dC = DictCobertura()
  For i = 2 To lr
    If Len(TXE(v(i, D_PED))) > 0 Then
      If TXE(v(i, D_PROV)) = "" Or UCase$(TXE(v(i, D_PROV))) = "VALIDAR" Then
        nFuera = nFuera + 1
      Else
        nuevo = DestinoPorReglas(reglas, TXE(v(i, D_PROV)), TXE(v(i, D_CANT)), TXE(v(i, D_PARR)), TXE(v(i, D_DIR)), mot)
        act = UCase$(TXE(v(i, D_DEST)))
        If (nuevo = "" Or nuevo = act) And (act = "UIO" Or act = "GYE") Then       ' lógica de PEDIDOS HCE
          k = ClaveTMS(TXE(v(i, D_H)), TXE(v(i, D_I)), TXE(v(i, D_J))): g = ""
          If dC.Exists(k) Then g = CStr(dC(k)(0))
          If Len(g) > 0 And InStr(1, g, "ITSANET", vbTextCompare) = 0 And InStr(1, g, "LAAR", vbTextCompare) = 0 Then
            nuevo = "PRO": mot = "Cobertura TMS: " & g & " (sin ITSANET/LAAR para " & act & ")"
          End If
        End If
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

'  Alta y edición de productos o cajas en DATA CODIGO Y CAJAS con el formulario de datos de Excel
'  (Nuevo, Buscar criterios, Anterior/Siguiente, Eliminar). DATA_CODIGOS y DATA_CAJAS son columnas completas,
'  así que las filas nuevas se usan de inmediato en pesos, volúmenes y costos.
Public Sub EditarMaestro(ByVal tipo As String)
  Dim ws As Worksheet, rng As Range, lr As Long
  On Error GoTo fallo
  RuedaDesactivar
  Set ws = ThisWorkbook.Worksheets("DATA CODIGO Y CAJAS")
  MostrarExcel
  ws.Visible = xlSheetVisible
  ws.Activate
  If UCase$(tipo) = "CAJAS" Then
    lr = ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, 15), ws.Cells(IIf(lr < 2, 2, lr), 26))       ' O:Z
  Else
    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(IIf(lr < 2, 2, lr), 13))        ' A:M
  End If
  ws.Names.Add Name:="Database", RefersTo:="='" & ws.Name & "'!" & rng.Address
  rng.Cells(1, 1).Select
  LogE "MAESTRO: edición de " & LCase$(tipo) & " (" & (rng.Rows.Count - 1) & " registros)"
  ws.ShowDataForm
  LogE "MAESTRO: edición de " & LCase$(tipo) & " terminada"
  Exit Sub
fallo:
  LogE "MAESTRO: " & Err.Description, "ERROR"
  MsgBox "No se pudo abrir el formulario de datos: " & Err.Description, vbExclamation
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

' Índice de una columna de una tabla buscándola por el nombre del encabezado.
' Así, si algún día la consulta cambia de columnas, el panel no se descuadra.
' el código sigue leyendo el dato correcto. Si no está, devuelve 'predet'.
Public Function ColIdx(lo As ListObject, ByVal nombre As String, Optional ByVal predet As Long = 0) As Long
  Dim lc As ListColumn
  ColIdx = predet
  If lo Is Nothing Then Exit Function
  For Each lc In lo.ListColumns
    If UCase$(Trim$(lc.Name)) = UCase$(Trim$(nombre)) Then ColIdx = lc.Index: Exit Function
  Next
End Function

' Empaque y picking por pedido -> Array(estado, cajas, peso cajas, volumen %, unidades conf, unidades sol)
Public Function DictEmpaque() As Object
  Dim d As Object, sol As Object, conf As Object, cont1 As Object, cajas As Object, pesoC As Object, dPct As Object
  Dim cajasD As Object, nCajD As Object
  Dim lo As ListObject, a, i As Long, doc As String, k
  Dim cDoc As Long, cProd As Long, cConf As Long, cCont As Long
  Set d = CreateObject("Scripting.Dictionary"): Set sol = CreateObject("Scripting.Dictionary")
  Set conf = CreateObject("Scripting.Dictionary"): Set cont1 = CreateObject("Scripting.Dictionary")
  Set cajasD = CreateObject("Scripting.Dictionary"): Set nCajD = CreateObject("Scripting.Dictionary")
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
    ' las columnas se buscan por nombre: la consulta puede traer columnas nuevas
    cDoc = ColIdx(lo, "DOC_EXT", 3): cProd = ColIdx(lo, "PRODUCTO_ID", 4)
    cConf = ColIdx(lo, "Cantidad_Confirmada", 6): cCont = ColIdx(lo, "NRO_CONTENEDORA_EMPAQUE", 7)
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      ' Una fila por pedido + producto + contenedora: un producto repartido en dos
      ' cajas suma las dos, pero la misma combinación nunca se cuenta dos veces
      ' (pesos y costos correctos para facturar).
      ' Las cajas del pedido = contenedoras distintas que aparecen en sus filas,
      ' sin importar cuántos items lleve cada una.
      For i = 1 To UBound(a, 1)
        doc = TXE(a(i, cDoc))
        If Len(doc) > 0 Then
          k = doc & Chr(1) & TXE(a(i, cProd)) & Chr(1) & TXE(a(i, cCont))
          If Not cont1.Exists(k) Then
            cont1(k) = 1
            conf(doc) = conf(doc) + Val(TXE(a(i, cConf)))
          End If
          If Len(TXE(a(i, cCont))) > 0 Then
            If Not cajasD.Exists(doc & Chr(1) & TXE(a(i, cCont))) Then
              cajasD(doc & Chr(1) & TXE(a(i, cCont))) = 1
              nCajD(doc) = nCajD(doc) + 1
            End If
          End If
        End If
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
  ' las cajas reales las da DEPOT (contenedoras de ITEMS DEPOT); el Google Sheets es el respaldo
  For Each k In nCajD.Keys
    If d.Exists(k) Then
      d(k) = Array("EMPACADO", nCajD(k), IIf(cajas.Exists(k), pesoC(k), 0), _
                   IIf(dPct.Exists(k), Format(dPct(k), "0%"), ""), conf(k), sol(k))
    Else
      d(k) = Array("EMPACADO", nCajD(k), 0, "", conf(k), sol(k))
    End If
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
        If n <= 30 Then LogE "COBERTURA TMS fila " & i & " pedido " & TXE(v(i, D_PED)) & ": " & prob, "REVISAR"
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
  Dim nSin As Long, nPar As Long, nPick As Long, nEmp As Long, nTot As Long, nDup As Long
  Dim skuSinPeso As Object, skuSinPrecio As Object, cajaSin As Object
  Dim cDoc As Long, cProd As Long, cConf As Long, cCont As Long, cPeso As Long, cPrec As Long
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
  ' confirmado (Estado): una línea por pedido + producto + contenedora; si una se repite, se avisa
  Set lo = Nothing: Set lo = ThisWorkbook.Worksheets("ITEMS DEPOT").ListObjects("Estado")
  If Not lo Is Nothing Then
    cDoc = ColIdx(lo, "DOC_EXT", 3): cProd = ColIdx(lo, "PRODUCTO_ID", 4)
    cConf = ColIdx(lo, "Cantidad_Confirmada", 6): cCont = ColIdx(lo, "NRO_CONTENEDORA_EMPAQUE", 7)
    cPeso = ColIdx(lo, "PESO", 9): cPrec = ColIdx(lo, "PRECIO", 11)
    If lo.ListRows.Count > 0 Then
      a = lo.DataBodyRange.Value
      For i = 1 To UBound(a, 1)
        doc = TXE(a(i, cDoc))
        If Len(doc) > 0 Then
          k = doc & Chr(1) & TXE(a(i, cProd)) & Chr(1) & TXE(a(i, cCont))
          If cont1.Exists(k) Then
            nDup = nDup + 1
          Else
            cont1(k) = 1
            conf(doc) = conf(doc) + Val(TXE(a(i, cConf)))
          End If
        End If
        If cPeso > 0 Then
          If TXE(a(i, cPeso)) = "" And Len(TXE(a(i, cProd))) > 0 Then skuSinPeso(TXE(a(i, cProd))) = 1
        End If
        If cPrec > 0 Then
          If UCase$(TXE(a(i, cPrec))) = "VERIFICAR" And Len(TXE(a(i, cProd))) > 0 Then skuSinPrecio(TXE(a(i, cProd))) = 1
        End If
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
  If nDup > 0 Then LogE "ITEMS DEPOT: " & nDup & " fila(s) repetidas (mismo pedido, producto y contenedora). " & _
       "Las cantidades se contaron una sola vez. Corrige la consulta 'Estado' con SQL_ITEMS_DEPOT.sql para que los pesos y costos de la hoja salgan bien.", "AVISO"
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
' Cambia, SOLO en el bloque que se va a exportar, la columna auxiliar FILA_DATOS por la zona
' peligrosa del pedido. FILA_DATOS es la posición dentro de DATOS!A2:A500 (la misma que usan las
' fórmulas INDEX de la hoja), así que la fila de DATOS es esa posición + 1.
Private Sub FilaPorZona(ByRef vv As Variant, ByVal k As Long, ByVal lc As Long)
  Dim j As Long, cFila As Long, i As Long, f As Long, lr As Long, z, ws As Worksheet, n As Long
  For j = 1 To lc
    If UCase$(Replace(TXE(vv(1, j)), " ", "")) = "FILA_DATOS" Then cFila = j: Exit For
  Next
  If cFila = 0 Then Exit Sub
  Set ws = ThisWorkbook.Worksheets(HDAT)
  lr = UltimaFilaDatos()
  If lr >= 2 Then z = ws.Range(ws.Cells(2, D_ZONA), ws.Cells(lr, D_ZONA)).Value
  vv(1, cFila) = "ZONA"
  For i = 2 To k
    f = Val(TXE(vv(i, cFila)))
    vv(i, cFila) = ""
    If f >= 1 And IsArray(z) Then
      If f <= UBound(z, 1) Then
        If Not IsError(z(f, 1)) Then vv(i, cFila) = TXE(z(f, 1))
      End If
    End If
    If Len(TXE(vv(i, cFila))) > 0 Then n = n + 1
  Next
  LogE "EXPORTAR TRAMACO: la columna FILA_DATOS sale como ZONA (zona peligrosa); " & n & " de " & (k - 1) & " fila(s) con zona. La hoja TRAMACO no se modifica."
End Sub

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
  ' TRAMACO: la columna auxiliar FILA_DATOS no le sirve al courier. En el archivo que sale
  ' se reemplaza por la ZONA (zona peligrosa que viene de PEDIDOS HCE). La hoja no se toca.
  If UCase$(hojaN) = "TRAMACO" Then FilaPorZona vv, k, lc
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
  Dim bar As CommandBar
  On Error Resume Next
  Application.CommandBars("Despacho EGR HYCITE").Delete
  On Error GoTo 0
  Set bar = Application.CommandBars.Add(Name:="Despacho EGR HYCITE", Position:=msoBarTop, Temporary:=True)
  BotonMenu bar, "Panel EGR", "AbrirPanelEGR", 1087, False
  BotonMenu bar, "¿Qué sigue?", "MenuGuia", 984, False
  BotonMenu bar, "0 Limpiar día", "LimpiarDia", 358, True
  BotonMenu bar, "1 Revisar cobertura TMS", "MenuRevisarCobertura", 1098, True
  BotonMenu bar, "2 Ver cambios sugeridos", "MenuVerCambios", 1099, False
  BotonMenu bar, "2b Aplicar destinos sugeridos", "MenuAplicarDestinos", 1099, False
  BotonMenu bar, "2c Asignar destino a un pedido", "MenuAsignarDestino", 1099, False
  BotonMenu bar, "3 IMPRIMIR ETIQUETAS", "MenuEtiquetas", 4, True
  BotonMenu bar, "4 Avance empaque", "MenuAvance", 1016, True
  BotonMenu bar, "5 Exportar reportes", "MenuExportar", 3, False
  BotonMenu bar, "Trazabilidad", "MenuTrazabilidad", 1016, False
  BotonMenu bar, "Actualizar datos", "ActualizarTodo", 459, True
  BotonMenu bar, "Reglas de destino", "MenuReglas", 1017, False
  BotonMenu bar, "Productos", "MenuProductos", 1087, False
  BotonMenu bar, "Cajas", "MenuCajas", 1087, False
  BotonMenu bar, "Impresora", "ElegirImpresoraMenu", 4, False
  ListaHojasMenu bar
  BotonMenu bar, "Reparar fórmulas", "RepararFormulasEGR", 1100, True
  BotonMenu bar, "Desbloquear", "Desbloquear", 346, False
  bar.Visible = True
End Sub

Private Sub BotonMenu(bar As CommandBar, ByVal cap As String, ByVal macro As String, ByVal cara As Long, ByVal grupo As Boolean)
  Dim b As CommandBarButton
  Set b = bar.Controls.Add(Type:=msoControlButton)
  b.Caption = cap: b.OnAction = "'" & ThisWorkbook.Name & "'!" & macro
  b.Style = msoButtonIconAndCaption: b.BeginGroup = grupo
  On Error Resume Next
  b.FaceId = cara
End Sub

' ---------- operación desde Complementos (sin formularios) ----------
Private Function PanelOcupado() As Boolean
  RuedaDesactivar
  If gOcupadoE Then MsgBox "Hay un proceso en curso.", vbInformation: PanelOcupado = True
End Function

Sub MenuRevisarCobertura()
  Dim c As New Collection, n As Long, i As Long, s As String, ws As Worksheet
  If PanelOcupado() Then Exit Sub
  n = RevisarCoberturaTMS(c)
  If n = 0 Then MsgBox "Todos los pedidos están en la cobertura TMS, con código postal y teléfono.", vbInformation: Exit Sub
  For i = 1 To c.Count
    If i <= 15 Then s = s & vbCrLf & "Fila " & c(i)(0) & "  " & c(i)(1) & ": " & Left$(c(i)(4), 70)
  Next
  MsgBox n & " pedido(s) con observación:" & s & IIf(n > 15, vbCrLf & "...", "") & vbCrLf & vbCrLf & _
         "Los fuera de cobertura no tendrán destino ni etiqueta hasta corregirlos (detalle en LOG_EGR).", vbExclamation, "Cobertura TMS"
  Set ws = ThisWorkbook.Worksheets(HDAT)
  ws.Activate
  ws.Cells(c(1)(0), D_PED).Select
End Sub

Sub MenuAplicarDestinos()
  Dim n As Long, i As Long, s As String, p
  If PanelOcupado() Then Exit Sub
  n = ProponerDestinos()
  If n = 0 Then MsgBox "Ningún pedido cambia de destino según las reglas y la cobertura.", vbInformation: Exit Sub
  For i = 1 To gPropuestas.Count
    p = gPropuestas(i)
    If i <= 20 Then s = s & vbCrLf & p(1) & "  " & IIf(Len(p(6)) > 0, p(6), "(vacío)") & " -> " & p(7) & "  (" & Left$(p(8), 55) & ")"
  Next
  If MsgBox(n & " pedido(s) cambian de destino:" & s & IIf(n > 20, vbCrLf & "...", "") & vbCrLf & vbCrLf & _
            "¿Confirmar TODOS? (para elegir uno por uno usa el Panel EGR)", vbYesNo + vbQuestion, "Destinos sugeridos") <> vbYes Then Exit Sub
  AplicarDestinos
  TrazaAuto
  MsgBox n & " destino(s) confirmados. Si ya tenían etiqueta, reimprímelas y vuelve a exportar." & vbCrLf & _
         "La hoja TRAZABILIDAD se actualizó con el estado final y los cambios.", vbInformation
End Sub

Sub MenuAvance()
  Dim c As New Collection, res As String
  If PanelOcupado() Then Exit Sub
  AvancePedidos c, res
  MsgBox res & vbCrLf & vbCrLf & "Avisos de costos (SKU o cajas sin datos) en el registro LOG_EGR.", vbInformation, "Avance de empaque"
End Sub

Sub MenuExportar()
  Dim h As String, f As String, hojas, fmt As String
  If PanelOcupado() Then Exit Sub
  h = InputBox("¿Qué exportar?" & vbCrLf & "1 = TRAMACO (" & FilasExport("TRAMACO").Count & " filas)" & vbCrLf & _
               "2 = TMS (" & FilasExport("TMS").Count & " filas)" & vbCrLf & "3 = DESPACHOS (" & FilasExport("DESPACHOS").Count & " filas)" & vbCrLf & "4 = las tres", "Exportar reportes", "1")
  Select Case Trim$(h)
    Case "1": hojas = Array("TRAMACO")
    Case "2": hojas = Array("TMS")
    Case "3": hojas = Array("DESPACHOS")
    Case "4": hojas = Array("TMS", "TRAMACO", "DESPACHOS")
    Case Else: Exit Sub
  End Select
  f = InputBox("Formato:" & vbCrLf & "1 = CSV" & vbCrLf & "2 = XLSX" & vbCrLf & "3 = PDF", "Exportar reportes", "1")
  Select Case Trim$(f)
    Case "1": fmt = "CSV"
    Case "2": fmt = "XLSX"
    Case "3": fmt = "PDF"
    Case Else: Exit Sub
  End Select
  ExportarReportes fmt, hojas, (MsgBox("¿Crear también el correo de Outlook con los archivos?", vbYesNo + vbQuestion) = vbYes)
End Sub

' Desplegable con las hojas que se usan en el día
Private Sub ListaHojasMenu(bar As CommandBar)
  Dim c As CommandBarComboBox, h
  On Error Resume Next
  Set c = bar.Controls.Add(Type:=msoControlDropdown)
  If c Is Nothing Then Exit Sub
  c.Caption = "Ir a hoja": c.Width = 150: c.BeginGroup = True
  c.OnAction = "'" & ThisWorkbook.Name & "'!IrAHojaMenu"
  For Each h In Array("DATOS", "TMS", "TRAMACO", "DESPACHOS", "ETIQUETAS", "EMPAQUETADO", _
                      "TABLAS DINAMICAS", "ITEMS APIS", "ITEMS DEPOT", "COBERTURAS Y TARIFAS", _
                      "DATA CODIGO Y CAJAS", "REGLAS_DESTINO", "TRAZABILIDAD", "LOG_EGR", "PANEL")
    c.AddItem CStr(h)
  Next
  c.ListIndex = 1
End Sub

Sub IrAHojaMenu()
  Dim c As CommandBarComboBox, ws As Worksheet
  On Error Resume Next
  Set c = Application.CommandBars("Despacho EGR HYCITE").Controls("Ir a hoja")
  If c Is Nothing Then Exit Sub
  RuedaDesactivar
  Set ws = ThisWorkbook.Worksheets(c.Text)
  On Error GoTo 0
  If ws Is Nothing Then MsgBox "No existe la hoja " & c.Text & ".", vbExclamation: Exit Sub
  ws.Visible = xlSheetVisible
  ws.Activate
  LogE "MENU: ir a la hoja " & ws.Name
End Sub

' Dice en qué punto del día está el despacho y cuál es el siguiente paso
Sub MenuGuia()
  Dim ws As Worksheet, wsE As Worksheet, lr As Long, v, e, i As Long
  Dim nTot As Long, nFue As Long, nCam As Long, nPen As Long, nRei As Long, nEmp As Long
  Dim reglas, dC As Object, dE As Object, sug As String, mot As String, dest As String, k As String, g As String
  Dim txt As String, sig As String
  If PanelOcupado() Then Exit Sub
  Set ws = ThisWorkbook.Worksheets(HDAT): Set wsE = ThisWorkbook.Worksheets("ETIQUETAS")
  Application.Calculate
  lr = UltimaFilaDatos()
  If lr < 2 Then
    MsgBox "La hoja DATOS está vacía." & vbCrLf & vbCrLf & _
           "SIGUIENTE PASO: en PEDIDOS HCE, paso 7 'Enviar a EGR' con este archivo abierto.", vbInformation, "¿Qué sigue?"
    Exit Sub
  End If
  v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, D_RMOT)).Value
  e = wsE.Range(wsE.Cells(1, 1), wsE.Cells(lr, 14)).Value
  reglas = CargarReglas()
  Set dC = DictCobertura(): Set dE = DictEmpaque()
  For i = 2 To lr
    If Len(TXE(v(i, D_PED))) > 0 Then
      nTot = nTot + 1
      dest = UCase$(TXE(v(i, D_DEST)))
      If UCase$(TXE(v(i, D_VAL))) = "REVISAR" Or Len(dest) = 0 Then
        nFue = nFue + 1
      Else
        sug = ""
        If IsArray(reglas) Then sug = DestinoPorReglas(reglas, TXE(v(i, D_PROV)), TXE(v(i, D_CANT)), TXE(v(i, D_PARR)), TXE(v(i, D_DIR)), mot)
        If (sug = "" Or sug = dest) And (dest = "UIO" Or dest = "GYE") Then
          k = ClaveTMS(TXE(v(i, D_H)), TXE(v(i, D_I)), TXE(v(i, D_J))): g = ""
          If dC.Exists(k) Then g = CStr(dC(k)(0))
          If Len(g) > 0 Then
            If InStr(1, g, "ITSANET", vbTextCompare) = 0 And InStr(1, g, "LAAR", vbTextCompare) = 0 Then sug = "PRO"
          End If
        End If
        If Len(sug) > 0 And sug <> dest Then nCam = nCam + 1
        If UCase$(TXE(e(i, 6))) = "OK" Then
          If Len(TXE(e(i, 13))) > 0 And UCase$(TXE(e(i, 13))) <> dest Then nRei = nRei + 1
        Else
          nPen = nPen + 1
        End If
        If Left$(EstadoEmpaque(dE, TXE(v(i, D_PED))), 8) = "EMPACADO" Then nEmp = nEmp + 1
      End If
    End If
  Next
  txt = "ESTADO DE HOY" & vbCrLf & _
        "  Pedidos: " & nTot & vbCrLf & _
        "  Fuera de cobertura TMS: " & nFue & vbCrLf & _
        "  Con cambio de destino sugerido: " & nCam & vbCrLf & _
        "  Etiquetas pendientes: " & nPen & "   ·   por reimprimir: " & nRei & vbCrLf & _
        "  Empacados: " & nEmp & " de " & nTot & vbCrLf & vbCrLf
  If nFue > 0 Then
    sig = "1 Revisar cobertura TMS: hay " & nFue & " pedido(s) sin cobertura. No tendrán destino ni etiqueta hasta corregirlos."
  ElseIf nCam > 0 Then
    sig = "2 Ver cambios sugeridos: " & nCam & " pedido(s) cambian de destino. Revísalos y confirma."
  ElseIf nPen + nRei > 0 Then
    sig = "3 Etiquetas: faltan " & nPen & " por imprimir" & IIf(nRei > 0, " y " & nRei & " por reimprimir", "") & "."
  ElseIf nEmp < nTot Then
    sig = "4 Avance empaque: van " & nEmp & " de " & nTot & ". Pulsa 'Actualizar datos' cuando bodega avance."
  Else
    sig = "5 Exportar reportes: TRAMACO al portal, TMS al sistema y DESPACHOS a distribución."
  End If
  MsgBox txt & "SIGUIENTE PASO:" & vbCrLf & sig, vbInformation, "¿Qué sigue?"
End Sub

Sub MenuVerCambios()
  Dim n As Long, i As Long, s As String, p
  If PanelOcupado() Then Exit Sub
  n = ProponerDestinos()
  If n = 0 Then MsgBox "Ningún pedido cambia de destino según las reglas y la cobertura.", vbInformation, "Cambios sugeridos": Exit Sub
  For i = 1 To gPropuestas.Count
    p = gPropuestas(i)
    If i <= 25 Then s = s & vbCrLf & p(1) & "  " & p(3) & " / " & p(5) & vbCrLf & "     " & _
        IIf(Len(p(6)) > 0, p(6), "(vacío)") & " -> " & p(7) & "   (" & Left$(p(8), 60) & ")"
  Next
  MsgBox n & " pedido(s) con cambio sugerido:" & vbCrLf & s & IIf(n > 25, vbCrLf & "..." , "") & vbCrLf & vbCrLf & _
         "Para confirmarlos todos: '2b Aplicar destinos sugeridos'." & vbCrLf & _
         "Para elegir uno por uno: abre el Panel EGR.", vbInformation, "Cambios sugeridos"
End Sub

Sub MenuAsignarDestino()
  Dim ped As String, d As String, ws As Worksheet, lr As Long, i As Long, fila As Long
  If PanelOcupado() Then Exit Sub
  ped = Trim$(InputBox("Número de pedido:", "Asignar destino a mano"))
  If Len(ped) = 0 Then Exit Sub
  Set ws = ThisWorkbook.Worksheets(HDAT)
  lr = UltimaFilaDatos()
  For i = 2 To lr
    If TXE(ws.Cells(i, D_PED).Value) = ped Then fila = i: Exit For
  Next
  If fila = 0 Then MsgBox "No se encontró el pedido " & ped & " en la hoja DATOS.", vbExclamation: Exit Sub
  d = UCase$(Trim$(InputBox("Pedido " & ped & "  ·  " & TXE(ws.Cells(fila, D_NOM).Value) & vbCrLf & _
      "Destino actual: " & TXE(ws.Cells(fila, D_DEST).Value) & vbCrLf & vbCrLf & _
      "Escribe el destino nuevo: PRO, GYE, UIO o GPS", "Asignar destino a mano")))
  If Len(d) = 0 Then Exit Sub
  If d <> "PRO" And d <> "GYE" And d <> "UIO" And d <> "GPS" Then MsgBox "Destino no válido.", vbExclamation: Exit Sub
  If ConfirmarDestino(fila, d, "MANUAL OPERARIO") Then
    Application.Calculate
    RefrescarPanel
    MsgBox "Pedido " & ped & " asignado a " & d & "." & vbCrLf & _
           "Si ya tenía etiqueta, reimprímela y vuelve a exportar.", vbInformation
  End If
End Sub

' Un solo botón de etiquetas: abre la lista completa, se filtra y se imprime lo filtrado.
Sub MenuEtiquetas()
  If PanelOcupado() Then Exit Sub
  RuedaDesactivar
  frmEtiquetas.Show vbModal
End Sub

Sub MenuReglas()
  If PanelOcupado() Then Exit Sub
  CrearHojaReglas
  frmReglas.Show vbModal
End Sub

Sub MenuProductos()
  If PanelOcupado() Then Exit Sub
  EditarMaestro "PRODUCTOS"
End Sub

Sub MenuCajas()
  If PanelOcupado() Then Exit Sub
  EditarMaestro "CAJAS"
End Sub

' Inicio del día: borra la validación anterior de DATOS (destinos confirmados, datos de PEDIDOS HCE)
' y el estado de las etiquetas. Opcionalmente también los pedidos de ayer (C:N).
Sub LimpiarDia()
  Dim ws As Worksheet, wsE As Worksheet, r As VbMsgBoxResult
  If PanelOcupado() Then Exit Sub
  r = MsgBox("INICIAR EL DÍA: se borrará la validación anterior de DATOS:" & vbCrLf & _
             "  - destinos confirmados (AF:AH)" & vbCrLf & "  - datos de PEDIDOS HCE (AI:AO)" & vbCrLf & _
             "  - estado de etiquetas impresas (ETIQUETAS F, M, N)" & vbCrLf & vbCrLf & _
             "¿Borrar también los PEDIDOS de ayer (DATOS C:N)?" & vbCrLf & _
             "  Sí = borra todo    No = conserva los pedidos    Cancelar = no hace nada", vbYesNoCancel + vbQuestion, "Limpiar día")
  If r = vbCancel Then Exit Sub
  Respaldo "antes_limpiar_dia"
  Set ws = ThisWorkbook.Worksheets(HDAT): Set wsE = ThisWorkbook.Worksheets("ETIQUETAS")
  Application.ScreenUpdating = False
  ws.Range(ws.Cells(2, D_RDEST), ws.Cells(MAXF, 41)).ClearContents
  If r = vbYes Then ws.Range(ws.Cells(2, 3), ws.Cells(MAXF, 14)).ClearContents
  wsE.Range(wsE.Cells(2, 6), wsE.Cells(501, 6)).Value = "-"
  wsE.Range(wsE.Cells(2, 13), wsE.Cells(501, 14)).ClearContents
  Application.Calculate
  Application.ScreenUpdating = True
  LogE "LIMPIAR DÍA: validación anterior borrada" & IIf(r = vbYes, " (también los pedidos C:N)", " (pedidos conservados)")
  RefrescarPanel
  MsgBox "Listo para el nuevo día. Siguiente: en PEDIDOS HCE, paso 7 'Enviar a EGR'.", vbInformation
End Sub

Sub BorrarMenuEGR()
  On Error Resume Next
  Application.CommandBars("Despacho EGR HYCITE").Delete
End Sub

' =====================================================================================
'  TRAZABILIDAD DE PEDIDOS
'  Genera la hoja TRAZABILIDAD con tres bloques:
'    1. estado final de cada pedido (hoja DATOS, tal como quedó),
'    2. los cambios de destino confirmados (quién, cuándo y por qué),
'    3. el registro de acciones del día (LOG_EGR).
'  Es SOLO LECTURA sobre las hojas de trabajo: no cambia ni un dato ni una fórmula.
'  Si algo falla, lo anota en el registro y no interrumpe lo que se estaba haciendo.
' =====================================================================================

' Destino que saldría por la regla base de provincia, sin destino confirmado.
' Es la misma regla de la fórmula de DATOS!A, para poder mostrar el "antes".
Public Function DestinoBase(ByVal prov As String) As String
  Select Case UCase$(Trim$(prov))
    Case "", "VALIDAR": DestinoBase = ""
    Case "PICHINCHA": DestinoBase = "UIO"
    Case "GUAYAS": DestinoBase = "GYE"
    Case "GALAPAGOS": DestinoBase = "GPS"
    Case Else: DestinoBase = "PRO"
  End Select
End Function

' Arma (o rehace) la hoja TRAZABILIDAD. Devuelve cuántos pedidos quedaron listados.
Public Function GenerarTrazabilidad(Optional ByVal avisar As Boolean = True) As Long
  Dim ws As Worksheet, wsD As Worksheet, wsE As Worksheet, wsL As Worksheet
  Dim lr As Long, v, e, i As Long, n As Long, r As Long, j As Long
  Dim dC As Object, dE As Object, dSug As Object, cob, x, p
  Dim a() As Variant, enc, ped As String, dest As String, base As String, conf As Boolean
  Dim nConf As Long, nSug As Long, nFue As Long, k As String
  On Error GoTo fallo
  Set wsD = ThisWorkbook.Worksheets(HDAT)
  Set wsE = ThisWorkbook.Worksheets("ETIQUETAS")
  Application.Calculate
  lr = UltimaFilaDatos()
  If lr < 2 Then
    If avisar Then MsgBox "La hoja DATOS no tiene pedidos.", vbInformation, "Trazabilidad"
    Exit Function
  End If
  v = wsD.Range(wsD.Cells(1, 1), wsD.Cells(lr, 41)).Value
  e = wsE.Range(wsE.Cells(1, 1), wsE.Cells(lr, 14)).Value
  Set dC = DictCobertura(): Set dE = DictEmpaque()

  ' sugerencias vigentes de las reglas + cobertura (lo que propone "Ver cambios sugeridos")
  Set dSug = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  nSug = ProponerDestinos()
  If Not gPropuestas Is Nothing Then
    For i = 1 To gPropuestas.Count
      p = gPropuestas(i)
      dSug(CLng(p(0))) = Array(TXE(p(7)), TXE(p(8)))
    Next
  End If
  Err.Clear
  On Error GoTo fallo

  enc = Array("FILA", "PEDIDO", "DESTINATARIO", "PROVINCIA", "CANTÓN", "PARROQUIA", "PARR. TMS", "COD. POSTAL TMS", _
              "DIRECCIÓN", "TELÉFONO", "DESTINO FINAL", "ORIGEN DEL DESTINO", "DESTINO POR REGLA BASE", _
              "SUGERIDO AHORA (sin aplicar)", "MOTIVO DEL SUGERIDO", "MOTIVO / QUIÉN Y CUÁNDO", "COURIER", "TRAYECTO", _
              "ZONA PELIGROSA", "COBERTURA TMS", "DIAGNÓSTICO COBERTURA", "BULTOS", "PESO KG", _
              "ETIQUETA", "DESTINO IMPRESO", "FECHA IMPRESIÓN", "EMPAQUE", "CAJAS", "UNID. CONF/SOL")
  ReDim a(1 To lr, 1 To UBound(enc) + 1)
  For i = 2 To lr
    ped = TXE(v(i, D_PED))
    If Len(ped) > 0 Then
      n = n + 1
      dest = UCase$(TXE(v(i, D_DEST)))
      base = DestinoBase(TXE(v(i, D_PROV)))
      conf = (Len(TXE(v(i, D_RDEST))) > 0 And TXE(v(i, D_RPED)) = ped)
      If conf Then nConf = nConf + 1
      k = ClaveTMS(TXE(v(i, D_H)), TXE(v(i, D_I)), TXE(v(i, D_J)))
      cob = Empty: If dC.Exists(k) Then cob = dC(k)
      a(n, 1) = i: a(n, 2) = ped: a(n, 3) = TXE(v(i, D_NOM))
      a(n, 4) = TXE(v(i, D_PROV)): a(n, 5) = TXE(v(i, D_CANT)): a(n, 6) = TXE(v(i, D_PARR))
      a(n, 7) = TXE(v(i, D_PARRSP)): a(n, 8) = TXE(v(i, D_CPTMS))
      a(n, 9) = TXE(v(i, D_DIR)): a(n, 10) = TXE(v(i, D_TEL))
      a(n, 11) = dest
      If conf Then
        If InStr(1, TXE(v(i, D_RMOT)), "MANUAL", vbTextCompare) > 0 Then
          a(n, 12) = "CONFIRMADO A MANO"
        Else
          a(n, 12) = "CONFIRMADO POR REGLA"
        End If
      ElseIf Len(dest) = 0 Then
        a(n, 12) = "SIN DESTINO (fuera de cobertura)"
      Else
        a(n, 12) = "REGLA BASE POR PROVINCIA"
      End If
      a(n, 13) = base
      If dSug.Exists(i) Then
        x = dSug(i)
        If UCase$(x(0)) <> dest Then a(n, 14) = x(0): a(n, 15) = x(1)
      End If
      a(n, 16) = IIf(conf, TXE(v(i, D_RMOT)), "")
      a(n, 17) = TXE(v(i, D_COUR))
      a(n, 18) = TXE(v(i, 37)): If Len(a(n, 18)) = 0 Then a(n, 18) = CobVal2(cob, 2)
      a(n, 19) = TXE(v(i, 39))
      a(n, 20) = CobVal2(cob, 0)
      If UCase$(TXE(v(i, D_VAL))) = "REVISAR" Then
        nFue = nFue + 1
        a(n, 21) = "FUERA DE COBERTURA: " & TXE(v(i, D_DIAG)) & IIf(Len(TXE(v(i, D_SUG))) > 0, " (sugerida: " & TXE(v(i, D_SUG)) & ")", "")
      Else
        a(n, 21) = "OK"
      End If
      a(n, 22) = TXE(v(i, D_BUL)): a(n, 23) = TXE(v(i, D_PESO))
      If UCase$(TXE(e(i, 6))) = "OK" Then
        If Len(TXE(e(i, 13))) > 0 And UCase$(TXE(e(i, 13))) <> dest Then a(n, 24) = "REIMPRIMIR" Else a(n, 24) = "IMPRESA"
        a(n, 25) = TXE(e(i, 13)): a(n, 26) = TXE(e(i, 14))
      Else
        a(n, 24) = "PENDIENTE"
      End If
      a(n, 27) = EstadoEmpaque(dE, ped)
      If dE.Exists(ped) Then
        x = dE(ped)
        a(n, 28) = x(1): a(n, 29) = x(4) & "/" & x(5)
      End If
    End If
  Next

  Set ws = HojaTraza()
  Congelar
  If ws.AutoFilterMode Then ws.AutoFilterMode = False
  ws.Cells.Clear
  ws.Range("A1").Value = "TRAZABILIDAD DE PEDIDOS - HYCITE"
  ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
  ws.Range("A2").Value = "Generada el " & Format(Now, "dd/mm/yyyy hh:nn") & " por " & Application.UserName & _
                         "   ·   " & n & " pedido(s)   ·   " & nConf & " con destino confirmado   ·   " & _
                         nFue & " fuera de cobertura TMS   ·   " & dSug.Count & " con sugerencia vigente"
  ws.Range("A2").Font.Italic = True
  ws.Range("A4").Value = "1. ESTADO FINAL DE CADA PEDIDO (hoja DATOS tal como quedó)"
  Bloque ws, 4
  For j = 0 To UBound(enc): ws.Cells(5, j + 1).Value = enc(j): Next
  Encabezado ws.Range(ws.Cells(5, 1), ws.Cells(5, UBound(enc) + 1))
  If n > 0 Then
    ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, UBound(enc) + 1)).Value = a
    ws.Range(ws.Cells(5, 1), ws.Cells(5 + n, UBound(enc) + 1)).AutoFilter
  End If

  r = 5 + n + 2
  ws.Cells(r, 1).Value = "2. CAMBIOS DE DESTINO CONFIRMADOS (lo que se apartó de la regla base)"
  Bloque ws, r
  r = r + 1
  enc = Array("FILA", "PEDIDO", "DESTINATARIO", "PROVINCIA", "CANTÓN", "PARROQUIA", "DESTINO SIN CONFIRMAR (regla base)", _
              "DESTINO FINAL", "MOTIVO / QUIÉN Y CUÁNDO")
  For j = 0 To UBound(enc): ws.Cells(r, j + 1).Value = enc(j): Next
  Encabezado ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(enc) + 1))
  j = 0
  For i = 1 To n
    If Left$(TXE(a(i, 12)), 11) = "CONFIRMADO " Then
      j = j + 1
      ws.Cells(r + j, 1).Value = a(i, 1): ws.Cells(r + j, 2).Value = a(i, 2): ws.Cells(r + j, 3).Value = a(i, 3)
      ws.Cells(r + j, 4).Value = a(i, 4): ws.Cells(r + j, 5).Value = a(i, 5): ws.Cells(r + j, 6).Value = a(i, 6)
      ws.Cells(r + j, 7).Value = a(i, 13): ws.Cells(r + j, 8).Value = a(i, 11): ws.Cells(r + j, 9).Value = a(i, 16)
    End If
  Next
  If j = 0 Then ws.Cells(r + 1, 1).Value = "(ningún pedido tiene destino confirmado: todos salen por la regla base de provincia)"

  r = r + j + 3
  ws.Cells(r, 1).Value = "3. REGISTRO DE ACCIONES DEL DÍA (hoja oculta LOG_EGR)"
  Bloque ws, r
  r = r + 1
  enc = Array("FECHA", "USUARIO", "NIVEL", "ACCIÓN")
  For j = 0 To UBound(enc): ws.Cells(r, j + 1).Value = enc(j): Next
  Encabezado ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
  Set wsL = Hoja(HLOGE)
  j = 0
  If Not wsL Is Nothing Then
    Dim lrl As Long, w
    lrl = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row
    If lrl >= 2 Then
      w = wsL.Range("A2:D" & lrl).Value
      For i = 1 To UBound(w, 1)
        If IsDate(w(i, 1)) Then
          If Int(CDate(w(i, 1))) = Date Then
            j = j + 1
            ws.Cells(r + j, 1).Value = Format(w(i, 1), "dd/mm/yyyy hh:nn:ss")
            ws.Cells(r + j, 2).Value = w(i, 2): ws.Cells(r + j, 3).Value = w(i, 3): ws.Cells(r + j, 4).Value = w(i, 4)
          End If
        End If
      Next
    End If
  End If
  If j = 0 Then ws.Cells(r + 1, 1).Value = "(no hay acciones registradas hoy)"

  ws.Columns("A:AC").AutoFit
  For j = 1 To 29
    If ws.Columns(j).ColumnWidth > 42 Then ws.Columns(j).ColumnWidth = 42
  Next
  Liberar
  LogE "TRAZABILIDAD: hoja generada con " & n & " pedido(s), " & nConf & " destino(s) confirmado(s) y el registro del día"
  GenerarTrazabilidad = n
  Exit Function
fallo:
  Liberar
  LogE "TRAZABILIDAD: " & Err.Description, "ERROR"
  If avisar Then MsgBox "No se pudo generar la trazabilidad: " & Err.Description, vbExclamation
End Function

' Rehace la hoja TRAZABILIDAD sin interrumpir nunca lo que se estaba haciendo.
' Se llama después de aplicar destinos: si fallara, solo queda anotado en el registro.
Public Sub TrazaAuto()
  On Error Resume Next
  GenerarTrazabilidad False
End Sub

' Lee un dato de la cobertura sin que IIf evalúe las dos ramas (mismo criterio que el panel)
Private Function CobVal2(ByVal cob As Variant, ByVal i As Long) As String
  If IsArray(cob) Then
    If i >= LBound(cob) And i <= UBound(cob) Then CobVal2 = TXE(cob(i))
  End If
End Function

Private Function HojaTraza() As Worksheet
  Dim ws As Worksheet
  Set ws = Hoja(HTRZ)
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = HTRZ
  End If
  ws.Visible = xlSheetVisible
  Set HojaTraza = ws
End Function

Private Sub Bloque(ws As Worksheet, ByVal fila As Long)
  With ws.Cells(fila, 1)
    .Font.Bold = True: .Font.Size = 11: .Font.Color = RGB(48, 84, 150)
  End With
End Sub

Private Sub Encabezado(rg As Range)
  With rg
    .Font.Bold = True: .Font.Color = vbWhite
    .Interior.Color = RGB(48, 84, 150)
    .HorizontalAlignment = xlCenter
  End With
End Sub

' Deja una copia de la hoja TRAZABILIDAD en la carpeta de exportes. Devuelve la ruta.
Public Function GuardarCopiaTraza() As String
  Dim wb As Workbook, ruta As String
  On Error GoTo fallo
  Application.ScreenUpdating = False: Application.DisplayAlerts = False
  ThisWorkbook.Worksheets(HTRZ).Copy
  Set wb = ActiveWorkbook
  ruta = CarpetaExportes() & Application.PathSeparator & "TRAZABILIDAD_" & Format(Now, "yyyymmdd_hhnn") & ".xlsx"
  wb.SaveAs Filename:=ruta, FileFormat:=51
  wb.Close SaveChanges:=False
  Application.DisplayAlerts = True: Application.ScreenUpdating = True
  LogE "TRAZABILIDAD: copia guardada en " & ruta
  GuardarCopiaTraza = ruta
  Exit Function
fallo:
  On Error Resume Next
  If Not wb Is Nothing Then wb.Close SaveChanges:=False
  Application.DisplayAlerts = True: Application.ScreenUpdating = True
  LogE "TRAZABILIDAD: no se pudo guardar la copia: " & Err.Description, "ERROR"
  MsgBox "La hoja quedó lista, pero no se pudo guardar la copia: " & Err.Description, vbExclamation
End Function

' Texto del aviso, igual desde el menú y desde el panel
Public Function TextoTraza(ByVal n As Long) As String
  TextoTraza = "Hoja TRAZABILIDAD lista con " & n & " pedido(s):" & vbCrLf & vbCrLf & _
               "  1. estado final de cada pedido" & vbCrLf & _
               "  2. cambios de destino confirmados (quién, cuándo y por qué)" & vbCrLf & _
               "  3. registro de acciones de hoy" & vbCrLf & vbCrLf & _
               "¿Guardar además una copia en la carpeta de exportes?"
End Function

' Genera la hoja, la muestra y ofrece dejar una copia en la carpeta de exportes.
Sub MenuTrazabilidad()
  Dim n As Long, ruta As String
  If PanelOcupado() Then Exit Sub
  RuedaDesactivar
  n = GenerarTrazabilidad(True)
  If n = 0 Then Exit Sub
  On Error Resume Next
  ThisWorkbook.Worksheets(HTRZ).Activate
  On Error GoTo 0
  If MsgBox(TextoTraza(n), vbYesNo + vbQuestion, "Trazabilidad") <> vbYes Then Exit Sub
  ruta = GuardarCopiaTraza()
  If Len(ruta) > 0 Then MsgBox "Copia guardada en:" & vbCrLf & ruta, vbInformation, "Trazabilidad"
End Sub
