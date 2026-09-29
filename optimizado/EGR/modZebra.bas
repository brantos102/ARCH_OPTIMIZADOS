Option Explicit
' =====================================================================================
'  modZebra - Etiquetas de pedido para Zebra ZD230 (203 dpi, ZPL), etiqueta térmica 10 x 5 cm
'  Una etiqueta por pedido (fila de DATOS), igual al modelo:
'     código de barras Code 128 con el número de pedido y el número debajo
'     DESTINO grande (GYE / UIO / PRO / GPS) arriba a la derecha y la PARROQUIA debajo
'     NOMBRES (línea grande) y APELLIDOS (línea chica) abajo a la izquierda
'  Se envía ZPL directo a la impresora (sin controlador de página ni fuentes de código de barras).
'  La hoja ETIQUETAS queda como lista de control: columna F (STATUS) = OK cuando se imprimió.
' =====================================================================================
Private Type DOCINFO
  pDocName As String
  pOutputFile As String
  pDatatype As String
End Type
Private Declare PtrSafe Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As LongPtr, ByVal pDefault As LongPtr) As Long
Private Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare PtrSafe Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long

Private Const HETQ As String = "ETIQUETAS"

' ---------- texto seguro para ZPL (sin tildes, sin ^ ni ~) ----------
Private Function Limpio(ByVal s As String) As String
  Dim a, b, i As Long
  a = Array("Á", "É", "Í", "Ó", "Ú", "Ä", "Ë", "Ï", "Ö", "Ü", "Ñ", "À", "È", "Ì", "Ò", "Ù")
  b = Array("A", "E", "I", "O", "U", "A", "E", "I", "O", "U", "N", "A", "E", "I", "O", "U")
  s = UCase$(Trim$(s))
  For i = LBound(a) To UBound(a): s = Replace(s, a(i), b(i)): Next
  s = Replace(s, "^", " "): s = Replace(s, "~", " "): s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
  Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
  Limpio = s
End Function

' Nombres / apellidos: 4+ palabras -> 2 y resto; 3 -> 1 y 2; 2 -> 1 y 1
Public Sub PartirNombre(ByVal completo As String, ByRef nombres As String, ByRef apellidos As String)
  Dim p, n As Long, i As Long, k As Long
  completo = Limpio(completo)
  nombres = completo: apellidos = ""
  If Len(completo) = 0 Then Exit Sub
  p = Split(completo, " "): n = UBound(p) + 1
  If n <= 1 Then Exit Sub
  If n >= 4 Then k = 2 Else k = 1
  nombres = "": apellidos = ""
  For i = 0 To n - 1
    If i < k Then nombres = nombres & IIf(Len(nombres) > 0, " ", "") & p(i) Else apellidos = apellidos & IIf(Len(apellidos) > 0, " ", "") & p(i)
  Next
End Sub

' ZPL de una etiqueta 10 x 5 cm a 203 dpi (800 x 400 puntos)
Public Function ZplEtiqueta(ByVal pedido As String, ByVal destinatario As String, ByVal destino As String, ByVal parroquia As String) As String
  Dim nom As String, ape As String, ox As Long, oy As Long, z As String
  PartirNombre destinatario, nom, ape
  pedido = Limpio(pedido): destino = Limpio(destino): parroquia = Limpio(parroquia)
  ox = Val(Cfg("ETIQ_OFFSET_X", "0")): oy = Val(Cfg("ETIQ_OFFSET_Y", "0"))
  z = "^XA^CI0^PW800^LL400^LH" & ox & "," & oy & "^MD10"
  z = z & "^FO40,40^BY3,3,100^BCN,100,N,N,N,A^FD" & pedido & "^FS"                 ' código de barras
  z = z & "^FO40,148^A0N,36,30^FD" & EspaciarDigitos(pedido) & "^FS"              ' número legible
  z = z & "^FO520,40^A0N,110,80^FB260,1,0,C^FD" & destino & "^FS"                   ' GYE / UIO / PRO
  z = z & "^FO420,160^A0N,36,30^FB360,2,4,C^FD" & parroquia & "^FS"                 ' parroquia
  z = z & "^FO40,250^A0N,50,42^FB740,1,0,L^FD" & nom & "^FS"                        ' nombres
  z = z & "^FO40,315^A0N,34,30^FB740,1,0,L^FD" & ape & "^FS"                        ' apellidos
  z = z & "^PQ1^XZ"
  ZplEtiqueta = z
End Function

Private Function EspaciarDigitos(ByVal s As String) As String
  Dim i As Long, o As String
  For i = 1 To Len(s): o = o & Mid$(s, i, 1) & IIf(i < Len(s), " ", ""): Next
  EspaciarDigitos = o
End Function

' ---------- envío RAW a la impresora ----------
Public Function EnviarRaw(ByVal impresora As String, ByVal datos As String) As Boolean
  Dim h As LongPtr, di As DOCINFO, b() As Byte, esc As Long
  If Len(datos) = 0 Then Exit Function
  If OpenPrinter(impresora, h, 0) = 0 Then LogE "ETIQUETAS: no se pudo abrir la impresora '" & impresora & "'", "ERROR": Exit Function
  di.pDocName = "Etiquetas HYCITE": di.pOutputFile = vbNullString: di.pDatatype = "RAW"
  If StartDocPrinter(h, 1, di) = 0 Then ClosePrinter h: LogE "ETIQUETAS: la impresora rechazó el trabajo", "ERROR": Exit Function
  StartPagePrinter h
  b = StrConv(datos, vbFromUnicode)
  WritePrinter h, b(0), UBound(b) + 1, esc
  EndPagePrinter h
  EndDocPrinter h
  ClosePrinter h
  EnviarRaw = (esc = UBound(b) + 1)
  If Not EnviarRaw Then LogE "ETIQUETAS: se enviaron " & esc & " de " & (UBound(b) + 1) & " bytes", "ERROR"
End Function

' Lista de impresoras instaladas
Public Function ListaImpresoras() As Collection
  Dim c As New Collection, wmi As Object, it As Object
  On Error Resume Next
  Set wmi = GetObject("winmgmts:\\.\root\cimv2")
  For Each it In wmi.ExecQuery("SELECT Name FROM Win32_Printer")
    c.Add CStr(it.Name)
  Next
  Set ListaImpresoras = c
End Function

Public Sub ElegirImpresora()
  Dim c As Collection, i As Long, s As String, r As String, sug As Long
  Set c = ListaImpresoras()
  If c.Count = 0 Then MsgBox "No se encontraron impresoras instaladas.", vbExclamation: Exit Sub
  For i = 1 To c.Count
    s = s & i & ") " & c(i) & vbCrLf
    If sug = 0 And (InStr(1, c(i), "ZD", vbTextCompare) > 0 Or InStr(1, c(i), "ZEBRA", vbTextCompare) > 0 Or InStr(1, c(i), "ZDESIGNER", vbTextCompare) > 0) Then sug = i
  Next
  r = InputBox("Impresora de etiquetas (Zebra ZD230):" & vbCrLf & vbCrLf & s & vbCrLf & "Escribe el número:", "Elegir impresora", IIf(sug > 0, CStr(sug), ""))
  If Len(r) = 0 Then Exit Sub
  If Val(r) < 1 Or Val(r) > c.Count Then MsgBox "Número no válido.", vbExclamation: Exit Sub
  SetCfg "IMPRESORA_ZEBRA", c(Val(r))
  LogE "ETIQUETAS: impresora elegida = " & c(Val(r))
End Sub

' ---------- impresión ----------
' modo: "PENDIENTES" (STATUS <> OK), "TODAS", "PEDIDOS" (lista de números separados por coma)
Public Sub ImprimirEtiquetas(ByVal modo As String, Optional ByVal listaPedidos As String = "", Optional ByVal soloArchivo As Boolean = False)
  Dim wsD As Worksheet, wsE As Worksheet, lr As Long, v, i As Long, zpl As String, n As Long, nSin As Long
  Dim filas As New Collection, imp As String, ped As String, f, lst As String, ruta As String, ff As Integer
  Set wsD = ThisWorkbook.Worksheets(HDAT)
  Set wsE = ThisWorkbook.Worksheets(HETQ)
  Application.Calculate
  lr = UltimaFilaDatos()
  If lr < 2 Then MsgBox "DATOS está vacío.", vbExclamation: Exit Sub
  v = wsD.Range(wsD.Cells(1, 1), wsD.Cells(lr, D_PARRSP)).Value
  lst = "," & Replace(Replace(listaPedidos, " ", ""), ";", ",") & ","
  For i = 2 To lr
    ped = TXE(v(i, D_PED))
    If Len(ped) > 0 Then
      Select Case UCase$(modo)
        Case "TODAS": filas.Add i
        Case "PENDIENTES": If UCase$(TXE(wsE.Cells(i, 6).Value)) <> "OK" Then filas.Add i
        Case "PEDIDOS": If InStr(lst, "," & ped & ",") > 0 Then filas.Add i
      End Select
    End If
  Next
  If filas.Count = 0 Then MsgBox "No hay etiquetas para imprimir con esa opción.", vbInformation: Exit Sub
  For Each f In filas
    If Len(TXE(v(f, D_DEST))) = 0 Then
      nSin = nSin + 1
      LogE "ETIQUETAS fila " & f & " pedido " & TXE(v(f, D_PED)) & ": sin DESTINO (fuera de cobertura TMS); no se imprime", "AVISO"
    Else
      zpl = zpl & ZplEtiqueta(TXE(v(f, D_PED)), TXE(v(f, D_NOM)), TXE(v(f, D_DEST)), IIf(Len(TXE(v(f, D_PARRSP))) > 0, TXE(v(f, D_PARRSP)), TXE(v(f, D_PARR)))) & vbCrLf
      n = n + 1
    End If
  Next
  If n = 0 Then MsgBox "Ninguna etiqueta tiene destino (revisa la cobertura TMS).", vbExclamation: Exit Sub
  If soloArchivo Then
    ruta = CarpetaExportes() & Application.PathSeparator & "ETIQUETAS_" & Format(Now, "yyyymmdd_hhnnss") & ".zpl"
    ff = FreeFile: Open ruta For Output As #ff: Print #ff, zpl;: Close #ff
    LogE "ETIQUETAS: " & n & " etiqueta(s) guardadas en " & ruta & " (se pueden ver en labelary.com)"
    MsgBox n & " etiqueta(s) guardadas en:" & vbCrLf & ruta, vbInformation
    Exit Sub
  End If
  imp = Cfg("IMPRESORA_ZEBRA")
  If Len(imp) = 0 Then ElegirImpresora: imp = Cfg("IMPRESORA_ZEBRA")
  If Len(imp) = 0 Then Exit Sub
  If MsgBox("Imprimir " & n & " etiqueta(s) en '" & imp & "'?" & IIf(nSin > 0, vbCrLf & nSin & " pedido(s) sin destino se omiten.", ""), vbYesNo + vbQuestion, "Etiquetas") <> vbYes Then Exit Sub
  If EnviarRaw(imp, zpl) Then
    For Each f In filas
      If Len(TXE(v(f, D_DEST))) > 0 Then wsE.Cells(f, 6).Value = "OK"
    Next
    LogE "ETIQUETAS: " & n & " etiqueta(s) enviadas a " & imp & " (" & modo & IIf(Len(listaPedidos) > 0, ": " & listaPedidos, "") & ")"
  Else
    MsgBox "No se pudo imprimir. Revisa que la Zebra esté encendida y elegida (registro del panel).", vbExclamation
  End If
End Sub
