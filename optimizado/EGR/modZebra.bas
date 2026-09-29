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
Public gEtiqFilas As Collection      ' filas de DATOS para frmEtiquetas

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

' Lista de impresoras instaladas (primero WScript.Network, más liviano; WMI solo si hace falta)
Public Function ListaImpresoras() As Collection
  Dim c As New Collection, net As Object, cn As Object, i As Long, wmi As Object, it As Object, u As Object
  Set u = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  Set net = CreateObject("WScript.Network")
  Set cn = net.EnumPrinterConnections
  If Not cn Is Nothing Then
    For i = 0 To cn.Count - 1 Step 2
      If Len(cn.Item(i + 1)) > 0 And Not u.Exists(cn.Item(i + 1)) Then u(cn.Item(i + 1)) = 1: c.Add CStr(cn.Item(i + 1))
    Next
  End If
  If c.Count = 0 Then
    Err.Clear
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If Not wmi Is Nothing Then
      For Each it In wmi.ExecQuery("SELECT Name FROM Win32_Printer")
        If Not u.Exists(CStr(it.Name)) Then u(CStr(it.Name)) = 1: c.Add CStr(it.Name)
      Next
    End If
  End If
  On Error GoTo 0
  Set ListaImpresoras = c
End Function

' Imprime (o guarda en archivo) las etiquetas de las filas de DATOS indicadas.
' Marca en ETIQUETAS: F = OK, M = destino impreso, N = fecha. Devuelve cuántas se enviaron.
Public Function ImprimirFilas(filas As Collection, ByVal impresora As String, ByVal soloArchivo As Boolean) As Long
  Dim wsD As Worksheet, wsE As Worksheet, f, zpl As String, n As Long, ruta As String, ff As Integer
  Dim ped As String, nom As String, dest As String, parr As String
  Set wsD = ThisWorkbook.Worksheets(HDAT)
  Set wsE = ThisWorkbook.Worksheets(HETQ)
  For Each f In filas
    ped = TXE(wsD.Cells(f, D_PED).Value): dest = TXE(wsD.Cells(f, D_DEST).Value)
    nom = TXE(wsD.Cells(f, D_NOM).Value): parr = TXE(wsD.Cells(f, D_PARRSP).Value)
    If Len(parr) = 0 Then parr = TXE(wsD.Cells(f, D_PARR).Value)
    If Len(ped) > 0 And Len(dest) > 0 Then zpl = zpl & ZplEtiqueta(ped, nom, dest, parr) & vbCrLf: n = n + 1
  Next
  If n = 0 Then Exit Function
  If soloArchivo Then
    ruta = CarpetaExportes() & Application.PathSeparator & "ETIQUETAS_" & Format(Now, "yyyymmdd_hhnnss") & ".zpl"
    ff = FreeFile: Open ruta For Output As #ff: Print #ff, zpl;: Close #ff
    LogE "ETIQUETAS: " & n & " etiqueta(s) guardadas en " & ruta & " (se pueden ver en labelary.com)"
    ImprimirFilas = n
    Exit Function
  End If
  If Not EnviarRaw(impresora, zpl) Then Exit Function
  For Each f In filas
    dest = TXE(wsD.Cells(f, D_DEST).Value)
    If Len(dest) > 0 Then
      wsE.Cells(f, 6).Value = "OK"
      wsE.Cells(f, 13).Value = dest
      wsE.Cells(f, 14).Value = Now
    End If
  Next
  If Len(TXE(wsE.Cells(1, 13).Value)) = 0 Then wsE.Cells(1, 13).Value = "DESTINO IMPRESO": wsE.Cells(1, 14).Value = "FECHA IMPRESION"
  LogE "ETIQUETAS: " & n & " etiqueta(s) enviadas a " & impresora
  ImprimirFilas = n
End Function

' ---------- vista previa: Code 128 (subconjunto B) solo para dibujar en pantalla ----------
Public Function Code128Modulos(ByVal s As String) As String
  Dim pat, i As Long, v As Long, chk As Long, o As String
  pat = Split("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 113222 123122 123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 212123 212321 232121 111323 131123 131321 112313 132113 132311 211313 231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 311321 331121 312113 312311 332111 314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 122411 142112 142211 241211 221114 413111 241112 134111 111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 214121 412121 111143 111341 131141 114113 114311 411113 411311 113141 114131 311141 411131 211412 211214 211232 2331112", " ")
  chk = 104: o = pat(104)
  For i = 1 To Len(s)
    v = Asc(Mid$(s, i, 1)) - 32
    If v < 0 Or v > 94 Then v = 0
    chk = chk + v * i: o = o & pat(v)
  Next
  o = o & pat(chk Mod 103) & pat(106)
  Code128Modulos = o          ' anchos alternados barra/espacio en módulos
End Function
