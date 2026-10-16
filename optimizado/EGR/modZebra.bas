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
  ' Diseño (800 x 400 puntos = 10 x 5 cm a 203 ppp), con bandas que NO se pisan:
  '    26..138  código de barras            140..196  número del pedido
  '   186..282  destino en grande, derecha  216..322  destinatario, 1 o 2 renglones, izquierda
  '   326..362  parroquia, derecha
  ' El nombre se parte AQUÍ, no con ^FB: cuando el texto no entraba en los renglones
  ' de ^FB, la impresora escribía el sobrante ENCIMA del último renglón.
  Dim ox As Long, oy As Long, osc As Long, z As String, nom As String
  Dim l1 As String, l2 As String, alto As Long, anc As Long, y2 As Long
  nom = Limpio(destinatario)
  pedido = Limpio(pedido): destino = Limpio(destino): parroquia = Limpio(parroquia)
  parroquia = Recortar(parroquia, 23)
  NombreEnDos nom, 340, l1, l2, alto, anc
  ox = Val(Cfg("ETIQ_OFFSET_X", "0")): oy = Val(Cfg("ETIQ_OFFSET_Y", "0"))
  osc = Val(Cfg("ETIQ_OSCURIDAD", "12")): If osc < 0 Or osc > 30 Then osc = 12
  z = "^XA^CI0^PW800^LL400^LH" & ox & "," & oy & "^MD" & osc & "^PR3"
  z = z & "^FO140,26^BY4,2.5,112^BCN,112,N,N,N^FD" & pedido & "^FS"            ' barras
  z = z & "^FO140,150^A0N,46,40^FB400,1,0,C^FD" & pedido & "^FS"               ' número centrado bajo las barras
  z = z & "^FO390,186^A0N,96,86^FB380,1,0,R^FD" & destino & "^FS"              ' UIO / GYE / PRO / GPS
  z = z & "^FO30,216^A0N," & alto & "," & anc & "^FB340,1,0,L^FD" & l1 & "^FS" ' destinatario, renglón 1
  If Len(l2) > 0 Then
    y2 = 216 + alto + 6
    z = z & "^FO30," & y2 & "^A0N," & alto & "," & anc & "^FB340,1,0,L^FD" & l2 & "^FS"
  End If
  z = z & "^FO330,326^A0N,36,30^FB440,1,0,R^FD" & parroquia & "^FS"            ' parroquia
  z = z & "^PQ1^XZ"
  ZplEtiqueta = z
End Function

' Parte el nombre en uno o dos renglones que quepan en 'ancho' puntos, con la letra
' más grande con la que entre. Devuelve también el alto y el ancho de la letra.
Private Sub NombreEnDos(ByVal nom As String, ByVal ancho As Long, ByRef l1 As String, ByRef l2 As String, _
                         ByRef alto As Long, ByRef anc As Long)
  Dim tam, k As Long, cmax As Long
  tam = Array(48, 42, 36, 30, 26)
  For k = 0 To UBound(tam)
    alto = CLng(tam(k)): anc = CLng(tam(k) * 0.85)
    cmax = CLng(ancho / (anc * 0.62))
    If cmax < 4 Then cmax = 4
    If Reparte(nom, cmax, l1, l2) Then Exit Sub
  Next
  ' ni con la letra más chica entra: se recorta, pero nunca se escribe encima
  l1 = Recortar(nom, cmax)
  l2 = Recortar(Trim$(Mid$(nom, Len(l1) + 1)), cmax)
End Sub

' Reparte el texto en dos renglones de como máximo cmax caracteres, sin cortar palabras.
' Devuelve False si no alcanza (hay que probar con letra más chica).
Private Function Reparte(ByVal s As String, ByVal cmax As Long, ByRef l1 As String, ByRef l2 As String) As Boolean
  Dim pal, i As Long, cur As String
  l1 = "": l2 = ""
  pal = Split(Trim$(s), " ")
  For i = 0 To UBound(pal)
    If Len(pal(i)) > 0 Then
      If Len(pal(i)) > cmax Then Exit Function
      If Len(l2) = 0 Then
        cur = Trim$(l1 & " " & pal(i))
        If Len(cur) <= cmax Then
          l1 = cur
        Else
          l2 = pal(i)
        End If
      Else
        cur = l2 & " " & pal(i)
        If Len(cur) <= cmax Then l2 = cur Else Exit Function
      End If
    End If
  Next
  Reparte = (Len(l1) > 0)
End Function

Private Function Recortar(ByVal s As String, ByVal n As Long) As String
  If Len(s) <= n Then Recortar = s Else Recortar = Trim$(Left$(s, n))
End Function

' ---------- envío RAW a la impresora ----------
Public Function EnviarRaw(ByVal impresora As String, ByVal datos As String) As Boolean
  Dim nombres As New Collection, nm, p As Long
  If Len(datos) = 0 Or Len(impresora) = 0 Then Exit Function
  nombres.Add impresora
  ' Windows muestra las impresoras compartidas como "IMPRESORA en SERVIDOR"; el nombre real es \\SERVIDOR\IMPRESORA
  p = InStrRev(impresora, " en ")
  If p > 0 And Left$(impresora, 2) <> "\\" Then nombres.Add "\\" & Mid$(impresora, p + 4) & "\" & Left$(impresora, p - 1)
  For Each nm In nombres
    If EnviarRawA(CStr(nm), datos) Then EnviarRaw = True: Exit Function
  Next
  ' último recurso para impresoras compartidas: copiar el ZPL directo al recurso compartido
  For Each nm In nombres
    If Left$(nm, 2) = "\\" Then
      If CopiarAImpresora(CStr(nm), datos) Then EnviarRaw = True: Exit Function
    End If
  Next
End Function

Private Function EnviarRawA(ByVal impresora As String, ByVal datos As String) As Boolean
  Dim h As LongPtr, di As DOCINFO, b() As Byte, esc As Long, r As Long
  On Error GoTo fallo
  If OpenPrinter(impresora, h, 0) = 0 Then
    LogE "ETIQUETAS: no se pudo abrir '" & impresora & "' (código Windows " & Err.LastDllError & ")", "AVISO"
    Exit Function
  End If
  di.pDocName = "Etiquetas HYCITE": di.pOutputFile = vbNullString: di.pDatatype = "RAW"
  If StartDocPrinter(h, 1, di) = 0 Then
    LogE "ETIQUETAS: '" & impresora & "' rechazó el trabajo (código Windows " & Err.LastDllError & ")", "AVISO"
    ClosePrinter h: Exit Function
  End If
  StartPagePrinter h
  b = StrConv(datos, vbFromUnicode)
  r = WritePrinter(h, b(0), UBound(b) + 1, esc)
  EndPagePrinter h
  EndDocPrinter h
  ClosePrinter h
  EnviarRawA = (r <> 0 And esc = UBound(b) + 1)
  If Not EnviarRawA Then LogE "ETIQUETAS: '" & impresora & "' recibió " & esc & " de " & (UBound(b) + 1) & " bytes", "AVISO"
  Exit Function
fallo:
  LogE "ETIQUETAS: error al enviar a '" & impresora & "': " & Err.Description, "AVISO"
  On Error Resume Next
  If h <> 0 Then ClosePrinter h
End Function

Private Function CopiarAImpresora(ByVal destino As String, ByVal datos As String) As Boolean
  Dim f As String, ff As Integer, r As Long
  On Error GoTo fallo
  f = Environ$("TEMP") & "\etiquetas_hycite.zpl"
  ff = FreeFile: Open f For Output As #ff: Print #ff, datos;: Close #ff
  r = CreateObject("WScript.Shell").Run("cmd /c copy /b """ & f & """ """ & destino & """", 0, True)
  CopiarAImpresora = (r = 0)
  LogE "ETIQUETAS: copia directa a " & destino & IIf(r = 0, " correcta", " falló (código " & r & ")"), IIf(r = 0, "INFO", "AVISO")
  Exit Function
fallo:
  LogE "ETIQUETAS: copia directa a " & destino & ": " & Err.Description, "AVISO"
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
' Devuelve las mismas filas, ordenadas por DESTINO, PARROQUIA y PEDIDO.
' Si algo fallara, devuelve la colección tal como entró: nunca deja de imprimir.
Public Function OrdenarEtiquetas(filas As Collection) As Collection
  Dim ws As Worksheet, n As Long, i As Long, j As Long, f, cl() As String, fi() As Long
  Dim t As Long, ct As String, h As Long, p As String, c As New Collection
  On Error GoTo salir
  n = filas.Count
  If n < 2 Then Set OrdenarEtiquetas = filas: Exit Function
  Set ws = ThisWorkbook.Worksheets(HDAT)
  ReDim cl(1 To n): ReDim fi(1 To n)
  i = 0
  For Each f In filas
    i = i + 1
    fi(i) = CLng(f)
    p = TXE(ws.Cells(fi(i), D_PARRSP).Value)
    If Len(p) = 0 Then p = TXE(ws.Cells(fi(i), D_PARR).Value)
    cl(i) = UCase$(TXE(ws.Cells(fi(i), D_DEST).Value)) & Chr(1) & UCase$(p) & Chr(1) & TXE(ws.Cells(fi(i), D_PED).Value)
  Next
  h = 1
  Do While h < n \ 3: h = h * 3 + 1: Loop
  Do While h >= 1
    For i = h + 1 To n
      t = fi(i): ct = cl(i)
      j = i
      Do While j > h
        If cl(j - h) <= ct Then Exit Do
        fi(j) = fi(j - h): cl(j) = cl(j - h)
        j = j - h
      Loop
      fi(j) = t: cl(j) = ct
    Next
    h = (h - 1) \ 3
  Loop
  For i = 1 To n: c.Add fi(i): Next
  Set OrdenarEtiquetas = c
  Exit Function
salir:
  LogE "ETIQUETAS: no se pudo ordenar la lista (" & Err.Description & "); se imprime en el orden recibido.", "AVISO"
  Set OrdenarEtiquetas = filas
End Function

Public Function ImprimirFilas(filas As Collection, ByVal impresora As String, ByVal soloArchivo As Boolean) As Long
  Dim wsD As Worksheet, wsE As Worksheet, f, zpl As String, n As Long, ruta As String, ff As Integer
  Dim ped As String, nom As String, dest As String, parr As String, orden As Collection
  Set wsD = ThisWorkbook.Worksheets(HDAT)
  Set wsE = ThisWorkbook.Worksheets(HETQ)
  ' Las etiquetas salen SIEMPRE en el mismo orden: destino, parroquia y pedido. Así un
  ' lote mezclado sale ya agrupado por ruta, venga de donde venga la lista de filas.
  Set orden = OrdenarEtiquetas(filas)
  For Each f In orden
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
  For Each f In orden
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


' ---------- impresión desde Complementos (sin formularios) ----------
Public Sub ElegirImpresoraMenu()
  Dim c As Collection, i As Long, s As String, r As String, sug As Long
  RuedaDesactivar
  Set c = ListaImpresoras()
  If c.Count = 0 Then MsgBox "No se encontraron impresoras instaladas.", vbExclamation: Exit Sub
  For i = 1 To c.Count
    s = s & i & ") " & c(i) & vbCrLf
    If sug = 0 And (InStr(1, c(i), "ZD", vbTextCompare) > 0 Or InStr(1, c(i), "ZEBRA", vbTextCompare) > 0) Then sug = i
  Next
  r = InputBox("Impresora de etiquetas:" & vbCrLf & vbCrLf & s & vbCrLf & "Escribe el número:", "Elegir impresora", IIf(sug > 0, CStr(sug), ""))
  If Len(r) = 0 Then Exit Sub
  If Val(r) < 1 Or Val(r) > c.Count Then MsgBox "Número no válido.", vbExclamation: Exit Sub
  SetCfg "IMPRESORA_ZEBRA", c(Val(r))
  LogE "ETIQUETAS: impresora elegida = " & c(Val(r))
  MsgBox "Impresora guardada: " & c(Val(r)), vbInformation
End Sub

' modo: "PENDIENTES" (sin imprimir o por reimprimir), "TODAS", "PEDIDOS" (lista)
Public Sub ImprimirMenu(ByVal modo As String, Optional ByVal lista As String = "")
  Dim wsD As Worksheet, wsE As Worksheet, lr As Long, i As Long, filas As New Collection, dest As String, ped As String
  Dim impr As String, n As Long, l As String
  RuedaDesactivar
  Set wsD = ThisWorkbook.Worksheets(HDAT): Set wsE = ThisWorkbook.Worksheets(HETQ)
  Application.Calculate
  lr = UltimaFilaDatos()
  l = "," & Replace(Replace(lista, " ", ""), ";", ",") & ","
  For i = 2 To lr
    ped = TXE(wsD.Cells(i, D_PED).Value): dest = UCase$(TXE(wsD.Cells(i, D_DEST).Value))
    If Len(ped) > 0 And Len(dest) > 0 Then
      Select Case modo
        Case "TODAS": filas.Add i
        Case "PEDIDOS": If InStr(l, "," & ped & ",") > 0 Then filas.Add i
        Case Else
          If UCase$(TXE(wsE.Cells(i, 6).Value)) <> "OK" Or (Len(TXE(wsE.Cells(i, 13).Value)) > 0 And UCase$(TXE(wsE.Cells(i, 13).Value)) <> dest) Then filas.Add i
      End Select
    End If
  Next
  If filas.Count = 0 Then MsgBox "No hay etiquetas para imprimir con esa opción.", vbInformation: Exit Sub
  impr = Cfg("IMPRESORA_ZEBRA")
  If Len(impr) = 0 Then ElegirImpresoraMenu: impr = Cfg("IMPRESORA_ZEBRA")
  If Len(impr) = 0 Then Exit Sub
  If MsgBox("Imprimir " & filas.Count & " etiqueta(s) en:" & vbCrLf & impr & "?", vbYesNo + vbQuestion, "Etiquetas") <> vbYes Then Exit Sub
  Application.Cursor = xlWait
  n = ImprimirFilas(filas, impr, False)
  Application.Cursor = xlDefault
  If n > 0 Then
    MsgBox n & " etiqueta(s) enviadas.", vbInformation
  Else
    MsgBox "No se pudo imprimir en '" & impr & "'. Revisa el registro (LOG_EGR) o elige otra vez la impresora.", vbExclamation
  End If
End Sub

Sub MenuImprimirPendientes()
  ImprimirMenu "PENDIENTES"
End Sub

Sub MenuImprimirTodas()
  ImprimirMenu "TODAS"
End Sub

Sub MenuReimprimir()
  Dim s As String
  RuedaDesactivar
  s = InputBox("Número(s) de pedido a imprimir, separados por coma:", "Reimprimir etiquetas")
  If Len(Trim$(s)) > 0 Then ImprimirMenu "PEDIDOS", s
End Sub
