Option Explicit
' =====================================================================================
'  PEDIDOS HCE - Módulo1 (versión optimizada v2)
'  Reemplaza COMPLETO el contenido de Módulo1. Conserva todos los nombres de macros
'  existentes (los botones y el formulario siguen funcionando) y agrega:
'   - Gestor / destino (UIO-GYE-GPS-PRO) / trayecto TRAMACO / zona peligrosa (cols V:Z)
'   - Sugerencia de cobertura cercana (cabecera cantonal, ciudad principal/secundaria)
'   - Cantón del cliente respetado cuando la parroquia existe en varios cantones
'   - CAMBIOS con historial acumulado + exportación
'   - Aviso final con opción de abrir el validador solo con los casos REVISAR
'   - Panel central (frmPanel) y barra de botones en el orden del flujo
' =====================================================================================

Const HD As String = "DEPOT"
Const H1 As String = "COBERTURA"
Const HZ As String = "ZONAS PELIGROSAS"
Const HLOG As String = "CAMBIOS"
Const C_DIR As Long = 2, C_REF As Long = 3
Const C_PROV As Long = 6, C_CANT As Long = 7, C_PARR As Long = 8
Const C_SIG As Long = 12
Const C_NN As Long = 14                 ' N:U panel de validación (igual que antes)
Const C_EXT As Long = 22                ' V:AD columnas nuevas
Const N_EXT As Long = 9                 ' V:AD
Const N_EXTC As Long = 7                ' V:AB calculadas por DatosExt (AC = original, AD = tipo de corrección)
Const SOLO_FILTRADO As Boolean = True
Const ZONA_PELIGROSA_A_REVISAR As Boolean = False   ' True = una zona peligrosa deja el pedido en REVISAR

Public gSoloRevisar As Boolean          ' frmValidar: mostrar solo REVISAR
Public gFilaInicio As Long              ' frmValidar: fila con la que empieza

Dim mTriple As Object, mCantonProvM As Object
Dim mPC As Object, mRev As Object, mProvParr As Object, mParrList As Object
Dim mCantonProv As Object, mParrGlobal As Object, mPcSet As Object, mCantonParr As Object, mParrOrig As Object
Dim mSiglas As Object, mCodeCant As Object
Dim mCobInfo As Object, mPrincProv As Object, mZonas As Collection
Dim cGest As Long, cGestSug As Long, cTray As Long


Sub Desbloquear()
  On Error Resume Next
  Application.ScreenUpdating = True: Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic: Application.Cursor = xlDefault
  Application.DisplayAlerts = True
  Dim sh As Worksheet
  For Each sh In ThisWorkbook.Worksheets: sh.ScrollArea = "": Next
  MsgBox "Pantalla desbloqueada.", vbInformation
End Sub

Function Mapa() As Object
  Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
  d("AZUAY") = "01": d("BOLIVAR") = "02": d("CANAR") = "03": d("CARCHI") = "04": d("COTOPAXI") = "05"
  d("CHIMBORAZO") = "06": d("EL ORO") = "07": d("ESMERALDAS") = "08": d("GUAYAS") = "09": d("IMBABURA") = "10"
  d("LOJA") = "11": d("LOS RIOS") = "12": d("MANABI") = "13": d("MORONA SANTIAGO") = "14": d("NAPO") = "15"
  d("PASTAZA") = "16": d("PICHINCHA") = "17": d("TUNGURAHUA") = "18": d("ZAMORA CHINCHIPE") = "19": d("GALAPAGOS") = "20"
  d("SUCUMBIOS") = "21": d("ORELLANA") = "22": d("FRANCISCO DE ORELLANA") = "22"
  d("SANTO DOMINGO DE LOS TSACHILAS") = "23": d("SANTA ELENA") = "24"
  Set Mapa = d
End Function

Sub ColorearCambios(ws As Worksheet, lr As Long)
  Dim i As Long
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo sc
    MarcaDif ws.Cells(i, C_NN), ws.Cells(i, C_PROV)
    MarcaDif ws.Cells(i, C_NN + 1), ws.Cells(i, C_CANT)
    MarcaDif ws.Cells(i, C_NN + 2), ws.Cells(i, C_PARR)
sc:
  Next
End Sub
Private Sub MarcaDif(celProp As Range, celOrig As Range)
  If Len(TX(celProp)) > 0 And Normaliza(TX(celProp)) <> Normaliza(TX(celOrig)) Then
    celProp.Interior.Color = RGB(255, 242, 204)      ' amarillo suave = corregido
  Else
    celProp.Interior.ColorIndex = xlNone
  End If
End Sub

Function Normaliza(ByVal s As String) As String
  Dim t As String, i As Long, ch As String, o As String, p1 As Long, p2 As Long, a, b
  s = Replace(s, "-", " ")
  Static reN As Object
  If reN Is Nothing Then Set reN = CreateObject("VBScript.RegExp"): reN.Global = True: reN.Pattern = ChrW(195) & "[^A-Za-z0-9 ]?"
  s = reN.Replace(s, "N")
  a = Array("Á", "É", "Í", "Ó", "Ú", "Ä", "Ë", "Ï", "Ö", "Ü", "Ñ", "À", "È", "Ì", "Ò", "Ù", "Â", "Ê")
  b = Array("A", "E", "I", "O", "U", "A", "E", "I", "O", "U", "N", "A", "E", "I", "O", "U", "A", "E")
  t = UCase$(Trim$(s & ""))
  For i = LBound(a) To UBound(a): t = Replace(t, a(i), b(i)): Next
  Do While InStr(t, "(") > 0
    p1 = InStr(t, "("): p2 = InStr(t, ")")
    If p2 > p1 Then t = Left$(t, p1 - 1) & " " & Mid$(t, p2 + 1) Else t = Replace(t, "(", " ")
  Loop
  For i = 1 To Len(t)
    ch = Mid$(t, i, 1)
    If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then o = o & ch Else o = o & " "
  Next
  Do While InStr(o, "  ") > 0: o = Replace(o, "  ", " "): Loop
  Normaliza = Trim$(o)
End Function

Function Pretty(P As String) As String
  Select Case P
    Case "CANAR": Pretty = "CAÑAR"
    Case Else: Pretty = P
  End Select
End Function

' Provincia tal como la espera el archivo EGR (sin Ñ): evita el problema CAÑAR/CANAR
Function ProvExport(ByVal s As String) As String
  If Normaliza(s) = "CANAR" Then ProvExport = "CANAR" Else ProvExport = Trim$(s)
End Function

' Levenshtein sin Application.Min (mucho más rápido, mismo resultado)
Function Lev(ByVal a As String, ByVal b As String) As Long
  Dim m As Long, n As Long, i As Long, j As Long, prev As Long, t As Long, c As Long, mn As Long
  m = Len(a): n = Len(b)
  If m = 0 Then Lev = n: Exit Function
  If n = 0 Then Lev = m: Exit Function
  Dim P() As Long: ReDim P(0 To n)
  For j = 0 To n: P(j) = j: Next
  For i = 1 To m
    prev = P(0): P(0) = i
    For j = 1 To n
      t = P(j)
      If Mid$(a, i, 1) = Mid$(b, j, 1) Then c = 0 Else c = 1
      mn = P(j) + 1
      If P(j - 1) + 1 < mn Then mn = P(j - 1) + 1
      If prev + c < mn Then mn = prev + c
      P(j) = mn: prev = t
    Next
  Next
  Lev = P(n)
End Function

' ---------- utilidades ----------
Function TXV(v As Variant) As String
  If IsError(v) Then TXV = "" Else TXV = Trim$(CStr(v & ""))
End Function

Private Function ColV(v As Variant, ByVal i As Long, ByVal c As Long) As String
  If c > 0 Then ColV = TXV(v(i, c))
End Function

Private Function HdrNorm(ByVal s As String) As String
  Dim i As Long, ch As String, o As String
  s = UCase$(s)
  For i = 1 To Len(s)
    ch = Mid$(s, i, 1)
    If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then o = o & ch Else o = o & " "
  Next
  Do While InStr(o, "  ") > 0: o = Replace(o, "  ", " "): Loop
  HdrNorm = " " & Trim$(o) & " "
End Function

' Busca una columna por su encabezado (por palabra inicial), entre c1 y c2
Private Function ColHdr(ws As Worksheet, ByVal hdrRow As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal clave As String) As Long
  Dim j As Long, k As String
  k = Trim$(HdrNorm(clave))
  For j = c1 To c2
    If InStr(HdrNorm(TX(ws.Cells(hdrRow, j))), " " & k) > 0 Then ColHdr = j: Exit Function
  Next
End Function

Private Sub CerrarForm(ByVal nombre As String)
  Dim i As Long
  For i = VBA.UserForms.Count - 1 To 0 Step -1
    If VBA.UserForms(i).Name = nombre Then Unload VBA.UserForms(i)
  Next
End Sub

' ---------- carga de maestros ----------
Sub CargarBases()
  Set mPC = Mapa()
  Set mRev = CreateObject("Scripting.Dictionary")
  Dim kk
  For Each kk In mPC.Keys
    If Not mRev.Exists(mPC(kk)) Then mRev(mPC(kk)) = kk
  Next
  Set mProvParr = CreateObject("Scripting.Dictionary")
  Set mParrList = CreateObject("Scripting.Dictionary")
  Set mCantonProv = CreateObject("Scripting.Dictionary")
  Set mParrGlobal = CreateObject("Scripting.Dictionary")
  Set mPcSet = CreateObject("Scripting.Dictionary")
  Set mCantonParr = CreateObject("Scripting.Dictionary")
  Set mParrOrig = CreateObject("Scripting.Dictionary")
  Set mTriple = CreateObject("Scripting.Dictionary")
  Set mCantonProvM = CreateObject("Scripting.Dictionary")
  Set mSiglas = CreateObject("Scripting.Dictionary")
  Set mCodeCant = CreateObject("Scripting.Dictionary")
  Set mCobInfo = CreateObject("Scripting.Dictionary")
  Set mPrincProv = CreateObject("Scripting.Dictionary")
  Dim ws As Worksheet, lr As Long, lc As Long, arr, i As Long
  Set ws = ThisWorkbook.Worksheets(H1)
  lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
  If lc < 4 Then lc = 4
  cGest = ColHdr(ws, 1, 1, lc, "GESTOR DE ENTREGAS")
  cGestSug = ColHdr(ws, 1, 1, lc, "GESTOR SUGERIDO")
  cTray = ColHdr(ws, 1, 1, lc, "TRAYECTO TRAMACO")
  If cTray = 0 And cGestSug > 0 And cGestSug < lc Then cTray = ColHdr(ws, 1, cGestSug + 1, lc, "TRAYECTO")
  If lr >= 2 Then
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lr, lc)).Value
    For i = 1 To UBound(arr, 1)
      AddBase TXV(arr(i, 1)), TXV(arr(i, 2)), TXV(arr(i, 3)), TXV(arr(i, 4))
      AddInfo TXV(arr(i, 1)), TXV(arr(i, 2)), TXV(arr(i, 3)), ColV(arr, i, cGest), ColV(arr, i, cGestSug), ColV(arr, i, cTray)
    Next
  End If
  CargarZonas
End Sub

Private Sub AddBase(vP, vC, vQ, Optional vSig)
  Dim P As String, Co As String, nc As String, Q As String, oQ As String, k As String
  P = Normaliza(vP): Co = Trim$(vC & ""): nc = Normaliza(Co): Q = Normaliza(vQ): oQ = Trim$(vQ & "")
  If Len(P) = 0 Or Len(Q) = 0 Then Exit Sub
  If Not mProvParr.Exists(P) Then mProvParr(P) = "|"
  If InStr(mProvParr(P), "|" & Q & "|") = 0 Then mProvParr(P) = mProvParr(P) & Q & "|"
  If Not mParrOrig.Exists(P & "|" & Q) Then mParrOrig(P & "|" & Q) = oQ
  If Not mParrList.Exists(P) Then Set mParrList(P) = New Collection
  mParrList(P).Add Array(Q, Co, oQ)
  If Len(nc) >= 4 And Not mCantonProv.Exists(nc) Then mCantonProv(nc) = P & Chr(1) & Co
  If Not mParrGlobal.Exists(Q) Then mParrGlobal(Q) = "|"
  If InStr(mParrGlobal(Q), "|" & P & Chr(1) & Co & "|") = 0 Then mParrGlobal(Q) = mParrGlobal(Q) & P & Chr(1) & Co & "|"
  k = P & "|" & Q
  If Not mPcSet.Exists(k) Then Set mPcSet(k) = New Collection
  mPcSet(k).Add Array(nc, Co)
  k = P & "|" & nc
  If Not mCantonParr.Exists(k) Then Set mCantonParr(k) = New Collection
  mCantonParr(k).Add Array(Q, Co, oQ)
  If Len(nc) > 0 Then
    If Not mTriple.Exists(P & "|" & nc & "|" & Q) Then mTriple(P & "|" & nc & "|" & Q) = Co
    If Not mCantonProvM.Exists(nc) Then mCantonProvM(nc) = "|"
    If InStr(mCantonProvM(nc), "|" & P & "|") = 0 Then mCantonProvM(nc) = mCantonProvM(nc) & P & "|"
  End If
  Dim sg As String: sg = Trim$(vSig & "")
  If Not mSiglas.Exists(P & "|" & nc & "|" & Q) Then mSiglas(P & "|" & nc & "|" & Q) = sg
  If Not mSiglas.Exists(P & "||" & Q) Then mSiglas(P & "||" & Q) = sg
  If Len(sg) > 0 And Len(nc) > 0 Then
    Dim cod As String: cod = CodigoSig(sg)
    If Len(cod) > 0 And Not mCodeCant.Exists(P & "|" & nc) Then mCodeCant(P & "|" & nc) = cod
  End If
End Sub

' Gestor, gestor sugerido y trayecto de cada destino + lista de ciudades principales/secundarias
Private Sub AddInfo(ByVal vP As String, ByVal vC As String, ByVal vQ As String, ByVal gest As String, ByVal gsug As String, ByVal tray As String)
  Dim P As String, nc As String, Q As String, k As String, tu As String, tipo As Long
  P = Normaliza(vP): nc = Normaliza(vC): Q = Normaliza(vQ)
  If Len(P) = 0 Or Len(Q) = 0 Then Exit Sub
  k = P & "|" & nc & "|" & Q
  If Not mCobInfo.Exists(k) Then mCobInfo(k) = Array(gest, gsug, tray, vC, vQ)
  If Not mCobInfo.Exists(P & "||" & Q) Then mCobInfo(P & "||" & Q) = Array(gest, gsug, tray, vC, vQ)
  tu = UCase$(Trim$(tray))
  If tu = "CP" Or InStr(tu, "PRINCIPAL") > 0 Then
    tipo = 1
  ElseIf tu = "CS" Or InStr(tu, "SECUNDARIA") > 0 Then
    tipo = 2
  End If
  If tipo > 0 Then
    If Not mPrincProv.Exists(P) Then Set mPrincProv(P) = New Collection
    mPrincProv(P).Add Array(nc, vC, Q, vQ, tipo)
  End If
End Sub

Private Sub CargarZonas()
  ' Hoja ZONAS PELIGROSAS: PROVINCIA | CIUDAD | PARROQUIA | ZONA PELIGROSA | SECTOR |
  ' PUNTO DE ATENCION MAS CERCANO | DIRECCIÓN DEL PUNTO | VALIDACION ZONA PELIGROSA (SI/NO)
  ' Las columnas se detectan por encabezado (el orden puede cambiar).
  Set mZonas = New Collection
  Dim wz As Worksheet
  On Error Resume Next
  Set wz = ThisWorkbook.Worksheets(HZ)
  On Error GoTo 0
  If wz Is Nothing Then Exit Sub
  If Application.WorksheetFunction.CountA(wz.UsedRange) = 0 Then Exit Sub
  Dim v, nr As Long, ncol As Long, i As Long, j As Long, hr As Long, lim As Long, r0 As Long
  Dim cP As Long, cC As Long, cQ As Long, cZ As Long, cS As Long, cPto As Long, cDir As Long, cVal As Long, h As String
  v = wz.UsedRange.Value
  If Not IsArray(v) Then Exit Sub
  nr = UBound(v, 1): ncol = UBound(v, 2)
  lim = nr: If lim > 15 Then lim = 15
  For i = 1 To lim
    cP = 0: cC = 0: cQ = 0: cZ = 0: cS = 0: cPto = 0: cDir = 0: cVal = 0
    For j = 1 To ncol
      h = HdrNorm(TXV(v(i, j)))
      If InStr(h, " VALIDACION") > 0 Then
        If cVal = 0 Then cVal = j
      ElseIf InStr(h, " PUNTO") > 0 And InStr(h, " DIRECC") > 0 Then
        If cDir = 0 Then cDir = j
      ElseIf InStr(h, " PUNTO") > 0 Then
        If cPto = 0 Then cPto = j
      ElseIf InStr(h, " PROV") > 0 Then
        If cP = 0 Then cP = j
      ElseIf InStr(h, " CANT") > 0 Or InStr(h, " CIUDAD") > 0 Then
        If cC = 0 Then cC = j
      ElseIf InStr(h, " PARR") > 0 Then
        If cQ = 0 Then cQ = j
      ElseIf InStr(h, " ZONA") > 0 Or InStr(h, " BARRIO") > 0 Then
        If cZ = 0 Then cZ = j
      ElseIf InStr(h, " SECTOR") > 0 Then
        If cS = 0 Then cS = j
      End If
    Next
    If cZ > 0 And (cP > 0 Or cC > 0 Or cQ > 0) Then hr = i: Exit For
  Next
  If hr = 0 Then Exit Sub
  For i = hr + 1 To nr
    If cVal > 0 Then
      If UCase$(ColV(v, i, cVal)) = "NO" Then GoTo sigZ
    End If
    If Len(ColV(v, i, cZ)) + Len(ColV(v, i, cS)) > 0 Then
      ' 0 prov | 1 ciudad | 2 parroquia | 3 zona | 4 sector | 5 punto TMC | 6 dirección punto | 7 texto
      mZonas.Add Array(Normaliza(ColV(v, i, cP)), Normaliza(ColV(v, i, cC)), Normaliza(ColV(v, i, cQ)), _
                       Normaliza(ColV(v, i, cZ)), Normaliza(ColV(v, i, cS)), ColV(v, i, cPto), ColV(v, i, cDir), _
                       Trim$(ColV(v, i, cZ) & IIf(Len(ColV(v, i, cS)) > 0, " / " & ColV(v, i, cS), "")))
    End If
sigZ:
  Next
End Sub

' Texto demasiado general para identificar un barrio (NORTE, SUR, CENTRO, un cantón, una provincia...)
Private Function EsGenerico(ByVal s As String, ByVal z1 As String, ByVal z2 As String) As Boolean
  EsGenerico = True
  If Len(s) < 4 Then Exit Function
  If s = z1 Or s = z2 Then Exit Function
  Select Case s
    Case "NORTE", "SUR", "ESTE", "OESTE", "CENTRO", "NOROESTE", "NORESTE", "SUROESTE", "SURESTE", "SUBURBIO", "CENTRO NORTE", "CENTRO SUR"
      Exit Function
  End Select
  If mCantonProv.Exists(s) Or mPC.Exists(s) Then Exit Function
  EsGenerico = False
End Function

Private Function QuitaPrefijo(ByVal s As String) As String
  Dim pf
  For Each pf In Array("BARRIO ", "B ", "CDLA ", "CIUDADELA ", "COOPERATIVA ", "COOP ", "SECTOR ", "URBANIZACION ", "URB ", "RECINTO ")
    If Left$(s, Len(pf)) = pf Then QuitaPrefijo = Trim$(Mid$(s, Len(pf) + 1)): Exit Function
  Next
  QuitaPrefijo = s
End Function

Private Function EnDireccion(ByVal s As String, ByVal a As String) As Boolean
  Dim s2 As String
  If Len(s) = 0 Then Exit Function
  If InStr(a, " " & s & " ") > 0 Then EnDireccion = True: Exit Function
  s2 = QuitaPrefijo(s)
  If s2 <> s And Len(s2) >= 5 Then EnDireccion = (InStr(a, " " & s2 & " ") > 0)
End Function

' Devuelve "ZONA PELIGROSA: ..." si el pedido cae en una zona de la hoja; punto = oficina TMC más cercana
' Si solo coincide una parroquia con sector direccional (NORTE, SUR...) devuelve "POSIBLE ZONA PELIGROSA: ..."
Public Function ZonaRiesgo(ByVal P As String, ByVal nc As String, ByVal Q As String, ByVal a As String, Optional ByRef punto As String) As String
  Dim z, ok As Boolean, hit As Boolean, zEsp As Boolean, sEsp As Boolean, blanda As String, ptoB As String
  If mZonas Is Nothing Then CargarBases
  punto = ""
  If Len(a) = 0 Then a = " "
  For Each z In mZonas
    ok = True
    If Len(z(0)) > 0 Then
      If P <> z(0) And Left$(P, Len(z(0)) + 1) <> z(0) & " " Then ok = False
    End If
    If ok And Len(z(1)) > 0 Then
      If z(1) <> nc And InStr(a, " " & z(1) & " ") = 0 Then ok = False
    End If
    If ok Then
      zEsp = Not EsGenerico(z(3), z(1), z(2))
      sEsp = Not EsGenerico(z(4), z(1), z(2))
      hit = False
      If zEsp Then hit = EnDireccion(z(3), a)
      If Not hit And sEsp Then hit = EnDireccion(z(4), a)
      If hit Then
        punto = Trim$(z(5))
        ZonaRiesgo = "ZONA PELIGROSA: " & z(7) & IIf(Len(punto) > 0, " | Punto: " & punto & IIf(Len(z(6)) > 0, " (" & z(6) & ")", ""), "")
        Exit Function
      End If
      ' la zona es la parroquia (ej. PUERTO LOPEZ, JOSE LUIS TAMAYO)
      If Not zEsp And Not sEsp And Len(z(3)) > 0 And z(3) = z(2) And Q = z(2) Then
        If Len(z(4)) = 0 Then
          punto = Trim$(z(5))
          ZonaRiesgo = "ZONA PELIGROSA: parroquia " & z(7) & IIf(Len(punto) > 0, " | Punto: " & punto & IIf(Len(z(6)) > 0, " (" & z(6) & ")", ""), "")
          Exit Function
        ElseIf Len(blanda) = 0 Then
          ' solo un sector de la parroquia (SUR, NORTE...): aviso, no C.O.D automático
          ptoB = Trim$(z(5))
          blanda = "POSIBLE ZONA PELIGROSA: sector " & z(4) & " de " & z(2) & IIf(Len(ptoB) > 0, " | Punto: " & ptoB, "")
        End If
      End If
    End If
  Next
  If Len(blanda) > 0 Then ZonaRiesgo = blanda: punto = ptoB
End Function

' Códigos de TRAYECTO(TRAMACO) de COBERTURA
Public Function TrayectoTexto(ByVal cod As String) As String
  Select Case UCase$(Trim$(cod))
    Case "CP": TrayectoTexto = "CIUDAD PRINCIPAL"
    Case "CS": TrayectoTexto = "CIUDAD SECUNDARIA"
    Case "TE": TrayectoTexto = "TRAYECTO ESPECIAL"
    Case "TD": TrayectoTexto = "TRAYECTO DIFERENCIADO"
    Case Else: TrayectoTexto = Trim$(cod)
  End Select
End Function

' Valores de las columnas V:Z (gestor, destino, trayecto, tipo de entrega, zona)
Public Function DatosExt(ByVal P As String, ByVal nc As String, ByVal Q As String, ByVal a As String) As Variant
  Dim dest As String, tray As String, gest As String, zona As String, tipo As String, pto As String
  gest = InfoGestor(P, nc, Q, dest, tray)
  zona = ZonaRiesgo(P, nc, Q, a, pto)
  tipo = "NORMAL"
  If Left$(zona, 8) = "POSIBLE " Then
    tipo = "VERIFICAR SECTOR (posible zona peligrosa)"
  ElseIf Len(zona) > 0 Then
    If dest = "PRO" Then
      tipo = "C.O.D - RETIRO OFICINA " & IIf(Len(pto) > 0, pto, "TMC")
    Else
      tipo = "CONFIRMAR ENTREGA (ZONA PELIGROSA)"
    End If
  End If
  Dim inf, gq As String, gr As String
  If mCobInfo.Exists(P & "|" & nc & "|" & Q) Then
    inf = mCobInfo(P & "|" & nc & "|" & Q)
  ElseIf mCobInfo.Exists(P & "||" & Q) Then
    inf = mCobInfo(P & "||" & Q)
  End If
  If IsArray(inf) Then gq = inf(0): gr = inf(1)
  DatosExt = Array(gest, dest, tray, tipo, zona, gq, gr)
End Function

' ---------- resolución de cantón / cobertura ----------
Function ResolveCanton(P As String, Q As String, aC As Collection) As String
  Dim setc As Collection, it, ac2, u As String, same As Boolean
  If mPcSet.Exists(P & "|" & Q) Then Set setc = mPcSet(P & "|" & Q) Else Set setc = New Collection
  If Not aC Is Nothing Then
    For Each ac2 In aC
      For Each it In setc
        If it(0) = ac2 Then ResolveCanton = it(1): Exit Function
      Next
    Next
  End If
  If setc.Count > 0 Then
    same = True: u = setc(1)(1)
    For Each it In setc
      If it(1) <> u Then same = False
    Next
    If same Then ResolveCanton = u: Exit Function
  End If
  If Not aC Is Nothing Then
    If aC.Count > 0 Then
      If mCantonProv.Exists(aC(1)) Then ResolveCanton = Split(mCantonProv(aC(1)), Chr(1))(1): Exit Function
    End If
  End If
  If setc.Count > 0 Then ResolveCanton = setc(1)(1)
End Function

' Cantón que escribió el cliente, si es válido para esa parroquia (por nombre o por sufijo "(TUL)")
Private Function CantonCliente(ByVal P As String, ByVal nG As String, ByVal Q As String, ByVal rawParr As String) As String
  Dim suf As String, it
  If Len(nG) > 0 Then
    If mTriple.Exists(P & "|" & nG & "|" & Q) Then CantonCliente = mTriple(P & "|" & nG & "|" & Q): Exit Function
  End If
  suf = CodigoSig(rawParr)
  If Len(suf) > 0 And mPcSet.Exists(P & "|" & Q) Then
    For Each it In mPcSet(P & "|" & Q)
      If mSiglas.Exists(P & "|" & it(0) & "|" & Q) Then
        If CodigoSig(mSiglas(P & "|" & it(0) & "|" & Q)) = suf Then CantonCliente = it(1): Exit Function
      End If
    Next
  End If
End Function

Private Function NumCantones(ByVal P As String, ByVal Q As String) As Long
  Dim it, u As Object
  If Not mPcSet.Exists(P & "|" & Q) Then Exit Function
  Set u = CreateObject("Scripting.Dictionary")
  For Each it In mPcSet(P & "|" & Q)
    u(it(0)) = 1
  Next
  NumCantones = u.Count
End Function

Private Function OrigParr(ByVal k As String, ByVal defecto As String) As String
  OrigParr = defecto
  If mCobInfo.Exists(k) Then OrigParr = mCobInfo(k)(4)
End Function

' Sugerencia de cobertura cercana. Devuelve Array(cantón, parroquia) o Empty.
' Orden: 0) misma parroquia mal escrita  1) cabecera cantonal (parroquia = cantón)
'        2) ciudad principal / secundaria del mismo cantón
'        3) ciudad principal / secundaria mencionada en la dirección
'        4) ciudad principal / secundaria de la provincia   5) parroquia más parecida del cantón
Public Function SugerirCercana(ByVal P As String, ByVal nc As String, ByVal Q As String, ByVal a As String, ByRef motivo As String) As Variant
  Dim it, best, bd As Long, dd As Long, lim As Long, k As String, tipo As Long, col As Collection
  If mTriple Is Nothing Then CargarBases
  SugerirCercana = Empty: motivo = ""
  If Len(a) = 0 Then a = " "
  If Len(nc) > 0 And Len(Q) > 0 And mCantonParr.Exists(P & "|" & nc) Then
    bd = 999
    For Each it In mCantonParr(P & "|" & nc)
      dd = Lev(Q, it(0))
      If dd < bd Then bd = dd: best = it
    Next
    If Len(Q) <= 5 Then lim = 1 Else lim = 2
    If bd <= lim Then SugerirCercana = Array(best(1), best(2)): motivo = "Parroquia similar (error de escritura)": Exit Function
  End If
  If Len(nc) > 0 Then
    k = P & "|" & nc & "|" & nc
    If mTriple.Exists(k) Then SugerirCercana = Array(mTriple(k), OrigParr(k, Pretty(nc))): motivo = "Parroquia principal del cantón (cabecera)": Exit Function
  End If
  If mPrincProv.Exists(P) Then
    For tipo = 1 To 2
      For Each it In mPrincProv(P)
        If it(0) = nc And it(4) = tipo Then
          SugerirCercana = Array(it(1), it(3))
          If tipo = 1 Then motivo = "Ciudad principal del cantón" Else motivo = "Ciudad secundaria del cantón"
          Exit Function
        End If
      Next
    Next
    For tipo = 1 To 2
      For Each it In mPrincProv(P)
        If it(4) = tipo And (InStr(a, " " & it(0) & " ") > 0 Or InStr(a, " " & it(2) & " ") > 0) Then
          SugerirCercana = Array(it(1), it(3))
          If tipo = 1 Then motivo = "Ciudad principal cercana (dirección)" Else motivo = "Ciudad secundaria cercana (dirección)"
          Exit Function
        End If
      Next
    Next
    For tipo = 1 To 2
      For Each it In mPrincProv(P)
        If it(4) = tipo And it(0) = it(2) Then
          SugerirCercana = Array(it(1), it(3))
          If tipo = 1 Then motivo = "Ciudad principal de la provincia" Else motivo = "Ciudad secundaria de la provincia"
          Exit Function
        End If
      Next
      For Each it In mPrincProv(P)
        If it(4) = tipo Then
          SugerirCercana = Array(it(1), it(3))
          If tipo = 1 Then motivo = "Ciudad principal de la provincia" Else motivo = "Ciudad secundaria de la provincia"
          Exit Function
        End If
      Next
    Next
  End If
  ' 5) capital provincial (provincias sin CP/CS en COBERTURA)
  Dim cap As String: cap = CapitalProv(P)
  If Len(cap) > 0 And cap <> nc Then
    k = P & "|" & cap & "|" & cap
    If mTriple.Exists(k) Then SugerirCercana = Array(mTriple(k), OrigParr(k, Pretty(cap))): motivo = "Capital provincial": Exit Function
    If mCantonParr.Exists(P & "|" & cap) Then
      Set col = mCantonParr(P & "|" & cap)
      If col.Count > 0 Then SugerirCercana = Array(col(1)(1), col(1)(2)): motivo = "Capital provincial": Exit Function
    End If
  End If
  ' 6) último recurso: parroquia más parecida del cantón
  If Len(nc) > 0 And mCantonParr.Exists(P & "|" & nc) Then
    Set col = mCantonParr(P & "|" & nc)
    If col.Count > 0 Then SugerirCercana = Array(col(1)(1), FirstFuzzy(col, Q)): motivo = "Parroquia más parecida del cantón"
  End If
End Function

' Cantón capital de cada provincia (nombre normalizado como en COBERTURA)
Private Function CapitalProv(ByVal P As String) As String
  Select Case P
    Case "AZUAY": CapitalProv = "CUENCA"
    Case "BOLIVAR": CapitalProv = "GUARANDA"
    Case "CANAR": CapitalProv = "AZOGUES"
    Case "CARCHI": CapitalProv = "TULCAN"
    Case "COTOPAXI": CapitalProv = "LATACUNGA"
    Case "CHIMBORAZO": CapitalProv = "RIOBAMBA"
    Case "EL ORO": CapitalProv = "MACHALA"
    Case "ESMERALDAS": CapitalProv = "ESMERALDAS"
    Case "GUAYAS": CapitalProv = "GUAYAQUIL"
    Case "IMBABURA": CapitalProv = "IBARRA"
    Case "LOJA": CapitalProv = "LOJA"
    Case "LOS RIOS": CapitalProv = "BABAHOYO"
    Case "MANABI": CapitalProv = "PORTOVIEJO"
    Case "MORONA SANTIAGO": CapitalProv = "MORONA"
    Case "NAPO": CapitalProv = "TENA"
    Case "PASTAZA": CapitalProv = "PASTAZA"
    Case "PICHINCHA": CapitalProv = "QUITO"
    Case "TUNGURAHUA": CapitalProv = "AMBATO"
    Case "ZAMORA CHINCHIPE": CapitalProv = "ZAMORA"
    Case "GALAPAGOS": CapitalProv = "SAN CRISTOBAL"
    Case "SUCUMBIOS": CapitalProv = "LAGO AGRIO"
    Case "ORELLANA": CapitalProv = "ORELLANA"
    Case "SANTO DOMINGO DE LOS TSACHILAS": CapitalProv = "SANTO DOMINGO"
    Case "SANTA ELENA": CapitalProv = "SANTA ELENA"
  End Select
End Function

' Regla de gestor: QUITO / GUAYAQUIL -> ITSANET o LAAR (según cobertura); GALÁPAGOS -> GPS; resto -> PRO / TRAMACO
Public Function InfoGestor(ByVal P As String, ByVal nc As String, ByVal Q As String, ByRef dest As String, ByRef tray As String) As String
  Dim inf, gest As String, gsug As String, k As String
  If mCobInfo Is Nothing Then CargarBases
  k = P & "|" & nc & "|" & Q
  If mCobInfo.Exists(k) Then
    inf = mCobInfo(k)
  ElseIf mCobInfo.Exists(P & "||" & Q) Then
    inf = mCobInfo(P & "||" & Q)
  End If
  tray = ""
  If IsArray(inf) Then gest = UCase$(inf(0)): gsug = UCase$(inf(1)): tray = TrayectoTexto(inf(2))
  If P = "PICHINCHA" And nc = "QUITO" Then
    dest = "UIO"
  ElseIf P = "GUAYAS" And nc = "GUAYAQUIL" Then
    dest = "GYE"
  ElseIf P = "GALAPAGOS" Then
    dest = "GPS"
  Else
    dest = "PRO"
  End If
  Select Case dest
    Case "UIO", "GYE"
      If InStr(gest, "ITSANET") > 0 Then
        InfoGestor = "ITSANET"
      ElseIf InStr(gest, "LAAR") > 0 Then
        InfoGestor = "LAAR COURIER"
      Else
        InfoGestor = "ITSANET"
      End If
    Case "GPS"
      If Len(gest) > 0 Then InfoGestor = gest Else InfoGestor = "TRAMACO"
    Case Else
      If Len(gsug) > 0 Then InfoGestor = gsug Else InfoGestor = "TRAMACO"
  End Select
End Function



' Recalcula V:Z de una fila a partir de su propuesta N:P (lo usa frmValidar)
Public Sub EscribirExtFila(ws As Worksheet, ByVal r As Long)
  Dim ex, j As Long, a As String
  a = " " & Normaliza(TX(ws.Cells(r, C_DIR))) & " "
  ex = DatosExt(Normaliza(TX(ws.Cells(r, C_NN))), Normaliza(TX(ws.Cells(r, C_NN + 1))), Normaliza(TX(ws.Cells(r, C_NN + 2))), a)
  For j = 0 To N_EXTC - 1: ws.Cells(r, C_EXT + j).Value = ex(j): Next
End Sub

' ---------- accesos para los formularios ----------
Public Function ExisteTriada(ByVal P As String, ByVal nc As String, ByVal Q As String) As Boolean
  If mTriple Is Nothing Then CargarBases
  ExisteTriada = mTriple.Exists(P & "|" & nc & "|" & Q)
End Function

Public Function ListaCantones(ByVal P As String) As Collection
  Dim c As New Collection, k, pre As String, col As Collection
  If mCantonParr Is Nothing Then CargarBases
  pre = P & "|"
  For Each k In mCantonParr.Keys
    If Left$(k, Len(pre)) = pre And Len(k) > Len(pre) Then
      Set col = mCantonParr(k)
      If col.Count > 0 Then c.Add col(1)(1)
    End If
  Next
  Set ListaCantones = c
End Function

Public Function ListaParroquias(ByVal P As String, ByVal nc As String) As Collection
  Dim c As New Collection, it, u As Object
  If mCantonParr Is Nothing Then CargarBases
  Set u = CreateObject("Scripting.Dictionary")
  If mCantonParr.Exists(P & "|" & nc) Then
    For Each it In mCantonParr(P & "|" & nc)
      If Not u.Exists(it(2)) Then u(it(2)) = 1: c.Add it(2)
    Next
  End If
  Set ListaParroquias = c
End Function

Function ProvDeCP(a As String) As String
  Static re As Object
  If re Is Nothing Then
    Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.Pattern = "(^|\D)(\d{6})(\D|$)"
  End If
  Dim m, mm, pre As String
  Set m = re.Execute(a)
  For Each mm In m
    pre = Left$(mm.SubMatches(1), 2)
    If mRev.Exists(pre) Then ProvDeCP = mRev(pre): Exit Function
  Next
End Function

Function FuzzyStr(pool As Collection, Q As String, topN As Long) As String
  Dim n As Long, used() As Boolean, x As Long, j As Long, bi As Long, bd As Long, dd As Long, s As String, lim As Long
  n = pool.Count: If n = 0 Then Exit Function
  ReDim used(1 To n)
  lim = topN: If n < lim Then lim = n
  For x = 1 To lim
    bd = 999999: bi = 0
    For j = 1 To n
      If Not used(j) Then
        dd = Lev(Q, pool(j)(0))
        If dd < bd Then bd = dd: bi = j
      End If
    Next
    If bi > 0 Then used(bi) = True: s = s & IIf(s = "", "", " | ") & pool(bi)(2) & " [" & pool(bi)(1) & "]"
  Next
  FuzzyStr = s
End Function

Function FirstFuzzy(pool As Collection, Q As String) As String
  Dim n As Long, j As Long, bi As Long, bd As Long, dd As Long
  n = pool.Count: If n = 0 Then Exit Function
  bd = 999999
  For j = 1 To n
    dd = Lev(Q, pool(j)(0))
    If dd < bd Then bd = dd: bi = j
  Next
  If bi > 0 Then FirstFuzzy = pool(bi)(2)
End Function


' =====================================================================================
'  VALIDACIÓN
' =====================================================================================
Sub ValidarPedidos()
  On Error GoTo cleanup
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
  CargarBases
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim okC As Long, revC As Long
  If lr < 2 Then GoTo fin
  Dim d, out, out2, ex, i As Long, j As Long
  d = ws.Range(ws.Cells(1, 1), ws.Cells(lr, C_PARR)).Value
  If SOLO_FILTRADO Then
    out = ws.Range(ws.Cells(2, C_NN), ws.Cells(lr, C_NN + 7)).Value
    out2 = ws.Range(ws.Cells(2, C_EXT), ws.Cells(lr, C_EXT + N_EXT - 1)).Value
  Else
    ReDim out(1 To lr - 1, 1 To 8)
    ReDim out2(1 To lr - 1, 1 To N_EXT)
  End If
  Dim nP As String, nQ As String, nG As String, a As String, votes As Object, pr, cn, prov As String
  Dim bvote As Long, vk, evid As String, aCinP As Collection, ci2, pg As String
  Dim parrO As String, parrFin As String, estParr As String, sug As String, cantonFin As String
  Dim scopeNc As String, pool As Collection, cand, it, ncF As String
  Dim alert As String, est As String, accion As String, cantonConf As Boolean, esOK As Boolean
  Dim locs, uDup As Object, z, pparts, keyd As String
  Dim usedCola As Boolean, cpStr As String, qAlert As String, ncCola As String
  Dim rawParr As String, ncS As String, sc, motivoS As String, cantonSug As String, ambig As Boolean
  Dim tipoCorr As String, origParts
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo sig
    nP = Normaliza(TXV(d(i, C_PROV))): nG = Normaliza(TXV(d(i, C_CANT))): nQ = Normaliza(TXV(d(i, C_PARR)))
    rawParr = TXV(d(i, C_PARR))
    If Len(nP) = 0 And Len(nQ) = 0 And Len(Normaliza(TXV(d(i, C_DIR)))) = 0 Then GoTo sig
    a = " " & Normaliza(TXV(d(i, C_DIR))) & " ": evid = "": sug = ""
    cantonSug = "": ambig = False: motivoS = ""
    Set votes = CreateObject("Scripting.Dictionary")
    ' --- 1) parser de la cola de la dirección (… PARROQUIA CP PROVINCIA) ---
    usedCola = False
    If ParseColaDireccion(a, prov, parrO, parrFin, cpStr) Then
      Set aCinP = New Collection
      cantonFin = CantonCliente(prov, nG, parrFin, rawParr)
      If Len(cantonFin) > 0 Then
        aCinP.Add Normaliza(cantonFin)
      Else
        cantonFin = ResolveCanton(prov, parrFin, aCinP)
        ambig = (NumCantones(prov, parrFin) > 1)
      End If
      ncCola = Normaliza(cantonFin)
      evid = "cp:" & cpStr & " prov:" & Pretty(prov)
      If mTriple.Exists(prov & "|" & ncCola & "|" & parrFin) Then
        estParr = "OK"
      ElseIf mProvParr.Exists(prov) And InStr(mProvParr(prov), "|" & parrFin & "|") > 0 Then
        estParr = "OK"
      Else
        estParr = "COLA"
      End If
      If ambig And estParr = "OK" Then estParr = "AMBIG"
      If Len(nP) > 0 And nP <> prov Then evid = evid & " (F=" & Pretty(nP) & "?)"
      usedCola = True
      GoTo tras_resolucion
    End If
    ' --- 2) votación de provincia (lógica original) ---
    If mPC.Exists(nP) Then votes(nP) = votes(nP) + 1
    If Len(nG) > 0 Then
      If mCantonProv.Exists(nG) Then pg = Split(mCantonProv(nG), Chr(1))(0): votes(pg) = votes(pg) + 1
    End If
    For Each pr In mPC.Keys
      If InStr(a, " " & pr & " ") > 0 Then votes(pr) = votes(pr) + 3: If InStr(evid, "prov") = 0 Then evid = evid & "prov:" & Pretty(CStr(pr)) & " "
    Next
    If InStr(a, " SANTO DOMINGO ") > 0 Then votes("SANTO DOMINGO DE LOS TSACHILAS") = votes("SANTO DOMINGO DE LOS TSACHILAS") + 3
    For Each cn In mCantonProv.Keys
      If InStr(a, " " & cn & " ") > 0 Then pg = Split(mCantonProv(cn), Chr(1))(0): votes(pg) = votes(pg) + 2
    Next
    pg = ProvDeCP(a): If Len(pg) > 0 Then votes(pg) = votes(pg) + 2
    prov = "": If mPC.Exists(nP) Then prov = nP
    bvote = -1
    For Each vk In votes.Keys
      If votes(vk) > bvote Then bvote = votes(vk): prov = vk
    Next
    If prov = "" And mPC.Exists(nP) Then prov = nP
    If prov = "" Then prov = "PICHINCHA"
    Set aCinP = New Collection
    If Len(nG) > 0 Then
      If mCantonProv.Exists(nG) Then
        If Split(mCantonProv(nG), Chr(1))(0) = prov Then aCinP.Add nG
      End If
    End If
    For Each cn In mCantonProv.Keys
      If Split(mCantonProv(cn), Chr(1))(0) = prov And InStr(a, " " & cn & " ") > 0 Then
        If Not (Len(nG) > 0 And CStr(cn) = nG) Then aCinP.Add cn
      End If
    Next
    If mProvParr.Exists(prov) And InStr(mProvParr(prov), "|" & nQ & "|") > 0 Then
      parrFin = nQ
      If mParrOrig.Exists(prov & "|" & nQ) Then parrO = mParrOrig(prov & "|" & nQ) Else parrO = rawParr
      estParr = "OK": sug = ""
      cantonFin = CantonCliente(prov, nG, nQ, rawParr)
      If Len(cantonFin) > 0 Then
        aCinP.Add Normaliza(cantonFin)
      Else
        cantonFin = ResolveCanton(prov, nQ, aCinP)
      End If
    Else
      scopeNc = "": If aCinP.Count > 0 Then scopeNc = aCinP(1)
      If Len(scopeNc) > 0 And mCantonParr.Exists(prov & "|" & scopeNc) Then
        Set pool = mCantonParr(prov & "|" & scopeNc)
      ElseIf mParrList.Exists(prov) Then
        Set pool = mParrList(prov)
      Else
        Set pool = New Collection
      End If
      cand = Empty
      For Each it In pool
        If it(0) = nQ And Len(nQ) >= 3 And InStr(a, " " & it(0) & " ") > 0 Then cand = it: Exit For
      Next
      If IsEmpty(cand) Then
        For Each it In pool
          If Len(it(0)) >= 4 And Not mPC.Exists(it(0)) And it(0) <> scopeNc And InStr(a, " " & it(0) & " ") > 0 Then cand = it: Exit For
        Next
      End If
      If IsEmpty(cand) Then
        For Each it In pool
          If Len(it(0)) >= 4 And Not mPC.Exists(it(0)) And InStr(a, " " & it(0) & " ") > 0 Then cand = it: Exit For
        Next
      End If
      If Not IsEmpty(cand) Then
        parrFin = cand(0): parrO = cand(2): estParr = "DIR": sug = ""
      Else
        ' --- NUEVO: fuera de cobertura -> cobertura cercana (cabecera / ciudad principal o secundaria) ---
        ncS = scopeNc: If Len(ncS) = 0 Then ncS = nG
        sug = FuzzyStr(pool, nQ, 3)
        sc = SugerirCercana(prov, ncS, nQ, a, motivoS)
        If IsArray(sc) Then
          parrO = sc(1): parrFin = Normaliza(parrO): cantonSug = sc(0)
          sug = parrO & " [" & sc(0) & "] (" & motivoS & ")" & IIf(Len(sug) > 0, " | " & sug, "")
        Else
          parrO = FirstFuzzy(pool, nQ): parrFin = Normaliza(parrO)
        End If
        If mParrGlobal.Exists(nQ) Then
          estParr = "OTRA_PROV"
        ElseIf motivoS = "Parroquia similar (error de escritura)" Or Not IsArray(sc) Then
          estParr = "SUGERIR"
        Else
          estParr = "FUERA_COB"
        End If
      End If
      cantonFin = ResolveCanton(prov, parrFin, aCinP)
      If Len(cantonSug) > 0 Then cantonFin = cantonSug
    End If
tras_resolucion:
    ' Regla Quito / DMQ (se mantiene): parroquia DMQ y sigla CALDERON (QUI)
    If prov = "PICHINCHA" And estParr <> "OK" Then
      If InStr(a, " QUITO ") > 0 Or InStr(a, " DMQ ") > 0 Or InStr(a, " DISTRITO METROPOLITANO ") > 0 Then
        cantonFin = "QUITO"
        If estParr = "OTRA_PROV" Or estParr = "FUERA_COB" Or (estParr = "SUGERIR" And Len(cantonSug) = 0) Then
          parrO = "DISTRITO METROPOLITANO DE QUITO": parrFin = "QUITO": estParr = "QUITO"
        End If
      End If
    End If
    If Len(parrO) = 0 Then parrO = rawParr
    If Len(parrO) = 0 Then parrO = Pretty(prov)
    If Len(cantonFin) = 0 Then cantonFin = parrO
    ncF = Normaliza(cantonFin)
    cantonConf = False
    For Each ci2 In aCinP
      If CStr(ci2) = ncF Then cantonConf = True: Exit For
    Next
    qAlert = IIf(usedCola, parrFin, nQ)
    alert = ""
    If mParrGlobal.Exists(qAlert) Then
      Set uDup = CreateObject("Scripting.Dictionary")
      locs = Split(mParrGlobal(qAlert), "|")
      For Each z In locs
        If InStr(CStr(z), Chr(1)) > 0 Then
          pparts = Split(CStr(z), Chr(1))
          If UBound(pparts) >= 1 Then
            keyd = Pretty(CStr(pparts(0))) & "/" & Normaliza(CStr(pparts(1)))
            If Not uDup.Exists(keyd) Then uDup(keyd) = Pretty(CStr(pparts(0))) & "/" & CStr(pparts(1))
          End If
        End If
      Next
      If uDup.Count > 1 Then alert = "Parroquia repetida en: " & Left$(Join(uDup.Items, " ; "), 90)
    End If
    ex = DatosExt(prov, ncF, parrFin, a)
    esOK = (estParr = "OK")
    If esOK And Not mPC.Exists(prov) Then esOK = False
    If esOK And Not (nP = "" Or nP = prov) And Not usedCola Then esOK = False
    If esOK And (Len(alert) > 0 And Not cantonConf) And Not usedCola Then esOK = False
    If esOK And ZONA_PELIGROSA_A_REVISAR And Len(ex(4)) > 0 Then esOK = False
    If esOK Then
      est = "OK": accion = "Sin acción": okC = okC + 1
    Else
      est = "REVISAR": revC = revC + 1
      Select Case True
        Case estParr = "COLA": accion = "Confirmar parroquia (dirección/CP)"
        Case estParr = "AMBIG": accion = "Verificar cantón (parroquia en varios cantones)"
        Case estParr = "OTRA_PROV": accion = "Parroquia de otra provincia"
        Case estParr = "FUERA_COB": accion = "Fuera de cobertura: confirmar sugerida"
        Case estParr = "SUGERIR": accion = "Elegir parroquia sugerida"
        Case estParr = "DIR": accion = "Parroquia de la dirección"
        Case estParr = "QUITO": accion = "Confirmar Quito/DMQ"
        Case (Len(alert) > 0 And Not cantonConf): accion = "Verificar duplicado"
        Case Len(ex(4)) > 0: accion = "Zona peligrosa: confirmar C.O.D / retiro TMC"
        Case Else: accion = "Revisar"
      End Select
    End If
    out(i - 1, 1) = Pretty(prov): out(i - 1, 2) = cantonFin: out(i - 1, 3) = parrO
    out(i - 1, 4) = est: out(i - 1, 5) = Trim$(evid): out(i - 1, 6) = sug: out(i - 1, 7) = accion: out(i - 1, 8) = alert
    For j = 0 To N_EXTC - 1: out2(i - 1, j + 1) = ex(j): Next
    ' AC: dato original del cliente (se conserva aunque se aplique la corrección)
    If Len(TXV(out2(i - 1, N_EXTC + 1))) = 0 Then out2(i - 1, N_EXTC + 1) = TXV(d(i, C_PROV)) & " / " & TXV(d(i, C_CANT)) & " / " & rawParr
    ' AD: tipo de corrección
    Select Case estParr
      Case "FUERA_COB": tipoCorr = "COBERTURA CERCANA: " & motivoS
      Case "SUGERIR"
        If Len(motivoS) > 0 Then tipoCorr = "CORRECCION DE ESCRITURA" Else tipoCorr = "SUGERENCIA POR SIMILITUD"
      Case "QUITO": tipoCorr = "DMQ -> CALDERON (QUI)"
      Case "AMBIG": tipoCorr = "VERIFICAR CANTON"
      Case "OTRA_PROV": tipoCorr = "PARROQUIA DE OTRA PROVINCIA"
      Case "DIR": tipoCorr = "PARROQUIA TOMADA DE LA DIRECCION"
      Case "COLA": tipoCorr = "CONFIRMAR PARROQUIA"
      Case Else
        origParts = Split(TXV(out2(i - 1, N_EXTC + 1)), " / ")
        tipoCorr = "CORREGIDO SISTEMA"
        If UBound(origParts) >= 2 Then
          If Normaliza(CStr(origParts(0))) = prov And Normaliza(CStr(origParts(1))) = ncF And Normaliza(CStr(origParts(2))) = Normaliza(parrO) Then tipoCorr = "SIN CAMBIO"
        End If
    End Select
    out2(i - 1, N_EXTC + 2) = tipoCorr
    ws.Cells(i, C_SIG).Value = SiglaFinal(prov, ncF, parrFin, parrO)
sig:
  Next
  ws.Range(ws.Cells(2, C_NN), ws.Cells(lr, C_NN + 7)).Value = out
  ws.Range(ws.Cells(2, C_EXT), ws.Cells(lr, C_EXT + N_EXT - 1)).Value = out2
  FormatoPanel ws, lr
  ColorearCambios ws, lr
fin:
  Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
  If revC > 0 Then
    If MsgBox("Validación lista." & vbCrLf & "OK: " & okC & "    Por revisar: " & revC & vbCrLf & vbCrLf & _
              "¿Revisar ahora los " & revC & " casos pendientes en el validador?", _
              vbYesNo + vbQuestion, "Validación HYCITE") = vbYes Then AbrirValidadorRevisar
  Else
    MsgBox "Validación lista." & vbCrLf & "OK: " & okC & "    Por revisar: 0", vbInformation, "Validación HYCITE"
  End If
  Exit Sub
cleanup:
  Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True: Application.EnableEvents = True
  MsgBox "Se detuvo por un error (pantalla liberada): " & Err.Description, vbExclamation
End Sub

Sub LimpiarValidacion()
  If MsgBox("Se limpiará el panel de validación (N:AD), la columna L y sus colores." & vbCrLf & _
     "La data del cliente (A:K) NO se toca. ¿Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  LimpiarPanel
  MsgBox "Panel y siglas limpiados. Listo para la nueva data del día.", vbInformation
End Sub

Private Sub LimpiarPanel()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  If lr < 2 Then lr = 2
  With ws.Range(ws.Cells(2, C_SIG), ws.Cells(lr, C_SIG)): .ClearContents: .Interior.ColorIndex = xlNone: End With
  With ws.Range(ws.Cells(2, C_NN), ws.Cells(lr, C_EXT + N_EXT - 1)): .ClearContents: .Interior.ColorIndex = xlNone: End With
End Sub

' Paso 0: refresca las consultas (DEPOT y TMS) y ofrece limpiar la validación anterior
Sub ActualizarDatosDepot()
  Dim ws As Worksheet, lo As ListObject, n As Long, lr As Long
  If MsgBox("Se actualizarán los pedidos desde DEPOT y la lista de TMS (consultas ODBC)." & vbCrLf & _
            "Tarda unos segundos. ¿Continuar?", vbYesNo + vbQuestion, "Paso 0 - Actualizar datos") <> vbYes Then Exit Sub
  On Error GoTo fallo
  Application.Cursor = xlWait
  For Each ws In ThisWorkbook.Worksheets
    For Each lo In ws.ListObjects
      If lo.SourceType = xlSrcQuery Then lo.QueryTable.Refresh BackgroundQuery:=False: n = n + 1
    Next
  Next
  Application.Cursor = xlDefault
  On Error GoTo 0
  Set ws = ThisWorkbook.Worksheets(HD)
  lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  If MsgBox(n & " consultas actualizadas. Pedidos en DEPOT: " & (lr - 1) & vbCrLf & vbCrLf & _
            "Los resultados de la validación anterior pueden no corresponder a los pedidos nuevos." & vbCrLf & _
            "¿Limpiar la validación ahora? (recomendado)", vbYesNo + vbQuestion, "Paso 0 - Actualizar datos") = vbYes Then LimpiarPanel
  Exit Sub
fallo:
  Application.Cursor = xlDefault
  MsgBox "No se pudo actualizar: " & Err.Description & vbCrLf & "Revisa la conexión ODBC (DEPOTUIO / TMS1) y vuelve a intentar.", vbExclamation
End Sub

Sub FormatoPanel(ws As Worksheet, lr As Long)
  Dim h, j As Long, r As Range
  h = Array("PROVINCIA_PROP", "CANTON/CIUDAD_PROP", "PARROQUIA_PROP", "ESTADO", "EVIDENCIA_DIRECCION", "SUGERENCIAS", "ACCION_OPERARIO", "ALERTAS/DUPLICADOS")
  For j = 0 To 7
    With ws.Cells(1, C_NN + j)
      .Value = h(j): .Interior.Color = RGB(48, 84, 150): .Font.Color = vbWhite: .Font.Bold = True
    End With
  Next
  h = Array("GESTOR_ASIGNADO", "DESTINO", "TRAYECTO_TRAMACO", "TIPO_ENTREGA", "ZONA_PELIGROSA", "GESTOR_COBERTURA(Q)", "GESTOR_SUGERIDO(R)", "ORIGINAL_CLIENTE", "TIPO_CORRECCION")
  For j = 0 To N_EXT - 1
    With ws.Cells(1, C_EXT + j)
      .Value = h(j): .Interior.Color = RGB(0, 128, 96): .Font.Color = vbWhite: .Font.Bold = True
    End With
  Next
  Set r = ws.Range(ws.Cells(2, C_NN + 3), ws.Cells(lr, C_NN + 3))
  r.FormatConditions.Delete
  With r.FormatConditions.Add(xlCellValue, xlEqual, "=""OK""")
    .Interior.Color = RGB(198, 239, 206)
  End With
  With r.FormatConditions.Add(xlCellValue, xlEqual, "=""REVISAR""")
    .Interior.Color = RGB(255, 235, 156)
  End With
  With r.FormatConditions.Add(xlCellValue, xlEqual, "=""APROBADO""")
    .Interior.Color = RGB(189, 215, 238)
  End With
  Set r = ws.Range(ws.Cells(2, C_EXT + 4), ws.Cells(lr, C_EXT + 4))
  r.FormatConditions.Delete
  With r.FormatConditions.Add(Type:=xlTextString, String:="ZONA PELIGROSA", TextOperator:=xlContains)
    .Interior.Color = RGB(255, 199, 206)
  End With
End Sub

Sub AprobarTodas()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim i As Long, cntRev As Long, cntFC As Long
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If UCase$(Trim$(TX(ws.Cells(i, C_NN + 3)))) = "REVISAR" Then
        cntRev = cntRev + 1
        If InStr(TX(ws.Cells(i, C_NN + 6)), "Fuera de cobertura") > 0 Then cntFC = cntFC + 1
      End If
    End If
  Next
  If cntRev = 0 Then MsgBox "No hay pedidos en REVISAR. Usa 'Aplicar aprobados'.", vbInformation: Exit Sub
  If MsgBox("Se aprobarán " & cntRev & " pedidos con las sugerencias del sistema" & _
     IIf(cntFC > 0, " (" & cntFC & " fuera de cobertura con cobertura cercana sugerida)", "") & "." & vbCrLf & _
     "Recomendado: revisarlos primero con 'Revisar pendientes'. ¿Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If UCase$(Trim$(TX(ws.Cells(i, C_NN + 3)))) = "REVISAR" Then
        ws.Cells(i, C_NN + 3).Value = "APROBADO"
        ws.Cells(i, C_NN + 6).Value = "Aprobado en lote " & Format(Now, "yyyy-mm-dd hh:nn") & " | " & TX(ws.Cells(i, C_NN + 6))
      End If
    End If
  Next
  MsgBox cntRev & " aprobados. Usa 'Aplicar aprobados'.", vbInformation
End Sub

' ---------- CAMBIOS: historial acumulado (ya no se borra) ----------
Public Function HojaLog() As Worksheet
  Dim lg As Worksheet
  On Error Resume Next
  Set lg = ThisWorkbook.Worksheets(HLOG)
  On Error GoTo 0
  If lg Is Nothing Then
    Set lg = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(HD)): lg.Name = HLOG
  End If
  If TX(lg.Cells(1, 12)) <> "ORIGEN" Or TX(lg.Cells(1, 17)) <> "TIPO_CORRECCION" Then
    lg.Range("A1:Q1").Value = Array("FECHA_PROCESO", "NRO_REFERENCIA", "DESTINATARIO", "DIRECCION_CLIENTE", "PROV_ORIGINAL", "PROV_NUEVA", _
      "CANTON_ORIGINAL", "CANTON_NUEVO", "PARROQUIA_ORIG", "PARROQUIA_NUEVA", "MOTIVO", "ORIGEN", "USUARIO", "SIGLA", "GESTOR", "DESTINO", "TIPO_CORRECCION")
    With lg.Range("A1:Q1"): .Font.Bold = True: .Interior.Color = RGB(48, 84, 150): .Font.Color = vbWhite: End With
  End If
  Set HojaLog = lg
End Function

Sub AplicarAprobados()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim lg As Worksheet: Set lg = HojaLog()
  Dim i As Long, n As Long, g As Long, g0 As Long, est As String, canton As String, parr As String
  Dim oF As String, oG As String, oH As String, nF As String, mot As String, origen As String
  g = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row + 1: If g < 2 Then g = 2
  g0 = g
  Application.ScreenUpdating = False
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo cont
    est = UCase$(Trim$(TX(ws.Cells(i, C_NN + 3))))
    If est = "APROBADO" Or est = "OK" Then
      oF = Trim$(TX(ws.Cells(i, C_PROV))): oG = Trim$(TX(ws.Cells(i, C_CANT))): oH = Trim$(TX(ws.Cells(i, C_PARR)))
      nF = ProvExport(TX(ws.Cells(i, C_NN))): canton = Trim$(TX(ws.Cells(i, C_NN + 1))): parr = Trim$(TX(ws.Cells(i, C_NN + 2)))
      If Len(nF) = 0 Or Len(parr) = 0 Then GoTo cont
      If Len(canton) = 0 Then canton = parr
      If oF <> nF Or oG <> canton Or oH <> parr Then
        mot = TX(ws.Cells(i, C_NN + 6))
        origen = "SISTEMA"
        If InStr(1, mot, "operario", vbTextCompare) > 0 Then origen = "OPERARIO"
        If InStr(1, mot, "lote", vbTextCompare) > 0 Then origen = "LOTE"
        lg.Cells(g, 1).NumberFormat = "yyyy-mm-dd hh:mm": lg.Cells(g, 1).Value = Now
        lg.Cells(g, 2).NumberFormat = "@": lg.Cells(g, 2).Value = TX(ws.Cells(i, C_REF))
        lg.Cells(g, 3).Value = TX(ws.Cells(i, 5))
        lg.Cells(g, 4).Value = TX(ws.Cells(i, C_DIR))
        lg.Cells(g, 5).Value = oF: lg.Cells(g, 6).Value = nF
        lg.Cells(g, 7).Value = oG: lg.Cells(g, 8).Value = canton
        lg.Cells(g, 9).Value = oH: lg.Cells(g, 10).Value = parr
        lg.Cells(g, 11).Value = mot
        lg.Cells(g, 12).Value = origen
        lg.Cells(g, 13).Value = Application.UserName
        lg.Cells(g, 14).Value = TX(ws.Cells(i, C_SIG))
        lg.Cells(g, 15).Value = TX(ws.Cells(i, C_EXT))
        lg.Cells(g, 16).Value = TX(ws.Cells(i, C_EXT + 1))
        lg.Cells(g, 17).Value = TX(ws.Cells(i, C_EXT + N_EXTC + 1))
        g = g + 1
      End If
      ws.Cells(i, C_PROV).Value = nF
      ws.Cells(i, C_CANT).Value = canton
      ws.Cells(i, C_PARR).Value = parr
      n = n + 1
    End If
cont:
  Next
  lg.Columns("A:Q").AutoFit
  Application.ScreenUpdating = True
  MsgBox n & " aplicados." & vbCrLf & (g - g0) & " cambios agregados al historial de la hoja CAMBIOS.", vbInformation
End Sub

Private Function EsHoy(v As Variant) As Boolean
  If IsError(v) Then Exit Function
  If IsDate(v) Then EsHoy = (Int(CDate(v)) = Date)
End Function

Sub ExportarCambios()
  Dim lg As Worksheet: Set lg = HojaLog()
  Dim lr As Long: lr = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row
  If lr < 2 Then MsgBox "No hay cambios registrados.", vbInformation: Exit Sub
  Dim r As VbMsgBoxResult
  r = MsgBox("¿Exportar solo los cambios de HOY?" & vbCrLf & "Sí = solo hoy    No = todo el historial", vbYesNoCancel + vbQuestion, "Exportar CAMBIOS")
  If r = vbCancel Then Exit Sub
  Dim wbN As Workbook, wsN As Worksheet, i As Long, g As Long, f As String
  Set wbN = Workbooks.Add(xlWBATWorksheet)
  Set wsN = wbN.Worksheets(1): wsN.Name = "CAMBIOS"
  lg.Range("A1:Q1").Copy wsN.Range("A1")
  g = 2
  For i = 2 To lr
    If r = vbNo Or EsHoy(lg.Cells(i, 1).Value) Then
      lg.Range(lg.Cells(i, 1), lg.Cells(i, 17)).Copy wsN.Cells(g, 1): g = g + 1
    End If
  Next
  wsN.Columns("A:Q").AutoFit
  If Len(ThisWorkbook.Path) > 0 Then
    f = ThisWorkbook.Path & Application.PathSeparator & "CAMBIOS_HYCITE_" & Format(Now, "yyyymmdd_hhnn") & ".xlsx"
    On Error Resume Next
    wbN.SaveAs f, xlOpenXMLWorkbook
    If Err.Number = 0 Then
      On Error GoTo 0
      MsgBox (g - 2) & " cambios exportados a:" & vbCrLf & f, vbInformation
      Exit Sub
    End If
    On Error GoTo 0
  End If
  MsgBox (g - 2) & " cambios copiados a un libro nuevo (guárdalo manualmente).", vbInformation
End Sub

Sub RevisarCambios()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  CargarBases
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim i As Long, prob As Long, msg As String, P As String, nc As String, Q As String, txtErr As String
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo cont
    If UCase$(Trim$(TX(ws.Cells(i, C_NN + 3)))) = "APROBADO" Then
      P = Normaliza(TX(ws.Cells(i, C_NN))): nc = Normaliza(TX(ws.Cells(i, C_NN + 1))): Q = Normaliza(TX(ws.Cells(i, C_NN + 2)))
      txtErr = ""
      If Not mPC.Exists(P) Then txtErr = txtErr & "provincia inválida; "
      If Q <> "QUITO" And Not mTriple.Exists(P & "|" & nc & "|" & Q) Then txtErr = txtErr & "provincia/cantón/parroquia no está en COBERTURA; "
      If Len(txtErr) > 0 Then
        prob = prob + 1: ws.Cells(i, C_NN + 3).Interior.Color = RGB(255, 199, 206)
        If prob <= 15 Then msg = msg & "Fila " & i & " (" & TX(ws.Cells(i, C_REF)) & "): " & txtErr & vbCrLf
      Else
        ws.Cells(i, C_NN + 3).Interior.Color = RGB(198, 239, 206)
      End If
    End If
cont:
  Next
  If prob = 0 Then MsgBox "Revisión OK: aprobados consistentes con COBERTURA.", vbInformation Else MsgBox prob & " con posibles errores (en rojo):" & vbCrLf & vbCrLf & msg, vbExclamation
End Sub

' ---------- exportación al archivo EGR (mismo mapeo C:N de siempre) ----------
Sub ExportarADatos()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim wbT As Workbook, wsT As Worksheet, wb As Workbook
  For Each wb In Application.Workbooks
    If InStr(UCase$(wb.Name), "EGR") > 0 And InStr(UCase$(wb.Name), "HYCITE") > 0 And InStr(UCase$(wb.Name), "COPIA") = 0 Then Set wbT = wb: Exit For
  Next
  If wbT Is Nothing Then
    For Each wb In Application.Workbooks
      If InStr(UCase$(wb.Name), "EGR") > 0 And InStr(UCase$(wb.Name), "HYCITE") > 0 Then Set wbT = wb: Exit For
    Next
  End If
  If wbT Is Nothing Then MsgBox "Abre el archivo 'Formato EGR ... HYCITE' y reintenta.", vbExclamation: Exit Sub
  On Error Resume Next
  Set wsT = wbT.Worksheets("DATOS")
  On Error GoTo 0
  If wsT Is Nothing Then MsgBox "No se encontró la hoja 'DATOS'.", vbExclamation: Exit Sub

  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim i As Long, cnt As Long, k As Long, nRev As Long, nSinAplicar As Long, e As String
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If Len(Trim$(TX(ws.Cells(i, C_REF)))) > 0 Then
        cnt = cnt + 1
        e = UCase$(Trim$(TX(ws.Cells(i, C_NN + 3))))
        If e = "REVISAR" Then nRev = nRev + 1
        If (e = "OK" Or e = "APROBADO") And Len(TX(ws.Cells(i, C_NN + 2))) > 0 Then
          If Normaliza(TX(ws.Cells(i, C_NN + 2))) <> Normaliza(TX(ws.Cells(i, C_PARR))) Or _
             Normaliza(TX(ws.Cells(i, C_NN + 1))) <> Normaliza(TX(ws.Cells(i, C_CANT))) Then nSinAplicar = nSinAplicar + 1
        End If
      End If
    End If
  Next
  If cnt = 0 Then MsgBox "No hay filas visibles para exportar.", vbInformation: Exit Sub
  If nRev > 0 Or nSinAplicar > 0 Then
    If MsgBox("Atención antes de enviar a '" & wbT.Name & "':" & vbCrLf & _
       IIf(nRev > 0, "  - " & nRev & " pedidos siguen en REVISAR" & vbCrLf, "") & _
       IIf(nSinAplicar > 0, "  - " & nSinAplicar & " correcciones aún no aplicadas (usa 'Aplicar aprobados')" & vbCrLf, "") & _
       vbCrLf & "¿Exportar de todas formas?", vbYesNo + vbExclamation, "Enviar a EGR") <> vbYes Then Exit Sub
  End If

  Dim data(): ReDim data(1 To cnt, 1 To 12)
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If Len(Trim$(TX(ws.Cells(i, C_REF)))) > 0 Then
        k = k + 1
        data(k, 1) = TX(ws.Cells(i, 1))                 ' C: CLIENTE (EMPRESA)
        data(k, 2) = TX(ws.Cells(i, C_DIR))             ' D: dirección
        data(k, 3) = TX(ws.Cells(i, C_REF))             ' E: referencia
        data(k, 4) = TX(ws.Cells(i, 4))                 ' F: código viaje
        data(k, 5) = TX(ws.Cells(i, 5))                 ' G: destinatario
        data(k, 6) = ProvExport(TX(ws.Cells(i, 6)))     ' H: provincia (sin Ñ)
        data(k, 7) = TX(ws.Cells(i, 7))                 ' I: cantón
        data(k, 8) = TX(ws.Cells(i, 8))                 ' J: parroquia (CP_DEST)
        data(k, 9) = TX(ws.Cells(i, 9))                 ' K: móvil
        data(k, 10) = TX(ws.Cells(i, 10))               ' L: fijo
        data(k, 11) = TX(ws.Cells(i, 11))               ' M: email
        data(k, 12) = TX(ws.Cells(i, C_SIG))            ' N: sigla
      End If
    End If
  Next
  Dim lrT As Long: lrT = wsT.Cells(wsT.Rows.Count, 3).End(xlUp).Row
  If lrT >= 2 Then wsT.Range(wsT.Cells(2, 3), wsT.Cells(lrT, 14)).ClearContents
  wsT.Range(wsT.Cells(2, 3), wsT.Cells(cnt + 1, 14)).Value = data
  MsgBox cnt & " filas exportadas a '" & wbT.Name & "' hoja DATOS (C2:N" & (cnt + 1) & ").", vbInformation
End Sub

' ---------- menú (pestaña Complementos) en el orden del flujo ----------
Sub CrearMenuHycite()
  On Error Resume Next
  Application.CommandBars("Validacion HYCITE").Delete
  On Error GoTo 0
  Dim bar As CommandBar
  Set bar = Application.CommandBars.Add(Name:="Validacion HYCITE", Position:=msoBarTop, Temporary:=True)
  AddBtn bar, "Panel HYCITE", "AbrirPanel", 2174, True
  AddBtn bar, "0 Actualizar datos", "ActualizarDatosDepot", 37, True
  AddBtn bar, "1 Limpiar", "LimpiarValidacion", 47, False
  AddBtn bar, "2 Validar", "ValidarPedidos", 25, False
  AddBtn bar, "3 Revisar pendientes", "AbrirValidadorRevisar", 6362, False
  AddBtn bar, "4 Aprobar lote", "AprobarTodas", 356, False
  AddBtn bar, "5 Aplicar aprobados", "AplicarAprobados", 358, False
  AddBtn bar, "6 Actualizar siglas", "ActualizarSiglas", 1017, False
  AddBtn bar, "7 Enviar a EGR", "ExportarADatos", 3, True
  AddBtn bar, "Exportar CAMBIOS", "ExportarCambios", 1561, True
  AddBtn bar, "Desbloquear", "Desbloquear", 225, False
  bar.Visible = True
End Sub

Private Sub AddBtn(bar As CommandBar, cap As String, macro As String, face As Long, grupo As Boolean)
  Dim b As CommandBarButton
  Set b = bar.Controls.Add(msoControlButton)
  b.Caption = cap
  b.OnAction = macro
  b.Style = msoButtonIconAndCaption
  On Error Resume Next
  b.FaceId = face
  On Error GoTo 0
  b.BeginGroup = grupo
End Sub

Function TX(c As Range) As String
  On Error Resume Next
  If IsError(c.Value) Then TX = "" Else TX = CStr(c.Value)
  On Error GoTo 0
End Function

' Se conserva, pero ya no está en la barra: la consulta Depot no trae columna FECHA
' (el filtro de fecha lo hace el SQL).
Sub FiltrarPorFecha()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  If lr < 2 Then Exit Sub
  Dim C_FEC As Long, j As Long
  For j = 1 To 40
    If InStr(UCase$(ws.Cells(1, j).Value & ""), "FECHA") > 0 Then C_FEC = j: Exit For
  Next
  If C_FEC = 0 Then MsgBox "No se encontró la columna FECHA en la fila 1 (la consulta Depot ya filtra por fecha).", vbInformation: Exit Sub
  Dim defd As String: defd = Format(Date, "dd/mm/yyyy")
  Dim resp As String: resp = InputBox("Fecha del día a CONSERVAR (dd/mm/aaaa)." & vbCrLf & _
     "Se eliminarán las filas de otras fechas.", "Filtrar por día", defd)
  If Trim$(resp) = "" Then Exit Sub
  If Not IsDate(resp) Then MsgBox "Fecha no válida.", vbExclamation: Exit Sub
  Dim target As Long: target = Int(CDate(resp))
  If MsgBox("Se conservarán SOLO los pedidos del " & Format(target, "dd/mm/yyyy") & _
     " y se ELIMINARÁN las filas de otras fechas. ¿Continuar?", vbYesNo + vbExclamation, "Filtrar por día") <> vbYes Then Exit Sub
  Application.ScreenUpdating = False
  Dim i As Long, v, borradas As Long
  For i = lr To 2 Step -1
    v = ws.Cells(i, C_FEC).Value
    If Not IsError(v) Then
      If IsDate(v) Then
        If Int(CDate(v)) <> target Then ws.Rows(i).Delete: borradas = borradas + 1
      End If
    End If
  Next
  Application.ScreenUpdating = True
  MsgBox borradas & " filas de otras fechas eliminadas.", vbInformation
End Sub

Function FindCP(ByVal a As String, ByRef cpPos As Long) As String
  Static re As Object
  If re Is Nothing Then
    Set re = CreateObject("VBScript.RegExp"): re.Global = True
    re.Pattern = "(^|\D)(\d{6})(\D|$)"
  End If
  Dim m, mm: Set m = re.Execute(a)
  cpPos = 0: FindCP = ""
  If m.Count = 0 Then Exit Function
  Set mm = m(m.Count - 1)
  FindCP = mm.SubMatches(1)
  cpPos = mm.FirstIndex + Len(mm.SubMatches(0)) + 1
End Function

Function ProvFromCP(ByVal cp As String) As String
  If Len(cp) < 2 Then Exit Function
  If mRev.Exists(Left$(cp, 2)) Then ProvFromCP = mRev(Left$(cp, 2))
End Function

Function ParseColaDireccion(ByVal a As String, ByRef prov As String, ByRef parrO As String, _
    ByRef parrFin As String, ByRef cp As String) As Boolean
  Dim cpPos As Long: cp = FindCP(a, cpPos)
  If Len(cp) = 0 Or cpPos <= 1 Then Exit Function
  Dim cands As New Collection, pcp As String
  pcp = ProvFromCP(cp): If Len(pcp) > 0 Then cands.Add pcp
  Dim tail As String: tail = " " & Trim$(Mid$(a, cpPos + 6)) & " "
  Dim pr
  For Each pr In mPC.Keys
    If InStr(tail, " " & pr & " ") > 0 Then cands.Add CStr(pr)
  Next
  Dim pre As String: pre = " " & Trim$(Left$(a, cpPos - 1)) & " "
  Dim L As Long: L = Len(pre)
  Dim cc, pool As Collection, it, best As String, bestOrig As String, maxLen As Long, cLen As Long, bok As Boolean
  For Each cc In cands
    If mParrList.Exists(cc) Then
      Set pool = mParrList(cc): maxLen = 0: best = "": bestOrig = ""
      For Each it In pool
        cLen = Len(it(0))
        If cLen > maxLen And cLen >= 3 Then
          If Right$(pre, cLen + 1) = it(0) & " " Then
            bok = (L - cLen - 1 = 0)
            If Not bok Then bok = (Mid$(pre, L - cLen - 1, 1) = " ")
            If bok Then maxLen = cLen: best = it(0): bestOrig = it(2)
          End If
        End If
      Next
      If maxLen > 0 Then prov = cc: parrFin = best: parrO = bestOrig: ParseColaDireccion = True: Exit Function
    End If
  Next
End Function

Function CodigoSig(ByVal sig As String) As String
  Dim p1 As Long, p2 As Long
  p2 = InStrRev(sig, ")"): p1 = InStrRev(sig, "(")
  If p1 > 0 And p2 > p1 Then CodigoSig = Normaliza(Mid$(sig, p1 + 1, p2 - p1 - 1))
End Function

Function SiglaFinal(ByVal P As String, ByVal ncF As String, ByVal Q As String, ByVal parrO As String) As String
  If mTriple Is Nothing Then CargarBases
  ' 1. Quito genérico (sin parroquia real) -> CALDERON (QUI)
  If P = "PICHINCHA" And ncF = "QUITO" Then
    If Q = "QUITO" Or InStr(Q, "DISTRITO METROPOLITANO") > 0 Or Len(Q) = 0 Then
      SiglaFinal = "CALDERON (QUI)": Exit Function
    End If
  End If
  ' 2. Fuera de cobertura -> sugerencia de cobertura cercana
  If Len(ncF) > 0 And Len(Q) > 0 Then
    If Not mTriple.Exists(P & "|" & ncF & "|" & Q) Then
      Dim scx, motx As String, sugx As String
      scx = SugerirCercana(P, ncF, Q, "", motx)
      If IsArray(scx) Then sugx = scx(1)
      If Len(sugx) > 0 Then
        SiglaFinal = "Fuera de cobertura (Sug: " & sugx & ")"
      Else
        SiglaFinal = "Fuera de cobertura"
      End If
      Exit Function
    End If
  End If
  ' 3. Sigla del maestro
  Dim sg As String
  If mSiglas.Exists(P & "|" & ncF & "|" & Q) Then sg = mSiglas(P & "|" & ncF & "|" & Q)
  If Len(sg) = 0 And mSiglas.Exists(P & "||" & Q) Then sg = mSiglas(P & "||" & Q)
  If Len(sg) > 0 Then SiglaFinal = sg: Exit Function
  ' 4. Generada con el código del cantón
  Dim cod As String
  If P = "PICHINCHA" And ncF = "QUITO" Then
    cod = "QUI"
  ElseIf mCodeCant.Exists(P & "|" & ncF) Then
    cod = mCodeCant(P & "|" & ncF)
  End If
  Dim baseName As String: baseName = Trim$(parrO)
  If Len(baseName) = 0 Then baseName = Q
  Dim pp As Long: pp = InStrRev(baseName, "(")
  If pp > 1 Then baseName = Trim$(Left$(baseName, pp - 1))
  If Len(cod) > 0 Then SiglaFinal = baseName & " (" & cod & ")" Else SiglaFinal = baseName
End Function

Sub ActualizarSiglas()   ' recalcula L y V:Z desde F/G/H (editados a mano)
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  CargarBases
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim i As Long, n As Long, j As Long, ex
  Dim P As String, ncF As String, Q As String, parrO As String
  Application.ScreenUpdating = False
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo sig
    P = Normaliza(TX(ws.Cells(i, C_PROV)))
    ncF = Normaliza(TX(ws.Cells(i, C_CANT)))
    Q = Normaliza(TX(ws.Cells(i, C_PARR)))
    parrO = Trim$(TX(ws.Cells(i, C_PARR)))
    If Len(P) > 0 And Len(Q) > 0 Then
      ws.Cells(i, C_SIG).Value = SiglaFinal(P, ncF, Q, parrO): n = n + 1
      ex = DatosExt(P, ncF, Q, " " & Normaliza(TX(ws.Cells(i, C_DIR))) & " ")
      For j = 0 To N_EXTC - 1: ws.Cells(i, C_EXT + j).Value = ex(j): Next
    End If
sig:
  Next
  Application.ScreenUpdating = True
  MsgBox n & " siglas y gestores recalculados (L y V:Z).", vbInformation
End Sub

' ---------- formularios ----------
Sub AbrirValidador()
  CerrarForm "frmValidar"
  gSoloRevisar = False
  frmValidar.Show vbModeless
End Sub

Sub AbrirValidadorRevisar()
  CerrarForm "frmValidar"
  gSoloRevisar = True
  frmValidar.Show vbModeless
End Sub

Sub AbrirValidadorFila(ByVal fila As Long)
  CerrarForm "frmValidar"
  gSoloRevisar = False: gFilaInicio = fila
  frmValidar.Show vbModeless
End Sub

Sub AbrirPanel()
  CerrarForm "frmPanel"
  frmPanel.Show vbModeless
End Sub

' Antes llamaba a frmRevisar (no existe) y rompía "Depuración > Compilar". Ahora abre los pendientes.
Sub RevisarSugerencias()
  AbrirValidadorRevisar
End Sub
