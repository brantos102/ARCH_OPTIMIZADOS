Option Explicit

Const HD As String = "DEPOT"
Const H1 As String = "COBERTURA"
Const C_DIR As Long = 2, C_REF As Long = 3
Const C_PROV As Long = 6, C_CANT As Long = 7, C_PARR As Long = 8
Const C_NN As Long = 14
Const SOLO_FILTRADO As Boolean = True

Dim mTriple As Object, mCantonProvM As Object
Dim mPC As Object, mRev As Object, mProvParr As Object, mParrList As Object
Dim mCantonProv As Object, mParrGlobal As Object, mPcSet As Object, mCantonParr As Object, mParrOrig As Object
Dim mSiglas As Object, mCodeCant As Object


Sub Desbloquear()
  On Error Resume Next
  Application.ScreenUpdating = True: Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic: Application.Cursor = xlDefault
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

Function Lev(ByVal a As String, ByVal b As String) As Long
  Dim m As Long, n As Long, i As Long, j As Long, prev As Long, t As Long, c As Long
  m = Len(a): n = Len(b)
  If m = 0 Then Lev = n: Exit Function
  If n = 0 Then Lev = m: Exit Function
  Dim P() As Long: ReDim P(0 To n)
  For j = 0 To n: P(j) = j: Next
  For i = 1 To m
    prev = P(0): P(0) = i
    For j = 1 To n
      t = P(j): c = IIf(Mid$(a, i, 1) = Mid$(b, j, 1), 0, 1)
      P(j) = Application.Min(P(j) + 1, P(j - 1) + 1, prev + c): prev = t
    Next
  Next
  Lev = P(n)
End Function

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
  Dim ws As Worksheet, lr As Long, arr, i As Long
  Set ws = ThisWorkbook.Worksheets(H1)
  lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  arr = ws.Range("A2:D" & lr).Value
   For i = 1 To UBound(arr, 1): AddBase arr(i, 1), arr(i, 2), arr(i, 3), arr(i, 4): Next
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
  ' --- NUEVO ---
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
  Dim n As Long, used() As Boolean, x As Long, j As Long, bi As Long, bd As Long, dd As Long, s As String
  n = pool.Count: If n = 0 Then Exit Function
  ReDim used(1 To n)
  For x = 1 To Application.Min(topN, n)
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


Sub ValidarPedidos()
  On Error GoTo cleanup
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
  CargarBases
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  If lr < 2 Then GoTo fin
  Dim d, out, i As Long
  d = ws.Range(ws.Cells(1, 1), ws.Cells(lr, C_PARR)).Value
  If SOLO_FILTRADO Then
    out = ws.Range(ws.Cells(2, C_NN), ws.Cells(lr, C_NN + 7)).Value
  Else
    ReDim out(1 To lr - 1, 1 To 8)
  End If
  Dim nP As String, nQ As String, nG As String, a As String, votes As Object, pr, cn, prov As String
  Dim bvote As Long, vk, evid As String, aCinP As Collection, ci2, pg As String
  Dim parrO As String, parrFin As String, estParr As String, sug As String, cantonFin As String
  Dim scopeNc As String, pool As Collection, cand, it, ncF As String
  Dim alert As String, est As String, accion As String, cantonConf As Boolean, esOK As Boolean, okC As Long, revC As Long
  Dim locs, uDup As Object, z, pparts, keyd As String
  Dim usedCola As Boolean, cpStr As String, qAlert As String, ncCola As String
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo sig
    nP = Normaliza(d(i, C_PROV)): nG = Normaliza(d(i, C_CANT)): nQ = Normaliza(d(i, C_PARR))
    If Len(nP) = 0 And Len(nQ) = 0 And Len(Normaliza(d(i, C_DIR))) = 0 Then GoTo sig
    a = " " & Normaliza(d(i, C_DIR)) & " ": evid = "": sug = ""
    Set votes = CreateObject("Scripting.Dictionary")
    ' >>> NUEVO: parser de cola + CP como fuente primaria
    usedCola = False
    If ParseColaDireccion(a, prov, parrO, parrFin, cpStr) Then
      Set aCinP = New Collection
      cantonFin = ResolveCanton(prov, parrFin, aCinP)
      ncCola = Normaliza(cantonFin)
      evid = "cp:" & cpStr & " prov:" & Pretty(prov)
      If mTriple.Exists(prov & "|" & ncCola & "|" & parrFin) Then
        estParr = "OK"
      ElseIf mProvParr.Exists(prov) And InStr(mProvParr(prov), "|" & parrFin & "|") > 0 Then
        estParr = "OK"
      Else
        estParr = "COLA"
      End If
      If Len(nP) > 0 And nP <> prov Then evid = evid & " (F=" & Pretty(nP) & "?)"
      usedCola = True
      GoTo tras_resolucion
    End If
    ' <<< fin NUEVO — a partir de aquí, lógica de voteo original
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
      If mParrOrig.Exists(prov & "|" & nQ) Then parrO = mParrOrig(prov & "|" & nQ) Else parrO = Trim$(d(i, C_PARR) & "")
      estParr = "OK": sug = ""
      cantonFin = ResolveCanton(prov, nQ, aCinP)
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
        sug = FuzzyStr(pool, nQ, 3)
        parrO = FirstFuzzy(pool, nQ): parrFin = Normaliza(parrO)
        If mParrGlobal.Exists(nQ) Then estParr = "OTRA_PROV" Else estParr = "SUGERIR"
      End If
      cantonFin = ResolveCanton(prov, parrFin, aCinP)
    End If
tras_resolucion:
    If prov = "PICHINCHA" And estParr <> "OK" Then
      If InStr(a, " QUITO ") > 0 Or InStr(a, " DMQ ") > 0 Or InStr(a, " DISTRITO METROPOLITANO ") > 0 Then
        cantonFin = "QUITO"
        If estParr = "SUGERIR" Or estParr = "OTRA_PROV" Then parrO = "DISTRITO METROPOLITANO DE QUITO": parrFin = "QUITO": estParr = "QUITO"
      End If
    End If
    If Len(parrO) = 0 Then parrO = Trim$(d(i, C_PARR) & "")
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
    esOK = (estParr = "OK")
    If esOK And Not mPC.Exists(prov) Then esOK = False
    If esOK And Not (nP = "" Or nP = prov) And Not usedCola Then esOK = False
    If esOK And (Len(alert) > 0 And Not cantonConf) And Not usedCola Then esOK = False
    If esOK Then
      est = "OK": accion = "Sin acción": okC = okC + 1
    Else
      est = "REVISAR": revC = revC + 1
      Select Case True
        Case estParr = "COLA": accion = "Confirmar parroquia (dirección/CP)"
        Case estParr = "OTRA_PROV": accion = "Parroquia de otra provincia"
        Case estParr = "SUGERIR": accion = "Elegir parroquia sugerida"
        Case estParr = "DIR": accion = "Parroquia de la dirección"
        Case estParr = "QUITO": accion = "Confirmar Quito/DMQ"
        Case (Len(alert) > 0 And Not cantonConf): accion = "Verificar duplicado"
        Case Else: accion = "Revisar"
      End Select
    End If
    out(i - 1, 1) = Pretty(prov): out(i - 1, 2) = cantonFin: out(i - 1, 3) = parrO
    out(i - 1, 4) = est: out(i - 1, 5) = Trim$(evid): out(i - 1, 6) = sug: out(i - 1, 7) = accion: out(i - 1, 8) = alert
    ws.Cells(i, 12).Value = SiglaFinal(prov, ncF, parrFin, parrO)
sig:
  Next
  ws.Range(ws.Cells(2, C_NN), ws.Cells(lr, C_NN + 7)).Value = out
  FormatoPanel ws, lr
  ColorearCambios ws, lr
fin:
  Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
  MsgBox "Validación lista." & vbCrLf & "OK: " & okC & "    Por revisar: " & revC, vbInformation
  Exit Sub
cleanup:
  Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True: Application.EnableEvents = True
  MsgBox "Se detuvo por un error (pantalla liberada): " & err.Description, vbExclamation
End Sub

Sub LimpiarValidacion()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  If lr < 2 Then lr = 2
  If MsgBox("Se limpiará el panel de validación (N:U), la columna L y sus colores." & vbCrLf & _
     "La data del cliente (A:K) NO se toca. ¿Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  With ws.Range(ws.Cells(2, 12), ws.Cells(lr, 12)): .ClearContents: .Interior.ColorIndex = xlNone: End With
  With ws.Range(ws.Cells(2, C_NN), ws.Cells(lr, C_NN + 7)): .ClearContents: .Interior.ColorIndex = xlNone: End With
  MsgBox "Panel y siglas limpiados. Listo para la nueva data del día.", vbInformation
End Sub

Sub FormatoPanel(ws As Worksheet, lr As Long)
  Dim h, j As Long, r As Range
  h = Array("PROVINCIA_PROP", "CANTON/CIUDAD_PROP", "PARROQUIA_PROP", "ESTADO", "EVIDENCIA_DIRECCION", "SUGERENCIAS", "ACCION_OPERARIO", "ALERTAS/DUPLICADOS")
  For j = 0 To 7
    With ws.Cells(1, C_NN + j)
      .Value = h(j): .Interior.Color = RGB(48, 84, 150): .Font.Color = vbWhite: .Font.bold = True
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
End Sub

Sub AprobarTodas()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim i As Long, cntRev As Long
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If UCase$(Trim$(TX(ws.Cells(i, C_NN + 3)))) = "REVISAR" Then cntRev = cntRev + 1
    End If
  Next
  If cntRev = 0 Then MsgBox "No hay pedidos en REVISAR. Usa '4) Aplicar aprobados'.", vbInformation: Exit Sub
  If MsgBox("Se aprobarán " & cntRev & " pedidos con las sugerencias. ¿Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If UCase$(Trim$(TX(ws.Cells(i, C_NN + 3)))) = "REVISAR" Then
        ws.Cells(i, C_NN + 3).Value = "APROBADO"
        ws.Cells(i, C_NN + 6).Value = "Aprobado en lote " & Format(Now, "yyyy-mm-dd hh:nn")
      End If
    End If
  Next
  MsgBox cntRev & " aprobados. Usa '4) Aplicar aprobados'.", vbInformation
End Sub

Sub AplicarAprobados()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim lg As Worksheet
  On Error Resume Next
  Set lg = ThisWorkbook.Worksheets("CAMBIOS")
  On Error GoTo 0
  If lg Is Nothing Then
    Set lg = ThisWorkbook.Worksheets.Add(After:=ws): lg.Name = "CAMBIOS"
  Else
    lg.Cells.Clear
  End If
  lg.Range("A1:K1").Value = Array("FECHA_PROCESO", "NRO_REFERENCIA", "DESTINATARIO", "DIRECCION_CLIENTE", "PROV_ORIGINAL", "PROV_NUEVA", "CANTON_ORIGINAL", "CANTON_NUEVO", "PARROQUIA_ORIG", "PARROQUIA_NUEVA", "MOTIVO")
  lg.Range("A1:K1").Font.bold = True: lg.Range("A1:K1").Interior.Color = RGB(48, 84, 150): lg.Range("A1:K1").Font.Color = vbWhite
  Dim i As Long, n As Long, g As Long, est As String, canton As String, parr As String
  Dim oF As String, oG As String, oH As String, nF As String
  g = 2
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo cont
    est = UCase$(Trim$(TX(ws.Cells(i, C_NN + 3))))
    If est = "APROBADO" Or est = "OK" Then
      oF = Trim$(TX(ws.Cells(i, C_PROV))): oG = Trim$(TX(ws.Cells(i, C_CANT))): oH = Trim$(TX(ws.Cells(i, C_PARR)))
      nF = Trim$(TX(ws.Cells(i, C_NN))): canton = Trim$(TX(ws.Cells(i, C_NN + 1))): parr = Trim$(TX(ws.Cells(i, C_NN + 2)))
      If Len(canton) = 0 Then canton = parr
      If oF <> nF Or oG <> canton Or oH <> parr Then
        lg.Cells(g, 1).Value = Format(Now, "yyyy-mm-dd hh:nn")
        lg.Cells(g, 2).Value = TX(ws.Cells(i, C_REF))
        lg.Cells(g, 3).Value = TX(ws.Cells(i, 5))
        lg.Cells(g, 4).Value = TX(ws.Cells(i, C_DIR))
        lg.Cells(g, 5).Value = oF: lg.Cells(g, 6).Value = nF
        lg.Cells(g, 7).Value = oG: lg.Cells(g, 8).Value = canton
        lg.Cells(g, 9).Value = oH: lg.Cells(g, 10).Value = parr
        lg.Cells(g, 11).Value = TX(ws.Cells(i, C_NN + 6))
        g = g + 1
      End If
      ws.Cells(i, C_PROV).Value = nF
      ws.Cells(i, C_CANT).Value = canton
      ws.Cells(i, C_PARR).Value = parr
      n = n + 1
    End If
cont:
  Next
  lg.Columns("A:K").AutoFit
  MsgBox n & " aplicados." & vbCrLf & (g - 2) & " cambios registrados en la hoja CAMBIOS.", vbInformation
End Sub

Sub RevisarCambios()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  CargarBases
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim i As Long, prob As Long, msg As String, P As String, Q As String, err As String
  For i = 2 To lr
    If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo cont
    If UCase$(Trim$(TX(ws.Cells(i, C_NN + 3)))) = "APROBADO" Then
      P = Normaliza(TX(ws.Cells(i, C_NN))): Q = Normaliza(TX(ws.Cells(i, C_NN + 2)))
      err = ""
      If Not mPC.Exists(P) Then err = err & "provincia inválida; "
      If mProvParr.Exists(P) Then
        If InStr(mProvParr(P), "|" & Q & "|") = 0 And Q <> "QUITO" Then err = err & "parroquia no está en la provincia; "
      End If
      If Len(err) > 0 Then
        prob = prob + 1: ws.Cells(i, C_NN + 3).Interior.Color = RGB(255, 199, 206)
        If prob <= 15 Then msg = msg & "Fila " & i & " (" & TX(ws.Cells(i, C_REF)) & "): " & err & vbCrLf
      Else
        ws.Cells(i, C_NN + 3).Interior.Color = RGB(198, 239, 206)
      End If
    End If
cont:
  Next
  If prob = 0 Then MsgBox "Revisión OK: aprobados consistentes con COBERTURA.", vbInformation Else MsgBox prob & " con posibles errores (en rojo):" & vbCrLf & vbCrLf & msg, vbExclamation
End Sub

Sub ExportarADatos()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim wbT As Workbook, wsT As Worksheet, wb As Workbook
  
  ' Buscar el archivo de destino
  For Each wb In Application.Workbooks
    If InStr(UCase$(wb.Name), "EGR") > 0 And InStr(UCase$(wb.Name), "HYCITE") > 0 Then Set wbT = wb: Exit For
  Next
  
  If wbT Is Nothing Then MsgBox "Abre el archivo 'Formato EGR ... HYCITE' y reintenta.", vbExclamation: Exit Sub
  
  On Error Resume Next
  Set wsT = wbT.Worksheets("DATOS")
  On Error GoTo 0
  
  If wsT Is Nothing Then MsgBox "No se encontró la hoja 'DATOS'.", vbExclamation: Exit Sub
  
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  Dim i As Long, cnt As Long, k As Long
  
  ' Contar filas a exportar
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If Len(Trim$(TX(ws.Cells(i, C_REF)))) > 0 Then cnt = cnt + 1
    End If
  Next
  
  If cnt = 0 Then MsgBox "No hay filas visibles para exportar.", vbInformation: Exit Sub
  
  ' Dimensionar el arreglo para 12 columnas exactas (C a N)
  Dim data(): ReDim data(1 To cnt, 1 To 12)
  
  For i = 2 To lr
    If Not (SOLO_FILTRADO And ws.Rows(i).Hidden) Then
      If Len(Trim$(TX(ws.Cells(i, C_REF)))) > 0 Then
        k = k + 1
        
        data(k, 1) = TX(ws.Cells(i, 1))         ' Col C: Fecha
        data(k, 2) = TX(ws.Cells(i, C_DIR))     ' Col D: Dirección
        data(k, 3) = TX(ws.Cells(i, C_REF))     ' Col E: Referencia
        data(k, 4) = TX(ws.Cells(i, 4))         ' Col F: CODIGO_VI
        data(k, 5) = TX(ws.Cells(i, 5))         ' Col G: DESTINATARIO
        
        ' --- MAPPING GEOGRÁFICO ---
        data(k, 6) = TX(ws.Cells(i, 6))         ' Col H: PROVINCIA_DEST (Origen F)
        data(k, 7) = TX(ws.Cells(i, 7))         ' Col I: LOCALIDAD_DEST (Origen G)
        data(k, 8) = TX(ws.Cells(i, 8))         ' Col J: CP_DEST (Origen H)
        
        ' --- MAPPING DE CONTACTO Y SIGLAS ---
        data(k, 9) = TX(ws.Cells(i, 9))         ' Col K: TELEFONO MÓVIL (Origen I)
        data(k, 10) = TX(ws.Cells(i, 10))       ' Col L: TELEFONO FIJO (Origen J)
        data(k, 11) = TX(ws.Cells(i, 11))       ' Col M: EMAIL (Origen K)
        data(k, 12) = TX(ws.Cells(i, 12))       ' Col N: LOCALIDAD_DEST_SIG (Origen L)
        
      End If
    End If
  Next
  
  ' Pegar los datos en bloque en la hoja DATOS a partir de la columna C hasta la N (índice 14)
  Dim lrT As Long: lrT = wsT.Cells(wsT.Rows.Count, 3).End(xlUp).Row
  If lrT >= 2 Then wsT.Range(wsT.Cells(2, 3), wsT.Cells(lrT, 14)).ClearContents
  wsT.Range(wsT.Cells(2, 3), wsT.Cells(cnt + 1, 14)).Value = data
  
  MsgBox cnt & " filas exportadas a '" & wbT.Name & "' hoja DATOS (C2:N" & (cnt + 1) & ").", vbInformation
End Sub

Sub CrearMenuHycite()
   On Error Resume Next
   Application.CommandBars("Validacion HYCITE").Delete
   On Error GoTo 0
   Dim bar As CommandBar
   Set bar = Application.CommandBars.Add(Name:="Validacion HYCITE", Position:=msoBarTop, Temporary:=True)
   AddBtn bar, "Validar", "ValidarPedidos", 25, True
   AddBtn bar, "Aprobar", "AprobarTodas", 356, False
   AddBtn bar, "Aplicar aprobados", "AplicarAprobados", 358, False
   AddBtn bar, "Actualizar Parroquia_siglas", "ActualizarSiglas", 1017, True
   AddBtn bar, "Exportar datos", "ExportarADatos", 3, True
   AddBtn bar, "Desbloquear", "Desbloquear", 225, False
   AddBtn bar, "Filtrar día", "FiltrarPorFecha", 461, False
   AddBtn bar, "Validar línea a línea", "AbrirValidador", 6362, True
   AddBtn bar, "Limpiar validación", "LimpiarValidacion", 47, True
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

Sub FiltrarPorFecha()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
  Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
  If lr < 2 Then Exit Sub
  ' localizar la columna FECHA por su encabezado
  Dim C_FEC As Long, j As Long
  For j = 1 To 40
    If InStr(UCase$(ws.Cells(1, j).Value & ""), "FECHA") > 0 Then C_FEC = j: Exit For
  Next
  If C_FEC = 0 Then MsgBox "No se encontró la columna FECHA en la fila 1.", vbExclamation: Exit Sub
  ' fecha por defecto = la de la celda seleccionada (si es fecha), si no, hoy
  Dim defd As String: defd = Format(Date, "dd/mm/yyyy")
  On Error Resume Next
  If IsDate(ws.Cells(ActiveCell.Row, C_FEC).Value) Then defd = Format(ws.Cells(ActiveCell.Row, C_FEC).Value, "dd/mm/yyyy")
  On Error GoTo 0
  ' pedir la fecha a conservar
  Dim resp As String: resp = InputBox("Fecha del día a CONSERVAR (dd/mm/aaaa)." & vbCrLf & _
     "Se eliminarán las filas de otras fechas.", "Filtrar por día", defd)
  If Trim$(resp) = "" Then Exit Sub
  If Not IsDate(resp) Then MsgBox "Fecha no válida.", vbExclamation: Exit Sub
  Dim target As Long: target = Int(CDate(resp))
  If MsgBox("Se conservarán SOLO los pedidos del " & Format(target, "dd/mm/yyyy") & _
     " y se ELIMINARÁN las filas de otras fechas. ¿Continuar?", vbYesNo + vbExclamation, "Filtrar por día") <> vbYes Then Exit Sub
  Application.ScreenUpdating = False
  Dim i As Long, v, borradas As Long, sinFecha As Long
  For i = lr To 2 Step -1
    v = ws.Cells(i, C_FEC).Value
    If IsError(v) Then
      sinFecha = sinFecha + 1                       ' sin fecha válida -> se conserva
    ElseIf IsDate(v) Then
      If Int(CDate(v)) <> target Then ws.Rows(i).Delete: borradas = borradas + 1
    Else
      sinFecha = sinFecha + 1                       ' texto/no fecha -> se conserva
    End If
  Next
  Application.ScreenUpdating = True
  MsgBox "Listo. Quedan solo los pedidos del " & Format(target, "dd/mm/yyyy") & "." & vbCrLf & _
    borradas & " filas de otras fechas eliminadas." & _
    IIf(sinFecha > 0, vbCrLf & sinFecha & " filas SIN fecha válida se conservaron (revisa la columna FECHA).", ""), vbInformation
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
  Set mm = m(m.Count - 1)                       ' último bloque de 6 dígitos = cola
  FindCP = mm.SubMatches(1)
  cpPos = mm.FirstIndex + Len(mm.SubMatches(0)) + 1
End Function

Function ProvFromCP(ByVal cp As String) As String
  If Len(cp) < 2 Then Exit Function
  If mRev.Exists(Left$(cp, 2)) Then ProvFromCP = mRev(Left$(cp, 2))
End Function

' Devuelve True y rellena prov/parrFin/parrO/cp si logra parsear la cola de la dirección
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
    ' 1. Quito genérico (sin parroquia real) -> default CALDERON (QUI)
    If P = "PICHINCHA" And ncF = "QUITO" Then
        If Q = "QUITO" Or InStr(Q, "DISTRITO METROPOLITANO") > 0 Or Len(Q) = 0 Then
            SiglaFinal = "CALDERON (QUI)": Exit Function
        End If
    End If

    ' --- 2. NUEVA REGLA: FUERA DE COBERTURA Y SUGERENCIA INTELIGENTE ---
    If Len(ncF) > 0 And Len(Q) > 0 Then
        ' Si la triada exacta (Provincia|Cantón|Parroquia) no existe en el maestro
        If Not mTriple.Exists(P & "|" & ncF & "|" & Q) Then
            Dim sug As String
            
            ' Intento A: Sugerir la cabecera cantonal (ej. Cantón Quevedo -> Parroquia Quevedo)
            If mTriple.Exists(P & "|" & ncF & "|" & ncF) Then
                sug = Pretty(ncF)
            ' Intento B: Si no hay cabecera, usar el motor Fuzzy para la parroquia más similar
            ElseIf mCantonParr.Exists(P & "|" & ncF) Then
                sug = FirstFuzzy(mCantonParr(P & "|" & ncF), Q)
            End If
            
            ' Retornar la etiqueta de fuera de cobertura con la sugerencia
            If Len(sug) > 0 Then
                SiglaFinal = "Fuera de cobertura (Sug: " & sug & ")"
            Else
                SiglaFinal = "Fuera de cobertura"
            End If
            Exit Function
        End If
    End If
    ' -------------------------------------------------------------------

    ' 3. Lógica original si el destino SÍ es válido
    Dim sg As String
    If mSiglas.Exists(P & "|" & ncF & "|" & Q) Then sg = mSiglas(P & "|" & ncF & "|" & Q)
    If Len(sg) = 0 And mSiglas.Exists(P & "||" & Q) Then sg = mSiglas(P & "||" & Q)
    If Len(sg) > 0 Then SiglaFinal = sg: Exit Function

    ' Generar con el código del cantón si el maestro no tiene sigla
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

Sub ActualizarSiglas()   ' botón adicional: recalcula L desde F/G/H (editados a mano)
   Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(HD)
   CargarBases
   Dim lr As Long: lr = ws.Cells(ws.Rows.Count, C_REF).End(xlUp).Row
   Dim i As Long, n As Long
   Application.ScreenUpdating = False
   For i = 2 To lr
       If SOLO_FILTRADO And ws.Rows(i).Hidden Then GoTo sig
       Dim P As String, ncF As String, Q As String, parrO As String
       P = Normaliza(ws.Cells(i, C_PROV).Value)
       ncF = Normaliza(ws.Cells(i, C_CANT).Value)
       Q = Normaliza(ws.Cells(i, C_PARR).Value)
       parrO = Trim$(TX(ws.Cells(i, C_PARR)))
       If Len(P) > 0 And Len(Q) > 0 Then ws.Cells(i, 12).Value = SiglaFinal(P, ncF, Q, parrO): n = n + 1
sig:
   Next
   Application.ScreenUpdating = True
   MsgBox n & " siglas recalculadas en la columna L.", vbInformation
End Sub

Sub AbrirValidador()
  frmValidar.Show vbModeless
End Sub

Sub RevisarSugerencias()
  frmRevisar.Show vbModeless
End Sub
