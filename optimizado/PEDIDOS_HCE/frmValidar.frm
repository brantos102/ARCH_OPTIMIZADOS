Option Explicit
' =====================================================================================
'  frmValidar (versión optimizada v2) - reemplaza TODO el código del formulario frmValidar.
'  Los controles se crean al abrir: no hay que dibujar nada en el diseñador.
'  Novedades: listas Provincia > Cantón > Parroquia dependientes, sugerencias con su cantón,
'  botón "Sugerir cobertura cercana" (cabecera cantonal / ciudad principal o secundaria),
'  vista previa de sigla, gestor, destino y zona peligrosa, y aviso si queda fuera de cobertura.
' =====================================================================================
Private rws As Collection, idx As Long, wsD As Worksheet, nAprob As Long, mCarga As Boolean, mMotivo As String
Private lblRef As MSForms.Label, lblDir As MSForms.Label, lblCli As MSForms.Label
Private lblSug As MSForms.Label, lblInfo As MSForms.Label
Private WithEvents cboProv As MSForms.ComboBox
Private WithEvents cboCant As MSForms.ComboBox
Private WithEvents cboParr As MSForms.ComboBox
Private WithEvents btnAplicar As MSForms.CommandButton
Private WithEvents btnSugerir As MSForms.CommandButton
Private WithEvents btnOmitir As MSForms.CommandButton
Private WithEvents btnAnterior As MSForms.CommandButton
Private WithEvents btnCerrar As MSForms.CommandButton
Private Const C_NN As Long = 14, C_REF As Long = 3, C_DIR As Long = 2, C_SIG As Long = 12
Private Const C_PROV As Long = 6, C_CANT As Long = 7, C_PARR As Long = 8, C_CORR As Long = 30

Private Sub UserForm_Initialize()
  Set wsD = ThisWorkbook.Worksheets("DEPOT")
  CargarBases
  Me.Caption = "Revisión de cobertura - Provincia / Cantón / Parroquia"
  Me.Width = 540: Me.Height = 360
  AddLbl "Pedido:", 10, 8, 55, True
  Set lblRef = AddLbl("", 70, 8, 450, True)
  AddLbl "Dirección:", 10, 26, 55
  Set lblDir = AddLbl("", 70, 26, 455): lblDir.Height = 30: lblDir.WordWrap = True
  AddLbl "Cliente:", 10, 58, 55
  Set lblCli = AddLbl("", 70, 58, 455): lblCli.ForeColor = RGB(90, 90, 90)
  AddLbl "Provincia:", 10, 80, 100
  Set cboProv = AddCbo(120, 78, 240)
  AddLbl "Cantón/Ciudad:", 10, 106, 100
  Set cboCant = AddCbo(120, 104, 240)
  AddLbl "Parroquia:", 10, 132, 100
  Set cboParr = AddCbo(120, 130, 400)
  AddLbl "Sugerencias:", 10, 158, 100
  Set lblSug = AddLbl("", 120, 158, 400): lblSug.Height = 44: lblSug.WordWrap = True: lblSug.ForeColor = RGB(150, 80, 0)
  AddLbl "Resultado:", 10, 206, 100
  Set lblInfo = AddLbl("", 120, 206, 400, True): lblInfo.Height = 44: lblInfo.WordWrap = True
  Dim P
  For Each P In Array("AZUAY", "BOLIVAR", "CAÑAR", "CARCHI", "CHIMBORAZO", "COTOPAXI", "EL ORO", "ESMERALDAS", "GALAPAGOS", "GUAYAS", "IMBABURA", "LOJA", "LOS RIOS", "MANABI", "MORONA SANTIAGO", "NAPO", "ORELLANA", "PASTAZA", "PICHINCHA", "SANTA ELENA", "SANTO DOMINGO DE LOS TSACHILAS", "SUCUMBIOS", "TUNGURAHUA", "ZAMORA CHINCHIPE")
    cboProv.AddItem P
  Next
  Set btnAplicar = PlaceBtn("Aplicar y siguiente", 10, 262, 120, RGB(112, 173, 71))
  Set btnSugerir = PlaceBtn("Sugerir cobertura cercana", 136, 262, 150, RGB(47, 117, 181))
  Set btnOmitir = PlaceBtn("Omitir", 292, 262, 70, RGB(150, 150, 150))
  Set btnAnterior = PlaceBtn("Anterior", 368, 262, 70, RGB(150, 150, 150))
  Set btnCerrar = PlaceBtn("Cerrar", 444, 262, 76, RGB(192, 80, 77))

  Dim soloRev As Boolean, ini As Long, lastD As Long, i As Long, e As String
  soloRev = gSoloRevisar: gSoloRevisar = False
  ini = gFilaInicio: gFilaInicio = 0
  Set rws = New Collection
  lastD = wsD.Cells(wsD.Rows.Count, C_REF).End(xlUp).Row
  For i = 2 To lastD
    e = UCase$(Trim$(TX(wsD.Cells(i, C_NN + 3))))
    If Not wsD.Rows(i).Hidden Then
      If e = "REVISAR" Or (e = "APROBADO" And Not soloRev) Or i = ini Then rws.Add i
    End If
  Next
  idx = 1
  If ini > 0 Then
    For i = 1 To rws.Count
      If rws(i) = ini Then idx = i: Exit For
    Next
  End If
  If rws.Count = 0 Then MsgBox "No hay pedidos pendientes de revisar.", vbInformation
  If soloRev Then Me.Caption = Me.Caption & "   [solo REVISAR: " & rws.Count & "]"
  CargarFila
End Sub

Private Function AddLbl(t As String, L As Single, tp As Single, Optional w As Single = 60, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = t: c.Left = L: c.Top = tp: c.Width = w: c.Font.Bold = bold
  Set AddLbl = c
End Function
Private Function AddCbo(L As Single, tp As Single, w As Single) As MSForms.ComboBox
  Dim c As MSForms.ComboBox: Set c = Me.Controls.Add("Forms.ComboBox.1")
  c.Left = L: c.Top = tp: c.Width = w: c.ListRows = 14: Set AddCbo = c
End Function
Private Function PlaceBtn(cap As String, L As Single, tp As Single, w As Single, col As Long) As MSForms.CommandButton
  Dim b As MSForms.CommandButton: Set b = Me.Controls.Add("Forms.CommandButton.1")
  b.Caption = cap: b.Left = L: b.Top = tp: b.Width = w: b.Height = 28
  b.BackColor = col: b.ForeColor = vbWhite: b.Font.Bold = True
  Set PlaceBtn = b
End Function

' ---------- listas dependientes ----------
Private Sub LlenarCantones()
  Dim c As Collection, it
  cboCant.Clear
  Set c = ListaCantones(Normaliza(cboProv.Text))
  For Each it In c: cboCant.AddItem it: Next
End Sub

Private Sub LlenarParroquias(ByVal sugs As String)
  Dim c As Collection, it, parts, s As String
  cboParr.Clear
  If Len(sugs) > 0 Then
    parts = Split(sugs, " | ")
    For Each it In parts
      s = Trim$(it)
      If Len(s) > 0 Then cboParr.AddItem s
    Next
  End If
  Set c = ListaParroquias(Normaliza(cboProv.Text), Normaliza(cboCant.Text))
  For Each it In c: cboParr.AddItem it: Next
End Sub

' "PARROQUIA [CANTON] (motivo)" -> parroquia y cantón
Private Sub ParseOpcion(ByVal s As String, ByRef parr As String, ByRef cant As String)
  Dim p1 As Long, p2 As Long
  parr = Trim$(s): cant = ""
  p1 = InStr(s, " ["): If p1 = 0 Then Exit Sub
  p2 = InStr(p1, s, "]")
  parr = Trim$(Left$(s, p1 - 1))
  If p2 > p1 Then cant = Trim$(Mid$(s, p1 + 2, p2 - p1 - 2))
End Sub

Private Sub CargarFila()
  If rws.Count = 0 Then Exit Sub
  If idx < 1 Then idx = 1
  If idx > rws.Count Then idx = rws.Count
  Dim r As Long: r = rws(idx)
  mCarga = True: mMotivo = ""
  lblRef.Caption = TX(wsD.Cells(r, C_REF)) & "   (fila " & r & " · " & idx & "/" & rws.Count & ")   " & _
                   TX(wsD.Cells(r, C_NN + 3)) & " - " & TX(wsD.Cells(r, C_NN + 6))
  lblDir.Caption = TX(wsD.Cells(r, C_DIR))
  lblCli.Caption = TX(wsD.Cells(r, C_PROV)) & " / " & TX(wsD.Cells(r, C_CANT)) & " / " & TX(wsD.Cells(r, C_PARR))
  cboProv.Text = TX(wsD.Cells(r, C_NN))
  LlenarCantones
  cboCant.Text = TX(wsD.Cells(r, C_NN + 1))
  LlenarParroquias TX(wsD.Cells(r, C_NN + 5))
  cboParr.Text = TX(wsD.Cells(r, C_NN + 2))
  lblSug.Caption = TX(wsD.Cells(r, C_NN + 5)) & "   " & TX(wsD.Cells(r, C_NN + 7))
  mCarga = False
  ActualizarInfo
End Sub

Private Sub ActualizarInfo()
  If rws.Count = 0 Then Exit Sub
  Dim P As String, nc As String, Q As String, parr As String, cant As String, ex, sg As String, ok As Boolean
  ParseOpcion cboParr.Text, parr, cant
  P = Normaliza(cboProv.Text): nc = Normaliza(cboCant.Text): Q = Normaliza(parr)
  If Len(P) = 0 Or Len(Q) = 0 Then lblInfo.Caption = "": Exit Sub
  ok = ExisteTriada(P, nc, Q) Or (P = "PICHINCHA" And nc = "QUITO" And InStr(Q, "DISTRITO METROPOLITANO") > 0)
  sg = SiglaFinal(P, nc, Q, parr)
  ex = DatosExt(P, nc, Q, " " & Normaliza(lblDir.Caption) & " ")
  lblInfo.Caption = IIf(ok, "[EN COBERTURA]  ", "[FUERA DE COBERTURA]  ") & "Sigla: " & sg & vbCrLf & _
                    "Gestor: " & ex(0) & "   Destino: " & ex(1) & IIf(Len(ex(2)) > 0, "   Trayecto: " & ex(2), "") & _
                    IIf(Len(ex(4)) > 0, vbCrLf & ex(4) & " -> " & ex(3), "")
  If ok Then lblInfo.ForeColor = RGB(0, 110, 0) Else lblInfo.ForeColor = RGB(192, 0, 0)
End Sub

' ---------- eventos de las listas ----------
Private Sub cboProv_Change()
  If mCarga Then Exit Sub
  mCarga = True
  LlenarCantones: cboCant.Text = ""
  LlenarParroquias "": cboParr.Text = ""
  mCarga = False
  ActualizarInfo
End Sub

Private Sub cboCant_Change()
  If mCarga Then Exit Sub
  Dim t As String: t = cboParr.Text
  mCarga = True
  LlenarParroquias ""
  cboParr.Text = t
  mCarga = False
  ActualizarInfo
End Sub

Private Sub cboParr_Change()
  If mCarga Then Exit Sub
  Dim parr As String, cant As String
  ParseOpcion cboParr.Text, parr, cant
  If Len(cant) > 0 Then
    mCarga = True
    cboCant.Text = cant
    LlenarParroquias TX(wsD.Cells(rws(idx), C_NN + 5))
    cboParr.Text = parr
    mCarga = False
  End If
  ActualizarInfo
End Sub

' ---------- botones ----------
Private Sub btnSugerir_Click()
  If rws.Count = 0 Then Exit Sub
  Dim sc, motivo As String, parr As String, cant As String
  ParseOpcion cboParr.Text, parr, cant
  sc = SugerirCercana(Normaliza(cboProv.Text), Normaliza(cboCant.Text), Normaliza(parr), " " & Normaliza(lblDir.Caption) & " ", motivo)
  If Not IsArray(sc) Then MsgBox "No hay una cobertura cercana para esta provincia/cantón.", vbInformation: Exit Sub
  mCarga = True
  cboCant.Text = sc(0)
  LlenarParroquias TX(wsD.Cells(rws(idx), C_NN + 5))
  cboParr.Text = sc(1)
  mCarga = False
  ActualizarInfo
  lblSug.Caption = "Sugerido: " & sc(1) & " [" & sc(0) & "] - " & motivo
  mMotivo = motivo
End Sub

Private Sub btnAplicar_Click()
  If rws.Count = 0 Then Exit Sub
  Dim r As Long: r = rws(idx)
  Dim parr As String, cant As String, P As String, nc As String, Q As String
  ParseOpcion cboParr.Text, parr, cant
  If Len(cant) = 0 Then cant = Trim$(cboCant.Text)
  P = Normaliza(cboProv.Text): nc = Normaliza(cant): Q = Normaliza(parr)
  If Len(P) = 0 Or Len(Q) = 0 Then MsgBox "Falta provincia o parroquia.", vbExclamation: Exit Sub
  If Not ExisteTriada(P, nc, Q) And Not (P = "PICHINCHA" And nc = "QUITO" And InStr(Q, "DISTRITO METROPOLITANO") > 0) Then
    If MsgBox("La combinación " & cboProv.Text & " / " & cant & " / " & parr & " NO está en COBERTURA." & vbCrLf & _
              "Sugerencia: usa 'Sugerir cobertura cercana'." & vbCrLf & vbCrLf & "¿Aplicar de todas formas?", _
              vbYesNo + vbExclamation, "Fuera de cobertura") <> vbYes Then Exit Sub
  End If
  Dim cambio As Boolean
  cambio = (Normaliza(TX(wsD.Cells(r, C_NN))) <> P Or Normaliza(TX(wsD.Cells(r, C_NN + 1))) <> nc Or Normaliza(TX(wsD.Cells(r, C_NN + 2))) <> Q)
  If Len(mMotivo) > 0 Then
    wsD.Cells(r, C_CORR).Value = "COBERTURA CERCANA (operario): " & mMotivo
  ElseIf cambio Then
    wsD.Cells(r, C_CORR).Value = "MANUAL OPERARIO"
  End If
  wsD.Cells(r, C_NN).Value = Replace(cboProv.Text, "-", " ")
  wsD.Cells(r, C_NN + 1).Value = cant
  wsD.Cells(r, C_NN + 2).Value = parr
  wsD.Cells(r, C_NN + 3).Value = "APROBADO"
  wsD.Cells(r, C_NN + 6).Value = "Aprobado operario " & Format(Now, "yyyy-mm-dd hh:nn")
  wsD.Cells(r, C_SIG).Value = SiglaFinal(P, nc, Q, parr)
  EscribirExtFila wsD, r
  nAprob = nAprob + 1
  If idx >= rws.Count Then
    MsgBox "Revisión terminada. Aprobados: " & nAprob & vbCrLf & "Siguiente paso: '5 Aplicar aprobados'.", vbInformation
    Unload Me
  Else
    idx = idx + 1: CargarFila
  End If
End Sub

Private Sub btnOmitir_Click()
  If idx < rws.Count Then idx = idx + 1
  CargarFila
End Sub
Private Sub btnAnterior_Click()
  If idx > 1 Then idx = idx - 1
  CargarFila
End Sub
Private Sub btnCerrar_Click()
  Unload Me
End Sub
