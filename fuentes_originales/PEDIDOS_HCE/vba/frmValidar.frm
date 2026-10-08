Option Explicit
Private rws As Collection, idx As Long, wsD As Worksheet, nAprob As Long
Private lblRef As MSForms.Label, lblDir As MSForms.Label, lblSug As MSForms.Label
Private cboProv As MSForms.ComboBox, cboCant As MSForms.ComboBox, cboParr As MSForms.ComboBox
Private WithEvents btnAplicar As MSForms.CommandButton
Attribute btnAplicar.VB_VarHelpID = -1
Private WithEvents btnOmitir As MSForms.CommandButton
Attribute btnOmitir.VB_VarHelpID = -1
Private WithEvents btnAnterior As MSForms.CommandButton
Attribute btnAnterior.VB_VarHelpID = -1
Private WithEvents btnCerrar As MSForms.CommandButton
Attribute btnCerrar.VB_VarHelpID = -1
Private Const C_NN As Long = 14, C_REF As Long = 3, C_DIR As Long = 2, C_SIG As Long = 12

Private Sub UserForm_Initialize()
  Set wsD = ThisWorkbook.Worksheets("DEPOT")
  CargarBases
  Me.Caption = "Revisión — Provincia / Cantón / Parroquia"
  Me.Width = 470: Me.Height = 285
  AddLbl "Pedido:", 10, 8, 55, True
  Set lblRef = AddLbl("", 68, 8, 385, True)
  AddLbl "Dirección:", 10, 28, 55
  Set lblDir = AddLbl("", 68, 28, 388): lblDir.Height = 30: lblDir.WordWrap = True
  AddLbl "Provincia:", 10, 68, 100
  Set cboProv = AddCbo(120, 66, 220)
  AddLbl "Cantón/Ciudad:", 10, 98, 100
  Set cboCant = AddCbo(120, 96, 220)
  AddLbl "Parroquia:", 10, 128, 100
  Set cboParr = AddCbo(120, 126, 260)
  AddLbl "Sugerencias:", 10, 158, 100
  Set lblSug = AddLbl("", 120, 158, 330): lblSug.Height = 34: lblSug.WordWrap = True: lblSug.ForeColor = RGB(150, 80, 0)
  Dim P
  For Each P In Array("AZUAY", "BOLIVAR", "CAÑAR", "CARCHI", "CHIMBORAZO", "COTOPAXI", "EL ORO", "ESMERALDAS", "GALAPAGOS", "GUAYAS", "IMBABURA", "LOJA", "LOS RIOS", "MANABI", "MORONA SANTIAGO", "NAPO", "ORELLANA", "PASTAZA", "PICHINCHA", "SANTA ELENA", "SANTO DOMINGO DE LOS TSACHILAS", "SUCUMBIOS", "TUNGURAHUA", "ZAMORA CHINCHIPE")
    cboProv.AddItem P
  Next
  Set btnAplicar = PlaceBtn("Aplicar y siguiente", 12, 205, RGB(112, 173, 71))
  Set btnOmitir = PlaceBtn("Omitir", 140, 205, RGB(150, 150, 150))
  Set btnAnterior = PlaceBtn("Anterior", 270, 205, RGB(150, 150, 150))
  Set btnCerrar = PlaceBtn("Cerrar", 405, 205, RGB(192, 80, 77))
  Set rws = New Collection
  Dim lastD As Long, i As Long, e As String
  lastD = wsD.Cells(wsD.Rows.Count, C_REF).End(xlUp).Row
  For i = 2 To lastD
    e = UCase$(Trim$(wsD.Cells(i, C_NN + 3).Value & ""))
    If (e = "REVISAR" Or e = "APROBADO") And Not wsD.Rows(i).Hidden Then rws.Add i
  Next
  If rws.Count = 0 Then MsgBox "No hay filas visibles para revisar.", vbInformation
  idx = 1: CargarFila
End Sub

Private Function AddLbl(t As String, L As Single, tp As Single, Optional w As Single = 60, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = t: c.Left = L: c.Top = tp: c.Width = w: c.Font.bold = bold
  Set AddLbl = c
End Function
Private Function AddCbo(L As Single, tp As Single, w As Single) As MSForms.ComboBox
  Dim c As MSForms.ComboBox: Set c = Me.Controls.Add("Forms.ComboBox.1")
  c.Left = L: c.Top = tp: c.Width = w: Set AddCbo = c
End Function
Private Function PlaceBtn(cap As String, L As Single, tp As Single, col As Long) As MSForms.CommandButton
  Dim b As MSForms.CommandButton: Set b = Me.Controls.Add("Forms.CommandButton.1")
  b.Caption = cap: b.Left = L: b.Top = tp: b.Width = IIf(L > 380, 55, 120): b.Height = 26
  b.BackColor = col: b.ForeColor = vbWhite: b.Font.bold = True
  Set PlaceBtn = b
End Function

Private Sub CargarFila()
  If rws.Count = 0 Then Exit Sub
  If idx < 1 Then idx = 1
  If idx > rws.Count Then idx = rws.Count
  Dim r As Long: r = rws(idx)
  Dim prov As String, cant As String, parr As String
  prov = wsD.Cells(r, C_NN).Value: cant = wsD.Cells(r, C_NN + 1).Value: parr = wsD.Cells(r, C_NN + 2).Value
  lblRef.Caption = wsD.Cells(r, C_REF).Value & "   (fila " & r & " · " & idx & "/" & rws.Count & ")"
  lblDir.Caption = wsD.Cells(r, C_DIR).Value
  cboProv.Value = prov
  cboParr.Clear
  If Len(parr) > 0 Then cboParr.AddItem parr
  Dim s As String, parts, it, nm As String
  s = wsD.Cells(r, C_NN + 5).Value
  If Len(s) > 0 Then
    parts = Split(s, "|")
    For Each it In parts
      nm = Trim$(it)
      If InStr(nm, " [") > 0 Then nm = Trim$(Left$(nm, InStr(nm, " [") - 1))
      If Len(nm) > 0 Then cboParr.AddItem nm
    Next
  End If
  cboParr.Value = parr
  cboCant.Clear
  If Len(cant) > 0 Then cboCant.AddItem cant
  If Len(parr) > 0 And parr <> cant Then cboCant.AddItem parr
  cboCant.Value = cant
  lblSug.Caption = wsD.Cells(r, C_NN + 5).Value & "   " & wsD.Cells(r, C_NN + 7).Value
End Sub

Private Sub btnAplicar_Click()
  If rws.Count = 0 Then Exit Sub
  Dim r As Long: r = rws(idx)
  wsD.Cells(r, C_NN).Value = Replace(cboProv.Value, "-", " ")
  wsD.Cells(r, C_NN + 1).Value = cboCant.Value
  wsD.Cells(r, C_NN + 2).Value = cboParr.Value
  wsD.Cells(r, C_NN + 3).Value = "APROBADO"
  wsD.Cells(r, C_NN + 6).Value = "Aprobado operario " & Format(Now, "yyyy-mm-dd hh:nn")
  wsD.Cells(r, C_SIG).Value = SiglaFinal(Normaliza(cboProv.Value), Normaliza(cboCant.Value), Normaliza(cboParr.Value), cboParr.Value)
  nAprob = nAprob + 1
  If idx >= rws.Count Then
    MsgBox "Revisión terminada. Aprobados: " & nAprob & vbCrLf & "Usa '3) Aplicar aprobados'.", vbInformation
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
