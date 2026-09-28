Option Explicit
' =====================================================================================
'  frmValidar (versión final) - reemplaza TODO el código del formulario frmValidar.
'  Los controles se crean al abrir. Redimensionable: arrastra el borde y todo se escala.
'   - "¿Por qué está en revisión?" en lenguaje claro
'   - Opciones sugeridas en una lista con el motivo de cada una (clic = usarla)
'   - "Si aplicas esta selección": cobertura, sigla, gestor, destino, trayecto, zona
'   - RECOMENDACIÓN: aplicar / no aplicar / atención zona peligrosa
'   - Botón "Buscar cobertura" (frmCobertura) que devuelve la selección aquí
' =====================================================================================
Private Const BASE_W As Single = 720
Private Const BASE_H As Single = 540
Private Const C_NN As Long = 14, C_REF As Long = 3, C_DIR As Long = 2, C_SIG As Long = 12
Private Const C_PROV As Long = 6, C_CANT As Long = 7, C_PARR As Long = 8, C_CORR As Long = 30

Private rws As Collection, idx As Long, wsD As Worksheet, nAprob As Long, mCarga As Boolean, mMotivo As String
Private mTitulo As String, mInW0 As Single, mInH0 As Single
Private lblRef As MSForms.Label, lblDir As MSForms.Label, lblCli As MSForms.Label
Private lblPorQue As MSForms.Label, lblInfo As MSForms.Label, lblReco As MSForms.Label
Private WithEvents cboProv As MSForms.ComboBox
Private WithEvents cboCant As MSForms.ComboBox
Private WithEvents cboParr As MSForms.ComboBox
Private WithEvents lstSug As MSForms.ListBox
Private WithEvents btnAplicar As MSForms.CommandButton
Private WithEvents btnSugerir As MSForms.CommandButton
Private WithEvents btnBuscar As MSForms.CommandButton
Private WithEvents btnOmitir As MSForms.CommandButton
Private WithEvents btnAnterior As MSForms.CommandButton
Private WithEvents btnCerrar As MSForms.CommandButton

Private Sub UserForm_Initialize()
  Set wsD = ThisWorkbook.Worksheets("DEPOT")
  CargarBases
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H
  Dim t As MSForms.Label
  Set lblRef = AddLbl("", 10, 8, 690, True): lblRef.Font.Size = 10
  AddLbl "Dirección:", 10, 28, 80, True
  Set lblDir = AddLbl("", 95, 28, 610): lblDir.Height = 28: lblDir.WordWrap = True
  AddLbl "Cliente escribió:", 10, 58, 85, True
  Set lblCli = AddLbl("", 95, 58, 610): lblCli.ForeColor = RGB(90, 90, 90)
  Set lblPorQue = AddLbl("", 10, 76, 695): lblPorQue.Height = 44: lblPorQue.WordWrap = True
  lblPorQue.BackColor = RGB(255, 242, 204): lblPorQue.BorderStyle = fmBorderStyleSingle: lblPorQue.ForeColor = RGB(128, 64, 0)

  AddLbl "Provincia:", 10, 130, 85, True
  Set cboProv = AddCbo(95, 127, 250)
  AddLbl "Cantón/Ciudad:", 10, 154, 85, True
  Set cboCant = AddCbo(95, 151, 250)
  AddLbl "Parroquia:", 10, 178, 85, True
  Set cboParr = AddCbo(95, 175, 420)
  Set t = AddLbl("Tu selección. Puedes escribir, elegir de la lista o hacer clic en una opción sugerida.", 360, 130, 345)
  t.Height = 40: t.WordWrap = True: t.ForeColor = RGB(120, 120, 120)

  Set t = AddLbl("OPCIONES SUGERIDAS (clic para usar)", 10, 204, 400, True): t.ForeColor = RGB(48, 84, 150)
  AddLbl "PARROQUIA", 13, 220, 190, True
  AddLbl "CANTÓN", 203, 220, 130, True
  AddLbl "POR QUÉ SE SUGIERE", 333, 220, 370, True
  Set lstSug = Me.Controls.Add("Forms.ListBox.1")
  lstSug.Left = 10: lstSug.Top = 234: lstSug.Width = 695: lstSug.Height = 88
  lstSug.ColumnCount = 3: lstSug.ColumnWidths = "190;130;360": lstSug.Font.Size = 8

  Set t = AddLbl("SI APLICAS ESTA SELECCIÓN", 10, 328, 400, True): t.ForeColor = RGB(48, 84, 150)
  Set lblInfo = AddLbl("", 10, 342, 695): lblInfo.Height = 70: lblInfo.WordWrap = True
  lblInfo.BorderStyle = fmBorderStyleSingle: lblInfo.BackColor = RGB(248, 248, 248)
  Set lblReco = AddLbl("", 10, 416, 695, True): lblReco.Height = 30: lblReco.WordWrap = True: lblReco.Font.Size = 9

  Dim P
  For Each P In Array("AZUAY", "BOLIVAR", "CAÑAR", "CARCHI", "CHIMBORAZO", "COTOPAXI", "EL ORO", "ESMERALDAS", "GALAPAGOS", "GUAYAS", "IMBABURA", "LOJA", "LOS RIOS", "MANABI", "MORONA SANTIAGO", "NAPO", "ORELLANA", "PASTAZA", "PICHINCHA", "SANTA ELENA", "SANTO DOMINGO DE LOS TSACHILAS", "SUCUMBIOS", "TUNGURAHUA", "ZAMORA CHINCHIPE")
    cboProv.AddItem P
  Next
  Set btnAplicar = PlaceBtn("Aplicar y siguiente", 10, 454, 130, RGB(112, 173, 71), "Guarda la selección como APROBADO y pasa al siguiente pedido.")
  Set btnSugerir = PlaceBtn("Sugerir cobertura cercana", 146, 454, 150, RGB(47, 117, 181), _
                            "Propone: misma parroquia mal escrita, parroquia principal del cantón, ciudad principal/secundaria o capital provincial.")
  Set btnBuscar = PlaceBtn("Buscar cobertura...", 302, 454, 120, RGB(0, 112, 192), _
                           "Abre el buscador de COBERTURA filtrado por este pedido para elegir la parroquia exacta.")
  Set btnOmitir = PlaceBtn("Omitir", 428, 454, 80, RGB(150, 150, 150), "Pasa al siguiente sin guardar.")
  Set btnAnterior = PlaceBtn("Anterior", 514, 454, 80, RGB(150, 150, 150), "Vuelve al pedido anterior.")
  Set btnCerrar = PlaceBtn("Cerrar", 600, 454, 105, RGB(192, 80, 77), "Cierra el validador.")

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
  mTitulo = "Revisión de cobertura HYCITE" & IIf(soloRev, " - solo REVISAR (" & rws.Count & ")", "")
  Me.Caption = mTitulo
  mInW0 = Me.InsideWidth: mInH0 = Me.InsideHeight
  HacerRedimensionable mTitulo
  AjustarAPantalla Me, BASE_W, BASE_H, 0.8
  If rws.Count = 0 Then MsgBox "No hay pedidos pendientes de revisar.", vbInformation
  CargarFila
End Sub

Private Sub UserForm_Resize()
  If mInW0 = 0 Or mInH0 = 0 Then Exit Sub
  Dim f As Double
  f = Me.InsideWidth / mInW0
  If Me.InsideHeight / mInH0 < f Then f = Me.InsideHeight / mInH0
  If f < 0.4 Then f = 0.4
  If f > 3 Then f = 3
  If Abs(Me.Zoom - f * 100) > 1 Then Me.Zoom = f * 100
End Sub

Private Function AddLbl(t As String, L As Single, tp As Single, Optional w As Single = 60, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = t: c.Left = L: c.Top = tp: c.Width = w: c.Height = 14: c.Font.Bold = bold
  Set AddLbl = c
End Function
Private Function AddCbo(L As Single, tp As Single, w As Single) As MSForms.ComboBox
  Dim c As MSForms.ComboBox: Set c = Me.Controls.Add("Forms.ComboBox.1")
  c.Left = L: c.Top = tp: c.Width = w: c.Height = 18: c.ListRows = 14: Set AddCbo = c
End Function
Private Function PlaceBtn(cap As String, L As Single, tp As Single, w As Single, col As Long, tip As String) As MSForms.CommandButton
  Dim b As MSForms.CommandButton: Set b = Me.Controls.Add("Forms.CommandButton.1")
  b.Caption = cap: b.Left = L: b.Top = tp: b.Width = w: b.Height = 32
  b.BackColor = col: b.ForeColor = vbWhite: b.Font.Bold = True: b.ControlTipText = tip
  Set PlaceBtn = b
End Function

' ---------- textos de ayuda ----------
Private Function Explicacion(ByVal accion As String) As String
  Select Case True
    Case InStr(accion, "Confirmar parroquia") > 0
      Explicacion = "La parroquia leída al final de la dirección no coincide exactamente con COBERTURA. Confirma la parroquia correcta."
    Case InStr(accion, "Verificar cantón") > 0
      Explicacion = "Esa parroquia existe en varios cantones de la provincia y el cliente no indicó cuál. Elige el cantón que corresponde a la dirección."
    Case InStr(accion, "otra provincia") > 0
      Explicacion = "La parroquia que escribió el cliente pertenece a otra provincia. Revisa si está mal la provincia o la parroquia."
    Case InStr(accion, "Fuera de cobertura") > 0
      Explicacion = "La parroquia del cliente NO está en COBERTURA. El sistema propone la cobertura más cercana (ver el motivo abajo). Acéptala o busca otra con 'Buscar cobertura'."
    Case InStr(accion, "sugerida") > 0
      Explicacion = "La parroquia está mal escrita o no se reconoce. Elige la opción correcta de la lista."
    Case InStr(accion, "de la dirección") > 0
      Explicacion = "La parroquia del cliente no existe; el sistema tomó la parroquia que aparece escrita en la dirección. Confírmala."
    Case InStr(accion, "Quito") > 0
      Explicacion = "Dirección en Quito sin parroquia reconocible. Por regla se asigna DISTRITO METROPOLITANO DE QUITO con sigla CALDERON (QUI). Si la dirección menciona un sector, busca su parroquia."
    Case InStr(accion, "duplicado") > 0
      Explicacion = "El nombre de la parroquia existe en varias provincias o cantones. Verifica que sea el correcto."
    Case InStr(accion, "Zona peligrosa") > 0
      Explicacion = "El destino está en una zona peligrosa para el courier. Confirma entrega C.O.D / retiro en oficina TMC."
    Case InStr(accion, "Aprobado") > 0
      Explicacion = "Ya fue aprobado. Puedes cambiarlo y volver a aplicar."
    Case Else
      Explicacion = "Revisa la provincia, el cantón y la parroquia con la dirección."
  End Select
End Function

' ---------- listas ----------
Private Sub LlenarCantones()
  Dim it
  cboCant.Clear
  For Each it In ListaCantones(Normaliza(cboProv.Text)): cboCant.AddItem it: Next
End Sub

Private Sub LlenarParroquias()
  Dim it
  cboParr.Clear
  For Each it In ListaParroquias(Normaliza(cboProv.Text), Normaliza(cboCant.Text)): cboParr.AddItem it: Next
End Sub

' "PARROQUIA [CANTON] (motivo)" -> parroquia, cantón, motivo
Private Sub ParseOpcion(ByVal s As String, ByRef parr As String, ByRef cant As String, Optional ByRef motivo As String)
  Dim p1 As Long, p2 As Long, p3 As Long
  parr = Trim$(s): cant = "": motivo = ""
  p1 = InStr(s, " ["): If p1 = 0 Then Exit Sub
  p2 = InStr(p1, s, "]")
  parr = Trim$(Left$(s, p1 - 1))
  If p2 > p1 Then
    cant = Trim$(Mid$(s, p1 + 2, p2 - p1 - 2))
    p3 = InStr(p2, s, "(")
    If p3 > 0 Then motivo = Trim$(Replace(Mid$(s, p3 + 1), ")", ""))
  End If
End Sub

Private Sub AgregarSug(ByVal parr As String, ByVal cant As String, ByVal porque As String)
  Dim i As Long
  If Len(parr) = 0 Then Exit Sub
  For i = 0 To lstSug.ListCount - 1
    If Normaliza(lstSug.List(i, 0)) = Normaliza(parr) And Normaliza(lstSug.List(i, 1)) = Normaliza(cant) Then Exit Sub
  Next
  lstSug.AddItem parr
  lstSug.List(lstSug.ListCount - 1, 1) = cant
  lstSug.List(lstSug.ListCount - 1, 2) = porque
End Sub

Private Sub LlenarSugerencias(ByVal r As Long)
  Dim parts, it, parr As String, cant As String, mot As String, sc, motivo As String
  lstSug.Clear
  ' 1) propuesta del sistema
  AgregarSug TX(wsD.Cells(r, C_NN + 2)), TX(wsD.Cells(r, C_NN + 1)), "Propuesta del sistema: " & TX(wsD.Cells(r, C_CORR))
  ' 2) sugerencias guardadas en la validación
  If Len(TX(wsD.Cells(r, C_NN + 5))) > 0 Then
    parts = Split(TX(wsD.Cells(r, C_NN + 5)), " | ")
    For Each it In parts
      ParseOpcion CStr(it), parr, cant, mot
      AgregarSug parr, cant, IIf(Len(mot) > 0, mot, "Nombre parecido a lo que escribió el cliente")
    Next
  End If
  ' 3) cobertura cercana calculada ahora
  sc = SugerirCercana(Normaliza(TX(wsD.Cells(r, C_NN))), Normaliza(TX(wsD.Cells(r, C_NN + 1))), Normaliza(TX(wsD.Cells(r, C_PARR))), _
                      " " & Normaliza(TX(wsD.Cells(r, C_DIR))) & " ", motivo)
  If IsArray(sc) Then AgregarSug CStr(sc(1)), CStr(sc(0)), motivo
  ' 4) parroquia principal (cabecera) del cantón propuesto
  If ExisteTriada(Normaliza(TX(wsD.Cells(r, C_NN))), Normaliza(TX(wsD.Cells(r, C_NN + 1))), Normaliza(TX(wsD.Cells(r, C_NN + 1)))) Then
    AgregarSug TX(wsD.Cells(r, C_NN + 1)), TX(wsD.Cells(r, C_NN + 1)), "Parroquia principal del cantón (cabecera)"
  End If
End Sub

Private Sub CargarFila()
  If rws.Count = 0 Then Exit Sub
  If idx < 1 Then idx = 1
  If idx > rws.Count Then idx = rws.Count
  Dim r As Long: r = rws(idx)
  mCarga = True: mMotivo = ""
  lblRef.Caption = "Pedido " & TX(wsD.Cells(r, C_REF)) & "   (fila " & r & " · " & idx & " de " & rws.Count & ")   " & _
                   TX(wsD.Cells(r, C_NN + 3)) & "   Destinatario: " & TX(wsD.Cells(r, 5))
  lblDir.Caption = TX(wsD.Cells(r, C_DIR))
  lblCli.Caption = TX(wsD.Cells(r, C_PROV)) & "  /  " & TX(wsD.Cells(r, C_CANT)) & "  /  " & TX(wsD.Cells(r, C_PARR)) & _
                   IIf(Len(TX(wsD.Cells(r, 29))) > 0, "      (original: " & TX(wsD.Cells(r, 29)) & ")", "")
  lblPorQue.Caption = "¿POR QUÉ ESTÁ EN REVISIÓN?  " & TX(wsD.Cells(r, C_NN + 6)) & vbCrLf & Explicacion(TX(wsD.Cells(r, C_NN + 6))) & _
                      IIf(Len(TX(wsD.Cells(r, C_NN + 4))) > 0, "   [Evidencia: " & TX(wsD.Cells(r, C_NN + 4)) & "]", "")
  cboProv.Text = TX(wsD.Cells(r, C_NN))
  LlenarCantones
  cboCant.Text = TX(wsD.Cells(r, C_NN + 1))
  LlenarParroquias
  cboParr.Text = TX(wsD.Cells(r, C_NN + 2))
  LlenarSugerencias r
  mCarga = False
  ActualizarInfo
End Sub

Private Sub ActualizarInfo()
  If rws.Count = 0 Then Exit Sub
  Dim P As String, nc As String, Q As String, ex, sg As String, ok As Boolean, r As Long, igualProp As Boolean
  r = rws(idx)
  P = Normaliza(cboProv.Text): nc = Normaliza(cboCant.Text): Q = Normaliza(cboParr.Text)
  If Len(P) = 0 Or Len(Q) = 0 Then
    lblInfo.Caption = "Completa provincia y parroquia.": lblReco.Caption = "": Exit Sub
  End If
  ok = ExisteTriada(P, nc, Q) Or (P = "PICHINCHA" And nc = "QUITO" And InStr(Q, "DISTRITO METROPOLITANO") > 0)
  sg = SiglaFinal(P, nc, Q, cboParr.Text)
  ex = DatosExt(P, nc, Q, " " & Normaliza(lblDir.Caption) & " ")
  igualProp = (P = Normaliza(TX(wsD.Cells(r, C_NN))) And nc = Normaliza(TX(wsD.Cells(r, C_NN + 1))) And Q = Normaliza(TX(wsD.Cells(r, C_NN + 2))))
  lblInfo.Caption = cboProv.Text & " / " & cboCant.Text & " / " & cboParr.Text & "   ->   " & IIf(ok, "EN COBERTURA", "FUERA DE COBERTURA") & vbCrLf & _
                    "Sigla: " & sg & "      Gestor asignado: " & ex(0) & " (destino " & ex(1) & ")" & IIf(Len(ex(2)) > 0, "      Trayecto: " & ex(2), "") & vbCrLf & _
                    "Gestor en cobertura (Q): " & IIf(Len(ex(5)) > 0, ex(5), "-") & "      Sugerido (R): " & IIf(Len(ex(6)) > 0, ex(6), "-") & vbCrLf & _
                    "Tipo de entrega: " & ex(3) & IIf(Len(ex(4)) > 0, "      " & ex(4), "") & IIf(Len(mMotivo) > 0, vbCrLf & "Elegida por: " & mMotivo, "")
  If Not ok Then
    lblReco.ForeColor = RGB(192, 0, 0)
    lblReco.Caption = "NO RECOMENDADO: esta combinación no está en COBERTURA. Usa una opción sugerida, 'Sugerir cobertura cercana' o 'Buscar cobertura'."
  ElseIf Left$(ex(4), 4) = "ZONA" Then
    lblReco.ForeColor = RGB(191, 90, 0)
    lblReco.Caption = "ATENCIÓN: zona peligrosa. Entrega sugerida: " & ex(3) & ". Si estás de acuerdo, aplica."
  ElseIf igualProp Then
    lblReco.ForeColor = RGB(0, 110, 0)
    lblReco.Caption = "RECOMENDADO: coincide con la propuesta del sistema y está en cobertura. Pulsa 'Aplicar y siguiente'."
  Else
    lblReco.ForeColor = RGB(0, 110, 0)
    lblReco.Caption = "VÁLIDO: está en cobertura (cambio manual respecto a la propuesta del sistema). Pulsa 'Aplicar y siguiente' si es correcto."
  End If
End Sub

' Llamado desde frmCobertura al elegir "Asignar al pedido"
Public Sub AsignarCobertura(ByVal prov As String, ByVal cant As String, ByVal parr As String, ByVal motivo As String)
  If rws.Count = 0 Then Exit Sub
  mCarga = True
  cboProv.Text = prov
  LlenarCantones
  cboCant.Text = cant
  LlenarParroquias
  cboParr.Text = parr
  mCarga = False
  mMotivo = motivo
  ActualizarInfo
  Me.Show vbModeless
End Sub

' ---------- eventos ----------
Private Sub cboProv_Change()
  If mCarga Then Exit Sub
  mCarga = True
  LlenarCantones: cboCant.Text = ""
  cboParr.Clear: cboParr.Text = ""
  mCarga = False
  mMotivo = ""
  ActualizarInfo
End Sub

Private Sub cboCant_Change()
  If mCarga Then Exit Sub
  Dim t As String: t = cboParr.Text
  mCarga = True
  LlenarParroquias
  cboParr.Text = t
  mCarga = False
  mMotivo = ""
  ActualizarInfo
End Sub

Private Sub cboParr_Change()
  If mCarga Then Exit Sub
  mMotivo = ""
  ActualizarInfo
End Sub

Private Sub lstSug_Click()
  If lstSug.ListIndex < 0 Then Exit Sub
  Dim parr As String, cant As String
  parr = lstSug.List(lstSug.ListIndex, 0): cant = lstSug.List(lstSug.ListIndex, 1)
  mCarga = True
  If Len(cant) > 0 Then cboCant.Text = cant
  LlenarParroquias
  cboParr.Text = parr
  mCarga = False
  mMotivo = lstSug.List(lstSug.ListIndex, 2)
  ActualizarInfo
End Sub

Private Sub btnSugerir_Click()
  If rws.Count = 0 Then Exit Sub
  Dim sc, motivo As String
  sc = SugerirCercana(Normaliza(cboProv.Text), Normaliza(cboCant.Text), Normaliza(cboParr.Text), " " & Normaliza(lblDir.Caption) & " ", motivo)
  If Not IsArray(sc) Then MsgBox "No hay una cobertura cercana para esta provincia/cantón. Usa 'Buscar cobertura'.", vbInformation: Exit Sub
  AgregarSug CStr(sc(1)), CStr(sc(0)), motivo
  mCarga = True
  cboCant.Text = sc(0)
  LlenarParroquias
  cboParr.Text = sc(1)
  mCarga = False
  mMotivo = motivo
  ActualizarInfo
End Sub

Private Sub btnBuscar_Click()
  If rws.Count = 0 Then Exit Sub
  gCtxFila = rws(idx): gCtxPedido = TX(wsD.Cells(gCtxFila, C_REF))
  gCtxProv = cboProv.Text: gCtxCant = cboCant.Text: gCtxParr = TX(wsD.Cells(gCtxFila, C_PARR))
  gCtxDir = lblDir.Caption
  AbrirCobertura
End Sub

Private Sub btnAplicar_Click()
  If rws.Count = 0 Then Exit Sub
  Dim r As Long: r = rws(idx)
  Dim parr As String, cant As String, P As String, nc As String, Q As String, cambio As Boolean
  parr = Trim$(cboParr.Text): cant = Trim$(cboCant.Text)
  P = Normaliza(cboProv.Text): nc = Normaliza(cant): Q = Normaliza(parr)
  If Len(P) = 0 Or Len(Q) = 0 Then MsgBox "Falta provincia o parroquia.", vbExclamation: Exit Sub
  If Not ExisteTriada(P, nc, Q) And Not (P = "PICHINCHA" And nc = "QUITO" And InStr(Q, "DISTRITO METROPOLITANO") > 0) Then
    If MsgBox("La combinación " & cboProv.Text & " / " & cant & " / " & parr & " NO está en COBERTURA." & vbCrLf & _
              "Sugerencia: usa una opción sugerida o 'Buscar cobertura'." & vbCrLf & vbCrLf & "¿Aplicar de todas formas?", _
              vbYesNo + vbExclamation, "Fuera de cobertura") <> vbYes Then Exit Sub
    LogP "Fila " & r & " pedido " & TX(wsD.Cells(r, C_REF)) & ": aprobado FUERA DE COBERTURA por el operario (" & cboProv.Text & "/" & cant & "/" & parr & ")", "AVISO"
  End If
  cambio = (Normaliza(TX(wsD.Cells(r, C_NN))) <> P Or Normaliza(TX(wsD.Cells(r, C_NN + 1))) <> nc Or Normaliza(TX(wsD.Cells(r, C_NN + 2))) <> Q)
  If Len(mMotivo) > 0 And cambio Then
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
  LogP "Fila " & r & " pedido " & TX(wsD.Cells(r, C_REF)) & ": aprobado por el operario -> " & cboProv.Text & "/" & cant & "/" & parr & IIf(cambio, " (modificado)", " (propuesta del sistema)")
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
