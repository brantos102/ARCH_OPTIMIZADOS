Option Explicit
' =====================================================================================
'  frmLote - TABLA DE SUGERENCIAS (confirmación en lote)
'  Insertar > UserForm, nombre frmLote, pegar este código. Los controles se crean solos.
'
'  Para qué: revisar los pendientes uno por uno es lento cuando la mayoría son el mismo
'  caso (una letra cambiada, el nombre de la cobertura, la cabecera del cantón). Aquí se
'  ven TODOS a la vez con su propuesta y una CONFIANZA calculada, se marcan los que estén
'  bien y se aprueban de una sola vez.
'
'  Qué NO hace: no inventa nada. Aplica exactamente la propuesta que ya calculó la
'  validación (columnas N/O/P) y escribe igual que el validador de uno en uno: APROBADO,
'  sigla, columnas V:AD y registro. Lo que no se marca, se queda como está.
' =====================================================================================
Private Const BASE_W As Single = 1000
Private Const BASE_H As Single = 600
Private Const C_NN As Long = 14, C_REF As Long = 3, C_SIG As Long = 12
Private Const C_PROV As Long = 6, C_CANT As Long = 7, C_PARR As Long = 8, C_CORR As Long = 30

Private wsD As Worksheet, mCerrando As Boolean, mCarga As Boolean
Private mTitulo As String, mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean
Private mFila() As Long, mConf() As Long, mN As Long
Private lblTit As MSForms.Label, lblCnt As MSForms.Label, lblPie As MSForms.Label
Private WithEvents lst As MSForms.ListBox
Private WithEvents cboConf As MSForms.ComboBox
Private WithEvents txtBuscar As MSForms.TextBox
Private WithEvents bAlta As MSForms.CommandButton
Private WithEvents bTodo As MSForms.CommandButton
Private WithEvents bNada As MSForms.CommandButton
Private WithEvents bAplicar As MSForms.CommandButton
Private WithEvents bUno As MSForms.CommandButton
Private WithEvents bCerrar As MSForms.CommandButton

Private Sub UserForm_Initialize()
  Dim t As MSForms.Label, f
  mTitulo = "Sugerencias del día - confirmar en lote"
  Me.Caption = mTitulo
  Set wsD = ThisWorkbook.Worksheets("DEPOT")
  CargarBases
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H

  Set lblTit = AddLbl("", 10, 8, 960, True): lblTit.Font.Size = 10: lblTit.ForeColor = RGB(48, 84, 150)
  Set t = AddLbl("Marca los que estén bien y pulsa 'Aprobar marcados'. Lo que no marques se queda como está.", 10, 26, 960)
  t.ForeColor = RGB(90, 90, 90)

  AddLbl "Confianza:", 10, 50, 56, True
  Set cboConf = AddCbo(68, 47, 120)
  For Each f In Array("TODAS", "ALTA", "MEDIA", "BAJA"): cboConf.AddItem f: Next
  AddLbl "Buscar:", 200, 50, 40, True
  Set txtBuscar = Me.Controls.Add("Forms.TextBox.1")
  txtBuscar.Left = 242: txtBuscar.Top = 47: txtBuscar.Width = 180: txtBuscar.Height = 18
  txtBuscar.ControlTipText = "Pedido, provincia, cantón o parroquia."
  Set bAlta = PlaceBtn("Marcar las de confianza ALTA", 436, 45, 180, RGB(112, 173, 71), _
      "Marca las propuestas inequívocas: están en COBERTURA y el nombre coincide o cambia muy poco.")
  Set bTodo = PlaceBtn("Marcar todo", 622, 45, 90, RGB(120, 120, 120), "Marca todas las de la lista.")
  Set bNada = PlaceBtn("Quitar marcas", 716, 45, 90, RGB(120, 120, 120), "Quita todas las marcas.")

  Set lblCnt = AddLbl("", 10, 76, 960, True)
  AddLbl "PEDIDO", 13, 94, 70, True
  AddLbl "EL CLIENTE ESCRIBIÓ", 85, 94, 210, True
  AddLbl "PROPUESTA DEL SISTEMA", 297, 94, 210, True
  AddLbl "POR QUÉ", 509, 94, 300, True
  AddLbl "CONFIANZA", 811, 94, 80, True
  AddLbl "FILA", 893, 94, 50, True
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = 10: lst.Top = 108: lst.Width = 960: lst.Height = 400: lst.Font.Size = 8
  lst.ColumnCount = 6: lst.ColumnWidths = "70;210;210;300;80;50"
  lst.MultiSelect = fmMultiSelectExtended
  lst.ControlTipText = "Ctrl + clic o Shift + clic para marcar varios. Doble clic abre ese pedido en el validador de uno en uno."

  Set lblPie = AddLbl("", 10, 514, 960): lblPie.Height = 26: lblPie.WordWrap = True: lblPie.ForeColor = RGB(128, 64, 0)
  Set bAplicar = PlaceBtn("Aprobar marcados", 10, 544, 200, RGB(112, 173, 71), _
      "Escribe la propuesta como APROBADO en los pedidos marcados. Después hay que pulsar '5 Aplicar aprobados'.")
  Set bUno = PlaceBtn("Revisar uno por uno", 216, 544, 160, RGB(47, 117, 181), "Abre el validador de siempre con los pendientes.")
  Set bCerrar = PlaceBtn("Cerrar", 870, 544, 100, RGB(192, 80, 77), "Cierra sin cambiar nada.")

  Cargar
  mCarga = True: cboConf.ListIndex = 0: mCarga = False
  Filtrar
  mLay = CapturarLayout(Me, mCW, mCH): mEsc = 1
  HacerRedimensionable mTitulo
  AjustarAPantalla Me, BASE_W, BASE_H, 0.94
  Reescalar
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  mCerrando = True
End Sub
Private Sub UserForm_Resize()
  Reescalar
End Sub
Private Sub Reescalar()
  If Not IsArray(mLay) Or mEscalando Or mCerrando Then Exit Sub
  Dim f As Double
  f = EscalaAjuste(Me, mCW, mCH)
  If Abs(f - mEsc) < 0.01 Then Exit Sub
  mEscalando = True: mEsc = f
  EscalarLayout Me, mLay, f, mCW, mCH
  mEscalando = False
End Sub

Private Function AddLbl(t As String, L As Single, tp As Single, Optional w As Single = 60, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = t: c.Left = L: c.Top = tp: c.Width = w: c.Height = 14: c.Font.Bold = bold
  Set AddLbl = c
End Function
Private Function AddCbo(L As Single, tp As Single, w As Single) As MSForms.ComboBox
  Dim c As MSForms.ComboBox: Set c = Me.Controls.Add("Forms.ComboBox.1")
  c.Left = L: c.Top = tp: c.Width = w: c.Height = 18: c.Style = fmStyleDropDownList: c.ListRows = 8
  Set AddCbo = c
End Function
Private Function PlaceBtn(cap As String, L As Single, tp As Single, w As Single, col As Long, tip As String) As MSForms.CommandButton
  Dim b As MSForms.CommandButton: Set b = Me.Controls.Add("Forms.CommandButton.1")
  b.Caption = cap: b.Left = L: b.Top = tp: b.Width = w: b.Height = 28
  b.BackColor = col: b.ForeColor = vbWhite: b.Font.Bold = True: b.Font.Size = 8
  b.ControlTipText = tip: b.TakeFocusOnClick = False: b.WordWrap = True
  Set PlaceBtn = b
End Function

' ---------- datos ----------
Private Sub Cargar()
  Dim lr As Long, i As Long, n As Long
  lr = wsD.Cells(wsD.Rows.Count, C_REF).End(xlUp).Row
  ReDim mFila(1 To 1): ReDim mConf(1 To 1)
  If lr < 2 Then Exit Sub
  ReDim mFila(1 To lr): ReDim mConf(1 To lr)
  For i = 2 To lr
    If Not wsD.Rows(i).Hidden Then
      If UCase$(Trim$(TX(wsD.Cells(i, C_NN + 3)))) = "REVISAR" Then
        n = n + 1: mFila(n) = i: mConf(n) = Confianza(i)
      End If
    End If
  Next
  mN = n
  OrdenarPorConfianza
  lblTit.Caption = mN & " pedido(s) pendientes de revisar"
End Sub

' ALTA 3 / MEDIA 2 / BAJA 1. Alta = la propuesta está en COBERTURA y el nombre coincide
' con lo que escribió el cliente o cambia muy poco (error de escritura, nombre oficial).
Private Function Confianza(ByVal r As Long) As Long
  Dim P As String, nc As String, Q As String, qCli As String, d As Long, tipo As String
  P = Normaliza(TX(wsD.Cells(r, C_NN)))
  nc = Normaliza(TX(wsD.Cells(r, C_NN + 1)))
  Q = Normaliza(TX(wsD.Cells(r, C_NN + 2)))
  qCli = Normaliza(TX(wsD.Cells(r, C_PARR)))
  tipo = UCase$(TX(wsD.Cells(r, C_CORR)))
  If Len(P) = 0 Or Len(Q) = 0 Then Confianza = 1: Exit Function
  If Not ExisteTriada(P, nc, Q) Then Confianza = 1: Exit Function
  If InStr(tipo, "ZONA") > 0 Then Confianza = 1: Exit Function
  d = Lev(qCli, Q)
  If d = 0 Then Confianza = 3: Exit Function
  If InStr(tipo, "CORRECCION AUTOMATICA") > 0 Or InStr(tipo, "NOMBRE SEGUN COBERTURA") > 0 Then Confianza = 3: Exit Function
  If InStr(tipo, "DMQ") > 0 Then Confianza = 3: Exit Function
  If d <= 2 Then Confianza = 3: Exit Function
  If InStr(tipo, "CABECERA") > 0 Or InStr(tipo, "COBERTURA CERCANA") > 0 Then Confianza = 2: Exit Function
  Confianza = 2
End Function

Private Function TextoConf(ByVal c As Long) As String
  Select Case c
    Case 3: TextoConf = "ALTA"
    Case 2: TextoConf = "MEDIA"
    Case Else: TextoConf = "BAJA"
  End Select
End Function

Private Sub OrdenarPorConfianza()
  Dim i As Long, j As Long, t As Long
  If mN < 2 Then Exit Sub
  For i = 1 To mN - 1
    For j = i + 1 To mN
      If mConf(j) > mConf(i) Then
        t = mConf(i): mConf(i) = mConf(j): mConf(j) = t
        t = mFila(i): mFila(i) = mFila(j): mFila(j) = t
      End If
    Next
  Next
End Sub

' ---------- lista ----------
Private Sub Filtrar()
  Dim i As Long, r As Long, q As String, fc As String, cli As String, prop As String, pq As String
  If mCerrando Or lst Is Nothing Then Exit Sub
  q = UCase$(Trim$(txtBuscar.Text)): fc = cboConf.Text
  mCarga = True
  lst.Clear
  For i = 1 To mN
    r = mFila(i)
    If Len(fc) = 0 Or fc = "TODAS" Or fc = TextoConf(mConf(i)) Then
      cli = TX(wsD.Cells(r, C_PROV)) & " / " & TX(wsD.Cells(r, C_CANT)) & " / " & TX(wsD.Cells(r, C_PARR))
      prop = TX(wsD.Cells(r, C_NN)) & " / " & TX(wsD.Cells(r, C_NN + 1)) & " / " & TX(wsD.Cells(r, C_NN + 2))
      pq = TX(wsD.Cells(r, C_CORR))
      If Len(pq) = 0 Then pq = TX(wsD.Cells(r, C_NN + 6))
      If Len(q) = 0 Or InStr(UCase$(TX(wsD.Cells(r, C_REF)) & " " & cli & " " & prop & " " & pq), q) > 0 Then
        lst.AddItem TX(wsD.Cells(r, C_REF))
        lst.List(lst.ListCount - 1, 1) = cli
        lst.List(lst.ListCount - 1, 2) = prop
        lst.List(lst.ListCount - 1, 3) = pq
        lst.List(lst.ListCount - 1, 4) = TextoConf(mConf(i))
        lst.List(lst.ListCount - 1, 5) = CStr(r)
      End If
    End If
  Next
  mCarga = False
  Contar
End Sub

Private Sub Contar()
  Dim i As Long, n As Long
  If mCerrando Or lblCnt Is Nothing Or bAplicar Is Nothing Then Exit Sub
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then n = n + 1
  Next
  lblCnt.Caption = lst.ListCount & " en la lista   ·   " & n & " marcado(s)"
  If n = 0 Then
    bAplicar.Caption = "Aprobar marcados"
    lblPie.Caption = "Nada marcado. Marca con Ctrl + clic, o usa 'Marcar las de confianza ALTA'."
  Else
    bAplicar.Caption = "Aprobar los " & n & " marcados"
    lblPie.Caption = "Se escribirá la propuesta del sistema en esos " & n & " pedidos y quedarán como APROBADO. " & _
                     "Después hay que pulsar '5 Aplicar aprobados' para llevarlos a la data."
  End If
End Sub

Private Sub cboConf_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub
Private Sub txtBuscar_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub
Private Sub lst_Change()
  If mCerrando Or mCarga Then Exit Sub
  Contar
End Sub
Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  If mCerrando Or lst.ListIndex < 0 Then Exit Sub
  gFilaInicio = Val(lst.List(lst.ListIndex, 5))
  gSoloRevisar = True
  mCerrando = True
  Unload Me
  frmValidar.Show vbModal
End Sub

Private Sub bTodo_Click()
  Dim i As Long
  If mCerrando Then Exit Sub
  mCarga = True
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = True: Next
  mCarga = False: Contar
End Sub
Private Sub bNada_Click()
  Dim i As Long
  If mCerrando Then Exit Sub
  mCarga = True
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = False: Next
  mCarga = False: Contar
End Sub
Private Sub bAlta_Click()
  Dim i As Long
  If mCerrando Then Exit Sub
  mCarga = True
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = (lst.List(i, 4) = "ALTA"): Next
  mCarga = False: Contar
End Sub

Private Sub bUno_Click()
  gSoloRevisar = True: gFilaInicio = 0
  mCerrando = True
  Unload Me
  frmValidar.Show vbModal
End Sub

Private Sub bCerrar_Click()
  mCerrando = True
  Unload Me
End Sub

' ---------- aprobar en lote ----------
Private Sub bAplicar_Click()
  Dim i As Long, r As Long, n As Long, nOk As Long, P As String, nc As String, Q As String, parr As String
  If mCerrando Then Exit Sub
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then n = n + 1
  Next
  If n = 0 Then MsgBox "No hay nada marcado.", vbInformation, "Sugerencias": Exit Sub
  If MsgBox("Se aprobarán " & n & " pedido(s) con la propuesta del sistema." & vbCrLf & vbCrLf & _
            "Lo que no está marcado NO se toca." & vbCrLf & _
            "Después hay que pulsar '5 Aplicar aprobados'." & vbCrLf & vbCrLf & _
            "¿Continuar?", vbYesNo + vbQuestion, "Aprobar en lote") <> vbYes Then Exit Sub
  Application.ScreenUpdating = False
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then
      r = Val(lst.List(i, 5))
      If r >= 2 Then
        P = Normaliza(TX(wsD.Cells(r, C_NN)))
        nc = Normaliza(TX(wsD.Cells(r, C_NN + 1)))
        Q = Normaliza(TX(wsD.Cells(r, C_NN + 2)))
        parr = NombreCobertura(P, nc, Q, TX(wsD.Cells(r, C_NN + 2)))
        wsD.Cells(r, C_NN + 2).Value = parr
        wsD.Cells(r, C_NN + 3).Value = "APROBADO"
        wsD.Cells(r, C_NN + 6).Value = "Aprobado en tabla de sugerencias " & Format(Now, "yyyy-mm-dd hh:nn")
        wsD.Cells(r, C_SIG).Value = SiglaFinal(P, nc, Q, parr)
        EscribirExtFila wsD, r
        LogP "Fila " & r & " pedido " & TX(wsD.Cells(r, C_REF)) & ": aprobado en lote (confianza " & lst.List(i, 4) & _
             ") -> " & TX(wsD.Cells(r, C_NN)) & "/" & TX(wsD.Cells(r, C_NN + 1)) & "/" & parr
        nOk = nOk + 1
      End If
    End If
  Next
  Application.ScreenUpdating = True
  LogP "SUGERENCIAS: " & nOk & " pedido(s) aprobados desde la tabla de sugerencias"
  Cargar
  Filtrar
  MsgBox nOk & " pedido(s) aprobados." & vbCrLf & vbCrLf & _
         "Quedan " & mN & " por revisar." & vbCrLf & _
         "Siguiente paso: '5 Aplicar aprobados'.", vbInformation, "Sugerencias"
  If mN = 0 Then mCerrando = True: Unload Me
End Sub
