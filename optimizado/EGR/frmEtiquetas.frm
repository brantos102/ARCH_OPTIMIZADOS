Option Explicit
' =====================================================================================
'  frmEtiquetas - Imprimir etiquetas Zebra (10 x 5 cm).
'  Insertar > UserForm, nombre frmEtiquetas, pegar este código.
'  Funciona así: se abre con TODOS los pedidos que tienen destino, se filtra la lista
'  y se imprime exactamente lo que quedó en la lista. No hay que marcar nada.
'  Rendimiento: la lista se arma en memoria y se vuelca de una sola vez.
' =====================================================================================
Private Const BASE_W As Single = 860
Private Const BASE_H As Single = 500
Private Const TITULO As String = "Imprimir etiquetas HYCITE"
Private Const PX As Single = 356        ' esquina de la etiqueta de muestra
Private Const PY As Single = 58
Private Const ES As Single = 0.6        ' 800 x 400 puntos Zebra -> 480 x 240 en pantalla

Private mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean
Private mCerrando As Boolean, mCarga As Boolean
Private mFila() As Long, mPed() As String, mDest() As String, mNom() As String, mParr() As String, mEst() As String
Private mN As Long, mIdx() As Long, mNIdx As Long
Private lblInfo As MSForms.Label, lblCnt As MSForms.Label, fondo As MSForms.Label
Private lblNum As MSForms.Label, lblDest As MSForms.Label, lblParr As MSForms.Label, lblNom As MSForms.Label
Private barras(1 To 90) As MSForms.Label
Private WithEvents lst As MSForms.ListBox
Private WithEvents cboEstado As MSForms.ComboBox
Private WithEvents chkPRO As MSForms.CheckBox
Private WithEvents chkGYE As MSForms.CheckBox
Private WithEvents chkUIO As MSForms.CheckBox
Private WithEvents chkGPS As MSForms.CheckBox
Private WithEvents bTodo As MSForms.CommandButton
Private WithEvents bNada As MSForms.CommandButton
Private WithEvents txtBuscar As MSForms.TextBox
Private cboImp As MSForms.ComboBox
Private WithEvents bQuitar As MSForms.CommandButton
Private WithEvents bImprimir As MSForms.CommandButton
Private WithEvents bZpl As MSForms.CommandButton
Private WithEvents bCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
  Dim t As MSForms.Label, i As Long, f
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H

  Set lblInfo = NL("", 10, 8, 830, 14, True): lblInfo.ForeColor = RGB(48, 84, 150)

  ' ----- filtros -----
  NL "Etiqueta:", 10, 30, 46
  Set cboEstado = NC(58, 27, 118)
  For Each f In Array("PENDIENTES + REIMPRIMIR", "PENDIENTES", "REIMPRIMIR", "IMPRESAS", "TODAS"): cboEstado.AddItem f: Next
  NL "Buscar:", 184, 30, 38
  Set txtBuscar = Me.Controls.Add("Forms.TextBox.1")
  txtBuscar.Left = 224: txtBuscar.Top = 27: txtBuscar.Width = 158: txtBuscar.Height = 18
  txtBuscar.ControlTipText = "Pedido, destinatario o parroquia. Escribe y la lista se filtra sola."
  Set bQuitar = NB("Quitar filtros", 390, 26, 84, 20, RGB(120, 120, 120))

  ' ----- destinos: se pueden combinar (un lote PRO + GYE + UIO a la vez) -----
  NL "Destinos:", 10, 53, 46, 14, True
  Set chkPRO = NK("PRO", 58, 51)
  Set chkGYE = NK("GYE", 118, 51)
  Set chkUIO = NK("UIO", 178, 51)
  Set chkGPS = NK("GPS", 238, 51)
  Set t = NL("(marca los que quieras imprimir juntos)", 300, 53, 174): t.ForeColor = RGB(120, 120, 120)

  ' ----- lista -----
  Set lblCnt = NL("", 10, 72, 336, 14, True)
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = 10: lst.Top = 88: lst.Width = 336: lst.Height = 284: lst.Font.Size = 8
  lst.ColumnCount = 4: lst.ColumnWidths = "72;34;150;74"
  lst.MultiSelect = fmMultiSelectExtended
  lst.ControlTipText = "Sin seleccionar nada se imprime TODA la lista. Si seleccionas (Ctrl o Shift + clic) se imprime solo lo seleccionado."
  Set bTodo = NB("Seleccionar todo", 10, 376, 100, 18, RGB(120, 120, 120))
  Set bNada = NB("Quitar selección", 114, 376, 100, 18, RGB(120, 120, 120))
  Set t = NL("Sin selección se imprime toda la lista. Ctrl + clic o Shift + clic para elegir solo algunos. Un clic muestra la etiqueta.", 10, 398, 336)
  t.Height = 28: t.WordWrap = True: t.ForeColor = RGB(120, 120, 120)

  ' ----- vista previa -----
  Set t = NL("VISTA PREVIA (10 x 5 cm)", PX, PY - 18, 400, 14, True): t.ForeColor = RGB(48, 84, 150)
  Set fondo = NL("", PX, PY, 800 * ES, 400 * ES)
  fondo.BackColor = vbWhite: fondo.BorderStyle = fmBorderStyleSingle
  For i = 1 To 90
    Set barras(i) = NL("", PX, PY, 1, 1): barras(i).BackColor = vbBlack: barras(i).Visible = False
  Next
  Set lblNum = TL(PX + 140 * ES, PY + 150 * ES, 400 * ES, 13): lblNum.TextAlign = fmTextAlignCenter
  Set lblDest = TL(PX + 380 * ES, PY + 186 * ES, 390 * ES, 30, True): lblDest.TextAlign = fmTextAlignRight
  Set lblParr = TL(PX + 330 * ES, PY + 288 * ES, 440 * ES, 11): lblParr.TextAlign = fmTextAlignRight
  Set lblNom = TL(PX + 34 * ES, PY + 224 * ES, 330 * ES, 16, True): lblNom.WordWrap = True: lblNom.Height = 46

  ' ----- impresora y acciones -----
  NL "Impresora:", PX, 300, 56
  Set cboImp = NC(PX + 58, 297, 322)
  Set t = NL("Se guarda la impresora elegida: la próxima vez ya aparece seleccionada.", PX, 318, 390)
  t.ForeColor = RGB(120, 120, 120)
  Set bImprimir = NB("Imprimir", PX, 342, 244, 40, RGB(0, 128, 96))
  bImprimir.Font.Size = 11
  Set bZpl = NB("Guardar ZPL (prueba)", PX + 252, 342, 128, 40, RGB(120, 120, 120))
  Set bCancel = NB("Cerrar", PX + 252, 390, 128, 24, RGB(192, 80, 77))
  Set t = NL("Si la etiqueta sale clara o corrida, ajusta ETIQ_OSCURIDAD y ETIQ_OFFSET en la hoja CONFIG_EGR.", PX, 390, 244)
  t.Height = 28: t.WordWrap = True: t.ForeColor = RGB(120, 120, 120)

  CargarPedidos
  CargarImpresoras
  mCarga = True
  cboEstado.ListIndex = 0
  mCarga = False
  Filtrar
  mLay = CapturarLayout(Me, mCW, mCH): mEsc = 1
  HacerRedimensionable TITULO
  AjustarAPantalla Me, BASE_W, BASE_H, 0.9
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
  mEscalando = True
  mEsc = f
  EscalarLayout Me, mLay, f, mCW, mCH
  mEscalando = False
  If Not lst Is Nothing Then If lst.ListIndex >= 0 Then Muestra lst.ListIndex
End Sub

' ---------- controles ----------
Private Function NL(cap As String, x As Single, y As Single, w As Single, Optional h As Single = 14, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = h: c.Font.Bold = bold
  Set NL = c
End Function
Private Function TL(x As Single, y As Single, w As Single, tam As Single, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Left = x: c.Top = y: c.Width = w: c.Height = tam * 1.7: c.BackStyle = fmBackStyleTransparent
  c.Font.Name = "Arial": c.Font.Size = tam: c.Font.Bold = bold: c.ForeColor = vbBlack
  Set TL = c
End Function
Private Function NK(cap As String, x As Single, y As Single) As MSForms.CheckBox
  Dim c As MSForms.CheckBox: Set c = Me.Controls.Add("Forms.CheckBox.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = 56: c.Height = 16: c.Value = True
  c.Font.Bold = True
  Set NK = c
End Function

Private Function NC(x As Single, y As Single, w As Single) As MSForms.ComboBox
  Dim c As MSForms.ComboBox: Set c = Me.Controls.Add("Forms.ComboBox.1")
  c.Left = x: c.Top = y: c.Width = w: c.Height = 18: c.ListRows = 12: c.Style = fmStyleDropDownList
  Set NC = c
End Function
Private Function NB(cap As String, x As Single, y As Single, w As Single, h As Single, col As Long) As MSForms.CommandButton
  Dim c As MSForms.CommandButton: Set c = Me.Controls.Add("Forms.CommandButton.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = h
  c.BackColor = col: c.ForeColor = vbWhite: c.Font.Bold = True: c.Font.Size = 8: c.TakeFocusOnClick = False
  Set NB = c
End Function

' ---------- datos ----------
Private Sub CargarPedidos()
  Dim wsD As Worksheet, wsE As Worksheet, lr As Long, v, e, i As Long, n As Long, dest As String, est As String, nSin As Long
  Set wsD = ThisWorkbook.Worksheets(HDAT)
  Set wsE = ThisWorkbook.Worksheets("ETIQUETAS")
  lr = UltimaFilaDatos()
  ReDim mFila(1 To 1): ReDim mPed(1 To 1): ReDim mDest(1 To 1): ReDim mNom(1 To 1): ReDim mParr(1 To 1): ReDim mEst(1 To 1)
  If lr < 2 Then Exit Sub
  v = wsD.Range(wsD.Cells(1, 1), wsD.Cells(lr, D_PARRSP)).Value
  e = wsE.Range(wsE.Cells(1, 1), wsE.Cells(lr, 14)).Value
  ReDim mFila(1 To lr): ReDim mPed(1 To lr): ReDim mDest(1 To lr): ReDim mNom(1 To lr): ReDim mParr(1 To lr): ReDim mEst(1 To lr)
  For i = 2 To lr
    If Len(TXE(v(i, D_PED))) > 0 Then
      dest = UCase$(TXE(v(i, D_DEST)))
      If Len(dest) = 0 Then
        nSin = nSin + 1
      Else
        If UCase$(TXE(e(i, 6))) = "OK" Then
          If Len(TXE(e(i, 13))) > 0 And UCase$(TXE(e(i, 13))) <> dest Then est = "REIMPRIMIR" Else est = "IMPRESA"
        Else
          est = "PENDIENTE"
        End If
        n = n + 1
        mFila(n) = i: mPed(n) = TXE(v(i, D_PED)): mDest(n) = dest
        mNom(n) = TXE(v(i, D_NOM)): mEst(n) = est
        mParr(n) = TXE(v(i, D_PARRSP)): If Len(mParr(n)) = 0 Then mParr(n) = TXE(v(i, D_PARR))
      End If
    End If
  Next
  mN = n
  lblInfo.Caption = mN & " pedido(s) con destino" & _
    IIf(nSin > 0, "   ·   " & nSin & " sin destino (fuera de cobertura TMS) no se pueden imprimir", "")
End Sub

Private Sub CargarImpresoras()
  Dim c As Collection, it, guard As String, i As Long
  guard = Cfg("IMPRESORA_ZEBRA")
  Set c = ListaImpresoras()
  For Each it In c
    cboImp.AddItem it
    If it = guard Then cboImp.ListIndex = cboImp.ListCount - 1
  Next
  If cboImp.ListIndex < 0 Then
    For i = 0 To cboImp.ListCount - 1
      If InStr(1, cboImp.List(i), "ZD", vbTextCompare) > 0 Or InStr(1, cboImp.List(i), "ZEBRA", vbTextCompare) > 0 _
         Or InStr(1, cboImp.List(i), "ZDESIGNER", vbTextCompare) > 0 Then cboImp.ListIndex = i: Exit For
    Next
  End If
End Sub

' ---------- filtro ----------
Private Function Pasa(ByVal k As Long, ByVal est As String, ByVal des As String, ByVal q As String) As Boolean
  Select Case est
    Case "PENDIENTES": If mEst(k) <> "PENDIENTE" Then Exit Function
    Case "REIMPRIMIR": If mEst(k) <> "REIMPRIMIR" Then Exit Function
    Case "PENDIENTES + REIMPRIMIR": If mEst(k) = "IMPRESA" Then Exit Function
    Case "IMPRESAS": If mEst(k) <> "IMPRESA" Then Exit Function
  End Select
  ' des trae los destinos marcados entre barras: "|PRO|GYE|". Vacío = todos.
  If Len(des) > 0 Then
    If InStr(1, des, "|" & mDest(k) & "|", vbTextCompare) = 0 Then Exit Function
  End If
  If Len(q) > 0 Then
    If InStr(1, mPed(k) & " " & mNom(k) & " " & mParr(k) & " " & mDest(k), q, vbTextCompare) = 0 Then Exit Function
  End If
  Pasa = True
End Function

' Destinos marcados, en texto, para el aviso de confirmación.
Private Function TextoDestinos() As String
  Dim s As String
  If chkPRO.Value Then s = s & "PRO "
  If chkGYE.Value Then s = s & "GYE "
  If chkUIO.Value Then s = s & "UIO "
  If chkGPS.Value Then s = s & "GPS "
  If Len(s) = 0 Then TextoDestinos = "(ninguno)" Else TextoDestinos = Trim$(s)
End Function

' Destinos marcados, entre barras. Si están los cuatro (o ninguno) devuelve vacío = todos.
Private Function DestinosMarcados() As String
  Dim s As String
  If chkPRO Is Nothing Then Exit Function
  If chkPRO.Value Then s = s & "PRO|"
  If chkGYE.Value Then s = s & "GYE|"
  If chkUIO.Value Then s = s & "UIO|"
  If chkGPS.Value Then s = s & "GPS|"
  If Len(s) = 0 Then DestinosMarcados = "|(ninguno)|": Exit Function   ' sin destinos marcados: lista vacía
  If s = "PRO|GYE|UIO|GPS|" Then Exit Function                          ' los cuatro = todos
  DestinosMarcados = "|" & s
End Function

Private Sub Filtrar()
  Dim k As Long, est As String, des As String, q As String, arr() As String, n As Long
  If mCerrando Or lst Is Nothing Or cboEstado Is Nothing Then Exit Sub
  est = cboEstado.Text: des = DestinosMarcados(): q = Trim$(txtBuscar.Text)
  ReDim mIdx(1 To IIf(mN > 0, mN, 1)): mNIdx = 0
  For k = 1 To mN
    If Pasa(k, est, des, q) Then mNIdx = mNIdx + 1: mIdx(mNIdx) = k
  Next
  Ordenar                                   ' agrupa el lote por destino y dentro por pedido
  mCarga = True
  lst.Clear
  If mNIdx > 0 Then
    ReDim arr(0 To mNIdx - 1, 0 To 3)
    For n = 1 To mNIdx
      k = mIdx(n)
      arr(n - 1, 0) = mPed(k): arr(n - 1, 1) = mDest(k): arr(n - 1, 2) = mNom(k): arr(n - 1, 3) = mEst(k)
    Next
    lst.List = arr
    ' no se fija ListIndex: en una lista de selección múltiple eso dejaría marcada la
    ' primera fila y parecería que hay selección cuando no la hay
  End If
  mCarga = False
  If mNIdx > 0 Then Muestra 0
  Contar
End Sub

Private Function NSeleccionadas() As Long
  Dim i As Long, n As Long
  If lst Is Nothing Then Exit Function
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then n = n + 1
  Next
  NSeleccionadas = n
End Function

Private Sub Contar()
  Dim nSel As Long
  If mCerrando Or lblCnt Is Nothing Or bImprimir Is Nothing Then Exit Sub
  nSel = NSeleccionadas()
  If mNIdx = 0 Then
    lblCnt.Caption = "0 etiquetas en la lista"
    bImprimir.Caption = "No hay etiquetas que imprimir"
  ElseIf nSel > 0 Then
    lblCnt.Caption = mNIdx & " en la lista   ·   " & nSel & " seleccionada(s)"
    bImprimir.Caption = "Imprimir las " & nSel & " seleccionada(s)"
  Else
    lblCnt.Caption = mNIdx & " etiqueta(s) en la lista   ·   sin selección = se imprimen todas"
    bImprimir.Caption = "Imprimir las " & mNIdx & " de la lista"
  End If
End Sub

' Lo que se va a imprimir: lo seleccionado si hay selección; si no, toda la lista filtrada.
Private Function Listadas() As Collection
  Dim c As New Collection, n As Long, hay As Boolean
  hay = (NSeleccionadas() > 0)
  For n = 1 To mNIdx
    If Not hay Then
      c.Add mFila(mIdx(n))
    ElseIf n - 1 < lst.ListCount Then
      If lst.Selected(n - 1) Then c.Add mFila(mIdx(n))
    End If
  Next
  Set Listadas = c
End Function

' ---------- vista previa ----------
Private Sub Muestra(ByVal pos As Long)
  Dim k As Long, m As String, i As Long, x As Single, w As Single, nBar As Long, e As Double, ox As Single, oy As Single
  If mCerrando Or fondo Is Nothing Or lblNum Is Nothing Then Exit Sub
  If pos < 0 Or pos >= mNIdx Then Exit Sub
  k = mIdx(pos + 1)
  e = mEsc: If e <= 0 Then e = 1
  ox = fondo.Left: oy = fondo.Top
  m = Code128Modulos(mPed(k))
  x = ox + 140 * ES * e
  For i = 1 To 90: barras(i).Visible = False: Next
  For i = 1 To Len(m)
    w = Val(Mid$(m, i, 1)) * 4 * ES * e * 0.5
    If i Mod 2 = 1 Then
      nBar = nBar + 1
      If nBar <= 90 Then
        barras(nBar).Left = x: barras(nBar).Top = oy + 26 * ES * e
        barras(nBar).Width = w: barras(nBar).Height = 112 * ES * e
        barras(nBar).Visible = True
      End If
    End If
    x = x + w
  Next
  lblNum.Caption = mPed(k)
  lblDest.Caption = mDest(k)
  lblParr.Caption = UCase$(mParr(k))
  lblNom.Caption = UCase$(mNom(k))
End Sub

' ---------- eventos ----------
Private Sub cboEstado_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub
Private Sub chkPRO_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub
Private Sub chkGYE_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub
Private Sub chkUIO_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub
Private Sub chkGPS_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub

Private Sub bTodo_Click()
  Dim i As Long
  If mCerrando Or lst Is Nothing Then Exit Sub
  mCarga = True
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = True: Next
  mCarga = False
  Contar
End Sub

Private Sub bNada_Click()
  Dim i As Long
  If mCerrando Or lst Is Nothing Then Exit Sub
  mCarga = True
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = False: Next
  mCarga = False
  Contar
End Sub

' Ordena las filas filtradas por destino y, dentro de cada destino, por número de pedido.
Private Sub Ordenar()
  Dim i As Long, j As Long, h As Long, t As Long, cl() As String, ct As String
  If mNIdx < 2 Then Exit Sub
  ReDim cl(1 To mNIdx)
  For i = 1 To mNIdx: cl(i) = mDest(mIdx(i)) & Chr(1) & mPed(mIdx(i)): Next
  h = 1
  Do While h < mNIdx \ 3: h = h * 3 + 1: Loop
  Do While h >= 1
    For i = h + 1 To mNIdx
      t = mIdx(i): ct = cl(i)
      j = i
      Do While j > h
        If cl(j - h) <= ct Then Exit Do
        mIdx(j) = mIdx(j - h): cl(j) = cl(j - h)
        j = j - h
      Loop
      mIdx(j) = t: cl(j) = ct
    Next
    h = (h - 1) \ 3
  Loop
End Sub
Private Sub txtBuscar_Change()
  If mCarga Then Exit Sub
  Filtrar
End Sub
Private Sub bQuitar_Click()
  If mCerrando Then Exit Sub
  mCarga = True
  cboEstado.ListIndex = 0: txtBuscar.Text = ""
  chkPRO.Value = True: chkGYE.Value = True: chkUIO.Value = True: chkGPS.Value = True
  mCarga = False
  Filtrar
End Sub

Private Sub lst_Change()
  If mCerrando Or mCarga Or lst Is Nothing Then Exit Sub
  If lst.ListIndex >= 0 Then Muestra lst.ListIndex
  Contar
End Sub

Private Sub bImprimir_Click()
  Dim c As Collection, n As Long, det As String
  If mCerrando Then Exit Sub
  Set c = Listadas()
  If c.Count = 0 Then MsgBox "No hay etiquetas en la lista. Cambia los filtros.", vbInformation: Exit Sub
  If cboImp.ListIndex < 0 Then MsgBox "Elige la impresora Zebra.", vbExclamation: Exit Sub
  det = "Etiqueta: " & cboEstado.Text & "   ·   Destinos: " & TextoDestinos()
  If Len(Trim$(txtBuscar.Text)) > 0 Then det = det & "   ·   Buscar: " & Trim$(txtBuscar.Text)
  det = det & vbCrLf & IIf(NSeleccionadas() > 0, "Se imprime SOLO lo seleccionado.", "Se imprime toda la lista.")
  If MsgBox("Imprimir " & c.Count & " etiqueta(s) en:" & vbCrLf & cboImp.Text & vbCrLf & vbCrLf & det, _
            vbYesNo + vbQuestion, "Imprimir etiquetas") <> vbYes Then Exit Sub
  SetCfg "IMPRESORA_ZEBRA", cboImp.Text
  Application.Cursor = xlWait
  n = ImprimirFilas(c, cboImp.Text, False)
  Application.Cursor = xlDefault
  If n > 0 Then
    mCerrando = True
    MsgBox n & " etiqueta(s) enviadas.", vbInformation
    Unload Me
  Else
    MsgBox "No se pudo imprimir. Revisa que la Zebra esté encendida y conectada (detalle en el registro).", vbExclamation
  End If
End Sub

Private Sub bZpl_Click()
  Dim c As Collection, n As Long
  If mCerrando Then Exit Sub
  Set c = Listadas()
  If c.Count = 0 Then MsgBox "No hay etiquetas en la lista.", vbInformation: Exit Sub
  n = ImprimirFilas(c, "", True)
  If n > 0 Then MsgBox n & " etiqueta(s) guardadas en la carpeta de exportes (archivo .zpl)." & vbCrLf & _
                        "Se puede ver en labelary.com con 8 dpmm y 3,94 x 1,97 pulgadas.", vbInformation
End Sub

Private Sub bCancel_Click()
  mCerrando = True
  Unload Me
End Sub
