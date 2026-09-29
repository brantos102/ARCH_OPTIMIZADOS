Option Explicit
' =====================================================================================
'  frmEtiquetas - Vista previa e impresión de etiquetas Zebra (10 x 5 cm).
'  Insertar > UserForm, nombre frmEtiquetas, pegar este código. Lo abre frmEGR.
'  - Lista de pedidos a imprimir (marcados = se imprimen); clic = vista previa
'  - Impresora: se elige una vez y queda guardada (CONFIG_EGR)
' =====================================================================================
Private Const BASE_W As Single = 800
Private Const BASE_H As Single = 470
Private Const TITULO As String = "Etiquetas HYCITE - vista previa e impresión"
Private Const PX As Single = 300        ' posición de la etiqueta de muestra
Private Const PY As Single = 40
Private Const ES As Single = 0.5        ' 800 x 400 puntos Zebra -> 400 x 200 en pantalla

Private mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean
Private mFilas() As Long, mN As Long
Private lblInfo As MSForms.Label, fondo As MSForms.Label
Private lblNum As MSForms.Label, lblDest As MSForms.Label, lblParr As MSForms.Label, lblNom As MSForms.Label, lblApe As MSForms.Label
Private barras(1 To 80) As MSForms.Label
Private WithEvents lst As MSForms.ListBox
Private cboImp As MSForms.ComboBox
Private WithEvents bTodos As MSForms.CommandButton
Private WithEvents bNinguno As MSForms.CommandButton
Private WithEvents bImprimir As MSForms.CommandButton
Private WithEvents bZpl As MSForms.CommandButton
Private WithEvents bCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
  Dim wsD As Worksheet, f, i As Long, nSin As Long, t As MSForms.Label, colImp As Collection, guard As String, it
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H
  Set lblInfo = NL("", 10, 8, 770, 14, True): lblInfo.ForeColor = RGB(48, 84, 150)
  NL "Pedidos (marcados = se imprimen; clic = ver)", 10, 26, 280
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = 10: lst.Top = 40: lst.Width = 280: lst.Height = 330: lst.Font.Size = 8
  lst.ColumnCount = 3: lst.ColumnWidths = "70;40;160"
  lst.MultiSelect = fmMultiSelectMulti: lst.ListStyle = fmListStyleOption
  Set bTodos = NB("Marcar todos", 10, 374, 136, 22, RGB(120, 120, 120))
  Set bNinguno = NB("Desmarcar todos", 154, 374, 136, 22, RGB(120, 120, 120))
  ' etiqueta de muestra
  Set t = NL("VISTA PREVIA (10 x 5 cm, blanco y negro)", PX, PY - 16, 400, 14, True): t.ForeColor = RGB(48, 84, 150)
  Set fondo = NL("", PX, PY, 800 * ES, 400 * ES)
  fondo.BackColor = vbWhite: fondo.BorderStyle = fmBorderStyleSingle
  For i = 1 To 80
    Set barras(i) = NL("", PX, PY, 1, 1): barras(i).BackColor = vbBlack: barras(i).Visible = False
  Next
  Set lblNum = TL(PX + 40 * ES, PY + 148 * ES, 440 * ES, 11, False)
  Set lblDest = TL(PX + 520 * ES, PY + 34 * ES, 260 * ES, 40, True): lblDest.TextAlign = fmTextAlignCenter
  Set lblParr = TL(PX + 420 * ES, PY + 160 * ES, 360 * ES, 12, False): lblParr.TextAlign = fmTextAlignCenter: lblParr.WordWrap = True: lblParr.Height = 40
  Set lblNom = TL(PX + 40 * ES, PY + 246 * ES, 740 * ES, 17, True)
  Set lblApe = TL(PX + 40 * ES, PY + 312 * ES, 740 * ES, 12, False)
  ' impresora y botones
  NL "Impresora:", PX, 262, 60
  Set cboImp = Me.Controls.Add("Forms.ComboBox.1")
  cboImp.Left = PX + 62: cboImp.Top = 258: cboImp.Width = 338: cboImp.Height = 18: cboImp.Style = fmStyleDropDownList
  Set t = NL("Se guarda la impresora elegida; la próxima vez aparece seleccionada.", PX, 280, 400): t.ForeColor = RGB(120, 120, 120)
  Set bImprimir = NB("Imprimir marcadas", PX, 304, 196, 36, RGB(0, 128, 96))
  Set bZpl = NB("Guardar ZPL (prueba)", PX + 204, 304, 196, 36, RGB(120, 120, 120))
  Set bCancel = NB("Cancelar", PX + 204, 346, 196, 26, RGB(192, 80, 77))

  ' pedidos
  Set wsD = ThisWorkbook.Worksheets(HDAT)
  If gEtiqFilas Is Nothing Then Set gEtiqFilas = New Collection
  ReDim mFilas(1 To IIf(gEtiqFilas.Count > 0, gEtiqFilas.Count, 1))
  For Each f In gEtiqFilas
    If Len(TXE(wsD.Cells(f, D_DEST).Value)) > 0 Then
      mN = mN + 1: mFilas(mN) = f
      lst.AddItem TXE(wsD.Cells(f, D_PED).Value)
      lst.List(lst.ListCount - 1, 1) = TXE(wsD.Cells(f, D_DEST).Value)
      lst.List(lst.ListCount - 1, 2) = TXE(wsD.Cells(f, D_NOM).Value)
      lst.Selected(lst.ListCount - 1) = True
    Else
      nSin = nSin + 1
      LogE "ETIQUETAS fila " & f & " pedido " & TXE(wsD.Cells(f, D_PED).Value) & ": sin DESTINO (fuera de cobertura TMS); no se imprime", "AVISO"
    End If
  Next
  lblInfo.Caption = mN & " etiqueta(s) para imprimir" & IIf(nSin > 0, "   (" & nSin & " pedido(s) sin destino se omiten: ver registro)", "")
  ' impresoras (una sola vez)
  guard = Cfg("IMPRESORA_ZEBRA")
  Set colImp = ListaImpresoras()
  For Each it In colImp
    cboImp.AddItem it
    If it = guard Then cboImp.ListIndex = cboImp.ListCount - 1
  Next
  If cboImp.ListIndex < 0 Then
    For i = 0 To cboImp.ListCount - 1
      If InStr(1, cboImp.List(i), "ZD", vbTextCompare) > 0 Or InStr(1, cboImp.List(i), "ZEBRA", vbTextCompare) > 0 Or InStr(1, cboImp.List(i), "ZDESIGNER", vbTextCompare) > 0 Then cboImp.ListIndex = i: Exit For
    Next
  End If
  If mN > 0 Then lst.ListIndex = 0: Muestra 1
  ActualizarBoton
  mLay = CapturarLayout(Me, mCW, mCH): mEsc = 1
  HacerRedimensionable TITULO
  AjustarAPantalla Me, BASE_W, BASE_H, 0.85
  Reescalar
End Sub

Private Sub UserForm_Resize()
  Reescalar
End Sub

Private Sub Reescalar()
  If Not IsArray(mLay) Or mEscalando Then Exit Sub
  Dim f As Double
  f = EscalaAjuste(Me, mCW, mCH)
  If Abs(f - mEsc) < 0.01 Then Exit Sub
  mEscalando = True
  mEsc = f
  EscalarLayout Me, mLay, f, mCW, mCH
  If lst.ListIndex >= 0 Then Muestra lst.ListIndex + 1
  mEscalando = False
End Sub

Private Function NL(cap As String, x As Single, y As Single, w As Single, Optional h As Single = 14, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = h: c.Font.Bold = bold
  Set NL = c
End Function
Private Function TL(x As Single, y As Single, w As Single, tam As Single, bold As Boolean) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Left = x: c.Top = y: c.Width = w: c.Height = tam * 1.6: c.BackStyle = fmBackStyleTransparent
  c.Font.Name = "Arial": c.Font.Size = tam: c.Font.Bold = bold: c.ForeColor = vbBlack
  Set TL = c
End Function
Private Function NB(cap As String, x As Single, y As Single, w As Single, h As Single, col As Long) As MSForms.CommandButton
  Dim c As MSForms.CommandButton: Set c = Me.Controls.Add("Forms.CommandButton.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = h
  c.BackColor = col: c.ForeColor = vbWhite: c.Font.Bold = True: c.TakeFocusOnClick = False
  Set NB = c
End Function

' Dibuja la etiqueta de muestra del elemento k (1..mN)
Private Sub Muestra(ByVal k As Long)
  Dim wsD As Worksheet, f As Long, ped As String, nom As String, ape As String, parr As String, m As String
  Dim i As Long, x As Single, w As Single, nBar As Long, e As Double, ox As Single, oy As Single
  If k < 1 Or k > mN Then Exit Sub
  Set wsD = ThisWorkbook.Worksheets(HDAT)
  f = mFilas(k)
  ped = TXE(wsD.Cells(f, D_PED).Value)
  PartirNombre TXE(wsD.Cells(f, D_NOM).Value), nom, ape
  parr = TXE(wsD.Cells(f, D_PARRSP).Value): If Len(parr) = 0 Then parr = TXE(wsD.Cells(f, D_PARR).Value)
  e = mEsc: If e <= 0 Then e = 1
  ox = fondo.Left: oy = fondo.Top
  ' barras (aproximación en pantalla; la impresora genera el código real)
  m = Code128Modulos(ped)
  x = ox + 40 * ES * e
  For i = 1 To 80: barras(i).Visible = False: Next
  For i = 1 To Len(m)
    w = Val(Mid$(m, i, 1)) * 3 * ES * e * 0.5
    If i Mod 2 = 1 Then
      nBar = nBar + 1
      If nBar <= 80 Then
        barras(nBar).Left = x: barras(nBar).Top = oy + 40 * ES * e: barras(nBar).Width = w: barras(nBar).Height = 100 * ES * e
        barras(nBar).Visible = True
      End If
    End If
    x = x + w
  Next
  lblNum.Caption = EspaciosDig(ped)
  lblDest.Caption = UCase$(TXE(wsD.Cells(f, D_DEST).Value))
  lblParr.Caption = UCase$(parr)
  lblNom.Caption = nom
  lblApe.Caption = ape
End Sub

Private Function EspaciosDig(ByVal s As String) As String
  Dim i As Long, o As String
  For i = 1 To Len(s): o = o & Mid$(s, i, 1) & IIf(i < Len(s), " ", ""): Next
  EspaciosDig = o
End Function

Private Function Marcadas() As Collection
  Dim c As New Collection, i As Long
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then c.Add mFilas(i + 1)
  Next
  Set Marcadas = c
End Function

Private Sub ActualizarBoton()
  bImprimir.Caption = "Imprimir " & Marcadas().Count & " etiqueta(s)"
End Sub

Private Sub lst_Change()
  If lst.ListIndex >= 0 Then Muestra lst.ListIndex + 1
  ActualizarBoton
End Sub

Private Sub bTodos_Click()
  Dim i As Long
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = True: Next
  ActualizarBoton
End Sub
Private Sub bNinguno_Click()
  Dim i As Long
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = False: Next
  ActualizarBoton
End Sub

Private Sub bImprimir_Click()
  Dim c As Collection, n As Long
  Set c = Marcadas()
  If c.Count = 0 Then MsgBox "No hay etiquetas marcadas.", vbInformation: Exit Sub
  If cboImp.ListIndex < 0 Then MsgBox "Elige la impresora Zebra.", vbExclamation: Exit Sub
  If MsgBox("Imprimir " & c.Count & " etiqueta(s) en:" & vbCrLf & cboImp.Text, vbYesNo + vbQuestion, "Etiquetas") <> vbYes Then Exit Sub
  SetCfg "IMPRESORA_ZEBRA", cboImp.Text
  n = ImprimirFilas(c, cboImp.Text, False)
  If n > 0 Then
    MsgBox n & " etiqueta(s) enviadas a " & cboImp.Text & ".", vbInformation
    Unload Me
  Else
    MsgBox "No se pudo imprimir. Revisa que la Zebra esté encendida y conectada (detalle en el registro).", vbExclamation
  End If
End Sub

Private Sub bZpl_Click()
  Dim c As Collection, n As Long
  Set c = Marcadas()
  If c.Count = 0 Then MsgBox "No hay etiquetas marcadas.", vbInformation: Exit Sub
  n = ImprimirFilas(c, "", True)
  If n > 0 Then MsgBox n & " etiqueta(s) guardadas en la carpeta de exportes (archivo .zpl).", vbInformation
End Sub

Private Sub bCancel_Click()
  Unload Me
End Sub
