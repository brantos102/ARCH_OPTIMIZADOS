Option Explicit
' =====================================================================================
'  frmEtiquetas - Selección, vista previa e impresión de etiquetas Zebra (10 x 5 cm).
'  Insertar > UserForm, nombre frmEtiquetas, pegar este código. Lo abre frmEGR o el menú.
'   - filtros por estado de etiqueta, destino y texto libre
'   - lo marcado se conserva aunque cambie el filtro
'   - vista previa del pedido seleccionado
'   - la impresora se elige una vez y queda guardada
'  Rendimiento: la lista se arma en memoria y se vuelca de una sola vez (lst.List = arr).
' =====================================================================================
Private Const BASE_W As Single = 860
Private Const BASE_H As Single = 520
Private Const TITULO As String = "Etiquetas HYCITE - seleccionar e imprimir"
Private Const PX As Single = 356        ' esquina de la etiqueta de muestra
Private Const PY As Single = 60
Private Const ES As Single = 0.6        ' 800 x 400 puntos Zebra -> 480 x 240 en pantalla

Private mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean
Private mCerrando As Boolean, mCarga As Boolean
Private mFila() As Long, mPed() As String, mDest() As String, mNom() As String, mParr() As String, mEst() As String
Private mN As Long, mIdx() As Long, mNIdx As Long, mMarca As Object
Private lblInfo As MSForms.Label, lblCnt As MSForms.Label, fondo As MSForms.Label
Private lblNum As MSForms.Label, lblDest As MSForms.Label, lblParr As MSForms.Label, lblNom As MSForms.Label
Private barras(1 To 90) As MSForms.Label
Private WithEvents lst As MSForms.ListBox
Private WithEvents cboEstado As MSForms.ComboBox
Private WithEvents cboDest As MSForms.ComboBox
Private WithEvents txtBuscar As MSForms.TextBox
Private cboImp As MSForms.ComboBox
Private WithEvents bTodos As MSForms.CommandButton
Private WithEvents bNinguno As MSForms.CommandButton
Private WithEvents bLimpF As MSForms.CommandButton
Private WithEvents bImprimir As MSForms.CommandButton
Private WithEvents bZpl As MSForms.CommandButton
Private WithEvents bCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
  Dim t As MSForms.Label, i As Long, f
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H
  Set mMarca = CreateObject("Scripting.Dictionary")

  Set lblInfo = NL("", 10, 8, 830, 14, True): lblInfo.ForeColor = RGB(48, 84, 150)

  ' ----- filtros -----
  NL "Etiqueta:", 10, 30, 46
  Set cboEstado = NC(58, 27, 104)
  For Each f In Array("PENDIENTES", "REIMPRIMIR", "PENDIENTES + REIMPRIMIR", "IMPRESAS", "TODAS"): cboEstado.AddItem f: Next
  NL "Destino:", 170, 30, 42
  Set cboDest = NC(214, 27, 64)
  For Each f In Array("TODOS", "PRO", "GYE", "UIO", "GPS"): cboDest.AddItem f: Next
  NL "Buscar:", 286, 30, 38
  Set txtBuscar = Me.Controls.Add("Forms.TextBox.1")
  txtBuscar.Left = 326: txtBuscar.Top = 27: txtBuscar.Width = 150: txtBuscar.Height = 18
  txtBuscar.ControlTipText = "Pedido, destinatario o parroquia."
  Set bLimpF = NB("Quitar filtros", 482, 26, 86, 20, RGB(120, 120, 120))

  ' ----- lista -----
  Set lblCnt = NL("", 10, 50, 330)
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = 10: lst.Top = 66: lst.Width = 336: lst.Height = 352: lst.Font.Size = 8
  lst.ColumnCount = 4: lst.ColumnWidths = "72;34;150;74"
  lst.MultiSelect = fmMultiSelectMulti: lst.ListStyle = fmListStyleOption
  lst.ControlTipText = "Marca las etiquetas a imprimir. Lo marcado se conserva al cambiar el filtro."
  Set bTodos = NB("Marcar lo filtrado", 10, 422, 112, 22, RGB(120, 120, 120))
  Set bNinguno = NB("Desmarcar todo", 126, 422, 104, 22, RGB(120, 120, 120))

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
  Set bImprimir = NB("Imprimir", PX, 342, 200, 38, RGB(0, 128, 96))
  bImprimir.Font.Size = 10
  Set bZpl = NB("Guardar ZPL (prueba)", PX + 208, 342, 172, 38, RGB(120, 120, 120))
  Set bCancel = NB("Cerrar", PX + 208, 386, 172, 26, RGB(192, 80, 77))
  Set t = NL("Si la etiqueta sale clara o corrida, ajusta ETIQ_OSCURIDAD y ETIQ_OFFSET en la hoja CONFIG_EGR.", PX, 418, 390)
  t.Height = 28: t.WordWrap = True: t.ForeColor = RGB(120, 120, 120)

  CargarPedidos
  CargarImpresoras
  mCarga = True
  cboDest.ListIndex = 0
  If PreSeleccion() > 0 Then
    cboEstado.ListIndex = 4      ' TODAS: ya vienen marcados los elegidos en el panel
    mCarga = False
    Filtrar False
  Else
    cboEstado.ListIndex = 2      ' pendientes + reimprimir
    mCarga = False
    Filtrar True
  End If
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
  lblInfo.Caption = mN & " pedido(s) con destino" & IIf(nSin > 0, "   ·   " & nSin & " sin destino (fuera de cobertura TMS) no se pueden imprimir", "")
End Sub

' Pedidos que venían marcados desde el panel (gEtiqFilas). Devuelve cuántos se marcaron.
Private Function PreSeleccion() As Long
  Dim f, k As Long, n As Long
  If gEtiqFilas Is Nothing Then Exit Function
  If gEtiqFilas.Count = 0 Then Exit Function
  For Each f In gEtiqFilas
    For k = 1 To mN
      If mFila(k) = CLng(f) Then mMarca(mPed(k)) = True: n = n + 1: Exit For
    Next
  Next
  PreSeleccion = n
End Function

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
  If des <> "TODOS" And des <> "" Then
    If mDest(k) <> des Then Exit Function
  End If
  If Len(q) > 0 Then
    If InStr(1, mPed(k) & " " & mNom(k) & " " & mParr(k) & " " & mDest(k), q, vbTextCompare) = 0 Then Exit Function
  End If
  Pasa = True
End Function

' marcarNuevos: True = marca lo que entra al filtro (al abrir y al cambiar de filtro)
Private Sub Filtrar(Optional ByVal marcarNuevos As Boolean = False)
  Dim k As Long, est As String, des As String, q As String, arr() As String, n As Long
  If mCerrando Or lst Is Nothing Or cboEstado Is Nothing Then Exit Sub
  est = cboEstado.Text: des = cboDest.Text: q = Trim$(txtBuscar.Text)
  ReDim mIdx(1 To IIf(mN > 0, mN, 1)): mNIdx = 0
  For k = 1 To mN
    If Pasa(k, est, des, q) Then
      mNIdx = mNIdx + 1: mIdx(mNIdx) = k
      If marcarNuevos Then mMarca(mPed(k)) = True
    End If
  Next
  mCarga = True
  lst.Clear
  If mNIdx > 0 Then
    ReDim arr(0 To mNIdx - 1, 0 To 3)
    For n = 1 To mNIdx
      k = mIdx(n)
      arr(n - 1, 0) = mPed(k): arr(n - 1, 1) = mDest(k): arr(n - 1, 2) = mNom(k): arr(n - 1, 3) = mEst(k)
    Next
    lst.List = arr
    For n = 1 To mNIdx
      If mMarca.Exists(mPed(mIdx(n))) Then
        If mMarca(mPed(mIdx(n))) Then lst.Selected(n - 1) = True
      End If
    Next
    lst.ListIndex = 0
  End If
  mCarga = False
  If mNIdx > 0 Then Muestra 0
  Contar
End Sub

Private Sub Contar()
  Dim n As Long
  ' al descargar el formulario los controles se destruyen antes que el código: sin esta guarda
  ' un último evento de la lista intentaba escribir en un botón que ya no existe
  If mCerrando Or lblCnt Is Nothing Or bImprimir Is Nothing Then Exit Sub
  n = Marcadas().Count
  lblCnt.Caption = mNIdx & " en la lista   ·   " & n & " marcada(s) para imprimir"
  bImprimir.Caption = IIf(n = 0, "Imprimir", "Imprimir " & n & " etiqueta(s)")
End Sub

Private Function Marcadas() As Collection
  Dim c As New Collection, k As Long
  For k = 1 To mN
    If mMarca.Exists(mPed(k)) Then If mMarca(mPed(k)) Then c.Add mFila(k)
  Next
  Set Marcadas = c
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
  Filtrar True
End Sub
Private Sub cboDest_Change()
  If mCarga Then Exit Sub
  Filtrar True
End Sub
Private Sub txtBuscar_Change()
  If mCarga Then Exit Sub
  Filtrar False
End Sub
Private Sub bLimpF_Click()
  mCarga = True
  cboEstado.ListIndex = 2: cboDest.ListIndex = 0: txtBuscar.Text = ""
  mCarga = False
  Filtrar True
End Sub

Private Sub lst_Change()
  Dim i As Long
  If mCerrando Or mCarga Or lst Is Nothing Then Exit Sub
  For i = 0 To lst.ListCount - 1
    If i < mNIdx Then mMarca(mPed(mIdx(i + 1))) = lst.Selected(i)
  Next
  If lst.ListIndex >= 0 Then Muestra lst.ListIndex
  Contar
End Sub

Private Sub bTodos_Click()
  Dim i As Long
  If mCerrando Or lst Is Nothing Then Exit Sub
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = True: Next
End Sub
Private Sub bNinguno_Click()
  Dim i As Long, k As Long
  If mCerrando Or lst Is Nothing Then Exit Sub
  For k = 1 To mN: mMarca(mPed(k)) = False: Next
  For i = 0 To lst.ListCount - 1: lst.Selected(i) = False: Next
  Contar
End Sub

Private Sub bImprimir_Click()
  Dim c As Collection, n As Long
  Set c = Marcadas()
  If c.Count = 0 Then MsgBox "No hay etiquetas marcadas.", vbInformation: Exit Sub
  If cboImp.ListIndex < 0 Then MsgBox "Elige la impresora Zebra.", vbExclamation: Exit Sub
  If MsgBox("Imprimir " & c.Count & " etiqueta(s) en:" & vbCrLf & cboImp.Text, vbYesNo + vbQuestion, "Etiquetas") <> vbYes Then Exit Sub
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
  Set c = Marcadas()
  If c.Count = 0 Then MsgBox "No hay etiquetas marcadas.", vbInformation: Exit Sub
  n = ImprimirFilas(c, "", True)
  If n > 0 Then MsgBox n & " etiqueta(s) guardadas en la carpeta de exportes (archivo .zpl)." & vbCrLf & _
                        "Se puede ver en labelary.com con 8 dpmm y 3,94 x 1,97 pulgadas.", vbInformation
End Sub

Private Sub bCancel_Click()
  mCerrando = True
  Unload Me
End Sub
