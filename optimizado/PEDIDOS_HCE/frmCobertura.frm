Option Explicit
' =====================================================================================
'  frmCobertura (NUEVO) - Buscador de cobertura. Insertar > UserForm, (Name) = frmCobertura.
'  Se abre desde el panel ("Buscar cobertura") o desde el validador ("Buscar cobertura...").
'  Filtra COBERTURA por provincia / cantón / parroquia y marca cada opción:
'    SUGERIDA (más cercana a la dirección del pedido), EN LA DIRECCION, PRINCIPAL DEL CANTON,
'    CIUDAD PRINCIPAL / SECUNDARIA, ZONA PELIGROSA. "Asignar al pedido" la pasa al validador.
' =====================================================================================
Private Const BASE_W As Single = 1000
Private Const BASE_H As Single = 580
Private Const TITULO As String = "Buscar cobertura HYCITE"

Private mData() As String, mNorm() As String, mRank() As Long, mN As Long
Private mIdx() As Long, mNIdx As Long, mCarga As Boolean, mInW0 As Single, mInH0 As Single
Private mCtxP As String, mCtxNc As String, mCtxQ As String, mCtxA As String, mSugP As String, mSugQ As String, mSugNc As String, mSugMot As String
Private lblCtx As MSForms.Label, lblInfo As MSForms.Label, lblCnt As MSForms.Label
Private WithEvents cboProv As MSForms.ComboBox
Private WithEvents cboCant As MSForms.ComboBox
Private WithEvents txtParr As MSForms.TextBox
Private WithEvents txtLibre As MSForms.TextBox
Private WithEvents chkMarcadas As MSForms.CheckBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents bAsignar As MSForms.CommandButton
Private WithEvents bHoja As MSForms.CommandButton
Private WithEvents bLimpiar As MSForms.CommandButton
Private WithEvents bCerrar As MSForms.CommandButton
' columnas de mData: 1 MARCA | 2 PARROQUIA | 3 CANTON | 4 PROVINCIA | 5 SIGLA | 6 GESTOR Q | 7 SUGERIDO R | 8 TRAYECTO | 9 ZONAS | 10 FILA | 11 PUNTOS TMC

Private Sub UserForm_Initialize()
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H
  CargarBases
  Dim t As MSForms.Label, i As Long, anchos, titulos, x As Single
  Set lblCtx = NL("", 10, 6, 970, True): lblCtx.Height = 30: lblCtx.WordWrap = True: lblCtx.ForeColor = RGB(48, 84, 150)
  NL "Provincia:", 10, 42, 50
  Set cboProv = NC(62, 38, 170)
  NL "Cantón:", 242, 42, 40
  Set cboCant = NC(284, 38, 170)
  NL "Parroquia contiene:", 464, 42, 92
  Set txtParr = Me.Controls.Add("Forms.TextBox.1"): txtParr.Left = 556: txtParr.Top = 38: txtParr.Width = 130: txtParr.Height = 18
  NL "Buscar:", 696, 42, 38
  Set txtLibre = Me.Controls.Add("Forms.TextBox.1"): txtLibre.Left = 736: txtLibre.Top = 38: txtLibre.Width = 120: txtLibre.Height = 18
  txtLibre.ControlTipText = "Busca en todas las columnas (sigla, gestor, trayecto...)."
  Set chkMarcadas = Me.Controls.Add("Forms.CheckBox.1")
  chkMarcadas.Caption = "Solo marcadas": chkMarcadas.Left = 864: chkMarcadas.Top = 38: chkMarcadas.Width = 110: chkMarcadas.Height = 18
  chkMarcadas.ControlTipText = "Muestra solo sugerida, en la dirección, principal del cantón y ciudades principales/secundarias."
  Set t = NL("MARCAS:  SUGERIDA = más cercana a la dirección del pedido   DIRECCION = mencionada en la dirección   PRINCIPAL = cabecera del cantón   CP / CS = ciudad principal / secundaria   ZONA(n) = zonas peligrosas", 10, 62, 970)
  t.ForeColor = RGB(90, 90, 90)
  titulos = Array("MARCA", "PARROQUIA", "CANTON", "PROVINCIA", "SIGLA", "GESTOR (Q)", "SUGERIDO (R)", "TRAYECTO (S)", "ZONAS")
  anchos = Array(170, 150, 110, 95, 140, 125, 65, 90, 40)
  x = 13
  For i = 0 To 8
    Set t = NL(" " & titulos(i), x, 80, CSng(anchos(i)) - 1, True)
    t.BackColor = RGB(48, 84, 150): t.ForeColor = vbWhite: t.Font.Size = 8
    x = x + anchos(i)
  Next
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = 10: lst.Top = 94: lst.Width = 975: lst.Height = 300: lst.Font.Size = 8
  lst.ColumnCount = 9: lst.ColumnWidths = "170;150;110;95;140;125;65;90;40"
  lst.ControlTipText = "Clic = ver detalle. Doble clic = asignar al pedido (si el validador está abierto)."
  Set lblCnt = NL("", 10, 398, 400, True)
  Set lblInfo = NL("Selecciona una parroquia para ver el detalle.", 10, 414, 975): lblInfo.Height = 88
  lblInfo.WordWrap = True: lblInfo.BorderStyle = fmBorderStyleSingle: lblInfo.BackColor = RGB(248, 248, 248): lblInfo.Font.Size = 9
  Set bAsignar = BT("Asignar al pedido", 10, 508, 160, RGB(112, 173, 71), "Pasa esta provincia/cantón/parroquia al validador (luego pulsa 'Aplicar y siguiente' allí).")
  Set bHoja = BT("Ver en hoja COBERTURA", 176, 508, 150, RGB(120, 120, 120), "Selecciona la fila en la hoja COBERTURA.")
  Set bLimpiar = BT("Limpiar filtros", 332, 508, 110, RGB(150, 150, 150), "Quita los filtros de texto.")
  Set bCerrar = BT("Cerrar", 885, 508, 100, RGB(192, 80, 77), "Cierra el buscador.")

  ' contexto del pedido (si viene del validador o del panel)
  mCtxP = Normaliza(gCtxProv): mCtxNc = Normaliza(gCtxCant): mCtxQ = Normaliza(gCtxParr): mCtxA = " " & Normaliza(gCtxDir) & " "
  If Len(gCtxPedido) > 0 Then
    lblCtx.Caption = "Pedido " & gCtxPedido & " (fila " & gCtxFila & ")  -  Dirección: " & gCtxDir
    Dim sc
    sc = SugerirCercana(mCtxP, mCtxNc, mCtxQ, mCtxA, mSugMot)
    If IsArray(sc) Then mSugNc = Normaliza(CStr(sc(0))): mSugQ = Normaliza(CStr(sc(1))): mSugP = mCtxP
  Else
    lblCtx.Caption = "Sin pedido seleccionado: búsqueda libre en COBERTURA. (Para asignar, ábrelo desde el validador.)"
  End If
  bAsignar.Enabled = FormAbierto("frmValidar") And Len(gCtxPedido) > 0
  CargarDatos
  mInW0 = Me.InsideWidth: mInH0 = Me.InsideHeight
  HacerRedimensionable TITULO
  AjustarAPantalla Me, BASE_W, BASE_H, 0.85
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

Private Function NL(cap As String, L As Single, tp As Single, w As Single, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = cap: c.Left = L: c.Top = tp: c.Width = w: c.Height = 14: c.Font.Bold = bold
  Set NL = c
End Function
Private Function NC(L As Single, tp As Single, w As Single) As MSForms.ComboBox
  Dim c As MSForms.ComboBox: Set c = Me.Controls.Add("Forms.ComboBox.1")
  c.Left = L: c.Top = tp: c.Width = w: c.Height = 18: c.ListRows = 18: c.Style = fmStyleDropDownList
  Set NC = c
End Function
Private Function BT(cap As String, L As Single, tp As Single, w As Single, col As Long, tip As String) As MSForms.CommandButton
  Dim b As MSForms.CommandButton: Set b = Me.Controls.Add("Forms.CommandButton.1")
  b.Caption = cap: b.Left = L: b.Top = tp: b.Width = w: b.Height = 30
  b.BackColor = col: b.ForeColor = vbWhite: b.Font.Bold = True: b.ControlTipText = tip
  Set BT = b
End Function
Private Function FormAbierto(ByVal nombre As String) As Boolean
  Dim i As Long
  For i = 0 To VBA.UserForms.Count - 1
    If VBA.UserForms(i).Name = nombre Then FormAbierto = True: Exit Function
  Next
End Function
Private Function VS(v As Variant) As String
  If IsError(v) Then VS = "" Else VS = Trim$(CStr(v & ""))
End Function
Private Function FindCol(ws As Worksheet, ByVal lc As Long, ByVal txt As String) As Long
  Dim j As Long
  For j = 1 To lc
    If InStr(1, TX(ws.Cells(1, j)), txt, vbTextCompare) > 0 Then FindCol = j: Exit Function
  Next
End Function

' ---------- datos ----------
Private Sub CargarDatos()
  Dim wc As Worksheet, lr As Long, lc As Long, v, i As Long, n As Long, f As Long
  Dim cQ As Long, cR As Long, cS As Long, P As String, nc As String, Q As String, tr As String, marca As String, rk As Long
  Dim nz As Long, pts As String, provs As Object, k
  Set wc = ThisWorkbook.Worksheets("COBERTURA")
  lr = wc.Cells(wc.Rows.Count, 1).End(xlUp).Row
  lc = wc.Cells(1, wc.Columns.Count).End(xlToLeft).Column
  If lr < 2 Then Exit Sub
  cQ = FindCol(wc, lc, "GESTOR DE ENTREGAS"): cR = FindCol(wc, lc, "SUGERIDO"): cS = FindCol(wc, lc, "TRAMACO")
  v = wc.Range(wc.Cells(1, 1), wc.Cells(lr, lc)).Value
  ReDim mData(1 To lr - 1, 1 To 11): ReDim mNorm(1 To lr - 1, 1 To 11): ReDim mRank(1 To lr - 1)
  Set provs = CreateObject("Scripting.Dictionary")
  For i = 2 To lr
    If Len(VS(v(i, 1))) > 0 Then
      n = n + 1
      P = Normaliza(VS(v(i, 1))): nc = Normaliza(VS(v(i, 2))): Q = Normaliza(VS(v(i, 3)))
      If cS > 0 Then tr = UCase$(VS(v(i, cS))) Else tr = ""
      marca = "": rk = 9
      If Len(mSugQ) > 0 And P = mSugP And nc = mSugNc And Q = mSugQ Then marca = "SUGERIDA": rk = 1
      If Len(mCtxP) > 0 And P = mCtxP And Len(Q) >= 4 And InStr(mCtxA, " " & Q & " ") > 0 Then
        marca = marca & IIf(Len(marca) > 0, " + ", "") & "DIRECCION"
        If rk > 2 Then rk = 2
      End If
      If Q = nc Then
        marca = marca & IIf(Len(marca) > 0, " + ", "") & "PRINCIPAL"
        If rk > 3 Then rk = 3
      End If
      If tr = "CP" Then
        marca = marca & IIf(Len(marca) > 0, " + ", "") & "CP"
        If rk > 4 Then rk = 4
      End If
      If tr = "CS" Then
        marca = marca & IIf(Len(marca) > 0, " + ", "") & "CS"
        If rk > 5 Then rk = 5
      End If
      nz = ContarZonas(P, nc, Q, pts)
      If nz > 0 Then marca = marca & IIf(Len(marca) > 0, " + ", "") & "ZONA(" & nz & ")"
      mData(n, 1) = marca
      mData(n, 2) = VS(v(i, 3)): mData(n, 3) = VS(v(i, 2)): mData(n, 4) = VS(v(i, 1)): mData(n, 5) = VS(v(i, 4))
      If cQ > 0 Then mData(n, 6) = VS(v(i, cQ))
      If cR > 0 Then mData(n, 7) = VS(v(i, cR))
      mData(n, 8) = TrayectoTexto(tr)
      mData(n, 9) = IIf(nz > 0, CStr(nz), "")
      mData(n, 10) = CStr(i): mData(n, 11) = pts
      mRank(n) = rk
      For f = 1 To 9: mNorm(n, f) = Normaliza(mData(n, f)): Next
      If Not provs.Exists(mData(n, 4)) Then provs(mData(n, 4)) = 1
    End If
  Next
  mN = n
  mCarga = True
  cboProv.Clear: cboProv.AddItem "(todas)"
  For Each k In provs.Keys: cboProv.AddItem k: Next
  cboProv.ListIndex = 0
  For i = 1 To cboProv.ListCount - 1
    If Normaliza(cboProv.List(i)) = mCtxP And Len(mCtxP) > 0 Then cboProv.ListIndex = i: Exit For
  Next
  LlenarCantones
  If Len(mCtxNc) > 0 Then
    For i = 1 To cboCant.ListCount - 1
      If Normaliza(cboCant.List(i)) = mCtxNc Then cboCant.ListIndex = i: Exit For
    Next
  End If
  mCarga = False
  Refrescar
End Sub

Private Sub LlenarCantones()
  Dim u As Object, r As Long, pv As String
  cboCant.Clear: cboCant.AddItem "(todos)"
  Set u = CreateObject("Scripting.Dictionary")
  pv = IIf(cboProv.ListIndex > 0, Normaliza(cboProv.Text), "")
  For r = 1 To mN
    If Len(pv) = 0 Or mNorm(r, 4) = pv Then
      If Not u.Exists(mData(r, 3)) Then u(mData(r, 3)) = 1: cboCant.AddItem mData(r, 3)
    End If
  Next
  cboCant.ListIndex = 0
End Sub

Private Sub Refrescar()
  If mCarga Then Exit Sub
  Dim r As Long, k As Long, j As Long, i As Long, t As Long, arr(), pv As String, cv As String, tq As String, tl As String, busc As String
  pv = IIf(cboProv.ListIndex > 0, Normaliza(cboProv.Text), "")
  cv = IIf(cboCant.ListIndex > 0, Normaliza(cboCant.Text), "")
  tq = Normaliza(txtParr.Text): tl = Normaliza(txtLibre.Text)
  mNIdx = 0: lst.Clear
  If mN = 0 Then Exit Sub
  ReDim mIdx(1 To mN)
  For r = 1 To mN
    If (Len(pv) = 0 Or mNorm(r, 4) = pv) And (Len(cv) = 0 Or mNorm(r, 3) = cv) Then
      If Len(tq) = 0 Or InStr(mNorm(r, 2), tq) > 0 Then
        busc = " " & mNorm(r, 1) & " " & mNorm(r, 2) & " " & mNorm(r, 3) & " " & mNorm(r, 5) & " " & mNorm(r, 6) & " " & mNorm(r, 7) & " " & mNorm(r, 8) & " "
        If Len(tl) = 0 Or InStr(busc, tl) > 0 Then
          If Not chkMarcadas.Value Or mRank(r) < 9 Then mNIdx = mNIdx + 1: mIdx(mNIdx) = r
        End If
      End If
    End If
  Next
  ' orden: marcadas primero (sugerida, dirección, principal, CP, CS), luego cantón y parroquia
  For i = 2 To mNIdx
    t = mIdx(i): j = i - 1
    Do While j >= 1
      If Not Mayor(mIdx(j), t) Then Exit Do
      mIdx(j + 1) = mIdx(j): j = j - 1
    Loop
    mIdx(j + 1) = t
  Next
  If mNIdx > 0 Then
    ReDim arr(1 To mNIdx, 1 To 9)
    For k = 1 To mNIdx
      For j = 1 To 9
        arr(k, j) = mData(mIdx(k), j)
      Next
    Next
    lst.List = arr
  End If
  lblCnt.Caption = "Coberturas encontradas: " & mNIdx & IIf(Len(mSugQ) > 0, "     Sugerida para el pedido: " & mSugQ & " (" & mSugMot & ")", "")
  lblInfo.Caption = "Selecciona una parroquia para ver el detalle."
End Sub

Private Function Mayor(ByVal a As Long, ByVal b As Long) As Boolean
  If mRank(a) <> mRank(b) Then Mayor = (mRank(a) > mRank(b)): Exit Function
  If mNorm(a, 3) <> mNorm(b, 3) Then Mayor = (mNorm(a, 3) > mNorm(b, 3)): Exit Function
  Mayor = (mNorm(a, 2) > mNorm(b, 2))
End Function

Private Sub MostrarDetalle()
  If lst.ListIndex < 0 Or mNIdx = 0 Then Exit Sub
  Dim r As Long, s As String, ex, porque As String, mr As String
  r = mIdx(lst.ListIndex + 1)
  mr = mData(r, 1)
  If InStr(mr, "SUGERIDA") > 0 Then porque = porque & "Es la cobertura que el sistema considera MÁS CERCANA a la dirección del pedido actual (" & mSugMot & "). "
  If InStr(mr, "DIRECCION") > 0 Then porque = porque & "Esta parroquia se MENCIONA en la dirección del pedido actual. "
  If InStr(mr, "PRINCIPAL") > 0 Then porque = porque & "Es la PARROQUIA PRINCIPAL (cabecera) del cantón " & mData(r, 3) & ". "
  If InStr(mr, "CP") > 0 Then porque = porque & "Es CIUDAD PRINCIPAL para TRAMACO. "
  If InStr(mr, "CS") > 0 Then porque = porque & "Es CIUDAD SECUNDARIA para TRAMACO. "
  ex = DatosExt(mNorm(r, 4), mNorm(r, 3), mNorm(r, 2), mCtxA)
  s = mData(r, 2) & "  -  cantón " & mData(r, 3) & ", provincia " & mData(r, 4) & "      Sigla: " & mData(r, 5) & vbCrLf & _
      "Si se asigna: gestor " & ex(0) & " (destino " & ex(1) & ")" & IIf(Len(ex(2)) > 0, "   Trayecto: " & ex(2), "") & _
      "      Gestor cobertura (Q): " & IIf(Len(mData(r, 6)) > 0, mData(r, 6), "-") & "   Sugerido (R): " & IIf(Len(mData(r, 7)) > 0, mData(r, 7), "-") & vbCrLf & _
      IIf(Len(mData(r, 9)) > 0, "Zonas peligrosas registradas en esta parroquia: " & mData(r, 9) & IIf(Len(mData(r, 11)) > 0, " (punto TMC: " & mData(r, 11) & ")", "") & vbCrLf, "") & _
      IIf(Len(porque) > 0, "Por qué elegirla: " & porque, "Sin marca especial: es una cobertura válida del cantón.")
  lblInfo.Caption = s
  If Len(mData(r, 9)) > 0 Then lblInfo.ForeColor = RGB(191, 90, 0) Else lblInfo.ForeColor = RGB(30, 30, 30)
End Sub

' ---------- eventos ----------
Private Sub cboProv_Change()
  If mCarga Then Exit Sub
  mCarga = True: LlenarCantones: mCarga = False
  Refrescar
End Sub
Private Sub cboCant_Change()
  Refrescar
End Sub
Private Sub txtParr_Change()
  Refrescar
End Sub
Private Sub txtLibre_Change()
  Refrescar
End Sub
Private Sub chkMarcadas_Click()
  Refrescar
End Sub
Private Sub lst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  RuedaActivar lst, TITULO
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  RuedaDesactivar
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  RuedaDesactivar
End Sub
' flechas arriba/abajo del teclado también actualizan el detalle
Private Sub lst_Change()
  If lst.ListIndex >= 0 Then MostrarDetalle
End Sub

Private Sub lst_Click()
  MostrarDetalle
End Sub
Private Sub lst_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If KeyCode = vbKeyReturn And bAsignar.Enabled Then bAsignar_Click
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  If bAsignar.Enabled Then bAsignar_Click
End Sub

Private Sub bAsignar_Click()
  If lst.ListIndex < 0 Or mNIdx = 0 Then MsgBox "Selecciona una parroquia de la lista.", vbInformation: Exit Sub
  If Not FormAbierto("frmValidar") Then MsgBox "El validador está cerrado. Ábrelo y usa 'Buscar cobertura...'.", vbInformation: Exit Sub
  Dim r As Long, i As Long, mot As String
  r = mIdx(lst.ListIndex + 1)
  mot = "Buscar cobertura" & IIf(Len(mData(r, 1)) > 0, ": " & mData(r, 1), "")
  If InStr(mData(r, 1), "SUGERIDA") > 0 Then mot = mot & " (" & mSugMot & ")"
  For i = 0 To VBA.UserForms.Count - 1
    If VBA.UserForms(i).Name = "frmValidar" Then VBA.UserForms(i).AsignarCobertura mData(r, 4), mData(r, 3), mData(r, 2), mot: Exit For
  Next
  LogP "Buscar cobertura: pedido " & gCtxPedido & " -> " & mData(r, 4) & "/" & mData(r, 3) & "/" & mData(r, 2) & " asignado al validador"
  Unload Me
End Sub

Private Sub bHoja_Click()
  If lst.ListIndex < 0 Or mNIdx = 0 Then Exit Sub
  Dim r As Long: r = CLng(mData(mIdx(lst.ListIndex + 1), 10))
  On Error Resume Next
  ThisWorkbook.Worksheets("COBERTURA").Activate
  ThisWorkbook.Worksheets("COBERTURA").Cells(r, 3).Select
  On Error GoTo 0
End Sub

Private Sub bLimpiar_Click()
  mCarga = True
  txtParr.Text = "": txtLibre.Text = "": chkMarcadas.Value = False
  mCarga = False
  Refrescar
End Sub

Private Sub bCerrar_Click()
  Unload Me
End Sub
