Option Explicit
' =====================================================================================
'  frmEGR - Panel del despacho (Formato EGR_FL_HYCITE). Insertar > UserForm, nombre frmEGR,
'  pegar este código. Los controles se crean al abrir. Requiere modEGR, modZebra y modVentanas.
' =====================================================================================
Private Const BASE_W As Single = 1000
Private Const BASE_H As Single = 640
Private Const TITULO As String = "Panel de despacho HYCITE - Etapa 2: EGR, etiquetas y reportes"
Private Const X0 As Single = 206

Private mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean
Private mVista As String, mFilas As Collection
Private lblVista As MSForms.Label, lblKpi As MSForms.Label, lblGuia As MSForms.Label
Private hdr(1 To 6) As MSForms.Label, txtLog As MSForms.TextBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents b1 As MSForms.CommandButton
Private WithEvents b2 As MSForms.CommandButton
Private WithEvents b3 As MSForms.CommandButton
Private WithEvents b4a As MSForms.CommandButton
Private WithEvents b4b As MSForms.CommandButton
Private WithEvents b4c As MSForms.CommandButton
Private WithEvents b4d As MSForms.CommandButton
Private WithEvents b4e As MSForms.CommandButton
Private WithEvents b5 As MSForms.CommandButton
Private WithEvents b6 As MSForms.CommandButton
Private WithEvents b6m As MSForms.CommandButton
Private cboFmt As MSForms.ComboBox
Private chkTMS As MSForms.CheckBox, chkTRA As MSForms.CheckBox, chkDES As MSForms.CheckBox
Private WithEvents bRep As MSForms.CommandButton
Private WithEvents bReg As MSForms.CommandButton
Private WithEvents bQuit As MSForms.CommandButton
Private WithEvents bDesb As MSForms.CommandButton
Private WithEvents bExcel As MSForms.CommandButton
Private WithEvents bCer As MSForms.CommandButton
Private WithEvents bAplSel As MSForms.CommandButton
Private WithEvents bAplTod As MSForms.CommandButton
Private WithEvents bIr As MSForms.CommandButton
Private WithEvents bLogL As MSForms.CommandButton

Private Sub UserForm_Initialize()
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H
  Dim t As MSForms.Label, i As Long
  Set t = L("PASOS DEL DESPACHO", 8, 6, 190, True): t.ForeColor = RGB(48, 84, 150)
  Set b1 = B("1  Actualizar todo", 8, 22, 190, 28, RGB(89, 89, 89), "Trae ITEMS API, ITEMS DEPOT (ODBC) y EMPAQUETADO, y las tablas dinámicas. Guarda respaldo antes. Si algo falla, lo dice aquí sin cerrar Excel.")
  Set b2 = B("2  Revisar cobertura TMS", 8, 54, 190, 28, RGB(47, 117, 181), "Segunda comprobación contra la cobertura TMS (no revalida lo de PEDIDOS HCE): fuera de cobertura, sin código postal, sin teléfono.")
  Set b3 = B("3  Proponer destinos (reglas)", 8, 86, 190, 28, RGB(237, 125, 49), "Aplica REGLAS_DESTINO (Guayas, Daule, cooperativas, Azuay...) y muestra los pedidos que cambian de destino para confirmarlos.")
  Set t = L("4  ETIQUETAS (Zebra 10 x 5 cm)", 8, 120, 190, True): t.ForeColor = RGB(0, 128, 96)
  Set b4a = B("Imprimir pendientes", 8, 136, 190, 24, RGB(0, 128, 96), "Imprime una etiqueta por pedido que aún no tiene STATUS = OK en la hoja ETIQUETAS.")
  Set b4b = B("Reimprimir pedido(s)...", 8, 163, 93, 22, RGB(0, 150, 110), "Escribe uno o varios números de pedido separados por coma.")
  Set b4c = B("Imprimir todas", 105, 163, 93, 22, RGB(0, 150, 110), "Imprime todas las etiquetas del día (pide confirmación).")
  Set b4d = B("Guardar ZPL (prueba)", 8, 188, 93, 22, RGB(120, 120, 120), "Guarda las etiquetas pendientes en un archivo .zpl sin imprimir (para revisar el diseño).")
  Set b4e = B("Elegir impresora", 105, 188, 93, 22, RGB(120, 120, 120), "Elige la Zebra ZD230 entre las impresoras instaladas.")
  Set b5 = B("5  Avance picking / empaque", 8, 218, 190, 28, RGB(112, 48, 160), "Por pedido: sin picking, picking parcial, pickeado sin empacar, empacado. Avisa SKU y cajas sin peso/precio.")
  Set t = L("6  EXPORTAR REPORTES", 8, 252, 190, True): t.ForeColor = RGB(48, 84, 150)
  Set cboFmt = Me.Controls.Add("Forms.ComboBox.1")
  cboFmt.Left = 8: cboFmt.Top = 268: cboFmt.Width = 70: cboFmt.Height = 18: cboFmt.Style = fmStyleDropDownList
  cboFmt.AddItem "CSV": cboFmt.AddItem "XLSX": cboFmt.AddItem "PDF": cboFmt.ListIndex = 0
  Set chkTMS = Chk("TMS", 84, 268, 42): Set chkTRA = Chk("TRAMACO", 126, 268, 72)
  Set chkDES = Chk("DESPACHOS (distribución)", 8, 288, 190)
  chkTMS.Value = True: chkTRA.Value = True: chkDES.Value = True
  Set b6 = B("Exportar", 8, 308, 93, 26, RGB(0, 97, 0), "Valida (errores, VALIDAR, campos vacíos) y exporta solo las filas con datos, sin cambiar las columnas.")
  Set b6m = B("Exportar + correo", 105, 308, 93, 26, RGB(0, 97, 0), "Exporta y abre un correo de Outlook con los archivos adjuntos.")
  Set lblGuia = L("", 8, 342, 190): lblGuia.Height = 70: lblGuia.WordWrap = True
  lblGuia.BackColor = RGB(255, 242, 204): lblGuia.BorderStyle = fmBorderStyleSingle: lblGuia.ForeColor = RGB(128, 64, 0)
  lblGuia.Caption = "Orden del día: PEDIDOS HCE (paso 7 Enviar a EGR) -> 1 -> 2 -> 3 -> 4 etiquetas -> empaquetado (Google Sheets) -> 1 otra vez -> 5 -> 6."
  Set t = L("HERRAMIENTAS", 8, 420, 190, True): t.ForeColor = RGB(48, 84, 150)
  Set bRep = B("Reparar fórmulas (una vez)", 8, 436, 190, 24, RGB(192, 80, 77), "Corrige los #REF! y prepara DATOS!A para las reglas. Guarda respaldo antes.")
  Set bReg = B("Editar reglas de destino", 8, 463, 190, 24, RGB(120, 120, 120), "Abre la hoja REGLAS_DESTINO (supervisor).")
  Set bQuit = B("Quitar destinos confirmados", 8, 490, 190, 24, RGB(120, 120, 120), "Vuelve todos los pedidos a la regla base por provincia.")
  Set bDesb = B("Desbloquear", 8, 517, 93, 24, RGB(120, 120, 120), "Restaura pantalla, eventos y cálculo si Excel quedó bloqueado.")
  Set bExcel = B("Ver Excel", 105, 517, 93, 24, RGB(0, 97, 0), "Oculta el panel. Para volver: Complementos > Panel EGR.")
  Set bCer = B("Cerrar panel", 8, 548, 190, 26, RGB(192, 80, 77), "Cierra el panel.")

  Set lblVista = L("Pulsa un paso para ver aquí su resultado.", X0, 6, 780, True): lblVista.ForeColor = RGB(48, 84, 150)
  Set lblKpi = L("", X0, 22, 780, True)
  For i = 1 To 6
    Set hdr(i) = L("", X0, 40, 10, True)
    hdr(i).BackColor = RGB(48, 84, 150): hdr(i).ForeColor = vbWhite: hdr(i).Font.Size = 8
  Next
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = X0: lst.Top = 54: lst.Width = 786: lst.Height = 300: lst.Font.Size = 8
  lst.ColumnCount = 6: lst.MultiSelect = fmMultiSelectExtended
  Set bAplSel = B("Aplicar seleccionados", X0, 358, 150, 24, RGB(237, 125, 49), "Confirma el nuevo destino de los pedidos seleccionados (Ctrl/Shift + clic para varios).")
  Set bAplTod = B("Aplicar todos", X0 + 156, 358, 110, 24, RGB(237, 125, 49), "Confirma el nuevo destino de todos los pedidos de la lista.")
  Set bIr = B("Ir a la fila en DATOS", X0 + 272, 358, 140, 24, RGB(120, 120, 120), "Oculta el panel y selecciona la fila en DATOS.")
  Set t = L("REGISTRO DE ACTIVIDAD (acciones y errores; historial en la hoja oculta LOG_EGR)", X0, 390, 600, True): t.ForeColor = RGB(48, 84, 150)
  Set bLogL = B("Limpiar", X0 + 726, 387, 60, 18, RGB(150, 150, 150), "Vacía el registro en pantalla.")
  Set txtLog = Me.Controls.Add("Forms.TextBox.1")
  txtLog.Left = X0: txtLog.Top = 408: txtLog.Width = 786: txtLog.Height = 196
  txtLog.MultiLine = True: txtLog.ScrollBars = fmScrollBarsVertical: txtLog.Locked = True: txtLog.WordWrap = True
  txtLog.Font.Name = "Consolas": txtLog.Font.Size = 8: txtLog.BackColor = RGB(30, 30, 30): txtLog.ForeColor = RGB(220, 220, 220)
  CargarLogPrevio
  VistaBotones
  mLay = CapturarLayout(Me, mCW, mCH): mEsc = 1
  HacerRedimensionable TITULO
  AjustarAPantalla Me, BASE_W, BASE_H, 0.92
  Reescalar
  LogE "PANEL EGR: abierto por " & Application.UserName
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
  If Len(mVista) > 0 Then PonerColumnas
  mEscalando = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  RuedaDesactivar
End Sub

' ---------- creación de controles ----------
Private Function L(cap As String, x As Single, y As Single, w As Single, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = 14: c.Font.Bold = bold
  Set L = c
End Function
Private Function B(cap As String, x As Single, y As Single, w As Single, h As Single, col As Long, tip As String) As MSForms.CommandButton
  Dim c As MSForms.CommandButton: Set c = Me.Controls.Add("Forms.CommandButton.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = h
  c.BackColor = col: c.ForeColor = vbWhite: c.Font.Bold = True: c.Font.Size = 8
  c.ControlTipText = tip: c.TakeFocusOnClick = False
  Set B = c
End Function
Private Function Chk(cap As String, x As Single, y As Single, w As Single) As MSForms.CheckBox
  Dim c As MSForms.CheckBox: Set c = Me.Controls.Add("Forms.CheckBox.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = 18
  Set Chk = c
End Function

' ---------- registro ----------
Public Sub AgregarLog(ByVal linea As String)
  If txtLog Is Nothing Then Exit Sub
  Dim t As String
  t = txtLog.Text
  If Len(t) > 60000 Then t = Right$(t, 40000)
  txtLog.Text = t & IIf(Len(t) > 0, vbCrLf, "") & linea
  txtLog.SelStart = Len(txtLog.Text)
  Me.Repaint
End Sub

Private Sub CargarLogPrevio()
  Dim it, t As String
  If gLogE Is Nothing Then Exit Sub
  For Each it In gLogE: t = t & IIf(Len(t) > 0, vbCrLf, "") & it: Next
  txtLog.Text = t: txtLog.SelStart = Len(t)
End Sub

' ---------- lista ----------
Private Sub DefinirColumnas(ByVal vista As String)
  mVista = vista
  PonerColumnas
End Sub

Private Sub PonerColumnas()
  Dim nom, anc, i As Long, x As Single, w As String, e As Double
  Select Case mVista
    Case "DESTINOS"
      nom = Array("FILA", "PEDIDO", "DESTINATARIO", "PROVINCIA / CANTÓN / PARROQUIA", "ACTUAL > NUEVO", "MOTIVO (regla)")
      anc = Array(34, 70, 150, 200, 80, 250)
    Case "COBERTURA"
      nom = Array("FILA", "PEDIDO", "DESTINATARIO", "PROV / CANTÓN / PARROQUIA (TMS)", "OBSERVACIÓN", "")
      anc = Array(34, 70, 150, 220, 310, 0)
    Case Else
      nom = Array("FILA", "PEDIDO", "DESTINATARIO", "DESTINO", "ESTADO", "")
      anc = Array(34, 70, 200, 60, 420, 0)
  End Select
  e = mEsc: If e <= 0 Then e = 1
  x = lst.Left + 3 * e
  For i = 1 To 6
    hdr(i).Caption = " " & nom(i - 1): hdr(i).Left = x: hdr(i).Width = IIf(anc(i - 1) > 0, anc(i - 1) * e - 1, 0.1)
    hdr(i).Visible = (anc(i - 1) > 0)
    x = x + anc(i - 1) * e
    w = w & IIf(i > 1, ";", "") & CStr(CLng(anc(i - 1) * e))
  Next
  lst.ColumnWidths = w
End Sub

Private Sub VistaBotones()
  bAplSel.Visible = (mVista = "DESTINOS"): bAplTod.Visible = (mVista = "DESTINOS")
End Sub

Private Sub LlenarLista(filas As Collection, ByVal nCols As Long)
  Dim r, j As Long
  lst.Clear
  Set mFilas = New Collection
  For Each r In filas
    lst.AddItem CStr(r(0))
    For j = 1 To nCols - 1
      If j <= UBound(r) Then lst.List(lst.ListCount - 1, j) = CStr(r(j))
    Next
    mFilas.Add CLng(r(0))
  Next
End Sub

' ---------- pasos ----------
Private Function Ocupado() As Boolean
  If gOcupadoE Then MsgBox "Hay un proceso en curso. Espera a que termine.", vbInformation: Ocupado = True
End Function

Private Sub b1_Click()
  If Ocupado() Then Exit Sub
  ActualizarTodo
End Sub

Private Sub b2_Click()
  If Ocupado() Then Exit Sub
  Dim c As New Collection, n As Long
  n = RevisarCoberturaTMS(c)
  DefinirColumnas "COBERTURA": VistaBotones
  LlenarLista c, 5
  lblVista.Caption = "COBERTURA TMS: pedidos con observación (fuera de cobertura, sin código postal, sin teléfono)"
  lblKpi.Caption = IIf(n = 0, "Todo en cobertura TMS. Sigue con el paso 3.", n & " pedido(s) con observación. Los de 'Fuera de cobertura' no tendrán destino ni etiqueta hasta corregir la cobertura.")
End Sub

Private Sub b3_Click()
  If Ocupado() Then Exit Sub
  Dim n As Long, c As New Collection, p, i As Long
  n = ProponerDestinos()
  If Not gPropuestas Is Nothing Then
    For i = 1 To gPropuestas.Count
      p = gPropuestas(i)
      c.Add Array(p(0), p(1), p(2), p(3) & " / " & p(4) & " / " & p(5), IIf(Len(p(6)) > 0, p(6), "(vacío)") & " > " & p(7), p(8))
    Next
  End If
  DefinirColumnas "DESTINOS": VistaBotones
  LlenarLista c, 6
  lblVista.Caption = "DESTINOS PROPUESTOS POR REGLAS: estos pedidos cambiarían de destino (afecta TMS, TRAMACO, DESPACHOS y etiquetas)"
  lblKpi.Caption = IIf(n = 0, "Ningún pedido cambia de destino. Sigue con el paso 4 (etiquetas).", n & " pedido(s). Revisa y pulsa 'Aplicar seleccionados' o 'Aplicar todos'.")
End Sub

Private Sub bAplSel_Click()
  Dim sel As New Collection, i As Long
  If mVista <> "DESTINOS" Then Exit Sub
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then sel.Add i + 1
  Next
  If sel.Count = 0 Then MsgBox "Selecciona uno o varios pedidos (Ctrl + clic).", vbInformation: Exit Sub
  If MsgBox("¿Confirmar el nuevo destino de " & sel.Count & " pedido(s)?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  AplicarDestinos sel
  b3_Click
End Sub

Private Sub bAplTod_Click()
  If mVista <> "DESTINOS" Or lst.ListCount = 0 Then Exit Sub
  If MsgBox("¿Confirmar el nuevo destino de los " & lst.ListCount & " pedido(s) de la lista?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  AplicarDestinos
  b3_Click
End Sub

Private Sub b4a_Click()
  If Ocupado() Then Exit Sub
  ImprimirEtiquetas "PENDIENTES"
End Sub
Private Sub b4b_Click()
  Dim s As String
  If Ocupado() Then Exit Sub
  s = InputBox("Número(s) de pedido a reimprimir, separados por coma:", "Reimprimir etiquetas")
  If Len(Trim$(s)) > 0 Then ImprimirEtiquetas "PEDIDOS", s
End Sub
Private Sub b4c_Click()
  If Ocupado() Then Exit Sub
  ImprimirEtiquetas "TODAS"
End Sub
Private Sub b4d_Click()
  If Ocupado() Then Exit Sub
  ImprimirEtiquetas "PENDIENTES", "", True
End Sub
Private Sub b4e_Click()
  ElegirImpresora
End Sub

Private Sub b5_Click()
  If Ocupado() Then Exit Sub
  Dim c As New Collection, res As String
  AvancePedidos c, res
  DefinirColumnas "AVANCE": VistaBotones
  LlenarLista c, 5
  lblVista.Caption = "AVANCE DEL DÍA: picking (ITEMS API vs ITEMS DEPOT) y empaque (EMPAQUETADO)"
  lblKpi.Caption = res
End Sub

Private Function HojasElegidas() As Variant
  Dim s As String
  If chkTMS.Value Then s = s & "|TMS"
  If chkTRA.Value Then s = s & "|TRAMACO"
  If chkDES.Value Then s = s & "|DESPACHOS"
  If Len(s) = 0 Then HojasElegidas = Empty Else HojasElegidas = Split(Mid$(s, 2), "|")
End Function

Private Sub b6_Click()
  If Ocupado() Then Exit Sub
  Dim h: h = HojasElegidas()
  If IsEmpty(h) Then MsgBox "Marca al menos una hoja.", vbInformation: Exit Sub
  ExportarReportes cboFmt.Text, h, False
End Sub
Private Sub b6m_Click()
  If Ocupado() Then Exit Sub
  Dim h: h = HojasElegidas()
  If IsEmpty(h) Then MsgBox "Marca al menos una hoja.", vbInformation: Exit Sub
  ExportarReportes cboFmt.Text, h, True
End Sub

' ---------- herramientas ----------
Private Sub bRep_Click()
  If Ocupado() Then Exit Sub
  RepararFormulasEGR
End Sub
Private Sub bReg_Click()
  CrearHojaReglas
  RuedaDesactivar
  Me.Hide
  MostrarExcel
  With ThisWorkbook.Worksheets(HREG)
    .Visible = xlSheetVisible
    .Activate
  End With
End Sub
Private Sub bQuit_Click()
  If MsgBox("¿Quitar TODOS los destinos confirmados por reglas?", vbYesNo + vbExclamation) = vbYes Then QuitarDestinosConfirmados
End Sub
Private Sub bDesb_Click()
  Desbloquear
End Sub
Private Sub bExcel_Click()
  RuedaDesactivar
  Me.Hide
  MostrarExcel
End Sub
Private Sub bCer_Click()
  Unload Me
End Sub
Private Sub bLogL_Click()
  txtLog.Text = ""
End Sub

Private Sub bIr_Click()
  Dim i As Long, r As Long
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then r = Val(lst.List(i, 0)): Exit For
  Next
  If r = 0 Then MsgBox "Selecciona un registro.", vbInformation: Exit Sub
  RuedaDesactivar
  Me.Hide
  MostrarExcel
  With ThisWorkbook.Worksheets(HDAT)
    .Activate
    .Rows(r).Hidden = False
    .Cells(r, 2).Select
  End With
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  bIr_Click
End Sub

Private Sub lst_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  RuedaActivar lst, TITULO
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  RuedaDesactivar
End Sub
