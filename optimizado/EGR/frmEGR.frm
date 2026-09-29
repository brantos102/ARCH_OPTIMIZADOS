Option Explicit
' =====================================================================================
'  frmEGR - Panel de despacho (Formato EGR_FL_HYCITE). Insertar > UserForm, nombre frmEGR,
'  pegar este código. Los controles se crean al abrir. Requiere modEGR, modZebra, modVentanas
'  y el formulario frmEtiquetas.
'  Vista inicial: TODOS los pedidos de DATOS con cobertura TMS, destino actual, destino sugerido
'  por reglas, courier, gestor, trayecto, zona peligrosa (desde PEDIDOS HCE), etiqueta y empaque.
' =====================================================================================
Private Const BASE_W As Single = 1220
Private Const BASE_H As Single = 700
Private Const TITULO As String = "Panel de despacho HYCITE - Etapa 2: destinos, etiquetas, empaque y reportes"
Private Const X0 As Single = 206
Private Const ANCHO As Single = 1000
Private Const NF As Long = 30

Private mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean
Private mPrimera As Boolean, mVista As String, mData() As Variant, mN As Long, mIdx() As Long, mNIdx As Long, mListo As Boolean
Private lblVista As MSForms.Label, lblKpi As MSForms.Label, lblDet As MSForms.Label, lblFilasExp As MSForms.Label
Private hdr(1 To 14) As MSForms.Label, txtLog As MSForms.TextBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents cboFiltro As MSForms.ComboBox
Private WithEvents txtBuscar As MSForms.TextBox
Private WithEvents bLista As MSForms.CommandButton
Private WithEvents bCob As MSForms.CommandButton
Private WithEvents bCamb As MSForms.CommandButton
Private WithEvents bSugSel As MSForms.CommandButton
Private WithEvents bSugTod As MSForms.CommandButton
Private cboDest As MSForms.ComboBox
Private WithEvents bAsig As MSForms.CommandButton
Private WithEvents bQuitar As MSForms.CommandButton
Private WithEvents bEtqSel As MSForms.CommandButton
Private WithEvents bEtqTod As MSForms.CommandButton
Private WithEvents bVPed As MSForms.CommandButton
Private WithEvents bVEmp As MSForms.CommandButton
Private cboHoja As MSForms.ComboBox, cboFmt As MSForms.ComboBox
Private WithEvents bExp As MSForms.CommandButton
Private WithEvents bExpM As MSForms.CommandButton
Private WithEvents bAct As MSForms.CommandButton
Private WithEvents bRep As MSForms.CommandButton
Private WithEvents bReg As MSForms.CommandButton
Private WithEvents bDesb As MSForms.CommandButton
Private WithEvents bExcel As MSForms.CommandButton
Private WithEvents bIr As MSForms.CommandButton
Private WithEvents bCer As MSForms.CommandButton
Private WithEvents bLogL As MSForms.CommandButton
' columnas de mData
'  1 SEÑAL  2 FILA  3 PEDIDO  4 DESTINATARIO  5 PROV  6 CANTON  7 PARROQUIA  8 DESTINO  9 SUGERIDO
' 10 COURIER 11 TRAYECTO 12 ZONA 13 ETIQUETA 14 EMPAQUE 15 MOTIVO SUG 16 DIRECCION 17 GESTOR COB
' 18 GESTOR SUG 19 TIPO ENTREGA 20 COB TMS 21 CONFIRMADO 22 CAJAS 23 PESO 24 BULTOS 25 DESTINO HCE
' 26 CP TMS 27 VOL% 28 UNIDADES 29 PESO CAJAS 30 TELEFONO

Private Sub UserForm_Initialize()
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H
  Dim t As MSForms.Label, i As Long, f
  ' ----- columna izquierda -----
  Set t = NL("1  COBERTURA Y DESTINOS", 8, 6, 190, True): t.ForeColor = RGB(48, 84, 150)
  Set bCob = NB("Revisar cobertura TMS", 8, 22, 190, 24, RGB(47, 117, 181), "Muestra solo los pedidos fuera de la cobertura TMS, sin código postal o sin teléfono (segunda comprobación; no revalida PEDIDOS HCE).")
  Set bCamb = NB("Ver cambios sugeridos", 8, 49, 190, 24, RGB(237, 125, 49), "Muestra los pedidos cuyo destino cambia por REGLAS_DESTINO (columna SUGER.).")
  Set bSugSel = NB("Aplicar sugerencia (selecc.)", 8, 76, 190, 24, RGB(237, 125, 49), "Confirma el destino sugerido en los pedidos seleccionados (Ctrl / Shift + clic para varios).")
  Set bSugTod = NB("Aplicar todas las sugerencias", 8, 103, 190, 24, RGB(191, 90, 0), "Confirma el destino sugerido en todos los pedidos que tienen cambio.")
  NL "Asignar a mano:", 8, 134, 80
  Set cboDest = Me.Controls.Add("Forms.ComboBox.1")
  cboDest.Left = 88: cboDest.Top = 130: cboDest.Width = 50: cboDest.Height = 18: cboDest.Style = fmStyleDropDownList
  For Each f In Array("PRO", "GYE", "UIO", "GPS"): cboDest.AddItem f: Next
  cboDest.ListIndex = 0
  Set bAsig = NB("Asignar", 142, 130, 56, 20, RGB(112, 48, 160), "Asigna el destino elegido a los pedidos seleccionados (decisión del operario).")
  Set bQuitar = NB("Quitar confirmación (selecc.)", 8, 155, 190, 22, RGB(120, 120, 120), "Los pedidos seleccionados vuelven a la regla base por provincia.")

  Set t = NL("2  ETIQUETAS (Zebra 10 x 5 cm)", 8, 186, 190, True): t.ForeColor = RGB(0, 128, 96)
  Set bEtqSel = NB("Imprimir seleccionadas", 8, 202, 93, 28, RGB(0, 128, 96), "Vista previa e impresión de las etiquetas de los pedidos seleccionados en la lista (usa el filtro para elegirlos).")
  Set bEtqTod = NB("Imprimir todas", 105, 202, 93, 28, RGB(0, 150, 110), "Vista previa e impresión de todas las etiquetas del día (pedidos con destino).")

  Set t = NL("3  SEGUIMIENTO", 8, 238, 190, True): t.ForeColor = RGB(112, 48, 160)
  Set bVPed = NB("Ver pedidos", 8, 254, 93, 24, RGB(89, 89, 89), "Vista principal: todos los pedidos con destino, cobertura y señales.")
  Set bVEmp = NB("Avance empaque", 105, 254, 93, 24, RGB(112, 48, 160), "Por pedido: picking (ITEMS API vs ITEMS DEPOT), cajas, peso, volumen % (datos de las tablas dinámicas).")

  Set t = NL("4  EXPORTAR REPORTES", 8, 286, 190, True): t.ForeColor = RGB(0, 97, 0)
  NL "Hoja:", 8, 305, 30
  Set cboHoja = Me.Controls.Add("Forms.ComboBox.1")
  cboHoja.Left = 38: cboHoja.Top = 301: cboHoja.Width = 90: cboHoja.Height = 18: cboHoja.Style = fmStyleDropDownList
  For Each f In Array("TRAMACO", "TMS", "DESPACHOS", "LAS TRES"): cboHoja.AddItem f: Next
  cboHoja.ListIndex = 0
  Set cboFmt = Me.Controls.Add("Forms.ComboBox.1")
  cboFmt.Left = 132: cboFmt.Top = 301: cboFmt.Width = 66: cboFmt.Height = 18: cboFmt.Style = fmStyleDropDownList
  For Each f In Array("CSV", "XLSX", "PDF"): cboFmt.AddItem f: Next
  cboFmt.ListIndex = 0
  Set lblFilasExp = NL("", 8, 323, 190): lblFilasExp.Height = 26: lblFilasExp.WordWrap = True: lblFilasExp.ForeColor = RGB(90, 90, 90)
  Set bExp = NB("Exportar", 8, 350, 93, 26, RGB(0, 97, 0), "Valida y exporta solo las filas con datos de la hoja elegida, sin cambiar sus columnas (plantillas de carga).")
  Set bExpM = NB("Exportar + correo", 105, 350, 93, 26, RGB(0, 97, 0), "Exporta y deja un borrador de Outlook con los archivos adjuntos.")

  Set t = NL("OPERACIÓN (en cualquier momento)", 8, 388, 190, True): t.ForeColor = RGB(89, 89, 89)
  Set bAct = NB("Actualizar datos (items, empaque, tablas)", 8, 404, 190, 26, RGB(89, 89, 89), "Trae ITEMS API, ITEMS DEPOT, EMPAQUETADO (Google Sheets) y actualiza las tablas dinámicas. Si una fuente falla, pregunta si sigue.")
  Set bRep = NB("Reparar fórmulas (una vez)", 8, 434, 190, 22, RGB(192, 80, 77), "Corrige los #REF! y prepara DATOS!A para los destinos confirmados. Guarda respaldo antes.")
  Set bReg = NB("Editar reglas", 8, 459, 93, 22, RGB(120, 120, 120), "Abre la hoja REGLAS_DESTINO (supervisor).")
  Set bIr = NB("Ir a la fila", 105, 459, 93, 22, RGB(120, 120, 120), "Oculta el panel y selecciona el pedido en DATOS.")
  Set bDesb = NB("Desbloquear", 8, 484, 93, 22, RGB(120, 120, 120), "Restaura pantalla, eventos y cálculo si Excel quedó bloqueado.")
  Set bExcel = NB("Ver Excel", 105, 484, 93, 22, RGB(0, 97, 0), "Oculta el panel. Para volver: Complementos > Panel EGR.")
  Set bCer = NB("Cerrar panel", 8, 514, 190, 24, RGB(192, 80, 77), "Cierra el panel.")
  Set t = NL("Flujo: PEDIDOS HCE (Enviar a EGR) > revisar cobertura TMS > confirmar destinos > etiquetas > empaque (Actualizar datos) > exportar. Si cambia un destino: reimprimir etiqueta y volver a exportar.", 8, 544, 190)
  t.Height = 70: t.WordWrap = True: t.BackColor = RGB(255, 242, 204): t.BorderStyle = fmBorderStyleSingle: t.ForeColor = RGB(128, 64, 0)

  ' ----- área derecha -----
  Set lblVista = NL("PEDIDOS DEL DÍA", X0, 6, 400, True): lblVista.ForeColor = RGB(48, 84, 150)
  NL "Filtro:", X0 + 410, 8, 32
  Set cboFiltro = Me.Controls.Add("Forms.ComboBox.1")
  cboFiltro.Left = X0 + 444: cboFiltro.Top = 4: cboFiltro.Width = 180: cboFiltro.Height = 18: cboFiltro.Style = fmStyleDropDownList: cboFiltro.ListRows = 16
  For Each f In Array("TODOS", "CON SEÑAL (no OK)", "CAMBIO SUGERIDO", "FUERA COBERTURA TMS", "ZONA PELIGROSA", "CONFIRMADOS", _
                      "ETIQUETA PENDIENTE", "REIMPRIMIR ETIQUETA", "SIN EMPACAR", "EMPACADOS", "DESTINO PRO", "DESTINO GYE", "DESTINO UIO", "DESTINO GPS")
    cboFiltro.AddItem f
  Next
  NL "Buscar:", X0 + 634, 8, 38
  Set txtBuscar = Me.Controls.Add("Forms.TextBox.1")
  txtBuscar.Left = X0 + 674: txtBuscar.Top = 4: txtBuscar.Width = 200: txtBuscar.Height = 18
  txtBuscar.ControlTipText = "Busca en todas las columnas (pedido, nombre, parroquia, courier...)."
  Set bLista = NB("Actualizar lista", X0 + 884, 3, 116, 20, RGB(120, 120, 120), "Vuelve a leer DATOS (después de editar en Excel o de actualizar datos).")
  Set t = NL("SEÑAL:  X fuera cobertura TMS   !! zona peligrosa   > cambio sugerido   @ reimprimir etiqueta   ! verificar sector   * destino confirmado   OK sin observación", X0, 27, ANCHO)
  t.ForeColor = RGB(90, 90, 90)
  Set lblKpi = NL("", X0, 42, ANCHO, True): lblKpi.ForeColor = RGB(48, 84, 150)
  For i = 1 To 14
    Set hdr(i) = NL("", X0, 58, 10, True)
    hdr(i).BackColor = RGB(48, 84, 150): hdr(i).ForeColor = vbWhite: hdr(i).Font.Size = 8: hdr(i).Visible = False
  Next
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = X0: lst.Top = 72: lst.Width = ANCHO: lst.Height = 380: lst.Font.Size = 8
  lst.MultiSelect = fmMultiSelectExtended
  lst.ControlTipText = "Clic = detalle. Ctrl/Shift + clic = varios. Doble clic = ir a la fila en DATOS."
  Set t = NL("DETALLE DEL PEDIDO", X0, 458, 590, True): t.ForeColor = RGB(48, 84, 150)
  Set lblDet = NL("Selecciona un pedido para ver el detalle.", X0, 474, 590)
  lblDet.Height = 212: lblDet.WordWrap = True: lblDet.BorderStyle = fmBorderStyleSingle: lblDet.BackColor = RGB(248, 248, 248): lblDet.Font.Size = 9
  Set t = NL("REGISTRO DE ACTIVIDAD (historial en la hoja oculta LOG_EGR)", X0 + 600, 458, 330, True): t.ForeColor = RGB(48, 84, 150)
  Set bLogL = NB("Limpiar", X0 + 940, 456, 60, 16, RGB(150, 150, 150), "Vacía el registro en pantalla.")
  Set txtLog = Me.Controls.Add("Forms.TextBox.1")
  txtLog.Left = X0 + 600: txtLog.Top = 474: txtLog.Width = 400: txtLog.Height = 212
  txtLog.MultiLine = True: txtLog.ScrollBars = fmScrollBarsVertical: txtLog.Locked = True: txtLog.WordWrap = True
  txtLog.Font.Name = "Consolas": txtLog.Font.Size = 8: txtLog.BackColor = RGB(30, 30, 30): txtLog.ForeColor = RGB(220, 220, 220)
  CargarLogPrevio

  mLay = CapturarLayout(Me, mCW, mCH): mEsc = 1
  HacerRedimensionable TITULO
  AjustarAPantalla Me, BASE_W, BASE_H, 0.96
  Reescalar
  mVista = "PEDIDOS"
  cboFiltro.ListIndex = 0
  mListo = True
  Recargar
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
  If mListo Then PonerColumnas
  mEscalando = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  RuedaDesactivar
End Sub

Private Sub UserForm_Activate()
  If Not mPrimera Then mPrimera = True: Exit Sub       ' la primera vez ya cargó Initialize
  If mListo Then Recargar
End Sub

' ---------- creación de controles ----------
Private Function NL(cap As String, x As Single, y As Single, w As Single, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = 14: c.Font.Bold = bold
  Set NL = c
End Function
Private Function NB(cap As String, x As Single, y As Single, w As Single, h As Single, col As Long, tip As String) As MSForms.CommandButton
  Dim c As MSForms.CommandButton: Set c = Me.Controls.Add("Forms.CommandButton.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = h
  c.BackColor = col: c.ForeColor = vbWhite: c.Font.Bold = True: c.Font.Size = 8
  c.ControlTipText = tip: c.TakeFocusOnClick = False: c.WordWrap = True
  Set NB = c
End Function

' ---------- registro ----------
Public Sub AgregarLog(ByVal linea As String)
  If txtLog Is Nothing Then Exit Sub
  Dim t As String
  t = txtLog.Text
  If Len(t) > 60000 Then t = Right$(t, 40000)
  txtLog.Text = t & IIf(Len(t) > 0, vbCrLf, "") & linea
  txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub CargarLogPrevio()
  Dim it, t As String
  If gLogE Is Nothing Then Exit Sub
  For Each it In gLogE: t = t & IIf(Len(t) > 0, vbCrLf, "") & it: Next
  txtLog.Text = t: txtLog.SelStart = Len(t)
End Sub

' Antes de cualquier acción: sin gancho de rueda (evita cierres de Excel con mensajes o procesos largos)
Private Function Listo() As Boolean
  RuedaDesactivar
  If gOcupadoE Then MsgBox "Hay un proceso en curso. Espera a que termine.", vbInformation: Exit Function
  Listo = True
End Function

' =====================================================================================
'  Datos
' =====================================================================================
Public Sub Recargar()
  Dim ws As Worksheet, wsE As Worksheet, lr As Long, v, e, i As Long, n As Long, reglas, dC As Object, dE As Object
  Dim sug As String, mot As String, dest As String, zona As String, k As String, cob, fuera As Boolean, conf As Boolean
  Dim etq As String, emp As String, x
  On Error GoTo fallo
  Application.Cursor = xlWait
  Set ws = ThisWorkbook.Worksheets(HDAT)
  Set wsE = ThisWorkbook.Worksheets("ETIQUETAS")
  lr = UltimaFilaDatos()
  mN = 0
  If lr < 2 Then ReDim mData(1 To 1, 1 To NF): GoTo mostrar
  v = ws.Range(ws.Cells(1, 1), ws.Cells(lr, 41)).Value
  e = wsE.Range(wsE.Cells(1, 1), wsE.Cells(lr, 14)).Value
  reglas = CargarReglas()
  Set dC = DictCobertura()
  Set dE = DictEmpaque()
  ReDim mData(1 To lr, 1 To NF)
  For i = 2 To lr
    If Len(TXE(v(i, D_PED))) > 0 Then
      n = n + 1
      dest = UCase$(TXE(v(i, D_DEST)))
      fuera = (UCase$(TXE(v(i, D_VAL))) = "REVISAR" Or Len(TXE(v(i, D_PROV))) = 0 Or UCase$(TXE(v(i, D_PROV))) = "VALIDAR")
      sug = "": mot = ""
      If Not fuera Then sug = DestinoPorReglas(reglas, TXE(v(i, D_PROV)), TXE(v(i, D_CANT)), TXE(v(i, D_PARR)), TXE(v(i, D_DIR)), mot)
      conf = (Len(TXE(v(i, D_RDEST))) > 0 And TXE(v(i, D_RPED)) = TXE(v(i, D_PED)))
      k = ClaveTMS(TXE(v(i, D_H)), TXE(v(i, D_I)), TXE(v(i, D_J)))
      cob = Empty: If dC.Exists(k) Then cob = dC(k)
      zona = TXE(v(i, 39))                                          ' AM zona (desde PEDIDOS HCE)
      ' etiqueta
      If UCase$(TXE(e(i, 6))) = "OK" Then
        If Len(TXE(e(i, 13))) > 0 And UCase$(TXE(e(i, 13))) <> dest Then etq = "REIMPRIMIR" Else etq = "IMPRESA"
      Else
        etq = "PENDIENTE"
      End If
      emp = EstadoEmpaque(dE, TXE(v(i, D_PED)))
      mData(n, 2) = i: mData(n, 3) = TXE(v(i, D_PED)): mData(n, 4) = TXE(v(i, D_NOM))
      mData(n, 5) = TXE(v(i, D_PROV)): mData(n, 6) = TXE(v(i, D_CANT)): mData(n, 7) = TXE(v(i, D_PARR))
      mData(n, 8) = dest: mData(n, 9) = IIf(Len(sug) > 0 And sug <> dest, sug, "")
      mData(n, 10) = TXE(v(i, D_COUR))
      mData(n, 11) = IIf(Len(TXE(v(i, 37))) > 0, TXE(v(i, 37)), IIf(IsArray(cob), cob(2), ""))      ' AK trayecto o cobertura V
      mData(n, 12) = zona: mData(n, 13) = etq: mData(n, 14) = emp
      mData(n, 15) = mot: mData(n, 16) = TXE(v(i, D_DIR))
      mData(n, 17) = IIf(IsArray(cob), cob(0), "")
      mData(n, 18) = IIf(Len(TXE(v(i, 41))) > 0, TXE(v(i, 41)), IIf(IsArray(cob), cob(1), ""))      ' AO sugerido R
      mData(n, 19) = TXE(v(i, 38))                                                                   ' AL tipo de entrega
      mData(n, 20) = IIf(fuera, "FUERA: " & TXE(v(i, D_DIAG)) & IIf(Len(TXE(v(i, D_SUG))) > 0, " (sugerida: " & TXE(v(i, D_SUG)) & ")", ""), "OK")
      mData(n, 21) = IIf(conf, TXE(v(i, D_RMOT)), "")
      mData(n, 23) = TXE(v(i, D_PESO)): mData(n, 24) = TXE(v(i, D_BUL)): mData(n, 25) = TXE(v(i, 36))
      mData(n, 26) = TXE(v(i, D_CPTMS)): mData(n, 30) = TXE(v(i, D_TEL))
      If dE.Exists(TXE(v(i, D_PED))) Then
        x = dE(TXE(v(i, D_PED)))
        mData(n, 22) = x(1): mData(n, 27) = x(3): mData(n, 28) = x(4) & "/" & x(5): mData(n, 29) = x(2)
      End If
      ' señal (prioridad)
      Select Case True
        Case fuera: mData(n, 1) = "X FUERA TMS"
        Case Left$(UCase$(zona), 4) = "ZONA": mData(n, 1) = "!! ZONA"
        Case Len(mData(n, 9)) > 0: mData(n, 1) = "> CAMBIO"
        Case etq = "REIMPRIMIR": mData(n, 1) = "@ REIMPRIMIR"
        Case InStr(UCase$(zona), "POSIBLE") > 0: mData(n, 1) = "! SECTOR"
        Case conf: mData(n, 1) = "* CONFIRMADO"
        Case Else: mData(n, 1) = "OK"
      End Select
    End If
  Next
  mN = n
mostrar:
  Application.Cursor = xlDefault
  ContarExport
  Filtrar
  Exit Sub
fallo:
  Application.Cursor = xlDefault
  LogE "PANEL: error al leer los datos: " & Err.Description, "ERROR"
End Sub

Private Sub ContarExport()
  Dim s As String, h
  On Error Resume Next
  For Each h In Array("TRAMACO", "TMS", "DESPACHOS")
    s = s & IIf(Len(s) > 0, " · ", "") & h & " " & FilasExport(CStr(h)).Count
  Next
  lblFilasExp.Caption = "Filas con datos: " & s
End Sub

Private Function Pasa(ByVal r As Long) As Boolean
  Dim f As String
  f = cboFiltro.Text
  Select Case f
    Case "", "TODOS": Pasa = True
    Case "CON SEÑAL (no OK)": Pasa = (mData(r, 1) <> "OK")
    Case "CAMBIO SUGERIDO": Pasa = (Len(mData(r, 9)) > 0)
    Case "FUERA COBERTURA TMS": Pasa = (Left$(mData(r, 1), 1) = "X" Or Len(mData(r, 26)) = 0 Or Len(mData(r, 30)) = 0)
    Case "ZONA PELIGROSA": Pasa = (Len(mData(r, 12)) > 0)
    Case "CONFIRMADOS": Pasa = (Len(mData(r, 21)) > 0)
    Case "ETIQUETA PENDIENTE": Pasa = (mData(r, 13) = "PENDIENTE")
    Case "REIMPRIMIR ETIQUETA": Pasa = (mData(r, 13) = "REIMPRIMIR")
    Case "SIN EMPACAR": Pasa = (Left$(mData(r, 14), 8) <> "EMPACADO")
    Case "EMPACADOS": Pasa = (Left$(mData(r, 14), 8) = "EMPACADO")
    Case Else
      If Left$(f, 8) = "DESTINO " Then Pasa = (mData(r, 8) = Mid$(f, 9))
  End Select
End Function

Private Sub Filtrar()
  Dim r As Long, j As Long, q As String, s As String, cols, arr(), k As Long, nc As Long
  Dim nFue As Long, nZon As Long, nCam As Long, nRei As Long, nEmp As Long, nPen As Long, nPro As Long, nGye As Long, nUio As Long, nGps As Long
  If Not mListo Then Exit Sub
  q = UCase$(Trim$(txtBuscar.Text))
  ReDim mIdx(1 To IIf(mN > 0, mN, 1)): mNIdx = 0
  For r = 1 To mN
    If Pasa(r) Then
      If Len(q) > 0 Then
        s = ""
        For j = 1 To NF: s = s & " " & UCase$(CStr(mData(r, j) & "")): Next
        If InStr(s, q) = 0 Then GoTo sig
      End If
      mNIdx = mNIdx + 1: mIdx(mNIdx) = r
    End If
sig:
    If Left$(mData(r, 1), 1) = "X" Then nFue = nFue + 1
    If Len(mData(r, 12)) > 0 Then nZon = nZon + 1
    If Len(mData(r, 9)) > 0 Then nCam = nCam + 1
    If mData(r, 13) = "REIMPRIMIR" Then nRei = nRei + 1
    If mData(r, 13) = "PENDIENTE" Then nPen = nPen + 1
    If Left$(mData(r, 14), 8) = "EMPACADO" Then nEmp = nEmp + 1
    Select Case mData(r, 8)
      Case "PRO": nPro = nPro + 1
      Case "GYE": nGye = nGye + 1
      Case "UIO": nUio = nUio + 1
      Case "GPS": nGps = nGps + 1
    End Select
  Next
  lblKpi.Caption = mN & " pedidos (" & mNIdx & " en lista)  |  PRO " & nPro & " · GYE " & nGye & " · UIO " & nUio & " · GPS " & nGps & _
                   "  |  fuera TMS " & nFue & " · zonas " & nZon & " · cambios sugeridos " & nCam & _
                   "  |  etiquetas pendientes " & nPen & " · reimprimir " & nRei & "  |  empacados " & nEmp & "/" & mN & _
                   IIf(mN > 0, " (" & Format(nEmp / mN, "0%") & ")", "")
  cols = ColumnasVista()
  nc = UBound(cols) + 1
  PonerColumnas
  lst.Clear
  If mNIdx = 0 Then Exit Sub
  ReDim arr(0 To mNIdx - 1, 0 To nc - 1)
  For k = 1 To mNIdx
    For j = 0 To nc - 1
      arr(k - 1, j) = CStr(mData(mIdx(k), cols(j)) & "")
    Next
  Next
  lst.ColumnCount = nc
  lst.List = arr
End Sub

Private Function ColumnasVista() As Variant
  If mVista = "EMPAQUE" Then
    ColumnasVista = Array(2, 3, 4, 8, 14, 22, 29, 27, 28, 24, 23, 13, 10)
  Else
    ColumnasVista = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
  End If
End Function

Private Sub PonerColumnas()
  Dim nom, anc, i As Long, x As Single, w As String, e As Double
  If mVista = "EMPAQUE" Then
    nom = Array("FILA", "PEDIDO", "DESTINATARIO", "DEST.", "EMPAQUE", "CAJAS", "PESO CAJAS", "VOL. %", "UNID. CONF/SOL", "BULTOS", "PESO KG", "ETIQUETA", "COURIER")
    anc = Array(30, 62, 170, 40, 110, 40, 64, 50, 86, 46, 56, 70, 90)
  Else
    nom = Array("SEÑAL", "FILA", "PEDIDO", "DESTINATARIO", "PROVINCIA", "CANTÓN", "PARROQUIA", "DEST.", "SUGER.", "COURIER", "TRAYECTO", "ZONA", "ETIQUETA", "EMPAQUE")
    anc = Array(70, 28, 60, 120, 72, 80, 100, 36, 42, 70, 88, 88, 62, 72)
  End If
  e = mEsc: If e <= 0 Then e = 1
  x = lst.Left + 3 * e
  For i = 1 To 14
    If i - 1 <= UBound(nom) Then
      hdr(i).Caption = " " & nom(i - 1): hdr(i).Left = x: hdr(i).Width = anc(i - 1) * e - 1: hdr(i).Visible = True
      x = x + anc(i - 1) * e
      w = w & IIf(Len(w) > 0, ";", "") & CStr(CLng(anc(i - 1) * e))
    Else
      hdr(i).Visible = False
    End If
  Next
  lst.ColumnWidths = w
End Sub

Private Function Seleccionados() As Collection
  Dim c As New Collection, i As Long
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then c.Add mIdx(i + 1)
  Next
  Set Seleccionados = c
End Function

Private Sub MostrarDetalle(ByVal r As Long)
  Dim s As String
  s = "Pedido " & mData(r, 3) & "  (fila " & mData(r, 2) & ")   " & mData(r, 4) & "   Tel: " & mData(r, 30) & vbCrLf & _
      "Dirección: " & mData(r, 16) & vbCrLf & _
      "Cobertura TMS: " & mData(r, 5) & " / " & mData(r, 6) & " / " & mData(r, 7) & "   CP TMS: " & mData(r, 26) & "   -> " & mData(r, 20) & vbCrLf & vbCrLf & _
      "DESTINO ACTUAL: " & IIf(Len(mData(r, 8)) > 0, mData(r, 8), "(sin destino)") & "     Destino en PEDIDOS HCE: " & IIf(Len(mData(r, 25)) > 0, mData(r, 25), "-") & vbCrLf
  If Len(mData(r, 9)) > 0 Then
    s = s & "SUGERIDO POR REGLAS: " & mData(r, 9) & "  <- " & mData(r, 15) & vbCrLf & "  (pulsa 'Aplicar sugerencia' para confirmarlo o déjalo como está)" & vbCrLf
  ElseIf Len(mData(r, 15)) > 0 Then
    s = s & "Regla aplicada: " & mData(r, 15) & vbCrLf
  End If
  If Len(mData(r, 21)) > 0 Then s = s & "Confirmado: " & mData(r, 21) & vbCrLf
  s = s & vbCrLf & "Courier: " & mData(r, 10) & "   Gestor en cobertura: " & mData(r, 17) & "   Gestor sugerido (R): " & IIf(Len(mData(r, 18)) > 0, mData(r, 18), "-") & vbCrLf & _
      "Trayecto: " & mData(r, 11) & "   Tipo de entrega: " & IIf(Len(mData(r, 19)) > 0, mData(r, 19), "NORMAL") & vbCrLf & _
      "Zona peligrosa: " & IIf(Len(mData(r, 12)) > 0, mData(r, 12), "no") & vbCrLf & _
      "Etiqueta: " & mData(r, 13) & IIf(mData(r, 13) = "REIMPRIMIR", " (se imprimió con otro destino: vuelve a imprimirla y a exportar)", "") & vbCrLf & _
      "Empaque: " & mData(r, 14) & "   Cajas: " & mData(r, 22) & "   Bultos: " & mData(r, 24) & "   Peso: " & mData(r, 23) & " kg   Vol.: " & mData(r, 27)
  lblDet.Caption = s
End Sub

' =====================================================================================
'  Eventos de la lista y filtros
' =====================================================================================
Private Sub cboFiltro_Change()
  Filtrar
End Sub
Private Sub txtBuscar_Change()
  Filtrar
End Sub
Private Sub bLista_Click()
  If Not Listo() Then Exit Sub
  Recargar
End Sub

Private Sub lst_Change()
  Dim i As Long
  If lst.ListIndex >= 0 And lst.ListIndex < mNIdx Then
    If lst.Selected(lst.ListIndex) Then MostrarDetalle mIdx(lst.ListIndex + 1): Exit Sub
  End If
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then MostrarDetalle mIdx(i + 1): Exit Sub
  Next
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

' =====================================================================================
'  1. Cobertura y destinos
' =====================================================================================
Private Sub bCob_Click()
  If Not Listo() Then Exit Sub
  Dim c As New Collection
  RevisarCoberturaTMS c                     ' deja el detalle en el registro
  mVista = "PEDIDOS": lblVista.Caption = "COBERTURA TMS: pedidos con observación"
  cboFiltro.Text = "FUERA COBERTURA TMS"
  Filtrar
End Sub

Private Sub bCamb_Click()
  If Not Listo() Then Exit Sub
  mVista = "PEDIDOS": lblVista.Caption = "CAMBIOS DE DESTINO SUGERIDOS POR REGLAS (columna SUGER.)"
  cboFiltro.Text = "CAMBIO SUGERIDO"
  Filtrar
End Sub

Private Sub bSugSel_Click()
  If Not Listo() Then Exit Sub
  AplicarSug Seleccionados()
End Sub

Private Sub bSugTod_Click()
  If Not Listo() Then Exit Sub
  Dim c As New Collection, r As Long
  For r = 1 To mN
    If Len(mData(r, 9)) > 0 Then c.Add r
  Next
  AplicarSug c
End Sub

Private Sub AplicarSug(c As Collection)
  Dim r, n As Long, s As String
  For Each r In c
    If Len(mData(r, 9)) > 0 Then n = n + 1: If n <= 15 Then s = s & vbCrLf & "  " & mData(r, 3) & "  " & mData(r, 8) & " -> " & mData(r, 9) & "  (" & mData(r, 15) & ")"
  Next
  If n = 0 Then MsgBox "Ninguno de los pedidos elegidos tiene cambio sugerido.", vbInformation: Exit Sub
  If MsgBox("Confirmar el destino sugerido en " & n & " pedido(s):" & s & IIf(n > 15, vbCrLf & "  ...", "") & vbCrLf & vbCrLf & _
            "Afecta TMS, TRAMACO, DESPACHOS y etiquetas.", vbYesNo + vbQuestion, "Confirmar destinos") <> vbYes Then Exit Sub
  For Each r In c
    If Len(mData(r, 9)) > 0 Then If Not ConfirmarDestino(mData(r, 2), mData(r, 9), mData(r, 15)) Then Exit For
  Next
  Application.Calculate
  LogE "DESTINOS: " & n & " confirmado(s). Si ya tenían etiqueta, aparecen como '@ REIMPRIMIR'; vuelve a exportar los reportes."
  Recargar
End Sub

Private Sub bAsig_Click()
  If Not Listo() Then Exit Sub
  Dim c As Collection, r, d As String
  Set c = Seleccionados()
  If c.Count = 0 Then MsgBox "Selecciona uno o varios pedidos en la lista.", vbInformation: Exit Sub
  d = cboDest.Text
  If MsgBox("Asignar destino " & d & " a " & c.Count & " pedido(s) seleccionado(s)?", vbYesNo + vbQuestion, "Asignar destino") <> vbYes Then Exit Sub
  For Each r In c
    If Not ConfirmarDestino(mData(r, 2), d, "MANUAL OPERARIO") Then Exit For
  Next
  Application.Calculate
  Recargar
End Sub

Private Sub bQuitar_Click()
  If Not Listo() Then Exit Sub
  Dim c As Collection, r
  Set c = Seleccionados()
  If c.Count = 0 Then MsgBox "Selecciona uno o varios pedidos en la lista.", vbInformation: Exit Sub
  If MsgBox("Quitar la confirmación de " & c.Count & " pedido(s)? Vuelven a la regla base por provincia.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  For Each r In c: QuitarConfirmacion mData(r, 2): Next
  Application.Calculate
  Recargar
End Sub

' =====================================================================================
'  2. Etiquetas
' =====================================================================================
Private Sub bEtqSel_Click()
  If Not Listo() Then Exit Sub
  Dim c As Collection, f As New Collection, r
  Set c = Seleccionados()
  If c.Count = 0 Then MsgBox "Selecciona en la lista los pedidos a imprimir (Ctrl + clic o Shift + clic para varios). Tip: usa el filtro 'ETIQUETA PENDIENTE' o 'REIMPRIMIR ETIQUETA'.", vbInformation: Exit Sub
  For Each r In c: f.Add mData(r, 2): Next
  AbrirEtiquetas f
End Sub

Private Sub bEtqTod_Click()
  If Not Listo() Then Exit Sub
  Dim f As New Collection, r As Long
  For r = 1 To mN: f.Add mData(r, 2): Next
  AbrirEtiquetas f
End Sub

Private Sub AbrirEtiquetas(filas As Collection)
  Set gEtiqFilas = filas
  frmEtiquetas.Show vbModal
  Recargar
End Sub

' =====================================================================================
'  3. Seguimiento
' =====================================================================================
Private Sub bVPed_Click()
  If Not Listo() Then Exit Sub
  mVista = "PEDIDOS": lblVista.Caption = "PEDIDOS DEL DÍA"
  cboFiltro.Text = "TODOS"
  Filtrar
End Sub

Private Sub bVEmp_Click()
  If Not Listo() Then Exit Sub
  Dim c As New Collection, res As String
  AvancePedidos c, res                       ' KPI + avisos de costos en el registro
  mVista = "EMPAQUE": lblVista.Caption = "AVANCE DE EMPAQUE (picking, cajas, peso, volumen %)"
  Recargar
End Sub

' =====================================================================================
'  4. Exportar
' =====================================================================================
Private Sub Exportar(ByVal correo As Boolean)
  Dim h
  If cboHoja.Text = "LAS TRES" Then h = Array("TMS", "TRAMACO", "DESPACHOS") Else h = Array(cboHoja.Text)
  ExportarReportes cboFmt.Text, h, correo
  ContarExport
End Sub
Private Sub bExp_Click()
  If Not Listo() Then Exit Sub
  Exportar False
End Sub
Private Sub bExpM_Click()
  If Not Listo() Then Exit Sub
  Exportar True
End Sub

' =====================================================================================
'  Operación
' =====================================================================================
Private Sub bAct_Click()
  If Not Listo() Then Exit Sub
  ActualizarTodo
End Sub
Private Sub bRep_Click()
  If Not Listo() Then Exit Sub
  RepararFormulasEGR
  Recargar
End Sub
Private Sub bReg_Click()
  If Not Listo() Then Exit Sub
  CrearHojaReglas
  Me.Hide
  MostrarExcel
  With ThisWorkbook.Worksheets(HREG)
    .Visible = xlSheetVisible
    .Activate
  End With
End Sub
Private Sub bDesb_Click()
  RuedaDesactivar
  Desbloquear
End Sub
Private Sub bExcel_Click()
  RuedaDesactivar
  Me.Hide
  MostrarExcel
End Sub
Private Sub bCer_Click()
  RuedaDesactivar
  Unload Me
End Sub
Private Sub bLogL_Click()
  txtLog.Text = ""
End Sub

Private Sub bIr_Click()
  Dim c As Collection, r As Long
  RuedaDesactivar
  Set c = Seleccionados()
  If c.Count = 0 Then MsgBox "Selecciona un pedido.", vbInformation: Exit Sub
  r = mData(c(1), 2)
  Me.Hide
  MostrarExcel
  With ThisWorkbook.Worksheets(HDAT)
    .Activate
    .Rows(r).Hidden = False
    .Cells(r, 2).Select
  End With
End Sub
