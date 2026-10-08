Option Explicit
' =====================================================================================
'  frmPanel - Panel de validación HYCITE (versión final PEDIDOS HCE)
'  Crear: Insertar > UserForm, (Name) = frmPanel, pegar este código. Los controles se crean
'  solos al abrir. Funciona a la par con Excel (no modal).
'   - Pasos 0..7 en vertical con explicación al pasar el mouse y guía del SIGUIENTE PASO
'   - Vistas: PEDIDOS / COBERTURA / ZONAS PELIGROSAS, encabezados alineados a las columnas
'   - Filtro rápido, búsqueda general, 2 filtros dinámicos por campo y orden por cualquier campo
'   - Columna SEÑAL: zona peligrosa, verificar sector, revisar, cobertura cercana, aprobado, corregido
'   - Detalle del pedido: original vs propuesta, gestor asignado vs cobertura (Q) vs sugerido (R), trayecto (S)
' =====================================================================================
Private Const BASE_W As Single = 1180
Private Const BASE_H As Single = 700
Private Const X0 As Single = 176          ' inicio del área derecha
Private Const ANCHO As Single = 994       ' ancho del área derecha
Private Const TITULO As String = "Panel de validación HYCITE - Etapa 1: pedidos y cobertura"

Private wsD As Worksheet
Private mVista As String, mCargando As Boolean, mListo As Boolean
Private mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean, txtLog As MSForms.TextBox
Private mData() As String, mNorm() As String, mBus() As String
Private mN As Long, mNf As Long, mNd As Long
Private mHdr() As String, mW() As Single, mColFila As Long, mHoja As String
Private mIdx() As Long, mNIdx As Long
Private hdr(1 To 14) As MSForms.Label
Private bArr(0 To 7) As MSForms.CommandButton, mCap(0 To 7) As String, mColBase(0 To 7) As Long
Private lblKpi As MSForms.Label, lblGuia As MSForms.Label, lblDet As MSForms.Label

Private WithEvents cboVista As MSForms.ComboBox
Private WithEvents cboRapido As MSForms.ComboBox
Private WithEvents txtBuscar As MSForms.TextBox
Private WithEvents cboCampo1 As MSForms.ComboBox
Private WithEvents txtF1 As MSForms.TextBox
Private WithEvents cboCampo2 As MSForms.ComboBox
Private WithEvents txtF2 As MSForms.TextBox
Private WithEvents cboOrden As MSForms.ComboBox
Private WithEvents chkDesc As MSForms.CheckBox
Private WithEvents chkSeguir As MSForms.CheckBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents b0 As MSForms.CommandButton
Private WithEvents b1 As MSForms.CommandButton
Private WithEvents b2 As MSForms.CommandButton
Private WithEvents b3 As MSForms.CommandButton
Private WithEvents b4 As MSForms.CommandButton
Private WithEvents b5 As MSForms.CommandButton
Private WithEvents b6 As MSForms.CommandButton
Private WithEvents b7 As MSForms.CommandButton
Private WithEvents bCamb As MSForms.CommandButton
Private WithEvents bIr As MSForms.CommandButton
Private WithEvents bRef As MSForms.CommandButton
Private WithEvents bDesb As MSForms.CommandButton
Private WithEvents bCer As MSForms.CommandButton
Private WithEvents bLimpF As MSForms.CommandButton
Private WithEvents bCob As MSForms.CommandButton      ' Buscar cobertura
Private WithEvents bDiag As MSForms.CommandButton     ' Diagnóstico
Private WithEvents bExcel As MSForms.CommandButton    ' Ver Excel
Private WithEvents bMenos As MSForms.CommandButton    ' A-
Private WithEvents bMas As MSForms.CommandButton      ' A+
Private WithEvents bAjus As MSForms.CommandButton     ' ajustar a pantalla
Private WithEvents bLogL As MSForms.CommandButton     ' limpiar registro

' =====================================================================================
'  Construcción de la interfaz
' =====================================================================================
Private Sub UserForm_Initialize()
  Set wsD = ThisWorkbook.Worksheets("DEPOT")
  mCargando = True
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H

  ' ----- pasos (columna izquierda) -----
  Dim t As MSForms.Label, i As Long
  Set t = NewLbl("PASOS DEL PROCESO", 8, 6, 160, True): t.ForeColor = RGB(48, 84, 150)
  Set b0 = PasoBtn(0, "0 Actualizar datos", RGB(89, 89, 89), _
    "Trae los pedidos del día desde DEPOT y la lista de TMS (consultas). Luego ofrece limpiar la validación anterior.")
  Set b1 = PasoBtn(1, "1 Limpiar", RGB(127, 127, 127), _
    "Borra la validación anterior (columna L y N:AD). No toca los datos del cliente (A:K).")
  Set b2 = PasoBtn(2, "2 Validar", RGB(47, 117, 181), _
    "Valida provincia/cantón/parroquia contra COBERTURA y asigna sigla, gestor, destino, trayecto y zona peligrosa.")
  Set b3 = PasoBtn(3, "3 Revisar pendientes", RGB(237, 125, 49), _
    "Abre el validador solo con los pedidos en REVISAR para confirmarlos uno a uno.")
  Set b4 = PasoBtn(4, "4 Aprobar lote", RGB(191, 143, 0), _
    "Aprueba en bloque las sugerencias pendientes. Úsalo solo después de revisar la lista.")
  Set b5 = PasoBtn(5, "5 Aplicar aprobados", RGB(112, 173, 71), _
    "Copia la propuesta a los datos del pedido (F:H) y registra cada cambio en CAMBIOS.")
  Set b6 = PasoBtn(6, "6 Actualizar siglas", RGB(47, 117, 181), _
    "Recalcula sigla, gestor y zona si editaste a mano las columnas F:H.")
  Set b7 = PasoBtn(7, "7 Enviar a EGR", RGB(0, 128, 96), _
    "Envía los pedidos a la hoja DATOS del archivo Formato EGR_FL_HYCITE (debe estar abierto).")
  Set lblGuia = NewLbl("", 8, 284, 160, True)
  lblGuia.Height = 66: lblGuia.WordWrap = True: lblGuia.BackColor = RGB(255, 242, 204)
  lblGuia.BorderStyle = fmBorderStyleSingle: lblGuia.ForeColor = RGB(128, 64, 0)
  Set t = NewLbl("HERRAMIENTAS", 8, 356, 160, True): t.ForeColor = RGB(48, 84, 150)
  Set bCob = Btn("Buscar cobertura", 8, 372, 160, 26, RGB(47, 117, 181), _
    "Abre el buscador de COBERTURA (provincia, cantón, parroquia) con la parroquia principal, la más cercana a la dirección y el gestor.")
  Set bDiag = Btn("Diagnóstico de errores", 8, 402, 160, 26, RGB(192, 80, 77), _
    "Revisa todos los pedidos: parroquia inexistente, cantón mal redactado, duplicados, sin teléfono, zonas. El detalle queda en el registro.")
  Set bCamb = Btn("Exportar CAMBIOS", 8, 432, 160, 26, RGB(84, 130, 53), "Guarda en un archivo los cambios de hoy o todo el historial.")
  Set bIr = Btn("Ir a la fila en Excel", 8, 462, 160, 26, RGB(120, 120, 120), "Oculta el panel y selecciona en la hoja la fila del registro elegido.")
  Set bRef = Btn("Actualizar lista", 8, 492, 160, 26, RGB(120, 120, 120), "Vuelve a leer la hoja (usa después de editar en Excel).")
  Set bDesb = Btn("Desbloquear", 8, 522, 160, 26, RGB(120, 120, 120), "Restaura pantalla, eventos y cálculo si Excel quedó bloqueado.")
  Set bExcel = Btn("Ver Excel", 8, 560, 160, 28, RGB(0, 97, 0), _
    "Oculta el panel para trabajar en Excel. Para volver: Complementos > Panel HYCITE (los filtros se conservan).")
  Set bMenos = Btn("A -", 8, 594, 50, 22, RGB(120, 120, 120), "Achicar el panel.")
  Set bMas = Btn("A +", 63, 594, 50, 22, RGB(120, 120, 120), "Agrandar el panel.")
  Set bAjus = Btn("Ajustar", 118, 594, 50, 22, RGB(120, 120, 120), "Ajustar el panel a la pantalla. También puedes arrastrar el borde de la ventana.")
  Set bCer = Btn("Cerrar panel", 8, 640, 160, 26, RGB(192, 80, 77), "Cierra este panel (los datos quedan en la hoja).")

  ' ----- filtros (área derecha) -----
  NewLbl "Vista:", X0, 9, 32
  Set cboVista = NewCbo(X0 + 34, 5, 120, True)
  NewLbl "Filtro rápido:", X0 + 164, 9, 62
  Set cboRapido = NewCbo(X0 + 228, 5, 170, True)
  NewLbl "Buscar:", X0 + 408, 9, 38
  Set txtBuscar = NewTxt(X0 + 448, 5, 170)
  txtBuscar.ControlTipText = "Busca en todas las columnas mientras escribes (pedido, nombre, parroquia, gestor...)."
  Set chkSeguir = Me.Controls.Add("Forms.CheckBox.1")
  chkSeguir.Caption = "Seguir fila en Excel": chkSeguir.Left = X0 + 628: chkSeguir.Top = 5: chkSeguir.Width = 130: chkSeguir.Height = 18
  chkSeguir.ControlTipText = "Al elegir un registro, selecciona su fila en la hoja."

  NewLbl "Filtro 1:", X0, 33, 40
  Set cboCampo1 = NewCbo(X0 + 42, 29, 130, True)
  Set txtF1 = NewTxt(X0 + 176, 29, 110)
  NewLbl "Filtro 2:", X0 + 296, 33, 40
  Set cboCampo2 = NewCbo(X0 + 338, 29, 130, True)
  Set txtF2 = NewTxt(X0 + 472, 29, 110)
  NewLbl "Orden:", X0 + 592, 33, 32
  Set cboOrden = NewCbo(X0 + 626, 29, 110, True)
  Set chkDesc = Me.Controls.Add("Forms.CheckBox.1")
  chkDesc.Caption = "Desc.": chkDesc.Left = X0 + 740: chkDesc.Top = 29: chkDesc.Width = 44: chkDesc.Height = 18
  Set bLimpF = Btn("Limpiar", X0 + 786, 29, 40, 18, RGB(150, 150, 150), "Quita todos los filtros.")

  Set t = NewLbl("SEÑAL:  !! zona peligrosa   ! verificar sector   ? revisar   ~ cobertura cercana   + aprobado   * corregido   - sin validar   OK sin cambio", X0, 54, ANCHO, False)
  t.ForeColor = RGB(90, 90, 90)
  Set lblKpi = NewLbl("", X0, 70, ANCHO, True): lblKpi.ForeColor = RGB(48, 84, 150)

  For i = 1 To 14
    Set hdr(i) = NewLbl("", X0, 88, 10, True)
    hdr(i).BackColor = RGB(48, 84, 150): hdr(i).ForeColor = vbWhite: hdr(i).Height = 14
    hdr(i).Font.Size = 8: hdr(i).Visible = False
  Next
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = X0: lst.Top = 102: lst.Width = ANCHO: lst.Height = 380: lst.Font.Size = 8
  lst.ControlTipText = "Clic = ver detalle abajo. Doble clic = abrir el pedido en el validador (o ir a la fila en COBERTURA/ZONAS)."
  Set t = NewLbl("DETALLE DEL REGISTRO", X0, 488, 590, True): t.ForeColor = RGB(48, 84, 150)
  Set lblDet = NewLbl("Selecciona un registro para ver el detalle.", X0, 504, 590, False)
  lblDet.Height = 166: lblDet.WordWrap = True: lblDet.BorderStyle = fmBorderStyleSingle
  lblDet.BackColor = RGB(248, 248, 248): lblDet.Font.Size = 9

  Set t = NewLbl("REGISTRO DE ACTIVIDAD (avance de procesos y errores)", X0 + 600, 488, 330, True): t.ForeColor = RGB(48, 84, 150)
  Set bLogL = Btn("Limpiar", X0 + 934, 486, 60, 16, RGB(150, 150, 150), "Vacía el registro en pantalla (la hoja oculta LOG_PROCESO conserva el historial).")
  Set txtLog = Me.Controls.Add("Forms.TextBox.1")
  txtLog.Left = X0 + 600: txtLog.Top = 504: txtLog.Width = 394: txtLog.Height = 166
  txtLog.MultiLine = True: txtLog.ScrollBars = fmScrollBarsVertical: txtLog.Locked = True: txtLog.WordWrap = True
  txtLog.Font.Name = "Consolas": txtLog.Font.Size = 8: txtLog.BackColor = RGB(30, 30, 30): txtLog.ForeColor = RGB(220, 220, 220)
  CargarLogPrevio

  Dim f
  For Each f In Array("PEDIDOS", "COBERTURA", "ZONAS PELIGROSAS"): cboVista.AddItem f: Next
  For Each f In Array("TODOS", "CON SEÑAL (no OK)", "REVISAR", "APROBADO", "OK", "ZONA PELIGROSA", "COBERTURA CERCANA", _
                      "CORREGIDOS", "SIN APLICAR", "SIN VALIDAR", "DESTINO PRO (TRAMACO)", "DESTINO UIO", "DESTINO GYE", "DESTINO GPS")
    cboRapido.AddItem f
  Next
  mLay = CapturarLayout(Me, mCW, mCH): mEsc = 1       ' diseño al 100 %
  HacerRedimensionable TITULO
  AjustarAPantalla Me, BASE_W, BASE_H, 0.96
  Reescalar
  mCargando = False
  cboVista.ListIndex = 0          ' carga PEDIDOS
  mListo = True
  LogP "PANEL: abierto por " & Application.UserName
End Sub

' Al arrastrar el borde o maximizar, toda la interfaz se escala
Private Sub UserForm_Resize()
  Reescalar
End Sub

' Reacomoda todos los controles al tamaño actual de la ventana (sin Zoom)
Private Sub Reescalar()
  If Not IsArray(mLay) Or mEscalando Then Exit Sub
  Dim f As Double
  f = EscalaAjuste(Me, mCW, mCH)
  If Abs(f - mEsc) < 0.01 Then Exit Sub
  mEscalando = True
  mEsc = f
  EscalarLayout Me, mLay, f, mCW, mCH
  If mNd > 0 Then PonerEncabezados
  mEscalando = False
End Sub

' ---------- registro de actividad ----------
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
  If gLog Is Nothing Then Exit Sub
  For Each it In gLog
    t = t & IIf(Len(t) > 0, vbCrLf, "") & it
  Next
  txtLog.Text = t
  txtLog.SelStart = Len(t)
End Sub

' Ejecuta un paso evitando lanzar dos procesos a la vez
Private Function Ejecutar(ByVal macro As String) As Boolean
  RuedaDesactivar                      ' sin gancho de rueda durante procesos y mensajes
  If gOcupado Then MsgBox "Hay un proceso en curso. Espera a que termine.", vbInformation: Exit Function
  gOcupado = True
  On Error GoTo fallo
  Application.Run macro
  gOcupado = False
  Ejecutar = True
  Exit Function
fallo:
  gOcupado = False
  LogP "Error en " & macro & ": " & Err.Description, "ERROR"
  MsgBox "Error en " & macro & ": " & Err.Description, vbExclamation
End Function


Private Function NewLbl(cap As String, L As Single, tp As Single, w As Single, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = cap: c.Left = L: c.Top = tp: c.Width = w: c.Height = 14: c.Font.Bold = bold
  Set NewLbl = c
End Function
Private Function NewCbo(L As Single, tp As Single, w As Single, Optional soloLista As Boolean = False) As MSForms.ComboBox
  Dim c As MSForms.ComboBox: Set c = Me.Controls.Add("Forms.ComboBox.1")
  c.Left = L: c.Top = tp: c.Width = w: c.Height = 18: c.ListRows = 16
  If soloLista Then c.Style = fmStyleDropDownList
  Set NewCbo = c
End Function
Private Function NewTxt(L As Single, tp As Single, w As Single) As MSForms.TextBox
  Dim c As MSForms.TextBox: Set c = Me.Controls.Add("Forms.TextBox.1")
  c.Left = L: c.Top = tp: c.Width = w: c.Height = 18
  Set NewTxt = c
End Function
Private Function Btn(cap As String, L As Single, tp As Single, w As Single, h As Single, col As Long, tip As String) As MSForms.CommandButton
  Dim b As MSForms.CommandButton: Set b = Me.Controls.Add("Forms.CommandButton.1")
  b.Caption = cap: b.Left = L: b.Top = tp: b.Width = w: b.Height = h
  b.BackColor = col: b.ForeColor = vbWhite: b.Font.Bold = True: b.Font.Size = 8
  b.ControlTipText = tip: b.TakeFocusOnClick = False
  Set Btn = b
End Function
Private Function PasoBtn(ByVal i As Long, cap As String, col As Long, tip As String) As MSForms.CommandButton
  Dim b As MSForms.CommandButton
  Set b = Btn(cap, 8, 24 + i * 32, 160, 28, col, tip)
  Set bArr(i) = b: mCap(i) = cap: mColBase(i) = col
  Set PasoBtn = b
End Function

' =====================================================================================
'  Datos por vista
' =====================================================================================
Private Function VS(v As Variant) As String
  If IsError(v) Then VS = "" Else VS = Trim$(CStr(v & ""))
End Function

Private Sub SetCampos(nombres As Variant, anchos As Variant, ByVal nd As Long)
  Dim i As Long
  mNf = UBound(nombres) - LBound(nombres) + 1
  ReDim mHdr(1 To mNf): ReDim mW(1 To mNf)
  For i = 1 To mNf
    mHdr(i) = nombres(LBound(nombres) + i - 1)
    If i - 1 <= UBound(anchos) - LBound(anchos) Then mW(i) = anchos(LBound(anchos) + i - 1) Else mW(i) = 80
  Next
  mNd = nd
End Sub

Private Function FindCol(ws As Worksheet, ByVal lc As Long, ByVal txt As String) As Long
  Dim j As Long
  For j = 1 To lc
    If InStr(1, TX(ws.Cells(1, j)), txt, vbTextCompare) > 0 Then FindCol = j: Exit Function
  Next
End Function

Private Function CV(v As Variant, ByVal i As Long, ByVal c As Long) As String
  If c > 0 Then CV = VS(v(i, c))
End Function

Private Function Senal(ByVal e As String, ByVal zona As String, ByVal corr As String) As String
  If Left$(zona, 4) = "ZONA" Then
    Senal = "!! ZONA"
  ElseIf Left$(zona, 7) = "POSIBLE" Then
    Senal = "! SECTOR"
  ElseIf e = "REVISAR" Then
    Senal = "? REVISAR"
  ElseIf InStr(corr, "COBERTURA CERCANA") > 0 Then
    Senal = "~ CERCANA"
  ElseIf e = "APROBADO" Then
    Senal = "+ APROBADO"
  ElseIf Len(corr) > 0 And corr <> "SIN CAMBIO" Then
    Senal = "* CORREGIDO"
  ElseIf Len(e) = 0 Then
    Senal = "- SIN VALIDAR"
  Else
    Senal = "OK"
  End If
End Function

Private Sub CargarPedidos()
  SetCampos Array("SEÑAL", "FILA", "PEDIDO", "DESTINATARIO", "PROVINCIA", "CANTON", "PARROQUIA", "ESTADO", "CORRECCION", "GESTOR", "DEST", "TRAYECTO", _
                  "GESTOR COBERTURA (Q)", "GESTOR SUGERIDO (R)", "TIPO ENTREGA", "ZONA PELIGROSA", "ACCION", "SIGLA", "ORIGINAL CLIENTE", _
                  "DIRECCION", "SUGERENCIAS", "APLICADO"), _
            Array(56, 30, 64, 136, 86, 88, 136, 58, 122, 70, 36, 92), 12
  mColFila = 2: mHoja = "DEPOT"
  Dim lr As Long, v, i As Long, n As Long, e As String, zona As String, corr As String, apl As String
  lr = wsD.Cells(wsD.Rows.Count, 3).End(xlUp).Row
  mN = 0
  If lr < 2 Then Exit Sub
  v = wsD.Range(wsD.Cells(1, 1), wsD.Cells(lr, 30)).Value
  ReDim mData(1 To lr - 1, 1 To mNf)
  For i = 2 To lr
    If Not wsD.Rows(i).Hidden And Len(VS(v(i, 3))) > 0 Then
      n = n + 1
      e = UCase$(VS(v(i, 17))): zona = VS(v(i, 26)): corr = VS(v(i, 30))
      If Len(e) = 0 Or Len(VS(v(i, 16))) = 0 Then
        apl = "-"
      ElseIf Normaliza(VS(v(i, 14))) = Normaliza(VS(v(i, 6))) And Normaliza(VS(v(i, 15))) = Normaliza(VS(v(i, 7))) _
             And Normaliza(VS(v(i, 16))) = Normaliza(VS(v(i, 8))) Then
        apl = "SI"
      Else
        apl = "NO"
      End If
      mData(n, 1) = Senal(e, zona, corr)
      mData(n, 2) = CStr(i)
      mData(n, 3) = VS(v(i, 3))
      mData(n, 4) = VS(v(i, 5))
      mData(n, 5) = VS(v(i, 14)): mData(n, 6) = VS(v(i, 15)): mData(n, 7) = VS(v(i, 16))
      If Len(mData(n, 5)) = 0 Then mData(n, 5) = VS(v(i, 6)): mData(n, 6) = VS(v(i, 7)): mData(n, 7) = VS(v(i, 8))
      mData(n, 8) = e
      mData(n, 9) = corr
      mData(n, 10) = VS(v(i, 22))
      mData(n, 11) = VS(v(i, 23))
      mData(n, 12) = VS(v(i, 24))
      mData(n, 13) = VS(v(i, 27))
      mData(n, 14) = VS(v(i, 28))
      mData(n, 15) = VS(v(i, 25))
      mData(n, 16) = zona
      mData(n, 17) = VS(v(i, 20))
      mData(n, 18) = VS(v(i, 12))
      mData(n, 19) = VS(v(i, 29))
      mData(n, 20) = VS(v(i, 2))
      mData(n, 21) = VS(v(i, 19))
      mData(n, 22) = apl
    End If
  Next
  mN = n
End Sub

Private Sub CargarCobertura()
  SetCampos Array("FILA", "PROVINCIA", "CANTON", "PARROQUIA", "SIGLA", "GESTOR (Q)", "SUGERIDO (R)", "TRAYECTO (S)", "ZONA", _
                  "TIEMPO", "DIAS FRECUENCIA", "TIPO DE TRAYECTO"), _
            Array(32, 110, 112, 150, 150, 150, 72, 120, 60), 9
  mColFila = 1: mHoja = "COBERTURA"
  Dim wc As Worksheet, lr As Long, lc As Long, v, i As Long, n As Long
  Dim cQ As Long, cR As Long, cS As Long, cZ As Long, cT As Long, cF As Long, cTT As Long
  mN = 0
  On Error Resume Next
  Set wc = ThisWorkbook.Worksheets("COBERTURA")
  On Error GoTo 0
  If wc Is Nothing Then Exit Sub
  lr = wc.Cells(wc.Rows.Count, 1).End(xlUp).Row
  lc = wc.Cells(1, wc.Columns.Count).End(xlToLeft).Column
  If lr < 2 Or lc < 4 Then Exit Sub
  cQ = FindCol(wc, lc, "GESTOR DE ENTREGAS"): cR = FindCol(wc, lc, "SUGERIDO"): cS = FindCol(wc, lc, "TRAMACO")
  cZ = FindCol(wc, lc, "ZONA"): cT = FindCol(wc, lc, "TIEMPO"): cF = FindCol(wc, lc, "FRECUENCIA"): cTT = FindCol(wc, lc, "TIPO DE TRAYECTO")
  v = wc.Range(wc.Cells(1, 1), wc.Cells(lr, lc)).Value
  ReDim mData(1 To lr - 1, 1 To mNf)
  For i = 2 To lr
    If Len(VS(v(i, 1))) > 0 Then
      n = n + 1
      mData(n, 1) = CStr(i)
      mData(n, 2) = VS(v(i, 1)): mData(n, 3) = VS(v(i, 2)): mData(n, 4) = VS(v(i, 3)): mData(n, 5) = VS(v(i, 4))
      mData(n, 6) = CV(v, i, cQ): mData(n, 7) = CV(v, i, cR): mData(n, 8) = TrayectoTexto(CV(v, i, cS))
      mData(n, 9) = CV(v, i, cZ): mData(n, 10) = CV(v, i, cT): mData(n, 11) = CV(v, i, cF): mData(n, 12) = CV(v, i, cTT)
    End If
  Next
  mN = n
End Sub

Private Sub CargarZonas()
  SetCampos Array("FILA", "PROVINCIA", "CIUDAD", "PARROQUIA", "ZONA PELIGROSA", "SECTOR", "PUNTO TMC", "DIRECCION DEL PUNTO", "VALIDACION"), _
            Array(32, 105, 110, 140, 200, 120, 200), 7
  mColFila = 1: mHoja = "ZONAS PELIGROSAS"
  Dim wz As Worksheet, lr As Long, lc As Long, v, i As Long, n As Long
  Dim cP As Long, cC As Long, cQ As Long, cZ As Long, cS As Long, cPt As Long, cD As Long, cVal As Long
  mN = 0
  On Error Resume Next
  Set wz = ThisWorkbook.Worksheets("ZONAS PELIGROSAS")
  On Error GoTo 0
  If wz Is Nothing Then Exit Sub
  lr = wz.Cells(wz.Rows.Count, 1).End(xlUp).Row
  lc = wz.Cells(1, wz.Columns.Count).End(xlToLeft).Column
  If lr < 2 Then Exit Sub
  cP = FindCol(wz, lc, "PROVINCIA"): cC = FindCol(wz, lc, "CIUDAD"): cQ = FindCol(wz, lc, "PARROQUIA")
  cZ = FindCol(wz, lc, "ZONA PELIGROSA"): cS = FindCol(wz, lc, "SECTOR"): cPt = FindCol(wz, lc, "PUNTO DE ATENCION")
  cD = FindCol(wz, lc, "DIRECC"): cVal = FindCol(wz, lc, "VALIDACION")
  v = wz.Range(wz.Cells(1, 1), wz.Cells(lr, lc)).Value
  ReDim mData(1 To lr - 1, 1 To mNf)
  For i = 2 To lr
    If Len(CV(v, i, cZ)) + Len(CV(v, i, cQ)) > 0 Then
      n = n + 1
      mData(n, 1) = CStr(i)
      mData(n, 2) = CV(v, i, cP): mData(n, 3) = CV(v, i, cC): mData(n, 4) = CV(v, i, cQ)
      mData(n, 5) = CV(v, i, cZ): mData(n, 6) = CV(v, i, cS): mData(n, 7) = CV(v, i, cPt)
      mData(n, 8) = CV(v, i, cD): mData(n, 9) = CV(v, i, cVal)
    End If
  Next
  mN = n
End Sub

Private Sub CargarDatos()
  Dim r As Long, f As Long, s As String
  Application.Cursor = xlWait
  Select Case mVista
    Case "COBERTURA": CargarCobertura
    Case "ZONAS PELIGROSAS": CargarZonas
    Case Else: CargarPedidos
  End Select
  If mN > 0 Then
    ReDim mNorm(1 To mN, 1 To mNf): ReDim mBus(1 To mN)
    For r = 1 To mN
      s = " "
      For f = 1 To mNf
        mNorm(r, f) = Normaliza(mData(r, f))
        s = s & mNorm(r, f) & " "
      Next
      mBus(r) = s
    Next
  End If
  Application.Cursor = xlDefault
  GuiaFlujo
End Sub

Private Sub PonerEncabezados()
  Dim i As Long, x As Single, w As String, e As Double
  e = mEsc: If e <= 0 Then e = 1          ' escala actual de la ventana
  x = lst.Left + 3 * e
  For i = 1 To 14
    If i <= mNd Then
      hdr(i).Caption = " " & mHdr(i): hdr(i).Left = x: hdr(i).Width = mW(i) * e - 1: hdr(i).Visible = True
      x = x + mW(i) * e
      w = w & IIf(Len(w) > 0, ";", "") & CStr(CLng(mW(i) * e))
    Else
      hdr(i).Visible = False
    End If
  Next
  lst.ColumnCount = mNd
  lst.ColumnWidths = w
End Sub

' =====================================================================================
'  Filtros, orden y lista
' =====================================================================================
Private Function PasaRapido(ByVal r As Long) As Boolean
  If mVista <> "PEDIDOS" Then PasaRapido = True: Exit Function
  Dim e As String, corr As String
  e = mData(r, 8): corr = mData(r, 9)
  Select Case cboRapido.Text
    Case "CON SEÑAL (no OK)": PasaRapido = (mData(r, 1) <> "OK")
    Case "REVISAR", "APROBADO", "OK": PasaRapido = (e = cboRapido.Text)
    Case "ZONA PELIGROSA": PasaRapido = (Len(mData(r, 16)) > 0)
    Case "COBERTURA CERCANA": PasaRapido = (InStr(corr, "COBERTURA CERCANA") > 0)
    Case "CORREGIDOS": PasaRapido = (Len(corr) > 0 And corr <> "SIN CAMBIO")
    Case "SIN APLICAR": PasaRapido = (mData(r, 22) = "NO")
    Case "SIN VALIDAR": PasaRapido = (Len(e) = 0)
    Case "DESTINO PRO (TRAMACO)": PasaRapido = (mData(r, 11) = "PRO")
    Case "DESTINO UIO": PasaRapido = (mData(r, 11) = "UIO")
    Case "DESTINO GYE": PasaRapido = (mData(r, 11) = "GYE")
    Case "DESTINO GPS": PasaRapido = (mData(r, 11) = "GPS")
    Case Else: PasaRapido = True
  End Select
End Function

Private Function Coincide(ByVal r As Long, ByVal c As Long, ByVal t As String) As Boolean
  If Len(t) = 0 Then Coincide = True: Exit Function
  If c <= 0 Then
    Coincide = (InStr(mBus(r), t) > 0)
  Else
    Coincide = (InStr(mNorm(r, c), t) > 0)
  End If
End Function

Private Function Cmp(ByVal a As Long, ByVal b As Long, ByVal f As Long) As Long
  Dim x As String, y As String
  x = mData(a, f): y = mData(b, f)
  If Len(x) > 0 And Len(y) > 0 And IsNumeric(x) And IsNumeric(y) Then
    Cmp = Sgn(CDbl(x) - CDbl(y))
  Else
    Cmp = StrComp(mNorm(a, f), mNorm(b, f), vbBinaryCompare)
  End If
  If chkDesc.Value Then Cmp = -Cmp
End Function

Private Sub QS(ByVal lo As Long, ByVal hi As Long, ByVal f As Long)
  Dim i As Long, j As Long, p As Long, t As Long
  i = lo: j = hi: p = mIdx((lo + hi) \ 2)
  Do While i <= j
    Do While Cmp(mIdx(i), p, f) < 0
      i = i + 1
    Loop
    Do While Cmp(mIdx(j), p, f) > 0
      j = j - 1
    Loop
    If i <= j Then
      t = mIdx(i): mIdx(i) = mIdx(j): mIdx(j) = t
      i = i + 1: j = j - 1
    End If
  Loop
  If lo < j Then QS lo, j, f
  If i < hi Then QS i, hi, f
End Sub

Private Sub Refrescar()
  If mCargando Then Exit Sub
  Dim r As Long, k As Long, j As Long, arr(), t1 As String, t2 As String, tb As String
  t1 = Normaliza(txtF1.Text): t2 = Normaliza(txtF2.Text): tb = Normaliza(txtBuscar.Text)
  mNIdx = 0
  lst.Clear
  If mN > 0 Then
    ReDim mIdx(1 To mN)
    For r = 1 To mN
      If PasaRapido(r) Then
        If Coincide(r, cboCampo1.ListIndex, t1) And Coincide(r, cboCampo2.ListIndex, t2) And Coincide(r, 0, tb) Then
          mNIdx = mNIdx + 1: mIdx(mNIdx) = r
        End If
      End If
    Next
    If mNIdx > 1 And cboOrden.ListIndex > 0 Then QS 1, mNIdx, cboOrden.ListIndex
    If mNIdx > 0 Then
      ReDim arr(1 To mNIdx, 1 To mNd)
      For k = 1 To mNIdx
        For j = 1 To mNd
          arr(k, j) = mData(mIdx(k), j)
        Next
      Next
      lst.List = arr
    End If
  End If
  MostrarKpi
  lblDet.ForeColor = RGB(60, 60, 60)
  lblDet.Caption = "Selecciona un registro para ver el detalle." & IIf(mVista = "PEDIDOS", "  Doble clic = abrir el pedido en el validador.", "")
End Sub

Private Sub MostrarKpi()
  Dim r As Long, cZ As Long, cRev As Long, cAp As Long, cOk As Long, cCer As Long, cCor As Long, cSinAp As Long, cPro As Long
  If mVista <> "PEDIDOS" Then
    lblKpi.Caption = mVista & ": " & mN & " registros   |   mostrando " & mNIdx
    Exit Sub
  End If
  For r = 1 To mN
    If Len(mData(r, 16)) > 0 Then cZ = cZ + 1
    Select Case mData(r, 8)
      Case "REVISAR": cRev = cRev + 1
      Case "APROBADO": cAp = cAp + 1
      Case "OK": cOk = cOk + 1
    End Select
    If InStr(mData(r, 9), "COBERTURA CERCANA") > 0 Then cCer = cCer + 1
    If Len(mData(r, 9)) > 0 And mData(r, 9) <> "SIN CAMBIO" Then cCor = cCor + 1
    If mData(r, 22) = "NO" Then cSinAp = cSinAp + 1
    If mData(r, 11) = "PRO" Then cPro = cPro + 1
  Next
  lblKpi.Caption = "Pedidos: " & mN & "   OK: " & cOk & "   REVISAR: " & cRev & "   APROBADO: " & cAp & "   Corregidos: " & cCor & _
                   "   Cob. cercana: " & cCer & "   Sin aplicar: " & cSinAp & "   Zona peligrosa: " & cZ & "   PRO: " & cPro & _
                   "   |   mostrando " & mNIdx
End Sub

Private Sub MostrarDetalle()
  If lst.ListIndex < 0 Or mNIdx = 0 Then Exit Sub
  Dim r As Long, s As String, f As Long, txtCoin As String
  r = mIdx(lst.ListIndex + 1)
  If mVista = "PEDIDOS" Then
    If Len(mData(r, 14)) > 0 Then
      If UCase$(mData(r, 14)) = UCase$(mData(r, 10)) Then txtCoin = "  (coincide con el asignado)" Else txtCoin = "  (DIFIERE del asignado)"
    End If
    s = "PEDIDO " & mData(r, 3) & "  (fila " & mData(r, 2) & ")    Señal: " & mData(r, 1) & "    Estado: " & mData(r, 8) & _
        "    Aplicado a F:H: " & mData(r, 22) & vbCrLf & _
        "Dirección: " & mData(r, 20) & vbCrLf & _
        "Cliente original: " & mData(r, 19) & "    ->    Propuesta: " & mData(r, 5) & " / " & mData(r, 6) & " / " & mData(r, 7) & _
        "    Sigla: " & mData(r, 18) & vbCrLf & _
        "Corrección: " & mData(r, 9) & "    Acción: " & mData(r, 17) & vbCrLf & _
        "Gestor asignado: " & mData(r, 10) & " (" & mData(r, 11) & ")    Gestor cobertura (Q): " & mData(r, 13) & _
        "    Sugerido (R): " & mData(r, 14) & txtCoin & "    Trayecto (S): " & mData(r, 12) & vbCrLf & _
        "Tipo de entrega: " & mData(r, 15) & IIf(Len(mData(r, 16)) > 0, vbCrLf & mData(r, 16), "") & _
        IIf(Len(mData(r, 21)) > 0, vbCrLf & "Sugerencias: " & mData(r, 21), "")
    If Left$(mData(r, 16), 4) = "ZONA" Then
      lblDet.ForeColor = RGB(192, 0, 0)
    ElseIf Len(mData(r, 16)) > 0 Or mData(r, 8) = "REVISAR" Then
      lblDet.ForeColor = RGB(191, 90, 0)
    Else
      lblDet.ForeColor = RGB(30, 30, 30)
    End If
  Else
    For f = 1 To mNf
      If Len(mData(r, f)) > 0 Then s = s & mHdr(f) & ": " & mData(r, f) & "    "
    Next
    lblDet.ForeColor = RGB(30, 30, 30)
  End If
  lblDet.Caption = s
  If chkSeguir.Value Then IrAFila False
End Sub

Private Sub IrAFila(ByVal avisar As Boolean)
  If lst.ListIndex < 0 Or mNIdx = 0 Then
    If avisar Then MsgBox "Selecciona un registro de la lista.", vbInformation
    Exit Sub
  End If
  Dim r As Long, ws As Worksheet
  r = CLng(mData(mIdx(lst.ListIndex + 1), mColFila))
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets(mHoja)
  ws.Activate
  ws.Cells(r, IIf(mHoja = "DEPOT", 3, 1)).Select
  On Error GoTo 0
  If avisar Then
    LogP "Ir a Excel: hoja " & mHoja & " fila " & r & " (volver: Complementos > Panel HYCITE)"
    RuedaDesactivar
    Me.Hide
    MostrarExcel
  End If
End Sub

' =====================================================================================
'  Guía de flujo: marca el siguiente paso recomendado
' =====================================================================================
Private Sub GuiaFlujo()
  Dim lr As Long, v, i As Long, tot As Long, nVal As Long, rev As Long, sinAp As Long, e As String, paso As Long, msg As String
  lr = wsD.Cells(wsD.Rows.Count, 3).End(xlUp).Row
  If lr >= 2 Then
    v = wsD.Range(wsD.Cells(1, 1), wsD.Cells(lr, 17)).Value
    For i = 2 To lr
      If Not wsD.Rows(i).Hidden And Len(VS(v(i, 3))) > 0 Then
        tot = tot + 1: e = UCase$(VS(v(i, 17)))
        If Len(e) > 0 Then nVal = nVal + 1
        If e = "REVISAR" Then rev = rev + 1
        If e = "OK" Or e = "APROBADO" Then
          If Normaliza(VS(v(i, 16))) <> Normaliza(VS(v(i, 8))) Or Normaliza(VS(v(i, 15))) <> Normaliza(VS(v(i, 7))) _
             Or Normaliza(VS(v(i, 14))) <> Normaliza(VS(v(i, 6))) Then sinAp = sinAp + 1
        End If
      End If
    Next
  End If
  If tot = 0 Then
    paso = 0: msg = "No hay pedidos. Actualiza los datos desde DEPOT."
  ElseIf nVal = 0 Then
    paso = 2: msg = tot & " pedidos sin validar. Valida (si hay resultados viejos, primero 1 Limpiar)."
  ElseIf nVal < tot Then
    paso = 2: msg = (tot - nVal) & " pedidos nuevos sin validar. Vuelve a validar."
  ElseIf rev > 0 Then
    paso = 3: msg = rev & " pedidos en REVISAR. Revísalos uno a uno (o 4 para aprobar en lote)."
  ElseIf sinAp > 0 Then
    paso = 5: msg = sinAp & " correcciones sin aplicar. Aplica los aprobados."
  Else
    paso = 7: msg = "Todo validado y aplicado. Envía a EGR (el archivo EGR debe estar abierto)."
  End If
  For i = 0 To 7
    bArr(i).Caption = mCap(i): bArr(i).BackColor = mColBase(i)
  Next
  bArr(paso).Caption = ">> " & mCap(paso) & " <<": bArr(paso).BackColor = RGB(255, 120, 0)
  lblGuia.Caption = "SIGUIENTE PASO: " & paso & vbCrLf & msg
End Sub

Private Sub Recargar()
  CargarDatos
  Refrescar
End Sub

' =====================================================================================
'  Eventos
' =====================================================================================
Private Sub UserForm_Activate()
  If mListo Then Recargar          ' al volver del validador se actualiza sola
End Sub

Private Sub cboVista_Change()
  If mCargando Then Exit Sub
  Dim i As Long
  mVista = cboVista.Text
  mCargando = True
  CargarDatos
  cboCampo1.Clear: cboCampo2.Clear: cboOrden.Clear
  cboCampo1.AddItem "(todas)": cboCampo2.AddItem "(todas)": cboOrden.AddItem "(sin orden)"
  For i = 1 To mNf
    cboCampo1.AddItem mHdr(i): cboCampo2.AddItem mHdr(i): cboOrden.AddItem mHdr(i)
  Next
  cboCampo1.ListIndex = 0: cboCampo2.ListIndex = 0: cboOrden.ListIndex = 0
  txtF1.Text = "": txtF2.Text = "": txtBuscar.Text = ""
  cboRapido.ListIndex = 0: cboRapido.Enabled = (mVista = "PEDIDOS")
  mCargando = False
  PonerEncabezados
  Refrescar
End Sub

Private Sub cboRapido_Change()
  Refrescar
End Sub
Private Sub txtBuscar_Change()
  Refrescar
End Sub
Private Sub cboCampo1_Change()
  Refrescar
End Sub
Private Sub txtF1_Change()
  Refrescar
End Sub
Private Sub cboCampo2_Change()
  Refrescar
End Sub
Private Sub txtF2_Change()
  Refrescar
End Sub
Private Sub cboOrden_Change()
  Refrescar
End Sub
Private Sub chkDesc_Click()
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
  If KeyCode = vbKeyReturn Then lst_DblClick Nothing
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  If lst.ListIndex < 0 Or mNIdx = 0 Then Exit Sub
  If mVista = "PEDIDOS" Then
    AbrirValidadorFila CLng(mData(mIdx(lst.ListIndex + 1), 2))
  Else
    IrAFila True
  End If
End Sub

Private Sub bLimpF_Click()
  mCargando = True
  txtF1.Text = "": txtF2.Text = "": txtBuscar.Text = ""
  cboCampo1.ListIndex = 0: cboCampo2.ListIndex = 0: cboOrden.ListIndex = 0: chkDesc.Value = False
  If cboRapido.ListCount > 0 Then cboRapido.ListIndex = 0
  mCargando = False
  Refrescar
End Sub

Private Sub b0_Click()
  Ejecutar "ActualizarDatosDepot"
  Recargar
End Sub
Private Sub b1_Click()
  Ejecutar "LimpiarValidacion"
  Recargar
End Sub
Private Sub b2_Click()
  Ejecutar "ValidarPedidos"
  Recargar
End Sub
Private Sub b3_Click()
  AbrirValidadorRevisar
End Sub
Private Sub b4_Click()
  Ejecutar "AprobarTodas"
  Recargar
End Sub
Private Sub b5_Click()
  Ejecutar "AplicarAprobados"
  Recargar
End Sub
Private Sub b6_Click()
  Ejecutar "ActualizarSiglas"
  Recargar
End Sub
Private Sub b7_Click()
  Ejecutar "ExportarADatos"
End Sub
Private Sub bCamb_Click()
  Ejecutar "ExportarCambios"
End Sub
Private Sub bIr_Click()
  IrAFila True
End Sub
Private Sub bRef_Click()
  Recargar
End Sub
Private Sub bDesb_Click()
  Desbloquear
End Sub
Private Sub bCer_Click()
  Unload Me
End Sub

Private Sub bCob_Click()
  ' contexto: el pedido seleccionado (si hay uno en la vista PEDIDOS)
  gCtxFila = 0: gCtxPedido = "": gCtxProv = "": gCtxCant = "": gCtxParr = "": gCtxDir = ""
  If mVista = "PEDIDOS" And lst.ListIndex >= 0 And mNIdx > 0 Then
    Dim r As Long: r = mIdx(lst.ListIndex + 1)
    gCtxFila = CLng(mData(r, 2)): gCtxPedido = mData(r, 3)
    gCtxProv = mData(r, 5): gCtxCant = mData(r, 6): gCtxParr = mData(r, 7): gCtxDir = mData(r, 20)
  End If
  AbrirCobertura
End Sub
Private Sub bDiag_Click()
  Ejecutar "DiagnosticarPedidos"
End Sub
Private Sub bExcel_Click()
  LogP "PANEL: oculto para ver Excel (volver: Complementos > Panel HYCITE)"
  RuedaDesactivar
  Me.Hide
  MostrarExcel
End Sub
Private Sub bMenos_Click()
  Me.Width = Me.Width * 0.9: Me.Height = Me.Height * 0.9
End Sub
Private Sub bMas_Click()
  Me.Width = Me.Width * 1.1: Me.Height = Me.Height * 1.1
End Sub
Private Sub bAjus_Click()
  AjustarAPantalla Me, BASE_W, BASE_H, 0.96
  Reescalar
End Sub
Private Sub bLogL_Click()
  txtLog.Text = ""
End Sub
