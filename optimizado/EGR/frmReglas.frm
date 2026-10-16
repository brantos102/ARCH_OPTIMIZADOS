Option Explicit
' =====================================================================================
'  frmReglas - Editor de la hoja REGLAS_DESTINO sin tocar la hoja ni el código.
'  Insertar > UserForm, nombre frmReglas, pegar este código.
'  Se abre desde el panel (Editar reglas) o desde Complementos > Reglas de destino.
'   - lista de reglas en orden de prioridad
'   - alta, cambio, eliminación y activar/desactivar
'   - probador: escribe un pedido y mira qué regla gana
'  Rendimiento: lee y escribe la hoja en bloque; la lista se vuelca de una sola vez.
' =====================================================================================
Private Const BASE_W As Single = 900
Private Const BASE_H As Single = 560
Private Const TITULO As String = "Reglas de destino HYCITE"

Private mLay As Variant, mCW As Single, mCH As Single, mEsc As Double, mEscalando As Boolean
Private mCerrando As Boolean, mCarga As Boolean, mSucio As Boolean
Private mD() As String, mN As Long          ' 1 activa, 2 prioridad, 3 prov, 4 cant, 5 parr, 6 texto, 7 excepto, 8 destino, 9 motivo
Private lblAyuda As MSForms.Label, lblRes As MSForms.Label
Private WithEvents lst As MSForms.ListBox
Private WithEvents chkAct As MSForms.CheckBox
Private txtPri As MSForms.TextBox, txtProv As MSForms.TextBox, txtCant As MSForms.TextBox
Private txtParr As MSForms.TextBox, txtTxt As MSForms.TextBox, txtExc As MSForms.TextBox, txtMot As MSForms.TextBox
Private cboDest As MSForms.ComboBox
Private WithEvents bNueva As MSForms.CommandButton
Private WithEvents bGuardar As MSForms.CommandButton
Private WithEvents bBorrar As MSForms.CommandButton
Private WithEvents bHoja As MSForms.CommandButton
Private WithEvents bCerrar As MSForms.CommandButton
Private WithEvents bProbar As MSForms.CommandButton
Private tProv As MSForms.TextBox, tCant As MSForms.TextBox, tParr As MSForms.TextBox, tDir As MSForms.TextBox

Private Sub UserForm_Initialize()
  Dim t As MSForms.Label, f
  Me.Caption = TITULO
  Me.Zoom = 100: Me.Width = BASE_W: Me.Height = BASE_H

  Set t = NL("REGLAS ACTIVAS, EN ORDEN DE PRIORIDAD", 10, 8, 420, True): t.ForeColor = RGB(48, 84, 150)
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = 10: lst.Top = 24: lst.Width = 448: lst.Height = 330: lst.Font.Size = 8
  lst.ColumnCount = 6: lst.ColumnWidths = "30;34;74;80;44;180"
  Set lblAyuda = NL("", 10, 358, 448)
  lblAyuda.Height = 58: lblAyuda.WordWrap = True: lblAyuda.BackColor = RGB(255, 242, 204)
  lblAyuda.BorderStyle = fmBorderStyleSingle: lblAyuda.ForeColor = RGB(128, 64, 0)
  lblAyuda.Caption = "Gana la regla ACTIVA de menor prioridad que coincide con el pedido. " & _
                     "Vacío = cualquiera. Varias opciones separadas por punto y coma. " & _
                     "En TEXTO, KM>=15 significa del kilómetro 15 en adelante."

  ' ----- ficha de la regla -----
  Set t = NL("REGLA SELECCIONADA", 470, 8, 300, True): t.ForeColor = RGB(48, 84, 150)
  Set chkAct = Me.Controls.Add("Forms.CheckBox.1")
  chkAct.Caption = "Activa": chkAct.Left = 470: chkAct.Top = 26: chkAct.Width = 60: chkAct.Height = 18
  NL "Prioridad", 540, 28, 48
  Set txtPri = NT(592, 25, 50)
  txtPri.ControlTipText = "Número menor = se evalúa antes. Las reglas base son 900, 910 y 999."
  NL "Destino", 652, 28, 42
  Set cboDest = Me.Controls.Add("Forms.ComboBox.1")
  cboDest.Left = 696: cboDest.Top = 25: cboDest.Width = 64: cboDest.Height = 18: cboDest.Style = fmStyleDropDownList
  For Each f In Array("PRO", "GYE", "UIO", "GPS"): cboDest.AddItem f: Next

  NL "Provincia", 470, 54, 54: Set txtProv = NT(528, 51, 232)
  NL "Cantón", 470, 78, 54: Set txtCant = NT(528, 75, 232)
  txtCant.ControlTipText = "Uno o varios, separados por punto y coma: DURAN;MILAGRO;PLAYAS"
  NL "Parroquia", 470, 102, 54: Set txtParr = NT(528, 99, 232)
  NL "Texto", 470, 126, 54: Set txtTxt = NT(528, 123, 232)
  txtTxt.ControlTipText = "Palabras buscadas en la dirección y en la parroquia: COOP;GUASMO;KM>=15"
  NL "Excepto", 470, 150, 54: Set txtExc = NT(528, 147, 232)
  txtExc.ControlTipText = "Parroquias que quedan fuera de esta regla."
  NL "Motivo", 470, 174, 54: Set txtMot = NT(528, 171, 232)
  txtMot.ControlTipText = "Texto que verá el operador cuando se proponga este destino."

  Set bNueva = NB("Nueva regla", 470, 200, 92, 24, RGB(47, 117, 181))
  Set bGuardar = NB("Guardar", 568, 200, 92, 24, RGB(0, 128, 96))
  Set bBorrar = NB("Eliminar", 666, 200, 94, 24, RGB(192, 80, 77))

  ' ----- probador -----
  Set t = NL("PROBAR CON UN PEDIDO", 470, 238, 300, True): t.ForeColor = RGB(48, 84, 150)
  NL "Provincia", 470, 260, 54: Set tProv = NT(528, 257, 232)
  NL "Cantón", 470, 284, 54: Set tCant = NT(528, 281, 232)
  NL "Parroquia", 470, 308, 54: Set tParr = NT(528, 305, 232)
  NL "Dirección", 470, 332, 54: Set tDir = NT(528, 329, 232)
  Set bProbar = NB("Probar", 470, 356, 92, 24, RGB(112, 48, 160))
  Set lblRes = NL("", 568, 354, 192)
  lblRes.Height = 60: lblRes.WordWrap = True: lblRes.BorderStyle = fmBorderStyleSingle: lblRes.BackColor = RGB(248, 248, 248)

  Set bHoja = NB("Ver la hoja REGLAS_DESTINO", 470, 426, 160, 24, RGB(120, 120, 120))
  Set bCerrar = NB("Cerrar", 636, 426, 124, 24, RGB(120, 120, 120))

  Cargar
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
  mEscalando = True: mEsc = f
  EscalarLayout Me, mLay, f, mCW, mCH
  mEscalando = False
End Sub

Private Function NL(cap As String, x As Single, y As Single, w As Single, Optional bold As Boolean = False) As MSForms.Label
  Dim c As MSForms.Label: Set c = Me.Controls.Add("Forms.Label.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = 14: c.Font.Bold = bold: c.Font.Size = 8
  Set NL = c
End Function
Private Function NT(x As Single, y As Single, w As Single) As MSForms.TextBox
  Dim c As MSForms.TextBox: Set c = Me.Controls.Add("Forms.TextBox.1")
  c.Left = x: c.Top = y: c.Width = w: c.Height = 18: c.Font.Size = 9
  Set NT = c
End Function
Private Function NB(cap As String, x As Single, y As Single, w As Single, h As Single, col As Long) As MSForms.CommandButton
  Dim c As MSForms.CommandButton: Set c = Me.Controls.Add("Forms.CommandButton.1")
  c.Caption = cap: c.Left = x: c.Top = y: c.Width = w: c.Height = h
  c.BackColor = col: c.ForeColor = vbWhite: c.Font.Bold = True: c.Font.Size = 8: c.TakeFocusOnClick = False
  Set NB = c
End Function

' ---------- datos ----------
Private Sub Cargar()
  Dim ws As Worksheet, lr As Long, v, i As Long, j As Long
  CrearHojaReglas
  Set ws = ThisWorkbook.Worksheets(HREG)
  lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
  mN = 0
  ReDim mD(1 To IIf(lr > 1, lr + 30, 40), 1 To 9)
  If lr >= 2 Then
    v = ws.Range("A2:I" & lr).Value
    For i = 1 To UBound(v, 1)
      If Len(TXE(v(i, 8))) > 0 Then
        mN = mN + 1
        For j = 1 To 9: mD(mN, j) = TXE(v(i, j)): Next
      End If
    Next
  End If
  Ordenar
  Pintar
End Sub

' VBA solo deja redimensionar la última dimensión conservando datos, así que se copia a un arreglo mayor
Private Sub Crecer()
  Dim nuevo() As String, i As Long, j As Long
  ReDim nuevo(1 To UBound(mD, 1) + 30, 1 To 9)
  For i = 1 To mN
    For j = 1 To 9: nuevo(i, j) = mD(i, j): Next
  Next
  mD = nuevo
End Sub

Private Sub Ordenar()
  Dim i As Long, j As Long, k As Long, tmp As String
  For i = 2 To mN
    For j = i To 2 Step -1
      If Val(mD(j, 2)) < Val(mD(j - 1, 2)) Then
        For k = 1 To 9: tmp = mD(j, k): mD(j, k) = mD(j - 1, k): mD(j - 1, k) = tmp: Next
      Else
        Exit For
      End If
    Next
  Next
End Sub

Private Sub Pintar()
  Dim arr() As String, i As Long, sel As Long
  sel = lst.ListIndex
  mCarga = True
  lst.Clear
  If mN > 0 Then
    ReDim arr(0 To mN - 1, 0 To 5)
    For i = 1 To mN
      arr(i - 1, 0) = IIf(UCase$(mD(i, 1)) = "NO", "·", "SI")
      arr(i - 1, 1) = mD(i, 2)
      arr(i - 1, 2) = Left$(mD(i, 3), 22)
      arr(i - 1, 3) = Left$(mD(i, 4), 26)
      arr(i - 1, 4) = mD(i, 8)
      arr(i - 1, 5) = Left$(mD(i, 9), 60)
    Next
    lst.List = arr
    If sel >= 0 And sel < mN Then lst.ListIndex = sel Else lst.ListIndex = 0
  End If
  mCarga = False
  If lst.ListIndex >= 0 Then VerFicha lst.ListIndex + 1
End Sub

Private Sub VerFicha(ByVal k As Long)
  If k < 1 Or k > mN Then Exit Sub
  mCarga = True
  chkAct.Value = (UCase$(mD(k, 1)) <> "NO")
  txtPri.Text = mD(k, 2): txtProv.Text = mD(k, 3): txtCant.Text = mD(k, 4)
  txtParr.Text = mD(k, 5): txtTxt.Text = mD(k, 6): txtExc.Text = mD(k, 7)
  cboDest.Text = mD(k, 8): txtMot.Text = mD(k, 9)
  mCarga = False
End Sub

Private Sub lst_Change()
  If mCarga Or mCerrando Then Exit Sub
  If lst.ListIndex >= 0 Then VerFicha lst.ListIndex + 1
End Sub

Private Sub bNueva_Click()
  mCarga = True
  chkAct.Value = True: txtPri.Text = "50": cboDest.Text = "PRO"
  txtProv.Text = "": txtCant.Text = "": txtParr.Text = "": txtTxt.Text = "": txtExc.Text = ""
  txtMot.Text = "Nueva regla"
  mCarga = False
  lst.ListIndex = -1
  txtProv.SetFocus
  MsgBox "Completa la ficha de la derecha y pulsa Guardar. La regla se agrega al final y se ordena por prioridad.", vbInformation
End Sub

Private Sub bGuardar_Click()
  Dim k As Long
  If Len(Trim$(cboDest.Text)) = 0 Then MsgBox "Elige el destino: PRO, GYE, UIO o GPS.", vbExclamation: Exit Sub
  If Val(txtPri.Text) <= 0 Then MsgBox "La prioridad debe ser un número mayor que cero.", vbExclamation: Exit Sub
  If lst.ListIndex >= 0 Then
    k = lst.ListIndex + 1
  Else
    If mN + 1 > UBound(mD, 1) Then Crecer
    mN = mN + 1
    k = mN
  End If
  mD(k, 1) = IIf(chkAct.Value, "SI", "NO")
  mD(k, 2) = CStr(CLng(Val(txtPri.Text)))
  mD(k, 3) = Trim$(txtProv.Text): mD(k, 4) = Trim$(txtCant.Text): mD(k, 5) = Trim$(txtParr.Text)
  mD(k, 6) = Trim$(txtTxt.Text): mD(k, 7) = Trim$(txtExc.Text)
  mD(k, 8) = UCase$(Trim$(cboDest.Text)): mD(k, 9) = Trim$(txtMot.Text)
  Ordenar
  Escribir
  Pintar
  LogE "REGLAS: guardada la regla " & mD(k, 2) & " -> " & mD(k, 8)
End Sub

Private Sub bBorrar_Click()
  Dim k As Long, i As Long, j As Long
  If lst.ListIndex < 0 Then MsgBox "Elige una regla de la lista.", vbInformation: Exit Sub
  k = lst.ListIndex + 1
  If MsgBox("¿Eliminar la regla " & mD(k, 2) & " (" & mD(k, 8) & ")?" & vbCrLf & mD(k, 9), vbYesNo + vbExclamation) <> vbYes Then Exit Sub
  LogE "REGLAS: eliminada la regla " & mD(k, 2) & " -> " & mD(k, 8) & " (" & mD(k, 9) & ")", "AVISO"
  For i = k To mN - 1
    For j = 1 To 9: mD(i, j) = mD(i + 1, j): Next
  Next
  mN = mN - 1
  Escribir
  Pintar
End Sub

Private Sub Escribir()
  Dim ws As Worksheet, i As Long, j As Long, v() As Variant, lr As Long
  Set ws = ThisWorkbook.Worksheets(HREG)
  Application.ScreenUpdating = False
  lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
  If lr >= 2 Then ws.Range("A2:I" & lr).ClearContents
  If mN > 0 Then
    ReDim v(1 To mN, 1 To 9)
    For i = 1 To mN
      For j = 1 To 9: v(i, j) = mD(i, j): Next
    Next
    ws.Range(ws.Cells(2, 1), ws.Cells(mN + 1, 9)).Value = v
  End If
  Application.ScreenUpdating = True
  mSucio = True
End Sub

' ---------- probador ----------
Private Sub bProbar_Click()
  Dim reglas, dest As String, mot As String
  reglas = CargarReglas()
  If Not IsArray(reglas) Then lblRes.Caption = "No hay reglas activas.": Exit Sub
  dest = DestinoPorReglas(reglas, tProv.Text, tCant.Text, tParr.Text, tDir.Text, mot)
  If Len(dest) = 0 Then
    lblRes.Caption = "Ninguna regla coincide."
    lblRes.ForeColor = RGB(192, 80, 77)
  Else
    lblRes.Caption = "Destino: " & dest & vbCrLf & mot
    lblRes.ForeColor = RGB(0, 110, 0)
  End If
End Sub

Private Sub bHoja_Click()
  mCerrando = True
  Unload Me
  MostrarExcel
  With ThisWorkbook.Worksheets(HREG)
    .Visible = xlSheetVisible
    .Activate
  End With
End Sub

Private Sub bCerrar_Click()
  mCerrando = True
  Unload Me
  If mSucio Then LogE "REGLAS: cambios guardados. Vuelve a 'Ver cambios sugeridos' para aplicarlas a los pedidos de hoy."
End Sub
