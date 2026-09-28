Option Explicit
' =====================================================================================
'  frmPanel - Panel central de validación (NUEVO).
'  Crear: Insertar > UserForm, nombre (Name) = frmPanel, pegar este código. No hay que dibujar
'  controles: se crean al abrir. Muestra la hoja DEPOT como matriz con filtros, indicadores
'  y los botones del flujo en orden. Doble clic en un pedido = abrirlo en el validador.
' =====================================================================================
Private wsD As Worksheet
Private lblKpi As MSForms.Label
Private WithEvents cboFiltro As MSForms.ComboBox
Private WithEvents lst As MSForms.ListBox
Private WithEvents b1 As MSForms.CommandButton    ' 1 Limpiar
Private WithEvents b2 As MSForms.CommandButton    ' 2 Validar
Private WithEvents b3 As MSForms.CommandButton    ' 3 Revisar pendientes
Private WithEvents b4 As MSForms.CommandButton    ' 4 Aprobar lote
Private WithEvents b5 As MSForms.CommandButton    ' 5 Aplicar aprobados
Private WithEvents b6 As MSForms.CommandButton    ' 6 Actualizar siglas
Private WithEvents b7 As MSForms.CommandButton    ' 7 Enviar a EGR
Private WithEvents bCamb As MSForms.CommandButton ' Exportar CAMBIOS
Private WithEvents bDesb As MSForms.CommandButton ' Desbloquear
Private WithEvents bIr As MSForms.CommandButton   ' Ir a la fila
Private WithEvents bRef As MSForms.CommandButton  ' Actualizar lista
Private WithEvents bCer As MSForms.CommandButton  ' Cerrar
Private Const C_REF As Long = 3, C_DEST As Long = 5, C_NN As Long = 14, C_EXT As Long = 22

Private Sub UserForm_Initialize()
  Set wsD = ThisWorkbook.Worksheets("DEPOT")
  Me.Caption = "Panel de validación HYCITE"
  Me.Width = 880: Me.Height = 500
  Dim t As MSForms.Label
  Set t = Me.Controls.Add("Forms.Label.1")
  t.Caption = "Filtro:": t.Left = 10: t.Top = 12: t.Width = 40
  Set cboFiltro = Me.Controls.Add("Forms.ComboBox.1")
  cboFiltro.Left = 50: cboFiltro.Top = 8: cboFiltro.Width = 170: cboFiltro.Style = fmStyleDropDownList
  Dim f
  For Each f In Array("TODOS", "REVISAR", "OK", "APROBADO", "ZONA PELIGROSA", "DESTINO PRO (TRAMACO)", "DESTINO UIO", "DESTINO GYE", "DESTINO GPS")
    cboFiltro.AddItem f
  Next
  Set lblKpi = Me.Controls.Add("Forms.Label.1")
  lblKpi.Left = 232: lblKpi.Top = 6: lblKpi.Width = 630: lblKpi.Height = 30: lblKpi.WordWrap = True: lblKpi.Font.Bold = True
  Set lst = Me.Controls.Add("Forms.ListBox.1")
  lst.Left = 10: lst.Top = 40: lst.Width = 852: lst.Height = 350
  lst.ColumnCount = 10: lst.ColumnHeads = False
  lst.ColumnWidths = "34;66;120;90;90;130;58;86;40;210"
  lst.Font.Size = 8
  Set b1 = Btn("1 Limpiar", 10, 400, RGB(150, 150, 150))
  Set b2 = Btn("2 Validar", 116, 400, RGB(47, 117, 181))
  Set b3 = Btn("3 Revisar pendientes", 222, 400, RGB(237, 125, 49))
  Set b4 = Btn("4 Aprobar lote", 328, 400, RGB(191, 143, 0))
  Set b5 = Btn("5 Aplicar aprobados", 434, 400, RGB(112, 173, 71))
  Set b6 = Btn("6 Actualizar siglas", 540, 400, RGB(47, 117, 181))
  Set b7 = Btn("7 Enviar a EGR", 646, 400, RGB(0, 128, 96))
  Set bCamb = Btn("Exportar CAMBIOS", 10, 434, RGB(84, 130, 53))
  Set bIr = Btn("Ir a la fila", 116, 434, RGB(120, 120, 120))
  Set bRef = Btn("Actualizar lista", 222, 434, RGB(120, 120, 120))
  Set bDesb = Btn("Desbloquear", 328, 434, RGB(120, 120, 120))
  Set bCer = Btn("Cerrar", 752, 434, RGB(192, 80, 77))
  cboFiltro.ListIndex = 0   ' dispara Refrescar
End Sub

Private Function Btn(cap As String, L As Single, tp As Single, col As Long) As MSForms.CommandButton
  Dim b As MSForms.CommandButton: Set b = Me.Controls.Add("Forms.CommandButton.1")
  b.Caption = cap: b.Left = L: b.Top = tp: b.Width = 100: b.Height = 28
  b.BackColor = col: b.ForeColor = vbWhite: b.Font.Bold = True: b.Font.Size = 8
  Set Btn = b
End Function

Private Function PasaFiltro(ByVal f As String, ByVal e As String, ByVal dest As String, ByVal zona As String) As Boolean
  Select Case f
    Case "REVISAR", "OK", "APROBADO": PasaFiltro = (e = f)
    Case "ZONA PELIGROSA": PasaFiltro = (Len(zona) > 0)
    Case "DESTINO PRO (TRAMACO)": PasaFiltro = (dest = "PRO")
    Case "DESTINO UIO": PasaFiltro = (dest = "UIO")
    Case "DESTINO GYE": PasaFiltro = (dest = "GYE")
    Case "DESTINO GPS": PasaFiltro = (dest = "GPS")
    Case Else: PasaFiltro = True
  End Select
End Function

Private Sub Refrescar()
  Dim lr As Long, i As Long, n As Long, k As Long, j As Long, e As String, dest As String, zona As String
  Dim nOk As Long, nRev As Long, nAp As Long, nZona As Long, nPro As Long, nTot As Long
  Dim arr(), arr2()
  lr = wsD.Cells(wsD.Rows.Count, C_REF).End(xlUp).Row
  lst.Clear
  If lr < 2 Then lblKpi.Caption = "Sin pedidos en DEPOT.": Exit Sub
  ReDim arr(1 To lr - 1, 1 To 10)
  For i = 2 To lr
    If Not wsD.Rows(i).Hidden And Len(TX(wsD.Cells(i, C_REF))) > 0 Then
      e = UCase$(Trim$(TX(wsD.Cells(i, C_NN + 3))))
      dest = TX(wsD.Cells(i, C_EXT + 1)): zona = TX(wsD.Cells(i, C_EXT + 4))
      nTot = nTot + 1
      If e = "OK" Then nOk = nOk + 1
      If e = "REVISAR" Then nRev = nRev + 1
      If e = "APROBADO" Then nAp = nAp + 1
      If Len(zona) > 0 Then nZona = nZona + 1
      If dest = "PRO" Then nPro = nPro + 1
      If PasaFiltro(cboFiltro.Text, e, dest, zona) Then
        n = n + 1
        arr(n, 1) = i
        arr(n, 2) = TX(wsD.Cells(i, C_REF))
        arr(n, 3) = TX(wsD.Cells(i, C_DEST))
        arr(n, 4) = TX(wsD.Cells(i, C_NN))
        arr(n, 5) = TX(wsD.Cells(i, C_NN + 1))
        arr(n, 6) = TX(wsD.Cells(i, C_NN + 2))
        arr(n, 7) = e
        arr(n, 8) = TX(wsD.Cells(i, C_EXT))
        arr(n, 9) = dest
        arr(n, 10) = IIf(Len(zona) > 0, "ZONA PELIGROSA | ", "") & TX(wsD.Cells(i, C_NN + 6))
      End If
    End If
  Next
  lblKpi.Caption = "Pedidos: " & nTot & "   OK: " & nOk & "   REVISAR: " & nRev & "   APROBADO: " & nAp & _
                   "   Zona peligrosa: " & nZona & "   PRO (TRAMACO): " & nPro & vbCrLf & _
                   "Columnas: fila | pedido | destinatario | provincia | cantón | parroquia | estado | gestor | destino | acción"
  If n = 0 Then Exit Sub
  ReDim arr2(1 To n, 1 To 10)
  For k = 1 To n
    For j = 1 To 10: arr2(k, j) = arr(k, j): Next
  Next
  lst.List = arr2
End Sub

Private Function FilaSel() As Long
  If lst.ListIndex < 0 Then MsgBox "Selecciona un pedido de la lista.", vbInformation: Exit Function
  FilaSel = CLng(lst.List(lst.ListIndex, 0))
End Function

Private Sub cboFiltro_Change()
  Refrescar
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim r As Long: r = FilaSel(): If r = 0 Then Exit Sub
  AbrirValidadorFila r
End Sub

Private Sub b1_Click()
  LimpiarValidacion
  Refrescar
End Sub
Private Sub b2_Click()
  ValidarPedidos
  Refrescar
End Sub
Private Sub b3_Click()
  AbrirValidadorRevisar
End Sub
Private Sub b4_Click()
  AprobarTodas
  Refrescar
End Sub
Private Sub b5_Click()
  AplicarAprobados
  Refrescar
End Sub
Private Sub b6_Click()
  ActualizarSiglas
  Refrescar
End Sub
Private Sub b7_Click()
  ExportarADatos
End Sub
Private Sub bCamb_Click()
  ExportarCambios
End Sub
Private Sub bDesb_Click()
  Desbloquear
End Sub
Private Sub bRef_Click()
  Refrescar
End Sub
Private Sub bCer_Click()
  Unload Me
End Sub

Private Sub bIr_Click()
  Dim r As Long: r = FilaSel(): If r = 0 Then Exit Sub
  wsD.Activate
  wsD.Cells(r, C_REF).Select
End Sub
