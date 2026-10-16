Option Explicit
' ThisWorkbook de Formato EGR_FL_HYCITE: crea la barra "Despacho EGR HYCITE" (pestaña Complementos)
Private Sub Workbook_Open()
  CrearMenuEGR
End Sub

Private Sub Workbook_Activate()
  CrearMenuEGR
End Sub

Private Sub Workbook_Deactivate()
  BorrarMenuEGR
End Sub
