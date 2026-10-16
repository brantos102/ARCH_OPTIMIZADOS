Option Explicit
' ThisWorkbook - reemplaza TODO el código de ThisWorkbook.
' La barra "Validacion HYCITE" solo se muestra mientras este libro está activo.

Private Sub Workbook_Open()
  CrearMenuHycite
End Sub

Private Sub Workbook_Activate()
  CrearMenuHycite
End Sub

Private Sub Workbook_Deactivate()
  On Error Resume Next
  Application.CommandBars("Validacion HYCITE").Delete
End Sub
