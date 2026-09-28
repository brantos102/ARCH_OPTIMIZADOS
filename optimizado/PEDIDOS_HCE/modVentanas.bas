Option Explicit
' =====================================================================================
'  modVentanas (NUEVO) - Insertar > Módulo, nombre modVentanas, pegar este código.
'  Hace que los formularios se puedan agrandar/achicar con el mouse (borde y maximizar)
'  y calcula el área útil de la pantalla para abrir el panel a pantalla completa.
'  Requiere Office 2010 o superior (VBA7), 32 o 64 bits.
' =====================================================================================
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SPI_GETWORKAREA As Long = 48
Private Const LOGPIXELSX As Long = 88

' Agrega borde redimensionable y botón maximizar al formulario (se busca por su título)
Public Sub HacerRedimensionable(ByVal titulo As String)
  Dim h As LongPtr, st As LongPtr
  On Error GoTo salir
  h = FindWindow("ThunderDFrame", titulo)
  If h = 0 Then h = FindWindow("ThunderXFrame", titulo)
  If h = 0 Then Exit Sub
  st = GetWindowLongPtr(h, GWL_STYLE)
  st = st Or WS_THICKFRAME Or WS_MAXIMIZEBOX
  SetWindowLongPtr h, GWL_STYLE, st
  DrawMenuBar h
salir:
End Sub

' Área útil de la pantalla (sin barra de tareas) en puntos
Public Function AreaTrabajo(ByRef L As Single, ByRef T As Single, ByRef W As Single, ByRef H As Single) As Boolean
  Dim r As RECT, hdc As LongPtr, dpi As Long
  On Error GoTo fallo
  If SystemParametersInfo(SPI_GETWORKAREA, 0, r, 0) = 0 Then GoTo fallo
  hdc = GetDC(0): dpi = GetDeviceCaps(hdc, LOGPIXELSX): ReleaseDC 0, hdc
  If dpi <= 0 Then dpi = 96
  L = r.Left * 72 / dpi: T = r.Top * 72 / dpi
  W = (r.Right - r.Left) * 72 / dpi: H = (r.Bottom - r.Top) * 72 / dpi
  AreaTrabajo = (W > 200 And H > 200)
  Exit Function
fallo:
  AreaTrabajo = False
End Function

' Ajusta un formulario al porcentaje indicado de la pantalla, centrado, y devuelve el zoom aplicado
Public Function AjustarAPantalla(frm As Object, ByVal baseW As Single, ByVal baseH As Single, ByVal pct As Double) As Double
  Dim L As Single, T As Single, W As Single, H As Single, f As Double
  If Not AreaTrabajo(L, T, W, H) Then
    L = 0: T = 0: W = Application.UsableWidth: H = Application.UsableHeight + 60
  End If
  f = (W * pct) / baseW
  If (H * pct) / baseH < f Then f = (H * pct) / baseH
  If f < 0.5 Then f = 0.5
  If f > 2.5 Then f = 2.5
  frm.StartUpPosition = 0
  frm.Zoom = f * 100
  frm.Width = baseW * f: frm.Height = baseH * f
  frm.Left = L + (W - frm.Width) / 2: frm.Top = T + (H - frm.Height) / 2
  AjustarAPantalla = f
End Function

' Mostrar Excel (el panel se oculta; se vuelve con Complementos > Panel HYCITE)
Public Sub MostrarExcel()
  On Error Resume Next
  If Application.WindowState = xlMinimized Then Application.WindowState = xlMaximized
  AppActivate Application.Caption
End Sub
