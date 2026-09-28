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
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Type MSLLHOOKSTRUCT
  pt As POINTAPI
  mouseData As Long
  flags As Long
  time As Long
  dwExtraInfo As LongPtr
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

Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)

#If Win64 Then
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal Point As LongLong) As LongPtr
#Else
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
#End If
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hWnd As LongPtr, ByVal gaFlags As Long) As LongPtr

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private mHook As LongPtr
Private mLista As Object        ' ListBox que recibe la rueda del mouse
Private mHwndForm As LongPtr    ' ventana del formulario dueño de esa lista

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

' Ajusta un formulario al porcentaje indicado de la pantalla y devuelve el zoom aplicado.
' alinear: 0 = centrado, 1 = a la derecha (deja ver Excel a la izquierda)
Public Function AjustarAPantalla(frm As Object, ByVal baseW As Single, ByVal baseH As Single, ByVal pct As Double, Optional ByVal alinear As Long = 0) As Double
  Dim L As Single, T As Single, W As Single, H As Single, f As Double
  If Not AreaTrabajo(L, T, W, H) Then
    L = 0: T = 0: W = Application.UsableWidth: H = Application.UsableHeight + 60
  End If
  ' con escalado de Windows (125 %, 150 %) el área puede venir sobrestimada: se limita con la ventana de Excel
  On Error Resume Next
  If Application.WindowState = xlMaximized Then
    If Application.Width > 400 And Application.Width < W Then W = Application.Width: L = Application.Left
    If Application.Height > 300 And Application.Height < H Then H = Application.Height: T = Application.Top
  End If
  On Error GoTo 0
  f = (W * pct) / baseW
  If (H * pct) / baseH < f Then f = (H * pct) / baseH
  If f < 0.5 Then f = 0.5
  If f > 2.5 Then f = 2.5
  frm.StartUpPosition = 0
  frm.Zoom = f * 100
  frm.Width = baseW * f: frm.Height = baseH * f
  If alinear = 1 Then
    frm.Left = L + W - frm.Width - 4: frm.Top = T + (H - frm.Height) / 2
  Else
    frm.Left = L + (W - frm.Width) / 2: frm.Top = T + (H - frm.Height) / 2
  End If
  If frm.Top < T Then frm.Top = T
  AjustarAPantalla = f
End Function

' Mostrar Excel (el panel se oculta; se vuelve con Complementos > Panel HYCITE)
Public Sub MostrarExcel()
  On Error Resume Next
  If Application.WindowState = xlMinimized Then Application.WindowState = xlMaximized
  AppActivate Application.Caption
End Sub

' =====================================================================================
'  RUEDA DEL MOUSE en listas (las ListBox de VBA no la manejan solas).
'  El gancho se activa solo mientras el mouse está sobre una lista y se quita al salir
'  o al cerrar el formulario.
' =====================================================================================
Public Sub RuedaActivar(lb As Object, ByVal tituloForm As String)
  Set mLista = lb
  mHwndForm = FindWindow("ThunderDFrame", tituloForm)
  If mHook = 0 Then
    On Error Resume Next
    mHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf ProcRueda, Application.HinstancePtr, 0)
    On Error GoTo 0
  End If
End Sub

Public Sub RuedaDesactivar()
  If mHook <> 0 Then UnhookWindowsHookEx mHook
  mHook = 0
  Set mLista = Nothing
  mHwndForm = 0
End Sub

Private Function ProcRueda(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
  On Error GoTo pasar
  If nCode = 0 And wParam = WM_MOUSEWHEEL And Not mLista Is Nothing Then
    Dim ms As MSLLHOOKSTRUCT, delta As Long, t As Long
    CopyMemory ms, lParam, LenB(ms)
    If mHwndForm <> 0 Then
      If CursorFuera(ms.pt.X, ms.pt.Y) Then GoTo pasar      ' el mouse ya no está sobre el formulario
    End If
    delta = ms.mouseData \ &H10000                    ' parte alta con signo: +120 arriba / -120 abajo
    If mLista.ListCount > 0 Then
      t = mLista.TopIndex - Sgn(delta) * 3
      If t < 0 Then t = 0
      If t > mLista.ListCount - 1 Then t = mLista.ListCount - 1
      mLista.TopIndex = t
    End If
    ProcRueda = 1                                     ' la rueda ya se usó en la lista
    Exit Function
  End If
pasar:
  ProcRueda = CallNextHookEx(mHook, nCode, wParam, lParam)
End Function

Private Function CursorFuera(ByVal X As Long, ByVal Y As Long) As Boolean
  Dim h As LongPtr
  On Error GoTo fin
#If Win64 Then
  Dim pt As LongLong
  pt = (CLngLng(Y) * &H100000000^) Or (CLngLng(X) And &HFFFFFFFF^)
  h = WindowFromPoint(pt)
#Else
  h = WindowFromPoint(X, Y)
#End If
  If h <> 0 Then h = GetAncestor(h, 2)          ' GA_ROOT
  CursorFuera = (h <> mHwndForm)
fin:
End Function
