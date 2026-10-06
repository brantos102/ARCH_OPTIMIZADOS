' ===========================================================================================
'  Modulo1 - VERSION PARA PEGAR en el editor de VBA (sin lineas Attribute)
'
'  Use este archivo si copia y pega el codigo dentro del modulo Modulo1 ya existente.
'  Las lineas 'Attribute' solo son validas dentro de un .bas IMPORTADO: pegadas en el
'  editor producen el error 'Atributo no valido en Sub o Function'.
'
'  Al no llevar el atributo, el atajo de teclado se vuelve a asignar a mano en:
'     Vista > Macros > VistaPreviaEImpresion > Opciones...  (tecla P)
'
'  Si prefiere conservar el atajo automaticamente, use src/Modulo1.bas y la opcion
'  Archivo > Importar archivo... del editor de VBA.
' ===========================================================================================

Option Explicit

'==============================================================================================
' Módulo1  -  Entradas de impresión de etiquetas
'----------------------------------------------------------------------------------------------
' VistaPreviaEImpresion ........... mismo nombre y mismo atajo de siempre. Ahora llama al
'                                   motor de lotes (modEtiquetas): respeta el filtro de
'                                   "Consolidado", permite elegir copias por etiqueta y envía
'                                   lotes de 100 etiquetas en un solo trabajo de impresión.
' VistaPreviaEImpresion_Clasica ... lógica original (una etiqueta por trabajo de impresión),
'                                   conservada como respaldo y ya corregida:
'                                     - omite las filas ocultas por el autofiltro
'                                     - imprime también la primera etiqueta (antes se perdía)
'                                     - contador Long (antes Integer, fallaba sobre 32767)
' ValidarDato ..................... corregida: ya no falla cuando la celda tiene un error
'                                   (#N/A, #REF!, ...). VBA evalúa los dos lados de "Or",
'                                   por lo que la comparación valor = "" rompía la función.
'==============================================================================================

Sub VistaPreviaEImpresion()
    ImprimirEtiquetasZebra
End Sub


Sub VistaPreviaEImpresion_Clasica()

    Dim wsETQ As Worksheet, hojaOrigen As Worksheet
    Dim celda As Range
    Dim filaSeleccionada As Long
    Dim filasProcesadas As Long
    Dim valorA6 As String, valorA15 As String
    Dim respuesta As VbMsgBoxResult

    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, seleccione una o más filas en la hoja correspondiente.", vbExclamation
        Exit Sub
    End If

    Set wsETQ = ThisWorkbook.Sheets("ETQ")

    Select Case Selection.Worksheet.Name
        Case "Consolidado"
            Set hojaOrigen = ThisWorkbook.Sheets("Consolidado")
        Case "IMPRIMIR"
            Set hojaOrigen = ThisWorkbook.Sheets("IMPRIMIR")
        Case Else
            MsgBox "Debe seleccionar filas en 'Consolidado' o 'IMPRIMIR'.", vbExclamation
            Exit Sub
    End Select

    For Each celda In Selection.Rows

        filaSeleccionada = celda.Row

        ' fila de encabezados y filas ocultas por el filtro: no generan etiqueta
        If filaSeleccionada > 1 Then
            If Not hojaOrigen.Rows(filaSeleccionada).Hidden Then

                wsETQ.Range("A2").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 2).Value)

                valorA6 = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 3).Value)
                If valorA6 <> "" Then valorA6 = "*" & UCase$(valorA6) & "*"
                wsETQ.Range("A6").Value = valorA6

                wsETQ.Range("A9").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 3).Value)
                wsETQ.Range("A11").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 4).Value)
                wsETQ.Range("H10").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 13).Value)
                wsETQ.Range("I10").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 12).Value)
                wsETQ.Range("H3").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 9).Value)
                wsETQ.Range("A17").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 5).Value)
                wsETQ.Range("J10").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 10).Value)

                valorA15 = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 5).Value)
                If valorA15 <> "" Then valorA15 = "*" & UCase$(valorA15) & "*"
                wsETQ.Range("A15").Value = valorA15

                wsETQ.Range("H14").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 8).Value)
                wsETQ.Range("H5").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 7).Value)

                filasProcesadas = filasProcesadas + 1

                If filasProcesadas = 1 Then
                    wsETQ.PrintPreview
                    respuesta = MsgBox("¿Desea continuar con la impresión de las etiquetas?", _
                                       vbYesNo + vbQuestion, "Confirmar impresión")
                    If respuesta = vbNo Then Exit Sub
                End If

                wsETQ.PrintOut

            End If
        End If
    Next celda

    If filasProcesadas = 0 Then
        MsgBox "No había filas visibles seleccionadas para imprimir.", vbExclamation, "Impresión"
    Else
        MsgBox "Se enviaron " & filasProcesadas & " etiquetas a imprimir.", _
               vbInformation, "Impresión completada"
    End If
End Sub


Function ValidarDato(valor)

    If IsError(valor) Then
        ValidarDato = ""
    ElseIf IsNull(valor) Then
        ValidarDato = ""
    ElseIf IsEmpty(valor) Then
        ValidarDato = ""
    ElseIf CStr(valor & "") = "" Then
        ValidarDato = ""
    Else
        ValidarDato = valor
    End If
End Function
