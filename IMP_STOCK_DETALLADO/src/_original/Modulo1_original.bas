Attribute VB_Name = "Módulo1"
Sub VistaPreviaEImpresion()
Attribute VistaPreviaEImpresion.VB_ProcData.VB_Invoke_Func = "P\n14"
    Dim wsStock As Worksheet, wsImprimir As Worksheet, wsETQ As Worksheet
    Dim celda As Range
    Dim filaSeleccionada As Long
    Dim filasProcesadas As Integer
    Dim valorA6 As String, valorA15 As String
    Dim hojaOrigen As Worksheet

    ' Definir hojas
    Set wsStock = ThisWorkbook.Sheets("Consolidado")
    Set wsImprimir = ThisWorkbook.Sheets("IMPRIMIR")
    Set wsETQ = ThisWorkbook.Sheets("ETQ")
    
    ' Determinar de qué hoja proviene la selección
    Select Case Selection.Worksheet.Name
        Case "Consolidado"
            Set hojaOrigen = wsStock
        Case "IMPRIMIR"
            Set hojaOrigen = wsImprimir
        Case Else
            MsgBox "Debe seleccionar filas en 'STOCK DETALLADO UIO' o 'IMPRIMIR'.", vbExclamation
            Exit Sub
    End Select

    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, seleccione una o más filas en la hoja correspondiente.", vbExclamation
        Exit Sub
    End If

    filasProcesadas = 0

    For Each celda In Selection.Rows
        filaSeleccionada = celda.Row

        If filaSeleccionada = 1 Then GoTo SiguienteFila

        wsETQ.Activate

        wsETQ.Range("A2").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 2).Value)
        
        ' Obtener los valores y agregar asteriscos solo si no están vacíos
        valorA6 = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 3).Value)
        If valorA6 <> "" Then valorA6 = "*" & valorA6 & "*"
        wsETQ.Range("A6").Value = valorA6

        wsETQ.Range("A9").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 3).Value)
        wsETQ.Range("A11").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 4).Value)
        wsETQ.Range("H10").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 13).Value)
        wsETQ.Range("I10").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 12).Value)
        wsETQ.Range("H3").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 9).Value)
        wsETQ.Range("A17").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 5).Value)
        wsETQ.Range("J10").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 10).Value)

        ' Obtener el valor de A15 y agregar asteriscos solo si no está vacío
        valorA15 = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 5).Value)
        If valorA15 <> "" Then valorA15 = "*" & valorA15 & "*"
        wsETQ.Range("A15").Value = valorA15

        wsETQ.Range("H14").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 8).Value)
        wsETQ.Range("H5").Value = ValidarDato(hojaOrigen.Cells(filaSeleccionada, 7).Value)

        filasProcesadas = filasProcesadas + 1

        If filasProcesadas = 1 Then
            wsETQ.PrintPreview
            Dim respuesta As VbMsgBoxResult
            respuesta = MsgBox("¿Desea continuar con la impresión de las demás etiquetas?", vbYesNo + vbQuestion, "Confirmar impresión")
            If respuesta = vbNo Then Exit Sub
        Else
            wsETQ.PrintOut
        End If

SiguienteFila:
    Next celda

    MsgBox "Se enviaron " & filasProcesadas & " etiquetas a imprimir.", vbInformation, "Impresión completada"
End Sub

Function ValidarDato(valor)
    If IsError(valor) Or valor = "" Then
        ValidarDato = ""
    Else
        ValidarDato = valor
    End If
End Function


