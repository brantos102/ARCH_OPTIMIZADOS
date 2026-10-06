Attribute VB_Name = "Módulo1"
Option Explicit

'==============================================================================================
' Módulo1
'----------------------------------------------------------------------------------------------
' La impresión de etiquetas vive ahora en el módulo "modEtiquetas", con UNA sola entrada:
'
'       ImprimirEtiquetas        <- asignarle aquí la combinación de teclas
'                                   (Vista > Macros > Opciones...)
'       EtiquetasDiagnostico     <- herramienta de verificación, no imprime
'
' La antigua VistaPreviaEImpresion se eliminó a propósito: imprimía una etiqueta por trabajo
' de impresión (saturaba la cola de la Zebra), no respetaba el filtro de "Consolidado" y
' perdía la primera etiqueta en la vista previa. El código original queda guardado en
' src/_original/ del repositorio por si hiciera falta consultarlo.
'
' Aquí sólo queda ValidarDato, por si alguna fórmula de alguna hoja la usa.
' Corregida: VBA evalúa los dos lados de "Or", así que el "IsError(valor) Or valor = """""
' original lanzaba error 13 en cuanto la celda tenía #N/A o #REF!.
'==============================================================================================

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
