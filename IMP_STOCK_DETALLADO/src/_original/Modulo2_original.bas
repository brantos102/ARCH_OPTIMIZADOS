Attribute VB_Name = "Módulo2"
Option Explicit

Sub Checkboxes_masivos()
Dim celda As Range

For Each celda In Selection

    ActiveSheet.CheckBoxes.Add(celda.Left, celda.Top, 72, 17.25).Select
    
        With Selection
            .Caption = ""
            .LinkedCell = celda.Address
            .Display3DShading = False
        
        End With
        
celda.Value = False
celda.Font.Color = vbWhite

Next celda

End Sub
