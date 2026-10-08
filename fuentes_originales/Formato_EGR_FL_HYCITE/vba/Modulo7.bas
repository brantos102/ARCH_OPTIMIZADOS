Sub TABLAS()
Attribute TABLAS.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("TABLAS DINAMICAS").Select
    Range("B3").Select
    ActiveSheet.PivotTables("TablaDinámica2").PivotCache.Refresh
    Range("H3").Select
    ActiveSheet.PivotTables("TablaDinámica1").PivotCache.Refresh
    Range("M3").Select
    ActiveSheet.PivotTables("TablaDinámica7").PivotCache.Refresh
    Range("P3").Select
    ActiveSheet.PivotTables("TablaDinámica4").PivotCache.Refresh
    Sheets("EMPAQUETADO").Select
End Sub
Sub TABLA_APIS()
    Sheets("ITEMS APIS").PivotTables("TablaDinámica3").PivotCache.Refresh
End Sub
Sub ActualizarTodo()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cn As WorkbookConnection

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1) Refrescar datos externos (consultas/APIs) de forma SÍNCRONA
    For Each cn In ThisWorkbook.Connections
        On Error Resume Next
        cn.OLEDBConnection.BackgroundQuery = False
        cn.ODBCConnection.BackgroundQuery = False
        On Error GoTo 0
        cn.Refresh
    Next cn

    ' 2) Refrescar TODAS las tablas dinámicas de TODAS las hojas
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Datos y tablas dinámicas actualizados correctamente.", vbInformation, "Listo"
End Sub
