Attribute VB_Name = "Module1"
Sub AutoFitAllSheets()
    Dim ws As Worksheet
    Dim lastCol As Long
    
    ' Перебор всех листов в активной книге
    For Each ws In ThisWorkbook.Worksheets
        ' Найти последний заполненный столбец на листе
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Автоподбор ширины всех заполненных столбцов
        ws.Columns("A:" & Split(ws.Cells(1, lastCol).Address, "$")(1)).AutoFit
    Next ws
End Sub

